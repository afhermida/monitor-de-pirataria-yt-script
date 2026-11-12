#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Monitor de possíveis audiolivros não autorizados no YouTube.

Lê uma planilha XLSX com títulos e autores, consulta a API oficial do
YouTube Data API v3, calcula um score heurístico e gera um relatório XLSX
no formato da planilha de referência fornecida pelo usuário.

O programa NÃO baixa áudio ou vídeo. Ele serve para triagem de proteção de
catálogo; todo resultado deve passar por revisão humana antes de qualquer
medida jurídica ou pedido de remoção.

Exemplo com uma chave:
    export YOUTUBE_API_KEY="sua-chave"
    python monitor_youtube_audiobooks.py catalogo.xlsx \
        -o relatorio_pirataria_youtube.xlsx \
        --profile equilibrado \
        --authorized-channels canais_autorizados.txt

Exemplo com várias chaves e rotação automática:
    export YOUTUBE_API_KEYS="chave-1,chave-2,chave-3"
    python monitor_youtube_audiobooks.py catalogo.xlsx \
        -o relatorio_pirataria_youtube.xlsx

Para reproduzir mais de perto a lógica da planilha de referência:
    python monitor_youtube_audiobooks.py catalogo.xlsx --profile planilha
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import re
import sys
import time
import unicodedata
from dataclasses import asdict, dataclass, field
from datetime import datetime, timedelta, timezone
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence
from zoneinfo import ZoneInfo

SCRIPT_VERSION = "1.1.0"
CACHE_VERSION = 1
DEFAULT_TIMEZONE = "America/Sao_Paulo"
MAX_API_KEYS = 10
DEFAULT_SEARCH_BUDGET_PER_KEY = 95

# Pesos reconstruídos da planilha de referência.
WEIGHTS = {
    "title_piracy_phrase": 0.35,
    "suspicious_channel": 0.25,
    "long_duration": 0.20,
    "unapproved_channel": 0.10,
    "book_title_match": 0.10,
    "description_phrase": 0.10,
}

TITLE_PIRACY_PHRASES: tuple[tuple[str, str], ...] = (
    ("audiobook completo", "audiobook completo"),
    ("audio book completo", "audiobook completo"),
    ("audiolivro completo", "livro completo"),
    ("audio livro completo", "livro completo"),
    ("livro completo", "livro completo"),
    ("full audiobook", "full audiobook"),
    ("complete audiobook", "complete audiobook"),
    ("audiobook full", "full audiobook"),
    ("full audio book", "full audiobook"),
)

DESCRIPTION_PHRASES: tuple[tuple[str, str], ...] = (
    ("completo", "completo"),
    ("full", "full"),
    ("gratuito", "gratuito"),
    ("gratis", "grátis"),
    ("free", "free"),
    ("download", "download"),
    ("mp3", "mp3"),
)

SUMMARY_PHRASES: tuple[str, ...] = (
    "resumo",
    "resenha",
    "analise",
    "review",
    "explicado",
    "explicacao",
    "sinopse",
    "comentado",
    "comentario sobre",
)

AUDIOBOOK_MARKERS: tuple[str, ...] = (
    "audiobook",
    "audio book",
    "audiolivro",
    "audio livro",
    "livro narrado",
    "narracao humana",
    "voz humana",
)

DEFAULT_SUSPICIOUS_CHANNEL_TERMS: tuple[str, ...] = (
    "audio livros",
    "audiolivros",
    "audiobooks",
    "audiobook br",
    "acervo audio",
    "livros narrados",
    "ouca livros",
    "voz das letras",
    "livros gratis",
    "livros gratuitos",
)

HEADER_TITLE_ALIASES = {
    "titulo",
    "titulo do livro",
    "livro",
    "nome do livro",
    "title",
    "book title",
}
HEADER_AUTHOR_ALIASES = {
    "autor",
    "autora",
    "autores",
    "author",
    "authors",
}

STOPWORDS = {
    "a",
    "o",
    "as",
    "os",
    "um",
    "uma",
    "uns",
    "umas",
    "de",
    "da",
    "do",
    "das",
    "dos",
    "e",
    "em",
    "no",
    "na",
    "nos",
    "nas",
    "por",
    "para",
    "com",
    "sem",
    "the",
    "of",
    "and",
    "a",
    "an",
    "in",
    "on",
    "to",
    "for",
}

EDITION_MARKERS = {
    "edicao",
    "nova edicao",
    "bolso",
    "luxo",
    "traducao",
    "traducao direta",
    "revista",
    "ampliada",
    "volume",
    "vol",
    "capa dura",
    "ed especial",
    "edicao especial",
}

ISO_DURATION_RE = re.compile(
    r"^P(?:(?P<days>\d+)D)?(?:T(?:(?P<hours>\d+)H)?(?:(?P<minutes>\d+)M)?(?:(?P<seconds>\d+)S)?)?$"
)


class MonitorError(RuntimeError):
    """Erro esperado e apresentável ao usuário."""


class QuotaReached(MonitorError):
    """Todas as chaves disponíveis ficaram indisponíveis para a execução."""


@dataclass(frozen=True)
class Book:
    position: int
    title: str
    author: str

    @property
    def identity(self) -> str:
        return f"{normalize(self.title)}|{normalize(self.author)}"


@dataclass
class Video:
    video_id: str
    title: str
    description: str
    channel_id: str
    channel_title: str
    published_at: str
    duration_seconds: int
    view_count: int
    search_rank: int = 0

    @property
    def url(self) -> str:
        return f"https://www.youtube.com/watch?v={self.video_id}"


@dataclass
class Evaluation:
    book: Book
    video: Video
    score: float
    reasons: list[str]
    title_match_score: float
    author_match: bool
    accepted: bool
    rejected_reason: str | None = None


@dataclass
class ScanState:
    results_by_book: dict[str, list[Video]] = field(default_factory=dict)
    pending_books: set[str] = field(default_factory=set)
    api_search_calls: int = 0
    cached_books: int = 0
    api_keys_configured: int = 0
    api_keys_used: int = 0
    api_keys_unavailable: int = 0
    api_key_rotations: int = 0


# ---------------------------------------------------------------------------
# Texto, correspondência e heurísticas
# ---------------------------------------------------------------------------


def normalize(value: Any) -> str:
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower().replace("&", " e ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def clean_display_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def meaningful_tokens(value: str) -> list[str]:
    return [
        token
        for token in normalize(value).split()
        if token not in STOPWORDS and len(token) >= 2
    ]


def title_aliases(title: str) -> list[str]:
    """Gera formas do título sem informações editoriais e com barras separadas."""
    original = clean_display_text(title)
    no_parentheses = re.sub(r"\s*\([^)]*\)\s*", " ", original).strip()

    candidates = [no_parentheses]
    if ":" in no_parentheses:
        head, tail = no_parentheses.split(":", 1)
        if any(marker in normalize(tail) for marker in EDITION_MARKERS):
            candidates.append(head.strip())

    expanded: list[str] = []
    for candidate in candidates:
        expanded.append(candidate)
        expanded.extend(part.strip() for part in re.split(r"\s*/\s*", candidate) if part.strip())

    normalized_unique: list[str] = []
    seen: set[str] = set()
    for candidate in expanded:
        n = normalize(candidate)
        if n and n not in seen:
            seen.add(n)
            normalized_unique.append(n)
    return normalized_unique or [normalize(title)]


def author_tokens(author: str) -> list[str]:
    tokens = meaningful_tokens(author)
    if not tokens:
        return []
    # Sobrenomes são mais discriminantes; ainda mantemos os demais para autores compostos.
    ordered = [tokens[-1], *tokens[:-1]]
    return list(dict.fromkeys(ordered))


def contains_token(text_normalized: str, token: str) -> bool:
    return bool(re.search(rf"(?<![a-z0-9]){re.escape(token)}(?![a-z0-9])", text_normalized))


def author_matches(author: str, video_title: str, description: str) -> bool:
    haystack = f"{normalize(video_title)} {normalize(description)}"
    tokens = author_tokens(author)
    if not tokens:
        return False
    # O sobrenome basta; para autores de um único token, exige correspondência exata.
    return any(contains_token(haystack, token) for token in tokens[:2])


def compute_title_match(book_title: str, video_title: str) -> tuple[float, bool, str]:
    target = normalize(video_title)
    best_score = 0.0
    exact = False
    best_alias = ""

    for alias in title_aliases(book_title):
        alias_tokens = [t for t in alias.split() if t not in STOPWORDS]
        if not alias_tokens:
            continue

        if alias in target:
            score = 1.0
            is_exact = True
        else:
            matched = sum(1 for token in alias_tokens if contains_token(target, token))
            coverage = matched / len(alias_tokens)
            char_ratio = SequenceMatcher(None, alias, target).ratio()
            compact_target = " ".join(token for token in target.split() if token not in STOPWORDS)
            token_ratio = SequenceMatcher(None, " ".join(alias_tokens), compact_target).ratio()
            score = max(coverage * 0.82 + token_ratio * 0.18, char_ratio * 0.85)
            is_exact = False

        if score > best_score:
            best_score = score
            exact = is_exact
            best_alias = alias

    return round(best_score, 4), exact, best_alias


def bibliographic_gate(
    book: Book,
    video: Video,
    title_match_score: float,
    exact_title_match: bool,
    author_match: bool,
) -> tuple[bool, str | None]:
    """Bloqueia falsos positivos bibliográficos no perfil equilibrado."""
    aliases = title_aliases(book.title)
    alias = max(aliases, key=len) if aliases else normalize(book.title)
    tokens = [t for t in alias.split() if t not in STOPWORDS]
    compact_len = len(alias.replace(" ", ""))

    if not tokens:
        return False, "título bibliográfico vazio após normalização"

    if len(tokens) >= 4:
        ok = title_match_score >= 0.58 or (title_match_score >= 0.48 and author_match)
    elif len(tokens) == 3:
        ok = exact_title_match or title_match_score >= 0.68 or (
            title_match_score >= 0.55 and author_match
        )
    elif len(tokens) == 2:
        ok = exact_title_match and (author_match or compact_len >= 11)
    else:
        token = tokens[0]
        distinctive_numeric = token.isdigit() and len(token) >= 3
        ok = exact_title_match and (author_match or distinctive_numeric)

    if ok:
        return True, None
    return False, "correspondência bibliográfica fraca"


def first_phrase(text: str, phrases: Sequence[tuple[str, str]]) -> tuple[str, str] | None:
    normalized = normalize(text)
    for normalized_phrase, display_phrase in phrases:
        if normalized_phrase in normalized:
            return normalized_phrase, display_phrase
    return None


def channel_is_suspicious(channel_title: str, suspicious_terms: Sequence[str]) -> bool:
    normalized = normalize(channel_title)
    return any(normalize(term) in normalized for term in suspicious_terms if normalize(term))


def load_channel_rules(path: Path | None) -> tuple[set[str], set[str]]:
    """Lê IDs e nomes autorizados. Linhas podem usar `id:` ou `nome:`."""
    ids: set[str] = set()
    names: set[str] = set()
    if path is None:
        return ids, names
    if not path.exists():
        raise MonitorError(f"Arquivo de canais autorizados não encontrado: {path}")

    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.lower().startswith("id:"):
            value = line.split(":", 1)[1].strip()
            if value:
                ids.add(value)
        elif line.lower().startswith(("nome:", "name:")):
            value = line.split(":", 1)[1].strip()
            if value:
                names.add(normalize(value))
        elif line.startswith("UC") and len(line) >= 20:
            ids.add(line)
        else:
            names.add(normalize(line))
    return ids, names


def is_authorized_channel(video: Video, authorized_ids: set[str], authorized_names: set[str]) -> bool:
    return video.channel_id in authorized_ids or normalize(video.channel_title) in authorized_names


def parse_iso_duration(value: str | None) -> int:
    if not value:
        return 0
    match = ISO_DURATION_RE.match(value)
    if not match:
        return 0
    parts = {name: int(raw or 0) for name, raw in match.groupdict().items()}
    return (
        parts["days"] * 86400
        + parts["hours"] * 3600
        + parts["minutes"] * 60
        + parts["seconds"]
    )


def format_duration(seconds: int) -> str:
    seconds = max(0, int(seconds or 0))
    total_minutes = seconds // 60
    hours, minutes = divmod(total_minutes, 60)
    if hours:
        return f"{hours}h{minutes}m"
    return f"{minutes}m"


def evaluate_video(
    book: Book,
    video: Video,
    *,
    profile: str,
    threshold: float,
    long_duration_minutes: int,
    authorized_ids: set[str],
    authorized_names: set[str],
    suspicious_channel_terms: Sequence[str],
    include_summaries: bool,
) -> Evaluation:
    reasons: list[str] = []
    score = 0.0

    authorized = is_authorized_channel(video, authorized_ids, authorized_names)
    if authorized:
        return Evaluation(
            book=book,
            video=video,
            score=0.0,
            reasons=["canal autorizado"],
            title_match_score=0.0,
            author_match=False,
            accepted=False,
            rejected_reason="canal autorizado",
        )

    title_match_score, exact_title_match, _ = compute_title_match(book.title, video.title)
    author_match = author_matches(book.author, video.title, video.description)

    title_phrase = first_phrase(video.title, TITLE_PIRACY_PHRASES)
    if title_phrase:
        score += WEIGHTS["title_piracy_phrase"]
        reasons.append(f'título contém "{title_phrase[1]}"')

    if channel_is_suspicious(video.channel_title, suspicious_channel_terms):
        score += WEIGHTS["suspicious_channel"]
        reasons.append(f'canal suspeito: "{normalize(video.channel_title)}"')

    duration_minutes = video.duration_seconds // 60
    if duration_minutes >= long_duration_minutes:
        score += WEIGHTS["long_duration"]
        reasons.append(f"duração longa: {duration_minutes} min")

    # A API pública não fornece um selo de verificação utilizável para esta finalidade.
    # Aqui, "não verificado" significa simplesmente "fora da allowlist do titular".
    score += WEIGHTS["unapproved_channel"]
    reasons.append("canal não verificado")

    title_reason_match = exact_title_match or title_match_score >= 0.72
    if title_reason_match:
        score += WEIGHTS["book_title_match"]
        reasons.append("título do livro no vídeo")

    description_phrase = first_phrase(video.description, DESCRIPTION_PHRASES)
    if description_phrase:
        score += WEIGHTS["description_phrase"]
        reasons.append(f'descrição contém "{description_phrase[1]}"')

    rejected_reason: str | None = None

    if profile == "equilibrado":
        gate_ok, gate_reason = bibliographic_gate(
            book, video, title_match_score, exact_title_match, author_match
        )
        if not gate_ok:
            rejected_reason = gate_reason

        normalized_title = normalize(video.title)
        has_summary_signal = any(phrase in normalized_title for phrase in SUMMARY_PHRASES)
        if has_summary_signal and not include_summaries:
            score -= 0.30
            reasons.append("possível resumo/resenha: -0,30")

        has_audio_marker = any(marker in normalized_title for marker in AUDIOBOOK_MARKERS)
        if (
            duration_minutes < 45
            and not title_phrase
            and not has_audio_marker
            and not include_summaries
        ):
            score -= 0.25
            reasons.append("vídeo curto sem indicação clara de audiobook: -0,25")

    score = round(min(1.0, max(0.0, score)) + 1e-9, 2)
    accepted = rejected_reason is None and score >= threshold

    return Evaluation(
        book=book,
        video=video,
        score=score,
        reasons=reasons,
        title_match_score=title_match_score,
        author_match=author_match,
        accepted=accepted,
        rejected_reason=rejected_reason,
    )


# ---------------------------------------------------------------------------
# Leitura do catálogo
# ---------------------------------------------------------------------------


def detect_header(
    rows: Iterable[Sequence[Any]],
    *,
    explicit_title_column: str | None,
    explicit_author_column: str | None,
) -> tuple[int, int, int]:
    """Retorna (linha do cabeçalho, índice título, índice autor), tudo zero-based."""
    title_aliases_set = {normalize(x) for x in HEADER_TITLE_ALIASES}
    author_aliases_set = {normalize(x) for x in HEADER_AUTHOR_ALIASES}

    if explicit_title_column or explicit_author_column:
        if not (explicit_title_column and explicit_author_column):
            raise MonitorError(
                "Use --title-column e --author-column juntos, ou nenhum dos dois."
            )
        explicit_title = normalize(explicit_title_column)
        explicit_author = normalize(explicit_author_column)
    else:
        explicit_title = explicit_author = ""

    for row_index, row in enumerate(rows):
        normalized_cells = [normalize(cell) for cell in row]
        title_index = -1
        author_index = -1
        for col_index, value in enumerate(normalized_cells):
            if value == explicit_title or (not explicit_title and value in title_aliases_set):
                title_index = col_index
            if value == explicit_author or (not explicit_author and value in author_aliases_set):
                author_index = col_index
        if title_index >= 0 and author_index >= 0:
            return row_index, title_index, author_index

    raise MonitorError(
        "Não encontrei colunas de título e autor nas primeiras linhas. "
        "Informe --title-column e --author-column se os nomes forem diferentes."
    )


def read_catalog(
    input_path: Path,
    *,
    sheet_name: str | None,
    title_column: str | None,
    author_column: str | None,
) -> list[Book]:
    try:
        from openpyxl import load_workbook
    except ImportError as exc:  # pragma: no cover - mensagem operacional
        raise MonitorError("Dependência ausente: instale openpyxl.") from exc

    if not input_path.exists():
        raise MonitorError(f"Planilha de entrada não encontrada: {input_path}")

    workbook = load_workbook(input_path, read_only=True, data_only=True)
    try:
        candidate_sheets = [workbook[sheet_name]] if sheet_name else list(workbook.worksheets)
    except KeyError as exc:
        raise MonitorError(f"Aba inexistente: {sheet_name}") from exc

    detected_sheets: list[tuple[int, int, Any, tuple[int, int, int]]] = []

    for sheet in candidate_sheets:
        buffered_rows = [
            tuple(row)
            for row in sheet.iter_rows(min_row=1, max_row=20, values_only=True)
        ]
        try:
            candidate_header = detect_header(
                buffered_rows,
                explicit_title_column=title_column,
                explicit_author_column=author_column,
            )
        except MonitorError:
            continue
        normalized_name = normalize(sheet.title)
        preferred = int(
            any(
                marker in normalized_name
                for marker in ("catalogo", "catalog", "livros", "titulos")
            )
        )
        detected_sheets.append((preferred, int(sheet.max_row or 0), sheet, candidate_header))

    if not detected_sheets:
        chosen_sheet = None
        header_info = None
    else:
        _, _, chosen_sheet, header_info = max(
            detected_sheets, key=lambda item: (item[0], item[1])
        )

    if chosen_sheet is None or header_info is None:
        workbook.close()
        raise MonitorError(
            "Nenhuma aba contém um cabeçalho reconhecível de título e autor."
        )

    header_row_zero, title_index, author_index = header_info
    books: list[Book] = []
    seen: set[str] = set()

    for row in chosen_sheet.iter_rows(min_row=header_row_zero + 2, values_only=True):
        title = clean_display_text(row[title_index] if title_index < len(row) else None)
        author = clean_display_text(row[author_index] if author_index < len(row) else None)
        if not title:
            continue
        key = f"{normalize(title)}|{normalize(author)}"
        if key in seen:
            continue
        seen.add(key)
        books.append(Book(position=len(books) + 1, title=title, author=author))

    workbook.close()
    if not books:
        raise MonitorError("A planilha não contém títulos válidos após o cabeçalho.")
    return books


# ---------------------------------------------------------------------------
# YouTube API e cache
# ---------------------------------------------------------------------------


def build_query(book: Book) -> str:
    display_alias = title_aliases(book.title)[0]
    author = clean_display_text(book.author)
    # O operador | é OR na API do YouTube. A biblioteca cuida da URL-encoding.
    return f'"{display_alias}" "{author}" audiobook|audiolivro'.strip()


def cache_key(
    book: Book,
    *,
    query: str,
    max_results: int,
    region: str,
    language: str,
    published_after: str | None,
) -> str:
    payload = "|".join(
        [
            book.identity,
            query,
            str(max_results),
            region,
            language,
            published_after or "",
        ]
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def load_cache(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {"version": CACHE_VERSION, "entries": {}}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise MonitorError(f"Não foi possível ler o cache {path}: {exc}") from exc
    if data.get("version") != CACHE_VERSION or not isinstance(data.get("entries"), dict):
        return {"version": CACHE_VERSION, "entries": {}}
    return data


def save_cache(path: Path, cache: Mapping[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(
        json.dumps(cache, ensure_ascii=False, indent=2, sort_keys=True),
        encoding="utf-8",
    )
    temp.replace(path)


def parse_cached_videos(raw_videos: Sequence[Mapping[str, Any]]) -> list[Video]:
    videos: list[Video] = []
    for item in raw_videos:
        try:
            videos.append(
                Video(
                    video_id=str(item["video_id"]),
                    title=str(item.get("title", "")),
                    description=str(item.get("description", "")),
                    channel_id=str(item.get("channel_id", "")),
                    channel_title=str(item.get("channel_title", "")),
                    published_at=str(item.get("published_at", "")),
                    duration_seconds=int(item.get("duration_seconds", 0) or 0),
                    view_count=int(item.get("view_count", 0) or 0),
                    search_rank=int(item.get("search_rank", 0) or 0),
                )
            )
        except (KeyError, TypeError, ValueError):
            continue
    return videos


def cache_is_fresh(entry: Mapping[str, Any], ttl_days: int) -> bool:
    if ttl_days < 0:
        return True
    fetched_at = entry.get("fetched_at")
    if not fetched_at:
        return False
    try:
        fetched = datetime.fromisoformat(str(fetched_at).replace("Z", "+00:00"))
    except ValueError:
        return False
    if fetched.tzinfo is None:
        fetched = fetched.replace(tzinfo=timezone.utc)
    return datetime.now(timezone.utc) - fetched <= timedelta(days=ttl_days)


def create_youtube_client(api_key: str):
    try:
        from googleapiclient.discovery import build
    except ImportError as exc:  # pragma: no cover - mensagem operacional
        raise MonitorError(
            "Dependência ausente: instale google-api-python-client."
        ) from exc
    return build("youtube", "v3", developerKey=api_key, cache_discovery=False)


def api_error_reasons(exc: Exception) -> set[str]:
    """Extrai códigos de erro do corpo JSON de um HttpError, quando disponível."""
    reasons: set[str] = set()
    content = getattr(exc, "content", None)
    if isinstance(content, bytes):
        content = content.decode("utf-8", errors="replace")
    if not content:
        return reasons

    try:
        payload = json.loads(str(content))
    except (TypeError, ValueError, json.JSONDecodeError):
        return reasons
    if not isinstance(payload, Mapping):
        return reasons

    error = payload.get("error", {})
    if not isinstance(error, Mapping):
        return reasons
    status = error.get("status")
    if status:
        reasons.add(str(status).casefold())
    details = error.get("errors", [])
    if isinstance(details, Sequence) and not isinstance(details, (str, bytes)):
        for detail in details:
            if isinstance(detail, Mapping) and detail.get("reason"):
                reasons.add(str(detail["reason"]).casefold())
    return reasons


def is_quota_error(exc: Exception) -> bool:
    status = getattr(getattr(exc, "resp", None), "status", None)
    text = str(exc).casefold()
    reasons = api_error_reasons(exc)
    quota_reasons = {
        "quotaexceeded",
        "dailylimitexceeded",
        "dailylimitexceededunreg",
        "ratelimitexceeded",
        "userratelimitexceeded",
        "uploadlimitexceeded",
        "resource_exhausted",
    }
    if status == 429:
        return True
    return status == 403 and (
        bool(reasons & quota_reasons)
        or any(
            marker in text
            for marker in ("quota", "ratelimit", "rate limit", "dailylimit")
        )
    )


def is_key_configuration_error(exc: Exception) -> bool:
    """Identifica chave inválida, API desativada ou restrição incompatível."""
    status = getattr(getattr(exc, "resp", None), "status", None)
    text = str(exc).casefold()
    reasons = api_error_reasons(exc)
    key_reasons = {
        "accessnotconfigured",
        "apikeynotvalid",
        "autherror",
        "forbidden",
        "iprefererblocked",
        "keyexpired",
        "keyinvalid",
        "permission_denied",
    }
    if status == 401:
        return True
    return status in {400, 403} and (
        bool(reasons & key_reasons)
        or any(
            marker in text
            for marker in (
                "api key not valid",
                "keyinvalid",
                "access not configured",
                "accessnotconfigured",
                "iprefererblocked",
                "permission_denied",
                "forbidden",
            )
        )
    )


def redact_api_keys(text: str, api_keys: Sequence[str]) -> str:
    """Remove valores de chave de mensagens de erro antes de exibi-las."""
    redacted = str(text)
    for api_key in api_keys:
        if api_key:
            redacted = redacted.replace(api_key, "[CHAVE_OCULTA]")
    return re.sub(
        r"([?&]key=)[^&\s]+",
        r"\1[CHAVE_OCULTA]",
        redacted,
        flags=re.IGNORECASE,
    )


class YouTubeKeyPool:
    """Gerencia até dez chaves e troca de chave sem revelar seus valores."""

    def __init__(self, api_keys: Sequence[str], *, client_factory=None) -> None:
        keys = [str(key).strip() for key in api_keys if str(key).strip()]
        keys = list(dict.fromkeys(keys))
        if not keys:
            raise MonitorError("Nenhuma chave de API foi informada.")
        if len(keys) > MAX_API_KEYS:
            raise MonitorError(f"Informe no máximo {MAX_API_KEYS} chaves de API.")

        self.api_keys = keys
        self._client_factory = client_factory or create_youtube_client
        self._clients: dict[int, Any] = {}
        self._current_index = 0
        self._used_indices: set[int] = set()
        self._unavailable_indices: set[int] = set()
        self.rotations = 0

    @property
    def configured_count(self) -> int:
        return len(self.api_keys)

    @property
    def used_count(self) -> int:
        return len(self._used_indices)

    @property
    def unavailable_count(self) -> int:
        return len(self._unavailable_indices)

    def _current_client(self):
        index = self._current_index
        self._used_indices.add(index)
        if index not in self._clients:
            self._clients[index] = self._client_factory(self.api_keys[index])
        return self._clients[index]

    def _disable_current_and_rotate(self, reason: str) -> bool:
        old_index = self._current_index
        self._unavailable_indices.add(old_index)
        next_index = next(
            (
                index
                for index in range(old_index + 1, len(self.api_keys))
                if index not in self._unavailable_indices
            ),
            None,
        )
        if next_index is None:
            print(
                f"Chave {old_index + 1}/{len(self.api_keys)} indisponível ({reason}). "
                "Não há outra chave configurada para esta execução.",
                file=sys.stderr,
            )
            return False

        self._current_index = next_index
        self.rotations += 1
        print(
            f"Chave {old_index + 1}/{len(self.api_keys)} indisponível ({reason}); "
            f"alternando para a chave {next_index + 1}/{len(self.api_keys)}.",
            file=sys.stderr,
        )
        return True

    def execute(self, request_factory, retries: int = 3) -> Mapping[str, Any]:
        """Executa uma chamada, rotacionando em erros de cota ou configuração."""
        while True:
            youtube = self._current_client()
            delay = 1.5
            last_exc: Exception | None = None

            for attempt in range(retries + 1):
                try:
                    return request_factory(youtube).execute()
                except Exception as exc:  # HttpError é opcional em import-time
                    quota_error = is_quota_error(exc)
                    key_error = is_key_configuration_error(exc)
                    if quota_error or key_error:
                        reason = (
                            "cota ou limite atingido"
                            if quota_error
                            else "chave inválida, restrita ou API desativada"
                        )
                        if self._disable_current_and_rotate(reason):
                            break
                        raise QuotaReached(
                            f"Todas as {len(self.api_keys)} chaves de API ficaram "
                            "indisponíveis nesta execução."
                        ) from exc

                    last_exc = exc
                    status = getattr(getattr(exc, "resp", None), "status", None)
                    retryable = status in {429, 500, 502, 503, 504} or status is None
                    if not retryable or attempt >= retries:
                        safe_error = redact_api_keys(str(last_exc), self.api_keys)
                        raise MonitorError(
                            f"Falha ao consultar a API do YouTube: {safe_error}"
                        ) from exc
                    time.sleep(delay)
                    delay *= 2
            else:  # pragma: no cover - o laço sempre retorna ou lança erro
                safe_error = redact_api_keys(str(last_exc), self.api_keys)
                raise MonitorError(f"Falha ao consultar a API do YouTube: {safe_error}")


def search_youtube(
    key_pool: YouTubeKeyPool,
    book: Book,
    *,
    max_results: int,
    region: str,
    language: str,
    published_after: str | None,
) -> list[Video]:
    query = build_query(book)
    request_args: dict[str, Any] = {
        "part": "id",
        "q": query,
        "type": "video",
        "maxResults": max_results,
        "order": "relevance",
        "safeSearch": "none",
        "regionCode": region,
        "relevanceLanguage": language,
    }
    if published_after:
        request_args["publishedAfter"] = published_after

    search_response = key_pool.execute(
        lambda youtube: youtube.search().list(**request_args)
    )
    ordered_ids = [
        item.get("id", {}).get("videoId")
        for item in search_response.get("items", [])
        if item.get("id", {}).get("videoId")
    ]
    ordered_ids = list(dict.fromkeys(ordered_ids))
    if not ordered_ids:
        return []

    details_response = key_pool.execute(
        lambda youtube: youtube.videos().list(
            part="snippet,contentDetails,statistics,status",
            id=",".join(ordered_ids),
            maxResults=len(ordered_ids),
        )
    )
    by_id = {str(item.get("id")): item for item in details_response.get("items", [])}

    videos: list[Video] = []
    for rank, video_id in enumerate(ordered_ids, start=1):
        item = by_id.get(video_id)
        if not item:
            continue
        snippet = item.get("snippet", {})
        details = item.get("contentDetails", {})
        statistics = item.get("statistics", {})
        videos.append(
            Video(
                video_id=video_id,
                title=clean_display_text(snippet.get("title")),
                description=str(snippet.get("description", "") or ""),
                channel_id=str(snippet.get("channelId", "") or ""),
                channel_title=clean_display_text(snippet.get("channelTitle")),
                published_at=str(snippet.get("publishedAt", "") or ""),
                duration_seconds=parse_iso_duration(details.get("duration")),
                view_count=int(statistics.get("viewCount", 0) or 0),
                search_rank=rank,
            )
        )
    return videos


def scan_books(
    books: Sequence[Book],
    *,
    api_keys: Sequence[str],
    cache_path: Path,
    cache_ttl_days: int,
    refresh: bool,
    search_budget: int,
    max_results: int,
    region: str,
    language: str,
    published_after: str | None,
    progress_every: int,
) -> tuple[ScanState, dict[str, Any]]:
    cache = load_cache(cache_path)
    entries: dict[str, Any] = cache.setdefault("entries", {})
    state = ScanState(api_keys_configured=len(api_keys))
    key_pool = YouTubeKeyPool(api_keys) if api_keys else None
    quota_blocked = False

    for index, book in enumerate(books, start=1):
        query = build_query(book)
        key = cache_key(
            book,
            query=query,
            max_results=max_results,
            region=region,
            language=language,
            published_after=published_after,
        )
        entry = entries.get(key)
        fresh_cache = bool(
            entry and not refresh and cache_is_fresh(entry, cache_ttl_days)
        )

        if fresh_cache:
            state.results_by_book[book.identity] = parse_cached_videos(
                entry.get("videos", [])
            )
            state.cached_books += 1
        else:
            budget_exhausted = (
                search_budget > 0 and state.api_search_calls >= search_budget
            )
            cannot_search = budget_exhausted or key_pool is None or quota_blocked

            if cannot_search:
                # Mesmo em refresh parcial ou após esgotar a cota, um cache antigo
                # é melhor do que transformar um título já pesquisado em pendente.
                if entry:
                    state.results_by_book[book.identity] = parse_cached_videos(
                        entry.get("videos", [])
                    )
                    state.cached_books += 1
                else:
                    state.pending_books.add(book.identity)
            else:
                try:
                    videos = search_youtube(
                        key_pool,
                        book,
                        max_results=max_results,
                        region=region,
                        language=language,
                        published_after=published_after,
                    )
                except QuotaReached:
                    quota_blocked = True
                    if entry:
                        state.results_by_book[book.identity] = parse_cached_videos(
                            entry.get("videos", [])
                        )
                        state.cached_books += 1
                    else:
                        state.pending_books.add(book.identity)
                else:
                    state.api_search_calls += 1
                    state.results_by_book[book.identity] = videos
                    entries[key] = {
                        "book": {"title": book.title, "author": book.author},
                        "query": query,
                        "fetched_at": datetime.now(timezone.utc).isoformat(),
                        "videos": [asdict(video) for video in videos],
                    }
                    if state.api_search_calls % 5 == 0:
                        save_cache(cache_path, cache)

        if progress_every > 0 and index % progress_every == 0:
            print(
                f"Processados {index}/{len(books)} títulos "
                f"(API: {state.api_search_calls}, cache: {state.cached_books}, "
                f"pendentes: {len(state.pending_books)})",
                file=sys.stderr,
            )

    save_cache(cache_path, cache)
    if key_pool is not None:
        state.api_keys_used = key_pool.used_count
        state.api_keys_unavailable = key_pool.unavailable_count
        state.api_key_rotations = key_pool.rotations
    return state, cache


# ---------------------------------------------------------------------------
# Relatório XLSX
# ---------------------------------------------------------------------------


def score_font_color(score: float) -> str:
    if score >= 0.75:
        return "D93025"  # vermelho
    if score >= 0.50:
        return "EF6C00"  # laranja
    return "188038"  # verde


def style_table_sheet(ws, widths: Mapping[str, float], freeze: str = "A2") -> None:
    from openpyxl.styles import Alignment, Font, PatternFill

    header_fill = PatternFill("solid", fgColor="28588F")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 36
    ws.freeze_panes = freeze
    ws.auto_filter.ref = ws.dimensions
    for column, width in widths.items():
        ws.column_dimensions[column].width = width


def add_alternating_rows(ws, start_row: int, end_row: int) -> None:
    from openpyxl.styles import PatternFill

    fill = PatternFill("solid", fgColor="EAF3F8")
    for row in range(start_row, end_row + 1):
        if row % 2 == 0:
            for cell in ws[row]:
                cell.fill = fill


def add_hyperlink(cell, url: str, label: str) -> None:
    cell.value = label
    cell.hyperlink = url
    cell.style = "Hyperlink"


def build_report(
    output_path: Path,
    books: Sequence[Book],
    state: ScanState,
    *,
    profile: str,
    threshold: float,
    long_duration_minutes: int,
    authorized_ids: set[str],
    authorized_names: set[str],
    suspicious_channel_terms: Sequence[str],
    include_summaries: bool,
    timezone_name: str,
) -> tuple[list[Evaluation], dict[str, list[Evaluation]]]:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font, PatternFill
    except ImportError as exc:  # pragma: no cover
        raise MonitorError("Dependência ausente: instale openpyxl.") from exc

    evaluations: list[Evaluation] = []
    accepted_by_book: dict[str, list[Evaluation]] = {}

    for book in books:
        raw_videos = state.results_by_book.get(book.identity, [])
        seen_video_ids: set[str] = set()
        for video in raw_videos:
            if video.video_id in seen_video_ids:
                continue
            seen_video_ids.add(video.video_id)
            evaluation = evaluate_video(
                book,
                video,
                profile=profile,
                threshold=threshold,
                long_duration_minutes=long_duration_minutes,
                authorized_ids=authorized_ids,
                authorized_names=authorized_names,
                suspicious_channel_terms=suspicious_channel_terms,
                include_summaries=include_summaries,
            )
            evaluations.append(evaluation)
            if evaluation.accepted:
                accepted_by_book.setdefault(book.identity, []).append(evaluation)

    ranked_books: list[tuple[Book, list[Evaluation]]] = []
    for book in books:
        accepted = accepted_by_book.get(book.identity, [])
        if accepted:
            ranked_books.append((book, accepted))
    ranked_books.sort(
        key=lambda item: (
            max(ev.video.view_count for ev in item[1]),
            max(ev.score for ev in item[1]),
            item[0].title.casefold(),
        ),
        reverse=True,
    )

    workbook = Workbook()
    ws_summary = workbook.active
    ws_summary.title = "Resumo"
    ws_ranking = workbook.create_sheet("Ranking por Views")
    ws_videos = workbook.create_sheet("Vídeos Suspeitos")
    ws_catalog = workbook.create_sheet("Catálogo Completo")

    # Resumo
    try:
        tz = ZoneInfo(timezone_name)
    except Exception:
        tz = timezone.utc
    now = datetime.now(tz)
    total = len(books)
    suspicious_titles = len(accepted_by_book)
    pending = sum(1 for book in books if book.identity in state.pending_books)
    processed = total - pending
    clean_titles = max(0, processed - suspicious_titles)
    percentage = (suspicious_titles / processed) if processed else 0.0

    ws_summary.merge_cells("A1:B1")
    ws_summary["A1"] = "Relatório de Monitoramento de Pirataria — YouTube"
    ws_summary["A1"].font = Font(size=16, bold=True, color="24528A")
    ws_summary["A2"] = f"Gerado em: {now:%d/%m/%Y %H:%M}"
    ws_summary["A2"].font = Font(color="808080")

    summary_rows = [
        ("Total de títulos no catálogo", total),
        ("Títulos processados", processed),
        ("Títulos pendentes", pending),
        ("Títulos com suspeita de pirataria", suspicious_titles),
        ("Títulos sem suspeita", clean_titles),
        ("% dos processados com pirataria suspeita", percentage),
        ("Consultas search.list nesta execução", state.api_search_calls),
        ("Títulos reaproveitados do cache", state.cached_books),
        ("Chaves de API configuradas", state.api_keys_configured),
        ("Chaves de API utilizadas", state.api_keys_used),
        ("Chaves indisponíveis nesta execução", state.api_keys_unavailable),
        ("Trocas automáticas de chave", state.api_key_rotations),
        ("Perfil de classificação", profile),
    ]
    for row_index, (label, value) in enumerate(summary_rows, start=4):
        ws_summary.cell(row=row_index, column=1, value=label)
        ws_summary.cell(row=row_index, column=2, value=value)
        if row_index % 2 == 0:
            for col in range(1, 3):
                ws_summary.cell(row=row_index, column=col).fill = PatternFill(
                    "solid", fgColor="F2F2F2"
                )
        ws_summary.cell(row=row_index, column=2).font = Font(bold=True)
    ws_summary.cell(row=9, column=2).number_format = "0.0%"
    ws_summary.column_dimensions["A"].width = 49
    ws_summary.column_dimensions["B"].width = 20

    # Ranking
    ranking_headers = [
        "Posição",
        "Título",
        "Autor",
        "Qtd suspeitos",
        "Maior nº de views",
        "Score pirataria",
        "URL mais acessada",
        "Razões",
    ]
    ws_ranking.append(ranking_headers)
    for position, (book, accepted) in enumerate(ranked_books, start=1):
        most_viewed = max(
            accepted,
            key=lambda ev: (ev.video.view_count, ev.score, -ev.video.search_rank),
        )
        max_score = max(ev.score for ev in accepted)
        ws_ranking.append(
            [
                position,
                book.title,
                book.author,
                len(accepted),
                most_viewed.video.view_count,
                max_score,
                None,
                " | ".join(most_viewed.reasons),
            ]
        )
        row = ws_ranking.max_row
        add_hyperlink(ws_ranking.cell(row=row, column=7), most_viewed.video.url, "▶ Abrir vídeo")
        ws_ranking.cell(row=row, column=6).font = Font(
            bold=True, color=score_font_color(max_score)
        )
        ws_ranking.cell(row=row, column=8).alignment = Alignment(wrap_text=True, vertical="top")
    style_table_sheet(
        ws_ranking,
        {"A": 10, "B": 52, "C": 31, "D": 15, "E": 19, "F": 16, "G": 20, "H": 72},
    )
    add_alternating_rows(ws_ranking, 2, ws_ranking.max_row)

    # Vídeos suspeitos, agrupados pela ordem do ranking e por views.
    video_headers = [
        "Título do livro",
        "Autor",
        "Título do vídeo",
        "Canal",
        "Views",
        "Duração",
        "Score",
        "Razões",
        "Data publicação",
        "URL",
    ]
    ws_videos.append(video_headers)
    for book, accepted in ranked_books:
        for evaluation in sorted(
            accepted,
            key=lambda ev: (ev.video.view_count, ev.score, -ev.video.search_rank),
            reverse=True,
        ):
            published_date = evaluation.video.published_at[:10]
            ws_videos.append(
                [
                    book.title,
                    book.author,
                    evaluation.video.title,
                    evaluation.video.channel_title,
                    evaluation.video.view_count,
                    format_duration(evaluation.video.duration_seconds),
                    evaluation.score,
                    " | ".join(evaluation.reasons),
                    published_date,
                    None,
                ]
            )
            row = ws_videos.max_row
            add_hyperlink(ws_videos.cell(row=row, column=10), evaluation.video.url, "▶ Abrir")
            ws_videos.cell(row=row, column=7).font = Font(
                bold=True, color=score_font_color(evaluation.score)
            )
            for col in (3, 4, 8):
                ws_videos.cell(row=row, column=col).alignment = Alignment(
                    wrap_text=True, vertical="top"
                )
    style_table_sheet(
        ws_videos,
        {
            "A": 40,
            "B": 28,
            "C": 48,
            "D": 28,
            "E": 14,
            "F": 11,
            "G": 11,
            "H": 64,
            "I": 16,
            "J": 14,
        },
    )
    add_alternating_rows(ws_videos, 2, ws_videos.max_row)

    # Catálogo completo
    catalog_headers = [
        "Título",
        "Autor",
        "Suspeita de pirataria",
        "Qtd suspeitos",
        "URL mais visto",
    ]
    ws_catalog.append(catalog_headers)
    for book in books:
        accepted = accepted_by_book.get(book.identity, [])
        is_pending = book.identity in state.pending_books
        if accepted:
            status = "SIM"
            most_viewed = max(accepted, key=lambda ev: (ev.video.view_count, ev.score))
        elif is_pending:
            status = "pendente"
            most_viewed = None
        else:
            status = "não"
            most_viewed = None

        ws_catalog.append([book.title, book.author, status, len(accepted), None])
        row = ws_catalog.max_row
        status_cell = ws_catalog.cell(row=row, column=3)
        if status == "SIM":
            status_cell.font = Font(bold=True, color="D93025")
        elif status == "pendente":
            status_cell.font = Font(bold=True, color="EF6C00")
        else:
            status_cell.font = Font(color="188038")
        status_cell.alignment = Alignment(horizontal="center")
        if most_viewed:
            add_hyperlink(ws_catalog.cell(row=row, column=5), most_viewed.video.url, "▶ Abrir")

    style_table_sheet(ws_catalog, {"A": 58, "B": 34, "C": 23, "D": 16, "E": 18})
    add_alternating_rows(ws_catalog, 2, ws_catalog.max_row)

    # Ajustes numéricos e de alinhamento.
    for ws in (ws_ranking, ws_videos, ws_catalog):
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                current = cell.alignment
                cell.alignment = Alignment(
                    horizontal=current.horizontal,
                    vertical=current.vertical or "top",
                    text_rotation=current.textRotation,
                    wrap_text=current.wrap_text,
                    shrink_to_fit=current.shrinkToFit,
                    indent=current.indent,
                )

    for cell in ws_ranking["F"][1:]:
        cell.number_format = "0.00"
    for cell in ws_videos["G"][1:]:
        cell.number_format = "0.00"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    return [ev for ev in evaluations if ev.accepted], accepted_by_book


# ---------------------------------------------------------------------------
# CLI e autotestes
# ---------------------------------------------------------------------------


def split_api_key_values(values: Sequence[str]) -> list[str]:
    """Aceita valores repetidos, separados por vírgula, ponto e vírgula ou espaço."""
    keys: list[str] = []
    for raw_value in values:
        for candidate in re.split(r"[\s,;]+", str(raw_value).strip()):
            candidate = candidate.strip()
            if candidate:
                keys.append(candidate)
    return keys


def load_api_keys(
    cli_values: Sequence[str] | None,
    keys_file: Path | None,
    *,
    environment: Mapping[str, str] | None = None,
) -> list[str]:
    """Carrega até dez chaves, priorizando CLI/arquivo sobre variáveis de ambiente."""
    explicit_values = list(cli_values or [])
    if keys_file is not None:
        if not keys_file.exists():
            raise MonitorError(f"Arquivo de chaves não encontrado: {keys_file}")
        for raw_line in keys_file.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if line and not line.startswith("#"):
                explicit_values.append(line)

    if explicit_values:
        raw_values = explicit_values
    else:
        env = os.environ if environment is None else environment
        plural = str(env.get("YOUTUBE_API_KEYS", "") or "").strip()
        singular = str(env.get("YOUTUBE_API_KEY", "") or "").strip()
        raw_values = [plural] if plural else ([singular] if singular else [])

    keys = list(dict.fromkeys(split_api_key_values(raw_values)))
    if len(keys) > MAX_API_KEYS:
        raise MonitorError(
            f"Foram encontradas {len(keys)} chaves; informe no máximo {MAX_API_KEYS}."
        )
    return keys


def parse_published_after(value: str | None) -> str | None:
    if not value:
        return None
    try:
        parsed = datetime.strptime(value, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("Use a data no formato AAAA-MM-DD.") from exc
    return parsed.isoformat().replace("+00:00", "Z")


def load_terms_file(path: Path | None, defaults: Sequence[str]) -> list[str]:
    terms = list(defaults)
    if path is None:
        return terms
    if not path.exists():
        raise MonitorError(f"Arquivo de termos não encontrado: {path}")
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if line and not line.startswith("#"):
            terms.append(line)
    return list(dict.fromkeys(normalize(term) for term in terms if normalize(term)))


def make_test_video(**overrides: Any) -> Video:
    data = {
        "video_id": "test",
        "title": "",
        "description": "",
        "channel_id": "UC_TEST",
        "channel_title": "Canal comum",
        "published_at": "2024-01-01T00:00:00Z",
        "duration_seconds": 0,
        "view_count": 100,
        "search_rank": 1,
    }
    data.update(overrides)
    return Video(**data)


def run_self_test() -> None:
    book = Book(1, "Uma breve história do tempo", "Stephen Hawking")
    common = dict(
        profile="planilha",
        threshold=0.50,
        long_duration_minutes=20,
        authorized_ids=set(),
        authorized_names=set(),
        suspicious_channel_terms=DEFAULT_SUSPICIOUS_CHANNEL_TERMS,
        include_summaries=False,
    )

    cases = [
        (
            make_test_video(
                title="AudioBook Completo - UMA BREVE HISTÓRIA DO TEMPO (Stephen Hawking)",
                duration_seconds=323 * 60,
            ),
            0.75,
        ),
        (
            make_test_video(
                title="Uma breve história do tempo - conversa",
                description="versão full",
                duration_seconds=210 * 60,
            ),
            0.50,
        ),
        (
            make_test_video(
                title="Uma breve história do tempo - voz humana",
                channel_title="Audio Livros",
                duration_seconds=693 * 60,
            ),
            0.65,
        ),
        (
            make_test_video(
                title="Outro livro [Full Audiobook Free]",
                description="full",
                duration_seconds=2 * 60,
            ),
            0.55,
        ),
    ]

    failures: list[str] = []
    for index, (video, expected) in enumerate(cases, start=1):
        result = evaluate_video(book, video, **common)
        if not math.isclose(result.score, expected, abs_tol=1e-9):
            failures.append(f"caso {index}: esperado {expected}, obtido {result.score}")

    # Todos os sinais devem ser limitados a 1,00.
    all_signals = make_test_video(
        title="Uma breve história do tempo - Audiobook Completo",
        description="full completo gratuito",
        channel_title="Acervo Audio Livros",
        duration_seconds=400 * 60,
    )
    capped = evaluate_video(book, all_signals, **common)
    if capped.score != 1.0:
        failures.append(f"limite do score: esperado 1.0, obtido {capped.score}")

    # Carregamento e deduplicação das chaves.
    parsed_keys = load_api_keys(
        ["chave-a,chave-b", "chave-b", "chave-c"],
        None,
        environment={},
    )
    if parsed_keys != ["chave-a", "chave-b", "chave-c"]:
        failures.append(f"leitura de chaves: resultado inesperado {parsed_keys!r}")

    env_keys = load_api_keys(
        None,
        None,
        environment={"YOUTUBE_API_KEYS": "env-a; env-b\nenv-c"},
    )
    if env_keys != ["env-a", "env-b", "env-c"]:
        failures.append(f"variável YOUTUBE_API_KEYS: resultado inesperado {env_keys!r}")

    explicit_keys = load_api_keys(
        ["cli-a"],
        None,
        environment={"YOUTUBE_API_KEYS": "env-a,env-b"},
    )
    if explicit_keys != ["cli-a"]:
        failures.append("chaves explícitas não tiveram prioridade sobre o ambiente")

    try:
        load_api_keys([",".join(f"k{i}" for i in range(11))], None, environment={})
    except MonitorError:
        pass
    else:
        failures.append("limite de 10 chaves não foi aplicado")

    # Simula quota esgotada na primeira chave e sucesso imediato na segunda.
    class FakeResponse:
        status = 403

    class FakeQuotaError(Exception):
        def __init__(self) -> None:
            super().__init__("quotaExceeded")
            self.resp = FakeResponse()
            self.content = json.dumps(
                {
                    "error": {
                        "errors": [{"reason": "quotaExceeded"}],
                        "status": "RESOURCE_EXHAUSTED",
                    }
                }
            ).encode("utf-8")

    class FakeRequest:
        def __init__(self, client: str) -> None:
            self.client = client

        def execute(self) -> Mapping[str, Any]:
            if self.client == "chave-1":
                raise FakeQuotaError()
            return {"client": self.client, "items": []}

    pool = YouTubeKeyPool(
        ["chave-1", "chave-2"],
        client_factory=lambda key: key,
    )
    rotation_result = pool.execute(lambda client: FakeRequest(client), retries=0)
    if rotation_result.get("client") != "chave-2":
        failures.append("rotação automática não utilizou a segunda chave")
    if (pool.rotations, pool.used_count, pool.unavailable_count) != (1, 2, 1):
        failures.append(
            "estatísticas de rotação incorretas: "
            f"{pool.rotations=}, {pool.used_count=}, {pool.unavailable_count=}"
        )

    secret = "AIza-SEGREDO-TESTE"
    redacted = redact_api_keys(
        f"Falha em https://example.test/?key={secret}&x=1; valor={secret}",
        [secret],
    )
    if secret in redacted or "[CHAVE_OCULTA]" not in redacted:
        failures.append("redação de chave em mensagens de erro não funcionou")

    if failures:
        raise MonitorError("Autoteste falhou: " + "; ".join(failures))
    print(
        "Autoteste concluído: calibração e rotação de chaves funcionando "
        "corretamente."
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Lê títulos/autores de uma planilha, busca possíveis audiolivros "
            "não autorizados no YouTube e gera um relatório XLSX."
        )
    )
    parser.add_argument("input", nargs="?", type=Path, help="Planilha XLSX de entrada")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("relatorio_pirataria_youtube.xlsx"),
        help="Arquivo XLSX de saída",
    )
    parser.add_argument("--sheet", help="Nome da aba de entrada")
    parser.add_argument("--title-column", help="Nome exato da coluna de título")
    parser.add_argument("--author-column", help="Nome exato da coluna de autor")
    parser.add_argument(
        "--profile",
        choices=("equilibrado", "planilha"),
        default="equilibrado",
        help=(
            "equilibrado reduz falsos positivos; planilha reproduz de perto a "
            "pontuação da referência"
        ),
    )
    parser.add_argument(
        "--threshold",
        type=float,
        default=0.50,
        help="Score mínimo para classificar como suspeito (padrão: 0.50)",
    )
    parser.add_argument(
        "--long-duration-minutes",
        type=int,
        default=20,
        help="Duração a partir da qual o vídeo recebe o sinal de duração longa",
    )
    parser.add_argument(
        "--max-results",
        type=int,
        default=15,
        choices=range(1, 51),
        metavar="1..50",
        help="Resultados retornados por título em uma única busca",
    )
    parser.add_argument("--region", default="BR", help="Código de região ISO 3166-1")
    parser.add_argument("--language", default="pt", help="Idioma de relevância da busca")
    parser.add_argument(
        "--published-after",
        type=parse_published_after,
        help="Considerar vídeos publicados a partir de AAAA-MM-DD",
    )
    parser.add_argument(
        "--api-key",
        action="append",
        dest="api_keys",
        metavar="CHAVE",
        help=(
            "Chave da API. Pode ser repetido até 10 vezes ou receber valores "
            "separados por vírgula. Sem esta opção, usa YOUTUBE_API_KEYS ou "
            "YOUTUBE_API_KEY."
        ),
    )
    parser.add_argument(
        "--api-keys-file",
        type=Path,
        help="TXT com uma ou mais chaves por linha; comentários começam com #",
    )
    parser.add_argument(
        "--search-budget",
        type=int,
        default=None,
        help=(
            "Máximo total de novos títulos consultados; 0 = sem limite. Se omitido, "
            "usa 95 por chave configurada."
        ),
    )
    parser.add_argument(
        "--cache",
        type=Path,
        default=Path(".youtube_monitor_cache.json"),
        help="Arquivo JSON de cache/checkpoint",
    )
    parser.add_argument(
        "--cache-ttl-days",
        type=int,
        default=30,
        help="Validade do cache em dias; -1 = não expira",
    )
    parser.add_argument("--refresh", action="store_true", help="Ignorar entradas válidas do cache")
    parser.add_argument(
        "--authorized-channels",
        type=Path,
        help="TXT com IDs ou nomes de canais autorizados",
    )
    parser.add_argument(
        "--suspicious-channel-terms",
        type=Path,
        help="TXT com termos adicionais de canal suspeito",
    )
    parser.add_argument(
        "--include-summaries",
        action="store_true",
        help="Não penalizar vídeos identificados como resumo/resenha",
    )
    parser.add_argument(
        "--timezone",
        default=DEFAULT_TIMEZONE,
        help="Fuso horário IANA usado no carimbo do relatório",
    )
    parser.add_argument(
        "--progress-every",
        type=int,
        default=10,
        help="Exibir progresso a cada N títulos; 0 desativa",
    )
    parser.add_argument("--self-test", action="store_true", help="Executar testes de calibração")
    parser.add_argument("--version", action="version", version=SCRIPT_VERSION)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        if args.self_test:
            run_self_test()
            return 0
        if args.input is None:
            parser.error("informe a planilha de entrada ou use --self-test")
        if not 0.0 <= args.threshold <= 1.0:
            raise MonitorError("--threshold deve estar entre 0 e 1.")
        if args.search_budget is not None and args.search_budget < 0:
            raise MonitorError("--search-budget deve ser 0 ou um inteiro positivo.")

        api_keys = load_api_keys(args.api_keys, args.api_keys_file)
        search_budget = args.search_budget
        if search_budget is None:
            search_budget = DEFAULT_SEARCH_BUDGET_PER_KEY * len(api_keys)

        books = read_catalog(
            args.input,
            sheet_name=args.sheet,
            title_column=args.title_column,
            author_column=args.author_column,
        )
        authorized_ids, authorized_names = load_channel_rules(args.authorized_channels)
        suspicious_terms = load_terms_file(
            args.suspicious_channel_terms,
            DEFAULT_SUSPICIOUS_CHANNEL_TERMS,
        )

        state, _ = scan_books(
            books,
            api_keys=api_keys,
            cache_path=args.cache,
            cache_ttl_days=args.cache_ttl_days,
            refresh=args.refresh,
            search_budget=search_budget,
            max_results=args.max_results,
            region=args.region,
            language=args.language,
            published_after=args.published_after,
            progress_every=args.progress_every,
        )

        accepted, accepted_by_book = build_report(
            args.output,
            books,
            state,
            profile=args.profile,
            threshold=args.threshold,
            long_duration_minutes=args.long_duration_minutes,
            authorized_ids=authorized_ids,
            authorized_names=authorized_names,
            suspicious_channel_terms=suspicious_terms,
            include_summaries=args.include_summaries,
            timezone_name=args.timezone,
        )

        print(f"Relatório salvo em: {args.output}")
        print(
            f"Títulos: {len(books)} | com suspeita: {len(accepted_by_book)} | "
            f"vídeos suspeitos: {len(accepted)} | pendentes: {len(state.pending_books)}"
        )
        print(
            f"Consultas novas à API: {state.api_search_calls} | "
            f"títulos do cache: {state.cached_books}"
        )
        print(
            f"Chaves: {state.api_keys_configured} configuradas | "
            f"{state.api_keys_used} usadas | "
            f"{state.api_keys_unavailable} indisponíveis | "
            f"trocas automáticas: {state.api_key_rotations}"
        )
        if state.pending_books:
            print(
                "Há títulos pendentes. Execute novamente quando houver cota disponível, "
                "configure outra chave elegível ou solicite ampliação de cota; o cache "
                "permite retomar sem repetir buscas.",
                file=sys.stderr,
            )
        return 0

    except MonitorError as exc:
        print(f"Erro: {exc}", file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        print("Interrompido pelo usuário.", file=sys.stderr)
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
