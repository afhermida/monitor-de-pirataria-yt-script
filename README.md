# Monitor de possíveis audiolivros não autorizados no YouTube

Este pacote contém um script de triagem antipirataria. Ele lê uma planilha `.xlsx` com **títulos** e **autores**, consulta a **YouTube Data API v3**, classifica resultados por indícios de publicação integral e gera um relatório no mesmo formato geral da planilha de referência:

- `Resumo`
- `Ranking por Views`
- `Vídeos Suspeitos`
- `Catálogo Completo`

O script **não baixa áudio nem vídeo**. Os resultados são indícios para revisão humana; o score não comprova infração e não deve, sozinho, fundamentar uma notificação ou pedido de remoção.

## 1. Requisitos

- Python 3.10 ou superior
- De uma a dez chaves da YouTube Data API v3
- Dependências do arquivo `requirements_youtube_monitor.txt`

Instalação sugerida:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements_youtube_monitor.txt
```

No Windows PowerShell:

```powershell
.venv\Scripts\Activate.ps1
pip install -r requirements_youtube_monitor.txt
```

## 2. Configurar uma ou até dez chaves da API

O script aceita no máximo **10 chaves únicas** e faz rotação automática quando a chave ativa retorna erro de cota, limite, chave inválida, restrição incompatível ou API desativada.

A troca ocorre dentro da chamada que falhou. Por exemplo, se `search.list` terminar e a cota acabar durante `videos.list`, somente `videos.list` é repetida com a chave seguinte.

Os valores das chaves não são gravados no relatório, no cache nem nas mensagens de progresso. O log mostra apenas identificadores como `chave 2/5`.

### Opção A — variável de ambiente com várias chaves

Linux/macOS:

```bash
export YOUTUBE_API_KEYS="chave-1,chave-2,chave-3"
```

Windows PowerShell:

```powershell
$env:YOUTUBE_API_KEYS="chave-1,chave-2,chave-3"
```

A variável antiga, com uma única chave, continua funcionando:

```bash
export YOUTUBE_API_KEY="sua-chave"
```

### Opção B — repetir `--api-key`

```bash
python monitor_youtube_audiobooks.py catalogo.xlsx \
  --api-key "chave-1" \
  --api-key "chave-2" \
  --api-key "chave-3" \
  -o relatorio.xlsx
```

Também é possível passar uma lista em uma única opção:

```bash
--api-key "chave-1,chave-2,chave-3"
```

A linha de comando pode aparecer no histórico do shell ou na lista de processos. Para uso recorrente, prefira variável de ambiente ou arquivo protegido.

### Opção C — arquivo de chaves

Crie um TXT, por exemplo `chaves_youtube.txt`:

```text
# Uma chave por linha. Não coloque este arquivo no Git.
chave-1
chave-2
chave-3
```

Execute:

```bash
python monitor_youtube_audiobooks.py catalogo.xlsx \
  --api-keys-file chaves_youtube.txt \
  -o relatorio.xlsx
```

Em Linux/macOS, limite a leitura do arquivo:

```bash
chmod 600 chaves_youtube.txt
```

Quando `--api-key` ou `--api-keys-file` é usado, essas fontes explícitas têm prioridade sobre `YOUTUBE_API_KEYS` e `YOUTUBE_API_KEY`.

### Observação importante sobre cotas

A quota é associada ao projeto da chave. Chaves diferentes do mesmo projeto normalmente compartilham a mesma quota e, portanto, podem ficar indisponíveis em sequência. Use somente chaves e projetos legitimamente provisionados para a organização, respeite as políticas da API e solicite ampliação oficial de quota quando o volume recorrente exigir.

Documentação oficial:

- Autenticação: https://developers.google.com/youtube/data/guides/authentication
- Busca: https://developers.google.com/youtube/data/reference/rest/v3/search/list
- Metadados de vídeos: https://developers.google.com/youtube/data/reference/rest/v3/videos/list
- Cotas: https://developers.google.com/youtube/data/determine_quota_cost

## 3. Formato da planilha de entrada

O script procura automaticamente colunas com nomes como:

| Título | Autor |
|---|---|
| Meditações | Marco Aurélio |
| O poder do hábito | Charles Duhigg |

Também reconhece variantes como `Título do livro`, `Title`, `Autora` e `Author`. Se houver várias abas compatíveis, prioriza uma aba cujo nome contenha `Catálogo`, `Livros` ou `Títulos`; em seguida, escolhe a maior.

Para indicar explicitamente a aba ou os cabeçalhos:

```bash
python monitor_youtube_audiobooks.py catalogo.xlsx \
  --sheet "Catálogo" \
  --title-column "Nome da obra" \
  --author-column "Responsável"
```

## 4. Execução recomendada

Com chaves na variável `YOUTUBE_API_KEYS`:

```bash
python monitor_youtube_audiobooks.py catalogo.xlsx \
  -o relatorio_pirataria_youtube.xlsx \
  --profile equilibrado \
  --authorized-channels canais_autorizados_exemplo.txt
```

O perfil padrão, `equilibrado`, mantém a pontuação da planilha de referência, mas acrescenta duas proteções:

1. exige uma correspondência bibliográfica razoável entre o livro e o vídeo;
2. penaliza resumos, resenhas e vídeos curtos sem indicação clara de audiobook.

Isso evita falsos positivos como um vídeo de **Meditações**, de Marco Aurélio, aparecer como resultado de **Manual do estoicismo**, de Epicteto, apenas por conter “livro completo”.

## 5. Perfil compatível com a planilha de referência

Para reproduzir mais de perto a lógica observada na planilha fornecida:

```bash
python monitor_youtube_audiobooks.py catalogo.xlsx \
  -o relatorio_compativel.xlsx \
  --profile planilha
```

Pesos calibrados a partir do arquivo de referência:

| Indício | Peso |
|---|---:|
| Título contém “audiobook completo”, “livro completo” ou equivalente | 0,35 |
| Nome do canal contém termos típicos de repositório de audiolivros | 0,25 |
| Duração de pelo menos 20 minutos | 0,20 |
| Canal fora da lista de canais autorizados | 0,10 |
| Título do livro aparece no título do vídeo | 0,10 |
| Descrição contém “completo”, “full”, “gratuito” ou equivalente | 0,10 |

O score é limitado a `1,00`. O limiar padrão é `0,50` e pode ser alterado com `--threshold`.

## 6. Canais autorizados

A API pública não oferece um campo adequado para usar o selo visual de “verificado” como critério de autorização. Por isso, na coluna `Razões`, **“canal não verificado” significa “canal fora da allowlist do titular”**.

Crie um TXT com IDs ou nomes exatos de canais autorizados:

```text
# IDs são preferíveis porque nomes podem mudar.
id:UCxxxxxxxxxxxxxxxxxxxxxx

# Também é possível cadastrar o nome exibido.
nome:Companhia das Letras
nome:Canal oficial do selo
```

Use:

```bash
--authorized-channels canais_autorizados.txt
```

Resultados de canais cadastrados são descartados antes da pontuação.

## 7. Rotação, orçamento local, cache e retomada

O parâmetro `--search-budget` é um teto **local** para a quantidade de novos títulos pesquisados na execução. Ele não altera a quota do Google.

Quando o parâmetro é omitido, o script calcula automaticamente:

```text
95 × número de chaves configuradas
```

Exemplos:

| Chaves configuradas | Orçamento local padrão |
|---:|---:|
| 1 | 95 títulos |
| 3 | 285 títulos |
| 10 | 950 títulos |

A chave não é trocada simplesmente ao chegar a 95 títulos. O pool continua com a chave ativa até a API informar cota/limite ou outro erro que torne a chave indisponível. O orçamento acima apenas impede que uma execução automática ultrapasse o volume total planejado.

Para definir um teto total específico:

```bash
--search-budget 250
```

Para remover somente o teto local e deixar a execução seguir até terminar o catálogo ou todas as chaves ficarem indisponíveis:

```bash
--search-budget 0
```

O cache `.youtube_monitor_cache.json` guarda as buscas já concluídas. Ao executar novamente, o script continua pelos títulos pendentes sem repetir entradas válidas.

Opções úteis:

- `--cache caminho.json`: escolhe o arquivo de checkpoint.
- `--cache-ttl-days 30`: validade do cache; `-1` significa que não expira.
- `--refresh`: tenta atualizar entradas existentes; se todas as chaves acabarem, o relatório ainda aproveita o cache anterior.
- `--published-after 2026-01-01`: limita a busca a vídeos publicados a partir dessa data.

Quando nem a API nem o cache têm dados para um título, ele aparece como `pendente`, e não como “sem suspeita”.

A aba `Resumo` e o terminal informam:

- chaves configuradas;
- chaves efetivamente utilizadas;
- chaves marcadas como indisponíveis na execução;
- quantidade de trocas automáticas.

## 8. Termos adicionais para canais suspeitos

Os termos padrão incluem variantes como `audiolivros`, `audio livros`, `audiobooks` e `livros narrados`. Para acrescentar termos próprios, crie um TXT, um por linha:

```text
acervo falado
biblioteca em áudio
narração integral
```

Use:

```bash
--suspicious-channel-terms termos_canais.txt
```

## 9. Outros exemplos

Buscar até 30 resultados por título:

```bash
python monitor_youtube_audiobooks.py catalogo.xlsx \
  --max-results 30 \
  -o relatorio.xlsx
```

Elevar o limiar para reduzir o número de casos enviados à revisão:

```bash
python monitor_youtube_audiobooks.py catalogo.xlsx \
  --threshold 0.65 \
  -o relatorio.xlsx
```

Não penalizar vídeos marcados como resumo ou resenha:

```bash
python monitor_youtube_audiobooks.py catalogo.xlsx \
  --include-summaries \
  -o relatorio.xlsx
```

Executar os testes internos da calibração, leitura das chaves e rotação:

```bash
python monitor_youtube_audiobooks.py --self-test
```

Consultar todas as opções:

```bash
python monitor_youtube_audiobooks.py --help
```

## 10. Interpretação do relatório

- **Resumo:** cobertura da execução, pendências, quantidade de títulos sinalizados, uso do cache/API e estatísticas das chaves.
- **Ranking por Views:** ordena títulos pelo vídeo suspeito com mais visualizações. O score exibido é o maior score encontrado para o livro; a URL e as razões pertencem ao vídeo mais visto, reproduzindo o comportamento observado na referência.
- **Vídeos Suspeitos:** uma linha por candidato aceito, com canal, views, duração, score, razões, data e hyperlink.
- **Catálogo Completo:** preserva a ordem da entrada e marca cada obra como `SIM`, `não` ou `pendente`.

## 11. Limitações importantes

- A busca do YouTube é probabilística e muda ao longo do tempo.
- Metadados podem ser enganosos ou incompletos.
- Obras em domínio público podem envolver traduções, narrações, edições ou gravações ainda protegidas.
- Um canal com “audiolivros” no nome não é necessariamente infrator.
- Um vídeo classificado como resumo pode, em certos casos, conter trechos extensos; por isso, a revisão humana continua necessária.
- A rotação não cria quota adicional, não contorna quotas associadas ao mesmo projeto e não substitui uma solicitação oficial de ampliação.
