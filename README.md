# hidroweb_precip — chuva mensal/anual do Hidroweb para EI e fator R

Lê um `.txt` com códigos de estações pluviométricas, obtém as séries diárias da ANA, avalia a qualidade,
organiza a chuva total mensal em colunas (Jan–Dez) e calcula médias mensais e anual. Entrega uma planilha
por estação aprovada.

## Instalar e rodar

```
pip install pandas numpy xlsxwriter openpyxl requests

# o programa pergunta a janela de anos e os limites no terminal:
python hidroweb_precip.py estacoes.txt --fonte local --pasta baixados --saida saida

# ou tudo por opção (sem perguntas):
python hidroweb_precip.py estacoes.txt --fonte local --pasta baixados --saida saida \
    --ano-inicial 1991 --ano-final 2020 --min-anos 20 --max-falhas-pct 10

# baixando pela API da ANA (precisa do token; ver a seção sobre fontes):
export HIDROWEB_TOKEN='cole_o_tokenautenticacao'      # vale ~1 hora; não vai para arquivo nem log
# ou, para o programa obter e renovar o token sozinho (listas grandes):
read -rp "Identificador: " HIDROWEB_IDENTIFICADOR; read -rsp "Senha: " HIDROWEB_SENHA; echo
export HIDROWEB_IDENTIFICADOR HIDROWEB_SENHA
python hidroweb_precip.py estacoes.txt --fonte api --saida saida \
    --ano-inicial 1991 --ano-final 2020 --min-anos 20 --max-falhas-pct 10

python testar_sintetico.py        # valida tudo com dados inventados
```

`estacoes.txt`: um código por linha (ou separados por vírgula/espaço). `#` inicia comentário.
Códigos com 7 ou 8 dígitos funcionam; duplicatas e lixo são ignorados.

## De onde vêm os dados (leia antes de usar)

| Fonte | Como | Situação |
|---|---|---|
| `--fonte api` | API HidroWebService da ANA (`HidroSerieChuva/v1`, JSON). Uma requisição **por ano** (a API recusa intervalos acima de 366 dias). Token JWT (`tokenautenticacao`) em `HIDROWEB_TOKEN`; se faltar e houver terminal, é pedido sem eco. Guarda o JSON em `dados_brutos/` e o reaproveita (`--refazer` baixa de novo). | **Formato e limite de 366 dias confirmados com a API real** (estação 1646001, 09/11/2026). Ainda não rodado ao vivo: várias estações em sequência, anos sem dado (tratados como vazios), renovação do token. Se o token expirar no meio, a execução para com mensagem clara; sem `--ano-inicial` ele busca desde 1900, o que é lento. |
| `--fonte ws` | WebService público da ANA (`HidroSerieHistorica`, chuva = tipo 2), com 6 tentativas e espera crescente | **Não testado ao vivo** (o ambiente de desenvolvimento não alcança o servidor da ANA). A lógica foi validada contra um servidor falso que imita o formato documentado. A ANA indica que uma API nova substituirá este serviço, então ele pode mudar ou sair do ar. |
| `--fonte local` | Arquivos baixados do portal Hidroweb: `.zip` (inclusive zip dentro de zip), `.csv`, `.txt`, `.xlsx` | Formato assumido: uma linha por mês, colunas `EstacaoCodigo`, `NivelConsistencia`, `Data`, `Chuva01…Chuva31`, `Chuva01Status…`. Linhas `//`, separador `;`, vírgula decimal e latin-1 são aceitos. `.mdb` não é lido. |
| Token por identificador/senha | Com `HIDROWEB_IDENTIFICADOR` e `HIDROWEB_SENHA`, o programa obtém o token (rota `EstacoesTelemetricas/OAUth/v1`, cabeçalhos `Identificador` e `Senha`, token em `items.tokenautenticacao`) e o renova a menos de 5 min de expirar ou quando a API o recusa. Credenciais só por variável de ambiente; nunca em arquivo, log ou opção de linha de comando. | A **estrutura da resposta** foi confirmada com a API real. Os cabeçalhos e o endereço dessa rota vieram de memória: se a ANA recusar (mensagem "identificador/senha"), confira o endereço no manual e use `HIDROWEB_AUTH_URL`. |

Se o seu arquivo real divergir do formato assumido, o programa avisa qual arquivo não entendeu. Mande as primeiras linhas dele para ajustar.

O portal em si (cliques no site) não é automatizado de propósito: não há caminho estável e permitido para isso.

## O que o programa decide

- **Modelo base de datas × do zero:** o programa monta sozinho o calendário completo (todos os anos × 12 meses) e marca o que falta. Não precisa de modelo pronto. Meses que nem aparecem no arquivo viram falha "mês ausente na fonte", em vez de sumirem sem aviso.
- **Qual série usar em cada mês:** nível 2 (consistido, definitivo) é sempre preferido. Se o mês de nível 2 não tem leituras diárias mas traz o total mensal consolidado, esse total é usado (e contado na aba Resumo). O nível 1 (bruto) só entra quando o consistido não tem nem leituras nem total. `--so-consistido` desliga o bruto.
- **Mês válido:** todos os dias com leitura (`--max-falhas-dia N` tolera N dias sem leitura, mas o total então subestima).
- **Média mensal:** média do mês entre os anos em que ele é válido. **Média anual** = soma das 12 médias. Como conferência, a planilha traz também a média dos totais anuais (só anos com 12 meses).
- Status 3 (duvidoso) e 4 (acumulado) são contados e informados, mas não eliminam nada.

## Janela de anos e critérios eliminatórios (definidos por você)

Não há valores padrão: você informa a janela e os limites a cada execução, por opção ou, se omitir, o programa pergunta no terminal.

| O que | Opção | No terminal |
|---|---|---|
| Janela de anos: ano inicial | `--ano-inicial` | Enter em branco = sem limite |
| Janela de anos: ano final | `--ano-final` | Enter em branco = sem limite |
| Anos de histórico (1º ao último mês com dado) ≥ N | `--min-anos` | obrigatório |
| Meses com falha (%) ≤ N | `--max-falhas-pct` | obrigatório |

Além desses dois limites, a estação precisa ter ao menos 1 mês completo.

- **Janela de anos:** só os anos dentro dela entram em tudo (histórico, falhas, médias). O "histórico" e as "falhas" são medidos do primeiro ao último mês com dado **dentro da janela**.
- **Sem terminal** (agendador, script): se faltar `--min-anos` ou `--max-falhas-pct`, o programa recusa e diz o que falta. A janela omitida vira "sem limite".
- O critério "anos válidos em cada mês do calendário" foi **removido**. Atenção: sem ele, uma estação sem nenhum dado em algum mês do ano pode ser aprovada, e a média anual dela fica em branco (a soma das 12 médias exige os 12 meses).
- A janela e os limites usados ficam gravados no log e na aba `Resumo` de cada planilha. Estações reprovadas ficam no `resumo_geral` com o motivo.

## Meses sem dado

O programa **não estima** meses faltantes. Eles ficam em branco (amarelo) na aba `Mensal`, são listados na aba `Falhas` com o motivo (mês ausente na fonte, sem leituras no mês, N dias sem leitura), entram na contagem de "meses com falha" do critério eliminatório e ficam de fora das médias.

## Fournier, EI e fator R

A planilha sempre traz o coeficiente de Fournier mensal `p²/P` e o índice de Fournier modificado (soma dos 12). O EI só é calculado se você informar a equação da sua região:

```
--ei-a 67.355 --ei-b 0.85      # EI = a·(p²/P)^b ; R = soma dos 12 EI
```

Esses valores são só o formato do exemplo. Os coeficientes são regionais e vêm da literatura da sua área de estudo.

## Saídas (`--saida`)

- `estacoes/<código>.xlsx`: abas `Resumo` (critérios, médias, Fournier), `Mensal` (ano × Jan–Dez, total anual, médias) e `Falhas`.
- `resumo_geral.xlsx` / `.csv`: todas as estações, situação, motivo, médias e Fournier (ou R).
- `dados_diarios/`: série diária escolhida de cada estação. `dados_brutos/`: XML baixado (reaproveitado nas próximas execuções; `--refazer` baixa de novo). `log_execucao.txt`.

## Validação feita

`testar_sintetico.py` gera 6 estações inventadas (csv latin-1, xlsx, zip dentro de zip, zip inválido, XML, duplicatas de nível, meses ausentes/vazios/incompletos) e confere: lista de estações, aprovação/reprovação, 372 células mensais contra valores calculados à parte, médias, tolerância de dias, retentativa após erro 503, cache e R. Isso prova a lógica, **não** o comportamento com os arquivos reais da ANA.
