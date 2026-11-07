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

# baixando do WebService da ANA (ver a seção sobre fontes):
python hidroweb_precip.py estacoes.txt --fonte ws --saida saida --min-anos 20 --max-falhas-pct 10
python testar_sintetico.py        # valida tudo com dados inventados
```

`estacoes.txt`: um código por linha (ou separados por vírgula/espaço). `#` inicia comentário.
Códigos com 7 ou 8 dígitos funcionam; duplicatas e lixo são ignorados.

## De onde vêm os dados (leia antes de usar)

| Fonte | Como | Situação |
|---|---|---|
| `--fonte ws` | WebService público da ANA (`HidroSerieHistorica`, chuva = tipo 2), com 6 tentativas e espera crescente | **Não testado ao vivo** (o ambiente de desenvolvimento não alcança o servidor da ANA). A lógica foi validada contra um servidor falso que imita o formato documentado. A ANA indica que uma API nova substituirá este serviço, então ele pode mudar ou sair do ar. |
| `--fonte local` | Arquivos baixados do portal Hidroweb: `.zip` (inclusive zip dentro de zip), `.csv`, `.txt`, `.xlsx` | Formato assumido: uma linha por mês, colunas `EstacaoCodigo`, `NivelConsistencia`, `Data`, `Chuva01…Chuva31`, `Chuva01Status…`. Linhas `//`, separador `;`, vírgula decimal e latin-1 são aceitos. `.mdb` não é lido. |
| API oficial nova | Exige cadastro por e-mail (hidro@ana.gov.br) | Não implementada. Se você conseguir as credenciais, dá para acrescentar. |

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

## Correção de meses sem dado (`--preencher regional`, padrão)

Ponderação regional: `Px = média[(Mx/Mi)·Pi]`, com até 3 estações doadoras **da própria lista**, escolhidas por correlação (r ≥ 0,7 no mesmo mês, ≥ 10 anos em comum), exigindo ≥ 2 doadoras com valor naquele mês. Só preenche dentro da janela da estação e só em estações aprovadas. Precisa de ≥ 3 estações com dados na lista. `--preencher nenhum` desliga.

Valores preenchidos aparecem em azul, listados na aba `Falhas`, e **nunca entram nas médias "observadas"**. A coluna "com preenchimento" existe só para comparação. Doadoras distantes ou de clima diferente deveriam ser evitadas: use uma lista de estações da mesma região.

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

`testar_sintetico.py` gera 6 estações inventadas (csv latin-1, xlsx, zip dentro de zip, zip inválido, XML, duplicatas de nível, meses ausentes/vazios/incompletos) e confere: lista de estações, aprovação/reprovação, 372 células mensais contra valores calculados à parte, médias, tolerância de dias, preenchimento (erro mediano ≈ 6 % nos dados sintéticos), retentativa após erro 503, cache e R. Isso prova a lógica, **não** o comportamento com os arquivos reais da ANA.
