# estrategico — Automação do Dashboard de Risco (Farol)

Automatiza a atualização da base do dashboard a partir da planilha de **KRIs**
recebida todo mês. Basta colocar a nova planilha na pasta e rodar o script.

## O que o script faz

`atualizar_dashboard.py` executa, em sequência:

1. **Localiza a planilha de KRI mais recente** dentro da pasta `KRIs/`.
   Reconhece nomes no padrão `01. Janeiro`, `02. Fevereiro`, ... (usa o número
   do mês como referência; em caso de empate, a data de modificação).
2. **Lê as 4 abas de risco** (identificadas pelo número inicial do nome da aba:
   `1-`, `2-`, `3 -`, `4-`). Em cada aba procura a coluna **"Nota Final"** —
   exceto na aba de *Desequilíbrio Atuarial*, onde a coluna é **"EPS (KRI)"** —
   e usa o valor da **última linha preenchida** (a linha-resumo "KRI"), que
   representa a nota do risco de 0% a 100%.
3. **Preenche a aba `Fato_Indicadores`** seguindo o padrão existente:
   `Data`, `ID_Risco`, `Codigo_Risco`, `Real_Atingimento`, `Meta` (90%),
   `Variacao_pp_vs_anterior` e `Status_Farol`.
4. **Calcula a variação** comparando o valor do mês atual com o do mês anterior
   do mesmo risco.
5. **Define o farol**:
   - 🟢 **Verde** — valor ≥ 90%
   - 🟡 **Amarelo** — valor > 70% e < 90%
   - 🔴 **Vermelho** — valor < 70%
6. **Atualiza a aba `Dim_Tempo`** inserindo o mês processado e marcando-o como
   o mês atual (`Eh_Mes_Atual = "Sim"`).

O script é **idempotente**: rodar de novo para o mesmo mês atualiza as linhas já
existentes em vez de duplicá-las.

## Como usar

```bash
pip install openpyxl

# Coloque a nova planilha de KRI na pasta KRIs/ (ex.: "03. Março.xlsx") e rode:
python atualizar_dashboard.py
```

Por padrão o script lê de `KRIs/` e atualiza `dados farol ficticios_8710.xlsx`,
criando antes uma cópia de backup com data/hora no nome.

### Parâmetros opcionais

```bash
python atualizar_dashboard.py \
    --pasta KRIs \
    --destino "dados farol ficticios_8710.xlsx" \
    --ano 2026 \
    --sem-backup
```

| Parâmetro     | Descrição                                              | Padrão                          |
|---------------|--------------------------------------------------------|---------------------------------|
| `--pasta`     | Pasta com as planilhas de KRI                          | `KRIs`                          |
| `--destino`   | Planilha base do dashboard                             | `dados farol ficticios_8710.xlsx` |
| `--ano`       | Ano de referência do mês processado                    | ano atual                       |
| `--sem-backup`| Não criar cópia de backup da planilha destino          | (cria backup)                   |

## Estrutura

```
KRIs/                              # planilhas mensais de KRI (entrada)
  01. Janeiro.xlsx
dados farol ficticios_8710.xlsx    # base do dashboard (saída, modelo estrela)
atualizar_dashboard.py             # script de automação
```

## Observações

- A leitura das notas usa os **valores calculados** da planilha de KRI (cache do
  Excel). Os arquivos exportados pelo Excel já contêm esse cache; abra e salve a
  planilha no Excel caso alguma célula de "Nota Final"/"EPS (KRI)" venha sem
  valor calculado.
- O farol e a variação são gravados como **valores** (não fórmulas), garantindo
  leitura direta pelo Power BI sem depender de recálculo.
