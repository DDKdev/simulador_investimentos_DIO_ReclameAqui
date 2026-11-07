# 📊 Simulador de Investimentos em Fundos Imobiliários (FIIs)

Projeto desenvolvido como parte do desafio da **DIO (Digital Innovation One)** com mentoria do expert **Felipe Aguiar**. O objetivo deste projeto é construir uma ferramenta prática e interativa em Excel/Google Sheets para simulação de patrimônio acumulado e renda passiva gerada por Fundos Imobiliários ao longo do tempo.

---

## 📌 Perguntas de Negócio Respondidas pela Ferramenta

A planilha foi projetada no formato de painel/aplicativo para responder diretamente às 5 perguntas essenciais de planejamento financeiro:

| Pergunta de Negócio | Onde Encontrar na Planilha | Função / Lógica Utilizada |
| :--- | :--- | :--- |
| **1. Quanto investir por mês?** | Painel de Entrada de Dados (Célula `B4` / `aporte_mensal`) | Input direto do usuário ou calculado como 30% do salário |
| **2. Por quantos anos?** | Painel de Entrada de Dados (Célula `B5` / `prazo_anos`) | Input direto do usuário (convertido em meses multiplicando por 12) |
| **3. Qual a taxa de rendimento mensal?** | Painel de Entrada de Dados (Célula `B6` / `taxa_mensal`) | Percentual estimado de *Dividend Yield* mensal médio |
| **4. Quanto de patrimônio vai acumular?** | Painel de Resultados (Célula `E4`) | Calculado via função `=VF()` |
| **5. Quanto vai receber de dividendos por mês?** | Painel de Resultados (Célula `E5`) | Multiplicação do Patrimônio Acumulado pela `taxa_mensal` |

---

## 🧠 Fórmulas e Funções Principais

### 1. Função `VF` (Valor Futuro)
A função `VF` é o coração do simulador, responsável por calcular a projeção de juros compostos considerando aportes recorrentes:

```excel
=VF(taxa_mensal; prazo_anos * 12; -aporte_mensal; 0; 0)
```

* **`taxa` (`taxa_mensal`):** A taxa de juros/rendimento aplicada mensalmente.
* **`nper` (`prazo_anos * 12`):** O número total de períodos (meses). Multiplica-se os anos por 12.
* **`pgto` (`-aporte_mensal`):** O valor do depósito mensal. Usa-se o sinal negativo (`-`) para que o Excel entenda como uma saída de caixa e retorne o resultado final em valor positivo.
* **`vp` (`0`):** Valor presente/inicial (definido como 0 no cenário padrão sem patrimônio prévio).
* **`tipo` (`0`):** Indica que os aportes ocorrem no final de cada período (padrão).

### 2. Busca Dinâmica de Alocação (`PROCV` com Chave Composta)
Para dividir o aporte entre os 6 tipos de FIIs de acordo com o perfil de investidor selecionado (ex: Conservador, Moderado, Arrojado), utiliza-se o `PROCV` combinado a uma **chave composta**:

```excel
=PROCV(perfil_selecionado & "_" & tipo_fundo; tabela_apoio_alocacao; 3; FALSO)
```

* **Chave Composta (`perfil_selecionado & "_" & tipo_fundo`):** Concatena o perfil do investidor com o tipo de fundo (exemplo: `"Conservador_Tijolo"`). Isso permite buscar em uma tabela bidimensional de suporte uma combinação única de dois critérios.
* **Tabela de Apoio:** Contém a lista de combinações e o respectivo percentual de alocação correspondente (garantindo que cada perfil atinja o somatório de 100%).

---

## 🏷️ Intervalos Nomeados Criados

Para tornar as fórmulas legíveis e evitar erros com endereços de células fixos, foram definidos os seguintes intervalos nomeados na pasta de trabalho:

* `salario_base`: Valor base informado para cálculo de sugestão de aporte.
* `aporte_mensal`: Valor mensal investido.
* `prazo_anos`: Quantidade de anos da simulação.
* `taxa_mensal`: Taxa de Dividend Yield esperada ao mês.
* `perfil_investidor`: Célula com validação de dados em lista contendo o perfil selecionado.
* `tb_alocacao`: Intervalo da aba de suporte onde constam a chave composta, o perfil, o tipo de fundo e o percentual.

---

## 📊 Perfis de Investimento e Alocação

A divisão dos aportes por tipo de FII segue a seguinte distribuição percentual (somando 100% por perfil):

| Tipo de Fundo | Conservador | Moderado | Arrojado |
| :--- | :---: | :---: | :---: |
| **FIIs de Papel (CRI)** | 40% | 25% | 15% |
| **FIIs de Tijolo (Logística)** | 30% | 30% | 20% |
| **FIIs de Tijolo (Shoppings/Varejo)** | 15% | 20% | 20% |
| **FoFs (Fundos de Fundos)** | 15% | 15% | 15% |
| **FIIs de Desenvolvimento** | 0% | 5% | 15% |
| **FIAGRO / Outros** | 0% | 5% | 15% |
| **Total** | **100%** | **100%** | **100%** |

---

## 🛠️ Destaques e Melhorias em Relação à Versão Base

* **Proteção Visual de Células:** Células que possuem fórmulas possuem fundo cinza claro para sinalizar que não devem ser editadas manualmente.
* **Layout Limpo:** Linhas de grade, cabeçalhos de linha/coluna e barra de fórmulas foram ocultados na visão final para garantir a aparência de aplicativo.
* **Cenários Multi-Prazos:** Tabela rápida comparativa apresentando o resultado acumulado para prazos de 2, 5, 10, 20 e 30 anos.

---

## 🖼️ Capturas de Tela / Evidências de Funcionalidade

> *Adicione aqui os prints da sua planilha funcionando com dois perfis diferentes.*

1. **Simulação Perfil Conservador:**
   ![Perfil Conservador](docs/print_conservador.png)

2. **Simulação Perfil Arrojado:**
   ![Perfil Arrojado](docs/print_arrojado.png)

---

## 🚀 Como Executar o Projeto

1. Baixe o arquivo `.xlsx` deste repositório.
2. Abra no **Microsoft Excel** ou **Google Sheets**.
3. Altere os valores de **Aporte**, **Prazo (anos)**, **Taxa Mensal** ou o **Perfil de Investidor** no painel principal para ver todas as simulações e gráficos atualizarem automaticamente.