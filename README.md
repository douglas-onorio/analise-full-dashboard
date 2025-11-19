# 📦 Análise Full - Dashboard de Estoque Mercado Livre

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://estoque-full.streamlit.app/)

## 🚀 Sobre o Projeto
Esta aplicação web moderniza um processo que antes dependia de macros complexas em Excel (VBA). O sistema processa relatórios de estoque do **Mercado Livre (Full)**, cruza com custos internos e gera um dashboard interativo para tomada de decisão de reposição e análise de saúde do estoque.

O objetivo principal é automatizar a inteligência de estoque para múltiplas empresas simultaneamente, eliminando erros manuais e gargalos de processamento do Excel.

## ✨ Funcionalidades Principais

* **Processamento de Dados:** Ingestão de planilhas complexas (.xlsx) e limpeza de dados utilizando **Pandas**.
* **Lógica de Negócio Complexa:**
    * Replicada fielmente das regras originais de negócio (filtros de status, cálculo de dias de estoque, alertas de custo).
    * Algoritmo de sugestão de ação (Ex: "Repor imediatamente", "Campanha de giro", "Risco de descarte").
* **Multi-Empresa:** Capacidade de carregar e processar dados de várias contas (Ex: VALE RACE, VANPARTS) na mesma sessão, com consolidação final.
* **Simulação de Reposição (DBM):** Módulo que calcula a necessidade de compra baseada na média de vendas diária e fatores de segurança.
* **Visualização:** Dashboard interativo com KPIs, tags coloridas para alertas críticos e tabelas ordenáveis.
* **Exportação:** Gera um novo arquivo Excel consolidado e formatado com apenas um clique.

## 🛠 Tecnologias Utilizadas

* **Python 3.9+**
* **Streamlit:** Para interface frontend e interatividade.
* **Pandas & NumPy:** Para manipulação de dados de alta performance.
* **XlsxWriter:** Para exportação de relatórios Excel avançados.

## ⚙️ Como Rodar Localmente

1.  **Clone o repositório:**
    ```bash
    git clone [https://github.com/SEU_USUARIO/NOME_DO_REPO.git](https://github.com/SEU_USUARIO/NOME_DO_REPO.git)
    ```

2.  **Instale as dependências:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Execute a aplicação:**
    ```bash
    streamlit run app.py
    ```

## 🧠 O Desafio: VBA vs Python
Este projeto resolveu problemas de performance e usabilidade das planilhas antigas:
* **Antes (VBA):** Lento com grandes volumes de dados, travava o Excel, difícil de visualizar em celulares.
* **Agora (Python Web):** Processamento em segundos, acessível via navegador em qualquer lugar, interface limpa e amigável.

---
**Desenvolvido por Douglas Onorio**
