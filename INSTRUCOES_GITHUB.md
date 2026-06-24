# 📋 Passo a passo — Atualizar o dashboard no GitHub

Você precisa substituir **2 arquivos** no seu repositório (`Nay5MMO/Coco-CIA-2`):

## 1️⃣ Substituir `dashboard_vendas.py`

1. Abra o repositório no GitHub: https://github.com/Nay5MMO/Coco-CIA-2
2. Clique no arquivo **`dashboard_vendas.py`**
3. Clique no ícone do **lápis** (✏️) no canto superior direito (editar)
4. **Selecione tudo** que está lá (Ctrl+A) e **apague**
5. Abra o novo `dashboard_vendas.py` (o que eu te entreguei agora) num editor de texto, **copie tudo** (Ctrl+A, Ctrl+C)
6. **Cole** na caixa de edição do GitHub (Ctrl+V)
7. Role até o final da página, escreva uma mensagem como `"atualiza dashboard para v3"`
8. Clique no botão verde **"Commit changes"**

## 2️⃣ Substituir `requirements.txt`

1. Volte pra página principal do repositório
2. Clique no arquivo **`requirements.txt`**
3. Clique no ícone do **lápis** (✏️)
4. **Apague tudo** e cole o conteúdo novo:
   ```
   streamlit
   pandas
   plotly
   openpyxl
   reportlab
   matplotlib
   ```
5. Role pra baixo, mensagem: `"adiciona bibliotecas para PDF"`
6. Clique em **"Commit changes"**

## 3️⃣ Aguardar o Streamlit Cloud

- O Streamlit vai detectar a mudança e **redeployar automaticamente** (leva 2-5 minutos)
- Pode acompanhar o status em https://share.streamlit.io
- Quando terminar, recarregue a página do app — você verá as 9 abas novas:

  1. 🎯 **Diagnóstico** — resumo numa página só
  2. 📈 Visão Geral
  3. 📅 YoY
  4. 🥥 Matéria-Prima & R$/kg
  5. 💰 **Preço Médio** — distorção entre clientes
  6. ⚖️ **Comparação** — lado a lado
  7. 🔍 Estratégicas (ABC + Sazonalidade + Geo)
  8. ⚠️ Riscos
  9. 🗂️ Dados

## ✅ O que está incluído nesta versão

| Pedido da reunião | Onde está |
|---|---|
| Diagnóstico geral em uma página | Aba "🎯 Diagnóstico" |
| Comparar 2+ clientes ou estados | Aba "⚖️ Comparação" |
| Preço médio por cliente (detectar distorção tipo Sendas vs BBA) | Aba "💰 Preço Médio" |
| R$/kg por linha (com alerta de contra-senso Food vs Profissional) | Aba "🥥 Matéria-Prima & R$/kg" |
| Cliente como "Código — Nome" | Todas as tabelas |
| Produto como "Código — Descrição" | Todas as tabelas |
| Curva ABC | Aba "🔍 Estratégicas" + destaque no Diagnóstico |
| Exportar relatório completo | Botões Excel + PDF na barra lateral |

---

⚠️ **Importante:** quando o dono te passar o **cadastro de produto** (com peso real, validade, preço médio e saldo de estoque), me avise. Vou integrar para deixar os cálculos de R$/kg ainda mais precisos. Por enquanto, o peso está sendo extraído da descrição do produto (funciona para 32 dos 36 SKUs — só falta os 4 itens sem peso na descrição: casca, kit, óleo virgem genérico, torta).
