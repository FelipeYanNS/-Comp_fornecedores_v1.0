
---

## 📊 Funcionalidade Geral

- Selecionar uma empresa a partir do nome dos arquivos (o critério é até o sinal de menos - ).
- Ler e consolidar mútiplos arquivos de lançamentos contábeis da extenção `.csv`.
- Tratar e limpar dados (valores, documentos, contas, datas).
- Filtrar lançamentos da conta `2.1.2.01`(Código contábel da conta de Fornecedores do Passivo).
- Calcular saldos com base no tipo de ação (Crédito ou Débito).
- Agrupar dados por fornecedor e nota fiscal.
- Traz o saldo da conta para conferência (Botão de "Conferência").
- Gerar automaticamente um arquivo Excel por empresa (Botão "Baixar Arquivo").
- Consegue tratar os dados com a opção de selecionar o intervalo de tempo
- O código consegue considerar o saldo inicial apenas do primeiro ano

---

## 🛠️ Requisitos

- O código foi feito e pensado para ser utilizado no [Google Colab](https://colab.research.google.com/) em conjunto com o [Google Drive](https://drive.google.com/).
- Os dados devem ser extraido dos lançamentos contábeis no sistema [ERP UAU - Globaltec](https://www.globaltec.com.br/erp-uau/).
  
