[![author](https://img.shields.io/badge/author-mateusramos-red.svg)](https://www.linkedin.com/in/mateus-simoes-ramos/) ![contributions welcome](https://img.shields.io/badge/contributions-welcome-brightgreen.svg?style=flat)
# Análise de um Dataset de Ordens de Pedido de uma loja online usando Excel

<p align="center">
  <img alt="Dataset Inicial" width="60%" src="https://i.postimg.cc/K8XSJGvq/Dataset-inicial.png">
</p>

## Objetivo do Estudo
O objetivo deste projeto é realizar uma análise abrangente dos dados de pedidos de uma loja online, buscando insights que possam orientar decisões estratégicas individualmente para os clientes e apresentar as informações em um dashboard interativo.

**Para acessar o projeto completo, clique no link abaixo:**
 - [Projeto no Excel(OneDrive)](https://unipead-my.sharepoint.com/personal/mateus_ramos2_aluno_unip_br/_layouts/15/Doc.aspx?sourcedoc={7916d3b7-4548-4b83-b8ef-114df9611579}&action=embedview&wdAllowInteractivity=False&wdHideGridlines=True&wdHideHeaders=True&wdDownloadButton=True&wdInConfigurator=True&wdInConfigurator=True)

<!Projeto em Google Sheets também,
Botão pra pular para os resultados
Alterar imagem inicial e colocar imagem do dashboard
!>

## Fonte dos Dados
Os dados foram obtidos da plataforma [Kaggle](https://www.kaggle.com/). disponibilizado no link abaixo:

* [Dataset Sales Excel](https://community.tableau.com/s/question/0D54T00000CWeX8SAL/sample-superstore-sales-excelxls)


## Tecnologias Utilizadas
<p align="left">  
	<a href="https://www.microsoft.com/pt-br/microsoft-365/excel" target="_blank" rel="noreferrer"> <img src="https://i.postimg.cc/ncLxTcKJ/icons8-microsoft-excel-2019-48.png" alt="Excel" width="40" height="40"/> 
	</a>
	<a href="https://workspace.google.com/intl/pt-BR/products/sheets/" target="_blank" rel="noreferrer"> <img src="https://i.postimg.cc/V6x2BmFj/icons8-google-sheets-48.png" alt="Google Sheets" width="40" height="40"/> 
	</a> 
</p>

---

## Mãos a obra...
### 1. Importando e trabalhando com arquivos diferentes
Logo no início eu precisei mesclar dois arquivos, o Dataset Completo.xlsm estava apenas com o código do cliente e para esse projeto é importante sabermos o nome dos clientes para criar nossas análises e Dashboards.
<br><br>
<p align="center">
  <img alt="Dataset Inicial" width="75%" src="https://i.postimg.cc/K8XSJGvq/Dataset-inicial.png">
</p>

### 1.1 Tratando caracteres especiais
Ao verificar o arquivo clientes.pdf , que tem 793 clientes com 20 páginas, verifiquei que havia alguns caracteres especiais, onde constatei nesse arquivo os seguintes erros:

Onde é "$" deverá ser trocado por "C";<br>
Onde é "%" deverá ser trocado por "A";<br>
Onde é "-" deverá ser trocado por " ".
<br><br>
<p align="center">
  <img alt="PDF Nomes Clientes" width="75%" src="https://i.postimg.cc/3rkztKtM/3-PDF-Client-Names.gif">
</p>

### 1.2 Transformando Dados
Dentro do Power Query importei o arquivo, porém cada página do PDF gerou uma tabela diferente. Para resolver usei a função Acrescentar Consultas como Novas que coloca todas as tabelas em apenas uma página, para que todos os registros fiquem juntos.
<br><br>
<p align="center">
  <img alt="Dataset Inicial" width="75%" src="https://i.postimg.cc/66FDCfv4/4-meclando-tabelas-pdf.gif">
</p>

Ainda dentro do Power Query, usei a função localizar/substituir para fazer a limpeza dessa coluna.
<br><br>
<p align="center">
  <img alt="Dataset Inicial" width="75%" src="https://i.postimg.cc/QttRJ6m1/5-substituindo-caracteres.gif">
</p>

Alguns registros ficaram com a letra minúscula quando substituímos os caracteres especiais, então usei uma função para deixar a primeira letra maiúscula e também coloquei a primeira linha como cabeçalho. Finalmente finalizamos as alterações usando o Power Query.

>*Transformar/Formato/Colocar Cada Palavra Em Maiúsculo*<br>
>*Página Inicial/Usa a Primeira Linha Como Cabeçalho*
<br><br>
<p align="center">
  <img alt="Dataset Inicial" width="80%" src="https://i.postimg.cc/90yshLWR/6-letra-maiuscula-e-cabecalho.png">
  <img alt="Dataset Inicial" width="25%" src="https://i.postimg.cc/BvWTgtCg/8-Resultado-Power-Query.jpg">
</p>
<br>

### 2. Automação de preenchimento de dados
Importei os registros para a planilha e temos uma nova aba com os registros dos nomes dos cliente. Logo após foi necessário usar a função PROCV e SEERRO. Usei os dados inseridos do pdf como matriz onde o Costumer ID se comporta como chave e o Customer Name se comporta como valor de uma matriz, e depois populei todos os 9994 registros da planilha Dataset usando um script de macro. A função "SEERRO" serve para tratamento de erro caso tenha um código errado, e retorna um valor vazio como resultado, isso será útil posteriormente. 
Abaixo está o código usado no macro, onde optei por usar scripts do Office ao invés de VBA, sabendo que a grosso modo, o VBA são para soluções desktop e os scripts para soluções Web.

```typescript
function main(workbook: ExcelScript.Workbook) {
	// Obtém a aba Client_names
	let abaClient_names = workbook.getWorksheet("Client_names");

	// Obtém a aba Orders
	let abaOrders = workbook.getWorksheet("Orders");

	// Define a coluna pela qual irá iterar (no caso, coluna F)
	let coluna = abaOrders.getRange("F:F");

	// Obtém o alcance da coluna
	let usedRange = coluna.getUsedRange();

	// Obtém o índice da última linha na coluna
	let lastRow = usedRange.getLastRow().getRowIndex();

	// Loop através de cada célula na Coluna F
	for (var i = 1; i <= lastRow + 1; i++) {

	// Define a fórmula na célula D correspondente
	abaOrders.getRange("G" + i).setFormulaLocal('=SEERRO(PROCV(F' + i + '; ' + abaClient_names.getRange("A1:B793").getAddress() + '; 2; FALSO); "")');
	}
}
```

Então vamos ver a "magia" acontecendo. Após aproximadamente 8 minutos, os 9994 registros estavam preenchidos, quanto tempo você demoraria para preencher essa quantidade de registros?
<p align="center">
  <img alt="Dataset Inicial" width="75%" src="https://i.postimg.cc/C156tzM5/8-execucao-automacao-scor.gif">
</p>


