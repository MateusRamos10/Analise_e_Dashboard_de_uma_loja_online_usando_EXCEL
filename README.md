[![author](https://img.shields.io/badge/author-mateusramos-red.svg)](https://www.linkedin.com/in/mateus-simoes-ramos/) ![contributions welcome](https://img.shields.io/badge/contributions-welcome-brightgreen.svg?style=flat)
# Análise de um Dataset de Ordens de Pedido de uma loja online usando Excel

<p align="center">
  <img alt="Dataset Inicial" width="60%" src="https://i.postimg.cc/K8XSJGvq/Dataset-inicial.png">
</p>

## Objetivo do Estudo
O objetivo deste projeto é realizar uma análise abrangente dos dados de pedidos de uma loja online, buscando insights que possam orientar decisões estratégicas individualmente para os clientes e apresentar as informações em um dashboard interativo.

**Para acessar o projeto completo, clique no link abaixo:**
 - [**Projeto no Excel(OneDrive)**](https://unipead-my.sharepoint.com/personal/mateus_ramos2_aluno_unip_br/_layouts/15/Doc.aspx?sourcedoc={7916d3b7-4548-4b83-b8ef-114df9611579}&action=embedview&wdAllowInteractivity=False&wdHideGridlines=True&wdHideHeaders=True&wdDownloadButton=True&wdInConfigurator=True&wdInConfigurator=True)

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

