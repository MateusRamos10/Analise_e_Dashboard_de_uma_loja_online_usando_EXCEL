[![author](https://img.shields.io/badge/author-mateusramos-red.svg)](https://www.linkedin.com/in/mateus-simoes-ramos/) ![contributions welcome](https://img.shields.io/badge/contributions-welcome-brightgreen.svg?style=flat)
# Análise de um Dataset de Ordens de Pedido de uma loja online usando Excel
<p align="center">
  <img alt="Dashboard" width="85%" src="https://github.com/MateusRamos10/Excel_Clients/assets/43836795/9b526b74-a74c-4644-a6e6-ff5d83740101">
</p>

## Objetivo do Estudo
O objetivo deste projeto é realizar uma análise abrangente dos dados de pedidos de uma loja online, buscando insights que possam orientar decisões estratégicas e apresentar informações em um dashboard interativo.

**Para acessar o arquivo no OneDrive, clique no link abaixo:**
 - [Projeto no Excel(OneDrive)](https://1drv.ms/x/s!AiKF1yOmNO2VmFuJl5Yru03_DieR?e=kZ1J3m)

<!-- COMENTÁRIO Fazer projeto em Google Sheets também
icone sheets
<a href="https://workspace.google.com/intl/pt-BR/products/sheets/" target="_blank" rel="noreferrer"> <img src="https://i.postimg.cc/V6x2BmFj/icons8-google-sheets-48.png" alt="Google Sheets" width="40" height="40"/> 
</a> 
-->

## Fonte dos Dados
Os dados foram obtidos da plataforma [Kaggle](https://www.kaggle.com/). disponibilizado no link abaixo:

* [Dataset Sales Excel](https://community.tableau.com/s/question/0D54T00000CWeX8SAL/sample-superstore-sales-excelxls)


## Tecnologias Utilizadas
<p align="left">  
	<a href="https://www.microsoft.com/pt-br/microsoft-365/excel" target="_blank" rel="noreferrer"> <img src="https://i.postimg.cc/ncLxTcKJ/icons8-microsoft-excel-2019-48.png" alt="Excel" width="40" height="40"/> 
	</a>
</p>

---

<br>
<br>

## Mãos a obra...

> Caso prefira, clique abaixo para pular o desenvolvimento deste trabalho direto para os Resultados.
> <br>
> **[Resultados](#resultados)**

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

>*Transformar/Formato/Colocar Cada Palavra Em Maiúsculo*
>
>*Página Inicial/Usa a Primeira Linha Como Cabeçalho*

<br><br>
<p align="center">
  <img alt="Dataset Inicial" width="80%" src="https://i.postimg.cc/90yshLWR/6-letra-maiuscula-e-cabecalho.png">
  <img alt="Dataset Inicial" width="25%" src="https://i.postimg.cc/BvWTgtCg/8-Resultado-Power-Query.jpg">
</p>
<br>
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
<br>

Então vamos ver a "magia" acontecendo. Após aproximadamente 8 minutos, os 9994 registros estavam preenchidos, quanto tempo você demoraria para preencher essa quantidade de registros?

<p align="center">
  <img alt="execucao automacao scorsese" width="90%" src="https://i.imgur.com/d8uzqUo.gif">
</p>

Lembra da função SEERRO? Agora vamos verificar se tem algum valor nulo...
Temos aqui 13 valores nulos que representam 0,13% de todos os registros e nessa ocasião optei por excluir essas linhas. Também foi analisado se tem algum nome repetido na aba Client_names.

> *Dados/Filtrar/Desmarcar Selecionar Tudo/Marcar Vazias*
>
> *Página Inicial/Formatação Condicional/Regras de Realce das Células/Valores Duplicados*

<br>
<br>

### 3. Analisando tabelas e criando tabelas dinâmicas
Aqui é aquele momento quase sem registro, que a gente pensa e pensa e pensa... Mas quero destacar duas tabelas dinâmicas e duas **dicas** que só a experiência traz.
<p align="center">
  <img alt="Dataset Inicial" width="75%" src="https://github.com/MateusRamos10/Excel_Clients/assets/43836795/db40dd5f-c361-4d0a-99c6-b531b8a520a8">
</p>

### 3.1 Tabelas Dinâmicas
*\#1* Além de criações de tabelas dinâmicas *linkando* corretamente as colunas, criei a coluna que calcula a positividade.
Nesse contexto, chamei de positividade todo cliente que teve mais de uma compra no período filtrado.<br>
De uma maneira bem compreensível, se um cliente comprou em uma data e voltou a comprar em uma outra data, contamos como 1 para positividade, se ele comprou uma única vez e não voltou, a positividade recebe 0, e pra finalizar, usei a função SE para fazer a somatória de todos os registros com 1 subtraindo do total de clientes e temos a nossa positividade.
<br>

*\#2* Como eu queria trabalhar com o mapa do país, entendi que esse gráfico ele não pode ser criado a partir de uma tabela dinâmica, o próprio excel recomenda fazer uma cópia desses dados em uma coluna para criar esse gráfico, nas minhas pesquisas vi uma técnica para transformar uma tabela dinâmica em uma simples tabela, porém não obtive sucesso. 
Então para contornar esse problema e não ficar duplicando dados, eu crie uma coluna onde apenas coloquei o nome do país *United States* e plotei um gráfico dessa coluna e depois adicionei os dados da tabela dinâmica.

<p align="center">
  <img alt="Dataset Inicial" width="75%" src="https://github.com/MateusRamos10/Excel_Clients/assets/43836795/680b3b4b-2b59-4495-8aef-95c56877cfe1">
  <img alt="Dataset Inicial" width="60%" src="https://github.com/MateusRamos10/Excel_Clients/assets/43836795/520a0d16-4a35-4a53-97ce-71f301134f50">
</p>

### 3.2 Conselho de Amigo (Dicas)
Em vários momentos eu precisei simplesmente começar do zero por que ocorria algum erro que eu não conseguia solucionar, claro que me chateei no início, mas acredito que internalizei melhor tudo o que estava fazendo (extrair o melhor da situação).

<br>

*\#Tip1* Como eu queria fazer um único filtro para todas as tabelas e também já sabia antecipadamente a quantidade de tabelas dinânicas que eu queria fazer, entendi que era melhor eu criar o botão (que na verdade é uma segmentação de dados da tabela) e as 9 tabelas sem registro nenhum, do que criar *uma a uma* e programar individualmente. Em um caso que ao criar 8 tabelas e uma única céluar não ser selecionada ou uma simples coluna a menos, mesmo que eu não a use, a segmentação não vai funcionar e pode trazer uma dorzinha de cabeça, então é melhor pensar antes quais tabelas dinâmicas irá precisar.

<br>

*\#Tip2* Proteja sua planilha da maneira certa, primeiro marque a opção de movimentação da sua segmentação, depois desbloqueie a segmentação, e depois proteja a planilha. Dessa maneira, a única parte clicável do dashboard são os botões (a segmetação não pode estar agrupada).
> *Botão Direito na Segmentação/Posição e Layout/marque a opção Desabilitar redimensionamento e movimentação*
> 
> *Botão Direito na Segmentação/Tamanho e Propriedades/Propriedades/Desmarca a opção Bloqueado*
> 
> *Página Inicial/Formatar/Proteger Planilha*

<br>

*\#Tip3* Outro padrão que resolvi adotar é de criar uma aba para simplesmente visualizar as cores do meu dashboard e manter a segurança para não acabar estragando alguma programação do dashboard original.
Talvez seja a parte que as pessoas mais gostem, e eu demorei muito tempo pensando na harmonia e na mensagem que queria transmitir, mas foi bom testar as cores em uma planilha a parte e sim... Teoria das cores é fundamental e se você ter a curiosidade de ver quais foram as cores do meu primeiro Dashboard pode me enviar uma mensagem, não irei postar aqui rs.

<p align="center">
  <img alt="Dataset Inicial" width="75%" src="https://github.com/MateusRamos10/Excel_Clients/assets/43836795/2a7196d1-991f-4706-a33b-7c13aa1b2642">
</p>

## Resultados <a id="resultados"></a>
Projeto iniciado com uma planilha e um arquivo em pdf.
- Transformado em um único arquivo onde importei os dados do pdf;
- Criado uma aba com análises usando tabelas dinâmicas;
- Outra aba com a decisão final das cores;
- E por fim um dashboard interativo.

Deixo abaixo algumas conclusões extraídas do dashboard e caso tenha dúvida, sugestão ou alguma correção, deixo meu linkedin abaixo:
	
[Vamos no conectar!](https://www.linkedin.com/in/mateus-simoes-ramos/)

---

<p align="center">
  <img alt="Dataset Inicial" width="60%" src="https://github.com/MateusRamos10/Excel_Clients/assets/43836795/8e9b60bc-a078-43c1-8536-beb03800065c">
</p>

Categoria | Situação
----------|---------------
Bookcases | negativo em 2014,2015,2017
Tables 	  | negativo em 2014,2015,2016,2017
Supplies  | negativo em 2016,2017
Machines  | negativo em 2017

Chamo atenção para essas 4 categorias, em especial Bookcases e Tables que no quadro geral estão com o lucro negativo. Tables no último ano triplicou o prejuízo somando -$17,275,48.

---

<p align="center">
  <img alt="Dataset Inicial" width="30%" src="https://github.com/MateusRamos10/Excel_Clients/assets/43836795/14ccc81f-64b7-4d17-a9b0-ac5095f37ad0">
</p>

Por ser uma loja ampla e ter 17 subcategorias, temos produtos super baratos como envelopes de papel até mobília com mesas e cadeiras.
O ticket médio nesses 4 anos está em $60,67 e os 9 primeiros produtos estão com o ticket médio acima de mil dólares.

Produto | Categoria | Subcategoria
-------------------|----------|--------
Cisco TelePresence | Machines | Technology
Canon imageCLASS   | Copiers  | Technology
High Speed         | Suplies  | Office Suplies
HP Designjet       | Machines | Technology
Cubify CuveX       | Machines | Technology
Ibico EPK-21       | Binders  | Office Suplies
3D Systems         | Machines | Technology 
Zebra ZM400        | Machines | Technology 
Okidata MB760      | Machines | Technology 
Lexmark MX611dhe   | Machines | Technology

Sete de dez subcategorias pertencem a ***Machine***, buscar parcerias com fornecedores dessa área trará um retorno positivo.

---
<p align="center">
  <img alt="Dataset Inicial" width="60%" src="https://github.com/MateusRamos10/Excel_Clients/assets/43836795/ae546286-b154-421c-a8a7-869836dffe37">
</p>

Califórnia concentra 7663 das vendas ao longo dos 4 anos, em segundo vem New York com 4219, porém em terceiro é o Texas com 3724 pedidos. 
Seria interessante concentrar pontos de distribuição entre Califórnia e Texas que são mais próximos de acordo com os produtos que mais são vendidos nesses dois lugares, isso se for uma loja virtual independente.
Um serviço oferecido aos vendedores da Amazon é o FBA (**Fulfillment by Amazon**), é um serviço que engloba o armazenamento, entrega do produto e o pós venda. A Amazon possui centros de distribuição em todos os estados do Estados Unidos (e também tem centros no Brasil).
Como já fui um vendeor, essa logística é feita automaticamente, no momento em que escolhe enviar seus produtos para a Amazon, a própria decide para qual centro irá seus produtos, analisando de onde irá ser enviado os produtos e a demanda que eles tem na região.
<br>
<br>
<br>

Esse trabalho foi desenvolvimento a fim de mostrar algumas de minhas habilidades em Excel e espero que tenha gostado! Deixo abaixo meu contato:

[![linkedin](https://img.shields.io/badge/linkedin-0A66C2?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/mateus-simoes-ramos/)&nbsp;
[![Email](https://img.shields.io/badge/Gmail-D14836?style=for-the-badge&logo=gmail&logoColor=white)](mailto:mateusramos.oficial@gmail.com)&nbsp;
[![Perfil DIO](https://img.shields.io/badge/-Meu%20Perfil%20na%20DIO-000?style=for-the-badge)](https://web.dio.me/users/mateusramos_oficial?tab=skills)
*mailto: mateusramos.oficial@gmail.com*
<!-- Meu site -->
