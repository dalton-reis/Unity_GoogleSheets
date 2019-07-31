# Google Sheets Unity Plugin

Este repositório disponibiliza uma biblioteca desenvolvida com a **Unity 2018.4.4f1** para adicionar, atualizar, remover e listar itens à partir de uma planilha do Google Drive. Para isso foi utilizado a [Google Sheets API V4](https://developers.google.com/sheets/api/guides/concepts), que disponibiliza funcionalidades para trabalhar com essas planilhas. 

Caso queira apenas baixar o package para importar no seu projeto, basta [clicar aqui](https://github.com/tecedufurb/google-sheets-unity/blob/master/google-sheets-unity.unitypackage). O package já contém uma cena demonstrando um exemplo de uso do plugin.

O diagrama abaixo apresenta as classes e funcionalidades disponibilizadas pelo plugin:

![Diagrama da implementação](https://github.com/tecedufurb/google-sheets-unity/blob/master/google-sheets-unity.jpg)


## Pré-requisitos
Para que esse plugin funcione corretamente é necessário primeiro criar um projeto e gerar as credenciais necessárias no [Google Developer Console](https://console.developers.google.com/). Para isso recomenda-se seguir o passo a passo apresentado [aqui](https://flaviocopes.com/google-api-authentication/). Não é necessário seguir o passo a passo inteiro, apenas a parte referênte à geração das credenciais através do Google Developer Console.

Uma vez gerado o arquivo com as credenciais, deve-se substituir o conteúdo do arquivo [./Assets/GoogleSheetsAPI/credentials.json](https://github.com/tecedufurb/google-sheets-unity/blob/master/Assets/GoogleSheetsAPI/credentials.json) pelo conteúdo do arquivo de credenciais gerado.
