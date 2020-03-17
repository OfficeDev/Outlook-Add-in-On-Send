---
page_type: sample
products:
- office-outlook
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 3/20/2017 11:55:13 PM
---
# Exemplos de código ao enviar de suplementos do Outlook

Saiba como verificar a existência de palavras restritas no corpo da mensagem de email Outlook, adicionar um destinatário à linha Cc e verificar se há um assunto no email ao enviar.

>**Observação:** 

* O recurso ao enviar atualmente só tem suporte para o Outlook na Web no Office 365. 
* Para saber mais sobre o recurso ao enviar, confira [Recurso ao enviar para suplementos do Outlook](https://dev.office.com/docs/add-ins/outlook/outlook-on-send-addins).  
* Para obter uma explicação passo a passo do código, confira [Exemplos de códigos](https://docs.microsoft.com/en-us/outlook/add-ins/outlook-on-send-addins#code-examples).

## Sumário
* [Histórico de alterações](#change-history)
* [Pré-requisitos](#prerequisites)
* [Configurar e instalar o exemplo](#configure)
* [Carregar os manifestos](#manifests)
* [Executar o suplemento](#test-the-add-in)
* [Perguntas e comentários](#questions-and-comments)
* [Recursos adicionais](#additional-resources)

## Histórico de alterações

Abril de 2017

* Versão inicial.

## Pré-requisitos

* Um servidor Web confiável para hospedar os arquivos de exemplo. O servidor deve ser capaz de aceitar solicitações protegidas por SSL (https) e ter um certificado SSL válido.
* Uma conta de email do Office 365.
* Habilitar o recurso ao enviar, por padrão esta funcionalidade está desabilitada. Os suplementos para o Outlook na Web que usam o recurso ao enviar serão executados para os usuários atribuídos a uma política de caixa de correio do Outlook na Web que tem o sinalizador **OnSendAddinsEnabled** definido como **true**. Os administradores podem habilitar ao enviar executando cmdlets do PowerShell do Exchange Online. Para saber quais cmdlets executar, confira [Instalar suplementos do Outlook que usam ao enviar](https://docs.microsoft.com/en-us/outlook/add-ins/outlook-on-send-addins#installing-outlook-add-ins-that-use-on-send)

## Configurar e instalar o exemplo

1. Baixar ou bifurcar o repositório.
2. Abra o app.js. Na função `addCCOnSend`, altere `Contoso@contoso.onmicrosoft.com` para o seu endereço de email.
2. Implante os arquivos do suplemento em um diretório do seu servidor Web. Os arquivos a serem carregados são app.js e index.html.
3. Abra os arquivos de manifesto `Contoso Message Body Checker.xml` e `Contoso Subject and CC Checker.xml` em um editor de texto. Substitua todas as instâncias de `https://localhost:3000` com a URL HTTPS do diretório em que você carregou os arquivos na etapa anterior. Salve suas alterações.

   >  Para saber mais sobre:
   * Executar suplementos do Outlook, confira [Executar um suplemento do Outlook em uma conta do Office 365](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
   * Manifestos, confira [Manifestos de suplementos do Outlook](https://dev.office.com/docs/add-ins/outlook/manifests/manifests)

## Carregar os manifestos

1. Vá para o [Outlook Web App](https://outlook.office365.com).
2. Clique em **Configurações** (a engrenagem no canto superior direito da página) para abrir a página **Configurações** (como mostrado na captura de tela a seguir).

  ![A página Configurações](./readme-images/block-on-send-settings.png)

3. Na seção **Configurações do aplicativo** da página **Configurações**, escolha **Email**.
4. Na página **Opções**, selecione **Geral** e **Gerenciar suplementos** (como mostrado na captura de tela a seguir).

 ![A página Gerenciar suplementos](./readme-images/block-on-send-manage-addins.png)

5. Na página **Gerenciar suplementos**, clique no ícone '+' e selecione **Adicionar do arquivo**. Navegue até o arquivo de manifesto `Contoso Message Body Checker.xml` incluído no projeto. Clique em **Avançar** e em **Instalar**. Por fim, clique em **OK**.
6. Repita a etapa 5 para instalar o arquivo de manifesto `Contoso Subject and CC Checker.xml`.
7. Retorne ao modo de exibição de email no Outlook Web App.


## Executar o suplemento

### Assunto e verificador de CC

1. Redigir uma nova mensagem de email do Outlook Web App. 
2. Deixe a linha de assunto em branco.
3. Adicione um destinatário na linha **Para**. 
4. Clique em **Enviar**. 

* Uma cópia carbono é adicionada à linha CC. Neste exemplo, é `Contoso@contoso.onmicrosoft.com`.
* O envio do email é bloqueado e uma mensagem de erro é exibida na barra de informações para notificar o remetente para adicionar um assunto (como mostrado na captura de tela a seguir).  

 ![A barra de informações do verificador de assunto e CC](./readme-images/block-on-send-subject-cc-inforbar.png) 

6. Adicione uma linha de assunto.
7. Um `[Marcado]:` é adicionado à frente da linha de assunto e o email é enviado.

### Verificador de corpo da mensagem

1. Crie uma nova mensagem de email do Outlook Web App. 
2. No corpo da mensagem, digite `blockedword`, `blockedword1` ou `blockedword2`. (Esse é o conjunto de palavras restritas no arquivo app.js da função `checkBodyOnlyOnSendCallBack`).
3. Adicione um destinatário na linha **Para**. 
5. Clique em **Enviar**.  

* O envio do email é bloqueado devido a palavras bloqueadas encontradas no corpo da mensagem.  
* Uma mensagem de erro é exibida na barra de informações para notificar o remetente que foram encontradas palavras bloqueadas (como mostrado na captura de tela a seguir).  

 ![A barra de informações do verificador do corpo da mensagem](./readme-images/block-on-send-body.png)

5. Para enviar o email, remova as palavras bloqueadas.

## Perguntas e comentários

Gostaríamos de saber sua opinião sobre este exemplo. Você pode enviar comentários na seção *Problemas* deste repositório.

As perguntas sobre o desenvolvimento do Microsoft Office 365 em geral devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Se sua pergunta estiver relacionada às APIs JavaScript para Office, não deixe de marcá-la com as tags [office-js] e [API].

## Recursos adicionais

* [Documentação dos suplementos do Office](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Centro de Desenvolvimento do Office](http://dev.office.com/)
* Confira outros exemplos de Suplemento do Office em [OfficeDev no Github](https://github.com/officedev)

## Direitos autorais
Copyright (c) 2016 Microsoft Corporation. Todos os direitos reservados.



Este projeto adotou o [Código de Conduta do Código Aberto da Microsoft](https://opensource.microsoft.com/codeofconduct/). Para saber mais, confira [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou contate [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.
