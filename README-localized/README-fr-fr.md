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
# Complément Outlook sur un exemple de code d’envoi

Découvrez comment vérifier la présence de mots restreints dans le corps d’un message électronique Outlook, ajouter un destinataire à la ligne Cc et vérifier que l’e-mail à envoyer contient un objet.

>**Remarque :** 

* La fonctionnalité d’envoi est actuellement prise en charge par Outlook sur le web dans Office 365 uniquement. 
* Pour en savoir plus sur la fonctionnalité envoi, voir [Fonctionnalité d’envoi pour les compléments Outlook](https://dev.office.com/docs/add-ins/outlook/outlook-on-send-addins).  
* Pour obtenir une procédure pas à pas de code, voir [Exemples de code](https://docs.microsoft.com/en-us/outlook/add-ins/outlook-on-send-addins#code-examples).

## Table des matières
* [Historique des modifications](#change-history)
* [Conditions préalables](#prerequisites)
* [Configuration et installation de l'exemple](#configure)
* [Charger les manifestes](#manifests)
* [Exécution du complément](#test-the-add-in)
* [Questions et commentaires](#questions-and-comments)
* [Ressources supplémentaires](#additional-resources)

## Historique des modifications

Avril 2017

* Version d’origine.

## Conditions préalables

* Un serveur web approuvé pour héberger les fichiers de l'exemple. Le serveur doit pouvoir accepter des demandes protégées par SSL (https) et disposer d’un certificat SSL valide.
* Un compte de courrier Office 365.
* Activer la fonctionnalité d’envoi : par défaut, la fonctionnalité envoi est désactivée. Les compléments Outlook sur le web qui utilisent la fonctionnalité d’envoi s’exécutent pour les utilisateurs auxquels une stratégie de boîte aux lettres Outlook sur le web est attribuée, dont la valeur **OnSendAddinsEnabled** est définie sur **true**. Les administrateurs peuvent activer la fonctionnalité d’envoi en exécutant les cmdlets Exchange Online PowerShell. Pour découvrir les applets de commande à exécuter, voir [Installation des compléments Outlook qui utilisent l’envoi](https://docs.microsoft.com/en-us/outlook/add-ins/outlook-on-send-addins#installing-outlook-add-ins-that-use-on-send)

## Configuration et installation de l'exemple

1. Téléchargez ou dérivez par le référentiel.
2. Ouvrir app.js. Dans la fonction`addCCOnSend`, remplacez `Contoso@contoso.onmicrosoft.com` par votre propre adresse de messagerie.
2. Déployer les fichiers du complément dans un répertoire sur votre serveur Web. Les fichiers à charger sont app.js et index.html.
3. Ouvrez les fichiers manifesteContoso Message Body Checker.xml et `Contoso Subject and CC Checker.xml` dans un éditeur de texte. Remplacez toutes les instances de`https://localhost:3000` par les URL de HTTPS du répertoire dans lequel vous avez téléchargé les fichiers au cours de l’étape précédente. Enregistrez vos modifications.

   >  Pour plus d'informations sur :
   * exécuter les compléments Outlook, voir [Exécution d’un complément Outlook dans un compte Office 365](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
   * manifestes, voir [manifestes de compléments Outlook](https://dev.office.com/docs/add-ins/outlook/manifests/manifests)

## Charger les manifestes

1. Accédez à [Outlook Web App](https://outlook.office365.com).
2. Cliquez sur **Paramètres** (l’engrenage rouage dans le coin supérieur droit de la page) pour ouvrir la page **Paramètres** (comme illustré dans la capture d’écran suivante).

  ![Page Paramètres](./readme-images/block-on-send-settings.png)

3. Dans la section**Vos paramètres de l’application** de la page **Paramètres**, sélectionnez **E-mail**.
4. Dans la page **Options**, sélectionnez **Général**, puis **Gérer les compléments** (comme illustré dans la capture d’écran suivante).

 ![Page Gérer les compléments](./readme-images/block-on-send-manage-addins.png)

5. Dans la page **Gérer les compléments**, cliquez sur l’icône « + », puis sélectionnez **Ajouter à partir d’un fichier**. Accédez au fichier manifeste ` Contoso Message Body Checker.xml` inclus dans le projet. Cliquez sur **Suivant**, puis cliquez sur **Installer**. Pour terminer, cliquez sur **OK**.
6. Répétez l’étape 5 pour installer le fichier manifeste `Contoso Subject and CC Checker.xml`.
7. Revenir à l’affichage courrier dans Outlook Web App.


## Exécution du complément

### Vérificateur de l’objet et de la ligne Cc

1. Composez un nouveau message électronique Outlook Web App. 
2. Laissez la ligne d’objet vide.
3. Ajoutez un destinataire dans la ligne**à**. 
4. Cliquez sur **Envoyer**. 

* Une copie carbone est ajoutée à la ligne CC. Dans cet exemple, il s’agit de `Contoso@contoso.onmicrosoft.com`.
* L’envoi du message électronique n’est pas autorisé et un message d’erreur s’affiche dans la barre d’informations pour avertir l’expéditeur d’ajouter un objet. (comme illustré dans la capture d’écran ci-dessous).  

 ![Barre d’informations sur l’objet et le vérificateur CC](./readme-images/block-on-send-subject-cc-inforbar.png) 

6. Ajoutez une ligne d’objet.
7. Un `[vérifié] :` est ajouté au début de la ligne d’objet et le courrier électronique est envoyé.

### Vérificateur du corps du message

1. Créer un nouveau message électronique Outlook Web App. 
2. Dans le corps du message, tapez `blockedword`, `blockedword1` ou `blockedword2`. (Il s’agit de la matrice de mots restreints dans le fichier app.js de la fonction `checkBodyOnlyOnSendCallBack`).
3. Ajoutez un destinataire dans la ligne**à**. 
5. Cliquez sur **Envoyer**.  

* L’envoi du courrier électronique est bloqué en raison de mots bloqués trouvés dans le corps du message.  
* Un message d’erreur s’affiche dans la barre d’informations pour avertir l’expéditeur que des mots bloqués ont été trouvés. (comme illustré dans la capture d’écran ci-dessous).  

 ![Barre d’informations du vérificateur du corps du message](./readme-images/block-on-send-body.png)

5. Pour envoyer le courrier électronique, supprimez les mots bloqués.

## Questions et commentaires

Nous serions ravis de connaître votre opinion sur cet exemple. Vous pouvez nous envoyer vos commentaires via la section *Problèmes* de ce référentiel.

Les questions générales sur le développement de Microsoft Office 365 doivent être publiées sur [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Si votre question concerne les API Office JavaScript, assurez-vous qu’elle est marquée avec les balises [office js] et [API].

## Ressources supplémentaires

* [Documentation de complément Office](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Centre de développement Office](http://dev.office.com/)
* Plus d’exemples de complément Office sur [OfficeDev sur Github](https://github.com/officedev)

## Copyright
Copyright (c) 2016 Microsoft Corporation. Tous droits réservés.



Ce projet a adopté le [code de conduite Open Source de Microsoft](https://opensource.microsoft.com/codeofconduct/). Pour en savoir plus, reportez-vous à la [FAQ relative au code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.
