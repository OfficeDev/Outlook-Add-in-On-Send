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
# Outlook アドインの送信時コードのサンプル

Outlook メールのメッセージ本文で制限された単語をチェックし、CC 行に受信者を追加し、送信時にメールの件名が指定されていることを確認する方法を取り上げます。

>**注:** 

* 送信時機能は、現在 Office 365 の Outlook on the web でのみサポートされています。 
* 送信時機能の詳細については、「[Outlook アドインの送信時機能](https://dev.office.com/docs/add-ins/outlook/outlook-on-send-addins)」を参照してください。  
* コードのチュートリアルについては、「[コードの例](https://docs.microsoft.com/en-us/outlook/add-ins/outlook-on-send-addins#code-examples)」を参照してください。

## 目次
* [変更履歴](#change-history)
* [前提条件](#prerequisites)
* [サンプルの構成とインストール](#configure)
* [マニフェストをアップロードする](#manifests)
* [アドインを実行する](#test-the-add-in)
* [質問とコメント](#questions-and-comments)
* [その他のリソース](#additional-resources)

## 変更履歴

2017 年 4 月

* 初期バージョン。

## 前提条件

* サンプル ファイルをホストする信頼済み Web サーバー。サーバーは、SSL で保護された要求 (https) を受け入れることが可能で、有効な SSL 証明書を所有している必要があります。
* Office 365 メール アカウント。
* 送信時機能を有効にします。既定では、送信時機能は無効になっています。送信時機能を使用する Outlook on the web のアドインは、**OnSendAddinsEnabled** フラグが **true** に設定された Outlook on the web メールボックス ポリシーが割り当てられているユーザーに対して実行されます。管理者は、Exchange Online PowerShell コマンドレットを実行して、送信時機能を有効にできます。実行するコマンドレットについては、「[送信時機能を使用する Outlook アドインのインストール](https://docs.microsoft.com/en-us/outlook/add-ins/outlook-on-send-addins#installing-outlook-add-ins-that-use-on-send)」を参照してください。

## サンプルの構成とインストール

1. レポジトリをダウンロードまたはフォークします。
2. app.js を開きます。`addCCOnSend` 関数で、`Contoso@contoso.onmicrosoft.com` を自分のメール アドレスに変更します。
2. Web サーバー上のディレクトリにアドイン ファイルを展開します。アップロードするファイルは、app.js および index.html です。
3. テキスト エディターで `Contoso Message Body Checker.xml` と `Contoso Subject および CC Checker.xml` のマニフェスト ファイルを開きます。`https://localhost:3000` のすべてのインスタンスを、前の手順でファイルをアップロードしたディレクトリの HTTPS URL で置き換えます。変更内容を保存します。

   >  詳細については、以下を参照してください。
   * Outlook アドインの実行に関しては、「[Office 365 アカウントでの Outlook アドインの実行](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)」を参照してください。
   * マニフェストについては、「[Outlook アドインのマニフェスト](https://dev.office.com/docs/add-ins/outlook/manifests/manifests)」を参照してください。

## マニフェストをアップロードする

1. [Outlook Web App](https://outlook.office365.com) にログオンします。
2. [**設定**] (ページの右上隅にある歯車アイコン) をクリックして、[**設定**] ページ (次のスクリーンショットを参照) を開きます 。

  ![[設定] ページ](./readme-images/block-on-send-settings.png)

3. [**設定**] ページの [**アプリの設定**] セクションで、[**メール**] を選択します。
4. [**オプション**] ページで、[**全般**]、[**アドインの管理**] (次のスクリーンショットを参照) の順に選択します。

 ![[アドインの管理] ページ](./readme-images/block-on-send-manage-addins.png)

5. [**アドインの管理**] ページで、[+] アイコンをクリックし、[**ファイルから追加**] を選択します。プロジェクトに含まれている `Contoso Message Body Checker.xml` マニフェスト ファイルを参照します。[**次へ**]、[**インストール**] の順にクリックします。最後に、[**OK**] をクリックします。
6. 手順 5 を繰り返して、`Contoso Subject および CC Checker .xml` マニフェスト ファイルをインストールします。
7. Outlook Web App の [メール] ビューに戻ります。


## アドインを実行する

### 件名および CC のチェッカー

1. 新しい Outlook Web App のメール メッセージを作成します。 
2. [件名] は空白のままにします。
3. [**宛先**] に受信者を追加します。 
4. [**送信**] をクリックします。 

* カーボン コピーは CC に追加されます。このサンプルでは、`Contoso@contoso.onmicrosoft.com` です。
* メールの送信がブロックされ、情報バーに件名を追加するように送信者に通知するエラー メッセージが表示されます。 (次のスクリーンショットを参照)  

 ![件名および CC のチェッカーの情報バー](./readme-images/block-on-send-subject-cc-inforbar.png) 

6. 件名を追加します。
7. `[チェック済]:` 件名の先頭に追加され、メールが送信されます。

### メッセージ本文のチェッカー

1. 新しい Outlook Web App のメール メッセージを作成します。 
2. メッセージの本文に、`blockedword`、`blockedword1` または `blockedword2` を入力します。(これらは、`checkBodyOnlyOnSendCallBack` 関数の app.js ファイル内の制限された単語の配列です)。
3. [**宛先**] に受信者を追加します。 
5. [**送信**] をクリックします。  

* メッセージ本文でブロック対象の単語が見つかったため、メールの送信がブロックされています。  
* エラー メッセージが情報バーに表示され、検出されたブロック対象の単語を送信者に通知します (次のスクリーンショットを参照)。  

 ![メッセージ本文のチェッカーの情報バー](./readme-images/block-on-send-body.png)

5. メールを送信するには、ブロック対象の単語を削除します。

## 質問とコメント

このサンプルに関するフィードバックをお寄せください。このリポジトリの「*問題*」セクションでフィードバックを送信できます。

Microsoft Office 365 開発全般の質問につきましては、「[スタック オーバーフロー](http://stackoverflow.com/questions/tagged/office-js+API)」に投稿してください。Office JavaScript API に関する質問の場合は、必ず質問に [office-js] と [API] のタグを付けてください。

## その他の技術情報

* [Office アドインのドキュメント](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office デベロッパー センター](http://dev.office.com/)
* [Github の OfficeDev](https://github.com/officedev) にあるその他の Office アドイン サンプル

## 著作権
Copyright (c) 2016 Microsoft Corporation.All rights reserved.



このプロジェクトでは、[Microsoft オープン ソース倫理規定](https://opensource.microsoft.com/codeofconduct/)が採用されています。詳細については、「[倫理規定の FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。
