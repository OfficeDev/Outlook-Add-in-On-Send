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
# Outlook 加载项 Onsend 代码示例

了解如何在发送邮件时实现以下操作：检查 Outlook 电子邮件正文中是否有受限制的字词、在“抄送”行中添加收件人，以及检查电子邮件中是否包含主题。

>**注意：** 

* 目前仅 Office 365 中的 Outlook 网页版支持 Onsend 功能。 
* 若要了解 Onsend 功能，请参阅 [Outlook 加载项的 Onsend 功能](https://dev.office.com/docs/add-ins/outlook/outlook-on-send-addins)。  
* 如需获取代码演练，请参阅[代码示例](https://docs.microsoft.com/en-us/outlook/add-ins/outlook-on-send-addins#code-examples)。

## 目录
* [修订记录](#change-history)
* [先决条件](#prerequisites)
* [配置和安装示例](#configure)
* [上传清单](#manifests)
* [运行加载项](#test-the-add-in)
* [问题和意见](#questions-and-comments)
* [其他资源](#additional-resources)

## 修订记录

2017 年 4 月

* 首版。

## 先决条件

* 用于托管示例文件的可信 Web 服务器。该服务器必须能够接受受 SSL 保护的请求 (https)，并且具备有效的 SSL 证书。
* Office 365 电子邮件帐户。
* 启用 Onsend 功能 - 默认情况下，Onsend 功能处于禁用状态。对于分配了将 **OnSendAddinsEnabled** 标志设置为 **true** 的 Outlook 网页版邮箱策略的用户，系统将会为其运行使用 Onsend 功能的 Outlook 网页版加载项。管理员可以通过运行 Exchange Online PowerShell cmdlet 启用 Onsend。若要了解需要运行哪些 cmdlet，请参阅[安装使用 Onsend 的 Outlook 加载项](https://docs.microsoft.com/en-us/outlook/add-ins/outlook-on-send-addins#installing-outlook-add-ins-that-use-on-send)

## 配置和安装示例

1. 下载存储库或为其创建分支。
2. 打开 app.js。在 `addCCOnSend` 函数中，将 `Contoso@contoso.onmicrosoft.com` 改为自己的电子邮件地址。
2. 将加载项文件部署到 Web 服务器上的某个目录。要上传的文件是 app.js 和 index.html。
3. 在文本编辑器中打开 `Contoso Message Body Checker.xml` 和 `Contoso Subject and CC Checker.xml` 清单文件。将 `https://localhost:3000` 的所有实例都替换为上一步中的文件上传目录的 HTTPS URL。保存所做的更改。

   >  更多信息：
   * 有关运行 Outllook 加载项的更多信息，请参阅[在 Office 365 帐户中运行 Outlook 加载项](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
   * 有关清单的更多信息，请参阅 [Outlook 加载项清单](https://dev.office.com/docs/add-ins/outlook/manifests/manifests)

## 上传清单

1. 登录到 [Outlook Web App](https://outlook.office365.com)。
2. 单击“**设置**”（页面右上角的齿轮）以打开“**设置**”页面（如以下屏幕截图所示）。

  ![设置页面](./readme-images/block-on-send-settings.png)

3. 在“**设置**”页面的“**你的应用设置**”部分中，选择“**邮件**”。
4. 在“**选项**”页面中，选择“**常规**”，然后选择“**管理加载项**”（如下方屏幕截图所示）。

 ![管理加载项页面](./readme-images/block-on-send-manage-addins.png)

5. 在“**管理加载项**”页面上，单击“+”图标，选择“**从文件添加**”。浏览到项目中包含的 `Contoso Message Body Checker.xml` 清单文件。单击“**下一步**”，然后单击“**安装**”。最后，单击“**确定**”。
6. 重复步骤 5，以安装 `Contoso Subject and CC Checker.xml` 清单文件。
7. 在 Outlook Web App 中返回到“邮件”视图。


## 运行加载项

### 主题和抄送检查器

1. 撰写新的 Outlook Web App 电子邮件。 
2. 将主题行留空。
3. 在“**收件人**”行中添加一个收件人。 
4. 单击“**发送**”。 

* 系统会在“抄送”行中添加一个抄送收件人。在本示例中，抄送收件人为 `Contoso@contoso.onmicrosoft.com`。
* 系统会阻止发送该电子邮件，并在信息栏上显示一条错误消息，通知发件人添加主题（如下方屏幕截图所示）。  

 ![主题和抄送检查器的信息栏](./readme-images/block-on-send-subject-cc-inforbar.png) 

6. 添加主题行。
7. 系统会在主题行的前面添加一个 `[已检查]:` 标志并发送电子邮件。

### 邮件正文检查器

1. 创建新的 Outlook Web App 电子邮件。 
2. 在邮件正文中，键入 `blockedword`、`blockedword1` 或 `blockedword2`。（这些是 app.js 文件中 `checkBodyOnlyOnSendCallBack` 函数的一系列受限制字词）。
3. 在“**收件人**”行中添加一个收件人。 
5. 单击“**发送**”。  

* 系统会阻止发送该电子邮件，因为邮件正文中存在受阻字词。  
* 系统会在信息栏上显示一条错误消息，通知发件人存在受阻字词（如下方屏幕截图所示）。  

 ![邮件正文检查器的信息栏](./readme-images/block-on-send-body.png)

5. 若要发送电子邮件，请删除这些受阻字词。

## 问题和意见

我们乐意倾听你对此示例的反馈。可以在此存储库中的*“问题”*部分向我们发送反馈。

与 Microsoft Office 365 开发相关的一般问题应发布到 [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API)。如果你的问题是关于 Office JavaScript API，请务必为问题添加 [office-js] 和 [API].标记。

## 其他资源

* [Office 外接程序文档](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office 开发人员中心](http://dev.office.com/)
* 有关更多 Office 外接程序示例，请访问 [Github 上的 OfficeDev](https://github.com/officedev)。

## 版权信息
版权所有 (c) 2016 Microsoft Corporation。保留所有权利。



此项目已采用 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅[行为准则 FAQ](https://opensource.microsoft.com/codeofconduct/faq/)。如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。
