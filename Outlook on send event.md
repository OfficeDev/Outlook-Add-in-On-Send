# On send feature for Outlook add-ins

You can use the Outlook add-in events to handle, check, or block user actions when something of interest occurs. Events provide ways to:

- Raise an event and in response, handle the raised event appropriately
- Control user actions
- Handle changes
- Signal user actions such as button clicks
- Verify user data input
- Validate content in a message    

The on send feature for Outlook add-ins currently supports the **ItemSend** event type. The on send feature provides a way to block email users from certain actions and allows an add-in to set certain items on send. For example, you can use the on send feature to:

- Prevent a user from sending sensitive information or leaving the subject line blank.  
- Add a specific recipient to the CC line.

>  **Note:** <!-- We make this point later so removing from here. Better to have one concept covered in a Note. You can't publish add-ins that use the on send feature to the Office Store; they must be installed by an administrator.-->The on send feature is currently supported for Outlook Web App in Office 365 only. <!-- Removing because we don't provide any more information in that section. For more information, see **Guidelines and restrictions** later in this article. -->

<!-- Let's add something like this (below). Also, "limitations" seems like a friendlier word than "retrictions"; "restrictions" implies that the user is restricted, not the feature.-->

For information about additional limitations related to the on send feature, see [Limitations]() later in this article.
## How does the on send feature work?

You can use the on send feature to build an Outlook add-in that integrates the **ItemSend** synchronous event. This event detects that the user is pressing the **Send** button and blocks the email from being sent if the message validation fails.

Validation is done on the client side<!-- LG: This only applies to OWA, right? Do we need to qualify - i.e. might be different when additional platforms are supported?, in the browser -->. on trigger of the send event. For example, at message send event, an Outlook add-in that uses the on send feature can:

- Read and validate the email message contents
- Verify that the message includes a subject line
- Set a predetermined recipient 

If validation fails, the email is blocked from being sent, and an error message is displayed to let users know why they can't send the email.  

The following screenshot shows an information bar that notifies the sender to add a subject.

![Screenshot showing an error message prompting the user to enter a missing subject line](./readme-images/block-on-send-subject-cc-inforbar.png) 

The following screenshot shows an information bar that notifies the sender that blocked words were found.

![Screenshot showing an error message telling the user that blocked words were found](./readme-images/block-on-send-body.png)

<!-- Including a big long section about guidelines and restrictions near the beginning of the article might discourage folks from using the feature. I suggest that we include the code examples first, as that is what devs generally looking for, and have a section about limitations closer to the end of the article. 
                                                                                                                                                                                        
## Guidelines and restrictions
-->

<!-- This doesn't provide any more information; it just repeats what was stated earlier. I suggest we remove this as it is duplication. Esp because we're trying to make the article a little shorter. ;)

##  Only supported on Outlook Web App in Office 365 
Currently, the on send event is only supported on Outlook Web App in Office 365.  Support for other SKUs are coming soon.  
-->

## Limitations

The on send feature currently has the following limitations:

- You can't publish Outlook add-ins that use the on send feature to the Office Store. Add-ins that use the on send event will fail Office Store validation.   
- Only one **ItemSend** event is supported per add-in. If you have two or more **ItemSend** events in a manifest, the manifest will fail validation. 
- Multiple roundtrips to the web server that hosts the add-in can affect the performance of the add-in. Consider the effects on performance when you create add-ins that require multiple email message-based operations.

##  Multiple on send add-ins behavior

If there are more than one on send add-in installed, the add-ins will execute in order of installation.  After a first add-in allows sending, a second add-in could change something that would make the first one deny send (but the first one would not be run again as long as all add-ins have allowed sending).

For example, let's say *Add-in1* and *Add-in2* both use the on send event. *Add-in1* is installed first, and *Add-in2* installed after *Add-in1*.  Let's say *Add-in1* checks that the word *Fabrikam* appears in the message as a condition for the add-in to allow send.  However *Add-in2* removes any appearance of the word *Fabrikam* and then allows. The message would be allowed to send anyway with all *Fabrikam* removed due to the order of installation of  *Add-in1* and *Add-in2*.


## Deploying on send add-ins 

We recommend that administrators deploy Outlook add-ins that use the on send feature. Administrators have to ensure that the on send add-in:

- Is always present any time a compose item is opened (for email: new, reply, or forward).
- Can't be closed or disabled by the user.

## Installing a new on send add-in

On send functionality in Outlook Web App requires that add-ins are configured for the send event types. Newly installed on send add-ins for Outlook Web App will only be executed for users who are assigned an Outlook Web App mailbox policy that has the *OnSendAddinsEnabled* flag set to **true**.

  > **Note:** For scenario and cmdlets examples, see **Examples 1, 2 and 3** below.

To install a new add-in, run the following Exchange Online PowerShell cmdlets. 

```
$Data=Get-Content -Path '.\Contoso Message Body Checker.xml' -Encoding Byte –ReadCount 0 
```
```
New-App -OrganizationApp -FileData $Data -DefaultStateForUser Enabled
```

> **Note:** To learn how to use remote PowerShell to connect to Exchange Online, see [Connect to Exchange Online PowerShell](http://go.microsoft.com/fwlink/p/?LinkId=396554).

## Enable or disable on send add-in functionality 

By default, on send add-in functionality is disabled. Administrators can enable the functionality as required by running Exchange Online PowerShell cmdlets.
This section provides:

1. sample cmdlets for various scenarios related to enabling or disabling on send add-in functionality.   
2. list of mailbox types that are currently not supported.

### Mailbox type restrictions

The Outlook Web App on send functionality is only supported for user mailboxes. Currently the functionality is not supported for the following mailbox types:

1.	Shared mailboxes
2.	Offline mode
3.	Modern Group mailboxes

Outlook Web App won't allow sending if the Outlook Web App on send add-in feature is enabled for the above mentioned mailbox types. In cases where a user is responding in a modern group mailbox, the add-ins won't be executed and the message will be allowed.

### Example 1: Enabling on send add-ins for all users

To enable on send add-ins for all users, the steps are as follow. 

**Step 1: Create a new Outlook Web App mailbox policy**

```
New-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy
```

> **Note:** Administrators can use an existing policy, but it should be noted that on send add-in functionality is only supported on certain mailbox types (see **Mailbox type restrictions** section for more information). Unsupported mailboxes will be blocked from sending by default in Outlook Web App.

**Step 2: Enable on send add-in functionality**

```
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
```

**Step 3: Assign the policy to users**

```
Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy OWAOnSendAddinAllUserPolicy
```

### Example 2: Enable on send add-ins for a specific group of users

Let's say an administrator only wants to enable an Outlook Web App on send add-in functionality in an environment for Finance users (where the Finance users are in the Finance Department). The following are the steps.

**Step 1: Create a new Outlook Web App mailbox policy for the group**

```
New-OWAMailboxPolicy FinanceOWAPolicy
```

> **Note:** Administrators can use an existing policy, but it should be noted that on send add-in functionality is only supported on certain mailbox types (see **Mailbox type restrictions** section for more information). Unsupported mailboxes will be blocked from sending by default in Outlook Web App.


**Step 2: Enable on send add-in functionality**

```
Get-OWAMailboxPolicy FinanceOWAPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$true
```

**Step 3: Assign the policy to users**

```
$targetUsers = Get-Group 'Finance'|select -ExpandProperty members
$targetUsers | Get-User -Filter {RecipientTypeDetails -eq 'UserMailbox'}|Set-CASMailbox -OwaMailboxPolicy FinanceOWAPolicy
```

  >  **Note:** The administrator then waits up to 60 minutes for the policy to take effect or alternatively, restart Internet Information Services (IIS). Once the policy takes effect, all finance users will have Outlook Web App on send functionality enabled.

### Example 3:  Disable on send add-in functionality

If an administrator wants to disable the on send functionality for a user or assign an Outlook Web App mailbox policy that does not have the flag enabled (in this case ContosoCorpOWAPolicy), the cmdlets to run is as follows: 

```
Get-CASMailbox joe@contoso.com | Set-CASMailbox –OWAMailboxPolicy "ContosoCorpOWAPolicy"
```

  > **Note:** For more information about how to use the **Set-OwaMailboxPolicy** cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](https://technet.microsoft.com/en-us/library/dd297989(v=exchg.160).aspx)

If an administrator wants to disable the on send functionality for all users that have a specific Outlook Web App mailbox policy assigned that has the feature enabled, the cmdlets to run is as follows: 

```
Get-OWAMailboxPolicy OWAOnSendAddinAllUserPolicy | Set-OWAMailboxPolicy –OnSendAddinsEnabled:$false
```

 > **Note:** For more information about how to use the Set-OwaMailboxPolicy cmdlet to configure existing Outlook on the web mailbox policies, see [Set-OwaMailboxPolicy](https://technet.microsoft.com/en-us/library/dd297989(v=exchg.160).aspx).

# On send mailbox type scenarios

The following section describes the supported and unsupported usage scenarios for on send add-ins.

**Scenario 1: User mailbox with on send add-in feature enabled but no add-ins are installed**

- **Results:** In this scenario the user will be able to send mail without any add-ins executing.

**Scenario 2: User mailbox with on send add-in feature enabled and add-ins that supports on send are installed and enabled**

- **Results:** Add-ins will be initiated during the send event, which will then either allow or block the user from sending.

**Scenario 3: Mailbox delegation, where mailbox 1 has full access permissions to mailbox 2**

- Mailbox 1: On send add-in feature is enabled
- Mailbox 2: On send add-in feature is enabled
- Mailbox 1 opens Mailbox 2 in a new full Outlook Web App session (open another mailbox).
- **Results:** Mailbox 1 will not be able to send an email from Mailbox 2.

Scenario 3 is currently **not supported**. The workaround is to use scenario 5 for this use case.

**Scenario 4: Mailbox delegation, where mailbox 1 has full access permissions to mailbox 2**

- Mailbox 1: On send add-in feature is disabled
- Mailbox 2: On send add-in feature is enabled
- Mailbox 1 opens Mailbox 2 in a new full Outlook Web App session (open another mailbox).
- **Results:**  Mailbox 1 will not be able to send an email from Mailbox 2.

Scenario 4 is currently **not supported**. The workaround is to use scenario 5 for this use case.

**Scenario 5: Mailbox delegation, where mailbox 1 has full access permissions to mailbox 2**

- Mailbox 1: On send add-in feature is enabled
- Mailbox 2: On send add-in feature is enabled
- Mailbox 1 opens Mailbox 2 in the same Outlook Web App session (open shared folders).
- **Results:** The on send add-ins assigned to Mailbox 1 will be executed during send.

**Scenario 6: Mailbox delegation, where mailbox 1 has full access permissions to mailbox 2**

- Mailbox 1: On send add-in feature is enabled
- Mailbox 2: On send add-in feature is disabled
- Mailbox 1 opens Mailbox 2 in a new full Outlook Web App session (open another mailbox).
- **Results:** No on send add-ins will be executed. Mail will be sent as per normal.

**Scenario 7: Group 1 is a modern group mailbox and user mailbox 1 is a member of Group 1**

-	Mailbox 1: On send add-in feature is enabled and some on send add-ins are enabled.
-	Mailbox 1 composes a new message to Group 1.
- **Results:** The on send add-ins will be executed during send.

**Scenario 8: Group 1 is a modern group mailbox and user mailbox 1 is a member of Group 1**

-	Mailbox 1: On send add-in feature is enabled and some on send add-ins are enabled.
-	Mailbox 1 composes a new message to Group 1 within Group 1’s group window in Outlook Web App.
- **Results:** The on send add-ins will not be executed during send.

Scenario 8 is currently **not supported**. The workaround is to use scenario 7 for this use case.

**Scenario 9: User mailbox with on send add-in feature enabled, add-ins that support on send are installed and enabled and offline mode is enabled.**

- **Results:** The on send add-ins will be executed during send if the user is online. If the user is offline, the on send add-ins will not be executed during send and the email will not be sent.

# On send code sample scenario

In this section, we'll walkthrough a simple on send sample scenario and API usage.  To aid in the illustration and discussion, we'll use the [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) code sample.

## Manifest, version override and event

The [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) code sample includes two separate manifests:

- `Contoso Message Body Checker.xml` -- demonstrates how to check the body of a message for restricted words or sensitive information on send.  
- `Contoso Subject and CC Checker.xml` -- demonstrates how to add a recipient to the CC line and check that there is a subject in the message on send.  

In the `Contoso Message Body Checker.xml` manifest file, you include the function file and function name that should be called on **ItemSend** event (as shown below with comments). The operation is executed synchronously.

```
<Hosts>
        <Host xsi:type="MailHost">
          <DesktopFormFactor>
            <!-- The functionfile and function name to call on message send.  -->
            <!-- In this particular case the function validateBody will be called within the JavaScript code referenced in residUILessFunctionFileUrl. -->
            <FunctionFile resid="residUILessFunctionFileUrl" />
            <ExtensionPoint xsi:type="Events">
              <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="validateBody" />
            </ExtensionPoint>
          </DesktopFormFactor>
        </Host>
      </Hosts>
```

For the `Contoso Subject and CC Checker.xml` manifest file, the function file and function name to call on message send event looks as follows:

```
<Hosts>
        <Host xsi:type="MailHost">
          <DesktopFormFactor>
            <!-- The functionfile and function name to call on message send.  -->
            <!-- In this particular case the function validateSubjectAndCC will be called within the JavaScript code referenced in residUILessFunctionFileUrl. -->
            <FunctionFile resid="residUILessFunctionFileUrl" />
            <ExtensionPoint xsi:type="Events">
              <Event Type="ItemSend" FunctionExecution="synchronous" FunctionName="validateSubjectAndCC" />
            </ExtensionPoint>
          </DesktopFormFactor>
        </Host>
      </Hosts>
```



The on send API requires version **override V1_1**.  In your manifest, you add the version override as follows:

```
 <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <!-- On Send requires VersionOverridesV1_1 -->
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
```

 >  **Note:** For more information about:
   * running Outllook add-ins, see [Running an Outlook add-in in an Office 365 account](https://dev.outlook.com/MailAppsGettingStarted/GetStarted)
   * manifests, see [Outlook add-in manifests](https://dev.office.com/docs/add-ins/outlook/manifests/manifests), [VersionOverrides](https://dev.office.com/docs/add-ins/outlook/manifests/define-add-in-commands#versionoverrides), and [Office Add-ins XML manifest](https://dev.office.com/docs/add-ins/overview/add-in-manifests)


## Event, item, body getAsync and setAsync methods

To access the currently selected message which in this case is the newly composed  message, use the **Office.context.mailbox.item** namespace. The **ItemSend** event is automatically passed by the on send feature to the function specified in the manifest, which in this example is the `validateBody`function (as shown in the following code snippet taken from the [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) code sample).

```js
 var mailboxItem;

    Office.initialize = function (reason) {
        mailboxItem = Office.context.mailbox.item;
    }

    // Entry point for Contoso Message Body Checker add-in before send is allowed.
    // <param name="event">ItemSend event is automatically passed by on send code to the function specified in the manifest.</param>
    function validateBody(event) {
        mailboxItem.body.getAsync("html", { asyncContext: event }, checkBodyOnlyOnSendCallBack);
    }
```

In the `validateBody`function, the code then gets the current body in the specified format (html), passing the **ItemSend** event object the code wants to access in the callback method. In addition to the **getAsync** method, the **Body** object also provides a **setAsync** method that can be used to replace the entire body with the specified text. 

>  **Note:** For more information about:
   * Outlook events, see [Event Object](https://dev.outlook.com/reference/add-ins/Event.html).
   * message body **getAsync** and **setAsync** methods, see [getAsync](https://dev.outlook.com/reference/add-ins/Body.html)
  

## NotificationMessages object and event.completed method

In the `checkBodyOnlyOnSendCallBack` function, the code sample uses regular expression to check if the message body contains blocked words.  If it finds a match against an array of restricted words, it then blocks the email from being sent and notify the sender via the information bar.  To do this, it uses the **notificationMessages** property of the **Item** object to return a **NotificationMessages** object.  It then adds a notification to the item by calling the **addAsync** method (as shown in the following code snippet taken from the [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) code sample).

```js
  // Check if the body contains a specific set of blocked words. If it contains the blocked words, block email from being sent. Otherwise allows sending.
    // <param name="asyncResult">ItemeSend event passed from the calling function.</param>
    function checkBodyOnlyOnSendCallBack(asyncResult) {
        var listOfBlockedWords = new Array("blockedword", "blockedword1", "blockedword2");
        var wordExpression = listOfBlockedWords.join('|');

        // \b to perform a "whole words only" search using a regular expression in the form of \bword\b.
        // i to perform case-insensitive search.
        var regexCheck = new RegExp('\\b(' + wordExpression + ')\\b', 'i');
        var checkBody = regexCheck.test(asyncResult.value);

        if (checkBody) {
            mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Blocked words have been found in the body of this email. Please remove them.' });
            // Block send.
            asyncResult.asyncContext.completed({ allowEvent: false });
        }

        // Allow send.
        asyncResult.asyncContext.completed({ allowEvent: true });
    }
```

The following are the parameters for the **addAsync** method:
- 'NoSend' is a string and is a developer specified key to reference a notification message. Developers can use it to modify this message later. Key can’t be longer than 32 characters. 
- 'type' is one the properties of the  JSON object parameter.  Type is the type of a message corresponding to the [Office.MailboxEnums.ItemNotificationMessageType](https://dev.outlook.com/reference/add-ins/Office.MailboxEnums.html#.ItemNotificationMessageType) enumeration. The value can be progress indicator, information message or an error message. In this code sample, 'type' is an error message.  
- 'message'is one the properties of the JSON object parameter. In this case 'message' is the text of the notification message. 

To signal that the add-in has completed processing the **ItemSend** event triggered by the send operation, call the **event.completed({allowEvent:Boolean}** method.  The **allowEvent** property is a Boolean. If it's set to **true**, send is allowed. If false, then the email message is blocked from being sent.

>  **Note:** For more information about: 
   * notification messages for an item, see [notificationMessages](https://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html#notificationMessages)
   * **completed**  method, see [completed](https://dev.outlook.com/reference/add-ins/Event.html#completed)

## replaceAsync, removeAsync and getAllAsync methods

In addition to the **addAsync** method, the **NotificationMessages** object also includes **replaceAsync, removeAsync and getAllAsync** methods.  These methods are not used in this code sample.  For more information, see [NotificationMessages](https://dev.outlook.com/reference/add-ins/NotificationMessages.html).


## Subject and CC checker

In addition to the message body checker, the [Outlook-Add-in-On-Send](https://github.com/OfficeDev/Outlook-Add-in-On-Send) code sample also demonstrates how to add a recipient to the CC line and check that there is a subject in the message on send. It leverages the block on send event to allow or disallow an email from being sent.   The following code snippet shows how it's done.

```js
    // Invoke by Contoso Subject and CC Checker add-in before send is allowed.
    // <param name="event">ItemSend event is automatically passed by on send code to the function specified in the manifest.</param>
    function validateSubjectAndCC(event) {
        shouldChangeSubjectOnSend(event);
    }

    // Check if the subject should be changed. If it is already changed allow send. Otherwise change it.
    // <param name="event">ItemSend event passed from the calling function.</param>
    function shouldChangeSubjectOnSend(event) {
        mailboxItem.subject.getAsync(
            { asyncContext: event },
            function (asyncResult) {
                addCCOnSend(asyncResult.asyncContext);
                //console.log(asyncResult.value);
                // Match string.
                var checkSubject = (new RegExp(/\[Checked\]/)).test(asyncResult.value)
                // Add [Checked]: to subject line.
                subject = '[Checked]: ' + asyncResult.value;

                // Check if a string is blank, null or undefined.
                // If yes, block send and display information bar to notify sender to add a subject.
                if (asyncResult.value === null || (/^\s*$/).test(asyncResult.value)) {
                    mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Please enter a subject for this email.' });
                    asyncResult.asyncContext.completed({ allowEvent: false });
                }
                else {
                    // If can't find a [Checked]: string match in subject, call subjectOnSendChange function.
                    if (!checkSubject) {
                        subjectOnSendChange(subject, asyncResult.asyncContext);
                        //console.log(checkSubject);
                    }
                    else {
                        // Allow send.
                        asyncResult.asyncContext.completed({ allowEvent: true });
                    }
                }

            }
          )
    }

    // Add a CC to the email.  In this example, CC contoso@contoso.onmicrosoft.com
    // <param name="event">ItemSend event passed from calling function</param>
    function addCCOnSend(event) {
        mailboxItem.cc.setAsync(['Contoso@contoso.onmicrosoft.com'], { asyncContext: event });        
    }

    // Check if the subject should be changed. If it is already changed allow send, otherwise change it.
    // <param name="subject">Subject to set.</param>
    // <param name="event">ItemSend event passed from the calling function.</param>
    function subjectOnSendChange(subject, event) {
        mailboxItem.subject.setAsync(
            subject,
            { asyncContext: event },
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    mailboxItem.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Unable to set the subject.' });

                    // Block send.
                    asyncResult.asyncContext.completed({ allowEvent: false });
                }
                else {
                    // Allow send.
                    asyncResult.asyncContext.completed({ allowEvent: true });
                }

            });
    }
```

To learn more about how to add a recipient to the CC line and verify that the email message includes a subject line on send, and to see the APIs you can use, see the [Outlook-Add-in-On-Send sample](https://github.com/OfficeDev/Outlook-Add-in-On-Send). The code is well commented.   


## Additional resources

- [Overview of Outlook add-ins architecture and features](https://dev.office.com/docs/add-ins/outlook/overview?product=outlook)
    
- [Add-in Command Demo Outlook Add-in](https://github.com/jasonjoh/command-demo)
    
