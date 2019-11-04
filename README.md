
# OutlookCOMM

OutlookCOMM is a component which can be used in Microsoft Dynamics NAV as a workaround when Outlook can't be opened by NAV when composing or sending e-mails.


<p align="center">
  <img src="https://github.com/matteoparlato/OutlookCOMM/blob/master/Assets/OutlookCOMM_Diagram.png"/>
</p>


## Getting started

### Preparing the environment

* For NAV 2013R2 and later versions: 
  1. Copy the OutlookCOMM.NET.dll into the Add-in folder placed inside NAV service folder.
  2. Restart NAV service(s).

* For NAV 2013:
  1. Copy the OutlookCOMM.NET.dll into the Add-in folder placed inside both NAV service and RTC installation folders (the dll should be copied in every installed RTC).
  2. Restart NAV service(s).

* For NAV 2009R2 (RTC):
  1. Copy the OutlookCOMM.NET35.dll into the Add-in folder placed inside both NAV service and RTC installation folders (the dll should be copied in every installed RTC).

* For NAV 2009R2 Classic and previous versions:
  1. Run `RegisterOCOMMInterface.ps1` with PowerShell on every installed client.

  
### Applying changes to NAV objects

Here are the changes you need to apply to NAV objects in order to use OutlookCOMM:

**NOTE**

> May be necessary to apply further changes to the code for versions prior to NAV 2018 (looking at you NAV Classic :eyes:)


Codeunit 397 Mail
```perl
PROCEDURE NewMessageAsync@1000(ToAddresses@1001 : Text;CcAddresses@1002 : Text;BccAddresses@1000 : Text;Subject@1003 : Text;Body@1004 : Text;AttachFilename@1005 : Text;ShowNewMailDialogOnSend@1006 : Boolean) : Boolean;
VAR
  UserSetup@1101318000 : Record 91;
  MailUtilities@1101318001 : DotNet "'OutlookCOMM.NET, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'.OutlookCOMM.NET.MailUtilities" RUNONCLIENT;
BEGIN
  // Start OutlookCOMM/mpar
  // ORG: EXIT(CreateAndSendMessage(ToAddresses,CcAddresses,BccAddresses,Subject,Body,AttachFilename,ShowNewMailDialogOnSend,FALSE));
  IF UserSetup.GET(USERID) AND UserSetup."Use alternative E-Mail sending" AND ShowNewMailDialogOnSend THEN
    EXIT(MailUtilities.SaveEML(UserSetup."E-Mail", ToAddresses, CcAddresses, BccAddresses, Subject, Body, AttachFilename, TRUE, UserSetup."Use Outlook Account sender"))
  ELSE
    EXIT(CreateAndSendMessage(ToAddresses,CcAddresses,BccAddresses,Subject,Body,AttachFilename,ShowNewMailDialogOnSend,FALSE));
  // Stop OutlookCOMM/mpar
END;

PROCEDURE NewMessage@2(ToAddresses@1001 : Text;CcAddresses@1002 : Text;BccAddresses@1000 : Text;Subject@1003 : Text;Body@1004 : Text;AttachFilename@1005 : Text;ShowNewMailDialogOnSend@1006 : Boolean) : Boolean;
VAR
  UserSetup@1101318001 : Record 91;
  MailUtilities@1101318000 : DotNet "'OutlookCOMM.NET, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'.OutlookCOMM.NET.MailUtilities" RUNONCLIENT;
BEGIN
  // Start OutlookCOMM/mpar
  // ORG: EXIT(CreateAndSendMessage(ToAddresses,CcAddresses,BccAddresses,Subject,Body,AttachFilename,ShowNewMailDialogOnSend,TRUE));
  IF UserSetup.GET(USERID) AND UserSetup."Use alternative E-Mail sending" AND ShowNewMailDialogOnSend THEN
    EXIT(MailUtilities.SaveEML(UserSetup."E-Mail", ToAddresses, CcAddresses, BccAddresses, Subject, Body, AttachFilename, TRUE, UserSetup."Use Outlook Account sender"))
  ELSE
    EXIT(CreateAndSendMessage(ToAddresses,CcAddresses,BccAddresses,Subject,Body,AttachFilename,ShowNewMailDialogOnSend,TRUE));
  // Stop OutlookCOMM/mpar
END;
```

Table 91 User Setup
```perl
{ 50000;  ;Use alternative E-Mail sending;Boolean;
                                                CaptionML=[ENU=Use alternative E-Mail sending;
                                                          ITA=Usa metodo invio E-Mail alternativo];
                                                Description=OutlookCOMM, enables or disables the usage of the alternative E-Mail sending method for the user }
{ 50001;  ;Use Outlook Account sender;Boolean ;InitValue=Yes;
                                                CaptionML=[ENU=Use Outlook account as sender address;
                                                          ITA=Usa indirizzo mittente dell'account di Outlook];
                                                Description=OutlookCOMM, if enabled use Outlook E-Mail account as sender address otherwise use the value in the E-Mail field }
```

Codeunit 9520 Mail Management
```perl
LOCAL PROCEDURE SendMailOnWinClient@3() : Boolean;
VAR
  Mail@1003 : Codeunit 397;
  FileManagement@1006 : Codeunit 419;
  ClientAttachmentFilePath@1005 : Text;
  ClientAttachmentFullName@1009 : Text;
  BodyText@1000 : Text;
  MailUtilities@1101318001 : DotNet "'OutlookCOMM.NET, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'.OutlookCOMM.NET.MailUtilities" RUNONCLIENT;
  UserSetup@1101318000 : Record 91;
  EMLFilePath@1101318002 : Text;
BEGIN
  // Start OutlookCOMM/mpar
  IF UserSetup.GET(USERID) AND UserSetup."Use alternative E-Mail sending" THEN BEGIN
    CheckValidEmailAddress(TempEmailItem."Send to");
    EMLFilePath := FileManagement.DownloadTempCustomFile(TempEmailItem."Attachment File Path", TempEmailItem."Attachment Name");
    EXIT(MailUtilities.SaveEML(UserSetup."E-Mail",
                                TempEmailItem."Send to",
                                TempEmailItem."Send CC",
                                TempEmailItem."Send BCC",
                                TempEmailItem.Subject,
                                ImageBase64ToUrl(TempEmailItem.GetBodyText),
                                EMLFilePath,
                                TRUE,
                                UserSetup."Use Outlook Account sender"));
  END;
  // Stop OutlookCOMM/mpar
  .
  .
  .
END;
.
.
.
PROCEDURE SendMailOrDownload@17(TempEmailItem@1002 : TEMPORARY Record 9500;HideMailDialog@1000 : Boolean);
VAR
  MailManagement@1001 : Codeunit 9520;
  UserSetup@1101318001 : Record 91;
  MailUtilities@1101318000 : DotNet "'OutlookCOMM.NET, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'.OutlookCOMM.NET.MailUtilities" RUNONCLIENT;
  EMLFilePath@1101318002 : Text;
BEGIN
  // Start OutlookCOMM/mpar
  IF UserSetup.GET(USERID) AND UserSetup."Use alternative E-Mail sending" THEN BEGIN
    CheckValidEmailAddress(TempEmailItem."Send to");
    EMLFilePath := FileManagement.DownloadTempCustomFile(TempEmailItem."Attachment File Path", TempEmailItem."Attachment Name");
    MailUtilities.SaveEML(UserSetup."E-Mail",
                          TempEmailItem."Send to",
                          TempEmailItem."Send CC",
                          TempEmailItem."Send BCC",
                          TempEmailItem.Subject,
                          ImageBase64ToUrl(TempEmailItem.GetBodyText),
                          EMLFilePath,
                          TRUE,
                          UserSetup."Use Outlook Account sender");
    EXIT;
  END;
  // Stop OutlookCOMM/mpar
  .
  .
  .
END;
```

Codeunit 419 File Management
```perl
PROCEDURE DownloadTempCustomFile@1101318000(ServerFileName@1001 : Text;ClientFileName@1101318000 : Text) : Text;
VAR
  FileName@1102601003 : Text;
  Path@1102601004 : Text;
  DotPath@1101318001 : DotNet "'mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'.System.IO.Path";
BEGIN
  // Start OutlookCOMM/mpar
  FileName := DotPath.Combine(DotPath.GetDirectoryName(ServerFileName), ClientFileName);
  Path := Magicpath;
  DOWNLOAD(ServerFileName,'',Path,AllFilesDescriptionTxt,FileName);
  EXIT(FileName);
  // Stop OutlookCOMM/mpar
END;
```

Page 18006682 User Setup Card
```perl
ActionList=ACTIONS
{
  { 1101318004;  ;ActionContainer;
                  ActionContainerType=ActionItems }
  { 1101318006;1 ;Action    ;
                  Name=Test Alternative Sending Method;
                  CaptionML=[ENU=Test alternative E-Mail sending method;
                              ITA=Esegui test invio E-Mail alternativo];
                  Promoted=Yes;
                  Enabled="Use alternative E-Mail sending";
                  PromotedIsBig=Yes;
                  Image=MailSetup;
                  PromotedCategory=Process;
                  PromotedOnly=Yes;
                  OnAction=VAR
                              MailUtilities@1101318000 : DotNet "'OutlookCOMM.NET, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'.OutlookCOMM.NET.MailUtilities" RUNONCLIENT;
                            BEGIN
                              IF NOT "Use Outlook Account sender" THEN
                                TESTFIELD("E-Mail");
                              IF NOT MailUtilities.SaveEML("E-Mail", 'to@example.com', 'cc@example.com', 'bcc@example.com', 'Subject', 'Body', '', TRUE, "Use Outlook Account sender") THEN
                                MESSAGE(OutlookCOMMErrorText);
                            END;
                            }
}
.
.
.
{ 1101318002;1;Group  ;
            CaptionML=[ENU=OutlookCOMM;
                        ITA=OutlookCOMM];
            GroupType=Group }

{ 1101318003;2;Field  ;
            ToolTipML=[ENU=Enable this feature only if the traditional E-Mail sending through the direct call to the Outlook component does not work properly.;
                        ITA=Abilita questa funzione solo se l'invio E-Mail tradizionale tramite la chiamata diretta al componente di Outlook non funziona correttamente.];
            SourceExpr="Use alternative E-Mail sending" }

{ 1101318005;2;Field  ;
            ToolTipML=[ENU=If enabled use Outlook email account as sender address otherwise use the value defined in the E-Mail field.;
                        ITA=Se abilitato, utilizza l'account e-mail di Outlook come indirizzo mittente, altrimenti utilizza il valore definito nel campo E-mail.];
            SourceExpr="Use Outlook Account sender";
            Enabled="Use alternative E-Mail sending" }
```

### Test and setup OutlookCOMM in NAV

To enable the alternative e-mail sending method you need to enable the "Use alternative E-Mail sending" flags:

<p align="center">
  <img src="https://github.com/matteoparlato/OutlookCOMM/blob/master/Assets/UserSetupCard.png"/>
</p>

If you want to use the e-mail defined in the "User setup" instead of the one configured in Outlook you need to fill the field "E-Mail" with a valid e-mail address.

When enabled, you can test the alternative e-mail sending method just by pressing the "Test alternative E-Mail sending method" button in the "HOME" panel of the ribbon.

Any exception fired by OutlookCOMM will be registed inside Windows "Applications" event log so take a look at Windows Event Viewer logs when debugging in NAV.

OutlookCOMM can also be tested in Visual Studio using the OutlookCOMM.Test project.

## Downloads

|NAV Classic|NAV 2009 (RTC)|NAV 2013+|
| :---: | :---: | :---: |
|COM Events|.NET 3.5|.NET 4.0+|
|[Download](https://github.com/matteoparlato/OutlookCOMM/releases)|[Download](https://github.com/matteoparlato/OutlookCOMM/releases)|[Download](https://github.com/matteoparlato/OutlookCOMM/releases)| 

## Developing, testing and deploying OutlookCOMM

For all those who want to extend the functionality or fix bugs (hope not) here are the prerequisites for developing OutlookCOMM:

  - Visual Studio 2017+ with .NET desktop development workload.
  - Microsoft Dynamics NAV 2009R2 (Classic) or any previous version for COM interface tests.
  - Microsoft Dynamics NAV 2009R2 (RTC) or any later version for .NET library tests.
  - Any email client with *.eml file support (Outlook or Thunderbird recommended).

## Authors

* [**Matteo Parlato**](https://github.com/matteoparlato)

## Special thanks

* Alvise Giacomin
* Paolo Garlatti
* Everyone who will use, appreciate and share this tool!

## License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE) file for details.