Option Explicit

Dim thesubject, recipients, thebodyfile, theattachment , fso, qfile, thebody, theattachmentfolder, theattachmentfilepattern, fileFiles, fileFolder, objFile
Dim objOutlook, objNamespace, objFolder, objMail

dim mailbox
set mailbox = "yourname@domain.com"

thesubject               = WScript.Arguments(0)
recipients               = WScript.Arguments(1)
thebodyfile              = WScript.Arguments(2)
theattachmentfolder      = WScript.Arguments(3)
theattachmentfilepattern = Wscript.arguments(4)

    Set    fso           = CreateObject("Scripting.FileSystemObject")
    set    qfile         = fso.OpenTextFile(thebodyfile, 1, TRUE)
           thebody       = qfile.ReadAll
           qfile.Close
   
    ' Connect to Outlook

    Set    objOutlook    = CreateObject("Outlook.Application")
    Set    objNamespace  = objOutlook.GetNamespace("MAPI")
    Set    objFolder     = objNamespace.Folders(mailbox)
    Set    objMail       = objFolder.Items.Add

  ' Set email properties
  
    objMail.To           = recipients
  ' objMail.CC           = ccguys
    objMail.Subject      = thesubject
    objMail.HTMLBody     = thebody

    Set fileFolder       = fso.GetFolder(theattachmentfolder)
    Set fileFiles        = fileFolder.Files

    ' Loop through file collection and attach each file to email
    For Each objFile in fileFiles
    
          If      InStr(fso.GetFileName(objFile.Path), theattachmentfilepattern) > 0 _
          Then
                  objMail.Attachments.Add objFile.Path
          End If

    Next


   ' Send email
    objMail.Send
   'objMail.Display


    ' Release Objects and Referemces

    Set objMail         = Nothing
    Set objFolder       = Nothing
    Set objNamespace    = Nothing
    Set objOutlook      = Nothing
