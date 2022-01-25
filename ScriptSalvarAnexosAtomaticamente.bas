Attribute VB_Name = "Módulo"
Public Sub SaveAttachmentsToDiski9(MItem As Outlook.MailItem)

Dim oAttachment As Outlook.Attachment
Dim sSaveFolder As String

sSaveFolder = "xxxxxxxx"

For Each oAttachment In MItem.Attachments

    oAttachment.SaveAsFile sSaveFolder & oAttachment.DisplayName

Next

End Sub

