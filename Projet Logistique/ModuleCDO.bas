Attribute VB_Name = "ModuleCDO"
Option Explicit

Public Function MailEnvoi(Serveur, Identify As Boolean, User, PassWord, Port, Delay, Expediteur, Dest, DestEnCopy, Objet, Body, Pj) As Boolean
On Error GoTo Fin
' sub pour envoyer les mails
Dim msg
Dim Conf
Dim Config
Dim ess
Set msg = CreateObject("CDO.Message") 'pour la configuration du message
Set Conf = CreateObject("CDO.Configuration") '  pour la configuration de l'envoi
Dim strHTML
Set Config = Conf.Fields
MailEnvoi = True
' Configuration des parametres d'envoi
'(SMTP - Identification - SSL - Password - Nom Utilisateur - Adresse messagerie)
With Config
If Identify = False Then GoTo Anon
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = User
    .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = PassWord
Anon:
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = Port
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Serveur
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = Delay
    .Update
End With
DoEvents

'Configuration du message
'If E_mail.Sign.Value = Checked Then Convert ServeurFrm.SignTXT, ServeurFrm.Text1

With msg
    Set .Configuration = Conf
    .To = Dest
  .cc = DestEnCopy
    .From = Expediteur
    .Subject = Objet
'            If E_mail.Sign.Value = 1 Then _
    .htmlbody = E_mail.ZThtml.Text & "<p align=""left""><font face=""MS Sans Serif"" size=""1"" color=""#000000""><b>" & "---------------------------------------" & "<P></P>" & ServeurFrm.Text1.Text _
            Else _
.sender"toto"

    .htmlbody = Body '"<p align=""center""><font face=""Verdana"" size=""1"" color=""#9224FF""><b><br><font face=""Comic Sans MS"" size=""5"" color=""#FF0000""></b><i>" & body & "</i></font> " 'E_mail.ZThtml.Text
            If Pj <> "" Then
            Dim Pj2
            Dim i As Long
         Pj2 = Split(Pj & ";", ";")
         For i = 0 To UBound(Pj2) - 1
            If Trim(Pj2(i)) <> "" Then
               .AddAttachment Replace(Replace(Trim(Pj2(i)), Chr(10), ""), Chr(13), "")
            End If
         Next
         End If
    .Send 'envoi du message

End With
DoEvents
' reinitialisation des variables
Set msg = Nothing
Set Conf = Nothing
Set Config = Nothing

DoEvents
Exit Function
Fin:
MsgBox Err.Description
MailEnvoi = False
End Function

