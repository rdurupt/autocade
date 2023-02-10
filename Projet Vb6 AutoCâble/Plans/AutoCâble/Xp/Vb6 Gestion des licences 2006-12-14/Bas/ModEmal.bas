Attribute VB_Name = "ModEmal"
Option Explicit



Public Port As Long
Public SSL As Boolean
Public Autentic 'As Integer
Public PassWord As String
Public Serveur As String
'Public Identify As String
Public Expediteur As String
Public Delay As Long
Public Yn As Long
Public Yns As Long
Public User As String
Public ConV

Public sig As String
Function ReplaceHtml(Txt)
ReplaceHtml = Replace(Txt, Chr(10), "<br>")
End Function
Public Sub SendMal(Routine As String, Pj As String)
Dim Rs As Recordset
Dim Sql As String
Dim Destinataire As String
Dim Sujet As String
Dim Body As String
Sql = "SELECT T_Message_Mail.Sujet,T_Message_Mail.Body, T_Users.Email "
Sql = Sql & "FROM T_Users INNER JOIN (T_Message_Mail INNER JOIN T_Destinataire ON T_Message_Mail.Id = "
Sql = Sql & "T_Destinataire.Id_Message) ON T_Users.Id = T_Destinataire.Id_Useur "
Sql = Sql & "WHERE T_Message_Mail.Routine='" & MyReplace(Routine) & " ' "
Sql = Sql & "AND T_Users.Email Is Not Null;"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    While Rs.EOF = False
        Destinataire = Destinataire & Rs!EMail & ";"
        Sujet = Rs!Sujet
        Body = ReplaceHtml(Rs!Body)
        Rs.MoveNext
    Wend
    Destinataire = Left(Destinataire, Len(Destinataire) - 1)
    Set Rs = Con.OpenRecordSet("SELECT T_Serveur_Smtp.* FROM T_Serveur_Smtp;")
    MailEnvoi Rs!SMTP, Rs!Authentification, Rs!Utilisatuer, Rs!PassWord, Rs!Port, 15, Rs!Messagerie, Destinataire, "", Sujet, Body, Pj

End If



End Sub


'Public reg As ZebClass
'Public Sub MailEnvoi(Serveur, Identify As Boolean, User, PassWord, Port, Delay, Expediteur, Dest, DestEnCopy, Objet, Body, Pj)
'' sub pour envoyer les mails
'Dim msg
'Dim Conf
'Dim Config
'Dim ess
'Set msg = CreateObject("CDO.Message") 'pour la configuration du message
'Set Conf = CreateObject("CDO.Configuration") '  pour la configuration de l'envoi
'Dim strHTML
'Set Config = Conf.Fields
'
'' Configuration des parametres d'envoi
''(SMTP - Identification - SSL - Password - Nom Utilisateur - Adresse messagerie)
'With Config
'If Identify = False Then GoTo Anon
'    .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
'    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
'    .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = User
'    .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = PassWord
'Anon:
'    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = Port
'    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Serveur
'    .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = Delay
'    .Update
'End With
'DoEvents
'
''Configuration du message
''If E_mail.Sign.Value = Checked Then Convert ServeurFrm.SignTXT, ServeurFrm.Text1
'
'With msg
'    Set .Configuration = Conf
'    .To = Dest
'  .cc = DestEnCopy
'    .FROM = Expediteur
'    .Subject = Objet
''            If E_mail.Sign.Value = 1 Then _
'    .htmlbody = E_mail.ZThtml.Text & "<p align=""left""><font face=""MS Sans Serif"" size=""1"" color=""#000000""><b>" & "---------------------------------------" & "<P></P>" & ServeurFrm.Text1.Text _
'            Else _
'.sender"toto"
'
'    .htmlbody = Body '"<p align=""center""><font face=""Verdana"" size=""1"" color=""#9224FF""><b><br><font face=""Comic Sans MS"" size=""5"" color=""#FF0000""></b><i>" & body & "</i></font> " 'E_mail.ZThtml.Text
'    Dim splitPj
'    If Pj <> "" Then
'        splitPj = Split(Pj & ";", ";")
'        Dim IsplitPj As Long
'        For IsplitPj = 0 To UBound(splitPj)
'            If Trim("" & splitPj(IsplitPj)) <> "" Then
'                .AddAttachment Trim("" & splitPj(IsplitPj))
'            End If
'        Next
'
'    End If
'    .Send 'envoi du message
'
'End With
'DoEvents
'' reinitialisation des variables
'Set msg = Nothing
'Set Conf = Nothing
'Set Config = Nothing
'
'DoEvents
'End Sub

Public Sub MailEnvoi(Serveur, Identify As Boolean, User, PassWord, Port, Delay, Expediteur, Dest, DestEnCopy, Objet, Body, Pj)
' sub pour envoyer les mails
Dim msg
Dim Conf
Dim Config
Dim ess
Dim splitPj
Dim IsplitPj As Long
Set msg = CreateObject("CDO.Message") 'pour la configuration du message
Set Conf = CreateObject("CDO.Configuration") '  pour la configuration de l'envoi
Dim strHTML

Set Config = Conf.Fields

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
    .FROM = Expediteur
    .Subject = Objet
'            If E_mail.Sign.Value = 1 Then _
    .htmlbody = E_mail.ZThtml.Text & "<p align=""left""><font face=""MS Sans Serif"" size=""1"" color=""#000000""><b>" & "---------------------------------------" & "<P></P>" & ServeurFrm.Text1.Text _
            Else _
.sender"toto"

    .htmlbody = Body '"<p align=""center""><font face=""Verdana"" size=""1"" color=""#9224FF""><b><br><font face=""Comic Sans MS"" size=""5"" color=""#FF0000""></b><i>" & body & "</i></font> " 'E_mail.ZThtml.Text
            If Pj <> "" Then
        splitPj = Split(Pj & ";", ";")
        
        For IsplitPj = 0 To UBound(splitPj)
            If Trim("" & splitPj(IsplitPj)) <> "" Then
                .AddAttachment Trim("" & splitPj(IsplitPj))
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
End Sub
