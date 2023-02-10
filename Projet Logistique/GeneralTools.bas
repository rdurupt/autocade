Attribute VB_Name = "GeneralTools"
Const a = "t"
    Global DefaultFontFace
    Global BadWordFiler
    Global BadWords
    Global Application  As Application
    Global Request As Request
    Global Session As Session
    Global Response As Response
    Global Server As Server
    Public My_Conn As ADODB.Connection
    Public My_Conn2 As ADODB.Connection
    '             border,   headerFG,  headerBG,  headrHiFg, hdrHiBg ,itmFgColor, itmBgColor, itmHiFgColor, itmHiBgColor
    Global Const MenuColor = "'#FFFFFF', '#FFFFFF', '#0099CC', '#006666', '#C0E7EF', '#000080', '#F0AC07','#FFFFFF', '#000080' "
    '            <---------- HEADER ----------------><---------- ITEMS----------------------->
    Global Const MenuFonts = " 'Verdana', 'plain', 'bold', 'xx-small', 'Verdana', 'plain', 'bold', 'xx-small' "
    'BorderSize,Height,SepSize
    Global Const BorderSize = "1, 3, 1 "
    '            <---------- HEADER ----------------><---------- ITEMS----------------------->
    Global Const ImageSet = "'','bouton_on.jpg','bouton_off.jpg'"
    
    Const ForReading = 1
    Const ForWriting = 2
    Const ForAppending = 8
    Const GlobalStyle = "<link rel=STYLESHEET href=""PMainStyle1.css"" type=""text/css""> "
Public Function DsnTableMenu(candidat_ADOContact As String) As String
Dim MyConn As ADODB.Connection
Dim Rs As Recordset
Set MyConn = OpenDb(candidat_ADOContact)
    
    DsnTableMenu = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ="
   
   Set Rs = MyConn.Execute("SELECT BaseDefault.Path FROM BaseDefault;")

If Rs.EOF = False Then
    DsnTableMenu = DsnTableMenu & Rs("Path")

End If
MyConn.Close
Set MyConn = Nothing
End Function
Public Function Retourne_SMTP(candidat_ADOContact As String) As String
Dim MyCon
 Dim Sql As String
 Dim Rs
 
Set MyCon = OpenDb(DsnTableMenu(candidat_ADOContact))
    Sql = "SELECT T_Serveur_Smtp.* FROM T_Serveur_Smtp;"
Set Rs = MyCon.Execute(Sql)
Retourne_SMTP = Rs.GetString(, , ";", vbCrLf)
Rs.Close
Set Rs = Nothing
MyCon.Close
Set MyCon = Nothing
End Function
 Public Sub MailEnvoi(Serveur, Identify As Boolean, User, PassWord, Port, Delay, Expediteur, Dest, DestEnCopy, Objet, Body, Pj)
' sub pour envoyer les mails
Dim msg
Dim Conf
Dim Config
Dim ess
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
    .From = Expediteur
    .Subject = Objet
'            If E_mail.Sign.Value = 1 Then _
    .htmlbody = E_mail.ZThtml.Text & "<p align=""left""><font face=""MS Sans Serif"" size=""1"" color=""#000000""><b>" & "---------------------------------------" & "<P></P>" & ServeurFrm.Text1.Text _
            Else _
.sender"toto"

    .htmlbody = Body '"<p align=""center""><font face=""Verdana"" size=""1"" color=""#9224FF""><b><br><font face=""Comic Sans MS"" size=""5"" color=""#FF0000""></b><i>" & body & "</i></font> " 'E_mail.ZThtml.Text
            If Pj <> "" Then _
    .AddAttachment Pj
    .Send 'envoi du message

End With
DoEvents
' reinitialisation des variables
Set msg = Nothing
Set Conf = Nothing
Set Config = Nothing

DoEvents
End Sub
Public Function noquoteNum(s0)
    numtxt = s0
    Dim Numrique As Boolean
    
    If IsNumeric(numtxt) Then Numrique = True
    If Numrique = False Then numtxt = Replace(numtxt, ",", ".")
     
    If IsNumeric(numtxt) Then Numrique = True
    If Numrique = False Then numtxt = Replace(numtxt, ".", ",")
     
    If IsNumeric(numtxt) Then Numrique = True
       
    If Numrique = False Then
        noquoteNum = "NULL"
    Else
        noquoteNum = Replace(numtxt, ",", ".")
    End If
  
End Function
Public Function noquoteNumTxt(s0)
    numtxt = s0
    Dim Numrique As Boolean
    
    If IsNumeric(numtxt) Then Numrique = True
    If Numrique = False Then numtxt = Replace(numtxt, ",", ".")
     
    If IsNumeric(numtxt) Then Numrique = True
    If Numrique = False Then numtxt = Replace(numtxt, ".", ",")
     
    If IsNumeric(numtxt) Then Numrique = True
       
    If Numrique = False Then
        noquoteNumTxt = s0
    Else
        noquoteNumTxt = Replace(numtxt, ",", ".")
    End If

End Function
   
Public Function pr(strPrint)
    Response.Write strPrint & vbCrLf
End Function
Public Function safeEntry(strField)

    strSafe = Trim(strField)
    strSafe = funReplace(strSafe, "'", "´")
    strSafe = funReplace(strSafe, "<", "&lt;")
    strSafe = funReplace(strSafe, ">", "&gt;")
    safeEntry = strSafe
End Function

Public Function CheckHash(DataToCheck, CryptedData, Salt)

            Set CM = Server.CreateObject("AspCrypt.Crypt")
            If CM.Crypt(Salt, DataToCheck) = CryptedData Then
                CheckHash = 1
            Else
                CheckHash = 0
            End If
End Function
Sub HashIt(DataToHash)

            Randomize
            Salt = ""
           
            For I = 1 To 10
                '65 is ASCII for "A"
                Salt = Salt & Chr(Int(Rnd * 26) + 65)
            Next
            ' Calculate Hash of (Password & Salt)
            Set CM = Server.CreateObject("AspCrypt.Crypt")
            Session("candidat_HashData") = CM.Crypt(Salt, DataToHash)
            Session("candidat_HashSecure") = Salt
End Sub
Function funY2K(D)
    strDate = Trim(D)
    If InStr(strDate, " ") Then
        strDate = Left(strDate, InStr(strDate, " "))
        trailer = Right(D, Len(D) - InStr(D, " "))
    End If
    If IsDate(strDate) Then
        dateY2K = strDate
        If InStr(strDate, "/") = 2 Then
            strMonth = Left(strDate, 1)
            If InStr(3, strDate, "/") = 4 Then
                strDay = Mid(strDate, 3, 1)
            Else
                strDay = Mid(strDate, 3, 2)
            End If
            strYear = Right(strDate, Len(strDate) - InStr(3, strDate, "/"))
        ElseIf InStr(strDate, "/") = 3 Then
            strMonth = Left(strDate, 2)
            If InStr(4, strDate, "/") = 5 Then
                strDay = Mid(strDate, 4, 1)
            Else
                strDay = Mid(strDate, 4, 2)
            End If
            strYear = Right(strDate, Len(strDate) - InStr(4, strDate, "/"))
        End If
        intYear = CInt(strYear)
        If intYear >= 0 And intYear < 51 Then
            strYear = "20" & strYear
        ElseIf intYear > 50 And intYear < 100 Then
            strYear = "19" & strYear
        End If
        
        funY2K = strMonth & "/" & strDay & "/" & strYear & " " & trailer
    Else
        funY2K = ""
    End If
End Function
Public Function funReplace(a, b, C)
    funReplace = Replace(a, b, C)
End Function
Function getLetter(Num)
    If Num = 1 Then
        getLetter = "a"
    ElseIf Num = 2 Then
        getLetter = "b"
    ElseIf Num = 3 Then
        getLetter = "c"
    ElseIf Num = 4 Then
        getLetter = "d"
    ElseIf Num = 5 Then
        getLetter = "e"
    ElseIf Num = 6 Then
        getLetter = "f"
    ElseIf Num = 7 Then
        getLetter = "g"
    ElseIf Num = 8 Then
        getLetter = "h"
    ElseIf Num = 9 Then
        getLetter = "i"
    ElseIf Num = 10 Then
        getLetter = "j"
    ElseIf Num = 11 Then
        getLetter = "k"
    ElseIf Num = 12 Then
        getLetter = "l"
    ElseIf Num = 13 Then
        getLetter = "m"
    ElseIf Num = 14 Then
        getLetter = "n"
    ElseIf Num = 15 Then
        getLetter = "o"
    ElseIf Num = 16 Then
        getLetter = "p"
    ElseIf Num = 17 Then
        getLetter = "q"
    ElseIf Num = 18 Then
        getLetter = "r"
    ElseIf Num = 19 Then
        getLetter = "s"
    ElseIf Num = 20 Then
        getLetter = "t"
    ElseIf Num = 21 Then
        getLetter = "u"
    ElseIf Num = 22 Then
        getLetter = "v"
    ElseIf Num = 23 Then
        getLetter = "w"
    ElseIf Num = 24 Then
        getLetter = "x"
    ElseIf Num = 25 Then
        getLetter = "y"
    ElseIf Num = 26 Then
        getLetter = "z"
    Else
        getLetter = ""
    End If
End Function

Function ChkString(str)

     If str = "" Then
        str = " "
     Else
        If BadWordFiler = "true" Then
          bwords = Split(BadWords, "|")

          For I = 0 To UBound(bwords)
            str = Replace(str, bwords(I), String(Len(bwords(I)), "*"), 1, -1, 1)
          Next
        End If
     End If
     
     '  Do ASP Forum Code
     str = doCode(str, "[b]", "[/b]", "<b>", "</b>")
     str = doCode(str, "[i]", "[/i]", "<i>", "</i>")
     str = doCode(str, "[quote]", "[/quote]", "<BLOCKQUOTE><font size=1 face=arial>quote:<hr height=1 noshade>", "<hr height=1 noshade></BLOCKQUOTE></font><font face='" & DefaultFontFace & "' size=2>")
     str = doCode(str, "[a]", "[/a]", "<a>", "</a>")
     str = doCode(str, "[code]", "[/code]", "<pre>", "</pre>")
     
     
     str = Replace(str, "'", "''")
     str = Replace(str, "|", "/")
     
     ChkString = Trim(str)
End Function

Function doCode(str, oTag, cTag, roTag, rcTag)

    tx = Split(str, cTag)
    T = ""

    For I = 0 To UBound(tx)

      If LCase(oTag) = "[a]" Then
        P = InStr(1, tx(I), "[a]", 1)
        If P <> 0 Then
            tmp = Mid(tx(I), P)
            Url = Mid(tmp, 4)
            If LCase(Left(Url, 5)) = "http:" Then
                tmp1 = Replace(tmp, "[a]" & Url, "<A href='" & Url & "' Target=_Blank>Link</a>", 1, -1, 1)
            Else
                tmp1 = Replace(tmp, "[a]" & Url, "<A href='http://" & Url & "' Target=_Blank>Link</a>", 1, -1, 1)
            End If
            T = T & Replace(tx(I), tmp, tmp1)
        Else
            T = T & tx(I)
        End If
      Else
        cnt = InStr(1, tx(I), oTag, 1)
        Select Case cnt
            Case 0
                T = T & tx(I) & " "
            Case Else
                T = T & Replace(tx(I), oTag, roTag, 1, 1, 1)
                T = T & " " & rcTag & " "
        End Select
      End If
    Next
    doCode = T
End Function
Public Function CovertCommDate(MyNum As String) As String
Dim SplitNum
SplitNum = Split(MyNum & "__", "_")
CovertCommDate = Format(SplitNum(1), "dd/mm/yyyy")
End Function

Public Function OpenDb(MyDb)
Dim Con
Set Con = CreateObject("ADODB.Connection")
Con.Mode = 16
    Con.Open MyDb
Set OpenDb = Con
End Function
Public Function ReplaceChamp(txt As String, Form As String) As String
Dim Mytext As String
Dim SplitTxt
Debug.Print txt
Debug.Print UCase(Form)
ReplaceChamp = txt

Select Case UCase(Form)
Case "CONFIGCLI"
    ReplaceChamp = Replace(ReplaceChamp, "Société", "fld3")
        ReplaceChamp = Replace(ReplaceChamp, "Code Client", "[t_R_Social].[SocieteId]")
        ReplaceChamp = Replace(ReplaceChamp, "Cp", "Zip ")
        ReplaceChamp = Replace(ReplaceChamp, "Ville", "City")
        ReplaceChamp = Replace(ReplaceChamp, "Pays", "label_pays")
        ReplaceChamp = Replace(ReplaceChamp, "Créé le", "t_R_Social.DateCreation")
        ReplaceChamp = Replace(ReplaceChamp, "Contact email", "Users.UserEMail")
        ReplaceChamp = Replace(ReplaceChamp, "Contact login", "Users.UserLogin")
        ReplaceChamp = Replace(ReplaceChamp, "Liste rouge", "Users.Listerouge")
        '"Listerouge", "Users",
Case "HISTORIQUE_BR"
    '"Solder", "T_BR", "Soldé le"
     ReplaceChamp = Replace(ReplaceChamp, "N° Br", "NumBr")
     ReplaceChamp = Replace(ReplaceChamp, "Créé le", "T_BR.DateCration")
     ReplaceChamp = Replace(ReplaceChamp, "Emal le", "T_BR.BrMailDate")
     ReplaceChamp = Replace(ReplaceChamp, "Soldé le", "T_BR.Solder")
    
Case "AFFICHELISTECOMMANDDUJOUR"
    ReplaceChamp = Replace(ReplaceChamp, "Date Commande", "[Date Commande]")
    ReplaceChamp = Replace(ReplaceChamp, "Ref Groupe Commande", "[Ref Groupe Commande]")
     ReplaceChamp = Replace(ReplaceChamp, "[[Date Commande]]", "[Date Commande]")
     ReplaceChamp = Replace(ReplaceChamp, "[[Ref Groupe Commande]]", "[Ref Groupe Commande]")


Case "HISTORIQUE_COMANDE_GROUPEE"
     ReplaceChamp = Replace(ReplaceChamp, "Date Commande", "[Date Commande]")
     ReplaceChamp = Replace(ReplaceChamp, "Ref Groupe Commande", "[Ref Groupe Commande]")
     ReplaceChamp = Replace(ReplaceChamp, "[[Date Commande]]", "[Date Commande]")
     ReplaceChamp = Replace(ReplaceChamp, "[[Ref Groupe Commande]]", "[Ref Groupe Commande]")

Case "HISTORIQUE_EBOUTIQUE_FACTURE"
        ReplaceChamp = Replace(ReplaceChamp, "N° Facture", "T_Facturation_Cli.NumFacture")
        ReplaceChamp = Replace(ReplaceChamp, "Motant Versé", "Sum (T_Facturation_Cli.MotantVerse)")
        ReplaceChamp = Replace(ReplaceChamp, "Reste Payer", "Sum(T_Facturation_Cli.RestePayer)")
        ReplaceChamp = Replace(ReplaceChamp, "Montant à Encaiser", "Sum(T_Facturation_Cli.MontantEncaiser)")
        '
Case "HISTORIQUE_BL"
    ReplaceChamp = Replace(ReplaceChamp, "Créé le", "T_Commande_Liv.creation")
    ReplaceChamp = Replace(ReplaceChamp, "Créer le", "T_Commande_Liv.creation")
Case "HISTORIQUE_EBOUTIQUE_COMMANDE"
        ReplaceChamp = Replace(ReplaceChamp, "Ref Commande", "[Ref Commande]")
        ReplaceChamp = Replace(ReplaceChamp, "Prix Total HT", "[Total HT]")
        ReplaceChamp = Replace(ReplaceChamp, "Prix Total TTC", "[Expr2]+[Frais TTC]")
        ReplaceChamp = Replace(ReplaceChamp, "Mode de Transport", "[Mode de Transport]")
        
        '"creation", "T_Commande_Liv", "Créé le"
    Case "HISTORIQUE_CLIENT"
        
        ReplaceChamp = Replace(ReplaceChamp, "Société", "fld3")
        ReplaceChamp = Replace(ReplaceChamp, "Code Client", "[t_R_Social].[SocieteId] ")
        ReplaceChamp = Replace(ReplaceChamp, "Cp", "Zip ")
        ReplaceChamp = Replace(ReplaceChamp, "Ville", "City ")
        ReplaceChamp = Replace(ReplaceChamp, "Pays", "label_pays ")
        ReplaceChamp = Replace(ReplaceChamp, "Liste rouge", "t_R_Social.Listerouge ")
        '"Listerouge", "Users", "Liste rouge"
    Case "HISTORIQUE_COMPTEDETAL"
        ReplaceChamp = Replace(ReplaceChamp, "Code Client", "[t_R_Social].[SocieteId] ")
        ReplaceChamp = Replace(ReplaceChamp, "Société", "fld3 ")
         ReplaceChamp = Replace(ReplaceChamp, "N° Commande", "T_Num_Commande.NumCommande ")
         ReplaceChamp = Replace(ReplaceChamp, "Montant initial TTC", "val(Montantinitial) ")
         ReplaceChamp = Replace(ReplaceChamp, "Cloturéée", "T_Num_Commande.CloturerCommande ")
         ReplaceChamp = Replace(ReplaceChamp, "Verrouillée", "T_Num_Commande.Verouiller ")
         ReplaceChamp = Replace(ReplaceChamp, "Date de Création", "format( T_Num_Commande.DateMiseEnService,'yyyy-mm-dd-hh:mm:ss')")
         '"Verouiller", "T_Num_Commande", "Verrouillée""
    Case "HISTORIQUE_AVENANT"
        ReplaceChamp = Replace(ReplaceChamp, "N° Commande", "NumCommande") '
        ReplaceChamp = Replace(ReplaceChamp, "Avenant", "NumAvoire")
        ReplaceChamp = Replace(ReplaceChamp, "Credit TTC", "Credit")
        ReplaceChamp = Replace(ReplaceChamp, "Debit TTC", "Debit")
        ReplaceChamp = Replace(ReplaceChamp, "Solde TTC", "Solde")
         ReplaceChamp = Replace(ReplaceChamp, "Cloturé", "T_Avoire.Cloturer")
         ReplaceChamp = Replace(ReplaceChamp, "Facturé", "T_Avoire.Facturer")
        '"Facturer", "T_Avoire", "Facturé"
     Case "VOIRPANIERDETAL"
        ReplaceChamp = Replace(ReplaceChamp, "Cadie", "T_NumCadi.NumCadi")
        ReplaceChamp = Replace(ReplaceChamp, "Mail", "Users_1.UserEMail")
        ReplaceChamp = Replace(ReplaceChamp, "Date de création", "T_NumCadi.DateCreation")
        ReplaceChamp = Replace(ReplaceChamp, "validé le", "T_NumCadi.DateMiseaDisposition")
        ReplaceChamp = Replace(ReplaceChamp, "LastName", "Users_1.LastName")
        ReplaceChamp = Replace(ReplaceChamp, "FirstName", "Users_1.FirstName")
        ReplaceChamp = Replace(ReplaceChamp, "UserLogin", "Users_1.UserLogin")
'        ReplaceChamp = Replace(ReplaceChamp, "UserLogin", "Users_1.UserLogin")
     '      Users_1.UserLogin,      ,  , Users_1.UserEMail AS Mail"

End Select


'
'If Trim("" & Table) = "" Then
'    SplitTxt = Split("" & txt & "." & ".")
'Else
'    SplitTxt = Split(txt & ".", ".")
'End If
''ReplaceChamp = txt '
'
'
'
'ReplaceChamp = Replace(ReplaceChamp, "Listerouge", "[Users].[Listerouge]")
'
'    ReplaceChamp = Replace(ReplaceChamp, "Cadie", "[NumCadi]")
'    ReplaceChamp = Replace(ReplaceChamp, "Nom", "[Users_1].[LastName]")
'    ReplaceChamp = Replace(ReplaceChamp, "LastName", "[Users_1].[LastName]")
'    ReplaceChamp = Replace(ReplaceChamp, "FirstName", "[Users_1].[FirstName]")
'    ReplaceChamp = Replace(ReplaceChamp, "UserLogin", "[Users_1].[UserLogin]")
'
''    ReplaceChamp = Replace(ReplaceChamp, "Mail", "[Users_1].[UserEMail]")
'
'    ReplaceChamp = Replace(ReplaceChamp, "DateCreation", "[T_NumCadi].[DateCreation]")
'    ReplaceChamp = Replace(ReplaceChamp, "Date de créatin", "[T_NumCadi].[DateCreation]")
'    ReplaceChamp = Replace(ReplaceChamp, "DateMiseaDisposition", "[T_NumCadi].[DateMiseaDisposition]")
'    ReplaceChamp = Replace(ReplaceChamp, "validé le", "[T_NumCadi].[DateMiseaDisposition]")
'    ReplaceChamp = Replace(ReplaceChamp, "Solder", "Sol§der")
'
'
'
'ReplaceChamp = Replace(ReplaceChamp, "Liste rouge", "Listerouge")
'
''ReplaceChamp = Replace(ReplaceChamp, "Cloturer", "CloturerCommande")
'ReplaceChamp = Replace(ReplaceChamp, "Clotur&eacute;e", "CloturerCommande")
'ReplaceChamp = Replace(ReplaceChamp, "V&eacute;rouill&eacute;e", "Verouiller")
'
'
'ReplaceChamp = Replace(ReplaceChamp, "Solde", "val([Credit]-[Debit])")
'ReplaceChamp = Replace(ReplaceChamp, "ModeTransport", "Mode de Transport")
'ReplaceChamp = Replace(ReplaceChamp, "Créer le", "T_Commande_Liv.creation")
'ReplaceChamp = Replace(ReplaceChamp, "N° Comm Encelade", "numDevis")
'ReplaceChamp = Replace(ReplaceChamp, "BL", "NumLiv")
''ReplaceChamp = Replace(ReplaceChamp, "[Ref Groupe Commande]", "Ref Groupe Commande")
'
'ReplaceChamp = Replace(ReplaceChamp, "Fournisseur", "CatName")
'
'ReplaceChamp = Replace(ReplaceChamp, "Créé le", "T_Commande_Liv.creation")
'ReplaceChamp = Replace(ReplaceChamp, "Date Création", "format( T_BR.DateCration,'yyyy-mm-dd-hh:mm:ss')")
'
'
'
'ReplaceChamp = Replace(ReplaceChamp, "Soldé", "[Solder]")
'If UCase(ReplaceChamp) = UCase("Cloturé") Then
'    ReplaceChamp = Replace(ReplaceChamp, "Cloturé", "[Cloturer]")
'Else
'    ReplaceChamp = Replace(ReplaceChamp, "Cloturée", "[CloturerCommande]")
'End If
'ReplaceChamp = Replace(ReplaceChamp, "Verouillée", "[Verouiller]")
'ReplaceChamp = Replace(ReplaceChamp, "Facturée", "[Facturer]")
''Verouiller as [Verouillé]
'ReplaceChamp = Replace(ReplaceChamp, "§", "")
'SplitTxt = Split((ReplaceChamp & "Desc"), ("Desc"))
'SplitTxt = Trim(SplitTxt(0))
'ReplaceChamp = Replace(ReplaceChamp, "" & SplitTxt, "" & SplitTxt & "")
''ReplaceChamp = Replace(ReplaceChamp, "[", "")
''ReplaceChamp = Replace(ReplaceChamp, "]", "")
'' "T_Commande_Liv.creation"
End Function
Public Function translate(my_msg)

   
    translate = my_msg
    Exit Function
    my_cid = Session("candidat_Application")
    my_language = Session("candidat_language")
    strDSN = Session("candidat_ADOContact")
   
    Set GeneralTools.My_Conn = OpenDb(strDSN)
    Set Rs = GeneralTools.My_Conn.Execute("select txt_trans from dbp_trans where cid=" & my_cid & " and language='" & my_language & "' and txt ='" & my_msg & "'")
    If Not Rs.EOF Then
        res = Rs("txt_trans")
    Else
        res = my_msg
    End If
    Rs.Close
    Set Rs = Nothing
    GeneralTools.My_Conn.Close
    Set GeneralTools.My_Conn = Nothing
    translate = res

End Function


Public Function GetDefault(fld, def, DSN)
def = safeEntry(def)
    Set GeneralTools.My_Conn2 = OpenDb(DSN)
    Set RS100 = GeneralTools.My_Conn2.Execute("SELECT * FROM Defaults WHERE defName = '" & fld & "'")
    If Not RS100.EOF Then
           GetDefault = Trim(RS100("defValue"))
    Else

        GeneralTools.My_Conn2.Execute ("INSERT INTO Defaults(defName,defValue) VALUES('" & fld & "','" & def & "')")
        GetDefault = def
    End If
    GeneralTools.My_Conn2.Close
     Set GeneralTools.My_Conn2 = Nothing
End Function
Public Function GetMessage(fld, def, DSN)
def = safeEntry(def)
    Set Conn = OpenDb(DSN)
    Set RS100 = Conn.Execute("SELECT * FROM T_Eboutique_Messages WHERE NameMessage = '" & fld & "'")
    If Not RS100.EOF Then
           GetMessage = RS100("LibMessage")
    Else

        Conn.Execute ("INSERT INTO T_Eboutique_Messages(NameMessage,LibMessage) VALUES('" & fld & "','" & def & "')")
        GetMessage = def
    End If
    Conn.Close
     Set Conn = Nothing
End Function


Public Function CaddyJava(Id_Caddie, ComFour, DSN As String)

CaddyJava = "<script language='javascript'>"
CaddyJava = CaddyJava & vbCrLf & "// texte contient le message à afficher par défaut"
CaddyJava = CaddyJava & vbCrLf & "var msgStatus=""Encelade"";"
CaddyJava = CaddyJava & vbCrLf & "top.defaultStatus=msgStatus;"
CaddyJava = CaddyJava & vbCrLf & ""
CaddyJava = CaddyJava & vbCrLf & "function message(txt) {"
 CaddyJava = CaddyJava & vbCrLf & "   top.status=txt;"
CaddyJava = CaddyJava & vbCrLf & "}"
CaddyJava = CaddyJava & vbCrLf & ""
CaddyJava = CaddyJava & vbCrLf & "function plus_un(id,IdM){"
CaddyJava = CaddyJava & vbCrLf & "   myForm = this.document.forms[0];"
'CaddyJava = CaddyJava & vbCrLf & "   var f = document.form_caddie;"
CaddyJava = CaddyJava & vbCrLf & "    var qte = 0;"
CaddyJava = CaddyJava & vbCrLf & "    eval (""qte = myForm.qte_""+id+"".value;"");"
'CaddyJava = CaddyJava & vbCrLf & "    qte = My_Qte;"
'CaddyJava = CaddyJava & vbCrLf & "   alert(""function plus_un "" + qte);"
'CaddyJava = CaddyJava & vbCrLf & "    if (Math.abs(parseInt(qte)) == qte)"
'CaddyJava = CaddyJava & vbCrLf & "    {"
CaddyJava = CaddyJava & vbCrLf & "        qte++;"
CaddyJava = CaddyJava & vbCrLf & "        myForm.qte_produit.value = qte;"
CaddyJava = CaddyJava & vbCrLf & "         myForm.id_produit.value = id;"
CaddyJava = CaddyJava & vbCrLf & "         myForm.ID_Menu.value = IdM;"
CaddyJava = CaddyJava & vbCrLf & "       myForm.act.value = ""modif"";"
CaddyJava = CaddyJava & vbCrLf & "        myForm.mode.value= ""CANDIDAT_CADDY"";"
CaddyJava = CaddyJava & vbCrLf & "        myForm.action = ""Contact.asp?mode=CANDIDAT_CADDY&ComFour=" & ComFour & "&Id_Caddie_0=" & Id_Caddie & """;"

'CaddyJava = CaddyJava & vbCrLf & "        myForm.action = ""Contact.asp?mode=CANDIDAT_CADDY"";"
CaddyJava = CaddyJava & vbCrLf & "        myForm.submit();"
'CaddyJava = CaddyJava & vbCrLf & "    }"
'CaddyJava = CaddyJava & vbCrLf & "   else"
'CaddyJava = CaddyJava & vbCrLf & "    {"
'CaddyJava = CaddyJava & vbCrLf & "        alert('" & GetMessage("QuantiteSaisieErronee", "La quantité saisie est erronée...", DsnTableMenu(DSN)) & "';"
'CaddyJava = CaddyJava & vbCrLf & "    }"
CaddyJava = CaddyJava & vbCrLf & "}"

CaddyJava = CaddyJava & vbCrLf & "function moins_un(id,IdM){"
CaddyJava = CaddyJava & vbCrLf & "   myForm = this.document.forms[0];"
'CaddyJava = CaddyJava & vbCrLf & "   var f = document.form_caddie;"
CaddyJava = CaddyJava & vbCrLf & "    var qte = 0;"
CaddyJava = CaddyJava & vbCrLf & "    eval (""qte = myForm.qte_""+id+"".value;"");"
'CaddyJava = CaddyJava & vbCrLf & "   alert(""function plus_un "" + qte);"
CaddyJava = CaddyJava & vbCrLf & "    if (qte>1)"
'CaddyJava = CaddyJava & vbCrLf & "    if (Math.abs(parseInt(qte)) == qte)"
CaddyJava = CaddyJava & vbCrLf & "    {"
CaddyJava = CaddyJava & vbCrLf & "        qte--;"
CaddyJava = CaddyJava & vbCrLf & "        myForm.qte_produit.value = qte;"
CaddyJava = CaddyJava & vbCrLf & "         myForm.id_produit.value = id;"
CaddyJava = CaddyJava & vbCrLf & "         myForm.ID_Menu.value = IdM;"
CaddyJava = CaddyJava & vbCrLf & "       myForm.act.value = ""modif"";"
CaddyJava = CaddyJava & vbCrLf & "        myForm.mode.value= ""CANDIDAT_CADDY"";"
CaddyJava = CaddyJava & vbCrLf & "        myForm.action = ""Contact.asp?mode=CANDIDAT_CADDY&ComFour=" & ComFour & "&Id_Caddie_0=" & Id_Caddie & """;"
'CaddyJava = CaddyJava & vbCrLf & "        myForm.action = ""Contact.asp?mode=CANDIDAT_CADDY"";"
CaddyJava = CaddyJava & vbCrLf & "        myForm.submit();"
CaddyJava = CaddyJava & vbCrLf & "    }"
CaddyJava = CaddyJava & vbCrLf & "   else"
CaddyJava = CaddyJava & vbCrLf & "    {"

CaddyJava = CaddyJava & vbCrLf & "        alert('" & GetMessage("QuantiteSaisieErronee", "La quantité saisie est erronée...", DsnTableMenu(DSN)) & "');"
CaddyJava = CaddyJava & vbCrLf & "    }"
CaddyJava = CaddyJava & vbCrLf & "}"



CaddyJava = CaddyJava & vbCrLf & "function modifie(id,IdM){"
CaddyJava = CaddyJava & vbCrLf & "   myForm = this.document.forms[0];"
'CaddyJava = CaddyJava & vbCrLf & "    var f = document.form_caddie;"
CaddyJava = CaddyJava & vbCrLf & "    var qte = 0;"
CaddyJava = CaddyJava & vbCrLf & "    eval (""qte = myForm.qte_""+id+"".value;"");"
CaddyJava = CaddyJava & vbCrLf & "    if (Math.abs(parseInt(qte)) == qte)"
CaddyJava = CaddyJava & vbCrLf & "    {"
CaddyJava = CaddyJava & vbCrLf & "        if (qte>0){"
CaddyJava = CaddyJava & vbCrLf & "            myForm.qte_produit.value = qte;"
CaddyJava = CaddyJava & vbCrLf & "         myForm.id_produit.value = id;"
CaddyJava = CaddyJava & vbCrLf & "         myForm.ID_Menu.value = IdM;"
CaddyJava = CaddyJava & vbCrLf & "            myForm.act.value=""modif"";"
CaddyJava = CaddyJava & vbCrLf & "        myForm.mode.value= ""CANDIDAT_CADDY"";"
CaddyJava = CaddyJava & vbCrLf & "        myForm.action = ""Contact.asp?mode=CANDIDAT_CADDY&ComFour=" & ComFour & "&Id_Caddie_0=" & Id_Caddie & """;"
CaddyJava = CaddyJava & vbCrLf & "            myForm.submit();"
CaddyJava = CaddyJava & vbCrLf & "        }"
CaddyJava = CaddyJava & vbCrLf & "        else"
CaddyJava = CaddyJava & vbCrLf & "        {"
CaddyJava = CaddyJava & vbCrLf & "            suppr(id);"
CaddyJava = CaddyJava & vbCrLf & "        }"
CaddyJava = CaddyJava & vbCrLf & "    }"
CaddyJava = CaddyJava & vbCrLf & "    else"
CaddyJava = CaddyJava & vbCrLf & "    {"
CaddyJava = CaddyJava & vbCrLf & "        alert('" & GetMessage("QuantiteSaisieErronee", "La quantité saisie est erronée...", DsnTableMenu(DSN)) & "');"
CaddyJava = CaddyJava & vbCrLf & "    }"
CaddyJava = CaddyJava & vbCrLf & "    "
CaddyJava = CaddyJava & vbCrLf & "}"




CaddyJava = CaddyJava & vbCrLf & "function suppr(id,IdM){"
CaddyJava = CaddyJava & vbCrLf & "    if (confirm('" & GetMessage("VouloirSupprimerArticle", "Êtes-vous sûr de vouloir supprimer cet article ?", DsnTableMenu(DSN)) & "'))"
CaddyJava = CaddyJava & vbCrLf & "    {"
CaddyJava = CaddyJava & vbCrLf & "   myForm = this.document.forms[0];"
'CaddyJava = CaddyJava & vbCrLf & "        var f = document.form_caddie;"
CaddyJava = CaddyJava & vbCrLf & "         myForm.id_produit.value = id;"
CaddyJava = CaddyJava & vbCrLf & "         myForm.ID_Menu.value = IdM;"
CaddyJava = CaddyJava & vbCrLf & "            myForm.act.value=""suppr"";"
CaddyJava = CaddyJava & vbCrLf & "        myForm.mode.value= ""CANDIDAT_CADDY"";"
CaddyJava = CaddyJava & vbCrLf & "        myForm.action = ""Contact.asp?mode=CANDIDAT_CADDY&ComFour=" & ComFour & "&Id_Caddie_0=" & Id_Caddie & """;"
CaddyJava = CaddyJava & vbCrLf & "            myForm.submit();"
CaddyJava = CaddyJava & vbCrLf & "    }"
CaddyJava = CaddyJava & vbCrLf & "}"
CaddyJava = CaddyJava & vbCrLf & "</script>"
End Function
Public Function ListeSherch(Conn As ADODB.Connection, Table As String, txtSelect, NameListe As String)
Dim Sql As String
Dim Rs As Recordset
Dim RsAttrib As Recordset
 Set RsAttrib = GeneralTools.My_Conn.Execute("SELECT con_FieldDefs.FieldName, con_FieldDefs.FieldAttribut From con_FieldDefs WHERE con_FieldDefs.FieldName='" & Table & "';")

If InStr(1, UCase("" & RsAttrib("FieldAttribut")), UCase("Num")) <> 0 Then
Sql = "SELECT DISTINCT val(" & Table & ".CatName) as txtCatName FROM " & Table & " ORDER BY " & " "

 Sql = Sql & "val(" & Table & ".CatName);"
Else
Sql = "SELECT DISTINCT " & Table & ".CatName as txtCatName FROM " & Table & " ORDER BY " & " "

 Sql = Sql & Table & ".CatName;"
End If
Set Rs = Conn.Execute(Sql)

ListeSherch = ""
ListeSherch = ListeSherch & vbCrLf & " <font class=""smallerheader"">  <SELECT NAME=""" & NameListe & """ >"
   While Rs.EOF = False
        If noquoteNumTxt(Rs("txtCatName")) = txtSelect Then
             ListeSherch = ListeSherch & vbCrLf & " <option value=""" & noquoteNumTxt(Rs("txtCatName")) & """ selected>" & noquoteNumTxt(Rs("txtCatName"))
        Else
            ListeSherch = ListeSherch & vbCrLf & " <option value=""" & noquoteNumTxt(Rs("txtCatName")) & """>" & noquoteNumTxt(Rs("txtCatName"))
        End If
        Rs.MoveNext
    Wend
     ListeSherch = ListeSherch & vbCrLf & "</SELECT></font>"
    RsAttrib.Close
    Set RsAttrib = Nothing
End Function
Public Function FomatNum(Value, NbDec) As Double
Value = "" & Value
Dim txt
Dim ModuloVal As Integer
Dim MulTipl As Integer
Dim ValMultiple As Long
On Error Resume Next
FomatNum = Round(val(Replace(Value, ",", ".")), NbDec)
'MulTipl = 1
''Calcul le multiplicateur 1 * 10 exposant NbDec
'For I = 1 To NbDec
'    MulTipl = MulTipl * 10
'Next
''Arrondit le chiffre value * MulTipl
'Err.Clear
'ValMultiple = val(Replace(Value, ",", ".")) * MulTipl
'If Err Then
'Err.Clear
'    FomatNum = Value
'Else
'    FomatNum = ValMultiple * (1 / MulTipl)
'End If
'Restitue le chiffre initial arrondit au nombre de décimaux.



End Function

Public Function ChampsNameAs(Champs As String, Table As String, ChampsReplace As String) As String
Dim txtReplaceChamp As String
If Trim("" & Table) <> "" Then
    txtReplaceChamp = " [" & Table & "].[" & Champs & "] AS [" & ChampsReplace & "] "
Else
    txtReplaceChamp = " [" & Table & "].[" & Champs & "] "
End If
ChampsNameAs = txtReplaceChamp
'ChampsNameAs = Replace(ChampsNameAs, "T_BR.Solder", "T_BR.Solder as [Soldé]")
'ChampsNameAs = Replace(ChampsNameAs, "creation", "creation as [Créé le]")
'ChampsNameAs = Replace(ChampsNameAs, "DateCration", "DateCration as [Date Création]")
'ChampsNameAs = Replace(ChampsNameAs, "T_Num_Commande.CloturerCommande", "T_Num_Commande.CloturerCommande as [Cloturée]")
'ChampsNameAs = Replace(ChampsNameAs, "T_Num_Commande.CloturerCommande", "T_Num_Commande.CloturerCommande as [Cloturée]")
'ChampsNameAs = Replace(ChampsNameAs, "T_Num_Commande.Verouiller", "T_Num_Commande.Verouiller as [Verouillée]")
'ChampsNameAs = Replace(ChampsNameAs, "T_Avoire.Cloturer", "T_Avoire.Cloturer as [Cloturée]")
'ChampsNameAs = Replace(ChampsNameAs, "T_Avoire.Facturer", "T_Avoire.Facturer as [Facturée]")
End Function
