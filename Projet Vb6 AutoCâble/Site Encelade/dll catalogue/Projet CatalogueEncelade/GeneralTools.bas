Attribute VB_Name = "GeneralTools"

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
           
            For i = 1 To 10
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
Public Function funReplace(A, b, C)
    funReplace = Replace(A, b, C)
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

          For i = 0 To UBound(bwords)
            str = Replace(str, bwords(i), String(Len(bwords(i)), "*"), 1, -1, 1)
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
    t = ""

    For i = 0 To UBound(tx)

      If LCase(oTag) = "[a]" Then
        p = InStr(1, tx(i), "[a]", 1)
        If p <> 0 Then
            tmp = Mid(tx(i), p)
            url = Mid(tmp, 4)
            If LCase(Left(url, 5)) = "http:" Then
                tmp1 = Replace(tmp, "[a]" & url, "<A href='" & url & "' Target=_Blank>Link</a>", 1, -1, 1)
            Else
                tmp1 = Replace(tmp, "[a]" & url, "<A href='http://" & url & "' Target=_Blank>Link</a>", 1, -1, 1)
            End If
            t = t & Replace(tx(i), tmp, tmp1)
        Else
            t = t & tx(i)
        End If
      Else
        cnt = InStr(1, tx(i), oTag, 1)
        Select Case cnt
            Case 0
                t = t & tx(i) & " "
            Case Else
                t = t & Replace(tx(i), oTag, roTag, 1, 1, 1)
                t = t & " " & rcTag & " "
        End Select
      End If
    Next
    doCode = t
End Function
Public Function CovertCommDate(MyNum As String) As String
Dim SplitNum
SplitNum = Split(MyNum & "__", "_")
CovertCommDate = Format(SplitNum(1), "dd/mm/yyyy")
End Function

Public Function OpenDb(MyDb)
Dim con As ADODB.Connection
Set con = New ADODB.Connection
    con.Mode = 16
    con.Open MyDb
Set OpenDb = con
End Function
Public Function ReplaceChamp(txt As String) As String
Dim Mytext As String
ReplaceChamp = txt
ReplaceChamp = Replace(ReplaceChamp, "Solder", "Sol§der")
ReplaceChamp = Replace(ReplaceChamp, "Date Commande", "[Date Commande]")
ReplaceChamp = Replace(ReplaceChamp, "Code Client", "NumIndiveClient")
ReplaceChamp = Replace(ReplaceChamp, "Société", "fld3")
ReplaceChamp = Replace(ReplaceChamp, "Cp", "Zip")
ReplaceChamp = Replace(ReplaceChamp, "Ville", "City")
ReplaceChamp = Replace(ReplaceChamp, "Pays", "label_pays")
ReplaceChamp = Replace(ReplaceChamp, "Liste rouge", "Listerouge")
ReplaceChamp = Replace(ReplaceChamp, "N° Commande", "NumCommande") '
ReplaceChamp = Replace(ReplaceChamp, "Montant initial", "val(Montantinitial)")
ReplaceChamp = Replace(ReplaceChamp, "Cloturer", "CloturerCommande")
ReplaceChamp = Replace(ReplaceChamp, "Clotur&eacute;e", "CloturerCommande")
ReplaceChamp = Replace(ReplaceChamp, "V&eacute;rouill&eacute;e", "Verouiller")
ReplaceChamp = Replace(ReplaceChamp, "Date de Création", "format( T_Num_Commande.DateMiseEnService,'yyyy-mm-dd-hh:mm:ss')")
ReplaceChamp = Replace(ReplaceChamp, "Avenant", "NumAvoire")
ReplaceChamp = Replace(ReplaceChamp, "Solde", "val([Credit]-[Debit])")
ReplaceChamp = Replace(ReplaceChamp, "ModeTransport", "Mode de Transport")
ReplaceChamp = Replace(ReplaceChamp, "Créer le", "T_Commande_Liv.creation")
ReplaceChamp = Replace(ReplaceChamp, "N° Comm Encelade", "numDevis")
ReplaceChamp = Replace(ReplaceChamp, "BL", "NumLiv")
ReplaceChamp = Replace(ReplaceChamp, "[Ref Groupe Commande]", "Ref Groupe Commande")
ReplaceChamp = Replace(ReplaceChamp, "Ref Groupe Commande", "[Ref Groupe Commande]")
ReplaceChamp = Replace(ReplaceChamp, "Fournisseur", "CatName")
ReplaceChamp = Replace(ReplaceChamp, "§", "")
' "T_Commande_Liv.creation"
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

Public Function CaddyJava()

CaddyJava = ""
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
CaddyJava = CaddyJava & vbCrLf & "        myForm.action = ""Contact.asp?mode=CANDIDAT_CADDY"";"
'CaddyJava = CaddyJava & vbCrLf & "        myForm.action = ""Contact.asp?mode=CANDIDAT_CADDY"";"
CaddyJava = CaddyJava & vbCrLf & "        myForm.submit();"
'CaddyJava = CaddyJava & vbCrLf & "    }"
'CaddyJava = CaddyJava & vbCrLf & "   else"
'CaddyJava = CaddyJava & vbCrLf & "    {"
'CaddyJava = CaddyJava & vbCrLf & "        alert (""La quantité saisie est erronée..."");"
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
CaddyJava = CaddyJava & vbCrLf & "        myForm.action = ""Contact.asp?mode=CANDIDAT_CADDY"";"
'CaddyJava = CaddyJava & vbCrLf & "        myForm.action = ""Contact.asp?mode=CANDIDAT_CADDY"";"
CaddyJava = CaddyJava & vbCrLf & "        myForm.submit();"
CaddyJava = CaddyJava & vbCrLf & "    }"
CaddyJava = CaddyJava & vbCrLf & "   else"
CaddyJava = CaddyJava & vbCrLf & "    {"
CaddyJava = CaddyJava & vbCrLf & "        alert (""La quantité saisie est erronée..."");"
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
CaddyJava = CaddyJava & vbCrLf & "        myForm.action = ""Contact.asp?mode=CANDIDAT_CADDY"";"
CaddyJava = CaddyJava & vbCrLf & "            myForm.submit();"
CaddyJava = CaddyJava & vbCrLf & "        }"
CaddyJava = CaddyJava & vbCrLf & "        else"
CaddyJava = CaddyJava & vbCrLf & "        {"
CaddyJava = CaddyJava & vbCrLf & "            suppr(id);"
CaddyJava = CaddyJava & vbCrLf & "        }"
CaddyJava = CaddyJava & vbCrLf & "    }"
CaddyJava = CaddyJava & vbCrLf & "    else"
CaddyJava = CaddyJava & vbCrLf & "    {"
CaddyJava = CaddyJava & vbCrLf & "        alert (""La quantité saisie est erronée..."");"
CaddyJava = CaddyJava & vbCrLf & "    }"
CaddyJava = CaddyJava & vbCrLf & "    "
CaddyJava = CaddyJava & vbCrLf & "}"




CaddyJava = CaddyJava & vbCrLf & "function suppr(id,IdM){"
CaddyJava = CaddyJava & vbCrLf & "    if (confirm(""Êtes-vous sûr de vouloir supprimer cet article ?""))"
CaddyJava = CaddyJava & vbCrLf & "    {"
CaddyJava = CaddyJava & vbCrLf & "   myForm = this.document.forms[0];"
'CaddyJava = CaddyJava & vbCrLf & "        var f = document.form_caddie;"
CaddyJava = CaddyJava & vbCrLf & "         myForm.id_produit.value = id;"
CaddyJava = CaddyJava & vbCrLf & "         myForm.ID_Menu.value = IdM;"
CaddyJava = CaddyJava & vbCrLf & "            myForm.act.value=""suppr"";"
CaddyJava = CaddyJava & vbCrLf & "        myForm.mode.value= ""CANDIDAT_CADDY"";"
CaddyJava = CaddyJava & vbCrLf & "        myForm.action = ""Contact.asp?mode=CANDIDAT_CADDY"";"
CaddyJava = CaddyJava & vbCrLf & "            myForm.submit();"
CaddyJava = CaddyJava & vbCrLf & "    }"
CaddyJava = CaddyJava & vbCrLf & "}"

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
MulTipl = 1
'Calcul le multiplicateur 1 * 10 exposant NbDec
For i = 1 To NbDec
    MulTipl = MulTipl * 10
Next
'Arrondit le chiffre value * MulTipl
ValMultiple = val(Replace(Value, ",", ".")) * MulTipl
If Err Then
Err.Clear
    FomatNum = Value
Else
    FomatNum = ValMultiple * (1 / MulTipl)
End If
'Restitue le chiffre initial arrondit au nombre de décimaux.



End Function


