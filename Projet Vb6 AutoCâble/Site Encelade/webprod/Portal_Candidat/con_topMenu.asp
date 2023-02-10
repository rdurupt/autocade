<% Response.Expires = 0 %>
<!--#INCLUDE FILE="ADOConnect.asp"-->
<!--#INCLUDE FILE="MenuTools.asp"-->
<%

set Rs =server.CreateObject("ADODB.Recordset")
set Myconn = Server.CreateObject("ADODB.Connection")
Myconn.Open session("candidat_ADOContact")
set rs=Myconn.Execute("SELECT BaseDefault.Path FROM BaseDefault;")

if Rs.EOF=false then

	DSN	="DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Rs("Path")

end if
Myconn.close
set Myconn = Nothing
set Myconn = Server.CreateObject("ADODB.Connection")

Myconn.Open DSN
set Rs= Nothing

Public Function GetChildrenNb(ByRef tempSort, leCatId, posPar)
    leCompteur = 0
    For i = 0 To UBound(tempSort, 2)
        If CInt(tempSort(posPar, i)) = CInt(leCatId) Then
            leCompteur = leCompteur + 1
        End If
    Next
    GetChildrenNb = leCompteur
End Function

Public Function GetMenuArray(tempSort, posId, posPar, posCaption, racine)
    indx = 1
    nbC = UBound(tempSort, 1) + 3
    nbR = UBound(tempSort, 2)
    ReDim outArray(nbC, nbR)
    For i = 0 To UBound(tempSort, 2)
        For j = 0 To UBound(tempSort, 1)
            outArray(j, i) = tempSort(j, i)
            outArray(nbC, i) = 0
        Next
    Next
    outArray(nbC - 2, 0) = "Menu" & indx
    indx = indx + 1
    outArray(nbC - 1, 0) = GetChildrenNb(tempSort, tempSort(posId, 0), posPar)
    outArray(nbC, 0) = Len(tempSort(posCaption, 0))
    For i = 1 To UBound(outArray, 2)
        If CInt(tempSort(posPar, i)) = CInt(racine) Then
            outArray(nbC - 2, i) = "Menu" & indx
            indx = indx + 1
            outArray(nbC - 1, i) = GetChildrenNb(tempSort, tempSort(posId, i), posPar)
            outArray(nbC, i) = Len(tempSort(posCaption, i))
        Else
            leStrEl = ""
            For j = 0 To UBound(outArray, 2)
                If CInt(tempSort(posId, j)) = CInt(tempSort(posPar, i)) Then
                    leStrEl = outArray(nbC - 2, j)
                    leParent = CInt(tempSort(posId, j))
                    Exit For
                End If
            Next
            leNiveau = 0
            MaxLength = 0
            For j = 0 To UBound(tempSort, 2)
                If CInt(tempSort(posPar, j)) = leParent Then
                    If CInt(outArray(nbC, j)) > MaxLength Then
                        MaxLength = outArray(nbC, j)
                    End If
                    leNiveau = leNiveau + 1
                    If MaxLength < CInt(Len(tempSort(posCaption, i))) Then
                        outArray(nbC, j) = Len(tempSort(posCaption, i))
                        MaxLength = Len(tempSort(posCaption, i))
                    End If
                End If
                If CInt(tempSort(posId, j)) = CInt(tempSort(posId, i)) Then
                    Exit For
                End If
            Next
            outArray(nbC - 2, i) = leStrEl & "_" & leNiveau
            outArray(nbC - 1, i) = GetChildrenNb(tempSort, tempSort(posId, i), posPar)
            outArray(nbC, i) = MaxLength
        End If
    Next
    GetMenuArray = outArray
End Function

Function GetDefault(fld,def)

    Set ConnSP = Server.CreateObject("ADODB.Connection")
	ConnSP.Open session("candidat_ADOContact")

    Set RS100 = ConnSP.Execute("SELECT * FROM Defaults WHERE defName = '" & fld & "'")
    If Not RS100.EOF Then
        GetDefault = trim(RS100("defValue"))
	else
		ConnSP.Execute("INSERT INTO Defaults(defName,defValue) VALUES('" & fld & "','" & def & "')")
		GetDefault = def
    End If
    ConnSP.Close
        Set ConnSP = Nothing
      Set RS100 = Nothing
End Function


If Request("Mode")<>"con_frmcandidat" then
	response.write ("<body background=""background.asp"">")
	response.write ("<br><br><br><br>")
End if

nbFirstLine = 1
indx = 0
ReDim tempSort(3, 0)
tempSort(0, indx) = 1000
tempSort(1, indx) = -1
tempSort(2, indx) = "Type de pièces"
tempSort(3, indx) = ""
indx = indx + 1
Set RS1 = Myconn.Execute("SELECT * FROM menu  WHERE menu.PasVisible=false ORDER BY libelle")
do while not RS1.EOF 
	ReDim Preserve tempSort(3, indx)		
	tempSort(0, indx) = rs1("CatId")
	tempSort(1, indx) = 1000
	tempSort(2, indx) = RS1("libelle")
	tempSort(3, indx) = "switchbase.asp?CatId=" & rs1("CatId")
	indx = indx + 1
	RS1.movenext
loop

ReDim Preserve tempSort(3, indx)		
tempSort(0, indx) = 2000
tempSort(1, indx) = -1
tempSort(2, indx) = "Outils"
tempSort(3, indx) = ""
indx = indx + 1
'nbFirstLine = nbFirstLine + 1
ReDim Preserve tempSort(3, indx)		
tempSort(0, indx) = 2001
tempSort(1, indx) = 2000
tempSort(2, indx) = "Recherche"
tempSort(3, indx) = "Contact.asp?mode=search"
indx = indx + 1
nbFirstLine = nbFirstLine + 1
ReDim Preserve tempSort(3, indx)		
	tempSort(0, indx) = 2002
	tempSort(1, indx) =2000
	tempSort(2, indx) = "Histrorique Client"
	tempSort(3, indx) = "javascript:parent.frames['main'].location='Contact.asp?mode=EspaceClient&NumFrm=1'"


if session("candidat_UserType") = "Administrator" then	
indx = indx + 1
	nbFirstLine = nbFirstLine + 1
	ReDim Preserve tempSort(3, indx)		
	tempSort(0, indx) = 3000
	tempSort(1, indx) = -1
	tempSort(2, indx) = "Administration"
	tempSort(3, indx) = ""
	indx = indx + 1
	ReDim Preserve tempSort(3, indx)		
	tempSort(0, indx) = 3001
	tempSort(1, indx) = 3000
	tempSort(2, indx) = getdefault("tables","Liste des tables")
	tempSort(3, indx) = "javascript:parent.frames['main'].location='con_lstCategory.asp'"
	indx = indx + 1
	ReDim Preserve tempSort(3, indx)		
	tempSort(0, indx) = 3002
	tempSort(1, indx) = 3000
	tempSort(2, indx) = "Paramétrages"
	tempSort(3, indx) = "javascript:parent.frames['main'].location='Contact.asp?mode=con_frmSetting'"
	indx = indx + 1
	ReDim Preserve tempSort(3, indx)		
	tempSort(0, indx) = 3003
	tempSort(1, indx) = 3000
	tempSort(2, indx) = "Configuration"
	tempSort(3, indx) = "javascript:parent.frames['main'].location='Contact.asp?mode=Config'"
	indx = indx + 1
	ReDim Preserve tempSort(3, indx)		
	tempSort(0, indx) = 3004
	tempSort(1, indx) = 3000
	tempSort(2, indx) = "Gestion Client"
	tempSort(3, indx) = "javascript:parent.frames['main'].location='Contact.asp?mode=ConfigCLI'"
	indx = indx + 1
	ReDim Preserve tempSort(3, indx)		
	tempSort(0, indx) = 3005
	tempSort(1, indx) = 3000
	tempSort(2, indx) = "Historiques"
	tempSort(3, indx) = "javascript:parent.frames['main'].location='Contact.asp?mode=Historiques'"
	

End if

	


leStr = leStr & "<script language=""JavaScript"">" & vbCrLf
leStr = leStr & "var NoOffFirstLineMenus = " & nbFirstLine & ";" & vbCrLf
leStr = leStr & "var BaseHref = """";" & vbCrLf
tempARR = GetMenuArray(tempSort, 0, 1, 2, -1)
For i = 0 To UBound(tempARR, 2)
	leLink = tempARR(3, i)
    leWidth = Fix((tempARR(6, i) * 12) / 1.5) + 20
    leStr = leStr & tempARR(4, i) & " = new Array(""" & tempARR(2, i) & """, """ & leLink & """, """", " & tempARR(5, i) & ", 18, " & leWidth & ", """", """", """", """", """", """", -1, -1, -1, """", """ & Replace(tempARR(2, i), "'", "\'") & """);" & vbCrLf
Next
leStr = leStr & "</script>" & vbCrLf
leStr = leStr & "<script language=""JavaScript"">function Go(){return}</script>" & vbCrLf
leStr = leStr & "<script language=""JavaScript"" src=""Portal_Menu_Format.js""></script>" & vbCrLf
leStr = leStr & "<script language=""JavaScript"" src=""Portal_Menu.js""></script>" & vbCrLf
myconn.close
set Myconn = Nothing

Response.Write leStr

%>