VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Logistique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Application  As Application
Public Request As Request
Public Session As Session
Public Response As Response
Public SERVER As SERVER
Public FormCaption As String
Public FlagHTMLEditor As Boolean
 Public MyConnecSting As String
 Dim con
Dim ConexString As String
Function GenererBlBr()
Dim Bl As String
Dim Br As String
Response.Write "<div align='center'>"
 Response.Write "<table border='1'><tr>"
    Response.Write "<td bgcolor='#CCCCCC'align='center'>Affaire</td>"
    Response.Write "<td bgcolor='#CCCCCC'align='center'>Projet</td>"
    Response.Write "<td bgcolor='#CCCCCC'align='center'>Véhicule</td>"
    Response.Write "<td bgcolor='#CCCCCC'align='center'>Code S/Ensemble</td>"
    Response.Write "<td bgcolor='#CCCCCC'align='center'>Pièce</td>"
    Response.Write "<td bgcolor='#CCCCCC'align='center'>Indice</td>"
    Response.Write "<td bgcolor='#CCCCCC'align='center'>BL</td>"
    Response.Write "<td bgcolor='#CCCCCC'align='center'>BR</td>"
     Response.Write "</tr>"

'Recherche tous les N° Affaire lier à une création de BL
'*******************************************************
Sql = "SELECT  T_Projects.Id,T_Projects.NumAffaire "
Sql = Sql & "FROM T_Nomenclature INNER JOIN "
Sql = Sql & "T_SousEnsembles ON T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] =  "
Sql = Sql & "T_SousEnsembles.Id INNER JOIN "
Sql = Sql & "t_Vehicule ON T_SousEnsembles.ID_Vehicule_Véhicule = t_Vehicule.Id INNER JOIN "
Sql = Sql & "T_Projects ON t_Vehicule.Id_Projects_NumAffaire = T_Projects.Id "
Sql = Sql & "WHERE     T_Nomenclature.Num_de_retour = N'########' "
Sql = Sql & "GROUP BY T_Projects.Id,T_Projects.NumAffaire"
Set Rs = con.Execute(Sql)
While Rs.EOF = False
'**********************************************************************************************
'Affecte un N° de BL
   Bl = NumeroChrono("NumJob", "")
   Bl = "BL_" & Rs("NumAffaire") & "_" & Bl
'*********************************************************************************************
'MAJ N° de BL dans T_Nomenclature
   Sql = "UPDATE T_Nomenclature "
Sql = Sql & "SET Num_de_retour = '" & Bl & "' "
Sql = Sql & "FROM T_Projects INNER JOIN "
Sql = Sql & "t_Vehicule ON T_Projects.Id = t_Vehicule.Id_Projects_NumAffaire INNER JOIN "
Sql = Sql & "T_SousEnsembles ON t_Vehicule.Id = T_SousEnsembles.ID_Vehicule_Véhicule INNER JOIN "
Sql = Sql & "T_Nomenclature ON T_SousEnsembles.Id = T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] "
Sql = Sql & "Where T_Nomenclature.Num_de_retour = N'########' AND T_Projects.Id = " & Rs("Id") & " "
con.Execute (Sql)
'***************************************************************************************************
'Afficher les modifications
'*********************************************************************************************
Sql = "SELECT T_Projects.NumAffaire, T_Projects.Projet, t_Vehicule.Véhicule, "
Sql = Sql & "T_SousEnsembles.Code, T_Nomenclature.Pièce, "
Sql = Sql & "T_Nomenclature.Indice , T_Nomenclature.Num_de_retour "
Sql = Sql & "FROM T_Nomenclature INNER JOIN "
Sql = Sql & "T_SousEnsembles ON T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] =  "
Sql = Sql & "T_SousEnsembles.Id INNER JOIN "
Sql = Sql & "t_Vehicule ON T_SousEnsembles.ID_Vehicule_Véhicule = t_Vehicule.Id INNER JOIN "
Sql = Sql & "T_Projects ON t_Vehicule.Id_Projects_NumAffaire = T_Projects.Id "
Sql = Sql & "WHERE  T_Nomenclature.Num_de_retour = '" & Bl & "'"
Set RsBl = con.Execute(Sql)
While RsBl.EOF = False
    Response.Write "<tr>"
    Response.Write "<td>" & RsBl("NumAffaire") & "</td>"
    Response.Write "<td>" & RsBl("Projet") & "</td>"
    Response.Write "<td>" & RsBl("Véhicule") & "</td>"
    Response.Write "<td>" & RsBl("Code") & "</td>"
    Response.Write "<td>" & RsBl("Pièce") & "</td>"
    Response.Write "<td>" & RsBl("Indice") & "</td>"
    Response.Write "<td>" & RsBl("Num_de_retour") & "</td>"
    Response.Write "<td>&nbsp;</td>"
     Response.Write "</tr>"
    RsBl.MoveNext
Wend
RsBl.Close
Set RsBl = Nothing

'***************************************************************************************************
    Rs.MoveNext
Wend
'Recherche tous les N° Affaire lier à une création de BR
'*******************************************************
Sql = "SELECT     T_Projects.Id,T_Projects.NumAffaire "
Sql = Sql & "FROM         T_Nomenclature INNER JOIN "
Sql = Sql & "T_SousEnsembles ON T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] = "
Sql = Sql & "T_SousEnsembles.Id INNER JOIN "
Sql = Sql & "t_Vehicule ON T_SousEnsembles.ID_Vehicule_Véhicule = t_Vehicule.Id INNER JOIN "
Sql = Sql & "T_Projects ON t_Vehicule.Id_Projects_NumAffaire = T_Projects.Id "
Sql = Sql & "WHERE T_Nomenclature.N_de_BL = N'########' "
Sql = Sql & "GROUP BY T_Projects.Id,T_Projects.NumAffaire "


Set Rs = con.Execute(Sql)
While Rs.EOF = False
'**********************************************************************************************
'Affecte un N° de BR
   Br = NumeroChrono("NumJob", "")
   Br = "BR_" & Rs("NumAffaire") & "_" & Br
'*********************************************************************************************
'MAJ N° de BL dans T_Nomenclature
   Sql = "UPDATE T_Nomenclature "
Sql = Sql & "SET N_de_BL = '" & Br & "' "
Sql = Sql & "FROM T_Projects INNER JOIN "
Sql = Sql & "t_Vehicule ON T_Projects.Id = t_Vehicule.Id_Projects_NumAffaire INNER JOIN "
Sql = Sql & "T_SousEnsembles ON t_Vehicule.Id = T_SousEnsembles.ID_Vehicule_Véhicule INNER JOIN "
Sql = Sql & "T_Nomenclature ON T_SousEnsembles.Id = T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] "
Sql = Sql & "Where T_Nomenclature.N_de_BL = N'########' AND T_Projects.Id = " & Rs("Id") & " "
con.Execute (Sql)
'***********************************************************************************************
'Afficher les modifications
'*********************************************************************************************
Sql = "SELECT T_Projects.NumAffaire, T_Projects.Projet, t_Vehicule.Véhicule, "
Sql = Sql & "T_SousEnsembles.Code, T_Nomenclature.Pièce, "
Sql = Sql & "T_Nomenclature.Indice , T_Nomenclature.N_de_BL  "
Sql = Sql & "FROM T_Nomenclature INNER JOIN "
Sql = Sql & "T_SousEnsembles ON T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] =  "
Sql = Sql & "T_SousEnsembles.Id INNER JOIN "
Sql = Sql & "t_Vehicule ON T_SousEnsembles.ID_Vehicule_Véhicule = t_Vehicule.Id INNER JOIN "
Sql = Sql & "T_Projects ON t_Vehicule.Id_Projects_NumAffaire = T_Projects.Id "
Sql = Sql & "WHERE  T_Nomenclature.N_de_BL  = '" & Br & "'"
Set RsBl = con.Execute(Sql)
While RsBl.EOF = False
    Response.Write "<tr>"
    Response.Write "<td>" & RsBl("NumAffaire") & "</td>"
    Response.Write "<td>" & RsBl("Projet") & "</td>"
    Response.Write "<td>" & RsBl("Véhicule") & "</td>"
    Response.Write "<td>" & RsBl("Code") & "</td>"
    Response.Write "<td>" & RsBl("Pièce") & "</td>"
    Response.Write "<td>" & RsBl("Indice") & "</td>"
    Response.Write "<td>&nbsp;</td>"
    Response.Write "<td>" & RsBl("N_de_BL") & "</td>"
    Response.Write "</tr>"
    RsBl.MoveNext
Wend
RsBl.Close
Set RsBl = Nothing

'***************************************************************************************************
    Rs.MoveNext
Wend
Response.Write "</table></div>"
End Function
Function OpenDbN(TYPEBASE, BASE, SERVER, Fichier, USER, PASSWORD)
'NewSettion
'1 - ODBC
'2 - ORACLE
'3 - ACCESS 2000
'4 - ACCESS 97
'5 - SQL SERVER
'6 - SQL SERVER 2005 Express
Dim I As Integer
Dim GenereCSTRING As String
Select Case TYPEBASE
    Case 1
        GenereCSTRING = "Provider=MSDASQL.1;Password=" & PASSWORD & ";Persist Security Info=True;User ID=" & USER & ";Data Source=" & BASE
    Case 2
        GenereCSTRING = "Provider=OraOLEDB.Oracle.1;Password=" & PASSWORD & ";Persist Security Info=True;User ID=" & USER & ";Data Source=" & BASE
   
    Case 3
        GenereCSTRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Fichier & ";Persist Security Info=False"
    Case 4
        GenereCSTRING = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & Fichier
    Case 5
        GenereCSTRING = "Provider=SQLOLEDB.1;Password=" & PASSWORD & ";Persist Security Info=True;User ID=" & USER & ";Initial Catalog=" & BASE & ";Data Source=" & SERVER
   Case 6
    GenereCSTRING = "Provider=SQLNCLI.1;Password=" & PASSWORD & ";Persist Security Info=True;User ID=" & USER & ";Initial Catalog=" & BASE & ";Data Source=" & SERVER
     Case 7
         GenereCSTRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Fichier & ";Extended Properties=""Excel 8.0;"";"

    Case Else
        GenereCSTRING = ""

End Select
  On Error Resume Next
    NotErreur = False
Reprise:
    If Trim("" & GenereCSTRING) <> "" Then
    Set OpenDbN = CreateObject("ADODB.Connection")
    OpenDbN.Open (GenereCSTRING)
    If Err = 0 Then
        NotErreur = True
    Else
      I = I + 1
      If I < 50 Then
      Err.Clear
      GoTo Reprise
      End If
    End If
    Err.Clear
  End If
 On Error GoTo 0
'Response.Write GenereCSTRING
'Response.End
End Function

Sub NewSettion()
Dim Sql As String
Dim Rs
blnResult = RevertToSelf()
Sql = "SELECT T_Admin.UnserName,  T_Admin.PassWord,   T_Admin.Domain "
Sql = Sql & "FROM  T_Admin;"
Set Rs = con.Execute(Sql)
    lpUsername = "" & Rs!UnserName
    lpDomain = "" & Rs!Domain
    lpPassword = "" & Rs!PASSWORD
If LogonUser( _
lpUsername, _
lpDomain, _
lpPassword, _
          LOGON32_LOGON_INTERACTIVE, _
         LOGON32_PROVIDER_DEFAULT, _
            lngTokenHandle) = 0 Then
    MsgBox "Impossible d'ouvrir la session : " & lpUsername & ". "
    GoTo Fin
End If


If blnResult = False Then
    MsgBox "Impossible d'ouvrir LogonUser"
   GoTo Fin
End If
'MsgBox "Session avec le jeton" & lngTokenHandle & " et " & strAdminUser & ", " & strAdminDomain & ", " & strAdminPassword & " ouverte !"


blnResult = ImpersonateLoggedOnUser(lngTokenHandle)
Fin:
End Sub
Function OpenRecordSet(Sql, Conn)
    On Error Resume Next
    NotErreur = False
    Err.Clear
    Debug.Print Sql
    'Response.Write sql
    'Response.End
      Set OpenRecordSet = Conn.Execute(Sql)
  If Err = 0 Then
        NotErreur = True
    End If
    Err.Clear
 On Error GoTo 0
End Function


Function CloseRecordSet(Rs)
 On Error Resume Next
   Rs.Close
      Set CloseRecordSet = Nothing
End Function
Function CloseDb(Conn)
    Conn.Close
     Set Conn = Nothing
End Function

Sub ExecuteSql(Sql, Conn)
   On Error Resume Next
    NotErreur = False
    Err.Clear
    Debug.Print Sql
    'Response.Write sql
   ' Response.End
    Conn.Execute (Sql)
   If Err = 0 Then
        NotErreur = True
    End If
    Err.Clear
 On Error GoTo 0
End Sub


 Function RecapSqlNomenclatuer(FormName, champMasquer, ChDefautValeur, ValeurDefaule, Updatable)
RecapSqlNomenclatuer = RecapSqlNomenclatuer & vbCrLf & "<SCRIPT langage='JAVASCRIPT'>"
    RecapSqlNomenclatuer = RecapSqlNomenclatuer & vbCrLf & "function PageSuvante(Logistique_Page,id,id_Vehicule,Pièce,ID__SousEnsembles,Id_Projects_NumAffaire){"

    RecapSqlNomenclatuer = RecapSqlNomenclatuer & vbCrLf & "document." & FormName & ".Logistique_Page.value=Logistique_Page;"
    RecapSqlNomenclatuer = RecapSqlNomenclatuer & vbCrLf & "document." & FormName & ".id.value=id;"
    RecapSqlNomenclatuer = RecapSqlNomenclatuer & vbCrLf & "document." & FormName & ".id_Vehicule.value=id_Vehicule;"
    RecapSqlNomenclatuer = RecapSqlNomenclatuer & vbCrLf & "document." & FormName & ".Pièce.value=Pièce;"
    RecapSqlNomenclatuer = RecapSqlNomenclatuer & vbCrLf & "document." & FormName & ".ID__SousEnsembles.value=ID__SousEnsembles;"
    RecapSqlNomenclatuer = RecapSqlNomenclatuer & vbCrLf & "document." & FormName & ".Id_Projects_NumAffaire.value=Id_Projects_NumAffaire;"
'    Id_Projects_NumAffaire

    RecapSqlNomenclatuer = RecapSqlNomenclatuer & vbCrLf & "document." & FormName & ".action='Default.asp';"
    RecapSqlNomenclatuer = RecapSqlNomenclatuer & vbCrLf & "document." & FormName & ".submit();"



    RecapSqlNomenclatuer = RecapSqlNomenclatuer & vbCrLf & "}"
    RecapSqlNomenclatuer = RecapSqlNomenclatuer & vbCrLf & "</SCRIPT>"
    
    RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<form name='" & FormName & "' method='post'>" & vbCrLf
     
     RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<input name='Updatable' type='hidden' value='OK'/>" & vbCrLf
      RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<input name='FormName' type='hidden' value='" & FormName & "'/>" & vbCrLf
         RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<input name='id' type='hidden' value='" & Request("id") & "' />"
       RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<input name='id_Vehicule' type='hidden' value='" & Request("id_Vehicule") & "' />"
       RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<input name='Pièce' type='hidden' value='" & Request("Pièce") & "' />"
       RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<input name='ID__SousEnsembles' type='hidden' value='" & Request("ID__SousEnsembles") & "' />"

    RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<input name='Id_Projects_NumAffaire' type='hidden' value='" & Request("Id_Projects_NumAffaire") & "' />"
     RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<input name='Logistique_Page' type='hidden' value='4'/>" & vbCrLf
     
    
      
     
   
    
   

        
        Sql = "SELECT T_Projects.Projet, T_SousEnsembles.ID_Vehicule_Véhicule, t_Vehicule.Véhicule, T_SousEnsembles.[Code], T_Nomenclature.Designation_Piece,  "
        Sql = Sql & "T_Nomenclature.Pièce, T_Nomenclature.Quantite, T_Nomenclature.Reste_a_recevoir AS QtsManqaute,  "
        Sql = Sql & "T_Actions.Libellé, T_Delais.Libellé AS Delais "
        Sql = Sql & "FROM         t_Vehicule INNER JOIN "
        Sql = Sql & "T_SousEnsembles INNER JOIN "
        Sql = Sql & "T_Nomenclature LEFT OUTER JOIN "
        Sql = Sql & "T_Actions ON T_Nomenclature.Id_T_Actions_Libellé = T_Actions.Id ON "
        Sql = Sql & "T_SousEnsembles.Id = T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] ON "
        Sql = Sql & "t_Vehicule.Id = T_SousEnsembles.ID_Vehicule_Véhicule INNER JOIN "
        Sql = Sql & "T_Projects ON t_Vehicule.Id_Projects_NumAffaire = T_Projects.Id LEFT OUTER JOIN "
        Sql = Sql & "T_Delais ON T_Nomenclature.Id_T_Delais_Libellé = T_Delais.Id "
        
 
        
           MyFilter = RetourneCloseWhere()
          If Request("nbSelect") > 0 Then
            Sql = Sql & MyCloseWhere()
             If Trim("" & MyFilter) <> "" Then
                Sql = Sql & " AND (" & MyFilter & ") "
             End If
          Else
            If Trim("" & MyFilter) <> "" Then
                Sql = Sql & " WHERE " & MyFilter
             End If
        End If
        If Trim("" & Request("RecapFiltreVehicule")) <> "" Then
                     If InStr(1, UCase(Sql), "WHERE") <> 0 Then
                    
                Sql = Sql & " AND t_Vehicule.Id=" & Request("RecapFiltreVehicule") & " "
                
            Else
                MyFilter = "?"
                 Sql = Sql & " WHERE t_Vehicule.Id=" & Request("RecapFiltreVehicule") & " "
            End If
      End If
    
      
       Sql = Sql & "GROUP BY T_Projects.Projet, T_SousEnsembles.ID_Vehicule_Véhicule, t_Vehicule.Véhicule, T_SousEnsembles.[Code], T_Nomenclature.Designation_Piece,  "
        Sql = Sql & "T_Nomenclature.Pièce, T_Nomenclature.Quantite, T_Nomenclature.Reste_a_recevoir ,T_Actions.Libellé, T_Delais.Libellé, T_Nomenclature.Modifiee_le "
        Sql = Sql & "HAVING T_Nomenclature.Modifiee_le Is Null;"


                


 
       'Response.Write sql
        'response.End
        
        'sql=sql & "WHERE T_Nomenclature.Id_SousEnsemble=" & rs("id") & ";"
      
        Set RsNomenclature = OpenRecordSet(Sql, con)
         RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<table border='1'  width='100%'>" & vbCrLf
          If Request("nbSelect") > 0 Then
            RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<tr>" & vbCrLf
            RecapSqlNomenclatuer = RecapSqlNomenclatuer & "         <td colspan='10' >" & RecapSersh & "</td>" & vbCrLf
            
            RecapSqlNomenclatuer = RecapSqlNomenclatuer & "</tr>" & vbCrLf
         End If
         RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<tr >" & vbCrLf
         RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<td class='TitreTableau' colspan ='10'>Liste des pièces :</td>" & vbCrLf
         RecapSqlNomenclatuer = RecapSqlNomenclatuer & "</tr>" & vbCrLf
        
          
         RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<tr>" & vbCrLf
          RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<TD  colspan ='10'><table border='0'  width='100%'><tr>"
         
   RecapSqlNomenclatuer = RecapSqlNomenclatuer & vbCrLf & "<SCRIPT langage='JAVASCRIPT'>"
    
        
    RecapSqlNomenclatuer = RecapSqlNomenclatuer & "function ModifierRechereche(){" & vbCrLf
    RecapSqlNomenclatuer = RecapSqlNomenclatuer & "this.document." & FormName & ".ID__SousEnsembles.value='';"
    RecapSqlNomenclatuer = RecapSqlNomenclatuer & "this.document." & FormName & ".action='Default.asp?Logistique_Page=11'" & vbCrLf
    RecapSqlNomenclatuer = RecapSqlNomenclatuer & vbCrLf & "document." & FormName & ".submit();"
    RecapSqlNomenclatuer = RecapSqlNomenclatuer & "}" & vbCrLf
      
RecapSqlNomenclatuer = RecapSqlNomenclatuer & "function valider3(){" & vbCrLf
RecapSqlNomenclatuer = RecapSqlNomenclatuer & "this.document." & FormName & ".ID__SousEnsembles.value='';"

RecapSqlNomenclatuer = RecapSqlNomenclatuer & "this.document." & FormName & ".action='Default.asp?Logistique_Page=10'" & vbCrLf
RecapSqlNomenclatuer = RecapSqlNomenclatuer & vbCrLf & "document." & FormName & ".submit();"
RecapSqlNomenclatuer = RecapSqlNomenclatuer & "}" & vbCrLf

RecapSqlNomenclatuer = RecapSqlNomenclatuer & "function valider2(){" & vbCrLf
RecapSqlNomenclatuer = RecapSqlNomenclatuer & "this.document." & FormName & ".action='Default.asp?Logistique_Page=8'" & vbCrLf
RecapSqlNomenclatuer = RecapSqlNomenclatuer & vbCrLf & "document." & FormName & ".submit();"
RecapSqlNomenclatuer = RecapSqlNomenclatuer & "}" & vbCrLf
RecapSqlNomenclatuer = RecapSqlNomenclatuer & vbCrLf & "</SCRIPT >"

       RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<TD width='2' valign='top'>" & AfficheBounton("Retour", 8, "'javascript:valider3();'") & "</td>" & vbCrLf
      RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<TD>&nbsp;</td><TD>&nbsp;</td>" & vbCrLf
        RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<TD width='4%' valign='top'>Filtrés sur véhicule :" & vbCrLf
        Sql = "SELECT t_Vehicule.Id, [Projet] , [Véhicule]   FROM t_Vehicule ORDER BY t_Vehicule.Véhicule;"
        
         Sql = "SELECT     t_Vehicule.Id, T_Projects.NumAffaire, t_Vehicule.Véhicule "
        Sql = Sql & "FROM t_Vehicule INNER JOIN "
        Sql = Sql & "T_Projects ON t_Vehicule.Id_Projects_NumAffaire = T_Projects.Id "
        Sql = Sql & "ORDER BY t_Vehicule.Véhicule"
        Set RsFiltre = OpenRecordSet(Sql, con)

       RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<br>" & CreatListe(RsFiltre, Request("RecapFiltreVehicule"), "Id", "1;2", "RecapFiltreVehicule", " onChange='valider2();' ") & "</td>" & vbCrLf
      RsFiltre.Close
      Set RsFiltre = Nothing
      
     
     RecapSqlNomenclatuer = RecapSqlNomenclatuer & "</Tr></table></td>"
    
       RecapSqlNomenclatuer = RecapSqlNomenclatuer & "</tr> " & vbCrLf
       
        
        RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<tr>" & vbCrLf
        
         RecapSqlNomenclatuer = RecapSqlNomenclatuer & "         <td align='center' bgcolor='#CCCCCC'>&nbsp;</td>" & vbCrLf
        RecapSqlNomenclatuer = RecapSqlNomenclatuer & "         <td align='center' bgcolor='#CCCCCC'>Projet</td>" & vbCrLf
        RecapSqlNomenclatuer = RecapSqlNomenclatuer & "         <td   align='center' bgcolor='#CCCCCC'>Véhicule</td>" & vbCrLf
        RecapSqlNomenclatuer = RecapSqlNomenclatuer & "         <td   align='center' bgcolor='#CCCCCC'>Code Sous-Ensseble</td>" & vbCrLf
        RecapSqlNomenclatuer = RecapSqlNomenclatuer & "         <td   align='center' bgcolor='#CCCCCC'>Désignation pièce</td>" & vbCrLf
        RecapSqlNomenclatuer = RecapSqlNomenclatuer & "         <td   align='center' bgcolor='#CCCCCC'>Pièce</td>" & vbCrLf
        RecapSqlNomenclatuer = RecapSqlNomenclatuer & "         <td   align='center' bgcolor='#CCCCCC'>Qts</td>" & vbCrLf
        RecapSqlNomenclatuer = RecapSqlNomenclatuer & "         <td   align='center' bgcolor='#CCCCCC'>Qts Manquante</td>" & vbCrLf
        RecapSqlNomenclatuer = RecapSqlNomenclatuer & "         <td   align='center' bgcolor='#CCCCCC'>Action</td>" & vbCrLf
        RecapSqlNomenclatuer = RecapSqlNomenclatuer & "         <td   align='center' bgcolor='#CCCCCC'>Délais</td>" & vbCrLf

        RecapSqlNomenclatuer = RecapSqlNomenclatuer & "</tr> "
         
        While RsNomenclature.EOF = False
        RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<tr>"
           

      RecapSqlNomenclatuer = RecapSqlNomenclatuer & "             <td align='center'><a href =""javascript:PageSuvante('4','" & Request("id") & "','" & _
                         RsNomenclature("ID_Vehicule_Véhicule") & "','" & RsNomenclature("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Request("Id_Projects_NumAffaire") & "');""><img src='./Images/zoom.gif' border='0' /></td>" & vbCrLf
         
                         
        RecapSqlNomenclatuer = RecapSqlNomenclatuer & "             <td><a href =""javascript:PageSuvante('4','" & Request("id") & "','" & _
                         RsNomenclature("ID_Vehicule_Véhicule") & "','" & RsNomenclature("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & RsNomenclature("Projet") & "&nbsp;</a> </td>" & vbCrLf
       RecapSqlNomenclatuer = RecapSqlNomenclatuer & "             <td><a href =""javascript:PageSuvante('4','" & Request("id") & "','" & _
                         RsNomenclature("ID_Vehicule_Véhicule") & "','" & RsNomenclature("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & RsNomenclature("Véhicule") & "&nbsp;</a> </td>" & vbCrLf
        RecapSqlNomenclatuer = RecapSqlNomenclatuer & "             <td><a href =""javascript:PageSuvante('4','" & Request("id") & "','" & _
                         RsNomenclature("ID_Vehicule_Véhicule") & "','" & RsNomenclature("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & RsNomenclature("Code") & "&nbsp;</a> </td>" & vbCrLf
   RecapSqlNomenclatuer = RecapSqlNomenclatuer & "             <td><a href =""javascript:PageSuvante('4','" & Request("id") & "','" & _
                         RsNomenclature("ID_Vehicule_Véhicule") & "','" & RsNomenclature("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & RsNomenclature("Designation_Piece") & "&nbsp;</a> </td>" & vbCrLf
        RecapSqlNomenclatuer = RecapSqlNomenclatuer & "             <td><a href =""javascript:PageSuvante('4','" & Request("id") & "','" & _
                         RsNomenclature("ID_Vehicule_Véhicule") & "','" & RsNomenclature("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & RsNomenclature("Pièce") & "&nbsp;</a> </td>" & vbCrLf
        
        RecapSqlNomenclatuer = RecapSqlNomenclatuer & "             <td><a href =""javascript:PageSuvante('4','" & Request("id") & "','" & _
                         RsNomenclature("ID_Vehicule_Véhicule") & "','" & RsNomenclature("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & RsNomenclature("Quantite") & "&nbsp;</a> </td>" & vbCrLf
                         
 RecapSqlNomenclatuer = RecapSqlNomenclatuer & "             <td><a href =""javascript:PageSuvante('4','" & Request("id") & "','" & _
                         RsNomenclature("ID_Vehicule_Véhicule") & "','" & RsNomenclature("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & RsNomenclature("QtsManqaute") & "&nbsp;</a> </td>" & vbCrLf
 RecapSqlNomenclatuer = RecapSqlNomenclatuer & "             <td><a href =""javascript:PageSuvante('4','" & Request("id") & "','" & _
                         Request("ID_Vehicule") & "','" & RsNomenclature("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & RsNomenclature("Libellé") & "&nbsp;</a> </td>" & vbCrLf
 
 RecapSqlNomenclatuer = RecapSqlNomenclatuer & "             <td><a href =""javascript:PageSuvante('4','" & Request("id") & "','" & _
                         RsNomenclature("ID_Vehicule_Véhicule") & "','" & RsNomenclature("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & RsNomenclature("Delais") & "&nbsp;</a> </td>" & vbCrLf
        
        RecapSqlNomenclatuer = RecapSqlNomenclatuer & "</tr>"
       RsNomenclature.MoveNext
        Wend
       
    
       
       
       
          RecapSqlNomenclatuer = RecapSqlNomenclatuer & "</table>" & vbCrLf
          
          
          RecapSqlNomenclatuer = RecapSqlNomenclatuer & "<input name='NbS_Enssemble__' type='hidden' value='" & L2 & "'/>"
          RecapSqlNomenclatuer = RecapSqlNomenclatuer & "</form>"
      
       
End Function

 Function RetourLien_SousEnssemble_Nomenclature(ChaineSQL, FormName, champMasquer, ChDefautValeur, ValeurDefaule, Updatable, Url)
    If Request("MAJ") = "OK" Then
        For I = 0 To Request("NbLigne")
            
        Next
    End If
     RetourLien_SousEnssemble_Nomenclature = "<form name='" & FormName & "' method='post'>" & vbCrLf
     RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<input name='MAJ' type='hidden' value='OK'/>" & vbCrLf
     RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<input name='Updatable' type='hidden' value='OK'/>" & vbCrLf
      RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<input name='FormName' type='hidden' value='" & FormName & "'/>" & vbCrLf
       'Response.Write ChaineSQL
      'Response.End
      Set Rs = OpenRecordSet(ChaineSQL, con)
      Set rs_Nomenclaturee = OpenRecordSet("SELECT T_Nomenclature.* FROM T_Nomenclature;", con)
      Set rs_SousEnsembles = OpenRecordSet("SELECT T_SousEnsembles.*FROM T_SousEnsembles;", con)


      NbChamps = Rs.Fields.Count - 1
      RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<input name='NbCol' type='hidden' value='" & Rs.Fields.Count - 1 & "'/>" & vbCrLf
      RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<table border='1' >" & vbCrLf
    If Updatable <> "OK" Then
       
    Else
    
       
       
        
        RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<tr>" & vbCrLf
        RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<TD >" & AfficheBounton("Valider", 1, "'javascript:valider();'") & "</td>" & vbCrLf
        RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<TD colspan='" & NbChamps & "'>&nbsp;</td></tr>" & vbCrLf
        
     
      
     Dim L
     Dim C
     RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<tr>"
    
   RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<tr><td align='center' bgcolor='#CCCCCC'><input name='Champs0' type='hidden' value='id'/>Supprimer</td>" & vbCrLf
 RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "</td>" & vbCrLf

 RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<td align='center' bgcolor='#CCCCCC'>" & vbCrLf
   
 
   RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "Vehicule</td>" & vbCrLf
 RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<td align='center' bgcolor='#CCCCCC'>S/Ensembles</td>" & vbCrLf
     If Trim("" & Url) <> "" Then
     RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<td>Lien</td>"
     End If
     RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "</tr>"
     L = 0
     
     While Rs.EOF = False
     RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<tr>"
    
     
        Pasfaire = False
        RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<td>" & vbCrLf
        RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "    <input type='checkbox' name='Delet_L" & L & "'  value='Delete'>" & vbCrLf
        RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "    <input type='hidden' name='DI_L" & L & "' value='" & Rs.Fields("id").Value & "'>" & vbCrLf
        RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "</td>" & vbCrLf
                                                                                   
        RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<td>" & CreatListe_SousEnssemble_Nomenclature(rs_SousEnsembles, Rs("ID_SousEnsembles"), "id", "Code", "Sous-Ensseble", "txt_L" & L & "_C0") & "</td>" & vbCrLf '"<input type='text' name='txt_L" & L & "_C0'  value='" & rs.fields("ID_Vehicule").value & "'></td>" & vbcrlf
        RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<td>" & CreatListe_SousEnssemble_Nomenclature(rs_Nomenclaturee, Rs("Id_Nomenclature"), "id", "Type", "Pièce", "txt_L" & L & "_C1") & "</td>" & vbCrLf
      
           
          
  
      If Trim("" & Url) <> "" Then
        RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<td><a href='" & Url & Rs.Fields(champMasquer) & "'>Cliquez ici</a></td>"
      End If
     RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "</tr>"
      L = L + 1
     Rs.MoveNext
     
     Wend
     RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<input name='NbLigne' type='hidden' value='" & L - 1 & "'/>"
     For L = 0 To 9
            RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<tr>"
            
            RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<td>New</td>" & vbCrLf

            RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<td>" & CreatListe_SousEnssemble_Nomenclature(rs_SousEnsembles, "", "id", "Code", "Sous-Ensseble", "Newtxt_L" & L & "_C0") & "</td>" & vbCrLf
            RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "<td>" & CreatListe_SousEnssemble_Nomenclature(rs_Nomenclaturee, "", "id", "Type", "Pièce", "Newtxt_L" & L & "_C1") & "</td>" & vbCrLf




            RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "</tr>"
    Next
            RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "</table>" & vbCrLf
   
        
    End If
    RetourLien_SousEnssemble_Nomenclature = RetourLien_SousEnssemble_Nomenclature & "</form >" & vbCrLf
End Function
Function CreatListe_SousEnssemble_Nomenclature(Rs, Id, NameChampsId, NamChampListe, NamChampListe2, NameListe)
Rs.Requery

CreatListe_SousEnssemble_Nomenclature = CreatListe_SousEnssemble_Nomenclature & "<select name='" & NameListe & "' width='100%'>" & vbCrLf
CreatListe_SousEnssemble_Nomenclature = CreatListe_SousEnssemble_Nomenclature & "CreatListe_SousEnssemble_Nomenclature<option value='' width='100%'>&nbsp;</option>" & vbCrLf
While Rs.EOF = False
      If Rs(NameChampsId) = Id Then
        Selected = "selected"
      Else
        Selected = ""
      End If
        CreatListe_SousEnssemble_Nomenclature = CreatListe_SousEnssemble_Nomenclature & "<option width='100%' " & Selected & " value='" & Rs(NameChampsId) & "'>" & Rs(NamChampListe) & " | " & Rs(NamChampListe2) & "</option>" & vbCrLf
    Rs.MoveNext
Wend
 CreatListe_SousEnssemble_Nomenclature = CreatListe_SousEnssemble_Nomenclature & "</select>" & vbCrLf
End Function



 
Function RetourLien_Veicule_Souensembles(ChaineSQL, FormName, champMasquer, ChDefautValeur, ValeurDefaule, Updatable, Url)
    If Request("MAJ") = "OK" Then
        For I = 0 To Request("NbLigne")
            Sql = "UPDATE T_lien_Veicule_Souensembles SET "
            Sql = Sql & "T_lien_Veicule_Souensembles.ID_Vehicule = " & Request("txt_L" & I & "_C0") & ",  "
            Sql = Sql & "T_lien_Veicule_Souensembles.ID__SousEnsembles = " & Request("txt_L" & I & "_C1") & " "
            Sql = Sql & "WHERE T_lien_Veicule_Souensembles.id=" & Request("DI_L" & I) & ";"
            'Response.Write sql
                'Response.End
             ExecuteSql Sql, con
              If Request("Delet_L" & I) = "Delete" Then
            Sql = "DELETE  FROM T_lien_Veicule_Souensembles "
            Sql = Sql & "WHERE T_lien_Veicule_Souensembles.Id=" & Request("DI_L" & I) & ";"
            ExecuteSql Sql, con
          End If
        Next
        For I = 0 To 9
            If (Trim("" & Request("Newtxt_L" & I & "_C0")) & Trim("" & Request("Newtxt_L" & I & "_C10"))) <> "" Then
                Sql = "INSERT INTO T_lien_Veicule_Souensembles ( ID_Vehicule, ID__SousEnsembles ) "
                Sql = Sql & "Values ( " & Request("Newtxt_L" & I & "_C0") & "," & Request("Newtxt_L" & I & "_C1") & ");"
                
                 ExecuteSql Sql, con
             End If
        Next
    End If
     RetourLien_Veicule_Souensembles = "<form name='" & FormName & "' method='post'>" & vbCrLf
     RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<input name='MAJ' type='hidden' value='OK'/>" & vbCrLf
     RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<input name='Updatable' type='hidden' value='OK'/>" & vbCrLf
      RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<input name='FormName' type='hidden' value='" & FormName & "'/>" & vbCrLf
      Set Rs = OpenRecordSet(ChaineSQL, con)
      Set rs_Vehicule = OpenRecordSet("SELECT t_Vehicule.* FROM t_Vehicule;", con)
      Set rs_SousEnsembles = OpenRecordSet("SELECT T_SousEnsembles.*FROM T_SousEnsembles;", con)


      NbChamps = Rs.Fields.Count - 1
      RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<input name='NbCol' type='hidden' value='" & Rs.Fields.Count - 1 & "'/>" & vbCrLf
      RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<table border='1' >" & vbCrLf
    If Updatable <> "OK" Then
       
    Else
    
       
       
        
        RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<tr>" & vbCrLf
        RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<TD >" & AfficheBounton("Valider", 1, "'javascript:valider();'") & "</td>" & vbCrLf
        RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<TD colspan='" & NbChamps & "'>&nbsp;</td></tr>" & vbCrLf
        
     
      
     Dim L
     Dim C
     RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<tr>"
    
   RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<tr><td align='center' bgcolor='#CCCCCC'><input name='Champs0' type='hidden' value='id'/>Supprimer</td>" & vbCrLf
 RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "</td>" & vbCrLf

 RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<td align='center' bgcolor='#CCCCCC'>" & vbCrLf
   
 
   RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "Vehicule</td>" & vbCrLf
 RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<td align='center' bgcolor='#CCCCCC'>S/Ensembles</td>" & vbCrLf
     If Trim("" & Url) <> "" Then
     RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<td>Lien</td>"
     End If
     RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "</tr>"
     L = 0
     
     While Rs.EOF = False
     RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<tr>"
    
     
        Pasfaire = False
        RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<td>" & vbCrLf
        RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "    <input type='checkbox' name='Delet_L" & L & "'  value='Delete'>" & vbCrLf
        RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "    <input type='hidden' name='DI_L" & L & "' value='" & Rs.Fields("id").Value & "'>" & vbCrLf
        RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "</td>" & vbCrLf
                                                                                   
        RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<td>" & CreatListeVeicule_Souensembles(rs_Vehicule, Rs("ID_Vehicule"), "id", "Programme", "Véhicule", "txt_L" & L & "_C0") & "</td>" & vbCrLf '"<input type='text' name='txt_L" & L & "_C0'  value='" & rs.fields("ID_Vehicule").value & "'></td>" & vbcrlf
        RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<td>" & CreatListeVeicule_Souensembles(rs_SousEnsembles, Rs("ID__SousEnsembles"), "id", "Code", "Sous-Ensseble", "txt_L" & L & "_C1") & "</td>" & vbCrLf
      
           
          
  
      If Trim("" & Url) <> "" Then
        RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<td><a href='" & Url & Rs.Fields(champMasquer) & "'>Cliquez ici</a></td>"
      End If
     RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "</tr>"
      L = L + 1
     Rs.MoveNext
     
     Wend
     RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<input name='NbLigne' type='hidden' value='" & L - 1 & "'/>"
     For L = 0 To 9
            RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<tr>"
            
            RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<td>New</td>" & vbCrLf

            RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<td>" & CreatListeVeicule_Souensembles(rs_Vehicule, "", "id", "Programme", "Véhicule", "Newtxt_L" & L & "_C0") & "</td>" & vbCrLf
            RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "<td>" & CreatListeVeicule_Souensembles(rs_SousEnsembles, "", "id", "Code", "Sous-Ensseble", "Newtxt_L" & L & "_C1") & "</td>" & vbCrLf




            RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "</tr>"
    Next
            RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "</table>" & vbCrLf
   
        
    End If
    RetourLien_Veicule_Souensembles = RetourLien_Veicule_Souensembles & "</form >" & vbCrLf
End Function
Function CreatListeVeicule_Souensembles(Rs, Id, NameChampsId, NamChampListe, NamChampListe2, NameListe)
Rs.Requery

CreatListeVeicule_Souensembles = CreatListeVeicule_Souensembles & "<select name='" & NameListe & "' width='100%'>" & vbCrLf
CreatListeVeicule_Souensembles = CreatListeVeicule_Souensembles & "CreatListeVeicule_Souensembles<option value='' width='100%'>&nbsp;</option>" & vbCrLf
While Rs.EOF = False
      If Rs(NameChampsId) = Id Then
        Selected = "selected"
      Else
        Selected = ""
      End If
        CreatListeVeicule_Souensembles = CreatListeVeicule_Souensembles & "<option width='100%' " & Selected & " value='" & Rs(NameChampsId) & "'>" & Rs(NamChampListe) & " | " & Rs(NamChampListe2) & "</option>" & vbCrLf
    Rs.MoveNext
Wend
 CreatListeVeicule_Souensembles = CreatListeVeicule_Souensembles & "</select>" & vbCrLf
End Function



 
Function RetourSqlNomenclatuer(FormName, champMasquer, ChDefautValeur, ValeurDefaule, Updatable)
Dim Sql
Dim SqlValue
Dim BbParP As Integer
Dim Checked
On Error Resume Next
BbParP = 6
If Session("Updatable") = "True" And Request("MAJ") = "OK" Then

con.Execute "UPDATE T_SousEnsembles SET T_SousEnsembles.flag = Null WHERE T_SousEnsembles.flag='" & Session("NumJob") & "';"
'Response.Write Request("NbLigne")
'Response.End
   
    For I = 1 To Request("NbLigne")
    OPAF_Ok = False
            REAF_OK = False
    Modification = ""
 Sav_annuler_Change = False
  Sql = "SELECT T_SousEnsembles.flag FROM T_SousEnsembles "
                Sql = Sql & "WHERE T_SousEnsembles.Id=" & noquote(Request("Sav_ID_SousEnsembles_Sous-Ensseble_" & I), "num") & " "
                Sql = Sql & "and  Not T_SousEnsembles.flag Is Null  ;"
            'Response.Write sql
            'Response.End
                Set Rs_S_Ens = OpenRecordSet(Sql, con)
               If Rs_S_Ens.EOF = True Then
                    Sql = "UPDATE T_SousEnsembles  "
                    Sql = Sql & "SET T_SousEnsembles.flag = '" & Session("NumJob") & "' "
                    Sql = Sql & "WHERE T_SousEnsembles.Id=" & noquote(Request("Sav_ID_SousEnsembles_Sous-Ensseble_" & I), "num")
                    con.Execute Sql
               End If
                Rs_S_Ens.Close
                  Set rsSens = Nothing
'72870
            If Request("id_" & I) <> "" Then
            ModifieL = ""
            
             
            
            txt = Request("Désination FE_" & I)                     '1
            txt2 = Request("Sav_Désination FE_" & I)
            txt = Val("" & Request("ID_SousEnsembles_Sous-Ensseble_" & I))
             txt2 = Val("" & Request("Sav_ID_SousEnsembles_Sous-Ensseble_" & I))
            txt = txt & Request("Annuler_" & I)
            txt2 = txt2 & Request("Sav_Annuler_" & I)
            txt = txt & Request("Type_" & I)
            txt2 = txt2 & Request("Sav_Type_" & I) '2
            txt = txt & Request("Pièce_" & I)
            txt2 = txt2 & Request("Sav_Pièce_" & I)  '3
            txt = txt & Request("Indice_" & I)
            txt2 = txt2 & Request("Sav_Indice_" & I) '4
            txt = txt & Request("Commantaire_" & I)
            txt2 = txt2 & Request("Sav_Commantaire_" & I)    '5
            txt = txt & Request("Designation_Piece_" & I)
            txt2 = txt2 & Request("Sav_Designation_Piece_" & I)  '6
            txt = txt & Request("Famille/FE/V_" & I)
            txt2 = txt2 & Request("Sav_Famille/FE/V_" & I) '7
            txt = txt & Request("Num_appro_ S_" & I)
            txt2 = txt2 & Request("Sav_Num_appro_ S_" & I)    '8
            txt = txt & Val("" & Request("Quantite_" & I))
            txt2 = txt2 & Val("" & Request("Sav_Quantite_" & I)) '9
            txt = txt & Request("Bon_sortie_" & I)
            txt2 = txt2 & Request("Sav_Bon_sortie_" & I)  '10
            txt = txt & Request("Observation_ou_magasin_" & I)
            txt2 = txt2 & Request("Sav_Observation_ou_magasin_" & I) '11
            txt = txt & Request("Num_appro_" & I)
            txt2 = txt2 & Request("Sav_Num_appro_" & I) '12
            txt = txt & Val("" & Request("Masse_" & I))
            txt2 = txt2 & Val("" & Request("Sav_Masse_" & I)) '13
            txt = txt & Request("Commandee_le_" & I)
            txt2 = txt2 & Request("Sav_Commandee_le_" & I) '14
             txt = txt & Request("Collectee_le_" & I)
             txt2 = txt2 & Request("Sav_Collectee_le_" & I) '15
            txt = txt & Val("" & Request("Delais_" & I))
            txt2 = txt2 & Val("" & Request("Sav_Delais_" & I)) '16
            txt = txt & Request("Reçue_le_" & I)
            txt2 = txt2 & Request("Sav_Reçue_le_" & I) '17
            txt = txt & Val("" & Request("Qt_reçue_" & I))
            txt2 = txt2 & Val("" & Request("Sav_Qt_reçue_" & I)) '18
            txt = txt & Val("" & Request("Reste_a_recevoir_" & I))
            txt2 = txt2 & Val("" & Request("Sav_Reste_a_recevoir_" & I)) '19
            txt = txt & Request("N_de_BL_" & I)
            txt2 = txt2 & Request("Sav_N_de_BL_" & I) '20
            txt = txt & Val("" & Request("Reaffectee_a_" & I))
            txt2 = txt2 & Val("" & Request("Sav_Reaffectee_a_" & I))  '21
            txt = txt & Val("" & Request("Qt_Reaffectee_" & I))
            txt2 = txt2 & Val("" & Request("Sav_Qt_Reaffectee_" & I)) '22
            txt = txt & Val("" & Request("En_provenance_de_" & I))
            txt2 = txt2 & Val("" & Request("Sav_En_provenance_de_" & I)) '23
            txt = txt & Request("Retournee_le_" & I)
            txt2 = txt2 & Request("Sav_Retournee_le_" & I)  '24
            txt = txt & Val("" & Request("QT_Retournee_" & I))
            txt2 = txt2 & Val("" & Request("Sav_QT_Retournee_" & I)) '25
            txt = txt & Request("Num_de_retour_" & I)
            txt2 = txt2 & Request("Sav_Num_de_retour_" & I)  '26
            txt = txt & Request("Modifiee_le_" & I)
            txt2 = txt2 & Request("Sav_Modifiee_le_" & I)
            txt = txt & Val("" & Request("Action_" & I))
             txt2 = txt2 & Val("" & Request("Sav_Action_" & I))
             If Val("" & Request("Sav_Reste_a_recevoir_" & I)) <> Val("" & Request("Reste_a_recevoir_" & I)) Then
                 Modification = Modification & "Reste a recevoir = " & Val("" & Request("Sav_Reste_a_recevoir_" & I)) & Chr(10)
            End If
            Modification = ""
                               '17
                     
            
               If txt <> txt2 Then
               annuler_Change = True
               If Request("Sav_Annuler_" & I) <> Request("Annuler_" & I) Then
                    If Trim("" & Request("sav_Annuler_" & I)) = "" Then
                        Modification = Modification & "Activé = Non" & Chr(10)
                    Else
                        Modification = Modification & "Activé = Oui" & Chr(10)
                    End If
               End If
               If Val("" & Request("Sav_ID_SousEnsembles_Sous-Ensseble_" & I)) <> Val("" & Request("ID_SousEnsembles_Sous-Ensseble_" & I)) Then
                    Modification = Modification & "Sous Ensembles = " & Request("Sav_ID_SousEnsembles_Sous-Ensseble_" & I) & Chr(10)
               End If
                If Request("Sav_Designation_Piece_" & I) <> Request("Designation_Piece_" & I) Then
                 Modification = Modification & "Designation Piece = " & Request("Sav_Designation_Piece_" & I) & Chr(10)
            End If
               If Request("Sav_Pièce_" & I) <> Request("Pièce_" & I) Then
             Modification = Modification & "Pièce = " & Request("Sav_Pièce_" & I) & Chr(10)
             End If
            If Request("Sav_Indice_" & I) <> Request("Indice_" & I) Then
                 Modification = Modification & "Indice = " & Request("Sav_Indice_" & I) & Chr(10)
            End If
            If Val("" & Request("Sav_Quantite_" & I)) <> Val("" & Request("Quantite_" & I)) Then
             Modification = Modification & "Quantite = " & Request("Sav_Quantite_" & I) & Chr(10)
           End If
            If Val("" & Request("Sav_Reste_a_recevoir_" & I)) <> Val("" & Request("Reste_a_recevoir_" & I)) Then
                 Modification = Modification & "Reste a recevoir = " & Val("" & Request("Sav_Reste_a_recevoir_" & I)) & Chr(10)
            End If
            If Val("" & Request("Sav_Action_" & I)) <> Val("" & Request("Action_" & I)) Then
                 Modification = Modification & "Action = " & Request("Sav_Action_" & I) & Chr(10)
              End If
              If Request("Sav_Désination FE_" & I) <> Request("Désination FE_" & I) Then
                 Modification = Modification & "Désination FE = " & Request("Sav_Désination FE_" & I) & Chr(10)
              End If
               
            If Request("Sav_Type_" & I) <> Request("Type_" & I) Then
                 Modification = Modification & "Type = " & Request("Sav_Type_" & I) & Chr(10)
            End If
            
            If Request("Sav_Commantaire_" & I) <> Request("Commantaire_" & I) Then
                 Modification = Modification & "Commantaire = " & Request("Sav_Commantaire_" & I) & Chr(10)
            End If
           
            If Request("Sav_Famille/FE/V_" & I) <> Request("Famille/FE/V_" & I) Then
                 Modification = Modification & "Famille/FE/V = " & Request("Sav_Famille/FE/V_" & I) & Chr(10)
            End If
           If Request("Sav_Num_appro_ S_" & I) <> Request("Num_appro_ S_" & I) Then
             Modification = Modification & "Num approS = " & Request("Sav_ANum_appro_ S_" & I) & Chr(10)
           End If
          
            If Request("Sav_Bon_sortie_" & I) <> Request("Bon_sortie_" & I) Then
                 Modification = Modification & "Bon sortie = " & Request("Sav_Bon_sortie_" & I) & Chr(10)
            End If
            If Request("Sav_Observation_ou_magasin_" & I) <> Request("Observation_ou_magasin_" & I) Then
                 Modification = Modification & "Observation ou magasin = " & Request("Sav_Observation_ou_magasin_" & I) & Chr(10)
            End If
            If Request("Sav_Num_appro_" & I) <> Request("Num_appro_" & I) Then
                 Modification = Modification & "Num appro = " & Request("Sav_Num_appro_" & I) & Chr(10)
            End If
            If Val("" & Request("Sav_Masse_" & I)) <> Val("" & Request("Masse_" & I)) Then
                 Modification = Modification & "ActMasseion = " & Request("Sav_Masse_" & I) & Chr(10)
            End If
            If Request("Sav_Commandee_le_" & I) <> Request("Commandee_le_" & I) Then
                 Modification = Modification & "Commandee le = " & Request("Sav_Commandee_le_" & I) & Chr(10)
            End If
             If Request("Sav_Collectee_le_" & I) <> Request("Collectee_le_" & I) Then
                 Modification = Modification & "Collectee le = " & Request("Sav_Collectee_le_" & I) & Chr(10)
             End If
            If Val("" & Request("Sav_Delais_" & I)) <> Val("" & Request("Delais_" & I)) Then
                 Modification = Modification & "Delais = " & Request("Sav_Delais_" & I) & Chr(10)
            End If
           
            If Val("" & Request("Sav_Qt_reçue_" & I)) <> Val("" & Request("Qt_reçue_" & I)) Then
                 Modification = Modification & "Qt reçue = " & Val("" & Request("Sav_Qt_reçue_" & I)) & Chr(10)
            End If
           
             If Val("" & Request("Sav_Qt_Reaffectee_" & I)) <> Val("" & Request("Qt_Reaffectee_" & I)) Then
                 Modification = Modification & "Qt Reaffectee = " & Val("" & Request("Sav_Qt_Reaffectee_" & I)) & Chr(10)
            End If
            If Request("Sav_Reaffectee_a_" & I) <> Request("Reaffectee_a_" & I) Then
                 Modification = Modification & "Reaffectee a = " & Request("Sav_Reaffectee_a_" & I) & Chr(10)
            End If
           
            If Request("Sav_En_provenance_de_" & I) <> Request("Sav_En_provenance_de_" & I) Then
                 Modification = Modification & "En provenance de = " & Request("Sav_En_provenance_de_" & I) & Chr(10)
            End If
           If Request("Sav_Retournee_le_" & I) <> Request("Retournee_le_" & I) Then
                 Modification = Modification & "Retournee le = " & Request("Sav_Retournee_le_" & I) & Chr(10)
            End If
           If Val("" & Request("Sav_QT_Retournee_" & I)) <> Val("" & Request("QT_Retournee_" & I)) Then
                 Modification = Modification & "QT Retournee = " & Request("Sav_QT_Retournee_" & I) & Chr(10)
            End If
                ModifieL = Date
                
               End If
               
               
               
            Sql = "UPDATE T_Nomenclature SET  "
            SqlValue = "T_Nomenclature.Modification=" & noquote(Modification, "txt") & " ,"
             'SqlValue=SqlValue & "Id_T_Actions_Libellé= " &  noquote( Request("Sav_Action_" & i),"num") & ","
             If Val("" & Request("save_Action_" & I)) <> Val("" & Request("Action_" & I)) Then annuler_Change = True
            SqlValue = SqlValue & "T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] = " & noquote(Request("ID_SousEnsembles_Sous-Ensseble_" & I), "num") & ",  "
            SqlValue = SqlValue & "T_Nomenclature.[Désination FE] = " & noquote(Request("Désination FE_" & I), "txt") & ",  "
            SqlValue = SqlValue & "T_Nomenclature.Type = " & noquote(Request("Type_" & I), "txt") & ",  "
            SqlValue = SqlValue & "T_Nomenclature.Pièce = " & noquote(Request("Pièce_" & I), "txt") & ",  "
            'Response.Write i & " : " & trim("" & Request("annuler_" & i)) & "<br>"
            
            
             If Trim("" & Request("sav_annuler_" & I)) <> Trim("" & Request("annuler_" & I)) Then annuler_Change = True
                
           
            
            SqlValue = SqlValue & "T_Nomenclature.Indice = " & noquote(Request("Indice_" & I), "txt") & ",  "
            SqlValue = SqlValue & "T_Nomenclature.Commantaire =" & noquote(Request("Commantaire_" & I), "txt") & ",  "
            SqlValue = SqlValue & "T_Nomenclature.Designation_Piece = " & noquote(Request("Designation_Piece_" & I), "txt") & ",  "
            SqlValue = SqlValue & "T_Nomenclature.[Famille/FE/V] = " & noquote(Request("Famille/FE/V_" & I), "txt") & ",  "
            SqlValue = SqlValue & "T_Nomenclature.[Num_appro_ S] = " & noquote(Request("Num_appro_ S_" & I), "txt") & ",  "
            SqlValue = SqlValue & "T_Nomenclature.Quantite = " & noquote(Request("Quantite_" & I), "num") & ",  "
             SqlValue = SqlValue & "T_Nomenclature.Bon_sortie = " & noquote(Request("Bon_sortie_" & I), "txt") & ","
            SqlValue = SqlValue & "T_Nomenclature.Observation_ou_magasin = " & noquote(Request("Observation_ou_magasin_" & I), "txt") & ",  "
            SqlValue = SqlValue & "T_Nomenclature.Num_appro = " & noquote(Request("Num_appro_" & I), "txt") & ",  "
            SqlValue = SqlValue & "T_Nomenclature.Masse = " & noquote(Request("Masse_" & I), "num") & ",  "
            SqlValue = SqlValue & "T_Nomenclature.Commandee_le =" & noquote(Request("Commandee_le_" & I), "date") & ",  "
             SqlValue = SqlValue & "T_Nomenclature.Collectee_le = " & noquote(Request("Collectee_le_" & I), "date") & ",  "
            
            SqlValue = SqlValue & "T_Nomenclature.Id_T_Delais_Libellé = " & noquote(Request("Delais_" & I), "num") & ",  "

            SqlValue = SqlValue & "T_Nomenclature.Qt_reçue = " & noquote(Request("Qt_reçue_" & I), "num") & ",  "
            SqlValue = SqlValue & "T_Nomenclature.Reste_a_recevoir =" & noquote(Request("Reste_a_recevoir_" & I), "num") & ",  "
            SqlValue = SqlValue & "T_Nomenclature.Reaffectee_a = " & noquote(Request("Reaffectee_a_" & I), "num") & ",  "
            SqlValue = SqlValue & "T_Nomenclature.Qt_Reaffectee = " & noquote(Request("Qt_Reaffectee_" & I), "num") & ",  "
            SqlValue = SqlValue & "T_Nomenclature.En_provenance_de = " & noquote(Request("En_provenance_de_" & I), "num") & ",  "
            SqlValue = SqlValue & "T_Nomenclature.QT_Retournee = " & noquote(Request("QT_Retournee_" & I), "num") & ",  "
              
            
            Dim IsInsert
            
            IsInsert = Replace(Replace(Session("Qt_reçue"), "false", ""), "Faux", "")
            
             If Val(Trim("" & Request("Qt_reçue_" & I))) + Val(Trim("" & Request("Qt_Reaffectee_" & I))) + Val(Trim("" & Request("QT_Retournee_" & I))) <> 0 Then
                   OPAF_Ok = True
                 '***********************************************************
                 If Session("Qt_reçue") = "True" Then
                 If Request("Qt_reçue_" & I) + 0 <> 0 Then
                 
                    SqlValue = SqlValue & "T_Nomenclature.Reçue_le = " & noquote(Now(), "date") & ",  "
                    OPAF_BR = "########"
                        If OPAF_BR = "" Then
                        
                        Sql_Choix_Vehicule = "SELECT t_Vehicule.Véhicule "
                        Sql_Choix_Vehicule = Sql_Choix_Vehicule & "FROM (t_Vehicule INNER JOIN T_SousEnsembles ON t_Vehicule.Id =  "
                        Sql_Choix_Vehicule = Sql_Choix_Vehicule & "T_SousEnsembles.ID_Vehicule_Véhicule) INNER JOIN  "
                        Sql_Choix_Vehicule = Sql_Choix_Vehicule & "T_Nomenclature ON T_SousEnsembles.Id =  "
                        Sql_Choix_Vehicule = Sql_Choix_Vehicule & "T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] "
                        Sql_Choix_Vehicule = Sql_Choix_Vehicule & "WHERE T_Nomenclature.Id=" & Request("id_" & I) & " ;"
                        Set Rs_Sql_Choix_Vehicule = con.Execute(Sql_Choix_Vehicule)
                        
                        OPAF_BR = NumeroChrono("CHRONO", "BR_" & Rs_Sql_Choix_Vehicule("Véhicule") & "_")
                        Rs_Sql_Choix_Vehicule.Close
                        Set Rs_Sql_Choix_Vehicule = Nothing
                     End If
                    SqlValue = SqlValue & "T_Nomenclature.N_de_BL = '" & OPAF_BR & "',  "
                   Set Rs_br = OpenRecordSet("SELECT T_Stock.Id FROM T_Stock WHERE T_Stock.NumELib='AOPF_BR';", con)
                   If Rs_br.EOF = False Then
                    sql_br = "INSERT INTO T_Mouvement_de_stock ( Id_Nomenclature, Id_Stock, [Qts Entrer], [Qts Solde] ) "
                    sql_br = sql_br & "VALUES( " & Request("id_" & I) & " , " & Rs_br("id") & ", " & Request("Qt_reçue_" & I) & ", " & noquote(Request("Reste_a_recevoir_" & I), "num") & " );"
                     con.Execute sql_br

                   End If
                    Set Rs_br = CloseRecordSet(Rs_br)
                    
                Else
                    
                     SqlValue = SqlValue & "T_Nomenclature.Reçue_le = null,  "
                    SqlValue = SqlValue & "T_Nomenclature.N_de_BL = null,  "
                End If
                End If
                '***********************************************************
                    If Session("Qt Réaffectée") = "True" Then
                    If Request("QT_Retournee_" & I) + 0 <> 0 And Request("Reaffectee_a_" & I) <> "" Then
                        SqlValue = SqlValue & "T_Nomenclature.Reaffectee_Le = " & noquote(Now(), "date") & ",  "
                        REAF_OK = True
                       
                    Else
                         
                         SqlValue = SqlValue & "T_Nomenclature.Reaffectee_Le = null,  "
                    End If
                End If
                '***************************************************************
                    If Session("QT_Retournée") = "True" Then
                    If Request("QT_Retournee_" & I) + 0 <> 0 Then
                    
                    Sql = Sql & "T_Nomenclature.Retournee_le = " & noquote(Now(), "date") & ",  "
                    OPAF_BL = "########"
                     
                    Sql = Sql & "T_Nomenclature.Num_de_retour = '" & OPAF_BL & "',  "
                    Set Rs_br = OpenRecordSet("SELECT T_Stock.Id FROM T_Stock WHERE T_Stock.NumELib='OPAF_BL';", con)
                   If Rs_br.EOF = False Then
                    sql_br = "INSERT INTO T_Mouvement_de_stock ( Id_Nomenclature, Id_Stock, [Qts Entrer], [Qts Solde] ) "
                    sql_br = sql_br & "VALUES( " & Request("id_" & I) & " , " & Rs_br("id") & ", " & Request("Qt_reçue_" & I) & ", " & noquote(Request("Reste_a_recevoir_" & I), "num") & " );"
                    
                      con.Execute sql_br

                   End If
                    Set Rs_br = CloseRecordSet(Rs_br)
                Else
                     
                     SqlValue = SqlValue & "T_Nomenclature.Retournee_le = null,  "
                    SqlValue = SqlValue & "T_Nomenclature.Num_de_retour = null,  "
                End If
                
            SqlValue = SqlValue & "T_Nomenclature.Modifiee_le = " & noquote(Now(), "date") & ", "
                
            Else
                  SqlValue = SqlValue & "T_Nomenclature.Reçue_le = null,  "
                    SqlValue = SqlValue & "T_Nomenclature.N_de_BL = null,  "
                     SqlValue = SqlValue & "T_Nomenclature.Reaffectee_Le = null,  "
                     SqlValue = SqlValue & "T_Nomenclature.Retournee_le = null,  "
                    SqlValue = SqlValue & "T_Nomenclature.Num_de_retour = null,  "
                SqlValue = SqlValue & "T_Nomenclature.Modifiee_le = null, "
                
            End If

                                'if Request("QT_Retournee_" & i)+0 <>0 then
'
 '                                   SqlValue=SqlValue & "T_Nomenclature.Retournee_le = #" & now() & "#,  "
                                    OPAF_BL = "########"
                                    
                                     NewReafectation I, "########"
   '                                 SqlValue=SqlValue & "T_Nomenclature.Num_de_retour = '" & OPAF_BL & "',  "
    '                                set  Rs_br =OpenRecordSet("SELECT T_Stock.Id FROM T_Stock WHERE T_Stock.NumELib='OPAF_BL';",Con)
     '                               if Rs_br.eof=false then
      '                                  sql_br ="INSERT INTO T_Mouvement_de_stock ( Id_Nomenclature, Id_Stock, [Qts Entrer], [Qts Solde] ) "
       '                                 sql_br=sql_br & "VALUES( " & Request("id_" & i) & " , " & Rs_br("id") & ", " & Request("Qt_reçue_" & i) & ", " & noquote(Request("Reste_a_recevoir_" & i),"num") & " );"
        '                                 Con.execute  sql_br
'
 '                                   end if
 
  '                                  set Rs_br=CloseRecordSet(Rs_br)
   '                                 else
'
 '                                   SqlValue=SqlValue & "T_Nomenclature.Retournee_le = null,  "
  '                                  SqlValue=SqlValue & "T_Nomenclature.Num_de_retour = null,  "
   '                                 end if
'
 '                                   SqlValue=SqlValue & "T_Nomenclature.Modifiee_le = #" & now() & "#, "
'
 '                               else
  '                                  SqlValue=SqlValue & "T_Nomenclature.Reçue_le = null,  "
   '                                 SqlValue=SqlValue & "T_Nomenclature.N_de_BL = null,  "
    '                                SqlValue=SqlValue & "T_Nomenclature.Reaffectee_Le = null,  "
     '                               SqlValue=SqlValue & "T_Nomenclature.Retournee_le = null,  "
      '                              SqlValue=SqlValue & "T_Nomenclature.Num_de_retour = null,  "
       '                             SqlValue=SqlValue & "T_Nomenclature.Modifiee_le = null, "
'
 '                               end if
                 End If
            If Trim("" & SqlValue) <> "" Then
            SqlValue = Left(Trim(SqlValue), Len(Trim(SqlValue)) - 1) & " "
            End If
             
            If Trim(SqlValue) = "" Then
                NoExecuteEq = True
            Else
                NoExecuteEq = False
            End If
            If Trim("" & ModifieL) <> "" Then
                If InStr(1, SqlValue, "T_Nomenclature.Modifiee_le") = 0 Then
                    SqlValue = "T_Nomenclature.Modifiee_le=" & noquote(ModifieL, "date") & "," & SqlValue
                End If
            End If
            Sql = Sql & SqlValue
            Sql = Sql & "WHERE T_Nomenclature.Id=" & Request("id_" & I) & " and ("
            Sql = Sql & "T_Nomenclature.Flag<>" & noquote(Session("NumJob"), "txt") & " or T_Nomenclature.Flag Is Null);"
            If NoExecuteEq = False Then
           ' Response.Write sql & "<br>"
            'Response.End
            
               con.Execute Sql
             End If
            If OPAF_Ok = True And annuler_Change = True Then
                annuler_Change = False
                  MajNewNomenclateur I
                   
            End If
            
            If annuler_Change = True Then
                annuler_Change = False
                'If Session("Supprimer") <> "True" Then
                    
                    MajNewNomenclateur I
                'End If
            End If
           
      Else
      
      
       
            txt = Request("Sav_Désination FE_" & I)                     '1
            txt = txt & Request("Sav_Type_" & I)                        '2
            txt = txt & Request("Sav_Pièce_" & I)                       '3
            txt = txt & Request("Sav_Indice_" & I)                      '4
            txt = txt & Request("Sav_Commantaire_" & I)                 '5
            txt = txt & Request("Sav_Designation_Piece_" & I)           '6
            txt = txt & Request("Sav_Famille/FE/V_" & I)                '7
            txt = txt & Request("Sav_Num_appro_ S_" & I)                '8
            txt = txt & Request("Sav_Quantite_" & I)                    '9
            txt = txt & Request("Sav_Bon_sortie_" & I)                  '10
            txt = txt & Request("Sav_Observation_ou_magasin_" & I)      '11
            txt = txt & Request("Sav_Num_appro_" & I)                   '12
            txt = txt & Request("Sav_Masse_" & I)                       '13
            txt = txt & Request("Sav_Commandee_le_" & I)                '14
             txt = txt & Request("Sav_Collectee_le_" & I)               '15
            txt = txt & Request("Sav_Delais_" & I)                  '16
            txt = txt & Request("Sav_Reçue_le_" & I)                    '17
            txt = txt & Request("Sav_Qt_reçue_" & I)                    '18
            txt = txt & Request("Sav_Reste_a_recevoir_" & I)            '19
            txt = txt & Request("Sav_N_de_BL_" & I)                     '20
            txt = txt & Request("Sav_Reaffectee_a_" & I)                '21
            txt = txt & Request("Sav_Qt_Reaffectee_" & I)               '22
            txt = txt & Request("Sav_En_provenance_de_" & I)            '23
            txt = txt & Request("Sav_Retournee_le_" & I)                '24
            txt = txt & Request("Sav_QT_Retournee_" & I)                '25
            txt = txt & Request("Sav_Num_de_retour_" & I)               '26
            txt = txt & Request("Sav_Modifiee_le_" & I)                 '27
            
            If Trim("" & txt) <> "" Then
            
                    Sql = "insert into T_Nomenclature ("                '1
                    Sql = Sql & "[ID_SousEnsembles_Sous-Ensseble],"     '2
                    
                   Sql = Sql & "[Désination FE],  "                    '3
                    Sql = Sql & "Type,  "                               '4
                    Sql = Sql & "Pièce,  "                              '5
                    Sql = Sql & "Indice,  "                             '6
                    Sql = Sql & "Commantaire,  "                        '7
                    Sql = Sql & "Designation_Piece,  "                  '8
                    Sql = Sql & "[Famille/FE/V],  "                     '9
                    Sql = Sql & "[Num_appro_ S],  "                     '10
                    Sql = Sql & "Quantite,  "                           '11
                    Sql = Sql & "Bon_sortie,  "                         '12
                    Sql = Sql & "Observation_ou_magasin,  "             '13
                    Sql = Sql & "Num_appro,  "                          '14
                    Sql = Sql & "Masse,  "                              '15
                    Sql = Sql & "Commandee_le,  "                       '16
                    Sql = Sql & "Collectee_le,  "                      '17
                    Sql = Sql & "Id_T_Delais_Libellé,  "                         '18
                    Sql = Sql & "Reçue_le,  "                           '19
                    Sql = Sql & "Qt_reçue,  "                           '20
                    Sql = Sql & "Reste_a_recevoir,  "                   '21
                    Sql = Sql & "N_de_BL,  "                            '22
                    Sql = Sql & "Reaffectee_a,  "                       '23
                    Sql = Sql & "Qt_Reaffectee,  "                      '24
                    Sql = Sql & "En_provenance_de,  "                   '25
                    Sql = Sql & "Retournee_le,  "                       '26
                    Sql = Sql & "QT_Retournee,  "                       '27
                    Sql = Sql & "Num_de_retour,  "                      '28
                    Sql = Sql & "Modifiee_le,"                          '29
                    Sql = Sql & "Id_T_Actions_Libellé) "               '30
                    
                    Sql = Sql & "values ("
                    Sql = Sql & noquote(Request("Sav_ID_SousEnsembles_Sous-Ensseble_" & I), "num") & "," '2
                    
                    Sql = Sql & noquote(Request("Sav_Désination FE_" & I), "txt") & ",  "               '3
                    Sql = Sql & noquote(Request("Sav_Type_" & I), "txt") & ",  "                        '4
                    Sql = Sql & noquote(Request("Sav_Pièce_" & I), "txt") & ",  "                       '5
                    Sql = Sql & noquote(Request("Sav_Indice_" & I), "txt") & ",  "                      '6
                    Sql = Sql & noquote(Request("Sav_Commantaire_" & I), "txt") & ",  "                 '7
                    Sql = Sql & noquote(Request("Sav_Designation_Piece_" & I), "txt") & ",  "           '8
                    Sql = Sql & noquote(Request("Sav_Famille/FE/V_" & I), "txt") & ",  "                '9
                    Sql = Sql & noquote(Request("Sav_Num_appro_ S_" & I), "txt") & ",  "                '10
                    Sql = Sql & noquote(Request("Sav_Quantite_" & I), "num") & ",  "                    '11
                    Sql = Sql & noquote(Request("Sav_Bon_sortie_" & I), "txt") & ",  "                  '12
                    Sql = Sql & noquote(Request("Sav_Observation_ou_magasin_" & I), "txt") & ",  "      '13
                    Sql = Sql & noquote(Request("Sav_Num_appro_" & I), "txt") & ",  "                   '14
                    Sql = Sql & noquote(Request("Sav_Masse_" & I), "num") & ",  "                       '15
                    Sql = Sql & noquote(Request("Sav_Commandee_le_" & I), "date") & ",  "               '16
                    Sql = Sql & noquote(Request("Sav_Collectee_le_" & I), "date") & ",  "               '17
                    Sql = Sql & noquote(Request("Sav_Delais_" & I), "num") & ",  "                 '18
                    Sql = Sql & noquote(Request("Sav_Reçue_le_" & I), "date") & ",  "                   '19
                    Sql = Sql & noquote(Request("Sav_Qt_reçue_" & I), "num") & ",  "                    '20
                    Sql = Sql & noquote(Request("Sav_Reste_a_recevoir_" & I), "num") & ",  "            '21
                    Sql = Sql & noquote(Request("Sav_N_de_BL_" & I), "txt") & ",  "                     '22
                    Sql = Sql & noquote(Request("Sav_Reaffectee_a_" & I), "num") & ",  "                '23
                    Sql = Sql & noquote(Request("Sav_Qt_Reaffectee_" & I), "num") & ",  "               '24
                    Sql = Sql & noquote(Request("Sav_En_provenance_de_" & I), "num") & ",  "            '25
                    Sql = Sql & noquote(Request("Sav_Retournee_le_" & I), "date") & ",  "               '26
                    Sql = Sql & noquote(Request("Sav_QT_Retournee_" & I), "num") & ",  "                '27
                    Sql = Sql & noquote(Request("Sav_Num_de_retour_" & I), "txt") & ",  "               '28
                    Sql = Sql & noquote(Request("Sav_Modifiee_le_" & I), "date") & ", "                 '29
                    Sql = Sql & noquote(Request("Sav_Action_" & I), "num")                                  '30
                    Sql = Sql & ");"
                   'Response.Write sql
                   'Response.End
                      con.Execute Sql
             End If
      End If
            
          If Session("Supprimer") = "True" Then
              If Request("Nomenclature_" & I) = "Delete" Then
                Sql = "DELETE  FROM T_Nomenclature "
                Sql = Sql & "WHERE T_Nomenclature.Id=" & Request("id_" & I) & ";"
                 con.Execute Sql
              End If
         End If
    Next

'**************************************************************************************************************
' Mis à jour de la table SousEnsembles
' MASSE total (KG)
' Nb lignes sans masse
' Nb total de lignes
'**************************************************************************************************************
    Sql = "SELECT MyFROM.Id, MyFROM.SommeDeMasse "
    Sql = Sql & "FROM (SELECT T_SousEnsembles.Id, T_SousEnsembles.flag, Count(T_Nomenclature.Masse) AS CompteDeMasse,  "
    Sql = Sql & "Sum(T_Nomenclature.Masse) AS SommeDeMasse FROM T_SousEnsembles INNER JOIN T_Nomenclature ON T_SousEnsembles.Id =  "
    Sql = Sql & "T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] WHERE T_Nomenclature.Masse>0 and  "
    Sql = Sql & "T_SousEnsembles.flag=" & noquote(Session("NumJob"), "txt") & " and T_Nomenclature.Modifiee_le IS null GROUP  "
    Sql = Sql & "BY T_SousEnsembles.Id, T_SousEnsembles.flag) AS MyFROM INNER JOIN T_SousEnsembles ON MyFROM.Id =  "
    Sql = Sql & "T_SousEnsembles.Id;"
    Set Rs_Maj_S_Ens = con.Execute(Sql)
    While Rs_Maj_S_Ens.EOF = False
    Sql = "UPDATE T_SousEnsembles SET T_SousEnsembles.[MASSE total (KG)] = " & Replace("" & Rs_Maj_S_Ens("SommeDeMasse"), ",", ".") & " "
    Sql = Sql & "WHERE T_SousEnsembles.Id=" & Rs_Maj_S_Ens("id") & ";"
'Response.Write sql
    'Response.End
        con.Execute Sql
        Rs_Maj_S_Ens.MoveNext
    Wend
    
   Sql = "SELECT T_SousEnsembles.Id, T_SousEnsembles.flag, Count(T_Nomenclature.Masse) AS CompteDeMasse "
    Sql = Sql & "FROM T_SousEnsembles INNER JOIN T_Nomenclature ON T_SousEnsembles.Id =  "
    Sql = Sql & "T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] "
    Sql = Sql & "WHERE T_Nomenclature.Modifiee_le Is Null  "
    Sql = Sql & "AND T_SousEnsembles.flag=" & noquote(Session("NumJob"), "txt") & " "
    Sql = Sql & "GROUP BY T_SousEnsembles.Id, T_SousEnsembles.flag;"
 Set Rs_Maj_S_Ens = con.Execute(Sql)
    While Rs_Maj_S_Ens.EOF = False
    Sql = "UPDATE T_SousEnsembles SET  T_SousEnsembles.[Nb total de lignes] = " & Replace("" & Rs_Maj_S_Ens("CompteDeMasse"), ",", ".") & " "
    Sql = Sql & "WHERE T_SousEnsembles.Id=" & Rs_Maj_S_Ens("id") & ";"
'Response.Write sql
    'Response.End
        con.Execute Sql
        Rs_Maj_S_Ens.MoveNext
    Wend
   
   Sql = "SELECT T_SousEnsembles.Id, T_SousEnsembles.flag, Count(T_Nomenclature.Masse) AS CompteDeMasse "
    Sql = Sql & "FROM T_SousEnsembles INNER JOIN T_Nomenclature ON T_SousEnsembles.Id =  "
    Sql = Sql & "T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] "
    Sql = Sql & "WHERE (T_Nomenclature.Masse=0 Or T_Nomenclature.Masse Is Null)  "
    Sql = Sql & "AND T_Nomenclature.Modifiee_le Is Null  "
    Sql = Sql & "AND T_SousEnsembles.flag=" & noquote(Session("NumJob"), "txt") & " "
    Sql = Sql & "GROUP BY T_SousEnsembles.Id, T_SousEnsembles.flag;"

 Set Rs_Maj_S_Ens = con.Execute(Sql)
    While Rs_Maj_S_Ens.EOF = False
    Sql = "UPDATE T_SousEnsembles SET  T_SousEnsembles.[Nb lignes sans masse] = " & Replace("" & Rs_Maj_S_Ens("CompteDeMasse"), ",", ".") & " "
    Sql = Sql & "WHERE T_SousEnsembles.Id=" & Rs_Maj_S_Ens("id") & ";"
'Response.Write sql
    'Response.End
        con.Execute Sql
        Rs_Maj_S_Ens.MoveNext
    Wend
 '**************************************************************************************************************
   
    con.Execute "UPDATE T_SousEnsembles SET T_SousEnsembles.flag = Null WHERE T_SousEnsembles.flag='" & Session("NumJob") & "';"
    
     Sql = "UPDATE T_Nomenclature SET T_Nomenclature.Flag = Null "
                Sql = Sql & "WHERE T_Nomenclature.Flag=" & noquote(Session("NumJob"), "txt") & ";"
                con.Execute Sql
End If

Response.Write vbCrLf & "<SCRIPT langage='JAVASCRIPT'>"
    
    Response.Write "function ModifierRechereche(){" & vbCrLf
    Response.Write "this.document." & FormName & ".ID__SousEnsembles.value='';"
   Response.Write "this.document." & FormName & ".action='Default.asp?Logistique_Page=11'" & vbCrLf
  Response.Write "document." & FormName & ".submit();"
  Response.Write "}" & vbCrLf
  
Response.Write "function NaviguePage(NumPage){" & vbCrLf

Response.Write "this.document." & FormName & ".action='Default.asp?Logistique_Page=4&ID=" & Request("id") & "&NEXTPAGE=' +NumPage " & vbCrLf
Response.Write vbCrLf & "document." & FormName & ".submit();"
Response.Write "}" & vbCrLf
  
Response.Write "function valider3(){" & vbCrLf
'Response.Write "this.document." & FormName & ".Id_Projects_NumAffaire.value=Id_Projects_NumAffaire;"
Response.Write "this.document." & FormName & ".Pièce.value='';"
Response.Write "this.document." & FormName & ".id_Vehicule.value='';"
Response.Write "this.document." & FormName & ".action='Default.asp?Logistique_Page=8" & MyId & "&NEXTPAGE=" & Request("NEXTPAGE") & "'" & vbCrLf
Response.Write vbCrLf & "document." & FormName & ".submit();"
Response.Write "}" & vbCrLf
    Response.Write "function valider2(){" & vbCrLf
    If Request("id") <> "" Then
MyId = "&ID=" & Request("id")
Else
MyId = ""
End If
    Response.Write "var i=this.document." & FormName & ".Sav_NbLigne.value*1;" & vbCrLf
    'Response.Write "alert(this.document." & FormName & ".Sav_NbLigne.value);"
     'Response.Write "alert(this.document." & FormName & ".NbLigne.value);"
'    Response.Write "while (i<this.document." & FormName & ".NbLigne.value){" & vbCrLf
'    Response.Write "i++;" & vbCrLf
'    Response.Write "this.document." & FormName & ".elements['Sav_Action_' + i].value=this.document." & FormName & ".elements['Action_' + i].value;" & vbCrLf
'
'    Response.Write " if (this.document." & FormName & ".elements['Annuler_' + i].checked){"
'     Response.Write "this.document." & FormName & ".elements['Sav_Annuler_' + i].value='Activer';"
'     Response.Write "}else{"
'     Response.Write "this.document." & FormName & ".elements['Sav_Annuler_' + i].value='';"
'
'    Response.Write "}"
'    Response.Write "this.document." & FormName & ".elements['Sav_ID_SousEnsembles_Sous-Ensseble_' + i].value=this.document." & FormName & ".elements['ID_SousEnsembles_Sous-Ensseble_' + i].value;" & vbCrLf
''    Response.Write "this.document." & FormName & ".elements['Sav_Désination FE_' + i].value=this.document." & FormName & ".elements['Désination FE_' + i].value;" & vbCrLf
'
''Response.Write "this.document." & FormName & ".elements['Sav_Type_' + i].value=this.document." & FormName & ".elements['Type_' + i].value;" & vbCrLf
'Response.Write "this.document." & FormName & ".elements['Sav_Pièce_' + i].value=this.document." & FormName & ".elements['Pièce_' + i].value;" & vbCrLf
'Response.Write "this.document." & FormName & ".elements['Sav_Indice_' + i].value=this.document." & FormName & ".elements['Indice_' + i].value;" & vbCrLf
'Response.Write "this.document." & FormName & ".elements['Sav_Commantaire_' + i].value=this.document." & FormName & ".elements['Commantaire_' + i].value;" & vbCrLf
'
'Response.Write "this.document." & FormName & ".elements['Sav_Designation_Piece_' + i].value=this.document." & FormName & ".elements['Designation_Piece_' + i].value;" & vbCrLf
''Response.Write "this.document." & FormName & ".elements['Sav_Famille/FE/V_' + i].value=this.document." & FormName & ".elements['Famille/FE/V_' + i].value;" & vbCrLf
'
''Response.Write "this.document." & FormName & ".elements['Sav_Num_appro_ S_' + i].value=this.document." & FormName & ".elements['Num_appro_ S_' + i].value;" & vbCrLf
'
'Response.Write "this.document." & FormName & ".elements['Sav_Quantite_' + i].value=this.document." & FormName & ".elements['Quantite_' + i].value;" & vbCrLf
'
''Response.Write "this.document." & FormName & ".elements['Sav_Bon_sortie_' + i].value=this.document." & FormName & ".elements['Bon_sortie_' + i].value;" & vbCrLf
''Response.Write "this.document." & FormName & ".elements['Sav_Observation_ou_magasin_' + i].value=this.document." & FormName & ".elements['Observation_ou_magasin_' + i].value;" & vbCrLf
''
''Response.Write "this.document." & FormName & ".elements['Sav_Num_appro_' + i].value=this.document." & FormName & ".elements['Num_appro_' + i].value;" & vbCrLf
''
''Response.Write "this.document." & FormName & ".elements['Sav_Masse_' + i].value=this.document." & FormName & ".elements['Masse_' + i].value;" & vbCrLf
''Response.Write "this.document." & FormName & ".elements['Sav_Commandee_le_' + i].value=this.document." & FormName & ".elements['Commandee_le_' + i].value;" & vbCrLf
''Response.Write "this.document." & FormName & ".elements['Sav_Collectee_le_' + i].value=this.document." & FormName & ".elements['Collectee_le_' + i].value;" & vbCrLf
''Response.Write "this.document." & FormName & ".elements['Sav_Expedee_le_' + i].value=this.document." & FormName & ".elements['Expedee_le_' + i].value;" & vbCrLf
'Response.Write "this.document." & FormName & ".elements['Sav_Qt_reçue_' + i].value=this.document." & FormName & ".elements['Qt_reçue_' + i].value;" & vbCrLf
'Response.Write "this.document." & FormName & ".elements['Sav_Reste_a_recevoir_' + i].value=this.document." & FormName & ".elements['Reste_a_recevoir_' + i].value;" & vbCrLf
'Response.Write "this.document." & FormName & ".elements['Sav_Reaffectee_a_' + i].value=this.document." & FormName & ".elements['Reaffectee_a_' + i].value;" & vbCrLf
'Response.Write "this.document." & FormName & ".elements['Sav_Delais_' + i].value=this.document." & FormName & ".elements['Delais_' + i].value;" & vbCrLf
'Response.Write "this.document." & FormName & ".elements['Sav_En_provenance_de_' + i].value=this.document." & FormName & ".elements['En_provenance_de_' + i].value;" & vbCrLf
'Response.Write "this.document." & FormName & ".elements['Sav_QT_Retournee_' + i].value=this.document." & FormName & ".elements['QT_Retournee_' + i].value;" & vbCrLf
'      Response.Write "}" & vbCrLf
   
   Response.Write "this.document." & FormName & ".action='Default.asp?Logistique_Page=4" & MyId & "&NEXTPAGE=" & Request("NEXTPAGE") & "'" & vbCrLf
  Response.Write "this.document." & FormName & ".submit();" & vbCrLf
Response.Write "}" & vbCrLf


    Response.Write vbCrLf & "function PageSuvante(Logistique_Page,id,id_Vehicule,Pièce,ID__SousEnsembles,Id_Projects_NumAffaire){"

    Response.Write vbCrLf & "document." & FormName & ".Logistique_Page.value=Logistique_Page;"
    Response.Write vbCrLf & "document." & FormName & ".id.value=id;"
    Response.Write vbCrLf & "document." & FormName & ".id_Vehicule.value=id_Vehicule;"
    Response.Write vbCrLf & "document." & FormName & ".Pièce.value=Pièce;"
    Response.Write vbCrLf & "document." & FormName & ".ID__SousEnsembles.value=ID__SousEnsembles;"
    Response.Write vbCrLf & "document." & FormName & ".Id_Projects_NumAffaire.value=Id_Projects_NumAffaire;"
'
    Response.Write vbCrLf & "document." & FormName & ".action='Default.asp';"
    Response.Write vbCrLf & "document." & FormName & ".submit();"



    Response.Write vbCrLf & "}"
    Response.Write vbCrLf & "</SCRIPT>"
Response.Write "<script language='javascript' >" & vbCrLf
Response.Write "function RetourneSoldSurQts2(Qts,Reste_recevoir){"
Response.Write "Qts.value=MyTrim(Qts.value);"
Response.Write "if(Qts.value!=''){"
Response.Write "IsFloat(Qts);"
Response.Write "}"
Response.Write "Reste_recevoir.value=Qts.value;"
Response.Write "}"
Response.Write "function RetourneSoldSurQts(Qts,SaveQts,Reste_recevoir,Deja_recevoir){"

Response.Write "if (IsFloat(Qts)==false){" & vbCrLf
Response.Write "Qts.value=SaveQts;"
Response.Write "}"
Response.Write "MyValeur=funMath(SaveQts,Reste_recevoir.value,'-',false);" & vbCrLf

Response.Write "if (Qts.value<MyValeur){"
Response.Write "alert('La Quantité : ' + Qts.value +'\nEst inférieur\nA la Qts déjà réceptionnée :'+MyValeur);"
Response.Write "Qts.value=SaveQts;"
Response.Write "return;"
Response.Write "}"
Response.Write "MyValeur=funMath(Qts.value,Reste_recevoir.value,'-',false);" & vbCrLf
Response.Write "Reste_recevoir.value=funMath(MyValeur,Reste_recevoir.value,'+',false);" & vbCrLf



Response.Write "}"
Response.Write "function RetourneSold(Qts,Qt_reçue,QT_Retournee,Reste_recevoir,Deja_recevoir ){" & vbCrLf
Response.Write "IsFloat(Qt_reçue)" & vbCrLf
'Response.Write "IsFloat(Qt_Reaffectee);" & vbCrLf
'Response.Write "IsFloat(Qt_Reaffectee);" & vbCrLf
Response.Write "IsFloat(QT_Retournee);" & vbCrLf
Response.Write "IsFloat(QT_Retournee);" & vbCrLf
'*********************************************************************************
'QTS Recu
'**********************************************************************************
Response.Write "var MyValeur=0;" & vbCrLf
Response.Write "MyValeur=funMath(Deja_recevoir,Qt_reçue.value,'-',true);" & vbCrLf
Response.Write "if (MyValeur=='Err'){" & vbCrLf
Response.Write "Qt_reçue.value=0;" & vbCrLf
Response.Write "Reste_recevoir.value=Deja_recevoir;" & vbCrLf
Response.Write "QT_Retournee.value=0;" & vbCrLf
Response.Write "Qt_reçue.onchange();" & vbCrLf
Response.Write "Qt_reçue.focus();"
Response.Write "return;"
Response.Write "}" & vbCrLf

'*********************************************************************************
'QTS Retournee
'**********************************************************************************
Response.Write "MyValeur=funMath(MyValeur,QT_Retournee.value,'+',true);" & vbCrLf
Response.Write "if (MyValeur>Qts.value){" & vbCrLf
Response.Write "alert('La QT Retournée : ' + QT_Retournee.value +'\nEst supérieur\nA la Quantitée : '+ Qts.value );" & vbCrLf
Response.Write "Qt_reçue.value=0;" & vbCrLf
Response.Write "Reste_recevoir.value=Deja_recevoir;" & vbCrLf
Response.Write "QT_Retournee.value=0;" & vbCrLf
Response.Write "QT_Retournee.onchange();" & vbCrLf
Response.Write "QT_Retournee.focus();"
Response.Write "return;"
Response.Write "}" & vbCrLf


'*********************************************************************************
'QTS Reaffectee
'**********************************************************************************
'Response.Write "MyValeur=funMath(MyValeur,Qt_Reaffectee.value,'+',true);" & vbCrLf
'Response.Write "if (MyValeur>Qts.value){" & vbCrLf
'Response.Write "alert('La Qt Réaffectée : ' +Qt_Reaffectee.value +'\nEst supérieur\nA la Quantitée : '+ Qts.value );" & vbCrLf
'Response.Write "MyValeur=funMath(MyValeur,Qt_Reaffectee.value,'-',true);" & vbCrLf
'Response.Write "Qt_Reaffectee.value=0;" & vbCrLf
'Response.Write "Qt_Reaffectee.onchange();" & vbCrLf
'Response.Write "}" & vbCrLf

Response.Write "Reste_recevoir.value=MyValeur;" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "</script>" & vbCrLf
    Response.Write "<form name='" & FormName & "' method='post'>" & vbCrLf
     Response.Write "<input name='MAJ' type='hidden' value='OK'/>" & vbCrLf
     Response.Write "<input name='Updatable' type='hidden' value='OK'/>" & vbCrLf
      Response.Write "<input name='FormName' type='hidden' value='" & FormName & "'/>" & vbCrLf
      
    
      
      Dim L
     Dim C
    
      'Response.Write sql
        'Response.End
    
    
       
     
    
   
      Response.Write "<table border='0'  width='100%'>" & vbCrLf
       If Request("nbSelect") > 0 Then
            Response.Write "<tr>" & vbCrLf
            Response.Write "         <td >" & RecapSersh & "</td>" & vbCrLf
            
            Response.Write "</tr>" & vbCrLf
         End If
         Response.Write "<tr >" & vbCrLf
         Response.Write "<td class='TitreTableau'>Liste des pièces :</td>" & vbCrLf
         Response.Write "</tr>" & vbCrLf
        
          
         Response.Write "<tr>" & vbCrLf
          Response.Write "<TD width='100%'><table border='0'  width='100%'><tr>"
        
    If Session("Updatable") = "True" Then
         
        Response.Write "<TD width='2' valign='top'>" & AfficheBounton("Valider", 1, "'javascript:valider2();'") & "</td>" & vbCrLf
      End If
       Response.Write "<TD width='2' valign='top'>" & AfficheBounton("Retour", 8, """javascript:valider3();""") & "</td>" & vbCrLf
      Response.Write "<TD>&nbsp;</td><TD>&nbsp;</td>" & vbCrLf
        Response.Write "<TD width='4%' valign='top'>Filtrés sur véhicule :" & vbCrLf
        Sql = "SELECT t_Vehicule.Id,  Projet,  Véhicule FROM t_Vehicule ORDER BY t_Vehicule.Véhicule;"
        
         Sql = "SELECT     t_Vehicule.Id, T_Projects.NumAffaire, t_Vehicule.Véhicule "
        Sql = Sql & "FROM t_Vehicule INNER JOIN "
        Sql = Sql & "T_Projects ON t_Vehicule.Id_Projects_NumAffaire = T_Projects.Id "
        Sql = Sql & "ORDER BY t_Vehicule.Véhicule"
        Set RsFiltre = OpenRecordSet(Sql, con)

       Response.Write "<br>" & CreatListe(RsFiltre, Request("FiltreVehicule"), "Id", "1;2", "FiltreVehicule", " onChange='valider2();' ") & "</td>" & vbCrLf
      RsFiltre.Close
      Set RsFiltre = Nothing
      
     Response.Write "<TD width='4%' valign='top'>Masquer l´historique:&nbsp;" & vbCrLf
     If Trim("" & Request("MasquerHistoPiece")) <> "" Then
                Checked = "  Checked "
            Else
                Checked = ""
       End If
      Response.Write "<input name='MasquerHistoPiece' type='checkbox' value='Modifiee_le=Null' " & Checked & " onclick='valider2();'>"
     Response.Write "</td>" & vbCrLf
      Response.Write "<TD width='100%' valign='top'>&nbsp;</td>" & vbCrLf
     Response.Write "</Tr></table></td>"
    
       Response.Write "</tr> " & vbCrLf
Response.Write "<tr><td> <table border='1' width='100%'>" & vbCrLf
    L = 0
    L2 = 0
   
    
    

    
    sqlListe = "SELECT DISTINCT t_Vehicule.Id,  t_Vehicule.Id,t_Vehicule.Projet,t_Vehicule.Véhicule) "
    sqlListe = sqlListe & "FROM (((T_Nomenclature LEFT JOIN ((T_Nomenclature AS T_Nomenclature_1 LEFT  "
    sqlListe = sqlListe & "JOIN T_Lien_SousEnssemble_Nomenclature AS T_Lien_SousEnssemble_Nomenclature_1 ON  "
    sqlListe = sqlListe & "T_Nomenclature_1.Id = T_Lien_SousEnssemble_Nomenclature_1.Id_Nomenclature)  "
    sqlListe = sqlListe & "LEFT JOIN T_lien_Veicule_Souensembles AS T_lien_Veicule_Souensembles_1  "
    sqlListe = sqlListe & "ON T_Lien_SousEnssemble_Nomenclature_1.ID_SousEnsembles = T_lien_Veicule_Souensembles_1.ID__SousEnsembles)  "
    sqlListe = sqlListe & "ON (T_Nomenclature.Indice = T_Nomenclature_1.Indice) AND (T_Nomenclature.Pièce = T_Nomenclature_1.Pièce))  "
    sqlListe = sqlListe & "LEFT JOIN T_Lien_SousEnssemble_Nomenclature ON T_Nomenclature.Id = T_Lien_SousEnssemble_Nomenclature.Id_Nomenclature)  "
    sqlListe = sqlListe & "LEFT JOIN T_lien_Veicule_Souensembles ON T_Lien_SousEnssemble_Nomenclature.ID_SousEnsembles =  "
    sqlListe = sqlListe & "T_lien_Veicule_Souensembles.ID__SousEnsembles) LEFT JOIN t_Vehicule ON T_lien_Veicule_Souensembles_1.ID_Vehicule =  "
    sqlListe = sqlListe & "t_Vehicule.Id "
    sqlListe = sqlListe & "WHERE   T_Nomenclature_1.Modifiee_le Is Null "
    sqlListe = sqlListe & "AND T_lien_Veicule_Souensembles_1.ID_Vehicule<>[T_lien_Veicule_Souensembles].[ID_Vehicule]  "
    sqlListe = sqlListe & "AND T_Nomenclature.Id="
    
    sqlListe = "SELECT DISTINCT T_Vehicule.Id, t_Vehicule.Projet, t_Vehicule.Véhicule "
    sqlListe = sqlListe & "FROM (((T_Nomenclature LEFT JOIN T_Nomenclature AS T_Nomenclature_1 ON  "
    sqlListe = sqlListe & "(T_Nomenclature.Indice = T_Nomenclature_1.Indice) AND (T_Nomenclature.Pièce = T_Nomenclature_1.Pièce))  "
    sqlListe = sqlListe & "LEFT JOIN T_SousEnsembles ON T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] = T_SousEnsembles.Id) LEFT  "
    sqlListe = sqlListe & "JOIN t_Vehicule ON T_SousEnsembles.ID_Vehicule_Véhicule = t_Vehicule.Id) LEFT JOIN T_SousEnsembles AS T_SousEnsembles_1  "
    sqlListe = sqlListe & "ON T_Nomenclature_1.[ID_SousEnsembles_Sous-Ensseble] = T_SousEnsembles_1.Id "
   

    
    
    

        
        'Response.Write "<input name='S_Enssemble__" & L2 & "' type='hidden' value='" & rs("Id") & "'/>"
        'Response.Write "<textarea class='Champ' name='TtxtS_Enssemble__" & L2 & "' rows='4' cols='150'>" & rs("Designation S/Enssemble") & "</textarea>" & vbcrlf
       
       
        Sql = "SELECT T_Nomenclature.Modification,T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble], T_Nomenclature.Id, T_Nomenclature.Id_T_Actions_Libellé,  "
        Sql = Sql & "T_Nomenclature.Modifier_Par, T_Nomenclature.Activer, T_Projects.Projet, T_Nomenclature.[Désination FE],  "
        Sql = Sql & "T_Nomenclature.Type, T_Nomenclature.Pièce, T_Nomenclature.Indice, T_Nomenclature.Commantaire,  "
        Sql = Sql & "T_Nomenclature.Designation_Piece, T_Nomenclature.[Famille/FE/V], T_Nomenclature.[Num_appro_ S],  "
        Sql = Sql & "T_Nomenclature.Quantite, T_Nomenclature.Bon_sortie, T_Nomenclature.Observation_ou_magasin,  "
        Sql = Sql & "T_Nomenclature.Num_appro, T_Nomenclature.Masse, T_Nomenclature.Commandee_le, T_Nomenclature.Collectee_le,  "
        Sql = Sql & "T_Nomenclature.Expedee_le, T_Nomenclature.Reçue_le, T_Nomenclature.Qt_reçue, T_Nomenclature.Reste_a_recevoir,  "
        Sql = Sql & "T_Nomenclature.N_de_BL, T_Nomenclature.Reaffectee_a, T_Nomenclature.Qt_Reaffectee, T_Nomenclature.En_provenance_de,  "
        Sql = Sql & "T_Nomenclature.Retournee_le, T_Nomenclature.QT_Retournee, T_Nomenclature.Num_de_retour, T_Nomenclature.Modifiee_le,  "
        Sql = Sql & "T_Nomenclature.Id_T_Delais_Libellé "
                Sql = Sql & "FROM t_Vehicule INNER JOIN "
        Sql = Sql & "T_Projects ON t_Vehicule.Id_Projects_NumAffaire = T_Projects.Id LEFT OUTER JOIN "
        Sql = Sql & "T_SousEnsembles ON t_Vehicule.Id = T_SousEnsembles.ID_Vehicule_Véhicule LEFT OUTER JOIN "
        Sql = Sql & "T_Nomenclature ON T_SousEnsembles.Id = T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] "

        
       MyFilter = RetourneCloseWhere()
          If Request("nbSelect") > 0 Then
            Sql = Sql & MyCloseWhere()
             If Trim("" & MyFilter) <> "" Then
                Sql = Sql & " AND (" & MyFilter & ") "
             End If
          Else
            If Trim("" & MyFilter) <> "" Then
                Sql = Sql & " WHERE " & MyFilter
             End If
        End If
         If Trim("" & Request("FiltreVehicule")) <> "" Then
                    If InStr(1, UCase(Sql), "WHERE") <> 0 Then
                Sql = Sql & " AND t_Vehicule.Id=" & Request("FiltreVehicule") & " "
                
            Else
                 Sql = Sql & " WHERE t_Vehicule.Id=" & Request("FiltreVehicule") & " "
            End If
      End If
'        Sql = Sql & "GROUP BY T_Nomenclature.Modification,T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble], T_Nomenclature.Id, T_Nomenclature.Id_T_Actions_Libellé,  "
'        Sql = Sql & "T_Nomenclature.Modifier_Par, T_Nomenclature.Activer, t_Vehicule.Projet, T_Nomenclature.[Désination FE],  "
'        Sql = Sql & "T_Nomenclature.Type, T_Nomenclature.Pièce, T_Nomenclature.Indice, T_Nomenclature.Commantaire,  "
'        Sql = Sql & "T_Nomenclature.Designation_Piece, T_Nomenclature.[Famille/FE/V], T_Nomenclature.[Num_appro_ S],  "
'        Sql = Sql & "T_Nomenclature.Quantite, T_Nomenclature.Bon_sortie, T_Nomenclature.Observation_ou_magasin,  "
'        Sql = Sql & "T_Nomenclature.Num_appro, T_Nomenclature.Masse, T_Nomenclature.Commandee_le, T_Nomenclature.Collectee_le,  "
'        Sql = Sql & "T_Nomenclature.Expedee_le, T_Nomenclature.Reçue_le, T_Nomenclature.Qt_reçue, T_Nomenclature.Reste_a_recevoir,  "
'        Sql = Sql & "T_Nomenclature.N_de_BL, T_Nomenclature.Reaffectee_a, T_Nomenclature.Qt_Reaffectee, T_Nomenclature.En_provenance_de,  "
'        Sql = Sql & "T_Nomenclature.Retournee_le, T_Nomenclature.QT_Retournee, T_Nomenclature.Num_de_retour, T_Nomenclature.Modifiee_le,  "
'        Sql = Sql & "T_Nomenclature.Designation_Piece, T_Nomenclature.Id, T_Nomenclature.Id_T_Delais_Libellé "
        Sql = Sql & "ORDER BY  T_Nomenclature.[Désination FE], T_Nomenclature.Id DESC;"

       Response.Write "<input name='id' type='hidden' value='" & Request("id") & "' />"
       Response.Write "<input name='id_Vehicule' type='hidden' value='" & Request("id_Vehicule") & "' />"
       Response.Write "<input name='Pièce' type='hidden' value='" & Request("Pièce") & "' />"
       Response.Write "<input name='ID__SousEnsembles' type='hidden' value='" & Request("ID__SousEnsembles") & "' />"
        Response.Write "<input name='Id_Projects_NumAffaire' type='hidden' value='" & Request("Id_Projects_NumAffaire") & "' />"
       
  
 
       'Response.Write sql
        'Response.End
        
        'sql=sql & "WHERE T_Nomenclature.Id_SousEnsemble=" & rs("id") & ";"
      
        Set RsNomenclature = OpenRecordSet(Sql, con)
        SqlAction = "SELECT T_Actions.* FROM T_Actions ORDER BY T_Actions.Id;"

 Set RsSqlAction = OpenRecordSet(SqlAction, con)
 T_RsSqlAction = RsSqlAction.GetRows
  RsSqlAction.Close
  Set RsSqlAction = Nothing
  T_RsSqlAction = T_CreatListe(T_RsSqlAction, "", 0, "1;2", "T_RsSqlAction", "")
  
  
  SqlDelais = "SELECT T_Delais.* FROM T_Delais ORDER BY T_Delais.Id;"
 Set RsSqlDelais = OpenRecordSet(SqlDelais, con)
 T_RsSqlDelais = RsSqlDelais.GetRows
  RsSqlDelais.Close
  Set RsSqlDelais = Nothing
  T_RsSqlDelais = T_CreatListe(T_RsSqlDelais, "", 0, "1", "T_RsSqlDelais", "")
       
'        Response.Write "         <td align='center' bgcolor='#CCCCCC'>Type</td>" & vbCrLf
'        Response.Write "<td align='center' bgcolor='#CCCCCC'>Désignation FE</td>" & vbCrLf
'        Response.Write "<td align='center' bgcolor='#CCCCCC'>Famille/FE/V</td>" & vbCrLf
'        Response.Write "<td align='center' bgcolor='#CCCCCC'> N° d'appro S </td>"
'       Response.Write "<td align='center' bgcolor='#CCCCCC'>Bon_sortie</td>" & vbCrLf
'        Response.Write "<td align='center' bgcolor='#CCCCCC'>Observation ou magasin </td>" & vbCrLf
'        Response.Write "<td align='center' bgcolor='#CCCCCC'>N° d'appro  </td>" & vbCrLf
'        Response.Write "<td align='center' bgcolor='#CCCCCC'>Masse </td>" & vbCrLf
        
        
        
        
          
        Sql = "SELECT T_SousEnsembles.Id, [Véhicule],   [code] "
        Sql = Sql & "FROM t_Vehicule INNER JOIN T_SousEnsembles ON t_Vehicule.Id = T_SousEnsembles.ID_Vehicule_Véhicule;"

          Set RsT_SousEnsembles = OpenRecordSet(Sql, con)
          T_RsT_SousEnsembles = RsT_SousEnsembles.GetRows
          RsT_SousEnsembles.Close
          Set RsT_SousEnsembles = Nothing
          T_RsT_SousEnsembles = T_CreatListe(T_RsT_SousEnsembles, "", 0, "1;2", "T_RsT_SousEnsembles", "") & "</td>" & vbCrLf
        
         RsNomenclature.Filter = "" & Request("MasquerHistoPiece")
         Response.Write TitreNomenclatur
         If Session("Nouveau") = "True" Then
            If Trim("" & MyFilter) = "" Then
                L = L + 1
           
                    sqlListeWhere = "WHERE T_Nomenclature.Id=null;"
                  
                    Set Rsvehicule = OpenRecordSet(sqlListe & sqlListeWhere, con)

'---------------------------------------------------------------------------------------------
' New 1
'---------------------------------------------------------------------------------------------
   
                    Response.Write "<tr>"
                    

                    
                        Response.Write "         <td align='center'>New</td>" & vbCrLf
                        Response.Write "         <td align='center'>---" & vbCrLf
                        Response.Write "         <input Name='Sav_Annuler_" & L & "' type='hidden' value='Activer' />"
                        Response.Write "         <input Name='Annuler2_" & L & "' type='hidden' value='Activer' /> </td>" & vbCrLf
                        Response.Write "         <td align='center'>---</td>" & vbCrLf
                        Response.Write "<td>" & Replace(T_RsT_SousEnsembles, "name='T_RsT_SousEnsembles'", "name='Sav_ID_SousEnsembles_Sous-Ensseble_" & L & "'") & "</td>" & vbCrLf
                         
                        Response.Write "<td><input Name='Sav_Designation_Piece_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
                        Response.Write "             <td><input Name='Sav_Pièce_" & L & "' type='text' class='Champ'  value=''/> </td>" & vbCrLf
                        Response.Write "             <td><input Name='Sav_Indice_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
                        Response.Write "<td><input Name='Sav_Quantite_" & L & "' type='text' class='Champ'  value='' onchange='RetourneSoldSurQts2(this.document." & FormName & ".Sav_Quantite_" & L & ",this.document." & FormName & ".Sav_Reste_a_recevoir_" & L & ");'/></td>" & vbCrLf
                        Response.Write "<td><input Name='Sav_Reste_a_recevoir_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
                        Response.Write "<td>" & Replace(T_RsSqlAction, "name='T_RsSqlAction'", "name='Sav_Action_" & L & "'") & "</td>" & vbCrLf
                        Response.Write "<td>" & Replace(T_RsSqlDelais, "name='T_RsSqlDelais'", "name='Sav_Delais_" & L & "'") & "</td>" & vbCrLf
                        
                        Response.Write "<td>---</td>" & vbCrLf
                        Response.Write "<td>---</td>" & vbCrLf
                        Response.Write "<td>---</td>" & vbCrLf
                        Response.Write "<td>---</td>" & vbCrLf
                        Response.Write "<td>---</td>" & vbCrLf
                        Response.Write "<td>---</td>" & vbCrLf
                        Response.Write "<td>---</td>" & vbCrLf
                        Response.Write "<td>---</td>" & vbCrLf
                         Response.Write "<td>---</td>" & vbCrLf
                        Response.Write "<td>---</td>" & vbCrLf
                        Response.Write "<td>---</td>" & vbCrLf
'                        Response.Write "         <td><input Name='Sav_Type_" & L & "' type='text' class='Champ'  value='' /></td>" & vbCrLf
'                        Response.Write " <td><input Name='Id_" & L & "' type='hidden' value=''/><input Name='Sav_Désination FE_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "         " & vbCrLf
'                        Response.Write "<td><input Name='Sav_Famille/FE/V_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Sav_Num_appro_ S_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Sav_Bon_sortie_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Sav_Observation_ou_magasin_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Sav_Num_appro_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Sav_Masse_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Sav_Modifiee_le_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Modifier_Par" & L & "' type='text' class='Champ'  value='' disabled /></td>" & vbCrLf
                            Response.Write "<td><input Name='Sav_Commantaire_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "        <td><textarea class='Champ' name='Sav_Commantaire_" & L & "' rows='2' cols='20' ></textarea> </td>" & vbCrLf
                   
                    Response.Write "</tr>"
'
         
                    L = L + 1
                    Response.Write "<tr>"
'---------------------------------------------------------------------------------------------
' New 2
'---------------------------------------------------------------------------------------------
   
                    Response.Write "<tr>"
                    
                        Response.Write "         <td align='center'>New</td>" & vbCrLf
                        Response.Write "         <td align='center'>---" & vbCrLf
                        Response.Write "         <input Name='Sav_Annuler_" & L & "' type='hidden' value='Activer' />"
                        Response.Write "         <input Name='Annuler2_" & L & "' type='hidden' value='Activer' /> </td>" & vbCrLf
                        Response.Write "         <td align='center'>---</td>" & vbCrLf
                        Response.Write "<td>" & Replace(T_RsT_SousEnsembles, "name='T_RsT_SousEnsembles'", "name='Sav_ID_SousEnsembles_Sous-Ensseble_" & L & "'") & "</td>" & vbCrLf
                       Response.Write "<td><input Name='Sav_Designation_Piece_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
                        Response.Write "             <td><input Name='Sav_Pièce_" & L & "' type='text' class='Champ'  value=''/> </td>" & vbCrLf
                        Response.Write "             <td><input Name='Sav_Indice_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
                        Response.Write "<td><input Name='Sav_Quantite_" & L & "' type='text' class='Champ'  value='' onchange='RetourneSoldSurQts2(this.document." & FormName & ".Sav_Quantite_" & L & ",this.document." & FormName & ".Sav_Reste_a_recevoir_" & L & ");'/></td>" & vbCrLf
                        Response.Write "<td><input Name='Sav_Reste_a_recevoir_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
                        Response.Write "<td>" & Replace(T_RsSqlAction, "name='T_RsSqlAction'", "name='Sav_Action_" & L & "'") & "</td>" & vbCrLf
                        Response.Write "<td>" & Replace(T_RsSqlDelais, "name='T_RsSqlDelais'", "name='Sav_Delais_" & L & "'") & "</td>" & vbCrLf
                        Response.Write "<td>---</td>" & vbCrLf
                        Response.Write "<td>---</td>" & vbCrLf
                        Response.Write "<td>---</td>" & vbCrLf
                        Response.Write "<td>---</td>" & vbCrLf
                        Response.Write "<td>---</td>" & vbCrLf
                        Response.Write "<td>---</td>" & vbCrLf
                        Response.Write "<td>---</td>" & vbCrLf
                        Response.Write "<td>---</td>" & vbCrLf
                        Response.Write "<td>---</td>" & vbCrLf
                        Response.Write "<td>---</td>" & vbCrLf
                        Response.Write "<td>---</td>" & vbCrLf
                        Response.Write "<td><input Name='Sav_Commantaire_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'
'                        Response.Write "         <td><input Name='Sav_Type_" & L & "' type='text' class='Champ'  value='' /></td>" & vbCrLf
'                        Response.Write " <td><input Name='Id_" & L & "' type='hidden' value=''/><input Name='Sav_Désination FE_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "         " & vbCrLf
'                        Response.Write "<td><input Name='Sav_Famille/FE/V_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Sav_Num_appro_ S_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Sav_Bon_sortie_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Sav_Observation_ou_magasin_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Sav_Num_appro_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Sav_Masse_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Sav_Modifiee_le_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Modifier_Par" & L & "' type='text' class='Champ'  value='' disabled /></td>" & vbCrLf
'                        Response.Write "        <td><textarea class='Champ' name='Sav_Commantaire_" & L & "' rows='2' cols='20' ></textarea> </td>" & vbCrLf
                   
                    Response.Write "</tr>"
        
                    L = L + 1
                    Response.Write "<tr>"
''---------------------------------------------------------------------------------------------
'' New 3
''---------------------------------------------------------------------------------------------
'
'                    Response.Write "<tr>"
'
'                        Response.Write "         <td align='center'>New</td>" & vbCrLf
'                        Response.Write "         <td align='center'>---" & vbCrLf
'                        Response.Write "         <input Name='Sav_Annuler_" & L & "' type='hidden' value='Activer' />"
'                        Response.Write "         <input Name='Annuler2_" & L & "' type='hidden' value='Activer' /> </td>" & vbCrLf
'                        Response.Write "         <td align='center'>---</td>" & vbCrLf
'                        Response.Write "<td>" & Replace(T_RsT_SousEnsembles, "name='T_RsT_SousEnsembles'", "name='Sav_ID_SousEnsembles_Sous-Ensseble_" & L & "'") & "</td>" & vbCrLf
'                        Response.Write "<td><textarea class='Champ' name='Sav_Designation_Piece_" & L & "' rows='2' cols='20' ></textarea></td>" & vbCrLf
'                        Response.Write "             <td><input Name='Sav_Pièce_" & L & "' type='text' class='Champ'  value=''/> </td>" & vbCrLf
'                        Response.Write "             <td><input Name='Sav_Indice_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Sav_Quantite_" & L & "' type='text' class='Champ'  value='' onchange='RetourneSoldSurQts2(this.document." & FormName & ".Sav_Quantite_" & L & ",this.document." & FormName & ".Sav_Reste_a_recevoir_" & L & ");'/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Sav_Reste_a_recevoir_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td>" & Replace(T_RsSqlAction, "name='T_RsSqlAction'", "name='Sav_Action_" & L & "'") & "</td>" & vbCrLf
'                        Response.Write "<td>" & Replace(T_RsSqlDelais, "name='T_RsSqlDelais'", "name='Sav_Delais_" & L & "'") & "</td>" & vbCrLf
'                        Response.Write "<td>---</td>" & vbCrLf
'                        Response.Write "<td>---</td>" & vbCrLf
'                        Response.Write "<td>---</td>" & vbCrLf
'                        Response.Write "<td>---</td>" & vbCrLf
'                        Response.Write "<td>---</td>" & vbCrLf
'                        Response.Write "<td>---</td>" & vbCrLf
'                        Response.Write "<td>---</td>" & vbCrLf
'                        Response.Write "<td>---</td>" & vbCrLf
'                        Response.Write "         <td><input Name='Sav_Type_" & L & "' type='text' class='Champ'  value='' /></td>" & vbCrLf
'                        Response.Write " <td><input Name='Id_" & L & "' type='hidden' value=''/><input Name='Sav_Désination FE_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "         " & vbCrLf
'                        Response.Write "<td><input Name='Sav_Famille/FE/V_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Sav_Num_appro_ S_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Sav_Bon_sortie_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Sav_Observation_ou_magasin_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Sav_Num_appro_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Sav_Masse_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Sav_Modifiee_le_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
'                        Response.Write "<td><input Name='Modifier_Par" & L & "' type='text' class='Champ'  value='' disabled /></td>" & vbCrLf
'                        Response.Write "        <td><textarea class='Champ' name='Sav_Commantaire_" & L & "' rows='2' cols='20' ></textarea> </td>" & vbCrLf
'
'                    Response.Write "</tr>"
'                    L = L + 1
'                   Response.Write "<tr>"
''---------------------------------------------------------------------------------------------
'' New 4
''---------------------------------------------------------------------------------------------
'
''                    Response.Write "<tr>"
''
''                        Response.Write "         <td align='center'>New</td>" & vbCrLf
''                        Response.Write "         <td align='center'>---" & vbCrLf
''                        Response.Write "         <input Name='Sav_Annuler_" & L & "' type='hidden' value='Activer' />"
''                        Response.Write "         <input Name='Annuler2_" & L & "' type='hidden' value='Activer' /> </td>" & vbCrLf
''                        Response.Write "         <td align='center'>---</td>" & vbCrLf
''                        Response.Write "<td>" & Replace(T_RsT_SousEnsembles, "name='T_RsT_SousEnsembles'", "name='Sav_ID_SousEnsembles_Sous-Ensseble_" & L & "'") & "</td>" & vbCrLf
''                        Response.Write "<td><textarea class='Champ' name='Sav_Designation_Piece_" & L & "' rows='2' cols='20' ></textarea></td>" & vbCrLf
''                        Response.Write "             <td><input Name='Sav_Pièce_" & L & "' type='text' class='Champ'  value=''/> </td>" & vbCrLf
''                        Response.Write "             <td><input Name='Sav_Indice_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
''                        Response.Write "<td><input Name='Sav_Quantite_" & L & "' type='text' class='Champ'  value='' onchange='RetourneSoldSurQts2(this.document." & FormName & ".Sav_Quantite_" & L & ",this.document." & FormName & ".Sav_Reste_a_recevoir_" & L & ");'/></td>" & vbCrLf
''                        Response.Write "<td><input Name='Sav_Reste_a_recevoir_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
''                        Response.Write "<td>" & Replace(T_RsSqlAction, "name='T_RsSqlAction'", "name='Sav_Action_" & L & "'") & "</td>" & vbCrLf
''                        Response.Write "<td>" & Replace(T_RsSqlDelais, "name='T_RsSqlDelais'", "name='Sav_Delais_" & L & "'") & "</td>" & vbCrLf
''                        Response.Write "<td>---</td>" & vbCrLf
''                        Response.Write "<td>---</td>" & vbCrLf
''                        Response.Write "<td>---</td>" & vbCrLf
''                        Response.Write "<td>---</td>" & vbCrLf
''                        Response.Write "<td>---</td>" & vbCrLf
''                        Response.Write "<td>---</td>" & vbCrLf
''                        Response.Write "<td>---</td>" & vbCrLf
''                        Response.Write "<td>---</td>" & vbCrLf
''                        Response.Write "         <td><input Name='Sav_Type_" & L & "' type='text' class='Champ'  value='' /></td>" & vbCrLf
''                        Response.Write " <td><input Name='Id_" & L & "' type='hidden' value=''/><input Name='Sav_Désination FE_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
''                        Response.Write "         " & vbCrLf
''                        Response.Write "<td><input Name='Sav_Famille/FE/V_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
''                        Response.Write "<td><input Name='Sav_Num_appro_ S_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
''                        Response.Write "<td><input Name='Sav_Bon_sortie_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
''                        Response.Write "<td><input Name='Sav_Observation_ou_magasin_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
''                        Response.Write "<td><input Name='Sav_Num_appro_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
''                        Response.Write "<td><input Name='Sav_Masse_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
''                        Response.Write "<td><input Name='Sav_Modifiee_le_" & L & "' type='text' class='Champ'  value=''/></td>" & vbCrLf
''                        Response.Write "<td><input Name='Modifier_Par" & L & "' type='text' class='Champ'  value='' disabled /></td>" & vbCrLf
''                        Response.Write "        <td><textarea class='Champ' name='Sav_Commantaire_" & L & "' rows='2' cols='20' ></textarea> </td>" & vbCrLf
''
''                    L2 = L2 + 1
''                    Response.Write "</tr>"
'                    Set Rsvehicule = CloseRecordSet(Rsvehicule)
                    
                End If
      End If
      Response.Write "<input name='Sav_NbLigne' type='hidden' value='" & CStr(L) & "'/>"
      Dim IndexRecord As Long
      Dim Bigin As Long
      Dim Stoped As Long
      Dim currPageNum As Long
      currPageNum = Request("NEXTPAGE")
       
      If currPageNum = 0 Then currPageNum = 1
     Bigin = (currPageNum * BbParP + 1) - BbParP
    
      Stoped = Bigin + BbParP
      countRecurd = 0
       Do While RsNomenclature.EOF = False
        countRecurd = countRecurd + 1
        RsNomenclature.MoveNext
       Loop
       RsNomenclature.MoveFirst
       ModcountRecurd = countRecurd Mod BbParP
        countRecurd2 = (countRecurd - ModcountRecurd) / BbParP + 1
        If ModcountRecurd <> 0 Then countRecurd2 = countRecurd2 + 1
       Do While RsNomenclature.EOF = False
        IndexRecord = IndexRecord + 1
        If Bigin <= IndexRecord Then
            If Trim("" & RsNomenclature("Modifiee_le")) <> "" Then
                 Lname = ""
                 Disabled = " Disabled "
            Else
                Disabled = ""
                L = L + 1
                 Lname = L
        End If
        'Response.Write "Affiche = " & Disabled & "<br>"
     
        sqlListeWhere = "WHERE  [Pièce]='" & RsNomenclature("Pièce") & "' and [Indice]='" & RsNomenclature("Indice") & "'  "
        sqlListeWhere = sqlListeWhere & "AND T_Nomenclature.Id<>" & RsNomenclature("ID") & ";"
        
        sqlListeWhere = "WHERE T_Nomenclature.Modifiee_le Is Null  "
        sqlListeWhere = sqlListeWhere & "AND T_SousEnsembles_1.ID_Vehicule_Véhicule<>[T_SousEnsembles].[ID_Vehicule_Véhicule] "
        sqlListeWhere = sqlListeWhere & "AND T_Nomenclature_1.Id=" & RsNomenclature("ID") & ";"
       'Response.Write  sqlListe & sqlListeWhere
       'Response.End
     Set Rsvehicule = OpenRecordSet(sqlListe & sqlListeWhere, con)
    
    'RetournHeder "Designation_Piece"
    
        Response.Write "         <td align='center'><input Name='Nomenclature_" & Lname & "'" & RetournHeder("Supprimer", Disabled) & " value='Delete' type='checkbox' /> "
        Response.Write "         <input Name='Id_" & Lname & "' type='hidden' value='" & RsNomenclature("ID") & "' /></td>" & vbCrLf
       
        If RsNomenclature("Activer") = True Then
            Response.Write "         <td align='center'><input Name='Annuler_" & Lname & "' " & RetournHeder("Activer", Disabled) & " type='checkbox' value='Annuler' checked/> "
                If Disabled = "" Then
                    Response.Write "         <input Name='Sav_Annuler_" & Lname & "' type='hidden' value='Annuler' />"
                End If
                Response.Write " </td>" & vbCrLf
           
        Else
             Response.Write "         <td align='center'><input  Name='Annuler_" & Lname & "' " & RetournHeder("Activer", Disabled) & " type='checkbox' value='Annuler'  />"
             If Disabled = "" Then
                Response.Write "         <input Name='Sav_Annuler_" & Lname & "' type='hidden' value='' />"
               
             
              End If
               Response.Write "</td>" & vbCrLf
              
        End If
        Response.Write "         <td align='center'>" & RsNomenclature("Projet") & "</td>" & vbCrLf
        Response.Write "<td>"
        If Disabled = "" Then
         Response.Write "<input name='Sav_ID_SousEnsembles_Sous-Ensseble_" & Lname & "' type='hidden' value='" & RsNomenclature("ID_SousEnsembles_Sous-Ensseble") & "' />"
       End If

       
        Response.Write Replace(Replace(T_RsT_SousEnsembles, "name='T_RsT_SousEnsembles'", "name='ID_SousEnsembles_Sous-Ensseble_" & Lname & "'" & RetournHeder("Sous_Ensseble", Disabled)), "value='" & Trim("" & RsNomenclature("ID_SousEnsembles_Sous-Ensseble")) & "'", "value='" & Trim("" & RsNomenclature("ID_SousEnsembles_Sous-Ensseble")) & "' Selected ") & "</td>" & vbCrLf
        
        Response.Write "<td>"

        If Disabled = "" Then
         Response.Write "<input name='Sav_Designation_Piece_" & Lname & "' type='hidden' value='" & RsNomenclature("Designation_Piece") & "' />"
         End If
'        Response.Write "<textarea class='Champ' name='Designation_Piece_" & Lname & "'" & RetournHeder("Designation_Piece", Disabled) & " rows='2' cols='20' >" & RsNomenclature("Designation_Piece") & "</textarea></td>" & vbCrLf
        Response.Write "<input name='Designation_Piece_" & Lname & "'" & RetournHeder("Designation_Piece", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("Designation_Piece") & "'/></td>" & vbCrLf
        
         Response.Write "             <td>"
       
        If Disabled = "" Then
         Response.Write "<input name='Sav_Pièce_" & Lname & "' type='hidden' value='" & RsNomenclature("Pièce") & "' />"
        End If
         
        Response.Write "<input Name='Pièce_" & Lname & "'" & RetournHeder("Pièce", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("Pièce") & "'/> </td>" & vbCrLf
        Response.Write "             <td>"
        
        If Disabled = "" Then
         Response.Write "<input name='Sav_Indice_" & Lname & "' type='hidden' value='" & RsNomenclature("Indice") & "' />"
         End If
         
        Response.Write "<input Name='Indice_" & Lname & "'" & RetournHeder("Indice", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("Indice") & "'/></td>" & vbCrLf
        Response.Write "         " & vbCrLf
        Response.Write "<td>"
        If Disabled = "" Then
         Response.Write "<input name='Sav_Quantite_" & Lname & "' type='hidden' value='" & RsNomenclature("Quantite") & "' />"
         End If
         
        Response.Write "<input Name='Quantite_" & Lname & "'" & RetournHeder("Quantité", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("Quantite") & "' onchange='RetourneSoldSurQts(this.document." & FormName & ".Quantite_" & L & "," & RsNomenclature("Quantite") & ",this.document." & FormName & ".Reste_a_recevoir_" & Lname & "," & RsNomenclature("Quantite") - RsNomenclature("Reste_a_recevoir") & ");'/></td>" & vbCrLf
         Response.Write "<td>"
        If Disabled = "" Then
         Response.Write "<input name='Sav_Reste_a_recevoir_" & Lname & "' type='hidden' value='" & RsNomenclature("Reste_a_recevoir") & "' />"
         End If
         
        Response.Write "<input Name='Reste_a_recevoir_" & Lname & "'" & RetournHeder("Pièce", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("Reste_a_recevoir") & "'/></td>" & vbCrLf
         Response.Write "<td>" & Replace(Replace(T_RsSqlAction, "name='T_RsSqlAction'", "name='Action_" & Lname & "'" & RetournHeder("Action", Disabled)), "value='" & Trim("" & RsNomenclature("Id_T_Actions_Libellé")) & "'", "value='" & Trim("" & RsNomenclature("Id_T_Actions_Libellé")) & "' Selected ") 'T_RsSqlAction
'         Response.Write "<input Name='Action_" & Lname & "' type='hidden' value='" & RsNomenclature("Id_T_Actions_Libellé") & "'/> "
         Response.Write "<input Name='Sav_Action_" & Lname & "' type='hidden' value='" & RsNomenclature("Id_T_Actions_Libellé") & "'/> "
         Response.Write "</td>" & vbCrLf
         
          Response.Write "<td>" & Replace(Replace(T_RsSqlDelais, "name='T_RsSqlDelais'", "name='Delais_" & Lname & "'" & RetournHeder("Action", Disabled)), "value='" & Trim("" & RsNomenclature("Id_T_Delais_Libellé")) & "'", "value='" & Trim("" & RsNomenclature("Id_T_Delais_Libellé")) & "' Selected ")  'T_RsSqlAction
        
         Response.Write "<input Name='Sav_Delais_" & Lname & "' type='hidden' value='" & RsNomenclature("Id_T_Delais_Libellé") & "'/> "
         Response.Write "</td>" & vbCrLf
         Response.Write "<td>"
        If Disabled = "" Then
         Response.Write "<input name='Sav_Qt_reçue_" & Lname & "' type='hidden' value='" & RsNomenclature("Qt_reçue") & "' />"
         End If
         
        Response.Write "<input Name='Qt_reçue_" & Lname & "'" & RetournHeder("Qt_reçue", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("Qt_reçue") & "' onchange='RetourneSold(this.document." & FormName & ".Quantite_" & Lname & ",this.document." & FormName & ".Qt_reçue_" & Lname & ",this.document." & FormName & ".QT_Retournee_" & Lname & ",this.document." & FormName & ".Reste_a_recevoir_" & Lname & "," & RsNomenclature("Reste_a_recevoir") & ");'/></td>" & vbCrLf
        Response.Write "<td>"
        If Disabled = "" Then
         Response.Write "<input name='Sav_Reçue_le_" & Lname & "' type='hidden' value='" & RsNomenclature("Reçue_le") & "' />"
         End If
         
        Response.Write "<input Name='Reçue_le_" & Lname & "'" & RetournHeder("Reçue_le", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("Reçue_le") & "'/></td>" & vbCrLf
        Response.Write "<td>"
        If Disabled = "" Then
         Response.Write "<input name='Sav_En_provenance_de_" & Lname & "' type='hidden' value='" & RsNomenclature("En_provenance_de") & "' />"
         End If
         
       Response.Write CreatListe(Rsvehicule, RsNomenclature("En_provenance_de"), "Id", "1;2", "En_provenance_de_" & Lname & "' Disabled ", "") & "</td>" & vbCrLf
       Response.Write "<td>"
        If Disabled = "" Then
         Response.Write "<input name='Sav_N_de_BL_" & Lname & "' type='hidden' value='" & RsNomenclature("N_de_BL") & "' />"
         End If
         
        Response.Write "<input Name='N_de_BL_" & Lname & "'" & RetournHeder("N°_ BL", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("N_de_BL") & "'/></td>" & vbCrLf
      Response.Write "<td>"
        If Disabled = "" Then
         Response.Write "<input name='Sav_QT_Retournee_" & Lname & "' type='hidden' value='" & RsNomenclature("QT_Retournee") & "' />"
         End If
         
        Response.Write "<input Name='QT_Retournee_" & Lname & "'" & RetournHeder("Retournée", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("QT_Retournee") & "' onchange='RetourneSold(this.document." & FormName & ".Quantite_" & Lname & ",this.document." & FormName & ".Qt_reçue_" & Lname & ",this.document." & FormName & ".QT_Retournee_" & Lname & ",this.document." & FormName & ".Reste_a_recevoir_" & Lname & "," & RsNomenclature("Reste_a_recevoir") & ");'/></td>" & vbCrLf
        
       
        Response.Write "<td>"
        If Disabled = "" Then
         Response.Write "<input name='Sav_Retournee_le_" & Lname & "' type='hidden' value='" & RsNomenclature("Retournee_le") & "' />"
         End If
         
        Response.Write "<input Name='Retournee_le_" & Lname & "'" & RetournHeder("Retournée", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("Retournee_le") & "'/></td>" & vbCrLf
        
Response.Write "<td>"
        If Disabled = "" Then
         Response.Write "<input name='Sav_Reaffectee_a_" & Lname & "' type='hidden' value='" & RsNomenclature("Reaffectee_a") & "' />"
         End If
         
        Response.Write CreatListe(Rsvehicule, RsNomenclature("Reaffectee_a"), "Id", "1;2", "Reaffectee_a_" & Lname & "'" & RetournHeder("Réaffectée", Disabled), "") & "</td>" & vbCrLf
         Response.Write "<td>"
        If Disabled = "" Then
         Response.Write "<input name='Sav_Num_de_retour_" & Lname & "' type='hidden' value='" & RsNomenclature("Num_de_retour") & "' />"
         End If
         
        Response.Write "<input Name='Num_de_retour_" & Lname & "'" & RetournHeder("N°_retour", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("Num_de_retour") & "'/></td>" & vbCrLf
        
        
         Response.Write "<td> <textarea id='TextArea1' rows='2' cols='20' class='ChampAlert'>" & RsNomenclature("Modification") & "</textarea></td>"
'
'
'       If Disabled = "" Then
'         Response.Write "<input name='Sav_Type_" & Lname & "' type='hidden' value='" & RsNomenclature("Type") & "' />"
'        End If
'        Response.Write "<input Name='Type_" & Lname & "'" & RetournHeder("Type", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("Type") & "' /></td>"
'
'
'        Response.Write " <td>"
'
'
'       If Disabled = "" Then
'         Response.Write "<input name='Sav_Désination FE_" & Lname & "' type='hidden' value='" & RsNomenclature("Désination FE") & "' />"
'        End If
'        Response.Write "<input Name='Id_" & Lname & "'" & Disabled & " type='hidden' value='" & RsNomenclature("id") & "'/><input Name='Désination FE_" & Lname & "'" & RetournHeder("Désination FE", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("Désination FE") & "'/></td>" & vbCrLf
'
'
'        Response.Write "<td>"
'        If Disabled = "" Then
'         Response.Write "<input name='Sav_Famille/FE/V_" & Lname & "'type='hidden' value='" & RsNomenclature("Famille/FE/V") & "' />"
'         End If
'
'        Response.Write "<input Name='Famille/FE/V_" & Lname & "'" & RetournHeder("Famille/FE/V", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("Famille/FE/V") & "'/></td>" & vbCrLf
'        Response.Write "<td>"
'
'        If Disabled = "" Then
'         Response.Write "<input name='Sav_Num_appro_ S_" & Lname & "'type='hidden' value='" & RsNomenclature("Num_appro_ S") & "' />"
'         End If
'
'        Response.Write "<input Name='Num_appro_ S_" & Lname & "'" & RetournHeder("Appro_S", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("Num_appro_ S") & "'/></td>" & vbCrLf
'
'        Response.Write "<td>"
'        If Disabled = "" Then
'         Response.Write "<input name='Sav_Bon_sortie_" & Lname & "'type='hidden' value='" & RsNomenclature("Bon_sortie") & "' />"
'         End If
'         'Response.Write Session("Bon_sortie") & "<br>"
'        Response.Write "<input Name='Bon_sortie_" & Lname & "'" & RetournHeder("Bon_sortie", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("Bon_sortie") & "'/></td>" & vbCrLf
'        Response.Write "<td>"
'
'        If Disabled = "" Then
'         Response.Write "<input name='Sav_Observation_ou_magasin_" & Lname & "' type='hidden' value='" & RsNomenclature("Observation_ou_magasin") & "' />"
'         End If
'
'        Response.Write "<input Name='Observation_ou_magasin_" & Lname & "'" & RetournHeder("Pièce", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("Observation_ou_magasin") & "'/></td>" & vbCrLf
'        Response.Write "<td>"
'
'        If Disabled = "" Then
'         Response.Write "<input name='Sav_Num_appro_" & Lname & "' type='hidden' value='" & RsNomenclature("Num_appro") & "' />"
'         End If
'
'        Response.Write "<input Name='Num_appro_" & Lname & "'" & RetournHeder("Masse", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("Num_appro") & "'/></td>" & vbCrLf
'        Response.Write "<td>"
'
'        If Disabled = "" Then
'         Response.Write "<input name='Sav_Masse_" & Lname & "' type='hidden' value='" & RsNomenclature("Masse") & "' />"
'         End If
'
'        Response.Write "<input Name='Masse_" & Lname & "'" & RetournHeder("Pièce", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("Masse") & "'/></td>" & vbCrLf
'
'
'
        Response.Write "<td>"
         Response.Write "<input name='Sav_Modifiee_le_" & Lname & "' type='hidden' value='" & RsNomenclature("Modifiee_le") & "' />"
         
        Response.Write "<input Name='Modifiee_le_" & Lname & "'" & RetournHeder("Modifiée", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("Modifiee_le") & "'/></td>" & vbCrLf
         Response.Write "<td><input Name='Modifier_Par" & Lname & "' type='text' class='Champ'  value='" & RsNomenclature("Modifier_Par") & "' disabled /></td>" & vbCrLf
        Response.Write "        <td>"
        
        If Disabled = "" Then
         Response.Write "<input name='Sav_Commantaire_" & Lname & "' type='hidden' value='" & RsNomenclature("Commantaire") & "' />"
         End If
         
        Response.Write "<input name='Commantaire_" & Lname & "'" & RetournHeder("Commantaire", Disabled) & " type='text' class='Champ'  value='" & RsNomenclature("Commantaire") & "'/></td>" & vbCrLf
        Response.Write "         "
        Response.Write "</tr>"
       
       End If
       If IndexRecord >= Stoped Then Exit Do
       RsNomenclature.MoveNext
        Loop
        
         Set RsT_SousEnsembles = CloseRecordSet(RsT_SousEnsembles)
            Response.Write RetourneSuitePage(currPageNum, countRecurd2)
     If Session("Updatable") = "True" Then
       
        Response.Write "<tr>" & vbCrLf
        Response.Write "<TD width='2'>" & AfficheBounton("Valider", 1, "'javascript:valider2();'") & "</td>" & vbCrLf
        Response.Write "</tr></table>" & vbCrLf
    End If
          
          Response.Write "<input name='NbLigne' type='hidden' value='" & L & "'/>"
          
          Response.Write "<input name='NbS_Enssemble__' type='hidden' value='" & L2 & "'/>"
          Response.Write "</form>"
      'RsSqlAction.close
      ' set RsSqlAction=nothing
End Function
Function RetourneSuitePage(CurrenPage, NbPage)
Dim SelectCase As Integer
If Val("" & CurrenPage) = 0 Then CurrenPage = 1
If CurrenPage = 1 Then SelectCase = 1
If CurrenPage > 1 And CurrenPage < NbPage Then SelectCase = 2
If CurrenPage >= NbPage Then SelectCase = 3
RetourneSuitePage = RetourneSuitePage & "<table border=""0"" cellspacing=""0""  align='left'>"
RetourneSuitePage = RetourneSuitePage & "<tr valign=""middle""><td>&nbsp;</td><td valign=""middle"" nowrap class=""smallerheader"">Page:&nbsp;</td>"
Select Case SelectCase
        Case 1
            RetourneSuitePage = RetourneSuitePage & "<td><img src=""leftEnd_.gif"" border=""0""><img src=""leftOne_.gif"" border=""0""></td>"
            RetourneSuitePage = RetourneSuitePage & "<td><input align=""right"" class=""cmdflat"" type=""text""  disabled name=""PREVPAGE"" value='" & CurrenPage & "' size=""4""></td>"
            RetourneSuitePage = RetourneSuitePage & "<td><a href='javascript:NaviguePage(" & CurrenPage + 1 & ")'>"
            RetourneSuitePage = RetourneSuitePage & "<img src=""rightOne.gif"" border=""0""></a>"
            RetourneSuitePage = RetourneSuitePage & "<a href='javascript:NaviguePage(" & NbPage & ");'><img src=""rightEnd.gif"" border=""0""></a></td>"
        Case 2
             RetourneSuitePage = RetourneSuitePage & "<td><a href='javascript:NaviguePage(1);'><img src=""leftEnd.gif"" border=""0""></a>"
            RetourneSuitePage = RetourneSuitePage & "<a href='javascript:NaviguePage(" & CurrenPage - 1 & ");'><img src=""leftOne.gif"" border=""0""></a></td>"
            
           RetourneSuitePage = RetourneSuitePage & "<td><input align=""right"" class=""cmdflat"" type=""text"" disabled name=""PREVPAGE"" value='" & CurrenPage & "' size=""4""></td>"
           
'            RetourneSuitePage = RetourneSuitePage & "<a href='javascript:NaviguePage(" & NbPage & ");'><img src=""rightEnd.gif"" border=""0""></a></td>"
            RetourneSuitePage = RetourneSuitePage & "<td><a href='javascript:NaviguePage(" & CurrenPage + 1 & ")'>"
            RetourneSuitePage = RetourneSuitePage & "<img src=""rightOne.gif"" border=""0""></a>"
            RetourneSuitePage = RetourneSuitePage & "<a href='javascript:NaviguePage(" & NbPage & ");'><img src=""rightEnd.gif"" border=""0""></a></td>"
        Case 3
            RetourneSuitePage = RetourneSuitePage & "<td><a href='javascript:NaviguePage(1);'><img src=""leftEnd.gif"" border=""0""></a>"
            RetourneSuitePage = RetourneSuitePage & "<a href='javascript:NaviguePage(" & CurrenPage - 1 & ");'><img src=""leftOne.gif"" border=""0""></a></td>"
            
           RetourneSuitePage = RetourneSuitePage & "<td><input align=""right"" class=""cmdflat"" type=""text"" disabled name=""PREVPAGE"" value='" & CurrenPage & "' size=""4""></td>"
           
            RetourneSuitePage = RetourneSuitePage & "<a href='javascript:NaviguePage(" & NbPage & ");'><img src=""rightEnd.gif"" border=""0""></a></td>"
            RetourneSuitePage = RetourneSuitePage & "<td>"
            RetourneSuitePage = RetourneSuitePage & "<img src=""rightOne_.gif"" border=""0"">"
            RetourneSuitePage = RetourneSuitePage & "<img src=""rightEnd_.gif"" border=""0""></td>"
End Select

          
         
          
         
        '******************* Previous Page *************** --> countRecurd2
        
'        RetourneSuitePage = RetourneSuitePage & "<td><a href='javascript:NaviguePage(" & PREVPAGE + 1 & ")'></a><a href='javascript:NaviguePage(" & countRecurd2 & ");'><img src=""rightEnd.gif"" border=""0""></a></td>"
              
       
        
              
       
        
          
        
        
              
        
              
        
         RetourneSuitePage = RetourneSuitePage & "<td class='smallerheader'> Sur : " & NbPage & "</td>"
       RetourneSuitePage = RetourneSuitePage & "</tr></table></td></tr>" & vbCrLf
End Function
Function ReplaceChamp(ChampsName)
    ReplaceChamp = ChampsName
    Select Case ChampsName

            Case "Num_de_retour"
                ReplaceChamp = "BL"
            Case "N_de_BL"
                ReplaceChamp = "BR"
            Case "Bon_sortie"
                ReplaceChamp = "Bon sortie"
            Case "Sous-Ensseble"
                    ReplaceChamp = "Sous Ensseble"
            Case "Désination FE"
            Case "Type"
            Case "Pièce"
            Case "Indice"
            Case "Commantaire"
            Case "Designation_Piece"
               ReplaceChamp = "Désignation Pièce"
            Case "Famille/FE/V"
            Case "Num_appro_ S"
                ReplaceChamp = "N° d'appro  S"
            Case "Quantite"
                 ReplaceChamp = "Quantitée"
            Case "Bon_sortie"
            Case "Observation_ou_magasin"
                    ReplaceChamp = "Observation ou magasin"
            Case "Num_appro"
                ReplaceChamp = "N° d'appro"
            Case "Masse"
            Case "Commandee_le"
                ReplaceChamp = "Commandée le"
            Case "Collectee_le"
                ReplaceChamp = "Collectée le"
            Case "Expedee_le"
                  ReplaceChamp = "Expedée le"
            Case "Reçue_le"
                 ReplaceChamp = "Reçue le"
            Case "Qt_reçue"
                ReplaceChamp = "Qt reçue"
            Case "Reste_a_recevoir"
                ReplaceChamp = "Reste à recevoir"
            Case "N_de_BL"
                ReplaceChamp = "N° de BL"
            Case "Reaffectee_a"
                ReplaceChamp = "Reaffectée a"
            Case "Qt_Reaffectee"
                 ReplaceChamp = "Qt Réaffectée"
            Case "En_provenance_de"
                ReplaceChamp = "En provenance de"
            Case "Retournee_le"
                ReplaceChamp = "Retournée le"
            Case "QT_Retournee"
                 ReplaceChamp = "QT Retournée"
            Case "Num_de_retour"
                ReplaceChamp = "N° de retour"
            Case "Modifiee_le"
                ReplaceChamp = "Modifiée le"

    End Select
End Function
Function NewReafectation(I, Br)
 
 Sql = "UPDATE T_Nomenclature SET T_Nomenclature.NumJob = Null "
                Sql = Sql & "WHERE T_Nomenclature.NumJob=" & noquote(Session("NumJob"), "txt") & ";"
                con.Execute Sql
                
    Sql = "SELECT T_SousEnsembles.ID_Vehicule_Véhicule "
                    Sql = Sql & "FROM T_SousEnsembles INNER JOIN T_Nomenclature ON T_SousEnsembles.Id =  "
                    Sql = Sql & "T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] "
                    Sql = Sql & "WHERE T_Nomenclature.Id=" & Request("id_" & I) & ";"
  'Response.Write sql
  'Response.End
    
   Set RsIdvehicule = OpenRecordSet(Sql, con)
   If RsIdvehicule.EOF = False Then
  id_Vehicule = RsIdvehicule("ID_Vehicule_Véhicule")
  Set RsIdvehicule = CloseRecordSet(RsIdvehicule)

'                Sql = "UPDATE T_Nomenclature INNER JOIN (T_Lien_SousEnssemble_Nomenclature INNER JOIN T_lien_Veicule_Souensembles  "
'                    Sql = Sql & "ON T_Lien_SousEnssemble_Nomenclature.ID_SousEnsembles =  "
'                    Sql = Sql & "T_lien_Veicule_Souensembles.ID__SousEnsembles) ON T_Nomenclature.Id =  "
'                    Sql = Sql & "T_Lien_SousEnssemble_Nomenclature.Id_Nomenclature  "
'                    Sql = Sql & "SET T_Nomenclature.NumJob = " & noquote(Session("NumJob"), "txt") & ", "
'                     Sql = Sql & "Flag=" & noquote(Session("NumJob"), "txt") & " "
'                    Sql = Sql & "WHERE T_Nomenclature.Modifiee_le Is Null  "
'                    Sql = Sql & "AND T_Nomenclature.Pièce=" & noquote(Request("Sav_Pièce_" & I), "txt") & "  "
'                    Sql = Sql & "AND T_Nomenclature.Indice=" & noquote(Request("Sav_Indice_" & I), "txt") & "  "
'                    Sql = Sql & "AND T_SousEnsembles.ID_Vehicule_Véhicule=" & noquote(Request("Sav_Reaffectee_a_" & I), "num") & ";"
'
'                    Sql = "UPDATE T_SousEnsembles INNER JOIN T_Nomenclature ON T_SousEnsembles.Id =   "
'                    Sql = Sql & "T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble]   "
'                   Sql = Sql & "SET T_Nomenclature.NumJob = " & noquote(Session("NumJob"), "txt") & ", "
'                     Sql = Sql & "T_Nomenclature.Flag=" & noquote(Session("NumJob"), "txt") & " "
'                    Sql = Sql & "WHERE T_Nomenclature.Modifiee_le Is Null  "
'                    Sql = Sql & "AND T_Nomenclature.Pièce=" & noquote(Request("Sav_Pièce_" & I), "txt") & "  "
'                    Sql = Sql & "AND T_Nomenclature.Indice=" & noquote(Request("Indice_" & I), "txt") & "  "
'                    Sql = Sql & "AND T_SousEnsembles.ID_Vehicule_Véhicule=" & noquote(Request("Sav_Reaffectee_a_" & I), "num") & ";"

                    Sql = "Update T_Nomenclature "
                    Sql = Sql & "SET T_Nomenclature.NumJob = " & noquote(Session("NumJob"), "txt") & ",  "
                    Sql = Sql & "T_Nomenclature.Flag = " & noquote(Session("NumJob"), "txt") & " "
                    Sql = Sql & "WHERE     T_Nomenclature.Modifiee_le IS NULL AND T_Nomenclature.Pièce = " & noquote(Request("Sav_Pièce_" & I), "txt") & "  "
                    Sql = Sql & "AND T_Nomenclature.Indice =" & noquote(Request("Indice_" & I), "txt") & "   "
                    Sql = Sql & "AND T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] = "
                    Sql = Sql & "(SELECT     T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] "
                    Sql = Sql & "FROM T_SousEnsembles INNER JOIN "
                    Sql = Sql & "T_Nomenclature ON T_SousEnsembles.Id = T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] "
                    Sql = Sql & "WHERE T_Nomenclature.Modifiee_le IS NULL  "
                    Sql = Sql & "AND T_Nomenclature.Pièce = " & noquote(Request("Sav_Pièce_" & I), "txt") & "  "
                    Sql = Sql & "AND T_Nomenclature.Indice = " & noquote(Request("Sav_Indice_" & I), "txt") & "  "
                    Sql = Sql & "AND T_SousEnsembles.ID_Vehicule_Véhicule = " & noquote(Request("Reaffectee_a_" & I), "num") & ");"
                    
                    
                    
                    
                    
                     con.Execute Sql
    
                    Sql = "UPDATE T_Nomenclature "
                    Sql = Sql & "SET T_Nomenclature.Reçue_le = " & noquote(Now(), "date") & ",  "
                    Sql = Sql & "T_Nomenclature.Qt_reçue =" & noquote(Request("QT_Retournee_" & I), "num") & ",  "
                    Sql = Sql & "T_Nomenclature.Reste_a_recevoir = [Reste_a_recevoir]- " & noquote(Request("QT_Retournee_" & I), "num") & ",  "
                    Sql = Sql & "T_Nomenclature.Modifiee_le = " & noquote(Now(), "date") & ",  "
                    Sql = Sql & "T_Nomenclature.En_provenance_de = " & id_Vehicule & "  "
                    Sql = Sql & ",T_Nomenclature.N_de_BL = '" & Br & " ' "
                    Sql = Sql & "WHERE T_Nomenclature.NumJob= " & noquote(Session("NumJob"), "txt") & ";"
                 
                 
               
                  
                 
                     con.Execute Sql
                     
                     Sql = "SELECT T_Nomenclature.Id, T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] "
                    Sql = Sql & "FROM T_Nomenclature "
                    Sql = Sql & "WHERE T_Nomenclature.NumJob= " & noquote(Session("NumJob"), "txt") & ";"
                   
                    
                Set RsIdNomEnssemble = con.Execute(Sql)
                     Id_Nom = RsIdNomEnssemble("Id")
                     Id_Sens = RsIdNomEnssemble("ID_SousEnsembles_Sous-Ensseble")
                Set RsIdNomEnssemble = CloseRecordSet(RsIdNomEnssemble)
                
                Sql = "UPDATE T_Nomenclature SET T_Nomenclature.NumJob = Null "
                Sql = Sql & "WHERE T_Nomenclature.NumJob=" & noquote(Session("NumJob"), "txt") & ";"
                con.Execute Sql
                
                     Sql = "INSERT INTO T_Nomenclature ( Id_Stok,[ID_SousEnsembles_Sous-Ensseble], [Désination FE], Type, Pièce, Indice, Commantaire,  "
                    Sql = Sql & "Designation_Piece, [Famille/FE/V], [Num_appro_ S], Quantite, Bon_sortie,  "
                    Sql = Sql & "Observation_ou_magasin, Num_appro, Masse, Commandee_le, Collectee_le,  "
                    Sql = Sql & "Reste_a_recevoir, NumJob ) "
                    Sql = Sql & "SELECT T_Nomenclature.Id_Stok, T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble],T_Nomenclature.[Désination FE],  "
                    Sql = Sql & "T_Nomenclature.Type, T_Nomenclature.Pièce, T_Nomenclature.Indice,  "
                    Sql = Sql & "T_Nomenclature.Commantaire, T_Nomenclature.Designation_Piece, T_Nomenclature.[Famille/FE/V],  "
                    Sql = Sql & "T_Nomenclature.[Num_appro_ S], T_Nomenclature.Quantite, T_Nomenclature.Bon_sortie,  "
                    Sql = Sql & "T_Nomenclature.Observation_ou_magasin, T_Nomenclature.Num_appro, T_Nomenclature.Masse,  "
                    Sql = Sql & "T_Nomenclature.Commandee_le, T_Nomenclature.Collectee_le,  "
                    Sql = Sql & "T_Nomenclature.Reste_a_recevoir, " & noquote(Session("NumJob"), "txt") & " AS NumJob "
                    Sql = Sql & "FROM T_Nomenclature "
                    Sql = Sql & "WHERE T_Nomenclature.Id=" & Id_Nom & ";"
                   'Response.Write sql
                    'Response.End
                     con.Execute Sql
                     
                   ' sql="SELECT T_Nomenclature.Id "
                    'sql=sql & "FROM T_Nomenclature "
                    'sql=sql & "WHERE T_Nomenclature.NumJob=" & noquote( session("NumJob"),"txt") & ";"
                    ' set RsIdNomEnssemble= OpenRecordSet(sql,con)
                    ' Id_Nom=RsIdNomEnssemble("Id")
                    'set RsIdNomEnssemble =CloseRecordSet(RsIdNomEnssemble)
                    
                    'sql="INSERT INTO T_Lien_SousEnssemble_Nomenclature ( ID_SousEnsembles, Id_Nomenclature ) "
                    'sql=sql & "VALUES ( " & Id_Sens & "," & Id_Nom & ");"
                     'Con.execute sql


    End If
     
End Function
Function MajNewNomenclateur(I)
Sql = "insert into T_Nomenclature ("
                    Sql = Sql & "NumJob,"
                   Sql = Sql & "[ID_SousEnsembles_Sous-Ensseble],"
                    Sql = Sql & "[Désination FE],  "
                    Sql = Sql & "Type,  "
                    Sql = Sql & "Pièce,  "
                    Sql = Sql & "Indice,  "
                    Sql = Sql & "Commantaire,  "
                    Sql = Sql & "Designation_Piece,  "
                    Sql = Sql & "[Famille/FE/V],  "
                    Sql = Sql & "[Num_appro_ S],  "
                    Sql = Sql & "Quantite,  "
                    Sql = Sql & "Bon_sortie,  "
                    Sql = Sql & "Observation_ou_magasin,  "
                    Sql = Sql & "Num_appro,  "
                    Sql = Sql & "Masse,  "
                    Sql = Sql & "Commandee_le,"
                    Sql = Sql & "Reste_a_recevoir,"
                    Sql = Sql & "Activer,Id_T_Actions_Libellé,Id_T_Delais_Libellé) "
                    
                   
                    
                    
                  
                    
                    Sql = Sql & "values ("
                     Sql = Sql & noquote(Session("NumJob"), "txt") & ",  "
                     Sql = Sql & noquote(Request("ID_SousEnsembles_Sous-Ensseble_" & I), "num") & ",  "
                    Sql = Sql & noquote(Request("Désination FE_" & I), "txt") & ",  "
                    Sql = Sql & noquote(Request("Type_" & I), "txt") & ",  "
                    Sql = Sql & noquote(Request("Pièce_" & I), "txt") & ",  "
                    Sql = Sql & noquote(Request("Indice_" & I), "txt") & ",  "
                    Sql = Sql & noquote(Request("Commantaire_" & I), "txt") & ",  "
                    Sql = Sql & noquote(Request("Designation_Piece_" & I), "txt") & ",  "
                    Sql = Sql & noquote(Request("Famille/FE/V_" & I), "txt") & ",  "
                    Sql = Sql & noquote(Request("Num_appro_ S_" & I), "txt") & ",  "
                    Sql = Sql & noquote(Request("Quantite_" & I), "num") & ",  "
                    Sql = Sql & noquote(Request("Bon_sortie_" & I), "txt") & ",  "
                    Sql = Sql & noquote(Request("Observation_ou_magasin_" & I), "txt") & ",  "
                    Sql = Sql & noquote(Request("Num_appro_" & I), "txt") & ",  "
                    Sql = Sql & noquote(Request("Masse_" & I), "num") & ",  "
                    Sql = Sql & noquote(Request("Commandee_le_" & I), "date") & ",  "
                    Sql = Sql & noquote(Request("Reste_a_recevoir_" & I), "num") & ",  "
                    If Trim("" & Request("Annuler_" & I)) <> "" Then
                         Sql = Sql & -1 & " "
                         
                    Else
                        Sql = Sql & 0 & " "
                    End If
                    
                    
                    
                    Sql = Sql & "," & noquote(Request("Action_" & I), "num") & "," & noquote(Request("Delais_" & I), "num") & ");"
                    'Response.Write sql
                    'Response.End
                      con.Execute Sql
                      Sql = "UPDATE T_Nomenclature SET T_Nomenclature.Modifiee_le = " & noquote(Now(), "date") & ",T_Nomenclature.Modifier_Par ='" & Session("User") & "' "
                        Sql = Sql & "WHERE T_Nomenclature.Id=" & Request("id_" & I) & ";"
                        'Response.Write sql
                       ' Response.End
                        con.Execute Sql
                   
End Function
Function HISTORIQUE_Nommenclature()
Dim MyAtribut
Dim Sql
Dim MyL
On Error GoTo 0
 BlExiste = False
      
           
   

  currPageNum = txtGetDefault("HISTORIQUE_NB_BL_Record", "5")

   
 If Trim("" & Request("HISTORIQUE_Nommenclature_SortBy")) <> "" Then
        Session("HISTORIQUEsBlSortBy") = Request("HISTORIQUE_Nommenclature_SortBy")
 End If
  If Trim("" & Session("HISTORIQUEsBlSortBy")) = "" Then Session("HISTORIQUEsBlSortBy") = txtGetDefault("SetDefauleBLSortBy", "N_de_BL")
 
STRSQL = "SELECT DISTINCT  T_Nomenclature.Id, t_Vehicule.Projet, t_Vehicule.Véhicule, T_SousEnsembles.[Sous-Ensseble],  "
    STRSQL = STRSQL & "T_Nomenclature.Pièce, T_Nomenclature.Indice, T_Nomenclature.Bon_sortie, T_Nomenclature.Commandee_le,  "
    STRSQL = STRSQL & "T_Nomenclature.Quantite, T_Nomenclature.Reçue_le, T_Nomenclature.Qt_reçue, T_Nomenclature.N_de_BL ,  "
    STRSQL = STRSQL & "T_Nomenclature.Reste_a_recevoir, T_Nomenclature.Reaffectee_Le, T_Nomenclature.Qt_Reaffectee,  "
    STRSQL = STRSQL & "T_Nomenclature.Retournee_le, T_Nomenclature.QT_Retournee, T_Nomenclature.Num_de_retour  "
    STRSQL = STRSQL & "FROM t_Vehicule INNER JOIN ((T_SousEnsembles INNER JOIN (T_Nomenclature INNER JOIN "
    STRSQL = STRSQL & "T_Lien_SousEnssemble_Nomenclature ON T_Nomenclature.Id =  "
    STRSQL = STRSQL & "T_Lien_SousEnssemble_Nomenclature.Id_Nomenclature) ON T_SousEnsembles.Id =  "
    STRSQL = STRSQL & "T_Lien_SousEnssemble_Nomenclature.ID_SousEnsembles) INNER JOIN T_lien_Veicule_Souensembles  "
    STRSQL = STRSQL & "ON T_SousEnsembles.Id = T_lien_Veicule_Souensembles.ID__SousEnsembles) ON t_Vehicule.Id =  "
    STRSQL = STRSQL & "T_lien_Veicule_Souensembles.ID_Vehicule "
    STRSQL = STRSQL & "GROUP BY T_Nomenclature.Id, t_Vehicule.Projet, t_Vehicule.Véhicule, T_SousEnsembles.[Sous-Ensseble],  "
    STRSQL = STRSQL & "T_Nomenclature.Pièce, T_Nomenclature.Indice, T_Nomenclature.Bon_sortie, T_Nomenclature.Commandee_le,  "
    STRSQL = STRSQL & "T_Nomenclature.Quantite, T_Nomenclature.Reçue_le, T_Nomenclature.Qt_reçue, T_Nomenclature.N_de_BL,  "
    STRSQL = STRSQL & "T_Nomenclature.Reste_a_recevoir, T_Nomenclature.Reaffectee_Le, T_Nomenclature.Qt_Reaffectee,  "
    STRSQL = STRSQL & "T_Nomenclature.Retournee_le, T_Nomenclature.QT_Retournee, T_Nomenclature.Num_de_retour "
    If Session("HISTORIQUEsBlSortBy") <> "" Then
    '    STRSQL = STRSQL & "ORDER BY T_Commande_Liv.creation;"

    'Else
        
        STRSQL = STRSQL & "ORDER BY     " & Session("HISTORIQUEsBlSortBy") & ";"
    'End If

    
'Response.Write  STRSQL
'Response.End
 Set rs0 = OpenRecordSet(STRSQL, con)

    sqlcount = "SELECT Count([Id]) AS RecCount "
        sqlcount = sqlcount & "FROM T_Nomenclature ;"







    '*********** txtGet record count
   
    Set RS11 = OpenRecordSet(sqlcount, con)

    RecCount = RS11("RecCount")
    
    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<script language=""VBscript"">"
    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "function vbMouseOver(a)"
    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "   a.style.backgroundcolor = """ & txtGetDefault("mouseover", "#AFB4DC") & """"
    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "end function"
    
    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "function vbMouseOut(a)"
    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "   a.style.backgroundcolor = """ & txtGetDefault("mouseout", "#DBDEF1") & """"
    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "end function"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "</script>"
  
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<script language='javascript'>"
    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "function message(txt) {"
    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "   top.status=txt;"
    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "}"
    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & ""
   
    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "function javAnnuler() {"
    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "myForm = this.document.forms[0];"
    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "myForm.mode.value = 'CON_lst';"
    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "myForm.action='Contact.asp?mode=CON_lst';"
    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "myForm.submit();"
    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "}"
    
    
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "function javCadyRempli() {"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "myForm = this.document.forms[0];"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "myForm.mode.value = 'Candidat_caddy';"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "myForm.action='Contact.asp?mode=CONFIGCLI';"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "myForm.submit();"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "}"
        
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "function javSendEmail() {"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "   document.frm.submit()"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "}"
    
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "function javSelectAll(chk) {"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "   len = document.frm.elements.length;"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "   var i=0;"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "   for( i=0; i<len; i++) {"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "       if (document.frm.elements[i].name=='lstTO') {"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "           document.frm.elements[i].checked=chk;"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "       }"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "   }"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "}"
    
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "function javNewContact() {"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "   location.href = 'contact.asp?mode=CONFIGCLI_New&ID=0'"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "}"
      
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "function javResetSearch() {"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "   location.href = 'Contact.asp?mode=CONFIGCLI&reset=true'"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "}"
    
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "function javCat() {"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "   var ID = document.frmSearch.ID.options[document.frmSearch.ID.selectedIndex].value"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "   location.href = 'Contact.asp?mode=CONFIGCLI_New&ID=' + ID"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "}"
    
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "function javAddToCat() {"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "   window.open('con_dlgAddToCat.asp?ID=" & Session("candidat_SubCategory") & "&sid=" & Session("candidat_web_numDevis") & "&ListType=" & Session("candidat_Category") & "','dlgcat','resizable=yes,status=no,top=200,left=200,width=300,height=400')"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "}"
    
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "function javRemoveFromCat() {"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "   window.open('con_dlgRemoveFromCat.asp?ID=" & Session("candidat_SubCategory") & "&sid=" & Session("candidat_web_numDevis") & "&ListType=" & Session("candidat_Category") & "','dlgcat','resizable=yes,status=no,top=200,left=200,width=300,height=400')"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "}"
    
    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "function javDelete(ID) {"
       HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "    if (confirm('" & txtGetMessage("SuppFiche", "Supprimer cette fiche?") & "')) {"
       HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "        location.href = 'Default.asp?Logistique_Page=2ID='+ID"
       HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "    } else {"
       HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "        return"
       HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "    }"
       HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "}"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "</script>"
    
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<script language=""VBscript"">"
    
    
  
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "</script>"
    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<a name=""#Haut""></a>"

    '******** TitleBar
    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<table width=""100%"" border=""2"" bgcolor=""" & txtGetDefault("CONFIGCLI_bgcolor", "#DCC9B2") & """ cellpadding=""1"" cellspacing=""0"">"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<tr valign=""bottom"">"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<td>"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<table width=""100%"" border=""0"">"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<tr valign=""bottom"">"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<form name=""frmBL"" action='Default.asp?Logistique_Page=2'>"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<input type=""hidden"" name=""mode"" value=""HISTORIQUE_Eboutique_Affiche_Commande"">"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<td align=""left"" valign=""middle"">"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "Liste des BL"
    
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "</td>"

   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<td align=""right"" class=""smallerheader"" valign=""middle"">"
        If Session("candidat_UserType") = "Administrator" Then
    End If
     Set CaddyConn = Nothing
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "</td>"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "</form>"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "</tr>"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "</table>"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "</td>"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "</tr>"
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "</table>"
    
     '************** Headers ******
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<table border=""0"" width=""100%"" cellspacing=""1"" bgcolor=""" & txtGetDefault("bglstheadertr", "silver") & """>"
    
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<tr bgcolor=""" & txtGetDefault("bglsttr", "#E4D7C7") & """>"

    For I = 1 To rs0.Fields.Count - 1
        MyTxt = ReplaceChamp(rs0(I).Name)
        If Trim(MyTxt) = "" Then MyTxt = "&nbsp"
        
        If InStr(UCase(Session("HISTORIQUEsBlSortBy")), UCase(rs0(I).Name)) <> 0 Then
              
            If InStr(Session("HISTORIQUEsBlSortBy"), "Desc") Then
            
               HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<td nowrap><a href='Default.asp?Logistique_Page=2&RefCommande=" & numDevis & "&HISTORIQUE_Nommenclature_SortBy=[" & rs0(I).Name & "]'><img src='./Images/SortDesc.gif' border='0'><b>" & MyTxt & "</b></a></td>"
            Else
               HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<td nowrap><a href='Default.asp?Logistique_Page=2&RefCommande=" & numDevis & "&HISTORIQUE_Nommenclature_SortBy=[" & rs0(I).Name & "]+Desc'><img src='./Images/SortAss.gif' border='0'><b>" & MyTxt & "</b></a></td>"
            End If
        Else
            
           HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<td nowrap><a href='Default.asp?Logistique_Page=2&RefCommande=" & numDevis & "&HISTORIQUE_Nommenclature_SortBy=[" & rs0(I).Name & "]'><b>" & MyTxt & "</b></a></td>"
        End If
    Next
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "</tr>"
    
    
     
    '************  Initiate page counter
    
    lstPage = 1
    ENDITEM = currPageNum
    BEGINITEM = 1
    I = 0
    ICOUNTER = 0
    
    INTPAGE = Trim(Request("NEXTPAGE"))
    If Len(INTPAGE) > 0 And IsNumeric(INTPAGE) Then
        Session("HISTORIQUE_ClientPage") = INTPAGE
    End If
    
    If Len(Session("HISTORIQUE_ClientPage")) > 0 Then
    If CInt(Session("HISTORIQUE_ClientPage")) = 0 Then Session("HISTORIQUE_ClientPage") = 1
        lstPage = CInt(Session("HISTORIQUE_ClientPage"))
       
        ENDITEM = lstPage * currPageNum
        
        BEGINITEM = CLng(ENDITEM) - (currPageNum - 1)
    End If
     If RecCount < BEGINITEM Then
        lstPage = 1
        ENDITEM = currPageNum
        BEGINITEM = 1
        I = 0
        ICOUNTER = 0
    End If
    I = 0
    LastRegroupBlob = ""
    my_bglsttrok = txtGetDefault("bglsttrok", "#DCC9B2")
    my_bglsttrnook = txtGetDefault("bglsttrnook", "#DBDEF1")
    Do While Not rs0.EOF
       
        'if (LastRegroupBlob <> rs0("RegroupBlob")) then
        I = I + 1
       
        If (I >= BEGINITEM And I <= CLng(ENDITEM)) Then
             BlExiste = True
               HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<tr id=""CLI_FRM" & I & """ style=""background:" & my_bglsttrnook & ";"" valign=""top"" OnMouseOver=""vbMouseOver(CLI_FRM" & I & ")"" OnMouseOut=""vbMouseOut(CLI_FRM" & I & ")"" >" 'OnClick=""vbEditContact(" & rs0("N° Comm Encelade") & ")"">
            
            For j = 1 To (rs0.Fields.Count - 1)
                If Trim("" & rs0(j)) = "" Then MyTxt = "&nbsp" Else MyTxt = rs0(j)
                            'Debug.Print rs0(j).Type
                            If IsNumeric(MyTxt) Then
                                MyTxt = FomatNum("" & MyTxt, 3)
                            Else
                                If IsNumeric(Replace(MyTxt, ",", ".")) Then MyTxt = FomatNum("" & MyTxt, 3)
                            End If
                  
                    If Session("candidat_UserType") <> "Anonyme" Then
                   
                    
                        If UCase(rs0(j).Name) = UCase("Liste rouge") Then
                           
                            If rs0(j) = False Then
                                MyTxt = "NON"
                            Else
                                MyTxt = "OUI"
                            End If
                           HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<td nowrap><a  onMouseOver='message(""Ouvir BL N° "");return true;' href='Default.asp?Logistique_Page=4&ID=" & rs0("id") & "'>" & MyTxt & " </a></td>"
                        Else
                            HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<td nowrap><a  onMouseOver='message(""Ouvir BL N° "");return true;' href='Default.asp?Logistique_Page=4&ID=" & rs0("id") & "'>" & MyTxt & " </a></td>"
                        End If
                    Else
                       HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<td nowrap><a  onMouseOver='message(""Ouvir BL N° "");return true;' href='Default.asp?Logistique_Page=4&ID=" & rs0("id") & "'>" & MyTxt & " </a></td>"

                    End If

            
            Next

           HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "</tr>"
            ICOUNTER = ICOUNTER + 1
            
            '*** Build javascript for select all
        End If
        If I = CLng(ENDITEM) Then
            Exit Do
        End If
    
        rs0.MoveNext
    Loop
  
   HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "</table>"
    
    '***********************************  Page Count  *****************
    PageCount = CInt(RecCount / currPageNum)
     PageCount2 = RecCount / currPageNum
     If PageCount < PageCount2 Then PageCount = PageCount + 1
    If RecCount = 0 Then
        PageCount = 0
    ElseIf PageCount < 1 Then
        PageCount = 1
    Else
        INTLEFT = PageCount
        
        If InStr(PageCount, ".") Then INTLEFT = Left(PageCount, InStr(PageCount, ".") - 1)
        If InStr(PageCount, ",") Then INTLEFT = Left(PageCount, InStr(PageCount, ",") - 1)
        If CLng(Replace(INTLEFT, ",", ".")) > CLng(PageCount) Then PageCount = INTLEFT + 1
    End If
    
    If ((ICOUNTER > 0) And (Session("candidat_UserType") = "Administrator")) Or ((ICOUNTER > 0) And (Session("portal_candidat_Recherhe") = 0)) Then
        
        BRCOUNT = currPageNum - ICOUNTER
        PREVPAGE = lstPage - 1
        NEXTPAGE = lstPage + 1
       HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<table border=""0"" cellspacing=""0"">"
       HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<form name=""frmClientUser"" action=""Contact.asp"" method=""post"">"
       HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<input type=""hidden"" name=""mode"" value=""HISTORIQUE_Eboutique_Affiche_Commande"">"
       HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<input type=""hidden"" name=""RefCommande"" value=""" & numDevis & """>"
       HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<tr valign=""middle""><td>&nbsp;</td><td valign=""middle"" nowrap class=""smallerheader"">Page:&nbsp;</td>"
        
        '******************* Previous Page *************** -->
        If PREVPAGE < 1 Then
           HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<td><img src=""./Images/leftEnd_.gif"" border=""0""><img src=""./Images/leftOne_.gif"" border=""0""></td>"
        Else
           HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<td><a href=""Default.asp?Logistique_Page=2&RefCommande=" & numDevis & "&NEXTPAGE=1""><img src=""./Images/leftEnd.gif"" border=""0""></a><a href=""Default.asp?Logistique_Page=2&RefCommande=" & numDevis & "&NEXTPAGE=" & PREVPAGE & """><img src=""./Images/leftOne.gif"" border=""0""></a></td>"
        End If
        
       HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<td><input align=""right"" class=""cmdflat"" type=""text"" name=""lstPage"" value=""         " & lstPage & """ size=""4""></td>"
        
        If NEXTPAGE > PageCount Then
           HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<td><img src=""./Images/rightOne_.gif"" border=""0""><img src=""./Images/rightEnd_.gif"" border=""0""></td>"
        Else
           HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<td><a href=""Default.asp?Logistique_Page=2&RefCommande=" & numDevis & "&NEXTPAGE=" & NEXTPAGE & """><img src=""./Images/rightOne.gif"" border=""0""></a><a href=""Default.asp?Logistique_Page=2&RefCommande=" & numDevis & "&NEXTPAGE=" & PageCount & """><img src=""./Images/rightEnd.gif"" border=""0""></a></td>"
        End If
        
       HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<td valign=""middle""><font class=""smallerheader"">&nbsp;sur " & PageCount
       HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "&nbsp;&nbsp;&nbsp;&nbsp;Total :&nbsp;" & RecCount & "</font></td>"
       HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "</tr>"
       HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "</form>"
       HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "</table>"
        End If

    

    '***************** MESSAGE **********

    If Len(msg) > 0 Then
       HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<center><font color=""" & txtGetDefault("colormsg", "red") & "><b>" & msg & "</b></font></center>"
    End If

    HISTORIQUE_Nommenclature = HISTORIQUE_Nommenclature & vbCrLf & "<a  onMouseOver='message(""Haut de page"");return true;' href=""#Haut""><img src=""./Images/hautdepage.gif"" border=""0""></a>"
  End If
   
End Function
 Function Fun1()
   
  If Session("Updatable") = "True" Then
    majfunction
    Fun1 = JavaValider("T_User", 1) & vbCrLf
    End If
    Fun1 = Fun1 & RetourSqlTableau("select T_User.* from T_User ORDER BY T_User.id;", "T_User", "Id", "", "", "OK", "", "")
End Function
Function Fun18()
   
  If Session("Updatable") = "True" Then
    majfunction
    Fun18 = JavaValider("T_Societe", 1) & vbCrLf
    End If
    Fun18 = Fun18 & RetourSqlTableau("select T_Societe.* from T_Societe ORDER BY T_Societe.id;", "T_Societe", "Id", "", "", "OK", "", "")
End Function
Function Fun19()
   
  If Session("Updatable") = "True" Then
    majfunction
    Fun19 = JavaValider("T_Projects", 19) & vbCrLf
    End If
    Sql = "select DISTINCT T_Projects.* "
    Sql = Sql & "FROM t_Vehicule INNER JOIN "
        Sql = Sql & "T_Projects ON t_Vehicule.Id_Projects_NumAffaire = T_Projects.Id LEFT OUTER JOIN "
        Sql = Sql & "T_SousEnsembles ON t_Vehicule.Id = T_SousEnsembles.ID_Vehicule_Véhicule LEFT OUTER JOIN "
        Sql = Sql & "T_Nomenclature ON T_SousEnsembles.Id = T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] "
        Sql = Sql & ""
        
    Fun19 = Fun19 & RetourSqlTableau("select T_Projects.* from T_Projects ORDER BY T_Projects.id;", "T_Projects", "Id", "", "", "OK", "", "")
End Function

Function Fun2()
    'majfunction
    Fun2 = JavaValider("t_Vehicule", 2) & vbCrLf
    'Fun2=Fun2 & HISTORIQUE_Nommenclature()
   
    Fun2 = Fun2 & RetourVehicule("t_Vehicule", "OK", "Default.asp?Logistique_Page=2&ID=")
End Function

Function Fun3()
 If Session("Updatable") = "True" Then
    majfunction
    Fun3 = JavaValider("T_Stock", 3) & vbCrLf
    End If
    
    Fun3 = Fun3 & RetourSqlTableau("SELECT T_Stock.*  FROM T_Stock ORDER BY T_Stock.Ordre;", "T_Stock", "Id", "", "", "OK", "", "")
End Function

Function Fun14()
 If Session("Updatable") = "True" Then
    majfunction
    Fun14 = JavaValider("T_Actions", 14) & vbCrLf
    End If
    
    
    Fun14 = Fun14 & RetourSqlTableau("SELECT T_Actions.*  FROM T_Actions ORDER BY T_Actions.Id;", "T_Actions", "Id", "", "", "OK", "", "")
End Function
Function Fun15()
 If Session("Updatable") = "True" Then
    majfunction
    Fun15 = JavaValider("T_Delais", 15) & vbCrLf
    End If
    
    
    Fun15 = Fun15 & RetourSqlTableau("SELECT T_Delais.*  FROM T_Delais ORDER BY T_Delais.Id;", "T_Delais", "Id", "", "", "OK", "", "")
End Function

Function Fun4()
 If Session("Updatable") = "True" Then
    Fun4 = "" 'JavaValider("T_Nomenclature",4) &  vbcrlf
 End If
   Sql = Sql & "WHERE t_Vehicule.Id_Stock=1;"
    Fun4 = Fun4 & RetourSqlNomenclatuer("T_Nomenclature", "Id", "", "", "OK")
End Function

Function Fun5()
 If Session("Updatable") = "True" Then
    majfunction
    Fun5 = JavaValider("T_SousEnsembles", 5) & vbCrLf
      
 End If
        Sql = "SELECT T_SousEnsembles.Id, T_SousEnsembles.ID_Vehicule_Véhicule, T_SousEnsembles.Code, T_SousEnsembles.[Sous-Ensseble],  "
        Sql = Sql & "T_SousEnsembles.[MASSE total (KG)], T_SousEnsembles.[Nb lignes sans masse],  "
        Sql = Sql & "T_SousEnsembles.[Nb total de lignes]  "
        Sql = Sql & "FROM t_Vehicule INNER JOIN "
        Sql = Sql & "T_Projects ON t_Vehicule.Id_Projects_NumAffaire = T_Projects.Id LEFT OUTER JOIN "
        Sql = Sql & "T_SousEnsembles ON t_Vehicule.Id = T_SousEnsembles.ID_Vehicule_Véhicule LEFT OUTER JOIN "
        Sql = Sql & "T_Nomenclature ON T_SousEnsembles.Id = T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] "
           MyFilter = RetourneCloseWhere()
          If Request("nbSelect") > 0 Then
            Sql = Sql & MyCloseWhere()
             If Trim("" & MyFilter) <> "" Then
                Sql = Sql & " AND (" & MyFilter & ") "
             End If
          Else
            If Trim("" & MyFilter) <> "" Then
                Sql = Sql & " WHERE " & MyFilter
             End If
        End If
        
        If Trim("" & Request(Request("FormName") & "FiltreSouEnsemble")) <> "" Then
                    If InStr(1, UCase(Sql), "WHERE") <> 0 Then
                    
                Sql = Sql & " AND t_Vehicule.Id=" & Request(Request("FormName") & "FiltreSouEnsemble") & " "
                
            Else
                MyFilter = "?"
                 Sql = Sql & " WHERE t_Vehicule.Id=" & Request(Request("FormName") & "FiltreSouEnsemble") & " "
            End If
      End If
        
        Sql = Sql & "GROUP BY T_SousEnsembles.Id, T_SousEnsembles.ID_Vehicule_Véhicule, T_SousEnsembles.Code, T_SousEnsembles.[Sous-Ensseble],  "
        Sql = Sql & "T_SousEnsembles.[MASSE total (KG)], T_SousEnsembles.[Nb lignes sans masse],  "
        Sql = Sql & "T_SousEnsembles.[Nb total de lignes];"
        'Response.Write sql
        'Response.End
    Fun5 = Fun5 & RetourSqlTableau(Sql, "T_SousEnsembles", "Id", "", "1", "OK", "", MyFilter)
End Function
Function Fun6()
    If Session("Updatable") = "True" Then
    Fun6 = JavaValider("T_lien_Veicule_Souensembles", 6) & vbCrLf
   End If
    Fun6 = Fun6 & RetourLien_Veicule_Souensembles("select T_lien_Veicule_Souensembles.* from T_lien_Veicule_Souensembles ORDER BY T_lien_Veicule_Souensembles.id;", "T_lien_Veicule_Souensembles", "Id", "", "1", "OK", "")
End Function
Function Fun7()
   
    Fun7 = JavaValider("T_Lien_SousEnssemble_Nomenclature", 7) & vbCrLf
    Fun7 = Fun7 & RetourLien_SousEnssemble_Nomenclature("select T_Mouvement_de_stock.* from T_Mouvement_de_stock ORDER BY T_Mouvement_de_stock.id;", "T_Mouvement_de_stock", "Id", "", "1", "OK", "")
End Function

Function Fun8()
 If Session("Updatable") = "True" Then
    Fun8 = JavaValider("T_Nomenclature", 8) & vbCrLf
   Sql = Sql & "WHERE t_Vehicule.Id_Stock=1;"
   End If
    Fun8 = Fun8 & RecapSqlNomenclatuer("T_Nomenclature", "Id", "", "", "OK")
End Function
Function Fun9()
    Fun9 = RecapVehicule("T_Nomenclature")
End Function

Function majfunction()

If Request("MAJ") = "OK" Then
Sql = "SELECT " & Trim("" & Request("FormName")) & ".* FROM " & Trim("" & Request("FormName")) & ";"
Set Rs = OpenRecordSet(Sql, con)

    Sql = "UPDATE " & Request("FormName") & " Set "
sqlWere = "WHERE [" & Trim("" & Request("FormName")) & "].[" & Trim("" & Request("Champs0")) & "]="

   
    For L = 0 To Request("NbLigne")
        sqlval = ""
        For C = 1 To Request("NbCol")
        valeur = Request("txt_L" & L & "_C" & C)
        
        
            sqlval = sqlval & "[" & Trim("" & Request("FormName")) & "].[" & Trim("" & Request("Champs" & C)) & "]= " & noquote(valeur, RetourneTypeChamp(valeur, Rs.Fields(Trim("" & Request("Champs" & C))))) & ","
        Next
        sqlval = Trim("" & sqlval)
      sqlval = Left(sqlval, Len(sqlval) - 1) & " "
       'Response.Write sql  & sqlVal & sqlWere &  Request("DI_L" & L ) & ";"
'Response.End
If Trim("" & Request("DI_L" & L)) <> "" Then
       ExecuteSql Sql & sqlval & sqlWere & Request("DI_L" & L) & ";", con
End If
    If Request("Delet_L" & L) = "Del" Then
        ExecuteSql "DELETE FROM " & Request("FormName") & " WHERE [" & Request("FormName") & "].[" & _
        Request("Champs0") & "]=" & Request("DI_L" & L) & ";", con
    End If
    Next
If Request("Add") = "" Then
'Response.Write "toto"
'Response.End
Sql = "INSERT INTO " & Request("FormName") & " ("
   For C = 1 To Request("NbCol")
        Sql = Sql & "[" & Request("Champs" & C) & "],"
   Next
    Sql = Left(Sql, Len(Sql) - 1)
     Sql = Sql & ") Values ("


    For L = 0 To 9
        sqlval = ""
        If Request("Newtxt_L" & L & "_C1") <> "" Then
            For C = 1 To Request("NbCol")
            valeur = Request("Newtxt_L" & L & "_C" & C)
            
               sqlval = sqlval & noquote(valeur, RetourneTypeChamp(valeur, Rs.Fields(Trim("" & Request("Champs" & C))))) & ","
            Next
        End If
        If sqlval <> "" Then
            sqlval = Left(sqlval, Len(sqlval) - 1) & ");"
          ' Response.Write sql  & sqlval
           ' Response.End
            ExecuteSql Sql & sqlval, con
        End If
    Next
    End If
    Set Rs = CloseRecordSet(Rs)

End If
End Function

 
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'                OnStartPage
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub OnStartPage(PassedScriptingContext As ScriptingContext)
    Set MyScriptingContext = PassedScriptingContext
    Set Request = MyScriptingContext.Request
    Set Response = MyScriptingContext.Response
    Set Application = MyScriptingContext.Application
    Set SERVER = MyScriptingContext.SERVER
    Set Session = MyScriptingContext.Session
    'For i = 1 To Request.ServerVariables.Count
    '    Debug.Print i & " : " & Request.ServerVariables(i) & vbCrLf
    '
    'Next i
End Sub

Function RecapBR(FormName)
    RecapBR = RecapBR & vbCrLf & "<SCRIPT langage='JAVASCRIPT'>"
    RecapBR = RecapBR & vbCrLf & "function PageSuvante(Logistique_Page,id,id_Vehicule,Pièce,ID__SousEnsembles,BR){"

    RecapBR = RecapBR & vbCrLf & "document." & FormName & ".Logistique_Page.value=Logistique_Page;"
    RecapBR = RecapBR & vbCrLf & "document." & FormName & ".id.value=id;"
    RecapBR = RecapBR & vbCrLf & "document." & FormName & ".id_Vehicule.value=id_Vehicule;"
     RecapBR = RecapBR & vbCrLf & "document." & FormName & ".Pièce.value=Pièce;"
    RecapBR = RecapBR & vbCrLf & "document." & FormName & ".ID__SousEnsembles.value=ID__SousEnsembles;"
RecapBR = RecapBR & vbCrLf & "document." & FormName & ".BR.value=BR;"

    RecapBR = RecapBR & vbCrLf & "document." & FormName & ".action='BR.asp';"
    RecapBR = RecapBR & vbCrLf & "document." & FormName & ".submit();"



    RecapBR = RecapBR & vbCrLf & "}"
    RecapBR = RecapBR & vbCrLf & "</SCRIPT>"
    RecapBR = RecapBR & "<form name='" & FormName & "' method='post' >" & vbCrLf
    
    RecapBR = RecapBR & "<input name='Updatable' type='hidden' value='OK'/>" & vbCrLf
    RecapBR = RecapBR & "<input name='FormName' type='hidden' value='" & FormName & "'/>" & vbCrLf

    RecapBR = RecapBR & "<input name='Logistique_Page' type='hidden' value='10'/>" & vbCrLf
      
      
     
   
    
    
    sqlListe = sqlListe & "AND T_Nomenclature.Id="

        
        Sql = "SELECT T_SousEnsembles.ID_Vehicule_Véhicule, t_Vehicule.NumEncelade, t_Vehicule.Projet,  "
        Sql = Sql & "t_Vehicule.Programme, t_Vehicule.Désignation, t_Vehicule.Véhicule, T_Nomenclature.N_de_BL,  "
        Sql = Sql & "T_Nomenclature.Modifiee_le "
        Sql = Sql & "FROM (t_Vehicule LEFT JOIN T_SousEnsembles ON t_Vehicule.Id = T_SousEnsembles.ID_Vehicule_Véhicule)  "
        Sql = Sql & "LEFT JOIN T_Nomenclature ON T_SousEnsembles.Id = T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] "
        MyFilter = RetourneCloseWhere()
          If Request("nbSelect") > 0 Then
            Sql = Sql & MyCloseWhere()
             If Trim("" & MyFilter) <> "" Then
                Sql = Sql & " AND (" & MyFilter & ") "
             End If
          Else
            If Trim("" & MyFilter) <> "" Then
                Sql = Sql & " WHERE " & MyFilter
             End If
        End If
            
        Sql = Sql & "GROUP BY T_SousEnsembles.ID_Vehicule_Véhicule, t_Vehicule.NumEncelade, t_Vehicule.Projet, t_Vehicule.Programme,  "
        Sql = Sql & "t_Vehicule.Désignation, t_Vehicule.Véhicule, T_Nomenclature.N_de_BL, T_Nomenclature.Modifiee_le "
        Sql = Sql & "HAVING T_Nomenclature.N_de_BL Is Not Null "
        Sql = Sql & "ORDER BY T_Nomenclature.Modifiee_le DESC;"


     
'Response.Write sql
'Response.End

 
       
      
      
        Set Rsvehicule = OpenRecordSet(Sql, con)
      
         RecapBR = RecapBR & "<table border='1' width='100%'>" & vbCrLf
         If Request("nbSelect") > 0 Then
            RecapBR = RecapBR & "<tr>" & vbCrLf
            RecapBR = RecapBR & "         <td colspan='7' >" & RecapSersh & "</td>" & vbCrLf
            
            RecapBR = RecapBR & "</tr>" & vbCrLf
         End If
         RecapBR = RecapBR & "<tr >" & vbCrLf
         RecapBR = RecapBR & "<td class='TitreTableau' colspan ='8'>Liste des Bordereaux de Réception :</td>" & vbCrLf
         RecapBR = RecapBR & "</tr>" & vbCrLf
        
          
         RecapBR = RecapBR & "<tr>" & vbCrLf
          RecapBR = RecapBR & "<TD width='100%' colspan ='8'><table border='0'  width='100%'><tr>"
         
    RecapBR = RecapBR & vbCrLf & "<SCRIPT langage='JAVASCRIPT'>"
     
    RecapBR = RecapBR & "function ModifierRechereche(){" & vbCrLf
    RecapBR = RecapBR & "this.document." & FormName & ".ID__SousEnsembles.value='';"
    RecapBR = RecapBR & "this.document." & FormName & ".action='Default.asp?Logistique_Page=11'" & vbCrLf
    RecapBR = RecapBR & vbCrLf & "document." & FormName & ".submit();"
  RecapBR = RecapBR & "}" & vbCrLf
  
        RecapBR = RecapBR & "function valider3(){" & vbCrLf
        RecapBR = RecapBR & "this.document." & FormName & ".id_Vehicule.value='';"
        RecapBR = RecapBR & "this.document." & FormName & ".action='Default.asp?Logistique_Page=8'" & vbCrLf
        RecapBR = RecapBR & vbCrLf & "document." & FormName & ".submit();"
        RecapBR = RecapBR & "}" & vbCrLf

        RecapBR = RecapBR & "function valider2(){" & vbCrLf
        RecapBR = RecapBR & "this.document." & FormName & ".action='Default.asp?Logistique_Page=10'" & vbCrLf
        RecapBR = RecapBR & vbCrLf & "document." & FormName & ".submit();"
        RecapBR = RecapBR & "}" & vbCrLf
        RecapBR = RecapBR & vbCrLf & "</SCRIPT >"

       RecapBR = RecapBR & "<TD width='2' valign='top'>" & AfficheBounton("Retour", 8, "'javascript:valider3();'") & "</td>" & vbCrLf
      RecapBR = RecapBR & "<TD>&nbsp;</td><TD>&nbsp;</td>" & vbCrLf
        RecapBR = RecapBR & "<TD width='4%' valign='top'>Filtrés sur véhicule :" & vbCrLf
        Sql = "SELECT t_Vehicule.Id, [Projet] , [Véhicule]   FROM t_Vehicule ORDER BY t_Vehicule.Véhicule;"
        Set RsFiltre = OpenRecordSet(Sql, con)

       RecapBR = RecapBR & "<br>" & CreatListe(RsFiltre, Request("RecapFiltreSouEnsemble"), "Id", "1;2", "RecapFiltreSouEnsemble", " onChange='valider2();' ") & "</td>" & vbCrLf
      RsFiltre.Close
      Set RsFiltre = Nothing
      
     
     RecapBR = RecapBR & "</Tr></table></td>"
        
        RecapBR = RecapBR & "<tr>" & vbCrLf
       
        RecapBR = RecapBR & "         <td align='center' bgcolor='#CCCCCC'>N° Encelade</td>" & vbCrLf
        RecapBR = RecapBR & "         <td   align='center' bgcolor='#CCCCCC'>Projet</td>" & vbCrLf
        RecapBR = RecapBR & "         <td   align='center' bgcolor='#CCCCCC'>Programme</td>" & vbCrLf
        RecapBR = RecapBR & "         <td   align='center' bgcolor='#CCCCCC'>Désignation</td>" & vbCrLf
        RecapBR = RecapBR & "         <td   align='center' bgcolor='#CCCCCC'>Véhicule</td>" & vbCrLf
        RecapBR = RecapBR & "         <td   align='center' bgcolor='#CCCCCC'>BR</td>" & vbCrLf
         RecapBR = RecapBR & "         <td   align='center' bgcolor='#CCCCCC'>Date</td>" & vbCrLf
        RecapBR = RecapBR & "</tr> "
        RecapBR = RecapBR & "<input name='id' type='hidden' value='" & Request("id") & "' />"
       RecapBR = RecapBR & "<input name='id_Vehicule' type='hidden' value='" & Request("id_Vehicule") & "' />"
       RecapBR = RecapBR & "<input name='Pièce' type='hidden' value='" & Request("Pièce") & "' />"
       RecapBR = RecapBR & "<input name='ID__SousEnsembles' type='hidden' value='" & Request("ID__SousEnsembles") & "' />"
    RecapBR = RecapBR & "<input name='BR' type='hidden' value='" & Request("BR") & "' />"
        While Rsvehicule.EOF = False
        RecapBR = RecapBR & "<tr>"
           

        
        
       
        
        RecapBR = RecapBR & "             <td><a href =""javascript:PageSuvante('13','" & Request("id") & "','" & _
                          "0" & Request("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Rsvehicule("N_de_BL") & "');"">" & Rsvehicule("NumEncelade") & "&nbsp;</a> </td>" & vbCrLf
                         
         RecapBR = RecapBR & "             <td><a href =""javascript:PageSuvante('13','" & Request("id") & "','" & _
                          "0" & Request("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Rsvehicule("N_de_BL") & "');"">" & Rsvehicule("Projet") & "&nbsp;</a> </td>" & vbCrLf
                         
        RecapBR = RecapBR & "             <td><a href =""javascript:PageSuvante('13','" & Request("id") & "','" & _
                          "0" & Request("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Rsvehicule("N_de_BL") & "');"">" & Rsvehicule("Programme") & "&nbsp;</a> </td>" & vbCrLf
                         
        RecapBR = RecapBR & "             <td><a href =""javascript:PageSuvante('13','" & Request("id") & "','" & _
                          "0" & Request("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Rsvehicule("N_de_BL") & "');"">" & Replace("" & Rsvehicule("Désignation"), vbCrLf, "<br>") & "&nbsp;</a> </td>" & vbCrLf
                         
        RecapBR = RecapBR & "             <td><a href =""javascript:PageSuvante('13','" & Request("id") & "','" & _
                          "0" & Request("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Rsvehicule("N_de_BL") & "');"">" & Rsvehicule("Véhicule") & "&nbsp;</a> </td>" & vbCrLf
         RecapBR = RecapBR & "             <td><a href =""javascript:PageSuvante('13','" & Request("id") & "','" & _
                          "0" & Request("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Rsvehicule("N_de_BL") & "');"">" & Rsvehicule("N_de_BL") & "&nbsp;</a> </td>" & vbCrLf
     RecapBR = RecapBR & "             <td><a href =""javascript:PageSuvante('13','" & Request("id") & "','" & _
                          "0" & Request("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Rsvehicule("N_de_BL") & "');"">" & Rsvehicule("Modifiee_le") & "&nbsp;</a> </td>" & vbCrLf
        RecapBR = RecapBR & "</tr>"
       Rsvehicule.MoveNext
        Wend
     
    
    
       
       
       
          RecapBR = RecapBR & "</table>" & vbCrLf
          
          
          RecapBR = RecapBR & "<input name='NbS_Enssemble__' type='hidden' value='" & L2 & "'/>"
          RecapBR = RecapBR & "</form>"
      
       
End Function

Function RecapBL(FormName)
    RecapBL = RecapBL & vbCrLf & "<SCRIPT langage='JAVASCRIPT'>"
    RecapBL = RecapBL & vbCrLf & "function PageSuvante(Logistique_Page,id,id_Vehicule,Pièce,ID__SousEnsembles,BR){"

    RecapBL = RecapBL & vbCrLf & "document." & FormName & ".Logistique_Page.value=Logistique_Page;"
    RecapBL = RecapBL & vbCrLf & "document." & FormName & ".id.value=id;"
    RecapBL = RecapBL & vbCrLf & "document." & FormName & ".id_Vehicule.value=id_Vehicule;"
     RecapBL = RecapBL & vbCrLf & "document." & FormName & ".Pièce.value=Pièce;"
    RecapBL = RecapBL & vbCrLf & "document." & FormName & ".ID__SousEnsembles.value=ID__SousEnsembles;"
RecapBL = RecapBL & vbCrLf & "document." & FormName & ".BR.value=BR;"

    RecapBL = RecapBL & vbCrLf & "document." & FormName & ".action='BL.asp';"
    RecapBL = RecapBL & vbCrLf & "document." & FormName & ".submit();"



    RecapBL = RecapBL & vbCrLf & "}"
    RecapBL = RecapBL & vbCrLf & "</SCRIPT>"
    RecapBL = RecapBL & "<form name='" & FormName & "' method='post' >" & vbCrLf
    
    RecapBL = RecapBL & "<input name='Updatable' type='hidden' value='OK'/>" & vbCrLf
    RecapBL = RecapBL & "<input name='FormName' type='hidden' value='" & FormName & "'/>" & vbCrLf

    RecapBL = RecapBL & "<input name='Logistique_Page' type='hidden' value='10'/>" & vbCrLf
      
      
     
   
    
    
    sqlListe = sqlListe & "AND T_Nomenclature.Id="

        
        Sql = "SELECT T_SousEnsembles.ID_Vehicule_Véhicule, t_Vehicule.NumEncelade, t_Vehicule.Projet,  "
        Sql = Sql & "t_Vehicule.Programme, t_Vehicule.Désignation, t_Vehicule.Véhicule, T_Nomenclature.Num_de_retour,  "
        Sql = Sql & "T_Nomenclature.Modifiee_le "
        Sql = Sql & "FROM (t_Vehicule LEFT JOIN T_SousEnsembles ON t_Vehicule.Id = T_SousEnsembles.ID_Vehicule_Véhicule)  "
        Sql = Sql & "LEFT JOIN T_Nomenclature ON T_SousEnsembles.Id = T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] "
        MyFilter = RetourneCloseWhere()
          If Request("nbSelect") > 0 Then
            Sql = Sql & MyCloseWhere()
             If Trim("" & MyFilter) <> "" Then
                Sql = Sql & " AND (" & MyFilter & ") "
             End If
          Else
            If Trim("" & MyFilter) <> "" Then
                Sql = Sql & " WHERE " & MyFilter
             End If
        End If
            
        Sql = Sql & "GROUP BY T_SousEnsembles.ID_Vehicule_Véhicule, t_Vehicule.NumEncelade, t_Vehicule.Projet, t_Vehicule.Programme,  "
        Sql = Sql & "t_Vehicule.Désignation, t_Vehicule.Véhicule, T_Nomenclature.Num_de_retour, T_Nomenclature.Modifiee_le "
        Sql = Sql & "HAVING T_Nomenclature.Num_de_retour Is Not Null "
        Sql = Sql & "ORDER BY T_Nomenclature.Modifiee_le DESC;"


     
'Response.Write sql
'Response.End

 
       
      
      
        Set Rsvehicule = OpenRecordSet(Sql, con)
      
         RecapBL = RecapBL & "<table border='1' width='100%'>" & vbCrLf
         If Request("nbSelect") > 0 Then
            RecapBL = RecapBL & "<tr>" & vbCrLf
            RecapBL = RecapBL & "         <td colspan='7' >" & RecapBL & "</td>" & vbCrLf
            
            RecapBL = RecapBL & "</tr>" & vbCrLf
         End If
        RecapBL = RecapBL & "<tr >" & vbCrLf
         RecapBL = RecapBL & "<td class='TitreTableau' colspan ='8'>Liste des Bordereaux de Livraison :</td>" & vbCrLf
         RecapBL = RecapBL & "</tr>" & vbCrLf
        
          
         RecapBL = RecapBL & "<tr>" & vbCrLf
          RecapBL = RecapBL & "<TD width='100%' colspan ='8'><table border='0'  width='100%'><tr>"
         
    RecapBL = RecapBL & vbCrLf & "<SCRIPT langage='JAVASCRIPT'>"
    RecapBL = RecapBL & "function ModifierRechereche(){" & vbCrLf
    RecapBL = RecapBL & "this.document." & FormName & ".ID__SousEnsembles.value='';"
    RecapBL = RecapBL & "this.document." & FormName & ".action='Default.asp?Logistique_Page=11'" & vbCrLf
    RecapBL = RecapBL & vbCrLf & "document." & FormName & ".submit();"
    RecapBL = RecapBL & "}" & vbCrLf
    
        RecapBL = RecapBL & "function valider3(){" & vbCrLf
        RecapBL = RecapBL & "this.document." & FormName & ".id_Vehicule.value='';"
        RecapBL = RecapBL & "this.document." & FormName & ".action='Default.asp?Logistique_Page=8'" & vbCrLf
        RecapBL = RecapBL & vbCrLf & "document." & FormName & ".submit();"
        RecapBL = RecapBL & "}" & vbCrLf

        RecapBL = RecapBL & "function valider2(){" & vbCrLf
        RecapBL = RecapBL & "this.document." & FormName & ".action='Default.asp?Logistique_Page=10'" & vbCrLf
        RecapBL = RecapBL & vbCrLf & "document." & FormName & ".submit();"
        RecapBL = RecapBL & "}" & vbCrLf
        RecapBL = RecapBL & vbCrLf & "</SCRIPT >"

       RecapBL = RecapBL & "<TD width='2' valign='top'>" & AfficheBounton("Retour", 8, "'javascript:valider3();'") & "</td>" & vbCrLf
      RecapBL = RecapBL & "<TD>&nbsp;</td><TD>&nbsp;</td>" & vbCrLf
        RecapBL = RecapBL & "<TD width='4%' valign='top'>Filtrés sur véhicule :" & vbCrLf
        Sql = "SELECT t_Vehicule.Id, [Projet] , [Véhicule]   FROM t_Vehicule ORDER BY t_Vehicule.Véhicule;"
        Set RsFiltre = OpenRecordSet(Sql, con)

       RecapBL = RecapBL & "<br>" & CreatListe(RsFiltre, Request("RecapFiltreSouEnsemble"), "Id", "1;2", "RecapFiltreSouEnsemble", " onChange='valider2();' ") & "</td>" & vbCrLf
      RsFiltre.Close
      Set RsFiltre = Nothing
      
     
     RecapBL = RecapBL & "</Tr></table></td>"
        
        RecapBL = RecapBL & "<tr>" & vbCrLf
        RecapBL = RecapBL & "         <td align='center' bgcolor='#CCCCCC'>N° Encelade</td>" & vbCrLf
        RecapBL = RecapBL & "         <td   align='center' bgcolor='#CCCCCC'>Projet</td>" & vbCrLf
        RecapBL = RecapBL & "         <td   align='center' bgcolor='#CCCCCC'>Programme</td>" & vbCrLf
        RecapBL = RecapBL & "         <td   align='center' bgcolor='#CCCCCC'>Désignation</td>" & vbCrLf
        RecapBL = RecapBL & "         <td   align='center' bgcolor='#CCCCCC'>Véhicule</td>" & vbCrLf
        RecapBL = RecapBL & "         <td   align='center' bgcolor='#CCCCCC'>BL</td>" & vbCrLf
         RecapBL = RecapBL & "         <td   align='center' bgcolor='#CCCCCC'>Date</td>" & vbCrLf
        RecapBL = RecapBL & "</tr> "
        RecapBL = RecapBL & "<input name='id' type='hidden' value='" & Request("id") & "' />"
       RecapBL = RecapBL & "<input name='id_Vehicule' type='hidden' value='" & Request("id_Vehicule") & "' />"
       RecapBL = RecapBL & "<input name='Pièce' type='hidden' value='" & Request("Pièce") & "' />"
       RecapBL = RecapBL & "<input name='ID__SousEnsembles' type='hidden' value='" & Request("ID__SousEnsembles") & "' />"
    RecapBL = RecapBL & "<input name='BR' type='hidden' value='" & Request("BR") & "' />"
        While Rsvehicule.EOF = False
        RecapBL = RecapBL & "<tr>"
           

        
        
       
        
        RecapBL = RecapBL & "             <td><a href =""javascript:PageSuvante('13','" & Request("id") & "','" & _
                          "0" & Request("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Rsvehicule("Num_de_retour") & "');"">" & Rsvehicule("NumEncelade") & "&nbsp;</a> </td>" & vbCrLf
                         
         RecapBL = RecapBL & "             <td><a href =""javascript:PageSuvante('13','" & Request("id") & "','" & _
                          "0" & Request("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Rsvehicule("Num_de_retour") & "');"">" & Rsvehicule("Projet") & "&nbsp;</a> </td>" & vbCrLf
                         
        RecapBL = RecapBL & "             <td><a href =""javascript:PageSuvante('13','" & Request("id") & "','" & _
                          "0" & Request("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Rsvehicule("Num_de_retour") & "');"">" & Rsvehicule("Programme") & "&nbsp;</a> </td>" & vbCrLf
                         
        RecapBL = RecapBL & "             <td><a href =""javascript:PageSuvante('13','" & Request("id") & "','" & _
                          "0" & Request("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Rsvehicule("Num_de_retour") & "');"">" & Replace("" & Rsvehicule("Désignation"), vbCrLf, "<br>") & "&nbsp;</a> </td>" & vbCrLf
                         
        RecapBL = RecapBL & "             <td><a href =""javascript:PageSuvante('13','" & Request("id") & "','" & _
                          "0" & Request("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Rsvehicule("Num_de_retour") & "');"">" & Rsvehicule("Véhicule") & "&nbsp;</a> </td>" & vbCrLf
         RecapBL = RecapBL & "             <td><a href =""javascript:PageSuvante('13','" & Request("id") & "','" & _
                          "0" & Request("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Rsvehicule("Num_de_retour") & "');"">" & Rsvehicule("Num_de_retour") & "&nbsp;</a> </td>" & vbCrLf
     RecapBL = RecapBL & "             <td><a href =""javascript:PageSuvante('13','" & Request("id") & "','" & _
                          "0" & Request("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Rsvehicule("Num_de_retour") & "');"">" & Rsvehicule("Modifiee_le") & "&nbsp;</a> </td>" & vbCrLf
        RecapBL = RecapBL & "</tr>"
       Rsvehicule.MoveNext
        Wend
     
    
    
       
       
       
          RecapBL = RecapBL & "</table>" & vbCrLf
          
          
          RecapBL = RecapBL & "<input name='NbS_Enssemble__' type='hidden' value='" & L2 & "'/>"
          RecapBL = RecapBL & "</form>"
      
       
End Function

 
Function RecapSousEnssemble(FormName)
RecapSousEnssemble = RecapSousEnssemble & vbCrLf & "<SCRIPT langage='JAVASCRIPT'>"
    RecapSousEnssemble = RecapSousEnssemble & vbCrLf & "function PageSuvante(Logistique_Page,id,id_Vehicule,Pièce,ID__SousEnsembles,Id_Projects_NumAffaire){"

    RecapSousEnssemble = RecapSousEnssemble & vbCrLf & "document." & FormName & ".Logistique_Page.value=Logistique_Page;"
    RecapSousEnssemble = RecapSousEnssemble & vbCrLf & "document." & FormName & ".id.value=id;"
    RecapSousEnssemble = RecapSousEnssemble & vbCrLf & "document." & FormName & ".id_Vehicule.value=id_Vehicule;"
    RecapSousEnssemble = RecapSousEnssemble & vbCrLf & "document." & FormName & ".Pièce.value=Pièce;"
    RecapSousEnssemble = RecapSousEnssemble & vbCrLf & "document." & FormName & ".ID__SousEnsembles.value=ID__SousEnsembles;"
    RecapSousEnssemble = RecapSousEnssemble & vbCrLf & "document." & FormName & ".Id_Projects_NumAffaire.value=Id_Projects_NumAffaire;"

    RecapSousEnssemble = RecapSousEnssemble & vbCrLf & "document." & FormName & ".action='Default.asp';"
    RecapSousEnssemble = RecapSousEnssemble & vbCrLf & "document." & FormName & ".submit();"



    RecapSousEnssemble = RecapSousEnssemble & vbCrLf & "}"
    RecapSousEnssemble = RecapSousEnssemble & vbCrLf & "</SCRIPT>"
    
    
    RecapSousEnssemble = RecapSousEnssemble & "<form name='" & FormName & "' method='post'>" & vbCrLf
    
     
      RecapSousEnssemble = RecapSousEnssemble & "<input name='FormName' type='hidden' value='" & FormName & "'/>" & vbCrLf
        Sql = "SELECT T_SousEnsembles.Id, T_Projects.Projet, t_Vehicule.Véhicule, T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble],  "
        Sql = Sql & "T_SousEnsembles.Code, T_SousEnsembles.[Sous-Ensseble], T_SousEnsembles.[MASSE total (KG)],  "
        Sql = Sql & "T_SousEnsembles.[Nb lignes sans masse], T_SousEnsembles.[Nb total de lignes] "
        Sql = Sql & "FROM T_Projects INNER JOIN "
        Sql = Sql & "t_Vehicule ON T_Projects.Id = t_Vehicule.Id_Projects_NumAffaire RIGHT OUTER JOIN "
        Sql = Sql & "T_SousEnsembles LEFT OUTER JOIN "
        Sql = Sql & "T_Nomenclature ON T_SousEnsembles.Id = T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble]  "
        Sql = Sql & "ON t_Vehicule.Id = T_SousEnsembles.ID_Vehicule_Véhicule "
         MyFilter = RetourneCloseWhere()
          If Request("nbSelect") > 0 Then
            Sql = Sql & MyCloseWhere()
             If Trim("" & MyFilter) <> "" Then
                Sql = Sql & " AND (" & MyFilter & ") "
             End If
          Else
            If Trim("" & MyFilter) <> "" Then
                Sql = Sql & " WHERE " & MyFilter
             End If
        End If
            If Trim("" & Request("RecapFiltreSouEnsemble")) <> "" Then
                     If InStr(1, UCase(Sql), "WHERE") <> 0 Then
                    
                Sql = Sql & " AND t_Vehicule.Id=" & Request("RecapFiltreSouEnsemble") & " "
                
            Else
                MyFilter = "?"
                 Sql = Sql & " WHERE t_Vehicule.Id=" & Request("RecapFiltreSouEnsemble") & " "
            End If
      End If
       
            
        Sql = Sql & "GROUP BY T_SousEnsembles.Id, T_Projects.Projet, t_Vehicule.Véhicule,  "
        Sql = Sql & "T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble], T_SousEnsembles.Code, T_SousEnsembles.[Sous-Ensseble],  "
        Sql = Sql & "T_SousEnsembles.[MASSE total (KG)], T_SousEnsembles.[Nb lignes sans masse],  "
        Sql = Sql & "T_SousEnsembles.[Nb total de lignes];"
      
       'Response.Write sql
        'response.End
        
      
      
        Set RsSousEnsemble = OpenRecordSet(Sql, con)
       
         RecapSousEnssemble = RecapSousEnssemble & "<table border='1' width='100%'>" & vbCrLf
          If Request("nbSelect") > 0 Then
            RecapSousEnssemble = RecapSousEnssemble & "<tr>" & vbCrLf
            RecapSousEnssemble = RecapSousEnssemble & "         <td colspan='8' >" & RecapSersh & "</td>" & vbCrLf
            
            RecapSousEnssemble = RecapSousEnssemble & "</tr>" & vbCrLf
         End If
          RecapSousEnssemble = RecapSousEnssemble & "<tr >" & vbCrLf
         RecapSousEnssemble = RecapSousEnssemble & "<td class='TitreTableau' colspan ='8'>Liste des S/Enssembles :</td>" & vbCrLf
         RecapSousEnssemble = RecapSousEnssemble & "</tr>" & vbCrLf
        
          
         RecapSousEnssemble = RecapSousEnssemble & "<tr>" & vbCrLf
          RecapSousEnssemble = RecapSousEnssemble & "<TD width='100%' colspan ='8'><table border='0'  width='100%'><tr>"
         
    RecapSousEnssemble = RecapSousEnssemble & vbCrLf & "<SCRIPT langage='JAVASCRIPT'>"
    RecapSousEnssemble = RecapSousEnssemble & "function ModifierRechereche(){" & vbCrLf
    RecapSousEnssemble = RecapSousEnssemble & "this.document." & FormName & ".ID__SousEnsembles.value='';"
    RecapSousEnssemble = RecapSousEnssemble & "this.document." & FormName & ".action='Default.asp?Logistique_Page=11'" & vbCrLf
    RecapSousEnssemble = RecapSousEnssemble & vbCrLf & "document." & FormName & ".submit();"
    RecapSousEnssemble = RecapSousEnssemble & "}" & vbCrLf
    
        RecapSousEnssemble = RecapSousEnssemble & "function valider3(){" & vbCrLf
        RecapSousEnssemble = RecapSousEnssemble & "this.document." & FormName & ".id_Vehicule.value='';"
'        Id_Projects_NumAffaire
        
        RecapSousEnssemble = RecapSousEnssemble & "this.document." & FormName & ".action='Default.asp?Logistique_Page=9'" & vbCrLf
        RecapSousEnssemble = RecapSousEnssemble & vbCrLf & "document." & FormName & ".submit();"
        RecapSousEnssemble = RecapSousEnssemble & "}" & vbCrLf

        RecapSousEnssemble = RecapSousEnssemble & "function valider2(){" & vbCrLf
        RecapSousEnssemble = RecapSousEnssemble & "this.document." & FormName & ".action='Default.asp?Logistique_Page=10'" & vbCrLf
        RecapSousEnssemble = RecapSousEnssemble & vbCrLf & "document." & FormName & ".submit();"
        RecapSousEnssemble = RecapSousEnssemble & "}" & vbCrLf
        RecapSousEnssemble = RecapSousEnssemble & vbCrLf & "</SCRIPT >"

       RecapSousEnssemble = RecapSousEnssemble & "<TD width='2' valign='top'>" & AfficheBounton("Retour", 8, "'javascript:valider3();'") & "</td>" & vbCrLf
      RecapSousEnssemble = RecapSousEnssemble & "<TD>&nbsp;</td><TD>&nbsp;</td>" & vbCrLf
        RecapSousEnssemble = RecapSousEnssemble & "<TD width='4%' valign='top'>Filtrés sur véhicule :" & vbCrLf
        Sql = "SELECT t_Vehicule.Id, [Projet] , [Véhicule]   FROM t_Vehicule ORDER BY t_Vehicule.Véhicule;"
        
         Sql = "SELECT     t_Vehicule.Id, T_Projects.NumAffaire, t_Vehicule.Véhicule "
        Sql = Sql & "FROM t_Vehicule INNER JOIN "
        Sql = Sql & "T_Projects ON t_Vehicule.Id_Projects_NumAffaire = T_Projects.Id "
        Sql = Sql & "ORDER BY t_Vehicule.Véhicule"
        Set RsFiltre = OpenRecordSet(Sql, con)

       RecapSousEnssemble = RecapSousEnssemble & "<br>" & CreatListe(RsFiltre, Request("RecapFiltreSouEnsemble"), "Id", "1;2", "RecapFiltreSouEnsemble", " onChange='valider2();' ") & "</td>" & vbCrLf
      RsFiltre.Close
      Set RsFiltre = Nothing
      
     
     RecapSousEnssemble = RecapSousEnssemble & "</Tr></table></td>"
        RecapSousEnssemble = RecapSousEnssemble & "<tr>" & vbCrLf
            RecapSousEnssemble = RecapSousEnssemble & "         <td align='center' bgcolor='#CCCCCC'>&nbsp;</td>" & vbCrLf
        RecapSousEnssemble = RecapSousEnssemble & "         <td   align='center' bgcolor='#CCCCCC'>Projet</td>" & vbCrLf
        RecapSousEnssemble = RecapSousEnssemble & "         <td   align='center' bgcolor='#CCCCCC'>Véhicule</td>" & vbCrLf

        RecapSousEnssemble = RecapSousEnssemble & "         <td   align='center' bgcolor='#CCCCCC'>Code</td>" & vbCrLf
        RecapSousEnssemble = RecapSousEnssemble & "         <td   align='center' bgcolor='#CCCCCC'>Sous-Ensseble</td>" & vbCrLf
        RecapSousEnssemble = RecapSousEnssemble & "         <td   align='center' bgcolor='#CCCCCC'>MASSE total (KG)</td>" & vbCrLf
        RecapSousEnssemble = RecapSousEnssemble & "         <td   align='center' bgcolor='#CCCCCC'>Nb lignes sans masse</td>" & vbCrLf
        RecapSousEnssemble = RecapSousEnssemble & "         <td   align='center' bgcolor='#CCCCCC'>Nb total de lignes</td>" & vbCrLf
       
        RecapSousEnssemble = RecapSousEnssemble & "</tr> "
       RecapSousEnssemble = RecapSousEnssemble & "<input name='id' type='hidden' value='" & Request("id") & "' />"
       RecapSousEnssemble = RecapSousEnssemble & "<input name='id_Vehicule' type='hidden' value='" & Request("id_Vehicule") & "' />"
       RecapSousEnssemble = RecapSousEnssemble & "<input name='Pièce' type='hidden' value='" & Request("Pièce") & "' />"
       RecapSousEnssemble = RecapSousEnssemble & "<input name='ID__SousEnsembles' type='hidden' value='" & Request("ID__SousEnsembles") & "' />"
       RecapSousEnssemble = RecapSousEnssemble & "<input name='Id_Projects_NumAffaire' type='hidden' value='" & Request("Id_Projects_NumAffaire") & "' />"
    RecapSousEnssemble = RecapSousEnssemble & "<input name='Logistique_Page' type='hidden' value='9'/>" & vbCrLf
        While RsSousEnsemble.EOF = False
        RecapSousEnssemble = RecapSousEnssemble & "<tr>"
           


               
               
                         
  RecapSousEnssemble = RecapSousEnssemble & "             <td><a href =""javascript:PageSuvante('5','" & Request("id") & "','" & _
                         Request("ID_Vehicule") & "','" & Request("Pièce") & "','" & _
                         "0" & RsSousEnsemble("ID") & "','" & Request("Id_Projects_NumAffaire") & "');""><img src='./Images/zoom.gif' border='0' /></a></td>" & vbCrLf
  
   RecapSousEnssemble = RecapSousEnssemble & "             <td><a href =""javascript:PageSuvante('8','" & Request("id") & "','" & _
                         Request("ID_Vehicule") & "','" & Request("Pièce") & "','" & _
                         "0" & RsSousEnsemble("ID_SousEnsembles_Sous-Ensseble") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & RsSousEnsemble("Projet") & "&nbsp;</a> </td>" & vbCrLf
                         
  RecapSousEnssemble = RecapSousEnssemble & "             <td><a href =""javascript:PageSuvante('8','" & Request("id") & "','" & _
                         Request("ID_Vehicule") & "','" & Request("Pièce") & "','" & _
                         "0" & RsSousEnsemble("ID_SousEnsembles_Sous-Ensseble") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & RsSousEnsemble("Véhicule") & "&nbsp;</a> </td>" & vbCrLf
        
        RecapSousEnssemble = RecapSousEnssemble & "             <td><a href =""javascript:PageSuvante('8','" & Request("id") & "','" & _
                         Request("ID_Vehicule") & "','" & Request("Pièce") & "','" & _
                        "0" & RsSousEnsemble("ID_SousEnsembles_Sous-Ensseble") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & RsSousEnsemble("Code") & "&nbsp;</a> </td>" & vbCrLf
                         
       RecapSousEnssemble = RecapSousEnssemble & "             <td><a href =""javascript:PageSuvante('8','" & Request("id") & "','" & _
                         Request("ID_Vehicule") & "','" & Request("Pièce") & "','" & _
                         "0" & RsSousEnsemble("ID_SousEnsembles_Sous-Ensseble") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & RsSousEnsemble("Sous-Ensseble") & "&nbsp;</a> </td>" & vbCrLf
                         
      RecapSousEnssemble = RecapSousEnssemble & "             <td><a href =""javascript:PageSuvante('8','" & Request("id") & "','" & _
                         Request("ID_Vehicule") & "','" & Request("Pièce") & "','" & _
                        "0" & RsSousEnsemble("ID_SousEnsembles_Sous-Ensseble") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & RsSousEnsemble("MASSE total (KG)") & "&nbsp;</a> </td>" & vbCrLf
                         
        RecapSousEnssemble = RecapSousEnssemble & "             <td><a href =""javascript:PageSuvante('8','" & Request("id") & "','" & _
                         Request("ID_Vehicule") & "','" & Request("Pièce") & "','" & _
                         "0" & RsSousEnsemble("ID_SousEnsembles_Sous-Ensseble") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & RsSousEnsemble("Nb lignes sans masse") & "&nbsp;</a> </td>" & vbCrLf
                         
         RecapSousEnssemble = RecapSousEnssemble & "             <td><a href =""javascript:PageSuvante('8','" & Request("id") & "','" & _
                         Request("ID_Vehicule") & "','" & Request("Pièce") & "','" & _
                         "0" & RsSousEnsemble("ID_SousEnsembles_Sous-Ensseble") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & RsSousEnsemble("Nb total de lignes") & "&nbsp;</a> </td>" & vbCrLf
        RecapSousEnssemble = RecapSousEnssemble & "</tr>"
       RsSousEnsemble.MoveNext
        Wend
     
    
    
       
       
       
          RecapSousEnssemble = RecapSousEnssemble & "</table>" & vbCrLf
         
          
          RecapSousEnssemble = RecapSousEnssemble & "<input name='NbS_Enssemble__' type='hidden' value='" & L2 & "'/>"
          RecapSousEnssemble = RecapSousEnssemble & "</form>"
      
       
End Function

Function RecapVehicule(FormName)


    RecapVehicule = RecapVehicule & vbCrLf & "<SCRIPT langage='JAVASCRIPT'>"
    RecapVehicule = RecapVehicule & vbCrLf & "function PageSuvante(Logistique_Page,id,id_Vehicule,Pièce,ID__SousEnsembles,Id_Projects_NumAffaire){"

    RecapVehicule = RecapVehicule & vbCrLf & "document." & FormName & ".Logistique_Page.value=Logistique_Page;"
    RecapVehicule = RecapVehicule & vbCrLf & "document." & FormName & ".id.value=id;"
    RecapVehicule = RecapVehicule & vbCrLf & "document." & FormName & ".id_Vehicule.value=id_Vehicule;"
     RecapVehicule = RecapVehicule & vbCrLf & "document." & FormName & ".Pièce.value=Pièce;"
    RecapVehicule = RecapVehicule & vbCrLf & "document." & FormName & ".ID__SousEnsembles.value=ID__SousEnsembles;"
  RecapVehicule = RecapVehicule & vbCrLf & "document." & FormName & ".Id_Projects_NumAffaire.value=Id_Projects_NumAffaire;"
    RecapVehicule = RecapVehicule & vbCrLf & "document." & FormName & ".action='Default.asp';"
    RecapVehicule = RecapVehicule & vbCrLf & "document." & FormName & ".submit();"



    RecapVehicule = RecapVehicule & vbCrLf & "}"
    RecapVehicule = RecapVehicule & vbCrLf & "</SCRIPT>"
    RecapVehicule = RecapVehicule & "<form name='" & FormName & "' method='post'>" & vbCrLf
    
    RecapVehicule = RecapVehicule & "<input name='Updatable' type='hidden' value='OK'/>" & vbCrLf
    RecapVehicule = RecapVehicule & "<input name='FormName' type='hidden' value='" & FormName & "'/>" & vbCrLf

    RecapVehicule = RecapVehicule & "<input name='Logistique_Page' type='hidden' value='10'/>" & vbCrLf
      
      
     
   
    
    
    sqlListe = sqlListe & "AND T_Nomenclature.Id="

        
        Sql = "SELECT T_SousEnsembles.ID_Vehicule_Véhicule, t_Vehicule.NumEncelade, T_Projects.Projet,  "
        Sql = Sql & "t_Vehicule.Programme, t_Vehicule.Désignation, t_Vehicule.Véhicule "
        Sql = Sql & "FROM t_Vehicule INNER JOIN "
        Sql = Sql & "T_Projects ON t_Vehicule.Id_Projects_NumAffaire = T_Projects.Id LEFT OUTER JOIN "
        Sql = Sql & "T_SousEnsembles ON t_Vehicule.Id = T_SousEnsembles.ID_Vehicule_Véhicule LEFT OUTER JOIN "
        Sql = Sql & "T_Nomenclature ON T_SousEnsembles.Id = T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] "
        MyFilter = RetourneCloseWhere()
          If Request("nbSelect") > 0 Then
            Sql = Sql & MyCloseWhere()
             If Trim("" & MyFilter) <> "" Then
                Sql = Sql & " AND (" & MyFilter & ") "
             End If
          Else
            If Trim("" & MyFilter) <> "" Then
                Sql = Sql & " WHERE " & MyFilter
             End If
        End If
        If Trim("" & Request("RecapFiltreVehicule")) <> "" Then
                    If InStr(1, UCase(Sql), "WHERE") <> 0 Then
                    
                Sql = Sql & " AND t_Vehicule.Id=" & Request("RecapFiltreVehicule") & " "
                
            Else
                MyFilter = "?"
                 Sql = Sql & " WHERE t_Vehicule.Id=" & Request("RecapFiltreVehicule") & " "
            End If
      End If
            
        Sql = Sql & "GROUP BY T_SousEnsembles.ID_Vehicule_Véhicule, t_Vehicule.NumEncelade,   "
        Sql = Sql & "T_Projects.Projet, t_Vehicule.Programme, t_Vehicule.Désignation, t_Vehicule.Véhicule;"


     
'Response.Write sql
'Response.End

 
       
      
      
        Set Rsvehicule = OpenRecordSet(Sql, con)
      
         RecapVehicule = RecapVehicule & "<table border='1' width='100%'>" & vbCrLf
         If Request("nbSelect") > 0 Then
            RecapVehicule = RecapVehicule & "<tr>" & vbCrLf
            RecapVehicule = RecapVehicule & "         <td colspan='6' >" & RecapSersh & "</td>" & vbCrLf
            
            RecapVehicule = RecapVehicule & "</tr>" & vbCrLf
         End If
          RecapVehicule = RecapVehicule & "<tr >" & vbCrLf
         RecapVehicule = RecapVehicule & "<td class='TitreTableau' colspan ='8'>Liste des Véhicules :</td>" & vbCrLf
         RecapVehicule = RecapVehicule & "</tr>" & vbCrLf
         RecapVehicule = RecapVehicule & "<tr><td colspan ='8'><table><tr>" & vbCrLf
          
'        RecapVehicule = RecapVehicule & "<TD width='2' valign='top' colspan ='8'>" & AfficheBounton("Retour", 8, "'javascript:valider3();'") & "</td></tr>" & vbCrLf
          
         RecapVehicule = RecapVehicule & "<tr>" & vbCrLf
         
         
    RecapVehicule = RecapVehicule & vbCrLf & "<SCRIPT langage='JAVASCRIPT'>"
     
    RecapVehicule = RecapVehicule & "function ModifierRechereche(){" & vbCrLf
    RecapVehicule = RecapVehicule & "this.document." & FormName & ".ID__SousEnsembles.value='';"
    RecapVehicule = RecapVehicule & "this.document." & FormName & ".action='Default.asp?Logistique_Page=11'" & vbCrLf
    RecapVehicule = RecapVehicule & vbCrLf & "document." & FormName & ".submit();"
    RecapVehicule = RecapVehicule & "}" & vbCrLf
    
    
        RecapVehicule = RecapVehicule & "function valider3(){" & vbCrLf
        RecapVehicule = RecapVehicule & "this.document." & FormName & ".Id_Projects_NumAffaire.value='';"
        'Id_Projects_NumAffaire
         
        RecapVehicule = RecapVehicule & "this.document." & FormName & ".action='Default.asp?Logistique_Page=22'" & vbCrLf
        RecapVehicule = RecapVehicule & vbCrLf & "document." & FormName & ".submit();"
        RecapVehicule = RecapVehicule & "}" & vbCrLf

        RecapVehicule = RecapVehicule & "function valider2(){" & vbCrLf
        RecapVehicule = RecapVehicule & "this.document." & FormName & ".action='Default.asp?Logistique_Page=9'" & vbCrLf
        RecapVehicule = RecapVehicule & vbCrLf & "document." & FormName & ".submit();"
        RecapVehicule = RecapVehicule & "}" & vbCrLf
        RecapVehicule = RecapVehicule & vbCrLf & "</SCRIPT >"

      
      RecapVehicule = RecapVehicule & "<TD width='2%'> " & AfficheBounton("Retour", 8, "'javascript:valider3();'") & "</td><TD>&nbsp;</td>" & vbCrLf
        RecapVehicule = RecapVehicule & "<TD width='4%' valign='top'>Filtrés sur véhicule :" & vbCrLf
        Sql = "SELECT t_Vehicule.Id, [Projet] , [Véhicule]  FROM t_Vehicule ORDER BY t_Vehicule.Véhicule;"
        
         Sql = "SELECT     t_Vehicule.Id, T_Projects.NumAffaire, t_Vehicule.Véhicule "
        Sql = Sql & "FROM t_Vehicule INNER JOIN "
        Sql = Sql & "T_Projects ON t_Vehicule.Id_Projects_NumAffaire = T_Projects.Id "
        Sql = Sql & "ORDER BY t_Vehicule.Véhicule"
        Set RsFiltre = OpenRecordSet(Sql, con)

       RecapVehicule = RecapVehicule & "<br>" & CreatListe(RsFiltre, Request("RecapFiltreVehicule"), "Id", "1;2", "RecapFiltreVehicule", " onChange='valider2();' ") & "</td>" & vbCrLf
      RsFiltre.Close
      Set RsFiltre = Nothing
      
     
     RecapVehicule = RecapVehicule & "</Tr></table></td>"
        RecapVehicule = RecapVehicule & "<tr>" & vbCrLf
        RecapVehicule = RecapVehicule & "         <td align='center' bgcolor='#CCCCCC'>&nbsp;</td>" & vbCrLf
        RecapVehicule = RecapVehicule & "         <td align='center' bgcolor='#CCCCCC'>N° Encelade</td>" & vbCrLf
        RecapVehicule = RecapVehicule & "         <td   align='center' bgcolor='#CCCCCC'>Projet</td>" & vbCrLf
        RecapVehicule = RecapVehicule & "         <td   align='center' bgcolor='#CCCCCC'>Programme</td>" & vbCrLf
        RecapVehicule = RecapVehicule & "         <td   align='center' bgcolor='#CCCCCC'>Désignation</td>" & vbCrLf
        RecapVehicule = RecapVehicule & "         <td   align='center' bgcolor='#CCCCCC'>Véhicule</td>" & vbCrLf
       
        RecapVehicule = RecapVehicule & "</tr> "
        RecapVehicule = RecapVehicule & "<input name='id' type='hidden' value='" & Request("id") & "' />"
       RecapVehicule = RecapVehicule & "<input name='id_Vehicule' type='hidden' value='" & Request("id_Vehicule") & "' />"
       RecapVehicule = RecapVehicule & "<input name='Pièce' type='hidden' value='" & Request("Pièce") & "' />"
       RecapVehicule = RecapVehicule & "<input name='ID__SousEnsembles' type='hidden' value='" & Request("ID__SousEnsembles") & "' />"
    RecapVehicule = RecapVehicule & "<input name='Id_Projects_NumAffaire' type='hidden' value='" & Request("Id_Projects_NumAffaire") & "' />"
        While Rsvehicule.EOF = False
        RecapVehicule = RecapVehicule & "<tr>"
           

        
        
       
         RecapVehicule = RecapVehicule & "             <td align='center'><a href =""javascript:PageSuvante('2','" & Request("id") & "','" & _
                        "" & Rsvehicule("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Request("Id_Projects_NumAffaire") & "');""><img src='./Images/zoom.gif' border='0' /></a></td>" & vbCrLf
        
        RecapVehicule = RecapVehicule & "             <td><a href =""javascript:PageSuvante('10','" & Request("id") & "','" & _
                          "" & Rsvehicule("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & Rsvehicule("NumEncelade") & "&nbsp;</a> </td>" & vbCrLf
                         
         RecapVehicule = RecapVehicule & "             <td><a href =""javascript:PageSuvante('10','" & Request("id") & "','" & _
                          "" & Rsvehicule("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & Rsvehicule("Projet") & "&nbsp;</a> </td>" & vbCrLf
                         
        RecapVehicule = RecapVehicule & "             <td><a href =""javascript:PageSuvante('10','" & Request("id") & "','" & _
                          "" & Rsvehicule("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & Rsvehicule("Programme") & "&nbsp;</a> </td>" & vbCrLf
                         
        RecapVehicule = RecapVehicule & "             <td><a href =""javascript:PageSuvante('10','" & Request("id") & "','" & _
                          "" & Rsvehicule("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & Replace("" & Rsvehicule("Désignation"), vbCrLf, "<br>") & "&nbsp;</a> </td>" & vbCrLf
                         
        RecapVehicule = RecapVehicule & "             <td><a href =""javascript:PageSuvante('10','" & Request("id") & "','" & _
                          "" & Rsvehicule("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Request("Id_Projects_NumAffaire") & "');"">" & Rsvehicule("Véhicule") & "&nbsp;</a> </td>" & vbCrLf
        RecapVehicule = RecapVehicule & "</tr>"
       Rsvehicule.MoveNext
        Wend
     
    
    
       
       
       
          RecapVehicule = RecapVehicule & "</table>" & vbCrLf
          
          
          RecapVehicule = RecapVehicule & "<input name='NbS_Enssemble__' type='hidden' value='" & L2 & "'/>"
          RecapVehicule = RecapVehicule & "</form>"
      
       
End Function
Function RecapProjet(FormName)


    RecapProjet = RecapProjet & vbCrLf & "<SCRIPT langage='JAVASCRIPT'>"
    RecapProjet = RecapProjet & vbCrLf & "function PageSuvante(Logistique_Page,id,id_Vehicule,Pièce,ID__SousEnsembles,Id_Projects_NumAffaire){"

    RecapProjet = RecapProjet & vbCrLf & "document." & FormName & ".Logistique_Page.value=Logistique_Page;"
    RecapProjet = RecapProjet & vbCrLf & "document." & FormName & ".id.value=id;"
    RecapProjet = RecapProjet & vbCrLf & "document." & FormName & ".id_Vehicule.value=id_Vehicule;"
     RecapProjet = RecapProjet & vbCrLf & "document." & FormName & ".Pièce.value=Pièce;"
    RecapProjet = RecapProjet & vbCrLf & "document." & FormName & ".ID__SousEnsembles.value=ID__SousEnsembles;"
RecapProjet = RecapProjet & vbCrLf & "document." & FormName & ".Id_Projects_NumAffaire.value=Id_Projects_NumAffaire;"
    RecapProjet = RecapProjet & vbCrLf & "document." & FormName & ".action='Default.asp';"
    RecapProjet = RecapProjet & vbCrLf & "document." & FormName & ".submit();"



    RecapProjet = RecapProjet & vbCrLf & "}"
    RecapProjet = RecapProjet & vbCrLf & "</SCRIPT>"
    RecapProjet = RecapProjet & "<form name='" & FormName & "' method='post'>" & vbCrLf
    
    RecapProjet = RecapProjet & "<input name='Updatable' type='hidden' value='OK'/>" & vbCrLf
    RecapProjet = RecapProjet & "<input name='FormName' type='hidden' value='" & FormName & "'/>" & vbCrLf

    RecapProjet = RecapProjet & "<input name='Logistique_Page' type='hidden' value='10'/>" & vbCrLf
      
      
     
   
    
    
    sqlListe = sqlListe & "AND T_Nomenclature.Id="
    
'    SELECT     t_Vehicule.Id_Projects_NumAffaire, T_Societe.Code, T_Societe.Sosiété, T_Projects.NumAffaire, T_Projects.Projet
'FROM         t_Vehicule INNER JOIN
'                      T_Projects ON t_Vehicule.Id_Projects_NumAffaire = T_Projects.Id INNER JOIN
'                      T_Societe ON T_Projects.Id_Societe_code = T_Societe.Id LEFT OUTER JOIN
'                      T_SousEnsembles ON t_Vehicule.Id = T_SousEnsembles.ID_Vehicule_Véhicule LEFT OUTER JOIN
'                      T_Nomenclature ON T_SousEnsembles.Id = T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble]
'GROUP BY T_Projects.Projet, t_Vehicule.Id_Projects_NumAffaire, T_Projects.NumAffaire, T_Societe.Sosiété, T_Societe.Code

        
        Sql = "SELECT T_SousEnsembles.ID_Vehicule_Véhicule, t_Vehicule.NumEncelade, T_Projects.Projet,  "
        Sql = Sql & "t_Vehicule.Programme, t_Vehicule.Désignation, t_Vehicule.Véhicule "
        
        Sql = "SELECT t_Vehicule.Id_Projects_NumAffaire, T_Societe.Code, T_Societe.Sosiété, T_Projects.NumAffaire,  "
        Sql = Sql & "T_Projects.Projet "
        Sql = Sql & "FROM t_Vehicule INNER JOIN "
        Sql = Sql & "T_Projects ON t_Vehicule.Id_Projects_NumAffaire = T_Projects.Id INNER JOIN "
        Sql = Sql & "T_Societe ON T_Projects.Id_Societe_code = T_Societe.Id LEFT OUTER JOIN "
        Sql = Sql & "T_SousEnsembles ON t_Vehicule.Id = T_SousEnsembles.ID_Vehicule_Véhicule LEFT OUTER JOIN "
        Sql = Sql & "T_Nomenclature ON T_SousEnsembles.Id = T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] "
        MyFilter = RetourneCloseWhere()
          If Request("nbSelect") > 0 Then
            Sql = Sql & MyCloseWhere()
             If Trim("" & MyFilter) <> "" Then
                Sql = Sql & " AND (" & MyFilter & ") "
             End If
          Else
            If Trim("" & MyFilter) <> "" Then
                Sql = Sql & " WHERE " & MyFilter
             End If
        End If
        If Trim("" & Request("RecapFiltreVehicule")) <> "" Then
                    If InStr(1, UCase(Sql), "WHERE") <> 0 Then
                    
                Sql = Sql & " AND t_Vehicule.Id=" & Request("RecapFiltreVehicule") & " "
                
            Else
                MyFilter = "?"
                 Sql = Sql & " WHERE t_Vehicule.Id=" & Request("RecapFiltreVehicule") & " "
            End If
      End If
            
        Sql = Sql & "GROUP BY T_Projects.Projet, t_Vehicule.Id_Projects_NumAffaire, T_Projects.NumAffaire, T_Societe.Sosiété, T_Societe.Code"


     
'Response.Write sql
'Response.End

 
       
      
      
        Set Rsvehicule = OpenRecordSet(Sql, con)
      
         RecapProjet = RecapProjet & "<table border='1' width='100%'>" & vbCrLf
         If Request("nbSelect") > 0 Then
            RecapProjet = RecapProjet & "<tr>" & vbCrLf
            RecapProjet = RecapProjet & "         <td colspan='6' >" & RecapSersh & "</td>" & vbCrLf
            
            RecapProjet = RecapProjet & "</tr>" & vbCrLf
         End If
          RecapProjet = RecapProjet & "<tr >" & vbCrLf
         RecapProjet = RecapProjet & "<td class='TitreTableau' colspan ='8'>Liste des Projets :</td>" & vbCrLf
         RecapProjet = RecapProjet & "</tr>" & vbCrLf
         RecapProjet = RecapProjet & "<tr><td colspan ='8'><table><tr>" & vbCrLf
          
'        RecapProjet = RecapProjet & "<TD width='2' valign='top' colspan ='8'>" & AfficheBounton("Retour", 8, "'javascript:valider3();'") & "</td></tr>" & vbCrLf
          
         RecapProjet = RecapProjet & "<tr>" & vbCrLf
         
         
    RecapProjet = RecapProjet & vbCrLf & "<SCRIPT langage='JAVASCRIPT'>"
     
    RecapProjet = RecapProjet & "function ModifierRechereche(){" & vbCrLf
    RecapProjet = RecapProjet & "this.document." & FormName & ".ID__SousEnsembles.value='';"
    RecapProjet = RecapProjet & "this.document." & FormName & ".action='Default.asp?Logistique_Page=11'" & vbCrLf
    RecapProjet = RecapProjet & vbCrLf & "document." & FormName & ".submit();"
    RecapProjet = RecapProjet & "}" & vbCrLf
    
    
        RecapProjet = RecapProjet & "function valider3(){" & vbCrLf
        RecapProjet = RecapProjet & "this.document." & FormName & ".Id_Projects_NumAffaire.value='';"
        RecapProjet = RecapProjet & "this.document." & FormName & ".action='Default.asp?Logistique_Page=22'" & vbCrLf
        RecapProjet = RecapProjet & vbCrLf & "document." & FormName & ".submit();"
        RecapProjet = RecapProjet & "}" & vbCrLf

        RecapProjet = RecapProjet & "function valider2(){" & vbCrLf
        RecapProjet = RecapProjet & "this.document." & FormName & ".action='Default.asp?Logistique_Page=9'" & vbCrLf
        RecapProjet = RecapProjet & vbCrLf & "document." & FormName & ".submit();"
        RecapProjet = RecapProjet & "}" & vbCrLf
        RecapProjet = RecapProjet & vbCrLf & "</SCRIPT >"

      
      RecapProjet = RecapProjet & "<TD width='2%'> &nbsp;</td><TD>&nbsp;</td>" & vbCrLf
        RecapProjet = RecapProjet & "<TD width='4%' valign='top'>Filtrés sur véhicule :" & vbCrLf
        Sql = "SELECT t_Vehicule.Id, [Projet] , [Véhicule]  FROM t_Vehicule ORDER BY t_Vehicule.Véhicule;"
        
         Sql = "SELECT     t_Vehicule.Id, T_Projects.NumAffaire, t_Vehicule.Véhicule "
        Sql = Sql & "FROM t_Vehicule INNER JOIN "
        Sql = Sql & "T_Projects ON t_Vehicule.Id_Projects_NumAffaire = T_Projects.Id "
        Sql = Sql & "ORDER BY t_Vehicule.Véhicule"
        Set RsFiltre = OpenRecordSet(Sql, con)

       RecapProjet = RecapProjet & "<br>" & CreatListe(RsFiltre, Request("RecapFiltreVehicule"), "Id", "1;2", "RecapFiltreVehicule", " onChange='valider2();' ") & "</td>" & vbCrLf
      RsFiltre.Close
      Set RsFiltre = Nothing
      
     
     RecapProjet = RecapProjet & "</Tr></table></td>"
        RecapProjet = RecapProjet & "<tr>" & vbCrLf
        RecapProjet = RecapProjet & "         <td align='center' bgcolor='#CCCCCC'>&nbsp;</td>" & vbCrLf
        RecapProjet = RecapProjet & "         <td align='center' bgcolor='#CCCCCC'>Affaire</td>" & vbCrLf
        RecapProjet = RecapProjet & "         <td   align='center' bgcolor='#CCCCCC'>Code Société</td>" & vbCrLf
        RecapProjet = RecapProjet & "         <td   align='center' bgcolor='#CCCCCC'>Société</td>" & vbCrLf
        RecapProjet = RecapProjet & "         <td   align='center' bgcolor='#CCCCCC'>Projet</td>" & vbCrLf
'        RecapProjet = RecapProjet & "         <td   align='center' bgcolor='#CCCCCC'>Véhicule</td>" & vbCrLf
       
        RecapProjet = RecapProjet & "</tr> "
        RecapProjet = RecapProjet & "<input name='id' type='hidden' value='" & Request("id") & "' />"
       RecapProjet = RecapProjet & "<input name='id_Vehicule' type='hidden' value='" & Request("id_Vehicule") & "' />"
       RecapProjet = RecapProjet & "<input name='Pièce' type='hidden' value='" & Request("Pièce") & "' />"
       RecapProjet = RecapProjet & "<input name='ID__SousEnsembles' type='hidden' value='" & Request("ID__SousEnsembles") & "' />"
    RecapProjet = RecapProjet & "<input name='Id_Projects_NumAffaire' type='hidden' value='" & Request("Id_Projects_NumAffaire") & "' />"
        While Rsvehicule.EOF = False
        RecapProjet = RecapProjet & "<tr>"
           

        
        
       
         RecapProjet = RecapProjet & "             <td align='center'><a href =""javascript:PageSuvante('19','" & Request("id") & "','" & _
                        "0" & Request("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Rsvehicule("Id_Projects_NumAffaire") & "');""><img src='./Images/zoom.gif' border='0' /></a></td>" & vbCrLf
        
        RecapProjet = RecapProjet & "             <td><a href =""javascript:PageSuvante('9','" & Request("id") & "','" & _
                          "0" & Request("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Rsvehicule("Id_Projects_NumAffaire") & "');"">" & Rsvehicule("NumAffaire") & "&nbsp;</a> </td>" & vbCrLf
                         
         RecapProjet = RecapProjet & "             <td><a href =""javascript:PageSuvante('9','" & Request("id") & "','" & _
                          "0" & Request("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Rsvehicule("Id_Projects_NumAffaire") & "');"">" & Rsvehicule("Code") & "&nbsp;</a> </td>" & vbCrLf
                         
        RecapProjet = RecapProjet & "             <td><a href =""javascript:PageSuvante('9','" & Request("id") & "','" & _
                          "0" & Request("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Rsvehicule("Id_Projects_NumAffaire") & "');"">" & Rsvehicule("Sosiété") & "&nbsp;</a> </td>" & vbCrLf
                         
'        RecapProjet = RecapProjet & "             <td><a href =""javascript:PageSuvante('9','" & Request("id") & "','" & _
'                          "0" & Request("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
'                         Request("ID__SousEnsembles") & "');"">" & Replace("" & Rsvehicule("Désignation"), vbCrLf, "<br>") & "&nbsp;</a> </td>" & vbCrLf
                         
        RecapProjet = RecapProjet & "             <td><a href =""javascript:PageSuvante('9','" & Request("id") & "','" & _
                          "0" & Request("ID_Vehicule_Véhicule") & "','" & Request("Pièce") & "','" & _
                         Request("ID__SousEnsembles") & "','" & Rsvehicule("Id_Projects_NumAffaire") & "');"">" & Rsvehicule("Projet") & "&nbsp;</a> </td>" & vbCrLf
        RecapProjet = RecapProjet & "</tr>"
       Rsvehicule.MoveNext
        Wend
     
    
    
       
       
       
          RecapProjet = RecapProjet & "</table>" & vbCrLf
          
          
          RecapProjet = RecapProjet & "<input name='NbS_Enssemble__' type='hidden' value='" & L2 & "'/>"
          RecapProjet = RecapProjet & "</form>"
      
       
End Function
Function SelectNewProget(FormName)
SelectNewProget = SelectNewProget & "<SCRIPT langage='JAVASCRIPT'>" & vbCrLf
     SelectNewProget = SelectNewProget & "function mySubmit() {" & vbCrLf
            SelectNewProget = SelectNewProget & " myForm = this.document.forms[0];" & vbCrLf
             
            SelectNewProget = SelectNewProget & "myForm.action = 'Default.asp?Logistique_Page=20&fichier=" & Request("fichier") & "';" & vbCrLf
            
                     
            SelectNewProget = SelectNewProget & " myForm.submit();" & vbCrLf
            
          
              SelectNewProget = SelectNewProget & "}" & vbCrLf
     SelectNewProget = SelectNewProget & "</SCRIPT>" & vbCrLf
    SelectNewProget = SelectNewProget & "<form name='" & FormName & "' method='post'>" & vbCrLf
    
    SelectNewProget = SelectNewProget & "<input name='Updatable' type='hidden' value='OK'/>" & vbCrLf
    SelectNewProget = SelectNewProget & "<input name='FormName' type='hidden' value='" & FormName & "'/>" & vbCrLf

    SelectNewProget = SelectNewProget & "<input name='Logistique_Page' type='hidden' value='20'/>" & vbCrLf
      
      
     
   
    
    
    sqlListe = sqlListe & "AND T_Nomenclature.Id="

        
        
      
         SelectNewProget = SelectNewProget & "<table border='1' width='100%'>" & vbCrLf
         If Request("nbSelect") > 0 Then
            SelectNewProget = SelectNewProget & "<tr>" & vbCrLf
            SelectNewProget = SelectNewProget & "         <td colspan='6' >" & RecapSersh & "</td>" & vbCrLf
            
            SelectNewProget = SelectNewProget & "</tr>" & vbCrLf
         End If
          SelectNewProget = SelectNewProget & "<tr >" & vbCrLf
         SelectNewProget = SelectNewProget & "<td class='TitreTableau' colspan ='8'>Liste des Progets :</td>" & vbCrLf
         SelectNewProget = SelectNewProget & "</tr>" & vbCrLf
         SelectNewProget = SelectNewProget & "<tr><td>" & vbCrLf
         SelectNewProget = SelectNewProget & AfficheBountonSersh("Valider", "", """javascript:mySubmit();""")
        SelectNewProget = SelectNewProget & "</td></tr>" & vbCrLf
         SelectNewProget = SelectNewProget & "<tr>" & vbCrLf
          
        SelectNewProget = SelectNewProget & "<tr>" & vbCrLf
        SelectNewProget = SelectNewProget & "         <td align='center' bgcolor='#CCCCCC'>Sélectionner</td>" & vbCrLf
        SelectNewProget = SelectNewProget & "         <td align='center' bgcolor='#CCCCCC'>Projet</td>" & vbCrLf
          SelectNewProget = SelectNewProget & "         <td align='center' bgcolor='#CCCCCC'>Sociétés</td>" & vbCrLf
        SelectNewProget = SelectNewProget & "         <td align='center' bgcolor='#CCCCCC'>Affaire</td>" & vbCrLf
        
       
        SelectNewProget = SelectNewProget & "</tr> "
       
'    Set toto = CreateObject("excel.application")
'    toto.Visible = True
'    Set MyClass = toto.Workbooks.Add(SERVER.MapPath("./Excel") & "\" & Request("fichier"))
'    Sql = "if exists (select * from sysobjects where id = object_id(N'[" & Session("NumJob") & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf
'Sql = Sql & "drop table [" & Session("NumJob") & "]" & vbCrLf
''Sql = Sql & "GO" & vbCrLf
''con.Execute Sql
'
'Sql = Sql & "CREATE TABLE [" & Session("NumJob") & "] ("
'Sql = Sql & "    [NUMSENS] [char] (255) COLLATE French_CI_AS NULL ,"
'Sql = Sql & "    [PROJ] [char] (255) COLLATE French_CI_AS NULL ,"
'Sql = Sql & "    [AFF] [char] (255) COLLATE French_CI_AS NULL ,"
'Sql = Sql & "    [VAGUE] [char] (255) COLLATE French_CI_AS NULL ,"
'Sql = Sql & "    [NUMSE] [bigint] NULL ,"
'Sql = Sql & "    [DESGFE] [char] (255) COLLATE French_CI_AS NULL ,"
'Sql = Sql & "    [REFERENCE] [char] (255) COLLATE French_CI_AS NULL ,"
'Sql = Sql & "    [IND] [char] (255) COLLATE French_CI_AS NULL ,"
'Sql = Sql & "    [DESPIE] [char] (255) COLLATE French_CI_AS NULL ,"
'Sql = Sql & "    [QUANTITE] [bigint] NULL"
'Sql = Sql & ") ON [PRIMARY]" & vbCrLf
'con.Execute Sql
'Set RsExcel = MyClass.Worksheets("Feuil1").range("A1").currentregion
'For L = 2 To RsExcel.rows.Count
'   Sql = Sql & "INSERT INTO " & Session("NumJob") & " ( NUMSENS,PROJ, AFF, VAGUE, NUMSE, DESGFE, REFERENCE, IND, DESPIE, QUANTITE ) "
'Sql = Sql & "values(" & noquote(Trim("" & RsExcel("14")), "txt") & ","
'Sql = Sql & noquote(Trim("" & RsExcel("1")), "txt") & ","
'Sql = Sql & noquote(Trim(Replace(Replace("" & RsExcel(L, 3), Chr(10), ""), Chr(13), "")), "txt") & ","
'Sql = Sql & noquote(Trim(Replace(Replace("" & RsExcel(L, 5), Chr(10), ""), Chr(13), "")), "txt") & ","
'Sql = Sql & noquote(Trim(Replace(Replace("" & RsExcel(L, 6), Chr(10), ""), Chr(13), "")), "num") & ","
'Sql = Sql & noquote(Trim(Replace(Replace("" & RsExcel(L, 19), Chr(10), ""), Chr(13), "")), "txt") & ","
'Sql = Sql & noquote(Trim(Replace(Replace("" & RsExcel(L, 22), Chr(10), ""), Chr(13), "")), "txt") & ","
'Sql = Sql & noquote(Trim(Replace(Replace("" & RsExcel(L, 23), Chr(10), ""), Chr(13), "")), "txt") & ","
'Sql = Sql & noquote(Trim(Replace(Replace("" & RsExcel(L, 24), Chr(10), ""), Chr(13), "")), "txt") & ","
'Sql = Sql & noquote(Trim(Replace(Replace("" & RsExcel(L, 25), Chr(10), ""), Chr(13), "")), "num") & ")" & vbCrLf
'con.exe
'Next
'MyClass.Close False
'toto.quit
'Set MyClass = Nothing
'Set toto = Nothing
'Set RsExcel = Nothing
Set conexcel = OpenDbN(7, "", "", SERVER.MapPath("./Excel") & "\" & Request("fichier"), "", "")
Sql = "SELECT [Feuil1$].PROJ From [Feuil1$] GROUP BY [Feuil1$].PROJ;"
Set RsExcel = conexcel.Execute(Sql)
'Set ConNumerotation = OpenDbN(Session("ConexTypeNumerotation"), "", "", Session("BaseNumerotation"), "", "")
'Sql = "SELECT Activité.[Clé ac] AS ID, Activité.[Lib ac] "
'Sql = Sql & "From Activité "
'Sql = Sql & "Where Activité.[St ac] <> 4 "
'Sql = Sql & "GROUP BY Activité.[Clé ac], Activité.[Lib ac]  "
'Sql = Sql & "ORDER BY Activité.[Clé ac] DESC "
'
'
'Set RsAffair = ConNumerotation.Execute(Sql)
' T_RsAffair = RsAffair.GetRows
'  RsAffair.Close
'  Set RsAffair = Nothing
'   ConNumerotation.Close
'  Set ConNumerotation = Nothing
'   T_RsAffair = T_CreatListe(T_RsAffair, "", 0, "0;1", "T_RsAffair", "")
 Sql = "SELECT T_Societe.Id, T_Societe.Sosiété, T_Societe.Code From T_Societe ORDER BY T_Societe.Code"
 Set RsSociete = con.Execute(Sql)
 T_RsSociete = RsSociete.GetRows
  RsSociete.Close
  Set RsSociete = Nothing
  T_RsSociete = T_CreatListe(T_RsSociete, "", 0, "1;2", "T_RsSociete", "")
 
 
I = 0
While RsExcel.EOF = False
I = I + 1
'      SelectNewProget = SelectNewProget & "<tr><td width='2%' align='center'> <input name='Checkbox_" & I & "' type='checkbox' value='TRUE'/></td><td>"
'       SelectNewProget = SelectNewProget & "<input name='Projet_" & I & "' type='hidden' value='" & RsExcel!PROJ & "' />"
'    SelectNewProget = SelectNewProget & RsExcel!PROJ & "</td>" & vbCrLf
'    SelectNewProget = SelectNewProget & "<td>" & Replace(T_RsSociete, "name='T_RsSociete'", "name='T_RsSociete_" & I & "'") & "</td>" & vbCrLf
'    SelectNewProget = SelectNewProget & "<td>" & Replace(T_RsAffair, "name='T_RsAffair'", "name='T_RsAffair_" & I & "'") & "</td></tr>" & vbCrLf
 SelectNewProget = SelectNewProget & "<tr><td width='2%' align='center'> <input name='Checkbox_" & I & "' type='checkbox' value='TRUE'/></td><td>"
       SelectNewProget = SelectNewProget & "<input name='Projet_" & I & "' type='hidden' value='" & RsExcel!PROJ & "' />"
    SelectNewProget = SelectNewProget & RsExcel!PROJ & "</td>" & vbCrLf
    SelectNewProget = SelectNewProget & "<td width>" & Replace(T_RsSociete, "name='T_RsSociete'", "name='T_RsSociete_" & I & "'") & "</td>" & vbCrLf
    SelectNewProget = SelectNewProget & "<td><input name='T_RsAffair_" & I & "' type='text' value='' /></td></tr>" & vbCrLf
    
RsExcel.MoveNext
 Wend
   RsExcel.Close
   Set RsExcel = Nothing
   conexcel.Close
   Set conexcel = Nothing
       SelectNewProget = SelectNewProget & "<input Name='NbProget' type='hidden' value='" & I & "'/>"
          SelectNewProget = SelectNewProget & "</table>" & vbCrLf
          
          
          SelectNewProget = SelectNewProget & "<input name='NbS_Enssemble__' type='hidden' value='" & L2 & "'/>"
          SelectNewProget = SelectNewProget & "</form>"
      
       
End Function
Function SelectNewProgetVehicule(FormName)
SelectNewProgetVehicule = SelectNewProgetVehicule & "<SCRIPT langage='JAVASCRIPT'>" & vbCrLf
     SelectNewProgetVehicule = SelectNewProgetVehicule & "function mySubmit() {" & vbCrLf
            SelectNewProgetVehicule = SelectNewProgetVehicule & " myForm = this.document.forms[0];" & vbCrLf
             
            SelectNewProgetVehicule = SelectNewProgetVehicule & "myForm.action = 'Default.asp?Logistique_Page=21&fichier=" & Request("fichier") & "';" & vbCrLf
            
                     
            SelectNewProgetVehicule = SelectNewProgetVehicule & " myForm.submit();" & vbCrLf
            
          
              SelectNewProgetVehicule = SelectNewProgetVehicule & "}" & vbCrLf
     SelectNewProgetVehicule = SelectNewProgetVehicule & "</SCRIPT>" & vbCrLf
    SelectNewProgetVehicule = SelectNewProgetVehicule & "<form name='" & FormName & "' method='post'>" & vbCrLf
    
    SelectNewProgetVehicule = SelectNewProgetVehicule & "<input name='Updatable' type='hidden' value='OK'/>" & vbCrLf
    SelectNewProgetVehicule = SelectNewProgetVehicule & "<input name='FormName' type='hidden' value='" & FormName & "'/>" & vbCrLf

    SelectNewProgetVehicule = SelectNewProgetVehicule & "<input name='Logistique_Page' type='hidden' value='10'/>" & vbCrLf
      
      
     
   
    
    
    sqlListe = sqlListe & "AND T_Nomenclature.Id="

        
        
      
         SelectNewProgetVehicule = SelectNewProgetVehicule & "<table border='1' width='100%'>" & vbCrLf
         If Request("nbSelect") > 0 Then
            SelectNewProgetVehicule = SelectNewProgetVehicule & "<tr>" & vbCrLf
            SelectNewProgetVehicule = SelectNewProgetVehicule & "         <td colspan='6' >" & RecapSersh & "</td>" & vbCrLf
            
            SelectNewProgetVehicule = SelectNewProgetVehicule & "</tr>" & vbCrLf
         End If
          SelectNewProgetVehicule = SelectNewProgetVehicule & "<tr >" & vbCrLf
         SelectNewProgetVehicule = SelectNewProgetVehicule & "<td class='TitreTableau' colspan ='8'>Liste des Véhicules :</td>" & vbCrLf
         SelectNewProgetVehicule = SelectNewProgetVehicule & "</tr>" & vbCrLf
         SelectNewProgetVehicule = SelectNewProgetVehicule & "<tr><td>" & vbCrLf
         SelectNewProgetVehicule = SelectNewProgetVehicule & AfficheBountonSersh("Valider", "", """javascript:mySubmit();""")
        SelectNewProgetVehicule = SelectNewProgetVehicule & "</td></tr>" & vbCrLf
         SelectNewProgetVehicule = SelectNewProgetVehicule & "<tr>" & vbCrLf
          
        SelectNewProgetVehicule = SelectNewProgetVehicule & "<tr>" & vbCrLf
        SelectNewProgetVehicule = SelectNewProgetVehicule & "         <td align='center' bgcolor='#CCCCCC'>Sélectionner</td>" & vbCrLf
         SelectNewProgetVehicule = SelectNewProgetVehicule & "         <td align='center' bgcolor='#CCCCCC'>N° Encelade</td>" & vbCrLf
        SelectNewProgetVehicule = SelectNewProgetVehicule & "         <td align='center' bgcolor='#CCCCCC'>Projet</td>" & vbCrLf
        SelectNewProgetVehicule = SelectNewProgetVehicule & "         <td align='center' bgcolor='#CCCCCC'>Programme</td>" & vbCrLf
              SelectNewProgetVehicule = SelectNewProgetVehicule & "         <td align='center' bgcolor='#CCCCCC'>Véhicule</td>" & vbCrLf
        SelectNewProgetVehicule = SelectNewProgetVehicule & "</tr> "
       Sql = ""
    For I = 1 To Request("NbProget")
    
   
        If Request("Checkbox_" & I) = "TRUE" Then
            Sql = "INSERT T_Projects  ( Id_Societe_code, NumAffaire,Projet,NumJob) "
            Sql = Sql & "VALUES (" & noquote(Trim(Replace(Replace("" & Request("T_RsSociete_" & I), Chr(10), ""), Chr(13), "")), "num") & ","
            Sql = Sql & noquote(Trim(Replace(Replace("" & Request("T_RsAffair_" & I), Chr(10), ""), Chr(13), "")), "num") & ","
            Sql = Sql & noquote(Trim(Replace(Replace("" & Request("Projet_" & I), Chr(10), ""), Chr(13), "")), "txt") & "," & noquote(Session("NumJob"), "txt")
           
            Sql = Sql & ")" & vbCrLf
            con.Execute Sql
            Sql = ""
            WhereExcel = WhereExcel & " [Feuil1$].PROJ=" & noquote(Request("Projet_" & I), "txt") & " OR "
        End If
    Next
    On Error Resume Next
     
       WhereExcel = Trim("" & WhereExcel)
       WhereExcel = Left(WhereExcel, Len(WhereExcel) - 2)




       
Set conexcel = OpenDbN(7, "", "", SERVER.MapPath("./Excel") & "\" & Request("fichier"), "", "")
Sql = "SELECT  [Feuil1$].[NUMSENS],[Feuil1$].PROJ,[Feuil1$].AFF  ,[Feuil1$].DESPIE ,[Feuil1$].VAGUE,[Feuil1$].NUMSE,[Feuil1$].DESGFE,[Feuil1$].REFERENCE,[Feuil1$].IND,[Feuil1$].DESPIE,[Feuil1$].QUANTITE "
Sql = Sql & "FROM [Feuil1$] "
Sql = Sql & "Where " & WhereExcel & " and [Feuil1$].[NUMSENS] Is Not Null "
Sql = Sql & "GROUP BY [Feuil1$].[NUMSENS],[Feuil1$].PROJ,[Feuil1$].AFF  ,[Feuil1$].DESPIE ,[Feuil1$].VAGUE,[Feuil1$].NUMSE,[Feuil1$].DESGFE,[Feuil1$].REFERENCE,[Feuil1$].IND,[Feuil1$].DESPIE,[Feuil1$].QUANTITE;"
Set RsExcel = conexcel.Execute(Sql)
Sql = "if exists (select * from sysobjects where id = object_id(N'[" & Session("NumJob") & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)" & vbCrLf
Sql = Sql & "drop table [" & Session("NumJob") & "]" & vbCrLf
'Sql = Sql & "GO" & vbCrLf
'con.Execute Sql

Sql = Sql & "CREATE TABLE [" & Session("NumJob") & "] ("
Sql = Sql & "    [NUMSENS] [char] (255) COLLATE French_CI_AS NULL ,"
Sql = Sql & "    [PROJ] [char] (255) COLLATE French_CI_AS NULL ,"
Sql = Sql & "    [AFF] [char] (255) COLLATE French_CI_AS NULL ,"
Sql = Sql & "    [VAGUE] [char] (255) COLLATE French_CI_AS NULL ,"
Sql = Sql & "    [NUMSE] [bigint] NULL ,"
Sql = Sql & "    [DESGFE] [char] (255) COLLATE French_CI_AS NULL ,"
Sql = Sql & "    [REFERENCE] [char] (255) COLLATE French_CI_AS NULL ,"
Sql = Sql & "    [IND] [char] (255) COLLATE French_CI_AS NULL ,"
Sql = Sql & "    [DESPIE] [char] (255) COLLATE French_CI_AS NULL ,"
Sql = Sql & "    [QUANTITE] [bigint] NULL"
Sql = Sql & ") ON [PRIMARY]" & vbCrLf
'Sql = Sql & "GO" & vbCrLf
'con.Execute Sql

While RsExcel.EOF = False
Sql = Sql & "INSERT INTO " & Session("NumJob") & " ( NUMSENS,PROJ, AFF, VAGUE, NUMSE, DESGFE, REFERENCE, IND, DESPIE, QUANTITE ) "
Sql = Sql & "values(" & noquote(Trim("" & RsExcel("NUMSENS")), "txt") & ","
Sql = Sql & noquote(Trim("" & RsExcel("PROJ")), "txt") & ","
Sql = Sql & noquote(Trim(Replace(Replace("" & RsExcel("AFF"), Chr(10), ""), Chr(13), "")), "txt") & ","
Sql = Sql & noquote(Trim(Replace(Replace("" & RsExcel("VAGUE"), Chr(10), ""), Chr(13), "")), "txt") & ","
Sql = Sql & noquote(Trim(Replace(Replace("" & RsExcel("NUMSE"), Chr(10), ""), Chr(13), "")), "num") & ","
Sql = Sql & noquote(Trim(Replace(Replace("" & RsExcel("DESGFE"), Chr(10), ""), Chr(13), "")), "txt") & ","
Sql = Sql & noquote(Trim(Replace(Replace("" & RsExcel("REFERENCE"), Chr(10), ""), Chr(13), "")), "txt") & ","
Sql = Sql & noquote(Trim(Replace(Replace("" & RsExcel("IND"), Chr(10), ""), Chr(13), "")), "txt") & ","
Sql = Sql & noquote(Trim(Replace(Replace("" & RsExcel("DESPIE"), Chr(10), ""), Chr(13), "")), "txt") & ","
Sql = Sql & noquote(Trim(Replace(Replace("" & RsExcel("QUANTITE"), Chr(10), ""), Chr(13), "")), "num") & ")" & vbCrLf
'con.Execute Sql
RsExcel.MoveNext
Wend

con.Execute Sql


Sql = "SELECT     PROJ, AFF as Prorgramme, VAGUE, NUMSE "
Sql = Sql & "From " & Session("NumJob") & " "
Sql = Sql & "GROUP BY PROJ, AFF, VAGUE, NUMSE"
Set RsExcel = con.Execute(Sql)

conexcel.Close
Set conexcel = Nothing
'Set ConNumerotation = OpenDbN(Session("ConexTypeNumerotation"), "", "", Session("BaseNumerotation"), "", "")
'Sql = "SELECT Chrono" & Format(Date, "yyyy") & ".[Clé Ch], "
'Sql = Sql & "Chrono" & Format(Date, "yyyy") & ".[Clé ac],  "
'Sql = Sql & "Chrono" & Format(Date, "yyyy") & ".Année,  "
'Sql = Sql & "Chrono" & Format(Date, "yyyy") & ".Rév,  "
'Sql = Sql & "Chrono" & Format(Date, "yyyy") & ".[Clé ty] "
'Sql = Sql & "From Chrono" & Format(Date, "yyyy") & " "
'Sql = Sql & "Where Chrono" & Format(Date, "yyyy") & ".Année = '" & Format(Date, "yy") & "'  "
'Sql = Sql & "And Chrono" & Format(Date, "yyyy") & ".[Clé ty] = 'NM' "
'Sql = Sql & "ORDER BY Chrono" & Format(Date, "yyyy") & ".[Clé Ch] DESC;"
'
'
'Set RsAffair = ConNumerotation.Execute(Sql)
' T_RsAffair = CreatListeChrono(RsAffair, "")
'  RsAffair.Close
'  Set RsAffair = Nothing
'   ConNumerotation.Close
'  Set ConNumerotation = Nothing
 
 
 
I = 0
While RsExcel.EOF = False
I = I + 1

      SelectNewProgetVehicule = SelectNewProgetVehicule & "<tr><td width='2%' align='center'> <input name='Checkbox_" & I & "' type='checkbox' value='TRUE'/></td>"
'     SelectNewProgetVehicule = SelectNewProgetVehicule & "<td>" & Replace(T_RsAffair, "[NameListe]", "NumEns_" & I) & "</td>" & vbCrLf
      SelectNewProgetVehicule = SelectNewProgetVehicule & "<td><input name='NumEns_" & I & "' type='text' value=''/></td>" & vbCrLf
    SelectNewProgetVehicule = SelectNewProgetVehicule & "<td><input name='Projet_" & I & "' type='hidden' value='" & Trim("" & RsExcel!PROJ) & "'/>" & RsExcel!PROJ & "</td>" & vbCrLf
    SelectNewProgetVehicule = SelectNewProgetVehicule & "<td><input name='Programme_" & I & "' type='hidden' value='" & Trim("" & RsExcel!Prorgramme) & "'/>" & RsExcel!Prorgramme & "</td>" & vbCrLf
    
    
    
    SelectNewProgetVehicule = SelectNewProgetVehicule & " <td><input name='VAGUE_" & I & "' type='hidden' value='" & Trim("" & RsExcel!VAGUE) & "'/>"
    SelectNewProgetVehicule = SelectNewProgetVehicule & " <input name='NUMSE_" & I & "' type='hidden' value='" & Trim("" & RsExcel!NUMSE) & "'/>"
     SelectNewProgetVehicule = SelectNewProgetVehicule & Trim("" & RsExcel!VAGUE) & Trim("" & RsExcel!NUMSE) & "</td></tr>" & vbCrLf
   ' SelectNewProgetVehicule = SelectNewProgetVehicule & "<td>" & Replace(T_RsAffair, "name='T_RsAffair'", "name='T_RsAffair_" & I & "'") & "</td></tr>" & vbCrLf
     
RsExcel.MoveNext
 Wend
   RsExcel.Close
   Set RsExcel = Nothing
   
       SelectNewProgetVehicule = SelectNewProgetVehicule & "<input Name='NbProget' type='hidden' value='" & I & "'/>"
          SelectNewProgetVehicule = SelectNewProgetVehicule & "</table>" & vbCrLf
          
          
          SelectNewProgetVehicule = SelectNewProgetVehicule & "<input name='NbS_Enssemble__' type='hidden' value='" & L2 & "'/>"
          SelectNewProgetVehicule = SelectNewProgetVehicule & "</form>"
      
       
End Function
Function SelectNewProgetMAJ(FormName)


        
      Sql = ""
       
    For I = 1 To Request("NbProget")
    
   
        If Request("Checkbox_" & I) = "TRUE" Then
           
          
            Debug.Print
            Debug.Print noquote(Trim("" & Request("Programme_" & I)), "txt")
             Debug.Print noquote(Trim("" & Request("NumEns_" & I)), "txt")
            
             Debug.Print Trim("" & Request("VAGUE_" & I))
              Debug.Print Trim("" & Request("NUMSE_" & I))
               
      
            

MyWher = "WHERE "
MyWher = MyWher & "" & Session("NumJob") & ".NUMSE = " & noquote(Trim("" & Request("NUMSE_" & I)), "num") & "  "
MyWher = MyWher & "AND T_Projects.Projet = " & noquote(Trim("" & Request("Projet_" & I)), "txt") & "  "
MyWher = MyWher & "AND " & Session("NumJob") & ".AFF = " & noquote(Trim("" & Request("Programme_" & I)), "txt") & "  "
MyWher = MyWher & "AND " & Session("NumJob") & ".VAGUE = " & noquote(Trim("" & Request("VAGUE_" & I)), "txt") & "  "
MyWher = MyWher & "AND T_Projects.NumJob = '" & Session("NumJob") & "' "


'            véhicule
Sql1 = ""
Sql1 = Sql1 & "INSERT INTO t_Vehicule "
Sql1 = Sql1 & "(Id_Projects_NumAffaire,  NumEncelade, Programme, Véhicule) "
Sql1 = Sql1 & "SELECT Id AS Expr1, Expr1 AS Expr2, AFF AS Expr3, Expr2 AS Expr4 "
Sql1 = Sql1 & "FROM ("
Sql1 = Sql1 & "SELECT "
Sql1 = Sql1 & "T_Projects.Id, "
Sql1 = Sql1 & noquote(Trim("" & Request("NumEns_" & I)), "txt") & " AS Expr1, "
Sql1 = Sql1 & "" & Session("NumJob") & ".AFF, "
Sql1 = Sql1 & "REPLACE(" & Session("NumJob") & ".VAGUE, ' ', '') + REPLACE(STR(" & Session("NumJob") & ".NUMSE), ' ', '') AS Expr2 "
Sql1 = Sql1 & "FROM         " & Session("NumJob") & " INNER JOIN "
Sql1 = Sql1 & "T_Projects ON " & Session("NumJob") & ".PROJ = T_Projects.Projet "
Sql1 = Sql1 & MyWher & " "
Sql1 = Sql1 & "GROUP BY T_Projects.Projet, " & Session("NumJob") & ".AFF, " & Session("NumJob") & ".VAGUE, " & Session("NumJob") & ".NUMSE, REPLACE(" & Session("NumJob") & ".VAGUE,' ', '') + REPLACE(STR(" & Session("NumJob") & ".NUMSE), ' ', ''), T_Projects.NumJob, " & Session("NumJob") & ".DESGFE, T_Projects.NumJob, T_Projects.Id) as MyForm "
Sql1 = Sql1 & "GROUP BY Id, Expr1, AFF, Expr2"





'
'SELECT     Id AS Expr1, Expr1 AS Expr2, AFF AS Expr3, Expr2 AS Expr4
'FROM         (SELECT     T_Projects.Id, 'NM_1988_07_2051' AS Expr1, NumJob_2007_117_1.AFF, REPLACE(NumJob_2007_117_1.VAGUE, ' ', '')
'                                              + REPLACE(STR(NumJob_2007_117_1.NUMSE), ' ', '') AS Expr2
'                       FROM          NumJob_2007_117_1 INNER JOIN
'                                              T_Projects ON NumJob_2007_117_1.PROJ = T_Projects.Projet
'                       WHERE      (NumJob_2007_117_1.NUMSE = 1) AND (T_Projects.Projet = 'X44') AND (NumJob_2007_117_1.AFF = '1000') AND
'                                              (NumJob_2007_117_1.VAGUE = 'MP') AND (T_Projects.NumJob = 'NumJob_2007_117_1')
'                       GROUP BY T_Projects.Projet, NumJob_2007_117_1.AFF, NumJob_2007_117_1.VAGUE, NumJob_2007_117_1.NUMSE,
'                                              REPLACE(NumJob_2007_117_1.VAGUE, ' ', '') + REPLACE(STR(NumJob_2007_117_1.NUMSE), ' ', ''), T_Projects.NumJob,
'                                              NumJob_2007_117_1.DESGFE, T_Projects.NumJob, T_Projects.Id) MyForm
'GROUP BY Id, Expr1, AFF, Expr2


'sous enssemble:
Sql2 = ""
Sql2 = Sql2 & "INSERT INTO T_SousEnsembles "
Sql2 = Sql2 & "(ID_Vehicule_Véhicule, Code) "
Sql2 = Sql2 & "SELECT Id AS Expr1, NUMSENS AS Expr2 FROM ("
Sql2 = Sql2 & "SELECT t_Vehicule.Id, " & Session("NumJob") & ".NUMSENS "
Sql2 = Sql2 & "FROM " & Session("NumJob") & " INNER JOIN "
Sql2 = Sql2 & "T_Projects ON " & Session("NumJob") & ".PROJ = T_Projects.Projet INNER JOIN "
Sql2 = Sql2 & "t_Vehicule ON T_Projects.Id = t_Vehicule.Id_Projects_NumAffaire LEFT OUTER JOIN "
Sql2 = Sql2 & "T_SousEnsembles ON t_Vehicule.Id = T_SousEnsembles.ID_Vehicule_Véhicule "
Sql2 = Sql2 & MyWher & "and T_SousEnsembles.ID_Vehicule_Véhicule IS NULL "
Sql2 = Sql2 & "GROUP BY T_Projects.NumJob, " & Session("NumJob") & ".DESGFE, T_Projects.NumJob,  "
Sql2 = Sql2 & "t_Vehicule.Id, " & Session("NumJob") & ".NUMSENS) as MyFrom "
Sql2 = Sql2 & "GROUP BY Id, NUMSENS"





'Pièce:

 

Sql3 = ""
Sql3 = Sql3 & "INSERT INTO T_Nomenclature "
Sql3 = Sql3 & "([ID_SousEnsembles_Sous-Ensseble], [Désination FE], Pièce, Indice, Designation_Piece, Quantite,Reste_a_recevoir) "
Sql3 = Sql3 & "SELECT     T_SousEnsembles.Id, " & Session("NumJob") & ".DESGFE, " & Session("NumJob") & ".REFERENCE, " & Session("NumJob") & ".IND, "
 Sql3 = Sql3 & "" & Session("NumJob") & ".DESPIE , " & Session("NumJob") & ".QUANTITE, "
 Sql3 = Sql3 & Session("NumJob") & ".QUANTITE "
Sql3 = Sql3 & "FROM " & Session("NumJob") & " INNER JOIN "
Sql3 = Sql3 & "T_Projects ON " & Session("NumJob") & ".PROJ = T_Projects.Projet INNER JOIN "
Sql3 = Sql3 & "t_Vehicule ON T_Projects.Id = t_Vehicule.Id_Projects_NumAffaire INNER JOIN "
Sql3 = Sql3 & "T_SousEnsembles ON t_Vehicule.Id = T_SousEnsembles.ID_Vehicule_Véhicule AND "
Sql3 = Sql3 & "" & Session("NumJob") & ".NUMSENS = T_SousEnsembles.Code LEFT OUTER JOIN "
Sql3 = Sql3 & "T_Nomenclature ON T_SousEnsembles.Id = T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] "
Sql3 = Sql3 & MyWher & " and T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] IS NULL "
Sql3 = Sql3 & "GROUP BY T_SousEnsembles.Id, " & Session("NumJob") & ".DESGFE, " & Session("NumJob") & ".REFERENCE, " & Session("NumJob") & ".IND,"
Sql3 = Sql3 & " " & Session("NumJob") & ".DESPIE , " & Session("NumJob") & ".QUANTITE"


'SELECT     T_SousEnsembles.Id,  Session("NumJob") & ".DESGFE, Session("NumJob") & ".REFERENCE,
'                      Session("NumJob") & ".IND , Session("NumJob") & ".DESPIE, Session("NumJob") & ".QUANTITE
'From Session("NumJob") & " INNER JOIN"
'                      T_Projects ON Session("NumJob") & ".PROJ = T_Projects.Projet INNER JOIN
'                      t_Vehicule ON T_Projects.Id = t_Vehicule.Id_Projects_NumAffaire INNER JOIN
'                      T_SousEnsembles ON t_Vehicule.Id = T_SousEnsembles.ID_Vehicule_Véhicule AND
'                      Session ("NumJob") & ".NUMSENS = T_SousEnsembles.Code"
'Sql3 = Sql3 & MyWhere & " "
'GROUP BY T_Projects.NumJob, Session("NumJob") & ".DESGFE, T_Projects.NumJob, T_SousEnsembles.Id,
'                      Session("NumJob") & ".REFERENCE, Session("NumJob") & ".IND, Session("NumJob") & ".DESPIE, Session("NumJob") & ".QUANTITE,
'                      Session("NumJob") & ".DESGFE , Session("NumJob") & ".NUMSENS


con.Execute Sql1
con.Execute Sql2
con.Execute Sql3
Sql1 = ""
Sql2 = ""
Sql3 = ""
            
         
        End If
    Next
Sql = Sql & "UPDATE    T_Projects "
Sql = Sql & "SET              NumJob = NULL "
Sql = Sql & "WHERE NumJob='" & Session("NumJob") & "'" & vbCrLf
con.Execute Sql




 Response.Redirect "Default.asp"
       
End Function
Function RecapNewProget(FormName)
    
    RecapNewProget = RecapNewProget & "<form name='" & FormName & "' method='post'>" & vbCrLf
    
    RecapNewProget = RecapNewProget & "<input name='Updatable' type='hidden' value='OK'/>" & vbCrLf
    RecapNewProget = RecapNewProget & "<input name='FormName' type='hidden' value='" & FormName & "'/>" & vbCrLf

    RecapNewProget = RecapNewProget & "<input name='Logistique_Page' type='hidden' value='10'/>" & vbCrLf
      
      
     
   
    
    
    sqlListe = sqlListe & "AND T_Nomenclature.Id="

        
        
      
         RecapNewProget = RecapNewProget & "<table border='1' width='100%'>" & vbCrLf
         If Request("nbSelect") > 0 Then
            RecapNewProget = RecapNewProget & "<tr>" & vbCrLf
            RecapNewProget = RecapNewProget & "         <td colspan='6' >" & RecapSersh & "</td>" & vbCrLf
            
            RecapNewProget = RecapNewProget & "</tr>" & vbCrLf
         End If
          RecapNewProget = RecapNewProget & "<tr >" & vbCrLf
         RecapNewProget = RecapNewProget & "<td class='TitreTableau' colspan ='8'>Liste des Fichier Excel :</td>" & vbCrLf
         RecapNewProget = RecapNewProget & "</tr>" & vbCrLf
        
          
         RecapNewProget = RecapNewProget & "<tr>" & vbCrLf
          
        RecapNewProget = RecapNewProget & "<tr>" & vbCrLf
        RecapNewProget = RecapNewProget & "         <td align='center' bgcolor='#CCCCCC'>Supprimer</td>" & vbCrLf
        RecapNewProget = RecapNewProget & "         <td align='center' bgcolor='#CCCCCC'>Fichiers</td>" & vbCrLf
        
       
        RecapNewProget = RecapNewProget & "</tr> "
       
    
       Set FSO = SERVER.CreateObject("Scripting.FileSystemObject")

Set leRepertoire = FSO.GetFolder(SERVER.MapPath("./Excel") & "\")
I = 0
For Each Fichier In leRepertoire.Files
I = I + 1
      RecapNewProget = RecapNewProget & "<tr><td> <input name='Checkbox_" & I & "' type='checkbox' value='TRUE'/></td><td>"
    RecapNewProget = RecapNewProget & "<a href='Default.asp?Logistique_Page=17&fichier=" & Fichier.Name & "'>" & Fichier.Name & "</a></td></tr>" & vbCrLf
     
Next
       
       
       
          RecapNewProget = RecapNewProget & "</table>" & vbCrLf
          
          
          RecapNewProget = RecapNewProget & "<input name='NbS_Enssemble__' type='hidden' value='" & L2 & "'/>"
          RecapNewProget = RecapNewProget & "</form>"
      
       
End Function


 Function modSEARCH()
 
modSEARCH = ""
Dim Rs
Dim Sql
Dim IndexTab
Dim long_Ubound
Dim TableauSherch()
 Dim SelecetList
 Dim boolSelecetList
IndexTab = 0
boolSelecetList = False
' On recupere le nombre de lignes
If Request("nbSelect") > 0 Then
 nbLignes = CInt(Request("nbSelect")) + 1
Else
 nbLignes = 0
End If
Const IndexOperateurDefault = 2
Dim IndexChamps
Dim aSlt_OPERATEUR(1, 8)
aSlt_OPERATEUR(0, 0) = "<"
aSlt_OPERATEUR(0, 1) = "<="
aSlt_OPERATEUR(0, 2) = "="
aSlt_OPERATEUR(0, 3) = "=>"
aSlt_OPERATEUR(0, 4) = ">"
aSlt_OPERATEUR(0, 5) = "LIKE1"
aSlt_OPERATEUR(0, 6) = "LIKE2"
aSlt_OPERATEUR(0, 7) = "LIKE"
aSlt_OPERATEUR(0, 8) = "<>"

aSlt_OPERATEUR(1, 0) = "Inférieur à"
aSlt_OPERATEUR(1, 1) = "Inférieur ou Egal à"
aSlt_OPERATEUR(1, 2) = "Egal à"
aSlt_OPERATEUR(1, 3) = "Supérieur ou Egal à"
aSlt_OPERATEUR(1, 4) = "Supérieur à"
aSlt_OPERATEUR(1, 5) = " Commençant par"
aSlt_OPERATEUR(1, 6) = "Finissant par"
aSlt_OPERATEUR(1, 7) = "Contient"
aSlt_OPERATEUR(1, 8) = "Différent"

Dim aSlt_Logique(1, 2)
aSlt_Logique(0, 0) = "(Aucun)"
aSlt_Logique(0, 1) = "AND"
aSlt_Logique(0, 2) = "OR"

aSlt_Logique(1, 0) = "(Aucun)"
aSlt_Logique(1, 1) = "ET"
aSlt_Logique(1, 2) = "OU"








Dim aSlt2(1, 1)
aSlt2(0, 0) = 3
aSlt2(1, 0) = "Option 3"
aSlt2(0, 1) = 4
aSlt2(1, 1) = "Option 4"

Dim Champs(28, 1)
Dim I_Champ
I_Champ = 1
Champs(I_Champ, 0) = "t_Vehicule.Atelier"
I_Champ = I_Champ + 1
'Champs(I_Champ, 0) = "T_Nomenclature.Bon_sortie"
'
'I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "T_SousEnsembles.Code"
I_Champ = I_Champ + 1
'Champs(I_Champ, 0) = "T_Nomenclature.Collectee_le"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 0) = "T_Nomenclature.Commandee_le"
'I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "T_Nomenclature.Commantaire"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "t_Vehicule.DatePrevue"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "t_Vehicule.Demandeur"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "t_Vehicule.Désignation"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "T_Nomenclature.Designation_Piece"
I_Champ = I_Champ + 1
'Champs(I_Champ, 0) = "T_Nomenclature.[Désination FE]"
'I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "T_Nomenclature.En_provenance_de"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 0) = "T_Nomenclature.Expedee_le"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 0) = "T_Nomenclature.[Famille/FE/V]"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "T_Nomenclature.Indice"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "T_Nomenclature.Masse"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "T_SousEnsembles.[MASSE total (KG)]"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "T_Nomenclature.Modifiee_le"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "T_Nomenclature.Num_de_retour"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "T_Nomenclature.N_de_BL"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "T_SousEnsembles.[Nb lignes sans masse]"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "T_SousEnsembles.[Nb total de lignes]"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 0) = "T_Nomenclature.Num_appro"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 0) = "T_Nomenclature.[Num_appro_ S]"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "t_Vehicule.NumEncelade"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 0) = "T_Nomenclature.Observation_ou_magasin"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "T_Nomenclature.Pièce"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "t_Vehicule.Programme"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "t_Vehicule.Projet"
I_Champ = I_Champ + 1
'Champs(I_Champ, 0) = "T_Nomenclature.Qt_Reaffectee"
'I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "T_Nomenclature.Qt_reçue"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 0) = "T_Nomenclature.QT_Retournee"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "T_Nomenclature.Quantite"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "T_Nomenclature.Reaffectee_a"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 0) = "T_Nomenclature.Reaffectee_Le"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 0) = "t_Vehicule.Realignement"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "T_Nomenclature.Reçue_le"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "T_Nomenclature.Reste_a_recevoir"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "T_Nomenclature.Retournee_le"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "T_SousEnsembles.[Sous-Ensseble]"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 0) = "T_Nomenclature.Type"
I_Champ = I_Champ + 1
Champs(I_Champ, 0) = "t_Vehicule.Véhicule"

I_Champ = 1
Champs(I_Champ, 1) = "Atelier"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 1) = "Bon sortie"

I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Code"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 1) = "Collectee le"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 1) = "Commandee le"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Commantaire"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "DatePrevue"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Demandeur"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Désignation"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Designation_Piece"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 1) = "Désination FE"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "En_provenance de"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 1) = "Expedee le"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 1) = "Famille/FE/V"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Indice"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Masse"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "MASSE total (KG)"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Modifiee le"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "N° de levraison"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "N°de réception"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Nb lignes sans masse"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Nb total de lignes"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 1) = "Num_appro"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 1) = "Num appro S"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "NumEncelade"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 1) = "Observation ou magasin"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Pièce"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Programme"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Projet"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 1) = "Qt Reaffectee"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Qt reçue"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 1) = "QT Retournee"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Quantite"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Reaffectee a"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 1) = "Reaffectee Le"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 1) = "Realignement"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Reçue le"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Reste a recevoir"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Retournee le"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Sous-Ensseble"
'I_Champ = I_Champ + 1
'Champs(I_Champ, 1) = "Type"
I_Champ = I_Champ + 1
Champs(I_Champ, 1) = "Véhicule"


'For I = 0 To 48
'    Debug.Print I & " := "; Champs(I, 0) & " : " & Champs(I, 1)
'Next




 


    If Session("candidat_UserType") <> "Administrator" Then
        Sql = Sql & " and con_FieldDefs.FieldAttribut Not Like '%SYS%' "
 
    End If
    If Session("Droits_SockEncelade") <> 1 Then
         Sql = Sql & " and con_FieldDefs.FieldAttribut Not Like '%ENC%' "
    End If
    Sql = Sql & "ORDER BY con_FieldDefs.FieldAlias;"
   'Debug.modSEARCH=modSEARCH & int Replace(Sql, "%", "*")
    'Set Rs = GeneralTools.My_Conn.Execute(Sql)
    IndexChamps = 0
   ' While Rs.EOF = False
    'if InStr(1, Ucase(Rs!FieldAttribut), "CRI1") <> 0 Then
    ' IndexChamps = IndexTab
    'end if
    'ReDim preserve TableauSherch(IndexTab)
    ' TableauSherch(IndexTab) = Rs("FieldName") & ";" & Rs("FieldAlias") & ";" & Rs("FieldAttribut")
    'IndexTab = IndexTab + 1
        'Rs.MoveNext
    'Wend
            modSEARCH = modSEARCH & "<SCRIPT langage='JAVASCRIPT'>" & vbCrLf
            
            modSEARCH = modSEARCH & "function MyChangeFilde(MyChangeFilde_inFORM,MyChangeFilde_Obj) {" & vbCrLf
            modSEARCH = modSEARCH & "    MyChangeFilde_oSlt = eval(""document."" + MyChangeFilde_inFORM + ""."" + MyChangeFilde_Obj);" & vbCrLf
            modSEARCH = modSEARCH & "                    MyChangeFilde_oSlt.value='';" & vbCrLf
             modSEARCH = modSEARCH & " document.CloseWhere.submit();" & vbCrLf
           
            modSEARCH = modSEARCH & "}" & vbCrLf
            modSEARCH = modSEARCH & "function CheckedRadiobutton(IndexRadiobutton,nbLignes){" & vbCrLf
            modSEARCH = modSEARCH & "for(i=1;i<nbLignes+1;i++) {" & vbCrLf
            modSEARCH = modSEARCH & "}" & vbCrLf
             modSEARCH = modSEARCH & "            }" & vbCrLf
            
            modSEARCH = modSEARCH & "function MyChange( MyChange_inFORM,MyChange_Obj,MyChange_Obj2,MyChange_Index) {" & vbCrLf
            modSEARCH = modSEARCH & "    MyChange_oSlt = eval(""document."" + MyChange_inFORM + ""."" + MyChange_Obj);" & vbCrLf
            modSEARCH = modSEARCH & "    MyChange_oSlt2 = eval(""document."" + MyChange_inFORM + ""."" + MyChange_Obj2);" & vbCrLf
            modSEARCH = modSEARCH & "    MyChange_oSlt3 = eval(""document."" + MyChange_inFORM + "".nbSelect"" );" & vbCrLf
            modSEARCH = modSEARCH & "    for (MyChange_k = 0; MyChange_k < MyChange_oSlt.options.length; MyChange_k++) {" & vbCrLf
            modSEARCH = modSEARCH & "        if (MyChange_oSlt.options[MyChange_k].selected == true) {" & vbCrLf
            modSEARCH = modSEARCH & "            var MyChange_STRTEXTE =MyChange_oSlt.options[MyChange_k].text;" & vbCrLf
            modSEARCH = modSEARCH & "            if ((MyChange_STRTEXTE  ==""(Aucun)"") && (MyChange_oSlt2.value!=""(Aucun)"")) {" & vbCrLf
            modSEARCH = modSEARCH & "                    MyChange_oSlt3.value=MyChange_Index;" & vbCrLf
            modSEARCH = modSEARCH & "                    MyChange_oSlt2.value=MyChange_STRTEXTE;" & vbCrLf
             modSEARCH = modSEARCH & " document.CloseWhere.submit();" & vbCrLf
            modSEARCH = modSEARCH & "            } else {" & vbCrLf
            modSEARCH = modSEARCH & "                if ((MyChange_STRTEXTE  !=""(Aucun)"") && (MyChange_oSlt2.value==""(Aucun)"")) {" & vbCrLf
            modSEARCH = modSEARCH & "                    MyChange_oSlt2.value=MyChange_STRTEXTE;" & vbCrLf
             modSEARCH = modSEARCH & " document.CloseWhere.submit();" & vbCrLf
            modSEARCH = modSEARCH & "               }" & vbCrLf
            modSEARCH = modSEARCH & "            }" & vbCrLf
            modSEARCH = modSEARCH & "        }" & vbCrLf
            modSEARCH = modSEARCH & "    }" & vbCrLf
            modSEARCH = modSEARCH & "}" & vbCrLf
            '
            
          
            
            modSEARCH = modSEARCH & "function mySubmit(Mode) {" & vbCrLf
            modSEARCH = modSEARCH & " myForm = this.document.forms[0];" & vbCrLf
             modSEARCH = modSEARCH & "if  (Mode='EXE'){ " & vbCrLf
            modSEARCH = modSEARCH & "myForm.action = 'Default.asp?Logistique_Page='+myForm.SersheBASE.value;" & vbCrLf
             modSEARCH = modSEARCH & "} else {" & vbCrLf
             modSEARCH = modSEARCH & "myForm.action = ""Default.asp?Logistique_Page=9"";" & vbCrLf
             
            modSEARCH = modSEARCH & "}" & vbCrLf
            modSEARCH = modSEARCH & " myForm.EnceladeSearch.value=""Search"" + Mode;" & vbCrLf
            modSEARCH = modSEARCH & " myForm.SherchPage.value=myForm.SersheBASE.value;" & vbCrLf
           
            modSEARCH = modSEARCH & " document.CloseWhere.submit();" & vbCrLf
            
          
              modSEARCH = modSEARCH & "}" & vbCrLf
            
'            Session("candidat_MainTable") = "con_Contacts"
'
'
        modSEARCH = modSEARCH & "  </SCRIPT>" & vbCrLf
              
    modSEARCH = modSEARCH & "<form name=""CloseWhere"" method=""post"" action=""Default.asp?Logistique_Page=11"" >" & vbCrLf
    modSEARCH = modSEARCH & " <input type=""hidden"" name=""SherchPage"" value=''>" & vbCrLf
    modSEARCH = modSEARCH & "  <input type=""hidden"" name=""EnceladeSearch"" >" & vbCrLf
    modSEARCH = modSEARCH & " <table width=""100%"" border=""2"" bgcolor='#B7B8DE' cellpadding=""1"" cellspacing=""0"">" & vbCrLf
    modSEARCH = modSEARCH & "  <tr valign=""bottom"">" & vbCrLf
    modSEARCH = modSEARCH & "   <td>" & vbCrLf
    modSEARCH = modSEARCH & "    <table width=""100%"" border=""0"">" & vbCrLf
    modSEARCH = modSEARCH & "     <tr valign=""bottom"">" & vbCrLf
    modSEARCH = modSEARCH & "       <td align=""left"" valign=""middle"">" & vbCrLf
    modSEARCH = modSEARCH & "<font color='#FFFFFF' size='2' >" & vbCrLf
    modSEARCH = modSEARCH & "        " & Session("Encelade_Message")
    modSEARCH = modSEARCH & "</font>" & vbCrLf
    modSEARCH = modSEARCH & "       </td>" & vbCrLf
    modSEARCH = modSEARCH & "       <td>" & vbCrLf
'    modSEARCH=modSEARCH &  "        <font size=""1"" > Nombre de lignes a afficher </font> " & vbcrlf
    If Session("candidat_UserType") = "Administrator" Then
       ' modSEARCH=modSEARCH &  "        <INPUT TYPE=""hidden"" NAME =""NbLiges"" value=""" & GetDefault("con_PageNum", "15", Session("candidat_ADOContact")) & """>")

'   con_frmSetting = con_frmSetting & vbCrLf & ("<TD><input type=""text"" NAME=""NBLignesRecherche"" VALUE=" & GetDefault("NBLignesRecherche", "25", Session("candidat_ADOContact")) & "></TD>       ")
   Else
    modSEARCH = modSEARCH & "        <INPUT TYPE=""hidden"" NAME =""NbLiges"" value='25'>" & vbCrLf
    End If
    modSEARCH = modSEARCH & "       </td>" & vbCrLf
    modSEARCH = modSEARCH & "       <td align=""right"" class=""smallerheader"" valign=""middle""></td>" & vbCrLf
   
    modSEARCH = modSEARCH & "       <td align=""right"" color='#FFFFFF'  valign=""Right"" ><FONT  color='#FFFFFF' >Recherche sur</FONT></td>" & vbCrLf
     modSEARCH = modSEARCH & "       <td align=""left"" class=""smallerheader"" valign=""left"">" & vbCrLf
      modSEARCH = modSEARCH & "<select Name='SersheBASE' >" & vbCrLf
      If Request("SersheBASE") = "9" Then Selected = " selected " Else Selected = ""
       modSEARCH = modSEARCH & "<option value='22' " & Selected & " >Projets</option>" & vbCrLf
    modSEARCH = modSEARCH & "<option value='9' " & Selected & " >Véhicule</option>" & vbCrLf
     If Request("SersheBASE") = "10" Then Selected = " selected " Else Selected = ""
    modSEARCH = modSEARCH & "<option value='10'" & Selected & ">Sous-Ensembles</option>" & vbCrLf
     If Request("SersheBASE") = "8" Then Selected = " selected " Else Selected = ""
    modSEARCH = modSEARCH & "<option value='8'" & Selected & ">Pièces</option>" & vbCrLf
   
   
   If Request("SersheBASE") = "12" Then Selected = " selected " Else Selected = ""
    modSEARCH = modSEARCH & "<option value='12'" & Selected & ">B. Réception" & vbCrLf
    
    If Request("SersheBASE") = "13" Then Selected = " selected " Else Selected = ""
    modSEARCH = modSEARCH & "<option value='13'" & Selected & ">B. Livraison" & vbCrLf
      modSEARCH = modSEARCH & "</select>" & vbCrLf
           modSEARCH = modSEARCH & "</td>" & vbCrLf
    ' modSEARCH=modSEARCH &  "       <td align=""center"" class=""smallerheader"" valign=""middle"" width='3%' >" & vbcrlf
     ' modSEARCH=modSEARCH & AfficheBountonSersh("Annuler", "", """javascript:mySubmit('Annuler');""" )
      modSEARCH = modSEARCH & "     </td>  <td align=""center"" class=""smallerheader"" valign=""middle"" width='3%' >" & vbCrLf
       modSEARCH = modSEARCH & AfficheBountonSersh("Valider", "", """javascript:mySubmit('EXE');""")
   modSEARCH = modSEARCH & "       </td>" & vbCrLf
    modSEARCH = modSEARCH & "      </tr>" & vbCrLf
    modSEARCH = modSEARCH & "     </table>" & vbCrLf
    modSEARCH = modSEARCH & "    </td>" & vbCrLf
    modSEARCH = modSEARCH & "   </tr>" & vbCrLf
    modSEARCH = modSEARCH & "  </table>" & vbCrLf
    modSEARCH = modSEARCH & "     <table  width=""100%"" >" & vbCrLf
    modSEARCH = modSEARCH & "     <TR width=><td></td></tr>" & vbCrLf
    modSEARCH = modSEARCH & "     <TR ><td>" & vbCrLf
      modSEARCH = modSEARCH & "     <table bgcolor='#B7B8DE' width=""70%"" align=""center"">" & vbCrLf
    modSEARCH = modSEARCH & "    <TR>" & vbCrLf
    modSEARCH = modSEARCH & "     <TD  >" & vbCrLf
    modSEARCH = modSEARCH & "      <font  color='#FFFFFF' size='1' >    TRIER  A -> Z  </font>  " & vbCrLf
    modSEARCH = modSEARCH & "     </TD>" & vbCrLf
    modSEARCH = modSEARCH & "     <TD  >" & vbCrLf
    modSEARCH = modSEARCH & "      <font  color='#FFFFFF' size='1' >    TRIER  Z -> A    </font>  " & vbCrLf
    modSEARCH = modSEARCH & "     </TD>" & vbCrLf
    modSEARCH = modSEARCH & "     <TD  >" & vbCrLf
    modSEARCH = modSEARCH & "      <font  color='#FFFFFF' size='1' >    NOM DES CHAMPS   </font>  " & vbCrLf
    modSEARCH = modSEARCH & "     </TD>" & vbCrLf
    modSEARCH = modSEARCH & "     <TD>" & vbCrLf
    modSEARCH = modSEARCH & "      <font  color='#FFFFFF' size='1' >  OPERATEUR  </font> " & vbCrLf
    modSEARCH = modSEARCH & "     </TD>" & vbCrLf
    modSEARCH = modSEARCH & "     <TD>" & vbCrLf
    modSEARCH = modSEARCH & "      <font  color='#FFFFFF' size='1' >  VALEUR  </font> " & vbCrLf
    modSEARCH = modSEARCH & "     </TD>" & vbCrLf
    modSEARCH = modSEARCH & "     <TD>" & vbCrLf
    modSEARCH = modSEARCH & "      <font  color='#FFFFFF' size='1' > OPERATEUR LOGIQUE</font>" & vbCrLf
    modSEARCH = modSEARCH & "     </TD>" & vbCrLf
    modSEARCH = modSEARCH & "   </TR>" & vbCrLf
    
  I = 1
    
 '**************************************************************************************************************
 'Existe
 '**************************************************************************************************************
  If nbLignes > 0 Then
     For I = 1 To nbLignes - 1
    modSEARCH = modSEARCH & "   <TR>" & vbCrLf
    
       SelecetList = ""
                SelecetList = "" & Request("radiobutton")
                If SelecetList = CStr(I) Then
                 modSEARCH = modSEARCH & "<TD><input type=""radio"" name=""radiobutton"" value='" & CStr(I) & "' Checked></TD>" & vbCrLf
                Else
                 modSEARCH = modSEARCH & "<TD><input type=""radio"" name=""radiobutton"" value='" & CStr(I) & "'></TD>" & vbCrLf
                End If
                
               
                
                If SelecetList = CStr(I) & "DESC" Then
                modSEARCH = modSEARCH & "<TD><input type=""radio"" name=""radiobutton"" value='" & CStr(I) & "DESC' Checked></TD>" & vbCrLf
                Else
                 modSEARCH = modSEARCH & "<TD><input type=""radio"" name=""radiobutton"" value='" & CStr(I) & "DESC'></TD>" & vbCrLf
                End If
    modSEARCH = modSEARCH & "   <TD>" & vbCrLf
    modSEARCH = modSEARCH & "   <select name='Champ" & CStr(I) & "' >" & vbCrLf
    For j = 1 To UBound(Champs)
    If Request("Champ" & CStr(I)) = Champs(j, 0) Then Selected = " selected " Else Selected = ""
  modSEARCH = modSEARCH & " <option value='" & Champs(j, 0) & "'" & Selected & ">" & Champs(j, 1) & "</option>" & vbCrLf
Next
  modSEARCH = modSEARCH & " </select>" & vbCrLf
    modSEARCH = modSEARCH & "   </TD>" & vbCrLf
     modSEARCH = modSEARCH & "</SELECT>" & vbCrLf
                
                modSEARCH = modSEARCH & "         <TD>" & vbCrLf
                modSEARCH = modSEARCH & "<SELECT NAME=""OPERATEUR" & CStr(I) & """>" & vbCrLf
      
                If Request("OPERATEUR" & CStr(I)).Count > 0 Then
                    slt = Request("OPERATEUR" & I)
                Else
                    slt = ""
                End If
                For j = 0 To UBound(aSlt_OPERATEUR, 2)
                     If Trim(slt) = Trim(aSlt_OPERATEUR(0, j)) Then
                         modSEARCH = modSEARCH & " <option value=""" & aSlt_OPERATEUR(0, j) & """ selected>" & aSlt_OPERATEUR(1, j) & "</option>" & vbCrLf
                     Else
                        modSEARCH = modSEARCH & " <option value=""" & aSlt_OPERATEUR(0, j) & """>" & aSlt_OPERATEUR(1, j) & "</option>" & vbCrLf
                     End If
                 Next
       modSEARCH = modSEARCH & "</SELECT>" & vbCrLf
       modSEARCH = modSEARCH & " <td><input type=""text"" NAME =""VALEUR" & CStr(I) & """ value=" & Request("VALEUR" & CStr(I)) & "></td>" & vbCrLf
       modSEARCH = modSEARCH & "         <TD>" & vbCrLf
       modSEARCH = modSEARCH & "               <font class=""smallerheader"">  <select name=""Logique" & CStr(I) & """ onChange=""MyChange('CloseWhere',  'Logique" & CStr(I) & "','SaveLogique" & CStr(I) & "'," & CStr(I) & ");"">" & vbCrLf
        
       
        If Request("Logique" & CStr(I)).Count > 0 Then
            slt_Logique = Request("Logique" & I)
        Else
            slt_Logique = ""
        End If
       For j = 0 To UBound(aSlt_Logique, 2)
            If Trim(slt_Logique) = Trim(aSlt_Logique(0, j)) Then
                modSEARCH = modSEARCH & " <option value=""" & aSlt_Logique(0, j) & """ selected>" & aSlt_Logique(1, j) & "</option>" & vbCrLf
            Else
                modSEARCH = modSEARCH & " <option value=""" & aSlt_Logique(0, j) & """>" & aSlt_Logique(1, j) & "</option>" & vbCrLf
            End If
        Next
        
       modSEARCH = modSEARCH & "              </select></font>" & vbCrLf
       modSEARCH = modSEARCH & " <INPUT TYPE='hidden' NAME ='SaveLogique" & CStr(I) & "'  VALUE='" & Request("Logique" & CStr(I)) & "'>" & vbCrLf
       modSEARCH = modSEARCH & "         </TD>" & vbCrLf
    modSEARCH = modSEARCH & "   </TR>" & vbCrLf
    If Request("Logique" & CStr(I)) = "(Aucun)" Then
        Fin = True
        Exit For
        
    Else
        Fin = False
    End If
   Next
  End If
 '**************************************************************************************************************
 'New
 '**************************************************************************************************************
     If Fin = False Then
    modSEARCH = modSEARCH & "   <TR>" & vbCrLf
    
      modSEARCH = modSEARCH & "<TD><input type=""radio"" name=""radiobutton"" value=""" & CStr(I) & """></TD>" & vbCrLf
        
         modSEARCH = modSEARCH & "<TD><input type=""radio"" name=""radiobutton"" value=""" & CStr(I) & "DESC""></TD>" & vbCrLf
    modSEARCH = modSEARCH & "   <TD>" & vbCrLf
    modSEARCH = modSEARCH & "   <select name='Champ" & CStr(I) & "' >" & vbCrLf
    For j = 1 To UBound(Champs)
  modSEARCH = modSEARCH & " <option value='" & Champs(j, 0) & "'>" & Champs(j, 1) & "</option>" & vbCrLf
Next
  modSEARCH = modSEARCH & " </select>" & vbCrLf
    modSEARCH = modSEARCH & "   </TD>" & vbCrLf
     modSEARCH = modSEARCH & "</SELECT>" & vbCrLf
                
                modSEARCH = modSEARCH & "         <TD>" & vbCrLf
                modSEARCH = modSEARCH & "<SELECT NAME=""OPERATEUR" & CStr(I) & """>" & vbCrLf
      
                If Request("OPERATEUR" & CStr(I)).Count > 0 Then
                    slt = Request("OPERATEUR" & I)
                Else
                    slt = ""
                End If
                For j = 0 To UBound(aSlt_OPERATEUR, 2)
                     If j = 2 Then
                         modSEARCH = modSEARCH & " <option value=""" & aSlt_OPERATEUR(0, j) & """ selected>" & aSlt_OPERATEUR(1, j) & "</option>" & vbCrLf
                     Else
                        modSEARCH = modSEARCH & " <option value=""" & aSlt_OPERATEUR(0, j) & """>" & aSlt_OPERATEUR(1, j) & "</option>" & vbCrLf
                     End If
                 Next
       modSEARCH = modSEARCH & "</SELECT>" & vbCrLf
       modSEARCH = modSEARCH & " <td><input type=""text"" NAME =""VALEUR" & CStr(I) & """ value=" & txt & "></td>" & vbCrLf
       modSEARCH = modSEARCH & "         <TD>" & vbCrLf
       modSEARCH = modSEARCH & "               <font class=""smallerheader"">  <select name=""Logique" & CStr(I) & """ onChange=""MyChange('CloseWhere',  'Logique" & CStr(I) & "','SaveLogique" & CStr(I) & "'," & CStr(I) & ");"">" & vbCrLf
        
       
        If Request("Logique" & CStr(I)).Count > 0 Then
            slt_Logique = Request("Logique" & I)
        Else
            slt_Logique = ""
        End If
       For j = 0 To UBound(aSlt_Logique, 2)
            If Trim(slt_Logique) = Trim(aSlt_Logique(0, j)) Then
                modSEARCH = modSEARCH & " <option value=""" & aSlt_Logique(0, j) & """ selected>" & aSlt_Logique(1, j) & "</option>" & vbCrLf
            Else
                modSEARCH = modSEARCH & " <option value=""" & aSlt_Logique(0, j) & """>" & aSlt_Logique(1, j) & "</option>" & vbCrLf
            End If
        Next
        
       modSEARCH = modSEARCH & "              </select></font>" & vbCrLf
       modSEARCH = modSEARCH & " <INPUT TYPE='hidden' NAME ='SaveLogique" & CStr(I) & "'  VALUE='(Aucun)'>" & vbCrLf
       modSEARCH = modSEARCH & "         </TD>" & vbCrLf
    modSEARCH = modSEARCH & "   </TR>" & vbCrLf
   End If
    
    
     modSEARCH = modSEARCH & "</table>" & vbCrLf
       modSEARCH = modSEARCH & "    <td></td></tr>" & vbCrLf
       modSEARCH = modSEARCH & "    <TR >" & vbCrLf
       modSEARCH = modSEARCH & "    <td></td>" & vbCrLf
       modSEARCH = modSEARCH & "    </tr>" & vbCrLf
       modSEARCH = modSEARCH & "<input type='hidden' name='nbSelect' value='" & I & "'>" & vbCrLf
   modSEARCH = modSEARCH & "</form>" & vbCrLf
  
      

End Function


Function MyCloseWhere()
Dim txt
Dim TxtChamp
Dim TxtChamp2
Dim TxtOperateur
Dim TxtValeur
Dim TxtLogique
Dim BoolNull
Dim IsNuber

 Session("CloseWhereFrom") = ""
 Dim radiobutton
 Session("CloseWhereOrder") = ""
  Session("CloseWhereFrom") = ""
'.catname
txt = ""
   radiobutton = Request("radiobutton")
nbLignes = CInt(Request("nbSelect"))
  For I = 1 To nbLignes


 TxtOperateur = Request("Operateur" & CStr(I))
    TxtChamp = Request("Champ" & CStr(I))
    TxtChamp2 = TxtChamp
    TxtValeur = Request("Valeur" & CStr(I))
  
 



    If Trim(TxtValeur) = "" Then
        BoolNull = True
    Else
        BoolNull = False
       IsNuber = RetournIsNum(TxtChamp)
       
    End If
   
    TxtLogique = Request("Logique" & CStr(I))
 
    TxtChamp = TxtChamp
    
 
 
 
  
  
   If radiobutton = CStr(I) Then
        Session("CloseWhereOrder") = TxtChamp & " "
    End If
  If radiobutton = CStr(I) & "DESC" Then
        Session("CloseWhereOrder") = TxtChamp & " Desc "
    End If
  Select Case Trim(TxtOperateur)
       
        Case "<"
               
                If BoolNull = False Then
                    If IsNuber = "True" Then
                        txt = txt & "(" & TxtChamp & "  <   " & noquote("" & TxtValeur, "num") & " ) " & TxtLogique
                    Else
                        If IsNuber = "Date" Then
                            txt = txt & " (" & TxtChamp & " < " & noquote("" & TxtValeur, "date") & " ) " & TxtLogique
                        Else
                        txt = txt & " (" & TxtChamp & " < " & noquote("" & TxtValeur, "txt") & " ) " & TxtLogique
                        End If
                    End If
                    
                Else
                    
                    txt = txt & " (" & TxtChamp & " Is NULL ) " & TxtLogique
                End If
        Case "<="
                If BoolNull = False Then
                   If IsNuber = "True" Then
                  
                        txt = txt & TxtChamp & " <= " & noquote("" & TxtValeur, "num") & " ) " & TxtLogique
                   Else
                        If IsNuber = "Date" Then
                            txt = txt & " (" & TxtChamp & " <= " & noquote("" & TxtValeur, "date") & " ) " & TxtLogique
                        Else
                            txt = txt & " (" & TxtChamp & "<= " & noquote("" & TxtValeur, "txt") & " ) " & TxtLogique
                        End If
                   End If
                Else
                      
                    txt = txt & " (" & TxtChamp & " Is NULL ) " & TxtLogique
                End If
                
        Case "="
                If BoolNull = False Then
                    If IsNuber = "True" Then
                        txt = txt & "(" & TxtChamp & "  =   " & noquote("" & TxtValeur, "num") & " ) " & TxtLogique
                    Else
                        If IsNuber = "Date" Then
                            txt = txt & " (" & TxtChamp & " = " & noquote("" & TxtValeur, "date") & " ) " & TxtLogique
                        Else
                                txt = txt & " (" & TxtChamp & " = " & noquote("" & TxtValeur, "txt") & ") " & TxtLogique
                        End If
                    End If
                Else
                      
                    txt = txt & " (" & TxtChamp & " Is NULL ) " & TxtLogique
                End If
                 
        Case "=>"
                If BoolNull = False Then
                If IsNuber = "True" Then
                        txt = txt & "(" & TxtChamp & "  >=  " & noquote("" & TxtValeur, "num") & "  ) " & TxtLogique
                Else
                        If IsNuber = "Date" Then
                            txt = txt & " (" & TxtChamp & " >= " & noquote("" & TxtValeur, "date") & " ) " & TxtLogique
                        Else
                            txt = txt & " (" & TxtChamp & ">= " & noquote("" & TxtValeur, "txt") & " ) " & TxtLogique
                        End If
                  End If
                Else
                      
                    txt = txt & " (" & TxtChamp & " Is NULL ) " & TxtLogique
                End If
                 
        Case ">"
                If BoolNull = False Then
                    
                    If IsNuber = "True" Then
                        txt = txt & "(" & TxtChamp & "  >   " & noquote("" & TxtValeur, "num") & "  ) " & TxtLogique
                    Else
                        If IsNuber = "Date" Then
                            txt = txt & " (" & TxtChamp & " > " & noquote("" & TxtValeur, "date") & " ) " & TxtLogique
                        Else
                            txt = txt & " (" & TxtChamp & " > " & noquote("" & TxtValeur, "txt") & ") " & TxtLogique
                        End If
                    End If
                Else
                      
                    txt = txt & " (" & TxtChamp & " Is NULL ) " & TxtLogique
                End If
                 
        Case "<>"
                If BoolNull = False Then
                 
                    If IsNuber = "True" Then
                        txt = txt & "(" & TxtChamp & "  <>" & noquote("" & TxtValeur, "num") & " ) " & TxtLogique
                    Else
                        If IsNuber = "Date" Then
                            txt = txt & " (" & TxtChamp & " <> " & noquote("" & TxtValeur, "date") & " ) " & TxtLogique
                        Else
                            txt = txt & " (" & TxtChamp & " <> " & noquote("" & TxtValeur, "txt") & ") " & TxtLogique
                        End If
                    End If
                Else
                      If Session("portal_candidat_Recherhe") = 1 Then
                        MyCloseWhere = ""
                        Exit Function
                      End If
                    txt = txt & " (" & TxtChamp & " Is NULL ) " & TxtLogique
                End If
                 
        Case "LIKE1"
                If BoolNull = False Then
                     TxtValeur = TxtValeur & "*"
                Else
                      If Session("portal_candidat_Recherhe") = 1 Then
                        MyCloseWhere = ""
                        Exit Function
                      End If
                    txt = txt & " (" & TxtChamp & " Is NULL ) " & TxtLogique
                End If
               
                TxtValeur = Replace(TxtValeur, "*", "%", 1)
                TxtValeur = Replace(TxtValeur, "'", "''", 1)
                TxtValeur = Replace(TxtValeur, Chr(34), Chr(34) & Chr(34), 1)
                txt = txt & " (" & TxtChamp & " LIKE '" & TxtValeur & "') " & TxtLogique

        Case "LIKE2"
                If BoolNull = False Then
                    TxtValeur = "*" & TxtValeur
                Else
                      If Session("portal_candidat_Recherhe") = 1 Then
                        MyCloseWhere = ""
                        Exit Function
                      End If
                     
                    txt = txt & " (" & TxtChamp & " Is NULL ) " & TxtLogique
                End If
               
                TxtValeur = Replace(TxtValeur, "*", "%", 1)
                TxtValeur = Replace(TxtValeur, "'", "''", 1)
                TxtValeur = Replace(TxtValeur, Chr(34), Chr(34) & Chr(34), 1)
                txt = txt & " (" & TxtChamp & " LIKE '" & TxtValeur & "') " & TxtLogique
         
         Case "LIKE"
               
                      If Session("portal_candidat_Recherhe") = 1 Then
                        MyCloseWhere = ""
                        Exit Function
                     End If
                    'txt = txt & " (" & TxtChamp & " Is NULL ) " & TxtLogique
               
            
               
                TxtValeur = Replace(TxtValeur, "'", "''", 1)
                TxtValeur = Replace(TxtValeur, Chr(34), Chr(34) & Chr(34), 1)
                txt = txt & " (" & TxtChamp & " LIKE '%" & TxtValeur & "%') " & TxtLogique
  End Select
  
   
Next
txt = " " & Replace(txt, "(Aucun)", "", 1) & " "
MyCloseWhere = " WHERE " & txt

End Function


Function RetourVehicule(FormName, Updatable, Url)
If Request("MAJ") = "OK" Then
For I = 1 To Request("NbLigne")
    If Request("Id_L" & I) <> "" Then
        Sql = "UPDATE t_Vehicule SET t_Vehicule.NumEncelade = " & noquote(Request("NumEncelade_L" & I), "txt") & ", "
        Sql = Sql & "t_Vehicule.Id_Projects_NumAffaire = " & noquote(Request("Projet_L" & I), "num") & ",  "
        Sql = Sql & "t_Vehicule.Programme = " & noquote(Request("Programme_L" & I), "txt") & ",  "
        Sql = Sql & "t_Vehicule.Désignation = " & noquote(Request("Désignation_L" & I), "txt") & ",  "
        Sql = Sql & "t_Vehicule.Véhicule = " & noquote(Request("Véhicule_L" & I), "txt") & ",  "
        Sql = Sql & "t_Vehicule.Realignement = " & noquote(Request("Véhicule_L" & I), "txt") & ",  "
        Sql = Sql & "t_Vehicule.Demandeur = " & noquote(Request("Demandeur_L" & I), "txt") & ",  "
        Sql = Sql & "t_Vehicule.Atelier = " & noquote(Request("Atelier_L" & I), "txt") & ",  "
        Sql = Sql & "t_Vehicule.DatePrevue = " & noquote(Request("DatePrevue_L" & I), "date") & " "
        Sql = Sql & "WHERE t_Vehicule.Id=" & Request("Id_L" & I) & ";"
        'Response.Write sql
        'Response.End
        con.Execute Sql
       
        If Request("Delet_L" & I) = "Delete" Then
            Sql = "DELETE FROM t_Vehicule "
            Sql = Sql & "WHERE t_Vehicule.Id=" & Request("Id_L" & I) & ";"
             'Response.Write sql
        'Response.End
            con.Execute Sql
        End If
     Else
     txt = Request("NumEncelade_L" & I)
    txt = txt & Request("Projet_L" & I)
        txt = txt & Request("Programme_L" & I)
        txt = txt & Request("Désignation_L" & I)
        txt = txt & Request("Véhicule_L" & I)
        txt = txt & Request("Realignement_L" & I)
        txt = txt & Request("Demandeur_L" & I)
        txt = txt & Request("Atelier_L" & I)
        txt = txt & Request("DatePrevue_L" & I)
        If Trim("" & txt) <> "" Then
            Sql = "INSERT INTO t_Vehicule ( NumEncelade, Id_Projects_NumAffaire, Programme, Désignation, Véhicule, Realignement, "
            Sql = Sql & "Demandeur, Atelier, DatePrevue )"
            Sql = Sql & "VALUES ( " & noquote(Request("NumEncelade_L" & I), "txt") & " ,  "
            Sql = Sql & noquote(Request("Projet_L" & I), "num") & ",  "
            Sql = Sql & noquote(Request("Programme_L" & I), "txt") & " ,  "
            Sql = Sql & noquote(Request("Désignation_L" & I), "txt") & " ,  "
            Sql = Sql & noquote(Request("Véhicule_L" & I), "txt") & ",  "
            Sql = Sql & noquote(Request("Realignement_L" & I), "txt") & ",  "
            Sql = Sql & noquote(Request("Demandeur_L" & I), "txt") & ",  "
            Sql = Sql & noquote(Request("Atelier_L" & I), "txt") & " ,  "
            Sql = Sql & noquote(Request("DatePrevue_L" & I), "date") & " );"
            'Response.Write txt & " " & i
            'Response.End
            con.Execute Sql
        End If
     End If


Next
End If

  'RetourVehicule=HISTORIQUE_Nommenclature
     RetourVehicule = RetourVehicule & "<form name='" & FormName & "' method='post'>" & vbCrLf
     RetourVehicule = RetourVehicule & "<input name='MAJ' type='hidden' value='OK'/>" & vbCrLf
     RetourVehicule = RetourVehicule & "<input name='Updatable' type='hidden' value='OK'/>" & vbCrLf
      RetourVehicule = RetourVehicule & "<input name='FormName' type='hidden' value='" & FormName & "'/>" & vbCrLf
      RetourVehicule = RetourVehicule & "<input name='id_Vehicule' type='hidden' value='" & Request("id_Vehicule") & "' />"
       RetourVehicule = RetourVehicule & "<input name='Pièce' type='hidden' value='" & Request("Pièce") & "' />"
       RetourVehicule = RetourVehicule & "<input name='ID__SousEnsembles' type='hidden' value='" & Request("ID__SousEnsembles") & "' />"
 
           
      RetourVehicule = RetourVehicule & "<table border='0' >" & vbCrLf
       If Request("nbSelect") > 0 Then
            RetourVehicule = RetourVehicule & "<tr>" & vbCrLf
            RetourVehicule = RetourVehicule & "         <td colspan='6' >" & RecapSersh & "</td>" & vbCrLf
            
            RetourVehicule = RetourVehicule & "</tr>" & vbCrLf
         End If
      RetourVehicule = RetourVehicule & "<tr >" & vbCrLf
         RetourVehicule = RetourVehicule & "<td class='TitreTableau' >Liste des Véhicules :</td>" & vbCrLf
         RetourVehicule = RetourVehicule & "</tr>" & vbCrLf
          RetourVehicule = RetourVehicule & "<tr>" & vbCrLf
          RetourVehicule = RetourVehicule & "<TD width='100%'><table border='0'  width='100%'><tr>"
    If Session("Updatable") = "True" Then
    
        RetourVehicule = RetourVehicule & "<TD width='2' valign='top'>" & AfficheBounton("Valider", 2, "'javascript:valider();'") & "</td>" & vbCrLf
        
       
       
        
       End If
        
        
          
       
          RetourVehicule = RetourVehicule & "<TD width='2' valign='top'>"
         
    RetourVehicule = RetourVehicule & vbCrLf & "<SCRIPT langage='JAVASCRIPT'>"
       
    RetourVehicule = RetourVehicule & "function ModifierRechereche(){" & vbCrLf
    RetourVehicule = RetourVehicule & "this.document." & FormName & ".ID__SousEnsembles.value='';"
    RetourVehicule = RetourVehicule & "this.document." & FormName & ".action='Default.asp?Logistique_Page=11'" & vbCrLf
    RetourVehicule = RetourVehicule & vbCrLf & "document." & FormName & ".submit();"
      RetourVehicule = RetourVehicule & "}" & vbCrLf
    
        RetourVehicule = RetourVehicule & "function valider3(){" & vbCrLf
        RetourVehicule = RetourVehicule & "this.document." & FormName & ".id_Vehicule.value='';"
        RetourVehicule = RetourVehicule & "this.document." & FormName & ".action='Default.asp?Logistique_Page=9'" & vbCrLf
        RetourVehicule = RetourVehicule & vbCrLf & "document." & FormName & ".submit();"
        RetourVehicule = RetourVehicule & "}" & vbCrLf

        RetourVehicule = RetourVehicule & "function valider2(){" & vbCrLf
        RetourVehicule = RetourVehicule & "this.document." & FormName & ".action='Default.asp?Logistique_Page=2'" & vbCrLf
        RetourVehicule = RetourVehicule & vbCrLf & "document." & FormName & ".submit();"
        RetourVehicule = RetourVehicule & "}" & vbCrLf
        RetourVehicule = RetourVehicule & vbCrLf & "</SCRIPT >"
 RetourVehicule = RetourVehicule & "<TD width='2' valign='top'>" & AfficheBounton("Retour", 8, "'javascript:valider3();'") & "</td>" & vbCrLf
      RetourVehicule = RetourVehicule & "<TD>&nbsp;</td><TD>&nbsp;</td>" & vbCrLf
      
     
        RetourVehicule = RetourVehicule & "<TD width='100%' valign='top'>Filtrés sur véhicule :" & vbCrLf
        Sql = "SELECT t_Vehicule.Id, Projet,  Véhicule FROM t_Vehicule ORDER BY t_Vehicule.Véhicule;"
        
        
        Sql = "SELECT     t_Vehicule.Id, T_Projects.NumAffaire, t_Vehicule.Véhicule "
        Sql = Sql & "FROM t_Vehicule INNER JOIN "
        Sql = Sql & "T_Projects ON t_Vehicule.Id_Projects_NumAffaire = T_Projects.Id "
        Sql = Sql & "ORDER BY t_Vehicule.Véhicule"
        Set RsFiltre = OpenRecordSet(Sql, con)

       RetourVehicule = RetourVehicule & "<br>" & CreatListe(RsFiltre, Request("FiltreVehicule"), "Id", "1;2", "FiltreVehicule", " onChange='valider2();' ") & "</td>" & vbCrLf
      RsFiltre.Close
      Set RsFiltre = Nothing
      
     
     
     RetourVehicule = RetourVehicule & "</tr></table></td>"
     RetourVehicule = RetourVehicule & "</tr><tr><td> <table border='1' >" & vbCrLf
       Dim L
       Sql = "SELECT t_Vehicule.id,  t_Vehicule.id_Projects_NumAffaire,t_Vehicule.NumEncelade, t_Vehicule.Projet, t_Vehicule.Programme, t_Vehicule.Désignation,  "
        Sql = Sql & "t_Vehicule.Véhicule, t_Vehicule.Realignement, t_Vehicule.Demandeur, t_Vehicule.Atelier,  "
        Sql = Sql & "t_Vehicule.DatePrevue "
        Sql = Sql & "FROM t_Vehicule INNER JOIN "
        Sql = Sql & "T_Projects ON t_Vehicule.Id_Projects_NumAffaire = T_Projects.Id LEFT OUTER JOIN "
        Sql = Sql & "T_SousEnsembles ON t_Vehicule.Id = T_SousEnsembles.ID_Vehicule_Véhicule LEFT OUTER JOIN "
        Sql = Sql & "T_Nomenclature ON T_SousEnsembles.Id = T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble] "
        MyFilter = RetourneCloseWhere()
          If Request("nbSelect") > 0 Then
            Sql = Sql & MyCloseWhere()
             If Trim("" & MyFilter) <> "" Then
                Sql = Sql & " AND (" & MyFilter & ") "
                sqlw = "?"
             End If
          Else
            If Trim("" & MyFilter) <> "" Then
                Sql = Sql & " WHERE " & MyFilter
                sqlw = "?"
             End If
        End If
        If Trim("" & Request("FiltreVehicule")) <> "" Then
                    If InStr(1, UCase(Sql), "WHERE") <> 0 Then
                    
                Sql = Sql & " AND t_Vehicule.Id=" & Request("FiltreVehicule") & " "
                
            Else
                MyFilter = "?"
                 Sql = Sql & " WHERE t_Vehicule.Id=" & Request("FiltreVehicule") & " "
            End If
      End If
        Sql = Sql & "GROUP BY   t_Vehicule.id,t_Vehicule.id_Projects_NumAffaire, t_Vehicule.NumEncelade, t_Vehicule.Projet, t_Vehicule.Programme,  "
        Sql = Sql & "t_Vehicule.Désignation, t_Vehicule.Véhicule, t_Vehicule.Realignement, t_Vehicule.Demandeur,  "
        Sql = Sql & "t_Vehicule.Atelier, t_Vehicule.DatePrevue;"
'Response.Write sql
'Response.End
     Set rs_Vehicule = OpenRecordSet(Sql, con)
     
Sql = "SELECT    T_Projects.Id, T_Projects.NumAffaire, T_Projects.Projet From T_Projects ORDER BY T_Projects.NumAffaire"
       Set rs_Projet = OpenRecordSet(Sql, con)
        RetourVehicule = RetourVehicule & "<tr>" & vbCrLf
        RetourVehicule = RetourVehicule & "<td align='center' bgcolor='#CCCCCC'>Supprimer</td>" & vbCrLf
       
        
        RetourVehicule = RetourVehicule & "<td align='center' bgcolor='#CCCCCC'>N° Encelade</td>" & vbCrLf
        RetourVehicule = RetourVehicule & "<td align='center' bgcolor='#CCCCCC'>Projet</td>" & vbCrLf
        
        
         RetourVehicule = RetourVehicule & "<td align='center' bgcolor='#CCCCCC'>Programme</td>" & vbCrLf
          RetourVehicule = RetourVehicule & "<td align='center' bgcolor='#CCCCCC'>Désignation</td>" & vbCrLf
           RetourVehicule = RetourVehicule & "<td align='center' bgcolor='#CCCCCC'>Véhicule</td>" & vbCrLf
            RetourVehicule = RetourVehicule & "<td align='center' bgcolor='#CCCCCC'>Realignement</td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td align='center' bgcolor='#CCCCCC'>Demandeur</td>" & vbCrLf
        RetourVehicule = RetourVehicule & "<td align='center' bgcolor='#CCCCCC'>Atelier</td>" & vbCrLf
         RetourVehicule = RetourVehicule & "<td align='center' bgcolor='#CCCCCC'>Date Prevue</td>" & vbCrLf
             RetourVehicule = RetourVehicule & "</tr>" & vbCrLf
             L = 0
              If Trim("" & sqlw) = "" And Session("Nouveau") = "True" Then
            L = L + 1
            RetourVehicule = RetourVehicule & "<tr>" & vbCrLf
            RetourVehicule = RetourVehicule & "<td>"
            RetourVehicule = RetourVehicule & "<input name='Id_L" & L & "' type='hidden' value=''/>"
            RetourVehicule = RetourVehicule & "NEW</td>" & vbCrLf
            
             RetourVehicule = RetourVehicule & "<td><input name='NumEncelade_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
'             RetourVehicule = RetourVehicule & "<td><input name='Projet_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
                RetourVehicule = RetourVehicule & "<td>" & CreatListe(rs_Projet, "", "Id", "1;2", "Projet_L" & L, "") & "</td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Programme_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Désignation_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Véhicule_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Realignement_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Demandeur_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Atelier_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='DatePrevue_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
              L = L + 1
            RetourVehicule = RetourVehicule & "<tr>" & vbCrLf
            RetourVehicule = RetourVehicule & "<td>"
            RetourVehicule = RetourVehicule & "<input name='Id_L" & L & "' type='hidden' value=''/>"
            RetourVehicule = RetourVehicule & "NEW</td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='NumEncelade_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
'             RetourVehicule = RetourVehicule & "<td><input name='Projet_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td>" & CreatListe(rs_Projet, "", "Id", "1;2", "Projet_L" & L, "") & "</td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Programme_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Désignation_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Véhicule_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Realignement_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Demandeur_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Atelier_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='DatePrevue_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
               L = L + 1
            RetourVehicule = RetourVehicule & "<tr>" & vbCrLf
            RetourVehicule = RetourVehicule & "<td>"
            RetourVehicule = RetourVehicule & "<input name='Id_L" & L & "' type='hidden' value=''/>"
            RetourVehicule = RetourVehicule & "NEW</td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='NumEncelade_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
'             RetourVehicule = RetourVehicule & "<td><input name='Projet_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td>" & CreatListe(rs_Projet, "", "Id", "1;2", "Projet_L" & L, "") & "</td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Programme_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Désignation_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Véhicule_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Realignement_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Demandeur_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Atelier_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='DatePrevue_L" & L & "' type='text' class='Champ' value=''/></td>" & vbCrLf
            RetourVehicule = RetourVehicule & "</tr>" & vbCrLf
            End If
         If Session("Supprimer") <> "True" Then
        Supprimerdisabled = " disabled "
     Else
         Supprimerdisabled = ""
     End If
            While rs_Vehicule.EOF = False
            L = L + 1
            RetourVehicule = RetourVehicule & "<tr>" & vbCrLf
            RetourVehicule = RetourVehicule & "<td>"
            RetourVehicule = RetourVehicule & "<input name='Id_L" & L & "' type='hidden' value='" & rs_Vehicule("id") & "'>" & vbCrLf
            RetourVehicule = RetourVehicule & "<input type='checkbox' name='Delet_L" & L & "'  value='Delete' " & Supprimerdisabled & "></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='NumEncelade_L" & L & "' type='text' class='Champ' value='" & rs_Vehicule("NumEncelade") & "'/></td>" & vbCrLf
'             RetourVehicule = RetourVehicule & "<td><input name='Projet_L" & L & "' type='text' class='Champ' value='" & rs_Vehicule("d_Projects_NumAffaire") & "'/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td>" & CreatListe(rs_Projet, rs_Vehicule("id_Projects_NumAffaire"), "Id", "1;2", "Projet_L" & L, "") & "</td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Programme_L" & L & "' type='text' class='Champ' value='" & rs_Vehicule("Programme") & "'/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Désignation_L" & L & "' type='text' class='Champ' value='" & rs_Vehicule("Désignation") & "'/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Véhicule_L" & L & "' type='text' class='Champ' value='" & rs_Vehicule("Véhicule") & "'/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Realignement_L" & L & "' type='text' class='Champ' value='" & rs_Vehicule("Realignement") & "'/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Demandeur_L" & L & "' type='text' class='Champ' value='" & rs_Vehicule("Demandeur") & "'/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='Atelier_L" & L & "' type='text' class='Champ' value='" & rs_Vehicule("Atelier") & "'/></td>" & vbCrLf
             RetourVehicule = RetourVehicule & "<td><input name='DatePrevue_L" & L & "' type='text' class='Champ' value='" & rs_Vehicule("DatePrevue") & "'/></td>" & vbCrLf
            
            RetourVehicule = RetourVehicule & "</tr>" & vbCrLf
                rs_Vehicule.MoveNext
            Wend
            
           
            RetourVehicule = RetourVehicule & "</td></tr></table>" & vbCrLf
            RetourVehicule = RetourVehicule & "</table>" & vbCrLf
        
   
    RetourVehicule = RetourVehicule & "<input name='NbLigne' type='hidden' value=' " & L & "'/>"
    RetourVehicule = RetourVehicule & "</form >" & vbCrLf
      
    

     

End Function




Function OpenDb()
     On Error Resume Next
    NotErreur = False
    Set OpenDb = SERVER.CreateObject("ADODB.Connection")
    OpenDb.Mode = 16
      OpenDb.Open ConexString
   
    If Err = 0 Then
        NotErreur = True
    End If
    Err.Clear
 On Error GoTo 0
End Function










Function noquote(s0, MyFormat)

    noquote = safeEntry("" & s0)
    If Trim("" & noquote) = "" Then
        noquote = "Null"
        If "num" = MyFormat Then
            noquote = "0"
        End If
    Else
        Select Case MyFormat
                Case "txt"
                    noquote = "'" & noquote & "'"
                    
                Case "date"
               
                    If Session("ConexType") = 3 Or Session("ConexType") = 4 Then
                        noquote = "#" & Format(s0, "yyyy-mm-dd hh:mm:ss") & "#"
                    Else
                    
                    noquote = "CONVERT(DATETIME, '" & Format(s0, "yyyy-mm-dd hh:mm:ss") & "', 102)"
                    End If
                  ' Response.Write noquote
                  'Response.End
                    
                Case "num"
                If IsNumeric(Replace(Trim("" & s0), ",", ".")) Then
                    noquote = Replace(CDbl(Replace(Trim("" & s0), ",", ".")), ",", ".")
                Else
                    If IsNumeric(Replace(Trim("" & s0), ".", ",")) Then
                        noquote = Replace(CDbl(Replace(Trim("" & s0), ".", ",")), ",", ".")
                    Else
                        noquote = "0"
                    End If
                    
                End If
               
                   
        End Select
    End If
  
End Function
 Function safeEntry(strField)

    strSafe = Trim(strField)
    strSafe = funReplace(strSafe, "'", "´")
    strSafe = funReplace(strSafe, "<", "&lt;")
    strSafe = funReplace(strSafe, ">", "&gt;")
    safeEntry = strSafe
End Function
 Function funReplace(a, b, C)
    funReplace = Replace(a, b, C)
End Function


Function txtGetDefault(fld, def)
def = safeEntry(def)
   
    Set RS100 = OpenRecordSet("SELECT * FROM Defaults WHERE defName = '" & fld & "'", con)
    If Not RS100.EOF Then
           txtGetDefault = Trim(RS100("defValue"))
    Else

        ExecuteSql "INSERT INTO Defaults(defName,defValue) VALUES('" & fld & "','" & def & "')", con
        txtGetDefault = def
    End If
    
   Set RS100 = CloseRecordSet(RS100)
End Function

Function txtGetMessage(fld, def)
def = safeEntry(def)
    Set RS100 = OpenRecordSet("SELECT * FROM T_Messages WHERE NameMessage = '" & fld & "'", con)
    If Not RS100.EOF Then
           txtGetMessage = RS100("LibMessage")
    Else

        ExecuteSql "INSERT INTO T_Messages(NameMessage,LibMessage) VALUES('" & fld & "','" & def & "')", con
        txtGetMessage = def
    End If
   Set RS100 = CloseRecordSet(RS100)
End Function

Public Sub Main()

ConexString = Session("ConexString")

Response.Write "<script language='javascript' src='MyJava.js'></script>"
Set con = OpenDb

If Session("NumJob") = "" Then Session("NumJob") = NumeroChrono("NumJob", "NumJob_")
    
   
        Dim Logistique_Page
        Sql = "SELECT T_User.* FROM T_User WHERE T_User.Id=" & Session("Id_User") & " ;"
        'Response.Write sql
        'Response.End
Set Rs = OpenRecordSet(Sql, con)
If Rs.EOF = False Then
     Session("Id_User") = Rs("Id")
     Session("User") = Rs("Identifiant")
    Session("Désination FE") = Replace("" & Rs("Désination FE"), "Vrai", "True")
    Session("Type") = Replace("" & Rs("Type"), "Vrai", "True")
    Session("Pièce") = Replace("" & Rs("Pièce"), "Vrai", "True")
    Session("Indice") = Replace("" & Rs("Indice"), "Vrai", "True")
    Session("Commantaire") = Replace("" & Rs("Commantaires"), "Vrai", "True")
    Session("Designation_Piece") = Replace("" & Rs("Désignation Pièce"), "Vrai", "True")
    Session("Famille/FE/V") = Replace("" & Rs("Famille/FE/V"), "Vrai", "True")
    Session("Appro_S") = Replace("" & Rs("Appro_S"), "Vrai", "True")
    Session("Quantité") = Replace("" & Rs("Quantité"), "Vrai", "True")
    Session("Bon_sortie") = Replace("" & Rs("Bon_sortie"), "Vrai", "True")
    Session("Observation_ magasin") = Replace("" & Rs("Observation_ magasin"), "Vrai", "True")
    Session("Appro") = Replace("" & Rs("Appro"), "Vrai", "True")
    Session("Masse") = Replace("" & Rs("Masse"), "Vrai", "True")
    Session("Commandée_le") = Replace("" & Rs("Commandée_le"), "Vrai", "True")
    Session("Collectée_le") = Replace("" & Rs("Collectée_le"), "Vrai", "True")
    Session("Expédiée_le") = Replace("" & Rs("Expédiée_le"), "Vrai", "True")
    Session("Reçue_le") = Replace("" & Rs("Reçue_le"), "Vrai", "True")
    Session("Qt_reçue") = Replace("" & Rs("Qt_reçue"), "Vrai", "True")
    Session("Reste_recevoir") = Replace("" & Rs("Reste_recevoir"), "Vrai", "True")
    Session("N°_ BL") = Replace("" & Rs("N°_ BL"), "Vrai", "True")
    Session("Réaffectée") = Replace("" & Rs("Réaffectée"), "Vrai", "True")
    Session("Qt Réaffectée") = Replace("" & Rs("Qt Réaffectée"), "Vrai", "True")
    Session("provenance") = Replace("" & Rs("provenance"), "Vrai", "True")
    Session("Retournée") = Replace("" & Rs("Retournée"), "Vrai", "True")
    Session("QT_Retournée") = Replace("" & Rs("QT_Retournée"), "Vrai", "True")
    Session("N°_retour") = Replace("" & Rs("N°_retour"), "Vrai", "True")
    Session("Modifiée") = Replace("" & Rs("Modifiée"), "Vrai", "True")
    
    Session("Admin") = Replace("" & Rs("Admin"), "Vrai", "True")
    Session("Sous_Ensseble") = Replace("" & Rs("Sous_Ensseble"), "Vrai", "True")
    Session("Supprimer") = Replace("" & Rs("Supprimer"), "Vrai", "True")
     Session("Nouveau") = Replace("" & Rs("Nouveau"), "Vrai", "True")
     Session("Updatable") = Replace("" & Rs("Updatable"), "Vrai", "True")
     Session("Quantité") = Replace("" & Rs("Quantité"), "Vrai", "True")
     Session("Action") = Replace("" & Rs("Action"), "Vrai", "True")
     Session("Activer") = Replace("" & Rs("Activer"), "Vrai", "True")
     
End If
 If Session("Admin") = "True" Then
 
    Session("Désination FE") = "True"
    Session("Type") = "True"
    Session("Pièce") = "True"
    Session("Indice") = "True"
    Session("Commantaire") = "True"
    Session("Designation_Piece") = "True"
    Session("Famille/FE/V") = "True"
    Session("Appro_S") = "True"
    Session("Quantité") = "True"
    Session("Bon_sortie") = "True"
    Session("Observation_ magasin") = "True"
    Session("Appro") = "True"
    Session("Masse") = "True"
    Session("Commandée_le") = "True"
    Session("Collectée_le") = "True"
    Session("Expédiée_le") = "True"
    Session("Reçue_le") = "True"
    Session("Qt_reçue") = "True"
    Session("Reste_recevoir") = "True"
    Session("N°_ BL") = "True"
    Session("Réaffectée") = "True"
    Session("Qt Réaffectée") = "True"
    Session("provenance") = "True"
    Session("Retournée") = "True"
    Session("QT_Retournée") = "True"
    Session("N°_retour") = "True"
    Session("Modifiée") = "True"
    Session("Admin") = "True"
    Session("Sous_Ensseble") = "True"
    Session("Supprimer") = "True"
     Session("Nouveau") = "True"
     Session("Updatable") = "True"
     Session("Quantité") = "True"
     Session("Action") = "True"
     Session("Activer") = "True"
 End If
Rs.Close
Set Rs = Nothing
      
         Select Case Trim("" & Request("Logistique_Page"))
                Case "1"
                     Response.Write Fun1
                     
                Case "2"
                     Response.Write Fun2
                Case "3"
                      Response.Write Fun3
                 Case "4"
                       Response.Write Fun4
                 Case "5"
                       Response.Write Fun5
                  Case "6"
                       Response.Write Fun6
                  Case "7"
                       Response.Write Fun7
                  Case "8"
                       Response.Write Fun8
                  Case "9"
                     Response.Write RecapVehicule("T_Vehicule")
                  Case "10"
                     Response.Write RecapSousEnssemble("T_SousEnsembles")
                 Case "11"
                       Response.Write modSEARCH
                 Case "12"
                       Response.Write RecapBR("BR")
                 Case "13"
                       Response.Write RecapBL("BL")
                  Case "14"
                       Response.Write Fun14
                  Case "15"
                       Response.Write Fun15
                Case "16"
                       Response.Write RecapNewProget("LstProget")
                Case "17"
                       Response.Write SelectNewProget("LstProget")
                Case "18"
                       Response.Write Fun18
                Case "19"
                       Response.Write Fun19
                Case "20"
                       Response.Write SelectNewProgetVehicule("T_Vehicule")
                Case 21
                        SelectNewProgetMAJ ("T_Vehicule")
                Case 22
                        Response.Write RecapProjet("T_Projects")
                 Case 23
                        Response.Write GenererBlBr
                Case Else
                    
                   Response.Write modSEARCH
   End Select
    con.Close
   Set con = Nothing
  Response.Write "</body>"
 Response.Write "</html>"

End Sub
Function RetournRolbak(FormName, NbChamps)

Select Case Trim("" & FormName)
Case "T_Stock"
         RetournRolbak = RetournRolbak & "<tr >" & vbCrLf
         RetournRolbak = RetournRolbak & "<td class='TitreTableau' colspan ='" & NbChamps + 1 & "'>Liste des Stocks :</td></tr><tr>" & vbCrLf
         If Session("Updatable") = "True" Then
        RetournRolbak = RetournRolbak & "<TD width='2' valign='top' >" & AfficheBounton("Valider", 1, "'javascript:valider();'") & "</td>" & vbCrLf
        RetournRolbak = RetournRolbak & "<TD colspan='" & NbChamps & "'>&nbsp;</td>" & vbCrLf
        
     
       End If
         RetournRolbak = RetournRolbak & "</tr>" & vbCrLf
        Exit Function
Case "T_Actions"
         RetournRolbak = RetournRolbak & "<tr >" & vbCrLf
         RetournRolbak = RetournRolbak & "<td class='TitreTableau' colspan ='" & NbChamps + 1 & "'>Liste des Actions :</td></tr><tr>" & vbCrLf
         If Session("Updatable") = "True" Then
        RetournRolbak = RetournRolbak & "<TD width='2' valign='top' >" & AfficheBounton("Valider", 1, "'javascript:valider();'") & "</td>" & vbCrLf
        RetournRolbak = RetournRolbak & "<TD colspan='" & NbChamps & "'>&nbsp;</td>" & vbCrLf
        
     
       End If
         RetournRolbak = RetournRolbak & "</tr>" & vbCrLf
        Exit Function
Case "T_Delais"
         RetournRolbak = RetournRolbak & "<tr >" & vbCrLf
         RetournRolbak = RetournRolbak & "<td class='TitreTableau' colspan ='" & NbChamps + 1 & "'>Liste des Delais :</td></tr><tr>" & vbCrLf
         If Session("Updatable") = "True" Then
        RetournRolbak = RetournRolbak & "<TD width='2' valign='top' >" & AfficheBounton("Valider", 1, "'javascript:valider();'") & "</td>" & vbCrLf
        RetournRolbak = RetournRolbak & "<TD colspan='" & NbChamps & "'>&nbsp;</td>" & vbCrLf
        
     
       End If
         RetournRolbak = RetournRolbak & "</tr>" & vbCrLf
        Exit Function
        
Case "T_User"
         RetournRolbak = RetournRolbak & "<tr >" & vbCrLf
         RetournRolbak = RetournRolbak & "<td class='TitreTableau' colspan ='" & NbChamps + 1 & "'>Liste des Uesers :</td></tr><tr>" & vbCrLf
          If Session("Updatable") = "True" Then
        RetournRolbak = RetournRolbak & "<TD width='2' valign='top'>" & AfficheBounton("Valider", 1, "'javascript:valider();'") & "</td>" & vbCrLf
        RetournRolbak = RetournRolbak & "<TD colspan='" & NbChamps & "'>&nbsp;</td></tr>" & vbCrLf
        
     
       End If
         RetournRolbak = RetournRolbak & "</tr>" & vbCrLf
        Exit Function
        
Case "T_Projects"
     RetournRolbak = RetournRolbak & vbCrLf & "<SCRIPT langage='JAVASCRIPT'>"
        
    RetournRolbak = RetournRolbak & "function ModifierRechereche(){" & vbCrLf
    RetournRolbak = RetournRolbak & "this.document." & FormName & ".ID__SousEnsembles.value='';"
    RetournRolbak = RetournRolbak & "this.document." & FormName & ".action='Default.asp?Logistique_Page=11'" & vbCrLf
    RetournRolbak = RetournRolbak & vbCrLf & "document." & FormName & ".submit();"
    RetournRolbak = RetournRolbak & "}" & vbCrLf
'
       RetournRolbak = RetournRolbak & "function valider3(){" & vbCrLf
RetournRolbak = RetournRolbak & "this.document." & FormName & ".Id_Projects_NumAffaire.value='';"
RetournRolbak = RetournRolbak & "this.document." & FormName & ".action='Default.asp?Logistique_Page=22'" & vbCrLf
RetournRolbak = RetournRolbak & vbCrLf & "document." & FormName & ".submit();"
RetournRolbak = RetournRolbak & "}" & vbCrLf

        
        RetournRolbak = RetournRolbak & vbCrLf & "</SCRIPT >"
         RetournRolbak = RetournRolbak & "<tr >" & vbCrLf
         RetournRolbak = RetournRolbak & "<td class='TitreTableau' colspan ='" & NbChamps + 1 & "'>Liste des Projet :</td></tr><tr>" & vbCrLf
          aaa = 0
          RetournRolbak = RetournRolbak & "<TD TD colspan='" & NbChamps + 1 & "'>"
          RetournRolbak = RetournRolbak & "<table width='100%'><tr>"
          If Session("Updatable") = "True" Then
              
              RetournRolbak = RetournRolbak & "<td width='2%' valign='top'> " & AfficheBounton("Valider", 1, "'javascript:valider();'") & "</td>" & vbCrLf
       
            aaa = 1
     
        End If
       RetournRolbak = RetournRolbak & "<td width='2%' valign='top'>" & AfficheBounton("Retour", 1, "'javascript:valider3();'") & "</td><TD >&nbsp;</td></tr></table>" & vbCrLf
        
        RetournRolbak = RetournRolbak & "</tr>" & vbCrLf
        
     
       
        
        Exit Function
Case "T_Societe"
         RetournRolbak = RetournRolbak & "<tr >" & vbCrLf
         RetournRolbak = RetournRolbak & "<td class='TitreTableau' colspan ='" & NbChamps + 1 & "'>Liste des sociétées :</td></tr><tr>" & vbCrLf
          If Session("Updatable") = "True" Then
        RetournRolbak = RetournRolbak & "<TD width='2' valign='top'>" & AfficheBounton("Valider", 1, "'javascript:valider();'") & "</td>" & vbCrLf
        RetournRolbak = RetournRolbak & "<TD colspan='" & NbChamps & "'>&nbsp;</td></tr>" & vbCrLf
        
     
       End If
         RetournRolbak = RetournRolbak & "</tr>" & vbCrLf
        Exit Function
 Case "T_SousEnsembles"
        RetournRolbak = RetournRolbak & "<tr >" & vbCrLf
         RetournRolbak = RetournRolbak & "<td class='TitreTableau' colspan ='" & NbChamps + 1 & "'>Liste des S/Enssembles :</td>" & vbCrLf
         RetournRolbak = RetournRolbak & "</tr>" & vbCrLf
          RetournRolbak = RetournRolbak & "<tr>" & vbCrLf
          RetournRolbak = RetournRolbak & "<TD width='100%' colspan ='" & NbChamps + 1 & "'><table border='0'  width='100%'><tr>"
         
        RetournRolbak = RetournRolbak & vbCrLf & "<SCRIPT langage='JAVASCRIPT'>"
        
    RetournRolbak = RetournRolbak & "function ModifierRechereche(){" & vbCrLf
    RetournRolbak = RetournRolbak & "this.document." & FormName & ".ID__SousEnsembles.value='';"
    RetournRolbak = RetournRolbak & "this.document." & FormName & ".action='Default.asp?Logistique_Page=11'" & vbCrLf
    RetournRolbak = RetournRolbak & vbCrLf & "document." & FormName & ".submit();"
    RetournRolbak = RetournRolbak & "}" & vbCrLf
    
       RetournRolbak = RetournRolbak & "function valider3(){" & vbCrLf
RetournRolbak = RetournRolbak & "this.document." & FormName & ".ID__SousEnsembles.value='';"
RetournRolbak = RetournRolbak & "this.document." & FormName & ".action='Default.asp?Logistique_Page=10'" & vbCrLf
RetournRolbak = RetournRolbak & vbCrLf & "document." & FormName & ".submit();"
RetournRolbak = RetournRolbak & "}" & vbCrLf

        RetournRolbak = RetournRolbak & "function valider2(){" & vbCrLf
        RetournRolbak = RetournRolbak & "this.document." & FormName & ".action='Default.asp?Logistique_Page=5'" & vbCrLf
        RetournRolbak = RetournRolbak & vbCrLf & "document." & FormName & ".submit();"
        RetournRolbak = RetournRolbak & "}" & vbCrLf
        RetournRolbak = RetournRolbak & vbCrLf & "</SCRIPT >"
 If Session("Updatable") = "True" Then
        RetournRolbak = RetournRolbak & "<TD width='2' valign='top'>" & AfficheBounton("Valider", 1, "'javascript:valider();'") & "</td>" & vbCrLf
       
        
     
       End If
       RetournRolbak = RetournRolbak & "<TD width='2' valign='top'>" & AfficheBounton("Retour", 8, "'javascript:valider3();'") & "</td>" & vbCrLf
      RetournRolbak = RetournRolbak & "<TD>&nbsp;</td><TD>&nbsp;</td>" & vbCrLf
        RetournRolbak = RetournRolbak & "<TD width='4%' valign='top'>Filtrés sur véhicule :" & vbCrLf
         Sql = "SELECT     t_Vehicule.Id, T_Projects.NumAffaire, t_Vehicule.Véhicule "
        Sql = Sql & "FROM t_Vehicule INNER JOIN "
        Sql = Sql & "T_Projects ON t_Vehicule.Id_Projects_NumAffaire = T_Projects.Id "
        Sql = Sql & "ORDER BY t_Vehicule.Véhicule"
        Set RsFiltre = OpenRecordSet(Sql, con)

       RetournRolbak = RetournRolbak & "<br>" & CreatListe(RsFiltre, Request(FormName & "FiltreSouEnsemble"), "Id", "1;2", FormName & "FiltreSouEnsemble", " onChange='valider2();' ") & "</td>" & vbCrLf
      RsFiltre.Close
      Set RsFiltre = Nothing
      
     
     RetournRolbak = RetournRolbak & "</Tr></table></td>"
 
 
 End Select
        
End Function
Function RetourneTypeChamp(valeur, TypeChamp)
Debug.Print TypeChamp.Type
Select Case TypeChamp.Type
    'Texte : 10 202 Mémo : 12
    Case 10
        RetourneTypeChamp = "txt"
   Case 12
        RetourneTypeChamp = "txt"
   
    Case 202
        RetourneTypeChamp = "txt"
   Case 203
        RetourneTypeChamp = "txt"

    'Numérique :  4 or Monétaire : 5
    Case (4 Or 5)
        RetourneTypeChamp = "num"
    'Date/Heure : 8
    
    Case 3
        RetourneTypeChamp = "num"
    Case 8
        RetourneTypeChamp = "date"
   
  
    'Oui/Non : 1
    Case (1 Or 11)
        RetourneTypeChamp = "num"
        If Left(UCase("" & valeur), 1) = "T" Then
            valeur = -1
        Else
                valeur = 0
        End If
       
    
End Select
End Function
Function RetourneTypeControle(valeur, TypeChamp, Name)
'Response.Write TypeChamp.type & "<br>"

Dim Checked
Select Case TypeChamp.Type

    'Texte : 10 202
    Case 10
        RetourneTypeControle = "<input type='text' class='Champ' name='" & Name & "'  value='" & valeur & "'>"
       Case 202
        RetourneTypeControle = "<input type='text' class='Champ' name='" & Name & "'  value='" & valeur & "'>"
    'Mémo : 12
    Case 12
         RetourneTypeControle = "<textarea name='" & Name & "' rows='2' cols='20'>" & valeur & "</textarea>"
     Case 203
         RetourneTypeControle = "<textarea name='" & Name & "' rows='2' cols='20'>" & valeur & "</textarea>"
    'Numérique : 4
    Case 4
        RetourneTypeControle = "<input type='text' class='Champ' name='" & Name & "'  value='" & valeur & "'>"
    'Date/Heure : 8
    Case 8
        RetourneTypeControle = "<input type='text' class='Champ' name='" & Name & "'  value='" & valeur & "'>"
    'Monétaire : 5
    Case 5
         RetourneTypeControle = "<input type='text' class='Champ' name='" & Name & "'  value='" & valeur & "'>"
    'NuméroAuto : 4
    Case 4
        RetourneTypeControle = "<input type='hidden' name='" & Name & "'  value='" & valeur & "'>"
    'Oui/Non : 1
    Case (1 Or 11)
        
        If Left(UCase("" & valeur), 1) = "F" Then
            Checked = ""
        Else
            If Trim("" & valeur) <> "" Then
                Checked = "  Checked "
            Else
                Checked = ""
            End If
        End If
         RetourneTypeControle = "<input type='Checkbox' name='" & Name & "'  value='True' " & Checked & ">"
    'Objet OLE : 11
    Case 11
    'Lien hypertexte : 12
    Case 12
         RetourneTypeControle = "<input type='text' class='Champ' name='" & Name & "'  value='" & valeur & "'>"
    Case 3
            If InStr(1, TypeChamp.Name, "_") <> 0 Then
                SplitChamp = Split(TypeChamp.Name & "_", "_")
                
                Sql = "SELECT T_" & SplitChamp(1) & ".Id, T_" & SplitChamp(1) & "." & SplitChamp(2) & " "
                Sql = Sql & "FROM T_" & SplitChamp(1) & " GROUP BY T_" & SplitChamp(1) & ".Id, T_" & SplitChamp(1) & "." & SplitChamp(2) & ";"
               
                Set Rs = OpenRecordSet(Sql, con)
               
              
                RetourneTypeControle = CreatListe(Rs, valeur, "Id", "1", Name, "") & vbCrLf
                Rs.Close
                Set Rs = Nothing
            Else
                RetourneTypeControle = "<input type='text' class='Champ' name='" & Name & "'  value='" & valeur & "'>"
            End If
End Select

End Function
Function RetourneCloseWhere()
RetourneCloseWhere = ""
If Val(Trim("" & Request("id_Vehicule"))) <> 0 Then
'Response.Write   Request("id_Vehicule")
'Response.End
    RetourneCloseWhere = RetourneCloseWhere & "T_SousEnsembles.ID_Vehicule_Véhicule=" & Request("id_Vehicule") & " AND "

End If
If Trim("" & Request("Pièce")) <> "" Then
    RetourneCloseWhere = RetourneCloseWhere & "T_Nomenclature.Pièce=" & noquote(Request("Pièce"), "txt") & " AND "
End If

If Val(Trim("" & Request("id"))) <> 0 Then
    RetourneCloseWhere = RetourneCloseWhere & "T_Nomenclature.ID=" & Request("id") & " AND "
End If
If Val(Trim("" & Request("ID__SousEnsembles"))) <> 0 Then
    RetourneCloseWhere = RetourneCloseWhere & "T_Nomenclature.[ID_SousEnsembles_Sous-Ensseble]=" & Request("ID__SousEnsembles") & " AND "
End If
If Val(Trim("" & Request("Id_Projects_NumAffaire"))) <> 0 Then
    RetourneCloseWhere = RetourneCloseWhere & "t_Vehicule.[Id_Projects_NumAffaire]=" & Request("Id_Projects_NumAffaire") & " AND "
End If



If Trim("" & RetourneCloseWhere) <> "" Then
        RetourneCloseWhere = Left(RetourneCloseWhere, Len(RetourneCloseWhere) - 4)
        RetourneCloseWhere = RetourneCloseWhere
End If

End Function
Function JavaValider(FormName, SelectCase)
If Request("id") <> "" Then
MyId = "&ID=" & Request("id")
Else
MyId = ""
End If
    JavaValider = "<script language=javascript>" & vbCrLf
    JavaValider = JavaValider & "function valider(){" & vbCrLf
   
   JavaValider = JavaValider & "this.document." & FormName & ".action='Default.asp?Logistique_Page=" & SelectCase & MyId & "'" & vbCrLf
  JavaValider = JavaValider & "this.document." & FormName & ".submit();" & vbCrLf
JavaValider = JavaValider & "}" & vbCrLf
JavaValider = JavaValider & "</script >" & vbCrLf
End Function
Function AfficheBounton(Titre, Id, Url)
        AfficheBounton = AfficheBounton & vbCrLf & "<table border=""0"" background='./Images/Boutoncentre.bmp' cellpadding='0' cellspacing='0' width='100%' "
    If Trim("" & Id) <> "" Then
        AfficheBounton = AfficheBounton & "id='" & Id & "'"
    End If
       AfficheBounton = AfficheBounton & ">"
    
    AfficheBounton = AfficheBounton & vbCrLf & "    <tr>"
    AfficheBounton = AfficheBounton & vbCrLf & "        <td width='5%' align='left' ><img src='./Images/BoutonLeft.bmp' border='0'></td>"
    AfficheBounton = AfficheBounton & vbCrLf & ""
    AfficheBounton = AfficheBounton & vbCrLf & "        <td align='center'><a href=" & Url & " border='0'>" & Titre & "</a></td>"
    AfficheBounton = AfficheBounton & vbCrLf & ""
    AfficheBounton = AfficheBounton & vbCrLf & "        <td width='5%' align='right'><img src='./Images/Boutonrith.bmp' border='0'></td>"
    AfficheBounton = AfficheBounton & vbCrLf & "    </tr>"
    AfficheBounton = AfficheBounton & vbCrLf & "</table>"
  ' Debug.Print
End Function

Function AfficheBountonSersh(Titre, Id, Url)

        AfficheBountonSersh = AfficheBountonSersh & vbCrLf & "<table border=""0"" background='./Images/BoutoncentreSersh.bmp' cellpadding='0' cellspacing='0' width='100%' "
    If Trim("" & Id) <> "" Then
        AfficheBountonSersh = AfficheBountonSersh & "id='" & Id & "'>"
    Else
       AfficheBountonSersh = AfficheBountonSersh & ">"
    End If
    AfficheBountonSersh = AfficheBountonSersh & vbCrLf & "    <tr>"
    AfficheBountonSersh = AfficheBountonSersh & vbCrLf & "        <td width='5%' align='left' ><img src='./Images/BoutonLeftSersh.bmp' border='0'></td>"
    AfficheBountonSersh = AfficheBountonSersh & vbCrLf & ""
    AfficheBountonSersh = AfficheBountonSersh & vbCrLf & "        <td align='center'><a href=" & Url & " border='0'>" & Titre & "</a></td>"
    AfficheBountonSersh = AfficheBountonSersh & vbCrLf & ""
    AfficheBountonSersh = AfficheBountonSersh & vbCrLf & "        <td width='5%' align='right'><img src='./Images/BoutonrithSersh.bmp' border='0'></td>"
    AfficheBountonSersh = AfficheBountonSersh & vbCrLf & "    </tr>"
    AfficheBountonSersh = AfficheBountonSersh & vbCrLf & "</table>"
  ' Debug.Print
End Function

'**************************************************************
'permet de créer un tableau avec les information de mise à jour
 Function RetourSqlTableau(ChaineSQL, FormName, champMasquer, ChDefautValeur, ValeurDefaule, Updatable, Url, Add)
 Dim Rs
 'Response.Write ChaineSQL
 'Response.End
     RetourSqlTableau = "<form name='" & FormName & "' method='post'>" & vbCrLf
     RetourSqlTableau = RetourSqlTableau & "<input name='MAJ' type='hidden' value='OK'/>" & vbCrLf
     RetourSqlTableau = RetourSqlTableau & "<input name='Updatable' type='hidden' value='OK'/>" & vbCrLf
      RetourSqlTableau = RetourSqlTableau & "<input name='FormName' type='hidden' value='" & FormName & "'/>" & vbCrLf
       RetourSqlTableau = RetourSqlTableau & "<input name='id_Vehicule' type='hidden' value='" & Request("id_Vehicule") & "' />"
       RetourSqlTableau = RetourSqlTableau & "<input name='Pièce' type='hidden' value='" & Request("Pièce") & "' />"
       RetourSqlTableau = RetourSqlTableau & "<input name='ID__SousEnsembles' type='hidden' value='" & Request("ID__SousEnsembles") & "' />"
        RetourSqlTableau = RetourSqlTableau & "<input name='Id_Projects_NumAffaire' type='hidden' value='" & Request("Id_Projects_NumAffaire") & "' />"
  Dim L
     Dim C
      Set Rs = OpenRecordSet(ChaineSQL, con)
     
      NbChamps = Rs.Fields.Count - 1
      RetourSqlTableau = RetourSqlTableau & "<input name='NbCol' type='hidden' value='" & Rs.Fields.Count - 1 & "'/>" & vbCrLf
      RetourSqlTableau = RetourSqlTableau & "<table border='1'width='100%' >" & vbCrLf
       If Request("nbSelect") > 0 Then
            RetourSqlTableau = RetourSqlTableau & "<tr>" & vbCrLf
            RetourSqlTableau = RetourSqlTableau & "         <td colspan='" & Rs.Fields.Count & "' >" & RecapSersh & "</td>" & vbCrLf
            
            RetourSqlTableau = RetourSqlTableau & "</tr>" & vbCrLf
         End If
        
         RetourSqlTableau = RetourSqlTableau & "<tr>" & vbCrLf
         RetourSqlTableau = RetourSqlTableau & RetournRolbak(FormName, NbChamps)
    
     
     RetourSqlTableau = RetourSqlTableau & "<tr>"
     
    For C = 0 To Rs.Fields.Count - 1
    
        If UCase(champMasquer) <> UCase(Rs.Fields(C).Name) Then
            'if ucase(ChDefautValeur)<>ucase(rs.fields(C).name) then
                RetourSqlTableau = RetourSqlTableau & Rs.Fields(C).Name
            'end if
        Else
            RetourSqlTableau = RetourSqlTableau & "<td align='center' bgcolor='#CCCCCC'>Supprimer</td>" & vbCrLf
        End If
         RetourSqlTableau = RetourSqlTableau & "</td>" & vbCrLf
         RetourSqlTableau = RetourSqlTableau & "<td align='center' bgcolor='#CCCCCC'><input name='Champs" & C & "' type='hidden' value='" & Rs.Fields(C).Name & "'/>"
     
     Next
     If Trim("" & Url) <> "" Then
     RetourSqlTableau = RetourSqlTableau & "<td>Lien</td>"
     End If
     RetourSqlTableau = RetourSqlTableau & "</tr>"
     
     
       L = 0
    Dim boolPasfaire
     C = 0
     
     
     If Session("Nouveau") = "True" Then
      RetourSqlTableau = RetourSqlTableau & "<tr>"
      For L = 0 To 9
       
         'for C=0 to  rs.fields.Count - 1
         
         For C = 0 To Rs.Fields.Count - 1
         ChampName = "" & Rs.Fields(C).Name
    'boolPasfaire =false
      If champMasquer = ChampName Then
            RetourSqlTableau = RetourSqlTableau & "<td align='center' bgcolor='#FFFFFF'>" & vbCrLf
            RetourSqlTableau = RetourSqlTableau & "    New" & vbCrLf
           
            RetourSqlTableau = RetourSqlTableau & "</td>" & vbCrLf
      Else
            If ChDefautValeur = ChampName Then
                RetourSqlTableau = RetourSqlTableau & "<td align='center' bgcolor='#FFFFFF'>" & vbCrLf
                 RetourSqlTableau = RetourSqlTableau & "    New" & vbCrLf
                RetourSqlTableau = RetourSqlTableau & "</td>" & vbCrLf
            Else
                RetourSqlTableau = RetourSqlTableau & "<td align='center' bgcolor='#FFFFFF'>" & RetourneTypeControle("", Rs.Fields(C), "Newtxt_L" & L & "_C" & C) & "</td>" & vbCrLf
            End If
      End If
       
         
     Next
         
         
         
         'chamName="" & rs.fields(C).name
         
          '  if ucase(champMasquer)<>ucase(chamName) then
          '      RetourSqlTableau=RetourSqlTableau & "<td align='center' bgcolor='#FFFFFF'>" & RetourneTypeControle("",   rs.fields(C),"Newtxt_L" & L & "_C" & C )
          '      '<input type='text' class='Champ' name='Newtxt_L" & L & "_C" & C & "'  value=''></td>" & vbcrlf
          '       RetourSqlTableau=RetourSqlTableau & "</td>" & vbcrlf
           ' else
           '     RetourSqlTableau=RetourSqlTableau & "<td align='center' bgcolor='#FFFFFF'>New</td>" & vbcrlf
           ' end if
        ' next
         
         RetourSqlTableau = RetourSqlTableau & "</tr>"
      
     Next
      RetourSqlTableau = RetourSqlTableau & "</tr>"
     End If
    
     
     
     
     
     
     
          If Session("Supprimer") <> "True" Then
        Supprimerdisabled = " disabled "
     Else
         Supprimerdisabled = ""
     End If
     While Rs.EOF = False
        RetourSqlTableau = RetourSqlTableau & "<tr>"

        For C = 0 To Rs.Fields.Count - 1
          'Response.Write trim("" & rs.fields(C).name) & "<br/>"
              If UCase("" & champMasquer) = UCase("" & Rs.Fields(C).Name) Then
                    RetourSqlTableau = RetourSqlTableau & "<td align='center' bgcolor='#FFFFFF'>" & vbCrLf
                    RetourSqlTableau = RetourSqlTableau & "    <input type='checkbox' name='Delet_L" & L & "'  value='Del' " & Supprimerdisabled & ">" & vbCrLf
                    RetourSqlTableau = RetourSqlTableau & "    <input type='hidden' name='DI_L" & L & "' value='" & Rs.Fields(C).Value & "'>" & vbCrLf
                    RetourSqlTableau = RetourSqlTableau & "</td>" & vbCrLf
              Else
                If UCase("" & ChDefautValeur) = UCase(Rs.Fields(C).Name) Then
                    RetourSqlTableau = RetourSqlTableau & "<td align='center' bgcolor='#FFFFFF'>" & vbCrLf
                    RetourSqlTableau = RetourSqlTableau & "    <input type='checkbox' name='Delet_L" & L & "'  value='Del' " & Supprimerdisabled & ">" & vbCrLf
                    RetourSqlTableau = RetourSqlTableau & "    <input type='hidden' name='DI_L" & L & "' value='" & Rs.Fields(C).Value & "'>" & vbCrLf
                    RetourSqlTableau = RetourSqlTableau & "</td>" & vbCrLf
                Else
                    RetourSqlTableau = RetourSqlTableau & "<td align='center' bgcolor='#FFFFFF'>" & RetourneTypeControle("" & Rs.Fields(C).Value, Rs.Fields(C), "txt_L" & L & "_C" & C) & "</td>" & vbCrLf
                End If
              End If
              
            
            
         
        Next
        If Trim("" & Url) <> "" Then
        RetourSqlTableau = RetourSqlTableau & "<td><a href='" & Url & Rs.Fields(champMasquer) & "'>Cliquez ici</a></td>"
        End If
        RetourSqlTableau = RetourSqlTableau & "</tr>"
        L = L + 1
        Rs.MoveNext
     Wend
    
     RetourSqlTableau = RetourSqlTableau & "<input name='NbLigne' type='hidden' value='" & L - 1 & "'/>"
     
     
      RetourSqlTableau = RetourSqlTableau & "</table>" & vbCrLf
   
        
   
    If Trim("" & Add) <> "" Then
     RetourSqlTableau = RetourSqlTableau & "<input name='Add' type='hidden' value='NoN' />"
End If
    RetourSqlTableau = RetourSqlTableau & "</form >" & vbCrLf
End Function

Function CreatListe(Rs, Id, NameChampsId, NamChampListe, NameListe, OnChange)
Dim SplitChamp
SplitChamp = Split(NamChampListe & ";", ";")
Dim I As Long
Dim Valeurs As String
Rs.Requery
If Right(NameListe, 1) <> "'" Then NameListe = NameListe & "'"
CreatListe = CreatListe & "<select name='" & NameListe & " width='100%' " & OnChange & ">" & vbCrLf
CreatListe = CreatListe & "<option value='' width='100%' >&nbsp;</option>" & vbCrLf
While Rs.EOF = False
        For I = 0 To UBound(SplitChamp) - 1
            If I = 0 Then
                Valeurs = "" & Rs(Val(SplitChamp(I)))
            Else
                 Valeurs = Valeurs & " | " & Rs(Val(SplitChamp(I)))
            End If
        Next
      If Trim("" & Rs(NameChampsId)) = Trim("" & Id) Then
        Selected = "selected"
      Else
        Selected = ""
      End If
        CreatListe = CreatListe & "<option width='100%' " & Selected & " value='" & Rs(NameChampsId) & "' >" & Valeurs & "</option>" & vbCrLf
    Rs.MoveNext
Wend
 CreatListe = CreatListe & "</select>" & vbCrLf
End Function
Function CreatListeChrono(Rs, NameListe)
Dim SplitChamp
Dim I As Long
Dim Valeurs As String
Rs.Requery
If Trim("" & NameListe) = "" Then NameListe = "[NameListe]"
CreatListeChrono = CreatListeChrono & "<select name='" & NameListe & "' width='100%' " & OnChange & ">" & vbCrLf
CreatListeChrono = CreatListeChrono & "<option value='' width='100%' >&nbsp;</option>" & vbCrLf
While Rs.EOF = False
         
            
                 Valeurs = Rs("Clé ty") & "_" & Rs("Clé ac") & "_" & Rs("Année") & "_" & Rs("Clé Ch") & "_" & Rs("Rév")
            
        
      
        CreatListeChrono = CreatListeChrono & "<option width='100%' " & Selected & " value='" & Valeurs & "' >" & Valeurs & "</option>" & vbCrLf
    Rs.MoveNext
Wend
 CreatListeChrono = CreatListeChrono & "</select>" & vbCrLf
End Function

Function T_CreatListe(Rs, Id, NameChampsId, NamChampListe, NameListe, OnChange)

Dim I
Dim SplitChamp
Dim Valeurs As String
SplitChamp = Split(NamChampListe & ";", ";")
T_CreatListe = T_CreatListe & "<select name='" & NameListe & "' width='100%' " & OnChange & ">" & vbCrLf
T_CreatListe = T_CreatListe & "<option value='' width='100%' >&nbsp;</option>" & vbCrLf
Dim L

For L = 0 To UBound(Rs, 2)

     For I = 0 To UBound(SplitChamp) - 1
            If I = 0 Then
                Valeurs = "" & Rs(Val(SplitChamp(I)), L)
            Else
                 Valeurs = Valeurs & " | " & Rs(Val(SplitChamp(I)), L)
            End If
        Next
        T_CreatListe = T_CreatListe & "<option width='100%'   value='" & Rs(NameChampsId, L) & "' >" & Valeurs & "</option>" & vbCrLf
  Next
 T_CreatListe = T_CreatListe & "</select>" & vbCrLf
End Function
 Function SousRubique(Titre)
'SousRubique = SousRubique & vbCrLf & "</table>" & vbcrlf
        SousRubique = SousRubique & vbCrLf & "<table border='0' background='./Images/bandec.gif' cellpadding='0' cellspacing='0' width='100%'>" & vbCrLf
    SousRubique = SousRubique & vbCrLf & "    <tr>" & vbCrLf
    SousRubique = SousRubique & vbCrLf & "        <td width='5%' align='left' ><img src='./Images/bandeg.gif' border='0'></td>" & vbCrLf
    SousRubique = SousRubique & vbCrLf & "        <td>" & Titre & "</td>" & vbCrLf
    SousRubique = SousRubique & vbCrLf & "        <td width='5%' align='right'><img src='./Images/banded.gif' border='0'></td>" & vbCrLf
    SousRubique = SousRubique & vbCrLf & "    </tr>" & vbCrLf
    SousRubique = SousRubique & vbCrLf & "</table>" & vbCrLf
End Function
 Function NumeroChrono(Recherche, Prefix)


Dim txtValue
Set rsnumdevis = OpenRecordSet("SELECT t_Chrono.* FROM t_Chrono where t_Chrono.type='" & Recherche & "';", con)
    If rsnumdevis.EOF = True Then
        ExecuteSql "INSERT INTO t_Chrono ( Chrono, MyAnnee,type ) values( 1 , Year(date()) ,'" & Recherche & "');", con
        NumeroChrono = 1
    Else
        If rsnumdevis("MyAnnee") = Year(Date) Then
            NumeroChrono = rsnumdevis("Chrono") + 1
        Else
            NumeroChrono = 1
        End If
        
    End If
    Sql = "UPDATE t_Chrono SET t_Chrono.Chrono =" & NumeroChrono & ", t_Chrono.MyAnnee = YEAR({ fn NOW() }) where t_Chrono.type='" & Recherche & "';"
   ' Session("MySql") =' Session("MySql") & vbCrLf & Sql & "<br><br>"
    ExecuteSql Sql, con
     rsnumdevis.Close
     Set rsnumdevis = Nothing
    NumeroChrono = Prefix & Format(Date, "yy") & "_" & NumeroChrono & "_1"
End Function
Function FomatNum(Value, NbDec)
FomatNum = Replace(Value, ",", ".")
FomatNum = Round(FomatNum, NbDec)
End Function



Function ReplaceChamp2(txt, Form)
Dim Mytext
Dim SplitTxt

ReplaceChamp2 = txt

Select Case UCase(Form)
Case UCase("HISTORIQUE_Nommenclature")
        ReplaceChamp2 = Replace(ReplaceChamp2, "BR", "N_de_BL")
        ReplaceChamp2 = Replace(ReplaceChamp2, "BL", "Num_de_retour")
         ReplaceChamp2 = Replace(ReplaceChamp2, "Sous-Ensseble", "[Sous-Ensseble]")
Case "CONFIGCLI"
   
Case "HISTORIQUE_BR"
    
    
Case "AFFICHELISTECOMMANDDUJOUR"


Case "HISTORIQUE_COMANDE_GROUPEE"

Case "HISTORIQUE_EBOUTIQUE_FACTURE"

Case "HISTORIQUE_BL"
Case "HISTORIQUE_EBOUTIQUE_COMMANDE"
Case "HISTORIQUE_CLIENT"
Case "HISTORIQUE_COMPTEDETAL"
Case "HISTORIQUE_AVENANT"
Case "VOIRPANIERDETAL"

End Select



End Function
Function RecapSersh()
RecapSersh = ""
Dim Rs
Dim Sql
Dim IndexTab
Dim long_Ubound
Dim TableauSherch()
 Dim SelecetList
 Dim boolSelecetList
IndexTab = 0
boolSelecetList = False
' On recupere le nombre de lignes
If Request("nbSelect") > 0 Then
 nbLignes = CInt(Request("nbSelect"))
Else
 nbLignes = 0
End If
Const IndexOperateurDefault = 2
Dim IndexChamps
Dim aSlt_OPERATEUR(1, 8)
aSlt_OPERATEUR(0, 0) = "<"
aSlt_OPERATEUR(0, 1) = "<="
aSlt_OPERATEUR(0, 2) = "="
aSlt_OPERATEUR(0, 3) = "=>"
aSlt_OPERATEUR(0, 4) = ">"
aSlt_OPERATEUR(0, 5) = "LIKE1"
aSlt_OPERATEUR(0, 6) = "LIKE2"
aSlt_OPERATEUR(0, 7) = "LIKE"
aSlt_OPERATEUR(0, 8) = "<>"

aSlt_OPERATEUR(1, 0) = "Inférieur à"
aSlt_OPERATEUR(1, 1) = "Inférieur ou Egal à"
aSlt_OPERATEUR(1, 2) = "Egal à"
aSlt_OPERATEUR(1, 3) = "Supérieur ou Egal à"
aSlt_OPERATEUR(1, 4) = "Supérieur à"
aSlt_OPERATEUR(1, 5) = " Commençant par"
aSlt_OPERATEUR(1, 6) = "Finissant par"
aSlt_OPERATEUR(1, 7) = "Contient"
aSlt_OPERATEUR(1, 8) = "Différent"

Dim aSlt_Logique(1, 2)
aSlt_Logique(0, 0) = "(Aucun)"
aSlt_Logique(0, 1) = "AND"
aSlt_Logique(0, 2) = "OR"

aSlt_Logique(1, 0) = "(Aucun)"
aSlt_Logique(1, 1) = "ET"
aSlt_Logique(1, 2) = "OU"








Dim aSlt2(1, 1)
aSlt2(0, 0) = 3
aSlt2(1, 0) = "Option 3"
aSlt2(0, 1) = 4
aSlt2(1, 1) = "Option 4"

Dim Champs(44, 1)

Champs(1, 0) = "T_SousEnsembles.Code"
Champs(2, 0) = "T_SousEnsembles.[Sous-Ensseble]"
Champs(3, 0) = "T_SousEnsembles.[MASSE total (KG)]"
Champs(4, 0) = "T_SousEnsembles.[Nb lignes sans masse]"
Champs(5, 0) = "T_SousEnsembles.[Nb total de lignes]"
Champs(6, 0) = "t_Vehicule.NumEncelade"
Champs(7, 0) = "t_Vehicule.Projet"
Champs(8, 0) = "t_Vehicule.Programme"
Champs(9, 0) = "t_Vehicule.Désignation"
Champs(10, 0) = "t_Vehicule.Véhicule"
Champs(11, 0) = "t_Vehicule.Realignement"
Champs(12, 0) = "t_Vehicule.Demandeur"
Champs(13, 0) = "t_Vehicule.Atelier"
Champs(14, 0) = "t_Vehicule.DatePrevue"
Champs(15, 0) = "T_Nomenclature.[Désination FE]"
Champs(16, 0) = "T_Nomenclature.Type"
Champs(17, 0) = "T_Nomenclature.Pièce"
Champs(18, 0) = "T_Nomenclature.Indice"
Champs(19, 0) = "T_Nomenclature.Commantaire"
Champs(20, 0) = "T_Nomenclature.Designation_Piece"
Champs(21, 0) = "T_Nomenclature.[Famille/FE/V]"
Champs(22, 0) = "T_Nomenclature.[Num_appro_ S]"
Champs(23, 0) = "T_Nomenclature.Quantite"
Champs(24, 0) = "T_Nomenclature.Bon_sortie"
Champs(25, 0) = "T_Nomenclature.Observation_ou_magasin"
Champs(26, 0) = "T_Nomenclature.Num_appro"
Champs(27, 0) = "T_Nomenclature.Masse"
Champs(28, 0) = "T_Nomenclature.Commandee_le"
Champs(29, 0) = "T_Nomenclature.Collectee_le"
Champs(30, 0) = "T_Nomenclature.Expedee_le"
Champs(31, 0) = "T_Nomenclature.Reçue_le"
Champs(32, 0) = "T_Nomenclature.Qt_reçue"
Champs(33, 0) = "T_Nomenclature.Reste_a_recevoir"
Champs(34, 0) = "T_Nomenclature.N_de_BL"
Champs(35, 0) = "T_Nomenclature.Reaffectee_a"
Champs(36, 0) = "T_Nomenclature.Qt_Reaffectee"
Champs(37, 0) = "T_Nomenclature.Reaffectee_Le"
Champs(38, 0) = "T_Nomenclature.En_provenance_de"
Champs(39, 0) = "T_Nomenclature.Retournee_le"
Champs(40, 0) = "T_Nomenclature.QT_Retournee"
Champs(41, 0) = "T_Nomenclature.Num_de_retour"
Champs(42, 0) = "T_Nomenclature.Modifiee_le"


Champs(1, 1) = "Code"
Champs(2, 1) = "Sous-Ensseble"
Champs(3, 1) = "MASSE total (KG)"
Champs(4, 1) = "Nb lignes sans masse"
Champs(5, 1) = "Nb total de lignes"
Champs(6, 1) = "NumEncelade"
Champs(7, 1) = "Projet"
Champs(8, 1) = "Programme"
Champs(9, 1) = "Désignation"
Champs(10, 1) = "Véhicule"
Champs(11, 1) = "Realignement"
Champs(12, 1) = "Demandeur"
Champs(13, 1) = "Atelier"
Champs(14, 1) = "DatePrevue"
Champs(15, 1) = "Désination FE"
Champs(16, 1) = "Type"
Champs(17, 1) = "Pièce"
Champs(18, 1) = "Indice"
Champs(19, 1) = "Commantaire"
Champs(20, 1) = "Designation_Piece"
Champs(21, 1) = "Famille/FE/V"
Champs(22, 1) = "[Num_appro S"
Champs(23, 1) = "Quantite"
Champs(24, 1) = "Bon_sortie"
Champs(25, 1) = "Observation_ou_magasin"
Champs(26, 1) = "N° appro"
Champs(27, 1) = "Masse"
Champs(28, 1) = "Commandee le"
Champs(29, 1) = "Collectee le"
Champs(30, 1) = "Expedee le"
Champs(31, 1) = "Reçue le"
Champs(32, 1) = "Qt reçue"
Champs(33, 1) = "Reste à recevoir"
Champs(34, 1) = "N°de réception"
Champs(35, 1) = "Reaffectee a"
Champs(36, 1) = "Qt Reaffectee"
Champs(37, 1) = "Reaffectee Le"
Champs(38, 1) = "En provenance de"
Champs(39, 1) = "Retournee le"
Champs(40, 1) = "QT Retournee"
Champs(41, 1) = "N° de levraison"
Champs(42, 1) = "Modifiee le"


If Request("nbSelect") > 0 Then
 
 nbLignes = CInt(Request("nbSelect"))

RecapSersh = RecapSersh & " <input type=""hidden"" name=""SherchPage"" value='" & Request("SherchPage") & "'>" & vbCrLf
RecapSersh = RecapSersh & "<table width=""100%"" border=""2"" bgcolor='#B7B8DE' cellpadding=""1"" cellspacing=""0"">" & vbCrLf
RecapSersh = RecapSersh & "  <tr valign=""bottom"">" & vbCrLf
RecapSersh = RecapSersh & "   <td>" & vbCrLf
RecapSersh = RecapSersh & "    <table width=""100%"" border=""0"">" & vbCrLf
RecapSersh = RecapSersh & "     <tr valign=""bottom"">" & vbCrLf
RecapSersh = RecapSersh & "       <td align=""left"" valign=""middle"">" & vbCrLf
RecapSersh = RecapSersh & "<font color='#FFFFFF' size='2' >" & vbCrLf
RecapSersh = RecapSersh & "        " & Session("Encelade_Message")
RecapSersh = RecapSersh & "</font>" & vbCrLf
RecapSersh = RecapSersh & "       </td>" & vbCrLf
RecapSersh = RecapSersh & "       <td>" & vbCrLf
'    RecapSersh=RecapSersh &  "        <font size=""1"" > Nombre de lignes a afficher </font> " & vbcrlf
    If Session("candidat_UserType") = "Administrator" Then
       ' RecapSersh=RecapSersh &  "        <INPUT TYPE=""hidden"" NAME =""NbLiges"" value=""" & GetDefault("con_PageNum", "15", Session("candidat_ADOContact")) & """>")

'   con_frmSetting = con_frmSetting & vbCrLf & ("<TD><input type=""text"" NAME=""NBLignesRecherche"" VALUE=" & GetDefault("NBLignesRecherche", "25", Session("candidat_ADOContact")) & "></TD>       ")
   Else
    RecapSersh = RecapSersh & "        <INPUT TYPE=""hidden"" NAME =""NbLiges"" value='25'>" & vbCrLf
    End If
    RecapSersh = RecapSersh & "       </td>" & vbCrLf
    RecapSersh = RecapSersh & "       <td align=""right"" class=""smallerheader"" valign=""middle""></td>" & vbCrLf
    RecapSersh = RecapSersh & "<TD width='2' valign='top'>" & AfficheBountonSersh("Modifier", 11, "'javascript:ModifierRechereche();'") & "</td>"
 

    RecapSersh = RecapSersh & "       <td align=""right"" color='#FFFFFF'  valign=""Right"" ><FONT  color='#FFFFFF' >Recherche sur : "
  
   
       Select Case Trim("" & Request("SersheBASE"))
        Case "22"
            RecapSersh = RecapSersh & "<input  Name='SersheBASE' type='hidden' value='22'>Projets" & vbCrLf
         
         Case "9"
            RecapSersh = RecapSersh & "<input  Name='SersheBASE' type='hidden' value='9'>Véhicule" & vbCrLf
    
        Case "10"
            RecapSersh = RecapSersh & "<input  Name='SersheBASE' type='hidden' value='10'>Sous-Ensembles" & vbCrLf
    
        Case "8"
            RecapSersh = RecapSersh & "<input  Name='SersheBASE' type='hidden' value='8'>Pièces" & vbCrLf
        Case "12"
            RecapSersh = RecapSersh & "<input  Name='SersheBASE' type='hidden' value='12'>B.Réception" & vbCrLf
        Case "13"
            RecapSersh = RecapSersh & "<input  Name='SersheBASE' type='hidden' value='13'>B.Livraison" & vbCrLf
    End Select
   
      
           RecapSersh = RecapSersh & "</FONT></td>" & vbCrLf
    
    RecapSersh = RecapSersh & "      </tr>" & vbCrLf
    RecapSersh = RecapSersh & "</td>" & vbCrLf
    RecapSersh = RecapSersh & "      </tr>" & vbCrLf
    RecapSersh = RecapSersh & "</td>" & vbCrLf
    RecapSersh = RecapSersh & "     </table>" & vbCrLf
    RecapSersh = RecapSersh & "    </td>" & vbCrLf
    RecapSersh = RecapSersh & "   </tr>" & vbCrLf
    RecapSersh = RecapSersh & "  <table>" & vbCrLf
    RecapSersh = RecapSersh & "     <table  width=""100%"" >" & vbCrLf
    RecapSersh = RecapSersh & "     <TR width=></td></tr>" & vbCrLf
    RecapSersh = RecapSersh & "     <TR ><td>" & vbCrLf
      RecapSersh = RecapSersh & "     <table bgcolor='#B7B8DE' width=""70%"" align=""center"">" & vbCrLf
    RecapSersh = RecapSersh & "    <TR>" & vbCrLf
    RecapSersh = RecapSersh & "     <TD  >" & vbCrLf
    RecapSersh = RecapSersh & "      <font  color='#FFFFFF' size='1' >    TRIER </font>  " & vbCrLf
    RecapSersh = RecapSersh & "     </TD>" & vbCrLf
    RecapSersh = RecapSersh & "     <TD  >" & vbCrLf
    RecapSersh = RecapSersh & "      <font  color='#FFFFFF' size='1' >    NOM DES CHAMPS   </font>  " & vbCrLf
    RecapSersh = RecapSersh & "     </TD>" & vbCrLf
    RecapSersh = RecapSersh & "     <TD>" & vbCrLf
    RecapSersh = RecapSersh & "      <font  color='#FFFFFF' size='1' >  OPERATEUR  </font> " & vbCrLf
    RecapSersh = RecapSersh & "     </TD>" & vbCrLf
    RecapSersh = RecapSersh & "     <TD>" & vbCrLf
    RecapSersh = RecapSersh & "      <font  color='#FFFFFF' size='1' >  VALEUR  </font> " & vbCrLf
    RecapSersh = RecapSersh & "     </TD>" & vbCrLf
    RecapSersh = RecapSersh & "     <TD>" & vbCrLf
    RecapSersh = RecapSersh & "      <font  color='#FFFFFF' size='1' > OPERATEUR LOGIQUE</font>" & vbCrLf
    RecapSersh = RecapSersh & "     </TD>" & vbCrLf
    RecapSersh = RecapSersh & "   </TR>" & vbCrLf
    I = 1
    
 '**************************************************************************************************************
 'Existe
 '**************************************************************************************************************
  If nbLignes > 0 Then
     For I = 1 To nbLignes
    RecapSersh = RecapSersh & "   <TR>" & vbCrLf
     SelecetList = "" & Request("radiobutton")
       RecapSersh = RecapSersh & "<TD>"
   If SelecetList = CStr(I) Then
    RecapSersh = RecapSersh & "<input NAME='radiobutton" & CStr(I) & "' type='hidden' value='" & CStr(I) & "'/> A -> Z"
   End If
    If SelecetList = CStr(I) & "DESC" Then
    RecapSersh = RecapSersh & "<input NAME='radiobutton" & CStr(I) & "' type='hidden' value='" & CStr(I) & "DESC'/> Z -> A"
   End If
    RecapSersh = RecapSersh & "  &nbsp;</TD><TD>" & vbCrLf
     For j = 1 To UBound(Champs)
    
    
    If Request("Champ" & CStr(I)) = Champs(j, 0) Then
     RecapSersh = RecapSersh & "<input name='Champ" & CStr(I) & "' type='hidden'value='" & Champs(j, 0) & "'/>" & Champs(j, 1)
    End If
  
Next
  
    RecapSersh = RecapSersh & "   </TD>" & vbCrLf
     
                
                RecapSersh = RecapSersh & "         <TD>" & vbCrLf
      
                If Request("OPERATEUR" & CStr(I)).Count > 0 Then
                    slt = Request("OPERATEUR" & I)
                Else
                    slt = ""
                End If
                For j = 0 To UBound(aSlt_OPERATEUR, 2)
                     If Trim(slt) = Trim(aSlt_OPERATEUR(0, j)) Then
                     RecapSersh = RecapSersh & "<input NAME='OPERATEUR" & CStr(I) & "' type='hidden' value='" & aSlt_OPERATEUR(0, j) & "'/>" & aSlt_OPERATEUR(1, j)
                     
                     
                     End If
                 Next
       RecapSersh = RecapSersh & " <td><input type=""hidden"" NAME =""VALEUR" & CStr(I) & """ value='" & Request("VALEUR" & CStr(I)) & "'>" & Request("VALEUR" & CStr(I)) & "</td>" & vbCrLf
       RecapSersh = RecapSersh & "         <TD>" & vbCrLf
        
       
        If Request("Logique" & CStr(I)).Count > 0 Then
            slt_Logique = Request("Logique" & I)
        Else
            slt_Logique = ""
        End If
       For j = 0 To UBound(aSlt_Logique, 2)
            If Trim(slt_Logique) = Trim(aSlt_Logique(0, j)) Then
            RecapSersh = RecapSersh & "<input NAME='Logique" & CStr(I) & "' type='hidden'value='" & aSlt_Logique(0, j) & "'/>" & aSlt_Logique(1, j)
            End If
        Next
        
      
       RecapSersh = RecapSersh & " <INPUT TYPE='hidden' NAME ='SaveLogique" & CStr(I) & "'  VALUE='" & Request("Logique" & CStr(I)) & "'>" & vbCrLf
       RecapSersh = RecapSersh & "         </TD>" & vbCrLf
    RecapSersh = RecapSersh & "   </TR>" & vbCrLf
      
   Next
  End If
    RecapSersh = RecapSersh & "     </table>" & vbCrLf
 
    
     RecapSersh = RecapSersh & "</table>" & vbCrLf
     
       
       RecapSersh = RecapSersh & "<input type='hidden' name='nbSelect' value='" & I - 1 & "'>" & vbCrLf
   
  End If
End Function

Function RetournHeder(ChampName, Disabled)

RetournHeder = Disabled
If Disabled <> " Disabled " Then
  
    If Trim("" & Session(ChampName)) = "True" Then
        RetournHeder = ""
        
    Else
         RetournHeder = " Disabled "
    End If
End If
End Function
Function RetournIsNum(valeur)
RetournIsNum = "false"
Select Case valeur
    Case "T_SousEnsembles.Code"
    Case "T_SousEnsembles.[Sous-Ensseble]"
    Case "T_SousEnsembles.[MASSE total (KG)]"
        RetournIsNum = "True"
    Case "T_SousEnsembles.[Nb lignes sans masse]"
        RetournIsNum = "True"
    Case "T_SousEnsembles.[Nb total de lignes]"
        RetournIsNum = "True"
    Case "t_Vehicule.NumEncelade"
    Case "t_Vehicule.Projet"
    Case "t_Vehicule.Programme"
    Case "t_Vehicule.Désignation"
    Case "t_Vehicule.Véhicule"
    Case "t_Vehicule.Realignement"
    Case "t_Vehicule.Demandeur"
    Case "t_Vehicule.Atelier"
    Case "t_Vehicule.DatePrevue"
    Case "T_Nomenclature.[Désination FE]"
    Case "T_Nomenclature.Type"
    Case "T_Nomenclature.Pièce"
    Case "T_Nomenclature.Indice"
    Case "T_Nomenclature.Commantaire"
    Case "T_Nomenclature.Designation_Piece"
    Case "T_Nomenclature.[Famille/FE/V]"
    Case "T_Nomenclature.[Num_appro_ S]"
    Case "T_Nomenclature.Quantite"
        RetournIsNum = "True"
    Case "T_Nomenclature.Bon_sortie"
    Case "T_Nomenclature.Observation_ou_magasin"
    Case "T_Nomenclature.Num_appro"
    Case "T_Nomenclature.Masse"
            RetournIsNum = "True"
    Case "T_Nomenclature.Commandee_le"
        RetournIsNum = "Date"
    Case "T_Nomenclature.Collectee_le"
        RetournIsNum = "Date"
    Case "T_Nomenclature.Expedee_le"
        RetournIsNum = "Date"
    Case "T_Nomenclature.Reçue_le"
        RetournIsNum = "Date"
    Case "T_Nomenclature.Qt_reçue"
            RetournIsNum = "True"
    Case "T_Nomenclature.Reste_a_recevoir"
        RetournIsNum = "True"
    Case "T_Nomenclature.N_de_BL"
    Case "T_Nomenclature.Reaffectee_a"
    Case "T_Nomenclature.Qt_Reaffectee"
            RetournIsNum = "True"
    Case "T_Nomenclature.Reaffectee_Le"
        RetournIsNum = "Date"
    Case "T_Nomenclature.En_provenance_de"
        RetournIsNum = "True"
    Case "T_Nomenclature.Retournee_le"
        RetournIsNum = "Date"
    Case "T_Nomenclature.QT_Retournee"
        RetournIsNum = "True"
    Case "T_Nomenclature.Num_de_retour"
    Case "T_Nomenclature.Modifiee_le"
        RetournIsNum = "Date"
End Select
End Function
Function TitreNomenclatur() As String
 
 TitreNomenclatur = TitreNomenclatur & "<tr>" & vbCrLf

        TitreNomenclatur = TitreNomenclatur & "         <td align='center' bgcolor='#CCCCCC'>Supprimer</td>" & vbCrLf
        TitreNomenclatur = TitreNomenclatur & "         <td align='center' bgcolor='#CCCCCC'>Activer</td>" & vbCrLf
        TitreNomenclatur = TitreNomenclatur & "<td align='center' bgcolor='#CCCCCC'>Projet</td>" & vbCrLf
        TitreNomenclatur = TitreNomenclatur & "         <td   align='center' bgcolor='#CCCCCC'>Véhicule | Sous-Ensseble</td>" & vbCrLf
        TitreNomenclatur = TitreNomenclatur & "         <td   align='center' bgcolor='#CCCCCC'>Désination pièce </td>" & vbCrLf
        TitreNomenclatur = TitreNomenclatur & "<td align='center' bgcolor='#CCCCCC'>Pièce </td>" & vbCrLf
        TitreNomenclatur = TitreNomenclatur & "             <td align='center' bgcolor='#CCCCCC'>Indice</td>" & vbCrLf
        TitreNomenclatur = TitreNomenclatur & "<td align='center' bgcolor='#CCCCCC'>Quantitée </td>" & vbCrLf
        TitreNomenclatur = TitreNomenclatur & "<td align='center' bgcolor='#CCCCCC'>Reste à recevoir </td>" & vbCrLf
         TitreNomenclatur = TitreNomenclatur & "        <td align='center' bgcolor='#CCCCCC'>Action</td>"
        TitreNomenclatur = TitreNomenclatur & "        <td align='center' bgcolor='#CCCCCC'>Délais</td>"
 
        TitreNomenclatur = TitreNomenclatur & "<td align='center' bgcolor='#CCCCCC'>Qt reçue </td>" & vbCrLf
        TitreNomenclatur = TitreNomenclatur & "<td align='center' bgcolor='#CCCCCC'>Reçue le</td>" & vbCrLf
        TitreNomenclatur = TitreNomenclatur & "<td align='center' bgcolor='#CCCCCC'>En provenance de</td>" & vbCrLf
        TitreNomenclatur = TitreNomenclatur & "<td align='center' bgcolor='#CCCCCC'>N° de BR</td>" & vbCrLf
        TitreNomenclatur = TitreNomenclatur & "<td align='center' bgcolor='#CCCCCC'>QT Retournée </td>"
        TitreNomenclatur = TitreNomenclatur & "<td align='center' bgcolor='#CCCCCC'> Retournée le </td>" & vbCrLf
        TitreNomenclatur = TitreNomenclatur & "<td align='center' bgcolor='#CCCCCC'>Reaffectée à </td>" & vbCrLf
        TitreNomenclatur = TitreNomenclatur & "<td align='center' bgcolor='#CCCCCC'>N° BL</td>"
         TitreNomenclatur = TitreNomenclatur & "<td align='center' bgcolor='#CCCCCC'>Avant modification</td>"
        TitreNomenclatur = TitreNomenclatur & "<td align='center' bgcolor='#CCCCCC'>Modifiée le</td>"
        TitreNomenclatur = TitreNomenclatur & "<td align='center' bgcolor='#CCCCCC'>Modifiée par </td>" & vbCrLf
         TitreNomenclatur = TitreNomenclatur & "<td align='center' bgcolor='#CCCCCC'>Commentaires</td>" & vbCrLf

        TitreNomenclatur = TitreNomenclatur & "</tr>"
        
End Function
