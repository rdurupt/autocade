VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "Supprimer/Archiver:"
   ClientHeight    =   10485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13530
   OleObjectBlob   =   "UserForm4.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Noquite As Boolean
Public Sub Charge(MyForm As Object, Optional Filtre As String, Optional boolTxt As Boolean, Optional boolArchive As Boolean)
Dim sql As String
Dim Rs As Recordset
Dim IndexRow As Long
Dim IndexCol As Long
Me.Spreadsheet1.Columns(1).Locked = False
  Me.Spreadsheet1.Columns(2).Locked = False
 Me.Spreadsheet1.Columns(1).NumberFormat = "Yes/No"
 Me.Spreadsheet1.Columns(2).NumberFormat = "Yes/No"
 Me.Spreadsheet1.Columns(1).Locked = True
  Me.Spreadsheet1.Columns(2).Locked = True

 
boolTxts = boolTxt
IndexRow = 1
IndexCol = 0

sql = "SELECT  0 AS Suprimers  , 0 AS Archivers,SelectProjets.* "
sql = sql & "FROM SelectProjets; "
Set Rs = Con.OpenRecordSet(sql)
Rs.Filter = Filtre
Set MyFormCible = MyForm
While Rs.EOF = False
IndexRow = IndexRow + 1
For IndexCol = 0 To Rs.Fields.Count - 12
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Locked = False
    If IndexCol > 3 And IndexCol < 8 Then
         Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1) = Trim("" & Rs.Fields(IndexCol)) & Trim("_" & Rs.Fields(IndexCol + 18))
    Else
        Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1) = Trim("" & Rs.Fields(IndexCol))
    End If
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Interior.Color = ChoixCouleur(Rs.Fields(Rs.Fields.Count - 11))
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Locked = MyLocked(Me.Spreadsheet1.Cells(1, IndexCol + 1), Rs.Fields(Rs.Fields.Count - 11))

Next IndexCol

Rs.MoveNext
Wend

Set Rs = Con.CloseRecordSet(Rs)
 
Me.Show
End Sub
Function ChoixCouleur(Mode As Long) As Long
   
  
   Select Case Mode
    Case 1
        ChoixCouleur = 16777164
    Case 2
    ChoixCouleur = 10079487
    Case 3
        ChoixCouleur = 13434828
   End Select

End Function
Function MyLocked(Mytype As String, Statues As Long) As Boolean
MyLocked = True
If Mytype = "Supprimer O/N" And Statues = 1 Then MyLocked = False
If Mytype = "Archiver O/N" And Statues = 3 Then MyLocked = False

End Function

Private Sub CommandButton1_Click()
Dim Msg As String
Dim MyRange
Dim IndiceProjet As Long
Dim Id_Pieces As Long
Dim IdProjet As Long
Dim sql As String
Dim LibPice As String
Dim Rs As Recordset
Msg = "Attention tous les enregistrements supprimés seront définitivement perdus." & vbCrLf & vbCrLf
Msg = Msg & "Tous les enregistrements archivés pourront être réinsérés par la suite" & vbCrLf & vbCrLf
Msg = Msg & "Voulez-vous continuer." & vbCrLf & vbCrLf
If MsgBox(Msg, vbYesNo + vbQuestion, "Supprimer/Archiver") = vbNo Then Exit Sub
Set MyRange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
ProgressBar1.Value = 0
ProgressBar1.Max = MyRange.Rows.Count
For i = 2 To MyRange.Rows.Count
ProgressBar1.Value = i
    If UCase(MyRange(i, 2)) <> 0 Then
        IndiceProjet = CInt(Me.Spreadsheet1.Cells(i, 15))
        sql = "SELECT  T_indiceProjet.Id_Pieces FROM T_indiceProjet "
        sql = sql & "WHERE T_indiceProjet.Id=" & IndiceProjet & ";"
        Set Rs = Con.OpenRecordSet(sql)
        If Rs.EOF = False Then
            Id_Pieces = Rs!Id_Pieces
            
             sql = "SELECT  T_Pieces.IdProjet,T_Pieces.Description  FROM T_Pieces "
            sql = sql & "WHERE T_Pieces.Id=" & Id_Pieces & ";"
             Set Rs = Con.OpenRecordSet(sql)
              If Rs.EOF = False Then
                    IdProjet = Rs!IdProjet
                    LibPice = "" & Rs!Description
                    sql = "SELECT Archive_T_Projet.id FROM Archive_T_Projet "
                    sql = sql & "WHERE Archive_T_Projet.id=" & IdProjet & ";"
                     Set Rs = Con.OpenRecordSet(sql)
                     If Rs.EOF = True Then
                        sql = "INSERT INTO Archive_T_Projet ( id, Projet, Description, CleAc ) "
                        sql = sql & "SELECT T_Projet.id, T_Projet.Projet, T_Projet.Description, T_Projet.CleAc "
                        sql = sql & "FROM T_Projet "
                   sql = sql & "WHERE T_Projet.id=" & IdProjet & ";"

                        Con.Exequte sql

                     End If
                     sql = "SELECT Archive_T_Pieces.Id FROM Archive_T_Pieces "
                     sql = sql & "WHERE Archive_T_Pieces.Id=" & Id_Pieces & ";"
                        Set Rs = Con.OpenRecordSet(sql)
                     If Rs.EOF = True Then
                        sql = "INSERT INTO Archive_T_Pieces ( id, IdProjet, Description ) "
                        sql = sql & "SELECT T_Pieces.Id, T_Pieces.IdProjet, T_Pieces.Description "
                        sql = sql & "FROM T_Pieces "
                        sql = sql & "WHERE T_Pieces.Id=" & Id_Pieces & ";"
                        Con.Exequte sql
                     End If
                     sql = "SELECT Archive_T_indiceProjet.Id FROM Archive_T_indiceProjet "
                     sql = sql & "WHERE Archive_T_indiceProjet.Id=" & IndiceProjet & ";"
                       Set Rs = Con.OpenRecordSet(sql)
                     If Rs.EOF = True Then
                        sql = "INSERT INTO Archive_T_indiceProjet  "
                        sql = sql & "( Id, Id_Pieces, Description, PL, PL_Indice, OU, OU_Indice,  "
                        sql = sql & "Li, LI_Indice, PI, IdStatus, PI_Indice, IdStatusSave,  "
                        sql = sql & "IdApprobateur, PlAutoCadSaveAs, PlAutoCadSave, NbErr, Client,  "
                        sql = sql & "Destinataire, Service, DessineDate, DessineNOM, VerifieDate,  "
                        sql = sql & "VerifieNom, ApprouveDate, ApprouveNom, Responsable, Vague,  "
                        sql = sql & "Equipement, RefPF, Ensemble, CleAc, RefP, Masse, LiAutoCadSaveAs,  "
                        sql = sql & "LiAutoCadSave, OuAutoCadSaveAs, OuAutoCadSave, Archiver, Cartouche,  "
                        sql = sql & "Version, Pere ) "
                        sql = sql & "SELECT T_indiceProjet.Id, T_indiceProjet.Id_Pieces,  "
                        sql = sql & "T_indiceProjet.Description, T_indiceProjet.PL, "
                        sql = sql & "T_indiceProjet.PL_Indice, T_indiceProjet.[OU],  "
                        sql = sql & "T_indiceProjet.OU_Indice, T_indiceProjet.Li,  "
                        sql = sql & "T_indiceProjet.LI_Indice, T_indiceProjet.PI,  "
                        sql = sql & "T_indiceProjet.IdStatus, T_indiceProjet.PI_Indice,  "
                        sql = sql & "T_indiceProjet.IdStatusSave, T_indiceProjet.IdApprobateur,  "
                        sql = sql & "T_indiceProjet.PlAutoCadSaveAs, T_indiceProjet.PlAutoCadSave, "
                        sql = sql & " T_indiceProjet.NbErr, T_indiceProjet.Client,  "
                        sql = sql & "T_indiceProjet.Destinataire, T_indiceProjet.Service, "
                        sql = sql & " T_indiceProjet.DessineDate, T_indiceProjet.DessineNOM, "
                        sql = sql & " T_indiceProjet.VerifieDate, T_indiceProjet.VerifieNom, "
                        sql = sql & " T_indiceProjet.ApprouveDate, T_indiceProjet.ApprouveNom, "
                        sql = sql & " T_indiceProjet.Responsable, T_indiceProjet.Vague, "
                        sql = sql & " T_indiceProjet.Equipement, T_indiceProjet.RefPF, "
                        sql = sql & " T_indiceProjet.Ensemble, T_indiceProjet.CleAc, "
                        sql = sql & " T_indiceProjet.RefP, T_indiceProjet.Masse, "
                        sql = sql & " T_indiceProjet.LiAutoCadSaveAs, T_indiceProjet.LiAutoCadSave, "
                        sql = sql & " T_indiceProjet.OuAutoCadSaveAs, T_indiceProjet.OuAutoCadSave, "
                        sql = sql & " T_indiceProjet.Archiver, T_indiceProjet.Cartouche ,  "
                        sql = sql & "T_indiceProjet.Version, T_indiceProjet.Pere "
                        sql = sql & "FROM T_indiceProjet "
                        sql = sql & "WHERE T_indiceProjet.Id=" & IndiceProjet & ";"
                        Con.Exequte sql

                     End If
                     sql = "SELECT Archive_Connecteurs.Id_IndiceProjet FROM Archive_Connecteurs "
                     sql = sql & "WHERE Archive_Connecteurs.Id_IndiceProjet=" & IndiceProjet & ";"
                    Set Rs = Con.OpenRecordSet(sql)
                    If Rs.EOF = True Then
                        sql = "INSERT INTO Archive_Connecteurs ( Numéro, Id_IndiceProjet, CONNECTEUR, "
                        sql = sql & "[O/N], DESIGNATION, CODE_APP, N°, POS, [POS-OUT], PRECO1, PRECO2, [100%] )"
                        sql = sql & "SELECT Connecteurs.Numéro, Connecteurs.Id_IndiceProjet, "
                        sql = sql & "Connecteurs.CONNECTEUR, Connecteurs.[O/N], "
                        sql = sql & "Connecteurs.DESIGNATION, Connecteurs.CODE_APP, Connecteurs.N°, "
                        sql = sql & "Connecteurs.POS, Connecteurs.[POS-OUT], Connecteurs.PRECO1, "
                        sql = sql & "Connecteurs.PRECO2, Connecteurs.[100%]"
                        sql = sql & "FROM Connecteurs "
                        sql = sql & "WHERE Connecteurs.Id_IndiceProjet=" & IndiceProjet & ";"
                        Con.Exequte sql
                    End If
                        sql = "SELECT Archive_Ligne_Tableau_fils.Id_IndiceProjet FROM Archive_Ligne_Tableau_fils "
                        sql = sql & "WHERE Archive_Ligne_Tableau_fils.Id_IndiceProjet=" & IndiceProjet & ";"
                      
                        Set Rs = Con.OpenRecordSet(sql)
                        If Rs.EOF = True Then
                            sql = "INSERT INTO Archive_Ligne_Tableau_fils ( Numéro, Id_IndiceProjet, LIAI, "
                        sql = sql & "DESIGNATION, FIL, SECT, TEINT, TEINT2, ISO, [LONG], [LONG CP], COUPE,  "
                        sql = sql & "POS, [POS-OUT], FA, APP, VOI, POS2, [POS-OUT2], FA2, APP2, VOI2,  "
                        sql = sql & "PRECO, [OPTION] ) "
                        sql = sql & "SELECT Ligne_Tableau_fils.Numéro, Ligne_Tableau_fils.Id_IndiceProjet,  "
                        sql = sql & "Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION,  "
                        sql = sql & "Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.SECT,  "
                        sql = sql & "Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2,  "
                        sql = sql & "Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.LONG,  "
                        sql = sql & "Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.COUPE,  "
                        sql = sql & "Ligne_Tableau_fils.POS, Ligne_Tableau_fils.[POS-OUT],  "
                        sql = sql & "Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP,  "
                        sql = sql & "Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.POS2,  "
                        sql = sql & "Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2,  "
                        sql = sql & "Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2,  "
                        sql = sql & "Ligne_Tableau_fils.PRECO, Ligne_Tableau_fils.OPTION "
                        sql = sql & "FROM Ligne_Tableau_fils "
                        sql = sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IndiceProjet & ";"
                        Con.Exequte sql

                        End If
                        
                         sql = "SELECT Archive_Composants.Id_IndiceProjet FROM Archive_Composants "
                        sql = sql & "WHERE Archive_Composants.Id_IndiceProjet=" & IndiceProjet & ";"
                      
                        Set Rs = Con.OpenRecordSet(sql)
                        If Rs.EOF = True Then
                            sql = "INSERT INTO Archive_Composants ( Id, Id_IndiceProjet, DESIGNCOMP,  "
                            sql = sql & "NUMCOMP, REFCOMP, Path ) "
                            sql = sql & "SELECT Composants.Id, Composants.Id_IndiceProjet,  "
                            sql = sql & "Composants.DESIGNCOMP, Composants.NUMCOMP, Composants.REFCOMP,  "
                            sql = sql & "Composants.Path "
                            sql = sql & "FROM Composants "
                            sql = sql & "WHERE Composants.Id_IndiceProjet=" & IndiceProjet & ";"
                            Con.Exequte sql
                        End If
                        
                        sql = "SELECT Archive_Nota.Id_IndiceProjet FROM Archive_Nota "
                        sql = sql & "WHERE Archive_Nota.Id_IndiceProjet=" & IndiceProjet & ";"
                      
                        Set Rs = Con.OpenRecordSet(sql)
                        If Rs.EOF = True Then
                            sql = "INSERT INTO Archive_Nota ( Id, Id_IndiceProjet, NOTA, NUMNOTA ) "
                            sql = sql & "SELECT Nota.Id, Nota.Id_IndiceProjet, Nota.NOTA, Nota.NUMNOTA "
                            sql = sql & "FROM Nota "
                            sql = sql & "WHERE Nota.Id_IndiceProjet=" & IndiceProjet & ";"
                            Con.Exequte sql
            
                        End If
                        sql = "DELETE T_Pieces.*, T_Pieces.Id FROM T_Pieces "
                        sql = sql & "WHERE T_Pieces.Id=" & Id_Pieces & ";"
                        Con.Exequte sql
                        sql = "SELECT T_Pieces.Id FROM T_Pieces "
                        sql = sql & "WHERE T_Pieces.IdProjet=" & IdProjet & ";"
                        Set Rs = Con.OpenRecordSet(sql)
                        If Rs.EOF = True Then
                            sql = "DELETE T_Projet.* FROM T_Projet "
                            sql = sql & "WHERE T_Projet.id=" & IdProjet & ";"
                            Con.Exequte sql

                        End If
                       
              End If
'               Sql = "SELECT T_indiceProjet.Id FROM T_indiceProjet "
'            Sql = Sql & "WHERE T_indiceProjet.PI='" & MyReplace(LibPice) & "' "
'            Sql = Sql & "ORDER BY T_indiceProjet.Id DESC;"
'            Set Rs = Con.OpenRecordSet(Sql)
'            If Rs.EOF = False Then
'                Sql = "UPDATE T_indiceProjet SET T_indiceProjet.Archiver = true "
'                Sql = Sql & "WHERE T_indiceProjet.Id=" & Rs!Id & ";"
'                Con.Exequte Sql
'
'            End If
               Set Rs = Con.CloseRecordSet(Rs)
        End If
        
        
    End If
    If UCase(MyRange(i, 1)) <> 0 Then
         IndiceProjet = CInt(Me.Spreadsheet1.Cells(i, 15))
        sql = "SELECT  T_indiceProjet.Id_Pieces FROM T_indiceProjet "
        sql = sql & "WHERE T_indiceProjet.Id=" & IndiceProjet & ";"
        Set Rs = Con.OpenRecordSet(sql)
        If Rs.EOF = False Then
            Id_Pieces = Rs!Id_Pieces
            
             sql = "SELECT  T_Pieces.IdProjet,T_Pieces.Description  FROM T_Pieces "
            sql = sql & "WHERE T_Pieces.Id=" & Id_Pieces & ";"
             Set Rs = Con.OpenRecordSet(sql)
              If Rs.EOF = False Then
                    IdProjet = Rs!IdProjet
                    LibPice = "" & Rs!Description
              End If
          End If
          
          sql = "DELETE T_Pieces.*, T_Pieces.Id FROM T_Pieces "
                        sql = sql & "WHERE T_Pieces.Id=" & Id_Pieces & ";"
                        Con.Exequte sql
                        sql = "SELECT T_Pieces.Id FROM T_Pieces "
                        sql = sql & "WHERE T_Pieces.IdProjet=" & IdProjet & ";"
                        Set Rs = Con.OpenRecordSet(sql)
                        If Rs.EOF = True Then
                            sql = "DELETE T_Projet.* FROM T_Projet "
                            sql = sql & "WHERE T_Projet.id=" & IdProjet & ";"
                            Con.Exequte sql

                        End If
          
          sql = "SELECT T_indiceProjet.Id FROM T_indiceProjet "
            sql = sql & "WHERE T_indiceProjet.PI='" & MyReplace(LibPice) & "' "
            sql = sql & "ORDER BY T_indiceProjet.Id DESC;"
            Set Rs = Con.OpenRecordSet(sql)
            If Rs.EOF = False Then
                sql = "UPDATE T_indiceProjet SET T_indiceProjet.Archiver = true "
                sql = sql & "WHERE T_indiceProjet.Id=" & Rs!Id & ";"
                Con.Exequte sql

            End If
               Set Rs = Con.CloseRecordSet(Rs)
    End If
Next i
Noquite = False
Me.Hide
End Sub

Private Sub CommandButton2_Click()
Noquite = False
Me.Hide
End Sub

Private Sub UserForm_Activate()
 Noquite = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite
End Sub
