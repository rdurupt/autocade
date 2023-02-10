VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "Importer Archives :"
   ClientHeight    =   11835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14010
   OleObjectBlob   =   "UserForm5.dsx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Noquite As Boolean
Private Sub CommandButton1_Click()
Dim Msg As String
Dim MyRange
Dim IndiceProjet As Long
Dim Id_Pieces As Long
Dim IdProjet As Long
Dim sql As String
Dim LibPice As String
Dim Rs As Recordset
If MsgBox("vous réimporter les enregistrements Archivés", vbYesNo + vbQuestion, "Importer Archives") = vbNo Then Exit Sub
Set MyRange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
ProgressBar1.Value = 0
ProgressBar1.Max = MyRange.Rows.Count
For i = 2 To MyRange.Rows.Count
ProgressBar1.Value = i
    If UCase(MyRange(i, 1)) <> 0 Then
         IndiceProjet = CInt(Me.Spreadsheet1.Cells(i, 14))
        sql = "SELECT  Archive_T_indiceProjet.Id_Pieces FROM Archive_T_indiceProjet "
        sql = sql & "WHERE Archive_T_indiceProjet.Id=" & IndiceProjet & ";"
        Set Rs = Con.OpenRecordSet(sql)
        If Rs.EOF = False Then
            Id_Pieces = Rs!Id_Pieces
            
             sql = "SELECT  Archive_T_Pieces.IdProjet,Archive_T_Pieces.Description  FROM Archive_T_Pieces "
            sql = sql & "WHERE Archive_T_Pieces.Id=" & Id_Pieces & ";"
             Set Rs = Con.OpenRecordSet(sql)
              If Rs.EOF = False Then
                    IdProjet = Rs!IdProjet
                    LibPice = "" & Rs!Description
           End If
            

            
            sql = "SELECT T_Projet.id FROM T_Projet "
            sql = sql & "WHERE T_Projet.id=" & IdProjet & ";"
            Set Rs = Con.OpenRecordSet(sql)
            If Rs.EOF = True Then
                sql = "INSERT INTO T_Projet ( id, Projet, Description, CleAc ) "
            sql = sql & "SELECT Archive_T_Projet.id, Archive_T_Projet.Projet, "
            sql = sql & "Archive_T_Projet.Description, Archive_T_Projet.CleAc "
            sql = sql & "FROM Archive_T_Projet "
            sql = sql & "WHERE Archive_T_Projet.id=" & IdProjet & ";"
            Con.Exequte sql
            End If
            
            
              
            sql = "SELECT T_Pieces.Id FROM T_Pieces "
            sql = sql & "WHERE T_Pieces.Id=" & Id_Pieces & ";"
             Set Rs = Con.OpenRecordSet(sql)
            If Rs.EOF = True Then
                sql = "INSERT INTO T_Pieces ( id, IdProjet, Description ) "
                sql = sql & "SELECT Archive_T_Pieces.Id, Archive_T_Pieces.IdProjet, Archive_T_Pieces.Description "
                sql = sql & "FROM Archive_T_Pieces "
                sql = sql & "WHERE Archive_T_Pieces.Id=" & Id_Pieces & ";"
                Con.Exequte sql

            End If
            sql = "SELECT T_indiceProjet.Id FROM T_indiceProjet "
            sql = sql & "WHERE T_indiceProjet.Id=" & IndiceProjet & ";"
            Set Rs = Con.OpenRecordSet(sql)
            If Rs.EOF = True Then
                sql = "INSERT INTO T_indiceProjet ( Id, Id_Pieces, Description, PL, PL_Indice,  "
                sql = sql & "OU, OU_Indice, Li, LI_Indice, PI, IdStatus, PI_Indice, IdStatusSave,  "
                sql = sql & "IdApprobateur, PlAutoCadSaveAs, PlAutoCadSave, NbErr, Client, Destinataire,  "
                sql = sql & "Service, DessineDate, DessineNOM, VerifieDate, VerifieNom, ApprouveDate,  "
                sql = sql & "ApprouveNom, Responsable, Vague, Equipement, RefPF, Ensemble, CleAc, RefP,  "
                sql = sql & "Masse, LiAutoCadSaveAs, LiAutoCadSave, OuAutoCadSaveAs, OuAutoCadSave,  "
                sql = sql & "Archiver, Cartouche, Version, Pere ) "
                sql = sql & "SELECT Archive_T_indiceProjet.Id,Archive_T_indiceProjet.Id_Pieces,  "
                sql = sql & "Archive_T_indiceProjet.Description, Archive_T_indiceProjet.PL,  "
                sql = sql & "Archive_T_indiceProjet.PL_Indice, Archive_T_indiceProjet.[OU],  "
                sql = sql & "Archive_T_indiceProjet.OU_Indice, Archive_T_indiceProjet.Li,  "
                sql = sql & "Archive_T_indiceProjet.LI_Indice, Archive_T_indiceProjet.PI,  "
                sql = sql & "Archive_T_indiceProjet.IdStatus, Archive_T_indiceProjet.PI_Indice,  "
                sql = sql & "Archive_T_indiceProjet.IdStatusSave, Archive_T_indiceProjet.IdApprobateur,  "
                sql = sql & "Archive_T_indiceProjet.PlAutoCadSaveAs, Archive_T_indiceProjet.PlAutoCadSave,  "
                sql = sql & "Archive_T_indiceProjet.NbErr, Archive_T_indiceProjet.Client,  "
                sql = sql & "Archive_T_indiceProjet.Destinataire, Archive_T_indiceProjet.Service,  "
                sql = sql & "Archive_T_indiceProjet.DessineDate, Archive_T_indiceProjet.DessineNOM,  "
                sql = sql & "Archive_T_indiceProjet.VerifieDate, Archive_T_indiceProjet.VerifieNom,  "
                sql = sql & "Archive_T_indiceProjet.ApprouveDate, Archive_T_indiceProjet.ApprouveNom,  "
                sql = sql & "Archive_T_indiceProjet.Responsable, Archive_T_indiceProjet.Vague,  "
                sql = sql & "Archive_T_indiceProjet.Equipement, Archive_T_indiceProjet.RefPF,  "
                sql = sql & "Archive_T_indiceProjet.Ensemble , Archive_T_indiceProjet.CleAc,  "
                sql = sql & "Archive_T_indiceProjet.RefP, Archive_T_indiceProjet.Masse,  "
                sql = sql & "Archive_T_indiceProjet.LiAutoCadSaveAs, Archive_T_indiceProjet.LiAutoCadSave,  "
                sql = sql & "Archive_T_indiceProjet.OuAutoCadSaveAs, Archive_T_indiceProjet.OuAutoCadSave,  "
                sql = sql & "Archive_T_indiceProjet.Archiver, Archive_T_indiceProjet.Cartouche,  "
                sql = sql & "Archive_T_indiceProjet.Version, Archive_T_indiceProjet.Pere "
                sql = sql & "FROM Archive_T_indiceProjet "
                sql = sql & "WHERE Archive_T_indiceProjet.Id=" & IndiceProjet & ";"
                Con.Exequte sql
            End If
            sql = "SELECT Connecteurs.Id_IndiceProjet FROM Connecteurs "
            sql = sql & "WHERE Connecteurs.Id_IndiceProjet=" & IndiceProjet & ";"
             Set Rs = Con.OpenRecordSet(sql)
            If Rs.EOF = True Then
                sql = "INSERT INTO Connecteurs ( Numéro, Id_IndiceProjet, CONNECTEUR,  "
                sql = sql & "[O/N], DESIGNATION, CODE_APP, N°, POS, [POS-OUT], PRECO1, PRECO2, [100%] ) "
                sql = sql & "SELECT Archive_Connecteurs.Numéro, Archive_Connecteurs.Id_IndiceProjet,  "
                sql = sql & "Archive_Connecteurs.CONNECTEUR,Archive_Connecteurs.[O/N],  "
                sql = sql & "Archive_Connecteurs.DESIGNATION, Archive_Connecteurs.CODE_APP,  "
                sql = sql & "Archive_Connecteurs.N°, Archive_Connecteurs.POS, Archive_Connecteurs.[POS-OUT],  "
                sql = sql & "Archive_Connecteurs.PRECO1, Archive_Connecteurs.PRECO2, Archive_Connecteurs.[100%] "
                sql = sql & "FROM Archive_Connecteurs "
                sql = sql & "WHERE Archive_Connecteurs.Id_IndiceProjet=" & IndiceProjet & ";"
                Con.Exequte sql
            End If
            sql = "SELECT Ligne_Tableau_fils.Id_IndiceProjet FROM Ligne_Tableau_fils "
            sql = sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IndiceProjet & ";"
            Set Rs = Con.OpenRecordSet(sql)
            If Rs.EOF = True Then
                sql = "INSERT INTO Ligne_Tableau_fils ( Numéro, Id_IndiceProjet, LIAI, DESIGNATION, FIL,  "
                sql = sql & "SECT, TEINT, TEINT2, ISO, [LONG], [LONG CP], COUPE, POS, [POS-OUT], FA, APP,  "
                sql = sql & "VOI, POS2, [POS-OUT2], FA2, APP2, VOI2, PRECO, [OPTION] ) "
                sql = sql & "SELECT Archive_Ligne_Tableau_fils.Numéro,  "
                sql = sql & "Archive_Ligne_Tableau_fils.Id_IndiceProjet,Archive_Ligne_Tableau_fils.LIAI,  "
                sql = sql & "Archive_Ligne_Tableau_fils.DESIGNATION, Archive_Ligne_Tableau_fils.FIL, "
                sql = sql & "Archive_Ligne_Tableau_fils.SECT, Archive_Ligne_Tableau_fils.TEINT,  "
                sql = sql & "Archive_Ligne_Tableau_fils.TEINT2, Archive_Ligne_Tableau_fils.ISO,  "
                sql = sql & "Archive_Ligne_Tableau_fils.LONG, Archive_Ligne_Tableau_fils.[LONG CP],  "
                sql = sql & "Archive_Ligne_Tableau_fils.COUPE, Archive_Ligne_Tableau_fils.POS,  "
                sql = sql & "Archive_Ligne_Tableau_fils.[POS-OUT], Archive_Ligne_Tableau_fils.FA,  "
                sql = sql & "Archive_Ligne_Tableau_fils.APP, Archive_Ligne_Tableau_fils.VOI,  "
                sql = sql & "Archive_Ligne_Tableau_fils.POS2, Archive_Ligne_Tableau_fils.[POS-OUT2],  "
                sql = sql & "Archive_Ligne_Tableau_fils.FA2, Archive_Ligne_Tableau_fils.APP2,  "
                sql = sql & "Archive_Ligne_Tableau_fils.VOI2, Archive_Ligne_Tableau_fils.PRECO,  "
                sql = sql & "Archive_Ligne_Tableau_fils.OPTION "
                sql = sql & "FROM Archive_Ligne_Tableau_fils "
                sql = sql & "WHERE Archive_Ligne_Tableau_fils.Id_IndiceProjet=" & IndiceProjet & ";"
            Con.Exequte sql
            End If
            
             sql = "SELECT Composants.Id_IndiceProjet FROM Composants "
             sql = sql & "WHERE Composants.Id_IndiceProjet=" & IndiceProjet & ";"
             Set Rs = Con.OpenRecordSet(sql)
            If Rs.EOF = True Then
                sql = "INSERT INTO Composants ( Id, Id_IndiceProjet, DESIGNCOMP, NUMCOMP, REFCOMP, Path ) "
                sql = sql & "SELECT Archive_Composants.Id, Archive_Composants.Id_IndiceProjet,  "
                sql = sql & "Archive_Composants.DESIGNCOMP, Archive_Composants.NUMCOMP,  "
                sql = sql & "Archive_Composants.REFCOMP, Archive_Composants.Path "
                sql = sql & "FROM Archive_Composants "
                sql = sql & "WHERE Archive_Composants.Id_IndiceProjet=" & IndiceProjet & ";"
            Con.Exequte sql

            End If
            
            sql = "SELECT Nota.Id_IndiceProjet FROM Nota "
            sql = sql & "WHERE Nota.Id_IndiceProjet=" & IndiceProjet & ";"
             Set Rs = Con.OpenRecordSet(sql)
            If Rs.EOF = True Then
                sql = "INSERT INTO Nota ( Id, Id_IndiceProjet, NOTA, NUMNOTA )  "
                sql = sql & "SELECT Archive_Nota.Id, Archive_Nota.Id_IndiceProjet, Archive_Nota.NOTA,   "
                sql = sql & "Archive_Nota.NUMNOTA  "
                sql = sql & "FROM Archive_Nota  "
                sql = sql & "WHERE Archive_Nota.Id_IndiceProjet=" & IndiceProjet & ";"
                Con.Exequte sql

            End If
             sql = "DELETE Archive_T_Pieces.* FROM Archive_T_Pieces "
                        sql = sql & "WHERE Archive_T_Pieces.Id=" & Id_Pieces & ";"
                        Con.Exequte sql
                        sql = "SELECT Archive_T_Pieces.Id FROM Archive_T_Pieces "
                        sql = sql & "WHERE Archive_T_Pieces.IdProjet=" & IdProjet & ";"
                        Set Rs = Con.OpenRecordSet(sql)
                        If Rs.EOF = True Then
                            sql = "DELETE Archive_T_Projet.* FROM Archive_T_Projet "
                            sql = sql & "WHERE Archive_T_Projet.id=" & IdProjet & ";"
                            Con.Exequte sql

                        End If
        End If
    End If
Next i
 Noquite = False
Me.Hide
End Sub

Public Sub Charge(MyForm As Object, Optional Filtre As String, Optional boolTxt As Boolean, Optional boolArchive As Boolean)
Dim sql As String
Dim Rs As Recordset
Dim IndexRow As Long
Dim IndexCol As Long

boolTxts = boolTxt
IndexRow = 1
IndexCol = 0
Me.Spreadsheet1.Columns(1).Locked = False

 Me.Spreadsheet1.Columns(1).NumberFormat = "Yes/No"
 Me.Spreadsheet1.Columns(1).Locked = True
  
sql = "SELECT 0 AS Importer ,Archive_SelectProjets.* "
sql = sql & "FROM Archive_SelectProjets; "
Set Rs = Con.OpenRecordSet(sql)
Rs.Filter = Filtre
Set MyFormCible = MyForm
While Rs.EOF = False
IndexRow = IndexRow + 1
For IndexCol = 0 To Rs.Fields.Count - 11
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Locked = False
    If IndexCol > 2 And IndexCol < 7 Then
         Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1) = Trim("" & Rs.Fields(IndexCol)) & Trim("_" & Rs.Fields(IndexCol + 17))
    Else
        Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1) = Trim("" & Rs.Fields(IndexCol))
    End If
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Interior.Color = &HFFC0FF
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
If Mytype = "Importer O/N" Then MyLocked = False

End Function

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
