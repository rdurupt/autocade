VERSION 5.00
Object = "{0002E540-0000-0000-C000-000000000046}#1.1#0"; "MSOWC.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form UserForm5 
   Caption         =   "Importer Archives :"
   ClientHeight    =   12645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16890
   Icon            =   "CherchPices3.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12645
   ScaleWidth      =   16890
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   14640
      TabIndex        =   7
      Top             =   0
      Width           =   2175
      Begin VB.Image Image1 
         Height          =   855
         Left            =   40
         Picture         =   "CherchPices3.frx":08CA
         Stretch         =   -1  'True
         Top             =   150
         Width           =   2055
      End
   End
   Begin VB.CommandButton CommandButton2 
      Caption         =   "Annuler"
      Height          =   495
      Left            =   9840
      TabIndex        =   1
      Top             =   11640
      Width           =   3015
   End
   Begin VB.CommandButton CommandButton1 
      Caption         =   "Valider"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   11760
      Width           =   3135
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   12360
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin OWC.Spreadsheet Spreadsheet1 
      Height          =   10110
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   16890
      HTMLURL         =   ""
      HTMLData        =   $"CherchPices3.frx":4C71
      DataType        =   "HTMLDATA"
      AutoFit         =   -1  'True
      DisplayColHeaders=   0   'False
      DisplayGridlines=   -1  'True
      DisplayHorizontalScrollBar=   -1  'True
      DisplayRowHeaders=   0   'False
      DisplayTitleBar =   0   'False
      DisplayToolbar  =   0   'False
      DisplayVerticalScrollBar=   -1  'True
      EnableAutoCalculate=   -1  'True
      EnableEvents    =   -1  'True
      MoveAfterReturn =   0   'False
      MoveAfterReturnDirection=   0
      RightToLeft     =   0   'False
      ViewableRange   =   "1:65536"
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   195
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "Archive"
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   540
   End
   Begin VB.Label Label28 
      Caption         =   "Légendes :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Noquite As Boolean
Private Sub CommandButton1_Click()
Dim Msg As String
Dim Myrange
Dim IndiceProjet As Long
Dim Id_Pieces As Long
Dim IdProjet As Long
Dim Sql As String
Dim LibPice As String
Dim Rs As Recordset
Dim IdPIndice As Long
If MsgBox("vous réimporter les enregistrements Archivés", vbYesNo + vbQuestion, "Importer Archives") = vbNo Then Exit Sub
RazFiltreEditExcel Me.Spreadsheet1
Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
ProgressBar1.Value = 0
ProgressBar1.Max = Myrange.Rows.Count
For i = 2 To Myrange.Rows.Count
ProgressBar1.Value = i
    If UCase(Myrange(i, 1)) <> 0 Then
         IndiceProjet = CInt(Me.Spreadsheet1.Cells(i, 14))
         Sql = "SELECT   Archive_T_indiceProjet.Id "
        Sql = Sql & "FROM  Archive_T_indiceProjet "
        Sql = Sql & "WHERE [PI] & '_' & [PI_Indice]='" & Replace(Me.Spreadsheet1.Cells(i, 6), " ", "") & "'"
        Set Rs = Con.OpenRecordSet(Sql)
         If Rs.EOF = False Then
            IdPIndice = Rs!Id
         Else
         IdPIndice = 0
         End If
         
        Sql = "SELECT  Archive_T_indiceProjet.Id_Pieces FROM Archive_T_indiceProjet "
        Sql = Sql & "WHERE Archive_T_indiceProjet.Id=" & IndiceProjet & ";"
        Set Rs = Con.OpenRecordSet(Sql)
        If Rs.EOF = False Then
            Id_Pieces = Rs!Id_Pieces
            
             Sql = "SELECT  Archive_T_Pieces.IdProjet,Archive_T_Pieces.Description  FROM Archive_T_Pieces "
            Sql = Sql & "WHERE Archive_T_Pieces.Id=" & Id_Pieces & ";"
             Set Rs = Con.OpenRecordSet(Sql)
              If Rs.EOF = False Then
                    IdProjet = Rs!IdProjet
                    LibPice = "" & Rs!Description
           End If
            

            
            Sql = "SELECT T_Projet.id FROM T_Projet "
            Sql = Sql & "WHERE T_Projet.id=" & IdProjet & ";"
            Set Rs = Con.OpenRecordSet(Sql)
            If Rs.EOF = True Then
                Sql = "INSERT INTO T_Projet ( id, Projet, Description, CleAc ) "
            Sql = Sql & "SELECT Archive_T_Projet.id, Archive_T_Projet.Projet, "
            Sql = Sql & "Archive_T_Projet.Description, Archive_T_Projet.CleAc "
            Sql = Sql & "FROM Archive_T_Projet "
            Sql = Sql & "WHERE Archive_T_Projet.id=" & IdProjet & ";"
            Con.Exequte Sql
            End If
            
            
              
            Sql = "SELECT T_Pieces.Id FROM T_Pieces "
            Sql = Sql & "WHERE T_Pieces.Id=" & Id_Pieces & ";"
             Set Rs = Con.OpenRecordSet(Sql)
            If Rs.EOF = True Then
                Sql = "INSERT INTO T_Pieces ( id, IdProjet, Description ) "
                Sql = Sql & "SELECT Archive_T_Pieces.Id, Archive_T_Pieces.IdProjet, Archive_T_Pieces.Description "
                Sql = Sql & "FROM Archive_T_Pieces "
                Sql = Sql & "WHERE Archive_T_Pieces.Id=" & Id_Pieces & ";"
                Con.Exequte Sql

            End If
            
           
            
            Sql = "SELECT T_indiceProjet.Id FROM T_indiceProjet "
            Sql = Sql & "WHERE T_indiceProjet.Id=" & IndiceProjet & ";"
            Set Rs = Con.OpenRecordSet(Sql)
            If Rs.EOF = True Then
                Sql = "INSERT INTO T_indiceProjet ( Id, Id_Pieces, Description, PL, PL_Indice,  "
                Sql = Sql & "OU, OU_Indice, Li, LI_Indice, PI, IdStatus, PI_Indice, IdStatusSave,  "
                Sql = Sql & "IdApprobateur, PlAutoCadSaveAs, PlAutoCadSave, NbErr, Client, Destinataire,  "
                Sql = Sql & "Service, DessineDate, DessineNOM, VerifieDate, VerifieNom, ApprouveDate,  "
                Sql = Sql & "ApprouveNom, Responsable, Vague, Equipement, RefPF, Ensemble, CleAc, RefP,  "
                Sql = Sql & "Masse, LiAutoCadSaveAs, LiAutoCadSave, OuAutoCadSaveAs, OuAutoCadSave,  "
                Sql = Sql & "Archiver, Cartouche, Version, Pere,NbCartouche ) "
                Sql = Sql & "SELECT Archive_T_indiceProjet.Id,Archive_T_indiceProjet.Id_Pieces,  "
                Sql = Sql & "Archive_T_indiceProjet.Description, Archive_T_indiceProjet.PL,  "
                Sql = Sql & "Archive_T_indiceProjet.PL_Indice, Archive_T_indiceProjet.[OU],  "
                Sql = Sql & "Archive_T_indiceProjet.OU_Indice, Archive_T_indiceProjet.Li,  "
                Sql = Sql & "Archive_T_indiceProjet.LI_Indice, Archive_T_indiceProjet.PI,  "
                Sql = Sql & "Archive_T_indiceProjet.IdStatus, Archive_T_indiceProjet.PI_Indice,  "
                Sql = Sql & "Archive_T_indiceProjet.IdStatusSave, Archive_T_indiceProjet.IdApprobateur,  "
                Sql = Sql & "Archive_T_indiceProjet.PlAutoCadSaveAs, Archive_T_indiceProjet.PlAutoCadSave,  "
                Sql = Sql & "Archive_T_indiceProjet.NbErr, Archive_T_indiceProjet.Client,  "
                Sql = Sql & "Archive_T_indiceProjet.Destinataire, Archive_T_indiceProjet.Service,  "
                Sql = Sql & "Archive_T_indiceProjet.DessineDate, Archive_T_indiceProjet.DessineNOM,  "
                Sql = Sql & "Archive_T_indiceProjet.VerifieDate, Archive_T_indiceProjet.VerifieNom,  "
                Sql = Sql & "Archive_T_indiceProjet.ApprouveDate, Archive_T_indiceProjet.ApprouveNom,  "
                Sql = Sql & "Archive_T_indiceProjet.Responsable, Archive_T_indiceProjet.Vague,  "
                Sql = Sql & "Archive_T_indiceProjet.Equipement, Archive_T_indiceProjet.RefPF,  "
                Sql = Sql & "Archive_T_indiceProjet.Ensemble , Archive_T_indiceProjet.CleAc,  "
                Sql = Sql & "Archive_T_indiceProjet.RefP, Archive_T_indiceProjet.Masse,  "
                Sql = Sql & "Archive_T_indiceProjet.LiAutoCadSaveAs, Archive_T_indiceProjet.LiAutoCadSave,  "
                Sql = Sql & "Archive_T_indiceProjet.OuAutoCadSaveAs, Archive_T_indiceProjet.OuAutoCadSave,  "
                Sql = Sql & "Archive_T_indiceProjet.Archiver, Archive_T_indiceProjet.Cartouche,  "
                Sql = Sql & "Archive_T_indiceProjet.Version, Archive_T_indiceProjet.Pere,Archive_T_indiceProjet.NbCartouche "
                Sql = Sql & "FROM Archive_T_indiceProjet "
                Sql = Sql & "WHERE Archive_T_indiceProjet.Id=" & IndiceProjet & ";"
                Con.Exequte Sql
                If IdPIndice = 0 Then
                    Sql = "UPDATE T_indiceProjet SET T_indiceProjet.Archiver = False "
                    Sql = Sql & "WHERE T_indiceProjet.Id=" & IndiceProjet & ";"
                    Con.Exequte Sql
                End If
            End If
            
              Sql = "SELECT T_Critères.Id FROM T_Critères "
            Sql = Sql & "WHERE T_Critères.Id=" & Id_Pieces & ";"
             Set Rs = Con.OpenRecordSet(Sql)
            If Rs.EOF = True Then
            Sql = "INSERT INTO T_Critères SELECT Archive_T_Critères.* FROM Archive_T_Critères "
                Sql = Sql & "WHERE Archive_T_Critères.Id_IndiceProjet=" & IndiceProjet & ";"
            
'                Sql = "INSERT INTO T_Critères "
'                Sql = Sql & "SELECT Archive_T_Critères.* "
'                Sql = Sql & "FROM Archive_T_Critères "
'                Sql = Sql & "WHERE Archive_T_Critères.Id_IndiceProjet=" & IndiceProjet & ";"
                Con.Exequte Sql

            End If
            Sql = "SELECT Connecteurs.Id_IndiceProjet FROM Connecteurs "
            Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & IndiceProjet & ";"
             Set Rs = Con.OpenRecordSet(Sql)
            If Rs.EOF = True Then
            
            Sql = "INSERT INTO Connecteurs SELECT Archive_Connecteurs.* FROM Archive_Connecteurs "
                Sql = Sql & "WHERE Archive_Connecteurs.Id_IndiceProjet=" & IndiceProjet & ";"
            
'                Sql = "INSERT INTO Connecteurs ( Numéro, Id_IndiceProjet, CONNECTEUR,  "
'                Sql = Sql & "[O/N], DESIGNATION, CODE_APP, N°, POS, [POS-OUT], PRECO1, PRECO2, [100%] ) "
'                Sql = Sql & "SELECT Archive_Connecteurs.Numéro, Archive_Connecteurs.Id_IndiceProjet,  "
'                Sql = Sql & "Archive_Connecteurs.CONNECTEUR,Archive_Connecteurs.[O/N],  "
'                Sql = Sql & "Archive_Connecteurs.DESIGNATION, Archive_Connecteurs.CODE_APP,  "
'                Sql = Sql & "Archive_Connecteurs.N°, Archive_Connecteurs.POS, Archive_Connecteurs.[POS-OUT],  "
'                Sql = Sql & "Archive_Connecteurs.PRECO1, Archive_Connecteurs.PRECO2, Archive_Connecteurs.[100%] "
'                Sql = Sql & "FROM Archive_Connecteurs "
'                Sql = Sql & "WHERE Archive_Connecteurs.Id_IndiceProjet=" & IndiceProjet & ";"
                Con.Exequte Sql
            End If
            Sql = "SELECT Ligne_Tableau_fils.Id_IndiceProjet FROM Ligne_Tableau_fils "
            Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IndiceProjet & ";"
            Set Rs = Con.OpenRecordSet(Sql)
            If Rs.EOF = True Then
            
              Sql = "INSERT INTO Ligne_Tableau_fils SELECT Archive_Ligne_Tableau_fils.* FROM Archive_Ligne_Tableau_fils "
                Sql = Sql & "WHERE Archive_Ligne_Tableau_fils.Id_IndiceProjet=" & IndiceProjet & ";"
            
'                Sql = "INSERT INTO Ligne_Tableau_fils ( Numéro, Id_IndiceProjet, LIAI, DESIGNATION, FIL,  "
'                Sql = Sql & "SECT, TEINT, TEINT2, ISO, [LONG], [LONG CP], COUPE, POS, [POS-OUT], FA, APP,  "
'                Sql = Sql & "VOI, POS2, [POS-OUT2], FA2, APP2, VOI2, PRECO, [OPTION] ) "
'                Sql = Sql & "SELECT Archive_Ligne_Tableau_fils.Numéro,  "
'                Sql = Sql & "Archive_Ligne_Tableau_fils.Id_IndiceProjet,Archive_Ligne_Tableau_fils.LIAI,  "
'                Sql = Sql & "Archive_Ligne_Tableau_fils.DESIGNATION, Archive_Ligne_Tableau_fils.FIL, "
'                Sql = Sql & "Archive_Ligne_Tableau_fils.SECT, Archive_Ligne_Tableau_fils.TEINT,  "
'                Sql = Sql & "Archive_Ligne_Tableau_fils.TEINT2, Archive_Ligne_Tableau_fils.ISO,  "
'                Sql = Sql & "Archive_Ligne_Tableau_fils.LONG, Archive_Ligne_Tableau_fils.[LONG CP],  "
'                Sql = Sql & "Archive_Ligne_Tableau_fils.COUPE, Archive_Ligne_Tableau_fils.POS,  "
'                Sql = Sql & "Archive_Ligne_Tableau_fils.[POS-OUT], Archive_Ligne_Tableau_fils.FA,  "
'                Sql = Sql & "Archive_Ligne_Tableau_fils.APP, Archive_Ligne_Tableau_fils.VOI,  "
'                Sql = Sql & "Archive_Ligne_Tableau_fils.POS2, Archive_Ligne_Tableau_fils.[POS-OUT2],  "
'                Sql = Sql & "Archive_Ligne_Tableau_fils.FA2, Archive_Ligne_Tableau_fils.APP2,  "
'                Sql = Sql & "Archive_Ligne_Tableau_fils.VOI2, Archive_Ligne_Tableau_fils.PRECO,  "
'                Sql = Sql & "Archive_Ligne_Tableau_fils.OPTION "
'                Sql = Sql & "FROM Archive_Ligne_Tableau_fils "
'                Sql = Sql & "WHERE Archive_Ligne_Tableau_fils.Id_IndiceProjet=" & IndiceProjet & ";"
            Con.Exequte Sql
            End If
            
             Sql = "SELECT Composants.Id_IndiceProjet FROM Composants "
             Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & IndiceProjet & ";"
             Set Rs = Con.OpenRecordSet(Sql)
            If Rs.EOF = True Then
            
            Sql = "INSERT INTO Composants SELECT Archive_Composants.* FROM Archive_Composants "
                Sql = Sql & "WHERE Archive_Composants.Id_IndiceProjet=" & IndiceProjet & ";"
'
'                Sql = "INSERT INTO Composants ( Id, Id_IndiceProjet, DESIGNCOMP, NUMCOMP, REFCOMP, Path ) "
'                Sql = Sql & "SELECT Archive_Composants.Id, Archive_Composants.Id_IndiceProjet,  "
'                Sql = Sql & "Archive_Composants.DESIGNCOMP, Archive_Composants.NUMCOMP,  "
'                Sql = Sql & "Archive_Composants.REFCOMP, Archive_Composants.Path "
'                Sql = Sql & "FROM Archive_Composants "
'                Sql = Sql & "WHERE Archive_Composants.Id_IndiceProjet=" & IndiceProjet & ";"
            Con.Exequte Sql

            End If
            
            Sql = "SELECT Nota.Id_IndiceProjet FROM Nota "
            Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & IndiceProjet & ";"
             Set Rs = Con.OpenRecordSet(Sql)
            If Rs.EOF = True Then
            
             Sql = "INSERT INTO Nota SELECT Archive_Nota.* FROM Archive_Nota "
                Sql = Sql & "WHERE Archive_Nota.Id_IndiceProjet=" & IndiceProjet & ";"
            
'                Sql = "INSERT INTO Nota ( Id, Id_IndiceProjet, NOTA, NUMNOTA )  "
'                Sql = Sql & "SELECT Archive_Nota.Id, Archive_Nota.Id_IndiceProjet, Archive_Nota.NOTA,   "
'                Sql = Sql & "Archive_Nota.NUMNOTA  "
'                Sql = Sql & "FROM Archive_Nota  "
'                Sql = Sql & "WHERE Archive_Nota.Id_IndiceProjet=" & IndiceProjet & ";"
                Con.Exequte Sql

            End If
            Sql = "SELECT T_Noeuds.Id_IndiceProjet FROM T_Noeuds  "
                Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & IndiceProjet & ";"
              Set Rs = Con.OpenRecordSet(Sql)
        If Rs.EOF = True Then
                Sql = "INSERT INTO T_Noeuds SELECT Archive_T_Noeuds.*  "
                Sql = Sql & "FROM Archive_T_Noeuds "
                Sql = Sql & "WHERE Archive_T_Noeuds.Id_IndiceProjet=" & IndiceProjet & ";"

                Con.Exequte Sql

            End If

             Sql = "DELETE Archive_T_Pieces.*"
                Sql = Sql & "FROM Archive_T_Pieces INNER JOIN Archive_T_indiceProjet  "
                Sql = Sql & "ON Archive_T_Pieces.Id = Archive_T_indiceProjet.Id_Pieces "
                Sql = Sql & "WHERE Archive_T_indiceProjet.Id=" & IndiceProjet & ";"
                Con.Exequte Sql
                        
        End If
    End If
Next i
Sql = "DELETE Archive_T_Projet.*, Archive_T_Pieces.Id "
            Sql = Sql & "FROM Archive_T_Projet LEFT JOIN Archive_T_Pieces  "
            Sql = Sql & "ON Archive_T_Projet.id = Archive_T_Pieces.IdProjet "
            Sql = Sql & "WHERE Archive_T_Pieces.Id Is Null;"
            Con.Exequte Sql
Sql = "DELETE Archive_T_Pieces.* "
Sql = Sql & "FROM Archive_T_Pieces LEFT JOIN Archive_T_indiceProjet ON Archive_T_Pieces.Id = Archive_T_indiceProjet.Id_Pieces  "
Sql = Sql & "WHERE Archive_T_indiceProjet.Id Is Null;"
Con.Exequte Sql

 Noquite = False
Me.Hide
End Sub

Public Sub Charge(MyForm As Object, Optional Filtre As String, Optional boolTxt As Boolean, Optional boolArchive As Boolean)
Dim Sql As String
Dim Rs As Recordset
Dim IndexRow As Long
Dim IndexCol As Long

boolTxts = boolTxt
IndexRow = 1
IndexCol = 0
Me.Spreadsheet1.Columns(1).Locked = False

 Me.Spreadsheet1.Columns(1).NumberFormat = "Yes/No"
 Me.Spreadsheet1.Columns(1).Locked = True
  
Sql = "SELECT 0 AS Importer ,Archive_SelectProjets.* "
Sql = Sql & "FROM Archive_SelectProjets; "
Set Rs = Con.OpenRecordSet(Sql)
Rs.Filter = Filtre
Set MyFormCible = MyForm
While Rs.EOF = False
IndexRow = IndexRow + 1
For IndexCol = 0 To Rs.Fields.Count - 11
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Locked = False
    If IndexCol > 2 And IndexCol < 7 Then
         Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1) = Trim("'" & Rs.Fields(IndexCol)) & Trim("_" & Rs.Fields(IndexCol + 17))
    Else
        Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1) = Trim("'" & Rs.Fields(IndexCol))
    End If
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Interior.Color = &HFFC0FF
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Locked = MyLocked(Me.Spreadsheet1.Cells(1, IndexCol + 1), Rs.Fields(Rs.Fields.Count - 11))

Next IndexCol

Rs.MoveNext
Wend

Set Rs = Con.CloseRecordSet(Rs)

Dim Myrange
Set Myrange = Me.Spreadsheet1.Range("A1").CurrentRegion
Myrange.AutoFitColumns
Set Myrange = Nothing

Me.Show vbModal
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


