VERSION 5.00
Object = "{0002E540-0000-0000-C000-000000000046}#1.1#0"; "MSOWC.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form UserForm4 
   Caption         =   "Supprimer/Archiver:"
   ClientHeight    =   11715
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   18840
   Icon            =   "CherchPices2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11715
   ScaleWidth      =   18840
   StartUpPosition =   3  'Windows Default
   Begin OWC.Spreadsheet Spreadsheet1 
      Height          =   9360
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   18840
      HTMLURL         =   ""
      HTMLData        =   $"CherchPices2.frx":08CA
      DataType        =   "HTMLDATA"
      AutoFit         =   0   'False
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
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   16560
      TabIndex        =   11
      Top             =   0
      Width           =   2175
      Begin VB.Image Image1 
         Height          =   855
         Left            =   40
         Picture         =   "CherchPices2.frx":1BAA
         Stretch         =   -1  'True
         Top             =   150
         Width           =   2055
      End
   End
   Begin VB.CommandButton CommandButton2 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   10080
      TabIndex        =   2
      Top             =   10800
      Width           =   3135
   End
   Begin VB.CommandButton CommandButton1 
      Caption         =   "Valider"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   10800
      Width           =   3135
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   -240
      TabIndex        =   10
      Top             =   11400
      Width           =   19095
      _ExtentX        =   33681
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   195
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   795
      TabIndex        =   8
      Top             =   720
      Width           =   195
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   1485
      TabIndex        =   7
      Top             =   720
      Width           =   195
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "CRE"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   330
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "MOD"
      Height          =   195
      Index           =   1
      Left            =   1050
      TabIndex        =   5
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "VAL"
      Height          =   195
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   300
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Noquite As Boolean
Public Sub Charge(MyForm As Object, Optional Filtre As String, Optional boolTxt As Boolean, Optional boolArchive As Boolean)
Dim Sql As String
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

Sql = "SELECT  0 AS Suprimers  , 0 AS Archivers,SelectProjets.* "
Sql = Sql & "FROM SelectProjets; "
Set Rs = Con.OpenRecordSet(Sql)
Rs.Filter = Filtre
Set MyFormCible = MyForm
While Rs.EOF = False
IndexRow = IndexRow + 1
For IndexCol = 0 To Rs.Fields.Count - 12
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Locked = False
    If IndexCol > 3 And IndexCol < 8 Then
         Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1) = Trim("'" & Rs.Fields(IndexCol)) & Trim("_" & Rs.Fields(IndexCol + 18))
    Else
        Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1) = Trim("'" & Rs.Fields(IndexCol))
    End If
    Me.Spreadsheet1.Cells(IndexRow, IndexCol + 1).Interior.Color = ChoixCouleur(Rs.Fields(Rs.Fields.Count - 11))
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
If Mytype = "Supprimer O/N" And Statues = 1 Then MyLocked = False
If Mytype = "Supprimer O/N" And Statues = 2 Then MyLocked = False
If Mytype = "Supprimer O/N" And Statues = 3 Then MyLocked = False
If Mytype = "Archiver O/N" And Statues = 3 Then MyLocked = False

End Function

Private Sub Command1_Click()

End Sub

Private Sub CommandButton1_Click()
Dim Msg As String
Dim Myrange
Dim IndiceProjet As Long
Dim Id_Pieces As Long
Dim IdProjet As Long
Dim Sql As String
Dim LibPice As String
Dim Rs As Recordset
Msg = "Attention tous les enregistrements supprimés seront définitivement perdus." & vbCrLf & vbCrLf
Msg = Msg & "Tous les enregistrements archivés pourront être réinsérés par la suite" & vbCrLf & vbCrLf
Msg = Msg & "Voulez-vous continuer." & vbCrLf & vbCrLf
If MsgBox(Msg, vbYesNo + vbQuestion, "Supprimer/Archiver") = vbNo Then Exit Sub
RazFiltreEditExcel Me.Spreadsheet1
Set Myrange = Me.Spreadsheet1.ActiveSheet.Range("a1").CurrentRegion
ProgressBar1.Value = 0
ProgressBar1.Max = Myrange.Rows.Count
For i = 2 To Myrange.Rows.Count
ProgressBar1.Value = i
    If UCase(Myrange(i, 2)) <> 0 Then
        IndiceProjet = CInt(Me.Spreadsheet1.Cells(i, 15))
        Sql = "SELECT  T_indiceProjet.Id_Pieces FROM T_indiceProjet "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & IndiceProjet & ";"
        Set Rs = Con.OpenRecordSet(Sql)
        If Rs.EOF = False Then
            Id_Pieces = Rs!Id_Pieces
            
             Sql = "SELECT  T_Pieces.IdProjet,T_Pieces.Description  FROM T_Pieces "
            Sql = Sql & "WHERE T_Pieces.Id=" & Id_Pieces & ";"
             Set Rs = Con.OpenRecordSet(Sql)
              If Rs.EOF = False Then
                    IdProjet = Rs!IdProjet
                    LibPice = "" & Rs!Description
                    Sql = "SELECT Archive_T_Projet.id FROM Archive_T_Projet "
                    Sql = Sql & "WHERE Archive_T_Projet.id=" & IdProjet & ";"
                     Set Rs = Con.OpenRecordSet(Sql)
                     If Rs.EOF = True Then
                        Sql = "INSERT INTO Archive_T_Projet ( id, Projet, Description, CleAc ) "
                        Sql = Sql & "SELECT T_Projet.id, T_Projet.Projet, T_Projet.Description, T_Projet.CleAc "
                        Sql = Sql & "FROM T_Projet "
                   Sql = Sql & "WHERE T_Projet.id=" & IdProjet & ";"

                        Con.Exequte Sql

                     End If
                     Sql = "SELECT Archive_T_Pieces.Id FROM Archive_T_Pieces "
                     Sql = Sql & "WHERE Archive_T_Pieces.Id=" & Id_Pieces & ";"
                        Set Rs = Con.OpenRecordSet(Sql)
                     If Rs.EOF = True Then
                        Sql = "INSERT INTO Archive_T_Pieces ( id, IdProjet, Description ) "
                        Sql = Sql & "SELECT T_Pieces.Id, T_Pieces.IdProjet, T_Pieces.Description "
                        Sql = Sql & "FROM T_Pieces "
                        Sql = Sql & "WHERE T_Pieces.Id=" & Id_Pieces & ";"
                        Con.Exequte Sql
                     End If
                     Sql = "SELECT Archive_T_indiceProjet.Id FROM Archive_T_indiceProjet "
                     Sql = Sql & "WHERE Archive_T_indiceProjet.Id=" & IndiceProjet & ";"
                       Set Rs = Con.OpenRecordSet(Sql)
                     If Rs.EOF = True Then
                     
                     
                        Sql = "INSERT INTO Archive_T_indiceProjet  "
                        Sql = Sql & "( Id, Id_Pieces, Description, PL, PL_Indice, OU, OU_Indice,  "
                        Sql = Sql & "Li, LI_Indice, PI, IdStatus, PI_Indice, IdStatusSave,  "
                        Sql = Sql & "IdApprobateur, PlAutoCadSaveAs, PlAutoCadSave, NbErr, Client,  "
                        Sql = Sql & "Destinataire, Service, DessineDate, DessineNOM, VerifieDate,  "
                        Sql = Sql & "VerifieNom, ApprouveDate, ApprouveNom, Responsable, Vague,  "
                        Sql = Sql & "Equipement, RefPF, Ensemble, CleAc, RefP, Masse, LiAutoCadSaveAs,  "
                        Sql = Sql & "LiAutoCadSave, OuAutoCadSaveAs, OuAutoCadSave, Archiver, Cartouche,  "
                        Sql = Sql & "Version, Pere,NbCartouche) "
                        Sql = Sql & "SELECT T_indiceProjet.Id, T_indiceProjet.Id_Pieces,  "
                        Sql = Sql & "T_indiceProjet.Description, T_indiceProjet.PL, "
                        Sql = Sql & "T_indiceProjet.PL_Indice, T_indiceProjet.[OU],  "
                        Sql = Sql & "T_indiceProjet.OU_Indice, T_indiceProjet.Li,  "
                        Sql = Sql & "T_indiceProjet.LI_Indice, T_indiceProjet.PI,  "
                        Sql = Sql & "T_indiceProjet.IdStatus, T_indiceProjet.PI_Indice,  "
                        Sql = Sql & "T_indiceProjet.IdStatusSave, T_indiceProjet.IdApprobateur,  "
                        Sql = Sql & "T_indiceProjet.PlAutoCadSaveAs, T_indiceProjet.PlAutoCadSave, "
                        Sql = Sql & " T_indiceProjet.NbErr, T_indiceProjet.Client,  "
                        Sql = Sql & "T_indiceProjet.Destinataire, T_indiceProjet.Service, "
                        Sql = Sql & " T_indiceProjet.DessineDate, T_indiceProjet.DessineNOM, "
                        Sql = Sql & " T_indiceProjet.VerifieDate, T_indiceProjet.VerifieNom, "
                        Sql = Sql & " T_indiceProjet.ApprouveDate, T_indiceProjet.ApprouveNom, "
                        Sql = Sql & " T_indiceProjet.Responsable, T_indiceProjet.Vague, "
                        Sql = Sql & " T_indiceProjet.Equipement, T_indiceProjet.RefPF, "
                        Sql = Sql & " T_indiceProjet.Ensemble, T_indiceProjet.CleAc, "
                        Sql = Sql & " T_indiceProjet.RefP, T_indiceProjet.Masse, "
                        Sql = Sql & " T_indiceProjet.LiAutoCadSaveAs, T_indiceProjet.LiAutoCadSave, "
                        Sql = Sql & " T_indiceProjet.OuAutoCadSaveAs, T_indiceProjet.OuAutoCadSave, "
                        Sql = Sql & " T_indiceProjet.Archiver, T_indiceProjet.Cartouche ,  "
                        Sql = Sql & "T_indiceProjet.Version, T_indiceProjet.Pere,T_indiceProjet.NbCartouche "
                        
                        Sql = Sql & "FROM T_indiceProjet "
                        Sql = Sql & "WHERE T_indiceProjet.Id=" & IndiceProjet & ";"
                        Con.Exequte Sql

                     End If
                     Sql = "SELECT Archive_Connecteurs.Id_IndiceProjet FROM Archive_Connecteurs "
                     Sql = Sql & "WHERE Archive_Connecteurs.Id_IndiceProjet=" & IndiceProjet & ";"
                    Set Rs = Con.OpenRecordSet(Sql)
                    If Rs.EOF = True Then
                    
                     Sql = "INSERT INTO Archive_Connecteurs SELECT Connecteurs.* FROM Connecteurs "
                            Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & IndiceProjet & ";"
                            
                            
'                        Sql = "INSERT INTO Archive_Connecteurs ( Numéro, Id_IndiceProjet, CONNECTEUR, "
'                        Sql = Sql & "[O/N], DESIGNATION, CODE_APP, N°, POS, [POS-OUT], PRECO1, PRECO2, [100%] )"
'                        Sql = Sql & "SELECT Connecteurs.Numéro, Connecteurs.Id_IndiceProjet, "
'                        Sql = Sql & "Connecteurs.CONNECTEUR, Connecteurs.[O/N], "
'                        Sql = Sql & "Connecteurs.DESIGNATION, Connecteurs.CODE_APP, Connecteurs.N°, "
'                        Sql = Sql & "Connecteurs.POS, Connecteurs.[POS-OUT], Connecteurs.PRECO1, "
'                        Sql = Sql & "Connecteurs.PRECO2, Connecteurs.[100%]"
'                        Sql = Sql & "FROM Connecteurs "
'                        Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & IndiceProjet & ";"
                        Con.Exequte Sql
                    End If
                        Sql = "SELECT Archive_Ligne_Tableau_fils.Id_IndiceProjet FROM Archive_Ligne_Tableau_fils "
                        Sql = Sql & "WHERE Archive_Ligne_Tableau_fils.Id_IndiceProjet=" & IndiceProjet & ";"
                      
                        Set Rs = Con.OpenRecordSet(Sql)
                        If Rs.EOF = True Then
                        
                         Sql = "INSERT INTO Archive_Ligne_Tableau_fils SELECT Ligne_Tableau_fils.* FROM Ligne_Tableau_fils "
                            Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IndiceProjet & ";"
                            
'
'                            Sql = "INSERT INTO Archive_Ligne_Tableau_fils ( Numéro, Id_IndiceProjet, LIAI, "
'                        Sql = Sql & "DESIGNATION, FIL, SECT, TEINT, TEINT2, ISO, [LONG], [LONG CP], COUPE,  "
'                        Sql = Sql & "POS, [POS-OUT], FA, APP, VOI, POS2, [POS-OUT2], FA2, APP2, VOI2,  "
'                        Sql = Sql & "PRECO, [OPTION] ) "
'                        Sql = Sql & "SELECT Ligne_Tableau_fils.Numéro, Ligne_Tableau_fils.Id_IndiceProjet,  "
'                        Sql = Sql & "Ligne_Tableau_fils.LIAI, Ligne_Tableau_fils.DESIGNATION,  "
'                        Sql = Sql & "Ligne_Tableau_fils.FIL, Ligne_Tableau_fils.SECT,  "
'                        Sql = Sql & "Ligne_Tableau_fils.TEINT, Ligne_Tableau_fils.TEINT2,  "
'                        Sql = Sql & "Ligne_Tableau_fils.ISO, Ligne_Tableau_fils.LONG,  "
'                        Sql = Sql & "Ligne_Tableau_fils.[LONG CP], Ligne_Tableau_fils.COUPE,  "
'                        Sql = Sql & "Ligne_Tableau_fils.POS, Ligne_Tableau_fils.[POS-OUT],  "
'                        Sql = Sql & "Ligne_Tableau_fils.FA, Ligne_Tableau_fils.APP,  "
'                        Sql = Sql & "Ligne_Tableau_fils.VOI, Ligne_Tableau_fils.POS2,  "
'                        Sql = Sql & "Ligne_Tableau_fils.[POS-OUT2], Ligne_Tableau_fils.FA2,  "
'                        Sql = Sql & "Ligne_Tableau_fils.APP2, Ligne_Tableau_fils.VOI2,  "
'                        Sql = Sql & "Ligne_Tableau_fils.PRECO, Ligne_Tableau_fils.OPTION "
'                        Sql = Sql & "FROM Ligne_Tableau_fils "
'                        Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IndiceProjet & ";"
                        Con.Exequte Sql

                        End If
                        
                         Sql = "SELECT Archive_Composants.Id_IndiceProjet FROM Archive_Composants "
                        Sql = Sql & "WHERE Archive_Composants.Id_IndiceProjet=" & IndiceProjet & ";"
                      
                        Set Rs = Con.OpenRecordSet(Sql)
                        If Rs.EOF = True Then
                        
                        Sql = "INSERT INTO Archive_Composants SELECT Composants.* FROM Composants "
                            Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & IndiceProjet & ";"
                        
'                            Sql = "INSERT INTO Archive_Composants ( Id, Id_IndiceProjet, DESIGNCOMP,  "
'                            Sql = Sql & "NUMCOMP, REFCOMP, Path ) "
'                            Sql = Sql & "SELECT Composants.Id, Composants.Id_IndiceProjet,  "
'                            Sql = Sql & "Composants.DESIGNCOMP, Composants.NUMCOMP, Composants.REFCOMP,  "
'                            Sql = Sql & "Composants.Path "
'                            Sql = Sql & "FROM Composants "
'                            Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & IndiceProjet & ";"
                            Con.Exequte Sql
                        End If
                        
                        Sql = "SELECT Archive_Nota.Id_IndiceProjet FROM Archive_Nota "
                        Sql = Sql & "WHERE Archive_Nota.Id_IndiceProjet=" & IndiceProjet & ";"
                      
                        Set Rs = Con.OpenRecordSet(Sql)
                        If Rs.EOF = True Then
                            
                        Sql = "INSERT INTO Archive_Nota SELECT Nota.* FROM Nota "
                            Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & IndiceProjet & ";"
                            
'                            Sql = "INSERT INTO Archive_Nota ( Id, Id_IndiceProjet, NOTA, NUMNOTA ) "
'                            Sql = Sql & "SELECT Nota.Id, Nota.Id_IndiceProjet, Nota.NOTA, Nota.NUMNOTA "
'                            Sql = Sql & "FROM Nota "
'                            Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & IndiceProjet & ";"
                            Con.Exequte Sql
            
                        End If
                        
                        Sql = "SELECT Archive_T_Critères.Id_IndiceProjet FROM Archive_T_Critères "
                        Sql = Sql & "WHERE Archive_T_Critères.Id_IndiceProjet=" & IndiceProjet & ";"
                      
                        Set Rs = Con.OpenRecordSet(Sql)
                        If Rs.EOF = True Then
                        
                        Sql = "INSERT INTO Archive_T_Critères SELECT T_Critères.* FROM T_Critères "
                            Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & IndiceProjet & ";"
                        
'                            Sql = "INSERT INTO Archive_T_Critères ( Id, Id_IndiceProjet, CODE_CRITERE, CRITERES ) "
'                            Sql = Sql & "SELECT T_Critères.Id, T_Critères.Id_IndiceProjet, T_Critères.CODE_CRITERE, T_Critères.CRITERES "
'                            Sql = Sql & "FROM T_Critères "
'                            Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & IndiceProjet & ";"
                            Con.Exequte Sql
            
                        End If
                        Sql = "SELECT Archive_T_Noeuds.Id_IndiceProjet FROM Archive_T_Noeuds "
                            Sql = Sql & "WHERE Archive_T_Noeuds.Id_IndiceProjet=" & IndiceProjet & ";"

                        Set Rs = Con.OpenRecordSet(Sql)
                        If Rs.EOF = True Then
                           Sql = "INSERT INTO Archive_T_Noeuds SELECT T_Noeuds.* FROM T_Noeuds "
                            Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & IndiceProjet & ";"
                            Con.Exequte Sql
                        End If
                        Sql = "DELETE T_Pieces.*, T_Pieces.Id FROM T_Pieces "
                        Sql = Sql & "WHERE T_Pieces.Id=" & Id_Pieces & ";"
                        
                        Sql = "DELETE  T_indiceProjet.*  "
                        Sql = Sql & "FROM T_Pieces INNER JOIN T_indiceProjet ON  "
                        Sql = Sql & "T_Pieces.Id = T_indiceProjet.Id_Pieces "
                        Sql = Sql & "WHERE T_indiceProjet.Id=" & IndiceProjet & ";"

                        Con.Exequte Sql
'                        Sql = "SELECT T_Pieces.Id FROM T_Pieces "
'                        Sql = Sql & "WHERE T_Pieces.IdProjet=" & IdProjet & ";"
'                        Set Rs = Con.OpenRecordSet(Sql)
'                        If Rs.EOF = True Then
'                            Sql = "DELETE T_Projet.* FROM T_Projet "
'                            Sql = Sql & "WHERE T_Projet.id=" & IdProjet & ";"
'                            Con.Exequte Sql
'
'                        End If
                       
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
    If UCase(Myrange(i, 1)) <> 0 Then
         IndiceProjet = CInt(Me.Spreadsheet1.Cells(i, 15))
         Sql = "DELETE  T_indiceProjet.*  "
                        Sql = Sql & "FROM T_Pieces INNER JOIN T_indiceProjet ON  "
                        Sql = Sql & "T_Pieces.Id = T_indiceProjet.Id_Pieces "
                        Sql = Sql & "WHERE T_indiceProjet.Id=" & IndiceProjet & ";"

                        Con.Exequte Sql
                        zz = Split(Replace(Me.Spreadsheet1.Cells(i, 7), " ", ""), "_")
                        y = 0
                        Dim yy As Long
                        aa = ""
                        For yy = LBound(zz) To UBound(zz) - 1
                        aa = aa & zz(yy) & "_"
                        Next
                        aa = Left(aa, Len(aa) - 1)
                        Sql = "SELECT T_indiceProjet.Id "
                        Sql = Sql & "FROM T_indiceProjet "
                        Sql = Sql & "WHERE [PI]='" & aa & "' "
                        Sql = Sql & "ORDER BY T_indiceProjet.Id DESC; "
                        Set Rs = Con.OpenRecordSet(Sql)
                        If Rs.EOF = False Then
                            Sql = "UPDATE T_indiceProjet SET T_indiceProjet.Archiver = False "
                            Sql = Sql & "WHERE T_indiceProjet.Id=" & Rs!Id & ";"
                            Con.Exequte Sql
                        End If
       aa = Me.Spreadsheet1.Cells(i, 7)
       
      
    End If
Next i
 Sql = "DELETE T_Projet.*, T_Pieces.Id  "
Sql = Sql & "FROM T_Projet LEFT JOIN T_Pieces ON T_Projet.id = T_Pieces.IdProjet "
Sql = Sql & "WHERE T_Pieces.Id Is Null;"

Con.Exequte Sql
Sql = "DELETE T_Pieces.*, T_indiceProjet.Id  "
Sql = Sql & "FROM T_Pieces LEFT JOIN T_indiceProjet ON T_Pieces.Id = T_indiceProjet.Id_Pieces  "
Sql = Sql & "WHERE T_indiceProjet.Id Is Null;"
Con.Exequte Sql


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


