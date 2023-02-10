VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Job 
   Caption         =   "Form1"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   10080
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   1508
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   8040
      Picture         =   "Job.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label ProgressBar1Caption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "Job"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Id_Pice As Long



Private Sub Form_Activate()
Dim Sql As String
Dim NbErr As Long
Dim IdIndice As Long
Dim I As Long
Dim Equipement
Dim Equipement2
Dim Equipement3
Dim Nomenclature_Appareil As Boolean
Dim Par_Fournisseur As Boolean
Dim Par_Options As Boolean
Dim PreparNomOk As Integer
Dim Action As String
Dim NbPieces As Long
On Error Resume Next
Reprise:
Sql = "SELECT T_Job.* FROM T_Job "
Sql = Sql & "Where  T_Job.Job = " & Command & " "
Sql = Sql & "ORDER BY T_Job.Job;"
Debug.Print Sql
Set RsBarGraph = Con.OpenRecordSet(Sql)
If Err Then
    If NbErr > 10 Then End
    NbErr = NbErr + 1
    GoTo Reprise
End If
DoEvents
If RsBarGraph.EOF = False Then
Me.Caption = "JOB N°:" & RsBarGraph!Job

'On Error Resume Next
'RetournIdApp "acad.exe"
'Set AutoApp = New AutoCAD.AcadApplication
'        If Err = 0 Then
'            sql = "UPDATE T_Job SET T_Job.IdAutocad = " & RetournIdApp("acad.exe", True) & " "
'            sql = sql & "WHERE T_Job.Job=" & RsBarGraph!Job & ";"
'            Con.Execute sql
'
'            AutoApp.Documents(0).Close False
'            'AutoApp.Visible = False
            boolAutoCAD = True
'        Else
'            Err.Clear
''            MsgBox "Plus de licence Autocad disponible", vbInformation, "AutoCâble  licence :"
'            boolAutoCAD = False
'        End If
' On Error GoTo 0
'Con.Execute "UPDATE T_Job SET T_Job.DateDebut = Now() WHERE T_Job.Job=" & RsBarGraph!Job & ;"
Set FormBarGrah = Me
bool_MiseEnPage = True
PreparNomOk = RsBarGraph!PreparNomOk
Action = RsBarGraph!Action
If Action = "Modifier" Then Action = "Modifier Plan"
IdFils = RsBarGraph!Id_Fils
NomenclatureOk = RsBarGraph!NomenclatureOk
bool_Plan_L_Connecteurs = RsBarGraph!Plan_L_Connecteurs
 bool_Plan_L_Fils = RsBarGraph!Plan_L_Fils
 bool_Plan_L_Vignettes = RsBarGraph!Plan_L_Vignettes
 bool_Plan_L_Etiquettes = RsBarGraph!Plan_L_Etiquettes
 bool_Plan_L_Composants = RsBarGraph!Plan_L_Composants
 bool_Plan_L_Notas = RsBarGraph!Plan_L_Notas
 bool_Plan_L_cartouches = RsBarGraph!Plan_L_cartouches
 bool_Plan_L_Preconisations = RsBarGraph!Plan_L_Preconisations
 bool_Plan_L_Options = RsBarGraph!Plan_L_Options
 bool_Plan_L_Criteres = RsBarGraph!Plan_L_Criteres
 bool_Plan_L_Noeuds = RsBarGraph!Plan_L_Noeuds
'
 bool_Plan_E_Connecteurs = RsBarGraph!Plan_E_Connecteurs
 bool_Plan_E_Fils = RsBarGraph!Plan_E_Fils
 bool_Plan_E_Vignettes = RsBarGraph!Plan_E_Vignettes
 bool_Plan_E_Etiquettes = RsBarGraph!Plan_E_Etiquettes
 bool_Plan_E_Composants = RsBarGraph!Plan_E_Composants
 bool_Plan_E_Notas = RsBarGraph!Plan_E_Notas
 bool_Plan_E_cartouches = RsBarGraph!Plan_E_cartouches
 bool_Plan_E_Preconisations = RsBarGraph!Plan_E_Preconisations
 bool_Plan_E_Options = RsBarGraph!Plan_E_Options
 bool_Plan_E_Criteres = RsBarGraph!Plan_E_Criteres
 bool_Plan_E_Noeuds = RsBarGraph!Plan_E_Noeuds
'
'
 bool_Outil_L_Connecteurs = RsBarGraph!Outil_L_Connecteurs
 bool_Outil_L_Fils = RsBarGraph!Outil_L_Fils
 bool_Outil_L_Vignettes = RsBarGraph!Outil_L_Vignettes
 bool_Outil_L_Etiquettes = RsBarGraph!Outil_L_Etiquettes
 bool_Outil_L_Composants = RsBarGraph!Outil_L_Composants
 bool_Outil_L_Notas = RsBarGraph!Outil_L_Notas
 bool_Outil_L_cartouches = RsBarGraph!Outil_L_cartouches
 bool_Outil_L_Preconisations = RsBarGraph!Outil_L_Preconisations
 bool_Outil_L_Options = RsBarGraph!Outil_L_Options
 bool_Outil_L_Criteres = RsBarGraph!Outil_L_Criteres
 bool_Outil_L_Noeuds = RsBarGraph!Outil_L_Noeuds

'
 bool_Outil_E_Connecteurs = RsBarGraph!Outil_E_Connecteurs
 bool_Outil_E_Fils = RsBarGraph!Outil_E_Fils
 bool_Outil_E_Vignettes = RsBarGraph!Outil_E_Vignettes
 bool_Outil_E_Etiquettes = RsBarGraph!Outil_E_Etiquettes
 bool_Outil_E_Composants = RsBarGraph!Outil_E_Composants
 bool_Outil_E_Notas = RsBarGraph!Outil_E_Notas
 bool_Outil_E_cartouches = RsBarGraph!Outil_E_cartouches
 bool_Outil_E_Preconisations = RsBarGraph!Outil_E_Preconisations
 bool_Outil_E_Options = RsBarGraph!Outil_E_Options
 bool_Outil_E_Criteres = RsBarGraph!Outil_E_Criteres
 bool_Outil_E_Noeuds = RsBarGraph!Outil_E_Noeuds
'
 bool_Plan_Ouvrir = RsBarGraph!Plan_Ouvrir
 bool_Outil_Ouvrir = RsBarGraph!Outil_Ouvrir

  Nomenclature_Appareil = RsBarGraph!Nomenclature_Appareil
 Par_Fournisseur = RsBarGraph!Par_Fournisseur
 Par_Options = RsBarGraph!Par_Options
 NbPieces = RsBarGraph!NbPieces
 On Error Resume Next
Id_Pice = RsBarGraph!Id_Piece
DoEvents
Select Case RsBarGraph!Action
Case "Maj Eboutique"
         MajStock Id_Pice, IdFils, Val(NbPieces), Me

Case "Nomenclature"
      NomenclatureOk = False
Select Case PreparNomOk
    Case 0
            subExporteXls Id_Pice, False
    Case 1
        PreparationNomenclatuer Val(Id_Pice)
    Case 2
        Generer_Nomenclatuer Val(Id_Pice)
        Generer_Nomenclatuer2 Val(Id_Pice)
    Case 3
        Generer_NomenclatuerFinal Id_Pice
    Case 4
         Generer_NomenclatuerFinal Id_Pice
    
        
End Select
Case "Modifier Plan"

Sql = "SELECT T_Status.Status, T_indiceProjet.Id_Pieces FROM T_Status INNER JOIN T_indiceProjet ON T_Status.Id = T_indiceProjet.IdStatus "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_Pice & ";"
Set RsBarGraph = Con.OpenRecordSet(Sql)
strStatus = "" & RsBarGraph!Status
IdIndice = "" & RsBarGraph!Id_Pieces
Set RsBarGraph = Con.CloseRecordSet(RsBarGraph)
 subDessinerPlan Id_Pice
 
        subDessinerOutil Id_Pice
        If UCase(strStatus) = "VAL" Then
            MajEcartIndice IdIndice
        End If
 Case "Créer Ettiquettes"
       Sql = "SELECT T_indiceProjet.Equipement FROM T_indiceProjet "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_Pice & ";"
Set RsBarGraph = Con.OpenRecordSet(Sql)
 
      If RsBarGraph.EOF = False Then
 
 
 
  If Nomenclature_Appareil = True Then
Equipement = RsBarGraph!Equipement
Equipement = Split(Equipement & ";", ";")
For I = 0 To UBound(Equipement)
    Equipement2 = Split(Equipement(I) & "_", "_")
    If Trim("" & Equipement2(0)) <> "" Then
        Equipement3 = Equipement3 & ";" & Equipement2(0) & ";"
    End If
Next
    GenairEtiquette2 Val(Id_Pice), "" & Equipement3, Par_Options, Par_Fournisseur
    
Else
    GenairEtiquette Val(Val(Id_Pice))
End If
      End If
        
        
        
        
        
        
End Select
If boolAutoCAD = True Then
    Sql = "UPDATE T_Job SET T_Job.FinTraitement = True, T_Job.Status = 'NB Erreurs : " & NbError & "', T_Job.ValBarGraph = 0 "
    If NbError <> 0 Then
        Sql = Sql & ",FichierErr='" & FichierErr & "' "
    End If
    Sql = Sql & "WHERE T_Job.Job= " & Command & ";"
    Con.Execute Sql
    
     Con.Execute Sql
 End If
End If

  Sql = "DELETE [Utilise_Par].* FROM [Utilise_Par] "
  Sql = Sql & "WHERE [Utilise_Par].Machine='" & MyReplace(Machine) & "' "
  Sql = Sql & "and [Utilise_Par].User='" & MyReplace(UserName) & "';"
Con.Execute Sql
      
         
     
CodageX.DcrJenton
funCloseConnextion
Unload Me
End Sub

