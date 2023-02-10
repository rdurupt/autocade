VERSION 5.00
Object = "{0002E540-0000-0000-C000-000000000046}#1.1#0"; "MSOWC.DLL"
Begin VB.Form frmDetailOnglet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Détail Onglet Etat"
   ClientHeight    =   10305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   10305
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Paysage 
      Alignment       =   1  'Right Justify
      CausesValidation=   0   'False
      DisabledPicture =   "frmDetailOnglet.frx":0000
      DownPicture     =   "frmDetailOnglet.frx":0D62
      Height          =   615
      Left            =   11400
      Picture         =   "frmDetailOnglet.frx":1C14
      Style           =   1  'Graphical
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   600
      UseMaskColor    =   -1  'True
      Value           =   2  'Grayed
      Width           =   615
   End
   Begin VB.CheckBox PerfEntete 
      Alignment       =   1  'Right Justify
      Caption         =   "Inscrire préfixe dans entête"
      Height          =   315
      Left            =   9720
      TabIndex        =   59
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Frame Frame4 
      Caption         =   "Gestion des onglets:"
      Height          =   1215
      Left            =   240
      TabIndex        =   55
      Top             =   600
      Width           =   9255
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   58
         Top             =   840
         Width           =   9015
         Begin VB.OptionButton OptOnglets2 
            Caption         =   "Trier Par Volume de 0 à ?"
            Height          =   375
            Index           =   4
            Left            =   6840
            TabIndex        =   13
            Top             =   0
            Width           =   2175
         End
         Begin VB.OptionButton OptOnglets2 
            Caption         =   "Trier  Par Volume de ? à 0"
            Height          =   375
            Index           =   3
            Left            =   4560
            TabIndex        =   12
            Top             =   0
            Width           =   2175
         End
         Begin VB.OptionButton OptOnglets2 
            Caption         =   "Déplacer les Vides"
            Height          =   375
            Index           =   2
            Left            =   2760
            TabIndex        =   11
            Top             =   0
            Width           =   1695
         End
         Begin VB.OptionButton OptOnglets2 
            Caption         =   "Supprimer les Vides"
            Height          =   375
            Index           =   1
            Left            =   960
            TabIndex        =   10
            Top             =   0
            Width           =   1695
         End
         Begin VB.OptionButton OptOnglets2 
            Caption         =   "Aucune"
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   56
         Top             =   240
         Width           =   9015
         Begin VB.OptionButton OptOnglets 
            Caption         =   "Aucune"
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptOnglets 
            Caption         =   "Supprimer les Vides"
            Height          =   375
            Index           =   1
            Left            =   960
            TabIndex        =   5
            Top             =   0
            Width           =   1695
         End
         Begin VB.OptionButton OptOnglets 
            Caption         =   "Déplacer les Vides"
            Height          =   375
            Index           =   2
            Left            =   2760
            TabIndex        =   6
            Top             =   0
            Width           =   1695
         End
         Begin VB.OptionButton OptOnglets 
            Caption         =   "Trier  Par Volume de ? à 0"
            Height          =   375
            Index           =   3
            Left            =   4560
            TabIndex        =   7
            Top             =   0
            Width           =   2175
         End
         Begin VB.OptionButton OptOnglets 
            Caption         =   "Trier Par Volume de 0 à ?"
            Height          =   375
            Index           =   4
            Left            =   6840
            TabIndex        =   8
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.Label Label15 
         Caption         =   "ET"
         Height          =   255
         Index           =   0
         Left            =   4680
         TabIndex        =   57
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      Picture         =   "frmDetailOnglet.frx":2976
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Efface le Choix de l'onglet"
      Top             =   3600
      Width           =   375
   End
   Begin VB.CheckBox VueEpisure 
      Alignment       =   1  'Right Justify
      Caption         =   "Ajouter Vue Epissure"
      Height          =   315
      Left            =   9720
      TabIndex        =   15
      Top             =   1680
      Width           =   2295
   End
   Begin VB.CheckBox FiltreEquipement 
      Alignment       =   1  'Right Justify
      Caption         =   "Filtre sur équipement"
      Height          =   315
      Left            =   9720
      TabIndex        =   14
      Top             =   1440
      Width           =   2295
   End
   Begin OWC.Spreadsheet FiltreActif 
      Height          =   2775
      Left            =   240
      TabIndex        =   33
      Top             =   6720
      Width           =   12135
      HTMLURL         =   ""
      HTMLData        =   $"frmDetailOnglet.frx":2A1F
      DataType        =   "HTMLDATA"
      AutoFit         =   0   'False
      DisplayColHeaders=   0   'False
      DisplayGridlines=   -1  'True
      DisplayHorizontalScrollBar=   -1  'True
      DisplayRowHeaders=   0   'False
      DisplayTitleBar =   -1  'True
      DisplayToolbar  =   0   'False
      DisplayVerticalScrollBar=   -1  'True
      EnableAutoCalculate=   -1  'True
      EnableEvents    =   -1  'True
      MoveAfterReturn =   -1  'True
      MoveAfterReturnDirection=   0
      RightToLeft     =   0   'False
      ViewableRange   =   "1:65536"
   End
   Begin VB.Frame Frame2 
      Caption         =   "Maro Complémentaire"
      Height          =   1575
      Left            =   240
      TabIndex        =   42
      Top             =   1920
      Width           =   12015
      Begin VB.ComboBox DapresDoc 
         Height          =   315
         Left            =   2040
         TabIndex        =   16
         Top             =   360
         Width           =   1695
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   285
         Left            =   5400
         Max             =   9999
         Min             =   -9999
         TabIndex        =   19
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox FiltreSequentielle 
         Alignment       =   1  'Right Justify
         Caption         =   "Filtre séquentielle"
         Height          =   315
         Left            =   4080
         TabIndex        =   17
         Top             =   360
         Width           =   1815
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   285
         Left            =   10920
         Max             =   9999
         Min             =   -9999
         TabIndex        =   22
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox OngleEnd 
         Height          =   285
         Left            =   7440
         TabIndex        =   21
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox OngletStrat 
         Height          =   285
         Left            =   2040
         TabIndex        =   18
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "A partir d'un Document ?"
         Height          =   315
         Left            =   120
         TabIndex        =   49
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label16 
         Caption         =   "Décalé de:"
         Height          =   285
         Left            =   10080
         TabIndex        =   48
         Top             =   960
         Width           =   855
      End
      Begin VB.Label DecaleAppres 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   5640
         TabIndex        =   20
         Top             =   960
         Width           =   615
      End
      Begin VB.Label DecaleAvant 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   285
         Left            =   11160
         TabIndex        =   23
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "N°, Nom ou Dernier"
         Height          =   255
         Index           =   1
         Left            =   7440
         TabIndex        =   47
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "N°, Nom ou Premier"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   46
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label13 
         Caption         =   "Décalé de:"
         Height          =   285
         Left            =   4560
         TabIndex        =   45
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "au N°/Nom :"
         Height          =   285
         Left            =   6360
         TabIndex        =   44
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Scane onglet du N°/Nom :"
         Height          =   285
         Left            =   120
         TabIndex        =   43
         Top             =   960
         Width           =   1935
      End
   End
   Begin VB.TextBox SaveOnglet 
      Height          =   315
      Left            =   10320
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtMacro 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Left            =   10080
      Top             =   9480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Anuller"
      Height          =   615
      Left            =   7320
      TabIndex        =   35
      Top             =   9600
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Valider"
      Height          =   615
      Left            =   2400
      TabIndex        =   34
      Top             =   9600
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmDetailOnglet.frx":3818
      Left            =   5160
      List            =   "frmDetailOnglet.frx":381A
      TabIndex        =   2
      Text            =   "Critères"
      Top             =   240
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Liste des Champ"
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000007&
      Height          =   2775
      Left            =   1080
      TabIndex        =   0
      Top             =   3840
      Width           =   11175
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   0
         ScaleHeight     =   2295
         ScaleWidth      =   2655
         TabIndex        =   36
         Top             =   120
         Width           =   2655
         Begin VB.OptionButton Xls 
            Height          =   405
            Index           =   0
            Left            =   1440
            Picture         =   "frmDetailOnglet.frx":381C
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Créer un onglet sur ce Champ ?"
            Top             =   1320
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox ChampsAs 
            Height          =   405
            Index           =   0
            Left            =   0
            TabIndex        =   26
            Top             =   880
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CheckBox TriN 
            Height          =   405
            Index           =   0
            Left            =   360
            Picture         =   "frmDetailOnglet.frx":3D9E
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Pas de Tri"
            Top             =   1320
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CheckBox TriH 
            Height          =   405
            Index           =   0
            Left            =   1080
            Picture         =   "frmDetailOnglet.frx":3FBC
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Tri Décroissent"
            Top             =   1320
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CheckBox TriB 
            Height          =   405
            Index           =   0
            Left            =   720
            Picture         =   "frmDetailOnglet.frx":41DA
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Tri Croissent"
            Top             =   1320
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CheckBox SousTautal 
            Height          =   405
            Index           =   0
            Left            =   960
            Picture         =   "frmDetailOnglet.frx":43F8
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Ajouter Sous Total ?"
            Top             =   1800
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CheckBox ChkVisible 
            Caption         =   "Oui"
            CausesValidation=   0   'False
            Height          =   405
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   0
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label Champs 
            BorderStyle     =   1  'Fixed Single
            Height          =   405
            Index           =   0
            Left            =   0
            TabIndex        =   37
            Top             =   435
            Visible         =   0   'False
            Width           =   2535
         End
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   270
         LargeChange     =   2400
         Left            =   0
         Max             =   9999
         Min             =   240
         SmallChange     =   240
         TabIndex        =   32
         Top             =   2520
         Value           =   240
         Width           =   11175
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Champs"
      Height          =   405
      Left            =   360
      TabIndex        =   54
      Top             =   4515
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Nom"
      Height          =   405
      Left            =   360
      TabIndex        =   53
      Top             =   4965
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Visible"
      Height          =   405
      Left            =   360
      TabIndex        =   52
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Trier"
      Height          =   405
      Left            =   360
      TabIndex        =   51
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "S/Toutal"
      Height          =   405
      Left            =   360
      TabIndex        =   50
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "Onglet ou Préfix de savegarde"
      Height          =   315
      Left            =   7680
      TabIndex        =   41
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Nom De la Maco"
      Height          =   315
      Left            =   240
      TabIndex        =   40
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Onglet"
      Height          =   315
      Left            =   4440
      TabIndex        =   39
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Liste des Champs:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   38
      Top             =   3600
      Width           =   2295
   End
End
Attribute VB_Name = "frmDetailOnglet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim boolPasClik As Boolean
Dim nb As Long
Dim delet As Boolean
Dim Id As Long
Dim Menu As Long
Dim CreateDate As Date

Public Sub chargement(id_Menu As Long, Id_Fichier As Long, Creele As Date)
Dim Sql As String
Dim I As Long
Dim Rs As Recordset
CreateDate = Creele
Sql = "SELECT T_Menu_Etat_Onglet.Menu FROM T_Menu_Etat_Onglet ORDER BY T_Menu_Etat_Onglet.Ordre;"
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    Combo2.AddItem "" & Replace(Rs!Menu, "œ", "oe")
    Rs.MoveNext
Wend
Sql = "SELECT T_ETATS.* From T_ETATS "
Sql = Sql & "WHERE T_ETATS.ID=" & Id_Fichier & ";"
Set Rs = Con.OpenRecordSet(Sql)
Menu = Rs!Id

Me.Tag = Id_Fichier
Me.Caption = Me.Caption & ": " & Rs!EtatName & " Menu: " & Rs!Menu

Sql = "SELECT T_Etats_Onglet.* From T_Etats_Onglet "
Sql = Sql & "WHERE T_Etats_Onglet.Id=" & id_Menu & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
Me.VScroll1.Value = Val(Trim("" & Rs!DecaleAvant))
Me.VScroll2.Value = Val(Trim("" & Rs!DecaleAppres))
Me.txtMacro = "" & Rs!Macro
Me.txtMacro.Tag = "" & Rs!Macro
 SaveOnglet = "" & Rs!SaveOnglet
 SaveOnglet.Tag = "" & Rs!SaveOnglet
     If Rs!Paysage = True Then
       Me.Paysage.Value = 1
    Else
        Me.Paysage.Value = 0
    End If

    If Rs!PerfEntete = True Then
       Me.PerfEntete.Value = 1
    Else
        Me.PerfEntete.Value = 0
    End If
 If Rs!FiltreSequentielle = True Then
    Me.FiltreSequentielle.Value = 1
 Else
    Me.FiltreSequentielle.Value = 1
 End If
 OngleEnd = "" & Rs!OngleEnd
 OngletStrat = "" & Rs!OngletStrat
 Me.FiltreSequentielle = 1
For I = 0 To Me.Combo2.ListCount - 1
     If Me.Combo2.List(I) = Rs!Onglet Then Me.Combo2.ListIndex = I: Exit For
     
Next
End If
Set Rs = Con.CloseRecordSet(Rs)
Id = id_Menu
Me.Show vbModal
End Sub



Private Sub ChkVisible_Click(Index As Integer)
If ChkVisible(Index) = 1 Then
    ChkVisible(Index).Caption = "OUI"
Else
    ChkVisible(Index).Caption = "NON"
End If
End Sub



Private Sub Combo2_Click()
Static Save As String
If Trim("" & Save) <> Me.Combo2 Then
delet = True
   
End If
Save = Me.Combo2
End Sub

Private Sub Command1_Click()
Dim Sql As String
Dim SqlValue As String
Dim Rs As Recordset
Dim Row As Long
Dim MyRange
Dim IndexOption, IndexOption2, IndexOption3 As Long
For IndexOption = 0 To Me.OptOnglets.Count - 1
    If Me.OptOnglets(IndexOption).Value = True Then Exit For
    
Next
For IndexOption2 = 0 To Me.OptOnglets2.Count - 1
    If Me.OptOnglets2(IndexOption2).Value = True Then Exit For
    
Next

If Trim("" & txtMacro) = "" Then MsgBox "Vous devez saisir le nom de la macro": txtMacro.SetFocus: Exit Sub
If Trim("" & SaveOnglet) = "" Then MsgBox "Vous devez saisir le nom de L'onglet de sauvegarde": SaveOnglet.SetFocus: Exit Sub

Sql = "SELECT T_Etats_Onglet.Macro, T_Etats_Onglet.Id From T_Etats_Onglet "
Sql = Sql & "WHERE T_Etats_Onglet.Macro='" & MyReplace(txtMacro) & "' and T_Etats_Onglet.Id<>" & Id & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then MsgBox "La macro: " & txtMacro & " existe déjà.", vbExclamation: txtMacro = txtMacro.Tag: txtMacro.SetFocus: Exit Sub

Sql = "SELECT T_Etats_Onglet.Macro, T_Etats_Onglet.Id,T_Etats_Onglet.Document,"
Sql = Sql & "T_Etats_Onglet.OngletStrat,T_Etats_Onglet.OngleEnd, T_Etats_Onglet.GestionOnglet, T_Etats_Onglet.GestionOnglet2   From T_Etats_Onglet "
'Sql = Sql & "T_Etats_Onglet.DecaleAvant,T_Etats_Onglet.DecaleAppres "
Sql = Sql & "WHERE T_Etats_Onglet.SaveOnglet='" & MyReplace(SaveOnglet) & "' and T_Etats_Onglet.Id<>" & Id & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then MsgBox "L'onglet de sauvegarde: " & SaveOnglet & " existe déjà.", vbExclamation: SaveOnglet = SaveOnglet.Tag: SaveOnglet.SetFocus: Exit Sub
Sql = "SELECT T_Etats_Onglet.* "
Sql = Sql & "From T_Etats_Onglet "
Sql = Sql & "WHERE T_Etats_Onglet.Id_Etat=" & Me.Tag & " "
If Id <> 0 Then
Sql = Sql & "and T_Etats_Onglet.Id=" & Id & " "
End If
Sql = Sql & "ORDER BY T_Etats_Onglet.Id DESC;"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Or Id = 0 Then
        Rs.AddNew
        Rs!Macro = txtMacro
         Rs!SaveOnglet = SaveOnglet
         Rs!Paysage = Paysage
        Rs!Id_Etat = Me.Tag
        Rs!Onglet = Me.Combo2
        Rs!GestionOnglet = IndexOption
        Rs!GestionOnglet2 = IndexOption2
        Rs!DecaleAvant = Me.DecaleAvant
        Rs!DecaleAppres = Me.DecaleAppres
        If Me.PerfEntete.Value = 1 Then
            Rs!PerfEntete = True
        Else
            Rs!PerfEntete = False
        End If
        
        If Trim("" & Me.DapresDoc) = "" Then
            Rs!Document = Null
        Else
            Rs!Document = Me.DapresDoc
        End If
        
        If Trim("" & Me.OngletStrat) = "" Then
             Rs!OngletStrat = Null
        Else
         Rs!OngletStrat = Me.OngletStrat
        End If
       
       If Trim("" & Me.OngleEnd) = "" Then
             Rs!OngleEnd = Null
        Else
         Rs!OngleEnd = Me.OngleEnd
        End If
       If Me.VueEpisure.Value = 1 Then
            Rs!VueEpissur = True
        Else
            Rs!VueEpissur = False
        End If
        Rs.Update
Else
        Rs!Paysage = Paysage
        If Me.PerfEntete.Value = 1 Then
            Rs!PerfEntete = True
        Else
            Rs!PerfEntete = False
        End If
        Rs!SaveOnglet = SaveOnglet
        Rs!Macro = txtMacro
        Rs!Onglet = Me.Combo2
        Rs!DecaleAvant = Me.DecaleAvant
        Rs!DecaleAppres = Me.DecaleAppres
        Rs!GestionOnglet = IndexOption
        Rs!GestionOnglet2 = IndexOption2
        
        If Trim("" & Me.DapresDoc) = "" Then
             Rs!Document = Null
        Else
            Rs!Document = Me.DapresDoc
        End If
        If Trim("" & Me.OngletStrat) = "" Then
             Rs!OngletStrat = Null
        Else
         Rs!OngletStrat = Me.OngletStrat
        End If
       
       If Trim("" & Me.OngleEnd) = "" Then
             Rs!OngleEnd = Null
        Else
         Rs!OngleEnd = Me.OngleEnd
        End If
        If Me.VueEpisure.Value = 1 Then
            Rs!VueEpissur = True
        Else
            Rs!VueEpissur = False
        End If
        Rs.Update
End If
Rs.Requery
Dim I As Long
Sql = "DELETE T_Etats_Select_Champs.*, T_Etats_Select_Champs.Id_Onglet "
Sql = Sql & "From T_Etats_Select_Champs "
Sql = Sql & "WHERE T_Etats_Select_Champs.Id_Onglet=" & Rs!Id & ";"
Con.Execute Sql
Sql = "INSERT INTO T_Etats_Select_Champs ( Id_Onglet, ChamsName, ChampAs, Trie ,SousTautal,Visible,CreatOnglet)"
Sql = Sql & "VALUES ( " & Rs!Id & ","

For I = 1 To ChkVisible.Count - 1
SqlValue = ""
If Trim("" & ChampsAs(I)) = "" Then ChampsAs(I) = Champs(I)
    
       SqlValue = SqlValue & "'" & MyReplace(Champs(I)) & "',"
       SqlValue = SqlValue & "'" & MyReplace(ChampsAs(I)) & "',"
       
     If TriN(I).Value = 1 Then SqlValue = SqlValue & 0 & ", "
        If TriB(I).Value = 1 Then SqlValue = SqlValue & 1 & ", "
         If TriH(I).Value = 1 Then SqlValue = SqlValue & 2 & ", "
        
        
    
    If SousTautal(I).Value = 1 Then
        SqlValue = SqlValue & "true,"
    Else
        SqlValue = SqlValue & "False,"
    End If
    If ChkVisible(I).Value = 1 Then
        SqlValue = SqlValue & "false,"
    Else
        SqlValue = SqlValue & "true,"
    End If
    
    If Me.Xls(I).Value = False Then
        SqlValue = SqlValue & "false"
    Else
        SqlValue = SqlValue & "true"
    End If
        Con.Execute Sql & SqlValue & ");"
   
Next
If FiltreSequentielle.Value = 1 Then
    Sql = "UPDATE T_Etats_Onglet SET T_Etats_Onglet.FiltreSequentielle = True "

Else
    Sql = "UPDATE T_Etats_Onglet SET T_Etats_Onglet.FiltreSequentielle = False "
End If
     Sql = Sql & "WHERE T_Etats_Onglet.Id=" & Rs!Id & ";"
     Con.Execute Sql
     
If FiltreEquipement.Value = 1 Then
    Sql = "UPDATE T_Etats_Onglet SET T_Etats_Onglet.FiltreEquipement = True "

Else
    Sql = "UPDATE T_Etats_Onglet SET T_Etats_Onglet.FiltreEquipement = False "
End If
     Sql = Sql & "WHERE T_Etats_Onglet.Id=" & Rs!Id & ";"
     Con.Execute Sql
     
     
     
Sql = "DELETE T_Etats_Select_Filtre.* From T_Etats_Select_Filtre "
Sql = Sql & "WHERE T_Etats_Select_Filtre.Id_Onglet=" & Rs!Id & ";"
Con.Execute Sql
Set MyRange = Me.FiltreActif.Range("A1").CurrentRegion

For Row = 2 To MyRange.Rows.Count
    For I = 2 To MyRange.Columns.Count
    Debug.Print Trim("" & MyRange(Row, I))
        If Trim("" & MyRange(Row, I)) <> "" Then
            Sql = "INSERT INTO T_Etats_Select_Filtre ( Id_Onglet, FiltreName, Colonne ,Valeur,Ligne) "
            Sql = Sql & "values(" & Rs!Id & ", '" & MyReplace("" & MyRange(Row, 1)) & "' , '" & MyReplace("" & MyRange(1, I)) & "', '" & MyReplace("" & MyRange(Row, I)) & "'," & Row & ");"
            Con.Execute Sql
        End If
    Next
Next
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Command3_Click()
Dim I As Long
For I = 0 To Xls.Count - 1
    Xls(I).Value = False
Next
End Sub

Private Sub Form_Load()
Me.HScroll1.Max = Picture1.Width
Me.HScroll1.Min = 0
Me.HScroll1.Value = 0
delet = True
Me.Timer1.Interval = 1
Me.Timer1.Enabled = True

End Sub

Private Sub HScroll1_Change()
If Me.HScroll1.Value = 0 Then
    Picture1.Left = 240
Else
Picture1.Left = Me.HScroll1.Value * (-1 * nb)
End If
End Sub

Private Sub Timer1_Timer()
If delet = True Then
    UnlaodContol
    LaodContol
    delet = False
End If
End Sub

Private Sub TriB_Click(Index As Integer)
If boolPasClik = True Then Exit Sub
boolPasClik = True
TriN(Index).Value = 0
TriH(Index).Value = 0
boolPasClik = False
End Sub

Private Sub TriH_Click(Index As Integer)
If boolPasClik = True Then Exit Sub
boolPasClik = True
TriN(Index).Value = 0
TriB(Index).Value = 0
boolPasClik = False
End Sub

Private Sub TriN_Click(Index As Integer)
If boolPasClik = True Then Exit Sub
boolPasClik = True
TriB(Index).Value = 0
TriH(Index).Value = 0
boolPasClik = False
End Sub

Private Sub Visible_Click(Index As Integer)

End Sub
Sub UnlaodContol()
Dim I As Long
Dim FRM
Set FRM = Me
For I = FRM.ChkVisible.Count - 1 To 1 Step -1
    Unload FRM.ChkVisible(I)
    Unload FRM.Champs(I)
    Unload FRM.ChampsAs(I)
    Unload FRM.TriN(I)
    Unload FRM.TriB(I)
    Unload FRM.TriH(I)
    Unload FRM.Xls(I)
    Debug.Print Me.Xls.Count
    Unload FRM.SousTautal(I)
Next
End Sub
Sub LaodContol()
  Dim Sql As String
Dim Rs As Recordset
Dim RsFiltreAcrif As Recordset
Dim I As Integer
Dim NonVisible As Long
Dim MyRange
Dim UnVisible As Boolean
Dim Row As Long
Dim MyDocument As String
Dim AppOk As Boolean
Me.Visible = True
Me.VueEpisure.Enabled = True
For I = Me.FiltreActif.ActiveSheet.Range("a1").CurrentRegion.Rows.Count To 1 Step -1
  Me.FiltreActif.ActiveSheet.Rows(I).DeleteRows
Next
nb = 0
Sql = "SELECT T_Etats_Onglet.Macro, T_Etats_Onglet.FiltreSequentielle, T_Etats_Onglet.FiltreEquipement, T_Etats_Onglet.Id , "
 Sql = Sql & "T_Etats_Onglet.Document,T_Etats_Onglet.VueEpissur, T_Etats_Onglet.GestionOnglet, T_Etats_Onglet.GestionOnglet2 "
Sql = Sql & "From T_Etats_Onglet "
 Sql = Sql & "WHERE T_Etats_Onglet.Id=" & Id & " "
 Sql = Sql & "AND T_Etats_Onglet.Onglet='" & MyReplace(Me.Combo2) & "'; "

Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
MyDocument = "" & Rs!Document
Me.OptOnglets(Val("" & Rs!GestionOnglet)).Value = True
Me.OptOnglets2(Val("" & Rs!GestionOnglet2)).Value = True

    If Rs!FiltreSequentielle = True Then
        FiltreSequentielle.Value = 1
    Else
        FiltreSequentielle.Value = 0
    End If
    If Rs!FiltreEquipement = True Then
        FiltreEquipement.Value = 1
    Else
        FiltreEquipement.Value = 0
    End If
    If Rs!VueEpissur = True Then
        Me.VueEpisure.Value = 1
    Else
        Me.VueEpisure.Value = 0
    End If
    
Else
    FiltreEquipement.Value = 0
    FiltreSequentielle.Value = 0
    Me.VueEpisure.Value = 0
End If
'Dim Champs(Champs.Count - 1) As Long
Sql = "SELECT T_Menu_Etat_Onglet.Menu, T_Menu_Etat_Onglet.TableName, T_Menu_Etat_Onglet.App "
Sql = Sql & "FROM T_Menu_Etat_Onglet "
Sql = Sql & "WHERE T_Menu_Etat_Onglet.Menu='" & MyReplace(Me.Combo2) & "';"
Set Rs = Con.OpenRecordSet(Sql)
AppOk = Rs!App
'Select Case Me.Combo2
'    Case "Critères"
        Sql = "SELECT " & Rs!TableName & ".* From " & Rs!TableName & " WHERE " & Rs!TableName & ".Id=0;"

'    Case "Connecteurs"
'        Sql = "SELECT Connecteurs.* From Connecteurs WHERE Connecteurs.Id=0;"
'    Case "Tableau de fils"
'        Sql = "SELECT Ligne_Tableau_fils.* From Ligne_Tableau_fils WHERE Ligne_Tableau_fils.Id=0;"
'    Case "Composants"
'        Sql = "SELECT Composants.* From Composants WHERE Composants.Id=0;"
'    Case "Notas"
'        Sql = "SELECT Nota.* From Nota WHERE Nota.Id=0;"
'    Case "Noeuds"
'        Sql = "SELECT T_Noeuds.* From T_Noeuds WHERE T_Noeuds.Id=0;"
'End Select
Set Rs = Con.OpenRecordSet(Sql)
NonVisible = -1
 Me.FiltreActif.ActiveSheet.Cells(1, 1) = "Nom Du Filtre"
 Me.FiltreActif.Cells(1, 1).Interior.Color = 12632256
For I = 0 To Rs.Fields.Count - 1
    If Left(UCase(Rs.Fields(I).Name) & "___", 3) <> "ID_" Then
   
        Load ChkVisible(ChkVisible.Count)
        Load Champs(Champs.Count)
        Load TriN(TriN.Count)
        Load TriB(TriB.Count)
        Load TriH(TriH.Count)
        Load ChampsAs(ChampsAs.Count)
        Load Me.Xls(Me.Xls.Count)
        Load SousTautal(SousTautal.Count)
        
        ChkVisible(ChkVisible.Count - 1).Visible = True
         Champs(Champs.Count - 1).Visible = True
         TriN(TriN.Count - 1).Visible = True
         TriB(TriB.Count - 1).Visible = True
         TriH(TriH.Count - 1).Visible = True
         Me.Xls(Me.Xls.Count - 1).Visible = True
         
         ChampsAs(ChampsAs.Count - 1).Visible = True
         SousTautal(SousTautal.Count - 1).Visible = True
         ChkVisible(ChkVisible.Count - 1).Left = ChkVisible(0).Left + (2535 * nb)
         Champs(Champs.Count - 1).Left = Champs(0).Left + (2535 * nb)
         ChampsAs(ChampsAs.Count - 1).Left = ChampsAs(0).Left + (2535 * nb)
         TriN(TriN.Count - 1).Left = TriN(0).Left + (2535 * nb)
         TriB(TriB.Count - 1).Left = TriB(0).Left + (2535 * nb)
         TriB(TriB.Count - 1).Left = TriB(0).Left + (2535 * nb)
         TriH(TriH.Count - 1).Left = TriH(0).Left + (2535 * nb)
         Me.Xls(Me.Xls.Count - 1).Left = Me.Xls(0).Left + (2535 * nb)
         SousTautal(SousTautal.Count - 1).Left = TriB(TriB.Count - 1).Left
         Me.Picture1.Width = 3375 + (2535 * (nb + 3))
         nb = nb + 1
         Champs(Champs.Count - 1).Caption = Rs.Fields(I).Name
          Me.FiltreActif.ActiveSheet.Cells(1, I + 1 - NonVisible) = Rs.Fields(I).Name
'          RsFiltreAcrif
          Me.FiltreActif.Cells(1, I + 1 - NonVisible).Interior.Color = 12632256
          Sql = "SELECT T_Etats_Select_Filtre.FiltreName, T_Etats_Select_Filtre.Colonne, T_Etats_Select_Filtre.Valeur, "
        Sql = Sql & "T_Etats_Select_Filtre.Ligne "
        Sql = Sql & "FROM T_Etats_Onglet INNER JOIN T_Etats_Select_Filtre ON T_Etats_Onglet.Id = T_Etats_Select_Filtre.Id_Onglet "
        Sql = Sql & "WHERE T_Etats_Select_Filtre.Colonne='" & MyReplace(Rs.Fields(I).Name) & "' "
        Sql = Sql & "AND T_Etats_Select_Filtre.Id_Onglet=" & Id & "  "
        Sql = Sql & "AND T_Etats_Onglet.Onglet='" & MyReplace(Me.Combo2) & "' "
        Sql = Sql & "ORDER BY T_Etats_Select_Filtre.Id;"
        Set RsFiltreAcrif = Con.OpenRecordSet(Sql)
            Row = 1
        While RsFiltreAcrif.EOF = False
        Row = RsFiltreAcrif!Ligne
            
            Me.FiltreActif.ActiveSheet.Cells(Row, 1) = "" & RsFiltreAcrif!FiltreName
            Me.FiltreActif.ActiveSheet.Cells(Row, I + 1 - NonVisible) = "" & RsFiltreAcrif!Valeur
            RsFiltreAcrif.MoveNext
        Wend
        Set RsFiltreAcrif = Con.CloseRecordSet(RsFiltreAcrif)

     Else
        NonVisible = NonVisible + 1
    End If
Next
Me.FiltreActif.Range("a1").CurrentRegion.Columns.AutoFitColumns
If Id <> 0 Then
    For I = 1 To Champs.Count - 1
        Sql = "SELECT T_Etats_Select_Champs.* "
        Sql = Sql & "FROM T_Etats_Onglet INNER JOIN T_Etats_Select_Champs ON T_Etats_Onglet.Id = T_Etats_Select_Champs.Id_Onglet "
        Sql = Sql & "WHERE T_Etats_Select_Champs.Id_Onglet=" & Id & " "
        Sql = Sql & "AND T_Etats_Select_Champs.ChamsName='" & MyReplace(Champs(I)) & "'  AND T_Etats_Onglet.Onglet='" & MyReplace(Me.Combo2) & "' ;"
        Set Rs = Con.OpenRecordSet(Sql)
        If Rs.EOF = False Then
        If Rs!Visible = False Then
            ChkVisible(I).Value = 1
        Else
            ChkVisible(I).Value = 0
        End If
            Me.Xls(I).Value = Rs!CreatOnglet
            ChampsAs(I) = "" & Rs!ChampAs
'            UnVisible = True
            If Rs!SousTautal = True Then
                SousTautal(I).Value = 1
            Else
                SousTautal(I).Value = 0
            End If
            
            Select Case Rs!Trie
                Case 1
                    TriB(I).Value = 1
                Case 2
                    TriH(I).Value = 1
                Case Else
                    TriN(I).Value = 1
            End Select
'        Else
'            ChkVisible(I).Value = 0
'            Me.Xls(I).Value = False
'        End If
        Else
             SousTautal(I).Value = 0
             TriN(I).Value = 1
            
        End If
    Next

End If
'If UnVisible = False Then
'    For I = 1 To ChkVisible.Count - 1
'        ChkVisible(I).Value = 1
'    Next
'End If
Set Rs = Con.CloseRecordSet(Rs)

Me.HScroll1.Value = 0
DapresDoc.Clear
DapresDoc.AddItem ""
DapresDoc.AddItem "Dossier Fab"
DapresDoc.AddItem "Dossier Control"
DapresDoc.AddItem "Li"
Sql = "SELECT T_ETATS.EtatName From T_ETATS "
Sql = Sql & "WHERE T_ETATS.id<>" & Menu & " and T_ETATS.CreateDate<#" & Format(CreateDate, "yyyy-mm-dd hh:mm:ss") & "# "
Sql = Sql & "ORDER BY T_ETATS.CreateDate;"
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    DapresDoc.AddItem "" & Rs!EtatName
    Rs.MoveNext
    Wend
    For I = 0 To DapresDoc.ListCount
    If UCase(DapresDoc.List(I)) = UCase(MyDocument) Then
        DapresDoc.ListIndex = I
        Exit For
    End If
    Next
If AppOk = False Then
    Me.VueEpisure.Value = 0
    Me.VueEpisure.Enabled = False
   
End If
End Sub

Private Sub VScroll1_Change()
DecaleAvant.Caption = Me.VScroll1.Value
End Sub

Private Sub VScroll2_Change()
DecaleAppres.Caption = Me.VScroll2.Value
End Sub
