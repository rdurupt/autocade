VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl RecherAutocable 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   795
   ControlContainer=   -1  'True
   DefaultCancel   =   -1  'True
   EditAtDesignTime=   -1  'True
   ForwardFocus    =   -1  'True
   ScaleHeight     =   990.476
   ScaleMode       =   0  'User
   ScaleWidth      =   1009.523
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   65
      ImageHeight     =   72
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":3772
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   0
      Picture         =   "UserControl1.ctx":6BD4
      Stretch         =   -1  'True
      ToolTipText     =   "Rechercher Pièce sur AutoCâble."
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "RecherAutocable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Event Action(Tableau_Valeur, Annuler)
Public Property Get Database() As String
Database = Mydb
End Property
Public Property Let Database(ByVal Database As String)
Mydb = Database
PropertyChanged "Database"

End Property

Public Property Get Filtre() As String
Filtre = MyFiltre
End Property
Public Property Let Filtre(ByVal Filtre As String)
MyFiltre = Filtre
PropertyChanged "Filtre"
End Property
Private Sub Action(ByVal Tableau_Valeur As Variant, ByVal Annuler As Integer)
'On appelle Click quand l'utilisateur clique sur le contrôle
RaiseEvent Action(ValeurTableau, MyAnnuler)
End Sub



Private Sub Image1_Click()
If Trim("" & Mydb) <> "" Then
Form1.Show vbModal
Unload Form1
RaiseEvent Action(ValeurTableau, MyAnnuler)
End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Image1.Picture = ImageList1.ListImages(2).Picture
End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Image1.Picture = ImageList1.ListImages(1).Picture
End If
End Sub




Private Sub Rechercher_Click()

If Trim("" & Mydb) <> "" Then
 Form1.Show vbModal
Unload Form1
RaiseEvent Action(ValeurTableau, MyAnnuler)
End If

End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Mydb = PropBag.ReadProperty("Database", "")
MyFiltre = PropBag.ReadProperty("Filtre", "VerifieDate<> null and Archiver=false and IdStatus<4 ")
End Sub

Private Sub UserControl_Resize()
Image1.Width = UserControl.Width * 1.24
Image1.Height = UserControl.Height * 1.24
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Database", Mydb, "")
Call PropBag.WriteProperty("Filtre", MyFiltre, "VerifieDate<> null and Archiver=false and IdStatus<4 ")

End Sub
