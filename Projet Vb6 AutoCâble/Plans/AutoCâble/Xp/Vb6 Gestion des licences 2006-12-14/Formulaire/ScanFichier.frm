VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ScanFichier 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Explorateur de Fichiers :"
   ClientHeight    =   9105
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   7095
   Icon            =   "ScanFichier.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "ScanFichier.dsx":1272
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ScanFichier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim bolSelec As Boolean
Dim Extension As String
Dim MyFichier As String
Dim Nclick As Boolean
Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub CommandButton1_Click()
If Me.ListBox2.ListIndex > -1 Then
    MyFichier = Me.ListBox1.List(Me.ListBox1.ListIndex, 1) & "\" & Me.ListBox2.List(Me.ListBox2.ListIndex, 0)
End If
Me.Hide
End Sub

Private Sub CommandButton2_Click()
Me.Hide
End Sub

Private Sub ListBox1_Click()
Dim Fso As New FileSystemObject
Dim f, f1, fc, s
On Error Resume Next
If Nclick = True Then Exit Sub
  Me.ListBox2.Clear
  DoEvents
    Set f = Fso.GetFolder(Me.ListBox1.List(Me.ListBox1.ListIndex, 1))
    Set fc = f.Files
    For Each f1 In fc
    If Trim("" & Extension) <> "" Then
      If Right(UCase(f1), 4) = "." & Extension Then
           a = Split(f1, "\")
         Me.ListBox2.AddItem a(UBound(a))
      End If
    Else
          a = Split(f1, "\")
         Me.ListBox2.AddItem a(UBound(a))
    End If
   
    Next

Set Fso = Nothing
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Nclick = True Then Exit Sub
Dim d, Fd
Dim Fso As New FileSystemObject
Dim RepActive As String
Static NRep As Long
Dim TxtRep As String
On Error GoTo Fin
Nclick = True
If Me.ListBox1.ListIndex = -1 Then Exit Sub
 RepActive = Me.ListBox1.List(Me.ListBox1.ListIndex, 0)
 RepActivetag = Me.ListBox1.List(Me.ListBox1.ListIndex, 1)
 
  If Left(RepActivetag, 2) = "\\" Then nbBoucle = 2 Else nbBoucle = -1
 If InStr(1, RepActive, "..") <> 0 Then
    NRep = NRep - 1
    If NRep > 0 Then
        RepM = Split(RepActivetag, "\")
        RepActivetag = ""
       
        For I = 0 To NRep + nbBoucle
           RepActivetag = RepActivetag & RepM(I) & "\"
        Next I
        
       

    Else
       
       Charge
         Nclick = False
       Exit Sub
    End If
 Else
    
     NRep = NRep + 1
 End If
    If RepActive = "Câblage" Then
        RepActive = RepActivetag
    End If
  
    TxtRep = "."
 For I = 1 To NRep
 TxtRep = TxtRep & "."
 Next I
 ListBox1.Clear
 Me.ListBox1.AddItem TxtRep
 RepActive = RepActivetag
 Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = RepActive
Set Fd = Fso.GetFolder(RepActivetag)

   Set fc = Fd.SubFolders
    For Each f1 In fc
    a = Split(f1, "\")
    
    Me.ListBox1.AddItem a(UBound(a))
     Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = f1
    Next
    Set Fso = Nothing
    Me.ListBox1.ListIndex = 0
     Nclick = False
    Exit Sub
Fin:
MsgBox Err.Description
Set Fso = Nothing
 Charge
NRep = 0
 Nclick = False
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
CommandButton1_Click
End Sub

Private Sub UserForm_Activate()
Extension = UCase(Extension)
Me.Caption = Me.Caption & "*." & Extension
Charge
End Sub

Sub Charge()
bolSelec = False
Dim Fso As New FileSystemObject
Dim d, dc
Dim I As Long
'LoadDb
Dim LstRep() As String
Me.ListBox1.Clear
 I = 0
    Set dc = Fso.Drives
    For Each d In dc
  
        ListBox1.AddItem d
         Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = d
         If UCase(d) = "C:" Then Me.ListBox1.ListIndex = I
         I = I + 1
    Next
   
Me.ListBox1.AddItem "Câblage: donnees d entreprise"
Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = DonneesEntreprise

Me.ListBox1.AddItem "Câblage: production"
Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = DonneesProduction

bolSelec = True

End Sub
Public Function Chargement(txtExtension As String, Txt As String) As String
Chargement = ""
MyFichier = ""
Extension = txtExtension
Me.Show vbModal
If MyFichier <> "" Then Chargement = MyFichier Else Chargement = Txt
End Function
