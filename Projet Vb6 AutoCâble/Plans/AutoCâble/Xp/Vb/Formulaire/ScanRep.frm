VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ScanRep 
   Caption         =   "Explorateur de Répertoires:"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4170
   Icon            =   "ScanRep.dsx":0000
   OleObjectBlob   =   "ScanRep.dsx":030A
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ScanRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyTxt As String
Dim bolSelec As Boolean
Dim Extension As String
Dim MyRepertoir As String
Dim Nclick As Boolean
Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub CommandButton1_Click()
If Me.ListBox1.ListIndex > -1 Then
    MyRepertoir = Me.ListBox1.List(Me.ListBox1.ListIndex, 1)
End If
Me.Hide
End Sub

Private Sub CommandButton2_Click()
Me.Hide
End Sub



Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
If Nclick = True Then Exit Sub
 Nclick = True
Dim d, Fd
Dim Fso As New FileSystemObject
Dim RepActive As String
Static NRep As Long
Dim TxtRep As String
On Error GoTo Fin
If Me.ListBox1.ListIndex = -1 Then Exit Sub
 RepActive = Me.ListBox1.List(Me.ListBox1.ListIndex, 0)
 RepActivetag = Me.ListBox1.List(Me.ListBox1.ListIndex, 1)
 
  If Left(RepActivetag, 2) = "\\" Then nbBoucle = 2 Else nbBoucle = -1
 If InStr(1, RepActive, "..") <> 0 Then
    NRep = NRep - 1
    If NRep > 0 Then
        RepM = Split(RepActivetag, "\")
        RepActivetag = ""
       
        For i = 0 To NRep + nbBoucle
           RepActivetag = RepActivetag & RepM(i) & "\"
        Next i
        
       

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
 For i = 1 To NRep
 TxtRep = TxtRep & "."
 Next i
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

Private Sub ListBox2_Click()

End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
CommandButton1_Click
End Sub

Private Sub UserForm_Activate()

Charge
End Sub

Sub Charge()
bolSelec = False
Dim Fso As New FileSystemObject
Dim d, dc
Dim i As Long
LoadDb
Dim LstRep() As String
Me.ListBox1.Clear
 i = 0
    Set dc = Fso.Drives
    For Each d In dc
  
        ListBox1.AddItem d
         Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = d
         If UCase(d) = "C:" Then Me.ListBox1.ListIndex = i
         i = i + 1
    Next
   
Me.ListBox1.AddItem "Câblage: donnees d entreprise"
Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = DonneesEntreprise

Me.ListBox1.AddItem "Câblage: production"
Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = DonneesProduction

bolSelec = True

End Sub
Public Function Chargement(txt As String) As String
Chargement = ""
MyRepertoir = ""
Me.Show vbModal
If MyRepertoir <> "" Then Chargement = MyRepertoir Else Chargement = txt
End Function

