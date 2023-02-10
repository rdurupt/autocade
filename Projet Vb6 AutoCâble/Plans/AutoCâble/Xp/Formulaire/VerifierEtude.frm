VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VerifierEtude 
   Caption         =   "Vérification Plan :"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12240
   OleObjectBlob   =   "VerifierEtude.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VerifierEtude"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IdProjet As Long
Public IdPieces As Long
Public IdIndiceProjet As Long
Dim Noquite As Boolean


Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then
    txt18 = Format(Date, "dd/mm/yyyy")
Else
    txt18 = ""
End If
End Sub

Private Sub CommandButton1_Click()
Dim sql As String
Dim Rs As Recordset

CherchPices.Charge Me, " LiAutoCadSave <>  Null and NbErr=0 and IdStatus<3 and Archiver=False", True
Unload CherchPices





sql = "SELECT RqCartouche.Projet AS txt1,  "
sql = sql & "RqCartouche.Vague AS txt2,  "
sql = sql & "RqCartouche.Equipement AS txt3,  "
sql = sql & "RqCartouche.Responsable AS txt4,  "
sql = sql & "RqCartouche.Ensemble AS txt5,  "
sql = sql & "RqCartouche.CleAc AS txt6,  "
sql = sql & "RqCartouche.PI AS txt7,  "
sql = sql & "RqCartouche.PL AS txt8,  "
sql = sql & "RqCartouche.[OU] AS txt9,  "
sql = sql & "RqCartouche.Li AS txt10,  "
sql = sql & "RqCartouche.Client AS txt11,  "
sql = sql & "RqCartouche.Destinataire AS txt12,  "
sql = sql & "RqCartouche.Service AS txt13,  "
sql = sql & "RqCartouche.RefPF AS txt14, "
sql = sql & " RqCartouche.RefP AS txt15,  "
sql = sql & "RqCartouche.DessineDate AS txt16,  "
sql = sql & "RqCartouche.DessineNOM AS txt17,  "
sql = sql & "RqCartouche.VerifieDate AS txt18,  "
sql = sql & "RqCartouche.VerifieNom AS txt19,  "
sql = sql & "RqCartouche.ApprouveDate AS txt20,  "
sql = sql & "RqCartouche.ApprouveNom AS txt21 "
sql = sql & "FROM RqCartouche "
sql = sql & "WHERE RqCartouche.T_indiceProjet.Id=" & Val(Me.Tag) & " ;"
Debug.Print sql
Set Rs = Con.OpenRecordSet(sql)
If Rs.EOF = False Then
 Me.Controls("txt" & CStr(1)).Caption = "" & Rs.Fields("txt" & CStr(1))
For i = 2 To 3
    Me.Controls("txt" & CStr(i)) = "" & Rs.Fields("txt" & CStr(i))
Next i
 Me.Controls("txt" & CStr(4)).Caption = "" & Rs.Fields("txt" & CStr(4))
  Me.Controls("txt" & CStr(5)) = "" & Rs.Fields("txt" & CStr(5))
For i = 6 To 12
    Me.Controls("txt" & CStr(i)).Caption = "" & Rs.Fields("txt" & CStr(i))
Next i
For i = 13 To 15
    Me.Controls("txt" & CStr(i)) = "" & Rs.Fields("txt" & CStr(i))
Next i
For i = 16 To 18 Step 2
    Me.Controls("txt" & CStr(i)) = "" & Rs.Fields("txt" & CStr(i))
Next i
 Me.Controls("txt" & CStr(20)) = "" & Rs.Fields("txt" & CStr(20))
For i = 17 To 21 Step 2
    Me.Controls("txt" & CStr(i)).Caption = "" & Rs.Fields("txt" & CStr(i))
Next i
End If
If txt18 <> "" Then
    Me.CheckBox1.Value = True
Else
     Me.CheckBox1.Value = False
End If
End Sub

Private Sub CommandButton7_Click()
Dim sql As String
Dim Rs As Recordset
Set FormBarGrah = Me
If MyFormat("DATE", txt16, "Déssiné par") = False Then Exit Sub
If MyFormat("DATE", txt18, "Vérifié par") = False Then Exit Sub
If MyFormat("DATE", txt20, "Approuvé par") = False Then Exit Sub
If Trim("" & Me.Tag) = "" Then
    CommandButton1_Click
    Exit Sub
End If


sql = "SELECT T_indiceProjet.Pere FROM T_indiceProjet "
sql = sql & "WHERE T_indiceProjet.Id=" & Me.Tag & ";"
Set Rs = Con.OpenRecordSet(sql)
If Rs!Pere > 0 Then Me.Tag = Rs!Pere


sql = "UPDATE RqCartouche SET "
sql = sql & "RqCartouche.VerifieDate = " & MyReplaceDate(txt18) & " "
sql = sql & "WHERE RqCartouche.T_indiceProjet.Id=" & Me.Tag & ";"
Con.Exequte sql

sql = "UPDATE RqCartouche SET "
sql = sql & "RqCartouche.VerifieDate = " & MyReplaceDate(txt18) & " "
sql = sql & "WHERE RqCartouche.T_indiceProjet.pere=" & Me.Tag & ";"
Con.Exequte sql
IdIndiceProjet = Me.Tag
sql = "SELECT T_indiceProjet.Id_Pieces FROM T_indiceProjet "
sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(sql)
IdPieces = Rs!Id_Pieces
sql = "SELECT T_Pieces.IdProjet FROM T_Pieces "
sql = sql & "WHERE T_Pieces.Id=" & IdPieces & ";"
Set Rs = Con.OpenRecordSet(sql)
IdProjet = Rs!IdProjet


 ModifierUnCartouche Me.Tag
Noquite = False
Me.Hide
End Sub

Private Sub CommandButton8_Click()
Noquite = False
Me.Hide
End Sub

Private Sub UserForm_Activate()
Noquite = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite
End Sub
