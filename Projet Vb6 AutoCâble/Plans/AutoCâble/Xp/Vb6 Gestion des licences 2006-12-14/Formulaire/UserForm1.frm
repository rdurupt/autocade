VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sélection critères :"
   ClientHeight    =   9480
   ClientLeft      =   30
   ClientTop       =   195
   ClientWidth     =   6840
   Icon            =   "UserForm1.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "UserForm1.dsx":08CA
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TextSelect
Dim TextContenu
Dim MyTxt As Object
Dim MySeparateur As String
Dim MySeparateurAutre As String
Dim MyIndexContenu As New Collection
Dim MyTableau() As String

Private Sub Apercu_Click()

End Sub

Private Sub CommandButton1_Click()
Dim Rs As Recordset
Dim Sql As String
Dim IndexCont As Long
Dim IndexListe As Long

IndexListe = 0
If Trim("" & Me.LibEnsemble) = "" Then Exit Sub
If Trim("" & Me.IdEnsemble) = "" Then
    Sql = "SELECT  " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib FROM  " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & " "
    Sql = Sql & "WHERE  " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib= '" & MyReplace(Me.LibEnsemble) & "' "
'    Sql = Sql & "AND  " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".id<> " & Trim("" & Me.IdEnsemble) & ";"
    Set Rs = Con.OpenRecordSet(Sql)
     If Rs.EOF = True Then
    Sql = "INSERT INTO " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & " ( Lib ) "
    Sql = Sql & "values( '" & MyReplace(Me.LibEnsemble) & "');"
    Con.Execute Sql
   
    ReDim Preserve MyTableau(UBound(MyTableau) + 1)
    On Error Resume Next
    MyIndexContenu.Add UBound(MyTableau), Me.LibEnsemble
    On Error GoTo 0
    Else
    MsgBox Me.LibEnsemble & " Existe déjà", vbExclamation
    Me.LibEnsemble.SetFocus
    Exit Sub
    End If
Else
    Sql = "SELECT  " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib FROM  " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & " "
    Sql = Sql & "WHERE  " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib= '" & MyReplace(Me.LibEnsemble) & "' "
    Sql = Sql & "AND  " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".id<> " & Trim("" & Me.IdEnsemble) & ";"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = True Then
   
        Sql = "UPDATE " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & "  SET " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib = '" & MyReplace(Me.LibEnsemble) & "' "
        Sql = Sql & "WHERE " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".id=" & Me.IdEnsemble & ";"
        Con.Execute Sql
         IndexCont = MyIndexContenu(LstEnsemble)
            MyIndexContenu.Remove LstEnsemble
            MyIndexContenu.Add IndexCont, Me.LibEnsemble
            MyTableau(MyIndexContenu(Me.LibEnsemble)) = Replace(MyTableau(MyIndexContenu(Me.LibEnsemble)), LstEnsemble, LibEnsemble)
            For I = 1 To MyIndexContenu.Count
                If MyTableau(I) <> "" Then
                    Txt = Txt & MyTableau(I)
                End If
            Next I
            Me.Apercu = Txt
    Else
    MsgBox Me.LibEnsemble & " Existe déjà", vbExclamation
    Me.LibEnsemble.SetFocus
    Exit Sub
    End If
End If
Sql = "SELECT " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".* FROM " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & "  ORDER BY " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib;"
Set Rs = Con.OpenRecordSet(Sql)
Me.LstEnsemble.Clear
'MyIndexContenu.Add "", Me.LibEnsemble
While Rs.EOF = False
    Me.LstEnsemble.AddItem Trim("" & Rs!Lib)
'    Me.LstEnsemble.List(Me.LstEnsemble.LineCount - 1, 1) = Trim("" & Rs!ID)
    Me.LstEnsemble.List(Me.LstEnsemble.ListCount - 1, 1) = Trim("" & Rs!Id)

    If Me.LibEnsemble = Trim("" & Rs!Lib) Then IndexListe = Me.LstEnsemble.ListCount - 1
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
Me.LstEnsemble.ListIndex = IndexListe
Me.LibEnsemble = ""
Me.IdEnsemble = ""
End Sub

Private Sub CommandButton10_Click()
If Me.ListBox1.ListIndex = -1 Then
    MsgBox "Vous devez sélectionner [Critères à insérer]", vbExclamation
    Exit Sub
End If
Me.ListBox2.AddItem Me.ListBox1.List(Me.ListBox1.ListIndex, 0)
Me.ListBox2.List(Me.ListBox2.ListCount - 1, 1) = Me.ListBox1.List(Me.ListBox1.ListIndex, 1)
 Me.ListBox1.RemoveItem Me.ListBox1.ListIndex


End Sub

Private Sub CommandButton11_Click()
Set MyIndexContenu = Nothing
Valider
Unload Me
End Sub

Private Sub CommandButton12_Click()
Unload Me
End Sub

Private Sub CommandButton13_Click()
    Appliquer LstEnsemble.Text
End Sub

Private Sub CommandButton14_Click()
MyTableau(MyIndexContenu(Me.LstEnsemble.Text)) = ""
Txt = ""
For I = 1 To MyIndexContenu.Count
If MyTableau(I) <> "" Then
    Txt = Txt & MyTableau(I)
    End If
Next I
Me.Apercu = Txt

LstEnsemble_Click
End Sub



Private Sub CommandButton2_Click()
 Me.LibEnsemble = Me.LstEnsemble.List(Me.LstEnsemble.ListIndex, 0)
 Me.IdEnsemble = Me.LstEnsemble.List(Me.LstEnsemble.ListIndex, 1)
 
End Sub

Private Sub CommandButton3_Click()
Me.IdEnsemble = ""
Me.LibEnsemble = ""
End Sub

Private Sub CommandButton4_Click()
Dim Sql As String
Dim Rs As Recordset
IndexListe = 0
If Trim("" & Me.IdEnsemble) = "" Then Exit Sub
If MsgBox("Voulez vous vraiment supprimer : " & Me.LibEnsemble, vbQuestion + vbYesNo) = vbNo Then
    Exit Sub
End If
MyTableau(MyIndexContenu(Me.LibEnsemble)) = ""

Sql = "DELETE " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".* FROM " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & " "
Sql = Sql & "WHERE " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".id=" & Me.IdEnsemble & ";"
Con.Execute Sql
Sql = "SELECT " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".* FROM " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & " ORDER BY " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib;"
Set Rs = Con.OpenRecordSet(Sql)
Me.LstEnsemble.Clear
While Rs.EOF = False
    Me.LstEnsemble.AddItem Trim("" & Rs!Lib)
'    Me.LstEnsemble.List(Me.LstEnsemble.LineCount - 1, 1) = Trim("" & Rs!ID)
    Me.LstEnsemble.List(Me.LstEnsemble.ListCount - 1, 1) = Trim("" & Rs!Id)

    If Me.LibEnsemble = Trim("" & Rs!Lib) Then IndexListe = Me.LstEnsemble.ListCount - 1
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
'MyIndexContenu.Remove Me.LibEnsemble
Me.LibEnsemble = ""
Me.IdEnsemble = ""
If Me.LstEnsemble.ListCount <> 0 Then
    Me.LstEnsemble.ListIndex = IndexListe
    LstEnsemble_Click

End If
For I = 1 To MyIndexContenu.Count
If MyTableau(I) <> "" Then
    Txt = Txt & MyTableau(I)
    End If
Next I
Me.Apercu = Txt

End Sub

Private Sub CommandButton5_Click()
Dim Sql As String
Dim Rs As Recordset
If Trim("" & Me.LibSousEnsemble) = "" Then Exit Sub
If Trim("" & Me.IdEnSoussemble) = "" Then
    Sql = "SELECT Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".* "
    Sql = Sql & "FROM Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & " "
    Sql = Sql & "WHERE Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib='" & MyReplace(Me.LibSousEnsemble) & "' "
    Sql = Sql & "AND Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Id_Ensemble=" & Me.LstEnsemble.List(Me.LstEnsemble.ListIndex, 1) & ";"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = True Then
        Sql = "INSERT INTO Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & " ( Id_Ensemble, Lib ) "
        Sql = Sql & "Values (" & Me.LstEnsemble.List(Me.LstEnsemble.ListIndex, 1) & " , '" & MyReplace(Me.LibSousEnsemble) & "');"
        Con.Execute Sql
    Else
        MsgBox Me.LibSousEnsemble & " Existe déjà", vbExclamation
        Me.LibSousEnsemble.SetFocus
        Exit Sub
    End If
Else
    Sql = "SELECT Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".id, "
    Sql = Sql & "Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Id_Ensemble, "
    Sql = Sql & "Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib "
    Sql = Sql & "FROM Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & " "
    Sql = Sql & "WHERE Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".id<> " & Me.IdEnSoussemble & " "
    Sql = Sql & "AND Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Id_Ensemble=" & Me.LstEnsemble.List(Me.LstEnsemble.ListIndex, 1) & ""
    Sql = Sql & "AND Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib='" & MyReplace(Me.LibSousEnsemble) & "';"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = True Then
        Sql = "UPDATE Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & " SET Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib = '" & MyReplace(Me.LibSousEnsemble) & "' "
        Sql = Sql & "WHERE Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".id=" & Me.IdEnSoussemble & " ;"
        Con.Execute Sql

    Else
    End If
    
End If
Set Rs = Con.CloseRecordSet(Rs)
Me.IdEnSoussemble = ""
Me.LibSousEnsemble = ""
LstEnsemble_Click
End Sub

Private Sub CommandButton6_Click()
Me.LibSousEnsemble = ""
Me.IdEnSoussemble = ""
End Sub

Private Sub CommandButton7_Click()
Dim Sql As String
Dim Rs As Recordset
If Trim("" & Me.IdEnSoussemble) = "" Then Exit Sub
If MsgBox("Voulez vous vraiment supprimer : " & Me.LibSousEnsemble, vbQuestion + vbYesNo) = vbNo Then
    Exit Sub
End If
Sql = "DELETE Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".* FROM Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & " "
Sql = Sql & "WHERE sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".id=" & Me.IdEnSoussemble & ";"
Con.Execute Sql
Me.LibSousEnsemble = ""
Me.IdEnSoussemble = ""
LstEnsemble_Click
End Sub

Private Sub CommandButton8_Click()
If Me.ListBox2.ListIndex = -1 Then
    MsgBox "Vous devez sélectionner [Critères à sélectionner]", vbExclamation
    Exit Sub
End If
Me.LibSousEnsemble = Me.ListBox2.List(Me.ListBox2.ListIndex, 0)
Me.IdEnSoussemble = Me.ListBox2.List(Me.ListBox2.ListIndex, 1)
End Sub

Private Sub CommandButton9_Click()
If Me.ListBox2.ListIndex = -1 Then
    MsgBox "Vous devez sélectionner [Critères à sélectionner]", vbExclamation
    Exit Sub
End If
Me.ListBox1.AddItem Me.ListBox2.List(Me.ListBox2.ListIndex, 0)
Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Me.ListBox2.List(Me.ListBox2.ListIndex, 1)
 Me.ListBox2.RemoveItem Me.ListBox2.ListIndex
End Sub

Private Sub Image1_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
CommandButton10_Click
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
CommandButton9_Click
End Sub

Private Sub LstEnsemble_Click()
Dim Rs As Recordset
Dim Sql As String
Dim IndexListe As Long
Dim Txt
Me.ListBox1.Clear
Me.ListBox2.Clear

IndexListe = 0
ChangeLstEnsemble Me.LstEnsemble.Text
Txt = Split(TextContenu, "§")
For I = 1 To UBound(Txt)
Sql = "SELECT Sous"
Sql = Sql & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1)
Sql = Sql & ".* FROM Sous"
Sql = Sql & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1)
Sql = Sql & " WHERE Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1)
Sql = Sql & ".Id_Ensemble=" & Me.LstEnsemble.List(Me.LstEnsemble.ListIndex, 1) & " "
Sql = Sql & "AND Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib"
Sql = Sql & "='" & MyReplace("" & Txt(I)) & "'"
Sql = Sql & "ORDER BY Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1)
Sql = Sql & ".Lib;"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    Me.ListBox1.AddItem Trim("" & Rs!Lib)
    Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Trim("" & Rs!Id)
End If
Next I





Sql = "SELECT Sous"
Sql = Sql & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1)
Sql = Sql & ".* FROM Sous"
Sql = Sql & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1)
Sql = Sql & " WHERE Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1)
Sql = Sql & ".Id_Ensemble=" & Me.LstEnsemble.List(Me.LstEnsemble.ListIndex, 1)
Sql = Sql & "  ORDER BY Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1)
Sql = Sql & ".Lib;"
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
If InStr(TextContenu, "§" & Trim("" & Rs!Lib) & "§") = 0 Then
    Me.ListBox2.AddItem Trim("" & Rs!Lib)
'    Me.LstEnsemble.List(Me.LstEnsemble.LineCount - 1, 1) = Trim("" & Rs!ID)
    Me.ListBox2.List(Me.ListBox2.ListCount - 1, 1) = Trim("" & Rs!Id)
End If
'    If Me.LibEnsemble = Trim("" & Rs!Lib) Then IndexListe = Me.LstEnsemble.ListCount - 1
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
'Me.LstEnsemble.ListIndex = IndexListe





'TextSelect
End Sub

Private Sub UserForm_Activate()
'Dim Rs As Recordset
'Dim sql As String
'Dim IndexListe As Long
'Dim IsCollection As String
'IndexListe = 0
'
'sql = "SELECT " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".* FROM " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & "  ORDER BY " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib;"
'Set Rs = Con.OpenRecordSet(sql)
'Me.LstEnsemble.Clear
'While Rs.EOF = False
'
'    Me.LstEnsemble.AddItem Trim("" & Rs!Lib)
''    Me.LstEnsemble.List(Me.LstEnsemble.LineCount - 1, 1) = Trim("" & Rs!ID)
'    Me.LstEnsemble.List(Me.LstEnsemble.ListCount - 1, 1) = Trim("" & Rs!Id)
'
'    If TextSelect = Trim("" & Rs!Lib) Then IndexListe = Me.LstEnsemble.ListCount - 1
'    DoEvents
'    Rs.MoveNext
'Wend
'Set Rs = Con.CloseRecordSet(Rs)
'If Me.LstEnsemble.ListCount <> 0 Then
'Me.LstEnsemble.ListIndex = IndexListe
'End If
'DoEvents
End Sub

Public Sub charger(Txt As Object, Separateur As String, Table As String, Optional SeparateurLigne As String)
On Error Resume Next
Dim Texte
Dim Sql As String
Dim Rs As Recordset
Dim I As Long

LabEnseble.Caption = Table
Sql = "SELECT T_Droits.Id_Useur, T_Boutons.Name "
Sql = Sql & "FROM T_Boutons INNER JOIN T_Droits ON T_Boutons.Id = T_Droits.Id_Bouton "
Sql = Sql & "WHERE T_Droits.Id_Useur=" & Id_Users & " "
Sql = Sql & "AND T_Boutons.Name='" & Left(LabEnseble.Caption, Len(LabEnseble.Caption) - 1) & "';"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
Me.Frame5.Enabled = True
Me.Frame6.Enabled = True
Else
Me.Frame5.Enabled = False
Me.Frame6.Enabled = False
End If
DoEvents
MySeparateurAutre = SeparateurLigne
Sql = "SELECT " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".* FROM " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & "  ORDER BY " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib;"
Set Rs = Con.OpenRecordSet(Sql)
I = 0
Set MyIndexContenu = Nothing
Set MyIndexContenu = New Collection
While Rs.EOF = False
DoEvents
    I = I + 1
    MyIndexContenu.Add I, Trim("" & Rs!Lib)
    Rs.MoveNext
Wend
ReDim MyTableau(I)

Texte = Split(Txt.Text, Separateur)
Set MyTxt = Txt
 MySeparateur = Separateur

TextContenu = ""
If UBound(Texte) > -1 Then
TextSelect = Trim(Texte(0))
For I = 0 To UBound(Texte) - 1
TextSelect = IsColect("" & TextSelect, SeparateurAutre(MySeparateurAutre, Trim("" & Texte(I)), 0))
MyTableau(MyIndexContenu(TextSelect)) = MyTableau(MyIndexContenu(TextSelect)) & SeparateurAutre(MySeparateurAutre, Trim("" & Texte(I)), 1) & MySeparateur
If InStr(1, MyTableau(MyIndexContenu(TextSelect)), TextSelect) = 0 Then MyTableau(MyIndexContenu(TextSelect)) = TextSelect & MySeparateur & MyTableau(MyIndexContenu(TextSelect))
Next I
End If
txt2 = ""
For I = 1 To MyIndexContenu.Count
If MyTableau(I) <> "" Then
    txt2 = txt2 & MyTableau(I)
    End If
Next I
Me.Apercu = "" & Txt
IndexListe = 0

Sql = "SELECT " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".* FROM " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & "  ORDER BY " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib;"
Set Rs = Con.OpenRecordSet(Sql)
Me.LstEnsemble.Clear
While Rs.EOF = False

    Me.LstEnsemble.AddItem Trim("" & Rs!Lib)
'    Me.LstEnsemble.List(Me.LstEnsemble.LineCount - 1, 1) = Trim("" & Rs!ID)
    Me.LstEnsemble.List(Me.LstEnsemble.ListCount - 1, 1) = Trim("" & Rs!Id)

    If TextSelect = Trim("" & Rs!Lib) Then IndexListe = Me.LstEnsemble.ListCount - 1
    DoEvents
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
If Me.LstEnsemble.ListCount <> 0 Then
Me.LstEnsemble.ListIndex = IndexListe
End If

Me.Show vbModal
End Sub
Sub Valider()
On Error Resume Next
MyTxt = Me.Apercu
Set MyTxt = Nothing
End Sub
Sub Appliquer(LstEnsemble As String)
Dim Txt
If MySeparateurAutre = "" Then
    Txt = LstEnsemble & MySeparateur
Else
    Txt = ""
End If
For I = 0 To Me.ListBox1.ListCount - 1
    If MySeparateurAutre = "" Then
        Txt = Txt & Me.ListBox1.List(I, 0) & MySeparateur
    Else
        Txt = Txt & LstEnsemble & MySeparateurAutre & Me.ListBox1.List(I, 0) & MySeparateur
    End If
Next I
If Txt = "" Then Txt = LstEnsemble & MySeparateur
 MyTableau(MyIndexContenu(LstEnsemble)) = Txt
Txt = ""
For I = 1 To MyIndexContenu.Count
If MyTableau(I) <> "" Then
    Txt = Txt & MyTableau(I)
    End If
Next I
Me.Apercu = Txt

End Sub
Sub ChangeLstEnsemble(LstEnsemble As String)
Dim Texte



Texte = Split(MyTableau(MyIndexContenu(LstEnsemble)), MySeparateur)



TextContenu = ""
If UBound(Texte) > -1 Then
TextSelect = SeparateurAutre(MySeparateurAutre, Trim("" & Texte(0)), 0)
If MySeparateurAutre <> "" Then
    For I = 0 To UBound(Texte) - 1
        TextContenu = TextContenu & SeparateurAutre(MySeparateurAutre, Trim(Texte(I)), 1) & "§"
    Next I
Else
    For I = 1 To UBound(Texte)
        TextContenu = TextContenu & SeparateurAutre(MySeparateurAutre, Trim(Texte(I)), 1) & "§"
    Next I
End If
TextContenu = "§" & TextContenu
End If
End Sub

Function IsColect(Colect As String, ValColect As String) As String
Dim Num As Long
On Error Resume Next
Num = MyIndexContenu(ValColect)
If Err Then
    IsColect = Colect
Else
    IsColect = ValColect
End If
Err.Clear
On Error GoTo 0
End Function
Function SeparateurAutre(Separateur As String, Txt, NumFild As Integer) As String
Dim MyTxt

If Separateur <> "" Then
    MyTxt = Split(Txt & Separateur, Separateur)
    If UBound(MyTxt) = 0 Then
        SeparateurAutre = Txt
    Else
         SeparateurAutre = MyTxt(NumFild)
    End If
Else
    SeparateurAutre = Txt
End If
End Function
