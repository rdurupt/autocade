VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Sélection critères :"
   ClientHeight    =   10560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
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
Dim sql As String
Dim IndexCont As Long
Dim IndexListe As Long

IndexListe = 0
If Trim("" & Me.LibEnsemble) = "" Then Exit Sub
If Trim("" & Me.IdEnsemble) = "" Then
    sql = "SELECT  " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib FROM  " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & " "
    sql = sql & "WHERE  " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib= '" & MyReplace(Me.LibEnsemble) & "' "
'    Sql = Sql & "AND  " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".id<> " & Trim("" & Me.IdEnsemble) & ";"
    Set Rs = Con.OpenRecordSet(sql)
     If Rs.EOF = True Then
    sql = "INSERT INTO " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & " ( Lib ) "
    sql = sql & "values( '" & MyReplace(Me.LibEnsemble) & "');"
    Con.Exequte sql
   
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
    sql = "SELECT  " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib FROM  " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & " "
    sql = sql & "WHERE  " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib= '" & MyReplace(Me.LibEnsemble) & "' "
    sql = sql & "AND  " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".id<> " & Trim("" & Me.IdEnsemble) & ";"
    Set Rs = Con.OpenRecordSet(sql)
    If Rs.EOF = True Then
   
        sql = "UPDATE " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & "  SET " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib = '" & MyReplace(Me.LibEnsemble) & "' "
        sql = sql & "WHERE " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".id=" & Me.IdEnsemble & ";"
        Con.Exequte sql
         IndexCont = MyIndexContenu(LstEnsemble)
            MyIndexContenu.Remove LstEnsemble
            MyIndexContenu.Add IndexCont, Me.LibEnsemble
            MyTableau(MyIndexContenu(Me.LibEnsemble)) = Replace(MyTableau(MyIndexContenu(Me.LibEnsemble)), LstEnsemble, LibEnsemble)
            For i = 1 To MyIndexContenu.Count
                If MyTableau(i) <> "" Then
                    txt = txt & MyTableau(i)
                End If
            Next i
            Me.Apercu = txt
    Else
    MsgBox Me.LibEnsemble & " Existe déjà", vbExclamation
    Me.LibEnsemble.SetFocus
    Exit Sub
    End If
End If
sql = "SELECT " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".* FROM " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & "  ORDER BY " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib;"
Set Rs = Con.OpenRecordSet(sql)
Me.LstEnsemble.Clear
'MyIndexContenu.Add "", Me.LibEnsemble
While Rs.EOF = False
    Me.LstEnsemble.AddItem Trim("" & Rs!LIB)
'    Me.LstEnsemble.List(Me.LstEnsemble.LineCount - 1, 1) = Trim("" & Rs!ID)
    Me.LstEnsemble.List(Me.LstEnsemble.ListCount - 1, 1) = Trim("" & Rs!Id)

    If Me.LibEnsemble = Trim("" & Rs!LIB) Then IndexListe = Me.LstEnsemble.ListCount - 1
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
txt = ""
For i = 1 To MyIndexContenu.Count
If MyTableau(i) <> "" Then
    txt = txt & MyTableau(i)
    End If
Next i
Me.Apercu = txt

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
Dim sql As String
Dim Rs As Recordset
IndexListe = 0
If Trim("" & Me.IdEnsemble) = "" Then Exit Sub
If MsgBox("Voulez vous vraiment supprimer : " & Me.LibEnsemble, vbQuestion + vbYesNo) = vbNo Then
    Exit Sub
End If
MyTableau(MyIndexContenu(Me.LibEnsemble)) = ""

sql = "DELETE " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".* FROM " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & " "
sql = sql & "WHERE " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".id=" & Me.IdEnsemble & ";"
Con.Exequte sql
sql = "SELECT " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".* FROM " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & " ORDER BY " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib;"
Set Rs = Con.OpenRecordSet(sql)
Me.LstEnsemble.Clear
While Rs.EOF = False
    Me.LstEnsemble.AddItem Trim("" & Rs!LIB)
'    Me.LstEnsemble.List(Me.LstEnsemble.LineCount - 1, 1) = Trim("" & Rs!ID)
    Me.LstEnsemble.List(Me.LstEnsemble.ListCount - 1, 1) = Trim("" & Rs!Id)

    If Me.LibEnsemble = Trim("" & Rs!LIB) Then IndexListe = Me.LstEnsemble.ListCount - 1
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
For i = 1 To MyIndexContenu.Count
If MyTableau(i) <> "" Then
    txt = txt & MyTableau(i)
    End If
Next i
Me.Apercu = txt

End Sub

Private Sub CommandButton5_Click()
Dim sql As String
Dim Rs As Recordset
If Trim("" & Me.LibSousEnsemble) = "" Then Exit Sub
If Trim("" & Me.IdEnSoussemble) = "" Then
    sql = "SELECT Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".* "
    sql = sql & "FROM Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & " "
    sql = sql & "WHERE Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib='" & MyReplace(Me.LibSousEnsemble) & "' "
    sql = sql & "AND Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Id_Ensemble=" & Me.LstEnsemble.List(Me.LstEnsemble.ListIndex, 1) & ";"
    Set Rs = Con.OpenRecordSet(sql)
    If Rs.EOF = True Then
        sql = "INSERT INTO Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & " ( Id_Ensemble, Lib ) "
        sql = sql & "Values (" & Me.LstEnsemble.List(Me.LstEnsemble.ListIndex, 1) & " , '" & MyReplace(Me.LibSousEnsemble) & "');"
        Con.Exequte sql
    Else
        MsgBox Me.LibSousEnsemble & " Existe déjà", vbExclamation
        Me.LibSousEnsemble.SetFocus
        Exit Sub
    End If
Else
    sql = "SELECT Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".id, "
    sql = sql & "Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Id_Ensemble, "
    sql = sql & "Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib "
    sql = sql & "FROM Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & " "
    sql = sql & "WHERE Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".id<> " & Me.IdEnSoussemble & " "
    sql = sql & "AND Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Id_Ensemble=" & Me.LstEnsemble.List(Me.LstEnsemble.ListIndex, 1) & ""
    sql = sql & "AND Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib='" & MyReplace(Me.LibSousEnsemble) & "';"
    Set Rs = Con.OpenRecordSet(sql)
    If Rs.EOF = True Then
        sql = "UPDATE Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & " SET Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib = '" & MyReplace(Me.LibSousEnsemble) & "' "
        sql = sql & "WHERE Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".id=" & Me.IdEnSoussemble & " ;"
        Con.Exequte sql

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
Dim sql As String
Dim Rs As Recordset
If Trim("" & Me.IdEnSoussemble) = "" Then Exit Sub
If MsgBox("Voulez vous vraiment supprimer : " & Me.LibSousEnsemble, vbQuestion + vbYesNo) = vbNo Then
    Exit Sub
End If
sql = "DELETE Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".* FROM Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & " "
sql = sql & "WHERE sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".id=" & Me.IdEnSoussemble & ";"
Con.Exequte sql
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
Dim sql As String
Dim IndexListe As Long
Dim txt
Me.ListBox1.Clear
Me.ListBox2.Clear

IndexListe = 0
ChangeLstEnsemble Me.LstEnsemble.Text
txt = Split(TextContenu, "§")
For i = 1 To UBound(txt)
sql = "SELECT Sous"
sql = sql & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1)
sql = sql & ".* FROM Sous"
sql = sql & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1)
sql = sql & " WHERE Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1)
sql = sql & ".Id_Ensemble=" & Me.LstEnsemble.List(Me.LstEnsemble.ListIndex, 1) & " "
sql = sql & "AND Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib"
sql = sql & "='" & MyReplace("" & txt(i)) & "'"
sql = sql & "ORDER BY Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1)
sql = sql & ".Lib;"
Set Rs = Con.OpenRecordSet(sql)
If Rs.EOF = False Then
    Me.ListBox1.AddItem Trim("" & Rs!LIB)
    Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = Trim("" & Rs!Id)
End If
Next i





sql = "SELECT Sous"
sql = sql & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1)
sql = sql & ".* FROM Sous"
sql = sql & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1)
sql = sql & " WHERE Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1)
sql = sql & ".Id_Ensemble=" & Me.LstEnsemble.List(Me.LstEnsemble.ListIndex, 1)
sql = sql & "  ORDER BY Sous" & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1)
sql = sql & ".Lib;"
Set Rs = Con.OpenRecordSet(sql)
While Rs.EOF = False
If InStr(TextContenu, "§" & Trim("" & Rs!LIB) & "§") = 0 Then
    Me.ListBox2.AddItem Trim("" & Rs!LIB)
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
Dim Rs As Recordset
Dim sql As String
Dim IndexListe As Long
Dim IsCollection As String
IndexListe = 0

sql = "SELECT " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".* FROM " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & "  ORDER BY " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib;"
Set Rs = Con.OpenRecordSet(sql)
Me.LstEnsemble.Clear
While Rs.EOF = False
    Me.LstEnsemble.AddItem Trim("" & Rs!LIB)
'    Me.LstEnsemble.List(Me.LstEnsemble.LineCount - 1, 1) = Trim("" & Rs!ID)
    Me.LstEnsemble.List(Me.LstEnsemble.ListCount - 1, 1) = Trim("" & Rs!Id)

    If TextSelect = Trim("" & Rs!LIB) Then IndexListe = Me.LstEnsemble.ListCount - 1
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
If Me.LstEnsemble.ListCount <> 0 Then
Me.LstEnsemble.ListIndex = IndexListe
End If
Me.Frame5.Enabled = Admin
Me.Frame6.Enabled = Admin
End Sub

Public Sub Charger(txt As Object, Separateur As String, Table As String, Optional SeparateurLigne As String)
Dim Texte
Dim sql As String
Dim Rs As Recordset
Dim i As Long

LabEnseble.Caption = Table
DoEvents
MySeparateurAutre = SeparateurLigne
sql = "SELECT " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".* FROM " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & "  ORDER BY " & Left(Me.LabEnseble.Caption, Len(Me.LabEnseble.Caption) - 1) & ".Lib;"
Set Rs = Con.OpenRecordSet(sql)
i = 0
Set MyIndexContenu = Nothing
Set MyIndexContenu = New Collection
While Rs.EOF = False
    i = i + 1
    MyIndexContenu.Add i, Trim("" & Rs!LIB)
    Rs.MoveNext
Wend
ReDim MyTableau(i)

Texte = Split(txt.Text, Separateur)
Set MyTxt = txt
 MySeparateur = Separateur

TextContenu = ""
If UBound(Texte) > -1 Then
TextSelect = Trim(Texte(0))
For i = 0 To UBound(Texte) - 1
TextSelect = IsColect("" & TextSelect, SeparateurAutre(MySeparateurAutre, Trim("" & Texte(i)), 0))
MyTableau(MyIndexContenu(TextSelect)) = MyTableau(MyIndexContenu(TextSelect)) & SeparateurAutre(MySeparateurAutre, Trim("" & Texte(i)), 1) & MySeparateur
If InStr(1, MyTableau(MyIndexContenu(TextSelect)), TextSelect) = 0 Then MyTableau(MyIndexContenu(TextSelect)) = TextSelect & MySeparateur & MyTableau(MyIndexContenu(TextSelect))
Next i
End If
txt2 = ""
For i = 1 To MyIndexContenu.Count
If MyTableau(i) <> "" Then
    txt2 = txt2 & MyTableau(i)
    End If
Next i
Me.Apercu = txt

Me.Show
End Sub
Sub Valider()
MyTxt = Me.Apercu
Set MyTxt = Nothing
End Sub
Sub Appliquer(LstEnsemble As String)
Dim txt
If MySeparateurAutre = "" Then
    txt = LstEnsemble & MySeparateur
Else
    txt = ""
End If
For i = 0 To Me.ListBox1.ListCount - 1
    If MySeparateurAutre = "" Then
        txt = txt & Me.ListBox1.List(i, 0) & MySeparateur
    Else
        txt = txt & LstEnsemble & MySeparateurAutre & Me.ListBox1.List(i, 0) & MySeparateur
    End If
Next i
If txt = "" Then txt = LstEnsemble & MySeparateur
 MyTableau(MyIndexContenu(LstEnsemble)) = txt
txt = ""
For i = 1 To MyIndexContenu.Count
If MyTableau(i) <> "" Then
    txt = txt & MyTableau(i)
    End If
Next i
Me.Apercu = txt

End Sub
Sub ChangeLstEnsemble(LstEnsemble As String)
Dim Texte



Texte = Split(MyTableau(MyIndexContenu(LstEnsemble)), MySeparateur)



TextContenu = ""
If UBound(Texte) > -1 Then
TextSelect = SeparateurAutre(MySeparateurAutre, Trim("" & Texte(0)), 0)
If MySeparateurAutre <> "" Then
    For i = 0 To UBound(Texte) - 1
        TextContenu = TextContenu & SeparateurAutre(MySeparateurAutre, Trim(Texte(i)), 1) & "§"
    Next i
Else
    For i = 1 To UBound(Texte)
        TextContenu = TextContenu & SeparateurAutre(MySeparateurAutre, Trim(Texte(i)), 1) & "§"
    Next i
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
Function SeparateurAutre(Separateur As String, txt, NumFild As Integer) As String
Dim MyTxt

If Separateur <> "" Then
    MyTxt = Split(txt, Separateur)
    If UBound(MyTxt) = 0 Then
        SeparateurAutre = txt
    Else
         SeparateurAutre = MyTxt(NumFild)
    End If
Else
    SeparateurAutre = txt
End If
End Function
