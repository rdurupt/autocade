Attribute VB_Name = "FunctionWord"
Dim L As Long
Dim C As Long
Public Sub MyWordSaveAs(Path As String)
Dim fso As FileSystemObject
Set fso = New FileSystemObject
If fso.FileExists(Path & "_ETIQUETTE.DOC") = True Then fso.DeleteFile Path & "_ETIQUETTE.DOC"
MyWord.ActiveDocument.SaveAs Path & "_ETIQUETTE"
MyWord.Quit False
Set MyWord = Nothing
Set fso = Nothing
End Sub

Public Sub WordCopyCase(L As Long, C As Long)
'********************************************************************************************
'Place dans la presse papier le contenu d'une cellule (Colonne et ligne) d'un tableau Word.
'********************************************************************************************
MyWord.ActiveDocument.Tables(1).Cell(L, C).Select
MyWord.Selection.Copy
End Sub
Public Sub WordPasteCase(L As Long, C As Long)
'********************************************************************************************
'Copie dans d'une cellule (Colonne et ligne) d'un tableau Word le contenu  du presse papier
'********************************************************************************************
MyWord.ActiveDocument.Tables(1).Cell(L, C).Select
MyWord.Selection.Paste
End Sub
Public Sub WordInsertLigneTableau(L As Long, C As Long)
'***********************************
'Insert une ligne à un tableau Word.
'***********************************
MyWord.ActiveDocument.Tables(1).Cell(L - 1, C).Select
    MyWord.Selection.InsertRowsBelow 1
End Sub

Public Sub WordReplaceText(L As Long, C As Long, Champ As String, ReplaceText As String)
'***********************************************
'Replace le nom d'un champ dans un tableau Word.
'***********************************************
On Error GoTo Fin
MyWord.ActiveDocument.Tables(1).Cell(L, C).Select
    MyWord.Selection.Find.ClearFormatting
    MyWord.Selection.Find.Replacement.ClearFormatting
    With MyWord.Selection.Find
        .Text = Champ
        .Replacement.Text = ReplaceText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    MyWord.Selection.Find.Execute
    With MyWord.Selection
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
        .Find.Execute Replace:=wdReplaceOne
        If .Find.Forward = True Then
            .Collapse Direction:=wdCollapseEnd
        Else
            .Collapse Direction:=wdCollapseStart
        End If
'        .Find.Execute
    End With
    Exit Sub
Fin:
    MsgBox Err.Description
    Err.Clear
End Sub
Public Function WordNewDoc(WordModel As String) As Word.Document
'*****************************************
'Créer un nouveau document suivant modèle.
'*****************************************
L = 0
C = 0
Set WordNewDoc = New Word.Document

    Set WordNewDoc = MyWord.Documents.Add(Template:=WordModel, NewTemplate:=False, DocumentType:=0)
End Function
Public Sub CreerEtiquette(App)
'**********************************************************
'Permet de créer une étiquette et de renseigner les champs.
'**********************************************************
If C = 0 Then C = 1
C = C + 1

If L = 0 Then L = 1
If C = 4 Then
    C = 1
    L = L + 1
    WordInsertLigneTableau L, C
End If
WordCopyCase 1, 1
WordPasteCase L, C
For i = LBound(App) To UBound(App)
txt = Space(255)

txt = "" & App(i, 1) & txt
If Len(Trim(App(i, 1))) > 254 Then
txt = Left(txt, 252)
txt = txt & " ?"
End If
txt = Trim(txt)
txt = Replace(txt, Chr(10), "")
txt = Replace(txt, Chr(13), " ")
txt = Replace(txt, "; ,", ";")
'txt = Chr(10)
    WordReplaceText L, C, "" & App(i, 0), "" & txt
Next

End Sub
