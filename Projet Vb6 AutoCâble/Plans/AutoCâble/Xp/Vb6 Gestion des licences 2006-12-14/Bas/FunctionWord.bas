Attribute VB_Name = "FunctionWord"
Dim L As Long
Dim C As Long
Dim L2 As Long
Dim C2 As Long

Sub MyWordSaveAs(Path As String, Optional Save As Boolean = True, Optional Options As String)
Dim LibOption As String
Dim Fso As FileSystemObject
Set Fso = New FileSystemObject

If Options <> "" Then LibOption = "_" & Options
If Fso.FileExists(Path & "_ETIQUETTE" & LibOption & ".DOC") = True Then Fso.DeleteFile Path & "_ETIQUETTE" & LibOption & ".DOC"
If Save = True Then MyWordDoc.SaveAs Path & "_ETIQUETTE" & LibOption

If Fso.FileExists(Path & "_ETIQUETTE_Marquage" & LibOption & ".DOC") = True Then Fso.DeleteFile Path & "_ETIQUETTE_MARQUAGE" & LibOption & ".DOC"
If Save = True Then MyWordDoc2.SaveAs Path & "_ETIQUETTE_MARQUAGE" & LibOption
MyWord.Quit False
Set MyWord = Nothing
Set Fso = Nothing
End Sub

Sub WordCopyCase(MyWordDoc As Object, L As Long, C As Long)
'********************************************************************************************
'Place dans la presse papier le contenu d'une cellule (Colonne et ligne) d'un tableau Word.
'********************************************************************************************
'MyWordDoc.Application.Visible = True
MyWordDoc.Select
MyWordDoc.Tables(1).Cell(L, C).Select
MyWord.Selection.Copy
End Sub
Sub WordPasteCase(MyWordDoc As Object, L As Long, C As Long)
'********************************************************************************************
'Copie dans d'une cellule (Colonne et ligne) d'un tableau Word le contenu  du presse papier
'********************************************************************************************
On Error Resume Next
MyWordDoc.Select
MyWordDoc.Tables(1).Cell(L, C).Select
MyWord.Selection.Paste
MyWordDoc.Tables(1).Cell(L, 4).Select
If Err = 0 Then MyWord.Selection.Columns.Delete
End Sub
Sub WordInsertLigneTableau(MyWordDoc As Object, L As Long, C As Long)
'***********************************
'Insert une ligne à un tableau Word.
'***********************************
'MyWordDoc.Application.Visible = True
MyWordDoc.Select
'MyWordDoc.Tables(1).Cell(L - 1, C).Select
    MyWord.Selection.InsertRowsBelow 1
End Sub

Sub WordReplaceText(MyWordDoc As Object, L As Long, C As Long, Champ As String, ReplaceText As String)
'***********************************************
'Replace le nom d'un champ dans un tableau Word.
'***********************************************
On Error GoTo Fin
MyWordDoc.Select
MyWordDoc.Tables(1).Cell(L, C).Select
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
            .Collapse Direction:=1
        Else
            .Collapse Direction:=0
        End If
        .Find.Execute Replace:=1
        If .Find.Forward = True Then
            .Collapse Direction:=0
        Else
            .Collapse Direction:=1
        End If
'        .Find.Execute
    End With
    Exit Sub
Fin:
    MsgBox Err.Description
    Err.Clear
End Sub
Sub RemplaceWord(MyWord As Object, Champ As String, Valeur As String)
On Error GoTo Fin
    
     MyWord.Selection.Find.ClearFormatting
    MyWord.Selection.Find.Replacement.ClearFormatting
    With MyWord.Selection.Find
    
        .Text = Champ
        .Replacement.Text = Valeur
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
            .Collapse Direction:=1
        Else
            .Collapse Direction:=0
        End If
        .Find.Execute Replace:=1
        If .Find.Forward = True Then
            .Collapse Direction:=0
        Else
            .Collapse Direction:=1
        End If
'        .Find.Execute
    End With
    Exit Sub
Fin:
    MsgBox Err.Description
    Err.Clear
End Sub
Function WordNewDocApp(WordModel As String, MyWord As Object) As Object
'*****************************************
'Créer un nouveau document suivant modèle.
'*****************************************
L = 0
C = 0
Set WordNewDocApp = CreateObject("Word.Document")

    Set WordNewDocApp = MyWord.Documents.Add(Template:=WordModel, NewTemplate:=False, DocumentType:=0)
End Function

Function WordNewDoc(WordModel As String) As Object
'*****************************************
'Créer un nouveau document suivant modèle.
'*****************************************
L = 0
C = 0
Set WordNewDoc = CreateObject("Word.Document")

'    Set WordNewDoc = MyWord.Documents.Add(Template:=WordModel, NewTemplate:=False, DocumentType:=0)
End Function
Sub CreerEtiquette(App)
Dim CMax As Long
Dim C2Max As Long
'**********************************************************
'Permet de créer une étiquette et de renseigner les champs.
'**********************************************************
If Trim("" & App(0, 1)) = "" Then Exit Sub
If C = 0 Then C = 1
C = C + 1


If L = 0 Then L = 1
If C = Val(GetDefault("NbPalachEt", "3")) + 1 Then
    C = 1
    L = L + 1
    WordInsertLigneTableau MyWordDoc, L, C
    
End If

If C2 = 0 Then C2 = 1
C2 = C2 + 1

If L2 = 0 Then L2 = 1
If C2 = Val(GetDefault("NbPlacheEtiM", "3")) + 1 Then
    C2 = 1
    L2 = L2 + 1
    WordInsertLigneTableau MyWordDoc2, L2, C

    
End If

'MyWordDoc.Application.Visible = True
WordCopyCase MyWordDoc, 1, 1
WordPasteCase MyWordDoc, L, C

WordCopyCase MyWordDoc2, 1, 1
WordPasteCase MyWordDoc2, L2, C2
For I = LBound(App) To UBound(App)
Txt = Space(255)

Txt = "" & App(I, 1) & Txt

If Len(Trim(App(I, 1))) > 254 Then
Txt = Left(Txt, 252)
Txt = Txt & " ?"
End If
Txt = Trim(Txt)
Txt = Replace(Txt, Chr(10), "")
Txt = Replace(Txt, Chr(13), " ")
Txt = Replace(Txt, "; ,", ";")
'txt = Chr(10)
    WordReplaceText MyWordDoc, L, C, "" & App(I, 0), "" & Txt
     WordReplaceText MyWordDoc2, L2, C2, "" & App(I, 0), "" & Txt
Next

End Sub

Sub OuvreEnteteWord(MyWord As Object)
'
' OuvreEntete Macro
' Macro enregistrée le 09/06/2005 par robert.durupt
'
    If MyWord.ActiveWindow.View.SplitSpecial <> 0 Then
        MyWord.ActiveWindow.Panes(2).Close
    End If
    If MyWord.ActiveWindow.ActivePane.View.Type = 1 Or MyWord.ActiveWindow. _
         ActivePane.View.Type = 2 Then
         MyWord.ActiveWindow.ActivePane.View.Type = 3
    End If
     MyWord.ActiveWindow.ActivePane.View.SeekView = 9
End Sub
Sub FermeEnteteWord(MyWord As Object)
'
' FermeEntete Macro
' Macro enregistrée le 09/06/2005 par robert.durupt
'
    MyWord.ActiveWindow.ActivePane.View.SeekView = 0
End Sub
Sub OuvrePiedWord(MyWord As Object)
'
' OuvrePied Macro
' Macro enregistrée le 09/06/2005 par robert.durupt
'
    If MyWord.ActiveWindow.View.SplitSpecial <> 0 Then
        MyWord.ActiveWindow.Panes(2).Close
    End If
    If MyWord.ActiveWindow.ActivePane.View.Type = 1 Or ActiveWindow. _
        MyWord.ActivePane.View.Type = 2 Then
        MyWord.ActiveWindow.ActivePane.View.Type = 3
    End If
    MyWord.ActiveWindow.ActivePane.View.SeekView = 9
    If MyWord.Selection.HeaderFooter.IsHeader = True Then
        MyWord.ActiveWindow.ActivePane.View.SeekView = 10
    Else
        MyWord.ActiveWindow.ActivePane.View.SeekView = 9
    End If
End Sub

