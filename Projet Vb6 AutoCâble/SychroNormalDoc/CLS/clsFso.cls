VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Fso As New FileSystemObject
Dim MesRep As New Collection
Dim MyPath As String
Public Sub ChargeRep(Rep As String)
MyPath = Rep

Set Fd = Fso.GetFolder(Rep)

   Set fc = Fd.SubFolders
    For Each f1 In fc
    Debug.Print f1
    CreatNewCle (f1)
   
    Next
End Sub
Public Sub CreatNewCle(Rep As String)
Dim NewRem As New clsFso
MesRep.Add NewRem
Set NewRem = Nothing
MesRep(MesRep.Count).ChargeRep (Rep)
End Sub
Public Sub ChangeNormalDoc()
Dim f, f1, fc, s
On Error Resume Next
If Nclick = True Then Exit Sub
  DoEvents
    Set f = Fso.GetFolder(MyPath)
    Set fc = f.Files
    For Each f1 In fc
    Debug.Print f1
    If InStr(f1.Name, ".doc") <> 0 Then
   
      NormalDoc f1.Path
    End If
'
    Next
End Sub
Public Sub ScanMesRep()
Dim I As Long
For I = 1 To MesRep.Count
MesRep(I).ScanMesRep
Next
ChangeNormalDoc
Set Fso = Nothing
End Sub
Sub NormalDoc(Fichier As String)
Dim MywWord As New Word.Application
MywWord.Visible = True
MywWord.Documents.Open Fichier
With MywWord.ActiveDocument
        .UpdateStylesOnOpen = False
        .AttachedTemplate = _
            "\\192.168.1.194\Autocable Access\AutoCable Client\Modèles\ModèleWord\Normal.dot"
        .XMLSchemaReferences.AutomaticValidation = True
        .XMLSchemaReferences.AllowSaveAsXMLWithoutValidation = False
    End With
MywWord.ActiveDocument.Save
MywWord.ActiveDocument.Close
MywWord.Quit
Set MywWord = Nothing
End Sub
'
' NormalDoc Macro
' Macro enregistrée le 14/01/2008 par robert.durupt
'
    
