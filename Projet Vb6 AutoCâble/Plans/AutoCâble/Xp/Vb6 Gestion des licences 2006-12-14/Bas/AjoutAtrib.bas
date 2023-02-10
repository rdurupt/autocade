Attribute VB_Name = "AjoutAtrib"

Public MyLayer
Public MyColor

Public Sub AddAtrib()
Dim Rep
If Dir(App.Path & "\ConnecteurCreatAttributs\Test.Ok") <> "" Then
    MsgBox "La Macro d'ajout Attributs  connecteur est déjà en cour d'exécution", vbInformation
    Exit Sub
End If
Dim Fso As New FileSystemObject
Fso.CreateTextFile App.Path & "\ConnecteurCreatAttributs\Test.Ok"
Rep = Dir(App.Path & "\ConnecteurCreatAttributs\*.DWG")
If Rep = "" Then Exit Sub

Dim i As Long
Dim aa
Dim BB
Dim Trouve As Boolean
Dim MyBloc As String
AutoApp.Visible = True
While Rep <> ""


a = Split(Rep & ".", ".")
a = Split(a(0) & "§", "§")
MyBloc = a(0)
OpenFichier App.Path & "\ConnecteurCreatAttributs\" & Rep
Trouve = False
On Error GoTo Fin
For i = 0 To AutoApp.Documents(0).ModelSpace.Count - 1
Set aa = AutoApp.Documents(0).ModelSpace(i)
 BB = aa.ObjectName
 If BB = "AcDbAttributeDefinition" Then
     If Trim(UCase("" & aa.TagString)) = "DESIGNATION" Then
        MyLayer = aa.Layer
        MyColor = aa.Color
        Exit For
     End If
     
 End If
Next
For i = 0 To AutoApp.Documents(0).ModelSpace.Count - 1
    Set aa = AutoApp.Documents(0).ModelSpace(i)
    BB = aa.ObjectName
    
    If BB = "AcDbText" Then
    Debug.Print aa.TextString
   
   
    If InStr(1, ReplaceTxtAttribut("" & aa.TextString), "XXXXXX") <> 0 Or InStr(1, ReplaceTxtAttribut("" & aa.TextString), "xxxxxxxxxx") <> 0 Or InStr(1, ReplaceTxtAttribut("" & aa.TextString), ReplaceTxtAttribut("" & "ATTENTE REF")) <> 0 Or InStr(1, ReplaceTxtAttribut("" & aa.TextString), ReplaceTxtAttribut("" & "EN ATT")) <> 0 Or InStr(1, ReplaceTxtAttribut("" & aa.TextString), ReplaceTxtAttribut("" & MyBloc)) <> 0 Or InStr(1, UCase("" & aa.TextString), "REFERENCE") <> 0 Then
    Trouve = True
        AddAtribut aa, MyBloc
        aa.Delete
        Set aa = Nothing
        'AutoApp.Documents(0).Application.ZoomAll
        Rep = App.Path & "\SaveConnecteurCreatAttributs\" & Rep
       SaveAs Rep
       Exit For
      
    End If
    

    End If
     
    If BB = "AcDbMText" Then
    Debug.Print aa.TextString
   
 If InStr(1, ReplaceTxtAttribut("" & aa.TextString), "XXXXXX") <> 0 Or InStr(1, ReplaceTxtAttribut("" & aa.TextString), "xxxxxxxxxx") <> 0 Or InStr(1, ReplaceTxtAttribut("" & aa.TextString), ReplaceTxtAttribut("" & "ATTENTE REF")) <> 0 Or InStr(1, ReplaceTxtAttribut("" & aa.TextString), ReplaceTxtAttribut("" & "EN ATT")) <> 0 Or InStr(1, ReplaceTxtAttribut("" & aa.TextString), ReplaceTxtAttribut("" & MyBloc)) <> 0 Or InStr(1, UCase("" & aa.TextString), "REFERENCE") <> 0 Then
 Trouve = True
        AddAtribut aa, MyBloc
        aa.Delete
        Set aa = Nothing
        'AutoApp.Documents(0).Application.ZoomAll
        Rep = App.Path & "\SaveConnecteurCreatAttributs\" & Rep
       SaveAs Rep
       Exit For
      
    End If
    

    End If
    If BB = "AcDbAttributeDefinition" Then
    Debug.Print aa.TextString
        If InStr(1, ReplaceTxtAttribut(UCase("" & aa.TagString)), ReplaceTxtAttribut("" & MyBloc)) <> 0 Or InStr(1, UCase("" & aa.TagString), "REFERENCE") <> 0 Then
            aa.TagString = "RefConnecteurCli"
            aa.PromptString = UCase("Ref Connecteur Client :")
            
             aa.TextString = "" & MyBloc
            aa.Layer = MyLayer
            aa.Color = MyColor

             Trouve = True
            If Fso.FileExists(App.Path & "\SaveConnecteurCreatAttributs\" & Rep) = True Then
                Fso.DeleteFile App.Path & "\SaveConnecteurCreatAttributs\" & Rep
            End If
           SaveAs App.Path & "\SaveConnecteurCreatAttributs\" & Rep
           
           Exit For
        End If
    End If
Next

If Trouve = False Then AutoApp.Documents(0).Close , False
Fin:

Err.Clear

 Rep = Dir
 Wend
 Rep = Dir(App.Path & "\SaveConnecteurCreatAttributs\*.DWG")
While Rep <> ""
    If Fso.FileExists(App.Path & "\ConnecteurCreatAttributs\" & Rep) = True Then
        Fso.DeleteFile App.Path & "\ConnecteurCreatAttributs\" & Rep
        DoEvents
    End If
    Rep = Dir
Wend

'
 If Fso.FileExists(App.Path & "\ConnecteurCreatAttributs\Test.Ok") = True Then Fso.DeleteFile App.Path & "\ConnecteurCreatAttributs\Test.Ok"
 Set Fso = Nothing
 AutoApp.Visible = False
End Sub
Sub AddAtribut(Etiqqette, MyBloc As String)
Dim InsertPoint(0 To 2) As Double
Dim EE
Dim MyWheet

Set EE = AutoApp.Documents(0).ModelSpace.AddAttribute(Etiqqette.Height, acAttributeModeNormal, "Ref Connecteur Client :", Etiqqette.InsertionPoint, "RefConnecteurCli", MyBloc)
'EE.Color = "Magenta"
EE.Rotation = Etiqqette.Rotation
EE.Layer = MyLayer
EE.Color = MyColor

EE.Alignment = acAlignmentMiddleCenter
'EE = Etiqqette.InsertionPoint
EE.TextAlignmentPoint = Etiqqette.InsertionPoint
'EE.TextAlignmentPoint(1) = aa(1)
End Sub
Function ReplaceTxtAttribut(txt As String)
ReplaceTxtAttribut = txt
ReplaceTxtAttribut = UCase(ReplaceTxtAttribut)
ReplaceTxtAttribut = Replace(ReplaceTxtAttribut, " ", "")
ReplaceTxtAttribut = Replace(ReplaceTxtAttribut, "-", "")
ReplaceTxtAttribut = Replace(ReplaceTxtAttribut, "_", "")
ReplaceTxtAttribut = Replace(ReplaceTxtAttribut, "-", "")
ReplaceTxtAttribut = Replace(ReplaceTxtAttribut, ".", "")
ReplaceTxtAttribut = Replace(ReplaceTxtAttribut, ":", "")
ReplaceTxtAttribut = Replace(ReplaceTxtAttribut, "MOLEX", "")
ReplaceTxtAttribut = Replace(ReplaceTxtAttribut, "FCI", "")
ReplaceTxtAttribut = Replace(ReplaceTxtAttribut, "TYCO", "")
ReplaceTxtAttribut = Replace(ReplaceTxtAttribut, "/", "")
ReplaceTxtAttribut = Replace(ReplaceTxtAttribut, "XXXXX", "XXXXXX")
ReplaceTxtAttribut = Replace(ReplaceTxtAttribut, "XXXXXXXXXX", "XXXXXX")
ReplaceTxtAttribut = Replace(ReplaceTxtAttribut, "FILSENCOUPESNETTE", "XXXXXX")
ReplaceTxtAttribut = Replace(ReplaceTxtAttribut, "FILENCOUPENET", "XXXXXX")
ReplaceTxtAttribut = Replace(ReplaceTxtAttribut, "FILSCOUPENETTE", "XXXXXX")

'
If Left(ReplaceTxtAttribut & " ", 1) = "0" Then
ReplaceTxtAttribut = Mid(ReplaceTxtAttribut, 2, Len(ReplaceTxtAttribut) - 1)
End If
ReplaceTxtAttribut = Trim("" & UCase(ReplaceTxtAttribut))
End Function
