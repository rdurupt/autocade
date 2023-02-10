Attribute VB_Name = "AjoutAtrib"

Global MyLayer
Global MyColor

Public Sub AddAtrib()
Dim Rep
Dim splitTextString
Dim SplitRep
'RefConnecteurCli
If Dir(App.Path & "\DossierAplication\ConnecteurAtributs\ConnecteurCreatAttributs\Test.Ok") <> "" Then
    MsgBox "La Macro d'ajout Attributs  connecteur est déjà en cour d'exécution", vbInformation
    Exit Sub
End If
Dim Fso As New FileSystemObject
Fso.CreateTextFile App.Path & "\DossierAplication\ConnecteurAtributs\ConnecteurCreatAttributs\Test.Ok"
Rep = Dir(App.Path & "\DossierAplication\ConnecteurAtributs\ConnecteurCreatAttributs\*.DWG")
If Rep = "" Then Exit Sub

Dim I As Long
Dim aa
Dim BB
Dim Trouve As Boolean
Dim MyBloc As String
'AutoApp.Visible = True
While Rep <> ""


a = Split(Rep & ".", ".")
a = Split(a(0) & "§", "§")
MyBloc = a(0)
OpenFichier App.Path & "\DossierAplication\ConnecteurAtributs\ConnecteurCreatAttributs\" & Rep

Trouve = False
On Error GoTo Fin
For I = 0 To DocAutoCad.ModelSpace.Count - 1
Set aa = DocAutoCad.ModelSpace(I)
 BB = aa.ObjectName
 If BB = "AcDbAttributeDefinition" Then
     If Trim(UCase("" & aa.TagString)) = "DESIGNATION" Then
        MyLayer = aa.Layer
        MyColor = aa.Color
        Exit For
     End If
     
 End If
Next
For I = 0 To DocAutoCad.ModelSpace.Count - 1
    Set aa = DocAutoCad.ModelSpace(I)
    BB = aa.ObjectName
    
    If BB = "AcDbText" Then
    Debug.Print aa.TextString
   
    splitTextString = Split(MyBloc & "@@@", "@@@")
    If InStr(1, ReplaceTxtAttribut("" & aa.TextString), "" & splitTextString(0)) <> 0 Or InStr(1, ReplaceTxtAttribut("" & aa.TextString), "XXXXXX") <> 0 Or InStr(1, ReplaceTxtAttribut("" & aa.TextString), "xxxxxxxxxx") <> 0 Or InStr(1, ReplaceTxtAttribut("" & aa.TextString), ReplaceTxtAttribut("" & "ATTENTE REF")) <> 0 Or InStr(1, ReplaceTxtAttribut("" & aa.TextString), ReplaceTxtAttribut("" & "EN ATT")) <> 0 Or InStr(1, ReplaceTxtAttribut("" & aa.TextString), ReplaceTxtAttribut("" & MyBloc)) <> 0 Or InStr(1, UCase("" & aa.TextString), "REFERENCE") <> 0 Then
    Trouve = True
      If Trim("" & splitTextString(1)) <> "" Then
            Rep = Replace(Rep, "" & splitTextString(0) & "@@@", "")
        Else
            Rep = Replace(Rep, "@@@" & splitTextString(1), "")
        End If
        AddAtribut aa, "" & Replace(Rep, ".dwg", "")
        aa.Delete
        Set aa = Nothing
        'DocAutoCad.Application.ZoomAll
       
        SaveAs App.Path & "\DossierAplication\ConnecteurAtributs\SaveConnecteurCreatAttributs\" & Rep
         If Trim("" & splitTextString(1)) <> "" Then
                Rep = splitTextString(0) & "@@@" & Rep
           End If
        Fso.DeleteFile App.Path & "\DossierAplication\ConnecteurAtributs\ConnecteurCreatAttributs\" & Rep
       Exit For
      
    End If
    

    End If
     
    If BB = "AcDbMText" Then
    Debug.Print aa.TextString
   splitTextString = Split(MyBloc & "@@@", "@@@")
 If InStr(1, ReplaceTxtAttribut("" & aa.TextString), "" & splitTextString(0)) <> 0 Or InStr(1, ReplaceTxtAttribut("" & aa.TextString), "XXXXXX") <> 0 Or InStr(1, ReplaceTxtAttribut("" & aa.TextString), "xxxxxxxxxx") <> 0 Or InStr(1, ReplaceTxtAttribut("" & aa.TextString), ReplaceTxtAttribut("" & "ATTENTE REF")) <> 0 Or InStr(1, ReplaceTxtAttribut("" & aa.TextString), ReplaceTxtAttribut("" & "EN ATT")) <> 0 Or InStr(1, ReplaceTxtAttribut("" & aa.TextString), ReplaceTxtAttribut("" & MyBloc)) <> 0 Or InStr(1, UCase("" & aa.TextString), "REFERENCE") <> 0 Then
 Trouve = True
            If Trim("" & splitTextString(1)) <> "" Then
            Rep = Replace(Rep, "" & splitTextString(0) & "@@@", "")
        Else
            Rep = Replace(Rep, "@@@" & splitTextString(1), "")
        End If
        AddAtribut aa, "" & Replace(Rep, ".dwg", "")
        aa.Delete
        Set aa = Nothing
        'DocAutoCad.Application.ZoomAll
        
            If Fso.FileExists(App.Path & "\DossierAplication\ConnecteurAtributs\SaveConnecteurCreatAttributs\" & Rep) = True Then
                Fso.DeleteFile App.Path & "\DossierAplication\ConnecteurAtributs\SaveConnecteurCreatAttributs\" & Rep
            End If
           SaveAs App.Path & "\DossierAplication\ConnecteurAtributs\SaveConnecteurCreatAttributs\" & Rep
           If Trim("" & splitTextString(1)) <> "" Then
                Rep = splitTextString(0) & "@@@" & Rep
           End If
            Fso.DeleteFile App.Path & "\DossierAplication\ConnecteurAtributs\ConnecteurCreatAttributs\" & Rep
'       SaveAs Rep
       Exit For
      
    End If
    

    End If
    If BB = "AcDbAttributeDefinition" Then
    Debug.Print aa.TextString
       splitTextString = Split(MyBloc & "@@@", "@@@")
 If InStr(1, "" & aa.TagString, "RefConnecteurFour") <> 0 Or InStr(1, ReplaceTxtAttribut("" & aa.TextString), "" & splitTextString(0)) <> 0 Or InStr(1, ReplaceTxtAttribut(UCase("" & aa.TagString)), ReplaceTxtAttribut("" & MyBloc)) <> 0 Or InStr(1, UCase("" & aa.TagString), "REFERENCE") <> 0 Then
            aa.TagString = "RefConnecteurCli"
            aa.PromptString = UCase("Ref Connecteur Client :")
              If Trim("" & splitTextString(1)) <> "" Then
            Rep = Replace(Rep, "" & splitTextString(0) & "@@@", "")
        Else
            Rep = Replace(Rep, "@@@" & splitTextString(1), "")
        End If
            
             aa.TextString = "" & "" & Replace(Rep, ".dwg", "")
            aa.Layer = MyLayer
            aa.Color = MyColor

             Trouve = True
            If Fso.FileExists(App.Path & "\DossierAplication\ConnecteurAtributs\SaveConnecteurCreatAttributs\" & Rep) = True Then
                Fso.DeleteFile App.Path & "\DossierAplication\ConnecteurAtributs\SaveConnecteurCreatAttributs\" & Rep
            End If
           SaveAs App.Path & "\DossierAplication\ConnecteurAtributs\SaveConnecteurCreatAttributs\" & Rep
            If Trim("" & splitTextString(1)) <> "" Then
                Rep = splitTextString(0) & "@@@" & Rep
           End If
           Fso.DeleteFile App.Path & "\DossierAplication\ConnecteurAtributs\ConnecteurCreatAttributs\" & Rep
           Exit For
        End If
    End If
Next

If Trouve = False Then DocAutoCad.Close , False
Fin:

Err.Clear
Rep = ""

 Rep = Dir
 DoEvents
 Wend
 Rep = Dir(App.Path & "\DossierAplication\ConnecteurAtributs\SaveConnecteurCreatAttributs\*.DWG")
While Rep <> ""
    If Fso.FileExists(App.Path & "\DossierAplication\ConnecteurAtributs\ConnecteurCreatAttributs\" & Rep) = True Then
        Fso.DeleteFile App.Path & "\DossierAplication\ConnecteurAtributs\ConnecteurCreatAttributs\" & Rep
        DoEvents
    End If
    Rep = Dir
Wend

'
 If Fso.FileExists(App.Path & "\DossierAplication\ConnecteurAtributs\ConnecteurCreatAttributs\Test.Ok") = True Then Fso.DeleteFile App.Path & "\DossierAplication\ConnecteurAtributs\ConnecteurCreatAttributs\Test.Ok"
 Set Fso = Nothing
 'AutoApp.Visible = False
End Sub
Sub AddAtribut(Etiqqette, MyBloc As String)
Dim InsertPoint(0 To 2) As Double
Dim EE
Dim MyWheet

Set EE = DocAutoCad.ModelSpace.AddAttribute(Etiqqette.Height, acAttributeModeNormal, "Ref Connecteur Client :", Etiqqette.InsertionPoint, "RefConnecteurCli", MyBloc)
'EE.Color = "Magenta"
EE.Rotation = Etiqqette.Rotation
EE.Layer = MyLayer
EE.Color = MyColor

EE.Alignment = acAlignmentMiddleCenter
'EE = Etiqqette.InsertionPoint
EE.TextAlignmentPoint = Etiqqette.InsertionPoint
'EE.TextAlignmentPoint(1) = aa(1)
End Sub
Function ReplaceTxtAttribut(Txt As String)
ReplaceTxtAttribut = Txt
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
