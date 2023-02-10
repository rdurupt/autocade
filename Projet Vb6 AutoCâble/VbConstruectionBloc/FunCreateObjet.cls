VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FunCreateObjet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function CrateCercle(MyDocumment, InsertPoint, Rayon As Double, Couleur As Long)
Set CrateCercle = MyDocumment.ModelSpace.AddCircle(InsertPoint, Rayon)
CrateCercle.Color = Couleur
End Function
Public Function CrateCadre(MyDocumment, InsertPointCadre, Couleur As Long)
  Dim points(0 To 14) As Double
    
 
    points(0) = InsertPointCadre(0): points(1) = InsertPointCadre(1): points(2) = 0
    points(3) = InsertPointCadre(2): points(4) = InsertPointCadre(1): points(5) = 0
    points(6) = InsertPointCadre(2): points(7) = InsertPointCadre(3): points(8) = 0
    points(9) = InsertPointCadre(0): points(10) = InsertPointCadre(3): points(11) = 0
    points(12) = InsertPointCadre(0): points(13) = InsertPointCadre(1): points(14) = 0
    Set CrateCadre = MyDocumment.ModelSpace.AddPolyline(points)

CrateCadre.Color = Couleur
End Function

Public Function DefinirEspace(Y As Double, x As Double, NbLigne As Double, NbColonne As Double, OfsetX As Double, OfsetY As Double, Crecle As Boolean)
Dim MyCadre(3) As Double
MyCadre(0) = OfsetX
MyCadre(1) = OfsetY
If Crecle = False Then
    MyCadre(2) = OfsetX + (x * NbColonne)
Else
    MyCadre(2) = OfsetX + (Y * NbColonne)
End If
MyCadre(3) = OfsetY + (Y * NbLigne)


DefinirEspace = MyCadre
End Function
Public Function CreateLigne(MyDocumment, XY1, XY2, Couleur As Long)
    Set CreateLigne = MyDocumment.ModelSpace.AddLine(XY1, XY2)
    CreateLigne.Color = Couleur
    CreateLigne.Application.ZoomAll
    CreateLigne.Application.ZoomPrevious
'
End Function
Public Function CreateEtiquette(MyDocumment, txt As String, XY1, Couleur As Long, Alignment, Optional Rotation = 0)
    Set CreateEtiquette = MyDocumment.ModelSpace.AddText(txt, XY1, 3)
    CreateEtiquette.Alignment = Alignment
    If Alignment <> 0 Then CreateEtiquette.TextAlignmentPoint = XY1
    CreateEtiquette.Rotation = Rotation
    CreateEtiquette.Color = Couleur
    CreateEtiquette.Application.ZoomPrevious
    CreateEtiquette.Application.ZoomAll
End Function
Public Function CreateEtiquetteM(MyDocumment, txt As String, XY1, Couleur As Long, AttachmentPoint, Optional Rotation = 0)
    Set CreateEtiquetteM = MyDocumment.ModelSpace.AddMText(XY1, 3, txt)
    CreateEtiquetteM.AttachmentPoint = AttachmentPoint
    CreateEtiquetteM.Rotation = Rotation
'acAttachmentPointMiddleCenter
    CreateEtiquetteM.TextAlignmentPoint = XY1
    CreateEtiquetteM.Color = Couleur
End Function


Public Function CreateAttribue(MyDocumment, Prompt As String, txt As String, Defaulttxt As String, XY1, Couleur As Long, Alignment, Optional Rotation = 0)
    Set CreateAttribue = MyDocumment.ModelSpace.AddAttribute(3, acAttributeModeNormal, Prompt, XY1, txt, Defaulttxt)
    CreateAttribue.Alignment = Alignment
    CreateAttribue.Rotation = Rotation
    If Alignment <> 0 Then _
    CreateAttribue.TextAlignmentPoint = XY1
    CreateAttribue.Color = Couleur
End Function

Public Function SelectCopy(Documen1, Document2, x As Double, Y As Double, Index)
Dim IniOk As Boolean
Dim InserCadre(4)
Dim XY1(0 To 2) As Double
Dim XY2(0 To 2) As Double
Dim SaveXy1(0 To 2) As Double
Dim SaveXy2(0 To 2) As Double
 Dim StartPoint
 Dim EndPoint
 Dim InseterPoint
' MyDocDefault.Layers(1).LayerOn = False
 SaveXy1(0) = x: SaveXy1(1) = Y: SaveXy2(0) = x: SaveXy2(1) = Y

  For i = 0 To Documen1.ModelSpace.Count - 1
' CreateLigne MyAutocad.Documents(0) , MyAutocad.Documents(0).ModelSpace(i).Select
  Set aa = Documen1.ModelSpace(i)
 Debug.Print aa.ObjectName
' aa.Copy
 Document2.Activate
 
Select Case aa.ObjectName
        Case "AcDbLine"
            StartPoint = aa.StartPoint
             EndPoint = aa.EndPoint
            XY1(0) = x + StartPoint(0)
            XY1(1) = StartPoint(1) + Y
            
            XY2(0) = EndPoint(0) + x
            XY2(1) = EndPoint(1) + Y
            
'            XY1(0) = x: XY1(1) = Y
'            XY2(0) = EndPoint(1): XY2(1) = Y
            If IniOk = False Then
                IniOk = True
                SaveXy1(0) = XY1(0):  SaveXy1(1) = XY1(1)
                SaveXy2(0) = XY2(0): SaveXy2(1) = XY2(1)
            End If
            If SaveXy1(0) > XY1(0) Then SaveXy1(0) = XY1(0): If SaveXy1(1) > XY1(1) Then SaveXy1(1) = XY1(1)
            If SaveXy2(0) < XY2(0) Then SaveXy2(0) = XY2(0): If SaveXy2(1) < XY2(1) Then SaveXy2(1) = XY2(1)
            CreateLigne Document2, XY1, XY2, aa.Color
        Case "AcDbAttributeDefinition"
            InseterPoint = aa.InsertionPoint
            InseterPoint(0) = InseterPoint(0) + x
            InseterPoint(1) = InseterPoint(1) + Y
            CreateAttribue Document2, aa.PromptString, Replace(UCase(aa.TagString), UCase("[RNum]"), Index), aa.TextString, InseterPoint, acMagenta, acAlignmentMiddleLeft, aa.Rotation
          Case "AcDbPolyline"
         
          InseterPoint = aa.Coordinates
          InserCadre(0) = InseterPoint(4) + x
           InserCadre(1) = InseterPoint(5) + Y
            InserCadre(2) = InseterPoint(0) + x
             InserCadre(3) = InseterPoint(1) + Y
             CrateCadre Document2, InserCadre, aa.Color
             If IniOk = False Then
                IniOk = True
                SaveXy1(0) = XY1(0):  SaveXy1(1) = XY1(1)
                SaveXy2(0) = XY2(0): SaveXy2(1) = XY2(1)
            End If
            If SaveXy1(0) > XY1(0) Then SaveXy1(0) = XY1(0): If SaveXy1(1) > XY1(1) Then SaveXy1(1) = XY1(1)
            If SaveXy2(0) < XY2(0) Then SaveXy2(0) = XY2(0): If SaveXy2(1) < XY2(1) Then SaveXy2(1) = XY2(1)
          Case "AcDbCircle"
            InseterPoint = aa.Center
            InseterPoint(0) = InseterPoint(0) + x
            InseterPoint(1) = InseterPoint(1) + Y
            CrateCercle Document2, InseterPoint, aa.Radius, aa.Color
          Case "AcDbMText"
           InseterPoint = aa.InsertionPoint
            InseterPoint(0) = InseterPoint(0) + x
            InseterPoint(1) = InseterPoint(1) + Y
          CreateEtiquetteM Document2, Replace(aa.TextString, "[RNum]", Index), InseterPoint, aa.Color, aa.AttachmentPoint, aa.Rotation
        Case "AcDbText"
           InseterPoint = aa.InsertionPoint
            InseterPoint(0) = InseterPoint(0) + x
            InseterPoint(1) = InseterPoint(1) + Y
          CreateEtiquette Document2, Replace(aa.TextString, "[RNum]", Index), InseterPoint, aa.Color, aa.Alignment, aa.Rotation
End Select
'    MyAutocad.Documents(0).Activate
    
  Next
 InserCadre(0) = SaveXy1(0): InserCadre(1) = SaveXy1(1): InserCadre(2) = SaveXy2(0): InserCadre(3) = SaveXy2(1)
 SelectCopy = InserCadre
End Function
Public Function CovertNumChar(Num As Long) As String
Dim txt As String

txt = "A"
Dim NumCal As Long

For i = 0 To Num - 1
    NumCal = Len(txt)
    If i > 0 Then
        Mid(txt, NumCal, 1) = Chr(Asc(Mid(txt, NumCal, 1)) + 1)
    End If
    Debug.Print Asc(Mid(txt, NumCal, 1))
    If Asc(Mid(txt, NumCal, 1)) = 91 Then
        Mid(txt, NumCal, 1) = Chr(97)
    End If
    If Asc(Mid(txt, NumCal, 1)) = 123 Then
       
            Mid(txt, NumCal, 1) = "A"
            NumCal = NumCal - 1
    End If
    If NumCal = 0 Then
        txt = txt & "A"
    End If
Next
CovertNumChar = txt
End Function
