VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
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

Public Function DefinirEspace(Y As Double, X As Double, NbLigne As Double, NbColonne As Double, OfsetX As Double, OfsetY As Double)
Dim MyCadre(3) As Double
MyCadre(0) = OfsetX
MyCadre(1) = OfsetY
MyCadre(2) = OfsetX + (X * NbColonne)
MyCadre(3) = OfsetY + (Y * NbLigne)


DefinirEspace = MyCadre
End Function
Public Function CreateLigne(MyDocumment, XY1, XY2, Couleur As Long)
    Set CreateLigne = MyDocumment.ModelSpace.AddLine(XY1, XY2)
    CreateLigne.Color = Couleur
End Function
Public Function CreateEtiquette(MyDocumment, txt As String, XY1, Couleur As Long, Alignment)
    Set CreateEtiquette = MyDocumment.ModelSpace.AddText(txt, XY1, 3)
    CreateEtiquette.Alignment = acAlignmentMiddleCenter
    CreateEtiquette.TextAlignmentPoint = XY1
    CreateEtiquette.Color = Couleur
End Function
Public Function CreateEtiquetteM(MyDocumment, txt As String, XY1, Couleur As Long, AttachmentPoint)
    Set CreateEtiquetteM = MyDocumment.ModelSpace.AddMText(XY1, 3, txt)
    CreateEtiquetteM.AttachmentPoint = AttachmentPoint
'acAttachmentPointMiddleCenter
'    CreateEtiquetteM.TextAlignmentPoint = XY1
    CreateEtiquetteM.Color = Couleur
End Function


Public Function CreateAttribue(MyDocumment, Prompt As String, txt As String, Defaulttxt As String, XY1, Couleur As Long, Alignment)
    Set CreateAttribue = MyDocumment.ModelSpace.AddAttribute(3, acAttributeModeNormal, Prompt, XY1, txt, Defaulttxt)
    CreateAttribue.Alignment = Alignment
    CreateAttribue.TextAlignmentPoint = XY1
    CreateAttribue.Color = Couleur
End Function

Public Sub SelectCopy(Documen1, Document2, X As Double, Y As Double)
Dim InserCadre(4)

 Dim InseterPoint
  For i = 0 To Documen1.ModelSpace.Count - 1
' CreateLigne ThisDrawing , ThisDrawing.ModelSpace(i).Select
  Set aa = ThisDrawing.ModelSpace(i)
 Debug.Print aa.ObjectName
Select Case aa.ObjectName
        Case "AcDbAttributeDefinition"
            InseterPoint = aa.InsertionPoint
            InseterPoint(0) = InseterPoint(0) + X
            InseterPoint(1) = InseterPoint(1) + Y
            CreateAttribue Document2, aa.PromptString, aa.TagString, aa.TextString, InseterPoint, aa.Color, aa.Alignment
          Case "AcDbPolyline"
         
          InseterPoint = aa.Coordinates
          InserCadre(0) = InseterPoint(4) + X
           InserCadre(1) = InseterPoint(5) + Y
            InserCadre(2) = InseterPoint(0) + X
             InserCadre(3) = InseterPoint(1) + Y
             CrateCadre Document2, InserCadre, aa.Color
          Case "AcDbCircle"
            InseterPoint = aa.Center
            InseterPoint(0) = InseterPoint(0) + X
            InseterPoint(1) = InseterPoint(1) + Y
            CrateCercle Document2, InseterPoint, aa.Radius, aa.Color
          Case "AcDbMText"
           InseterPoint = aa.InsertionPoint
            InseterPoint(0) = InseterPoint(0) + X
            InseterPoint(1) = InseterPoint(1) + Y
          CreateEtiquetteM Document2, aa.TextString, InseterPoint, aa.Color, aa.AttachmentPoint
        Case "AcDbText"
           InseterPoint = aa.InsertionPoint
            InseterPoint(0) = InseterPoint(0) + X
            InseterPoint(1) = InseterPoint(1) + Y
          CreateEtiquette Document2, aa.TextString, InseterPoint, aa.Color, aa.Alignment
End Select
'    ThisDrawing.Activate
    
  Next
 
End Sub
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
