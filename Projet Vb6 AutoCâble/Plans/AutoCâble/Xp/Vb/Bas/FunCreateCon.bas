Attribute VB_Name = "FunCreateCon"



Sub NewConnecteurCardre()
'For i = 0 To ThisDrawing.ModelSpace.Count - 1
'Debug.Print ThisDrawing.ModelSpace(i).ObjectName
'
'Next

Dim CadreX As Double
Dim CadreY As Double
Dim MyFBloc As New FunCreateObjet
Dim a(0 To 2) As Double
Dim MyCercle
Dim MyPos
Dim XY1(0 To 2) As Double
Dim XY2(0 To 2) As Double
Dim NbL As Double
Dim NbC As Double
Dim MyType As String
MyType = "A1"
MyType = "A10"
MyType = "1A"
MyType = "10A"
MyType = "Z1"
MyType = "Z10"
MyType = "1Z"
MyType = "10Z"
MyType = "1"
MyType = "10"

NbL = 10
NbC = 10
'MyFBloc.SelectCopy ThisDrawing, ThisDrawing, 100, 10
a(0) = 10
a(1) = 10
a(2) = 1
 CadreX = 13
 CadreY = 20
'Set MyCercle = MyFBloc.CrateCercle(a, 10, acByBlock)
'Set MyCercle = MyFBloc.CrateCadre(MyFBloc.DefinirEspace(20, 3, 10, 36, 0))
MyPos = MyFBloc.DefinirEspace(CadreY + 8, CadreX + 8, NbL, NbC, 36, 15.6908)
'a(0) = MyPos(1) + 10
'a(1) = MyPos(3) - 10
a(0) = MyPos(0) + 5
a(1) = MyPos(3) - 20
a(2) = 1
Cv = a(0)
Lv = a(1)
Dim i As Long
For L = 1 To NbL
DoEvents
    a(0) = Cv
    For C = 1 To NbC
    i = i + 1
    
'Ligne Colone
MyType = "A"


Select Case UCase(MyType)

'        1,2,3...
        Case "1"
                CreatAlveole ThisDrawing, MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * L)), acByBlock, _
                "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, Val(i)
                
'        10,9,8..
        Case "10"
                CreatAlveole ThisDrawing, MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * L)), acByBlock, _
                "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, ((NbC * NbL) + 1) - i
                
'        1-1,1-2,1-3...
         Case "11"
            CreatAlveole ThisDrawing, MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * L)), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, Val(L) & "-" & Val(C)
                    
'        10-1,10-2,10-3...
        Case "101"
            CreatAlveole ThisDrawing, MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * L)), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, NbL + 1 - Val(L) & "-" & Val(C)
                    
'        1-10,1-9,1-8...
        Case "110"
            CreatAlveole ThisDrawing, MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * L)), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, Val(L) & "-" & NbC + 1 - Val(C)
                    
'        10-10,10-9,10-8...
        Case "1010"
            CreatAlveole ThisDrawing, MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * L)), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, NbL + 1 - Val(L) & "-" & NbC + 1 - Val(C)
                    
'        A,B,C...
        Case "A"
            CreatAlveole ThisDrawing, MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * L)), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, MyFBloc.CovertNumChar(i)
'        A1,A2,A3...
         Case "A1"
            CreatAlveole ThisDrawing, MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * L)), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, MyFBloc.CovertNumChar(Val(L)) & Val(C)
                    
'        A10,A9,A8...
        Case "A10"
            CreatAlveole ThisDrawing, MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * L)), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, MyFBloc.CovertNumChar(Val(L)) & NbL + 1 - Val(C)
       
'       Z,Y,X...
       Case "Z"
            CreatAlveole ThisDrawing, MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * L)), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, MyFBloc.CovertNumChar((NbL * NbC) + 1 - Val(i))
'        A-A,A-B,A-C...
        Case "AA"
            CreatAlveole ThisDrawing, MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * L)), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, MyFBloc.CovertNumChar(Val(L)) & "-" & MyFBloc.CovertNumChar(Val(C))
                    
'        A-Z,A-Y,A-X...
        Case "AZ"
            CreatAlveole ThisDrawing, MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * L)), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, MyFBloc.CovertNumChar(Val(L)) & "-" & MyFBloc.CovertNumChar(NbC + 1 - Val(C))
                    
'        Z-A,Z-B,Z-C...
        Case "ZA"
            CreatAlveole ThisDrawing, MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * L)), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, MyFBloc.CovertNumChar(Val(L)) & "-" & MyFBloc.CovertNumChar(Val(C))
                    
'        Z-Z,Z-Y,Z-X...
        Case "ZZ"
            CreatAlveole ThisDrawing, MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * L)), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, MyFBloc.CovertNumChar(NbL + 1 - Val(L)) & "-" & MyFBloc.CovertNumChar(NbC + 1 - Val(C))
       
'        1A,1B,1C...
        Case "1A"
             CreatAlveole ThisDrawing, MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * L)), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, Val(L) & MyFBloc.CovertNumChar(Val(C))
        
'        10A,10B,10C...
        Case "10A"
                   CreatAlveole ThisDrawing, MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * L)), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, NbL + 1 - Val(L) & MyFBloc.CovertNumChar(Val(C))

'        Z1,Z2,Z3...
        Case "Z1"
            CreatAlveole ThisDrawing, MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * L)), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, MyFBloc.CovertNumChar(NbL + 1 - Val(L)) & Val(C)
                    
'        Z10,Z9,Z8...
        Case "Z10"
            CreatAlveole ThisDrawing, MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * L)), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, MyFBloc.CovertNumChar(NbL + 1 - Val(L)) & NbC + 1 - Val(C)
                    
'        1Z,1Y,1Z...
        Case "1Z"
             CreatAlveole ThisDrawing, MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * L)), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, Val(L) & MyFBloc.CovertNumChar(NbC + 1 - Val(C))
                    
'        10Z,10Y,10X...
        Case "10Z"
             CreatAlveole ThisDrawing, MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * L)), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, NbL + 1 - Val(L) & MyFBloc.CovertNumChar(NbC + 1 - Val(C))
       

       
        Case Else

End Select



       
        a(0) = a(0) + CadreX + 4
    Next
a(1) = a(1) - CadreY - 4
Next
MyPos(0) = 0
MyPos(1) = 0
MyPos(2) = MyPos(2) + 10
MyPos(3) = MyPos(3) + 10 + 15.6908
XY1(0) = MyPos(0)
XY1(1) = MyPos(3) - 10
XY2(0) = MyPos(2)
XY2(1) = MyPos(3) - 10
''MyFBloc.CreateLigne ThisDrawing, XY1, XY2, acByBlock
'XY1(0) = MyPos(0)
'XY1(1) = XY1(1) - 10
'XY2(0) = MyPos(2)
'XY2(1) = XY2(1) - 10
'MyFBloc.CreateLigne ThisDrawing, XY1, XY2, acByBlock
'MyFBloc.CrateCadre ThisDrawing, MyPos, acByBlock
XY1(0) = XY1(0) + (XY2(0) / 2)
XY1(1) = XY1(1) + 5
MyFBloc.CreateAttribue ThisDrawing, "Ref Connecteur Client :", "RefConnecteurCli", "123454678", XY1, acMagenta, acAlignmentMiddleCenter
XY1(1) = XY1(1) + 10
MyFBloc.CreateAttribue ThisDrawing, "DESIGNATION :", "DESIGNATION", "", XY1, acMagenta, acAlignmentMiddleCenter

End Sub
Sub NewConnecteurCercle()
'For i = 0 To ThisDrawing.ModelSpace.Count - 1
'Debug.Print ThisDrawing.ModelSpace(i).ObjectName
'
'Next
Dim MyFBloc As New FunCreateObjet
Dim a(0 To 2) As Double
Dim XY1(0 To 2) As Double
Dim XY2(0 To 2) As Double
Dim MyCercle
Dim MyPos
a(0) = 10
a(1) = 10
a(2) = 1
'Set MyCercle = MyFBloc.CrateCercle(a, 10, acByBlock)
'Set MyCercle = MyFBloc.CrateCadre(MyFBloc.DefinirEspace(20, 3, 10, 36, 0))
MyPos = MyFBloc.DefinirEspace(10, 20, 3, 10, 36, 15.6908)
'a(0) = MyPos(1) + 10
'a(1) = MyPos(3) - 10
a(0) = MyPos(1) + 5
a(1) = MyPos(3) - 20
a(2) = 1
Cv = a(0)
Lv = a(1)
For L = 1 To 3
    a(0) = Cv
    For C = 1 To 10
        Set MyCercle = MyFBloc.CrateCercle(ThisDrawing, a, 16.08 / 2, acByBlock)
'        MyFBloc.CrateCadre (MyFBloc.DefinirEspace(1, 16.08, 13.02, a(0), a(1))), acByBlock
'        XY1 = a
'       XY1(1) = XY1(1) + (16.08 * 2 / 3)
'         XY2 = a
'         XY2(1) = XY2(1) + (16.08 * 2 / 3)
'         XY2(0) = XY2(0) + 13.02
'        MyFBloc.CreateLigne XY1, XY2, acByBlock
        a(0) = a(0) + 20
    Next
a(1) = a(1) - 20
Next
MyPos(0) = 0
MyPos(1) = 0
MyPos(2) = MyPos(2) + 10
MyPos(3) = MyPos(3) + 10 + 15.6908
XY1(0) = MyPos(0)
XY1(1) = MyPos(3) - 10
XY2(0) = MyPos(2)
XY2(1) = MyPos(3) - 10
'MyFBloc.CreateLigne ThisDrawing, XY1, XY2, acByBlock
XY1(0) = MyPos(0)
XY1(1) = XY1(1) - 10
XY2(0) = MyPos(2)
XY2(1) = XY2(1) - 10
'MyFBloc.CreateLigne ThisDrawing, XY1, XY2, acByBlock
'MyFBloc.CrateCadre ThisDrawing, MyPos, acByBlock
XY1(0) = XY1(0) + (XY2(0) / 2)
XY1(1) = XY1(1) + 5
MyFBloc.CreateAttribue ThisDrawing, "Ref Connecteur Client :", "RefConnecteurCli", "123454678", XY1, acMagenta, acAlignmentMiddleCenter
XY1(1) = XY1(1) + 10
MyFBloc.CreateAttribue ThisDrawing, "DESIGNATION :", "DESIGNATION", "", XY1, acMagenta, acAlignmentMiddleCenter

End Sub
Sub CreatAlveole(MyDocumment, InsertPointCadre, Couleur As Long, Prompt As String, txt As String, Defaulttxt As String, CadreY As Double, CadreX As Double, MyFBloc As FunCreateObjet, Index As String)
Dim XY1(0 To 2) As Double
Dim XY2(0 To 2) As Double
Dim InsertCadre(0 To 4) As Double
InsertCadre(0) = InsertPointCadre(0)
InsertCadre(1) = InsertPointCadre(1)
InsertCadre(2) = InsertPointCadre(2)
InsertCadre(3) = InsertPointCadre(3)
InsertCadre(1) = InsertCadre(3) - (CadreY * 1 / 4)
MyFBloc.CrateCadre MyDocumment, InsertCadre, acByBlock
 
 XY1(0) = InsertCadre(0) + (CadreX / 2)
       XY1(1) = InsertCadre(1) + ((InsertCadre(3) - InsertCadre(1)) * 3 / 2)
       
       MyFBloc.CreateEtiquette MyDocumment, "" & Index, XY1, acBlockReference, acAlignmentMiddleCenter
       
       XY1(0) = InsertCadre(0) + (CadreX / 2)
       XY1(1) = InsertCadre(1) + ((InsertCadre(3) - InsertCadre(1)) * 1 / 2)
       MyFBloc.CreateAttribue MyDocumment, "LIASON " & Index & " :", "LIAI" & Index, "", XY1, acMagenta, acAlignmentMiddleCenter
InsertCadre(3) = InsertCadre(1)
InsertCadre(1) = InsertCadre(3) - (CadreY * 3 / 4)
MyFBloc.CrateCadre MyDocumment, InsertCadre, acByBlock
 XY1(0) = InsertCadre(0) + (CadreX / 2)
       XY1(1) = InsertCadre(1) + ((InsertCadre(3) - InsertCadre(1)) * 1 / 4)
       MyFBloc.CreateAttribue MyDocumment, "FILS " & Index & " :", "FILS" & Index, "", XY1, acMagenta, acAlignmentMiddleCenter
       XY1(0) = InsertCadre(0) + (CadreX / 2)
       XY1(1) = InsertCadre(1) + ((InsertCadre(3) - InsertCadre(1)) * 3 / 4)
       MyFBloc.CreateAttribue MyDocumment, "MARIAGE " & Index & " :", "MAR" & Index, "", XY1, acMagenta, acAlignmentMiddleCenter
       
       ThisDrawing.Application.ZoomAll
       
'MyFBloc.CrateCadre MyDocumment, InsertPointCadre, acMagenta
      
'        XY1(1) = InsertPointCadre(3)
'        XY2(0) = InsertPointCadre(1)
'        XY2(1) = InsertPointCadre(3)
'
'       XY1(1) = (XY1(1) + CadreY) * 3 / 4
'
'         XY2(1) = (XY2(1) + CadreY) * 3 / 4
'         XY2(0) = XY2(0) + CadreX
'        MyFBloc.CreateLigne ThisDrawing, XY1, XY2, acByBlock
'
'        XY1(0) = InsertPointCadre(1)
'        XY1(1) = InsertPointCadre(3)
'        XY2(0) = InsertPointCadre(1)
'        XY2(1) = InsertPointCadre(3)
'         XY1(1) = (XY1(1) + CadreY) * 5 / 4
'
'         XY2(1) = (XY2(1) + CadreY) * 5 / 4
'         XY2(0) = XY2(0) + CadreX
'        MyFBloc.CreateLigne ThisDrawing, XY1, XY2, acByBlock
End Sub



