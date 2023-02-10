Attribute VB_Name = "Module1"
Public MyAutocad As Object
Public MyDocDefault As AutoCAD.AcadDocument

Function NewRealai(OfsetX As Double, OfsetY As Double, MyType As String, Cercle As Boolean)
    Dim CadreX As Double
Dim CadreY As Double
Dim MyFBloc As New FunCreateObjet
Dim a(0 To 2) As Double
 
Dim MyPos
Dim XY1(0 To 2) As Double
Dim XY2(0 To 2) As Double
Dim Nbl As Double
Dim NBC As Double
Dim InsertCadre



Nbl = 10
NBC = 10

'MyFBloc.SelectCopy MyAutocad.Documents(0), MyAutocad.Documents(0), 100, 10
a(0) = 10
a(1) = 10
a(2) = 1
 CadreX = 67
 CadreY = 67
'Set MyCercle = MyFBloc.CrateCercle(a, 10, acByBlock)
'Set MyCercle = MyFBloc.CrateCadre(MyFBloc.DefinirEspace(20, 3, 10, 36, 0))
MyPos = MyFBloc.DefinirEspace(CadreY + 8, CadreX + 8, Nbl, NBC, OfsetX, OfsetY, Cercle)
'a(0) = MyPos(1) + 10
'a(1) = MyPos(3) - 10
a(0) = MyPos(0) + 5
a(1) = MyPos(3) - 20
a(2) = 1
Cv = a(0)
Lv = a(1)
Dim i As Long
For l = 1 To Nbl
DoEvents
    a(0) = Cv
    For c = 1 To NBC
    i = i + 1
    Set aa = MyAutocad.Documents.Open(App.Path & "\Relais\Relai1.dwg")
'  Set aa = MyAutocad.Documents(App.Path & "\Relais\Relai1.dwg").SendCommand("_COPYCLIP ")
 
    InsertCadre = MyFBloc.SelectCopy(aa, MyDocDefault, a(0), a(1), i)
    a(0) = InsertCadre(2) + 5
    aa.Close False
Next
   a(0) = MyPos(0) + 5
   Debug.Print InsertCadre(3) - InsertCadre(1)
   a(1) = a(1) - (InsertCadre(3) - InsertCadre(1))
Next

NewRealai = MyPos
End Function

Function NewConnecteur(OfsetX As Double, OfsetY As Double, MyType As String, Cercle As Boolean, Nbl As Double, NBC As Double)
'For i = 0 To MyAutocad.Documents(0).ModelSpace.Count - 1
'Debug.Print MyAutocad.Documents(0).ModelSpace(i).ObjectName
'
'Next

Dim CadreX As Double
Dim CadreY As Double
Dim MyFBloc As New FunCreateObjet
Dim a(0 To 2) As Double
 
Dim MyPos
Dim XY1(0 To 2) As Double
Dim XY2(0 To 2) As Double







'MyFBloc.SelectCopy MyAutocad.Documents(0), MyAutocad.Documents(0), 100, 10
a(0) = 10
a(1) = 10
a(2) = 1
 CadreX = 16.8
 CadreY = 21.6
'Set MyCercle = MyFBloc.CrateCercle(a, 10, acByBlock)
'Set MyCercle = MyFBloc.CrateCadre(MyFBloc.DefinirEspace(20, 3, 10, 36, 0))
MyPos = MyFBloc.DefinirEspace(CadreY + 8, CadreX + 8, Nbl, NBC, OfsetX, OfsetY, Cercle)
'a(0) = MyPos(1) + 10
'a(1) = MyPos(3) - 10
a(0) = MyPos(0) + 5
a(1) = MyPos(3) - 21.6
a(2) = 1
Cv = a(0)
Lv = a(1)
Dim i As Long
For l = 1 To Nbl
DoEvents
    a(0) = Cv
    For c = 1 To NBC
    i = i + 1
    
'Ligne Colone



Select Case UCase(MyType)

'        1,2,3...
        Case "1"
                CreatAlveole MyAutocad.Documents(0), MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * l), Cercle), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, Val(i), Cercle
                
'        10,9,8..
        Case "10"
                CreatAlveole MyAutocad.Documents(0), MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * l), Cercle), acByBlock, _
                "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, ((NBC * Nbl) + 1) - i, Cercle
                
'        1-1,1-2,1-3...
         Case "11"
            CreatAlveole MyAutocad.Documents(0), MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * l), Cercle), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, Val(l) & "-" & Val(c), Cercle
                    
'        10-1,10-2,10-3...
        Case "101"
            CreatAlveole MyAutocad.Documents(0), MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * l), Cercle), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, Nbl + 1 - Val(l) & "-" & Val(c), Cercle
                    
'        1-10,1-9,1-8...
        Case "110"
            CreatAlveole MyAutocad.Documents(0), MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * l), Cercle), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, Val(l) & "-" & NBC + 1 - Val(c), Cercle
                    
'        10-10,10-9,10-8...
        Case "1010"
            CreatAlveole MyAutocad.Documents(0), MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * l), Cercle), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, Nbl + 1 - Val(l) & "-" & NBC + 1 - Val(c), Cercle
                    
'        A,B,C...
        Case "A"
            CreatAlveole MyAutocad.Documents(0), MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * l), Cercle), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, MyFBloc.CovertNumChar(i), Cercle
'        A1,A2,A3...
         Case "A1"
            CreatAlveole MyAutocad.Documents(0), MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * l), Cercle), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, MyFBloc.CovertNumChar(Val(l)) & Val(c), Cercle
                    
'        A10,A9,A8...
        Case "A10"
            CreatAlveole MyAutocad.Documents(0), MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * l), Cercle), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, MyFBloc.CovertNumChar(Val(l)) & Nbl + 1 - Val(c), Cercle
       
'       Z,Y,X...
       Case "Z"
            CreatAlveole MyAutocad.Documents(0), MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * l), Cercle), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, MyFBloc.CovertNumChar((Nbl * NBC) + 1 - Val(i)), Cercle
'        A-A,A-B,A-C...
        Case "AA"
            CreatAlveole MyAutocad.Documents(0), MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * l), Cercle), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, MyFBloc.CovertNumChar(Val(l)) & "-" & MyFBloc.CovertNumChar(Val(c)), Cercle
                    
'        A-Z,A-Y,A-X...
        Case "AZ"
            CreatAlveole MyAutocad.Documents(0), MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * l), Cercle), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, MyFBloc.CovertNumChar(Val(l)) & "-" & MyFBloc.CovertNumChar(NBC + 1 - Val(c)), Cercle
                    
'        Z-A,Z-B,Z-C...
        Case "ZA"
            CreatAlveole MyAutocad.Documents(0), MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * l), Cercle), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, MyFBloc.CovertNumChar(Val(l)) & "-" & MyFBloc.CovertNumChar(Val(c)), Cercle
                    
'        Z-Z,Z-Y,Z-X...
        Case "ZZ"
            CreatAlveole MyAutocad.Documents(0), MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * l), Cercle), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, MyFBloc.CovertNumChar(Nbl + 1 - Val(l)) & "-" & MyFBloc.CovertNumChar(NBC + 1 - Val(c)), Cercle
       
'        1A,1B,1C...
        Case "1A"
             CreatAlveole MyAutocad.Documents(0), MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * l), Cercle), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, Val(l) & MyFBloc.CovertNumChar(Val(c)), Cercle
        
'        10A,10B,10C...
        Case "10A"
                   CreatAlveole MyAutocad.Documents(0), MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * l), Cercle), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, Nbl + 1 - Val(l) & MyFBloc.CovertNumChar(Val(c)), Cercle

'        Z1,Z2,Z3...
        Case "Z1"
            CreatAlveole MyAutocad.Documents(0), MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * l), Cercle), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, MyFBloc.CovertNumChar(Nbl + 1 - Val(l)) & Val(c), Cercle
                    
'        Z10,Z9,Z8...
        Case "Z10"
            CreatAlveole MyAutocad.Documents(0), MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * l), Cercle), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, MyFBloc.CovertNumChar(Nbl + 1 - Val(l)) & NBC + 1 - Val(c), Cercle
                    
'        1Z,1Y,1Z...
        Case "1Z"
             CreatAlveole MyAutocad.Documents(0), MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * l), Cercle), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, Val(l) & MyFBloc.CovertNumChar(NBC + 1 - Val(c)), Cercle
                    
'        10Z,10Y,10X...
        Case "10Z"
             CreatAlveole MyAutocad.Documents(0), MyFBloc.DefinirEspace(1, 1, CadreY, CadreX, a(0), a(1) - (4 * l), Cercle), acByBlock, _
                    "TOTO :", "toto", "12345678", CadreY, CadreX, MyFBloc, Nbl + 1 - Val(l) & MyFBloc.CovertNumChar(NBC + 1 - Val(c)), Cercle
       

       
        Case Else

End Select


    If Cercle = False Then
        a(0) = a(0) + CadreX + 4
     Else
        a(0) = a(0) + CadreY + 4
     End If
    Next
a(1) = a(1) - CadreY - 4
Next
MyPos(0) = 0
MyPos(1) = 0
MyPos(2) = MyPos(2) + 10
MyPos(3) = MyPos(3) + 10 + 15.6908

MyFBloc.CreateLigne MyAutocad.Documents(0), XY1, XY2, acByBlock
XY1(0) = MyPos(0)
XY1(1) = XY1(1) - 10
XY2(0) = MyPos(2)
XY2(1) = XY2(1) - 10
MyFBloc.CreateLigne MyAutocad.Documents(0), XY1, XY2, acByBlock
MyFBloc.CrateCadre MyAutocad.Documents(0), MyPos, acByBlock
NewConnecteur = MyPos


End Function

Sub CreatAlveole(MyDocumment, InsertPointCadre, Couleur As Long, Prompt As String, txt As String, Defaulttxt As String, CadreY As Double, CadreX As Double, MyFBloc As FunCreateObjet, Index As String, Crecle As Boolean)
Dim XY1(0 To 2) As Double
Dim XY2(0 To 2) As Double
Dim InsertCadre(0 To 4) As Double
InsertCadre(0) = InsertPointCadre(0)
InsertCadre(1) = InsertPointCadre(1)
InsertCadre(2) = InsertPointCadre(2)
InsertCadre(3) = InsertPointCadre(3)
InsertCadre(1) = InsertCadre(3) - (CadreY * 1 / 4)
If Crecle = False Then
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
       MyFBloc.CreateAttribue MyDocumment, "MARIAGE " & Index & " :", "MAR" & Index, "", XY1, acMagenta, acAlignmentMiddleCenter
       
       XY1(1) = InsertCadre(1) + ((InsertCadre(3) - InsertCadre(1)) * 3 / 4)
       MyFBloc.CreateAttribue MyDocumment, "FIL " & Index & " :", "FIL" & Index, "", XY1, acMagenta, acAlignmentMiddleCenter
       XY1(0) = InsertCadre(0) + (CadreX / 2)
       
Else
    XY1(0) = InsertCadre(0) + (CadreY / 2)
    XY1(1) = InsertCadre(1) + (CadreY / 2)
    MyFBloc.CrateCercle MyDocumment, XY1, (CadreY / 2), Couleur
    XY1(0) = InsertCadre(0) + (CadreY * 1 / 6)
    XY1(1) = InsertCadre(1) + (CadreY * 4 / 6)
    XY2(0) = XY1(0) + (CadreY * 4 / 6)
    XY2(1) = XY1(1)
    MyFBloc.CreateLigne MyDocumment, XY1, XY2, Couleur
    XY1(0) = InsertCadre(0) + (CadreY / 2)
    XY1(1) = XY1(1) + 3
    MyFBloc.CreateAttribue MyDocumment, "LIASON " & Index & " :", "LIAI" & Index, "", XY1, acMagenta, acAlignmentMiddleCenter
    XY1(1) = XY1(1) - 6
    MyFBloc.CreateAttribue MyDocumment, "FIL " & Index & " :", "FIL" & Index, "", XY1, acMagenta, acAlignmentMiddleCenter
    XY1(1) = XY1(1) - 6
    MyFBloc.CreateAttribue MyDocumment, "MARIAGE " & Index & " :", "MAR" & Index, "", XY1, acMagenta, acAlignmentMiddleCenter
    XY1(1) = XY1(1) + CadreY
    MyFBloc.CreateEtiquette MyDocumment, "" & Index, XY1, acBlockReference, acAlignmentMiddleCenter
End If
       MyAutocad.Documents(0).Application.ZoomAll
       
'MyFBloc.CrateCadre MyDocumment, InsertPointCadre, acMagenta
      
'        XY1(1) = InsertPointCadre(3)
'        XY2(0) = InsertPointCadre(1)
'        XY2(1) = InsertPointCadre(3)
'
'       XY1(1) = (XY1(1) + CadreY) * 3 / 4
'
'         XY2(1) = (XY2(1) + CadreY) * 3 / 4
'         XY2(0) = XY2(0) + CadreX
'        MyFBloc.CreateLigne MyAutocad.Documents(0), XY1, XY2, acByBlock
'
'        XY1(0) = InsertPointCadre(1)
'        XY1(1) = InsertPointCadre(3)
'        XY2(0) = InsertPointCadre(1)
'        XY2(1) = InsertPointCadre(3)
'         XY1(1) = (XY1(1) + CadreY) * 5 / 4
'
'         XY2(1) = (XY2(1) + CadreY) * 5 / 4
'         XY2(0) = XY2(0) + CadreX
'        MyFBloc.CreateLigne MyAutocad.Documents(0), XY1, XY2, acByBlock
End Sub

Sub Demarage(MyType As String, Cercle As Boolean, Nbl As Double, NBC As Double)
Dim XY1(0 To 2) As Double
Dim XY2(0 To 2) As Double
Dim MyPos
'Dim MyType(10) As String
'Choose Select Objects.
'
'MyAutocad.Documents("Relai1.dwg").Senedim Command " _ai_selall Choix des objets en cours...terminé."
'
'   Application.Documents("Relai1.dwg").SendCommand "_COPYCLIP "
'   Application.Documents("ModelBloc.dwg").SendCommand "_pasteclip Spécifiez le point d'insertion: 100,100 "
ReDim MyPos(3)
'MyType(0) = "A1"
'MyType(1) = "A10"
'MyType(2) = "1A"
'MyType(3) = "10A"
'MyType(4) = "Z1"
'MyType(5) = "Z10"
'MyType(6) = "1Z"
'MyType(7) = "10Z"
'MyType(8) = "1"
'MyType(9) = "10"
'MyType(10) = "A"
'MyPos = NewRealai(36, 15.6908, "a", False)
'MyPos = NewRealai(36, Val(MyPos(3)), "10", False)
MyPos = NewConnecteur(36, Val(MyPos(3)), MyType, Cercle, Nbl, NBC)
'MyPos = NewConnecteur(36, Val(MyPos(3)), "z10",  True)
XY1(0) = MyPos(0)
XY1(1) = MyPos(3) - 10
XY2(0) = MyPos(2)
XY2(1) = MyPos(3) - 10
XY1(0) = XY1(0) + (XY2(0) / 2)
XY1(1) = XY1(1) + 5
Dim MyFBloc As New FunCreateObjet
MyFBloc.CreateAttribue MyAutocad.Documents(0), "Ref Connecteur Client :", "RefConnecteurCli", "123454678", XY1, acMagenta, acAlignmentMiddleCenter
XY1(1) = XY1(1) + 10
MyFBloc.CreateAttribue MyAutocad.Documents(0), "DESIGNATION :", "DESIGNATION", "", XY1, acMagenta, acAlignmentMiddleCenter
End Sub
