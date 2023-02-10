Attribute VB_Name = "ImporterXml"
Option Explicit
Type TypeConnecteur
    Connecteur  As String
    DESIGNATION As String
    Code_APP As String
    PRECO1 As String
End Type




 Public Sub ImporteCatiaV5(FichierTxt As String, Conn As Object, Connecteur As Object, Fil As Object, NumCollonne)
 Dim TableauConnecteur() As TypeConnecteur
 Dim FileNumber As Long
 Dim LigneTxt As String
 Dim File
 Dim MesConnecteurs
 Dim FileLength As Long
 Dim L As Long
 Dim I As Long
 Dim Save_i As Long
 Dim I_Connecteur As Long
 Dim Tableaufils() As String
 Dim SplitConnecteur
 Dim SplitFils
 Dim MyExcel As New EXCEL.Application
 Dim MyWorkbook As Workbook
 Dim MyRange As Range
 Dim ValPin
 MyExcel.Visible = True
 MyExcel.DisplayAlerts = False
 Set MyWorkbook = MyExcel.Workbooks.Open(Filename:=FichierTxt)
 MyWorkbook.Sheets(1).Rows("1:1").Delete Shift:=xlUp
 MyWorkbook.Sheets(1).Rows("1:1").Delete Shift:=xlUp
 MyWorkbook.Sheets(1).Rows("1:1").Delete Shift:=xlUp


    MyWorkbook.Sheets(1).Copy After:=Sheets(1)
MyWorkbook.Sheets(2).Select
 Set MyRange = MyWorkbook.Sheets(2).Range("A1").CurrentRegion
 MyRange.Delete Shift:=xlUp
  MyWorkbook.Sheets(2).Rows("1:1").Delete Shift:=xlUp
 MyWorkbook.Sheets(2).Rows("1:1").Delete Shift:=xlUp
 MyWorkbook.Sheets(2).Rows("1:1").Delete Shift:=xlUp
 MyWorkbook.Sheets(1).Select
Set MyRange = MyWorkbook.Sheets(1).Range("A1").CurrentRegion

 For L = 2 To MyRange.Rows.Count
 MyRange(L, 1).Select
 I = 1
 Save_i = 0
RepriseSersh:
 MyRange(L, 3).Select
I = LingneFilConform(Trim("" & MyRange(L, 3)), RetourneCodeApp(RetourneCodeApp(Replace(UCase(Trim("" & MyRange(L, 11))), "*", "."))), RetourneCodeApp(Replace(UCase(Trim("" & MyRange(L, 14))), "*", ".")), Trim("" & MyRange(L, 8)), UCase(Trim("" & MyRange(L, 9))), UCase(Trim("" & MyRange(L, 12))), Fil, Connecteur, MyRange, MyWorkbook.Sheets(2), L, I)
If I <> 0 Then
    Fil.Cells(I, 5).Select
    Fil.Cells(I, 1) = 1
'    Fil.Cells(I, 5) = "'" & Replace(Round(((Val(Replace(Replace(Trim("" & MyRange(L, 7)), ",", "."), "m", "")) / 2) ^ 2) * 3.1416, 2), ",", ".")
     Fil.Cells(I, 6) = Trim("" & MyRange(L, 8))
     Fil.Cells(I, 9) = "'" & Replace(Replace(Trim("" & MyRange(L, 6)), "m", ""), ",", ".")
     Fil.Cells(I, 10) = LobgueurDeCoupe(Val(Fil.Cells(I, 9)), Val(Fil.Cells(I, 5)), Trim("" & Fil.Cells(I, 24)))
Else
    I = Fil.Range("a1").CurrentRegion.Rows.Count + 1
    Fil.Cells(I, 5).Select
    Fil.Cells(I, 1) = 1
     Fil.Cells(I, 2) = Trim("" & MyRange(L, 3))
'    Fil.Cells(I, 5) = "'" & Replace(Round(((Val(Replace(Replace(Trim("" & MyRange(L, 7)), ",", "."), "m", "")) / 2) ^ 2) * 3.1416, 2), ",", ".")
     Fil.Cells(I, 6) = Trim("" & MyRange(L, 8))
     Fil.Cells(I, 9) = "'" & Replace(Replace(Trim("" & MyRange(L, 6)), "m", ""), ",", ".")
     Fil.Cells(I, 10) = LobgueurDeCoupe(Val(Fil.Cells(I, 9)), Val(Fil.Cells(I, 5)), Trim("" & Fil.Cells(I, 24)))
      Fil.Cells(I, 17) = RetourneCodeApp(Replace(UCase(Trim("" & MyRange(L, 11))), "*", "."))
      Fil.Cells(I, 18) = UCase(Trim("" & MyRange(L, 9)))
      If Trim("" & Fil.Cells(I, 18)) = "" Then
        ValPin = 1
        ValPin = ChercheXls(RetourneCodeApp(Replace(Fil.Cells(I, 17), ".", "*")), MyWorkbook.Sheets(2).Cells, False, True, Val(ValPin))
        Fil.Cells(I, 18) = MyWorkbook.Sheets(2).Cells(ValPin + 1, 3).Value
        End If
        
        Fil.Cells(I, 30).Select
       Fil.Cells(I, 29) = RetourneCodeApp(Replace(UCase(Trim("" & MyRange(L, 14))), "*", "."))
      Fil.Cells(I, 30) = UCase(Trim("" & MyRange(L, 12)))
      If Trim("" & Fil.Cells(I, 30)) = "" Then
        ValPin = 1
        ValPin = ChercheXls(Replace(Fil.Cells(I, 29), ".", "*"), MyWorkbook.Sheets(2).Cells, False, True, Val(ValPin))
        Fil.Cells(I, 30) = MyWorkbook.Sheets(2).Cells(ValPin + 1, 3).Value
        End If
End If
 Next
  MyRange.Delete Shift:=xlUp
  MyWorkbook.Sheets(1).Rows("1:1").Delete Shift:=xlUp
 MyWorkbook.Sheets(1).Rows("1:1").Delete Shift:=xlUp
 MyWorkbook.Sheets(1).Rows("1:1").Delete Shift:=xlUp
  MyWorkbook.Sheets(1).Cells(1, 6) = "*"
 Set MyRange = MyWorkbook.Sheets(1).Range("A1").CurrentRegion
 While MyRange.Columns.Count <> 1 And MyRange.Rows.Count <> 1
    I = 1
    I = ChercheXls(RetourneCodeApp(Replace(MyRange(1, 1), "*", ".")), Connecteur.Cells, True, True, I)
    If I <> 0 Then
        Connecteur.Cells(I, 1) = 1
    Else
       I = Connecteur.Range("a1").CurrentRegion.Rows.Count + 1
       Connecteur.Cells(I, 1).Select
        Connecteur.Cells(I, 1) = 1
        Connecteur.Cells(I, 6) = "'" & RetourneCodeApp(Replace(MyRange(1, 1), "*", "."))
        Connecteur.Cells(I, 2) = "'" & MyRange(1, 5)
         Connecteur.Cells(I, 4) = 0
    End If
        
MyRange.Delete Shift:=xlUp
  MyWorkbook.Sheets(1).Rows("1:1").Delete Shift:=xlUp
    MyWorkbook.Sheets(1).Cells(1, 6) = "*"
  Set MyRange = MyWorkbook.Sheets(1).Range("A1").CurrentRegion
  Wend
MyWorkbook.Close False
Set MyWorkbook = Nothing
MyExcel.Quit
Set MyExcel = Nothing
 End Sub
Function LobgueurDeCoupe(Longueur As Double, Section As Double, Preco As String) As String
    If InStr(1, UCase(Preco), "TOR") <> 0 Then
        LobgueurDeCoupe = Replace("'" & (Longueur + 400), ",", ".")
        Exit Function
    End If
    If Section < 4 Then
        LobgueurDeCoupe = Replace("'" & (Longueur + 300), ",", ".")
        Exit Function
    End If
    LobgueurDeCoupe = Replace("'" & (Longueur + 150), ",", ".")
        Exit Function
End Function
Function LingneFilConform(Liai As String, app1 As String, app2 As String, Couleur As String, Pin1 As String, Pin2 As String, Fil As Object, Con As Object, MyRange As Range, SheetCon As Worksheet, L As Long, I As Long) As Long
Dim I_Liai As Long
Dim I_Pin1 As Long
Dim I_Pin2 As Long
Dim SpitPin
Dim Save_I_Liai As Long
I_Liai = I
Reprise:
I_Liai = ChercheXls(Liai, Fil.Range("a1").CurrentRegion, True, Start:=I_Liai)
If Save_I_Liai > I_Liai Then
    I_Liai = 0
    GoTo Fin
End If
If Save_I_Liai = I_Liai Then
 GoTo Fin
End If
Save_I_Liai = I_Liai
'fil.Cells(I_Liai, 29).Select
If I_Liai = 0 Then GoTo Fin

If UCase(Fil.Cells(I_Liai, 17).Value) <> UCase(app1) Then
    I_Liai = I_Liai + 1
    GoTo Reprise
End If
Fil.Cells(I_Liai, 29).Select
If UCase(Fil.Cells(I_Liai, 29).Value) <> UCase(app2) Then
     I_Liai = I_Liai + 1
    GoTo Reprise
End If
I_Pin1 = 1
ReprisePin1:
If Pin1 <> "" Then
    If UCase(Fil.Cells(I_Liai, 18).Value) <> UCase(Pin1) Then
          I_Liai = I_Liai + 1
    GoTo Reprise
    End If
Else
    I_Pin1 = ChercheXls(Replace(app1, ".", "*"), SheetCon.Cells, False, True, I_Pin1)
    Pin1 = SheetCon.Cells(I_Pin1 + 1, 3).Value
    GoTo ReprisePin1
End If
I_Pin1 = 1

ReprisePin2:
If Pin2 <> "" Then
        If UCase(Fil.Cells(I_Liai, 30).Value) <> UCase(Pin2) Then
          I_Liai = I_Liai + 1
    GoTo Reprise
    End If
Else
     I_Pin1 = ChercheXls(Replace(app2, ".", "*"), SheetCon.Cells, False, True, I_Pin1)
    Pin2 = SheetCon.Cells(I_Pin1 + 1, 3).Value
    GoTo ReprisePin2
End If
Fin:
LingneFilConform = I_Liai
End Function
Function RetourneCodeApp(Valeur As String) As String
If InStr(1, Valeur, ".") = 0 Then
    If InStr(1, Valeur, "-") = 0 Then
       RetourneCodeApp = Left(Valeur, Len(Valeur) - 2) & "." & Right(Valeur, 2)
    Else
       RetourneCodeApp = Valeur
    End If
    Else
         RetourneCodeApp = Valeur
End If
End Function
