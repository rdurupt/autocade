Attribute VB_Name = "ValiderConnecteurs"
Sub LireRep()
Dim Fso As New FileSystemObject
Dim Rep
Dim MyExcel As New EXCEL.Application
Dim MyWorkbook As EXCEL.Workbook
Dim MyWorksheet As EXCEL.Worksheet
Set MyWorkbook = MyExcel.Workbooks.Add
Set MyWorksheet = MyWorkbook.Sheets.Add
Dim MyRange As EXCEL.Range
MyExcel.Visible = True
i = 0
 Rep = Dir("\\Enc-srv-prod-01\donnees d entreprise\Utilitaires\cablage\Librairies\Connecteurs\*.*")
 While Rep <> ""

 If InStr(1, UCase(Rep), ".DWG") <> 0 Then
 i = i + 1
 MyWorksheet.Cells(i, 1) = "\\Enc-srv-prod-01\donnees d entreprise\Utilitaires\cablage\Librairies\Connecteurs\" & Rep
  MyWorksheet.Cells(i, 2) = Fso.GetFile("\\Enc-srv-prod-01\donnees d entreprise\Utilitaires\cablage\Librairies\Connecteurs\" & Rep).DateLastModified
 End If
 Rep = Dir
 Wend
  Rep = Dir("\\Enc-srv-prod-01\donnees d entreprise\Utilitaires\cablage\Librairies\Connecteurs\Construction connecteurs\*.*")
 While Rep <> ""

 If InStr(1, UCase(Rep), ".DWG") <> 0 Then
 i = i + 1
 MyWorksheet.Cells(i, 1) = "\\Enc-srv-prod-01\donnees d entreprise\Utilitaires\cablage\Librairies\Connecteurs\Construction connecteurs\" & Rep
  MyWorksheet.Cells(i, 2) = Fso.GetFile("\\Enc-srv-prod-01\donnees d entreprise\Utilitaires\cablage\Librairies\Connecteurs\Construction connecteurs\" & Rep).DateLastModified
 End If
 Rep = Dir
 Wend
 
 

  Rep = Dir("\\Enc-srv-prod-01\donnees d entreprise\Utilitaires\cablage\Librairies\Connecteurs RD\*.*")
 While Rep <> ""

 If InStr(1, UCase(Rep), ".DWG") <> 0 Then
 DoEvents
 i = i + 1
 MyWorksheet.Cells(i, 1) = "\\Enc-srv-prod-01\donnees d entreprise\Utilitaires\cablage\Librairies\Connecteurs RD\" & Rep
  MyWorksheet.Cells(i, 2) = Fso.GetFile("\\Enc-srv-prod-01\donnees d entreprise\Utilitaires\cablage\Librairies\Connecteurs RD\" & Rep).DateLastModified
 End If
 Rep = Dir
 Wend
 
 Set MyRange = MyWorksheet.Range("a1").CurrentRegion
    MyRange.Sort Key1:=Range("B1"), Order1:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
        
        For i = 1 To MyRange.Rows.Count
            Rep = Dir(MyRange(i, 1))
            If Fso.FileExists("\\Enc-srv-prod-01\donnees d entreprise\Utilitaires\cablage\Librairies\Blocs de construction\ConecteurXls\" & Rep) = True Then
                Fso.DeleteFile "\\Enc-srv-prod-01\donnees d entreprise\Utilitaires\cablage\Librairies\Blocs de construction\ConecteurXls\" & Rep
            End If
         Fso.CopyFile MyRange(i, 1), "\\Enc-srv-prod-01\donnees d entreprise\Utilitaires\cablage\Librairies\Blocs de construction\ConecteurXls\" & Rep
         DoEvents
        Next i
        ChargeConecteur
'
End Sub
Sub ChargeConecteur()
    Dim Rep As String
    Dim NuFichier As Long
   Dim pathUser As String
    Dim Block As AcadBlockReference
    Msg = ""
    
     AutoApp.Documents.Add
      Rep = Dir("\\Enc-srv-prod-01\donnees d entreprise\Utilitaires\cablage\Librairies\Blocs de construction\ConecteurXls\*.*")
      InsertPointLigneTableau_fils(0) = 1
      InsertPointLigneTableau_fils(1) = 1
      InsertPointLigneTableau_fils(2) = 1
       While Rep <> ""
            
           Set Block = FunInsBlock2("\\Enc-srv-prod-01\donnees d entreprise\Utilitaires\cablage\Librairies\Blocs de construction\ConecteurXls\" & Rep, InsertPointLigneTableau_fils, 1, 1, 1, 1)
           If ErrInsert = False Then
           Att = Block.GetAttributes
           If IsConnecteurs(Att) = False Then
           Msg = Msg & "***************************************************" & vbCrLf
           Msg = Msg & "\\Enc-srv-prod-01\donnees d entreprise\Utilitaires\cablage\Librairies\Blocs de construction\ConecteurXls\" & Rep
           Msg = Msg & " N 'est pas un connecteur" & vbCrLf
            Msg = Msg & "***************************************************" & vbCrLf & vbCrLf
           Else
              ScanAtt Att, "\\Enc-srv-prod-01\donnees d entreprise\Utilitaires\cablage\Librairies\Blocs de construction\ConecteurXls\" & Rep
           End If
           End If
            Rep = Dir
         DoEvents
        
       Wend
           pathUser = Environ("USERPROFILE")
    pathUser = pathUser + "\Mes Documents\"
pathUser = pathUser & "Erreur Connecteurs.txt"
NuFichier = FreeFile
 Open pathUser For Output As #NuFichier
    Print #NuFichier, Msg
    Close #NuFichier

    Shell "notepad.exe " & pathUser, vbMaximizedFocus
    AutoApp.ActiveDocument.Close , False
End Sub
Sub ScanAtt(Att, Conecteur As String)
Dim Liai As String
Dim Fil As String
Dim Mar As String
Dim txt  As String
Dim txt1  As String
Dim txt2  As String
Dim IsNum As Boolean
Dim Coupe As Boolean
Coupe = False
Liai = ""
Fil = ""
Mar = ""

    For i = LBound(Att) To UBound(Att)
   
        If InStr(1, UCase("" & Att(i).TagString), "LIAI") <> 0 Then
            If Trim("" & Liai) = "" Then
                Liai = Trim(UCase("" & Att(i).TagString))
                For i2 = Len("LIAI") To Len(Att(i).TagString)
                    txt = Mid(UCase("" & Att(i).TagString), i2 + 1, 1)
                    If Not IsNumeric(txt) Then
                        If IsNum = False Then
                            txt1 = txt1 & txt
                        Else
                             txt2 = txt2 & txt
                        End If
                    Else
                        IsNum = True
                    End If
            
        
            Next i2
            End If
        
        Exit For
           
        End If
           
    Next i
    If Liai = "" Then
     Msg = Msg & "***************************************************" & vbCrLf
            Msg = Msg & "Erreur d'attribut : " & vbCrLf
            Msg = Msg & "Attributs Liai, Fil et Mar non trouvés ? " & vbCrLf
            Msg = Msg & "pour le connecteur : " & Conecteur & vbCrLf
            Msg = Msg & "***************************************************" & vbCrLf & vbCrLf
            Exit Sub
    Else
        Liai = "LIAI"
        Fil = "FiL"
        Mar = "MAR"
    End If
  For i = LBound(Att) To UBound(Att)
Reprise:
    If InStr(1, UCase("" & Att(i).TagString), Liai) <> 0 Then
      txt = Mid(UCase("" & Att(i).TagString), Len(Liai) + Len(txt1) + 1, Len(Att(i).TagString) - (Len(Liai) + Len(txt1) + Len(txt2)))
        If IsNum = True Then
            If Trim(txt) <> "" Then
                If Not IsNumeric(txt) Then
                    Msg = Msg & "***************************************************" & vbCrLf
                Msg = Msg & "Vérifiez la pertinence de l'attribut" & vbCrLf
                Msg = Msg & "Erreur d'attribut : " & UCase("" & Att(i).TagString) & vbCrLf
                Msg = Msg & "pour le connecteur : " & Conecteur & vbCrLf
                Msg = Msg & "***************************************************" & vbCrLf & vbCrLf
                End If
            End If
       Else
        If Trim(txt) <> "" Then
               Msg = Msg & "***************************************************" & vbCrLf
                Msg = Msg & "Vérifiez la pertinence de l'attribut" & vbCrLf
                Msg = Msg & "Erreur d'attribut : " & UCase("" & Att(i).TagString) & vbCrLf
                Msg = Msg & "pour le connecteur : " & Conecteur & vbCrLf
                Msg = Msg & "***************************************************" & vbCrLf & vbCrLf
        End If
       End If
       
         
        End If
       
     
    If InStr(1, UCase("" & Att(i).TagString), Fil) <> 0 Then
      txt = Mid(UCase("" & Att(i).TagString), Len(Fil) + Len(txt1) + 1, Len(Att(i).TagString) - (Len(Fil) + Len(txt1) + Len(txt2)))
        If IsNum = True Then
            If Trim(txt) <> "" Then
                If Not IsNumeric(txt) Then
                    Msg = Msg & "***************************************************" & vbCrLf
                Msg = Msg & "Vérifiez la pertinence de l'attribut" & vbCrLf
                Msg = Msg & "Erreur d'attribut : " & UCase("" & Att(i).TagString) & vbCrLf
                Msg = Msg & "pour le connecteur : " & Conecteur & vbCrLf
                Msg = Msg & "***************************************************" & vbCrLf & vbCrLf
                End If
            End If
       Else
        If Trim(txt) <> "" Then
               Msg = Msg & "***************************************************" & vbCrLf
                Msg = Msg & "Vérifiez la pertinence de l'attribut" & vbCrLf
                Msg = Msg & "Erreur d'attribut : " & UCase("" & Att(i).TagString) & vbCrLf
                Msg = Msg & "pour le connecteur : " & Conecteur & vbCrLf
                Msg = Msg & "***************************************************" & vbCrLf & vbCrLf
        End If
       End If
       
         
        End If
        If InStr(1, UCase("" & Att(i).TagString), Mar) <> 0 Then
      txt = Mid(UCase("" & Att(i).TagString), Len(Mar) + Len(txt1) + 1, Len(Att(i).TagString) - (Len(Mar) + Len(txt1) + Len(txt2)))
        If IsNum = True Then
            If Trim(txt) <> "" Then
                If Not IsNumeric(txt) Then
                    Msg = Msg & "***************************************************" & vbCrLf
                Msg = Msg & "Vérifiez la pertinence de l'attribut" & vbCrLf
                Msg = Msg & "Erreur d'attribut : " & UCase("" & Att(i).TagString) & vbCrLf
                Msg = Msg & "pour le connecteur : " & Conecteur & vbCrLf
                Msg = Msg & "***************************************************" & vbCrLf & vbCrLf
                End If
            End If
       Else
        If Trim(txt) <> "" Then
               Msg = Msg & "***************************************************" & vbCrLf
                Msg = Msg & "Vérifiez la pertinence de l'attribut" & vbCrLf
                Msg = Msg & "Erreur d'attribut : " & UCase("" & Att(i).TagString) & vbCrLf
                Msg = Msg & "pour le connecteur : " & Conecteur & vbCrLf
                Msg = Msg & "***************************************************" & vbCrLf & vbCrLf
        End If
       End If
       
         
        End If
    
   
  Next i
End Sub
