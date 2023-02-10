Attribute VB_Name = "ValiderConnecteurs"
Public PasConnecteur As Boolean
Public Sub LireRepEval()
Dim Rep
If Dir(App.Path & "\ConnecteursTest\Test.Ok") <> "" Then
    MsgBox "La Macro de test connecteur est déjà en cour d'exécution", vbInformation
    Exit Sub
End If
Rep = Dir(App.Path & "\ConnecteursTest\*.DWG")
If Rep = "" Then Exit Sub
Dim Fso As New FileSystemObject
Dim MyExcel As New EXCEL.Application
Dim MyWorkbook As EXCEL.Workbook
Dim MyWorksheet As EXCEL.Worksheet
 
 Set MyWorkbook = MyExcel.Workbooks.Add
Set MyWorksheet = MyWorkbook.Sheets.Add
Dim Myrange As EXCEL.Range
Fso.CreateTextFile App.Path & "\ConnecteursTest\Test.Ok"

MyExcel.Visible = True
AutoApp.Documents.Add
i = 1
MyWorksheet.Cells(i, 1) = "Valider"
MyWorksheet.Cells(i, 2) = "BLOC"
MyWorksheet.Cells(i, 3) = "Date"
MyWorksheet.Cells(i, 4) = "ERREUR"
 While Rep <> ""

 If InStr(1, UCase(Rep), ".DWG") <> 0 Then
 i = i + 1
 MyWorksheet.Cells(i, 2) = App.Path & "\ConnecteursTest\" & Rep
  MyWorksheet.Cells(i, 3) = Fso.GetFile(App.Path & "\ConnecteursTest\" & Rep).DateLastModified
 End If
 Rep = Dir
 Wend
 
 
 
 Set Myrange = MyWorksheet.Range("a1").CurrentRegion
    Myrange.Sort Key1:=Range("C1"), Order1:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
        
        For i = 2 To Myrange.Rows.Count
        PasConnecteur = False
            Rep = Dir(Myrange(i, 2))
             Myrange(i, 1).Select
            If ChargeConecteur(App.Path & "\ConnecteursTest\" & Rep, Myrange(i, 4)) = True Then
                Myrange(i, 1) = "OUI"
                Myrange(i, 2) = App.Path & "\ConnecteursValider\" & Rep
                If Fso.FileExists(App.Path & "\ConnecteursValider\" & Rep) = True Then
                    Fso.DeleteFile App.Path & "\ConnecteursValider\" & Rep
                End If
                    Fso.CopyFile App.Path & "\ConnecteursTest\" & Rep, App.Path & "\ConnecteursValider\" & Rep
            Else
                Myrange(i, 1) = "NON"
                Myrange(i, 2) = App.Path & "\ConnecteursDouteux\" & Rep
                If PasConnecteur = False Then
                     If Fso.FileExists(App.Path & "\ConnecteursDouteux\" & Rep) = True Then
                        Fso.DeleteFile App.Path & "\ConnecteursDouteux\" & Rep
                    End If
                     Fso.CopyFile App.Path & "\ConnecteursTest\" & Rep, App.Path & "\ConnecteursDouteux\" & Rep
              Else
                    If Fso.FileExists(App.Path & "\PasConnecteurs\" & Rep) = True Then
                        Fso.DeleteFile App.Path & "\PasConnecteurs\" & Rep
                    End If
                      Fso.CopyFile App.Path & "\ConnecteursTest\" & Rep, App.Path & "\PasConnecteurs\" & Rep
                       Myrange(i, 2) = App.Path & "\PasConnecteurs\" & Rep

              End If
            End If
               
         DoEvents
        Next i
         Fso.DeleteFile App.Path & "\ConnecteursRapport\*.*"
         MyWorkbook.SaveAs App.Path & "\ConnecteursRapport\ConnecteursRapport"
        AutoApp.Documents(0).Close , False
        Fso.DeleteFile App.Path & "\ConnecteursTest\*.*"
        MyExcel.Quit
        Set Fso = Nothing
        Set MyExcel = Nothing
        Set MyWorkbook = Nothing
        Set MyWorksheet = Nothing
'
End Sub
Function ChargeConecteur(Bloc As String, Myrange As Range) As Boolean
    Dim Rep As String
    Dim NuFichier As Long
   Dim pathUser As String
    Dim Block As AcadBlockReference
    Msg = ""
    On Error Resume Next
    ChargeConecteur = True
     
      InsertPointLigneTableau_fils(0) = 1
      InsertPointLigneTableau_fils(1) = 1
      InsertPointLigneTableau_fils(2) = 1
                 
           Set Block = FunInsBlock2(Bloc, InsertPointLigneTableau_fils, 1, 1, 1, 1)
    If ErrInsert = False Then
   
        Att = Block.GetAttributes
'        AutoApp.Documents(0).Application.Visible = True
        Myrange = ScanAtt(Att)
        If Trim("" & Myrange.Value) <> "" Then
        ChargeConecteur = False
'            Msg = Msg & "***************************************************" & vbCrLf
'            Msg = Msg & "\\Enc-srv-prod-01\donnees d entreprise\Utilitaires\cablage\Librairies\Blocs de construction\ConecteurXls\" & Rep
'            Msg = Msg & " N 'est pas un connecteur" & vbCrLf
'            Msg = Msg & "***************************************************" & vbCrLf & vbCrLf
'            Else
'              ScanAtt Att, "\\Enc-srv-prod-01\donnees d entreprise\Utilitaires\cablage\Librairies\Blocs de construction\ConecteurXls\" & Rep
        End If
        Else
             ChargeConecteur = False
         Myrange = "Err à l'insertion"
    End If
'            Rep = Dir
'         DoEvents

        Block.Delete
        AutoApp.Documents(0).PurgeAll
      Set Block = Nothing
   
    
End Function
Function ScanAtt(Att) As String
Dim Liai As String
Dim Fil As String
Dim Mar As String
Dim txt  As String
Dim txt1  As String
Dim txt2  As String
Dim IsNum As Boolean
Dim Coupe As Boolean
Dim DoublonAtt As New Collection
Coupe = False
Liai = ""
Fil = ""
Mar = ""
Msg = ""
PasConnecteur = False
If IsConnecteurs(Att) = False Then
PasConnecteur = True
    Msg = "N'est pas un connecteur"
    GoTo Fin
End If
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
     Msg = Msg & "***************************************************" & Chr(10)
            Msg = Msg & "Erreur d'attribut : " & Chr(10)
            Msg = Msg & "Attributs Liai non trouvés ? " & Chr(10)
'           ' Msg = Msg & "pour le connecteur : " & Conecteur & chr(10)
            Msg = Msg & "***************************************************" & Chr(10) & Chr(10)
           GoTo Fin
    Else
        Liai = "LIAI"
        Fil = "FiL"
        Mar = "MAR"
    End If
  For i = LBound(Att) To UBound(Att)
  On Error Resume Next
  MyAtt = ""
  MyAtt = DoublonAtt(UCase("" & Att(i).TagString))
  If Err Then
    Err.Clear
    DoublonAtt.Add i, UCase("" & Att(i).TagString)
  Else
     Msg = Msg & "***************************************************" & Chr(10)
                Msg = Msg & "L'attribut existe déjà attention aux doublons!" & Chr(10)
                Msg = Msg & "Erreur d'attribut : " & UCase("" & Att(i).TagString) & Chr(10)
'               ' Msg = Msg & "pour le connecteur : " & Conecteur & chr(10)
                Msg = Msg & "***************************************************" & Chr(10) & Chr(10)
  End If
  On Error GoTo 0
Reprise:
    If InStr(1, UCase("" & Att(i).TagString), Liai) <> 0 Then
      txt = Mid(UCase("" & Att(i).TagString), Len(Liai) + Len(txt1) + 1, Len(Att(i).TagString) - (Len(Liai) + Len(txt1) + Len(txt2)))
        If IsNum = True Then
            If Trim(txt) <> "" Then
                If Not IsNumeric(txt) Then
                    Msg = Msg & "***************************************************" & Chr(10)
                Msg = Msg & "Vérifiez la pertinence de l'attribut" & Chr(10)
                Msg = Msg & "Erreur d'attribut : " & UCase("" & Att(i).TagString) & Chr(10)
'               ' Msg = Msg & "pour le connecteur : " & Conecteur & chr(10)
                Msg = Msg & "***************************************************" & Chr(10) & Chr(10)
                End If
            End If
       Else
        If Trim(txt) <> "" Then
               Msg = Msg & "***************************************************" & Chr(10)
                Msg = Msg & "Vérifiez la pertinence de l'attribut" & Chr(10)
                Msg = Msg & "Erreur d'attribut : " & UCase("" & Att(i).TagString) & Chr(10)
               ' Msg = Msg & "pour le connecteur : " & Conecteur & chr(10)
                Msg = Msg & "***************************************************" & Chr(10) & Chr(10)
        End If
       End If
       
         
        End If
       
     
    If InStr(1, UCase("" & Att(i).TagString), Fil) <> 0 Then
      txt = Mid(UCase("" & Att(i).TagString), Len(Fil) + Len(txt1) + 1, Len(Att(i).TagString) - (Len(Fil) + Len(txt1) + Len(txt2)))
        If IsNum = True Then
            If Trim(txt) <> "" Then
                If Not IsNumeric(txt) Then
                    Msg = Msg & "***************************************************" & Chr(10)
                Msg = Msg & "Vérifiez la pertinence de l'attribut" & Chr(10)
                Msg = Msg & "Erreur d'attribut : " & UCase("" & Att(i).TagString) & Chr(10)
               ' Msg = Msg & "pour le connecteur : " & Conecteur & chr(10)
                Msg = Msg & "***************************************************" & Chr(10) & Chr(10)
                End If
            End If
       Else
        If Trim(txt) <> "" Then
               Msg = Msg & "***************************************************" & Chr(10)
                Msg = Msg & "Vérifiez la pertinence de l'attribut" & Chr(10)
                Msg = Msg & "Erreur d'attribut : " & UCase("" & Att(i).TagString) & Chr(10)
               ' Msg = Msg & "pour le connecteur : " & Conecteur & chr(10)
                Msg = Msg & "***************************************************" & Chr(10) & Chr(10)
        End If
       End If
       
         
        End If
        If InStr(1, UCase("" & Att(i).TagString), Mar) <> 0 Then
      txt = Mid(UCase("" & Att(i).TagString), Len(Mar) + Len(txt1) + 1, Len(Att(i).TagString) - (Len(Mar) + Len(txt1) + Len(txt2)))
        If IsNum = True Then
            If Trim(txt) <> "" Then
                If Not IsNumeric(txt) Then
                    Msg = Msg & "***************************************************" & Chr(10)
                Msg = Msg & "Vérifiez la pertinence de l'attribut" & Chr(10)
                Msg = Msg & "Erreur d'attribut : " & UCase("" & Att(i).TagString) & Chr(10)
               ' Msg = Msg & "pour le connecteur : " & Conecteur & chr(10)
                Msg = Msg & "***************************************************" & Chr(10) & Chr(10)
                End If
            End If
       Else
        If Trim(txt) <> "" Then
               Msg = Msg & "***************************************************" & Chr(10)
                Msg = Msg & "Vérifiez la pertinence de l'attribut" & Chr(10)
                Msg = Msg & "Erreur d'attribut : " & UCase("" & Att(i).TagString) & Chr(10)
               ' Msg = Msg & "pour le connecteur : " & Conecteur & chr(10)
                Msg = Msg & "***************************************************" & Chr(10) & Chr(10)
        End If
       End If
       
         
        End If
    
   
  Next i
Fin:
  ScanAtt = Msg
End Function
