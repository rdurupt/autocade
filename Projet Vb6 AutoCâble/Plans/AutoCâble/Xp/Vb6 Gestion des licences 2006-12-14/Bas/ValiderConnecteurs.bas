Attribute VB_Name = "ValiderConnecteurs"
Global PasConnecteur As Boolean
Sub LireRepEval()
Dim Rep
If Dir(App.Path & "\DossierAplication\TestConnecteurs\Test.Ok") <> "" Then
    MsgBox "La Macro de test connecteur est déjà en cour d'exécution", vbInformation
    Exit Sub
End If
Rep = Dir(App.Path & "\DossierAplication\TestConnecteurs\ConnecteursTest\*.DWG")
If Rep = "" Then Exit Sub
Dim Fso As New FileSystemObject
Dim MyExcel As New EXCEL.Application
MyExcel.DisplayAlerts = False
Dim MyWorkbook As EXCEL.Workbook
Dim MyWorksheet As EXCEL.Worksheet
 
 Set MyWorkbook = MyExcel.Workbooks.Add
Set MyWorksheet = MyWorkbook.Sheets.Add
Dim MyRange As EXCEL.Range
Fso.CreateTextFile App.Path & "\DossierAplication\TestConnecteurs\Test.Ok"

MyExcel.Visible = True
Set DocAutoCad = AutoApp.Documents.Add
'AutoApp.Visible = False
I = 1
MyWorksheet.Cells(I, 1) = "Valider"
MyWorksheet.Cells(I, 2) = "BLOC"
MyWorksheet.Cells(I, 3) = "Date"
MyWorksheet.Cells(I, 4) = "ERREUR"
 While Rep <> ""

 If InStr(1, UCase(Rep), ".DWG") <> 0 Then
 I = I + 1
  MyWorksheet.Cells(I, 2).Select
 MyWorksheet.Cells(I, 2) = App.Path & "\DossierAplication\TestConnecteurs\ConnecteursTest\" & Rep
  MyWorksheet.Cells(I, 3) = Fso.GetFile(App.Path & "\DossierAplication\TestConnecteurs\ConnecteursTest\" & Rep).DateLastModified
 End If
 Rep = Dir
 DoEvents
 Wend
 
 
 
 Set MyRange = MyWorksheet.Range("a1").CurrentRegion
    MyRange.Sort Key1:=Range("C1"), Order1:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
        
        For I = 2 To MyRange.Rows.Count
        PasConnecteur = False
            Rep = Dir("" & MyRange(I, 2))
             MyRange(I, 1).Select
             AutoApp.Visible = True
            If ChargeConecteur(MyRange(I, 2), MyRange(I, 4)) = True Then
                MyRange(I, 1) = "OUI"
                MyRange(I, 2) = App.Path & "\DossierAplication\TestConnecteurs\ConnecteursValider\" & Rep
                If Fso.FileExists(App.Path & "\DossierAplication\TestConnecteurs\ConnecteursValider\" & Rep) = True Then
                    Fso.DeleteFile App.Path & "\DossierAplication\TestConnecteurs\ConnecteursValider\" & Rep
                End If
                    Fso.CopyFile App.Path & "\DossierAplication\TestConnecteurs\ConnecteursTest\" & Rep, App.Path & "\DossierAplication\TestConnecteurs\ConnecteursValider\" & Rep
            Else
                MyRange(I, 1) = "NON"
                MyRange(I, 2) = App.Path & "\DossierAplication\TestConnecteurs\ConnecteursDouteux\" & Rep
                If PasConnecteur = False Then
                     If Fso.FileExists(App.Path & "\DossierAplication\TestConnecteurs\ConnecteursDouteux\" & Rep) = True Then
                        Fso.DeleteFile App.Path & "\DossierAplication\TestConnecteurs\ConnecteursDouteux\" & Rep
                    End If
                     Fso.CopyFile App.Path & "\DossierAplication\TestConnecteurs\ConnecteursTest\" & Rep, App.Path & "\DossierAplication\TestConnecteurs\ConnecteursDouteux\" & Rep
              Else
                    If Fso.FileExists(App.Path & "\DossierAplication\TestConnecteurs\PasConnecteurs\" & Rep) = True Then
                        Fso.DeleteFile App.Path & "\DossierAplication\TestConnecteurs\PasConnecteurs\" & Rep
                    End If
                      Fso.CopyFile App.Path & "\DossierAplication\TestConnecteurs\ConnecteursTest\" & Rep, App.Path & "\DossierAplication\TestConnecteurs\PasConnecteurs\" & Rep
                       MyRange(I, 2) = App.Path & "\DossierAplication\TestConnecteurs\PasConnecteurs\" & Rep

              End If
            End If
               
         DoEvents
        Next I
         Fso.DeleteFile App.Path & "\DossierAplication\TestConnecteurs\ConnecteursRapport\*.*"
         MyWorkbook.SaveAs App.Path & "\DossierAplication\TestConnecteurs\ConnecteursRapport\ConnecteursRapport", ReadOnlyRecommended:=True
        DocAutoCad.Close , False
        Fso.DeleteFile App.Path & "\DossierAplication\TestConnecteurs\*.*"
        Fso.DeleteFile App.Path & "\DossierAplication\TestConnecteurs\ConnecteursTest\*.*"
        
        
        MyExcel.Quit
        AutoApp.Visible = False
        Set Fso = Nothing
        Set MyExcel = Nothing
        Set MyWorkbook = Nothing
        Set MyWorksheet = Nothing
'
End Sub
Function ChargeConecteur(Bloc As String, MyRange As Range) As Boolean
    Dim Rep As String
    Dim NuFichier As Long
   Dim pathUser As String
    Dim Block As Object
    msg = ""
    On Error Resume Next
    ChargeConecteur = True
     
      InsertPointLigneTableau_fils(0) = 1
      InsertPointLigneTableau_fils(1) = 1
      InsertPointLigneTableau_fils(2) = 1
                 
           Set Block = FunInsBlock2(Bloc, InsertPointLigneTableau_fils, "1", 0, 1, 1)
    If ErrInsert = False Then
   
        Att = Block.GetAttributes
'        DocAutoCad.Application.Visible = True
        MyRange = ScanAtt(Att)
        If Trim("" & MyRange.Value) <> "" Then
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
         MyRange = "Err à l'insertion"
    End If
'            Rep = Dir
'         DoEvents

        Block.Delete
        DocAutoCad.PurgeAll
      Set Block = Nothing
   
    
End Function
Function ScanAtt(Att) As String
Dim Liai As String
Dim Fil As String
Dim Mar As String
Dim Txt  As String
Dim txt1  As String
Dim txt2  As String
Dim IsNum As Boolean
Dim Coupe As Boolean
Dim DoublonAtt As New Collection
Coupe = False
Liai = ""
Fil = ""
Mar = ""
msg = ""
PasConnecteur = False
If IsConnecteurs(Att) = False Then
PasConnecteur = True
    msg = "N'est pas un connecteur"
    GoTo Fin
End If
    For I = LBound(Att) To UBound(Att)
   
        If InStr(1, UCase("" & Att(I).TagString), "LIAI") <> 0 Then
            If Trim("" & Liai) = "" Then
                Liai = Trim(UCase("" & Att(I).TagString))
                For I2 = Len("LIAI") To Len(Att(I).TagString)
                    Txt = Mid(UCase("" & Att(I).TagString), I2 + 1, 1)
                    If Not IsNumeric(Txt) Then
                        If IsNum = False Then
                            txt1 = txt1 & Txt
                        Else
                             txt2 = txt2 & Txt
                        End If
                    Else
                        IsNum = True
                    End If
            
        
            Next I2
            End If
        
        Exit For
           
        End If
           
    Next I
    If Liai = "" Then
     msg = msg & "***************************************************" & Chr(10)
            msg = msg & "Erreur d'attribut : " & Chr(10)
            msg = msg & "Attributs Liai non trouvés ? " & Chr(10)
'           ' Msg = Msg & "pour le connecteur : " & Conecteur & chr(10)
            msg = msg & "***************************************************" & Chr(10) & Chr(10)
           GoTo Fin
    Else
        Liai = "LIAI"
        Fil = "FiL"
        Mar = "MAR"
    End If
  For I = LBound(Att) To UBound(Att)
  On Error Resume Next
  MyAtt = ""
  MyAtt = DoublonAtt(UCase("" & Att(I).TagString))
  If Err Then
    Err.Clear
    DoublonAtt.Add I, UCase("" & Att(I).TagString)
  Else
     msg = msg & "***************************************************" & Chr(10)
                msg = msg & "L'attribut existe déjà attention aux doublons!" & Chr(10)
                msg = msg & "Erreur d'attribut : " & UCase("" & Att(I).TagString) & Chr(10)
'               ' Msg = Msg & "pour le connecteur : " & Conecteur & chr(10)
                msg = msg & "***************************************************" & Chr(10) & Chr(10)
  End If
  On Error GoTo 0
Reprise:
    If InStr(1, UCase("" & Att(I).TagString), Liai) <> 0 Then
      Txt = Mid(UCase("" & Att(I).TagString), Len(Liai) + Len(txt1) + 1, Len(Att(I).TagString) - (Len(Liai) + Len(txt1) + Len(txt2)))
        If IsNum = True Then
            If Trim(Txt) <> "" Then
                If Not IsNumeric(Txt) Then
                    msg = msg & "***************************************************" & Chr(10)
                msg = msg & "Vérifiez la pertinence de l'attribut" & Chr(10)
                msg = msg & "Erreur d'attribut : " & UCase("" & Att(I).TagString) & Chr(10)
'               ' Msg = Msg & "pour le connecteur : " & Conecteur & chr(10)
                msg = msg & "***************************************************" & Chr(10) & Chr(10)
                End If
            End If
       Else
        If Trim(Txt) <> "" Then
               msg = msg & "***************************************************" & Chr(10)
                msg = msg & "Vérifiez la pertinence de l'attribut" & Chr(10)
                msg = msg & "Erreur d'attribut : " & UCase("" & Att(I).TagString) & Chr(10)
               ' Msg = Msg & "pour le connecteur : " & Conecteur & chr(10)
                msg = msg & "***************************************************" & Chr(10) & Chr(10)
        End If
       End If
       
         
        End If
       
     
    If InStr(1, UCase("" & Att(I).TagString), Fil) <> 0 Then
      Txt = Mid(UCase("" & Att(I).TagString), Len(Fil) + Len(txt1) + 1, Len(Att(I).TagString) - (Len(Fil) + Len(txt1) + Len(txt2)))
        If IsNum = True Then
            If Trim(Txt) <> "" Then
                If Not IsNumeric(Txt) Then
                    msg = msg & "***************************************************" & Chr(10)
                msg = msg & "Vérifiez la pertinence de l'attribut" & Chr(10)
                msg = msg & "Erreur d'attribut : " & UCase("" & Att(I).TagString) & Chr(10)
               ' Msg = Msg & "pour le connecteur : " & Conecteur & chr(10)
                msg = msg & "***************************************************" & Chr(10) & Chr(10)
                End If
            End If
       Else
        If Trim(Txt) <> "" Then
               msg = msg & "***************************************************" & Chr(10)
                msg = msg & "Vérifiez la pertinence de l'attribut" & Chr(10)
                msg = msg & "Erreur d'attribut : " & UCase("" & Att(I).TagString) & Chr(10)
               ' Msg = Msg & "pour le connecteur : " & Conecteur & chr(10)
                msg = msg & "***************************************************" & Chr(10) & Chr(10)
        End If
       End If
       
         
        End If
        If InStr(1, UCase("" & Att(I).TagString), Mar) <> 0 Then
      Txt = Mid(UCase("" & Att(I).TagString), Len(Mar) + Len(txt1) + 1, Len(Att(I).TagString) - (Len(Mar) + Len(txt1) + Len(txt2)))
        If IsNum = True Then
            If Trim(Txt) <> "" Then
                If Not IsNumeric(Txt) Then
                    msg = msg & "***************************************************" & Chr(10)
                msg = msg & "Vérifiez la pertinence de l'attribut" & Chr(10)
                msg = msg & "Erreur d'attribut : " & UCase("" & Att(I).TagString) & Chr(10)
               ' Msg = Msg & "pour le connecteur : " & Conecteur & chr(10)
                msg = msg & "***************************************************" & Chr(10) & Chr(10)
                End If
            End If
       Else
        If Trim(Txt) <> "" Then
               msg = msg & "***************************************************" & Chr(10)
                msg = msg & "Vérifiez la pertinence de l'attribut" & Chr(10)
                msg = msg & "Erreur d'attribut : " & UCase("" & Att(I).TagString) & Chr(10)
               ' Msg = Msg & "pour le connecteur : " & Conecteur & chr(10)
                msg = msg & "***************************************************" & Chr(10) & Chr(10)
        End If
       End If
       
         
        End If
    
   
  Next I
Fin:
  ScanAtt = msg
End Function
