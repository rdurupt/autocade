VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsEtiqette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim App As Collection
Dim TableauEtiquette() As Etiqette
Const preFixFild = "<<-!"
Const SuFixFild = "!->>"
Public Sub PrpareEtiqet(tableau, TableauX)
Set App = Nothing
Set App = New Collection
tableau2 = Split(tableau & Chr(10), Chr(10))
tableau3 = Split(TableauX & Chr(10), Chr(10))
ReDim TableauEtiquette(0)
ReDim TableauEtiquette(UBound(tableau2) - 1)
For I = 0 To UBound(tableau2) - 1

On Error Resume Next
a = ""
a = App(tableau2(I))
If Err Then
Err.Clear
    App.Add I, tableau2(I)

End If
TableauEtiquette(App(tableau2(I))).Code_APP(0) = preFixFild & "Code_APP" & SuFixFild
TableauEtiquette(App(tableau2(I))).Code_APP(1) = tableau2(I)
TableauEtiquette(App(tableau2(I))).DESIGNATION(0) = preFixFild & "DESIGNATION" & SuFixFild
TableauEtiquette(App(tableau2(I))).DESIGNATION(1) = tableau3(I)

TableauEtiquette(App(tableau2(I))).Connecteur(0) = preFixFild & "CONNECTEUR" & SuFixFild
TableauEtiquette(App(tableau2(I))).Ensemble(0) = preFixFild & UCase("Ensemble") & SuFixFild
TableauEtiquette(App(tableau2(I))).Famille(0) = preFixFild & UCase("Famille") & SuFixFild
TableauEtiquette(App(tableau2(I))).PI(0) = preFixFild & UCase("PI") & SuFixFild

TableauEtiquette(App(tableau2(I))).Ref_Joint(0) = preFixFild & UCase("Ref Joint") & SuFixFild
TableauEtiquette(App(tableau2(I))).AlveRef(0) = preFixFild & UCase("Alvé Réf") & SuFixFild

TableauEtiquette(App(tableau2(I))).Verrou(0) = preFixFild & UCase("Ref Verrou") & SuFixFild
TableauEtiquette(App(tableau2(I))).Capot(0) = preFixFild & UCase("Ref Capot") & SuFixFild
TableauEtiquette(App(tableau2(I))).Bouchon(0) = preFixFild & UCase("Ref Bouchon") & SuFixFild



Next
End Sub
Public Sub RenseigneChamp(Champ As String, Valeur As String, Optional NoRepat As Boolean)
Dim NbFamille
If Trim("" & Valeur) = "" Then Exit Sub
 
For I = LBound(TableauEtiquette) To UBound(TableauEtiquette)

Select Case UCase(Champ)
    Case UCase("Ensemble")
        TableauEtiquette(I).Ensemble(1) = Valeur
    Case UCase("RefConnecteur")
        TableauEtiquette(I).Connecteur(1) = Valeur
        
'    Case UCase("RefClip")
'    NbFamille = Split("" & Valeur & "((((", "(")
'    NbFamille = Split(NbFamille(1) & ")))", ")")
'    If Val("" & NbFamille(0)) = 0 Then Valeur = ""
'        TableauEtiquette(I).Famille(1) = TableauEtiquette(I).Famille(1) & " " & Valeur
         
    Case UCase("Ref Joint")
        TableauEtiquette(I).Ref_Joint(1) = TableauEtiquette(I).Ref_Joint(1) & " " & Valeur
        
    Case UCase("PI")
        TableauEtiquette(I).PI(1) = Valeur
        
    Case UCase("Ensemble")
        If NoRepat = True Then
            TableauEtiquette(I).Ensemble(1) = Valeur
        Else
            TableauEtiquette(I).Ensemble(1) = TableauEtiquette(I).Ensemble(1) & " " & Valeur
        End If
     Case UCase("RefClip")
        TableauEtiquette(I).AlveRef(1) = TableauEtiquette(I).AlveRef(1) & " " & Valeur
        TableauEtiquette(I).Famille(1) = TableauEtiquette(I).Famille(1) & " " & Valeur
     
     Case UCase("Capot")
            If NoRepat = True Then
                TableauEtiquette(I).Capot(1) = Valeur
            Else
                TableauEtiquette(I).Capot(1) = TableauEtiquette(I).Capot(1) & " " & Valeur
            End If
     Case UCase("RefVerrou")
            If NoRepat = True Then
                TableauEtiquette(I).Verrou(1) = Valeur
            Else
                TableauEtiquette(I).Verrou(1) = TableauEtiquette(I).Verrou(1) & " " & Valeur
            End If
    Case UCase("RefBouchon")
        If NoRepat = True Then
            TableauEtiquette(I).Bouchon(1) = Valeur
        Else
            TableauEtiquette(I).Bouchon(1) = TableauEtiquette(I).Bouchon(1) & " " & Valeur
        End If
End Select
DoEvents
Next I

End Sub
Public Sub RenseigneChampApp(Champ As String, Valeur As String, tableau)
tableau2 = Split(tableau, Chr(13))

For I = 0 To UBound(tableau2) - 1
tableau3 = Split(tableau2(I), Chr(9))
Select Case UCase(Champ)
    Case "CONNECTEUR"
        TableauEtiquette(App(tableau3(0))).Connecteur(1) = Valeur
        
    Case UCase("Ensemble")
        TableauEtiquette(App(tableau3(0))).Connecteur(1) = Valeur
        
    Case UCase("PI")
        TableauEtiquette(App(tableau3(0))).Connecteur(1) = Valeur
        
    Case UCase("Famille")
        TableauEtiquette(App(tableau3(0))).Connecteur(1) = Valeur
             
    Case UCase("Ref Joint")
        TableauEtiquette(App(tableau3(0))).Connecteur(1) = Valeur
        
    Case UCase("Alvé Réf")
        TableauEtiquette(App(tableau3(0))).AlveRef(1) = Valeur
        
        
  
        
End Select
Next I

End Sub
Public Function TableMax() As Long
TableMax = UBound(TableauEtiquette)
End Function
Public Function TableMin() As Long
TableMin = LBound(TableauEtiquette)
End Function
Public Function RetournEtiquette(L As Long)
Dim Table(10, 1) As String
    Table(0, 0) = TableauEtiquette(L).Code_APP(0): Table(0, 1) = TableauEtiquette(L).Code_APP(1)
    Table(1, 0) = TableauEtiquette(L).Connecteur(0): Table(1, 1) = TableauEtiquette(L).Connecteur(1)
    Table(2, 0) = TableauEtiquette(L).DESIGNATION(0): Table(2, 1) = TableauEtiquette(L).DESIGNATION(1)
    Table(3, 0) = TableauEtiquette(L).Famille(0): Table(3, 1) = TableauEtiquette(L).Famille(1)
    Table(4, 0) = TableauEtiquette(L).Ref_Joint(0): Table(4, 1) = TableauEtiquette(L).Ref_Joint(1)
    Table(5, 0) = TableauEtiquette(L).Ensemble(0): Table(5, 1) = TableauEtiquette(L).Ensemble(1)
    Table(6, 0) = TableauEtiquette(L).PI(0): Table(6, 1) = TableauEtiquette(L).PI(1)
    Table(7, 0) = TableauEtiquette(L).AlveRef(0): Table(7, 1) = TableauEtiquette(L).AlveRef(1)
    Table(8, 0) = TableauEtiquette(L).Capot(0): Table(8, 1) = TableauEtiquette(L).Capot(1)
    Table(9, 0) = TableauEtiquette(L).Verrou(0): Table(9, 1) = TableauEtiquette(L).Verrou(1)
    Table(10, 0) = TableauEtiquette(L).Bouchon(0): Table(10, 1) = TableauEtiquette(L).Bouchon(1)
    
    Table(0, 1) = Replace(Table(0, 1), ";,", ";")
    Table(1, 1) = Replace(Table(1, 1), ";,", ";")
    Table(2, 1) = Replace(Table(2, 1), ";,", ";")
    Table(3, 1) = Replace(Table(3, 1), ";,", ";")
    Table(4, 1) = Replace(Table(4, 1), ";,", ";")
    Table(5, 1) = Replace(Table(5, 1), ";,", ";")
    Table(6, 1) = Replace(Table(6, 1), ";,", ";")
    Table(7, 1) = Replace(Table(7, 1), ";,", ";")
    Table(8, 1) = Replace(Table(8, 1), ";,", ";")
    Table(9, 1) = Replace(Table(9, 1), ";,", ";")
    Table(10, 1) = Replace(Table(10, 1), ";,", ";")
    
    Table(0, 1) = Replace(Table(0, 1), ", ;", ";")
    Table(1, 1) = Replace(Table(1, 1), ", ;", ";")
    Table(2, 1) = Replace(Table(2, 1), ", ;", ";")
    Table(3, 1) = Replace(Table(3, 1), ", ;", ";")
    Table(4, 1) = Replace(Table(4, 1), ", ;", ";")
    Table(5, 1) = Replace(Table(5, 1), ", ;", ";")
    Table(6, 1) = Replace(Table(6, 1), ", ;", ";")
    Table(7, 1) = Replace(Table(7, 1), ", ;", ";")
    Table(8, 1) = Replace(Table(8, 1), ", ;", ";")
    Table(9, 1) = Replace(Table(9, 1), ", ;", ";")
    Table(10, 1) = Replace(Table(10, 1), ", ;", ";")
    
  RetournEtiquette = Table
End Function
