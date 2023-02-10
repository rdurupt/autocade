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
Public Sub PrpareEtiqet(Tableau)
Set App = Nothing
Set App = New Collection
Tableau2 = Split(Tableau, Chr(13))
ReDim TableauEtiquette(0)
ReDim TableauEtiquette(UBound(Tableau2) - 1)
For i = 0 To UBound(Tableau2) - 1
Tableau3 = Split(Tableau2(i), Chr(9))
On Error Resume Next
a = ""
a = App(Tableau3(0))
If Err Then
Err.Clear
    App.Add i, Tableau3(0)

End If
TableauEtiquette(App(Tableau3(0))).Code_APP(0) = preFixFild & "Code_APP" & SuFixFild
TableauEtiquette(App(Tableau3(0))).Code_APP(1) = Tableau3(0)
TableauEtiquette(App(Tableau3(0))).DESIGNATION(0) = preFixFild & "DESIGNATION" & SuFixFild
TableauEtiquette(App(Tableau3(0))).DESIGNATION(1) = Tableau3(1)

TableauEtiquette(App(Tableau3(0))).Connecteur(0) = preFixFild & "CONNECTEUR" & SuFixFild
TableauEtiquette(App(Tableau3(0))).Ensemble(0) = preFixFild & UCase("Ensemble") & SuFixFild
TableauEtiquette(App(Tableau3(0))).Famille(0) = preFixFild & UCase("Famille") & SuFixFild
TableauEtiquette(App(Tableau3(0))).Pi(0) = preFixFild & UCase("PI") & SuFixFild
TableauEtiquette(App(Tableau3(0))).Ref_Joint(0) = preFixFild & UCase("Ref Joint") & SuFixFild
TableauEtiquette(App(Tableau3(0))).AlveRef(0) = preFixFild & UCase("Alvé Réf") & SuFixFild

Next
End Sub
Public Sub RenseigneChamp(Champ As String, Valeur As String)
For i = LBound(TableauEtiquette) To UBound(TableauEtiquette)
Select Case UCase(Champ)
    Case "CONNECTEUR"
        TableauEtiquette(i).Connecteur(1) = Valeur
        
    Case UCase("Famille")
        TableauEtiquette(i).Famille(1) = TableauEtiquette(i).Famille(1) & Valeur & ", "
         
    Case UCase("Ref Joint")
        TableauEtiquette(i).Ref_Joint(1) = TableauEtiquette(i).Ref_Joint(1) & Valeur & ","
        
    Case UCase("PI")
        TableauEtiquette(i).Pi(1) = Valeur
        
    Case UCase("Ensemble")
        TableauEtiquette(i).Ensemble(1) = TableauEtiquette(i).Ensemble(1) & Valeur
        
     Case UCase("Alvé Réf")
        TableauEtiquette(i).AlveRef(1) = TableauEtiquette(i).AlveRef(1) & Valeur & ","
End Select
Next i

End Sub
Public Sub RenseigneChampApp(Champ As String, Valeur As String, Tableau)
Tableau2 = Split(Tableau, Chr(13))

For i = 0 To UBound(Tableau2) - 1
Tableau3 = Split(Tableau2(i), Chr(9))
Select Case UCase(Champ)
    Case "CONNECTEUR"
        TableauEtiquette(App(Tableau3(0))).Connecteur(1) = Valeur
        
    Case UCase("Ensemble")
        TableauEtiquette(App(Tableau3(0))).Connecteur(1) = Valeur
        
    Case UCase("PI")
        TableauEtiquette(App(Tableau3(0))).Connecteur(1) = Valeur
        
    Case UCase("Famille")
        TableauEtiquette(App(Tableau3(0))).Connecteur(1) = Valeur
             
    Case UCase("Ref Joint")
        TableauEtiquette(App(Tableau3(0))).Connecteur(1) = Valeur
        
    Case UCase("Alvé Réf")
        TableauEtiquette(App(Tableau3(0))).AlveRef(1) = Valeur
        
  
        
End Select
Next i

End Sub
Public Function TableMax() As Long
TableMax = UBound(TableauEtiquette)
End Function
Public Function TableMin() As Long
TableMin = LBound(TableauEtiquette)
End Function
Public Function RetournEtiquette(L As Long)
Dim Table(7, 1) As String
    Table(0, 0) = TableauEtiquette(L).Code_APP(0): Table(0, 1) = TableauEtiquette(L).Code_APP(1)
    Table(1, 0) = TableauEtiquette(L).Connecteur(0): Table(1, 1) = TableauEtiquette(L).Connecteur(1)
    Table(2, 0) = TableauEtiquette(L).DESIGNATION(0): Table(2, 1) = TableauEtiquette(L).DESIGNATION(1)
    Table(3, 0) = TableauEtiquette(L).Famille(0): Table(3, 1) = TableauEtiquette(L).Famille(1)
    Table(4, 0) = TableauEtiquette(L).Ref_Joint(0): Table(4, 1) = TableauEtiquette(L).Ref_Joint(1)
    Table(5, 0) = TableauEtiquette(L).Ensemble(0): Table(5, 1) = TableauEtiquette(L).Ensemble(1)
    Table(6, 0) = TableauEtiquette(L).Pi(0): Table(6, 1) = TableauEtiquette(L).Pi(1)
    Table(7, 0) = TableauEtiquette(L).AlveRef(0): Table(7, 1) = TableauEtiquette(L).AlveRef(1)
    
    Table(0, 1) = Replace(Table(0, 1), ";,", ";")
    Table(1, 1) = Replace(Table(1, 1), ";,", ";")
    Table(2, 1) = Replace(Table(2, 1), ";,", ";")
    Table(3, 1) = Replace(Table(3, 1), ";,", ";")
    Table(4, 1) = Replace(Table(4, 1), ";,", ";")
    Table(5, 1) = Replace(Table(5, 1), ";,", ";")
    Table(6, 1) = Replace(Table(6, 1), ";,", ";")
    Table(7, 1) = Replace(Table(7, 1), ";,", ";")
    
    Table(0, 1) = Replace(Table(0, 1), ", ;", ";")
    Table(1, 1) = Replace(Table(1, 1), ", ;", ";")
    Table(2, 1) = Replace(Table(2, 1), ", ;", ";")
    Table(3, 1) = Replace(Table(3, 1), ", ;", ";")
    Table(4, 1) = Replace(Table(4, 1), ", ;", ";")
    Table(5, 1) = Replace(Table(5, 1), ", ;", ";")
    Table(6, 1) = Replace(Table(6, 1), ", ;", ";")
    Table(7, 1) = Replace(Table(7, 1), ", ;", ";")
    
  RetournEtiquette = Table
End Function
