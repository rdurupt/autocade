Attribute VB_Name = "VariableGlobal"
Option Explicit

Type MyLicGene
    Societe As String
    Tous As String
    AficheFrm As String
    DateDeb As String
    DateExecuter As String
    DateFin As String
    Enregistre As String
    NbJeton As String
    NbJetonActif As String
   
End Type
Type MyLic
    Serial As String
    PassWord As String
    Useur As String
    Enregistre As String
End Type
Type Licence
    Count As Long
    General As MyLicGene
    Record() As MyLic
   
End Type
Type MyDb
     UserDb As String
    PassWordDb As String
End Type
Public PassDb As MyDb
Public FiledLicence As Licence

Public PrixV As String

Public CodageX As New CDETXT

