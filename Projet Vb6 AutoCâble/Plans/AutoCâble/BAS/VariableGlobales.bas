Attribute VB_Name = "VariableGlobales"
Public AutoApp As AcadApplication
Public InsertPointLigneTableau_fils(0 To 2) As Double
Public InsertPointConnecteur(0 To 2) As Double
Public InsertPointLigneTableau_Vignette(0 To 2) As Double
Public MyEcel As EXCEL.Application
Public TableauPath As New Collection
Public NbContolClient As Long
Public boolExec As Boolean
Dim NewBlock  As AcadBlockReference
Public Admin As Boolean
Public Lecture As Boolean
Public Ecriture As Boolean
Public Creation As Boolean
Public Loguer As Boolean
Public NoClose As Boolean
Public boolQuitte As Boolean
Public ActifDoc As String
Public DbNumPlan As String
Public db As String
Public varProjet As String
Public varIndice As String
Public Con As New Ado
Public ConNumPlan As New Ado
Public MyCARTOUCHE_Client
Public LeCartouche As String
Public LeCartoucheE As String
Public AttribuCartouche As New Collection
Public JobError As Long
Public Fichier As String
Public NbLignesVignette As Long
Public ErrInsert As Boolean
Public boolFormClient As Boolean
Public MenuShow As Boolean
Type T_Con
    Kill As Boolean
    NewBlock  As AcadBlockReference
    Attribues As Collection
    NewVignette  As AcadBlockReference
    AttribuesVignette As Collection
    EPISSURE As Boolean
    indexFile As Long
    TableauFile() As String
    AttribuesFils As Collection
    ConnecteurExiste As Boolean
End Type
Public TableauDeConnecteurs() As T_Con
Public LeCient As String
Public VarPreced As Boolean
Public boolCreationPlan As Boolean
Public CollectionCon As Collection

