Attribute VB_Name = "VariableGlobales"
Public PathConnecteurs As String
Public PathArchiveAutocad As String
Public PathBlocs As String
Public DonneesEntreprise As String
Public DonneesProduction As String
Public NmJob As Long
Public RepPlacheClous As String
Public PlanchClous As String
Public Msg
Public AutoApp As AcadApplication
Public InsertPointLigneTableau_fils(0 To 2) As Double
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
Public ConBaseNum As New Ado
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
Public strStatus As String
Public boolValideMOD As Boolean
Public PlanArchive As Boolean
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
    InsertPointLigneC(0 To 2) As Double
    InsertPointLigneV(0 To 2) As Double
    InsertPointLigneE(0 To 2) As Double
    XScaleFactorC As Double
    YScaleFactorC As Double
    ZScaleFactorC As Double
    XScaleFactorV As Double
    YScaleFactorV As Double
    ZScaleFactorV As Double
    RotationC As Double
    RotationV As Double
    PosOk As Boolean
End Type
Type T_TorS
    TorName As String
    NewBlockTorDetail  As AcadBlockReference
    TableauFile As String
     TorExiste As Boolean
    Insert(0 To 2) As Double
    XScaleFactor As Double
    YScaleFactor As Double
    ZScaleFactor As Double
    Rotation As Double
    PosOk As Boolean
    Garder As Boolean
End Type
Type T_Tor
    Kill As Boolean
    CodeApp As String
    NewBlockTorTire  As AcadBlockReference
    Tor() As T_TorS
    CollectionTor As New Collection
    Attribues As Collection
    NumTor As Long
    TorExiste As Boolean
    InsertTorTitre(0 To 2) As Double
    XScaleFactor As Double
    YScaleFactor As Double
    ZScaleFactor As Double
    Rotation As Double
    PosOk As Boolean
    Garder As Boolean
End Type
Type T_Comp
    Kill As Boolean
    NewBlock  As AcadBlockReference
    Attribues As Collection
    ComposantsExiste As Boolean
    InsertPointLigneC(0 To 2) As Double
    XScaleFactorC As Double
    YScaleFactorC As Double
    ZScaleFactorC As Double
    RotationC As Double
    PosOk As Boolean
End Type

Type T_Notas
    Kill As Boolean
    NewBlock  As AcadBlockReference
    Attribues As Collection
    NotasExiste As Boolean
    InsertPointLigneC(0 To 2) As Double
    XScaleFactorC As Double
    YScaleFactorC As Double
    ZScaleFactorC As Double
    RotationC As Double
    PosOk As Boolean
End Type



Type PointConnecteur
    InsertPointConnecteur(0 To 2) As Double
End Type
Public NUMCOM As Long
Public NUMNOTA As Long
Public NUMNTOR As Long
Public NUMNTORBLOC As Long
Public Const IndexIstC = 9
Public InsertPointConnecteur(0 To IndexIstC) As PointConnecteur
Public TableauDeComposants() As T_Comp
Public TableauDeConnecteurs() As T_Con
Public TableauDeNotas() As T_Notas
Public TableuDeTor() As T_Tor
Public LeCient As String
Public VarPreced As Boolean
Public boolCreationPlan As Boolean
Public CollectionCon As Collection
Public CollectionComp As Collection
Public CollectionNota As Collection
Public NbError As Long
Public FormBarGrah As Object
Public CollectionTor As Collection
