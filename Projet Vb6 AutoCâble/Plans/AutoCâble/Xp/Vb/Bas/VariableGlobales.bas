Attribute VB_Name = "VariableGlobales"
Public MyPathXlsMoins1 As String
Public MyWord As Word.Application
Public MyWordDoc As Word.Document
Public bool_Plan_L_Connecteurs As Boolean
Public bool_Plan_L_Fils As Boolean
Public bool_Plan_L_Vignettes As Boolean
Public bool_Plan_L_Etiquettes As Boolean
Public bool_Plan_L_Composants As Boolean
Public bool_Plan_L_Notas As Boolean
Public bool_Plan_L_cartouches As Boolean
Public bool_Plan_L_Preconisations As Boolean
Public bool_Plan_L_Options As Boolean
Public bool_Plan_L_Criteres As Boolean
Public bool_Plan_L_Noeuds As Boolean
Public Bool_Fichier_Li As Boolean
Public PereFilsOk As Boolean

Public bool_Plan_E_Connecteurs As Boolean
Public bool_Plan_E_Fils As Boolean
Public bool_Plan_E_Vignettes As Boolean
Public bool_Plan_E_Etiquettes As Boolean
Public bool_Plan_E_Composants As Boolean
Public bool_Plan_E_Notas As Boolean
Public bool_Plan_E_cartouches As Boolean
Public bool_Plan_E_Preconisations As Boolean
Public bool_Plan_E_Options As Boolean
Public bool_Plan_E_Criteres As Boolean
Public bool_Plan_E_Noeuds As Boolean

Public bool_Outil_L_Connecteurs As Boolean
Public bool_Outil_L_Fils As Boolean
Public bool_Outil_L_Vignettes As Boolean
Public bool_Outil_L_Etiquettes As Boolean
Public bool_Outil_L_Composants As Boolean
Public bool_Outil_L_Notas As Boolean
Public bool_Outil_L_cartouches As Boolean
Public bool_Outil_L_Preconisations As Boolean
Public bool_Outil_L_Options As Boolean
Public bool_Outil_L_Criteres As Boolean
Public bool_Outil_L_Noeuds As Boolean
Public XlsPrix As String

Public bool_Outil_E_Connecteurs As Boolean
Public bool_Outil_E_Fils As Boolean
Public bool_Outil_E_Vignettes As Boolean
Public bool_Outil_E_Etiquettes As Boolean
Public bool_Outil_E_Composants As Boolean
Public bool_Outil_E_Notas As Boolean
Public bool_Outil_E_cartouches As Boolean
Public bool_Outil_E_Preconisations As Boolean
Public bool_Outil_E_Options As Boolean
Public bool_Outil_E_Criteres As Boolean
Public bool_Outil_E_Noeuds As Boolean

Public bool_Plan_Ouvrir As Boolean
Public bool_Outil_Ouvrir As Boolean
Public boolAutoCAD As Boolean
Public NbConnecteur As Long
Public NotSaveRacourci As Boolean
Public IdFils As Long
Public NbCartouche As Long
Public NomenclatureOk As Boolean
Public PathConnecteurs As String
Public PathArchiveAutocad As String
Public PathBlocs As String
Public DonneesEntreprise As String
Public DonneesProduction As String
Public NmJob As Long
Public RepPlacheClous As String
Public PlanchClous As String
Public Msg
Public AutoApp As Object  'AutoCAD.AcadApplication
Public InsertPointLigneTableau_fils(0 To 2) As Double
Public InsertPointLigneCritères(0 To 2) As Double
Public InsertPointLigneTableau_fils2(0 To 2) As Double
Public InsertPointLigneTableau_Vignette(0 To 2) As Double
Public MyEcel As EXCEL.Application
Public TableauPath As New Collection
Public NbContolClient As Long
Public boolExec As Boolean
Dim NewBlock  As AcadBlockReference
Public Admin As Boolean
Public Verifrificateur As Boolean
Public Approbateur As Boolean
Public Creation As Boolean
Public Loguer As Boolean
Public NoClose As Boolean
Public boolQuitte As Boolean
Public ActifDoc As String
Public DbNumPlan As String
Public BdDateTable As String
Public DbCatalogue As String
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
Public CollectionFils As Collection
Public CollectionEtiquettes As Collection
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
Type T_Composant
    Name As String
    BlockComp As AcadBlockReference
     InsertComp(0 To 2) As Double
    XScaleFactorComp As Double
    YScaleFactorComp As Double
    ZScaleFactorComp As Double
    RotationComp As Double
    PosOkComp As Boolean
    BlockDesing As AcadBlockReference
    InsertDesing(0 To 2) As Double
    XScaleFactorDesin As Double
    YScaleFactorDesin As Double
    ZScaleFactorDesin As Double
    RotationDesin As Double
    PosOkDesin As Boolean
  End Type
  Type T_Noeud
    Name As String
    BlockComp As AcadBlockReference
     InsertComp(0 To 2) As Double
    XScaleFactorComp As Double
    YScaleFactorComp As Double
    ZScaleFactorComp As Double
    RotationComp As Double
    PosOkComp As Boolean
    BlockDesing As AcadBlockReference
    InsertDesing(0 To 2) As Double
    XScaleFactorDesin As Double
    YScaleFactorDesin As Double
    ZScaleFactorDesin As Double
    RotationDesin As Double
    PosOkDesin As Boolean
    BlockFleche As AcadBlockReference
    InsertFleche(0 To 2) As Double
    XScaleFactorFleche  As Double
    YScaleFactorFleche As Double
    ZScaleFactorFleche As Double
    RotationFleche As Double
    PosOkFleche As Boolean
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
Public TableauComposant() As T_Composant
Public NUMCOM As Long
Public NUMNOTA As Long
Public NUMPRECO As Long

Public NUMNOEUDS As Long
Public NUMNETT As Long
Public NUMNTOR As Long
Public NUMNTORBLOC As Long
Public Const IndexIstC = 9
Public InsertPointConnecteur(0 To IndexIstC) As PointConnecteur
Public Const IndexIstN = 25
Public InsertNouds(0 To IndexIstN) As PointConnecteur
Public TableauDeComposants() As T_Comp
Public TableauDeConnecteurs() As T_Con
Public TableauDeNotas() As T_Notas
Public TableauDeNoeuds() As T_Noeud
Public TableuDeTor() As T_Tor
Public TableuEtiquettes() As T_Tor
Public LeCient As String
Public VarPreced As Boolean
Public boolCreationPlan As Boolean
Public CollectionCon As Collection
Public CollectionComp As Collection
Public CollectionNota As Collection
Public CollectionNoeuds As Collection
Public RefOption As Collection
Public RefCriteres As Collection
Public CollectionChartouche As Collection
Public NbError As Long
Public FormBarGrah As Object
Public CollectionTor As Collection
Public Type Etiqette
    Ensemble(1) As String
    Pi(1) As String
   Code_APP(1) As String
   DESIGNATION(1) As String
   Connecteur(1) As String
   Ref_Joint(1) As String
   Famille(1) As String
   AlveRef(1) As String
    
    
End Type
