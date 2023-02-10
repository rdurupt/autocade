Attribute VB_Name = "VariableGlobales"
Public ADO_TYPEBASE As Integer
Public ADO_BASE As String
Public ADO_SERVER As String
Public ADO_Fichier As String
Public ADO_User As String
Public ADO_PassWord As String
Public ChronoChario As Long
Public PossibleArretKill As Boolean
Public ArretKill As Boolean
Public SavIsClient As Boolean
Public IdUser As Long
Public PathAppliAutocad As String
Public PathAppliExcel As String
Public ComteurPass As Long
Global StarOnglet As Long
Global OfsetOnglet As Long
Global DocAutoCad As Object
Global StopOnglet As Long
Global GestionOnglets As Long
Global GestionOnglets2 As Long
Global PerfEntete As Boolean
Public ActionType As Long
Public ColecAplication As Collection
Public PreparNomOk As Boolean
Global BoolGenEtatEpisure As Boolean
'Global BooolBloque As Boolean
Global LstColecDoc As Collection
Global GenerateurDocOk As Boolean
Global RsBarGraph As Recordset
Global MyCollectionConnecteur As Collection
Global MyPathXlsMoins1 As String
Global MyWord As Object
Global MyWordDoc As Object
Global MyWordDoc2 As Object
Global bool_Plan_L_Connecteurs As Boolean
Global bool_Plan_L_Fils As Boolean
Global bool_Plan_L_Vignettes As Boolean
Global bool_Plan_L_Etiquettes As Boolean
Global bool_Plan_L_Composants As Boolean
Global bool_Plan_L_Notas As Boolean
Global bool_Plan_L_cartouches As Boolean
Global bool_Plan_L_Preconisations As Boolean
Global bool_Plan_L_Options As Boolean
Global bool_Plan_L_Criteres As Boolean
Global bool_Plan_L_Noeuds As Boolean
Global Bool_Fichier_Li As Boolean
Global PereFilsOk As Boolean
Global ClasseurXls As String
Global Cell(4) As Long
Global Id_Users As Long
Global IndexTableauDocGen As Long
Global IsCilent As Boolean
Global IsServeur As Boolean
Global bool_Plan_E_Connecteurs As Boolean
Global bool_Plan_E_Fils As Boolean
Global bool_Plan_E_Vignettes As Boolean
Global bool_Plan_E_Etiquettes As Boolean
Global bool_Plan_E_Composants As Boolean
Global bool_Plan_E_Notas As Boolean
Global bool_Plan_E_cartouches As Boolean
Global bool_Plan_E_Preconisations As Boolean
Global bool_Plan_E_Options As Boolean
Global bool_Plan_E_Criteres As Boolean
Global bool_Plan_E_Noeuds As Boolean
Global bool_MiseEnPage As Boolean
Global bool_Outil_L_Connecteurs As Boolean
Global bool_Outil_L_Fils As Boolean
Global bool_Outil_L_Vignettes As Boolean
Global bool_Outil_L_Etiquettes As Boolean
Global bool_Outil_L_Composants As Boolean
Global bool_Outil_L_Notas As Boolean
Global bool_Outil_L_cartouches As Boolean
Global bool_Outil_L_Preconisations As Boolean
Global bool_Outil_L_Options As Boolean
Global bool_Outil_L_Criteres As Boolean
Global bool_Outil_L_Noeuds As Boolean
Global XlsPrix As String
Global EnteteClasseurControle As String

Global bool_Outil_E_Connecteurs As Boolean
Global bool_Outil_E_Fils As Boolean
Global bool_Outil_E_Vignettes As Boolean
Global bool_Outil_E_Etiquettes As Boolean
Global bool_Outil_E_Composants As Boolean
Global bool_Outil_E_Notas As Boolean
Global bool_Outil_E_cartouches As Boolean
Global bool_Outil_E_Preconisations As Boolean
Global bool_Outil_E_Options As Boolean
Global bool_Outil_E_Criteres As Boolean
Global bool_Outil_E_Noeuds As Boolean
Global AdcFileName As String
Global bool_Plan_Ouvrir As Boolean
Global bool_Outil_Ouvrir As Boolean
Global boolAutoCAD As Boolean
Global NbConnecteur As Long
Global NotSaveRacourci As Boolean
Global IdFils As Long

Global NbCartouche As Long
Global NomenclatureOk As Boolean
Global PathConnecteurs As String
Global PathArchiveAutocad As String
Global PathBlocs As String
Global DonneesEntreprise As String
Global DonneesProduction As String
Global NmJob As Long
Global RepPlacheClous As String
Global PlanchClous As String
Global msg
Global AutoApp As Object  'AutoCAD.AcadApplication
Global InsertPointLigneTableau_fils(0 To 2) As Double
Global InsertPointLigneCritères(0 To 2) As Double
Global InsertPointLigneTableau_fils2(0 To 2) As Double
Global InsertPointLigneTableau_fils3(0 To 2) As Double
Global InsertPointLigneTableau_Vignette(0 To 2) As Double
Global MyExcel As EXCEL.Application
Global TableauPath As New Collection
Global NbContolClient As Long
Global boolExec As Boolean
Dim NewBlock  As Object
Global Admin As Boolean
Global Verifrificateur As Boolean
Global Approbateur As Boolean
Global Creation As Boolean
Global Loguer As Boolean
Global NoClose As Boolean
Global boolQuitte As Boolean
Global ActifDoc As String
Global DbNumPlan As String
Global BdDateTable As String
Global DbCatalogue As String
Global Db As String
Global varProjet As String
Global GeneEtatMacro As String

Global varIndice As String
Global Con As New Ado
Global ConBaseNum As New Ado
Global MyCARTOUCHE_Client
Global LeCartouche As String
Global LeCartoucheE As String
Global AttribuCartouche As New Collection
Global JobError As Long
Global Fichier As String
Global NbLignesVignette As Long
Global ErrInsert As Boolean
Global boolFormClient As Boolean
Global MenuShow As Boolean
Global strStatus As String
Global boolValideMOD As Boolean
Global PlanArchive As Boolean
Global CollectionFils As Collection
Global CollectionEtiquettes As Collection
Global MyWorkbookTravail As Workbook
Global MyWorkbookOnglet As Workbook
Global MyWorkbookAppli As Workbook
Global MyTableau() As String
Global MyTableaul() As String
Global AutocableDRIVE As String
Type T_Con
    Kill As Boolean
    NewBlock  As Object
    Attribues As Collection
    NewVignette  As Object
    AttribuesVignette As Collection
    Epissure As Boolean
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
    NewBlockTorDetail  As Object
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
    BlockComp As Object
     InsertComp(0 To 2) As Double
    XScaleFactorComp As Double
    YScaleFactorComp As Double
    ZScaleFactorComp As Double
    RotationComp As Double
    PosOkComp As Boolean
    BlockDesing As Object
    InsertDesing(0 To 2) As Double
    XScaleFactorDesin As Double
    YScaleFactorDesin As Double
    ZScaleFactorDesin As Double
    RotationDesin As Double
    PosOkDesin As Boolean
  End Type
  Type T_Noeud
    Name As String
    BlockComp As Object
     InsertComp(0 To 2) As Double
    XScaleFactorComp As Double
    YScaleFactorComp As Double
    ZScaleFactorComp As Double
    RotationComp As Double
    PosOkComp As Boolean
    BlockDesing As Object
    InsertDesing(0 To 2) As Double
    XScaleFactorDesin As Double
    YScaleFactorDesin As Double
    ZScaleFactorDesin As Double
    RotationDesin As Double
    PosOkDesin As Boolean
    BlockFleche As Object
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
    NewBlockTorTire  As Object
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
    NewBlock  As Object
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
    NewBlock  As Object
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
Global TableauComposant() As T_Composant
Global NUMCOM As Long
Global NUMNOTA As Long
Global NUMPRECO As Long

Global NUMNOEUDS As Long
Global NUMNETT As Long
Global NUMNTOR As Long
Global NUMNTORBLOC As Long
Global Const IndexIstC = 9
Global InsertPointConnecteur(0 To IndexIstC) As PointConnecteur
Global Const IndexIstN = 25
Global InsertNouds(0 To IndexIstN) As PointConnecteur
Global TableauDeComposants() As T_Comp
Global TableauDeConnecteurs() As T_Con
Global TableauDeNotas() As T_Notas
Global TableauDeNoeuds() As T_Noeud
Global TableuDeTor() As T_Tor
Global TableuEtiquettes() As T_Tor
Global LeCient As String
Global VarPreced As Boolean
Global boolCreationPlan As Boolean
Global CollectionCon As Collection
Global CollectionComp As Collection
Global CollectionNota As Collection
Global CollectionNoeuds As Collection
Global RefOption As Collection
Global RefCriteres As Collection
Global RefAcCorrective As Collection
Global CollectionChartouche As Collection
Global NbError As Long
Global FormBarGrah As Object
Global CollectionTor As Collection
Global FichierErr As String
Type Etiqette
    Ensemble(1) As String
    PI(1) As String
   Code_APP(1) As String
   DESIGNATION(1) As String
   Connecteur(1) As String
   Ref_Joint(1) As String
   Famille(1) As String
   AlveRef(1) As String
   Capot(1) As String
   Verrou(1) As String
   Bouchon(1) As String
End Type
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
    UserName As String
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
Global PassDb As MyDb

Global FiledLicence As Licence

Public Const MainTitle = "AutoCâble éditeur"
Global CodageX As New CDETXT
Global UserName As String
Global Machine As String
Public PortraitPaysage As Long
Public Const NORMAL_PRIORITY_CLASS = &H20&
Public Const REALTIME_PRIORITY_CLASS = &H100&
Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
Public Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Public Declare Function CreateProcessWithLogon Lib "advapi32" Alias "CreateProcessWithLogonW" (ByVal lpUsername As Long, ByVal lpDomain As Long, ByVal lpPassword As Long, ByVal dwLogonFlags As Long, ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInfo As PROCESS_INFORMATION) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

