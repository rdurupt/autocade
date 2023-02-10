VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsNomanclatureGenerer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 Private Type Voie
    IndexLongueur As Long
    Liaison As String
    Voie As String
    Section As Double
    Famille As String
    Clip As String
    FourClip As String
    Joint As String
    FourJoint As String
    Coss As String
    FourCoss As String
    LONG As Double
    LONG_CP As Double
    Coupe As Double
    Long_Add As Double
    SECT As Double
    TEINT As String
    TEINT2 As String
    ISO As String
    
   End Type
Private Type Clip
    Famille As String
    Clip As String
    Min As Double
    Max As Double
End Type
Private Type Connecteur
    Bouchon As String
    FourBouchon As String
    Capot As String
    FourCapot As String
    Verrou As String
    FourVerrou As String
End Type
Private Type Fils
    Section As Double
    DIAMETRE As Double
End Type
Private Type Options
    DESIGNATION As String
    LstVoie As New Collection
    Voies() As Voie
End Type
Public Connect2
Private Connect As Connecteur
Private TableClips() As Clip
Private ClipA As Clip
Private ClipZ As Clip
Public Prix As Double
Public RefConnecteur As String
Public FourConnecteur As String
Public App As String
Public DESIGNATION As String
Public IdOption As Long
Public NameOption As New Collection
Private VoieOption() As Options

Private initClipOk As Boolean
Public NewRsUpdate As New Collection
Public Sub RenseigneLongeur(L As Double, C As Double, L_CP As Double, L_add As Double, s As Double, ISO As String, Couleur As String, Liseret As String, Voie As String, Critaire As String)
On Error Resume Next
VoieOption(NameOption(Critaire)).Voies(VoieOption(NameOption(Critaire)).LstVoie(Voie)).LONG = L
VoieOption(NameOption(Critaire)).Voies(VoieOption(NameOption(Critaire)).LstVoie(Voie)).LONG_CP = L_CP
VoieOption(NameOption(Critaire)).Voies(VoieOption(NameOption(Critaire)).LstVoie(Voie)).Coupe = C
VoieOption(NameOption(Critaire)).Voies(VoieOption(NameOption(Critaire)).LstVoie(Voie)).Long_Add = L_add
VoieOption(NameOption(Critaire)).Voies(VoieOption(NameOption(Critaire)).LstVoie(Voie)).SECT = s
VoieOption(NameOption(Critaire)).Voies(VoieOption(NameOption(Critaire)).LstVoie(Voie)).ISO = ISO
VoieOption(NameOption(Critaire)).Voies(VoieOption(NameOption(Critaire)).LstVoie(Voie)).TEINT = Couleur
VoieOption(NameOption(Critaire)).Voies(VoieOption(NameOption(Critaire)).LstVoie(Voie)).TEINT2 = Liseret
On Error GoTo 0
End Sub
Public Function IniJoint(RsFils As Recordset)
Dim Sql As String
Dim Rs As Recordset
Dim DIAMETRE  As Double
Dim DiamtrFil() As Fils
Dim a As Fils
Dim Z As Fils
Static Nb120Ac As Long
Dim MyVoi As String
Dim I As Long
Dim I2 As Long
Dim lst
'RsFils
On Error Resume Next
'If App = "120.AC" Then
'MsgBox "120.AC"
'End If
If "" & RsFils!App = App Then
Err.Clear

On Error GoTo FinG
MyVoi = "" & RsFils!VOI
Reprise1:
'If Trim(Replace(Replace("" & RsFils![Ref Clip], Chr(10), ""), Chr(13), "")) <> "" Then _

    VoieOption(NameOption("" & RsFils!Option)).Voies(VoieOption(NameOption("" & RsFils!Option)).LstVoie("" & RsFils!VOI.Value)).Clip = Replace(Replace("" & RsFils![Ref Clip], Chr(10), ""), Chr(13), "")

'If Trim(Replace(Replace("" & RsFils![Ref Clip Four], Chr(10), ""), Chr(13), "")) <> "" Then _

    VoieOption(NameOption("" & RsFils!Option)).Voies(VoieOption(NameOption("" & RsFils!Option)).LstVoie("" & RsFils!VOI)).FourClip = Replace(Replace("" & RsFils![Ref Clip Four], Chr(10), ""), Chr(13), "")
    
'If Trim(Replace(Replace("" & RsFils![Ref Joint], Chr(10), ""), Chr(13), "")) <> "" Then _

    VoieOption(NameOption("" & RsFils!Option)).Voies(VoieOption(NameOption("" & RsFils!Option)).LstVoie("" & RsFils!VOI)).FourJoint = Replace(Replace("" & RsFils![Ref Joint Four], Chr(10), ""), Chr(13), "")
    
'If Trim(Replace(Replace("" & RsFils![Ref Joint Four], Chr(10), ""), Chr(13), "")) <> "" Then _

    VoieOption(NameOption("" & RsFils!Option)).Voies(VoieOption(NameOption("" & RsFils!Option)).LstVoie("" & RsFils!VOI)).Joint = Replace(Replace("" & RsFils![Ref Joint], Chr(10), ""), Chr(13), "")
    
Else
On Error GoTo 0
    On Error GoTo FinG
    
    MyVoi = "" & RsFils!VOI2
'    If VoieOption(NameOption("" & RsFils!Option)).Voies(VoieOption(NameOption("" & RsFils!Option)).LstVoie("" & RsFils!VOI2)).Clip = "" Then
'    End If
Reprise2:
     VoieOption(NameOption("" & RsFils!Option)).Voies(VoieOption(NameOption("" & RsFils!Option)).LstVoie("" & RsFils!VOI2.Value)).Clip = "" & RsFils![Ref Clip2]
    VoieOption(NameOption("" & RsFils!Option)).Voies(VoieOption(NameOption("" & RsFils!Option)).LstVoie("" & RsFils!VOI2)).FourClip = "" & RsFils![Ref Clip Four2]
    VoieOption(NameOption("" & RsFils!Option)).Voies(VoieOption(NameOption("" & RsFils!Option)).LstVoie("" & RsFils!VOI2)).FourJoint = "" & Replace(Replace("" & RsFils![Ref Joint Four2], Chr(10), ""), Chr(13), "")
    VoieOption(NameOption("" & RsFils!Option)).Voies(VoieOption(NameOption("" & RsFils!Option)).LstVoie("" & RsFils!VOI2)).Joint = "" & Replace(Replace("" & RsFils![Ref Joint2], Chr(10), ""), Chr(13), "")
End If
GoTo FinG
Fin:
Err.Clear
'VoieOption (NameOption("" & RsFils!Option))
AjouterLstVoie "" & RsFils!VOI, "" & RsFils!Option
'ReDim Preserve VoieOption(UBound(VoieOption))
'' Lst = VoieOption(UBound(VoieOption)).LstVoie
'VoieOption(UBound(VoieOption)).LstVoie.Add VoieOption(UBound(VoieOption)).LstVoie.Count, "" & RsFils!VOI.Value
GoTo Reprise1
Resume Next

'    If Left(UCase(App), 1) <> "E" Then MsgBox App & " Voie : " & MyVoi


Fin1:
Err.Clear
AjouterLstVoie "" & RsFils!VOI2, "" & RsFils!Option

'ReDim Preserve VoieOption(UBound(VoieOption))
'' Lst = VoieOption(UBound(VoieOption)).LstVoie
'VoieOption(UBound(VoieOption)).LstVoie.Add VoieOption(UBound(VoieOption)).LstVoie.Count, "" & RsFils!VOI2.Value
GoTo Reprise2


FinG:
Err.Clear
On Error GoTo 0
End Function
Public Function InitConnecteur(BouchonRf As String, BouchonRfFour As String, Capot As String, _
                                CapotFour As String, Verrou As String, VerrouFour)
Dim Sql As String
Dim Rs As Recordset
Dim DIAMETRE  As Double
Dim DiamtrFil() As Fils
Dim a As Fils
Dim Z As Fils
Dim I As Long
Dim I2 As Long
'RsFils
On Error Resume Next
Connect.Bouchon = BouchonRf
Connect.FourBouchon = BouchonRfFour
Connect.Capot = Capot
Connect.FourCapot = CapotFour
Connect.Verrou = Verrou
Connect.FourVerrou = VerrouFour

On Error GoTo 0
End Function

Public Function RetourneRecordset(RM As ReyRecordsetMaker, rr) As ADODB.Recordset
Dim a As Long
 Dim I As Long
 Dim I_Option As Long
 Dim I_Voie As Long
 Dim indexFild As Long
 Dim Irs As Long
 Dim aa As FildName
' Set RM.rs = rs
'    RM.AddField "Nom", FT_VarChar, 30
'    RM.AddField "Naiss", FT_VarChar, 10
On Error GoTo Fin1
'If App = "120.AC" Then
'MsgBox "rs"
'End If
    For I = 0 To UBound(VoieOption)
        
'      Set aa = New FildName
'      aa.retournTableur rr
'        NewRsUpdate.Add aa
'        Set aa = Nothing
    If Err Then
        Err.Clear
'        RM.AddRecord
Set aa = New FildName
      aa.retournTableur rr
        indexFild = 0
'        RM.Fill "" & rr(indexFild, 0), App
        aa.Remplie App, indexFild
'        rr(indexFild, 4) = App
        indexFild = indexFild + 1
'        RM.Fill "" & rr(indexFild, 0), DESIGNATION
        indexFild = indexFild + 1
'        RM.Fill "" & rr(indexFild, 0), RefConnecteur
        rr(indexFild, 4) = RefConnecteur
        indexFild = indexFild + 1
'        RM.Fill "" & rr(indexFild, 0), "" & FourConnecteur
        indexFild = 8
'        RM.Fill "" & rr(indexFild, 0), "" & Connect.Bouchon
                indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & Connect.Capot
                 indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & Connect.FourCapot
                
                indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & Connect.Verrou
                indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & Connect.FourVerrou
'        RM.AddRecord
           NewRsUpdate.Add aa
        Set aa = Nothing
        Exit For
    End If
   
        For I_Option = 0 To UBound(VoieOption)
            For I_Voie = 0 To UBound(VoieOption(I_Option).Voies)
               If VoieOption(I_Option).Voies(I_Voie).Liaison <> "" Then
                Set aa = New FildName
      aa.retournTableur rr
                indexFild = 0
'                RM.Fill "" & rr(indexFild, 0), App
                 aa.Remplie "" & App, indexFild
                indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), DESIGNATION
                aa.Remplie "" & DESIGNATION, indexFild
                indexFild = indexFild + 1
'                 RM.Fill "" & rr(indexFild, 0), RefConnecteur
                 aa.Remplie "" & RefConnecteur, indexFild
                indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & FourConnecteur
                aa.Remplie "" & FourConnecteur, indexFild
                indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & VoieOption(I_Option).Voies(I_Voie).Liaison
                 aa.Remplie "" & VoieOption(I_Option).Voies(I_Voie).Liaison, indexFild
                indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & VoieOption(I_Option).Voies(I_Voie).Voie
                aa.Remplie "" & VoieOption(I_Option).Voies(I_Voie).Voie, indexFild
                 indexFild = indexFild + 1
'                 VoieOption(I_Option).Voies(I_Voie).LongueurF_G = 1.1
'                 VoieOption(I_Option).Voies(I_Voie).LongueurF_C = 2.2
'                 VoieOption(I_Option).Voies(I_Voie).LongueurF_D = 3.3
'                 VoieOption(I_Option).Voies(I_Voie).LongueurF_Choix = "C"
'                 RM.Fill "" & rr(indexFild, 0), Val(Replace("" & VoieOption(I_Option).Voies(I_Voie).Long_Add, ",", ".")), True
                 aa.Remplie "" & Replace("" & VoieOption(I_Option).Voies(I_Voie).Long_Add, ",", "."), indexFild
'                 indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & VoieOption(I_Option).Voies(I_Voie).LongueurF_C
'                 indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & VoieOption(I_Option).Voies(I_Voie).LongueurF_D
'                 indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & VoieOption(I_Option).Voies(I_Voie).LongueurF_Choix
                indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), Replace("" & VoieOption(I_Option).Voies(I_Voie).SECT, ",", "."), True
                aa.Remplie "" & Replace("" & VoieOption(I_Option).Voies(I_Voie).SECT, ",", "."), indexFild
                 indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), Val(Replace("" & VoieOption(I_Option).Voies(I_Voie).LONG, ",", ".")), True
                aa.Remplie "" & Replace("" & VoieOption(I_Option).Voies(I_Voie).SECT, ",", "."), indexFild
                 indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), Val(Replace("" & VoieOption(I_Option).Voies(I_Voie).Coupe, ",", ".")), True
                 aa.Remplie "" & Replace("" & VoieOption(I_Option).Voies(I_Voie).Coupe, ",", "."), indexFild
                 indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), Val(Replace("" & VoieOption(I_Option).Voies(I_Voie).LONG_CP, ",", ".")), True
                aa.Remplie "" & Replace("" & VoieOption(I_Option).Voies(I_Voie).LONG_CP, ",", "."), indexFild
                 indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & VoieOption(I_Option).Voies(I_Voie).TEINT
                aa.Remplie "" & VoieOption(I_Option).Voies(I_Voie).TEINT, indexFild
                 indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & VoieOption(I_Option).Voies(I_Voie).TEINT2
                aa.Remplie "" & VoieOption(I_Option).Voies(I_Voie).TEINT2, indexFild
                 indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & VoieOption(I_Option).Voies(I_Voie).Famille
                aa.Remplie "" & VoieOption(I_Option).Voies(I_Voie).Famille, indexFild
                indexFild = indexFild + 1
                
'                RM.Fill "" & rr(indexFild, 0), "" & Connect.Bouchon
                aa.Remplie "" & Connect.Bouchon, indexFild
                indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & Connect.FourBouchon
                aa.Remplie "" & Connect.FourBouchon, indexFild
                 indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & Connect.Capot
                aa.Remplie "" & Connect.Capot, indexFild
                 indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & Connect.FourCapot
                aa.Remplie "" & Connect.FourCapot, indexFild
                
                indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & Connect.Verrou
                 aa.Remplie "" & Connect.Verrou, indexFild
                indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & Connect.FourVerrou
                aa.Remplie "" & Connect.FourVerrou, indexFild
                indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & VoieOption(I_Option).DESIGNATION
                    aa.Remplie "" & VoieOption(I_Option).DESIGNATION, indexFild
                indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & VoieOption(I_Option).Voies(I_Voie).Clip
                aa.Remplie "" & VoieOption(I_Option).Voies(I_Voie).Clip, indexFild
                indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & VoieOption(I_Option).Voies(I_Voie).FourClip
                aa.Remplie "" & VoieOption(I_Option).Voies(I_Voie).FourClip, indexFild
                indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & VoieOption(I_Option).Voies(I_Voie).Joint
                aa.Remplie "" & VoieOption(I_Option).Voies(I_Voie).Joint, indexFild
                indexFild = indexFild + 1
'                RM.Fill "" & rr(indexFild, 0), "" & VoieOption(I_Option).Voies(I_Voie).FourJoint
                aa.Remplie "" & VoieOption(I_Option).Voies(I_Voie).FourJoint, indexFild
'               RM.AddRecord
                NewRsUpdate.Add aa
        Set aa = Nothing
               End If
           Next
        Next
'         RM.AddRecord
    Next
    GoTo Fin
Fin1:
Resume Next
Fin:
   Err.Clear
     On Error GoTo 0
   
'    Set RetourneRecordset = RM.Recordset
    ' Une fois le recordset recupere , celui du constructeur est detruit !!!!)
    
'    Rs.MoveFirst
'    While Not Rs.EOF
'        For a = 0 To Rs.Fields.Count - 1
'            MsgBox Rs.Fields(a).Name & " = " & Rs.Fields(a).Value, vbOKOnly, Rs.AbsolutePosition
'        Next a
'        Rs.MoveNext
'    Wend
End Function

Public Function IntiConnecteur(ConnecteurName As String, Verrou As String, FourVerrou As String, Capot As String) As String
        RefConnecteur = ConnecteurName
        Connect.Verrou = Verrou
        Connect.FourVerrou = "" & FourVerrou
        Connect.Capot = Capot
'        Connect.FourCapot=
'        If Verrou <> "" Then MsgBox ""
'        If FourVerrou <> "" Then MsgBox ""
'            If Capot <> "" Then MsgBox ""
End Function
Public Sub DelBouchon()
Dim I As Long
'For I = 0 To LstVoie.Count - 1
''    If Voies(I).Section <> 0 Then Voies(I).Bouchon = ""
'Next
End Sub
Public Sub SubSection(Voie As String, Section As Double)
'On Error Resume Next
'    Voies(LstVoie(Voie)).Section = Voies(LstVoie(Voie)).Section + Section
'Err.Clear
   
End Sub
Public Sub AjouterCritaire(Critaire As String)
Dim Critaires As Long

On Error Resume Next
Critaires = NameOption(Critaire)
If Err Then
Err.Clear
On Error GoTo Fin
    NameOption.Add NameOption.Count, Critaire
    ReDim Preserve VoieOption(NameOption(Critaire))
    VoieOption(NameOption(Critaire)).DESIGNATION = Critaire
End If
Fin:
Err.Clear
On Error GoTo 0
End Sub

Public Sub AjouterLstVoie(Voie As String, Critaire As String)
On Error GoTo Fin
'VoieOption(NameOption(Critaire)).LstVoie
    VoieOption(NameOption(Critaire)).LstVoie.Add VoieOption(NameOption(Critaire)).LstVoie.Count, Voie
    ReDim Preserve VoieOption(NameOption(Critaire)).Voies(VoieOption(NameOption(Critaire)).LstVoie(Voie))
    VoieOption(NameOption(Critaire)).Voies(VoieOption(NameOption(Critaire)).LstVoie(Voie)).Voie = Voie
Fin:
Err.Clear
On Error GoTo 0
End Sub
Public Sub initLiaison(Voie As String, Critaire As String, Liaison As String)
'On Error Resume Next
On Error GoTo Fin
If VoieOption(NameOption(Critaire)).Voies(VoieOption(NameOption(Critaire)).LstVoie(Voie)).Liaison = "" Then
End If

    VoieOption(NameOption(Critaire)).Voies(VoieOption(NameOption(Critaire)).LstVoie(Voie)).Liaison = Liaison
Fin:
Err.Clear
On Error GoTo 0
End Sub

Private Sub Class_Initialize()
'    Set LstVoie = New Collection
End Sub
'Public Sub ChoixClip()
'Dim I As Long
'Dim IndexClip As Long
'Dim Sql As String
'Dim RS As Recordset
'
'For I = 0 To LstVoie.Count - 1
'    Sql = "SELECT DISTINCT T_Lien_Con_Famille_Voies_Clip.Référence, "
'    Sql = Sql & "T_Lien_Con_Famille_Voies_Clip.Famille, T_Lien_Con_Famille_Voies.Voie,  "
'    Sql = Sql & "T_Lien_Con_Famille_Voies_Clip.Min, T_Lien_Con_Famille_Voies_Clip.Max "
'    Sql = Sql & "FROM T_Lien_Con_Famille INNER  "
'    Sql = Sql & "JOIN (T_Lien_Con_Famille_Voies INNER JOIN T_Lien_Con_Famille_Voies_Clip  "
'    Sql = Sql & "ON T_Lien_Con_Famille_Voies.Id =  "
'    Sql = Sql & "T_Lien_Con_Famille_Voies_Clip.Id_T_Lien_Con_Famille_Voies)  "
'    Sql = Sql & "ON T_Lien_Con_Famille.Id = T_Lien_Con_Famille_Voies.Id_T_Lien_Con_Famille "
'    Sql = Sql & "Where T_Lien_Con_Famille.Connecteur = '" & Connecteur & "'  "
'    Sql = Sql & "And T_Lien_Con_Famille_Voies.Voie = '" & Voies(I).Voie & "' "
'    Sql = Sql & "And T_Lien_Con_Famille_Voies_Clip.Min <= " & Replace(Voies(I).Section, ",", ".") & "  "
'    Sql = Sql & "And T_Lien_Con_Famille_Voies_Clip.Max >= " & Replace(Voies(I).Section, ",", ".") & " "
'    Sql = Sql & "ORDER BY T_Lien_Con_Famille_Voies_Clip.Min, T_Lien_Con_Famille_Voies_Clip.Max;"
'    Set RS = Con.OpenRecordSet(Sql)
'    If RS.EOF = False Then
'        While RS.EOF = False
'            Voies(I).Famille = Voies(I).Famille & Chr(10) & "" & RS!Famille
'            Voies(I).Clip = Voies(I).Clip & Chr(10) & RS!Famille & " : " & RS!Référence
'            RS.MoveNext
'        Wend
'    Else
'      For IndexClip = 1 To UBound(TableClips)
'        If Voies(I).Section >= TableClips(IndexClip).Min And Voies(I).Section <= TableClips(IndexClip).Max Then
'            Voies(I).Famille = Voies(I).Famille & Chr(10) & TableClips(IndexClip).Famille
'            Voies(I).Clip = Voies(I).Clip & Chr(10) & TableClips(IndexClip).Clip
'
'        End If
'      Next
'   End If
'   Set RS = Con.CloseRecordSet(RS)
'Next
'End Sub
Public Sub InitCip()
Dim Sql As String
Dim Rs As Recordset
Dim I As Long
If initClipOk = True Then Exit Sub
initClipOk = True
Sql = "SELECT T_Lien_Connecteur_Clip.RefConnecteur, MyFrom.txt1, MyFrom.[Famille Lib], MyFrom.[Alvé Réf Fourr],  "
Sql = Sql & "[Alvéole Mini en mm2],  "
Sql = Sql & "[Alvéole Maxi en mm2] ,  "
Sql = Sql & "MyFrom.Qté, MyFrom.[Prix U], MyFrom.[Prix Total] "
Sql = Sql & "FROM T_Lien_Connecteur_Clip INNER JOIN (SELECT DISTINCT con_contacts.txt1, lst21.CatName AS  "
Sql = Sql & "[Famille Lib], con_contacts.txt3 AS [Alvé Réf Fourr], lst22.CatName AS [Alvéole Mini en mm2],  "
Sql = Sql & "lst23.CatName AS [Alvéole Maxi en mm2], 0 AS Qté, 0 AS [Prix U], 0 AS [Prix Total] "
Sql = Sql & "FROM ((con_contacts LEFT JOIN lst21 ON con_contacts.lst21 = lst21.CatID) LEFT JOIN lst22  "
Sql = Sql & "ON con_contacts.lst22 = lst22.CatID) LEFT JOIN lst23 ON con_contacts.lst23 = lst23.CatID IN '"
Sql = Sql & TableauPath("Eb_CONNECTIQUE")
Sql = Sql & "') AS MyFrom ON T_Lien_Connecteur_Clip.Refclip = MyFrom.txt1 "
Sql = Sql & "where  T_Lien_Connecteur_Clip.RefConnecteur='" & RefConnecteur & "';"
Set Rs = Con.OpenRecordSet(Sql)
ReDim TableClips(I)
While Rs.EOF = False
 I = I + 1
  ReDim Preserve TableClips(I)
    TableClips(I).Famille = "" & Rs![Famille Lib]
    TableClips(I).Clip = "" & Rs!txt1
    TableClips(I).Min = Val(Replace(Trim(Replace("" & Rs![Alvéole Mini en mm2], "mm2", "")), ",", "."))
    TableClips(I).Max = Val(Replace(Trim(Replace("" & Rs![Alvéole Maxi en mm2], "mm2", "")), ",", "."))
   
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
For I = 1 To UBound(TableClips) - 1
    If TableClips(I).Min > TableClips(I + 1).Min Then
        ClipZ.Clip = TableClips(I).Clip
        ClipZ.Famille = TableClips(I).Famille
        ClipZ.Max = TableClips(I).Max
        ClipZ.Min = TableClips(I).Min
        
        ClipA.Clip = TableClips(I + 1).Clip
        ClipA.Famille = TableClips(I + 1).Famille
        ClipA.Max = TableClips(I + 1).Max
        ClipA.Min = TableClips(I + 1).Min
        
        TableClips(I).Clip = ClipA.Clip
        TableClips(I).Famille = ClipA.Famille
        TableClips(I).Max = ClipA.Max
        TableClips(I).Min = ClipA.Min
        
        TableClips(I + 1).Clip = ClipZ.Clip
        TableClips(I + 1).Famille = ClipZ.Famille
        TableClips(I + 1).Max = ClipZ.Max
        TableClips(I + 1).Min = ClipZ.Min
        I = I - 2
    Else
        If TableClips(I).Min = TableClips(I + 1).Min Then
            If TableClips(I).Max > TableClips(I + 1).Max Then
                 ClipZ.Clip = TableClips(I).Clip
                ClipZ.Famille = TableClips(I).Famille
                ClipZ.Max = TableClips(I).Max
                ClipZ.Min = TableClips(I).Min
                
                ClipA.Clip = TableClips(I + 1).Clip
                ClipA.Famille = TableClips(I + 1).Famille
                ClipA.Max = TableClips(I + 1).Max
                ClipA.Min = TableClips(I + 1).Min
                
                TableClips(I).Clip = ClipA.Clip
                TableClips(I).Famille = ClipA.Famille
                TableClips(I).Max = ClipA.Max
                TableClips(I).Min = ClipA.Min
                
                TableClips(I + 1).Clip = ClipZ.Clip
                TableClips(I + 1).Famille = ClipZ.Famille
                TableClips(I + 1).Max = ClipZ.Max
                TableClips(I + 1).Min = ClipZ.Min
                I = I - 2
            End If
        End If
            
    End If
    
    
Next
End Sub

Public Sub RendeignePrix()
Dim Sql As String
Dim Rs As Recordset
Dim Champ As String

Sql = "SELECT MyForm.defValue "
Sql = Sql & "FROM ( SELECT Defaults.* "
Sql = Sql & " FROM Defaults  IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyForm "
Sql = Sql & "WHERE MyForm.defName='RefCaddyPrixU';"
Set Rs = Con.OpenRecordSet(Sql)

If Rs.EOF = False Then
    Champ = Rs!defValue
End If
'
Sql = "SELECT MyForm.txt1, MyForm." & Champ & " "
Sql = Sql & "FROM (SELECT con_contacts.* FROM con_contacts  IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS MyForm "
Sql = Sql & "WHERE MyForm.txt1='" & RefConnecteur & "';"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    Prix = Val(Replace(Trim("" & Rs(Champ)), ",", "."))
Else
    Sql = "SELECT MyForm.defValue "
    Sql = Sql & "FROM ( SELECT Defaults.* "
    Sql = Sql & " FROM Defaults  IN '"
    Sql = Sql & TableauPath("Eb_CONNECTIQUE")
    Sql = Sql & "') AS MyForm "
    Sql = Sql & "WHERE MyForm.defName='RefCaddyPrixU';"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = False Then
        Champ = Rs!defValue
    End If
    Sql = "SELECT MyForm.txt1, MyForm." & Champ & " "
    Sql = Sql & "FROM (SELECT con_contacts.* FROM con_contacts  IN '"
    Sql = Sql & TableauPath("Eb_CONNECTIQUE")
    Sql = Sql & "') AS MyForm "
    Sql = Sql & "WHERE MyForm.txt1='" & RefConnecteur & "';"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = False Then
        Prix = Val(Replace(Trim("" & Rs(Champ)), ",", "."))
    End If

End If
  Set Rs = Con.CloseRecordSet(Rs)
End Sub

Public Sub renseigneVoies(Connecteur As String, App As String, Critaire As String)
Dim Sql As String
Dim Rs As Recordset
Dim NbVoie As Long
Sql = "SELECT T_Lien_Con_Famille.Connecteur, T_Lien_Con_Famille_Voies.Voie  "
        Sql = Sql & "FROM T_Lien_Con_Famille INNER JOIN T_Lien_Con_Famille_Voies  "
        Sql = Sql & "ON T_Lien_Con_Famille.Id = T_Lien_Con_Famille_Voies.Id_T_Lien_Con_Famille "
        Sql = Sql & "WHERE T_Lien_Con_Famille.Connecteur='" & RefConnecteur & "' "
        Sql = Sql & "ORDER BY T_Lien_Con_Famille_Voies.Voie;"

Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
   AjouterLstVoie "" & Rs!Voie, Critaire
   
'    Connect.Capot = "" & rs!Capot
'    Connect.Bouchon = "" & rs!Bouchon
'     Connect.Verrou = "" & rs!Verrou
    Rs.MoveNext
Wend
 Set Rs = Con.CloseRecordSet(Rs)
 End Sub
'Public Sub MajTableauFils(RS As Recordset)
'On Error GoTo Fin
'If "" & RS!App = App Then
'    RS![Ref Clip] = Voies(LstVoie("V_" & RS!Voi)).Clip
'    RS![Ref Joint] = Voies(LstVoie("V_" & RS!Voi)).Joint
'End If
'Fin:
'Err.Clear
'On Error GoTo Fin2
'If "" & RS!App2 = App Then
'    RS![Ref Clip2] = Voies(LstVoie("V_" & RS!Voi2)).Clip
'    RS![Ref Joint2] = "" & Voies(LstVoie("V_" & RS!Voi2)).Joint
'End If
'Fin2:
'Err.Clear
'On Error GoTo 0
'RS.Update
'End Sub

Public Sub MajConnecteur(Rs As Recordset)
Dim Sql As String
Dim RsConnecteur As Recordset
On Error GoTo Fin
Dim Id As Long
Id = Rs!Id_IndiceProjet
Debug.Print Rs.Source
If "" & Rs!App = App Then
    Sql = "SELECT Connecteurs.* FROM Connecteurs "
    Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & Rs!Id_IndiceProjet & " "
    Sql = Sql & "AND Connecteurs.CONNECTEUR='" & safeEntry("" & Rs("Ref Connecteur")) & "';"
    Set RsConnecteur = Con.OpenRecordSet(Sql)
    
End If
Fin:
Err.Clear
On Error GoTo Fin2
If "" & Rs!App2 = App Then
Sql = "SELECT Connecteurs.* FROM Connecteurs "
    Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & Rs!Id_IndiceProjet & " "
    Sql = Sql & "AND Connecteurs.CONNECTEUR='" & safeEntry("" & Rs("Ref Connecteur2")) & "';"
    Set RsConnecteur = Con.OpenRecordSet(Sql)
End If
Fin2:
Err.Clear
On Error GoTo 0
If RsConnecteur.EOF = False Then
     RsConnecteur!RefBouchon = Connect.Bouchon
     RsConnecteur!RefCapot = Connect.Capot
     RsConnecteur!refVerrou = Connect.Verrou
    RsConnecteur!RefConnecteurFour = FourConnecteur
    
End If
If RsConnecteur.EOF = False Then
    RsConnecteur.Update
End If
End Sub

 Function retournTableur(rr)
Dim Txt(24, 4) As String
Dim I As Long
For I = 0 To 24
            Txt(I, 0) = rr(I, 0)
            Txt(I, 1) = rr(I, 1)
            Txt(I, 2) = rr(I, 2)
            Txt(I, 3) = rr(I, 3)
            Txt(I, 4) = rr(I, 4)
        Next
retournTableur = Txt
End Function
