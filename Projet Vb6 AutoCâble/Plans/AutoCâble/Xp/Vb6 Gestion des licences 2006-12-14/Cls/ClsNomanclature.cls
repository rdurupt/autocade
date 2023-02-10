VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsNomanclature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim NuInit As String
Private Type Voie
    Voie As String
    Section As Double
    LongueurF_G As Double
    LongueurF_C As Double
    LongueurF_D As Double
    Famille As String
    Clip As String
    FourClip As String
    Joint As String
    FourJoint As String
    Coss As String
    FourCoss As String
End Type
Private Type Clip
    Famille As String
    Clip As String
    ClipFour As String
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
Dim Connect As Connecteur
Public IsEpisure As Boolean
Dim TableClips() As Clip
Dim ClipA As Clip
Dim ClipZ As Clip
Dim LongueurF_Choix As String
Public Prix As Double
Public Connecteur As String
Public FourConnecteur As String
Public App As String
Private Voies() As Voie
Private LstVoie As Collection
Private initClipOk As Boolean
Public ChampCli As String
Public ChampReff As String

Public Function initFilsDirection(Id_Projet As Long)
Dim Rs As Recordset
Dim Sql As String
Sql = "SELECT Connecteurs.LongueurF_Choix FROM Connecteurs "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & Id_Projet & " "
Sql = Sql & "AND Connecteurs.CODE_APP='" & App & "';"

Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
    LongueurF_Choix = "" & Rs!LongueurF_Choix
End If
Set Rs = Con.CloseRecordSet(Rs)
End Function
Public Function IniJoint(RsFils As Recordset)
Dim Sql As String
Dim Rs As Recordset
Dim DIAMETRE  As Double
Dim DiamtrFil() As Fils
Dim a As Fils
Dim Z As Fils
Dim I As Long
Dim I2 As Long
Dim SpliConnecteur
SpliConnecteur = Split(Connecteur & "§", "§")
ReDim Preserve DiamtrFil(I)
Sql = "SELECT DISTINCT lst31.CatName2, con_contacts.txt28 as Diametre, lst29.CatName as Section, lst6.CatName2 "
Sql = Sql & "FROM ((con_contacts INNER JOIN lst31 ON con_contacts.lst31 = lst31.CatID)  "
Sql = Sql & "INNER JOIN lst29 ON con_contacts.lst29 = lst29.CatID)  "
Sql = Sql & "INNER JOIN lst6 ON con_contacts.lst6 = lst6.CatID IN  "
Sql = Sql & "'" & TableauPath("Eb_FILS") & "' "
Sql = Sql & "WHERE lst31.CatName2='" & RsFils!ISO & "'  "
Sql = Sql & "AND con_contacts.txt28<>''  "
Sql = Sql & "AND lst6.CatName2='" & RsFils!TEINT & "';"
Set Rs = Con.OpenRecordSet(Sql)

    While Rs.EOF = False
    I = I + 1
        ReDim Preserve DiamtrFil(I)
            DiamtrFil(I).DIAMETRE = Val(Trim(Replace(Replace("" & Rs!DIAMETRE, "m", ""), ",", ".")))
            DiamtrFil(I).Section = Val(Trim(Replace(Replace("" & Rs!Section, "mm2", ""), ",", ".")))
            
        Rs.MoveNext
    Wend
    For I = 1 To UBound(DiamtrFil) - 1
        If DiamtrFil(I).Section > DiamtrFil(I + 1).Section Then
            Z.DIAMETRE = DiamtrFil(I).DIAMETRE
            Z.Section = DiamtrFil(I).Section
            a.DIAMETRE = DiamtrFil(I + 1).DIAMETRE
            a.Section = DiamtrFil(I + 1).Section
            DiamtrFil(I + 1).DIAMETRE = Z.DIAMETRE
            DiamtrFil(I + 1).Section = Z.Section
            DiamtrFil(I).DIAMETRE = a.DIAMETRE
            DiamtrFil(I).Section = a.Section
            I = I - 2
        Else
            If DiamtrFil(I).DIAMETRE > DiamtrFil(I + 1).DIAMETRE And DiamtrFil(I).Section = DiamtrFil(I + 1).Section Then
                Z.DIAMETRE = DiamtrFil(I).DIAMETRE
                Z.Section = DiamtrFil(I).Section
                a.DIAMETRE = DiamtrFil(I + 1).DIAMETRE
                a.Section = DiamtrFil(I + 1).Section
                DiamtrFil(I + 1).DIAMETRE = Z.DIAMETRE
                DiamtrFil(I + 1).Section = Z.Section
                DiamtrFil(I).DIAMETRE = a.DIAMETRE
                DiamtrFil(I).Section = a.Section
                I = I - 2
            End If
        End If
    Next I
    For I = 1 To UBound(DiamtrFil)
    If Val(Replace("" & RsFils!SECT, ",", ".")) = DiamtrFil(I).Section Then
        DIAMETRE = DiamtrFil(I).DIAMETRE
        Exit For
    End If
        
    Next
     If "" & RsFils![App] = App Then
        Sql = "SELECT  Myform.txt3,T_Lien_Con_Famille_Connecteur_Joint.Joint, T_Lien_Con_Famille_Connecteur_Joint.[Isolant Min], "
        Sql = Sql & "T_Lien_Con_Famille_Connecteur_Joint.[Isolant Max] "
        Sql = Sql & "FROM (T_Lien_Con_Famille INNER JOIN T_Lien_Con_Famille_Connecteur_Joint ON T_Lien_Con_Famille.Id =  "
        Sql = Sql & "T_Lien_Con_Famille_Connecteur_Joint.Id_Connecteur) INNER JOIN (SELECT DISTINCT   con_contacts.*  "
        Sql = Sql & "FROM con_contacts IN '"
        Sql = Sql & TableauPath("Eb_JOINTS")
        Sql = Sql & "') AS Myform ON T_Lien_Con_Famille_Connecteur_Joint.Joint = Myform.txt1 "
        Sql = Sql & "WHERE T_Lien_Con_Famille.Connecteur='" & SpliConnecteur(0) & "' "
        Sql = Sql & "AND T_Lien_Con_Famille_Connecteur_Joint.[Isolant Min]>=" & Replace(DIAMETRE, ",", ".") & " "
        Sql = Sql & "AND T_Lien_Con_Famille_Connecteur_Joint.[Isolant Max]>=" & Replace(DIAMETRE, ",", ".") & ";"
        
'        Sql = "SELECT T_Lien_Con_Famille_Connecteur_Joint.RefFour, T_Lien_Con_Famille_Connecteur_Joint.Joint,  "
'        Sql = Sql & "T_Lien_Con_Famille_Connecteur_Joint.[Isolant Min], T_Lien_Con_Famille_Connecteur_Joint.[Isolant Max] "
'        Sql = Sql & "FROM T_Lien_Con_Famille INNER JOIN T_Lien_Con_Famille_Connecteur_Joint ON T_Lien_Con_Famille.Id = T_Lien_Con_Famille_Connecteur_Joint.Id_Connecteur"
'WHERE (((T_Lien_Con_Famille_Connecteur_Joint.[Isolant Min])>=1.7) AND ((T_Lien_Con_Famille_Connecteur_Joint.[Isolant Max])>=1.7) AND ((T_Lien_Con_Famille.Connecteur)='243427Y000'));

        
        Set Rs = Con.OpenRecordSet(Sql)
        While Rs.EOF = False
            Voies(LstVoie("V_" & RsFils!VOI)).Joint = Voies(LstVoie("V_" & RsFils!VOI)).Joint & "" & Rs!Joint & Chr(10)
             Voies(LstVoie("V_" & RsFils!VOI)).FourJoint = Voies(LstVoie("V_" & RsFils!VOI)).FourJoint & "" & Rs!txt3 & Chr(10)
            Rs.MoveNext
        Wend
  End If
   If "" & RsFils![App2] = App Then
        Sql = "SELECT  Myform.txt3,T_Lien_Con_Famille_Connecteur_Joint.Joint, T_Lien_Con_Famille_Connecteur_Joint.[Isolant Min], "
        Sql = Sql & "T_Lien_Con_Famille_Connecteur_Joint.[Isolant Max] "
        Sql = Sql & "FROM (T_Lien_Con_Famille INNER JOIN T_Lien_Con_Famille_Connecteur_Joint ON T_Lien_Con_Famille.Id =  "
        Sql = Sql & "T_Lien_Con_Famille_Connecteur_Joint.Id_Connecteur) INNER JOIN (SELECT DISTINCT   con_contacts.*  "
        Sql = Sql & "FROM con_contacts IN '"
        Sql = Sql & TableauPath("Eb_JOINTS")
        Sql = Sql & "') AS Myform ON T_Lien_Con_Famille_Connecteur_Joint.Joint = Myform.txt1 "
        Sql = Sql & "WHERE T_Lien_Con_Famille.Connecteur='" & SpliConnecteur(0) & "' "
        Sql = Sql & "AND T_Lien_Con_Famille_Connecteur_Joint.[Isolant Min]>=" & Replace(DIAMETRE, ",", ".") & " "
        Sql = Sql & "AND T_Lien_Con_Famille_Connecteur_Joint.[Isolant Max]>=" & Replace(DIAMETRE, ",", ".") & ";"
        Set Rs = Con.OpenRecordSet(Sql)
        While Rs.EOF = False
            Voies(LstVoie("V_" & RsFils!VOI2)).Joint = Voies(LstVoie("V_" & RsFils!VOI2)).Joint & "" & Rs!Joint & Chr(10)
             Voies(LstVoie("V_" & RsFils!VOI2)).FourJoint = Voies(LstVoie("V_" & RsFils!VOI2)).FourJoint & "" & Rs!txt3 & Chr(10)
            Rs.MoveNext
        Wend
    End If

Set Rs = Con.CloseRecordSet(Rs)
End Function

Public Function RendeigneConnecteur(Connecteur As String) As String
Dim Sql As String
Dim Rs As Recordset
If Trim("" & Connect.Verrou) <> "" Then Exit Function
Sql = "SELECT  T_Lien_Con_Verroux.Référence "
Sql = Sql & "FROM T_Lien_Con_Famille INNER JOIN T_Lien_Con_Verroux  "
Sql = Sql & "ON T_Lien_Con_Famille.Id = T_Lien_Con_Verroux.Id_T_Lien_Con_Famille "
Sql = Sql & "WHERE T_Lien_Con_Famille.Connecteur='" & Connecteur & "';"


Sql = "SELECT T_Lien_Con_Verroux.Référence, FromVerrou." & ChampReff & " As RefFour "
Sql = Sql & "FROM (SELECT con_contacts.* FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS FromConnecteur INNER JOIN (T_Lien_Con_Famille INNER  "
Sql = Sql & "JOIN (T_Lien_Con_Verroux INNER  "
Sql = Sql & "JOIN (SELECT con_contacts.* FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CAPOTS")
Sql = Sql & "') AS FromVerrou ON T_Lien_Con_Verroux.Référence = FromVerrou.txt1)  "
Sql = Sql & "ON T_Lien_Con_Famille.Id = T_Lien_Con_Verroux.Id_T_Lien_Con_Famille)  "
Sql = Sql & "ON FromConnecteur.txt1 = T_Lien_Con_Famille.Connecteur "
Sql = Sql & "WHERE FromConnecteur." & ChampCli & "='" & Connecteur & "';"



Set Rs = Con.OpenRecordSet(Sql)
 If Rs.EOF = False Then
    Connect.Verrou = "" & Rs!Référence
    Connect.FourVerrou = "" & Rs!refFour
End If
    
   Sql = "SELECT DISTINCT T_LientConnecteur.RefBouchon "
    Sql = Sql & "FROM T_LientConnecteur "
    Sql = Sql & "WHERE T_LientConnecteur.RefConnecteur='" & Connecteur & "';"
    
    
Sql = "SELECT  T_Lien_Con_Bouchons.Référence,FromBouchon." & ChampReff & " AS RefFour  "
Sql = Sql & "FROM (((SELECT con_contacts.* FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS FromConnecteur INNER   "
Sql = Sql & "JOIN T_Lien_Con_Famille ON FromConnecteur.txt1 = T_Lien_Con_Famille.Connecteur)   "
Sql = Sql & "INNER JOIN T_Lien_Con_Bouchons ON T_Lien_Con_Famille.Id = T_Lien_Con_Bouchons.Id_T_Lien_Con_Famille)   "
Sql = Sql & "INNER JOIN (SELECT con_contacts.* FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_JOINTS")
Sql = Sql & "') AS FromBouchon ON T_Lien_Con_Bouchons.Référence = FromBouchon.txt1  "
Sql = Sql & "WHERE FromConnecteur." & ChampCli & "='" & Connecteur & "';"
    
    Set Rs = Con.OpenRecordSet(Sql)

    While Rs.EOF = False
        Connect.Bouchon = Connect.Bouchon & Chr(10) & "" & Rs!Référence
        Connect.FourBouchon = Connect.FourBouchon & Chr(10) & "" & Rs!refFour
        Rs.MoveNext
    Wend
Sql = "SELECT T_Lien_Con_Capot.Référence, FromCapot.txt3 AS RefFour "
Sql = Sql & "FROM (((SELECT con_contacts.* FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "') AS FromConnecteur INNER JOIN T_Lien_Con_Famille ON FromConnecteur.txt1 = T_Lien_Con_Famille.Connecteur)  "
Sql = Sql & "INNER JOIN T_Lien_Con_Capot ON T_Lien_Con_Famille.Id = T_Lien_Con_Capot.Id_T_Lien_Con_Famille)  "
Sql = Sql & "INNER JOIN (SELECT con_contacts.* FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CAPOTS")
Sql = Sql & "') AS FromCapot ON T_Lien_Con_Capot.Référence = FromCapot.txt1 "
Sql = Sql & "WHERE FromConnecteur." & ChampCli & "='" & Connecteur & "';"
Set Rs = Con.OpenRecordSet(Sql)

    While Rs.EOF = False
        Connect.Capot = Connect.Capot & " " & Rs!Référence
        Connect.FourCapot = Connect.FourCapot & " " & Rs!refFour
        Rs.MoveNext
    Wend
Connect.Capot = Trim(Connect.Capot)
Connect.FourCapot = Trim(Connect.FourCapot)
Set Rs = Con.CloseRecordSet(Rs)
End Function
Public Sub DelBouchon()
Dim I As Long
'For I = 0 To LstVoie.Count - 1
''    If Voies(I).Section <> 0 Then Voies(I).Bouchon = ""
'Next
End Sub
Public Sub SubSection(Voie As String, Section As Double)
On Error Resume Next
    Voies(LstVoie("V_" & Voie)).Section = Voies(LstVoie("V_" & Voie)).Section + Section
Err.Clear
   
End Sub
Private Sub AjouterLstVoie(Voie As String)
On Error Resume Next
    LstVoie.Add LstVoie.Count, "V_" & Voie
    ReDim Preserve Voies(LstVoie("V_" & Voie))
    Voies(LstVoie("V_" & Voie)).Voie = Voie
Err.Clear
End Sub

Private Sub Class_Initialize()
    Set LstVoie = New Collection
End Sub
Public Sub ChoixClip()
Dim I As Long
Dim IndexClip As Long
Dim Sql As String
Dim Rs As Recordset
Dim SplitConnecteur
SplitConnecteur = Split(Connecteur & "§", "§")
    
For I = 0 To LstVoie.Count - 1
    Sql = "SELECT DISTINCT T_Lien_Con_Famille_Voies_Clip.Référence, T_Lien_Con_Famille_Voies_Clip.RefFour, "
    Sql = Sql & "T_Lien_Con_Famille_Voies_Clip.Famille, T_Lien_Con_Famille_Voies.Voie,  "
    Sql = Sql & "T_Lien_Con_Famille_Voies_Clip.Min, T_Lien_Con_Famille_Voies_Clip.Max "
    Sql = Sql & "FROM T_Lien_Con_Famille INNER  "
    Sql = Sql & "JOIN (T_Lien_Con_Famille_Voies INNER JOIN T_Lien_Con_Famille_Voies_Clip  "
    Sql = Sql & "ON T_Lien_Con_Famille_Voies.Id =  "
    Sql = Sql & "T_Lien_Con_Famille_Voies_Clip.Id_T_Lien_Con_Famille_Voies)  "
    Sql = Sql & "ON T_Lien_Con_Famille.Id = T_Lien_Con_Famille_Voies.Id_T_Lien_Con_Famille "
    Sql = Sql & "Where T_Lien_Con_Famille.Connecteur = '" & SplitConnecteur(0) & "'  "
    Sql = Sql & "And T_Lien_Con_Famille_Voies.Voie = '" & Voies(I).Voie & "' "
    Sql = Sql & "And T_Lien_Con_Famille_Voies_Clip.Min <= " & Replace(Voies(I).Section, ",", ".") & "  "
    Sql = Sql & "And T_Lien_Con_Famille_Voies_Clip.Max >= " & Replace(Voies(I).Section, ",", ".") & " "
    Sql = Sql & "ORDER BY T_Lien_Con_Famille_Voies_Clip.Min, T_Lien_Con_Famille_Voies_Clip.Max;"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = False Then
        While Rs.EOF = False
            Voies(I).Famille = Voies(I).Famille & Chr(10) & "" & Rs!Famille
            Voies(I).Clip = Voies(I).Clip & Chr(10) & Rs!Famille & " : " & Rs!Référence
            Voies(I).FourClip = Voies(I).FourClip & Chr(10) & Rs!Famille & " : " & Rs!Référence
            Rs.MoveNext
        Wend
    Else
      For IndexClip = 1 To UBound(TableClips)
        If Voies(I).Section >= TableClips(IndexClip).Min And Voies(I).Section <= TableClips(IndexClip).Max Then
            Voies(I).Famille = Voies(I).Famille & Chr(10) & TableClips(IndexClip).Famille
            Voies(I).Clip = Voies(I).Clip & Chr(10) & TableClips(IndexClip).Famille & " : " & TableClips(IndexClip).Clip
            Voies(I).FourClip = Voies(I).FourClip & Chr(10) & TableClips(IndexClip).Famille & " : " & TableClips(IndexClip).ClipFour
            
        End If
      Next
   End If
   Set Rs = Con.CloseRecordSet(Rs)
Next
End Sub
Public Sub InitCip()
Dim Sql As String
Dim Rs As Recordset
Dim I As Long
Dim SplitConnecteur
SplitConnecteur = Split(Connecteur & "§", "§")

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
Sql = Sql & "where  T_Lien_Connecteur_Clip.RefConnecteur='" & SplitConnecteur(0) & "';"
Set Rs = Con.OpenRecordSet(Sql)

If Rs.EOF = True Then
Sql = "SELECT MyForm.txt1 AS RefConnecteur, MyForm2.ClipCli AS Txt1, MyForm2.ClipFour AS [Alvé Réf Fourr], MyForm2.Famille AS [Famille Lib],  "
Sql = Sql & "MyForm2.Min AS [Alvéole Mini en mm2], MyForm2.Max AS [Alvéole Maxi en mm2] "
Sql = Sql & "FROM (SELECT con_contacts.* FROM con_contacts IN '"
Sql = Sql & TableauPath("Eb_CONNECTEURS")
Sql = Sql & "' where  con_contacts.txt1='" & SplitConnecteur(0) & "') AS MyForm  "
Sql = Sql & "INNER JOIN (SELECT con_contacts.txt1 AS ClipCli, con_contacts.txt3 AS ClipFour, lst21.CatName AS Famille, lst21.CatID AS IdFamille,  "
Sql = Sql & "lst22.CatName AS [Min], lst23.CatName AS [Max] "
Sql = Sql & "FROM ((con_contacts INNER JOIN lst21 ON con_contacts.lst21 = lst21.CatID) LEFT JOIN lst22 ON con_contacts.lst22 = lst22.CatID)  "
Sql = Sql & "LEFT JOIN lst23 ON con_contacts.lst23 = lst23.CatID IN '"
 Sql = Sql & TableauPath("Eb_CONNECTIQUE")
Sql = Sql & "') AS MyForm2 ON MyForm.lst21 = MyForm2.IdFamille;"
Set Rs = Con.OpenRecordSet(Sql)

    
End If
ReDim TableClips(I)
While Rs.EOF = False
 I = I + 1
  ReDim Preserve TableClips(I)
    TableClips(I).Famille = "" & Rs![Famille Lib]
    TableClips(I).Clip = "" & Rs!txt1
    TableClips(I).ClipFour = "" & Rs![Alvé Réf Fourr]
    TableClips(I).Min = Val(Replace(Trim(Replace("" & Rs![Alvéole Mini en mm2], "mm2", "")), ",", "."))
    TableClips(I).Max = Val(Replace(Trim(Replace("" & Rs![Alvéole Maxi en mm2], "mm2", "")), ",", "."))
   
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
For I = 1 To UBound(TableClips) - 1
    If TableClips(I).Min > TableClips(I + 1).Min Then
        ClipZ.Clip = TableClips(I).Clip
        ClipZ.ClipFour = TableClips(I).ClipFour
        ClipZ.Famille = TableClips(I).Famille
        ClipZ.Max = TableClips(I).Max
        ClipZ.Min = TableClips(I).Min
        
        ClipA.Clip = TableClips(I + 1).Clip
        ClipA.ClipFour = TableClips(I + 1).ClipFour
        ClipA.Famille = TableClips(I + 1).Famille
        ClipA.Max = TableClips(I + 1).Max
        ClipA.Min = TableClips(I + 1).Min
        
        TableClips(I).Clip = ClipA.Clip
        TableClips(I).ClipFour = ClipA.ClipFour
        TableClips(I).Famille = ClipA.Famille
        TableClips(I).Max = ClipA.Max
        TableClips(I).Min = ClipA.Min
        
        TableClips(I + 1).Clip = ClipZ.Clip
        TableClips(I + 1).ClipFour = ClipZ.ClipFour
        TableClips(I + 1).Famille = ClipZ.Famille
        TableClips(I + 1).Max = ClipZ.Max
        TableClips(I + 1).Min = ClipZ.Min
        I = I - 2
    Else
        If TableClips(I).Min = TableClips(I + 1).Min Then
            If TableClips(I).Max > TableClips(I + 1).Max Then
                 ClipZ.Clip = TableClips(I).Clip
                ClipZ.ClipFour = TableClips(I).ClipFour
                ClipZ.Famille = TableClips(I).Famille
                ClipZ.Max = TableClips(I).Max
                ClipZ.Min = TableClips(I).Min
                
                ClipA.Clip = TableClips(I + 1).Clip
                ClipA.ClipFour = TableClips(I + 1).ClipFour
                ClipA.Famille = TableClips(I + 1).Famille
                ClipA.Max = TableClips(I + 1).Max
                ClipA.Min = TableClips(I + 1).Min
                
                TableClips(I).Clip = ClipA.Clip
                TableClips(I).ClipFour = ClipA.ClipFour
                TableClips(I).Famille = ClipA.Famille
                TableClips(I).Max = ClipA.Max
                TableClips(I).Min = ClipA.Min
                
                TableClips(I + 1).Clip = ClipZ.Clip
                TableClips(I + 1).ClipFour = ClipZ.ClipFour
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
Dim SplitConnecteur
SplitConnecteur = Split(Connecteur & "§", "§")
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
Sql = Sql & "WHERE MyForm.txt1='" & SplitConnecteur(0) & "';"
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
    Sql = Sql & "WHERE MyForm.txt1='" & SplitConnecteur(0) & "';"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = False Then
        Prix = Val(Replace(Trim("" & Rs(Champ)), ",", "."))
    End If

End If
  Set Rs = Con.CloseRecordSet(Rs)
End Sub

Public Sub renseigneVoies(Connecteur As String, App As String)
Dim Sql As String
Dim Rs As Recordset
Dim NbVoie As Long
Dim SplitConnecteur
Dim IndexEpissure As Long
Sql = "SELECT T_Lien_Con_Famille.Connecteur, T_Lien_Con_Famille_Voies.Voie,  "
Sql = Sql & "T_Lien_Con_Famille_Voies.LongueurF_G, T_Lien_Con_Famille_Voies.LongueurF_C,  "
Sql = Sql & "T_Lien_Con_Famille_Voies.LongueurF_D, T_Lien_Con_Capot.Référence AS Capot,  "
Sql = Sql & "T_Lien_Con_Verroux.Référence AS Verrou, T_Lien_Con_Bouchons.Référence AS Bouchon  "
Sql = Sql & "FROM (((T_Lien_Con_Famille INNER JOIN T_Lien_Con_Famille_Voies  "
Sql = Sql & "ON T_Lien_Con_Famille.Id = T_Lien_Con_Famille_Voies.Id_T_Lien_Con_Famille)  "
Sql = Sql & "LEFT JOIN T_Lien_Con_Bouchons ON T_Lien_Con_Famille.Id =  "
Sql = Sql & "T_Lien_Con_Bouchons.Id_T_Lien_Con_Famille) LEFT JOIN T_Lien_Con_Capot  "
Sql = Sql & "ON T_Lien_Con_Famille.Id = T_Lien_Con_Capot.Id_T_Lien_Con_Famille)  "
Sql = Sql & "LEFT JOIN T_Lien_Con_Verroux ON T_Lien_Con_Famille.Id =  "
Sql = Sql & "T_Lien_Con_Verroux.Id_T_Lien_Con_Famille "
SplitConnecteur = Split(Connecteur & "§", "§")
Sql = Sql & "WHERE T_Lien_Con_Famille.Connecteur='" & SplitConnecteur(0) & "';"

Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then
   For SplitConnecteur = 1 To 99
    AjouterLstVoie "G" & SplitConnecteur
    AjouterLstVoie "D" & SplitConnecteur
   Next
End If
While Rs.EOF = False
   AjouterLstVoie "" & Rs!Voie
   
'    Connect.Capot = "" & Rs!Capot
'    Connect.Bouchon = "" & Rs!Bouchon
'     Connect.Verrou = "" & Rs!Verrou
    Rs.MoveNext
Wend
 Set Rs = Con.CloseRecordSet(Rs)
 End Sub
Public Sub MajTableauFils(Rs As Recordset)
Dim Clip
On Error GoTo Fin
If VoieExiste(Trim("V_" & Rs!VOI)) = False Then GoTo Fin

If "" & Rs!App = App Then
    Rs![Ref Clip] = Voies(LstVoie("V_" & Rs!VOI)).Clip
    Rs![Ref Clip Four] = Voies(LstVoie("V_" & Rs!VOI)).FourClip
    Rs![Ref Joint] = Voies(LstVoie("V_" & Rs!VOI)).Joint
     Rs![Ref Joint Four] = Voies(LstVoie("V_" & Rs!VOI)).FourJoint
    Rs![Ref Connecteur_Four] = FourConnecteur
    Select Case UCase(LongueurF_Choix)
        Case "D"
            Rs!Long_Add = Voies(LstVoie("V_" & Rs!VOI)).LongueurF_D
        Case "G"
            Rs!Long_Add = Voies(LstVoie("V_" & Rs!VOI)).LongueurF_G
        Case Else
            Rs!Long_Add = Voies(LstVoie("V_" & Rs!VOI)).LongueurF_C
    End Select
    
End If
Fin:
Err.Clear
On Error GoTo 0
On Error GoTo Fin2
DoEvents
If VoieExiste(Trim("V_" & Rs!VOI2)) = False Then GoTo Fin2
If "" & Rs!App2 = App Then

    Rs![Ref Clip2] = Voies(LstVoie("V_" & Rs!VOI2)).Clip
    Rs![Ref Clip Four2] = Voies(LstVoie("V_" & Rs!VOI2)).FourClip
    
    Rs![Ref Joint2] = "" & Voies(LstVoie("V_" & Rs!VOI2)).Joint
    Rs![Ref Joint Four2] = Voies(LstVoie("V_" & Rs!VOI)).FourJoint
    Rs![Ref Connecteur_Four2] = FourConnecteur
    Select Case UCase(LongueurF_Choix)
        Case "D"
            Rs!Long_Add2 = Voies(LstVoie("V_" & Rs!VOI2)).LongueurF_D
        Case "G"
            Rs!Long_Add2 = Voies(LstVoie("V_" & Rs!VOI2)).LongueurF_G
        Case Else
            Rs!Long_Add2 = Voies(LstVoie("V_" & Rs!VOI2)).LongueurF_C
    End Select
End If
Fin2:
Err.Clear
On Error GoTo 0
Rs.Update
End Sub
Private Function VoieExiste(Voies As String) As Boolean
Dim Voie As String
On Error Resume Next
Voie = LstVoie(Voies)
If Err = 0 Then
    VoieExiste = True
End If
End Function
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
    Sql = Sql & "AND Connecteurs.CODE_APP='" & safeEntry("" & Rs("app")) & "';"
    Set RsConnecteur = Con.OpenRecordSet(Sql)
    
End If
Fin:
Err.Clear
On Error GoTo Fin2
If "" & Rs!App2 = App Then
Sql = "SELECT Connecteurs.* FROM Connecteurs "
    Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & Rs!Id_IndiceProjet & " "
    Sql = Sql & "AND Connecteurs.CODE_APP='" & safeEntry("" & Rs("app2")) & "';"
    Set RsConnecteur = Con.OpenRecordSet(Sql)
End If
Fin2:
Err.Clear
On Error GoTo 0
If RsConnecteur.EOF = False Then
     RsConnecteur!RefBouchon = Connect.Bouchon
     RsConnecteur!RefBouchonFour = Connect.FourBouchon
     RsConnecteur!RefCapot = Connect.Capot
     RsConnecteur!ReFCapotFour = Connect.FourCapot
     RsConnecteur!refVerrou = Connect.Verrou
     RsConnecteur!RefVerrouFour = Connect.FourVerrou
    RsConnecteur!RefConnecteurFour = FourConnecteur
    
End If
If RsConnecteur.EOF = False Then
    RsConnecteur.Update
End If
End Sub
