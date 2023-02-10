VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PreparISO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public IsTableCrate As Boolean
Public colecISO As Collection
Private Type Fils
    Section As Double
    DIAMETRE As Double
    Ref As String
    refFour As String
    
End Type

Private Sub Class_Initialize()
Set colecISO = New Collection

End Sub

Private Sub Class_Terminate()
Set colecISO = Nothing
End Sub
Public Function DefinIso(RsFils As Recordset)
Dim Sql As String
Dim Rs As Recordset
Dim DIAMETRE  As Double
Dim DiamtrFil() As Fils
Dim a As Fils
Dim Z As Fils
Dim I As Long
Dim I2 As Long
ReDim Preserve DiamtrFil(I)
Sql = "SELECT DISTINCT lst31.CatName2, con_contacts.txt28 AS Diametre, lst29.CatName2 AS [Section], lst6.CatName2, con_contacts.txt1, con_contacts.txt3, Val('' & [txt81])+Val('' & [txt11]) AS Expr1 "
Sql = Sql & "FROM ((con_contacts INNER JOIN lst31 ON con_contacts.lst31 = lst31.CatID)  "
Sql = Sql & "INNER JOIN lst29 ON con_contacts.lst29 = lst29.CatID)  "
Sql = Sql & "INNER JOIN lst6 ON con_contacts.lst6 = lst6.CatID IN  "
Sql = Sql & "'" & TableauPath("Eb_FILS") & "' "
Sql = Sql & "WHERE lst31.CatName2='" & RsFils!ISO & "'  "
'Sql = Sql & "AND con_contacts.txt28<>''  "
Sql = Sql & "AND lst6.CatName2='" & RsFils!TEINT & "'  AND Val('' & [txt81])+Val('' & [txt11])>0 "
Sql = Sql & "ORDER BY con_contacts.txt28, Val('' & [txt81])+Val('' & [txt11]) DESC;"
Set Rs = Con.OpenRecordSet(Sql)
    While Rs.EOF = False
    I = I + 1
        ReDim Preserve DiamtrFil(I)
            DiamtrFil(I).DIAMETRE = Val(Trim(Replace(Replace("" & Rs!DIAMETRE, "m", ""), ",", ".")))
            DiamtrFil(I).Section = Val(Trim(Replace(Replace("" & Rs!Section, "mm2", ""), ",", ".")))
            DiamtrFil(I).Ref = Trim("" & Rs!txt1)
            DiamtrFil(I).refFour = Trim("" & Rs!txt3)
            
        Rs.MoveNext
    Wend
    For I = 1 To UBound(DiamtrFil) - 1
        If DiamtrFil(I).Section > DiamtrFil(I + 1).Section Then
        
            Z.DIAMETRE = DiamtrFil(I).DIAMETRE
            Z.Section = DiamtrFil(I).Section
            Z.Ref = DiamtrFil(I).Ref
            Z.Ref = DiamtrFil(I).refFour
            
            a.DIAMETRE = DiamtrFil(I + 1).DIAMETRE
            a.Section = DiamtrFil(I + 1).Section
            a.Ref = DiamtrFil(I + 1).Ref
            a.refFour = DiamtrFil(I + 1).refFour
            
            DiamtrFil(I + 1).DIAMETRE = Z.DIAMETRE
            DiamtrFil(I + 1).Section = Z.Section
            DiamtrFil(I + 1).Ref = Z.Ref
            DiamtrFil(I + 1).refFour = Z.refFour
            
            DiamtrFil(I).DIAMETRE = a.DIAMETRE
            DiamtrFil(I).Section = a.Section
            DiamtrFil(I).Ref = a.Ref
            DiamtrFil(I).refFour = a.refFour
            
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
    If Val(Replace("" & RsFils!Section, ",", ".")) = DiamtrFil(I).Section Then
        DIAMETRE = DiamtrFil(I).DIAMETRE
        If IsTableCrate = False Then
           Sql = " SELECT '" & RsFils!Liai & "' AS LIAI, '" & DiamtrFil(I).Ref & "' AS Ref, '" & DiamtrFil(I).refFour & "' AS RefFour INTO Temp_ISO_" & NmJob & " ;"
            Con.Execute Sql
            IsTableCrate = True
   Else
        Sql = "INSERT INTO Temp_ISO_" & NmJob & " ( LIAI, Ref, RefFour ) "
        Sql = Sql & " values ('" & RsFils!Liai & "', '" & DiamtrFil(I).Ref & "', '" & DiamtrFil(I).refFour & "');"
        Con.Execute Sql

   End If
        Exit For
    End If
        
    Next
   
'    If "" & RsFils![Ref Connecteur] = Connecteur Then
'        Sql = "SELECT  Myform.txt3,T_Lien_Con_Famille_Connecteur_Joint.Joint, T_Lien_Con_Famille_Connecteur_Joint.[Isolant Min], "
'        Sql = Sql & "T_Lien_Con_Famille_Connecteur_Joint.[Isolant Max] "
'        Sql = Sql & "FROM (T_Lien_Con_Famille INNER JOIN T_Lien_Con_Famille_Connecteur_Joint ON T_Lien_Con_Famille.Id =  "
'        Sql = Sql & "T_Lien_Con_Famille_Connecteur_Joint.Id_Connecteur) INNER JOIN (SELECT DISTINCT   con_contacts.*  "
'        Sql = Sql & "FROM con_contacts IN '"
'        Sql = Sql & TableauPath("Eb_JOINTS")
'        Sql = Sql & "') AS Myform ON T_Lien_Con_Famille_Connecteur_Joint.Joint = Myform.txt1 "
'        Sql = Sql & "WHERE T_Lien_Con_Famille.Connecteur='" & RsFils![Ref Connecteur] & "' "
'        Sql = Sql & "AND T_Lien_Con_Famille_Connecteur_Joint.[Isolant Min]>=" & Replace(DIAMETRE, ",", ".") & " "
'        Sql = Sql & "AND T_Lien_Con_Famille_Connecteur_Joint.[Isolant Max]>=" & Replace(DIAMETRE, ",", ".") & ";"
'        Set Rs = Con.OpenRecordSet(Sql)
'        While Rs.EOF = False
'            Voies(LstVoie("V_" & RsFils!VOI)).Joint = Voies(LstVoie("V_" & RsFils!VOI)).Joint & "" & Rs!Joint & Chr(10)
'             Voies(LstVoie("V_" & RsFils!VOI)).FourJoint = Voies(LstVoie("V_" & RsFils!VOI)).FourJoint & "" & Rs!txt3 & Chr(10)
'            Rs.MoveNext
'        Wend
'    Else
'        Sql = "SELECT  Myform.txt3,T_Lien_Con_Famille_Connecteur_Joint.Joint, T_Lien_Con_Famille_Connecteur_Joint.[Isolant Min], "
'        Sql = Sql & "T_Lien_Con_Famille_Connecteur_Joint.[Isolant Max] "
'        Sql = Sql & "FROM (T_Lien_Con_Famille INNER JOIN T_Lien_Con_Famille_Connecteur_Joint ON T_Lien_Con_Famille.Id =  "
'        Sql = Sql & "T_Lien_Con_Famille_Connecteur_Joint.Id_Connecteur) INNER JOIN (SELECT DISTINCT   con_contacts.*  "
'        Sql = Sql & "FROM con_contacts IN '"
'        Sql = Sql & TableauPath("Eb_JOINTS")
'        Sql = Sql & "') AS Myform ON T_Lien_Con_Famille_Connecteur_Joint.Joint = Myform.txt1 "
'        Sql = Sql & "WHERE T_Lien_Con_Famille.Connecteur='" & RsFils![Ref Connecteur2] & "' "
'        Sql = Sql & "AND T_Lien_Con_Famille_Connecteur_Joint.[Isolant Min]>=" & Replace(DIAMETRE, ",", ".") & " "
'        Sql = Sql & "AND T_Lien_Con_Famille_Connecteur_Joint.[Isolant Max]>=" & Replace(DIAMETRE, ",", ".") & ";"
'        Set Rs = Con.OpenRecordSet(Sql)
'        While Rs.EOF = False
'            Voies(LstVoie("V_" & RsFils!VOI2)).Joint = Voies(LstVoie("V_" & RsFils!VOI2)).Joint & "" & Rs!Joint & Chr(10)
'             Voies(LstVoie("V_" & RsFils!VOI2)).FourJoint = Voies(LstVoie("V_" & RsFils!VOI2)).FourJoint & "" & Rs!txt3 & Chr(10)
'            Rs.MoveNext
'        Wend
'    End If

Set Rs = Con.CloseRecordSet(Rs)
End Function

