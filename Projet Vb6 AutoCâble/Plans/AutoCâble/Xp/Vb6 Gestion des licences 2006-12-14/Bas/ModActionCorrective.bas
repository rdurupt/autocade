Attribute VB_Name = "ModActionCorrective"
Option Explicit

Sub SubActionCorrective(IdIndiceProjet As Long, IdFils As Long)
Dim Rs As Recordset
Dim RsModifier As Recordset
Dim sql As String
Dim PathPl As String
Dim PathPl2 As String
Dim MyWord As Object
'MyWord.Visible = True

sql = "SELECT RqCartouche.* "
sql = sql & "FROM RqCartouche "
sql = sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(sql)


MyWord.Documents.Add DefinirChemienComplet(TableauPath("PathServer"), DefinirChemienComplet(TableauPath("PathArchiveAutocad"), TableauPath("ModelNC")))
'MyWord.Visible = True
RemplaceWord MyWord, """Client""", "" & Rs!Client
OuvreEnteteWord MyWord
RemplaceWord MyWord, """NC""", "" & Rs!dnc
RemplaceWord MyWord, """date""", Format(Date, "dd/mm/yyyy")
FermeEnteteWord MyWord

PathPl = PathArchive(TableauPath.Item("PathArchiveAutocad"), "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "DNC", Rs.Fields("dnc"), IdIndiceProjet, Rs.Fields("PI_Indice"), "", Rs!Version, True)
MyWord.ActiveDocument.SaveAs PathPl
MyWord.ActiveDocument.Close , False

 If IdFils <> 0 Then
        sql = "SELECT RqCartouche.* "
        sql = sql & "FROM RqCartouche "
        sql = sql & "WHERE T_indiceProjet.Id=" & IdFils & ";"
        Set RsModifier = Con.OpenRecordSet(sql)
         PathPl2 = PathArchive(TableauPath.Item("PathArchiveAutocad"), "" & RsModifier!Client, "" & RsModifier!CleAc, "" & RsModifier!Pieces, "DNC", RsModifier.Fields("dnc"), RsModifier!Id, RsModifier.Fields("PI_Indice"), "", RsModifier!Version, True)
        Racourci "" & PathPl2, "" & PathPl, "DOC"

    End If



MyWord.Documents.Add DefinirChemienComplet(TableauPath("PathServer"), DefinirChemienComplet(TableauPath("PathArchiveAutocad"), TableauPath("ModelAC")))
RemplaceWord MyWord, """Client""", "" & Rs!Client
RemplaceWord MyWord, """LI""", "" & Rs!LIEC
RemplaceWord MyWord, """date""", Format(Date, "dd/mm/yyyy")
OuvreEnteteWord MyWord
RemplaceWord MyWord, """date""", Format(Date, "dd/mm/yyyy")
RemplaceWord MyWord, """AC""", "" & Rs!ReffIndice
FermeEnteteWord MyWord
PathPl = PathArchive(TableauPath.Item("PathArchiveAutocad"), "" & Rs!Client, "" & Rs!CleAc, "" & Rs!Pieces, "LIEC", Rs.Fields("ReffIndice"), IdIndiceProjet, Rs.Fields("PI_Indice"), "", Rs!Version, True)
MyWord.ActiveDocument.SaveAs PathPl
MyWord.ActiveDocument.Close , False

 If IdFils <> 0 Then
        sql = "SELECT RqCartouche.* "
        sql = sql & "FROM RqCartouche "
        sql = sql & "WHERE T_indiceProjet.Id=" & IdFils & ";"
        Set RsModifier = Con.OpenRecordSet(sql)
         PathPl2 = PathArchive(TableauPath.Item("PathArchiveAutocad"), "" & RsModifier!Client, "" & RsModifier!CleAc, "" & RsModifier!Pieces, "LIEC", Rs.Fields("ReffIndice"), RsModifier!Id, RsModifier.Fields("PI_Indice"), "", RsModifier!Version, True)
        Racourci "" & PathPl2, "" & PathPl, "DOC"

    End If

MyWord.Quit False
Set MyWord = Nothing
End Sub
