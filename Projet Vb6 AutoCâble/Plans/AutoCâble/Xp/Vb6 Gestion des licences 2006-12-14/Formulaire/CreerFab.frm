VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreerFab 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dossiers de Fabrication :"
   ClientHeight    =   8085
   ClientLeft      =   30
   ClientTop       =   300
   ClientWidth     =   9360
   Icon            =   "CreerFab.dsx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OleObjectBlob   =   "CreerFab.dsx":27A2
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CreerFab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim Noquite As Boolean
Public Affaire As String
Public PieceCLI As String

Public Sub chargement()
'DbNumPlan = MyWorkbookAppli.Worksheets("Configuration").Range("DbChrono").Value
'DbAcces = MyWorkbookAppli.Worksheets("Configuration").Range("DbAcces").Value
'
'Me.txt4 = Affaire
Me.Show vbModal
End Sub



'Private Sub CommandButton10_Click()
'Dim sql As String
'Dim Rs As Recordset
'CherchPices.Charge Me, " LiAutoCadSave <>  Null and NbErr=0 AND Pere=0 and CleAc=" & Val(txt6), True
'Unload CherchPices
'
'sql = "SELECT T_indiceProjet.Id, [PI] & '_' & [PI_Indice] AS Piece FROM T_indiceProjet "
'sql = sql & "WHERE T_indiceProjet.Id=" & Val(Me.Tag) & ";"
'Set Rs = Con.OpenRecordSet(sql)
'If Rs.EOF = False Then
''Me.Pere = Rs!Piece
'End If
''Me.Pere.Tag = Val(Me.Tag)
'Set Rs = Con.CloseRecordSet(Rs)
'
'End Sub

Private Sub CommandButton8_Click()
Noquite = False
Me.Hide

End Sub







Private Sub CommandButton1_Click()
Dim Sql As String
Dim Rs As Recordset
Dim CherchPicesAnnuler As Boolean
CherchPices.Charge Me, "(VerifieDate= Null   and Archiver=false) OR (IdStatus<4  and Archiver=false)"
CherchPicesAnnuler = CherchPices.Annuler
Unload CherchPices
If Me.txt3.Tag = "" Then CherchPicesAnnuler = True

If CherchPicesAnnuler = True Then Exit Sub

IdFils = 0
Sql = "SELECT T_indiceProjet.Pere FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Me.txt3.Tag & ";"
Set Rs = Con.OpenRecordSet(Sql)
IdFils = 0
Me.Tag = Me.txt3.Tag
If Rs!Pere > 0 Then
IdFils = Me.txt3.Tag
    Me.txt3.Tag = Rs!Pere
    Me.Tag = Me.txt3.Tag
End If
Set Rs = Con.CloseRecordSet(Rs)
Maj Me.txt3.Tag
End Sub

Private Sub CommandButton14_Click()
Maj Me.txt3.Tag
End Sub

'Private Sub txt17_Change()
'If Trim("" & Me.txt17.Text) = "" Then
'     Me.txt16 = ""
'Else
'    Me.txt16 = Format(Date, "dd/mm/yyyy")
'End If
'
'
'End Sub



'Private Sub CommandButton13_Click()
'Me.txt1 = ScanFichier.Chargement("XLS", Me.txt1)
'maj txt1
'End Sub

'Private Sub CommandButton14_Click()
'If Val(Me.Tag) <> 0 Then
'maj Me.Tag
'End If
'End Sub

Private Sub CommandButton15_Click()
Noquite = False
'frmAutoc�ble.DesEnabledMenu
Unload Me
End Sub

'Private Sub Croissant_Click()
'Me.Decroissant.Value = False
'End Sub

'Private Sub Decroissant_Click()
'Croissant.Value = False
'End Sub

'Private Sub txt4_Change()
'Dim Sql As String
'Dim I
'Dim RsBaseNum As Recordset
'Dim ChronoAnnee As String
'ChronoAnnee = Format(Date, "yyyy")
''If ConBaseNum.OpenConnetion(MyWorkbookAppli.Worksheets("Configuration").Range("DbChrono").Value) = True Then
'
''txt5.Clear
''txt6.Clear
''txt7.Clear
''txt8.Clear
'txt13.Clear
'txt14.Clear
'txt15.Clear
'txt16.Clear
'
'
'If txt4 = "" Then Exit Sub
'
'
'Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
'Sql = Sql & "Agent_2.[Pr�nom ag] AS Red_P_Nom,  "
'Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
'Sql = Sql & "Agent_1.[Pr�nom ag] AS Verif_P_Nom,  "
'Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
'Sql = Sql & "Agent.[Pr�nom ag] AS Apr_P_Nom,  "
'Sql = Sql & "" & ChronoAnnee & ".[Cl� ty], " & ChronoAnnee & ".[Cl� ac],  "
'Sql = Sql & "" & ChronoAnnee & ".Ann�e, " & ChronoAnnee & ".[Cl� Ch],  "
'Sql = Sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".R�v  "
'Sql = Sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Cl� ap] = Agent.[Cl� ag])  "
'Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Cl� ve] = Agent_1.[Cl� ag])  "
'Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Cl� re] = Agent_2.[Cl� ag] "
'Sql = Sql & "WHERE " & ChronoAnnee & ".[Cl� ac] = " & txt4 & " and " & ChronoAnnee & ".[Cl� ty] = 'PI' "
'Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Cl� Ch] DESC;"
'Set RsBaseNum = Con.OpenRecordSet(Sql)
''
''While RsBaseNum.EOF = False
''
''    txt5.AddItem "" & RsBaseNum![Cl� ty] & "_" & RsBaseNum![Cl� ac] & "_" & RsBaseNum![Ann�e] & _
''    "_" & RsBaseNum![Cl� Ch] & "_" '& RsBaseNum![R�v]
''    txt5.List(txt5.ListCount - 1, 1) = "" & RsBaseNum![Cl� Ch]
''    If Len(Trim(" " & RsBaseNum![Red_P_Nom])) > 0 Then
''        txt5.List(txt5.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & Left(UCase(RsBaseNum![Red_P_Nom]), 1)
''    Else
''         txt5.List(txt5.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom]
''    End If
''    If Len(Trim(" " & RsBaseNum![Verif_P_Nom])) > 0 Then
''         txt5.List(txt5.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & Left(UCase(RsBaseNum![Verif_P_Nom]), 1)
''    Else
''         txt5.List(txt5.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom]
''    End If
''
''   If Len(Trim(" " & RsBaseNum![Apr_P_Nom])) > 0 Then
''         txt5.List(txt5.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & Left(UCase(RsBaseNum![Apr_P_Nom]), 1)
''    Else
''        txt5.List(txt5.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom]
''    End If
''    txt5.List(txt5.ListCount - 1, 5) = "" & RsBaseNum![R�v]
''If txt5.List(txt5.ListCount - 1, 0) = Pi Then txt5.ListIndex = txt5.ListCount - 1
''' PL = ""
''' OU = ""
''' LI = ""
''    RsBaseNum.MoveNext
''Wend
''
'
'
'
'Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
'Sql = Sql & "Agent_2.[Pr�nom ag] AS Red_P_Nom,  "
'Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
'Sql = Sql & "Agent_1.[Pr�nom ag] AS Verif_P_Nom,  "
'Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
'Sql = Sql & "Agent.[Pr�nom ag] AS Apr_P_Nom,  "
'Sql = Sql & "" & ChronoAnnee & ".[Cl� ty], " & ChronoAnnee & ".[Cl� ac],  "
'Sql = Sql & "" & ChronoAnnee & ".Ann�e, " & ChronoAnnee & ".[Cl� Ch],  "
'Sql = Sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".R�v  "
'Sql = Sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Cl� ap] = Agent.[Cl� ag])  "
'Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Cl� ve] = Agent_1.[Cl� ag])  "
'Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Cl� re] = Agent_2.[Cl� ag] "
'Sql = Sql & "WHERE " & ChronoAnnee & ".[Cl� ac] = " & txt4 & " and " & ChronoAnnee & ".[Cl� ty] = 'PL' "
'Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Cl� Ch] DESC;"
'Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)
'
'While RsBaseNum.EOF = False
'    txt6.AddItem "" & RsBaseNum![Cl� ty] & "_" & RsBaseNum![Cl� ac] & "_" & RsBaseNum![Ann�e] & _
'    "_" & RsBaseNum![Cl� Ch] & "_" & RsBaseNum![R�v]
'    txt6.List(txt6.ListCount - 1, 1) = "" & RsBaseNum![Cl� Ch]
'    txt6.List(txt6.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
'    txt6.List(txt6.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
'    txt6.List(txt6.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
'    txt6.List(txt6.ListCount - 1, 5) = " " & RsBaseNum![R�v]
'    If txt6.List(txt6.ListCount - 1, 0) = PL Then txt6.ListIndex = txt6.ListCount - 1
'' PL = ""
'' OU = ""
'' LI = ""
'    RsBaseNum.MoveNext
'Wend
'
'Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
'Sql = Sql & "Agent_2.[Pr�nom ag] AS Red_P_Nom,  "
'Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
'Sql = Sql & "Agent_1.[Pr�nom ag] AS Verif_P_Nom,  "
'Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
'Sql = Sql & "Agent.[Pr�nom ag] AS Apr_P_Nom,  "
'Sql = Sql & "" & ChronoAnnee & ".[Cl� ty], " & ChronoAnnee & ".[Cl� ac],  "
'Sql = Sql & "" & ChronoAnnee & ".Ann�e, " & ChronoAnnee & ".[Cl� Ch],  "
'Sql = Sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".R�v  "
'Sql = Sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Cl� ap] = Agent.[Cl� ag])  "
'Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Cl� ve] = Agent_1.[Cl� ag])  "
'Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Cl� re] = Agent_2.[Cl� ag] "
'Sql = Sql & "WHERE " & ChronoAnnee & ".[Cl� ac] = " & txt4 & " and " & ChronoAnnee & ".[Cl� ty] = 'OU' "
'Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Cl� Ch] DESC;"
'Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)
'
'While RsBaseNum.EOF = False
'    txt7.AddItem "" & RsBaseNum![Cl� ty] & "_" & RsBaseNum![Cl� ac] & "_" & RsBaseNum![Ann�e] & _
'    "_" & RsBaseNum![Cl� Ch] & "_" & RsBaseNum![R�v]
'    txt7.List(txt7.ListCount - 1, 1) = "" & RsBaseNum![Cl� Ch]
'    txt7.List(txt7.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
'    txt7.List(txt7.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
'    txt7.List(txt7.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
'    txt7.List(txt7.ListCount - 1, 5) = " " & RsBaseNum![R�v]
'    If txt7.List(txt7.ListCount - 1, 0) = OU Then txt7.ListIndex = txt7.ListCount - 1
'' PL = ""
'' OU = ""
'' LI = ""
'    RsBaseNum.MoveNext
'Wend
'txt8.Clear
'Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
'Sql = Sql & "Agent_2.[Pr�nom ag] AS Red_P_Nom,  "
'Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
'Sql = Sql & "Agent_1.[Pr�nom ag] AS Verif_P_Nom,  "
'Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
'Sql = Sql & "Agent.[Pr�nom ag] AS Apr_P_Nom,  "
'Sql = Sql & "" & ChronoAnnee & ".[Cl� ty], " & ChronoAnnee & ".[Cl� ac],  "
'Sql = Sql & "" & ChronoAnnee & ".Ann�e, " & ChronoAnnee & ".[Cl� Ch],  "
'Sql = Sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".R�v  "
'Sql = Sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Cl� ap] = Agent.[Cl� ag])  "
'Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Cl� ve] = Agent_1.[Cl� ag])  "
'Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Cl� re] = Agent_2.[Cl� ag] "
'Sql = Sql & "WHERE " & ChronoAnnee & ".[Cl� ac] = " & txt4 & " and " & ChronoAnnee & ".[Cl� ty] = 'LI' "
'Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Cl� Ch] DESC;"
'Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)
'
'While RsBaseNum.EOF = False
'
'     txt8.AddItem "" & RsBaseNum![Cl� ty] & "_" & RsBaseNum![Cl� ac] & "_" & RsBaseNum![Ann�e] & _
'    "_" & RsBaseNum![Cl� Ch] & "_" & RsBaseNum![R�v]
'    txt8.List(txt8.ListCount - 1, 1) = "" & RsBaseNum![Cl� Ch]
'    txt8.List(txt8.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
'    txt8.List(txt8.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
'    txt8.List(txt8.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
'    txt8.List(txt8.ListCount - 1, 5) = " " & RsBaseNum![R�v]
'      If txt8.List(txt8.ListCount - 1, 0) = LI Then txt8.ListIndex = txt8.ListCount - 1
'' PL = ""
'' OU = ""
'' LI = ""
'    For I = 1 To 3
'    Me.Controls("txt1" & CStr(I)).AddItem "" & RsBaseNum![Cl� ty] & "_" & RsBaseNum![Cl� ac] & "_" & RsBaseNum![Ann�e] & _
'    "_" & RsBaseNum![Cl� Ch] & "_" & RsBaseNum![R�v]
'    Me.Controls("txt1" & CStr(I)).List(Me.Controls("txt1" & CStr(I)).ListCount - 1, 1) = "" & RsBaseNum![Cl� Ch]
'    Me.Controls("txt1" & CStr(I)).List(Me.Controls("txt1" & CStr(I)).ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
'    Me.Controls("txt1" & CStr(I)).List(Me.Controls("txt1" & CStr(I)).ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
'    Me.Controls("txt1" & CStr(I)).List(Me.Controls("txt1" & CStr(I)).ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
'    Me.Controls("txt1" & CStr(I)).List(Me.Controls("txt1" & CStr(I)).ListCount - 1, 5) = " " & RsBaseNum![R�v]
'    Next
'    RsBaseNum.MoveNext
'Wend
'
'
'
'Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
'Sql = Sql & "Agent_2.[Pr�nom ag] AS Red_P_Nom,  "
'Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
'Sql = Sql & "Agent_1.[Pr�nom ag] AS Verif_P_Nom,  "
'Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
'Sql = Sql & "Agent.[Pr�nom ag] AS Apr_P_Nom,  "
'Sql = Sql & "" & ChronoAnnee & ".[Cl� ty], " & ChronoAnnee & ".[Cl� ac],  "
'Sql = Sql & "" & ChronoAnnee & ".Ann�e, " & ChronoAnnee & ".[Cl� Ch],  "
'Sql = Sql & "" & ChronoAnnee & ".rv, " & ChronoAnnee & ".R�v  "
'Sql = Sql & "FROM ((" & ChronoAnnee & " INNER JOIN Agent ON " & ChronoAnnee & ".[Cl� ap] = Agent.[Cl� ag])  "
'Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee & ".[Cl� ve] = Agent_1.[Cl� ag])  "
'Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee & ".[Cl� re] = Agent_2.[Cl� ag] "
'Sql = Sql & "WHERE " & ChronoAnnee & ".[Cl� ac] = " & txt4 & " and " & ChronoAnnee & ".[Cl� ty] = 'NC' "
'Sql = Sql & "ORDER BY " & ChronoAnnee & ".[Cl� Ch] DESC;"
'Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)
'
'While RsBaseNum.EOF = False
'    txt16.AddItem "" & RsBaseNum![Cl� ty] & "_" & RsBaseNum![Cl� ac] & "_" & RsBaseNum![Ann�e] & _
'    "_" & RsBaseNum![Cl� Ch] & "_" & RsBaseNum![R�v]
'    txt16.List(txt16.ListCount - 1, 1) = "" & RsBaseNum![Cl� Ch]
'    txt16.List(txt16.ListCount - 1, 1) = "" & RsBaseNum![Cl� Ch]
'    txt16.List(txt16.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
'    txt16.List(txt16.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
'    txt16.List(txt16.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
'    txt16.List(txt16.ListCount - 1, 5) = "" & RsBaseNum![R�v]
'    RsBaseNum.MoveNext
'Wend
'
'
'
'
'
'Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
'Sql = Sql & "Agent_2.[Pr�nom ag] AS Red_P_Nom,  "
'Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
'Sql = Sql & "Agent_1.[Pr�nom ag] AS Verif_P_Nom,  "
'Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
'Sql = Sql & "Agent.[Pr�nom ag] AS Apr_P_Nom,  "
'Sql = Sql & "" & ChronoAnnee_Moins & ".[Cl� ty], " & ChronoAnnee_Moins & ".[Cl� ac],  "
'Sql = Sql & "" & ChronoAnnee_Moins & ".Ann�e, " & ChronoAnnee_Moins & ".[Cl� Ch],  "
'Sql = Sql & "" & ChronoAnnee_Moins & ".rv, " & ChronoAnnee_Moins & ".R�v  "
'Sql = Sql & "FROM ((" & ChronoAnnee_Moins & " INNER JOIN Agent ON " & ChronoAnnee_Moins & ".[Cl� ap] = Agent.[Cl� ag])  "
'Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee_Moins & ".[Cl� ve] = Agent_1.[Cl� ag])  "
'Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee_Moins & ".[Cl� re] = Agent_2.[Cl� ag] "
'Sql = Sql & "WHERE " & ChronoAnnee_Moins & ".[Cl� ac] = " & txt4 & " and " & ChronoAnnee_Moins & ".[Cl� ty] = 'PI' "
'Sql = Sql & "ORDER BY " & ChronoAnnee_Moins & ".[Cl� Ch] DESC;"
'Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)
'
'While RsBaseNum.EOF = False
'
'    txt5.AddItem "" & RsBaseNum![Cl� ty] & "_" & RsBaseNum![Cl� ac] & "_" & RsBaseNum![Ann�e] & _
'    "_" & RsBaseNum![Cl� Ch] & "_" & RsBaseNum![R�v]
'    txt5.List(txt5.ListCount - 1, 1) = "" & RsBaseNum![Cl� Ch]
'    If Len(Trim(" " & RsBaseNum![Red_P_Nom])) > 0 Then
'        txt5.List(txt5.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & Left(UCase(RsBaseNum![Red_P_Nom]), 1)
'    Else
'         txt5.List(txt5.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom]
'    End If
'    If Len(Trim(" " & RsBaseNum![Verif_P_Nom])) > 0 Then
'         txt5.List(txt5.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & Left(UCase(RsBaseNum![Verif_P_Nom]), 1)
'    Else
'         txt5.List(txt5.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom]
'    End If
'
'   If Len(Trim(" " & RsBaseNum![Apr_P_Nom])) > 0 Then
'         txt5.List(txt5.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & Left(UCase(RsBaseNum![Apr_P_Nom]), 1)
'    Else
'        txt5.List(txt5.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom]
'    End If
'    txt5.List(txt5.ListCount - 1, 5) = "" & RsBaseNum![R�v]
'   If txt5.List(txt5.ListCount - 1, 0) = Pi Then txt5.ListIndex = txt5.ListCount - 1
'
'    RsBaseNum.MoveNext
'Wend
'
'Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
'Sql = Sql & "Agent_2.[Pr�nom ag] AS Red_P_Nom,  "
'Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
'Sql = Sql & "Agent_1.[Pr�nom ag] AS Verif_P_Nom,  "
'Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
'Sql = Sql & "Agent.[Pr�nom ag] AS Apr_P_Nom,  "
'Sql = Sql & "" & ChronoAnnee_Moins & ".[Cl� ty], " & ChronoAnnee_Moins & ".[Cl� ac],  "
'Sql = Sql & "" & ChronoAnnee_Moins & ".Ann�e, " & ChronoAnnee_Moins & ".[Cl� Ch],  "
'Sql = Sql & "" & ChronoAnnee_Moins & ".rv, " & ChronoAnnee_Moins & ".R�v  "
'Sql = Sql & "FROM ((" & ChronoAnnee_Moins & " INNER JOIN Agent ON " & ChronoAnnee_Moins & ".[Cl� ap] = Agent.[Cl� ag])  "
'Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee_Moins & ".[Cl� ve] = Agent_1.[Cl� ag])  "
'Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee_Moins & ".[Cl� re] = Agent_2.[Cl� ag] "
'Sql = Sql & "WHERE " & ChronoAnnee_Moins & ".[Cl� ac] = " & txt4 & " and " & ChronoAnnee_Moins & ".[Cl� ty] = 'PL' "
'Sql = Sql & "ORDER BY " & ChronoAnnee_Moins & ".[Cl� Ch] DESC;"
'Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)
'
'While RsBaseNum.EOF = False
'    txt6.AddItem "" & RsBaseNum![Cl� ty] & "_" & RsBaseNum![Cl� ac] & "_" & RsBaseNum![Ann�e] & _
'    "_" & RsBaseNum![Cl� Ch] & "_" & RsBaseNum![R�v]
'    txt6.List(txt6.ListCount - 1, 1) = "" & RsBaseNum![Cl� Ch]
'    txt6.List(txt6.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
'    txt6.List(txt6.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
'    txt6.List(txt6.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
'    txt6.List(txt6.ListCount - 1, 5) = " " & RsBaseNum![R�v]
'    If txt6.List(txt6.ListCount - 1, 0) = PL Then txt6.ListIndex = txt6.ListCount - 1
'    RsBaseNum.MoveNext
'Wend
'
'
'
'
'Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
'Sql = Sql & "Agent_2.[Pr�nom ag] AS Red_P_Nom,  "
'Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
'Sql = Sql & "Agent_1.[Pr�nom ag] AS Verif_P_Nom,  "
'Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
'Sql = Sql & "Agent.[Pr�nom ag] AS Apr_P_Nom,  "
'Sql = Sql & "" & ChronoAnnee_Moins & ".[Cl� ty], " & ChronoAnnee_Moins & ".[Cl� ac],  "
'Sql = Sql & "" & ChronoAnnee_Moins & ".Ann�e, " & ChronoAnnee_Moins & ".[Cl� Ch],  "
'Sql = Sql & "" & ChronoAnnee_Moins & ".rv, " & ChronoAnnee_Moins & ".R�v  "
'Sql = Sql & "FROM ((" & ChronoAnnee_Moins & " INNER JOIN Agent ON " & ChronoAnnee_Moins & ".[Cl� ap] = Agent.[Cl� ag])  "
'Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee_Moins & ".[Cl� ve] = Agent_1.[Cl� ag])  "
'Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee_Moins & ".[Cl� re] = Agent_2.[Cl� ag] "
'Sql = Sql & "WHERE " & ChronoAnnee_Moins & ".[Cl� ac] = " & txt4 & " and " & ChronoAnnee_Moins & ".[Cl� ty] = 'OU' "
'Sql = Sql & "ORDER BY " & ChronoAnnee_Moins & ".[Cl� Ch] DESC;"
'Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)
'
'While RsBaseNum.EOF = False
'    txt7.AddItem "" & RsBaseNum![Cl� ty] & "_" & RsBaseNum![Cl� ac] & "_" & RsBaseNum![Ann�e] & _
'    "_" & RsBaseNum![Cl� Ch] & "_" & RsBaseNum![R�v]
'    txt7.List(txt7.ListCount - 1, 1) = "" & RsBaseNum![Cl� Ch]
'    txt7.List(txt7.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
'    txt7.List(txt7.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
'    txt7.List(txt7.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
'    txt7.List(txt7.ListCount - 1, 5) = " " & RsBaseNum![R�v]
'    If txt7.List(txt7.ListCount - 1, 0) = OU Then txt7.ListIndex = txt7.ListCount - 1
'    RsBaseNum.MoveNext
'Wend
'
'Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
'Sql = Sql & "Agent_2.[Pr�nom ag] AS Red_P_Nom,  "
'Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
'Sql = Sql & "Agent_1.[Pr�nom ag] AS Verif_P_Nom,  "
'Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
'Sql = Sql & "Agent.[Pr�nom ag] AS Apr_P_Nom,  "
'Sql = Sql & "" & ChronoAnnee_Moins & ".[Cl� ty], " & ChronoAnnee_Moins & ".[Cl� ac],  "
'Sql = Sql & "" & ChronoAnnee_Moins & ".Ann�e, " & ChronoAnnee_Moins & ".[Cl� Ch],  "
'Sql = Sql & "" & ChronoAnnee_Moins & ".rv, " & ChronoAnnee_Moins & ".R�v  "
'Sql = Sql & "FROM ((" & ChronoAnnee_Moins & " INNER JOIN Agent ON " & ChronoAnnee_Moins & ".[Cl� ap] = Agent.[Cl� ag])  "
'Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee_Moins & ".[Cl� ve] = Agent_1.[Cl� ag])  "
'Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee_Moins & ".[Cl� re] = Agent_2.[Cl� ag] "
'Sql = Sql & "WHERE " & ChronoAnnee_Moins & ".[Cl� ac] = " & txt4 & " and " & ChronoAnnee_Moins & ".[Cl� ty] = 'LI' "
'Sql = Sql & "ORDER BY " & ChronoAnnee_Moins & ".[Cl� Ch] DESC;"
'
'Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)
'While RsBaseNum.EOF = False
'     txt8.AddItem "" & RsBaseNum![Cl� ty] & "_" & RsBaseNum![Cl� ac] & "_" & RsBaseNum![Ann�e] & _
'    "_" & RsBaseNum![Cl� Ch] & "_" & RsBaseNum![R�v]
'    txt8.List(txt8.ListCount - 1, 1) = "" & RsBaseNum![Cl� Ch]
'    txt8.List(txt8.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
'    txt8.List(txt8.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
'    txt8.List(txt8.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
'    txt8.List(txt8.ListCount - 1, 5) = " " & RsBaseNum![R�v]
'    If txt8.List(txt8.ListCount - 1, 0) = LI Then txt8.ListIndex = txt8.ListCount - 1
'    For I = 1 To 3
'    Me.Controls("txt1" & CStr(I)).AddItem "" & RsBaseNum![Cl� ty] & "_" & RsBaseNum![Cl� ac] & "_" & RsBaseNum![Ann�e] & _
'    "_" & RsBaseNum![Cl� Ch] & "_" & RsBaseNum![R�v]
'    Me.Controls("txt1" & CStr(I)).List(Me.Controls("txt1" & CStr(I)).ListCount - 1, 1) = "" & RsBaseNum![Cl� Ch]
'    Me.Controls("txt1" & CStr(I)).List(Me.Controls("txt1" & CStr(I)).ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
'    Me.Controls("txt1" & CStr(I)).List(Me.Controls("txt1" & CStr(I)).ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
'    Me.Controls("txt1" & CStr(I)).List(Me.Controls("txt1" & CStr(I)).ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
'    Me.Controls("txt1" & CStr(I)).List(Me.Controls("txt1" & CStr(I)).ListCount - 1, 5) = " " & RsBaseNum![R�v]
'    Next
'    RsBaseNum.MoveNext
'Wend
'
'Sql = "SELECT Agent_2.[Nom ag] AS Red_Nom,  "
'Sql = Sql & "Agent_2.[Pr�nom ag] AS Red_P_Nom,  "
'Sql = Sql & "Agent_1.[Nom ag] AS Verif_Nom,  "
'Sql = Sql & "Agent_1.[Pr�nom ag] AS Verif_P_Nom,  "
'Sql = Sql & "Agent.[Nom ag] AS Apr_Nom,  "
'Sql = Sql & "Agent.[Pr�nom ag] AS Apr_P_Nom,  "
'Sql = Sql & "" & ChronoAnnee_Moins & ".[Cl� ty], " & ChronoAnnee_Moins & ".[Cl� ac],  "
'Sql = Sql & "" & ChronoAnnee_Moins & ".Ann�e, " & ChronoAnnee_Moins & ".[Cl� Ch],  "
'Sql = Sql & "" & ChronoAnnee_Moins & ".rv, " & ChronoAnnee_Moins & ".R�v  "
'Sql = Sql & "FROM ((" & ChronoAnnee_Moins & " INNER JOIN Agent ON " & ChronoAnnee_Moins & ".[Cl� ap] = Agent.[Cl� ag])  "
'Sql = Sql & "INNER JOIN Agent AS Agent_1 ON " & ChronoAnnee_Moins & ".[Cl� ve] = Agent_1.[Cl� ag])  "
'Sql = Sql & "INNER JOIN Agent AS Agent_2 ON " & ChronoAnnee_Moins & ".[Cl� re] = Agent_2.[Cl� ag] "
'Sql = Sql & "WHERE " & ChronoAnnee_Moins & ".[Cl� ac] = " & txt4 & " and " & ChronoAnnee_Moins & ".[Cl� ty] = 'NC' "
'Sql = Sql & "ORDER BY " & ChronoAnnee_Moins & ".[Cl� Ch] DESC;"
'Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)
'
'While RsBaseNum.EOF = False
'    txt16.AddItem "" & RsBaseNum![Cl� ty] & "_" & RsBaseNum![Cl� ac] & "_" & RsBaseNum![Ann�e] & _
'    "_" & RsBaseNum![Cl� Ch] & "_" & RsBaseNum![R�v]
'    txt16.List(txt16.ListCount - 1, 1) = "" & RsBaseNum![Cl� Ch]
'    txt16.List(txt16.ListCount - 1, 1) = "" & RsBaseNum![Cl� Ch]
'    txt16.List(txt16.ListCount - 1, 2) = "" & RsBaseNum![Red_Nom] & " " & RsBaseNum![Red_P_Nom]
'    txt16.List(txt16.ListCount - 1, 3) = "" & RsBaseNum![Verif_Nom] & " " & RsBaseNum![Verif_P_Nom]
'    txt16.List(txt16.ListCount - 1, 4) = "" & RsBaseNum![Apr_Nom] & " " & RsBaseNum![Apr_P_Nom]
'    txt16.List(txt16.ListCount - 1, 5) = "" & RsBaseNum![R�v]
'    RsBaseNum.MoveNext
'Wend
'
'
'Set RsBaseNum = ConBaseNum.CloseRecordSet(RsBaseNum)
'
'ConBaseNum.CloseConnection
'End If
'End Sub









Private Sub CommandButton7_Click()
Dim Exec As Boolean

If Trim("" & Me.Tag) = "" Then
    CommandButton1_Click
    Exit Sub
 End If
If MyFormatQRY(txt13) = False Then Exit Sub
If MyFormatQRY(txt14) = False Then Exit Sub
If MyFormatQRY(txt15) = False Then Exit Sub
If MyFormatQRY(txt16) = False Then Exit Sub
For I = 13 To 16
    For I2 = 13 To 16
        If I <> I2 Then
            If Me.Controls("txt" & CStr(I)) = Me.Controls("txt" & CStr(I2)) Then
                MsgBox "Vous devez saisir des valeurs diff�rentes dans les listes d�roulante" & vbCrLf & Me.Controls("txt" & CStr(I)) & " = " & Me.Controls("txt" & CStr(I2)), vbExclamation
                Exit Sub
            End If
        End If
    Next
Next
Dim Fso As New FileSystemObject
If Fso.FileExists(Environ("USERPROFILE") & "\Mes Documents\" & Replace(txt5.Caption, ":", "_", 1) & ".XLS") = True Then
    Fso.DeleteFile Environ("USERPROFILE") & "\Mes Documents\" & Replace(txt5.Caption, ":", "_", 1) & ".XLS"
End If

Set FormBarGrah = Me
If ExporteXls(Environ("USERPROFILE") & "\Mes Documents\" & Replace(txt5.Caption, ":", "_", 1) & ".XLS", CLng(Me.Tag)) = True Then

EnteteClasseurControle = "Contr�le"
bool_MiseEnPage = True
DossierDeFab Environ("USERPROFILE") & "\Mes Documents\" & Replace(txt5.Caption, ":", "_", 1) & ".XLS", CLng(Me.Tag), _
Me.txt1, Me.txt2, Me.txt3, Me.txt4, Me.txt5, Me.txt6, Me.txt7, Me.txt8, Me.txt9, Me.PieceCLI, Me.txt13, Me.txt14, Me.txt16, True, Me.Affaire, Val(Me.Tag)

EnteteClasseurControle = "Fabrication"

DossierDeFab Environ("USERPROFILE") & "\Mes Documents\" & Replace(txt5.Caption, ":", "_", 1) & ".XLS", CLng(Me.Tag), _
Me.txt1, Me.txt2, Me.txt3, Me.txt4, Me.txt5, Me.txt6, Me.txt7, Me.txt8, Me.txt9, Me.PieceCLI, Me.txt13, Me.txt15, Me.txt16, False, Me.Affaire, Val(Me.Tag)
bool_MiseEnPage = False
End If
If Fso.FileExists(Environ("USERPROFILE") & "\Mes Documents\" & Replace(txt5.Caption, ":", "_", 1) & ".XLS") = True Then
    Fso.DeleteFile Environ("USERPROFILE") & "\Mes Documents\" & Replace(txt5.Caption, ":", "_", 1) & ".XLS"
End If
Noquite = False
'frmAutoc�ble.DesEnabledMenu
Unload Me
End Sub

'Private Sub CommandButton3_Click()
'UserForm1.Charger txt3, vbCrLf, "Ensemble:"
'
'End Sub

'Private Sub CommandButton4_Click()
'UserForm1.Charger txt2, ";", "Equipement:", "_"
'End Sub

'Private Sub CommandButton5_Click()
'UserForm1.Charger txt2, " ", "Vagues:", " "
'End Sub

'Private Sub CommandButton6_Click()
'UserForm3.Show vbmodal
'Unload UserForm3
'If Me.txt1 <> "" Then maj Me.txt1
'
'End Sub

'Private Sub CommandButton7_Click()
'Dim Sql As String
'Dim Rs As Recordset
'Dim pose As Long
'Dim txt As String
'Dim I
'Dim AA
'Dim I2
'If ValideChampsTexte(Me, 14) = False Then Exit Sub
'
'For I = 0 To 3
'    If txt8.Text = Me.Controls("txt" & CStr(11 + I)).Text Then
'    AA = Split(Me.Controls("txt" & CStr(11 + I)).Tag, ";")
'        MsgBox AA(0) & " : " & Me.Controls("txt" & CStr(11 + I)).Text & " existe d�j� "
'        Me.Controls("txt" & CStr(11 + I)).SetFocus
'         Exit Sub
'    End If
'
'    For I2 = 0 To 3
'    If I <> I2 Then
'       If Me.Controls("txt" & CStr(11 + I)).Text = Me.Controls("txt" & CStr(11 + I2)).Text Then
'       AA = Split(Me.Controls("txt" & CStr(11 + I2)).Tag, ";")
'       MsgBox AA(0) & " : " & Me.Controls("txt" & CStr(11 + I2)).Text & " existe d�j� "
'        Me.Controls("txt" & CStr(11 + I2)).SetFocus
'        Exit Sub
'       End If
'    End If
'    Next
'Next
'Dim Fso As New FileSystemObject
'If Fso.FolderExists(DosserFab) = False Then Fso.CreateFolder DosserFab
'If Fso.FolderExists(DossierNc) = False Then Fso.CreateFolder DossierNc
'Set Fso = Nothing
'DosserFab = DosserFab & "\"
'DossierNc = DossierNc & "\"
'PageGarde = UCase(txt11)
'FicheEnCours = UCase(txt11)
'ClasseurControle = UCase(txt13)
'FicheNc = UCase(txt14)
'
'
'
'Noquite = False
'Me.Hide
'boolTrieCroissant = True
'ClasseurControle = UCase(txt13)
'EnteteClasseurControle = "Fabrication"
'
'ClasseurXls = UCase(txt1)
'CrateOnglet2
'boolTrieCroissant = False
'
'EnteteClasseurControle = "Contr�le"
'ClasseurControle = UCase(txt12)
'
'ClasseurXls = UCase(txt1)
'CrateOnglet2
'End
'End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = Noquite
End Sub
Sub Maj(Id_Pieces As String)
'If Trim("" & FichierXLS) = "" Then Exit Sub
Dim Rs As Recordset
Dim RsBaseNum As Recordset
Dim Sql As String
Dim indexClient As Long
Dim RqChronoAnnee As String
Dim ChronoAnneeEnCours As String
Dim ChronoAnnee_Moins As String

Dim T_Affaire
Dim I
Dim Liste


'DosserFab = DosserFab & "Dosser Fab"
Debug.Print DosserFab
DossierNc = DossierNc & "02-NC"
Debug.Print DossierNc
'T_Affaire = T_Affaire(UBound(T_Affaire))
'Liste = Split(T_Affaire, ".")
'LI = ""
'LI = Liste(0)
'T_Affaire = Split(T_Affaire, "_")
'
'Affaire = T_Affaire(1)
'Me.txt4 = Affaire
'ConBaseNum.OpenConnetion DbNumPlan
'Con.OpenConnetion DbAcces

Dim RqChronoAnne
Sql = "SELECT RqCartouche.* , [RefPieceClient] & '_' & Trim('' & [Ref_Piece_CLI]) AS PieceCLI "
Sql = Sql & "FROM RqCartouche "
 Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_Pieces & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
PieceCLI = "" & Rs!PieceCLI
End If
Set Rs = Con.CloseRecordSet(Rs)
Noquite = True
RqChronoAnne = "[Chrono Requ�te " & Format(Date, "yyyy]")
ChronoAnneeEnCours = "Chrono" & Format(Date, "yyyy")
ChronoAnnee_Moins = "Chrono" & Val(Format(Date, "yyyy") - 1)


 Client = ""
 Destinataire = ""
 Service = ""
 Vague = ""
 Equipement = ""
 Ensemble = ""
 MASSE = ""
' PieceCLI = ""



PI = ""
PL = ""
OU = ""
Client = ""
'txt10 = ""
'txt2 = ""
'txt3 = ""

Sql = "SELECT T_indiceProjet.CleAc "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "WHERE T_indiceProjet.Id=" & Id_Pieces & ";"

Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = False Then
   Affaire = Rs!CleAc
End If
Set Rs = Con.CloseRecordSet(Rs)

Sql = "SELECT CStr([Cl� ty]) & '_' & CStr([Cl� ac]) & '_' & CStr([Ann�e]) & '_' & CStr([Cl� Ch]) & '_' & CStr([R�v]) AS LI "
Sql = Sql & "FROM " & ChronoAnneeEnCours & " "
Sql = Sql & "WHERE " & ChronoAnneeEnCours & ".[Cl� ty]='LI' AND " & ChronoAnneeEnCours & ".[Cl� ac]=" & Affaire & " "
Sql = Sql & "ORDER BY " & ChronoAnneeEnCours & ".[Cl� ty] DESC;"

ConBaseNum.TYPEBASE = ADO_TYPEBASE
ConBaseNum.SERVER = ADO_SERVER
ConBaseNum.User = ADO_User
ConBaseNum.PassWord = ADO_PassWord
ConBaseNum.BASE = DbNumPlan




ConBaseNum.OpenConnetion
Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)


 Me.txt13.Clear
 Me.txt13.AddItem ""
While RsBaseNum.EOF = False
    Me.txt13.AddItem Trim("" & RsBaseNum!LI)
'        If Trim("" & Rs!Client) = Client Then Me.txt13.ListIndex = Me.txt13.ListCount - 1

    RsBaseNum.MoveNext
Wend


RsBaseNum.Requery

 Me.txt14.Clear
 Me.txt14.AddItem ""
While RsBaseNum.EOF = False
    Me.txt14.AddItem Trim("" & RsBaseNum!LI)
'        If Trim("" & Rs!Client) = Client Then Me.txt13.ListIndex = Me.txt13.ListCount - 1

    RsBaseNum.MoveNext
Wend
RsBaseNum.Requery
Me.txt15.Clear
 Me.txt15.AddItem ""
While RsBaseNum.EOF = False
    Me.txt15.AddItem Trim("" & RsBaseNum!LI)
'        If Trim("" & Rs!Client) = Client Then Me.txt13.ListIndex = Me.txt13.ListCount - 1

    RsBaseNum.MoveNext
Wend


Sql = "SELECT CStr([Cl� ty]) & '_' & CStr([Cl� ac]) & '_' & CStr([Ann�e]) & '_' & CStr([Cl� Ch]) & '_' & CStr([R�v]) AS LI "
Sql = Sql & "FROM " & ChronoAnnee_Moins & " "
Sql = Sql & "WHERE " & ChronoAnnee_Moins & ".[Cl� ty]='LI' AND " & ChronoAnnee_Moins & ".[Cl� ac]=" & Affaire & " "
Sql = Sql & "ORDER BY " & ChronoAnnee_Moins & ".[Cl� ty] DESC;"
Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)
While RsBaseNum.EOF = False
    Me.txt13.AddItem Trim("" & RsBaseNum!LI)
'        If Trim("" & Rs!Client) = Client Then Me.txt13.ListIndex = Me.txt13.ListCount - 1

    RsBaseNum.MoveNext
Wend


RsBaseNum.Requery


While RsBaseNum.EOF = False
    Me.txt14.AddItem Trim("" & RsBaseNum!LI)
'        If Trim("" & Rs!Client) = Client Then Me.txt13.ListIndex = Me.txt13.ListCount - 1

    RsBaseNum.MoveNext
Wend
RsBaseNum.Requery
While RsBaseNum.EOF = False
    Me.txt15.AddItem Trim("" & RsBaseNum!LI)
'        If Trim("" & Rs!Client) = Client Then Me.txt13.ListIndex = Me.txt13.ListCount - 1

    RsBaseNum.MoveNext
Wend
Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)


 
While RsBaseNum.EOF = False
    Me.txt13.AddItem Trim("" & RsBaseNum!LI)
'        If Trim("" & Rs!Client) = Client Then Me.txt13.ListIndex = Me.txt13.ListCount - 1

    RsBaseNum.MoveNext
Wend



Sql = "SELECT CStr([Cl� ty]) & '_' & CStr([Cl� ac]) & '_' & CStr([Ann�e]) & '_' & CStr([Cl� Ch]) & '_' & CStr([R�v]) AS LI "
Sql = Sql & "FROM " & ChronoAnneeEnCours & " "
Sql = Sql & "WHERE " & ChronoAnneeEnCours & ".[Cl� ty]='NC' AND " & ChronoAnneeEnCours & ".[Cl� ac]=" & Affaire & " "
Sql = Sql & "ORDER BY " & ChronoAnneeEnCours & ".[Cl� ty] DESC;"

Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)
 Me.txt16.Clear
 Me.txt16.AddItem ""
While RsBaseNum.EOF = False
    Me.txt16.AddItem Trim("" & RsBaseNum!LI)
'        If Trim("" & Rs!Client) = Client Then Me.txt13.ListIndex = Me.txt13.ListCount - 1

    RsBaseNum.MoveNext
Wend

Sql = "SELECT CStr([Cl� ty]) & '_' & CStr([Cl� ac]) & '_' & CStr([Ann�e]) & '_' & CStr([Cl� Ch]) & '_' & CStr([R�v]) AS LI "
Sql = Sql & "FROM " & ChronoAnnee_Moins & " "
Sql = Sql & "WHERE " & ChronoAnnee_Moins & ".[Cl� ty]='NC' AND " & ChronoAnnee_Moins & ".[Cl� ac]=" & Affaire & " "
Sql = Sql & "ORDER BY " & ChronoAnnee_Moins & ".[Cl� ty] DESC;"

Set RsBaseNum = ConBaseNum.OpenRecordSet(Sql)
 While RsBaseNum.EOF = False
    Me.txt16.AddItem Trim("" & RsBaseNum!LI)
'        If Trim("" & Rs!Client) = Client Then Me.txt13.ListIndex = Me.txt13.ListCount - 1

    RsBaseNum.MoveNext
Wend

Set RsBaseNum = ConBaseNum.CloseRecordSet(RsBaseNum)


'If RsBaseNum.EOF = False Then
'    txt4 = Trim("" & RsBaseNum!NOM) & " " & Trim("" & RsBaseNum!Pr�nom)
'End If

'    txt6.AddItem RsBaseNum![Cl� ac]
'    txt6.List(txt6.ListCount - 1, 1) = Trim("" & RsBaseNum!NOM) & " " & Trim("" & RsBaseNum!Pr�nom)
'      txt6.List(txt6.ListCount - 1, 2) = Trim("" & RsBaseNum!Int)
'
'    RsBaseNum.MoveNext
'Wend
' txt4_Change

ConBaseNum.CloseConnection

End Sub

