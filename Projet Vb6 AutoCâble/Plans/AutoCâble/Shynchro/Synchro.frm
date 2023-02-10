VERSION 5.00
Begin VB.Form Synchro 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Synchro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_DblClick()
Dim Con As New Ado
Dim Rs As Recordset
Dim Sql As String
Con.OpenConnetion "\\10.30.0.5\production\Cablage-production\AutoCable\Access\eBoutique\Encelade_CONNECTEURS.mdb"
Sql = "SELECT con_FieldDefs.FieldName, con_FieldDefs.FieldAlias "
Sql = Sql & "From con_FieldDefs "
Sql = Sql & "WHERE con_FieldDefs.FieldAlias<>'' "
Sql = Sql & "ORDER BY con_FieldDefs.FieldAlias;"
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
Con.CloseConnection
End Sub

