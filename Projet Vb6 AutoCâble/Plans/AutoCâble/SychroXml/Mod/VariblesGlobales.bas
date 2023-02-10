Attribute VB_Name = "VariblesGlobales"
Option Explicit

Public MyExcel As EXCEL.Application
Public Con As New Ado
Public MySIC_TERM As Range
Public MySIC_CONT As Range
Public MyIS As Range
Public MyShell As Range
Global TableauPath As Collection
Global DbCatalogue As String
Global LeCient As String
Global FormBarGrah As Object
Global JobError As Long
Global NbError As Long
Global bool_MiseEnPage As Boolean
Global NmJob As Long
Global MyWorkbook As EXCEL.Workbook
Global IdFils As Long
Public PortraitPaysage As Long

Public numST As Integer
Public numSC As Integer
Public numIS As Integer
Public numS As Integer
Public numE As Integer
Public numW As Integer
Public numWG As Integer
Public Tocount As Integer
Public NBERRORS As Integer
 Public NBERRORSW As Integer
Public BdDateTable As String
Public DbNumPlan As String
Public IsCilent As Boolean
Public IsServeur As Boolean
Public Db As String
Public AutocableDRIVE As String
Public DonneesEntreprise As String
Public DonneesProduction As String
