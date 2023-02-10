VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDETXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private StrClé As String
Private StrTexte As String

Public Property Let IniClé(Clé As String)
    StrClé = Clé
    Initialize
End Property

Public Property Let IniTexte(texte As String)
    StrTexte = texte
End Property

Public Property Get IniTexte() As String
    IniTexte = StrTexte
End Property

Public Sub DoXor()
    Dim lngC As Long
    Dim intB As Long
    Dim LngN As Long
    For LngN = 1 To Len(StrTexte)
        lngC = Asc(Mid(StrTexte, LngN, 1))
        intB = Int(Rnd * 256)
        Mid(StrTexte, LngN, 1) = Chr(lngC Xor intB)
    Next LngN
End Sub

Private Sub Initialize()
    Dim LngN As Long
    Randomize Rnd(-1)
    For LngN = 1 To Len(StrTexte)
        Randomize Rnd(-Rnd * Asc(Mid(StrClé, LngN, 1)))
    Next LngN
End Sub

Public Function Encrypt(Code As String) As String
Dim strHead As String
Dim strT As String
Dim strA As String
Dim CodageX As New CDETXT
Dim LngN As Long

CodageX.IniClé = Code
CodageX.IniTexte = Code
CodageX.DoXor
strT = CodageX.IniTexte
strHead = strT
CodageX.IniClé = strHead
CodageX.IniTexte = strA
CodageX.DoXor
strA = CodageX.IniTexte
Encrypt = strHead


End Function

Public Function Decrypt(Code As String, Verif As String) As String
Dim strA As String
Dim strT As String
Dim CodageX As New CDETXT
Dim lnlN As Long
Decrypt = True
strT = Code 'Mid(Code, Len("[Crypting_AutoCâble]") + 1, Len(Code) - (Len("[Crypting_AutoCâble]")))
CodageX.IniClé = Verif
CodageX.IniTexte = strT
CodageX.DoXor
 Decrypt = CodageX.IniTexte
'If CodageX.IniTexte <> Verif Then
'    MsgBox "Mot de passe incorrect", vbExclamation + vbOKOnly
'    Decrypt = False
'    Exit Function
'End If

    
End Function

Public Function DefinSerialPass(Useur As String, Serial As String, Pass As String)
On Error Resume Next

Dim x
 Dim Longcode
 Dim y
 Dim P As Integer
 Dim i As Long
x = Right(Useur, 1) 'code pour coder le serial
Longcode = Len(Useur) + 3  'code pour coder le serial
y = Longcode * 10
Serial = Chr(Longcode * 5) & Asc(Right(Useur, 1)) & Asc(Left(Useur, 1)) & Asc(x) & Chr(y) 'code pour coder le serial
Pass = ""
For i = 1 To Len(Useur)


P = Asc(Mid(Useur, i, 1))
    P = 90 - Asc(Mid(Useur, i, 1))
Reprise:
      
    If (P < 48) And (P < 58) Then P = 48 + (48 - P)
     If (P > 57) And (P < 90) Then P = 90 + (90 - P)
     
     
     If (P > 122) Then P = 122 + (122 - P)
  If (P > 90) And (P < 97) Then
    P = 97 - P
       
        GoTo Reprise
        End If
 
    Pass = Pass & Chr(P)
 Debug.Print Pass
'Pass
Next
End Function
Public Function LireJeton() As String
Dim NumFile As Long
Dim DateD As Date
Dim DateE As Date
Dim DateF As Date
Dim Msg2 As String
Dim Fso As New FileSystemObject
'    LireLicence = False
    If Fso.FileExists(App.path & "\AutoCâble.dll") = False Then
        MsgBox "La DLL AutoCâble.DLL non trouvé", vbCritical, "AotoCâble"
        End
    End If
NumFile = FreeFile
Open App.path & "\AutoCâble.dll" For Random As #NumFile
Get #NumFile, , FiledLicence

Close #NumFile
LireJeton = Val(CodageX.Decrypt(FiledLicence.General.NbJeton, "")) - Val(CodageX.Decrypt(FiledLicence.General.NbJetonActif, ""))

If LireJeton < 0 Then LireJeton = 0
End Function


Private Function PeriodeVal(DateD As Date, DateEncours As Date, DateF As Date) As Boolean
PeriodeVal = True
 If (DateD < DateEncours) And (DateEncours > DateF) Then
    PeriodeVal = False
 End If
End Function

Private Sub EcirLicence()
Dim NumFile As Long

FiledLicence.General.NbJeton = Encrypt(FiledLicence.General.NbJeton)
FiledLicence.General.NbJetonActif = Encrypt(FiledLicence.General.NbJetonActif)
FiledLicence.General.DateExecuter = Format(Date, "yyyy-mm-dd")
FiledLicence.General.DateDeb = Encrypt(FiledLicence.General.DateDeb)
 
FiledLicence.General.DateExecuter = Encrypt(FiledLicence.General.DateExecuter)
FiledLicence.General.DateFin = Encrypt(FiledLicence.General.DateFin)
FiledLicence.General.Enregistre = Encrypt("" & FiledLicence.General.Enregistre)
FiledLicence.Record(FiledLicence.Count - 1).PassWord = Encrypt(FiledLicence.Record(FiledLicence.Count - 1).PassWord)
FiledLicence.Record(FiledLicence.Count - 1).Serial = Encrypt(FiledLicence.Record(FiledLicence.Count - 1).Serial)
FiledLicence.Record(FiledLicence.Count - 1).Useur = Encrypt(FiledLicence.Record(FiledLicence.Count - 1).Useur)

FiledLicence.General.Societe = Encrypt("" & FiledLicence.General.Societe)
FiledLicence.General.AficheFrm = Encrypt("" & FiledLicence.General.AficheFrm)
FiledLicence.General.Tous = Encrypt("" & FiledLicence.General.Tous)
Dim Fso As New FileSystemObject
 If Fso.FileExists(App.path & "\AutoCâble.dll") = True Then
    Fso.DeleteFile App.path & "\AutoCâble.dll"
 End If
 Set Fso = Nothing
 NumFile = FreeFile
Open App.path & "\AutoCâble.dll" For Random As #NumFile
Put #NumFile, , FiledLicence

Close #NumFile
End Sub
Function IcrJenton(Jeton As String, NbJeton As String) As Boolean
Dim J As Long
Dim NbJ As Long
J = Val(Jeton)
NbJ = Val(NbJeton)
If NbJ = 0 Then
    IcrJenton = True
    Exit Function
End If
If J = NbJ Then Exit Function
J = J + 1
Jeton = CStr(J)
IcrJenton = True
End Function
Public Sub DcrJenton()
Dim NumFile As Long
NumFile = FreeFile
Open App.path & "\AutoCâble.dll" For Random As #NumFile
Get #NumFile, , FiledLicence

Close #NumFile
FiledLicence.General.DateDeb = CodageX.Decrypt(FiledLicence.General.DateDeb, "")
FiledLicence.General.DateExecuter = CodageX.Decrypt(FiledLicence.General.DateExecuter, "")
FiledLicence.General.DateFin = CodageX.Decrypt(FiledLicence.General.DateFin, "")
FiledLicence.General.NbJeton = CodageX.Decrypt(FiledLicence.General.NbJeton, "")
FiledLicence.General.NbJetonActif = CodageX.Decrypt(FiledLicence.General.NbJetonActif, "")
FiledLicence.General.Enregistre = CodageX.Decrypt(FiledLicence.General.Enregistre, "")
FiledLicence.Record(FiledLicence.Count - 1).PassWord = CodageX.Decrypt(FiledLicence.Record(FiledLicence.Count - 1).PassWord, "")
FiledLicence.Record(FiledLicence.Count - 1).Serial = CodageX.Decrypt(FiledLicence.Record(FiledLicence.Count - 1).Serial, "")
FiledLicence.Record(FiledLicence.Count - 1).Useur = CodageX.Decrypt(FiledLicence.Record(FiledLicence.Count - 1).Useur, "")
FiledLicence.Record(FiledLicence.Count - 1).Enregistre = CodageX.Decrypt(FiledLicence.Record(FiledLicence.Count - 1).Enregistre, "")
FiledLicence.General.Societe = CodageX.Decrypt(FiledLicence.General.Societe, "")
FiledLicence.General.AficheFrm = CodageX.Decrypt(FiledLicence.General.AficheFrm, "")
FiledLicence.General.Tous = CodageX.Decrypt(FiledLicence.General.Tous, "")
If FiledLicence.General.NbJeton <> "0" Then
    FiledLicence.General.NbJetonActif = CStr(Val(FiledLicence.General.NbJetonActif) - 1)
    If Val(FiledLicence.General.NbJetonActif) < 0 Then
         FiledLicence.General.NbJetonActif = "0"
    End If
    EcirLicence
End If
End Sub

