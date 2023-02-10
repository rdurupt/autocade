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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private StrClé As String
Private StrTexte As String

Public Property Let IniClé(Clé As String)
    StrClé = Clé
    Initialize
End Property

Public Property Let IniTexte(Texte As String)
    StrTexte = Texte
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

Public Function DefinSerialPass(UserName As String, Serial As String, Pass As String)
On Error Resume Next

Dim x
 Dim Longcode
 Dim y
 Dim P As Integer
 Dim I As Long
x = Right(UserName, 1) 'code pour coder le serial
Longcode = Len(UserName) + 3  'code pour coder le serial
y = Longcode * 10
Serial = Chr(Longcode * 5) & Asc(Right(UserName, 1)) & Asc(Left(UserName, 1)) & Asc(x) & Chr(y) 'code pour coder le serial
Pass = ""
For I = 1 To Len(UserName)


P = Asc(Mid(UserName, I, 1))
    P = 90 - Asc(Mid(UserName, I, 1))
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
    If Fso.FileExists(App.Path & "\AutoCâble.dll") = False Then
        MsgBox "La DLL AutoCâble.DLL non trouvé", vbCritical, "AotoCâble"
        End
    End If
NumFile = FreeFile
Open App.Path & "\AutoCâble.dll" For Random As #NumFile
Get #NumFile, , FiledLicence

Close #NumFile
LireJeton = CodageX.Decrypt(FiledLicence.General.NbJeton, "")
End Function
Public Function LireLicence() As Boolean
Dim NumFile As Long
Dim DateD As Date
Dim DateE As Date
Dim DateF As Date
Dim Msg2 As String
Dim Fso As New FileSystemObject
    LireLicence = False
    If Fso.FileExists(App.Path & "\AutoCâble.dll") = False Then
        MsgBox "La DLL AutoCâble.DLL non trouvé", vbCritical, "AotoCâble"
        End
    End If
NumFile = FreeFile
Open App.Path & "\AutoCâble.dll" For Random As #NumFile
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
FiledLicence.Record(FiledLicence.Count - 1).UserName = CodageX.Decrypt(FiledLicence.Record(FiledLicence.Count - 1).UserName, "")
FiledLicence.Record(FiledLicence.Count - 1).Enregistre = CodageX.Decrypt(FiledLicence.Record(FiledLicence.Count - 1).Enregistre, "")
FiledLicence.General.Societe = CodageX.Decrypt(FiledLicence.General.Societe, "")
FiledLicence.General.AficheFrm = CodageX.Decrypt(FiledLicence.General.AficheFrm, "")
FiledLicence.General.Tous = CodageX.Decrypt(FiledLicence.General.Tous, "")

If Trim("" & FiledLicence.General.DateExecuter) = "" Then
    FiledLicence.General.DateExecuter = Format(Date, "yyyy-mm-dd")
End If
DateD = CDate(FiledLicence.General.DateDeb)
DateE = CDate(FiledLicence.General.DateExecuter)
DateF = CDate(FiledLicence.General.DateFin)
If IcrJenton(FiledLicence.General.NbJetonActif, FiledLicence.General.NbJeton) = False Then
    If IsServeur = False Then MsgBox "Plus de Licence disponible", vbCritical
    End
End If
If PeriodeVal(DateD, Date, DateF) = False Then
   If FiledLicence.General.Enregistre <> "Yes" Then
   msg = "Vous travaillez en dehors de la plage de validité de votre licence:" & vbCrLf
   msg = msg & "Date de début:" & DateD & vbCrLf
   msg = msg & "Date de fin: " & DateF
   
    If FiledLicence.General.AficheFrm = "Yes" Then
        Msg2 = vbCrLf & vbCrLf & "Voulez vous vous enregistre"
    End If
    If Trim("" & Msg2) <> "" Then
        If MsgBox(msg & Msg2, vbYesNo + vbQuestion) = vbNo Then
            End
        Else
            frmRegistreLicence.Show vbModal
            If FiledLicence.General.Enregistre <> "Yes" Then End
        End If
    Else
        MsgBox msg, vbCritical
        End
    End If
   End If
   
End If
If PeriodeVal(DateD, DateE, Date) = False Then
    MsgBox "Petit malin vous avez changé la date de votre Ordinateur", vbCritical
    End
End If

If FiledLicence.General.AficheFrm = "Yes" And FiledLicence.General.Enregistre <> "Yes" Then
        msg = "Votre licence expire le: " & DateF
        msg = msg & vbCrLf & vbCrLf & "Voulez vous vous enregistre"
        If MsgBox(msg, vbYesNo + vbQuestion) = vbYes Then
        frmRegistreLicence.Show vbModal
        End If
End If
Set Fso = Nothing
EcirLicence
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
FiledLicence.Record(FiledLicence.Count - 1).UserName = Encrypt(FiledLicence.Record(FiledLicence.Count - 1).UserName)

FiledLicence.General.Societe = Encrypt("" & FiledLicence.General.Societe)
FiledLicence.General.AficheFrm = Encrypt("" & FiledLicence.General.AficheFrm)
FiledLicence.General.Tous = Encrypt("" & FiledLicence.General.Tous)
Dim Fso As New FileSystemObject
 If Fso.FileExists(App.Path & "\AutoCâble.dll") = True Then
    Fso.DeleteFile App.Path & "\AutoCâble.dll"
 End If
 Set Fso = Nothing
 NumFile = FreeFile
Open App.Path & "\AutoCâble.dll" For Random As #NumFile
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
Open App.Path & "\AutoCâble.dll" For Random As #NumFile
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
FiledLicence.Record(FiledLicence.Count - 1).UserName = CodageX.Decrypt(FiledLicence.Record(FiledLicence.Count - 1).UserName, "")
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
Public Sub ReInitJenton()
Dim NumFile As Long
NumFile = FreeFile
Open App.Path & "\AutoCâble.dll" For Random As #NumFile
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
FiledLicence.Record(FiledLicence.Count - 1).UserName = CodageX.Decrypt(FiledLicence.Record(FiledLicence.Count - 1).UserName, "")
FiledLicence.Record(FiledLicence.Count - 1).Enregistre = CodageX.Decrypt(FiledLicence.Record(FiledLicence.Count - 1).Enregistre, "")
FiledLicence.General.Societe = CodageX.Decrypt(FiledLicence.General.Societe, "")
FiledLicence.General.AficheFrm = CodageX.Decrypt(FiledLicence.General.AficheFrm, "")
FiledLicence.General.Tous = CodageX.Decrypt(FiledLicence.General.Tous, "")
FiledLicence.General.NbJetonActif = 0
    EcirLicence

End Sub
