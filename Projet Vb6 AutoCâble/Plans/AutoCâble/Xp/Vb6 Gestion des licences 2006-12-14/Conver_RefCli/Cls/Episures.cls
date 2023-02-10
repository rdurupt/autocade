VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Episures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public APP As String
Public Epissure As Boolean
Public Flag As Boolean
Public CloseWere As String
Dim Sql As String
Public Id_IndiceProjet As Long

Public MesConnecteur As New Collection
 Dim T_Laia As Liaison
 Dim I As Long
Public Sub AddEpisure(Rsliai As Recordset)
 Dim Num As Long


 

   

Set T_Laia = New Liaison
    
    
T_Laia.APP = "" & Rsliai("App")
T_Laia.App2 = "" & Rsliai("APP2")
T_Laia.LIAI = "" & Rsliai("LIAI")
T_Laia.POS_OUT = "" & Rsliai("POS-OUT")
T_Laia.POS_OUT2 = "" & Rsliai("POS-OUT2")
T_Laia.REF_CONNECTEUR = "" & Rsliai("REF CONNECTEUR")
T_Laia.REF_CONNECTEUR2 = "" & Rsliai("REF CONNECTEUR2")
T_Laia.SECT = "" & Rsliai("SECT")
T_Laia.TEINT = "" & Rsliai("TEINT")
T_Laia.VOI = "" & Rsliai("VOI")
T_Laia.VOI2 = "" & Rsliai("VOI2")
MesConnecteur.Add T_Laia
Set T_Laia = Nothing
End Sub
Public Function GestionEpissure(ColecLiason As Collection, InexColl As Long) As Long
Dim ItexLiason As Long
GestionEpissure = InexColl
Dim I As Long
'   754.AA

        For ItexLiason = 1 To MesConnecteur.Count
      
            If ColecLiason(MesConnecteur(ItexLiason).APP).Epissure = True Then
            
                    Debug.Print MesConnecteur(ItexLiason).REF_CONNECTEUR
                   
                    If MesConnecteur(ItexLiason).Flag = False Then
                        MesConnecteur(ItexLiason).Flag = True
                        ColecLiason(MesConnecteur(ItexLiason).APP).Flag = True
                         RechercheLesEpissure ColecLiason(MesConnecteur(ItexLiason).App2), ColecLiason, I, 1, MesConnecteur(ItexLiason).APP
                         GestionEpissure = 0
                    End If
                    
               
            End If
            If ColecLiason(MesConnecteur(ItexLiason).App2).Epissure = True Then
                
                Debug.Print MesConnecteur(ItexLiason).REF_CONNECTEUR2
                 
                 If MesConnecteur(ItexLiason).Flag = False Then
                 MesConnecteur(ItexLiason).Flag = True
                    ColecLiason(MesConnecteur(ItexLiason).App2).Flag = True
                    RechercheLesEpissure ColecLiason(MesConnecteur(ItexLiason).APP), ColecLiason, I, 2, MesConnecteur(ItexLiason).App2
                    GestionEpissure = 0
                End If
                 
            End If
        
        Next
   

End Function
Public Sub RechercheLesEpissure(Destination As Episures, ColecLiason As Collection, Index As Long, NumApp As Integer, CibleApp As String)
Dim Rs As Recordset

Sql = "SELECT T_Dossier_Contrôle.LIAI, T_Dossier_Contrôle.SECT, T_Dossier_Contrôle.TEINT, T_Dossier_Contrôle.[POS-OUT], "
Sql = Sql & "T_Dossier_Contrôle.APP, T_Dossier_Contrôle.VOI, T_Dossier_Contrôle.[REF CONNECTEUR], T_Dossier_Contrôle.[POS-OUT2],  "
Sql = Sql & "T_Dossier_Contrôle.APP2, T_Dossier_Contrôle.VOI2, T_Dossier_Contrôle.[REF CONNECTEUR2] "
Sql = Sql & "From T_Dossier_Contrôle "
Sql = Sql & "WHERE (T_Dossier_Contrôle.[APP]='" & CibleApp & "' or T_Dossier_Contrôle.[APP2]='" & CibleApp & "') and T_Dossier_Contrôle.Id_IndiceProjet=1156 and (" & CloseWere & ");"
Set Rs = Con.OpenRecordSet(Sql)
    Destination.AjouteLesEpissure Rs, Index, CibleApp
End Sub
Public Sub AjouteLesEpissure(Rs As Recordset, Index As Long, CibleApp As String)
Debug.Print I
Dim INel As Long
Dim Tll As Liaison
Dim MyColec As New Collection
Dim ItexLiason As Long

    For I = 1 To MesConnecteur.Count
        If (MesConnecteur(I).APP = CibleApp) Or (MesConnecteur(I).App2 = CibleApp) Then
            MesConnecteur(I).Flag = True
        End If
        
    Next

While Rs.EOF = False
    If UCase(APP) = UCase("" & Rs!APP) Then
        Set Tll = New Liaison
            Tll.APP = APP
            Tll.POS_OUT = "" & Rs![POS-OUT]
            Tll.REF_CONNECTEUR = Rs![REF CONNECTEUR]
            Tll.VOI = Rs!VOI
        MyColec.Add Tll
        Set Tll = Nothing
    End If
    If UCase(APP) = UCase("" & Rs!App2) Then
        Set Tll = New Liaison
        Tll.APP = APP
        Tll.POS_OUT2 = "" & Rs![POS-OUT2]
        Tll.REF_CONNECTEUR = Rs![REF CONNECTEUR2]
        Tll.VOI = Rs!VOI2
         MyColec.Add Tll
        Set Tll = Nothing
    End If
Rs.MoveNext
Wend
Rs.Requery
While Rs.EOF = False
     If UCase(CibleApp) = UCase("" & Rs!APP) Then
        For ItexLiason = 1 To MyColec.Count
            Set T_Laia = New Liaison
         T_Laia.APP = MyColec(ItexLiason).APP
         T_Laia.POS_OUT = MyColec(ItexLiason).POS_OUT
         T_Laia.REF_CONNECTEUR = MyColec(ItexLiason).REF_CONNECTEUR
         T_Laia.VOI = MyColec(ItexLiason).VOI
         T_Laia.SECT = "" & Rs!SECT
         T_Laia.TEINT = "" & Rs!TEINT
          T_Laia.LIAI = "" & Rs!LIAI
         
        T_Laia.App2 = "" & Rs!App2
         T_Laia.POS_OUT2 = "" & Rs![POS-OUT2]
         T_Laia.REF_CONNECTEUR2 = "" & Rs![REF CONNECTEUR2]
         T_Laia.VOI2 = "" & Rs!VOI2
        
         MesConnecteur.Add T_Laia
        Set T_Laia = Nothing
        Next
     End If
      If UCase(CibleApp) = UCase("" & Rs!App2) Then
        For ItexLiason = 1 To MyColec.Count
            Set T_Laia = New Liaison
             T_Laia.APP = MyColec(ItexLiason).APP
         T_Laia.POS_OUT = MyColec(ItexLiason).POS_OUT
         T_Laia.REF_CONNECTEUR = MyColec(ItexLiason).REF_CONNECTEUR
         T_Laia.VOI = MyColec(ItexLiason).VOI
         T_Laia.SECT = "" & Rs!SECT
         T_Laia.TEINT = "" & Rs!TEINT
          T_Laia.LIAI = "" & Rs!LIAI
         
        T_Laia.App2 = "" & Rs!APP
         T_Laia.POS_OUT2 = "" & Rs![POS-OUT]
         T_Laia.REF_CONNECTEUR2 = "" & Rs![REF CONNECTEUR]
         T_Laia.VOI2 = "" & Rs!VOI
             MesConnecteur.Add T_Laia
        Set T_Laia = Nothing
        Next
     End If
    Rs.MoveNext
Wend
End Sub

Public Sub RetourneTableau(TXT As String)
Dim I As Long
Dim Tableau As String
Dim Tableau2 As String
Tableau = ""
If Flag = False Then
    For I = 1 To MesConnecteur.Count
    
        If MesConnecteur(I).Flag = False Then
        
            If UCase("" & MesConnecteur(I).APP) <> UCase("" & MesConnecteur(I).App2) Then
                Tableau = Trim("" & MesConnecteur(I).LIAI) & ";"
                Tableau = Tableau & Trim("" & MesConnecteur(I).SECT) & ";"
                Tableau = Tableau & Trim("" & MesConnecteur(I).TEINT) & ";"
                Tableau = Tableau & Trim("" & MesConnecteur(I).POS_OUT) & ";"
                Tableau = Tableau & Trim("" & MesConnecteur(I).APP) & ";"
                Tableau = Tableau & Trim("" & MesConnecteur(I).VOI) & ";"
                Tableau = Tableau & Trim("" & MesConnecteur(I).REF_CONNECTEUR) & ";"
                Tableau = Tableau & Trim("" & MesConnecteur(I).POS_OUT2) & ";"
                Tableau = Tableau & Trim("" & MesConnecteur(I).App2) & ";"
                Tableau = Tableau & Trim("" & MesConnecteur(I).VOI2) & ";"
                Tableau = Tableau & Trim("" & MesConnecteur(I).REF_CONNECTEUR2) & vbCrLf
                
                Tableau2 = Trim("" & MesConnecteur(I).LIAI) & ";"
                Tableau2 = Tableau2 & Trim("" & MesConnecteur(I).SECT) & ";"
                Tableau2 = Tableau2 & Trim("" & MesConnecteur(I).TEINT) & ";"
                Tableau2 = Tableau2 & Trim("" & MesConnecteur(I).POS_OUT2) & ";"
                Tableau2 = Tableau2 & Trim("" & MesConnecteur(I).App2) & ";"
                Tableau2 = Tableau2 & Trim("" & MesConnecteur(I).VOI2) & ";"
                Tableau2 = Tableau2 & Trim("" & MesConnecteur(I).REF_CONNECTEUR2) & ";"
                Tableau2 = Tableau2 & Trim("" & MesConnecteur(I).POS_OUT) & ";"
                Tableau2 = Tableau2 & Trim("" & MesConnecteur(I).APP) & ";"
                Tableau2 = Tableau2 & Trim("" & MesConnecteur(I).VOI) & ";"
                Tableau2 = Tableau2 & Trim("" & MesConnecteur(I).REF_CONNECTEUR) & vbCrLf

                If InStr(1, TXT, Tableau) = 0 Then
                    If InStr(1, TXT, Tableau2) = 0 Then
                        If Trim("" & MesConnecteur(I).APP) = Trim(APP) Then
                            TXT = TXT & Tableau
                        Else
                        TXT = TXT & Tableau2
                        End If
                    End If
                End If
            End If
        End If
    Next

End If
'Debug.Print RetourneTableau
End Sub

Private Sub Class_Initialize()
Set MesConnecteur = Nothing
Set MesConnecteur = New Collection
End Sub

Private Sub Class_Terminate()
Set MesConnecteur = Nothing

End Sub
