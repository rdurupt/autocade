VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Logistique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Application  As Application
Public Request As Request
Public Session As Session
Public Response As Response
Public Server As Server
Public FormCaption As String
Public FlagHTMLEditor As Boolean
 Public MyConnecSting As String



Public Sub Main()
Response.Write "<script language='javascript' src='MyJava.js'></script>"
Set Con = OpenDb

If Session("NumJob") = "" Then Session("NumJob") = NumeroChrono("NumJob", "NumJob_")
    
   
        Dim Logistique_Page
        Sql = "SELECT T_Societe.* FROM T_Societe WHERE T_Societe.Id=" & Session("Id_User") & " ;"
        'Response.Write sql
        'Response.End
Set Rs = OpenRecordSet(Sql, Con)
If Rs.EOF = False Then
     Session("Id_User") = Replace("" & Rs("Id"), "Vrai", "True")
     Session("User") = Rs("Identifiant")
    Session("Désination FE") = Replace("" & Rs("Désination FE"), "Vrai", "True")
    Session("Type") = Replace("" & Rs("Type"), "Vrai", "True")
    Session("Pièce") = Replace("" & Rs("Pièce"), "Vrai", "True")
    Session("Indice") = Replace("" & Rs("Indice"), "Vrai", "True")
    Session("Commantaire") = Replace("" & Rs("Commantaires"), "Vrai", "True")
    Session("Designation_Piece") = Replace("" & Rs("Désignation Pièce"), "Vrai", "True")
    Session("Famille/FE/V") = Replace("" & Rs("Famille/FE/V"), "Vrai", "True")
    Session("Appro_S") = Replace("" & Rs("Appro_S"), "Vrai", "True")
    Session("Quantité") = Replace("" & Rs("Quantité"), "Vrai", "True")
    Session("Bon_sortie") = Replace("" & Rs("Bon_sortie"), "Vrai", "True")
    Session("Observation_ magasin") = Replace("" & Rs("Observation_ magasin"), "Vrai", "True")
    Session("Appro") = Replace("" & Rs("Appro"), "Vrai", "True")
    Session("Masse") = Replace("" & Rs("Masse"), "Vrai", "True")
    Session("Commandée_le") = Replace("" & Rs("Commandée_le"), "Vrai", "True")
    Session("Collectée_le") = Replace("" & Rs("Collectée_le"), "Vrai", "True")
    Session("Expédiée_le") = Replace("" & Rs("Expédiée_le"), "Vrai", "True")
    Session("Reçue_le") = Replace("" & Rs("Reçue_le"), "Vrai", "True")
    Session("Qt_reçue") = Replace("" & Rs("Qt_reçue"), "Vrai", "True")
    Session("Reste_recevoir") = Replace("" & Rs("Reste_recevoir"), "Vrai", "True")
    Session("N°_ BL") = Replace("" & Rs("N°_ BL"), "Vrai", "True")
    Session("Réaffectée") = Replace("" & Rs("Réaffectée"), "Vrai", "True")
    Session("Qt Réaffectée") = Replace("" & Rs("Qt Réaffectée"), "Vrai", "True")
    Session("provenance") = Replace("" & Rs("provenance"), "Vrai", "True")
    Session("Retournée") = Replace("" & Rs("Retournée"), "Vrai", "True")
    Session("QT_Retournée") = Replace("" & Rs("QT_Retournée"), "Vrai", "True")
    Session("N°_retour") = Replace("" & Rs("N°_retour"), "Vrai", "True")
    Session("Modifiée") = Replace("" & Rs("Modifiée"), "Vrai", "True")
    
    Session("Admin") = Replace("" & Rs("Admin"), "Vrai", "True")
    Session("Sous_Ensseble") = Replace("" & Rs("Sous_Ensseble"), "Vrai", "True")
    Session("Supprimer") = Replace("" & Rs("Supprimer"), "Vrai", "True")
     Session("Nouveau") = Replace("" & Rs("Nouveau"), "Vrai", "True")
     Session("Updatable") = Replace("" & Rs("Updatable"), "Vrai", "True")
     Session("Quantité") = Replace("" & Rs("Quantité"), "Vrai", "True")
     Session("Action") = Replace("" & Rs("Action"), "Vrai", "True")
     Session("Activer") = Replace("" & Rs("Activer"), "Vrai", "True")
     
End If
 If Session("Admin") = "True" Then
 Session("Id_User") = "True"
    Session("Désination FE") = "True"
    Session("Type") = "True"
    Session("Pièce") = "True"
    Session("Indice") = "True"
    Session("Commantaire") = "True"
    Session("Designation_Piece") = "True"
    Session("Famille/FE/V") = "True"
    Session("Appro_S") = "True"
    Session("Quantité") = "True"
    Session("Bon_sortie") = "True"
    Session("Observation_ magasin") = "True"
    Session("Appro") = "True"
    Session("Masse") = "True"
    Session("Commandée_le") = "True"
    Session("Collectée_le") = "True"
    Session("Expédiée_le") = "True"
    Session("Reçue_le") = "True"
    Session("Qt_reçue") = "True"
    Session("Reste_recevoir") = "True"
    Session("N°_ BL") = "True"
    Session("Réaffectée") = "True"
    Session("Qt Réaffectée") = "True"
    Session("provenance") = "True"
    Session("Retournée") = "True"
    Session("QT_Retournée") = "True"
    Session("N°_retour") = "True"
    Session("Modifiée") = "True"
    Session("Admin") = "True"
    Session("Sous_Ensseble") = "True"
    Session("Supprimer") = "True"
     Session("Nouveau") = "True"
     Session("Updatable") = "True"
     Session("Quantité") = "True"
     Session("Action") = "True"
     Session("Activer") = "True"
 End If
Rs.Close
Set Rs = Nothing
          
         Select Case Trim("" & Request("Logistique_Page"))
                Case "1"
                     Response.Write Case1
                     
                Case "2"
                     Response.Write Case2
                Case "3"
                      Response.Write Case3
                 Case "4"
                       Response.Write Case4
                 Case "5"
                       Response.Write Case5
                  Case "6"
                       Response.Write Case6
                  Case "7"
                       Response.Write Case7
                  Case "8"
                       Response.Write Case8
                  Case "9"
                     Response.Write RecapVehicule("T_Vehicule")
                  Case "10"
                     Response.Write RecapSousEnssemble("T_SousEnsembles")
                 Case "11"
                       Response.Write modSEARCH
                 Case "12"
                       Response.Write RecapBR("BR")
                 Case "13"
                       Response.Write RecapBL("BL")
                  Case "14"
                       Response.Write Case14
                Case Else
                    
                   Response.Write modSEARCH
   End Select
    Con.Close
   Set Con = Nothing
  Response.Write "</body>"
 Response.Write "</html>"

End Sub
