Attribute VB_Name = "LienConFamille"
Option Explicit

Sub LstConecteur()
Dim Con As New Ado
Dim fs As New FileSystemObject
Dim f
Dim fc
Dim f1
Dim sql As String
Dim Bloc
Dim XYZ(0 To 2) As Double
Dim RsConnecteur As Recordset
Dim RsVoie As Recordset

Dim Colec As Collection
Dim Att
Dim I As Long
XYZ(0) = 1: XYZ(1) = 1: XYZ(2) = 1
On Error Resume Next
'Con.OpenConnetion "Q:\Autocable.mdb"
AdcFileName = OpenNew
AutoApp.Visible = True
'AdcFileName.Application.Visible = True
Set f = fs.GetFolder("Q:\Autocad\Connecteurs")
    Set fc = f.Files
ReDim TableauFichier(fc.Count, 2)
'i = 0
For Each f1 In fc
'    i = i + 1
    If UCase(Right(f1.Name, 4)) = UCase(".dwg") Then
        If InStr(1, f1.Name, "§") = 0 Then
            Set Bloc = FunInsBlock("" & f1.Path, XYZ, "")
            If IsConnecteurs(Bloc.GetAttributes) = True Then
                              
                
                sql = "SELECT T_Lien_Con_Famille.* "
                sql = sql & "FROM T_Lien_Con_Famille "
                sql = sql & "WHERE T_Lien_Con_Famille.Connecteur='" & Trim(Left(f1.Name, Len(f1.Name) - 4)) & "';"
                Set RsConnecteur = Con.OpenRecordSet(sql)
                If RsConnecteur.EOF = True Then
                    RsConnecteur.AddNew
                    RsConnecteur!Connecteur = Trim(Left(f1.Name, Len(f1.Name) - 4))
                    RsConnecteur.Update
                End If
                RsConnecteur.Requery
                 Att = Bloc.GetAttributes
                For I = 0 To UBound(Att)
                    If InStr(UCase(Att(I).TagString), "LIAI") <> 0 Then
                        sql = "SELECT T_Lien_Con_Famille_Voies.Voie, T_Lien_Con_Famille_Voies.Id_T_Lien_Con_Famille "
                        sql = sql & "FROM T_Lien_Con_Famille_Voies "
                        sql = sql & "WHERE T_Lien_Con_Famille_Voies.Voie='" & Replace(Att(I).TagString, "LIAI", "") & "'  "
                        sql = sql & "AND T_Lien_Con_Famille_Voies.Id_T_Lien_Con_Famille=" & RsConnecteur!Id & ";"
                        
                        Set RsVoie = Con.OpenRecordSet(sql)
                        If RsVoie.EOF = True Then
                            RsVoie.AddNew
                            RsVoie!Voie = Replace(Att(I).TagString, "LIAI", "")
                            RsVoie!Id_T_Lien_Con_Famille = RsConnecteur!Id
                            RsVoie.Update
                        End If
                        RsVoie.Requery
                    End If
                    
                    
                Next
            End If
        End If
    '
        Debug.Print f1.Name
        DoEvents
    End If

Next

 

Set fs = Nothing
MsgBox "Fin:"
End Sub
