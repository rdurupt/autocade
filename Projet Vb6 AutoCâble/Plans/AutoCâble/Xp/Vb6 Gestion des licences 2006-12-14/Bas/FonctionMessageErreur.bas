Attribute VB_Name = "FonctionMessageErreur"
Function MsgErreur(NumErr As Long, Lib1 As String, Lib2 As String, ErrDetail As String) As String
    Select Case NumErr
                Case 1
                    MsgErreur = "Le connecteur : " & Lib1 & " R�f : " & Lib2 & " n''existe pas dans la biblioth�que de blocks."
                Case 2
                    MsgErreur = "L''attribut : " & Lib1 & " de : " & Lib2 & " n''existe pas."
                Case 3
                    MsgErreur = "Impossible d''affecter le fil N� : " & Lib1 & " au connecteur : " & Lib2 & " car celui-ci n''existe pas."
                Case 4
                    MsgErreur = "Erreur de num�rotation pour le connecteur : " & Lib1 & " v�rifiez s''il n''existe pas un trou dans la num�rotaion.  "
                Case 5
                    MsgErreur = "L''attribut : " & Lib1 & " du connecteur : " & Lib2 & " n''existe pas."
                Case 6
                    MsgErreur = "Le composant : " & Lib1 & " R�f : " & Lib2 & " n''existe pas dans la biblioth�que de blocks."
                Case 7
                    MsgErreur = "L''attribut : " & Lib1 & " du composant : " & Lib2 & " n''existe pas."
                Case 8
                    MsgErreur = "Le connecteur : " & Lib1 & " n''existe pas dans le catalogue Client."
                Case 9
                    MsgErreur = "Le Block : " & Lib1 & " n''existe pas dans la biblioth�que de blocks."
                Case 10
                    MsgErreur = "Le fichier :  " & Lib1 & vbCrLf & "est actuellement ouvert par un autre utilisateur  et ne peut pas �tre sauvegard�."
                 Case 11
                    MsgErreur = "Pb Excel :  le fichier EXCEL ne peut pas �tre enregistr�."
                 
    End Select
    MsgErreur = MsgErreur & vbCrLf & "D�tail de l''erreur :"
    MsgErreur = MsgErreur & vbCrLf & "********************************************************************************************"
    MsgErreur = MsgErreur & vbCrLf & MyReplace(ErrDetail)
    MsgErreur = MsgErreur & vbCrLf & "********************************************************************************************"
    MsgErreur = MsgErreur & vbCrLf
    MsgErreur = MsgErreur & vbCrLf
    Debug.Print MsgErreur
    NbError = NbError + 1
End Function

