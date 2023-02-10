Attribute VB_Name = "FonctionMessageErreur"
Public Function MsgErreur(NumErr As Long, Lib1 As String, Lib2 As String, ErrDetail As String) As String
    Select Case NumErr
                Case 1
                    MsgErreur = "Le connecteur N° : " & Lib1 & " Réf : " & Lib2 & " n''existe pas dans la bibliothèque de blocks."
                Case 2
                    MsgErreur = "L''attribut : " & Lib1 & " de : " & Lib2 & " n''existe pas."
                Case 3
                    MsgErreur = "Impossible d''affecter le fil N° : " & Lib1 & " au connecteur : " & Lib2 & " car celui-ci n''existe pas."
                Case 4
                    MsgErreur = "Erreur de numérotation pour le connecteur : " & Lib1 & " vérifiez s''il n''existe pas un trou dans la numérotaion.  "
                Case 5
                    MsgErreur = "L''attribut : " & Lib1 & " du connecteur : " & Lib2 & " n''existe pas."
    
                 
    End Select
    MsgErreur = MsgErreur & vbCrLf & "Détail de l''erreur :"
    MsgErreur = MsgErreur & vbCrLf & "********************************************"
    MsgErreur = MsgErreur & vbCrLf & MyReplace(ErrDetail)
    MsgErreur = MsgErreur & vbCrLf & "********************************************"
    MsgErreur = MsgErreur & vbCrLf
    MsgErreur = MsgErreur & vbCrLf
    
End Function

