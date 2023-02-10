<!-- #include file="HE_Upload.asp" -->
<%

Dim page
Dim mode
Dim formBack
Dim foo
Dim win
Dim form
Dim leMode
Dim lePathTo
Dim leOrderBy
Dim leOrderStr
Dim leDisp
Dim FSO
Dim leFolder
Dim Uploader
Dim File
Dim myError
Dim strExtensions

strExtensions = ""
strExtensions = strExtensions & "||"
strExtensions = strExtensions & "GIF"
strExtensions = strExtensions & "||"
strExtensions = strExtensions & "JPEG"
strExtensions = strExtensions & "||"
strExtensions = strExtensions & "JPG"
strExtensions = strExtensions & "||"
strExtensions = strExtensions & "PNG"
strExtensions = strExtensions & "||"
strExtensions = strExtensions & "BMP"
strExtensions = strExtensions & "||"
strExtensions = strExtensions & "DOC"
strExtensions = strExtensions & "||"
strExtensions = strExtensions & "RTF"
strExtensions = strExtensions & "||"
strExtensions = strExtensions & "ZIP"
strExtensions = strExtensions & "||"
strExtensions = strExtensions & "PDF"
strExtensions = strExtensions & "||"
strExtensions = strExtensions & "XLS"
strExtensions = strExtensions & "||"
strExtensions = strExtensions & "CSV"
strExtensions = strExtensions & "||"
strExtensions = strExtensions & "PPT"
strExtensions = strExtensions & "||"
strExtensions = strExtensions & "PPS"
strExtensions = strExtensions & "||"
strExtensions = strExtensions & "TXT"
strExtensions = strExtensions & "||"


page = Request("h_page")
mode = Request("mode")
foo = Request("foo")
win = Request("f_win")
form = Request("f_form")
leMode = Request("h_mode")
Response.Write Request("h_mode")
Response.End
lePathTo = Request("f_PathTo")
leOrderBy = Request("f_OrderBy")
leOrderStr = Request("f_OrderStr")
leDisp = Request("_display")
formBack = Request("h_formback")
imgChange = Request("h_imgchange")

If leMode = "AddFolder" Then
 Dim leNewFolder
 leNewFolder = Request("f_NewFolder")
 Set FSO = Server.CreateObject("Scripting.FileSystemObject")
 If Not FSO.FolderExists(Server.MapPath(lePathTo & leNewFolder & "/")) Then
  FSO.CreateFolder(Server.MapPath(lePathTo & leNewFolder & "/"))
  If page = "img" Then
   Response.Redirect "HE_Images.asp?mode=" & mode & "&f_form=" & form & "&f_win=" & win & "&foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&_display=" & leDisp & "&f_msg=Le repertoire " & leNewFolder & " a bien été crée"
  If page = "DVD" Then
   Response.Redirect "HE_Candidat_Video.asp?mode=" & mode & "&f_form=" & form & "&f_win=" & win & "&foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&_display=" & leDisp & "&f_msg=Le repertoire " & leNewFolder & " a bien été crée"
   
  ElseIf page = "fic" Then
   Response.Redirect "HE_Files.asp?foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "f_msg=Le repertoire " & leNewFolder & " a bien été crée"
  ElseIf page = "bgd" Then
   Response.Redirect "HE_Background.asp?f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "f_msg=Le repertoire " & leNewFolder & " a bien été crée"
  ElseIf page = "cand" Then
   Response.Redirect "HE_Candidat.asp?h_imgchange=" & imgChange & "&h_formback=" & formBack & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "f_msg=Le repertoire " & leNewFolder & " a bien été crée"
  ElseIf page = "cand_fic" Then
   Response.Redirect "HE_Candidat_Files.asp?h_imgchange=" & imgChange & "&h_formback=" & formBack & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "f_msg=Le repertoire " & leNewFolder & " a bien été crée"
  End If
 Else
  Set FSO = Nothing
  If page = "img" Then
   Response.Redirect "HE_Images.asp?mode=" & mode & "&f_form=" & form & "&f_win=" & win & "&foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&_display=" & leDisp & "&f_msg=Ce nom de repertoire est déjà utilisé"
  If page = "DVD" Then
   Response.Redirect "HE_Candidat_Video.asp?mode=" & mode & "&f_form=" & form & "&f_win=" & win & "&foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&_display=" & leDisp & "&f_msg=Ce nom de repertoire est déjà utilisé"
   
  ElseIf page = "fic" Then
   Response.Redirect "HE_Files.asp?foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Ce nom de repertoire est déjà utilisé"
  ElseIf page = "bgd" Then
   Response.Redirect "HE_Background.asp?f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Ce nom de repertoire est déjà utilisé"
  ElseIf page = "cand" Then
   Response.Redirect "HE_Candidat.asp?h_imgchange=" & imgChange & "&h_formback=" & formBack & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Ce nom de repertoire est déjà utilisé"
  ElseIf page = "cand_fic" Then
   Response.Redirect "HE_Candidat_Files.asp?h_imgchange=" & imgChange & "&h_formback=" & formBack & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Ce nom de repertoire est déjà utilisé"
  End If
 End If
ElseIf leMode = "DelFolder" Then
 leFolder = Request("f_folder")
 Set FSO = Server.CreateObject("Scripting.FileSystemObject")
 FSO.DeleteFolder Server.MapPath(lePathTo & leFolder & "/"), True
 Set FSO = Nothing
 If page = "img" Then
  Response.Redirect "HE_Images.asp?mode=" & mode & "&f_form=" & form & "&f_win=" & win & "&foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&_display=" & leDisp & "&f_msg=Le repertoire a bien été effacé"
  If page = "DVD" Then
   Response.Redirect "HE_Candidat_Video.asp?mode=" & mode & "&f_form=" & form & "&f_win=" & win & "&foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&_display=" & leDisp & "&f_msg=Le repertoire a bien été effacé"
  
 ElseIf page = "fic" Then
  Response.Redirect "HE_Files.asp?foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Le repertoire a bien été effacé"
 ElseIf page = "bgd" Then
  Response.Redirect "HE_Background.asp?f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Le repertoire a bien été effacé"
 ElseIf page = "cand" Then
  Response.Redirect "HE_Candidat.asp?h_imgchange=" & imgChange & "&h_formback=" & formBack & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Le repertoire a bien été effacé"
 ElseIf page = "cand_fic" Then
  Response.Redirect "HE_Candidat_Files.asp?h_imgchange=" & imgChange & "&h_formback=" & formBack & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Le repertoire a bien été effacé"
 End If
ElseIf leMode = "DelImage" Then
 Dim lImage
 lImage = Request("f_image")
 Set FSO = Server.CreateObject("Scripting.FileSystemObject")
 FSO.DeleteFile Server.MapPath(lePathTo & lImage), True
 Set FSO = Nothing
 If page = "img" Then
  Response.Redirect "HE_Images.asp?mode=" & mode & "&f_form=" & form & "&f_win=" & win & "&foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&_display=" & leDisp & "&f_msg=L'image a bien été effacée"
  If page = "DVD" Then
   Response.Redirect "HE_Candidat_Video.asp?mode=" & mode & "&f_form=" & form & "&f_win=" & win & "&foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&_display=" & leDisp & "&f_msg=L'image a bien été effacée"
  
 ElseIf page = "bgd" Then
  Response.Redirect "HE_Background.asp?f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=L'image a bien été effacée"
 ElseIf page = "cand" Then
  Response.Redirect "HE_Candidat.asp?h_imgchange=" & imgChange & "&h_formback=" & formBack & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=L'image a bien été effacée"
 ElseIf page = "cand_fic" Then
  Response.Redirect "HE_Candidat_Files.asp?h_imgchange=" & imgChange & "&h_formback=" & formBack & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Le fichier a bien été effacé"
 End If
ElseIf leMode = "DelFile" Then
 Dim leFile
 leFile = Request("f_file")
 Set FSO = Server.CreateObject("Scripting.FileSystemObject")
 FSO.DeleteFile Server.MapPath(lePathTo & leFile), True
 Set FSO = Nothing
 If page = "fic" Then
  Response.Redirect "HE_Files.asp?foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Le fichier a bien été effacé"
 ElseIf page = "cand_fic" Then
  Response.Redirect "HE_Candidat_Files.asp?h_imgchange=" & imgChange & "&h_formback=" & formBack & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Le fichier a bien été effacé"
 End If
ElseIf leMode = "AddImage" Then
 Set Uploader = New FileUploader
 Uploader.Upload()
 myError = 0
 If CInt(Uploader.Files.Count) > 0 Then
  For Each File In Uploader.Files.Items
   If InStr(1, UCase(File.ContentType), "IMAGE/", vbTextCompare) > 0 Then
    If InStr(1, UCase(File.ContentType), "GIF", vbTextCompare) > 0 Or _
     InStr(1, UCase(File.ContentType), "JPEG", vbTextCompare) > 0 Or _
     InStr(1, UCase(File.ContentType), "BMP", vbTextCompare) > 0 Or _
     InStr(1, UCase(File.ContentType), "PNG", vbTextCompare) > 0 Then
     File.SaveToDisk Server.MapPath(lePathTo)
     myError = 0
    Else
     myError = 1
     Exit For
    End If
   Else
    myError = 1
    Exit For
   End If
  Next
 Else
  myError = 2
 End If
 Set Uploader = Nothing
 
ElseIf leMode = "AddDVD" Then
 Set Uploader = New FileUploader
 Uploader.Upload()
 myError = 0
 If CInt(Uploader.Files.Count) > 0 Then
  For Each File In Uploader.Files.Items
   If InStr(1, UCase(File.ContentType), "IMAGE/", vbTextCompare) > 0 Then
    If InStr(1, UCase(File.ContentType), "GIF", vbTextCompare) > 0 Or _
     InStr(1, UCase(File.ContentType), "JPEG", vbTextCompare) > 0 Or _
     InStr(1, UCase(File.ContentType), "BMP", vbTextCompare) > 0 Or _
     InStr(1, UCase(File.ContentType), "PNG", vbTextCompare) > 0 Then
     File.SaveToDisk Server.MapPath(lePathTo)
     myError = 0
    Else
     myError = 1
     Exit For
    End If
   Else
    myError = 1
    Exit For
   End If
  Next
 Else
  myError = 2
 End If
 Set Uploader = Nothing
 
 If myError = 1 Then
  If page = "img" Then
   Response.Redirect "HE_Images.asp?mode=" & mode & "&f_form=" & form & "&f_win=" & win & "&foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&_display=" & leDisp & "&f_msg=Ce type de fichier n'est pas accepté"
  If page = "DVD" Then
   Response.Redirect "HE_Candidat_Video.asp?mode=" & mode & "&f_form=" & form & "&f_win=" & win & "&foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&_display=" & leDisp & "&f_msg=Ce type de fichier n'est pas accepté"
   
  ElseIf page = "bgd" Then
   Response.Redirect "HE_Background.asp?f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Ce type de fichier n'est pas accepté"
  ElseIf page = "cand" Then
   Response.Redirect "HE_Candidat.asp?h_imgchange=" & imgChange & "&h_formback=" & formBack & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Ce type de fichier n'est pas accepté"
  ElseIf page = "cand_fic" Then
   Response.Redirect "HE_Candidat_Files.asp?h_imgchange=" & imgChange & "&h_formback=" & formBack & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Ce type de fichier n'est pas accepté"
  End If
 ElseIf myError = 2 Then
  If page = "img" Then
   Response.Redirect "HE_Images.asp?mode=" & mode & "&f_form=" & form & "&f_win=" & win & "&foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&_display=" & leDisp & "&f_msg=Veuillez vérifier le fichier"
  If page = "DVD" Then
   Response.Redirect "HE_Candidat_Video.asp?mode="  & mode & "&f_form=" & form & "&f_win=" & win & "&foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&_display=" & leDisp & "&f_msg=Veuillez vérifier le fichier"
   
  ElseIf page = "bgd" Then
   Response.Redirect "HE_Background.asp?f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Veuillez vérifier le fichier"
  ElseIf page = "cand" Then
   Response.Redirect "HE_Candidat.asp?h_imgchange=" & imgChange & "&h_formback=" & formBack & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Veuillez vérifier le fichier"
  ElseIf page = "cand_fic" Then
   Response.Redirect "HE_Candidat_Files.asp?h_imgchange=" & imgChange & "&h_formback=" & formBack & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Veuillez vérifier le fichier"
  End If
 Else
  If page = "img" Then
   Response.Redirect "HE_Images.asp?mode=" & mode & "&f_form=" & form & "&f_win=" & win & "&foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&_display=" & leDisp & "&f_msg=L'image a bien été ajoutée"
  If page = "DVD" Then
   Response.Redirect "HE_Candidat_Video.asp?mode=" & mode & "&f_form=" & form & "&f_win=" & win & "&foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&_display=" & leDisp & "&f_msg=L'image a bien été ajoutée"
   
  ElseIf page = "bgd" Then
   Response.Redirect "HE_Background.asp?f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=L'image a bien été ajoutée"
  ElseIf page = "cand" Then
   Response.Redirect "HE_Candidat.asp?h_imgchange=" & imgChange & "&h_formback=" & formBack & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=L'image a bien été ajoutée"
  ElseIf page = "cand_fic" Then
   Response.Redirect "HE_Candidat_Files.asp?h_imgchange=" & imgChange & "&h_formback=" & formBack & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=L'image a bien été ajoutée"
  End If
 End If
ElseIf leMode = "AddFile" Then
 Set Uploader = New FileUploader
 Uploader.Upload()
 myError = 0
 If CInt(Uploader.Files.Count) > 0 Then
  Dim ext
  For Each File In Uploader.Files.Items
   ext = UCase(Right(File.FileName, Len(File.FileName) - InStrRev(File.FileName, ".")))
   If InStr(1, UCase(strExtensions), "||" & ext & "||", vbTextCompare) > 0 Then
    File.SaveToDisk Server.MapPath(lePathTo)
    myError = 0
   Else
    myError = 1
    Exit For
   End If
  Next
 Else
  myError = 2
 End If
 Set Uploader = Nothing
 If page = "fic" Then
	If myError = 1 Then
		Response.Redirect "HE_Files.asp?foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Ce type de fichier n'est pas accepté"
	ElseIf myError = 2 Then
		Response.Redirect "HE_Files.asp?foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Veuillez vérifier le fichier"
	Else
		Response.Redirect "HE_Files.asp?foo=" & foo & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Le fichier a bien été ajouté"
	End If
  ElseIf page = "cand_fic" Then
	If myError = 1 Then
		Response.Redirect "HE_Candidat_Files.asp?h_imgchange=" & imgChange & "&h_formback=" & formBack & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Ce type de fichier n'est pas accepté"
	ElseIf myError = 2 Then
		Response.Redirect "HE_Candidat_Files.asp?h_imgchange=" & imgChange & "&h_formback=" & formBack & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Veuillez vérifier le fichier"
	Else
		Response.Redirect "HE_Candidat_Files.asp?h_imgchange=" & imgChange & "&h_formback=" & formBack & "&f_PathTo=" & Left(lePathTo, Len(lePathTo) - 1) & "&f_OrderBy=" & leOrderBy & "&f_OrderStr=" & leOrderStr & "&f_msg=Le fichier a bien été ajouté"
	End If
  End If
End If

%>