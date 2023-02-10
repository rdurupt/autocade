Attribute VB_Name = "FonctionsGlobales"

Private Function Session(Valeur As String) As String

End Function


Function replaceAccent(Txt As String, Optional noEspace As Boolean, Optional NotMajuscule As Boolean) As String
If NotMajuscule = False Then
    Txt = UCase(Txt)
Else
     Txt = Txt
End If
If noEspace = False Then
Txt = Replace(Txt, " ", "_", 1)
End If
Txt = Replace(Txt, UCase("¨"), "_")
Txt = Replace(Txt, UCase("^"), "_")
Txt = Replace(Txt, UCase("é"), "E")
Txt = Replace(Txt, UCase("è"), "E")
Txt = Replace(Txt, UCase("ê"), "E")
Txt = Replace(Txt, UCase("ë"), "E")
Txt = Replace(Txt, UCase("à"), "A")

Txt = Replace(Txt, UCase("â"), "A")
Txt = Replace(Txt, UCase("î"), "I")
Txt = Replace(Txt, UCase("ô"), "O")
Txt = Replace(Txt, UCase("û"), "U")

Txt = Replace(Txt, UCase("ä"), "A")
Txt = Replace(Txt, UCase("ï"), "I")
Txt = Replace(Txt, UCase("ö"), "O")
Txt = Replace(Txt, UCase("ü"), "U")

'txt = Replace(txt, UCase("'"), "_")
Txt = Replace(Txt, Chr(34), "''")
Txt = Replace(Txt, "<", "‹")
Txt = Replace(Txt, ">", "›")
Txt = Replace(Txt, "'", "''")
Txt = Replace(Txt, Chr(34), Chr(34) & Chr(34))
replaceAccent = Txt





End Function
Private Function Request(Valeur As String) As String

End Function

Public Function CrerDevis(numDevis As String, NouContacter As String)
Dim Sql As String
Dim Rs As Recordset
Dim RsSociete As Recordset

Sql = "SELECT t_Devis.*, t_Devis.Id, t_Devis.id_caddie, t_Devis.id_r_social, t_Devis.Id_Menu, t_Devis.id_produit, "
Sql = Sql & "t_Devis.Id_User, t_Devis.date_modif, t_Devis.qte_produit, t_Devis.TypePiece, t_Devis.RefProduit,  "
Sql = Sql & "t_Devis.numDevis, t_Devis.PrixU_produit, t_Devis.Designation, t_Devis.creation, t_Devis.Id_Avoire,  "
Sql = Sql & "t_Devis.QtsDispo, t_Devis.Reste, t_Devis.MyDate, t_Devis.TVA, t_Devis.PrixRevient, t_Devis.Remise "
Sql = Sql & "FROM t_Devis IN '"
Sql = Sql & TableauPath("Eb_Menu")
Sql = Sql & "' "
Sql = Sql & "WHERE t_Devis.numDevis='" & numDevis & "';"
Set Rs = Con.OpenRecordSet(Sql)

Sql = "SELECT  t_pays.label_pays,t_R_Social.SocieteId, t_R_Social.Id_Contact, t_R_Social.LastName, t_R_Social.FirstName, t_R_Social.id_livraison,  "
Sql = Sql & "t_R_Social.fld2, t_R_Social.fld3, t_R_Social.Address, t_R_Social.Zip, t_R_Social.City, t_R_Social.id_pays,  "
Sql = Sql & "t_R_Social.Email, t_R_Social.CId, t_R_Social.phone, t_R_Social.Portable, t_R_Social.fax, t_R_Social.NumIndiveClient,  "
Sql = Sql & "t_R_Social.Listerouge, t_R_Social.NbJoursPaie, t_R_Social.Remise, t_R_Social.SockEncelade,  "
Sql = Sql & "t_R_Social.DateCreation "
Sql = Sql & "FROM t_R_Social INNER JOIN t_pays ON t_R_Social.id_pays = t_pays.id_pays IN '"
Sql = Sql & TableauPath("Eb_Menu")
Sql = Sql & "' "
Sql = Sql & "WHERE t_R_Social.Id_Contact=" & Rs!id_r_social & ";"
Set RsSociete = Con.OpenRecordSet(Sql)

CrerDevis = ""
CrerDevis = CrerDevis & vbCrLf & ""
CrerDevis = CrerDevis & vbCrLf & "<html>"
'CrerDevis = CrerDevis & vbCrLf & "<script language='javascript' src='../Portal_Java/FunJava.js'></script>"
'CrerDevis = CrerDevis & vbCrLf & "<SCRIPT LANGUAGE='JAVASCRIPT'>"
'CrerDevis = CrerDevis & vbCrLf & "function ValideQtsPrix(L){"
'CrerDevis = CrerDevis & vbCrLf & "var PrixU;"
'CrerDevis = CrerDevis & vbCrLf & "PrixU =  document.forms['MAJPrixLiveraison'].elements['L_PrixU_produit_'+L];"
'CrerDevis = CrerDevis & vbCrLf & "var QTS;"
'CrerDevis = CrerDevis & vbCrLf & "QTS =  document.forms['MAJPrixLiveraison'].elements['L_qte_produit_'+L];"
'CrerDevis = CrerDevis & vbCrLf & "var THT;"
'CrerDevis = CrerDevis & vbCrLf & "THT =  document.forms['MAJPrixLiveraison'].elements['L_Prix_THT_produit_'+L];"
'CrerDevis = CrerDevis & vbCrLf & "if (IsNumeric2(Remplacer(PrixU.value,',','.'))==false){"
'CrerDevis = CrerDevis & vbCrLf & "    alert('Vous devez saisir une valeur numérique.');"
'CrerDevis = CrerDevis & vbCrLf & "    Var PrixU_Av"
'CrerDevis = CrerDevis & vbCrLf & "    PrixU_Av = document.forms['MAJPrixLiveraison'].elements['L_PrixU_produit_Avant_'+L];"
'CrerDevis = CrerDevis & vbCrLf & "   PrixU.value=PrixU_Av.value;"
'CrerDevis = CrerDevis & vbCrLf & "   PrixU.focus();"
'CrerDevis = CrerDevis & vbCrLf & "}"
'CrerDevis = CrerDevis & vbCrLf & "THT.Value = Qts.Value * Remplacer(PrixU.Value, ',', '.')"
'CrerDevis = CrerDevis & vbCrLf & "}"
'CrerDevis = CrerDevis & vbCrLf & "function ValideQts(L){"
'CrerDevis = CrerDevis & vbCrLf & "var QtsChange;"
'CrerDevis = CrerDevis & vbCrLf & "QtsChange = document.forms['MAJPrixLiveraison'].elements['L_qte_produit_'+L];"
'CrerDevis = CrerDevis & vbCrLf & "var QTS;"
'CrerDevis = CrerDevis & vbCrLf & "QTS = document.forms['MAJPrixLiveraison'].elements['L_QTS_'+L];"
'CrerDevis = CrerDevis & vbCrLf & "var Reste;"
'CrerDevis = CrerDevis & vbCrLf & "Reste = document.forms['MAJPrixLiveraison'].elements['txt_Reste_'+L];"
'CrerDevis = CrerDevis & vbCrLf & "if (!testEntier(QtsChange.value)){"
'CrerDevis = CrerDevis & vbCrLf & "    alert('Vous devez saisir un nombre Entier.');"
'CrerDevis = CrerDevis & vbCrLf & "   Err='Err';"
'CrerDevis = CrerDevis & vbCrLf & "}"
'CrerDevis = CrerDevis & vbCrLf & "var Err='Err'"
'CrerDevis = CrerDevis & vbCrLf & " Err=funMath(QtsChange.value,Reste.value,'-',false)"
'CrerDevis = CrerDevis & vbCrLf & "  if (Err<0) {"
'CrerDevis = CrerDevis & vbCrLf & "      if (funMath(Reste.value,Err,'+',false)<0){"
'CrerDevis = CrerDevis & vbCrLf & "     alert('La quantité saisie doit tenir compte de la quantité déjà livrée.');"
'CrerDevis = CrerDevis & vbCrLf & "  Err='Err';"
'CrerDevis = CrerDevis & vbCrLf & "      }"
'CrerDevis = CrerDevis & vbCrLf & "  }"
'CrerDevis = CrerDevis & vbCrLf & "  if (Err=='Err') {"
'CrerDevis = CrerDevis & vbCrLf & "    QtsChange.value=QTS.value;"
'CrerDevis = CrerDevis & vbCrLf & "   }"
'CrerDevis = CrerDevis & vbCrLf & "   ValideQtsPrix(L);"
'CrerDevis = CrerDevis & vbCrLf & "  "
'CrerDevis = CrerDevis & vbCrLf & "FunReste2(L);"
'CrerDevis = CrerDevis & vbCrLf & "}"
'CrerDevis = CrerDevis & vbCrLf & "function Isdouble(Obj){"
'CrerDevis = CrerDevis & vbCrLf & "  var MyObj;"
'CrerDevis = CrerDevis & vbCrLf & "  MyObj = document.forms['MAJPrixLiveraison'].elements[Obj];"
'CrerDevis = CrerDevis & vbCrLf & "   MyObj.value=Remplacer(jsTrim(MyObj.value),',','.');"
'CrerDevis = CrerDevis & vbCrLf & "  if (!testFloat(MyObj.value)) {"
'CrerDevis = CrerDevis & vbCrLf & "       alert('Vous devez saisir une valeur numérique.');"
'CrerDevis = CrerDevis & vbCrLf & "       MyObj.value='';"
'CrerDevis = CrerDevis & vbCrLf & "       return;"
'CrerDevis = CrerDevis & vbCrLf & "   }"
'CrerDevis = CrerDevis & vbCrLf & "}"
'CrerDevis = CrerDevis & vbCrLf & "function IsNumeric2(sText){"
'CrerDevis = CrerDevis & vbCrLf & "   var ValidChars = '0123456789.';"
'CrerDevis = CrerDevis & vbCrLf & "   var IsNumber=true;"
'CrerDevis = CrerDevis & vbCrLf & "   var Char;"
'CrerDevis = CrerDevis & vbCrLf & "  "
'CrerDevis = CrerDevis & vbCrLf & " "
'CrerDevis = CrerDevis & vbCrLf & "   for (i = 0; i < sText.length && IsNumber == true; i++)"
'CrerDevis = CrerDevis & vbCrLf & "     {"
'CrerDevis = CrerDevis & vbCrLf & "      Char = sText.charAt(i);"
'CrerDevis = CrerDevis & vbCrLf & "      if (ValidChars.indexOf(Char) == -1)"
'CrerDevis = CrerDevis & vbCrLf & "         {"
'CrerDevis = CrerDevis & vbCrLf & "         IsNumber = false;"
'CrerDevis = CrerDevis & vbCrLf & "        }"
'CrerDevis = CrerDevis & vbCrLf & "      }"
'CrerDevis = CrerDevis & vbCrLf & "   return IsNumber;"
'CrerDevis = CrerDevis & vbCrLf & "    "
'CrerDevis = CrerDevis & vbCrLf & "  }"
'CrerDevis = CrerDevis & vbCrLf & ""
'CrerDevis = CrerDevis & vbCrLf & ""
'CrerDevis = CrerDevis & vbCrLf & ""
'CrerDevis = CrerDevis & vbCrLf & "function ValideLivraison(Txt){"
'CrerDevis = CrerDevis & vbCrLf & "this.document.MAJLiveraison.TypeLiv.value=Txt;"
'CrerDevis = CrerDevis & vbCrLf & "this.document.MAJLiveraison.ACTION='Contact.asp';"
'CrerDevis = CrerDevis & vbCrLf & "this.document.MAJLiveraison.submit();"
'CrerDevis = CrerDevis & vbCrLf & "}"
'CrerDevis = CrerDevis & vbCrLf & "function SuppirmerArticle(L){"
'CrerDevis = CrerDevis & vbCrLf & "var L_RefProduit;"
'CrerDevis = CrerDevis & vbCrLf & "var IdSource;"
'CrerDevis = CrerDevis & vbCrLf & "var IdCible;"
'CrerDevis = CrerDevis & vbCrLf & "IdSource= document.forms['MAJPrixLiveraison'].elements['L_'+L];"
'CrerDevis = CrerDevis & vbCrLf & "IdCible= document.forms['MAJPrixLiveraison'].elements['IdSup'];"
'CrerDevis = CrerDevis & vbCrLf & "L_RefProduit= document.forms['MAJPrixLiveraison'].elements['L_RefProduit_'+L];"
'CrerDevis = CrerDevis & vbCrLf & "var Qts;"
'CrerDevis = CrerDevis & vbCrLf & "var Reste;"
'CrerDevis = CrerDevis & vbCrLf & "Qts= document.forms['MAJPrixLiveraison'].elements['L_qte_produit_'+L];"
'CrerDevis = CrerDevis & vbCrLf & "Reste = document.forms['MAJPrixLiveraison'].elements['Reste_'+L];"
'CrerDevis = CrerDevis & vbCrLf & "Qts.value=Remplacer(Qts.value,',','.');"
'CrerDevis = CrerDevis & vbCrLf & "Reste.value=Remplacer(Reste.value,',','.');"
'CrerDevis = CrerDevis & vbCrLf & "if (Qts.value!==Reste.value) {"
'CrerDevis = CrerDevis & vbCrLf & "  alert('Vous avez déclaré avoir livré des produits sur cette Référence'+L_RefProduit.value+'.\nLa suppression ne pourra pas être effectuée.');"
'CrerDevis = CrerDevis & vbCrLf & "  return;"
'CrerDevis = CrerDevis & vbCrLf & "}"
'CrerDevis = CrerDevis & vbCrLf & "  if (confirm('Voulez vous vraiment  supprimer cette référence :'+L_RefProduit.value+'\n\nAttention les quantités supprimées ne seront pas réintégrées dans la base E boutique .')) {"
'CrerDevis = CrerDevis & vbCrLf & "      IdCible.value=IdSource.value;"
'CrerDevis = CrerDevis & vbCrLf & "      this.document.MAJPrixLiveraison.Maj.value='Sup';"
'CrerDevis = CrerDevis & vbCrLf & "this.document.MAJPrixLiveraison.ACTION='Contact.asp';"
'CrerDevis = CrerDevis & vbCrLf & "this.document.MAJPrixLiveraison.submit();"
'CrerDevis = CrerDevis & vbCrLf & "    }"
'CrerDevis = CrerDevis & vbCrLf & "}"
'CrerDevis = CrerDevis & vbCrLf & "function ValideBl(){"
'CrerDevis = CrerDevis & vbCrLf & "Var Execute"
'CrerDevis = CrerDevis & vbCrLf & "Execute=false;"
'CrerDevis = CrerDevis & vbCrLf & "Var aa"
'CrerDevis = CrerDevis & vbCrLf & "for(i=0;i<this.document.MAJPrixLiveraison.NumLigne.value ;i++) {"
'CrerDevis = CrerDevis & vbCrLf & "I2 = I + 1"
'CrerDevis = CrerDevis & vbCrLf & "eval ('aa = this.document.MAJPrixLiveraison.Dispo_'+i2+'.value;');"
'CrerDevis = CrerDevis & vbCrLf & "if (aa!=0){"
'CrerDevis = CrerDevis & vbCrLf & "Execute=true;"
'CrerDevis = CrerDevis & vbCrLf & "}"
'CrerDevis = CrerDevis & vbCrLf & "}"
'CrerDevis = CrerDevis & vbCrLf & "if (Execute==false){"
'CrerDevis = CrerDevis & vbCrLf & "alert('Vous devez saisir au moins une valeur dans le champs Qté Livrée');"
'CrerDevis = CrerDevis & vbCrLf & "return;"
'CrerDevis = CrerDevis & vbCrLf & "}"
'CrerDevis = CrerDevis & vbCrLf & "var MyForm;"
'CrerDevis = CrerDevis & vbCrLf & "eval('MyForm=this.document.MAJPrixLiveraison;');"
'CrerDevis = CrerDevis & vbCrLf & "MyForm.CreateBl.value='Ok';"
'CrerDevis = CrerDevis & vbCrLf & "MyForm.mode.value='Historique_Bl';"
'CrerDevis = CrerDevis & vbCrLf & "MyForm.ACTION='Contact.asp';"
'CrerDevis = CrerDevis & vbCrLf & "MyForm.submit();"
'CrerDevis = CrerDevis & vbCrLf & "}"
'CrerDevis = CrerDevis & vbCrLf & "function FunReste2(L){"
'CrerDevis = CrerDevis & vbCrLf & "   var Mytxt_Reste;"
'CrerDevis = CrerDevis & vbCrLf & "var QtsCommand;"
'CrerDevis = CrerDevis & vbCrLf & "var QtsDebit;"
'CrerDevis = CrerDevis & vbCrLf & "var QtsChange;"
'CrerDevis = CrerDevis & vbCrLf & "var QtsChange;"
'CrerDevis = CrerDevis & vbCrLf & "MyReste= document.forms['MAJPrixLiveraison'].elements['Reste_'+L];"
'CrerDevis = CrerDevis & vbCrLf & "QtsCommand = document.forms['MAJPrixLiveraison'].elements['L_qte_produit_'+L];"
'CrerDevis = CrerDevis & vbCrLf & "QtsDebit = document.forms['MAJPrixLiveraison'].elements['Dispo_'+L];"
'CrerDevis = CrerDevis & vbCrLf & "Qts_R_Entrer = document.forms['MAJPrixLiveraison'].elements['L_qte_produit_'+L];"
'CrerDevis = CrerDevis & vbCrLf & "Qts_R_Sortie = document.forms['MAJPrixLiveraison'].elements['txt_Reste_'+L];"
'CrerDevis = CrerDevis & vbCrLf & "if (IsNumeric2(QtsDebit.value)==false){"
'CrerDevis = CrerDevis & vbCrLf & "alert('Vous devez saisir une valeur numérique.');"
'CrerDevis = CrerDevis & vbCrLf & "QtsDebit.value=0;"
'CrerDevis = CrerDevis & vbCrLf & "return;"
'CrerDevis = CrerDevis & vbCrLf & "}"
'CrerDevis = CrerDevis & vbCrLf & "Qts_R_Sortie.value=Qts_R_Entrer.value-QtsDebit.value;"
'CrerDevis = CrerDevis & vbCrLf & "  if (Qts_R_Sortie.value+QtsDebit.value.value>QtsCommand.value){"
'CrerDevis = CrerDevis & vbCrLf & "      alert('La Qté Saisie :' + QtsCommand.value + 'est inférieur à la Qté Restante à livrer :' + QtsDebit.value);"
'CrerDevis = CrerDevis & vbCrLf & "QtsCommand.value=(MyReste.value*1) + (QtsDebit.value*1);"
'CrerDevis = CrerDevis & vbCrLf & "Qts_R_Sortie.value=(QtsCommand.value*1) - (QtsDebit.value*1);"
'CrerDevis = CrerDevis & vbCrLf & "   ValideQtsPrix(L);"
'CrerDevis = CrerDevis & vbCrLf & "return;"
'CrerDevis = CrerDevis & vbCrLf & "  }"
'CrerDevis = CrerDevis & vbCrLf & " if (QtsDebit.value<0){"
'CrerDevis = CrerDevis & vbCrLf & "      alert('La Qté Saisie :' + QtsDebit.value + ' < 0');"
'CrerDevis = CrerDevis & vbCrLf & "QtsDebit.value=0;"
'CrerDevis = CrerDevis & vbCrLf & "Qts_R_Sortie.value=Qts_R_Entrer.value;"
'CrerDevis = CrerDevis & vbCrLf & "return;"
'CrerDevis = CrerDevis & vbCrLf & "  }"
'CrerDevis = CrerDevis & vbCrLf & "}"
'CrerDevis = CrerDevis & vbCrLf & "function FunReste(L){"
'CrerDevis = CrerDevis & vbCrLf & "   var Mytxt_Reste;"
'CrerDevis = CrerDevis & vbCrLf & "var QtsCommand;"
'CrerDevis = CrerDevis & vbCrLf & "var QtsDebit;"
'CrerDevis = CrerDevis & vbCrLf & "var QtsChange;"
'CrerDevis = CrerDevis & vbCrLf & "var QtsChange;"
'CrerDevis = CrerDevis & vbCrLf & "QtsCommand = document.forms['MAJPrixLiveraison'].elements['L_qte_produit_'+L];"
'CrerDevis = CrerDevis & vbCrLf & "QtsDebit = document.forms['MAJPrixLiveraison'].elements['Dispo_'+L];"
'CrerDevis = CrerDevis & vbCrLf & "Qts_R_Entrer = document.forms['MAJPrixLiveraison'].elements['Reste_'+L];"
'CrerDevis = CrerDevis & vbCrLf & "Qts_R_Sortie = document.forms['MAJPrixLiveraison'].elements['txt_Reste_'+L];"
'CrerDevis = CrerDevis & vbCrLf & "if (IsNumeric2(QtsDebit.value)==false){"
'CrerDevis = CrerDevis & vbCrLf & "alert('Vous devez saisir une valeur numérique.');"
'CrerDevis = CrerDevis & vbCrLf & "QtsDebit.value=0;"
'CrerDevis = CrerDevis & vbCrLf & "return;"
'CrerDevis = CrerDevis & vbCrLf & "}"
'CrerDevis = CrerDevis & vbCrLf & "Qts_R_Sortie.value=QtsCommand.value-QtsDebit.value;"
'CrerDevis = CrerDevis & vbCrLf & "  if (Qts_R_Sortie.value+Qts_R_Entrer.value<0){"
'CrerDevis = CrerDevis & vbCrLf & "      alert('Vous avez changé la quantité initiale de la commande \n et vous n´avez pas validé votre choix\nOu La Qté Saisi :' + QtsDebit.value + 'est supérieur à la Qté Restante à livrer :' + QtsCommand);"
'CrerDevis = CrerDevis & vbCrLf & "QtsDebit.value=0;"
'CrerDevis = CrerDevis & vbCrLf & "Qts_R_Sortie.value=Qts_R_Entrer.value;"
'CrerDevis = CrerDevis & vbCrLf & "return;"
'CrerDevis = CrerDevis & vbCrLf & "  }"
'CrerDevis = CrerDevis & vbCrLf & "  if (QtsDebit.value<0){"
'CrerDevis = CrerDevis & vbCrLf & "      alert('La Qté Saisie :' + QtsDebit.value + ' < 0');"
'CrerDevis = CrerDevis & vbCrLf & "QtsDebit.value=0;"
'CrerDevis = CrerDevis & vbCrLf & "Qts_R_Sortie.value=Qts_R_Entrer.value;"
'CrerDevis = CrerDevis & vbCrLf & "return;"
'CrerDevis = CrerDevis & vbCrLf & "  }"
'CrerDevis = CrerDevis & vbCrLf & "}"
'CrerDevis = CrerDevis & vbCrLf & "</SCRIPT>"
CrerDevis = CrerDevis & vbCrLf & "<head>"
CrerDevis = CrerDevis & vbCrLf & "<title>Commande N° " & numDevis & "</title>"
CrerDevis = CrerDevis & vbCrLf & "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
CrerDevis = CrerDevis & vbCrLf & "<link rel='stylesheet' href='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "encelade.css'></head>"
CrerDevis = CrerDevis & vbCrLf & ""
CrerDevis = CrerDevis & vbCrLf & "<body bgcolor='#FFFFFF' text='#000000'>"
'CrerDevis = CrerDevis & vbCrLf & "<SCRIPT LANGUAGE='JAVASCRIPT'>"
'CrerDevis = CrerDevis & vbCrLf & "function Imprimer(){"
'CrerDevis = CrerDevis & vbCrLf & "document.getElementById('PRINT').style.display = 'none';"
'CrerDevis = CrerDevis & vbCrLf & "window.print();"
'CrerDevis = CrerDevis & vbCrLf & "   document.getElementById('PRINT').style.display ='block';"
'CrerDevis = CrerDevis & vbCrLf & ""
'CrerDevis = CrerDevis & vbCrLf & "}"
'CrerDevis = CrerDevis & vbCrLf & "function AppercuBl() {"
'CrerDevis = CrerDevis & vbCrLf & "window.open('contactSansMenu.asp?mode=Hisorique_Bl_Visu&RefCommande=CD_39112_1&Liv=BL_39128_1&ApprecuBL=Ok');"
'CrerDevis = CrerDevis & vbCrLf & "}"
CrerDevis = CrerDevis & vbCrLf & "</SCRIPT>"
CrerDevis = CrerDevis & vbCrLf & "<div align='right'>"
'CrerDevis = CrerDevis & vbCrLf & "<table width='15%'><tr>"
'CrerDevis = CrerDevis & vbCrLf & "<td width='47'>"
'CrerDevis = CrerDevis & vbCrLf & "<table border='0' background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
'CrerDevis = CrerDevis & vbCrLf & "Boutoncentre.bmp ' cellpadding='0' cellspacing='0' width='100%' id='PRINT'>"
'CrerDevis = CrerDevis & vbCrLf & "    <tr>"
'CrerDevis = CrerDevis & vbCrLf & "        <td width='5%' align='left' ><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
'CrerDevis = CrerDevis & vbCrLf & "BoutonLeft.bmp ' border='0'></td>"
'CrerDevis = CrerDevis & vbCrLf & ""
'CrerDevis = CrerDevis & vbCrLf & "        <td align='center'><a href='javascript:Imprimer();' border='0'>Imprimer</a></td>"
'CrerDevis = CrerDevis & vbCrLf & ""
'CrerDevis = CrerDevis & vbCrLf & "        <td width='5%' align='right'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
'CrerDevis = CrerDevis & vbCrLf & "Boutonrith.bmp ' border='0'></td>"
' CrerDevis = CrerDevis & vbCrLf & "   </tr>"
'CrerDevis = CrerDevis & vbCrLf & "</table></td>"
CrerDevis = CrerDevis & vbCrLf & "</tr></table>"
CrerDevis = CrerDevis & vbCrLf & "<center><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "logo.gif ' border='0' ></center><br>"
CrerDevis = CrerDevis & vbCrLf & "<table width='95%' border='0' cellpadding='0' cellspacing='0' align='center'>"
CrerDevis = CrerDevis & vbCrLf & "<tr><td width='1'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "hg.gif ' width='17' height='22'></td>"
CrerDevis = CrerDevis & vbCrLf & "<td background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "bgh.gif '><table width='100%' border='0' cellpadding='0' cellspacing='0'>"
CrerDevis = CrerDevis & vbCrLf & "<tr><td background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "bgh2.gif '>'"
CrerDevis = CrerDevis & vbCrLf & "<font face='Arial, Helvetica, sans-serif' size='1'><b><font size='2' color='#868AB4'>R&eacute;capitulatif de votre Commande </font></b></font>"
CrerDevis = CrerDevis & vbCrLf & "</td> <td><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "bgh3.gif ' width='16' height='22'></td>"
CrerDevis = CrerDevis & vbCrLf & "</tr></table></td><td width='1'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "hd.gif ' width='18' height='22'></td>"
CrerDevis = CrerDevis & vbCrLf & "</tr><tr>"
CrerDevis = CrerDevis & vbCrLf & "<td background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "bgg.gif ' width='1'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "bgg.gif ' width='17' height='8'></td>"
CrerDevis = CrerDevis & vbCrLf & "<td bgcolor='#EBE3D7'>"
CrerDevis = CrerDevis & vbCrLf & "<table border=0 width='100%' cellpadding='2' cellspacing='0' class='smallheader'>"
CrerDevis = CrerDevis & vbCrLf & "</Table>"
CrerDevis = CrerDevis & vbCrLf & "<table border='0' background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "bandebg.gif ' cellpadding='0' cellspacing='0' width='100%'>"
CrerDevis = CrerDevis & vbCrLf & "    <tr>"
CrerDevis = CrerDevis & vbCrLf & "        <td width='5%' align='left' ><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "bandeg.gif ' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "        <td class='ecomtitrecaddie2' >R&eacute;f&eacute;rence &agrave; rappeler  </td>"
CrerDevis = CrerDevis & vbCrLf & "        <td width='5%' align='right'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "banded.gif ' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "    </tr>"
CrerDevis = CrerDevis & vbCrLf & "</table>"
CrerDevis = CrerDevis & vbCrLf & "<table border=0 width='100%' cellpadding='0' cellspacing='1'>"
CrerDevis = CrerDevis & vbCrLf & "    <tr  bgcolor=#FFFFFF height='10'>"
CrerDevis = CrerDevis & vbCrLf & "       <td>Num&eacute;ro de Client</td>"
CrerDevis = CrerDevis & vbCrLf & "        <td>"
CrerDevis = CrerDevis & vbCrLf & Format("" & RsSociete!SocieteId, "000#")
 Dim RsCliFac As Recordset
 Set RsCliFac = RsEboutiqueAddFac
CrerDevis = CrerDevis & vbCrLf & "</td>"
CrerDevis = CrerDevis & vbCrLf & "    </tr>"
CrerDevis = CrerDevis & vbCrLf & "    <tr  bgcolor=#FFFFFF height='10'>"
CrerDevis = CrerDevis & vbCrLf & "        <td>Num&eacute;ro de Commande</td>"
CrerDevis = CrerDevis & vbCrLf & "        <td>"
CrerDevis = CrerDevis & vbCrLf & NumCommadEboutique
CrerDevis = CrerDevis & vbCrLf & "</td>"
CrerDevis = CrerDevis & vbCrLf & "    </tr>"
CrerDevis = CrerDevis & vbCrLf & "    <tr  bgcolor=#FFFFFF height='10'>"
CrerDevis = CrerDevis & vbCrLf & "        <td>N°de Commande Encelade</td>"
CrerDevis = CrerDevis & vbCrLf & "        <td>"
CrerDevis = CrerDevis & vbCrLf & numDevis
CrerDevis = CrerDevis & vbCrLf & "</td>"
CrerDevis = CrerDevis & vbCrLf & "    </tr>"
CrerDevis = CrerDevis & vbCrLf & "   <tr  bgcolor=#FFFFFF height='10'>"
CrerDevis = CrerDevis & vbCrLf & "        <td>En Date du </td>"
CrerDevis = CrerDevis & vbCrLf & "        <td>"
CrerDevis = CrerDevis & vbCrLf & Format(Date, "dd/mm/yyyy")
CrerDevis = CrerDevis & vbCrLf & "</td>"
CrerDevis = CrerDevis & vbCrLf & "   </tr>"
CrerDevis = CrerDevis & vbCrLf & "    <tr  bgcolor=#FFFFFF height='10'>"
CrerDevis = CrerDevis & vbCrLf & "        <td>Demandeur</td>"
CrerDevis = CrerDevis & vbCrLf & "        <td>Autocâble</td>"
CrerDevis = CrerDevis & vbCrLf & "    </tr>"
'CrerDevis = CrerDevis & vbCrLf & "    <tr  bgcolor=#FFFFFF height='10'>"
'CrerDevis = CrerDevis & vbCrLf & "        <td>N°de Livraison</td>"
'CrerDevis = CrerDevis & vbCrLf & "        <td>"
'CrerDevis = CrerDevis & vbCrLf & "BL_39128_1"
'CrerDevis = CrerDevis & vbCrLf & "</td>"
''CrerDevis = CrerDevis & vbCrLf & "    </tr>"
'CrerDevis = CrerDevis & vbCrLf & "    <tr  bgcolor=#FFFFFF height='10'>"
'CrerDevis = CrerDevis & vbCrLf & "        <td>En Date du</td>"
'CrerDevis = CrerDevis & vbCrLf & "       <td>"
'CrerDevis = CrerDevis & vbCrLf & "15/02/2007"
'CrerDevis = CrerDevis & vbCrLf & "</td>"
CrerDevis = CrerDevis & vbCrLf & "    </tr>"
CrerDevis = CrerDevis & vbCrLf & "</table>"
CrerDevis = CrerDevis & vbCrLf & "<br>"
CrerDevis = CrerDevis & vbCrLf & "<table border='0' background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "bandebg.gif ' cellpadding='0' cellspacing='0' width='100%'>"
CrerDevis = CrerDevis & vbCrLf & "    <tr>"
CrerDevis = CrerDevis & vbCrLf & "       <td width='5%' align='left' ><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "bandeg.gif ' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "        <td class='ecomtitrecaddie2' >Nous Contacter</td>"
CrerDevis = CrerDevis & vbCrLf & "        <td width='5%' align='right'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "banded.gif ' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "    </tr>"
CrerDevis = CrerDevis & vbCrLf & "</table>"
CrerDevis = CrerDevis & vbCrLf & "<table border=0 width='100%' cellpadding='0' cellspacing='1'>"
CrerDevis = CrerDevis & vbCrLf & "   <tr  bgcolor=#FFFFFF height='10'>"
CrerDevis = CrerDevis & vbCrLf & "       <td>Service technique </td>"
CrerDevis = CrerDevis & vbCrLf & "        <td><A href='" & NouContacter & "'>" & NouContacter & "</A></td>"
CrerDevis = CrerDevis & vbCrLf & "  </tr>"
CrerDevis = CrerDevis & vbCrLf & "</table>"
CrerDevis = CrerDevis & vbCrLf & "<br>"
CrerDevis = CrerDevis & vbCrLf & "<table border='0' background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "bandebg.gif ' cellpadding='0' cellspacing='0' width='100%'>"
CrerDevis = CrerDevis & vbCrLf & "    <tr>"
CrerDevis = CrerDevis & vbCrLf & "        <td width='5%' align='left' ><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "bandeg.gif ' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "       <td class='ecomtitrecaddie2' >Coordonn&eacute;es du demandeur</td>"
CrerDevis = CrerDevis & vbCrLf & "        <td width='5%' align='right'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "banded.gif ' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "    </tr>"
CrerDevis = CrerDevis & vbCrLf & "</table>"
CrerDevis = CrerDevis & vbCrLf & "<table border=0 width='100%' cellpadding='0' cellspacing='1'>"
CrerDevis = CrerDevis & vbCrLf & "   <tr  bgcolor=#FFFFFF height='10'>"
CrerDevis = CrerDevis & vbCrLf & "        <td>Nom / Pr&eacute;nom</td>"
CrerDevis = CrerDevis & vbCrLf & "        <td>" & RsCliFac!pNom & " " & RsCliFac!Nom & "</td>"
CrerDevis = CrerDevis & vbCrLf & "   </tr>"
CrerDevis = CrerDevis & vbCrLf & "    "
CrerDevis = CrerDevis & vbCrLf & "    <tr  bgcolor=#FFFFFF height='10'>"
CrerDevis = CrerDevis & vbCrLf & "        <td>Soci&eacute;t&eacute;</td>"
CrerDevis = CrerDevis & vbCrLf & "        <td >" & RsCliFac!RaisonSocialLiv & "</td>"
CrerDevis = CrerDevis & vbCrLf & "    </tr>"
CrerDevis = CrerDevis & vbCrLf & ""
CrerDevis = CrerDevis & vbCrLf & "    "
CrerDevis = CrerDevis & vbCrLf & "    <tr  bgcolor=#FFFFFF height='10'>"
CrerDevis = CrerDevis & vbCrLf & "        <td >Adresse</td>"
CrerDevis = CrerDevis & vbCrLf & "        <td>" & RsCliFac!AddFac1 & "<br>" & RsCliFac!AddFac2 & "<br>" & RsCliFac!AddFac3 & "</td>"
CrerDevis = CrerDevis & vbCrLf & "   </tr>"
CrerDevis = CrerDevis & vbCrLf & ""
CrerDevis = CrerDevis & vbCrLf & "    <tr  bgcolor=#FFFFFF height='10'>"
CrerDevis = CrerDevis & vbCrLf & "        <td>CP</td>"
CrerDevis = CrerDevis & vbCrLf & "        <td>" & RsCliFac!CpLiFac & "</td>"
CrerDevis = CrerDevis & vbCrLf & "    </tr>"
CrerDevis = CrerDevis & vbCrLf & ""
CrerDevis = CrerDevis & vbCrLf & "    <tr  bgcolor=#FFFFFF height='10'>"
CrerDevis = CrerDevis & vbCrLf & "        <td >Ville / Pays</td>"
CrerDevis = CrerDevis & vbCrLf & "        <td>" & RsCliFac!VilleFac & " / " & RsSociete!label_pays & "</td>"
CrerDevis = CrerDevis & vbCrLf & "    </tr>"
CrerDevis = CrerDevis & vbCrLf & "    "
CrerDevis = CrerDevis & vbCrLf & "    <tr  bgcolor=#FFFFFF height='10'>"
CrerDevis = CrerDevis & vbCrLf & "        <td>Email</td>"
CrerDevis = CrerDevis & vbCrLf & "        <td> <A href='" & RsSociete!EMail & "'>" & RsSociete!EMail & "</A></td>"
CrerDevis = CrerDevis & vbCrLf & "    </tr>"
CrerDevis = CrerDevis & vbCrLf & "    <tr  bgcolor=#FFFFFF height='10'>"
CrerDevis = CrerDevis & vbCrLf & "        <td>T&eacute;l&eacute;phone / Fax</td>"
CrerDevis = CrerDevis & vbCrLf & "        <td>" & RsCliFac!TelFac & "/ " & RsCliFac!faxfac & "</td>"
CrerDevis = CrerDevis & vbCrLf & "    </tr>"
Set RsSociete = Con.CloseRecordSet(RsSociete)
Set RsCliFac = Con.CloseRecordSet(RsCliFac)
CrerDevis = CrerDevis & vbCrLf & "</table>"
CrerDevis = CrerDevis & vbCrLf & "<br>"
CrerDevis = CrerDevis & vbCrLf & "<table border='0' background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "bandebg.gif ' cellpadding='0' cellspacing='0' width='100%'>"
CrerDevis = CrerDevis & vbCrLf & "    <tr>"
CrerDevis = CrerDevis & vbCrLf & "        <td width='5%' align='left' ><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "bandeg.gif ' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "        <td class='ecomtitrecaddie2' > Adresse de livraison</td>"
CrerDevis = CrerDevis & vbCrLf & "       <td width='5%' align='right'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "banded.gif ' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "    </tr>"
CrerDevis = CrerDevis & vbCrLf & "</table>"
Dim RsLiv As Recordset
Set RsLiv = RsEboutiqueCliLivEncelade
CrerDevis = CrerDevis & vbCrLf & "<table border=0 width='100%' cellpadding='0' cellspacing='1'>"
CrerDevis = CrerDevis & vbCrLf & "    <tr  bgcolor=#FFFFFF height='10'>"
CrerDevis = CrerDevis & vbCrLf & "        <td >B&eacute;n&eacute;ficiaire</td>"
CrerDevis = CrerDevis & vbCrLf & "        <td>" & RsLiv!beneficiare_liv & "</td>"
CrerDevis = CrerDevis & vbCrLf & "    </tr>"
CrerDevis = CrerDevis & vbCrLf & ""
CrerDevis = CrerDevis & vbCrLf & "    <tr  bgcolor=#FFFFFF height='10'>"
CrerDevis = CrerDevis & vbCrLf & "        <td >Adresse</td>"
CrerDevis = CrerDevis & vbCrLf & "        <td>" & RsLiv!adresse_liv & "</td>"
CrerDevis = CrerDevis & vbCrLf & "    </tr>"
CrerDevis = CrerDevis & vbCrLf & ""
CrerDevis = CrerDevis & vbCrLf & "    <tr  bgcolor=#FFFFFF height='10'>"
CrerDevis = CrerDevis & vbCrLf & "        <td>CP</td>"
CrerDevis = CrerDevis & vbCrLf & "        <td>" & RsLiv!cp_liv & "</td>"
CrerDevis = CrerDevis & vbCrLf & "    </tr>"
CrerDevis = CrerDevis & vbCrLf & ""
CrerDevis = CrerDevis & vbCrLf & "    <tr  bgcolor=#FFFFFF height='10'>"
CrerDevis = CrerDevis & vbCrLf & "        <td>Ville</td>"
CrerDevis = CrerDevis & vbCrLf & "       <td>" & RsLiv!ville_liv & "</td>"
CrerDevis = CrerDevis & vbCrLf & "    </tr>"
CrerDevis = CrerDevis & vbCrLf & "    <tr  bgcolor=#FFFFFF height='10'>"
CrerDevis = CrerDevis & vbCrLf & "        <td>Mode de livraison</td>"
CrerDevis = CrerDevis & vbCrLf & "        <td>"
Set RsLiv = Con.CloseRecordSet(RsLiv)
CrerDevis = CrerDevis & vbCrLf & EboutiqueEboutiqueGetDefault("LivraisonComtoire", "Je désire  que ma commande me soit remise au comptoir.", TableauPath("Eb_Menu"))
CrerDevis = CrerDevis & vbCrLf & "</td>"
CrerDevis = CrerDevis & vbCrLf & "    </tr>"
CrerDevis = CrerDevis & vbCrLf & "</table>"
CrerDevis = CrerDevis & vbCrLf & "<br>"
CrerDevis = CrerDevis & vbCrLf & "<table border='0' background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "bandebg.gif ' cellpadding='0' cellspacing='0' width='100%'>"
CrerDevis = CrerDevis & vbCrLf & "    <tr>"
CrerDevis = CrerDevis & vbCrLf & "        <td width='5%' align='left' ><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "bandeg.gif ' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "        <td class='ecomtitrecaddie2' >D&eacute;tail de la Commande</td>"
CrerDevis = CrerDevis & vbCrLf & "        <td width='5%' align='right'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "banded.gif ' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "    </tr>"
CrerDevis = CrerDevis & vbCrLf & "</table>"
CrerDevis = CrerDevis & vbCrLf & ""
CrerDevis = CrerDevis & vbCrLf & "<table border=0 width='100%' cellpadding='2' cellspacing='0' class='smallheader'>"
CrerDevis = CrerDevis & vbCrLf & "</tr></Table>"
CrerDevis = CrerDevis & vbCrLf & "<table border='0' cellspacing='3' cellpadding='0' width='100%'>"
CrerDevis = CrerDevis & vbCrLf & "    <tr>"
CrerDevis = CrerDevis & vbCrLf & "        <td class='TextTrEntete' valign='center' align='center'><table border='0' background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_bg.gif' cellpadding='0' cellspacing='0' >"
CrerDevis = CrerDevis & vbCrLf & "<tr>"
CrerDevis = CrerDevis & vbCrLf & "<td align='left'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_gauche.gif' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "<td class='EcomTitreCaddie'> D&eacute;signation</td>"
CrerDevis = CrerDevis & vbCrLf & "<td align='right'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_droite.gif' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "</tr></table></td>"
CrerDevis = CrerDevis & vbCrLf & "        <td class='TextTrEntete' valign='center' align='center'><table border='0' background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_bg.gif' cellpadding='0' cellspacing='0' >"
CrerDevis = CrerDevis & vbCrLf & "<tr>"
CrerDevis = CrerDevis & vbCrLf & "<td align='left'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_gauche.gif' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "<td class='EcomTitreCaddie'> R&eacute;f&eacute;rence</td>"
CrerDevis = CrerDevis & vbCrLf & "<td align='right'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_droite.gif' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "</tr></table></td>"
 CrerDevis = CrerDevis & vbCrLf & "       <td class='TextTrEntete' valign='center' align='center'><table border='0' background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_bg.gif' cellpadding='0' cellspacing='0' >"
CrerDevis = CrerDevis & vbCrLf & "<tr>"
CrerDevis = CrerDevis & vbCrLf & "<td align='left'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_gauche.gif' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "<td class='EcomTitreCaddie'> Qt&eacute;</td>"
CrerDevis = CrerDevis & vbCrLf & "<td align='right'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_droite.gif' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "</tr></table></td>"
CrerDevis = CrerDevis & vbCrLf & "        <td class='TextTrEntete' valign='center' align='center'><table border='0' background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_bg.gif' cellpadding='0' cellspacing='0' >"
CrerDevis = CrerDevis & vbCrLf & "<tr>"
CrerDevis = CrerDevis & vbCrLf & "<td align='left'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_gauche.gif' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "<td class='EcomTitreCaddie'> P. U.H.T. €</td>"
CrerDevis = CrerDevis & vbCrLf & "<td align='right'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_droite.gif' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "</tr></table></td>"
CrerDevis = CrerDevis & vbCrLf & "        <td class='TextTrEntete' valign='center' align='center'><table border='0' background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_bg.gif' cellpadding='0' cellspacing='0' >"
CrerDevis = CrerDevis & vbCrLf & "<tr>"
CrerDevis = CrerDevis & vbCrLf & "<td align='left'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_gauche.gif' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "<td class='EcomTitreCaddie'> Remise %</td>"
CrerDevis = CrerDevis & vbCrLf & "<td align='right'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_droite.gif' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "</tr></table></td>"
CrerDevis = CrerDevis & vbCrLf & "        <td class='TextTrEntete' valign='center' align='center'><table border='0' background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_bg.gif' cellpadding='0' cellspacing='0' >"
CrerDevis = CrerDevis & vbCrLf & "<tr>"
CrerDevis = CrerDevis & vbCrLf & "<td align='left'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_gauche.gif' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "<td class='EcomTitreCaddie'> P.H.T. €</td>"
CrerDevis = CrerDevis & vbCrLf & "<td align='right'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_droite.gif' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "</tr></table></td>"
CrerDevis = CrerDevis & vbCrLf & "       <td class='TextTrEntete' valign='center' align='center'><table border='0' background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_bg.gif' cellpadding='0' cellspacing='0' >"
CrerDevis = CrerDevis & vbCrLf & "<tr>"
CrerDevis = CrerDevis & vbCrLf & "<td align='left'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_gauche.gif' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "<td class='EcomTitreCaddie'> T.V.A. €</td>"
CrerDevis = CrerDevis & vbCrLf & "<td align='right'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_droite.gif' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "</tr></table></td>"
CrerDevis = CrerDevis & vbCrLf & "        <td class='TextTrEntete' valign='center' align='center'><table border='0' background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_bg.gif' cellpadding='0' cellspacing='0' >"
CrerDevis = CrerDevis & vbCrLf & "<tr>"
CrerDevis = CrerDevis & vbCrLf & "<td align='left'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_gauche.gif' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "<td class='EcomTitreCaddie'> P.T.T.C. €</td>"
CrerDevis = CrerDevis & vbCrLf & "<td align='right'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_droite.gif' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "</tr></table></td>"
CrerDevis = CrerDevis & vbCrLf & "        <td class='TextTrEntete' valign='center' align='center'><table border='0' background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_bg.gif' cellpadding='0' cellspacing='0' >"
CrerDevis = CrerDevis & vbCrLf & "<tr>"
CrerDevis = CrerDevis & vbCrLf & "<td align='left'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_gauche.gif' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "<td class='EcomTitreCaddie'> Qt&eacute; Livr&eacute;e</td>"
CrerDevis = CrerDevis & vbCrLf & "<td align='right'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_droite.gif' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "</tr></table></td>"
CrerDevis = CrerDevis & vbCrLf & "        <td class='TextTrEntete' valign='center' align='center'><table border='0' background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_bg.gif' cellpadding='0' cellspacing='0' >"
CrerDevis = CrerDevis & vbCrLf & "<tr>"
CrerDevis = CrerDevis & vbCrLf & "<td align='left'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_gauche.gif' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "<td class='EcomTitreCaddie'> Reste</td>"
CrerDevis = CrerDevis & vbCrLf & "<td align='right'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & "titre_droite.gif' border='0'></td>"
CrerDevis = CrerDevis & vbCrLf & "</tr></table></td>"
CrerDevis = CrerDevis & vbCrLf & "   </tr>"
CrerDevis = CrerDevis & vbCrLf & "<Form Name='MAJPrixLiveraison'  method='post' >"
CrerDevis = CrerDevis & vbCrLf & "  <input type='hidden' name='ModeFacture' value=''>"
CrerDevis = CrerDevis & vbCrLf & "<input type='hidden' name='CreateBl' value='' >"
CrerDevis = CrerDevis & vbCrLf & "  <input type='hidden' name='ModeFacture' value=''>"
Dim qte_produit As Long
Dim PrixU_produit As Double
Dim Remise As Double
Dim PrxTotalProduit As Double
Dim TVA As Double
Dim TotalTva As Double
Dim TTC As Double
Dim T_PrixU_produit As Double
Dim T_Remise As Double
Dim T_PrxTotalProduit As Double
Dim T_TVA As Double
Dim T_TotalTva As Double
Dim T_TTC As Double

While Rs.EOF = False
CrerDevis = CrerDevis & vbCrLf & "    <tr  bgcolor=#FFFFFF width='100%' >"
qte_produit = Rs!qte_produit
PrixU_produit = Rs!PrixU_produit
Remise = Rs!Remise
PrxTotalProduit = Rs!PrixU_produit * Rs!qte_produit * (1 - Rs!Remise)
TVA = Rs!PrixU_produit * Rs!qte_produit * (1 - Rs!Remise) * (Rs!TVA / 100)
TTC = Rs!PrixU_produit * Rs!qte_produit * (1 - Rs!Remise) * (1 + (Rs!TVA / 100))

T_PrixU_produit = T_PrixU_produit + PrixU_produit
T_Remise = T_Remise + Remise
T_PrxTotalProduit = T_PrxTotalProduit + PrxTotalProduit
T_TVA = T_TVA + TVA
T_TTC = T_TTC + TTC
CrerDevis = CrerDevis & vbCrLf & "          <td align='center' width='20%'>" & Rs!DESIGNATION & "</td>"
CrerDevis = CrerDevis & vbCrLf & "                    "
CrerDevis = CrerDevis & vbCrLf & "          <td align='center' width='10%'>" & Rs!RefProduit & "</td>"
CrerDevis = CrerDevis & vbCrLf & ""
CrerDevis = CrerDevis & vbCrLf & "         <td align='right' >" & qte_produit & "</td>"
CrerDevis = CrerDevis & vbCrLf & "          <td align='center' width='10%'>" & Round(PrixU_produit, 3) & "</td>"
CrerDevis = CrerDevis & vbCrLf & "         <td align='center' width='10%'>" & Round(Remise, 3) & "</td>"
 CrerDevis = CrerDevis & vbCrLf & "         <td align='center' width='10%'>" & Round(PrxTotalProduit, 3) & " </td>"
CrerDevis = CrerDevis & vbCrLf & "          <td align='center' width='10%'>" & Round(TVA, 3) & "</td>"
CrerDevis = CrerDevis & vbCrLf & "          <td align='center' width='10%'>" & Round(TTC, 3) & "</td>"
CrerDevis = CrerDevis & vbCrLf & "          <td align='center' width='10%'>" & Rs!QtsDispo & "</td>"
CrerDevis = CrerDevis & vbCrLf & "          <td align='center' width='10%'>" & Rs!qte_produit - Rs!QtsDispo - Rs!Reste & "</td>"
CrerDevis = CrerDevis & vbCrLf & "          "
CrerDevis = CrerDevis & vbCrLf & "                  </tr>"
Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)
CrerDevis = CrerDevis & vbCrLf & "    <tr  bgcolor=#FFFFFF height='10'>"
CrerDevis = CrerDevis & vbCrLf & "                    "

CrerDevis = CrerDevis & vbCrLf & "          <td align='center' colspan='5'>Total</td>"
CrerDevis = CrerDevis & vbCrLf & "          <td align='center' >" & Round(T_PrxTotalProduit, 3) & "</td>"
CrerDevis = CrerDevis & vbCrLf & "          <td align='center' >" & Round(T_TVA, 3) & "</td>"
CrerDevis = CrerDevis & vbCrLf & "          <td align='center' >" & Round(T_TTC, 3) & "</td>"
CrerDevis = CrerDevis & vbCrLf & "          <td align='center' >&nbsp</td>"
CrerDevis = CrerDevis & vbCrLf & "          <td align='center' >&nbsp</td>"
CrerDevis = CrerDevis & vbCrLf & "                  </tr>"
CrerDevis = CrerDevis & vbCrLf & "</table>"
CrerDevis = CrerDevis & vbCrLf & "<table border=0 width='100%' cellpadding='2' cellspacing='0' class='smallheader'>"
CrerDevis = CrerDevis & vbCrLf & "</Table>"
CrerDevis = CrerDevis & vbCrLf & "          <input type='hidden' name ='NumLigne' value='1'>"
CrerDevis = CrerDevis & vbCrLf & "  <input type='hidden' name='TypeLiv' value=''>"
CrerDevis = CrerDevis & vbCrLf & "  <input type='hidden' name='mode'  value= 'HISTORIQUE_Eboutique_Affiche_Commande'>"
CrerDevis = CrerDevis & vbCrLf & "  <input type='hidden' name='IdSup'  value= ''>"
CrerDevis = CrerDevis & vbCrLf & "  <input type='hidden' name='Maj'  value= 'Prix'>"
CrerDevis = CrerDevis & vbCrLf & "  <input type='hidden' name='RefCommande'  value= 'CD_39112_1' >"
CrerDevis = CrerDevis & vbCrLf & "</Form >"
CrerDevis = CrerDevis & vbCrLf & "</td><td background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "bgd.gif ' width='1'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "bgd.gif ' width='18' height='8'></td>"
CrerDevis = CrerDevis & vbCrLf & "</tr><tr><td width='1'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "bg.gif ' width='17' height='16'></td>"
CrerDevis = CrerDevis & vbCrLf & "<td background='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "bgb.gif '><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "bgb.gif ' width='11' height='16'></td>"
CrerDevis = CrerDevis & vbCrLf & "<td width='1'><img src='" & GetDefault("eboutique_euxia_net", "http://eboutique.euxia.net/Portal_Candidat/Devis/") & ""
CrerDevis = CrerDevis & vbCrLf & "BD.gif ' width='18' height='16'></td>"
CrerDevis = CrerDevis & vbCrLf & "</tr></table>"
CrerDevis = CrerDevis & vbCrLf & "</body>"
CrerDevis = CrerDevis & vbCrLf & "</html>"

End Function
Public Function NumCommadEboutiqueID() As Long
Dim Sql As String
Dim Rs As Recordset
Dim NUMCOM As String
Sql = "SELECT T_Num_Command_Eboutique.NumCommand FROM T_Num_Command_Eboutique;"
Set Rs = Con.OpenRecordSet(Sql)
NUMCOM = "" & Rs!NumCommand
Set Rs = Con.CloseRecordSet(Rs)
Sql = "SELECT  T_Num_Commande.NumCommande, Max(T_Avoire.IdAvoire) AS MaxDeIdAvoire, Max(T_Avoire.NumAvoire) AS MaxDeNumAvoire "
Sql = Sql & "FROM (t_R_Social INNER JOIN Users ON t_R_Social.SocieteId = Users.Id_Societes) INNER JOIN  "
Sql = Sql & "(T_Num_Commande INNER JOIN T_Avoire ON T_Num_Commande.IdNumCommande = T_Avoire.IdNumCommande)  "
Sql = Sql & "ON t_R_Social.SocieteId = T_Num_Commande.IdSociete IN '"
Sql = Sql & TableauPath("Eb_Menu")
Sql = Sql & "' "
Sql = Sql & "Where Users.UserID = " & EboutiqueUserId & "  "
Sql = Sql & "And T_Num_Commande.NumCommande = '" & NUMCOM & "' "
Sql = Sql & "GROUP BY Users.UserID, T_Num_Commande.NumCommande"
Set Rs = Con.OpenRecordSet(Sql)
NumCommadEboutiqueID = Rs!MaxDeIdAvoire
Set Rs = Con.CloseRecordSet(Rs)
End Function
Public Function RsEboutiqueAddFac() As Recordset
Dim Sql As String
Sql = "SELECT T_Raison_Sociale_Societe.*, t_pays.label_pays AS PaysLiv, t_pays_1.label_pays AS PaysFac "
Sql = Sql & "FROM (T_Raison_Sociale_Societe INNER JOIN t_pays AS t_pays_1  "
Sql = Sql & "ON T_Raison_Sociale_Societe.Id_PaysFac = t_pays_1.id_pays)  "
Sql = Sql & "INNER JOIN t_pays ON T_Raison_Sociale_Societe.Id_Pays_Liv = t_pays.id_pays IN  '"
Sql = Sql & TableauPath("Eb_Menu")
Sql = Sql & "' "
Set RsEboutiqueAddFac = Con.OpenRecordSet(Sql)

End Function
Public Function EboutiqueEboutiqueGetDefault(fld, def, DSN)
Dim Rs As Recordset
def = safeEntry(def)
   
    Set Rs = Con.OpenRecordSet("SELECT * FROM Defaults IN '" & DSN & "' WHERE defName = '" & fld & "'")
    If Not Rs.EOF Then
           EboutiqueEboutiqueGetDefault = Trim(Rs("defValue"))
    Else

        Con.Execute "INSERT INTO Defaults(defName,defValue) IN '" & DSN & "' VALUES('" & fld & "','" & def & "')"
        EboutiqueEboutiqueGetDefault = def
    End If
    Set Rs = Con.CloseRecordSet(Rs)
    
End Function

Public Function RsEboutiqueCliLivEncelade() As Recordset
Dim Sql As String
Dim Rs As Recordset
Dim NUMCOM As String
Dim IdUser As Long
IdUser = EboutiqueUserId


Sql = "SELECT Users.* FROM Users IN '"
Sql = Sql & TableauPath("Eb_Menu")
Sql = Sql & "' "
Sql = Sql & "Where Users.UserID =" & IdUser & "; "


Sql = "SELECT Users_1.UserEMail, Users_1.phone, Users_1.fax, t_livraison.* "
Sql = Sql & "FROM ((t_R_Social INNER JOIN Users ON t_R_Social.SocieteId = Users.Id_Societes) "
Sql = Sql & "INNER JOIN t_livraison ON t_R_Social.id_livraison = t_livraison.id_livraison)  "
Sql = Sql & "INNER JOIN Users AS Users_1 ON t_R_Social.Id_Contact = Users_1.UserID  IN '"
Sql = Sql & TableauPath("Eb_Menu")
Sql = Sql & "' "
Sql = Sql & "Where Users.UserID =" & IdUser & "; "

'Sql = Sql & "And T_Num_Commande.NumCommande = '" & NUMCOM & "' "
'Sql = Sql & "GROUP BY Users.UserID, T_Num_Commande.NumCommande"
Set RsEboutiqueCliLivEncelade = Con.OpenRecordSet(Sql)
'NumCommadEboutiqueID = Rs!MaxDeIdAvoire
End Function
Public Function NumCommadEboutique() As String
Dim Sql As String
Dim Rs As Recordset
Dim NUMCOM As String
Sql = "SELECT T_Num_Command_Eboutique.NumCommand FROM T_Num_Command_Eboutique;"
Set Rs = Con.OpenRecordSet(Sql)
NumCommadEboutique = "" & Rs!NumCommand
Set Rs = Con.CloseRecordSet(Rs)
End Function


Public Function EboutiqueUserId() As Long
Dim Rs As Recordset
Dim Sql As String
Sql = "SELECT UserEboutique.UserEboutique FROM UserEboutique;"

Set Rs = Con.OpenRecordSet(Sql)
Sql = "SELECT Users.UserID FROM Users IN '"
Sql = Sql & TableauPath("Eb_Menu")
Sql = Sql & "' WHERE Users.UserLogin='" & Rs!UserEboutique & "';"
Set Rs = Con.OpenRecordSet(Sql)
EboutiqueUserId = Rs!UserID
Set Rs = Con.CloseRecordSet(Rs)
End Function
Public Function EboutiqueSocieteId() As Long
Dim Rs As Recordset
Dim Sql As String
Sql = "SELECT UserEboutique.UserEboutique FROM UserEboutique;"

Set Rs = Con.OpenRecordSet(Sql)


Sql = "SELECT t_R_Social.Id_Contact  FROM t_R_Social INNER JOIN Users ON "
Sql = Sql & "t_R_Social.SocieteId = Users.Id_Societes IN '"
Sql = Sql & TableauPath("Eb_Menu")
Sql = Sql & "' WHERE Users.UserLogin='" & Rs!UserEboutique & "';"
Set Rs = Con.OpenRecordSet(Sql)
EboutiqueSocieteId = Rs!Id_Contact
Set Rs = Con.CloseRecordSet(Rs)
End Function
Public Function Fun_ID_Cadie(NumCadi As String) As Long
Dim Rs As Recordset
Dim Sql As String
Sql = "INSERT INTO T_NumCadi ( NumCadi,Id_User ) in '"
Sql = Sql & TableauPath("Eb_Menu")
Sql = Sql & "' VALUES( '" & NumCadi & "'," & EboutiqueUserId & " );"
        Con.Execute Sql
        Sql = "SELECT T_NumCadi.id From T_NumCadi IN '"
        Sql = Sql & TableauPath("Eb_Menu")
        Sql = Sql & "' WHERE T_NumCadi.NumCadi='" & NumCadi & "';"
        Set Rs = Con.OpenRecordSet(Sql)
        If Rs.EOF = False Then
           Fun_ID_Cadie = Rs("id")
        
           
        End If
        
        Set Rs = Con.CloseRecordSet(Rs)
End Function
Public Function GetRemiseEboutique() As Long
Dim IdUser As Long
Dim Rs As Recordset
Dim Sql As String
IdUser = EboutiqueUserId

'SELECT Users.UserID, t_R_Social.Remise
'FROM t_R_Social INNER JOIN Users ON t_R_Social.SocieteId = Users.Id_Societes IN '\\Autocable\webprod\dbsportail\wwwencelade\Encelade_menu.mdb'
'WHERE (((Users.UserID)=311));



Sql = "SELECT  t_R_Social.Remise "
Sql = Sql & "FROM t_R_Social INNER JOIN Users ON t_R_Social.SocieteId = Users.Id_Societes IN '"
Sql = Sql & TableauPath("Eb_Menu")
Sql = Sql & "' WHERE Users.UserID=" & IdUser & ";"
Set Rs = Con.OpenRecordSet(Sql)
GetRemiseEboutique = Rs!Remise
Set Rs = Con.CloseRecordSet(Rs)

End Function
Public Function GetDefaultEboutique(fld, def)
def = safeEntry(def)
Dim Rs As Recordset
Dim Sql As String
'SELECT Users.UserID
'FROM Users IN '\\Autocable\webprod\dbsportail\wwwencelade\\Encelade_menu.mdb'
'WHERE (((Users.UserLogin)="autocable")); TableauPath("" & Rs!Lib_Menu)
Sql = "SELECT * FROM Defaults IN '"
Sql = Sql & TableauPath("Eb_Menu")
Sql = Sql & "'"
Sql = Sql & "WHERE defName = '" & fld & "'"
    Set Rs = Con.OpenRecordSet(Sql)
    If Not Rs.EOF Then
           GetDefaultEboutique = Trim(Rs("defValue"))
    Else
        Sql = "INSERT INTO Defaults ( defName, defValue ) IN '"
        Sql = Sql & TableauPath("Eb_Menu")
        Sql = Sql & "' VALUES('" & fld & "','" & def & "') "
       Con.Execute Sql
        GetDefaultEboutique = def
    End If
    Set Rs = Con.CloseRecordSet(Rs)
End Function
Public Function NumeroChrono(Recherche, Prefix) As String
Dim Sql As String
Dim Rs As Recordset
Dim txtValue As String
txtValue = GetDefaultEboutique(Recherche, Prefix)
Sql = "SELECT t_NumDevis.* FROM t_NumDevis IN '"
Sql = Sql & TableauPath("Eb_Menu")
Sql = Sql & "'"
Sql = Sql & "where t_NumDevis.type='" & Recherche & "';"
Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = True Then
        Sql = "INSERT INTO t_NumDevis ( NumDevis, Mydate,type ) IN '"
        Sql = Sql & TableauPath("Eb_Menu")
        Sql = Sql & "' values( 1 , Now() ,'" & Recherche & "');"
        Con.Execute Sql
        NumeroChrono = 1
    Else
        If Format(Rs!MyDate, "dd/mm/yyyy") = Format(Date, "dd/mm/yyyy") Then
            NumeroChrono = Rs!numDevis + 1
        Else
            NumeroChrono = 1
        End If
        
    End If
    Sql = "UPDATE t_NumDevis in '"
    Sql = Sql & TableauPath("Eb_Menu")
    Sql = Sql & "' SET t_NumDevis.NumDevis =" & NumeroChrono & ", t_NumDevis.Mydate = Now() where t_NumDevis.type='" & Recherche & "';"
   ' Session("MySql") =' Session("MySql") & vbCrLf & Sql & "<br><br>"
    Con.Execute Sql
    
    NumeroChrono = Prefix & Format(Date, "#") & "_" & NumeroChrono
End Function
Public Function RetournLibMenu(MyPath) As String
Dim SplitPath
SplitPath = Split(TableauPath(MyPath), "\")
SplitPath = SplitPath(UBound(SplitPath))
SplitPath = Replace(SplitPath, ".mdb", "")
SplitPath = Replace(SplitPath, "Encelade_", "")
RetournLibMenu = SplitPath
End Function

Public Function RetournIdMenu(MyPath) As Long
Dim SplitPath
Dim RsMenu As Recordset
Dim Sql As String
SplitPath = Split(TableauPath(MyPath), "\")
SplitPath = SplitPath(UBound(SplitPath))
SplitPath = Replace(SplitPath, ".mdb", "")
SplitPath = Replace(SplitPath, "Encelade_", "")
Sql = "SELECT MyMenu.CatId FROM (SELECT Menu.*  from  Menu  IN '"
Sql = Sql & TableauPath("Eb_Menu")
Sql = Sql & "') AS MyMenu "
Sql = Sql & "WHERE MyMenu.libelle='" & SplitPath & "';"
Set RsMenu = Con.OpenRecordSet(Sql)
RetournIdMenu = Val("" & RsMenu!CatId)
Set RsMenu = Con.CloseRecordSet(RsMenu)
End Function

Public Function RetournMenuId(LibMenu) As Long
Dim SplitPath
Dim RsMenu As Recordset
Dim Sql As String
Sql = "SELECT MyMenu.CatId FROM (SELECT Menu.*  from  Menu  IN '"
Sql = Sql & TableauPath("Eb_Menu")
Sql = Sql & "') AS MyMenu "
Sql = Sql & "WHERE MyMenu.libelle='" & LibMenu & "';"
Set RsMenu = Con.OpenRecordSet(Sql)
RetournMenuId = Val("" & RsMenu!CatId)
Set RsMenu = Con.CloseRecordSet(RsMenu)
End Function
Public Function RetournPathMenu(id_Menu As Long) As String
Dim RsMenu As Recordset
Dim Sql As String
Sql = "SELECT MyMenu.base FROM (SELECT Menu.*  from  Menu  IN '"
Sql = Sql & TableauPath("Eb_Menu")
Sql = Sql & "') AS MyMenu "
Sql = Sql & "WHERE MyMenu.CatId='" & id_Menu & "';"
Set RsMenu = Con.OpenRecordSet(Sql)
RetournPathMenu = Val("" & RsMenu!CatId)
Set RsMenu = Con.CloseRecordSet(RsMenu)
End Function
Public Sub NewAutocadAdmin()
Dim Sql As String
Dim Rs As Recordset
blnResult = RevertToSelf()
Sql = "SELECT AdminAutocable.User, AdminAutocable.PassWord, AdminAutocable.Serveur,AdminAutocable.Service,AdminAutocable.Domain "
Sql = Sql & "FROM AdminAutocable;"
Set Rs = Con.OpenRecordSet(Sql)
    lpUsername = "" & Rs!User
    lpDomain = "" & Rs!domain
    lpPassword = "" & Rs!PassWord
If LogonUser( _
lpUsername, _
lpDomain, _
lpPassword, _
          LOGON32_LOGON_INTERACTIVE, _
         LOGON32_PROVIDER_DEFAULT, _
            lngTokenHandle) = 0 Then
    MsgBox "Impossible d'ouvrir la session : " & lpUsername & ". "
    GoTo Fin
End If


If blnResult = False Then
    MsgBox "Impossible d'ouvrir LogonUser"
   GoTo Fin
End If
'MsgBox "Session avec le jeton" & lngTokenHandle & " et " & strAdminUser & ", " & strAdminDomain & ", " & strAdminPassword & " ouverte !"


blnResult = ImpersonateLoggedOnUser(lngTokenHandle)
'MySeconde 15
On Error Resume Next
  SetAutocad

If Err = 0 Then
    AutoApp.Visible = False
    AutoApp.Documents(0).Close False
    Example_AutoAudit
    DoEvents
    IsCilent = False
    boolAutoCAD = True
    Logoff

Else
    MsgBox "Plus de licence Autocad disponible", vbInformation, "AutoCâble  licence :"
    boolAutoCAD = False
End If

Fin:
Set Rs = Con.CloseRecordSet(Rs)
End Sub
Public Sub Example_AutoAudit()
    ' This example returns the current setting of
    ' AutoAudit. It then changes the value, and finally
    ' it resets the value back to the original setting.
    
    Dim preferences As Object
    Dim currAutoAudit As Boolean
    Dim newAutoAudit As Boolean
    
    Set preferences = AutoApp.preferences
    
    ' Retrieve the current AutoAudit value
'    currAutoAudit = preferences.OpenSave.AutoAudit
'    MsgBox "The current value for AutoAudit is " & currAutoAudit, vbInformation, "AutoAudit Example"
'
    ' Toggle the value for AutoAudit
    newAutoAudit = Not (currAutoAudit)
    preferences.OpenSave.AutoAudit = newAutoAudit
'    MsgBox "The new value for AutoAudit is " & newAutoAudit, vbInformation, "AutoAudit Example"
'
'    ' Reset AutoAudit to its original value
'    preferences.OpenSave.AutoAudit = newAutoAudit
'    MsgBox "The AutoAudit value is reset to " & currAutoAudit, vbInformation, "AutoAudit Example"
End Sub

Public Sub CreatKlc(DocAutoCad As Object, OptionPiece As String)
On Error Resume Next
DocAutoCad.Layers.Add "KLC_" & Trim("" & OptionPiece)
DocAutoCad.ActiveLayer = DocAutoCad.Layers("KLC_" & Trim("" & OptionPiece))

End Sub
Public Sub LibertPice()
 Dim Sql As String
Sql = "UPDATE T_indiceProjet SET T_indiceProjet.UserName = Null "
Sql = Sql & "WHERE T_indiceProjet.UserName='" & Replace(Machine, "'", "''") & "';"
Con.Execute Sql
End Sub
Public Sub MajDroitsFrm(IdUser As Long)
'
Dim Sql As String
Dim Rs As Recordset
Dim NbMenu As Long



Sql = "SELECT T_Boutons.Name,T_Boutons.Bouton FROM T_Boutons "
Sql = Sql & "Where T_Boutons.ContonTotal = False "
Sql = Sql & "ORDER BY T_Boutons.Name;"
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
frmAutocâble.Controls(Trim("" & Rs!Name)).Enabled = True
frmAutocâble.Controls(Trim("" & Rs!Name)).Caption = Trim("" & Rs!Bouton)
DoEvents
    Rs.MoveNext
    
Wend

Sql = "SELECT T_Boutons.Name "
Sql = Sql & "FROM T_Boutons INNER JOIN T_Droits ON T_Boutons.Id = T_Droits.Id_Bouton "
Sql = Sql & "GROUP BY T_Boutons.Name, T_Boutons.ContonTotal "
Sql = Sql & "Having T_Boutons.ContonTotal = False "
Sql = Sql & "ORDER BY T_Boutons.Name;"
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
frmAutocâble.CommandButton9.Enabled = False
frmAutocâble.Controls(Trim("" & Rs("Name"))).Enabled = False
DoEvents
    Rs.MoveNext

Wend

Sql = "SELECT T_Boutons.Name "
Sql = Sql & "FROM T_Boutons INNER JOIN ((T_Users INNER JOIN (T_Groupe INNER JOIN  "
Sql = Sql & "T_Groupe_Users ON T_Groupe.id = T_Groupe_Users.Id_Groupe)  "
Sql = Sql & "ON T_Users.Id = T_Groupe_Users.Id_Users) INNER JOIN T_Droits  "
Sql = Sql & "ON T_Users.Id = T_Droits.Id_Useur) ON T_Boutons.Id = T_Droits.Id_Bouton "
Sql = Sql & "WHERE T_Boutons.ContonTotal = False and T_Users.Id=" & IdUser & ";"
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
frmAutocâble.Controls(Trim("" & Rs!Name)).Enabled = True
DoEvents
    Rs.MoveNext
    
Wend
Sql = "SELECT Utilitaire.NameBouton, Utilitaire.Utilitaire "
Sql = Sql & "FROM Utilitaire ORDER BY Utilitaire.NameBouton;"
Set Rs = Con.OpenRecordSet(Sql)


'Me.Modules.Visible = False
If Rs.EOF = True Then
    frmAutocâble.CommandButton10.Enabled = False
Else
For I = frmAutocâble.Utilitaire.Count - 1 To 1 Step -1
    If I <> 0 Then
        Unload frmAutocâble.Utilitaire(I)
    End If
Next

While Rs.EOF = False
If NbMenu = 0 Then
    frmAutocâble.Utilitaire(NbMenu).Caption = "" & Rs!NameBouton
Else
    Load frmAutocâble.Utilitaire(NbMenu)
    frmAutocâble.Utilitaire(NbMenu).Visible = True
    frmAutocâble.Utilitaire(NbMenu).Caption = "" & Rs!NameBouton
End If
'Me.List1.AddItem "" & Rs!NameBouton
NbMenu = NbMenu + 1
   Rs.MoveNext
Wend


'For I = 0 To Me.Controls.Count - 1
'    MyControl.Add I, Me.Controls(I).Name
'Next
End If
Set Rs = Con.CloseRecordSet(Rs)

Sql = "SELECT Module.NameBouton, Module.Utilitaire "
Sql = Sql & "FROM Module ORDER BY Module.NameBouton;"
Set Rs = Con.OpenRecordSet(Sql)


'Me.Modules.Visible = False
If Rs.EOF = True Then
    frmAutocâble.Modules.Enabled = False
Else
For I = frmAutocâble.ModuleDetail.Count - 1 To 1 Step -1
    If I <> 0 Then
'    If I = 1 Then
        Unload frmAutocâble.ModuleDetail(I)
'    Else
'        Unload ModuleDetail.Utilitaire(I)
'        End If
    End If
Next
NbMenu = 0
While Rs.EOF = False
If NbMenu = 0 Then
    frmAutocâble.ModuleDetail(NbMenu).Caption = "" & Rs!NameBouton
Else
    Load frmAutocâble.ModuleDetail(NbMenu)
    frmAutocâble.ModuleDetail(NbMenu).Visible = True
    frmAutocâble.ModuleDetail(NbMenu).Caption = "" & Rs!NameBouton
End If
'Me.List1.AddItem "" & Rs!NameBouton
NbMenu = NbMenu + 1
   Rs.MoveNext
Wend


'For I = 0 To Me.Controls.Count - 1
'    MyControl.Add I, Me.Controls(I).Name
'Next
End If
Set Rs = Con.CloseRecordSet(Rs)




End Sub

Public Function RetournIdApp(Application As String, Optional Retourn As Boolean, Optional SERVER As String, Optional PassWord As String) As Long
Dim Liste
Dim element
Dim Valid
Dim ColecAplication As New Collection
RetournIdApp = -1
If Trim("" & SERVER) <> "" Then
Set Liste = GetObject("winmgmts://" & SERVER).InstancesOf("Win32_Process")
Else
Set Liste = GetObject("winmgmts:").InstancesOf("Win32_Process")
End If
If Retourn = False Then
               

For Each element In Liste
    Debug.Print element.Name
    If UCase(element.Name) = UCase(Application) Then
        ColecAplication.Add element.Handle, element.Handle
    End If
Next element
Else
    On Error Resume Next
    For Each element In Liste
    Debug.Print element.Name & " : " & element.Handle
    If UCase(element.Name) = UCase(Application) Then
        Valid = ColecAplication(element.Handle)
        If Err Then
            Err.Clear
                RetournIdApp = element.Handle
                Exit For
        End If
    End If
Next element
End If

End Function
Public Sub StratProcess(lpApplicationName As String, lpCommandLine As String)
On Error Resume Next
    Dim Sql As String
    Dim Rs As Recordset
    Dim lpUsername As String, lpDomain As String, lpPassword As String
    Dim lpCurrentDirectory As String
    Dim StartInfo As STARTUPINFO, ProcessInfo As PROCESS_INFORMATION
    Sql = "SELECT AdminAutocable.User, AdminAutocable.PassWord, AdminAutocable.Serveur,AdminAutocable.Service "
Sql = Sql & "FROM AdminAutocable;"
Set Rs = Con.OpenRecordSet(Sql)
    lpUsername = "" & Rs!User
    lpDomain = ""
    lpPassword = "" & Rs!PassWord
    
    Set Rs = Con.CloseRecordSet(Rs)
    lpApplicationName = Trim("" & lpApplicationName) & " "
    lpCommandLine = " " & Trim("" & lpCommandLine) & " "
'    lpCommandLine = " \\10.30.0.5\production\Cablage-production\RENAULT\PI\662\16-PI\PI_662_05_1445_1\12-PL\PL_662_05_1444_1.dwg "

'    lpCommandLine = vbNullString 'use the same as lpApplicationName
    lpCurrentDirectory = ""  'use standard directory
    StartInfo.cb = LenB(StartInfo) 'initialize structure
    StartInfo.dwFlags = 0&
    CreateProcessWithLogon StrPtr(lpUsername), StrPtr(lpDomain), StrPtr(lpPassword), LOGON_WITH_PROFILE, StrPtr(lpApplicationName), StrPtr(lpCommandLine), CREATE_DEFAULT_ERROR_MODE Or CREATE_NEW_CONSOLE Or CREATE_NEW_PROCESS_GROUP, ByVal 0&, StrPtr(lpCurrentDirectory), StartInfo, ProcessInfo
    CloseHandle ProcessInfo.hThread 'close the handle to the main thread, since we don't use it
    CloseHandle ProcessInfo.hProcess 'close the handle to the process, since we don't use it
    'note that closing the handles of the main thread and the process do not terminate the process
    Set AutoApp = GetObject("", ProcessInfo.hProcess)
    'unload this application
End Sub
Public Function ConverOngletGneEtat(Ongl As String) As String
Dim Sql As String
Dim Rs As Recordset
ConverOngletGneEtat = Ongl
Sql = "SELECT T_Etats_Onglet.Onglet, T_Etats_Select_Filtre.FiltreName "
Sql = Sql & "FROM T_Etats_Onglet INNER JOIN T_Etats_Select_Filtre ON T_Etats_Onglet.Id = T_Etats_Select_Filtre.Id_Onglet "
Sql = Sql & "GROUP BY T_Etats_Onglet.Onglet, T_Etats_Select_Filtre.FiltreName;"
Set Rs = Con.OpenRecordSet(Sql)
ConverOngletGneEtat = Replace(Replace(UCase(ConverOngletGneEtat), UCase("cont_Ongt_"), ""), UCase("_FLT_S_Equ"), "")
 ConverOngletGneEtat = Replace(Replace(UCase(ConverOngletGneEtat), UCase("cont_Ongt_"), ""), UCase("S_Filtre_"), "")
 ConverOngletGneEtat = Replace(Replace(UCase(ConverOngletGneEtat), UCase("cont_Ongt_"), ""), UCase("S_Filtre"), "")
  ConverOngletGneEtat = Replace(Replace(ConverOngletGneEtat, "cont_Ongt_", ""), UCase("_Activ"), "")
While Rs.EOF = False
       ConverOngletGneEtat = Replace(UCase(ConverOngletGneEtat), UCase(Trim("" & Rs!Onglet)) & "_", "")
       ConverOngletGneEtat = Replace(UCase(ConverOngletGneEtat), UCase(Trim("" & Rs!FiltreName)) & "_", "")
       ConverOngletGneEtat = Replace(UCase(ConverOngletGneEtat), UCase("_" & Trim("" & Rs!FiltreName)), "")
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)

End Function

Public Sub DeletSheetEtat(MyWorkbook As EXCEL.Workbook, Action As Long, NumDoc As Long, Optional Deb As Long, Optional Fin As Long)
On Error Resume Next
Dim I As Long
Dim MyRange As Range
Dim DeletOk As Boolean
Dim I2 As Integer
Dim Dernier As Boolean
Dim I_OngletVide As Long
Dim PlusGrand As Boolean
Dim NameSheet As String
Dim CounRange1 As Long
Dim CounRange2 As Long
MyWorkbook.Application.DisplayAlerts = False
If Deb = 0 Then
    Deb = 1
End If
If Fin = 0 Then
    Fin = MyWorkbook.Sheets.Count
End If

If Deb > MyWorkbook.Sheets.Count Then
     Deb = MyWorkbook.Sheets.Count
End If
If Fin > MyWorkbook.Sheets.Count Then
     Fin = MyWorkbook.Sheets.Count
End If
If NumDoc = 0 Then
   
   Select Case Action
'*******************************************************
'*          Supprimer les Onglets Vide                 *
'*******************************************************
        Case 1
                    DeletOk = True
             
               
             For I = Fin To Deb Step -1
                MyWorkbook.Worksheets(I).Select
                PlusGrand = False
                    
                            Set MyRange = Nothing
                            Set MyRange = MyWorkbook.Worksheets(I).Range(Replace(MyWorkbook.Worksheets(I).Cells(1, 1).Address, "$", "")).CurrentRegion
                                If MyRange.Rows.Count > 1 And Trim("" & MyRange(1, 1)) <> "" Then
                                    PlusGrand = True
                                    
                                    
                                End If
                            Set MyRange = MyWorkbook.Worksheets(I).Range(Replace(MyWorkbook.Worksheets(I).Cells(5, 1).Address, "$", "")).CurrentRegion
                                If MyRange.Rows.Count > 1 And Trim("" & MyRange(1, 1)) <> "" Then
                                    PlusGrand = True
                                    
                                    
                                End If
                        
                    Debug.Print MyWorkbook.Worksheets(I).Name
                    If PlusGrand = False Then
                        MyWorkbook.Sheets(I).Delete
                    Else
                        MyWorkbook.Worksheets(I).Replace ";", ""
                        ReplaceNull MyWorkbook.Worksheets(I)
                    End If
                Next
'*******************************************************
'* Déplace les Onglets Vide en fin de classeur         *
'*******************************************************
        Case 2
            For I = Fin To Deb Step -1
            CounRange1 = 0
            CounRange2 = 0
                MyWorkbook.Worksheets(I).Select
                PlusGrand = False
                Set MyRange = Nothing
                            Set MyRange = MyWorkbook.Worksheets(I).Range(Replace(MyWorkbook.Worksheets(I).Cells(1, 1).Address, "$", "")).CurrentRegion
                                If MyRange.Rows.Count > 1 And Trim("" & MyRange(1, 1)) <> "" Then
                                    PlusGrand = True
                                    CounRange1 = MyRange.Rows.Count
                                    
                                End If
                            Set MyRange = MyWorkbook.Worksheets(I).Range(Replace(MyWorkbook.Worksheets(I).Cells(5, 1).Address, "$", "")).CurrentRegion
                                If MyRange.Rows.Count > 1 And Trim("" & MyRange(1, 1)) <> "" Then
                                    PlusGrand = True
                                     CounRange1 = MyRange.Rows.Count
                                    
                                End If
                Debug.Print MyWorkbook.Worksheets(I).Name
                If PlusGrand = False Then
                    For I_OngletVide = I + 1 To Fin
                        Dernier = True
                        PlusGrand = False
                        Set MyRange = Nothing
                            Set MyRange = MyWorkbook.Worksheets(I_OngletVide).Range(Replace(MyWorkbook.Worksheets(I_OngletVide).Cells(1, 1).Address, "$", "")).CurrentRegion
                                If MyRange.Rows.Count > 1 Then
                                    PlusGrand = True
                                    CounRange2 = MyRange.Rows.Count
                                    
                                End If
                            Set MyRange = MyWorkbook.Worksheets(I_OngletVide).Range(Replace(MyWorkbook.Worksheets(I_OngletVide).Cells(5, 1).Address, "$", "")).CurrentRegion
                                If MyRange.Rows.Count > 1 Then
                                    PlusGrand = True
                                    CounRange2 = MyRange.Rows.Count
                                    
                                End If
                        If PlusGrand = False Then
                            Dernier = False
                            NameSheet = MyWorkbook.Sheets(I).Name
                            If CounRange1 < CounRange2 Then
                                DepaceSheet MyWorkbook, I, I_OngletVide, Dernier
                                If NameSheet <> MyWorkbook.Sheets(I).Name Then
                                    I = I + 2
                                        If I > Fin Then I = Fin + 1
                                End If
                            End If
                            Exit For
                        Else
                            If I_OngletVide = Fin Then
                                Dernier = True
                            End If
                                If CounRange1 < CounRange2 Then
                                    DepaceSheet MyWorkbook, I, I_OngletVide, Dernier
                                    
                                    If NameSheet <> MyWorkbook.Sheets(I).Name Then
                                        I = I + 2
                                        If I > Fin Then I = Fin + 1
                                    End If
                                    Exit For
                                 End If
                            
                        End If
                    Next
                        
                        Debug.Print MyWorkbook.Worksheets(I_OngletVide).Name
                       
                    
            Else
                MyWorkbook.Worksheets(I).Replace ";", ""
                ReplaceNull MyWorkbook.Worksheets(I)
            End If
            Next
        
'*******************************************************
'*           Tier les Onglets du classeur >            *
'*******************************************************
        
         Case 3
             For I = Deb To Fin - 1
                 MyWorkbook.Worksheets(I).Select
                 If MyWorkbook.Worksheets(I).Range("A1").CurrentRegion.Rows.Count < _
                     MyWorkbook.Worksheets(I + 1).Range("A1").CurrentRegion.Rows.Count Then
                        DepaceSheet MyWorkbook, I, I + 1, True
                        I = I - 2
                        If I < Deb Then
                            I = Deb - 1
                        End If
                  End If
               
             Next
'*******************************************************
'*           Tier les Onglets du classeur <            *
'*******************************************************
             
         Case 4
             For I = Deb To Fin - 1
                 MyWorkbook.Worksheets(I).Select
                 If MyWorkbook.Worksheets(I).Range("A1").CurrentRegion.Rows.Count > _
                     MyWorkbook.Worksheets(I + 1).Range("A1").CurrentRegion.Rows.Count Then
                        DepaceSheet MyWorkbook, I, I + 1, True
                        I = I - 2
                        If I < Deb Then I = Deb - 1
                  End If

             Next
   End Select
End If

'*******************************************************
'* Supprimer les Onglets Présent à l'ouverture d'EXCEL *
'*******************************************************

    For I = MyWorkbook.Worksheets.Count To 1 Step -1
        If InStr(1, UCase(MyWorkbook.Worksheets(I).Name), UCase("§Feuil§")) <> 0 Then
            MyWorkbook.Sheets(I).Delete
        End If
    Next
'End If
   
   
'    MyWorkbook.Application.DisplayAlerts = True
    
End Sub


Public Sub insertExelAccess(MySheet As EXCEL.Worksheet, Table As String, RowStart As Long, Id_IndiceProjet As Long, _
                            Optional OnGletName As Boolean, Optional NotDeletTable As Boolean)
Dim Sql As String
Dim SqlValue As String
Dim MyRange As Range
Dim Rs As Recordset
On Error GoTo 0
If NotDeletTable = False Then
    Sql = "DELETE " & Table & ".* FROM " & Table & " WHERE " & Table & ".Id_IndiceProjet=" & Id_IndiceProjet & ";"
    Con.Execute Sql
End If
Set Rs = Con.OpenRecordSet("SELECT " & Table & ".* FROM " & Table & " WHERE " & Table & ".ID=0;")

Set MyRange = MySheet.Cells(RowStart, 1).CurrentRegion
'Myrange.Application'.Visible = True
Sql = "INSERT INTO " & Table & " ( Id_IndiceProjet,"
If OnGletName = True Then Sql = Sql & "Onglet,"
For I = 1 To MyRange.Columns.Count
    Sql = Sql & "[" & MyRange(1, I) & "],"
Next
Sql = Sql & "PoseChario,"
Sql = Left(Sql, Len(Sql) - 1) & ") Values (" & Id_IndiceProjet & ","
'Sql = Sql & ", PoseChario"
If OnGletName = True Then Sql = Sql & "'" & Replace(MySheet.Name, "'", "''") & "',"
If MyRange.Rows.Count = 1 Then

SqlValue = ""
        For I2 = 1 To MyRange.Columns.Count
            SqlValue = SqlValue & "null,"
        Next
        SqlValue = Left(SqlValue, Len(SqlValue) - 1) & ");"
    Con.Execute Sql & SqlValue

End If
For I = 2 To MyRange.Rows.Count
    SqlValue = ""
    ChronoChario = ChronoChario + 1
   
        For I2 = 1 To MyRange.Columns.Count
'        Debug.Print Myrange(I, I2).Address
'       Debug.Print Myrange(1, I2).Value & " = " & MySheet.Range(Myrange(I, I2).Address).FormulaR1C1
'       Myrange.Application'.Visible = True

Debug.Print MyRange(1, I2).Value & " : " & MyRange(1, I2).Value; a; " " & "" & MyRange(I, I2).FormulaR1C1
'Myrange.Application '.Visible = True
        Select Case Rs(MyRange(1, I2).Value).Type
        Case 11
            SqlValue = SqlValue & Replace(Replace(Replace(Replace(UCase(MyRange(I, I2)), "FALSE", 0), "TRUE", 1), "FAUX", 0), "VRAI", 1) & ","
        Case 202
            SqlValue = SqlValue & "'" & MyReplace("" & MyRange(I, I2).FormulaR1C1) & "',"
        Case 203
            SqlValue = SqlValue & "'" & MyReplace("" & MyRange(I, I2).FormulaR1C1) & "',"
        Case 5
            SqlValue = SqlValue & Replace(Val(Replace("" & MyReplace(MyRange(I, I2).FormulaR1C1), ",", ".")), ",", ".") & ","
        Case 3
            SqlValue = SqlValue & Replace(Val(Replace("" & MyReplace(MyRange(I, I2).FormulaR1C1), ",", ".")), ",", ".") & ","
        Case Else
            MsgBox ""
        End Select
    Next
    SqlValue = SqlValue & ChronoChario & ","
    SqlValue = Left(SqlValue, Len(SqlValue) - 1) & ");"
    Con.Execute Sql & SqlValue
    
    
Next

Set Rs = Con.CloseRecordSet(Rs)
End Sub



'Dans votre code, ici un click sur un label (les labels peuvent être transparents et donc se placer sur un BMP) :

Public Function RechercheSpreadsheet(MyRange, MyCellule, strRecherche) As Long
'Permet de rechercher une valeur dans un tableau Excel.
'MyxlWhole = MyxlWhole + 1
On Error Resume Next
'Recherche = Myrange.Find(What:=strRecherche, After:=Myrange.Cells(MyCellule, 1), _
'            LookIn:=xlFormulas, LookAt _
'         :=MyxlWhole, searchorder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
'        False).Row
        
RechercheSpreadsheet = MyRange.Find(What:=MyCellule, After:=MyRange.Cells(MyCellule.Row + 1, MyCellule.Column), findlookin:=ssFormulas, findlookat:=ssPart, SearchOrder:=ssByRows, SearchDirection:=ssNext, MatchCase:=False).Row
        
 
    If Err Then
        Err.Clear
        RechercheSpreadsheet = 0
    End If
End Function

Public Function safeEntry(Txt) As String
safeEntry = Txt
safeEntry = Replace(safeEntry, "'", "''")
safeEntry = Replace(safeEntry, Chr(34), Chr(34) & Chr(34))
End Function
Public Sub SetUpdate(fld, def)
Dim Sql As String
GetDefault fld, def

Sql = "UPDATE AutoCableDefaults SET AutoCableDefaults.defValue = '" & def & "' "
Sql = Sql & "WHERE AutoCableDefaults.defName='" & fld & "';"
Con.Execute Sql
End Sub
Public Function AutoIncrement(fld) As Long
Dim Sql As String
AutoIncrement = Val(GetDefault(fld, "0"))
AutoIncrement = AutoIncrement + 1
AutoIncrement = def
Sql = "UPDATE AutoCableDefaults SET AutoCableDefaults.defValue = '" & AutoIncrement & "' "
Sql = Sql & "WHERE AutoCableDefaults.defName='" & fld & "';"
Con.Execute Sql
End Function
Public Function GetDefault(fld, def)
Dim Sql As String
Dim Rs As Recordset
def = safeEntry(def)
Sql = "SELECT * FROM AutoCableDefaults WHERE defName = '" & fld & "'"
Set Rs = Con.OpenRecordSet(Sql)
    If Not Rs.EOF Then
           GetDefault = Trim(Rs("defValue"))
    Else

        Con.Execute "INSERT INTO AutoCableDefaults(defName,defValue) VALUES('" & fld & "','" & def & "')"
        GetDefault = def
    End If
    Set Rs = Con.CloseRecordSet(Rs)
End Function

Function LoadCalque()
  DocAutoCad.ActiveLayer = DocAutoCad.Layers("0")
   DocAutoCad.ActiveLayer.ViewportDefault = True
  DocAutoCad.PurgeAll
End Function

Function ValideChampsTexte(Formulaire, NbChamps As Long) As Boolean
   
    ValideChampsTexte = False
    For I = 1 To NbChamps
        If MyFormatQRY(Formulaire.Controls("txt" & CStr(I))) = False Then Exit Function
        DoEvents
       
    Next I
    ValideChampsTexte = True
    End Function
 Function MyFormatQRY(Txt As Object) As Boolean
 Dim MyTag
 MyFormatQRY = False
 MyTag = Split(Txt.Tag, ";")
    
        If Trim("" & Txt) = "" Then
            If UCase(Trim(MyTag(2))) = "QRY" Then
                MsgBox "Valeur de : " & MyTag(1) & " obligatoire", vbExclamation
                Txt.SetFocus
                Exit Function
            End If
        Else
            If MyFormat("" & MyTag(3), Txt, "" & MyTag(1)) = False Then
                Txt.SetFocus
                Exit Function
            End If
           
    
        End If
        MyFormatQRY = True
 End Function
 
 
 
 Function MyFormat(Mytype As String, MyText As Object, MyLib As String) As Boolean
 MyFormat = True
 If MyText = "" Then Exit Function
 
  Select Case UCase(Mytype)
                    Case "DATE"
                        If Not IsDate(MyText) Then
                            MsgBox "Vous devez saisir une date pour : " & MyLib, vbExclamation
                            MyText = ""
                            MyFormat = False
                            
                            Exit Function
                        Else
                            MyText = Format(MyText, "dd/mm/yyyy")
                        End If
                    Case "ENT"
                        If Not IsNumeric(MyText) Then
                            MsgBox "Vous devez saisir un nombre entier pour : " & MyLib, vbExclamation
                            MyText = ""
                            MyFormat = False
                            Exit Function
                        Else
                            If (InStr(1, (MyText), ",") <> 0) Or (InStr(1, (MyText), ".") <> 0) Then
                                MsgBox "Vous devez saisir un nombre entier pour : " & MyLib, vbExclamation
                                MyText = ""
                                MyFormat = False
                                Exit Function
                            End If
                        End If
                    Case "DBL"
                        If Not IsNumeric(MyText) Then
                            MyText = Replace(MyText, ".", ",")
                        End If
                        If Not IsNumeric(MyText) Then
                            MsgBox "Vous devez saisir un nombre à virgule pour : " & MyLib, vbExclamation
                            MyText = ""
                            MyFormat = False
                            Exit Function
                        End If
            End Select
 End Function
 
Function DecalInsertPointLigneTableau_fils_Bas(y, Ofset)
    DecalInsertPointLigneTableau_fils_Bas = y + Ofset
End Function
Function DecalInsertPointLigneTableau_fils_Gauche(x, Ofset)
    DecalInsertPointLigneTableau_fils_Gauche = x + Ofset
End Function
Function funAttributesLigne_Tableau_fils(MyName As String, Attributes, tableau, nb, Optional RangeTitre As Recordset, Optional BoolTirte As Boolean, Optional MyColection As Collection, Optional vignette As Boolean, Optional Epissure As Boolean)
Dim DESIGNATION As String
Dim MyNb As Long
Dim MyNbStart As Long
Dim msgAttib As String
Dim MyAttribute As Collection
Dim SlpitAttributes
Set MyAttribute = New Collection
For I = 0 To UBound(Attributes)
    MyAttribute.Add I, Attributes(I).TagString
Next
If vignette = True Then
    If Epissure = False Then
        DESIGNATION = ".HAUT"
        MyNb = 4
        MyNbStart = 2
    Else
        MyNbStart = 3
         DESIGNATION = " "
      MyNb = 3
    End If
Else
    DESIGNATION = ""
      MyNb = nb
      If BoolTirte = True Then
        MyNbStart = 2
      Else
      MyNbStart = 0
      End If
End If
On Error GoTo MsgError

For I = MyNbStart To MyNb
DoEvents
    If BoolTirte = False Then
    
    Debug.Print Attributes(I).TextString
    SlpitAttributes = "" & tableau(I) & ";"
        SlpitAttributes = Split(SlpitAttributes, ";")
        Debug.Print SlpitAttributes(0)
        Attributes(MyAttribute(Trim("" & SlpitAttributes(0)))).TextString = "" & SlpitAttributes(1)
    Else
        msgAttib = tableau(I)
       If Epissure = False Then
            If RangeTitre.Fields(I).Name = "RefConnecteurFour" Then
                'If Trim("" & tableau(I)) <> "" Then _
                    'Attributes(MyColection.Item(Replace(RangeTitre.Fields(I).Name, "PRECO", "PRECO.") & Designation)).TextString = Trim("" & tableau(I))
                                
            Else
                Attributes(MyColection.Item(Replace(RangeTitre.Fields(I).Name, "PRECO", "PRECO.") & DESIGNATION)).TextString = Trim("" & tableau(I))

            End If
        Else
            If RangeTitre.Fields(I).Name = "RefConnecteurFour" Then
                If Trim("" & tableau(I)) <> "" Then _
                Attributes(MyColection.Item("EPISSURE")).TextString = Trim("" & tableau(I))
                
            Else
                Attributes(MyColection.Item("EPISSURE")).TextString = Trim("" & tableau(I))
            End If
        End If
            DESIGNATION = ""
    End If
Next I
Exit Function
MsgError:
If BoolTirte = False Then
    FunError 2, "" & SlpitAttributes(0), "Tableau de Fils"
Else
    FunError 2, RangeTitre.Fields(I).Name, MyName
End If

Resume Next
End Function

Function FunInsBlock(PathName, InsertPoint, Name, Optional Rotation As Double, Optional XScaleFactor As Double, Optional YScaleFactor As Double, Optional ZScaleFactor As Double, Optional Options As String) As Object
'If Trim("" & Options) <> "" Then
'    CreatKlc DocAutoCad, UCase(Options)
'
'End If
On Error GoTo GesERR
Dim DDDD As Object

ErrInsert = False
'  AutoApp.Visible = True
    Set FunInsBlock = DocAutoCad.ModelSpace.InsertBlock(InsertPoint, PathName, 1#, 1#, 1#, 0#)
   Dim layerObj As Object
   a = FunInsBlock.Name
'  e = FunInsBlock.GetAttributes
  FunInsBlock.Rotation = Rotation
  If XScaleFactor = 0 Then XScaleFactor = 1
   If YScaleFactor = 0 Then YScaleFactor = 1
   If ZScaleFactor = 0 Then ZScaleFactor = 1
  FunInsBlock.XScaleFactor = XScaleFactor
   FunInsBlock.YScaleFactor = YScaleFactor
   FunInsBlock.ZScaleFactor = ZScaleFactor
   On Error Resume Next
    
    Err.Clear
'    AutoApp
DocAutoCad.ActiveLayer = DocAutoCad.Layers("0")
 
 DocAutoCad.Application.ZoomAll
    Exit Function
GesERR:
    FunError 100, "*", PathName & vbCrLf & Err.Description
    Err.Clear
    ErrInsert = True
'  DocAutoCad.ActiveLayer = DocAutoCad.Layers("0")
'   Resume Next
End Function

Function FunInsBlock2(PathName, InsertPoint, Name, Optional Rotation As Double, Optional XScaleFactor As Double, Optional YScaleFactor As Double, Optional ZScaleFactor As Double) As Object
On Error GoTo GesERR
ErrInsert = False
    Set FunInsBlock2 = DocAutoCad.ModelSpace.InsertBlock(InsertPoint, PathName, 1#, 1#, 1#, 0#)
   Dim layerObj As Object
   a = FunInsBlock2.Name
  e = FunInsBlock2.GetAttributes
  FunInsBlock2.Rotation = Rotation
  If XScaleFactor = 0 Then XScaleFactor = 1
   If YScaleFactor = 0 Then YScaleFactor = 1
   If ZScaleFactor = 0 Then ZScaleFactor = 1
  FunInsBlock2.XScaleFactor = XScaleFactor
   FunInsBlock2.YScaleFactor = YScaleFactor
   FunInsBlock2.ZScaleFactor = ZScaleFactor

   DocAutoCad.Application.ZoomAll
    Exit Function
GesERR:
' Msg = Msg & "***************************************************" & vbCrLf
'           Msg = Msg & PathName & vbCrLf
'            Msg = Msg & Err.Description & vbCrLf
'            Msg = Msg & "***************************************************" & vbCrLf & vbCrLf
ErrInsert = True
End Function
Function DirNUMEROFIL(Path As String, Index As Long) As String
    Dim MyDir As String
    Dim IndexDir As Long
    Dim Fso As New FileSystemObject
    DirNUMEROFIL = ""
    MyDir = ""
     
    For IndexDir = 0 To 3
        If Fso.FileExists(Path & "NUMEROFIL" & CStr(IndexDir + Index) & ".dwg") = True Then
            MyDir = Path & "NUMEROFIL" & CStr(IndexDir + Index) & ".dwg"
            Exit For
        End If
               
    Next IndexDir
    If Trim(MyDir) = "" Then Exit Function
    DirNUMEROFIL = MyDir
    Set Fso = Nothing
End Function
Function FunError(NumErr As Long, Lib1 As String, msg As String, Optional Lib2 As String)
Dim Sql As String
If Trim("" & Lib1) = "" Then Exit Function
If JobError = 0 Then JobError = AtrbNumError
msg = MsgErreur(NumErr, Lib1, Lib2, msg)
Sql = "INSERT INTO T_Error ( JobError, ValError ) "
Sql = Sql & "values(" & JobError & ",'" & msg & "' );"
Con.Execute Sql

End Function
Function AfficheErreur(Path As String, Entete)

    Dim NuFichier As Long
    Dim Text
    Dim MyTxtErr
    Dim Sql As String
    Dim RsErreur As Recordset
    Dim Fichier As String
    NuFichier = FreeFile
    
    Text = ""
    Sql = "SELECT T_Error.ValError FROM T_Error "
    Sql = Sql & "WHERE T_Error.JobError=" & JobError & ";"
   
    Set RsErreur = Con.OpenRecordSet(Sql)
    While RsErreur.EOF = False
        Text = Text & RsErreur!ValError & vbCrLf
        RsErreur.MoveNext
    Wend
    Set RsErreur = Con.CloseRecordSet(RsErreur)
    If Trim("" & Text) = "" Then Exit Function
    
    Dim Fso As New FileSystemObject

    MyTxtErr = Entete & Text
    pathUser = Environ("USERPROFILE")
    pathUser = pathUser + "\Mes Documents"
    
    If Fso.FolderExists(pathUser) = False Then
         Fso.CreateFolder pathUser
    End If
    
    If Fso.FolderExists(Path & "RepErrorLog") = False Then
        Fso.CreateFolder Path & "RepErrorLog"
    End If
    Fichier = Path & "RepErrorLog\Error_" & Format(Now, "yyyy-mm-dd_hh_mm_ss") & ".log"
    
    While Fso.FileExists(Fichier) = True
        Fichier = Path & "RepErrorLog\Error_" & Format(Now, "yyyy-mm-dd_hh_mm_ss") & ".log"
    
    Wend
        FichierErr = Fichier
    Open Fichier For Output As #NuFichier
    Print #NuFichier, MyTxtErr
    Close #NuFichier
    Sql = "DELETE T_Error.* FROM T_Error "
    Sql = Sql & "WHERE T_Error.JobError=" & JobError & ";"
    Con.Execute Sql
    Set Fso = Nothing
     If IsServeur = False Then
        Shell "notepad.exe " & Fichier, vbMaximizedFocus
     End If
End Function
 Public Function VueArriere(MySheet As Worksheet) As Long
 On Error Resume Next
Dim CollectionColon As New Collection
Dim MyRangeSource As Range
Dim TableFilsD() As String
Dim IndexFilsD As Long
Dim TableFilsG() As String
Dim IndexFilsG As Long
Dim TableFilsC() As String
Dim IndexFilsC As Long
Dim ColStart As Integer
Dim SheetName As String
Dim SheetName2 As String
Dim PosPrefix As Long
Dim C, L As Long
VueArriere = 0
Set MyRangeSource = MySheet.Range("A1").CurrentRegion
ColStart = MyRangeSource.Columns.Count / 2
ColStart = ColStart - 1
For C = 1 To MyRangeSource.Columns.Count
    CollectionColon.Add C, MyRangeSource(1, C).Value
Next
'MySheet.Application.Visible = True
SheetName = MySheet.Name
For L = 2 To MyRangeSource.Rows.Count
   If InStr(1, UCase(SheetName), UCase(MyRangeSource(L, CollectionColon("App")))) <> 0 Then
        SheetName = MyRangeSource(L, CollectionColon("App"))
        Exit For
   End If
   If InStr(1, UCase(SheetName), UCase(MyRangeSource(L, CollectionColon("App2")))) <> 0 Then
        SheetName = MyRangeSource(L, CollectionColon("App2"))
        Exit For
   End If
Next


'SheetName = Replace(SheetName, Sufix, "")
'PosPrefix = InStr(1, SheetName, Prefix)
'
'If PosPrefix <> 0 Then
'PosPrefix = PosPrefix + Len(Prefix)
'Debug.Print Mid(SheetName, PosPrefix, Len(SheetName) - (PosPrefix - 1))
'SheetName = Mid(SheetName, PosPrefix, Len(SheetName) - (PosPrefix - 1))
'End If
If Left(UCase(SheetName), 1) = "E" Then Epissur = True
 If Epissur = True Then
 SheetName2 = SheetName
' MyRangeSource.Application.Visible = True
'    Set MyRangeSource = MySheet.Cells(1, 1).CurrentRegion
    For I = 2 To MyRangeSource.Rows.Count
      
       If UCase(SheetName) = UCase(MyRangeSource(I, CollectionColon("App"))) Then
            If UCase(Left("" & MyRangeSource(I, CollectionColon("VOI")) & " ", 1)) = "G" Then
                IndexFilsG = IndexFilsG + 1
                ReDim Preserve TableFilsG(IndexFilsG)
                TableFilsG(IndexFilsG) = MyRangeSource(I, CollectionColon("App2")) & " : " & MyRangeSource(I, CollectionColon("VOI2")) & " FILS: " & MyRangeSource(I, CollectionColon("FIL"))
            Else
                 If UCase(Left("" & MyRangeSource(I, CollectionColon("VOI")) & " ", 1)) = "D" Then
                    IndexFilsD = IndexFilsD + 1
                    ReDim Preserve TableFilsD(IndexFilsD)
                    TableFilsD(IndexFilsD) = MyRangeSource(I, CollectionColon("App2")) & " : " & MyRangeSource(I, CollectionColon("VOI2")) & " FILS: " & MyRangeSource(I, CollectionColon("FIL"))
                 Else
                    IndexFilsC = IndexFilsC + 1
                    ReDim Preserve TableFilsC(IndexFilsC)
                    TableFilsC(IndexFilsC) = MyRangeSource(I, CollectionColon("App2")) & " : " & MyRangeSource(I, CollectionColon("VOI2")) & " FILS: " & MyRangeSource(I, CollectionColon("FIL"))
                 End If
            End If
       End If
         If UCase(SheetName) = UCase(MyRangeSource(I, CollectionColon("App2"))) Then
            If UCase(Left("" & MyRangeSource(I, CollectionColon("VOI2")) & " ", 1)) = "G" Then
                IndexFilsG = IndexFilsG + 1
                ReDim Preserve TableFilsG(IndexFilsG)
                TableFilsG(IndexFilsG) = MyRangeSource(I, CollectionColon("App")) & " : " & MyRangeSource(I, CollectionColon("VOI")) & " FILS: " & MyRangeSource(I, CollectionColon("FIL"))
            Else
                 If UCase(Left("" & MyRangeSource(I, CollectionColon("VOI2")) & " ", 1)) = "D" Then
                    IndexFilsD = IndexFilsD + 1
                    ReDim Preserve TableFilsD(IndexFilsD)
                    TableFilsD(IndexFilsD) = MyRangeSource(I, CollectionColon("App")) & " : " & MyRangeSource(I, CollectionColon("VOI")) & " FILS: " & MyRangeSource(I, CollectionColon("FIL"))
                 Else
                    IndexFilsC = IndexFilsC + 1
                    ReDim Preserve TableFilsC(IndexFilsC)
                    TableFilsC(IndexFilsC) = MyRangeSource(I, CollectionColon("App")) & " : " & MyRangeSource(I, CollectionColon("VOI")) & " FILS: " & MyRangeSource(I, CollectionColon("FIL"))
                 End If
            End If
       End If
    Next
'    MyRangeSource.Application.Visible = True
        MyRangeSource(MyRangeSource.Rows.Count + 3, ColStart) = "Gauche"
        MyRangeSource(MyRangeSource.Rows.Count + 3, ColStart + 1) = Replace(UCase(SheetName2), UCase(Prefix), "")
        MyRangeSource(MyRangeSource.Rows.Count + 3, ColStart + 2) = "Droite"
        FormatExcelPlage MyRangeSource(MyRangeSource.Rows.Count + 3, ColStart), 40, False, True, xlCenter, xlCenter
        FormatExcelPlage MyRangeSource(MyRangeSource.Rows.Count + 3, ColStart + 1), 40, False, True, xlCenter, xlCenter
        FormatExcelPlage MyRangeSource(MyRangeSource.Rows.Count + 3, ColStart + 2), 40, False, True, xlCenter, xlCenter
        For I = 1 To IndexFilsG
        MyRangeSource(MyRangeSource.Rows.Count + 3 + I, ColStart) = TableFilsG(I)
        Next
        For I = 1 To IndexFilsC
        MyRangeSource(MyRangeSource.Rows.Count + 3 + I, ColStart + 1) = TableFilsC(I)
        Next
        For I = 1 To IndexFilsD
        MyRangeSource(MyRangeSource.Rows.Count + 3 + I, ColStart + 2) = TableFilsD(I)
        Next
        I = IndexFilsG
        If I < IndexFilsC Then I = IndexFilsC
        
        
        If I < IndexFilsD Then I = IndexFilsD
            If MySheet.Range(Replace(MyRangeSource(MyRangeSource.Rows.Count + 3, ColStart).Address & ":" & MyRangeSource(MyRangeSource.Rows.Count + 3 + I, ColStart + 2).Address, "$", "")).Rows.Count = 1 Then
                SuprmerCells MySheet.Range(Replace(MyRangeSource(MyRangeSource.Rows.Count + 3, ColStart).Address & ":" & MyRangeSource(MyRangeSource.Rows.Count + 3 + I, ColStart + 2).Address, "$", "")), "H"
            Else
                FormatExcelPlage2 MySheet.Range(MyRangeSource(MyRangeSource.Rows.Count + 3, ColStart).Address & ":" & MyRangeSource(MyRangeSource.Rows.Count + 3 + I, ColStart + 2).Address), 40, False, True, xlCenter, xlCenter, CLng(I)
            End If
         If I <> 0 Then I = I + 3
End If
VueArriere = I
End Function

Function MyReplace(strVal As String) As String
strVal = Trim(strVal)
MyReplace = strVal
MyReplace = Replace(MyReplace, "'", "''")
MyReplace = Replace(MyReplace, Chr(34), Chr(34) & Chr(34))
MyReplace = Trim("" & MyReplace)
End Function
Function MyReplaceDate(strVal As String) As String
    If Trim(strVal) = "" Then
        MyReplaceDate = "NULL"
    Else
        MyReplaceDate = "#" & strVal & "#"
    End If
    
End Function
Function MyReplaceBool(strVal As Object) As String
    If strVal.Value = True Then
        MyReplaceBool = "true"
    Else
        MyReplaceBool = "false"
    End If
    
End Function

Function OpenFichier(Fichier As String) As String

On Error Resume Next
SecuFill Fichier, False
DoEvents
Example_AutoAudit
'AutoApp.VBE.Alert = True
'AutoApp.VBE.DisplayAlert = True
'    Dim MyDocument As New AutoCAD.AcadDocument
'GetAutocad

Set DocAutoCad = AutoApp.Documents.Open(Fichier)
DocAutoCad.Activate
DocAutoCad.PurgeAll
'DocAutoCad.SendCommand "_audit" & vbCrLf & "O" & vbCrLf & "O" & vbCrLf & "O" & vbCrLf

'MySeconde 10
DocAutoCad.PurgeAll
AutoCAD.Visible = True
DocAutoCad.SendCommand "Commande: _zoom _e *Annuler*" & Chr(10)


DoEvents
   OpenFichier = DocAutoCad.Name
'   Set MyDocument = Nothing
   
End Function
Function OpenNew() As String

  

    Set DocAutoCad = AutoApp.Documents.Add
    OpenNew = DocAutoCad.Name
End Function

Sub SaveAs(Fichier)
On Error Resume Next

Dim MyErr As String
'    DocAutoCad.ActiveLayer = DocAutoCad.Layers("0")
DocAutoCad.Activate
        DocAutoCad.PurgeAll
'        DocAutoCad.SendCommand "_audit" & vbCrLf & " O " & vbCrLf & " O " & vbCrLf
'        MySeconde 10
        DocAutoCad.PurgeAll
    DocAutoCad.SaveAs Fichier
    
        Err.Clear
        MySeconde 3
       DocAutoCad.Save
     
    If Err Then
        MyErr = Err.Description
         Err.Clear
        If IsServeur = False Then
            MsgBox MyErr
         Else
            FunError 10, "" & Fichier, "" & MyErr
         End If
     End If
   
   
  
   DocAutoCad.Close , False
   MySeconde 3
   If UCase(strStatus) = "VAL" Then
        SecuFill Fichier & ".dwg", True
   Else
     SecuFill Fichier & ".dwg", False
    End If
   If AutoApp.Documents.Count = 0 Then AutoApp.Visible = False
End Sub
Function MySeconde(NuSeconde As Integer)
 a = Second(Time)
    While Abs(a - Second(Time)) < NuSeconde
    DoEvents
    Wend
End Function
Sub CloseDocument()
On Error Resume Next
    DocAutoCad.Close False
    On Error GoTo 0
End Sub
Function IsConnecteurs(Attributes As Variant) As Boolean
    Dim Table(5) As String
    Dim Trouve As Boolean
    
Table(0) = "DESIGNATION"
Table(1) = "POS"
Table(2) = "N°"
Table(3) = "CODE_APP"
Table(4) = "PRECO1"
Table(5) = "PRECO2"
    IsConnecteurs = True
    
      For I = LBound(Table) To UBound(Table)
      DoEvents
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = PRECO(UCase("" & Attributes(I2).TagString), True) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next I2
         If Trouve = False Then
                  IsConnecteurs = False
                Exit Function
            End If
      Next I
End Function
Function IsComposants(Attributes As Variant) As Boolean
    Dim Table(2) As String
    Dim Trouve As Boolean
    
Table(0) = "DESIGNCOMP"
Table(1) = "NUMCOMP"
Table(2) = "REFCOMP"
    IsComposants = True
    
      For I = LBound(Table) To UBound(Table)
      DoEvents
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = PRECO(UCase("" & Attributes(I2).TagString), True) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next I2
         If Trouve = False Then
                  IsComposants = False
                Exit Function
            End If
      Next I
End Function
Function IsNotas(Attributes As Variant) As Boolean
    Dim Table(0) As String
    Dim Trouve As Boolean
    
Table(0) = "NUMNOTA"

    IsNotas = True
    
      For I = LBound(Table) To UBound(Table)
      DoEvents
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = PRECO(UCase("" & Attributes(I2).TagString), True) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next I2
         If Trouve = False Then
                  IsNotas = False
                Exit Function
            End If
      Next I
End Function

Function IsTor(Attributes As Variant) As Boolean
    Dim Table(0) As String
    Dim Trouve As Boolean
    
Table(0) = "TORDESIGNATION"

    IsTor = True
    
      For I = LBound(Table) To UBound(Table)
      DoEvents
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = UCase("" & Attributes(I2).TagString) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next I2
         If Trouve = False Then
                  IsTor = False
                Exit Function
            End If
      Next I
End Function

Function IsTorDetail(Attributes As Variant) As Boolean
    Dim Table(2) As String
    Dim Trouve As Boolean
    
Table(2) = "TORDESIGNATION"
Table(1) = "TORFILS"
Table(0) = "TORNUM"
    IsTorDetail = True
    
      For I = LBound(Table) To UBound(Table)
      DoEvents
        For I2 = LBound(Attributes) To UBound(Attributes)
        Debug.Print UCase("" & Attributes(I2).TagString)
            If Table(I) = UCase("" & Attributes(I2).TagString) Then
            
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next I2
         If Trouve = False Then
                  IsTorDetail = False
                Exit Function
            End If
      Next I
End Function
Function IsCartoucheClient(Attributes As Variant) As Boolean
    Dim Table(6) As String
    Dim Trouve As Boolean
Table(0) = "DESIGN.1.CART.RENAULT"
Table(1) = "DESGN.1.ANGL.CART.REN"
Table(2) = "REF.PF.CART.RENAULT"
Table(3) = "IND.PF"
Table(4) = "REF.PLAN.INDUSTRIEL"
Table(5) = "IND.PI"
Table(6) = "REF.PIECE.CART.RENAULT"
    IsCartoucheClient = True
    
      For I = LBound(Table) To UBound(Table)
      DoEvents
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = UCase("" & Attributes(I2).TagString) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next I2
         If Trouve = False Then
                  IsCartoucheClient = False
                Exit Function
            End If
      Next I
End Function
Function IsCartoucheEncelade(Attributes As Variant) As Boolean
    Dim Table(6) As String
    Dim Trouve As Boolean
Table(0) = ".NOM.DU.CLIENT"
Table(1) = ".RESPONSABLE.CLIENT"
Table(2) = ".NOM.DU.PROJET"
Table(3) = ".VAGUE"
Table(4) = ".DESIGNATION.LIGNE.1"
Table(5) = ".OPTION.ET.DIVERSITE"
Table(6) = "REFERENCE.PLAN.CLIENT"
    IsCartoucheEncelade = True
    
      For I = LBound(Table) To UBound(Table)
      DoEvents
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = UCase("" & Attributes(I2).TagString) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next I2
         If Trouve = False Then
                  IsCartoucheEncelade = False
                Exit Function
            End If
      Next I
End Function


Function IsNOMBRE_FILS(Attributes As Variant) As Boolean
    Dim Table(0) As String
    
Table(0) = "_FILS"
   IsNOMBRE_FILS = True
    

   For I = LBound(Table) To UBound(Table)
   DoEvents
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = PRECO(UCase("" & Attributes(I2).TagString), True) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next I2
         If Trouve = False Then
                  IsNOMBRE_FILS = False
                Exit Function
            End If
      Next I
End Function

Function IsEpissures(Attributes As Variant) As Boolean
    Dim Table(6) As String
    
Table(0) = "DESIGNATION"
Table(1) = "POS"
Table(2) = "N°"
Table(3) = "CODE_APP"
Table(4) = "PRECO1"
Table(5) = "PRECO2"
Table(6) = "FILG1"


    IsEpissures = True
    

   For I = LBound(Table) To UBound(Table)
   DoEvents
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = PRECO(UCase("" & Attributes(I2).TagString)) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next I2
         If Trouve = False Then
                  IsEpissures = False
                Exit Function
            End If
      Next I
      
End Function
Function IsTableauFils(Attributes As Variant) As Boolean
    Dim Table(13) As String
    Table(0) = UCase("LIAI")
    Table(1) = UCase("Designation")
    Table(2) = UCase("Fil")
    Table(3) = UCase("SECT")
    Table(4) = UCase("CO")
    Table(5) = UCase("CO")
    Table(6) = UCase("ISO")
    Table(7) = UCase("POS")
    Table(8) = UCase("Con")
    Table(9) = UCase("VOIE")
    Table(10) = UCase("POS")
    Table(11) = UCase("Con")
    Table(12) = UCase("VOIE")
    Table(13) = UCase("LONG")
    IsTableauFils = True
    
    For I = LBound(Table) To UBound(Table)
    DoEvents
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = PRECO(UCase("" & Attributes(I2).TagString)) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next I2
         If Trouve = False Then
                  IsTableauFils = False
                Exit Function
            End If
      Next I
    
    
'
'    For i = LBound(Attributes) To UBound(Attributes)
'        If Trim(UCase(Attributes(i).TagString)) <> Trim(Table(i)) Then
'            IsTableauFils = False
'            Exit Function
'        End If
'    Next i
End Function

Function IsCriteres(Attributes As Variant) As Boolean
    Dim Table(1) As String
    Table(0) = UCase("REFCRITERE")
    Table(1) = UCase("REFCRITERELIB")
    IsCriteres = True
    
    For I = LBound(Table) To UBound(Table)
    DoEvents
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = PRECO(UCase("" & Attributes(I2).TagString)) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next I2
         If Trouve = False Then
                  IsCriteres = False
                Exit Function
            End If
      Next I
    
    
End Function
Function IsActionCorrective(Attributes As Variant) As Boolean
    Dim Table(1) As String
   
  
    Table(0) = UCase("REFACINDICE")
    Table(1) = UCase("REFACORRECTIVE")
    IsActionCorrective = True
    
    For I = LBound(Table) To UBound(Table)
    DoEvents
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = PRECO(UCase("" & Attributes(I2).TagString)) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next I2
         If Trouve = False Then
                  IsActionCorrective = False
                Exit Function
            End If
      Next I
    
    
End Function



Function IsNoeuds(Attributes As Variant) As Boolean
    Dim Table(1) As String
    Table(0) = UCase("LONG")
    Table(1) = UCase("NOEUD")
    IsNoeuds = True
    
    For I = LBound(Table) To UBound(Table)
    DoEvents
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = PRECO(UCase("" & Attributes(I2).TagString)) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next I2
         If Trouve = False Then
                  IsNoeuds = False
                Exit Function
            End If
      Next I
    
    
End Function

Function IsRefOption(Attributes As Variant) As Boolean
    Dim Table(0) As String
    Table(0) = UCase("REFOPTION")
   
    IsRefOption = True
    
    For I = LBound(Table) To UBound(Table)
    DoEvents
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = PRECO(UCase("" & Attributes(I2).TagString)) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next I2
         If Trouve = False Then
                  IsRefOption = False
                Exit Function
            End If
      Next I
    
    
End Function
Function IsEnteteTableauFils(Attributes As Variant) As Boolean
    Dim Table(12) As String
    Table(0) = UCase("LIAI")
    Table(1) = UCase("DESIGNATION")
    Table(2) = UCase("FIL")
    Table(3) = UCase("SECT")
    Table(4) = UCase("CO")
    Table(5) = UCase("ISO")
    Table(6) = UCase("POS")
    Table(7) = UCase("CON")
    Table(8) = UCase("VOIE")
    Table(9) = UCase("POS")
    Table(10) = UCase("CON")
    Table(11) = UCase("VOIE")
    Table(12) = UCase("LONG")
    
    IsEnteteTableauFils = True
    
    For I = LBound(Table) To UBound(Table)
    DoEvents
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = PRECO(UCase("" & Attributes(I2).TagString)) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next I2
         If Trouve = False Then
                  IsEnteteTableauFils = False
                Exit Function
            End If
      Next I
    
'
'    For i = LBound(Attributes) To UBound(Attributes)
'        If Trim(UCase(Attributes(i).TagString)) <> Trim(Table(i)) Then
'            IsEnteteTableauFils = False
'            Exit Function
'        End If
'    Next i
End Function

Sub LoadDb()
BdDateTable = CherCheInFihier("BdDateTable")
DbNumPlan = CherCheInFihier("Bdnumero")
Db = CherCheInFihier("BdAutocable")
If UCase(CherCheInFihier("IsCilent")) = "TRUE" Then IsCilent = True

If UCase(CherCheInFihier("IsServeur")) = "TRUE" Then IsServeur = True
ADO_TYPEBASE = CherCheInFihier("ADO_TYPEBASE")
ADO_BASE = CherCheInFihier("ADO_BASE")
ADO_SERVER = CherCheInFihier("ADO_SERVER")
ADO_Fichier = CherCheInFihier("ADO_Fichier")
ADO_User = CherCheInFihier("ADO_User=")
ADO_PassWord = CherCheInFihier("ADO_PassWord")
AutocableDRIVE = CherCheInFihier("AutocableDRIVE")
DonneesEntreprise = CherCheInFihier("DonneesEntreprise")
DonneesProduction = CherCheInFihier("DonneesProduction")

funOpenDatabase
NmJob = LaodJob
If IsServeur = IsCilent Then IsServeur = False: IsCilent = False
End Sub
Function LaodJob() As Long
Dim Sql As String
Dim Rs As Recordset
If NmJob = 0 Then

Sql = "SELECT [NumErreur]+1 AS Job FROM T_NumErreur WHERE T_NumErreur.LibErreur='Job';"
Set Rs = Con.OpenRecordSet(Sql)
LaodJob = Rs!Job
Sql = "UPDATE T_NumErreur SET T_NumErreur.NumErreur = [NumErreur]+1 WHERE T_NumErreur.LibErreur='Job';"
Con.Execute Sql
Set Rs = Con.CloseRecordSet(Rs)

Else
    LaodJob = NmJob
End If
End Function
Function PathArchive(PathRacicine As String, Client As String, CleAc As String, Piece As String, Mytype As String, Fichier, IdPieces As Long, Indice_Pieces As String, Indice_Plan As String, Version As Long, Optional NoRegistre As Boolean) As String
Dim Fso As New FileSystemObject
Dim Sql As String
Dim Rs As Recordset
Indice_Pieces = Trim("" & Indice_Pieces)
Indice_Plan = Trim("" & Indice_Plan)
Piece = Replace(Piece, "/", "_", 1)
Piece = Replace(Piece, ":", "", 1)
Piece = Replace(Piece, ".", "", 1)
Piece = Piece & "_" & Indice_Pieces
If UCase(Mytype) = UCase("SyntG") Or UCase(Mytype) = UCase("pdf") Or UCase(Mytype) = UCase("Synt") Or Mytype = "LIEC" Or Mytype = "DAC" Or Mytype = "DNC" Or Mytype = "FAB" Then
Else
Fichier = Fichier & "_" & Indice_Plan
End If
Fichier = Replace(Fichier, "/", "_", 1)
Fichier = Replace(Fichier, ":", "", 1)
Fichier = Replace(Fichier, ".", "", 1)



PathArchive = TableauPath.Item(Mytype)
PathArchive = Replace(UCase(PathArchive), UCase("[VariableClient]"), Client)
PathArchive = Replace(UCase(PathArchive), UCase("[VariableAff]"), CleAc)
    PathArchive = Replace(UCase(PathArchive), UCase("[VaribleDoc]"), Fichier)

    


If Version > 1 Then
    
    PathArchive = Replace(UCase(PathArchive), UCase("[VARIABLEPI]"), Piece & "_MOD")
Else
    PathArchive = Replace(UCase(PathArchive), UCase("[VARIABLEPI]"), Piece)
End If

PathRacicine = DefinirChemienComplet(TableauPath.Item("PathServer"), PathRacicine)
MyPath = Split(PathArchive, "\")
aa = ""
For IndexP = 0 To UBound(MyPath) - 1
aa = aa & MyPath(IndexP) & "\"
Debug.Print Left(PathRacicine & "\" & aa, Len(PathRacicine & "\" & aa) - 1)
If Fso.FolderExists(Left(PathRacicine & "\" & aa, Len(PathRacicine & "\" & aa) - 1)) = False Then
    Fso.CreateFolder Left(PathRacicine & "\" & aa, Len(PathRacicine & "\" & aa) - 1)
End If
Next


If NoRegistre = False Then
    If NomenclatureOk = True Then
        Sql = "UPDATE T_indiceProjet SET T_indiceProjet." & UCase(Mytype) & "AutoCadSave = '" & MyReplace(PathArchive) & "', T_indiceProjet.NbErr = " & NbError & ",T_indiceProjet." & UCase(Mytype) & "Ok=true "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & IdPieces & ";"
     Else
        Sql = "UPDATE T_indiceProjet SET T_indiceProjet." & UCase(Mytype) & "AutoCadSave = '" & MyReplace(PathArchive) & "' "
        Sql = Sql & "WHERE T_indiceProjet.Id=" & IdPieces & ";"
     End If
    Con.Execute Sql
End If
    


PathArchive = PathRacicine & "\" & PathArchive
Debug.Print PathArchive
End Function

Function DelAttribues(Attributes As Variant)
 For I = LBound(Attributes) To UBound(Attributes)
    Attributes(I).TextString = ""
    
 Next
End Function
Function IsVignette(Attributes As Variant)
    Dim Table(5) As String
    Table(2) = UCase("DESIGNATION.HAUT")
    Table(0) = UCase("CODE_APP")
    Table(1) = UCase("N°")
    Table(3) = UCase("DESIGNATION.BAS")
    Table(4) = UCase("DESIGNATION.GAUCHE")
    Table(5) = UCase("DESIGNATION.DROITE")
  
    IsVignette = True
     For I = LBound(Table) To UBound(Table)
     DoEvents
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = PRECO(UCase("" & Attributes(I2).TagString), True) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next I2
         If Trouve = False Then
                  IsVignette = False
                Exit Function
            End If
      Next I
End Function
Function IsVignetteEtiquette(Attributes As Variant)
    Dim Table(3) As String
    Dim Trouve As Boolean
    Table(0) = UCase("FIL1")
    Table(1) = UCase("FIL2")
    Table(2) = UCase("FIL3")
    Table(3) = UCase("FIL4")
  
    IsVignetteEtiquette = True
   
    Trouve = False
    For I2 = LBound(Attributes) To UBound(Attributes)
    DoEvents
        If Trim(UCase(Attributes(I2).TagString)) = "N°" Then
           Trouve = True
           Exit For
        End If
     Next I2
        If Trouve = True Then
         IsVignetteEtiquette = False
            Exit Function
        End If
    For I = LBound(Table) To UBound(Table)
    DoEvents
    Trouve = False
    For I2 = LBound(Attributes) To UBound(Attributes)
        If Trim(UCase(Attributes(I2).TagString)) = Trim(Table(I)) Then
           Trouve = True
           Exit For
        End If
     Next I2
        If Trouve = False Then
         IsVignetteEtiquette = False
            Exit Function
        End If
    Next I
End Function

Function IsVignetteEPISSURE(Attributes As Variant)
    Dim Table(0) As String
    Table(0) = UCase("EPISSURE")
  
    IsVignetteEPISSURE = True
    
       For I = LBound(Table) To UBound(Table)
       DoEvents
        For I2 = LBound(Attributes) To UBound(Attributes)
            If Table(I) = PRECO(UCase("" & Attributes(I2).TagString), True) Then
                Trouve = True
                Exit For
            Else
                Trouve = False
            End If
           
        Next I2
         If Trouve = False Then
                  IsVignetteEPISSURE = False
                Exit Function
            End If
      Next I
   
    
    
    
End Function
Sub SubLoadFils(IdPieces As Long, Mytype As String)
If (bool_Outil_E_Fils = False And Mytype = "PL") Or (bool_Outil_L_Fils = False And Mytype = "OU") Then Exit Sub
    Dim RsLigne As Recordset
    Dim Sql As String
    Dim Fso As New FileSystemObject
    Dim NbSupprim As Long
   Dim NbFils As Long
'    Dim NewTor As classTor
   
     
    PathBlocs = TableauPath.Item("PathBlocs")
      PathBlocs = DefinirChemienComplet(TableauPath.Item("PathServer"), PathBlocs)
'         If Left(PathBlocs, 2) <> "\\" And Left(PathBlocs, 1) = "\" Then PathBlocs = TableauPath.Item("PathServer") & PathBlocs
'         If Right(PathBlocs, 2) = "\\" Then PathBlocs = Mid(PathBlocs, 1, Len(PathBlocs) - 1)

    InsertPointLigneTableau_fils(0) = -1096.5549: InsertPointLigneTableau_fils(1) = -76.8274: InsertPointLigneTableau_fils(2) = 0

Sql = "SELECT Rq_Select_Tableau_Fils.* " ',  "
'    Sql = Sql & "T_indiceProjet.Id_Pieces "
    Sql = Sql & "FROM Rq_Select_Tableau_Fils "
    Sql = Sql & "WHERE Rq_Select_Tableau_Fils.Id_IndiceProjet=" & IdPieces & " " 'and Ligne_Tableau_fils.ACTIVER=true "
    Sql = Sql & "ORDER BY Rq_Select_Tableau_Fils.FIL;"

    Set RsLigne = Con.OpenRecordSet(Sql)
   While RsLigne.EOF = False
    NbFils = NbFils + 1
    RsLigne.MoveNext
   Wend
   RsLigne.Requery
EcritureTor RsLigne, Mytype

Ecriturefils RsLigne, Mytype, NbFils

        SubLoadCirteres IdPieces, Mytype
Fin:
    Set MyRange = Nothing
    Set MySheet = Nothing
    ReDim TableauDeConnecteurs(0)
    Set Fso = Nothing
    Set RsLigne = Con.CloseRecordSet(RsLigne)

End Sub

Sub SubLoadCirteres(IdPieces As Long, Mytype As String)
If (bool_Plan_E_Criteres = False And Mytype = "PL") Or (bool_Outil_E_Criteres = False And Mytype = "OU") Then Exit Sub

    Dim RsLigne As Recordset
    Dim Sql As String
    Dim NbSupprim As Long
    Dim NewBlock As Object
    Dim Collec As New Collection
    Dim MyColec As New Collection
    Dim Index As Long
'    Dim NewTor As classTor
   
     
    PathBlocs = TableauPath.Item("PathBlocs")
      PathBlocs = DefinirChemienComplet(TableauPath.Item("PathServer"), PathBlocs)
'         If Left(PathBlocs, 2) <> "\\" And Left(PathBlocs, 1) = "\" Then PathBlocs = TableauPath.Item("PathServer") & PathBlocs
'         If Right(PathBlocs, 2) = "\\" Then PathBlocs = Mid(PathBlocs, 1, Len(PathBlocs) - 1)
InsertPointLigneCritères(1) = -46#
InsertPointLigneCritères(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointLigneCritères(0), 377) '377)
InsertPointLigneCritères(2) = 0

Sql = "SELECT T_Critères.* " ',  "
'    Sql = Sql & "T_indiceProjet.Id_Pieces "
Sql = Sql & "FROM T_Critères "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & IdPieces & "  AND T_Critères.ACTIVER=True "
Sql = Sql & "ORDER BY T_Critères.CODE_CRITERE ;"

Set RsLigne = Con.OpenRecordSet(Sql)
If RsLigne.EOF = False Then
While RsLigne.EOF = False
Index = Index + 1
RsLigne.MoveNext
Wend
RsLigne.Requery
  FormBarGrah.ProgressBar1.Value = 0
FormBarGrah.ProgressBar1.Max = Index
FormBarGrah.ProgressBar1Caption.Caption = " Chargement des Critères"
Set NewBlock = FunInsBlock(PathBlocs & "\RefCriteres.dwg", InsertPointLigneCritères, "")

    While RsLigne.EOF = False
     IncremanteBarGrah FormBarGrah
     IncrmentServer FormBarGrah, Mytype
    InsertPointLigneCritères(1) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointLigneCritères(1), -3)
            Set NewBlock = FunInsBlock(PathBlocs & "\RefCriteres.dwg", InsertPointLigneCritères, "")
            Att = NewBlock.GetAttributes
            Set Colec = ColectionAttribueConecteur(Att)
            Att(Colec("REFCRITERE")).TextString = "" & RsLigne!CODE_CRITERE
           Att(Colec("REFCRITERELIB")).TextString = "" & RsLigne!Criteres

        RsLigne.MoveNext
    Wend
End If


End Sub


Sub SacnConnecteur(Mytype As String)
If (bool_Plan_E_Etiquettes = False And bool_Plan_L_Connecteurs = False And Mytype = "PL") Or (bool_Outil_E_Vignettes = False And bool_Outil_E_Connecteurs = False And Mytype = "OU") Then Exit Sub

    Dim Index As Long
    Dim NewBlock  As Object
    Dim MyFichier As String
    Dim PathNUMEROFIL As String
     FormBarGrah.ProgressBar1.Value = 0
      PathNUMEROFIL = TableauPath.Item("PathNUMEROFIL") & "\"
      PathNUMEROFIL = DefinirChemienComplet(TableauPath.Item("PathServer"), PathNUMEROFIL)
'                If Left(PathNUMEROFIL, 2) <> "\\" And Left(PathNUMEROFIL, 1) = "\" Then PathNUMEROFIL = TableauPath.Item("PathServer") & PathNUMEROFIL
'                If Right(PathNUMEROFIL, 2) = "\\" Then PathNUMEROFIL = Mid(PathNUMEROFIL, 1, Len(PathNUMEROFIL) - 1)
    If UBound(TableauDeConnecteurs) > 0 Then
         FormBarGrah.ProgressBar1.Max = 1 + UBound(TableauDeConnecteurs)
     Else
         FormBarGrah.ProgressBar1.Max = 1 + 1
     End If
     FormBarGrah.ProgressBar1Caption.Caption = " Chargement des vignettes"
    For Index = 1 To UBound(TableauDeConnecteurs)
         IncremanteBarGrah FormBarGrah
        IncrmentServer FormBarGrah, Mytype
        DoEvents
        
        If TableauDeConnecteurs(Index).ConnecteurExiste = True Then
            If TableauDeConnecteurs(Index).indexFile > 0 Then
                TableauDeConnecteurs(Index).TableauFile = TriTableau(TableauDeConnecteurs(Index).TableauFile)
               
                MyFichier = DirNUMEROFIL(PathNUMEROFIL, TableauDeConnecteurs(Index).indexFile)
                If Trim(MyFichier) <> "" Then
                    InsertionPoint = TableauDeConnecteurs(Index).NewVignette.InsertionPoint
                    
                    InsertPointLigneTableau_fils(0) = InsertionPoint(0) + 18.022: InsertPointLigneTableau_fils(1) = InsertionPoint(1): InsertPointLigneTableau_fils(2) = InsertionPoint(2)
'                    Set NewBlock = FunInsBlock(MyFichier, InsertPointLigneTableau_fils, "NF" & CInt(Index))
InStre = RetournInsertEtiquette(CInt(Index), InsertPointLigneTableau_fils)
                     Set NewBlock = FunInsBlock(MyFichier, RetournInsertEtiquette(CInt(Index), InsertPointLigneTableau_fils), "NF" & CInt(Index), RetourneRotationEtiquette(CInt(Index)), RetourneXEtiquette(CInt(Index)), RetourneYEtiquette(CInt(Index)), RetourneZEtiquette(CInt(Index)))
                    AttribueLibV NewBlock, Index
                    InsertPointLigneTableau_fils(1) = InsertionPoint(1)
                    InsertPointLigneTableau_fils(0) = InsertionPoint(0) + 70
                    Ats = TableauDeConnecteurs(Index).NewBlock.GetAttributes
                    ReseingeTor RetoureCadeApp(Ats), InsertPointLigneTableau_fils
                End If
            End If
        End If
    Next Index

End Sub
Function TriTableau(MyTableau)
    Dim Index As Long
    Dim boolPlus As Boolean
    a = ""
    For Index = 1 To UBound(MyTableau) - 1
        DoEvents
        
        While Val(MyTableau(Index)) > Val(MyTableau(Index + 1))
            Z = MyTableau(Index)
            a = MyTableau(Index + 1)
            MyTableau(Index) = a
            MyTableau(Index + 1) = Z
            Index = Index - 1
        Wend
    Next Index
    TriTableau = MyTableau

End Function
Function TriTableau2(MyTableau)
    Dim Index As Long
    Dim boolPlus As Boolean
    a = ""
    For Index = 1 To UBound(MyTableau) - 1
        DoEvents
        
        While Val(MyTableau(Index, 1)) > Val(MyTableau(Index + 1, 1))
            z0 = MyTableau(Index, 0)
            A0 = MyTableau(Index + 1, 0)
            z1 = MyTableau(Index, 1)
            a1 = MyTableau(Index + 1, 1)
            MyTableau(Index, 0) = A0
            MyTableau(Index + 1, 0) = z0
            MyTableau(Index, 1) = a1
            MyTableau(Index + 1, 1) = z1
            Index = Index - 1
        Wend
    Next Index
    TriTableau2 = MyTableau

End Function
Function InsertCartoucheEncelad(Index As Long, Mytype As String, ParamArray InsertPointLigneTableau_fils())
Dim Coef As Double
Dim MyMod As Double
MyMod = Index Mod 2
 Coef = Index
If MyMod <> 0 Then
   Coef = Coef - 1
End If
Coef = Coef / 2


If Mytype = "OU" Then
     InsertPointLigneTableau_fils(0)(0) = 3711.7662: InsertPointLigneTableau_fils(0)(1) = 115.8474: InsertPointLigneTableau_fils(0)(2) = 0
    GoTo Fin
 End If
'-1188
Select Case Index
        Case 1
            InsertPointLigneTableau_fils(0)(0) = -1188: InsertPointLigneTableau_fils(0)(1) = 0: InsertPointLigneTableau_fils(0)(2) = 0

        Case 2
            InsertPointLigneTableau_fils(0)(0) = -1188: InsertPointLigneTableau_fils(0)(1) = -840: InsertPointLigneTableau_fils(0)(2) = 0


        Case Else
            
            If MyMod = 0 Then
            Coef = Coef - 1
                InsertPointLigneTableau_fils(0)(0) = -1188: InsertPointLigneTableau_fils(0)(1) = -840: InsertPointLigneTableau_fils(0)(2) = 0
                InsertPointLigneTableau_fils(0)(0) = InsertPointLigneTableau_fils(0)(0) - (-1188 * Coef)

            Else
                InsertPointLigneTableau_fils(0)(0) = -1188: InsertPointLigneTableau_fils(0)(1) = 0: InsertPointLigneTableau_fils(0)(2) = 0
                InsertPointLigneTableau_fils(0)(0) = InsertPointLigneTableau_fils(0)(0) - (-1188 * Coef)
            End If

End Select
Fin:
End Function
Function InsertCartoucheClient(Index As Long, Mytype As String, ParamArray InsertPointLigneTableau_fils())
If Mytype = "OU" Then
    InsertPointLigneTableau_fils(0)(0) = 3566.7253: InsertPointLigneTableau_fils(0)(1) = 115.8474: InsertPointLigneTableau_fils(0)(2) = 0
    Exit Function
End If
Dim MyMod As Double
Dim Coef As Double
MyMod = Index Mod 2
Coef = Index
If MyMod <> 0 Then
    Coef = Coef - 1
End If
Coef = Coef / 2
'-1188

Select Case Index
        Case 1
                InsertPointLigneTableau_fils(0)(0) = -165.0409: InsertPointLigneTableau_fils(0)(1) = 126.0753: InsertPointLigneTableau_fils(0)(2) = 0

        Case 2
                 InsertPointLigneTableau_fils(0)(0) = -165.0409: InsertPointLigneTableau_fils(0)(1) = -713.9247: InsertPointLigneTableau_fils(0)(2) = 0


        Case Else
        If MyMod <> 0 Then
            
            aa = -165.0409 - 1022.9591
            InsertPointLigneTableau_fils(0)(0) = -165.0409: InsertPointLigneTableau_fils(0)(1) = 126.0753: InsertPointLigneTableau_fils(0)(2) = 0
            InsertPointLigneTableau_fils(0)(0) = InsertPointLigneTableau_fils(0)(0) - (aa * Coef)
        Else
        Coef = Coef - 1
            aa = -165.0409 - 1022.9591
            InsertPointLigneTableau_fils(0)(0) = -165.0409: InsertPointLigneTableau_fils(0)(1) = -713.9247: InsertPointLigneTableau_fils(0)(2) = 0
            InsertPointLigneTableau_fils(0)(0) = InsertPointLigneTableau_fils(0)(0) - (aa * Coef)
        End If
        
        
End Select
End Function

Function ChargeCartoucheEncelade(IdIndiceProjet As Long, Mytype As String, NbCartouche As Long, Optional OuOk As Boolean) As Boolean
Dim AttClent2
If (bool_Plan_E_cartouches = False And Mytype = "PL") Or (bool_Outil_E_cartouches = False And Mytype = "OU") Then Exit Function
  
Dim Fso As New FileSystemObject
Dim Sql As String
Dim Rs As Recordset
Dim Status As String
Dim FichierCartouche As String
Dim Index As Long
Dim RsCartochePlache As Recordset
'Sql = "UPDATE T_indiceProjet SET T_indiceProjet.Cartouche = '" & Replace(MyReplace(RepPlacheClous), TableauPath.Item("PathServer"), "") & "' "
'Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
'Con.Execute Sql
Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
Sql = Sql & "WHERE RqCartouche.T_indiceProjet.Id=" & IdIndiceProjet & ";"
'RepPlacheClous = DefinirChemienComplet(TableauPath.Item("PathServer"), RepPlacheClous)
'If Left(RepPlacheClous, 2) <> "\\" And Left(RepPlacheClous, 1) = "\" Then RepPlacheClous = TableauPath.Item("PathServer") & RepPlacheClous
'If Right(RepPlacheClous, 2) = "\\" Then RepPlacheClous = Mid(RepPlacheClous, 1, Len(RepPlacheClous) - 1)
PathBlocs = TableauPath.Item("PathBlocs")
PathBlocs = DefinirChemienComplet(TableauPath.Item("PathServer"), PathBlocs)
' If Left(PathBlocs, 2) <> "\\" And Left(PathBlocs, 1) = "\" Then PathBlocs = TableauPath.Item("PathServer") & PathBlocs
' If Right(PathBlocs, 2) = "\\" Then PathBlocs = Mid(PathBlocs, 1, Len(PathBlocs) - 1)
 PereFilsOk = AtocatOption(IdIndiceProjet)
 Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then Exit Function
    Status = "" & Rs!Status
    CartoucheCleient = False
    For Index = 1 To NbCartouche
      InsertCartoucheEncelad Index, Mytype, InsertPointLigneTableau_fils
    If Mytype = "OU" Then
        Sql = "SELECT T_indiceProjet.Cartouche, T_indiceProjet.Id FROM T_indiceProjet WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
        Set RsCartochePlache = Con.OpenRecordSet(Sql)
        
        If RsCartochePlache.EOF = False Then
            FichierCartouche = DefinirChemienComplet(TableauPath.Item("PathServer"), "" & RsCartochePlache!Cartouche)
        End If
      
    Else
        If Index = 1 Then
            
             FichierCartouche = PathBlocs & "\1 CARTOUCHE ENCELADE.dwg"
        Else
         FichierCartouche = PathBlocs & "\CARTOUCHE ENCELADE.dwg"
        End If
    End If
    If Fso.FileExists(FichierCartouche) = False Then
        Set Fso = Nothing
        Exit Function
    End If

    Set NewBlock = FunInsBlock(FichierCartouche, InsertPointLigneTableau_fils, "LeCartouche1E")

    AttClent = NewBlock.GetAttributes

    Set AttribuCartouche = ColectionAttribueConecteur(NewBlock.GetAttributes)
    
    AttClent(AttribuCartouche(".BASE.VEHICULE")).TextString = Replace("" & Rs("BaseVehicule"), Chr(13), "")
    AttClent(AttribuCartouche(".NOM.DU.Client")).TextString = Replace("" & Rs!Client, Chr(13), "")
    AttClent(AttribuCartouche(".RESPONSABLE.Client")).TextString = Replace("" & Rs!Responsable, Chr(13), "")
    AttClent(AttribuCartouche(".NOM.DU.Projet")).TextString = Replace("" & Rs!Projet, Chr(13), "")
    AttClent(AttribuCartouche(".VAGUE")).TextString = Replace("" & Rs!Vague, Chr(13), "")
        AttClent(AttribuCartouche(".DESIGNATION.LIGNE.1")).TextString = Replace("" & Rs!Ensemble, Chr(13), "")
        Txt = "" & Rs!Equipement
        If PereFilsOk = True Then Txt = ""
            AttClent(AttribuCartouche(".OPTION.ET.DIVERSITE")).TextString = Replace(Txt, Chr(13), "")
        
       
    
       
   
        AttClent(AttribuCartouche("Reference.PLAN.Client")).TextString = Replace("" & Rs!PL, Chr(13), "")
         AttClent(AttribuCartouche("INDICE")).TextString = Replace("" & Rs!PL_Indice, Chr(13), "")
    
   
    AttClent(AttribuCartouche("Reference.PLAN.FONCTIONNEL")).TextString = Replace("" & Rs!RefPF, Chr(13), "")
  
    AttClent(AttribuCartouche("RF2")).TextString = Replace("" & Rs!Ref_PF, Chr(13), "")
   Txt = "" & Rs!PI
        If PereFilsOk = True Then Txt = ""
        AttClent(AttribuCartouche("Reference.ENCELADE")).TextString = Replace(Txt, Chr(13), "")
        Txt = "" & Rs!PI_Indice
        If PereFilsOk = True Then Txt = ""
        AttClent(AttribuCartouche("RF1")).TextString = Replace(Txt, Chr(13), "")
     
     AttClent(AttribuCartouche("Reference.OU.ENCELADE")).TextString = Replace("" & Rs!OU, Chr(13), "")
         AttClent(AttribuCartouche("RF3")).TextString = Replace("" & Rs!OU_Indice, Chr(13), "")
    AttClent(AttribuCartouche("DESSINE.PAR")).TextString = Replace("" & Rs!DessineNOM, Chr(13), "")
    AttClent(AttribuCartouche("DESSINELE")).TextString = Replace("" & Rs!DessineDate, Chr(13), "")
    AttClent(AttribuCartouche("VERIFIE.PAR")).TextString = Replace("" & Rs!VerifieNom, Chr(13), "")
    AttClent(AttribuCartouche("VERIFIELE")).TextString = Replace("" & Rs!VerifieDate, Chr(13), "")
    AttClent(AttribuCartouche("APPROUVE.PAR")).TextString = Replace("" & Rs!ApprouveNom, Chr(13), "")
    AttClent(AttribuCartouche("APPROUVELE")).TextString = Replace("" & Rs!ApprouveDate, Chr(13), "")
    AttClent(AttribuCartouche(".MASSE")).TextString = Replace("" & Rs!MASSE, Chr(13), "")
    AttClent(AttribuCartouche("ETAT")).TextString = Replace(Status, Chr(13), "")
    AttClent(AttribuCartouche("X/X")).TextString = Replace(CStr(Index) & "/" & CStr(NbCartouche), Chr(13), "")
   Next Index
ActionCorrective IdIndiceProjet
    ChargeCartoucheEncelade = True
Set Fso = Nothing
End Function

Sub ActionCorrective(IdIndiceProjet As Long)
Dim Sql As String
Dim Rs As Recordset
Dim C As Long
Dim L As Long
Dim Fso As New FileSystemObject
On Error Resume Next
L = 1
C = 0
Dim Insert(0 To 2) As Double
Sql = "SELECT T_indiceProjet.Id_Pieces,  T_indiceProjet.Id FROM T_indiceProjet WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then Exit Sub
Sql = "SELECT T_indiceProjet_1.Id, T_indiceProjet_1.PI_Indice,T_indiceProjet_1.ReffIndice, T_indiceProjet_1.Description, [T_indiceProjet_1].[PI] & '_' &  "
Sql = Sql & "Trim('' & [T_indiceProjet_1].[Pi_Indice]) AS Piece, [T_indiceProjet_1].[Pl] & '_' & Trim('' & [T_indiceProjet_1].[Li_Indice]) "
Sql = Sql & "AS Plan, [T_indiceProjet_1].[ou] & '_' & Trim('' & [T_indiceProjet_1].[ou_Indice]) AS Outil, [T_indiceProjet_1].[Li] & '_' &  "
Sql = Sql & "Trim('' & [T_indiceProjet_1].[Li_Indice]) AS Liste, T_indiceProjet.Id_Pieces "
Sql = Sql & "FROM T_indiceProjet INNER JOIN T_indiceProjet AS T_indiceProjet_1 ON T_indiceProjet.Id_Pieces = T_indiceProjet_1.Id_Pieces "
Sql = Sql & "Where T_indiceProjet.Id_Pieces = " & Rs!Id_Pieces & " and T_indiceProjet_1.ReffIndice Is Not Null "
Sql = Sql & "GROUP BY T_indiceProjet_1.Id,  T_indiceProjet_1.PI_Indice,T_indiceProjet_1.ReffIndice, T_indiceProjet_1.Description,  "
Sql = Sql & "[T_indiceProjet_1].[PI] & '_' & Trim('' & [T_indiceProjet_1].[Pi_Indice]), [T_indiceProjet_1].[Pl] & '_' &  "
Sql = Sql & "Trim('' & [T_indiceProjet_1].[Li_Indice]), [T_indiceProjet_1].[ou] & '_' & Trim('' & [T_indiceProjet_1].[ou_Indice]),  "
Sql = Sql & "[T_indiceProjet_1].[Li] & '_' & Trim('' & [T_indiceProjet_1].[Li_Indice]), T_indiceProjet.Id_Pieces "

Sql = Sql & "ORDER BY T_indiceProjet_1.Id;"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then Exit Sub

Insert(0) = -368.5657
Insert(1) = 20
Insert(2) = 1
'Rs.MoveNext

While Rs.EOF = False
C = C + 1
If Fso.FileExists(DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath.Item("PathBlocs") & "\ActionCorrective.dwg")) = True Then
 Set NewBlock = FunInsBlock(DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath.Item("PathBlocs") & "\ActionCorrective.dwg"), Insert, "")
 Set Att = ColectionAttribueConecteur(NewBlock.GetAttributes)
 MyAtt = NewBlock.GetAttributes
 MyAtt(Att("REFACINDICE")).TextString = "" & Rs!PI_Indice
  MyAtt(Att("REFACORRECTIVE")).TextString = "" & Rs!ReffIndice
Else
    FunError 9, "ActionCorrective.dwg", ""
End If
 Insert(1) = DecalInsertPointLigneTableau_fils_Bas(Insert(1), -14)
 If C = 3 Then
 C = 0
 Insert(1) = 20
 Insert(0) = DecalInsertPointLigneTableau_fils_Bas(Insert(0), 56.5657)
 End If
Rs.MoveNext
Wend

End Sub
Function ChargeCartoucheClient(IdIndiceProjet As Long, Mytype As String, NbCartouche As Long, Optional OuOk As Boolean) As Boolean
If (bool_Plan_E_cartouches = False And Mytype = "PL") Or (bool_Outil_E_cartouches = False And Mytype = "OU") Then Exit Function

Dim Sql As String
Dim FichierCartouche As String
Dim Rs As Recordset
Dim RsCartouche As Recordset
Dim Index As Long
Dim NbCar As Long
LeCartouche = "CARTOUCHE  RENAULT.dwg"
LeCartoucheE = "CARTOUCHE ENCELADE.dwg"
NbCar = 2
If OuOk = False Then NbCar = NbCartouche
'If boolFormClient = False Then Exit Function
Sql = "SELECT RqCartouche.* "
Sql = Sql & "FROM RqCartouche "
Sql = Sql & "WHERE RqCartouche.T_indiceProjet.Id=" & IdIndiceProjet & ";"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then Exit Function
MyCARTOUCHE_Client = Trim("" & Rs!Client)
Sql = "SELECT T_Clients.Formulaire FROM T_Clients "
Sql = Sql & "WHERE T_Clients.Client='" & MyReplace(Trim("" & Rs!Client)) & "';"
Set RsCartouche = Con.OpenRecordSet(Sql)
If RsCartouche.EOF = False Then
    LeCartouche = Trim("" & RsCartouche!Formulaire)
    LeCartouche = DefinirChemienComplet(TableauPath.Item("PathServer"), LeCartouche)
'     If Left(LeCartouche, 2) <> "\\" And Left(LeCartouche, 1) = "\" Then LeCartouche = TableauPath.Item("PathServer") & LeCartouche
'     If Right(LeCartouche, 2) = "\\" Then LeCartouche = Mid(LeCartouche, 1, Len(LeCartouche) - 1)
     
End If
Set RsCartouche = Con.CloseRecordSet(RsCartouche)
If LeCartouche = "" Then Exit Function
    Dim Fso As New FileSystemObject
    CartoucheCleient = False
    For Index = 1 To NbCar
    InsertCartoucheClient Index, Mytype, InsertPointLigneTableau_fils
   
    If Fso.FileExists(LeCartouche) = False Then
        Set Fso = Nothing
        Exit Function
    End If
    Set NewBlock = FunInsBlock(LeCartouche, InsertPointLigneTableau_fils, "LeCartouche1")
'    NewBlock.Application.Visible = True
    AttClent = NewBlock.GetAttributes

    Set AttribuCartouche = ColectionAttribueConecteur(NewBlock.GetAttributes)
AttClent(AttribuCartouche("DESIGN.1.CART.RENAULT")).TextString = Replace("" & Rs("Ensemble"), Chr(13), "")
AttClent(AttribuCartouche("MASSE")).TextString = Replace("" & Rs("Masse"), Chr(13), "")

AttClent(AttribuCartouche("REF.PF.CART.RENAULT")).TextString = Replace("" & Rs("RefPF"), Chr(13), "")
AttClent(AttribuCartouche("IND.PF")).TextString = Replace("" & Rs("Ref_PF"), Chr(13), "")
'If OuOk = True Then
'
'        AttClent(AttribuCartouche("REF.PLAN.INDUSTRIEL")).TextString = "" & Rs!OU
'        AttClent(AttribuCartouche("IND.PI")).TextString = "" & Rs("OU_Indice")
'    Else
    
        AttClent(AttribuCartouche("REF.PLAN.INDUSTRIEL")).TextString = Replace("" & Rs!RefP, Chr(13), "")
         AttClent(AttribuCartouche("IND.Pi")).TextString = Replace("" & Rs("Ref_Plan_CLI"), Chr(13), "")
    
'    End If
AttClent(AttribuCartouche("DESIGN.2.CART.RENAULT")).TextString = ""
AttClent(AttribuCartouche("DESGN.1.ANGL.CART.REN")).TextString = ""
AttClent(AttribuCartouche("DESGN.2.ANGL.CART.REN")).TextString = ""

'AttClent(AttribuCartouche("IND.PF")).TextString = ""

 
AttClent(AttribuCartouche("REF.PIECE.CART.RENAULT")).TextString = Replace("" & Rs("RefPieceClient") & "_" & Trim("" & Rs("Ref_Piece_CLI")), Chr(13), "")
AttClent(AttribuCartouche("SERVICE")).TextString = Replace("" & Rs("Service"), Chr(13), "")
AttClent(AttribuCartouche("UTILISATEURS")).TextString = Replace("" & Rs("Destinataire"), Chr(13), "")



AttClent(AttribuCartouche("REGLEMENT")).TextString = ""
AttClent(AttribuCartouche("NOTE.BE.1")).TextString = ""
AttClent(AttribuCartouche("NOTE.BE.2")).TextString = ""
AttClent(AttribuCartouche("Num.VISA")).TextString = ""
'AttClent(AttribuCartouche("REF.PIECE.CART." & MyCARTOUCHE_Client)).TextString = Trim("" & Rs!RefP) & "_" & Trim("" & Rs!Ref_PF)
AttClent(AttribuCartouche("X/X")).TextString = Replace(CStr(Index) & "/" & CStr(NbCartouche), Chr(13), "")
Next Index

'AttClent(AttribuCartouche("DESGN.1.ANGL.CART.REN")).TextString = ""
    CartoucheCleient = True
    Set Fso = Nothing
End Function
Sub Maj(MyControl As Object)
Dim Rs As Recordset
Dim Sql As String
Dim Txt As String
Txt = MyControl.Text
MyControl.Clear
Sql = "SELECT T_Clients.Client, T_Clients.Formulaire "
Sql = Sql & "FROM T_Clients "
Sql = Sql & "ORDER BY T_Clients.Client;"
MyControl.AddItem ""
Set Rs = Con.OpenRecordSet(Sql)
While Rs.EOF = False
    MyControl.AddItem Trim("" & Rs!Client)
        If UCase(MyControl.List(MyControl.ListCount - 1)) = UCase(Txt) Then MyControl.ListIndex = MyControl.ListCount - 1
         
    Rs.MoveNext
Wend
Set Rs = Con.CloseRecordSet(Rs)

End Sub
Function ChercheXls(Val, Myrange2, Optional Cherche2 As Boolean) As Long
ChercheXls = 1

Dim RowSave As Long
Dim RageTrouve
ReTante:
On Error Resume Next
Set RageTrouve = Myrange2.Cells.Find(Val, Myrange2.Cells(ChercheXls, 1), ssValues, ssPart)  'Myrange2.Find(Val, Myrange2.Cells(1, 1), Myrange2.xlValues, Myrange2.xlPart)

ChercheXls = RageTrouve.Row

If RowSave > ChercheXls Then
    ChercheXls = 0
    On Error GoTo 0
    GoTo Fin
End If

If UCase(Trim("" & RageTrouve)) <> UCase(Trim("" & Val)) Then
If Trim("" & RageTrouve) = "" Then
    ChercheXls = 0
    GoTo Fin
End If
If RowSave = ChercheXls Then
    If UCase(Trim("" & RageTrouve)) <> UCase(Trim("" & Val)) Then
        ChercheXls = 0
        GoTo Fin
    End If
End If

RowSave = ChercheXls
ChercheXls = ChercheXls + 1
Set RageTrouve = Nothing
On Error GoTo 0
GoTo ReTante
End If

' For I = 2 To Myrange.Count
'                If UCase(Trim("" & Myrange(I))) = UCase(Trim("" & Val)) Then
'                If Cherche2 = True Then
'                    If Myrange2(I) = 1 Then
'                         ChercheXls = I
'                        Exit For
'                    End If
'                Else
'                        ChercheXls = I
'                        Exit For
'                End If
'                End If
'            Next I

Fin:
On Error GoTo 0
End Function
Function CherCheInFihier(Cherher As String) As String
Dim FileNumber As Long
Dim MyString As String
Dim Spliligne
FileNumber = FreeFile

  
Open App.Path & "\Autocable.ini" For Input As #FileNumber
Do While Not EOF(FileNumber)    ' Effectue la boucle jusqu'à la fin du fichier.
    Input #FileNumber, MyString ' Lit les données dans deux variables.
    If InStr(1, MyString, Cherher) <> 0 Then
    Spliligne = Split(MyString & "====", "=")
       CherCheInFihier = Trim(Spliligne(1))
'       CherCheInFihier = Right(MyString, Len(MyString) - (Len(Cherher) + InStr(1, MyString, Cherher)))
       Exit Do
    End If
Loop
Close #FileNumber    ' Ferme le fichier
CherCheInFihier = Replace(CherCheInFihier, "§remplaceDate§", Format(Date, "yyyy"))
CherCheInFihier = Trim(CherCheInFihier)
End Function
Function RenseigneConnecteurBroches(RangeAttribue As Recordset, Mytype As String) As Boolean
If (bool_Plan_E_Connecteurs = False And Mytype = "PL") Or (bool_Outil_E_Connecteurs = False And Mytype = "OU") Then Exit Function

Dim txtErr As String
    On Error Resume Next
    RenseigneConnecteurBroches = False
'    CollectionTor
If RangeAttribue!Activer = False Then GoTo Fin
    If TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).ConnecteurExiste = True Then
        If TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).Epissure = False Then
            a = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).NewBlock.GetAttributes
            txtErr = "LIAI"
            IdAttrib = ""
            IdAttrib = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).Attribues.Item("LIAI" & RangeAttribue.Fields("VOI"))
            If Err Then
                FunError 5, txtErr & RangeAttribue.Fields("VOI"), Err.Description, "" & RangeAttribue.Fields("APP")
                Err.Clear
           End If
                dd = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).Attribues.Count
                If Trim("" & IdAttrib) <> "" Then a(IdAttrib).TextString = "" & RangeAttribue.Fields("LIAI")
            
             txtErr = "FIL"
             IdAttrib = ""
            IdAttrib = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).Attribues.Item("FIL" & RangeAttribue.Fields("VOI"))
           If Err Then
                FunError 5, txtErr & RangeAttribue.Fields("VOI"), Err.Description, "" & RangeAttribue.Fields("APP")
                Err.Clear
           End If
            If Trim("" & a(IdAttrib).TextString) <> "" Then
                 txtErr = "MAR"
                 IdAttrib = ""
                IdAttrib = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).Attribues.Item("MAR" & RangeAttribue.Fields("VOI"))
            If Err Then
                FunError 5, txtErr & RangeAttribue.Fields("VOI"), Err.Description, "" & RangeAttribue.Fields("APP")
                Err.Clear
           End If
            End If
            If Trim("" & IdAttrib) <> "" Then a(IdAttrib).TextString = "" & RangeAttribue.Fields("FIL")
        Else
             FunEPISSURE TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP")))).NewBlock.GetAttributes, "" & RangeAttribue.Fields("FIL"), "" & RangeAttribue.Fields("VOI"), CLng(CollectionCon("" & RangeAttribue.Fields("APP")))
        End If

    Else
        FunError 3, "" & RangeAttribue.Fields("FIL"), Err.Description, "" & RangeAttribue.Fields("APP")
    End If
    Err.Clear
    If TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).ConnecteurExiste = True Then

        If TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).Epissure = False Then
            a = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).NewBlock.GetAttributes
             txtErr = "LIAI"
             IdAttrib = ""
            IdAttrib = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).Attribues.Item("LIAI" & RangeAttribue.Fields("VOI2"))
            If Err Then
                FunError 5, txtErr & RangeAttribue.Fields("VOI2"), Err.Description, "" & RangeAttribue.Fields("APP2")
                Err.Clear
            End If
        
            dd = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).Attribues.Count
           If Trim("" & IdAttrib) <> "" Then a(IdAttrib).TextString = "" & RangeAttribue.Fields("LIAI")
             txtErr = "FIL"
             IdAttrib = ""
            IdAttrib = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).Attribues.Item("FIL" & RangeAttribue.Fields("VOI2"))
            If Err Then
                FunError 5, txtErr & RangeAttribue.Fields("VOI2"), Err.Description, "" & RangeAttribue.Fields("APP2")
                Err.Clear
           End If
            If Trim("" & a(IdAttrib).TextString) <> "" Then
             txtErr = "MAR"
             IdAttrib = ""
                IdAttrib = TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).Attribues.Item("MAR" & RangeAttribue.Fields("VOI2"))
                If Err Then
                FunError 5, txtErr & RangeAttribue.Fields("VOI2"), Err.Description, "" & RangeAttribue.Fields("APP2")
                Err.Clear
           End If
            End If
           If Trim("" & IdAttrib) <> "" Then a(IdAttrib).TextString = "" & RangeAttribue.Fields("Fil")
        Else
            If FunEPISSURE(TableauDeConnecteurs(CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))).NewBlock.GetAttributes, "" & RangeAttribue.Fields("Fil"), "" & RangeAttribue.Fields("VOI2"), CLng(CollectionCon("" & RangeAttribue.Fields("APP2")))) = False Then Exit Function

        End If
    Else
        FunError 3, "" & RangeAttribue.Fields("Fil"), Err.Description, "" & RangeAttribue.Fields("APP2")
    End If
    RenseigneConnecteurBroches = True
Fin:
End Function
Sub AttribueLibV(NewBlock As Object, Index As Long)
    Dim At
    Dim MyAtt As New Collection
    At = NewBlock.GetAttributes
    
    For I = 0 To UBound(At)
        DoEvents
       
        Debug.Print At(I).TagString
        MyAtt.Add CStr(I), At(I).TagString
    Next I
    Set TableauDeConnecteurs(Index).AttribuesFils = MyAtt
    Ats = TableauDeConnecteurs(Index).NewBlock.GetAttributes
    At(TableauDeConnecteurs(Index).AttribuesFils("DESIGNATION")).TextString = Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")"
    Set MyAtt = Nothing
    I3 = UBound(Ats) - 1
    Dim I2 As Long
    
    For I = 1 To TableauDeConnecteurs(Index).indexFile
        DoEvents
       
        At(TableauDeConnecteurs(Index).AttribuesFils("FIL" & CStr(I))).TextString = TableauDeConnecteurs(Index).TableauFile(I)
    Next I
'    ReseingeTor "aaa"
End Sub
Function RetoureCadeApp(Ats) As String
RetoureCadeApp = ""
For I = 0 To UBound(Ats) - 1
        If Ats(I).TagString = "CODE_APP" Then
        RetoureCadeApp = Ats(I).TextString
        Exit Function
        End If
    Next I
End Function
Function RetournInsertEtiquette(Index As Integer, InsertionPoint)
Ats = TableauDeConnecteurs(Index).NewBlock.GetAttributes
        a = ""
        On Error Resume Next
           a = CollectionEtiquettes("E" & Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")")
                    If Err = 0 Then
                        RetournInsertEtiquette = TableuEtiquettes(CollectionEtiquettes("E" & Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")")).InsertTorTitre
                     Else
                       RetournInsertEtiquette = InsertionPoint
                    End If
        Err.Clear
                       
        On Error GoTo 0


End Function
Function RetourneXEtiquette(Index As Integer)
Ats = TableauDeConnecteurs(Index).NewBlock.GetAttributes
        a = ""
        On Error Resume Next
           a = CollectionEtiquettes("E" & Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")")
                    If Err = 0 Then
                        RetourneXEtiquette = TableuEtiquettes(CollectionEtiquettes("E" & Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")")).XScaleFactor
                     Else
                        RetourneXEtiquette = 1
                    End If
        Err.Clear
                       
        On Error GoTo 0


End Function
Function RetourneYEtiquette(Index As Integer)
Ats = TableauDeConnecteurs(Index).NewBlock.GetAttributes
        a = ""
        On Error Resume Next
           a = CollectionEtiquettes("E" & Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")")
                    If Err = 0 Then
                        RetourneYEtiquette = TableuEtiquettes(CollectionEtiquettes("E" & Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")")).YScaleFactor
                     Else
                        RetourneYEtiquette = 1
                    End If
        Err.Clear
                       
        On Error GoTo 0


End Function
Function RetourneZEtiquette(Index As Integer)
Ats = TableauDeConnecteurs(Index).NewBlock.GetAttributes
        a = ""
        On Error Resume Next
           a = CollectionEtiquettes("E" & Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")")
                    If Err = 0 Then
                        RetourneZEtiquette = TableuEtiquettes(CollectionEtiquettes("E" & Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")")).ZScaleFactor
                     Else
                        RetourneZEtiquette = 1
                    End If
        Err.Clear
                       
        On Error GoTo 0


End Function
Function RetourneRotationEtiquette(Index As Integer)
Ats = TableauDeConnecteurs(Index).NewBlock.GetAttributes
        a = ""
        On Error Resume Next
           a = CollectionEtiquettes("E" & Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")")
                    If Err = 0 Then
                        RetourneRotationEtiquette = TableuEtiquettes(CollectionEtiquettes("E" & Ats(TableauDeConnecteurs(Index).Attribues("DESIGNATION")).TextString & " (" & Ats(TableauDeConnecteurs(Index).Attribues("N°")).TextString & ")")).Rotation
                     Else
                        RetourneRotationEtiquette = 0
                    End If
        Err.Clear
                       
        On Error GoTo 0


End Function


Sub TestFl()
a = CherCheInFihier("Bdnumero")
End Sub
Function LoadComposants(IdIndiceProjet As Long, Mytype As String) As Boolean
If (bool_Plan_E_Composants = False And Mytype = "PL") Or (bool_Outil_E_Composants = False And Mytype = "OU") Then Exit Function

  LoadComposants = False
    Dim RsComposants As Recordset
    Dim Sql As String
    Dim Myrep As String
    Dim Trouve As Boolean
   Dim NbCol  As Long
  
    Dim Fso As New FileSystemObject
    Dim NumErr As Long
    Dim Nb_L_C As Long
   Dim XMin As Double
   Dim YMin As Double
    Dim PathComposantsDefault As String
     Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1
    
      Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1

    Set RsConnecteur = Con.OpenRecordSet(Sql)
    LeCient = UCase(Trim("" & RsConnecteur!Client))
    Sql = "SELECT  T_Clients.PathComposants FROM T_Clients "
Sql = Sql & "WHERE T_Clients.Client='" & MyReplace(LeCient) & "';"

       Set RsConnecteur = Con.OpenRecordSet(Sql)
If RsConnecteur.EOF = False Then
    
    If Trim("" & RsConnecteur!PathComposants) = "" Then
         PathComposantsDefault = TableauPath.Item("PathComposantsDefault")
   Else
             PathComposantsDefault = RsConnecteur!PathComposants
             
'         If Left(PathComposantsDefault, 2) <> "\\" And Left(PathComposantsDefault, 1) = "\" Then PathComposantsDefault = TableauPath.Item("PathServer") & PathComposantsDefault
'             If Right(PathComposantsDefault, 2) = "\\" Then PathComposantsDefault = Mid(PathComposantsDefault, 1, Len(PathComposantsDefault) - 1)
    
    End If
Else
                 PathComposantsDefault = TableauPath.Item("PathComposantsDefault")

End If
PathComposantsDefault = DefinirChemienComplet(TableauPath.Item("PathServer"), PathComposantsDefault)
'If Left(PathComposantsDefault, 2) <> "\\" And Left(PathComposantsDefault, 1) = "\" Then PathComposantsDefault = TableauPath.Item("PathServer") & PathComposantsDefault
'If Right(PathComposantsDefault, 2) = "\\" Then PathComposantsDefault = Mid(PathComposantsDefault, 1, Len(PathComposantsDefault) - 1)
' PathComposantsDefault = PathComposantsDefault & "COMPOSANTS\"
 
 Sql = "SELECT Composants.* FROM Composants "
 Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & IdIndiceProjet & " AND Composants.ACTIVER=True;"
 Set RsComposants = Con.OpenRecordSet(Sql)
 While RsComposants.EOF = False
 On Error Resume Next
                    a = ""
                   a = CollectionComp(Trim("C" & RsComposants!NUMCOMP))
                If Err Then
                    If NUMCOM < RsComposants!NUMCOMP Then
                         ReDim Preserve TableauComposant(RsComposants!NUMCOMP)
                         
                         NUMCOM = RsComposants!NUMCOMP
                    End If
                     CollectionComp.Add RsComposants("NUMCOMP").Value, Trim("C" & RsComposants!NUMCOMP)
                End If
 
    
    RsComposants.MoveNext
 Wend
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = 1 + NUMCOM
 FormBarGrah.ProgressBar1Caption.Caption = " Chargement des Composants"
 
  RsComposants.Requery
   XMin = 823.5964: YMin = -954.9939
    For I = 0 To IndexIstC
  
    InsertPointConnecteur(I).InsertPointConnecteur(0) = XMin - (150 * I): InsertPointConnecteur(I).InsertPointConnecteur(1) = YMin - (150 * I): InsertPointConnecteur(I).InsertPointConnecteur(2) = 0
    Next I
 On Error GoTo GesERR
  
 While RsComposants.EOF = False
  IncremanteBarGrah FormBarGrah
   IncrmentServer FormBarGrah, Mytype
 If TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).PosOkComp = False Then
  InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(0), -300)
                Nb_L_C = Nb_L_C + 1
                If Nb_L_C = IndexIstC + 1 Then Nb_L_C = 0
                rr = InsertPointHiérarchie(Val(Trim("" & RsComposants!NUMCOMP)), 0, -1000, 150, -150, 10)
    TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).InsertComp(0) = rr(0)
    TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).InsertComp(1) = rr(1)
    TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).InsertComp(2) = rr(2)
   
    TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).XScaleFactorComp = 1
    TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).YScaleFactorComp = 1
    TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).ZScaleFactorComp = 1
    TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).RotationComp = 0
 End If

'    PathComposantsDefault = "\\10.30.0.5\\donnees d entreprise\Utilitaires\cablage\Librairies\COMPOSANTS"
    Lib1 = PathComposantsDefault & "\" & RsComposants!Path & "\" & RsComposants!REFCOMP & ".dwg"
'    If Fso.FileExists(Lib1) = False Then
'            Set TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).BlockComp = Nothing
'    Set TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).BlockComp = FunInsBlock(PathComposantsDefault & "\" & RsComposants!Path & "\" & RsComposants!REFCOMP & ".dwg", TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).InsertComp, "", TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).RotationComp, TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).XScaleFactorComp, TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).YScaleFactorComp, TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).ZScaleFactorComp)
'
'    Else
    
    Lib2 = "" & RsComposants!REFCOMP
    NumErr = 6
    Set TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).BlockComp = Nothing
    Set TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).BlockComp = FunInsBlock(PathComposantsDefault & "\" & RsComposants!Path & "\" & RsComposants!REFCOMP & ".dwg", TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).InsertComp, "", TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).RotationComp, TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).XScaleFactorComp, TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).YScaleFactorComp, TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).ZScaleFactorComp)
    If ErrInsert = True Then GoTo Fin
    Err.Clear
    Set Attribues = ColectionAttribueConecteur(TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).BlockComp.GetAttributes)

                Att = TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).BlockComp.GetAttributes
                Lib1 = "DESIGNCOMP"
                Lib2 = "" & RsComposants!REFCOMP
                NumErr = 7
                Att(Attribues("DESIGNCOMP")).TextString = "" & RsComposants!DESIGNCOMP
                Err.Clear
                 Lib1 = "NUMCOMP"
                Lib2 = "" & RsComposants!REFCOMP
                Att(Attribues("NUMCOMP")).TextString = "C" & RsComposants!NUMCOMP
                Err.Clear
                 Lib1 = "PATHCOMP"
                Lib1 = "NUMCOMP"
                Att(Attribues("PATHCOMP")).TextString = "" & RsComposants!Path
                  Err.Clear
                 Lib1 = "REFCOMP"
                Lib2 = ""
                Att(Attribues("REFCOMP")).TextString = "" & RsComposants!REFCOMP
                
                
                
                
     If TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).PosOkDesin = False Then
''  InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(0), -300)
'                Nb_L_C = Nb_L_C + 1
'                If Nb_L_C = IndexIstC + 1 Then Nb_L_C = 0
    TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).InsertDesing(0) = TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).InsertComp(0) - 41.6692
    TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).InsertDesing(1) = TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).InsertComp(1) - 13.186
    TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).InsertDesing(2) = TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).InsertComp(2)
   
    TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).XScaleFactorDesin = 1
    TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).YScaleFactorDesin = 1
    TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).ZScaleFactorDesin = 1
    TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).RotationDesin = 0
 End If

'    PathComposantsDefault = "\\10.30.0.5\\donnees d entreprise\Utilitaires\cablage\Librairies\COMPOSANTS"
    Lib1 = PathComposantsDefault & "\" & RsComposants!Path & "\" & RsComposants!REFCOMP & ".dwg"
    Lib2 = "" & RsComposants!REFCOMP
    NumErr = 6
    Set TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).BlockDesing = FunInsBlock(PathComposantsDefault & "\COMP_DESGN.dwg", TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).InsertDesing, "", TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).RotationDesin, TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).XScaleFactorDesin, TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).YScaleFactorDesin, TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).ZScaleFactorDesin)
    Err.Clear
    Set Attribues = ColectionAttribueConecteur(TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).BlockDesing.GetAttributes)
If ErrInsert = True Then GoTo Fin
                Att = TableauComposant(CollectionComp("C" & RsComposants!NUMCOMP)).BlockDesing.GetAttributes
                Lib1 = "DESIGNCOMP"
                Lib2 = "" & RsComposants!REFCOMP
                NumErr = 7
                Att(Attribues("DESIGNCOMP")).TextString = "" & RsComposants!DESIGNCOMP
                Err.Clear
                 Lib1 = "NUMCOMP"
                Lib2 = "" & RsComposants!REFCOMP
                Att(Attribues("NUMCOMP")).TextString = "C" & RsComposants!NUMCOMP
                Err.Clear
                 Lib1 = "PATHCOMP"
                Lib1 = "NUMCOMP"
                Att(Attribues("PATHCOMP")).TextString = "" & RsComposants!Path
                  Err.Clear
                 Lib1 = "REFCOMP"
                Lib2 = ""
                Att(Attribues("REFCOMP")).TextString = "" & RsComposants!REFCOMP
'   End If
Fin:
    RsComposants.MoveNext
 Wend
 Exit Function
GesERR:
    FunError NumErr, "" & Lib1, Err.Description, "" & Lib2
 Resume Next
End Function

Function LoadNoeuds(IdIndiceProjet As Long, Mytype As String) As Boolean
If (bool_Plan_E_Noeuds = False And Mytype = "PL") Or (bool_Outil_E_Noeuds = False And Mytype = "OU") Then Exit Function
  LoadNoeuds = False
    Dim RsComposants As Recordset
    Dim Sql As String
    Dim Myrep As String
    Dim Trouve As Boolean
   Dim NbCol  As Long
    
    Dim Fso As New FileSystemObject
    Dim NumErr As Long
    Dim Nb_L_C As Long
    Dim XMin As Double
   Dim YMin As Double
    Dim PathNotasDefault As String
     Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1
    
      Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1

    Set RsConnecteur = Con.OpenRecordSet(Sql)
    LeCient = UCase(Trim("" & RsConnecteur!Client))
    Sql = "SELECT  T_Clients.PathNotas FROM T_Clients "
Sql = Sql & "WHERE T_Clients.Client='" & MyReplace(LeCient) & "';"

       Set RsConnecteur = Con.OpenRecordSet(Sql)
If RsConnecteur.EOF = False Then
    
    If Trim("" & RsConnecteur!PathNotas) = "" Then
         PathNotasDefault = TableauPath.Item("PathNotasDefault")
   Else
             PathNotasDefault = RsConnecteur!PathNotas
'         If Left(PathNotasDefault, 2) <> "\\" And Left(PathNotasDefault, 1) = "\" Then PathNotasDefault = TableauPath.Item("PathServer") & PathNotasDefault
'         If Right(PathNotasDefault, 2) = "\\" Then PathNotasDefault = Mid(PathNotasDefault, 1, Len(PathNotasDefault) - 1)
'
    End If
Else
                 PathNotasDefault = TableauPath.Item("PathNotasDefault")

End If
PathNotasDefault = DefinirChemienComplet(TableauPath.Item("PathServer"), PathNotasDefault)
PathBlocs = TableauPath.Item("PathBlocs")
PathBlocs = DefinirChemienComplet(TableauPath.Item("PathServer"), PathBlocs)
' If Left(PathBlocs, 2) <> "\\" And Left(PathBlocs, 1) = "\" Then PathBlocs = TableauPath.Item("PathServer") & PathBlocs
' If Right(PathBlocs, 2) = "\\" Then PathBlocs = Mid(PathBlocs, 1, Len(PathBlocs) - 1)
' PathNotasDefault = PathNotasDefault & "Nota\"
 
 Sql = "SELECT T_Noeuds.* FROM T_Noeuds "
 Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & IdIndiceProjet & " and T_Noeuds.ACTIVER=true "
 Sql = Sql & "order by T_Noeuds.id;"
 Set RsComposants = Con.OpenRecordSet(Sql)

' Set CollectionNoeuds = New Collection
 While RsComposants.EOF = False
 On Error Resume Next
                  
                   a = CollectionNoeuds(Trim("N" & RsComposants!NŒUDS))
                If Err Then
                NUMNOEUDS = NUMNOEUDS + 1
                    Err.Clear
                         ReDim Preserve TableauDeNoeuds(NUMNOEUDS)
                         CollectionNoeuds.Add NUMNOEUDS, Trim("N" & RsComposants!NŒUDS)
                         
                  
 
                End If
 
    
    RsComposants.MoveNext
 Wend
   
  RsComposants.Requery
   XMin = -1337.8928: YMin = 870.4179
    For I = 0 To IndexIstN
  
    InsertNouds(I).InsertPointConnecteur(0) = XMin - (50 * I): InsertNouds(I).InsertPointConnecteur(1) = YMin + (50 * I): InsertNouds(I).InsertPointConnecteur(2) = 0
    Next I
 On Error GoTo GesERR
 
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = 1 + NUMNOEUDS
 FormBarGrah.ProgressBar1Caption.Caption = " Chargement des Noeuds"
 DoEvents
 
    
 While RsComposants.EOF = False
  IncremanteBarGrah FormBarGrah
  IncrmentServer FormBarGrah, Mytype
 If TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).PosOkComp = False Then
  zz = IndexationNoeuds(RsComposants!Noeuds)
   rr = InsertPointHiérarchie(Val(zz), -100, -1000, -50, -40, 10)
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).InsertComp(0) = rr(0)
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).InsertComp(1) = rr(1)
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).InsertComp(2) = rr(2)
   
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).XScaleFactorComp = 1
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).YScaleFactorComp = 1
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).ZScaleFactorComp = 1
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).RotationComp = 0
     InsertNouds(Nb_L_C).InsertPointConnecteur(1) = DecalInsertPointLigneTableau_fils_Gauche(InsertNouds(Nb_L_C).InsertPointConnecteur(1), -40)
                Nb_L_C = Nb_L_C + 1
                If Nb_L_C = IndexIstN + 1 Then
                    Nb_L_C = 0
                    For I = 0 To IndexIstN
                       InsertNouds(I).InsertPointConnecteur(0) = InsertNouds(I).InsertPointConnecteur(0) - 50: InsertNouds(I).InsertPointConnecteur(1) = YMin + (50 * I): InsertNouds(I).InsertPointConnecteur(2) = 0
                     Next
                End If
 End If

    Lib1 = PathBlocs & "\NOEUD.dwg"
    Lib2 = "" & RsComposants!Noeuds
    NumErr = 6
  
    Set TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).BlockComp = FunInsBlock(Lib1, TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).InsertComp, "", TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).RotationComp, TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).XScaleFactorComp, TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).YScaleFactorComp, TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).ZScaleFactorComp)
    Set Attribues = ColectionAttribueConecteur(TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).BlockComp.GetAttributes)

                Att = TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).BlockComp.GetAttributes
                
                NumErr = 7
                
                Att(Attribues("LONG")).TextString = "" & RsComposants!Longueur
                Att(Attribues("NOEUD")).TextString = "" & RsComposants!NŒUDS
                Att(Attribues("DIAM")).TextString = "" & RsComposants!DIAMETRE
                Att(Attribues("HAB")).TextString = "" & RsComposants!CODE_ENC
                Att(Attribues("CLASSE_T")).TextString = "" & RsComposants!CLASSE_T
                Err.Clear
                
                
                
 
If TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).PosOkDesin = False Then
 
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).InsertDesing(0) = TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).InsertComp(0) + 10
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).InsertDesing(1) = TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).InsertComp(1)
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).InsertDesing(2) = TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).InsertComp(2)
   
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).XScaleFactorDesin = 1
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).YScaleFactorDesin = 1
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).ZScaleFactorDesin = 1
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).RotationDesin = 0
 End If
If Trim("" & RsComposants!NŒUDS) <> "AA" Then
'    PathNotasDefault = "\\10.30.0.5\\donnees d entreprise\Utilitaires\cablage\Librairies\Nota"
'If "" & RsComposants!NŒUDS = "AA" Then
'    Lib1 = PathBlocs & "\NOEUD_0.dwg"
'Else
    Lib1 = PathBlocs & "\NOEUD_LONG.dwg"
'End If
    Lib2 = "" & RsComposants!Noeuds
    NumErr = 6
    Set TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).BlockDesing = FunInsBlock(Lib1, TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).InsertDesing, "", TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).RotationDesin, TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).XScaleFactorDesin, TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).YScaleFactorDesin, TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).ZScaleFactorDesin)
    Set Attribues = ColectionAttribueConecteur(TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).BlockDesing.GetAttributes)

                Att = TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).BlockDesing.GetAttributes
                
                NumErr = 7
                
'                 Lib1 = "NUMNOTA"
'                Lib2 = "" & RsComposants!Nota
                Att(Attribues("LONG")).TextString = "" & RsComposants!Longueur
                Att(Attribues("NOEUD")).TextString = "" & RsComposants!NŒUDS
                Att(Attribues("DIAM")).TextString = "" & RsComposants!DIAMETRE
                Att(Attribues("HAB")).TextString = "" & RsComposants!CODE_ENC
                Att(Attribues("CLASSE_T")).TextString = "" & RsComposants!CLASSE_T
                Err.Clear
 End If
 
                

If TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).PosOkFleche = False Then
 
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).InsertFleche(0) = TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).InsertComp(0)
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).InsertFleche(1) = TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).InsertComp(1) + 20
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).InsertFleche(2) = TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).InsertComp(2)
   
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).XScaleFactorFleche = 1
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).YScaleFactorFleche = 1
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).ZScaleFactorFleche = 1
    TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).RotationFleche = 0
 End If

'    PathNotasDefault = "\\10.30.0.5\\donnees d entreprise\Utilitaires\cablage\Librairies\Nota"
'If "" & RsComposants!NŒUDS = "AA" Then
'    Lib1 = PathBlocs & "\NOEUD_0.dwg"
'Else
If RsComposants!TORON_PRINCIPAL = True Then

    Lib1 = PathBlocs & "\NOEUD_PRINCIPAL"
Else
     Lib1 = PathBlocs & "\NOEUD_SECONDAIRE" '.dwg"
    End If
    If RsComposants!Fleche_Droite = False Then
        Lib1 = Lib1 & ".dwg"
    Else
        Lib1 = Lib1 & "1.dwg"
    End If
'End If
    Lib2 = "" & RsComposants!Noeuds
    NumErr = 6
    Set TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).BlockFleche = FunInsBlock(Lib1, TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).InsertFleche, "", TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).RotationFleche, TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).XScaleFactorFleche, TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).YScaleFactorFleche, TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).ZScaleFactorFleche)
    Set Attribues = ColectionAttribueConecteur(TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).BlockFleche.GetAttributes)

                Att = TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).BlockFleche.GetAttributes
                
                NumErr = 7
                
'                 Lib1 = "NUMNOTA"
'                Lib2 = "" & RsComposants!Nota
                Att(Attribues("LONG")).TextString = "" & RsComposants!Longueur
                Att(Attribues("NOEUD")).TextString = "" & RsComposants!NŒUDS
                Att(Attribues("DIAM")).TextString = "" & RsComposants!DIAMETRE
                Att(Attribues("HAB")).TextString = "" & RsComposants!CODE_ENC
                Att(Attribues("CLASSE_T")).TextString = "" & RsComposants!CLASSE_T
                Att(Attribues("LONG_CUMUL")).TextString = "" & RsComposants!LONGUEUR_CUMULEE
                Err.Clear

'                 Lib1 = "NOTA"
'                Lib2 = ""
'                Att(TableauDeNoeuds(CollectionNoeuds("N" & RsComposants!Noeuds)).Attribues("NOTA")).TextString = "" & RsComposants!NOTA
    RsComposants.MoveNext
 Wend
 Exit Function
GesERR:
    FunError NumErr, "" & Lib1, Err.Description, "" & Lib2
 Resume Next
End Function

Function LoadNotas(IdIndiceProjet As Long, Mytype As String) As Boolean
If (bool_Plan_E_Notas = False And Mytype = "PL") Or (bool_Outil_E_Notas = False And Mytype = "OU") Then Exit Function

  LoadNotas = False
    Dim RsComposants As Recordset
    Dim Sql As String
    Dim Myrep As String
    Dim Trouve As Boolean
   Dim NbCol  As Long
    
    Dim Fso As New FileSystemObject
    Dim NumErr As Long
    Dim Nb_L_C As Long
   Dim XMin As Double
   Dim YMin As Double
    Dim PathNotasDefault As String
     Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1
    
      Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1

    Set RsConnecteur = Con.OpenRecordSet(Sql)
    LeCient = UCase(Trim("" & RsConnecteur!Client))
    Sql = "SELECT  T_Clients.PathNotas FROM T_Clients "
Sql = Sql & "WHERE T_Clients.Client='" & MyReplace(LeCient) & "';"

       Set RsConnecteur = Con.OpenRecordSet(Sql)
If RsConnecteur.EOF = False Then
    
    If Trim("" & RsConnecteur!PathNotas) = "" Then
         PathNotasDefault = TableauPath.Item("PathNotasDefault")
   Else
             PathNotasDefault = RsConnecteur!PathNotas
'         If Left(PathNotasDefault, 2) <> "\\" And Left(PathNotasDefault, 1) = "\" Then PathNotasDefault = TableauPath.Item("PathServer") & PathNotasDefault
'         If Right(PathNotasDefault, 2) = "\\" Then PathNotasDefault = Mid(PathNotasDefault, 1, Len(PathNotasDefault) - 1)
'
    End If
Else
                 PathNotasDefault = TableauPath.Item("PathNotasDefault")

End If
PathNotasDefault = DefinirChemienComplet(TableauPath.Item("PathServer"), PathNotasDefault)
'If Left(PathNotasDefault, 2) <> "\\" And Left(PathNotasDefault, 1) = "\" Then PathNotasDefault = TableauPath.Item("PathServer") & PathNotasDefault
'If Right(PathNotasDefault, 2) = "\\" Then PathNotasDefault = Mid(PathNotasDefault, 1, Len(PathNotasDefault) - 1)
' PathNotasDefault = PathNotasDefault & "Nota\"
 
 Sql = "SELECT Nota.* FROM Nota "
 Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & IdIndiceProjet & "  AND Nota.ACTIVER=True "
 Sql = Sql & "order by Nota.NUMNOTA;"
 Set RsComposants = Con.OpenRecordSet(Sql)
 While RsComposants.EOF = False
 On Error Resume Next
                    a = ""
                   a = CollectionNota(Trim("N" & RsComposants!NUMNOTA))
                If Err Then
                
                    If NUMNOTA < RsComposants!NUMNOTA Then
                         ReDim Preserve TableauDeNotas(RsComposants!NUMNOTA)
                         
                         NUMNOTA = RsComposants!NUMNOTA
                    End If
                    CollectionNota.Add RsComposants("NUMNOTA").Value, Trim("N" & RsComposants!NUMNOTA)
                End If
 
    
    RsComposants.MoveNext
 Wend
   
  RsComposants.Requery
   XMin = -1498.3061: YMin = 482.3797
    For I = 0 To IndexIstC
  
    InsertPointConnecteur(I).InsertPointConnecteur(0) = XMin - (600 * I): InsertPointConnecteur(I).InsertPointConnecteur(1) = YMin + (600 * I): InsertPointConnecteur(I).InsertPointConnecteur(2) = 0
    Next I
 On Error GoTo GesERR
 
 FormBarGrah.ProgressBar1.Value = 0
 FormBarGrah.ProgressBar1.Max = 1 + NUMNOTA
 FormBarGrah.ProgressBar1Caption.Caption = " Chargement des Notas"
 
 While RsComposants.EOF = False
  IncremanteBarGrah FormBarGrah
  IncrmentServer FormBarGrah, Mytype
 If TableauDeNotas(CollectionNota("N" & RsComposants!NUMNOTA)).NotasExiste = False Then
 rr = InsertPointHiérarchie(Val(Trim("" & RsComposants!NUMNOTA)), -1300, -800, -600, 600, 4)
  InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(1) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(1), -600)
                Nb_L_C = Nb_L_C + 1
                If Nb_L_C = IndexIstC + 1 Then Nb_L_C = 0
    TableauDeNotas(CollectionNota("N" & RsComposants!NUMNOTA)).InsertPointLigneC(0) = rr(0)
    TableauDeNotas(CollectionNota("N" & RsComposants!NUMNOTA)).InsertPointLigneC(1) = rr(1)
    TableauDeNotas(CollectionNota("N" & RsComposants!NUMNOTA)).InsertPointLigneC(2) = rr(2)
   
    TableauDeNotas(CollectionNota("N" & RsComposants!NUMNOTA)).XScaleFactorC = 1
    TableauDeNotas(CollectionNota("N" & RsComposants!NUMNOTA)).YScaleFactorC = 1
    TableauDeNotas(CollectionNota("N" & RsComposants!NUMNOTA)).ZScaleFactorC = 1
    TableauDeNotas(CollectionNota("N" & RsComposants!NUMNOTA)).RotationC = 0
 End If

'    PathNotasDefault = "\\10.30.0.5\\donnees d entreprise\Utilitaires\cablage\Librairies\Nota"
    Lib1 = PathNotasDefault & "\" & RsComposants!Nota & ".dwg"
    Lib2 = "" & RsComposants!Nota
    NumErr = 6
    Set TableauDeNotas(CollectionNota("N" & RsComposants!NUMNOTA)).NewBlock = FunInsBlock(PathNotasDefault & "\" & RsComposants!Nota & ".dwg", TableauDeNotas(CollectionNota("N" & RsComposants!NUMNOTA)).InsertPointLigneC, "", TableauDeNotas(CollectionNota("N" & RsComposants!NUMNOTA)).RotationC, TableauDeNotas(CollectionNota("N" & RsComposants!NUMNOTA)).XScaleFactorC, TableauDeNotas(CollectionNota("N" & RsComposants!NUMNOTA)).YScaleFactorC, TableauDeNotas(CollectionNota("N" & RsComposants!NUMNOTA)).ZScaleFactorC)
    Set TableauDeNotas(CollectionNota("N" & RsComposants!NUMNOTA)).Attribues = ColectionAttribueConecteur(TableauDeNotas(CollectionNota("N" & RsComposants!NUMNOTA)).NewBlock.GetAttributes)

                Att = TableauDeNotas(CollectionNota("N" & RsComposants!NUMNOTA)).NewBlock.GetAttributes
                
                NumErr = 7
                
                 Lib1 = "NUMNOTA"
                Lib2 = "" & RsComposants!Nota
                Att(TableauDeNotas(CollectionNota("N" & RsComposants!NUMNOTA)).Attribues("NUMNOTA")).TextString = "" & RsComposants!NUMNOTA
                Err.Clear
'                 Lib1 = "NOTA"
'                Lib2 = ""
'                Att(TableauDeNotas(CollectionNota("N" & RsComposants!NUMNOTA)).Attribues("NOTA")).TextString = "" & RsComposants!NOTA
    RsComposants.MoveNext
 Wend
 Exit Function
GesERR:
    FunError NumErr, "" & Lib1, Err.Description, "" & Lib2
 Resume Next
End Function

Function LoadConnecteur(IdIndiceProjet As Long, Mytype As String) As Boolean
If (bool_Plan_E_Connecteurs = False And Mytype = "PL") Or (bool_Outil_E_Connecteurs = False And Mytype = "OU") Then Exit Function
    LoadConnecteur = False
    Dim RsConnecteur As Recordset
    Dim Sql As String
    Dim Myrep As String
    Dim Trouve As Boolean
   Dim NbCol  As Long
  
    Dim Fso As New FileSystemObject
    Dim NumErr As Long
    Dim Nb_L_C As Long
   Dim XMin As Double
   Dim YMin As Double
    Dim PathConnecteursDefault As String
  PathBlocs = TableauPath.Item("PathBlocs")
  PathBlocs = DefinirChemienComplet(TableauPath.Item("PathServer"), PathBlocs)
' If Left(PathBlocs, 2) <> "\\" And Left(PathBlocs, 1) = "\" Then PathBlocs = TableauPath.Item("PathServer") & PathBlocs
' If Right(PathBlocs, 2) = "\\" Then PathBlocs = Mid(PathBlocs, 1, Len(PathBlocs) - 1)
    Sql = "SELECT T_indiceProjet.Client FROM T_indiceProjet "
    Sql = Sql & "WHERE T_indiceProjet.Id=" & IdIndiceProjet & ";"

    NumErr = 1

    Set RsConnecteur = Con.OpenRecordSet(Sql)
    LeCient = UCase(Trim("" & RsConnecteur!Client))
    Sql = "SELECT T_Clients.Client, T_Clients.PathConnecteurs FROM T_Clients "
Sql = Sql & "WHERE T_Clients.Client='" & MyReplace(LeCient) & "';"

       Set RsConnecteur = Con.OpenRecordSet(Sql)
If RsConnecteur.EOF = False Then
    
    If Trim("" & RsConnecteur!PathConnecteurs) = "" Then
         PathConnecteursDefault = TableauPath.Item("PathConnecteursDefault")
   Else
             PathConnecteursDefault = RsConnecteur!PathConnecteurs
'         If Left(PathConnecteursDefault, 2) <> "\\" And Left(PathConnecteursDefault, 1) = "\" Then PathConnecteursDefault = TableauPath.Item("PathServer") & PathConnecteursDefault
'         If Right(PathConnecteursDefault, 2) = "\\" Then PathConnecteursDefault = Mid(PathConnecteursDefault, 1, Len(PathConnecteursDefault) - 1)
    
    End If
Else
                 PathConnecteursDefault = TableauPath.Item("PathConnecteursDefault")

End If
'If Left(PathConnecteursDefault, 2) <> "\\" And Left(PathConnecteursDefault, 1) = "\" Then PathConnecteursDefault = TableauPath.Item("PathServer") & PathConnecteursDefault
'If Right(PathConnecteursDefault, 2) = "\\" Then PathConnecteursDefault = Mid(PathConnecteursDefault, 1, Len(PathConnecteursDefault) - 1)

 PathConnecteursDefault = DefinirChemienComplet(TableauPath.Item("PathServer"), PathConnecteursDefault)

 Sql = "SELECT Connecteurs.CONNECTEUR, Connecteurs.[O/N],  "
Sql = Sql & "Connecteurs.DESIGNATION, Connecteurs.CODE_APP,  "
Sql = Sql & "Connecteurs.N°, Connecteurs.POS, Connecteurs.PRECO1,  "
Sql = Sql & "Connecteurs.PRECO2,Connecteurs.RefConnecteurFour  "
Sql = Sql & "FROM Connecteurs "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & IdIndiceProjet & " and Connecteurs.ACTIVER=true "
    Sql = Sql & "ORDER BY Connecteurs.N°;"
    NumErr = 1
    If Mytype = "OU" Then
        XMin = Val(GetDefault("Connecteur OU XMin", "1185.771"))
        YMin = Val(GetDefault("Connecteur OU YMin", "1667.3509"))
    Else
         XMin = Val(GetDefault("Connecteur PL XMin", "30"))
        XMin = Val(GetDefault("Connecteur Pl XMin", "870.4179"))
        
    End If
    Set RsConnecteur = Con.OpenRecordSet(Sql)
    InsertPointLigneTableau_Vignette(0) = Val(GetDefault("Vignette X", "-1146.1429")): InsertPointLigneTableau_Vignette(1) = Val(GetDefault("Vignette Y", "870.4179")): InsertPointLigneTableau_Vignette(2) = 0
    For I = 0 To IndexIstC
  
    InsertPointConnecteur(I).InsertPointConnecteur(0) = XMin + (150 * I): InsertPointConnecteur(I).InsertPointConnecteur(1) = YMin + (150 * I): InsertPointConnecteur(I).InsertPointConnecteur(2) = 0
    Next I
    I = 1
      NbConnecteur = 0
    For I = 1 To CollectionCon.Count
      If NbConnecteur < CollectionCon(I) Then
        NbConnecteur = CollectionCon(I)
      End If
    Next I
    I = 1
    While RsConnecteur.EOF = False
   
    If Trim(UCase("" & RsConnecteur.Fields(0))) <> "NEANT" Then
   
        On Error Resume Next
        NbCol = CLng(CollectionCon("" & RsConnecteur.Fields(3)))
        If Err Then
            Err.Clear
            
           NbConnecteur = NbConnecteur + 1
                CollectionCon.Add NbConnecteur, Trim("" & RsConnecteur.Fields(3))
            
        End If
         
         On Error GoTo 0
     Else
        NbConnecteur = NbConnecteur + 1
     End If
     
  
       RsConnecteur.MoveNext
    Wend
  
    ReDim Preserve TableauDeConnecteurs(NbConnecteur)
     FormBarGrah.ProgressBar1.Value = 0
    If NbConnecteur = 0 Then
     FormBarGrah.ProgressBar1.Max = 1 + 1
    Else
         FormBarGrah.ProgressBar1.Max = 1 + NbConnecteur
    End If
     FormBarGrah.ProgressBar1Caption.Caption = " Chargement des connecteurs"
    If NbConnecteur <> 0 Then
        RsConnecteur.Requery
    End If
      On Error GoTo GesERR
      FormBarGrah.ProgressBar1.Value = 0
       
    While RsConnecteur.EOF = False
'        If FormBarGrah.ProgressBar1.Max = FormBarGrah.ProgressBar1.Value Then
'            FormBarGrah.ProgressBar1.Max = FormBarGrah.ProgressBar1.Max + 1
'        End If
         IncremanteBarGrah FormBarGrah
        IncrmentServer FormBarGrah, Mytype
         
'       DoEvents
        
        DoEvents


   
    
        If UCase("" & RsConnecteur.Fields(0)) <> "NEANT" Then
        Debug.Print PathConnecteursDefault & "\" & RsConnecteur.Fields(0) & ".dwg"
            If Fso.FileExists(PathConnecteursDefault & "\" & RsConnecteur.Fields(0) & ".dwg") = True Then
                Myrep = PathConnecteursDefault
                Trouve = True
                NumErr = 4
              
                
                TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ConnecteurExiste = True
            Else
                NumErr = 1
                TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ConnecteurExiste = False
                Myrep = ""
                
GesERR:
'    If NumErr = 4 Then MsgBox "Err : " & NumErr
                Trouve = False
                FunError NumErr, "" & RsConnecteur.Fields(3), Err.Description, "" & RsConnecteur.Fields(0)
              
            End If
            If Trouve = True Then
            If TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).PosOk = True Then
                Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock = FunInsBlock(Myrep & "\" & RsConnecteur.Fields(0) & ".dwg", TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).InsertPointLigneC, "", TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).RotationC, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).XScaleFactorC, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).YScaleFactorC, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ZScaleFactorC)
            Else
                rr = InsertPointHiérarchie(Val(Trim("" & RsConnecteur.Fields(4))), Val(GetDefault("Connecteur X init", "0")), Val(GetDefault("Connecteur y init", "2000")), Val(GetDefault("Connecteur Decal X", "200")), Val(GetDefault("Connecteur Decal Y", "300")), Val(GetDefault("Connecteur Nb Bloc", "10")))
                Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock = FunInsBlock(Myrep & "\" & RsConnecteur.Fields(0) & ".dwg", rr, "", Options:="tous")
                InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointConnecteur(Nb_L_C).InsertPointConnecteur(0), Val(GetDefault("Connecteur Decal Droite", "-300")))
                Nb_L_C = Nb_L_C + 1
                If Nb_L_C = IndexIstC + 1 Then Nb_L_C = 0
             End If
                  If ErrInsert = True Then GoTo EnrSuinant
                If UCase("" & RsConnecteur.Fields(1)) = True Then
                    TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).Epissure = True
                    If TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).PosOk = True Then
                     If (bool_Plan_E_Etiquettes = True And Mytype = "PL") Or (bool_Outil_E_Vignettes = True And Mytype = "OU") Then
                        Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette = FunInsBlock(PathBlocs & "\EPISSURES.dwg", TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).InsertPointLigneV, "V" & "" & RsConnecteur.Fields(4), TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).RotationV, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).XScaleFactorV, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).YScaleFactorV, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ZScaleFactorV)
                    End If
                    Else
                         If (bool_Plan_E_Etiquettes = True And Mytype = "PL") Or (bool_Outil_E_Vignettes = True And Mytype = "OU") Then
                            rr = InsertPointHiérarchie(Val(Trim("" & RsConnecteur.Fields(4))), Val(GetDefault("Connecteur Ettiqette X init", "-100")), Val(GetDefault("Connecteur Ettiqette Y init", "2000")), Val(GetDefault("Connecteur Ettiqette Decal X", "-135")), Val(GetDefault("Connecteur Ettiqette Decal Y", "400")), Val(GetDefault("Connecteur Nb Bloc", "10")))
                        Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette = FunInsBlock(PathBlocs & "\EPISSURES.dwg", rr, "V" & "" & RsConnecteur.Fields(4))
                         
'                        InsertPointLigneTableau_Vignette(0) = DecalInsertPointLigneTableau_fils_Gauche(rr, 100)
                        NbLignesVignette = NbLignesVignette + 1
                        End If
                    End If
                Else
                    If (bool_Plan_E_Etiquettes = True And Mytype = "PL") Or (bool_Outil_E_Vignettes = True And Mytype = "OU") Then

                        TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).Epissure = False
                        If TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).PosOk = True Then
                            Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette = FunInsBlock(PathBlocs & "\VIGNETTE CONNECTEUR.dwg", TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).InsertPointLigneV, "V" & "" & RsConnecteur.Fields(4), TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).RotationV, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).XScaleFactorV, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).YScaleFactorV, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).ZScaleFactorV)
                        Else
'                             rr = InsertPointHiérarchie(Val(Trim("" & RsConnecteur.Fields(4))), -100, 2000, -135, 400, 10)
                              rr = InsertPointHiérarchie(Val(Trim("" & RsConnecteur.Fields(4))), Val(GetDefault("Connecteur Ettiqette X init", "-100")), Val(GetDefault("Connecteur Ettiqette Y init", "2000")), Val(GetDefault("Connecteur Ettiqette Decal X", "-135")), Val(GetDefault("Connecteur Ettiqette Decal Y", "400")), Val(GetDefault("Connecteur Nb Bloc", "10")))
                              
                            Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette = FunInsBlock(PathBlocs & "\VIGNETTE CONNECTEUR.dwg", rr, "V" & "" & RsConnecteur.Fields(4))
                        End If
                            
                            InsertPointLigneTableau_Vignette(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointLigneTableau_Vignette(0), Val(GetDefault("Connecteur Ettiqette Decal X", "-135")))
                            NbLignesVignette = NbLignesVignette + 1

                    End If
                End If
                Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).Attribues = ColectionAttribueConecteur(TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock.GetAttributes)
                At = TableauAtribCon(RsConnecteur, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).Epissure)
                funAttributesLigne_Tableau_fils TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock.Name, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewBlock.GetAttributes, At, "" & RsConnecteur.Fields.Count - 1, RsConnecteur, True, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).Attribues

                If (bool_Plan_E_Etiquettes = True And Mytype = "PL") Or (bool_Outil_E_Vignettes = True And Mytype = "OU") Then
                Set TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).AttribuesVignette = ColectionAttribueConecteur(TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette.GetAttributes)
                funAttributesLigne_Tableau_fils TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette.Name, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).NewVignette.GetAttributes, At, "" & RsConnecteur.Fields.Count - 1, RsConnecteur, True, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).AttribuesVignette, True, TableauDeConnecteurs(CLng(CollectionCon("" & RsConnecteur.Fields(3)))).Epissure
                End If
                
                
            End If
        End If
        If (bool_Plan_E_Etiquettes = True And Mytype = "PL") Or (bool_Outil_E_Vignettes = True And Mytype = "OU") Then
        If NbLignesVignette = 11 Then
            InsertPointLigneTableau_Vignette(0) = Val(GetDefault("Vignette X", "-1146.1429"))
            InsertPointLigneTableau_Vignette(1) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_Vignette(1), Val(GetDefault("Vignette Decale Bas", "-400")))
            NbLignesVignette = 0
         End If
        End If
EnrSuinant:
       RsConnecteur.MoveNext
        I = I + 1
    Wend
    LoadConnecteur = True
    Set Fso = Nothing
End Function
Function TableauAtribCon(MyAtrib As Recordset, Epissure As Boolean)
    Dim TabAt() As String
    ReDim TabAt(MyAtrib.Fields.Count)
    For Col = 0 To MyAtrib.Fields.Count - 1
    DoEvents
        If (Col = 0) And (Epissure = True) Then
            TabAt(Col) = "EPISSURE"
        Else
            TabAt(Col) = "" & MyAtrib.Fields(Col)
        End If
    Next Col
    TableauAtribCon = TabAt
End Function



Function PRECO(Var As String, Optional Iis As Boolean) As String
PRECO = Var
PRECO = Replace(UCase(PRECO), "CODE.APP", "CODE_APP")

If InStr(1, UCase(PRECO), "PRECO") <> 0 Then
    PRECO = "PRECO" & Right(PRECO, 1)
    
End If
If Iis = True Then

If (InStr(1, UCase(Var), "PRECO") <> 0) And (InStr(1, UCase(Var), "1") <> 0) Then
    PRECO = "PRECO1"
    
End If
If (InStr(1, UCase(Var), "PRECO") <> 0) And (InStr(1, UCase(Var), "2") <> 0) Then
    PRECO = "PRECO2"
    
End If
    If InStr(1, UCase(PRECO), "FIL") <> 0 Then
        PRECO = "FIL"
    
    End If
    
    If InStr(1, UCase(Var), "_FILS") <> 0 Then
        PRECO = "_FILS"
        
    End If
End If
End Function

Function DecodeCode_APP(Code_APP As String) As String
Dim SplitCode_APP
Dim NbUbound As Long
SlitCode_APP = Split(Code_APP, ".")
NbUbound = UBound(SlitCode_APP)
Select Case NbUbound
Case -1
    DecodeCode_APP = vbNullChar
Case 0
    DecodeCode_APP = SlitCode_APP(0)
Case 1
    DecodeCode_APP = SlitCode_APP(0)
Case 2
    DecodeCode_APP = SlitCode_APP(1)
Case Else
     DecodeCode_APP = SlitCode_APP(1)
    End Select
End Function
Sub test()
a = DecodeCode_APP("exzf_120_aa_ee")

a = DecodeCode_APP("exzf_120_aa")
a = DecodeCode_APP("120_aa")
a = DecodeCode_APP("120")
a = DecodeCode_APP("")

End Sub
Function funPath()
    Dim MyPath As New Collection
    Dim Rs As Recordset
        Set Rs = Con.OpenRecordSet("SELECT T_Path.* FROM T_Path;")
    While Rs.EOF = False
        MyPath.Add Rs.Fields("PathVar").Value, Rs.Fields("NameVar").Value
        Rs.MoveNext
    Wend
    Set Rs = Con.CloseRecordSet(Rs)
    Set funPath = MyPath
End Function
Function ChoixCouleur(Mode As Long, Optional BoolExcel As Boolean) As Long
   
  If BoolExcel = False Then
   Select Case Mode
   Case 0
        ChoixCouleur = 12632256
    Case 1
        ChoixCouleur = 16777164
    Case 2
    ChoixCouleur = 10079487
    Case 3
        ChoixCouleur = 13434828
    Case 4
        ChoixCouleur = &HFFC0FF
   End Select

Else
    Select Case Mode
    Case 0
        ChoixCouleur = 15
    Case 1
        ChoixCouleur = 34
    Case 2
    ChoixCouleur = 40
    Case 3
        ChoixCouleur = 35
    Case 4
        ChoixCouleur = 38
   End Select
End If
End Function
Function ColectionAttribueConecteur(Attribues) As Collection
    Dim MyAttribue As New Collection
    Dim IndexAt As Long


    IndexAt = 0
    On Error Resume Next
    For IndexAt = 0 To UBound(Attribues)
        
        Debug.Print UCase(Replace(Replace(UCase(Attribues(IndexAt).TagString), "PRECO", "PRECO."), "PRECO..", "PRECO."))
        MyAttribue.Add IndexAt, UCase(Replace(Replace(UCase(Attribues(IndexAt).TagString), "PRECO", "PRECO."), "PRECO..", "PRECO."))
        Set Atr = Nothing
        
     Next
    On Error GoTo 0
    Set ColectionAttribueConecteur = New Collection
    Set ColectionAttribueConecteur = MyAttribue
End Function
'Public Function RetournIdApp(Application As String, Optional Retourn As Boolean) As Long
'Dim Liste
'Dim element
'Set Liste = GetObject("winmgmts:").InstancesOf("Win32_Process")
'
'If Retourn = False Then
'Set ColecAplication = Nothing
'Set ColecAplication = New Collection
'
'
'For Each element In Liste
'    Debug.Print element.Name
'    If UCase(element.Name) = UCase(Application) Then
'        ColecAplication.Add element.Handle, element.Handle
'    End If
'Next element
'Else
'    On Error Resume Next
'    For Each element In Liste
'    Debug.Print element.Name & " : " & element.Handle
'    If UCase(element.Name) = UCase(Application) Then
'        Valid = ColecAplication(element.Handle)
'        If Err Then
'            Err.Clear
'                RetournIdApp = element.Handle
'                Exit For
'        End If
'    End If
'Next element
'End If
'
'End Function




Function FunEPISSURE(Attribues, Fil, Valeur, Connecteur As Long) As Boolean
    FunEPISSURE = False
    Dim bollInDif As Boolean
    Dim IbAttribue As Long
    Dim Fils As String
    Dim TouveFil As Boolean
    Dim boolNotExecute As Boolean
    Dim booD As Boolean
    Dim booG As Boolean
    On Error GoTo Fin
Valeur = Valeur & Space(50)
    bollInDif = True
    Fils = "FILG"
    If UCase(Left(Valeur, 1)) = "D" Then
        boolNotExecute = True
        booD = True
    End If
    If UCase(Left(Valeur, 1)) = "G" Then
       
        booG = True
    End If
    Valeur = Trim(Valeur)
    For I = 1 To UBound(Attribues)
        DoEvents
        
        IbAttribue = TableauDeConnecteurs(Connecteur).Attribues.Item(Fils & CStr(I))
        If (Trim("" & Attribues(IbAttribue).TextString) = "") And (boolNotExecute = False) Then
        Exit For
        End If
Retour:
    Next I
    Attribues(IbAttribue).TextString = Fil

    On Error GoTo 0
    FunEPISSURE = True
    Exit Function
Fin:

    If Fils = "FILG" Then
    If booG = True Then
        boolNotExecute = True
    Else
        boolNotExecute = False
     End If
        Fils = "FILD"
        I = 0
        Err.Clear
        GoTo Retour
    End If
    Err.Clear
End Function

 Function AtrbNumError() As Long
    Dim Sql As String
    Dim NErr As Long
    Dim RsNumError As Recordset
    Sql = "SELECT T_NumErreur.LibErreur, T_NumErreur.NumErreur "
    Sql = Sql & "FROM T_NumErreur "
    Sql = Sql & "WHERE T_NumErreur.LibErreur='ErrorApp';"
    Set RsNumError = Con.OpenRecordSet(Sql)
    If RsNumError.EOF = False Then
        Sql = "UPDATE T_NumErreur SET T_NumErreur.NumErreur = [NumErreur]+1;"
        Con.Execute Sql
        RsNumError.Requery
        AtrbNumError = RsNumError!NumErreur
    End If
End Function
Function VersionPices(Pieces As String) As Long
Dim Rs As Recordset
Dim Sql As String
Sql = "SELECT  VersionPices.Version FROM VersionPices "
Sql = Sql & "WHERE VersionPices.Pi='" & MyReplace(Pieces) & "';"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then
Sql = "INSERT INTO VersionPices ( Pi ) VALUES('" & MyReplace(Pieces) & "');"
Con.Execute Sql
End If
Sql = "UPDATE VersionPices SET VersionPices.Version = [Version] + 1 "
Sql = Sql & "WHERE VersionPices.Pi='" & MyReplace(Pieces) & "';"
Con.Execute Sql
Rs.Requery
VersionPices = Rs!Version
End Function
Function KilVersionXX(Version As String, Archive As String, Optional Kill As Boolean) As Boolean
Dim Fso As New FileSystemObject
MySeconde 5
Archive2 = Archive
Version2 = Version

Reprise:
SaveVersion2 = Version2
SaveArchive2 = Archive2
a = Split(Version2, "\")
Version2 = ""
For I = LBound(a) To UBound(a) - 1
Version2 = Version2 & a(I) & "\"
Next I
Version2 = Left(Version2, Len(Version2) - 1)



a = Split(Archive2, "\")
Archive2 = ""
For I = LBound(a) To UBound(a) - 1
Archive2 = Archive2 & a(I) & "\"
Next I
Archive2 = Left(Archive2, Len(Archive2) - 1)
Debug.Print Version2
Debug.Print Archive2
Debug.Print SaveVersion2
Debug.Print SaveArchive2
If Version2 <> Archive2 Then GoTo Reprise
If Kill = True Then

SetAttributs "" & SaveVersion2, False
    If SaveVersion2 <> SaveArchive2 Then On Error Resume Next: Fso.DeleteFolder SaveVersion2, True: Err.Clear: On Error GoTo 0
End If


KilVersionXX = True
End Function
Function funCloseConnextion()
Con.CloseConnection
'ConBaseNum.CloseConnection

End Function
Function ReseingeTor(CodeApp As String, InsertTorTitre) As Boolean
On Error GoTo Fin
   Dim PathTorDefault As String
 PathTorDefault = TableauPath.Item("PathTorDefault")
  PathTorDefault = DefinirChemienComplet(TableauPath.Item("PathServer"), PathTorDefault)
If TableuDeTor(CollectionTor(CodeApp)).Garder = False Then Exit Function
'If Left(PathTorDefault, 2) <> "\\" And Left(PathTorDefault, 1) = "\" Then PathTorDefault = TableauPath.Item("PathServer") & PathTorDefault
'
'If Right(PathTorDefault, 2) = "\\" Then PathTorDefault = Mid(PathTorDefault, 1, Len(PathTorDefault) - 1)

If TableuDeTor(CollectionTor(CodeApp)).TorExiste = False Then _
    TableuDeTor(CollectionTor(CodeApp)).InsertTorTitre(0) = Val(InsertTorTitre(0)): TableuDeTor(CollectionTor(CodeApp)).InsertTorTitre(1) = Val(InsertTorTitre(1)): TableuDeTor(CollectionTor(CodeApp)).InsertTorTitre(2) = Val(InsertTorTitre(2))
        Set TableuDeTor(CollectionTor(CodeApp)).NewBlockTorTire = FunInsBlock(PathTorDefault & "\TORDESIGNATION.dwg", TableuDeTor(CollectionTor(CodeApp)).InsertTorTitre, "", TableuDeTor(CollectionTor(CodeApp)).Rotation, TableuDeTor(CollectionTor(CodeApp)).XScaleFactor, TableuDeTor(CollectionTor(CodeApp)).YScaleFactor, TableuDeTor(I).ZScaleFactor)
            Set TableuDeTor(CollectionTor(CodeApp)).Attribues = ColectionAttribueConecteur(TableuDeTor(CollectionTor(CodeApp)).NewBlockTorTire.GetAttributes)
            a = TableuDeTor(CollectionTor(CodeApp)).NewBlockTorTire.GetAttributes
           a(TableuDeTor(CollectionTor(CodeApp)).Attribues("TORDESIGNATION")).TextString = TableuDeTor(CollectionTor(CodeApp)).CodeApp
           For I = 1 To UBound(TableuDeTor(CollectionTor(CodeApp)).Tor)
           
               If TableuDeTor(CollectionTor(CodeApp)).Tor(I).Garder = True Then
               If TableuDeTor(CollectionTor(CodeApp)).Tor(I).TorExiste = False Then
                    If I = 1 Then
                    
                        TableuDeTor(CollectionTor(CodeApp)).Tor(I).Insert(0) = TableuDeTor(CollectionTor(CodeApp)).InsertTorTitre(0)
                        TableuDeTor(CollectionTor(CodeApp)).Tor(I).Insert(1) = TableuDeTor(CollectionTor(CodeApp)).InsertTorTitre(1)
                        TableuDeTor(CollectionTor(CodeApp)).Tor(I).Insert(2) = TableuDeTor(CollectionTor(CodeApp)).InsertTorTitre(2)
                    Else
                        TableuDeTor(CollectionTor(CodeApp)).Tor(I).Insert(0) = TableuDeTor(CollectionTor(CodeApp)).Tor(I - I).Insert(0)
                        TableuDeTor(CollectionTor(CodeApp)).Tor(I).Insert(1) = TableuDeTor(CollectionTor(CodeApp)).Tor(I - I).Insert(1)
                        TableuDeTor(CollectionTor(CodeApp)).Tor(I).Insert(2) = TableuDeTor(CollectionTor(CodeApp)).Tor(I - I).Insert(2)
                    End If
                    TableuDeTor(CollectionTor(CodeApp)).Tor(I).Insert(1) = DecalInsertPointLigneTableau_fils_Bas(TableuDeTor(CollectionTor(CodeApp)).Tor(I).Insert(1), 5)
              End If
             Set TableuDeTor(CollectionTor(CodeApp)).Tor(I).NewBlockTorDetail = FunInsBlock(PathTorDefault & "\TORDETAIL.dwg", TableuDeTor(CollectionTor(CodeApp)).Tor(I).Insert, "", TableuDeTor(CollectionTor(CodeApp)).Tor(I).Rotation, TableuDeTor(I).Tor(I).XScaleFactor, TableuDeTor(CollectionTor(CodeApp)).YScaleFactor, TableuDeTor(CollectionTor(CodeApp)).Tor(I).ZScaleFactor)
            Set TableuDeTor(CollectionTor(CodeApp)).Attribues = ColectionAttribueConecteur(TableuDeTor(CollectionTor(CodeApp)).Tor(I).NewBlockTorDetail.GetAttributes)
            a = TableuDeTor(CollectionTor(CodeApp)).Tor(I).NewBlockTorDetail.GetAttributes
             a(TableuDeTor(CollectionTor(CodeApp)).Attribues("TORDESIGNATION")).TextString = "" & TableuDeTor(CollectionTor(CodeApp)).CodeApp
            a(TableuDeTor(CollectionTor(CodeApp)).Attribues("TORFILS")).TextString = "" & TableuDeTor(CollectionTor(CodeApp)).Tor(I).TableauFile
            a(TableuDeTor(CollectionTor(CodeApp)).Attribues("TORNUM")).TextString = "" & TableuDeTor(CollectionTor(CodeApp)).Tor(I).TorName
           End If
        Next I
Fin:
Err.Clear
End Function
Function funOpenDatabase()
Con.TYPEBASE = ADO_TYPEBASE
Con.BASE = ADO_BASE
Con.SERVER = ADO_SERVER
Con.Fichier = ADO_Fichier
Con.User = ADO_User
Con.PassWord = ADO_PassWord
If Trim("" & Con.BASE) = "" Then Con.BASE = Db
Con.OpenConnetion
'ConBaseNum.OpenConnetion DbNumPlan
End Function
'Global Function ScanServeur()
'Dim Fso As New FileSystemObject
'Dim PathFavorisreseau As String
'Dim FileNumber As Long
'
'Dim Txt
'txt2 = ""
'i = 0
'FileNumber = FreeFile
'PathFavorisreseau = "c:\Favoris_reseau.Bat"
'Open PathFavorisreseau For Output As FileNumber   ' Ouvre le fichier en lecture.
'    Print #FileNumber, "net View > %1"
'    Print #FileNumber, "cls"
'    Print #FileNumber, "Dir vue.Txt"
'    Print #FileNumber, "pause"
'    Close #FileNumber
'Close #FileNumber
'PathFavorisreseau2 = Environ("USERPROFILE") & "\vue.txt"
'
'Shell PathFavorisreseau & " " & PathFavorisreseau2
'Open PathFavorisreseau For Input As FileNumber   ' Ouvre le fichier en lecture.
'Do While Not EOF(FileNumber)
'i = i + 1 ' Effectue la boucle jusqu'à la fin du fichier.
'    Input #FileNumber, Txt    ' Lit les données dans deux variables.
'    Debug.Print Txt
'    If i > 3 Then
'    If Trim(Txt) = "La commande s'est termin‚e correctement." Then Exit Do
'    txt2 = Mid(Txt, 1, Len(Txt) - InStr(1, Txt, " ")) & vbCrLf
'    Debug.Print txt2
'    End If
'    ' Affiche les données dans la fenêtre Exécution.
'  Loop
'Close #FileNumber   ' Ferme le fichier.
'ScanServeur = Split(txt2, vbCrLf)
'Set Fso = Nothing
'End Function

Function ScanServeur()
Dim Fso As New FileSystemObject
Dim PathFavorisreseau As String
Dim FileNumber As Long
Dim Txt As String
Dim txt2 As String

txt2 = ""
I = 0
FileNumber = FreeFile
PathFavorisreseau = Environ("USERPROFILE") & "\Favoris_reseau"
If Fso.FolderExists(PathFavorisreseau) = False Then
    Fso.CreateFolder PathFavorisreseau
End If
Open PathFavorisreseau & "\Favoris_reseau.BAT" For Output As FileNumber   ' Ouvre le fichier en lecture.
Print #FileNumber, "net View>" & Chr(34) & PathFavorisreseau & "\Favoris_reseau.Txt" & Chr(34)
'Print #FileNumber, "dir " & PathFavorisreseau & "*.*"
'Print #FileNumber, "PAUSE"
Close #FileNumber

Shell Chr(34) & PathFavorisreseau & "\Favoris_reseau.BAT" & Chr(34)
PathFavorisreseau = PathFavorisreseau & "\Favoris_reseau.Txt"

Open PathFavorisreseau For Input As FileNumber   ' Ouvre le fichier en lecture.
Do While Not EOF(FileNumber)
    Input #FileNumber, Txt    ' Lit les données dans deux variables.
    Debug.Print Txt
    If InStr(1, Txt, Chr(32)) = 0 Then
         txt2 = txt2 & Txt & vbCrLf
    Else
        txt2 = txt2 & Txt 'Left(Txt, InStr(1, Txt, Chr(32)) - 1) & vbCrLf
    End If
    ' Affiche les données dans la fenêtre Exécution.
  Loop
Close #FileNumber   ' Ferme le fichier.
ScanServeur = Split(txt2, vbCrLf)
Set Fso = Nothing
End Function

Sub MajBase(IdIndice As Long)
Dim Sql As String
'***********************************************************************************************************************
'*                UPDATE Connecteurs SET Connecteurs.Supprimer = True
'*                  Supprime les données des tables de travail :                                                       *
Sql = "UPDATE T_Critères SET T_Critères.Supprimer = True    "
'Sql = Sql & "FROM T_Critères "
Sql = Sql & "WHERE T_Critères.Id_IndiceProjet=" & IdIndice & ";"
Con.Execute Sql

Sql = "UPDATE Ligne_Tableau_fils SET Ligne_Tableau_fils.Supprimer = True "
'Sql = Sql & "FROM Ligne_Tableau_fils "
Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndice & ";"
Con.Execute Sql

Sql = "UPDATE Connecteurs SET Connecteurs.Supprimer = True  "
'Sql = Sql & "FROM Connecteurs "
Sql = Sql & "WHERE Connecteurs.Id_IndiceProjet=" & IdIndice & ";"
Con.Execute Sql

Sql = "UPDATE Composants SET Composants.Supprimer = True  "
'Sql = Sql & "FROM Composants "
Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & IdIndice & ";"
Con.Execute Sql

Sql = "UPDATE Nota SET Nota.Supprimer = True "
'Sql = Sql & "FROM Nota "
Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & IdIndice & ";"
Con.Execute Sql


Sql = "UPDATE T_Noeuds SET T_Noeuds.Supprimer = True "
'Sql = Sql & "FROM T_Noeuds "
Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & IdIndice & ";"
Con.Execute Sql
'***********************************************************************************************************************
'*                                        Enrichie les données des tables de travail :                                 *

Sql = "UPDATE T_Critères INNER JOIN Xls_Critères ON T_Critères.Id = Xls_Critères.ID SET T_Critères.ACTIVER =  "
Sql = Sql & "[Xls_Critères].[ACTIVER], T_Critères.CODE_CRITERE = [Xls_Critères].[CODE_CRITERE], T_Critères.CRITERES =  "
Sql = Sql & "[Xls_Critères].[CRITERES], T_Critères.DESIGNATION = [Xls_Critères].[DESIGNATION], T_Critères.Commentaires =  "
Sql = Sql & "[Xls_Critères].[Commentaires], T_Critères.Supprimer = False "
Sql = Sql & "WHERE Xls_Critères.Job=" & NmJob & " AND T_Critères.Id_IndiceProjet=" & IdIndice & " ;"
Con.Execute Sql

'
Sql = "INSERT INTO T_Critères ( Id_IndiceProjet,ACTIVER,CODE_CRITERE, CRITERES,COMMENTAIRES)  "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet,Xls_Critères.ACTIVER, Xls_Critères.CODE_CRITERE, Xls_Critères.CRITERES ,Xls_Critères.COMMENTAIRES "
Sql = Sql & "FROM Xls_Critères LEFT JOIN T_Critères ON Xls_Critères.ID = T_Critères.Id "
Sql = Sql & "WHERE Xls_Critères.Job=" & NmJob & " AND T_Critères.Id Is Null;"






Con.Execute Sql


'
Sql = "DELETE FROM T_Critères WHERE T_Critères.Id_IndiceProjet=" & IdIndice & " AND T_Critères.Supprimer=True;"
Con.Execute Sql






Sql = "UPDATE xls_Ligne_Tableau_fils INNER JOIN Ligne_Tableau_fils ON xls_Ligne_Tableau_fils.ID =   "
Sql = Sql & "Ligne_Tableau_fils.Id SET Ligne_Tableau_fils.ACTIVER = [xls_Ligne_Tableau_fils].[ACTIVER], Ligne_Tableau_fils.LIAI =   "
Sql = Sql & "[xls_Ligne_Tableau_fils].[LIAI], Ligne_Tableau_fils.DESIGNATION = [xls_Ligne_Tableau_fils].[DESIGNATION], Ligne_Tableau_fils.FIL =   "
Sql = Sql & "[xls_Ligne_Tableau_fils].[FIL], Ligne_Tableau_fils.SECT = [xls_Ligne_Tableau_fils].[SECT], Ligne_Tableau_fils.TEINT =   "
Sql = Sql & "[xls_Ligne_Tableau_fils].[TEINT], Ligne_Tableau_fils.TEINT2 = [xls_Ligne_Tableau_fils].[TEINT2], Ligne_Tableau_fils.ISO =   "
Sql = Sql & "[xls_Ligne_Tableau_fils].[ISO], Ligne_Tableau_fils.[LONG] = [xls_Ligne_Tableau_fils].[LONG], Ligne_Tableau_fils.[LONG CP] =   "
Sql = Sql & "[xls_Ligne_Tableau_fils].[LONG CP], Ligne_Tableau_fils.Long_Add = [xls_Ligne_Tableau_fils].[Long_Add], Ligne_Tableau_fils.Long_Add2 =   "
Sql = Sql & "[xls_Ligne_Tableau_fils].[Long_Add2], Ligne_Tableau_fils.COUPE = [xls_Ligne_Tableau_fils].[COUPE], Ligne_Tableau_fils.POS =   "
Sql = Sql & "[xls_Ligne_Tableau_fils].[POS], Ligne_Tableau_fils.[POS-OUT] = [xls_Ligne_Tableau_fils].[POS-OUT] , Ligne_Tableau_fils.FA = "
Sql = Sql & "[xls_Ligne_Tableau_fils].[FA], Ligne_Tableau_fils.APP = [xls_Ligne_Tableau_fils].[App], Ligne_Tableau_fils.VOI =   "
Sql = Sql & "[xls_Ligne_Tableau_fils].[VOI], Ligne_Tableau_fils.[Ref Connecteur] = [xls_Ligne_Tableau_fils].[Ref Connecteur],   "
Sql = Sql & "Ligne_Tableau_fils.[Ref Connecteur_Four] = [xls_Ligne_Tableau_fils].[Ref Connecteur_Four], Ligne_Tableau_fils.[Ref Clip] =   "
Sql = Sql & "[xls_Ligne_Tableau_fils].[Ref Clip], Ligne_Tableau_fils.[Ref Clip Four] = [xls_Ligne_Tableau_fils].[Ref Clip Four],   "
Sql = Sql & "Ligne_Tableau_fils.PRECO = [xls_Ligne_Tableau_fils].[PRECO], Ligne_Tableau_fils.[Ref Joint] = [xls_Ligne_Tableau_fils].[Ref Joint],   "
Sql = Sql & "Ligne_Tableau_fils.[Ref Joint four] = [xls_Ligne_Tableau_fils].[Ref Joint four], Ligne_Tableau_fils.POS2 =   "
Sql = Sql & "[xls_Ligne_Tableau_fils].[POS2], Ligne_Tableau_fils.[POS-OUT2] = [xls_Ligne_Tableau_fils].[POS-OUT2], Ligne_Tableau_fils.FA2 =   "
Sql = Sql & "[xls_Ligne_Tableau_fils].[FA2], Ligne_Tableau_fils.APP2 = [xls_Ligne_Tableau_fils].[APP2 ], Ligne_Tableau_fils.VOI2 = [xls_Ligne_Tableau_fils].[VOI2],   "
Sql = Sql & "Ligne_Tableau_fils.[Ref Connecteur2] = [xls_Ligne_Tableau_fils].[Ref Connecteur2], Ligne_Tableau_fils.[Ref Connecteur_Four2] =   "
Sql = Sql & "[xls_Ligne_Tableau_fils].[Ref Connecteur_Four2], Ligne_Tableau_fils.[Ref Clip2] = [xls_Ligne_Tableau_fils].[Ref Clip2],   "
Sql = Sql & "Ligne_Tableau_fils.[Ref Clip Four2] = [xls_Ligne_Tableau_fils].[Ref Clip Four2], Ligne_Tableau_fils.PRECO2 =   "
Sql = Sql & "[xls_Ligne_Tableau_fils].[PRECO2], Ligne_Tableau_fils.[Ref Joint2] = [xls_Ligne_Tableau_fils].[Ref Joint2],   "
Sql = Sql & "Ligne_Tableau_fils.[Ref Joint Four2] = [xls_Ligne_Tableau_fils].[Ref Joint Four2], Ligne_Tableau_fils.PRECOG =   "
Sql = Sql & "[xls_Ligne_Tableau_fils].[PRECOG], Ligne_Tableau_fils.[OPTION] = [xls_Ligne_Tableau_fils].[OPTION],   "
Sql = Sql & "Ligne_Tableau_fils.[Critères spécifiques] = [xls_Ligne_Tableau_fils].[Critères spécifiques], Ligne_Tableau_fils.Commentaires =   "
Sql = Sql & "[xls_Ligne_Tableau_fils].[Commentaires], Ligne_Tableau_fils.Supprimer = False "
Sql = Sql & "WHERE xls_Ligne_Tableau_fils.Job=" & NmJob & " AND Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndice & " ;"


Con.Execute Sql

Sql = "INSERT INTO Ligne_Tableau_fils ( Id_IndiceProjet, LIAI, DESIGNATION, FIL, SECT, TEINT, TEINT2, ISO, [LONG], [LONG CP],  "
Sql = Sql & "COUPE, POS, [POS-OUT], FA, APP, VOI, [Ref Connecteur], [Ref Connecteur_Four], Long_Add, [Ref Clip], [Ref Clip Four],  "
Sql = Sql & "[Ref Joint], [Ref Joint Four], POS2, [POS-OUT2], FA2, APP2, VOI2, [Ref Connecteur2], [Ref Connecteur_Four2], Long_Add2,  "
Sql = Sql & "[Ref Clip2], [Ref Clip Four2], [Ref Joint2], [Ref Joint Four2], PRECOG, [OPTION], ACTIVER, [Critères spécifiques],  "
Sql = Sql & "PRECO, PRECO2, COMMENTAIRES ) "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, xls_Ligne_Tableau_fils.LIAI, xls_Ligne_Tableau_fils.DESIGNATION, xls_Ligne_Tableau_fils.FIL,  "
Sql = Sql & "xls_Ligne_Tableau_fils.SECT, xls_Ligne_Tableau_fils.TEINT, xls_Ligne_Tableau_fils.TEINT2, xls_Ligne_Tableau_fils.ISO,  "
Sql = Sql & "xls_Ligne_Tableau_fils.LONG, xls_Ligne_Tableau_fils.[LONG CP], xls_Ligne_Tableau_fils.COUPE, xls_Ligne_Tableau_fils.POS,  "
Sql = Sql & "xls_Ligne_Tableau_fils.[POS-OUT], xls_Ligne_Tableau_fils.FA, xls_Ligne_Tableau_fils.APP, xls_Ligne_Tableau_fils.VOI,  "
Sql = Sql & "xls_Ligne_Tableau_fils.[Ref Connecteur], xls_Ligne_Tableau_fils.[Ref Connecteur_Four], xls_Ligne_Tableau_fils.Long_Add,  "
Sql = Sql & "xls_Ligne_Tableau_fils.[Ref Clip], xls_Ligne_Tableau_fils.[Ref Clip Four], xls_Ligne_Tableau_fils.[Ref Joint],  "
Sql = Sql & "xls_Ligne_Tableau_fils.[Ref Joint four], xls_Ligne_Tableau_fils.POS2, xls_Ligne_Tableau_fils.[POS-OUT2], xls_Ligne_Tableau_fils.FA2,  "
Sql = Sql & "xls_Ligne_Tableau_fils.APP2, xls_Ligne_Tableau_fils.VOI2, xls_Ligne_Tableau_fils.[Ref Connecteur2],  "
Sql = Sql & "xls_Ligne_Tableau_fils.[Ref Connecteur_Four2], xls_Ligne_Tableau_fils.Long_Add2, xls_Ligne_Tableau_fils.[Ref Clip2],  "
Sql = Sql & "xls_Ligne_Tableau_fils.[Ref Clip Four2], xls_Ligne_Tableau_fils.[Ref Joint2], xls_Ligne_Tableau_fils.[Ref Joint four2],  "
Sql = Sql & "xls_Ligne_Tableau_fils.PRECOG, xls_Ligne_Tableau_fils.OPTION, xls_Ligne_Tableau_fils.ACTIVER,  "
Sql = Sql & "xls_Ligne_Tableau_fils.[Critères spécifiques], xls_Ligne_Tableau_fils.PRECO, xls_Ligne_Tableau_fils.PRECO2,  "
Sql = Sql & "xls_Ligne_Tableau_fils.Commentaires "
Sql = Sql & "FROM xls_Ligne_Tableau_fils LEFT JOIN Ligne_Tableau_fils ON xls_Ligne_Tableau_fils.ID = Ligne_Tableau_fils.Id "
Sql = Sql & "WHERE xls_Ligne_Tableau_fils.Job=" & NmJob & "  "
Sql = Sql & "AND Ligne_Tableau_fils.Id Is Null;"




Con.Execute Sql

'
Sql = "DELETE FROM Ligne_Tableau_fils WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndice & " AND Ligne_Tableau_fils.Supprimer=True;"
Con.Execute Sql



'
'Sql = "INSERT INTO Connecteurs ( Id_IndiceProjet, ACTIVER, CONNECTEUR, RefConnecteurFour, [O/N], DESIGNATION, CODE_APP, N°, "
'Sql = Sql & "POS, [POS-OUT], PRECO1, PRECO2, [100%], [OPTION], Pylone, Colonne, Ligne, RefBouchon, RefBouchonFour, ReFCapot, "
'Sql = Sql & "ReFCapotFour, RefVerrou, RefVerrouFour, LongueurF_Choix,COMMENTAIRES  )"
'Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, Xls_Connecteurs.ACTIVER, Xls_Connecteurs.CONNECTEUR, "
'Sql = Sql & "Xls_Connecteurs.RefConnecteurFour, Xls_Connecteurs.[O/N], Xls_Connecteurs.DESIGNATION, "
'Sql = Sql & "Xls_Connecteurs.CODE_APP, Xls_Connecteurs.N°, Xls_Connecteurs.POS, Xls_Connecteurs.[POS-OUT], "
'Sql = Sql & "Xls_Connecteurs.PRECO1,Xls_Connecteurs.PRECO2, Xls_Connecteurs.[100%], Xls_Connecteurs.OPTION, Xls_Connecteurs.Pylone, "
'Sql = Sql & "Xls_Connecteurs.Colonne, Xls_Connecteurs.Ligne, Xls_Connecteurs.RefBouchon, Xls_Connecteurs.RefBouchonFour, "
'Sql = Sql & "Xls_Connecteurs.ReFCapot, Xls_Connecteurs.ReFCapotFour, Xls_Connecteurs.RefVerrou, Xls_Connecteurs.RefVerrouFour, "
'Sql = Sql & "Xls_Connecteurs.LongueurF_Choix,Xls_Connecteurs.COMMENTAIRES  "
'Sql = Sql & "FROM Xls_Connecteurs "
'Sql = Sql & "WHERE Xls_Connecteurs.Job=" & NmJob & " AND Connecteurs.Id_IndiceProjet=" & IdIndice & ";"





Sql = "UPDATE Xls_Connecteurs INNER JOIN Connecteurs ON Xls_Connecteurs.ID = Connecteurs.Id SET Connecteurs.ACTIVER = [Xls_Connecteurs].[ACTIVER],  "
Sql = Sql & "Connecteurs.CONNECTEUR = [Xls_Connecteurs].[CONNECTEUR], Connecteurs.RefConnecteurFour = [Xls_Connecteurs].[RefConnecteurFour],  "
Sql = Sql & "Connecteurs.[O/N] = [Xls_Connecteurs].[O/N], Connecteurs.DESIGNATION = [Xls_Connecteurs].[DESIGNATION], Connecteurs.CODE_APP =  "
Sql = Sql & "[Xls_Connecteurs].[CODE_APP], Connecteurs.N° = [Xls_Connecteurs].[N°], Connecteurs.POS = [Xls_Connecteurs].[POS], Connecteurs.[POS-OUT] =  "
Sql = Sql & "[Xls_Connecteurs].[POS-OUT], Connecteurs.PRECO1 = [Xls_Connecteurs].[PRECO1], Connecteurs.PRECO2 = [Xls_Connecteurs].[PRECO2],  "
Sql = Sql & "Connecteurs.[OPTION] = [Xls_Connecteurs].[OPTION], Connecteurs.[100%] = [Xls_Connecteurs].[100%], Connecteurs.Pylone =  "
Sql = Sql & "[Xls_Connecteurs].[Pylone], Connecteurs.Colonne = [Xls_Connecteurs].[Colonne], Connecteurs.Ligne = [Xls_Connecteurs].[Ligne],  "
Sql = Sql & "Connecteurs.RefBouchon = [Xls_Connecteurs].[RefBouchon], Connecteurs.RefBouchonFour = [Xls_Connecteurs].[RefBouchonFour],  "
Sql = Sql & "Connecteurs.ReFCapot = [Xls_Connecteurs].[ReFCapot], Connecteurs.ReFCapotFour = [Xls_Connecteurs].[ReFCapotFour],  "
Sql = Sql & "Connecteurs.RefVerrou = [Xls_Connecteurs].[RefVerrou], Connecteurs.RefVerrouFour = [Xls_Connecteurs].[RefVerrouFour],  "
Sql = Sql & "Connecteurs.LongueurF_Choix = [Xls_Connecteurs].[LongueurF_Choix], Connecteurs.Commentaires = [Xls_Connecteurs].[Commentaires],  "
Sql = Sql & "Connecteurs.Supprimer = False "
Sql = Sql & "WHERE Xls_Connecteurs.Job=" & NmJob & " AND Connecteurs.Id_IndiceProjet=" & IdIndice & ";"


Con.Execute Sql





Sql = "INSERT INTO Connecteurs ( Id_IndiceProjet, ACTIVER, CONNECTEUR, RefConnecteurFour, [O/N], DESIGNATION, CODE_APP, N°, POS,  "
Sql = Sql & "[POS-OUT], PRECO1, PRECO2, [100%], [OPTION], Pylone, Colonne, Ligne, RefBouchon, RefBouchonFour, ReFCapot,  "
Sql = Sql & "ReFCapotFour, RefVerrou, RefVerrouFour, LongueurF_Choix, COMMENTAIRES ) "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, Xls_Connecteurs.ACTIVER, Xls_Connecteurs.CONNECTEUR, Xls_Connecteurs.RefConnecteurFour,  "
Sql = Sql & "Xls_Connecteurs.[O/N], Xls_Connecteurs.DESIGNATION, Xls_Connecteurs.CODE_APP, Xls_Connecteurs.N°,  "
Sql = Sql & "Xls_Connecteurs.POS, Xls_Connecteurs.[POS-OUT], Xls_Connecteurs.PRECO1, Xls_Connecteurs.PRECO2, Xls_Connecteurs.[100%],  "
Sql = Sql & "Xls_Connecteurs.OPTION, Xls_Connecteurs.Pylone, Xls_Connecteurs.Colonne, Xls_Connecteurs.Ligne, Xls_Connecteurs.RefBouchon,  "
Sql = Sql & "Xls_Connecteurs.RefBouchonFour, Xls_Connecteurs.ReFCapot, Xls_Connecteurs.ReFCapotFour, Xls_Connecteurs.RefVerrou,  "
Sql = Sql & "Xls_Connecteurs.RefVerrouFour, Xls_Connecteurs.LongueurF_Choix, Xls_Connecteurs.Commentaires "
Sql = Sql & "FROM Xls_Connecteurs LEFT JOIN Connecteurs ON Xls_Connecteurs.ID = Connecteurs.Id "
Sql = Sql & "WHERE Xls_Connecteurs.Job=" & NmJob & "  "
Sql = Sql & "AND Connecteurs.Id Is Null;"





Con.Execute Sql

Sql = "Delete FROM Connecteurs WHERE Connecteurs.Id_IndiceProjet=" & IdIndice & " And Connecteurs.Supprimer=true;"
Con.Execute Sql
Sql = "INSERT INTO Composants ( Id_IndiceProjet, ACTIVER,DESIGNCOMP, NUMCOMP, REFCOMP, Path  ,[OPTION],Code_APP_Lier,Voie,POS,[POS-OUT] ,COMMENTAIRES ) "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, Xls_Composants.ACTIVER,Xls_Composants.DESIGNCOMP, Xls_Composants.NUMCOMP,   "
Sql = Sql & "Xls_Composants.REFCOMP, Xls_Composants.Path  , Xls_Composants.[OPTION] ,Xls_Composants.Code_APP_Lier,Xls_Composants.Voie,Xls_Composants.POS,Xls_Composants.[POS-OUT] ,Xls_Composants.COMMENTAIRES  "
Sql = Sql & "FROM Xls_Composants "
Sql = Sql & "WHERE Xls_Composants.Job=" & NmJob & ";"

Sql = "UPDATE Xls_Composants INNER JOIN Composants ON Xls_Composants.ID = Composants.Id SET Composants.ACTIVER =  "
Sql = Sql & "[Xls_Composants].[ACTIVER], Composants.DESIGNCOMP = [Xls_Composants].[DESIGNCOMP], Composants.NUMCOMP = [Xls_Composants].[NUMCOMP],  "
Sql = Sql & "Composants.REFCOMP = [Xls_Composants].[REFCOMP], Composants.Path = [Xls_Composants].[Path], Composants.[OPTION] =  "
Sql = Sql & "[Xls_Composants].[OPTION], Composants.Code_APP_Lier = [Xls_Composants].[Code_APP_Lier], Composants.Voie = [Xls_Composants].[Voie],  "
Sql = Sql & "Composants.POS = [Xls_Composants].[POS], Composants.[POS-OUT] = [Xls_Composants].[POS-OUT], Composants.Commentaires =  "
Sql = Sql & "[Xls_Composants].[Commentaires], Composants.Supprimer = False "
Sql = Sql & "WHERE Composants.Id_IndiceProjet=" & IdIndice & "  "
Sql = Sql & "AND Xls_Composants.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "INSERT INTO Composants ( Id_IndiceProjet, ACTIVER, DESIGNCOMP, NUMCOMP, REFCOMP, Path, [OPTION], Code_APP_Lier, Voie, POS, [POS-OUT],  "
Sql = Sql & "COMMENTAIRES ) "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, Xls_Composants.ACTIVER, Xls_Composants.DESIGNCOMP, Xls_Composants.NUMCOMP, Xls_Composants.REFCOMP,  "
Sql = Sql & "Xls_Composants.Path, Xls_Composants.OPTION, Xls_Composants.Code_APP_Lier, Xls_Composants.Voie, Xls_Composants.POS,  "
Sql = Sql & "Xls_Composants.[POS-OUT], Xls_Composants.Commentaires "
Sql = Sql & "FROM Xls_Composants LEFT JOIN Composants ON Xls_Composants.ID = Composants.Id "
Sql = Sql & "WHERE Composants.Id Is Null "
Sql = Sql & " AND Xls_Composants.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "Delete FROM Composants WHERE Composants.Id_IndiceProjet=" & IdIndice & " And Composants.Supprimer=true;"
Con.Execute Sql

Sql = "INSERT INTO Nota ( Id_IndiceProjet,ACTIVER, NOTA, NUMNOTA,[OPTION],COMMENTAIRES  ) "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, Xls_Nota.ACTIVER,Xls_Nota.NOTA, Xls_Nota.NUMNOTA,Xls_Nota.[OPTION] ,Xls_Nota.COMMENTAIRES "
Sql = Sql & "FROM Xls_Nota "
Sql = Sql & "WHERE Xls_Nota.Job=" & NmJob & ";"

Sql = "UPDATE Xls_Nota INNER JOIN Nota ON Xls_Nota.ID = Nota.Id SET Nota.ACTIVER = [Xls_Nota].[ACTIVER], Nota.NOTA = [Xls_Nota].[NOTA],  "
Sql = Sql & "Nota.NUMNOTA = [Xls_Nota].[NUMNOTA], Nota.[OPTION] = [Xls_Nota].[OPTION], Nota.Commentaires = [Xls_Nota].[Commentaires],  Nota.Supprimer = False "
Sql = Sql & "WHERE Nota.Id_IndiceProjet=" & IdIndice & "  "
Sql = Sql & "AND Xls_Nota.Job=" & NmJob & ";"

Con.Execute Sql

Sql = "INSERT INTO Nota ( Id_IndiceProjet, ACTIVER, NOTA, NUMNOTA, [OPTION], COMMENTAIRES )  "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, Xls_Nota.ACTIVER, Xls_Nota.NOTA, Xls_Nota.NUMNOTA, Xls_Nota.OPTION, Xls_Nota.COMMENTAIRES "
Sql = Sql & "FROM Xls_Nota LEFT JOIN Nota ON Xls_Nota.ID = Nota.Id  "
Sql = Sql & "WHERE Xls_Nota.Job=" & NmJob & "   "
Sql = Sql & "AND Nota.Id Is Null;"
Con.Execute Sql

Sql = "Delete FROM Nota WHERE Nota.Id_IndiceProjet=" & IdIndice & " And Nota.Supprimer=true;"
Con.Execute Sql


Sql = "INSERT INTO T_Noeuds ( Id_IndiceProjet,Fleche_Droite, ACTIVER, NŒUDS,LONGUEUR,DESIGN_HAB, "
Sql = Sql & "CODE_RSA,CODE_PSA,CODE_ENC,DIAMETRE,CLASSE_T,TORON_PRINCIPAL, LONGUEUR_CUMULEE,[OPTION],COMMENTAIRES) "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, Xls_Noeuds.Fleche_Droite, Xls_Noeuds.ACTIVER, "
Sql = Sql & "Xls_Noeuds.NŒUDS,Xls_Noeuds.LONGUEUR,Xls_Noeuds.DESIGN_HAB,Xls_Noeuds.CODE_RSA, "
Sql = Sql & "Xls_Noeuds.CODE_PSA,Xls_Noeuds.CODE_ENC,Xls_Noeuds.DIAMETRE,Xls_Noeuds.CLASSE_T,Xls_Noeuds.TORON_PRINCIPAL, "
Sql = Sql & "Xls_Noeuds.LONGUEUR_CUMULEE ,Xls_Noeuds.[OPTION],Xls_Noeuds.COMMENTAIRES "
Sql = Sql & "FROM Xls_Noeuds "
Sql = Sql & "WHERE Xls_Noeuds.Job=" & NmJob & ";"

Sql = "UPDATE Xls_Noeuds INNER JOIN T_Noeuds ON Xls_Noeuds.ID = T_Noeuds.Id  SET T_Noeuds.Fleche_Droite = [Xls_Noeuds].[Fleche_Droite],  "
Sql = Sql & "T_Noeuds.ACTIVER = [Xls_Noeuds].[ACTIVER], T_Noeuds.NŒUDS = [Xls_Noeuds].[NŒUDS], T_Noeuds.LONGUEUR = [Xls_Noeuds].[LONGUEUR],  "
Sql = Sql & "T_Noeuds.DESIGN_HAB = [Xls_Noeuds].[DESIGN_HAB], T_Noeuds.CODE_RSA = [Xls_Noeuds].[CODE_RSA], T_Noeuds.CODE_PSA = [Xls_Noeuds].[CODE_PSA],  "
Sql = Sql & "T_Noeuds.CODE_ENC = [Xls_Noeuds].[CODE_ENC], T_Noeuds.DIAMETRE = [Xls_Noeuds].[DIAMETRE], T_Noeuds.CLASSE_T = [Xls_Noeuds].[CLASSE_T],  "
Sql = Sql & "T_Noeuds.TORON_PRINCIPAL = [Xls_Noeuds].[TORON_PRINCIPAL], T_Noeuds.LONGUEUR_CUMULEE = [Xls_Noeuds].[LONGUEUR_CUMULEE],  "
Sql = Sql & "T_Noeuds.[OPTION] = [Xls_Noeuds].[OPTION], T_Noeuds.Commentaires = [Xls_Noeuds].[Commentaires], T_Noeuds.Supprimer = False "
Sql = Sql & "WHERE T_Noeuds.Id_IndiceProjet=" & IdIndice & " "
Sql = Sql & "AND Xls_Noeuds.Job=" & NmJob & ";"



Con.Execute Sql
Sql = "INSERT INTO T_Noeuds ( Id_IndiceProjet, Fleche_Droite, ACTIVER, NŒUDS, LONGUEUR, DESIGN_HAB, CODE_RSA, CODE_PSA, CODE_ENC, DIAMETRE,  "
Sql = Sql & "CLASSE_T, TORON_PRINCIPAL, LONGUEUR_CUMULEE, [OPTION], COMMENTAIRES ) "
Sql = Sql & "SELECT " & IdIndice & " AS Id_IndiceProjet, Xls_Noeuds.Fleche_Droite, Xls_Noeuds.ACTIVER, Xls_Noeuds.NŒUDS, Xls_Noeuds.LONGUEUR,  "
Sql = Sql & "Xls_Noeuds.DESIGN_HAB, Xls_Noeuds.CODE_RSA, Xls_Noeuds.CODE_PSA, Xls_Noeuds.CODE_ENC, Xls_Noeuds.DIAMETRE, Xls_Noeuds.CLASSE_T,  "
Sql = Sql & "Xls_Noeuds.TORON_PRINCIPAL, Xls_Noeuds.LONGUEUR_CUMULEE, Xls_Noeuds.OPTION, Xls_Noeuds.Commentaires "
Sql = Sql & "FROM Xls_Noeuds LEFT JOIN T_Noeuds ON Xls_Noeuds.ID = T_Noeuds.Id "
Sql = Sql & "WHERE Xls_Noeuds.Job=" & NmJob & " "
Sql = Sql & "AND T_Noeuds.Id Is Null;"
Con.Execute Sql


Sql = "Delete FROM T_Noeuds WHERE T_Noeuds.Id_IndiceProjet=" & IdIndice & " And T_Noeuds.Supprimer=true;"
Con.Execute Sql


Sql = "DELETE Xls_Critères.*  FROM Xls_Critères "
Sql = Sql & "where Xls_Critères.Job=" & NmJob & ";"
Con.Execute Sql


 Sql = "DELETE xls_Ligne_Tableau_fils.*  FROM xls_Ligne_Tableau_fils "
Sql = Sql & "where xls_Ligne_Tableau_fils.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "DELETE Xls_Composants.*  FROM Xls_Composants "
Sql = Sql & "where Xls_Composants.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "DELETE Xls_Nota.*  FROM Xls_Nota "
Sql = Sql & "where Xls_Nota.Job=" & NmJob & ";"
Con.Execute Sql

Sql = "DELETE Xls_Connecteurs.* FROM Xls_Connecteurs "
Sql = Sql & "WHERE Xls_Connecteurs.Job=" & NmJob & ";"
Con.Execute Sql

'***********************************************************************************************************************
'*                                        Attribut les code appareil au tableau de fils :                              *
'
'sql = "UPDATE (Ligne_Tableau_fils LEFT JOIN Connecteurs ON Ligne_Tableau_fils.FA = Connecteurs.N°)  "
'sql = sql & "LEFT JOIN Connecteurs AS Connecteurs_1 ON Ligne_Tableau_fils.FA2 = Connecteurs_1.N°  "
'sql = sql & "SET Ligne_Tableau_fils.APP = [Connecteurs].[CODE_APP], Ligne_Tableau_fils.APP2 = [Connecteurs_1].[CODE_APP] "
'sql = sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet=" & IdIndice & " "
'sql = sql & "AND Connecteurs.Id_IndiceProjet=" & IdIndice & " "
'sql = sql & "AND Connecteurs_1.Id_IndiceProjet=" & IdIndice & ";"
'Con.Execute sql
'***********************************************************************************************************************

End Sub
Sub InsertColonneApres(MyWorksheet As Worksheet, MyChamp As String, MyNewChamp As String)
Dim MyRange As Range
Dim MyAddres As String
MyAddres = ""
Set MyRange = MyWorksheet.Range("A1").CurrentRegion
For I = 1 To MyRange.Columns.Count
    If MyRange(1, I) = MyChamp Then
        MyAddres = MyRange(1, I + 1).Address
        
        Exit For
    End If
Next
If Trim(MyAddres) <> "" Then
    IsertColonne MyNewChamp, MyWorksheet, , "" & MyAddres
End If
End Sub



Function AtocatOption(Id_Pieces As Long) As Boolean
If (bool_Plan_E_Options = False And Mytype = "PL") Or (bool_Outil_E_Options = False And Mytype = "OU") Then Exit Function

Dim Rs As Recordset
Dim RsSelect As Recordset
Dim Sql As String
Dim Index As Long
Dim MyPotionEntete As Collection
Dim Block As Object
'AutoCAD.Visible = True
AtocatOption = False
Index = 0
Sql = "SELECT T_indiceProjet.* "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "Where T_indiceProjet.Pere = " & Id_Pieces & " "
Sql = Sql & "ORDER BY T_indiceProjet.Id DESC;"
Set Rs = Con.OpenRecordSet(Sql)
If Rs.EOF = True Then Exit Function
Sql = "SELECT T_indiceProjet.* "
Sql = Sql & "FROM T_indiceProjet "
Sql = Sql & "Where T_indiceProjet.Id = " & Id_Pieces & " "
Sql = Sql & "Or T_indiceProjet.Pere = " & Id_Pieces & " "
Sql = Sql & "ORDER BY T_indiceProjet.Id DESC;"
Set Rs = Con.OpenRecordSet(Sql)
AtocatOption = True
While Rs.EOF = False
    Index = Index + 1
    Rs.MoveNext
Wend
  Rs.Requery
   FormBarGrah.ProgressBar1Caption = " Chargement des Options :"
     FormBarGrah.ProgressBar1.Value = 0
     FormBarGrah.ProgressBar1.Max = Index * 3
Set MyPotionEntete = New Collection
InsertPointLigneTableau_fils(0) = -1168#: InsertPointLigneTableau_fils(1) = 20#: InsertPointLigneTableau_fils(2) = 0
 InsertPointLigneTableau_fils2(1) = 20#: InsertPointLigneTableau_fils2(2) = 0
InsertPointLigneTableau_fils2(0) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils(0), -24.6369)
InsertPointLigneTableau_fils3(1) = 20#: InsertPointLigneTableau_fils3(2) = 0
InsertPointLigneTableau_fils3(0) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils2(0), -24.6369)
While Rs.EOF = False
IncremanteBarGrah FormBarGrah
Set Block = FunInsBlock(PathBlocs & "\RefOption.dwg", InsertPointLigneTableau_fils, "", 0, 0)
 InsertPointLigneTableau_fils(1) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils(1), -3)
 aa = Block.GetAttributes
 aa(0).TextString = UCase(Trim("" & Rs!PI) & "_" & Trim("" & Rs!PI_Indice))
 
 Set Block = FunInsBlock(PathBlocs & "\RefOption.dwg", InsertPointLigneTableau_fils2, "", 0, 0)
 InsertPointLigneTableau_fils2(1) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils2(1), -3)
 aa = Block.GetAttributes
 aa(0).TextString = UCase(Trim("" & Rs!RefPieceClient))
 

 
 Set Block = FunInsBlock(PathBlocs & "\RefOption.dwg", InsertPointLigneTableau_fils3, "", 0, 0)
 InsertPointLigneTableau_fils3(1) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils3(1), -3)
 aa = Block.GetAttributes
  Sql = "SELECT Rq_Select_Code_Critere.CloseWhere FROM Rq_Select_Code_Critere "
    Sql = Sql & "WHERE Rq_Select_Code_Critere.Id_IndiceProjet=" & Id_Pieces & "  "
    Sql = Sql & " AND Rq_Select_Code_Critere.CRITERES='" & MyReplace(Trim("" & Rs!Equipement)) & "';"
Set RsSelect = Con.OpenRecordSet(Sql)

 
 
 Sql = "SELECT Count(Ligne_Tableau_fils.Id_IndiceProjet) AS NbF "
    Sql = Sql & "FROM Ligne_Tableau_fils "
   
    Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet= " & Id_Pieces & "  "
    Sql = Sql & "AND Ligne_Tableau_fils.ACTIVER=True  "
    
     Sql = Sql & "AND " & MyReplaceSql(RsSelect, "Ligne_Tableau_fils", "CloseWhere", "OPTION") & " ;"
    


Set RsSelect = Con.OpenRecordSet(Sql)

 aa(0).TextString = UCase(Trim("" & RsSelect!NbF))
 
  Rs.MoveNext
Wend
 
 
 
   Set Block = FunInsBlock(PathBlocs & "\RefOption.dwg", InsertPointLigneTableau_fils, "", 0, 0)
 
' InsertPointLigneTableau_fils(0) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils(0), -24.6369)
  Set Block = FunInsBlock(PathBlocs & "\RefOption.dwg", InsertPointLigneTableau_fils2, "", 0, 0)
 
' Set Block = FunInsBlock("\\10.30.0.5\donnees d entreprise\Utilitaires\cablage\Librairies\Blocs de construction\RefOption.dwg", InsertPointLigneTableau_fils2, "", 0, 0)
 InsertPointLigneTableau_fils2(0) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils2(0), -24.6369)
  aa = Block.GetAttributes
 aa(0).TextString = UCase(Trim("ref client"))
 
 Set Block = FunInsBlock(PathBlocs & "\RefOption.dwg", InsertPointLigneTableau_fils3, "", 0, 0)
 
' Set Block = FunInsBlock("\\10.30.0.5\donnees d entreprise\Utilitaires\cablage\Librairies\Blocs de construction\RefOption.dwg", InsertPointLigneTableau_fils2, "", 0, 0)
 InsertPointLigneTableau_fils3(0) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils3(0), -24.6369)
  aa = Block.GetAttributes
 aa(0).TextString = UCase(Trim("Nb Fils"))
    Rs.Requery
 On Error Resume Next
 While Rs.EOF = False
 IncremanteBarGrah FormBarGrah
 az = Split(Trim("" & Rs!Equipement) & ";", ";")
    For I = LBound(aa) To UBound(az) - 1
    If Trim("" & az(I)) <> "" Then
    aa = MyPotionEntete(Trim("" & az(I)))
    If Err Then
        Err.Clear
'        1134.4342
   
        MyPotionEntete.Add Trim("" & az(I)), Trim("" & az(I))
            Set Block = FunInsBlock(PathBlocs & "\RefOption.dwg", InsertPointLigneTableau_fils3, "", 0, 0)
            InsertPointLigneTableau_fils3(0) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils3(0), -24.6369)
            aa = Block.GetAttributes
            aa(0).TextString = UCase(Trim("" & az(I)))
    End If
    End If
    Next
    Rs.MoveNext
Wend
InsertPointLigneTableau_fils(0) = -1168#: InsertPointLigneTableau_fils(1) = 20#: InsertPointLigneTableau_fils(2) = 0
    Rs.Requery
 
While Rs.EOF = False
 IncremanteBarGrah FormBarGrah
 DoEvents
InsertPointLigneTableau_fils(0) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils2(0), -24.6369)
    aa = MyPotionEntete(Trim("" & Rs!Equipement))
   For I = 1 To MyPotionEntete.Count
'   Sql = "SELECT  T_indiceProjet.* FROM T_indiceProjet "
'    Sql = Sql & "WHERE ([PI] & '_' & [PI_Indice]='" & Trim("" & Rs!PI) & "_" & Trim("" & Rs!PI_Indice) & "' "
'    Sql = Sql & "AND T_indiceProjet.Equipement='" & MyReplace(MyPotionEntete(I)) & "'"
'    Sql = Sql & "AND T_indiceProjet.Id=" & Id_Pieces & ") "
'    Sql = Sql & "OR ([PI] & '_' & [PI_Indice]='" & Trim("" & Rs!PI) & "_" & Trim("" & Rs!PI_Indice) & "'"
'    Sql = Sql & "AND T_indiceProjet.Equipement='" & MyReplace(MyPotionEntete(I)) & "'"
'    Sql = Sql & "AND T_indiceProjet.Pere=" & Id_Pieces & ");"
'
'
'     Sql = "SELECT Count(T_indiceProjet.Id) AS NbF "
'    Sql = Sql & "FROM T_indiceProjet INNER JOIN Ligne_Tableau_fils ON T_indiceProjet.Id = Ligne_Tableau_fils.Id_IndiceProjet "
'    Sql = Sql & "WHERE ([PI] & '_' & [PI_Indice]='" & Trim("" & Rs!PI) & "_" & Trim("" & Rs!PI_Indice) & "' "
'    Sql = Sql & "AND  [Equipement]  Like '%" & MyReplace(MyPotionEntete(I)) & ";%' "
'    Sql = Sql & "AND T_indiceProjet.Id=" & Id_Pieces & " and Ligne_Tableau_fils.ACTIVER=True) "
'    Sql = Sql & "OR ([PI] & '_' & [PI_Indice]='" & Trim("" & Rs!PI) & "_" & Trim("" & Rs!PI_Indice) & "' "
'    Sql = Sql & "AND  ';' & [Equipement] & ';' Like '%" & MyReplace(MyPotionEntete(I)) & ";%' "
'    Sql = Sql & "AND T_indiceProjet.Pere=" & Id_Pieces & ");"
'
'Set RsSelect = Con.OpenRecordSet(Sql)
'
'    Set Block = FunInsBlock(PathBlocs & "\RefOption.dwg", InsertPointLigneTableau_fils, "", 0, 0)
'
'             InsertPointLigneTableau_fils(0) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils(0), -24.6369)
'            aa = Block.GetAttributes
'            If RsSelect.EOF = False Then
'
'            aa(0).TextString = RsSelect!NbF
'            Else
'             aa(0).TextString = ""
'            End If
    
    
    
    Sql = "SELECT T_indiceProjet.* "
    Sql = Sql & "FROM T_indiceProjet "
    Sql = Sql & "WHERE ([PI] & '_' & [PI_Indice]='" & Trim("" & Rs!PI) & "_" & Trim("" & Rs!PI_Indice) & "' "
    Sql = Sql & "AND  ';' & [Equipement] & ';' Like '%;" & MyReplace(MyPotionEntete(I)) & ";%'  "
    Sql = Sql & "AND T_indiceProjet.Id=" & Rs!Id & ") "
    Sql = Sql & ";"

Set RsSelect = Con.OpenRecordSet(Sql)
       
            Set Block = FunInsBlock(PathBlocs & "\RefOption.dwg", InsertPointLigneTableau_fils, "", 0, 0)
            InsertPointLigneTableau_fils(0) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils(0), -24.6369)
            aa = Block.GetAttributes
'            RsSelect.Requery
'            While RsSelect.EOF = False
'            MsgBox RsSelect!Id
'                RsSelect.MoveNext
'            Wend
            If RsSelect.EOF = False Then
            
            aa(0).TextString = "X"
            Else
             aa(0).TextString = ""
            End If
    
    Next
    InsertPointLigneTableau_fils(1) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils(1), -3)
    Rs.MoveNext
Wend
InsertPointLigneTableau_fils(0) = -1168#
' Rs.Requery
'While Rs.EOF = False
' IncremanteBarGrah FormBarGrah
' DoEvents
'InsertPointLigneTableau_fils(0) = DecalInsertPointLigneTableau_fils_Bas(DecalInsertPointLigneTableau_fils_Bas(-1168#, -24.6369), -24.6369)
'    aa = MyPotionEntete(Trim("" & Rs!Equipement))
'   For I = 1 To MyPotionEntete.Count
'   Sql = "SELECT  T_indiceProjet.* FROM T_indiceProjet "
'    Sql = Sql & "WHERE ([PI] & '_' & [PI_Indice]='" & Trim("" & Rs!PI) & "_" & Trim("" & Rs!PI_Indice) & "' "
'    Sql = Sql & "AND T_indiceProjet.Equipement='" & MyReplace(MyPotionEntete(I)) & "'"
'    Sql = Sql & "AND T_indiceProjet.Id=" & Id_Pieces & ") "
'    Sql = Sql & "OR ([PI] & '_' & [PI_Indice]='" & Trim("" & Rs!PI) & "_" & Trim("" & Rs!PI_Indice) & "'"
'    Sql = Sql & "AND T_indiceProjet.Equipement='" & MyReplace(MyPotionEntete(I)) & "'"
'    Sql = Sql & "AND T_indiceProjet.Pere=" & Id_Pieces & ");"
'
'
'     Sql = "SELECT Count(T_indiceProjet.Id) AS NbF "
'    Sql = Sql & "FROM T_indiceProjet INNER JOIN Ligne_Tableau_fils ON T_indiceProjet.Id = Ligne_Tableau_fils.Id_IndiceProjet "
'    Sql = Sql & "WHERE ([PI] & '_' & [PI_Indice]='" & Trim("" & Rs!PI) & "_" & Trim("" & Rs!PI_Indice) & "' "
'    Sql = Sql & "AND  [Equipement]  Like '%" & MyReplace(MyPotionEntete(I)) & ";%' "
'    Sql = Sql & "AND T_indiceProjet.Id=" & Id_Pieces & " and Ligne_Tableau_fils.ACTIVER=True) "
'    Sql = Sql & "OR ([PI] & '_' & [PI_Indice]='" & Trim("" & Rs!PI) & "_" & Trim("" & Rs!PI_Indice) & "' "
'    Sql = Sql & "AND  ';' & [Equipement] & ';' Like '%" & MyReplace(MyPotionEntete(I)) & ";%' "
'    Sql = Sql & "AND T_indiceProjet.Pere=" & Id_Pieces & ");"
'
'    Sql = "SELECT Count(Ligne_Tableau_fils.Id_IndiceProjet) AS NbF "
'    Sql = Sql & "FROM Ligne_Tableau_fils "
'    Sql = Sql & "WHERE (';' & [OPTION] & ';' Like '%;" & MyReplace(MyPotionEntete(I)) & ";%' or ';' & [OPTION] & ';' Like '%;TOUS;%') "
'    Sql = Sql & "AND Ligne_Tableau_fils.Id_IndiceProjet= " & Id_Pieces & "  "
'    Sql = Sql & "AND Ligne_Tableau_fils.ACTIVER=True;"
'
'
'Sql = "SELECT Rq_Select_Code_Critere.CloseWhere FROM Rq_Select_Code_Critere "
'    Sql = Sql & "WHERE Rq_Select_Code_Critere.Id_IndiceProjet=" & Id_Pieces & "  "
'    Sql = Sql & "AND Rq_Select_Code_Critere.CRITERES='" & MyReplace(MyPotionEntete(I)) & "';"
'Set RsSelect = Con.OpenRecordSet(Sql)
'
'
'
'Sql = "SELECT Count(Ligne_Tableau_fils.Id_IndiceProjet) AS NbF "
'    Sql = Sql & "FROM Ligne_Tableau_fils "
'
'    Sql = Sql & "WHERE Ligne_Tableau_fils.Id_IndiceProjet= " & Id_Pieces & "  "
'    Sql = Sql & "AND Ligne_Tableau_fils.ACTIVER=True  "
'
'     Sql = Sql & "AND " & MyReplaceSql(RsSelect, "Ligne_Tableau_fils", "CloseWhere", "OPTION") & " ;"
'
'
'
'Set RsSelect = Con.OpenRecordSet(Sql)
'
'    Set Block = FunInsBlock(PathBlocs & "\RefOption.dwg", InsertPointLigneTableau_fils, "", 0, 0)
'
'             InsertPointLigneTableau_fils(0) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils(0), -24.6369)
'            aa = Block.GetAttributes
'            If RsSelect.EOF = False Then
''                If RsSelect!NbF = 0 Then
''                    Block.Delete
''                Else
'                    aa(0).TextString = RsSelect!NbF
''                End If
''            Else
''                Block.Delete
'            End If
'
'    Next
'    Rs.MoveNext
'    Wend
'MyPotionEntete.Add Block, i1
End Function
Sub Racourci(RaccourciName As String, RaccourciCible As String, Extension As String)
Dim Fso As New FileSystemObject
If Fso.FileExists(RaccourciName & ".Lnk") = True Then
     Fso.DeleteFile RaccourciName & ".Lnk"
End If
Set objshell = CreateObject("wscript.shell")
Set objraccourci = objshell.createshortcut(RaccourciName & ".Lnk")
objraccourci.targetpath = RaccourciCible & "." & Extension
objraccourci.Save
Set Fso = Nothing
Set objraccourci = Nothing
End Sub
Sub IncremanteBarGrah(obj As Object)
On Error Resume Next
If obj.ProgressBar1.Max = obj.ProgressBar1.Value Then
            obj.ProgressBar1.Max = obj.ProgressBar1.Max + 1
        End If
         obj.ProgressBar1.Value = obj.ProgressBar1.Value + 1
        
         DoEvents
 On Error GoTo 0
End Sub

Sub EcritureTor(RsLigne As Recordset, Mytype As String)
If (bool_Plan_E_Preconisations = False And Mytype = "PL") Or (bool_Outil_E_Preconisations = False And Mytype = "OU") Then Exit Sub

 While RsLigne.EOF = False
 
 
        If Trim("" & RsLigne!PRECOG) <> "" And RsLigne!Activer = True Then
        On Error Resume Next
        Set a = Nothing
        DoEvents
            a = ""
            a = CollectionTor("" & RsLigne!App)
            If Err Then
            Err.Clear
                NUMNTORBLOC = NUMNTORBLOC + 1
Reprise:        If Trim("" & RsLigne!App) = "" Then GoTo GotoMoveNext1
                CollectionTor.Add NUMNTORBLOC, Trim("" & RsLigne!App)
                ReDim Preserve TableuDeTor(CollectionTor(Trim("" & RsLigne!App)))
            End If
                TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).CodeApp = Trim("" & RsLigne!App)
                 TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).Garder = True
                If TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).CodeApp = "" Then
                   GoTo Reprise
                 End If
                  Set a = Nothing
        DoEvents
                a = TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).CollectionTor("" & RsLigne!PRECOG)
                If Err Then
            Err.Clear
                TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).NumTor = TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).NumTor + 1
               TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).CollectionTor.Add TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).NumTor, "" & RsLigne!PRECOG
                ReDim Preserve TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).CollectionTor("" & RsLigne!PRECOG))
            End If
            If InStr(1, TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).CollectionTor("" & RsLigne!PRECOG)).TableauFile, "" & RsLigne!Fil & " ") = 0 Then
              TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).CollectionTor("" & RsLigne!PRECOG)).TableauFile = TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).CollectionTor("" & RsLigne!PRECOG)).TableauFile & RsLigne!Fil & " "
               TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).CollectionTor("" & RsLigne!PRECOG)).TorName = "" & RsLigne!PRECOG
               TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!App))).CollectionTor("" & RsLigne!PRECOG)).Garder = True
           End If
            Set a = Nothing
        DoEvents
           On Error Resume Next
        Set a = Nothing
        DoEvents
        a = ""
            a = CollectionTor("" & RsLigne!App2)
            If Err Then
            Err.Clear
                NUMNTORBLOC = NUMNTORBLOC + 1
                CollectionTor.Add NUMNTORBLOC, Trim("" & RsLigne!App2)
Reprise2:
GotoMoveNext1:
               If Trim("" & RsLigne!App2) = "" Then GoTo GotoMoveNext
                ReDim Preserve TableuDeTor(CollectionTor(Trim("" & RsLigne!App2)))
            End If
                TableuDeTor(CollectionTor(Trim("" & RsLigne!App2))).CodeApp = Trim("" & RsLigne!App2)
                  TableuDeTor(CollectionTor(Trim("" & RsLigne!App2))).Garder = True
                If TableuDeTor(CollectionTor(Trim("" & RsLigne!App2))).CodeApp = "" Then
                   GoTo Reprise2
                 End If
                  Set a = Nothing
        DoEvents
            a = ""
                a = TableuDeTor(CollectionTor(Trim("" & RsLigne!App2))).CollectionTor("" & RsLigne!PRECOG)
                If Err Then
            Err.Clear
                TableuDeTor(CollectionTor(Trim("" & RsLigne!App2))).NumTor = TableuDeTor(CollectionTor(Trim("" & RsLigne!App2))).NumTor + 1
               TableuDeTor(CollectionTor(Trim("" & RsLigne!App2))).CollectionTor.Add TableuDeTor(CollectionTor(Trim("" & RsLigne!App2))).NumTor, "" & RsLigne!PRECOG
                ReDim Preserve TableuDeTor(CollectionTor(Trim("" & RsLigne!App2))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!App2))).CollectionTor("" & RsLigne!PRECOG))
            End If
            If InStr(1, TableuDeTor(CollectionTor(Trim("" & RsLigne!App2))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!App2))).CollectionTor("" & RsLigne!PRECOG)).TableauFile, "" & RsLigne!Fil & " ") = 0 Then
                 TableuDeTor(CollectionTor(Trim("" & RsLigne!App2))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!App2))).CollectionTor("" & RsLigne!PRECOG)).Garder = True
              TableuDeTor(CollectionTor(Trim("" & RsLigne!App2))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!App2))).CollectionTor("" & RsLigne!PRECOG)).TableauFile = TableuDeTor(CollectionTor(Trim("" & RsLigne!App2))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!App2))).CollectionTor("" & RsLigne!PRECOG)).TableauFile & RsLigne!Fil & " "
               TableuDeTor(CollectionTor(Trim("" & RsLigne!App2))).Tor(TableuDeTor(CollectionTor(Trim("" & RsLigne!App2))).CollectionTor("" & RsLigne!PRECOG)).TorName = "" & RsLigne!PRECOG
           End If
            Set a = Nothing
        DoEvents
            On Error GoTo 0
        End If
GotoMoveNext:
        RsLigne.MoveNext
    Wend

End Sub
Sub Ecriturefils(RsLigne As Recordset, Mytype As String, NbFils As Long)
Dim Fso As New FileSystemObject
Dim NbSuprime As Long
Dim Attributes
If (bool_Plan_E_Fils = True And Mytype = "PL") Or (bool_Outil_E_Fils = True And Mytype = "OU") Then
    
    If Val(NbFils) <> 0 Then
    RsLigne.MoveFirst
    End If
     FormBarGrah.ProgressBar1.Value = 0
    If Val(NbFils) <> 0 Then
     FormBarGrah.ProgressBar1.Max = 1 + NbFils
    Else
         FormBarGrah.ProgressBar1.Max = 1 + 1
    End If
     FormBarGrah.ProgressBar1Caption.Caption = " Chargement de la liste de fils"
InsertPointLigneTableau_fils(0) = Val(GetDefault("TableauxFilsX", "-1151")): InsertPointLigneTableau_fils(1) = Val(GetDefault("TableauxFilsY", "-43"))
    While RsLigne.EOF = False
         IncremanteBarGrah FormBarGrah
         IncrmentServer FormBarGrah, Mytype
        DoEvents
'        AutoApp.Documents(1).ZoomAll

        ReDim tableau(RsLigne.Fields.Count - 3)
        If "a" = "a" Then
            For Col = 0 To RsLigne.Fields.Count - 3
                DoEvents
                tableau(Col) = "" & RsLigne.Fields(Col).Name & ";" & RsLigne.Fields(Col)
            Next Col
    
            RenseigneConnecteurBroches RsLigne, Mytype
            Row = Row + 1
            If NbLignes = 100 Then
                InsertPointLigneTableau_fils(1) = -43: InsertPointLigneTableau_fils(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPointLigneTableau_fils(0), Val(GetDefault("TableauxFilsDroite", "377")))  '281.7719)
                NbLignes = 0
            End If

            If NbLignes = 0 Then
                If Fso.FileExists(PathBlocs & "\TITRE_TABLEAU_FILS.dwg") = False Then
                    MsgBox "err"
                End If
                Set NewBlock = FunInsBlock(PathBlocs & "\TITRE_TABLEAU_FILS.dwg", InsertPointLigneTableau_fils, "E" & CStr(Row))
                InsertPointLigneTableau_fils(1) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils(1), Val(GetDefault("TableauxFilsBas", "6")))
            End If
            If RsLigne!Activer = True Then
                filedwg = "1TABLEAU_FILS.dwg"
             Else
                filedwg = "0TABLEAU_FILS.dwg"
                NbSuprime = NbSuprime + 1
             End If
            If Fso.FileExists(PathBlocs & "\" & filedwg) = False Then
                MsgBox "err"
            End If
            Set NewBlock = FunInsBlock(PathBlocs & "\" & filedwg, InsertPointLigneTableau_fils, "L" & CInt(Row))
            Attributes = NewBlock.GetAttributes
            InsertPointLigneTableau_fils(1) = DecalInsertPointLigneTableau_fils_Bas(InsertPointLigneTableau_fils(1), Val(GetDefault("TableauxFilsBas", "6")))

            NbLignes = NbLignes + 1
            On Error GoTo Error1
            Lib1 = RsLigne.Fields(2)
            Lib2 = "" & RsLigne.Fields("APP")
            If RsLigne!Activer = True Then
            TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP")))).indexFile = TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP")))).indexFile + 1
            End If
            Lib1 = ""
            Lib2 = ""
            If RsLigne!Activer = True Then
            ReDim Preserve TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP")))).TableauFile(TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP")))).indexFile)
            TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP")))).TableauFile(TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP")))).indexFile) = RsLigne.Fields(2)
    End If
            Lib1 = RsLigne.Fields(2)
            Lib2 = "" & RsLigne.Fields("APP2")
             If RsLigne!Activer = True Then
            TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP2")))).indexFile = TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP2")))).indexFile + 1
            End If
            Lib1 = ""
            Lib2 = ""
            ReDim Preserve TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP2")))).TableauFile(TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP2")))).indexFile)

            TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP2")))).TableauFile(TableauDeConnecteurs(CLng(CollectionCon("" & RsLigne.Fields("APP2")))).indexFile) = RsLigne.Fields(2)
    
            a = RsLigne.Fields(0)
            funAttributesLigne_Tableau_fils NewBlock.Name, NewBlock.GetAttributes, tableau, RsLigne.Fields.Count - 3
            Else
           'NbSupprim = 'NbSupprim + 1
          End If
          InsertPointLigneCritères(0) = InsertPointLigneTableau_fils(0)
          InsertPointLigneCritères(1) = InsertPointLigneTableau_fils(1)
           InsertPointLigneCritères(2) = InsertPointLigneTableau_fils(2)
          RsLigne.MoveNext
     
        Wend
        If NbFils > 0 Then
            Set NewBlock = FunInsBlock(PathBlocs & "\Nombre_fils.dwg", InsertPointLigneTableau_fils, "N1")
            Attri = NewBlock.GetAttributes
            Attri(0).TextString = NbFils - NbSuprime
         End If
       Else
          If Val(NbFils) <> 0 Then
    RsLigne.MoveFirst
    End If
     FormBarGrah.ProgressBar1.Value = 0
    If Val(NbFils) <> 0 Then
     FormBarGrah.ProgressBar1.Max = 1 + NbFils
    Else
         FormBarGrah.ProgressBar1.Max = 1 + 1
    End If
     FormBarGrah.ProgressBar1Caption.Caption = " Chargrment de la liste de fils"

    While RsLigne.EOF = False
         IncremanteBarGrah FormBarGrah
        DoEvents
'        AutoApp.Documents(1).ZoomAll

        ReDim tableau(RsLigne.Fields.Count)
        If RsLigne!Activer = True Then
          
     
            RenseigneConnecteurBroches RsLigne, Mytype
           
           
          RsLigne.MoveNext
        End If
        Wend
       
      
       End If
        SacnConnecteur Mytype
        Set Fso = Nothing
        Exit Sub
Error1:
    FunError 3, CStr("" & Lib1), CStr("" & Lib2)
Resume Next
End Sub
Function NoeuName(Row As Long)
Dim Txt As String
Dim Ofset As Long
Dim nbTour As Long
Dim NbTord As Long
Dim txtColone As Long
Dim txtNuberColone As Long

Txt = "AA"
txtColone = Len(Txt)
txtNuberColone = Len(Txt)
Ofset = 0
nbTour = 0
NbTord = 0


For I = 0 To Row - 3
Reprise:
Mid(Txt, txtColone, 1) = Chr(Asc(Mid(Txt, txtColone, 1)) + 1)
DoEvents
If Asc(Mid(Txt, txtColone, 1)) = 91 Then
Mid(Txt, txtColone, 1) = "A"
txtColone = txtColone - 1
If txtColone = 0 Then
    Txt = Txt & "A"
    txtColone = Len(Txt)
Else
    GoTo Reprise
End If

End If
   If txtColone <> Len(Txt) Then txtColone = Len(Txt)



Next

NoeuName = Txt
End Function

Sub RazFiltreEditExcel(MySpreadsheet As Object)
On Error Resume Next
 MySpreadsheet.ActiveSheet.AutoFilterMode = False
'If MySpreadsheet.ActiveSheet.AutoFilterMode = True Then
'    For I = 1 To Myrange.Columns.Count
'    Set aa = MySpreadsheet.ActiveSheet.AutoFilter.Filters(I).Criteria
'    aa.Show All = True
'
'    Next
   MySpreadsheet.ActiveSheet.AutoFilterMode = True
    MySpreadsheet.ActiveSheet.Range("A1").AutoFilter
    MySpreadsheet.ActiveSheet.AutoFilter.Apply
DoEvents
'End If

End Sub

Function BackUp(Fichier As String, Optional LI As Boolean, Optional MyPathXlsMoins1 As String) As String
BackUp = MyPathXlsMoins1
If LI = True And Bool_Fichier_Li = True Then Exit Function
BackUp = ""
If Fichier = "" Then Exit Function
Dim Fso As New FileSystemObject
If Fso.FileExists(Fichier) = False Then
    Set Fso = Nothing
    Exit Function
End If
Dim Path
Dim PathAs As String
Dim SaveAs As String
Path = Split(Fichier, "\")
PathAs = ""
For I = LBound(Path) To UBound(Path) - 1
PathAs = PathAs & Path(I) & "\"
Debug.Print PathAs
Next
PathAs = PathAs & "Archives"
Debug.Print PathAs
If Fso.FolderExists(PathAs) = False Then
Fso.CreateFolder PathAs
End If
PathAs = PathAs & "\"
SaveAs = Format(Now, "yyyy-mm-dd-h-m-s_") & Path(UBound(Path))
While Fso.FileExists(PathAs & SaveAs) = True
    SaveAs = Format(Date, "yyyy-mm-dd-h-m-s_") & Path(UBound(Path))
Wend
Debug.Print PathAs & SaveAs
Fso.CopyFile Fichier, PathAs & SaveAs
BackUp = PathAs & SaveAs
If LI = True Then
    Bool_Fichier_Li = True
End If
Set Fso = Nothing
End Function
Function IndexationNoeuds(NumNoeud As String) As Long
Dim I As Long
Dim Txt As String
Txt = "AA"
 I = 2
While Txt <> NumNoeud
    I = I + 1
    Txt = NoeuName(I)
Wend
IndexationNoeuds = I - 1
End Function

Function InsertPointHiérarchie(Num As Long, PoseXinit As Double, PoseYinit As Double, DecalX As Double, DecalY As Double, NbBloc As Long)
Dim InsertPoint(0 To 2) As Double
Dim nbTour As Long
Dim NbTour2 As Long
nbTour = 0
InsertPoint(0) = PoseXinit
InsertPoint(1) = PoseYinit
For I = 1 To Num
    If nbTour = NbBloc Then
    NbTour2 = NbTour2 + 1
       InsertPoint(1) = PoseYinit
       InsertPoint(0) = DecalInsertPointLigneTableau_fils_Gauche(PoseXinit, DecalX * NbTour2)
       nbTour = 0
    End If
     If nbTour <> 0 Then
        InsertPoint(0) = DecalInsertPointLigneTableau_fils_Gauche(InsertPoint(0), DecalX)
        InsertPoint(1) = DecalInsertPointLigneTableau_fils_Gauche(InsertPoint(1), DecalY)
       
     End If
    nbTour = nbTour + 1
Next
InsertPointHiérarchie = InsertPoint
End Function
'Global Sub FormatExcelPlage(Plage As Range, Couleur As Long, Merge As Boolean, Grille As Boolean, HorizontalAlignment As Long, VerticalAlignment As Long, Optional ZoneImpressionOfset As Long)
'Plage.Interior.ColorIndex = Couleur
'If Merge = True Then Plage.Merge
'    Plage.HorizontalAlignment = HorizontalAlignment 'xlCenter
'    Plage.VerticalAlignment = VerticalAlignment 'xlCenter
'If Grille = True Then
'    Plage.Borders(xlEdgeLeft).LineStyle = xlContinuous
'    Plage.Borders(xlEdgeTop).LineStyle = xlContinuous
'    Plage.Borders(xlEdgeBottom).LineStyle = xlContinuous
'    Plage.Borders(xlEdgeRight).LineStyle = xlContinuous
'    Plage.Borders(xlContinuous).LineStyle = xlContinuous
'End If
'
'
'End Sub
Function DefinirChemienComplet(Serveur As String, Path As String) As String
If Right(Trim("" & Serveur), 1) <> "\" Then Serveur = Serveur & "\"
If Trim("" & Path) = "" Then
    DefinirChemienComplet = Serveur
Else
    If Left(Path, 1) = "\" And Left(Path, 2) <> "\\" Then Path = Right(Path, Len(Path) - 1)
DefinirChemienComplet = Path
End If
If Mid(DefinirChemienComplet, 2, 1) = ":" Then Exit Function
If Left(Path, 1) <> "\" Then
    If Right(Serveur, 1) <> "\" Then
        DefinirChemienComplet = Serveur & "\" & DefinirChemienComplet
    Else
         DefinirChemienComplet = Serveur & DefinirChemienComplet
    End If
End If
If Right(Trim(DefinirChemienComplet), 2) = "\\" Then DefinirChemienComplet = Mid(DefinirChemienComplet, 1, Len(DefinirChemienComplet) - 1)
If Left(DefinirChemienComplet, 1) = "\" And Left(DefinirChemienComplet, 2) <> "\\" Then DefinirChemienComplet = "\" & DefinirChemienComplet
Debug.Print DefinirChemienComplet
End Function

Function LirDoc() As Boolean
Dim MyDir As String
Dim Sql As String
Dim Rs As Recordset
Sql = "UPDATE Document SET Document.Supprimer = True;"
Con.Execute Sql

MyDir = Dir(App.Path & "\RepDoc\*.*")
While MyDir <> ""
    
    Sql = "SELECT Document.Documment FROM Document "
    Sql = Sql & "WHERE Document.Machine='" & MyReplace(Machine) & "'AND Document.Documment='" & MyReplace(MyDir) & "' "
    Sql = Sql & "AND Document.UserName='" & MyReplace(UserName) & "';"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = True Then
       
        Sql = "INSERT INTO Document ( Documment, UserName,Machine ) "
        Sql = Sql & "Values( '" & MyReplace(MyDir) & "', '" & MyReplace(UserName) & "', '" & MyReplace(Machine) & "' );"
        Con.Execute Sql
    End If
    
    Sql = "SELECT Document.Documment FROM Document "
    Sql = Sql & "WHERE Document.Machine='" & MyReplace(Machine) & "'AND Document.Documment='" & MyReplace(MyDir) & "' "
    Sql = Sql & "AND Document.UserName='" & MyReplace(UserName) & "' AND Document.PlusAficher=False;"
    Set Rs = Con.OpenRecordSet(Sql)
    If Rs.EOF = False Then
     LirDoc = True
     End If
    Set Rs = Con.CloseRecordSet(Rs)
    Sql = "UPDATE Document SET Document.Supprimer = False "
    Sql = Sql & "WHERE Document.Documment= '" & MyReplace(MyDir) & "';"
    Con.Execute Sql

    MyDir = Dir
Wend
Sql = "DELETE Document.*, Document.Supprimer FROM Document "
Sql = Sql & "WHERE Document.Supprimer=True;"
Con.Execute Sql
End Function
Sub Main()
Dim Poison As String
Dim Rs As Recordset
Dim RsCible As Recordset
Dim RsSource As Recordset
Dim DbOk As Boolean
Dim Sql As String
Dim ConCommposants As New Ado
Dim NbRegistrer As Long
Dim MyForm As New frmInit
Dim Maintenace As String
NomenclatureOk = True
PossibleArretKill = False
Poison = Format(Date, "mm-dd")
If Dir(App.Path & "\BdAutoCâble.dll") <> "" Then
NumFile = FreeFile
Open App.Path & "\BdAutoCâble.dll" For Random As #NumFile
Get #NumFile, , PassDb

Close #NumFile
PassDb.PassWordDb = CodageX.Decrypt(PassDb.PassWordDb, "")
PassDb.UserDb = CodageX.Decrypt(PassDb.UserDb, "")

End If
UserName = LirUserName
Machine = LirMachineName
LoadDb
Sql = "INSERT INTO [Utilise_Par] ( Machine,User ) "
Sql = Sql & "VALUES  ('" & MyReplace(Machine) & "','" & MyReplace(UserName) & "');"
Con.Execute Sql
Maintenace = CherCheInFihier("Maintenace")
If IsServeur = False Then
If Poison = "04-01" Then Poisson.Show vbModal ' MsgBox "Erreur du premier avril", vbCritical
    If UCase(Maintenace) = "TRUE" Then
        MsgBox "L'application à été arrêtée pour maintenance." & vbCrLf & "Veuillez contacter l'administrateur de l'application", vbCritical
        End
    End If
    If App.PrevInstance = True Then
        MsgBox "Une instance du programme à déjà été lancer, impossible de lancer une nouvelle instance.", vbOKOnly + vbCritical, "Autocâble"
        ' Ferme le programme
        End

    End If
Else
      If UCase(Maintenace) = "TRUE" Then
        
        End
    End If
    
   
End If


CodageX.LireLicence


'LoadDb
 If IsServeur = False Then
 loadFichierAppliWin
'    Set Rs = Con.OpenRecordSet("SELECT T_Boutons.Name, T_Boutons.ContonTotal FROM T_Boutons WHERE T_Boutons.ContonTotal=True;")

If GestionDesDroit("Application") = False Then End

End If

If LirDoc = True Then
    If IsServeur = False Then
    
        FremMsgMaj.Show vbModal
    End If
End If

 Set TableauPath = funPath
 
If IsServeur = True Then
    If Trim("" & Command) <> "" Then Job.Show vbModal
    Con.CloseConnection
End
End If
DbCatalogue = TableauPath("Catalogue")
DbCatalogue = DefinirChemienComplet(TableauPath.Item("PathServer"), DbCatalogue)
'If UCase(DefinirChemienComplet(TableauPath.Item("IsCilent"), DbCatalogue)) = "TRUE" Then IsCilent = t

'Sql = "SELECT TableIniDate.Lib, TableIniDate.MyDate FROM TableIniDate WHERE TableIniDate.Lib='BaseCatalogue';"
''Con.OpenConnetion db
'Set Rs = Con.OpenRecordSet(Sql)
'
'If Format(Rs!MyDate, "yyyy-mm-dd") <> Format(Date, "yyyy-mm-dd") Then
''    MyForm.Show
'    DoEvents
'    Rs!MyDate = Format(Date, "yyyy-mm-dd")
'    Rs.Update
'
'Debug.Print Db
'Dim DbEboutique As New Ado
'
'DbEboutique.TYPEBASE = ADO_TYPEBASE
'DbEboutique.SERVER = ADO_SERVER
'DbEboutique.User = ADO_User
'DbEboutique.PassWord = ADO_PassWord
'DbEboutique.BASE = DefinirChemienComplet(TableauPath.Item("PathServer"), TableauPath("Eb_CONNECTIQUE"))
'DbEboutique.OpenConnetion
'
'ConCommposants.TYPEBASE = ADO_TYPEBASE
'ConCommposants.SERVER = ADO_SERVER
'ConCommposants.User = ADO_User
'ConCommposants.PassWord = ADO_PassWord
'ConCommposants.BASE = DbCatalogue
'DbOk = ConCommposants.OpenConnetion
'
'    Sql = "SELECT con_contacts.txt9 AS [Prix u], con_contacts.txt1 AS [Alvé Réf], con_contacts.txt41 AS Mini,  "
'Sql = Sql & "con_contacts.txt6 AS Maxi, lst21.CatName AS Famille, con_contacts.txt3 AS [Alvé Réf Fourr],  "
'Sql = Sql & "lst9.CatName AS [Alvé Four] "
'Sql = Sql & "FROM (con_contacts LEFT JOIN lst21 ON con_contacts.lst21 = lst21.CatID)  "
'Sql = Sql & "LEFT JOIN lst9 ON con_contacts.lst9 = lst9.CatID "
'Sql = Sql & "WHERE con_contacts.txt41<>'';"
'
'
'   Set RsSource = DbEboutique.OpenRecordSet(Sql)
'   NbRegistrer = 0
'   While RsSource.EOF = False
'            NbRegistrer = NbRegistrer + 1
'            RsSource.MoveNext
'        Wend
'        Sql = "DELETE T_Alve_Eboutique.* FROM T_Alve_Eboutique;"
'Con.Execute Sql
'Sql = "SELECT T_Alve_Eboutique.* FROM T_Alve_Eboutique;"
'Set RsCible = Con.OpenRecordSet(Sql)
' RsSource.Requery
'        MyForm.ProgressBar1.Value = 0
'        MyForm.ProgressBar1.Max = NbRegistrer
'        MyForm.Visible = True
'        While RsSource.EOF = False
'        DoEvents
'            IncremanteBarGrah MyForm
'            RsCible.AddNew
'            RsCible![Alvé Fornisseur] = RsSource![Alvé Four]
''            RsCible![Nb Alvé] = RsSource![Nb Alvé]
''            RsCible!Voie = RsSource!Voie
'            RsCible![Famille Lib] = RsSource!Famille
''            RsCible![Famille Lib] = RsSource![Famille Lib]
'            RsCible![Alvé Réf] = RsSource![Alvé Réf]
''            RsCible!Qté = ConvertTxtAsDouble("" & RsSource!Qté)
'            RsCible![Prix u] = ConvertTxtAsDouble("" & RsSource![Prix u])
''            RsCible![Prix Total] = RsSource![Prix Total]
'            RsCible![Alvé Réf Fourr] = RsSource![Alvé Réf Fourr]
'            RsCible![Mini] = ConvertTxtAsDouble("" & RsSource![Mini])
'            RsCible!Maxi = ConvertTxtAsDouble("" & RsSource!Maxi)
'            RsCible.Update
'            RsSource.MoveNext
'        Wend
'        DbEboutique.CloseConnection
'End If
'Unload MyForm
'Set MyForm = Nothing
'Set Rs = Con.CloseRecordSet(Rs)
'LoadDb
'Restart.Show
frmAutocâble.Show
'Menu.Show


End Sub

Function ConvertTxtAsDouble(Txt) As Double
Txt = Replace(Txt, ",", ".")
Txt = Replace(Txt, "mm2", "")
Txt = Trim(Txt)
ConvertTxtAsDouble = Val(Txt)
DoEvents
End Function
Sub DeletSheet(MySheet As EXCEL.Worksheet)
    On Error Resume Next
    MySheet.Delete
Err.Clear
End Sub
Sub IncrmentServer(FormBarGrah As Object, Optional Mytype As String)
On Error Resume Next

Dim Sql As String

    If Trim("" & Mytype) = "" Then
    
        Sql = "UPDATE T_Job SET Status ='" & MyReplace(FormBarGrah.ProgressBar1Caption) & "' "
    Else
        Sql = "UPDATE T_Job SET Status ='" & MyReplace(Mytype) & " : " & MyReplace(FormBarGrah.ProgressBar1Caption) & "'"
    End If
    Sql = Sql & ",MaxBarGraph = " & FormBarGrah.ProgressBar1.Max & " "
    Sql = Sql & ",ValBarGraph = " & FormBarGrah.ProgressBar1.Value & ", T_Job.BarGraphMaj = Now() "
    Sql = Sql & "WHERE T_Job.Job= " & Command & ";"
    Con.Execute Sql
    Con.Execute Sql
    Con.Execute Sql
Err.Clear


End Sub

Public Function MyReplaceSql(Rs As Recordset, Table As String, ChampSource As String, Champ As String) As String
Dim T_Txt
Dim I As Long
Rs.Requery
MyReplaceSql = " [" & Table & "].[" & Champ & "] ='TOUS' "
While Rs.EOF = False
    If Trim("" & Rs(ChampSource)) <> "" Then
        If InStr(1, MyReplaceSql, " or [" & Table & "].[" & Champ & "] ='" & MyReplace(Rs(ChampSource)) & "' ") = 0 Then
            MyReplaceSql = MyReplaceSql & " or [" & Table & "].[" & Champ & "] ='" & MyReplace(Rs(ChampSource)) & "' "
            Debug.Print MyReplaceSql
        End If
    End If
    Rs.MoveNext
Wend
MyReplaceSql = " (" & MyReplaceSql & ") "
Debug.Print MyReplaceSql
End Function
Public Sub Copy_Rs_Spreadsheet(FRM As Form, Spreadsheet, Rs As Recordset, Mytype As String, FrmApelan As Object, LibAction As String)
FrmApelan.ProgressBar1Caption = LibAction
DoEvents
On Error Resume Next
Dim L As Long
Dim I As Long
Dim Save_ProgressBar1Value As Long
Dim Save_ProgressBar1Max As Long
Save_ProgressBar1Value = FrmApelan.ProgressBar1.Value
Save_ProgressBar1Max = FrmApelan.ProgressBar1.Value
For I = 0 To Rs.Fields.Count - 1
DoEvents
    Spreadsheet.Cells(1, I + 1) = "'" & Rs(I).Name
    Spreadsheet.Cells(1, I + 1).Interior.Color = ChoixCouleur(0)
Next
Const sDelimiteur$ = vbTab
If Rs.EOF = False Then
    toto = Rs.GetString(, , sDelimiteur$ & "'", "¤")
     toto = Replace(toto, Chr(10), "©")
      toto = Replace(toto, Chr(13), "")
    toto = Replace(toto, "¤", vbCrLf)
    Spreadsheet.Protection.Enabled = False
    Spreadsheet.Range("A2").ParseText _
    toto, sDelimiteur$
    
End If
FRM.Charger_Colection Spreadsheet, Mytype

Spreadsheet.AutoFilterMode = False
Spreadsheet.Range("a1").AutoFilter
On Error Resume Next
    Spreadsheet.Range("a1").CurrentRegion.AutoFitColumns
    Err.Clear
    Spreadsheet.Range("a1").CurrentRegion.Cells.EntireColumn.AutoFit
    Err.Clear
    On Error GoTo 0
    For I = 0 To Rs.Fields.Count - 1
   
       
        If Rs.Fields(I).Type = adBoolean Then
        
        DoEvents
            FrmApelan.ProgressBar1.Value = 0
            FrmApelan.ProgressBar1.Max = Spreadsheet.Cells(1, I + 1).CurrentRegion.Rows.Count
            For L = 2 To Spreadsheet.Cells(1, I + 1).CurrentRegion.Rows.Count
                IncremanteBarGrah FrmApelan
                 Spreadsheet.Cells(L, I + 1).Value = Replace(Spreadsheet.Cells(L, I + 1).Value, "'", "")
            Next
            Spreadsheet.Columns(I + 1).NumberFormat = "Yes/No"
        End If
         If InStr(UCase(Rs.Fields(I).Name), UCase("Prix Total")) <> 0 Then
           FrmApelan.ProgressBar1.Value = 0
            FrmApelan.ProgressBar1.Max = Spreadsheet.Cells(1, I + 1).CurrentRegion.Rows.Count
            For L = 2 To Spreadsheet.Cells(1, I + 1).CurrentRegion.Rows.Count
                IncremanteBarGrah FrmApelan
                
                Spreadsheet.Cells(L, I + 1).Formula = Replace(Spreadsheet.Cells(L, I + 1).Value, "'", "")
            Next
         End If
        
    Next
    FrmApelan.ProgressBar1.Value = 0
  If Save_ProgressBar1Max = 0 Then Save_ProgressBar1Max = Save_ProgressBar1Value + 1
 FrmApelan.ProgressBar1.Max = Save_ProgressBar1Max
  FrmApelan.ProgressBar1.Value = Save_ProgressBar1Value
'  Set Rs = Con.CloseRecordSet(Rs)

End Sub
Public Function SetAutocad()
On Error Resume Next
boolAutoCAD = True

Set AutoApp = CreateObject("AutoCAD.Application")
If Err = 0 Then
    AutoApp.Visible = False
    Example_AutoAudit
    AutoApp.Documents(0).Close False
    DoEvents
Else
    Err.Clear
    Set AutoApp = CreateObject("AutoCAD.AcadApplication.17")
    If Err = 0 Then
        AutoApp.Visible = False
        Example_AutoAudit
        AutoApp.Documents(0).Close False
        DoEvents
    Else
        Err.Clear
        MsgBox "Plus de licence Autocad disponible", vbInformation, "AutoCâble  licence :"
        boolAutoCAD = False
    End If

End If

End Function

Public Function GetAutocad()
On Error Resume Next
Set AutoApp = GetObject(, "autocad.application")
If Err Then
    Err.Clear
    Set AutoApp = GetObject(, "AutoCAD.AcadApplication.17")
End If
Err.Clear
End Function

