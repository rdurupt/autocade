<% response.addheader "Pragma","no-cache" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel='stylesheet' href='<%=session("PortalPath")%>Portal_Styles/euxiaadmin.css'>
<!--#include file="admin-lib.asp"-->
</HEAD>

<BODY>
<div align="center">
<center>
<%	

if Session("candidat_ADOContact") & "§§" = "§§" Then
 Session("candidat_ADOContact") = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=e:\webprod\dbsportail\wwwextranetencelade\Encelade_Menu.mdb"  
end if



Sql = "SELECT Path FROM BaseDefault"
Set Conn1 = Server.Createobject("adodb.connection")
Conn1.open Session("candidat_ADOContact")
Set Rs = Server.Createobject("adodb.recordset")
Set Rs = Conn1.Execute(Sql)
sPath = Rs("Path")
Rs.Close
Conn1.Close
Set Conn1 = Nothing

Set MyPortal = Server.CreateObject(Session("PortalObject") & ".Portal")

	set conn=server.createobject("adodb.connection")
	'conn.open myDSN
	conn.open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & sPath

currency_code =978	
	
dev = Request("dev")
id_commande = Request("id_commande")
select case dev
	case "update"

		sql = "SELECT userlogin FROM dbp_users WHERE userid = " & session("userid")
		set rs = conn.Execute(sql)
		sql = "update t_commande set etat_commande=" & etat & ", id_adm_user="
		if (rs.eof and rs.bof) then
			sql = sql & "0 "
		else
			sql = sql & session("userid") 
		end if
		if (request("f_card_type") = "chèque" or  request("f_card_type")= "cheque") and etat=2  then
			sql = sql & ", payment_date='" & now() & "'"
		end if
		sql = sql & " where id_commande=" & id_commande
		
		'response.write sql & "<br>"
		' response.end
		rs.close
		set rs = nothing
		'response.write sql&"<br>" 
		'response.end
		conn.execute(sql)
		etat=1
end select


jour = Request("jour")
anneemois = Request("anneemois")
card_type = request("card_type")

response.write("<br><br><h2>" & MyPortal.translate("Tous les devis") & "</h2>")

%>
<form action="commande.asp" method="post">
<% = MyPortal.Translate("Date") %> : 
<select name="anneemois">
	<option value="all"><% = MyPortal.Translate("toutes") %>
<% 
  	sql = "select distinct format(creation,'yyyy/mm') As Danneemois"
	sql = sql & " from t_devis "
	set cur = conn.execute(sql)
	
	while not cur.eof
    	response.write "<option value=""" & cur("Danneemois") & """" 
    	if anneemois=cur("Danneemois") then 
    		response.write " selected" 
    	end if 
    	response.write ">" & cur("Danneemois") & CRLF
	cur.movenext
	wend

%>
</select>
<%if anneemois<>"" and anneemois<>"all" then %>
<select name="jour">
<%	
	sql = "select distinct format(creation,'dd') As Djour "
	sql = sql & " from t_devis "
	sql = sql & " where format(creation,'yyyy/mm')='"& anneemois &"'"
	set cur = conn.execute(sql)
	while not cur.eof
    	response.write "<option value=""" & cur(0) &"""" 
    	if jour=cur(0) then 
    		response.write " selected" 
    	end if 
    	response.write ">"&cur(0)&CRLF
    	cur.movenext
	wend %> 
</select>
<%end if %>
<input type="Submit" name="envoi" value="ok">
</form>
<%if  dev<>"view" and anneemois<>"" then %>
<%if anneemois="all" or jour<>"" then%>
	<center>
	<table border=0 width="90%"><tr><td>
	<table border=1 width="100%">
	<!--  class="header" bgcolor="#808080" -->
	<tr>
		<td class="TrEntete"><% = MyPortal.Translate("Num devis") %></td>
		<td class="TrEntete"><% = MyPortal.Translate("Date") %></td>
		<td class="TrEntete"><% = MyPortal.Translate("Client") %></td>
		<td class="TrEntete"><% = MyPortal.Translate("Total") %></td>
	</tr>
<%
	
	sql = "select distinct cd.numdevis,"
	sql = sql & " cd.id_r_social,"
	sql = sql & " format(creation,'dd/mm/yyyy') As dDay,"
	sql = sql & " format(creation,'dd') As Djour,"
	sql = sql & " c.FirstName + ' ' + c.LastName As UserDisplayName,"
	sql = sql & " SUM(qte_produit*prixu_produit) As total"
	sql = sql & "  from t_devis cd, t_r_social c where " 
	sql = sql & " cd.id_r_social = c.userid "
	if anneemois<>"all" then
		sql = sql & " and format(creation,'yyyy/mm') ='"& anneemois &"'"
		if jour <> "" then
			sql = sql & " and format(creation,'dd')='"& jour &"'"
		end if
	end if
	sql = sql &" group by cd.numdevis, cd.id_r_social, format(creation,'dd/mm/yyyy'), format(creation,'dd'), c.FirstName + ' ' + c.LastName"
	sql = sql &" order by format(creation,'dd/mm/yyyy') desc"
	set cur = conn.execute(SQL)
	while not cur.eof

	%>
	<tr><td><a href="commande.asp?cc=<%'=currency_code%>&dev=view&id_commande=<%=cur("numdevis")%>&anneemois=<%=anneemois%>&jour=<%=cur("Djour")%>"><%=cur("numdevis")%></a></td><%
	response.write "<td>" & cur("dDay") & "</td>" & CRLF
	response.write "<td>" & cur("UserDisplayName") & "</td>" & CRLF
	if isNull(currency_code) then
		currency_code = 978
	end if
	response.write "<td>" & formatNumber(cur("total"),2) & "</td>" & CRLF
	cur.movenext
	wend
	
%>
</table>
</td></tr></table><%end if %>
<%end if %>
<%if dev = "view" then
id_commande = Request("id_commande")
anneemois = Request("anneemois")
jour = Request("jour")

total = 0.00
i_poids  = 0.00


sql = "SELECT t_devis.typepiece,"
sql = sql & " t_devis.id_produit,"
sql = sql & " t_devis.refproduit,"
sql = sql & " t_devis.qte_produit,"
sql = sql & " t_devis.prixu_produit,"
sql = sql & " (t_devis.prixu_produit * t_devis.qte_produit) as total,"
sql = sql & " t_devis.designation"
sql = sql & " FROM t_devis"
sql = sql & " WHERE numdevis = '" & id_commande & "'"

	set cur = conn.execute(SQL)
	%>
	<font class="titre1">Devis <%=id_commande %> :<br><br>
	<center>
	<table border=0 width="80%"><tr><td>
	<table border=1 width="100%">
	<!--  class="header" bgcolor="#808080" -->
	<tr nowrap>
	<th class="TrEntete"><% = MyPortal.Translate("Type Pièce") %></th>
	<th class="TrEntete"><% = MyPortal.Translate("Réference") %></th>
	<th class="TrEntete"><% = MyPortal.Translate("Produit") %></th>
	<th class="TrEntete"><% = MyPortal.Translate("Prix Unitaire") %></th>
	<th class="TrEntete"><% = MyPortal.Translate("Qté") %></th>
	<th class="TrEntete"><% = MyPortal.Translate("Prix Total") %></th>
	</tr>
	<%
	if not cur.eof then
	leTotal = 0
	leQte = 0
	While Not cur.EOF
            leStr = leStr & "<tr class=""smalltext"" width=""30""> "
            leStr = leStr & "<td>" & cur("typepiece") & "</td>"
            leStr = leStr & "<td>" & cur("designation") & "</td>"
            leStr = leStr & "<td>" & cur("refproduit") & "</td>"
            leStr = leStr & "<td align=""center"">" & cur("prixu_produit") & "</td>"
            leStr = leStr & "<td align=""center"">" & cur("qte_produit") & "</td>"
            leStr = leStr & "<td align=""center"">" & cur("total") & "</td>"
            leTotal = leTotal + cur("total")
            leQte = leQte + cur("qte_produit")
            cur.MoveNext
        Wend
        TDbg = 0
else
	leStr = leStr & "<tr><td colspan=4>Erreur : aucun produit trouvé !</td></tr>" 
end if
            leStr = leStr & "<tr class=""TrEntete"" height=""25"">"
            leStr = leStr & "<td colspan=""4"" class=""TextTrEntete"" align=""right"">" & MyPortal.translate("Total TTC") & " (€)&nbsp;:</td>"
            leStr = leStr & "<td width=""20%"" align=""center"" class=""TextTrEntete"">" & leQte & "</td>"
            leStr = leStr & "<td width=""20%"" align=""center"" class=""TextTrEntete"">" & leTotal & "</td></tr>"
			
response.write (leStr)

sql = "select distinct c.LastName, "
sql = sql & " c.FirstName, "
sql = sql & " c.address, "
sql = sql & " c.zip, "
sql = sql & " c.city, "
sql = sql & " p.label_pays, "
sql = sql & " c.id_livraison, "
sql = sql & " c.email "
sql = sql & " from	t_r_social c, t_pays p, t_devis cd " 
sql = sql & " where	cd.numdevis = '" & id_commande & "'"
sql = sql & " and	cd.id_r_social = c.userid " 
sql = sql & " and	p.id_pays = c.id_pays"

set cur=conn.execute(sql)

cf = cur(0) & " " & cur(1) & ",<br>" & cur(2) & ", " & cur(3) & " " & cur(4) & " - " & cur(5)
email = cur(7)
if isNull(email) then
	email = "[Aucun]"
else
	email = "<a href=""mailto:" & cstr(cur(7)) & """>" & cstr(cur(7)) & "</a>"
end if

cf = cf & "<br>" & email

Response.Write cf & "<br><br>"

%>
</table>
</td></tr></table>
<form action="commande.asp" method="post">
<input type="hidden" name="id_commande" value="<%=id_commande%>">
<input type="hidden" name="anneemois" value="<%=server.htmlencode(anneemois)%>">
<input type="hidden" name="jour" value="<%=server.htmlencode(anneemois)%>">
<input type="hidden" name="dev" value="update">
</form>
<br><br>
<a href="commande.asp?id_commande=<%=id_commande %>&anneemois=<%=anneemois %>&jour=<%=jour %>&etat=<%=etat%>"><% = MyPortal.Translate("retour") %></a>
<%end if %>
<%
	conn.close()
	set conn = nothing
	Set MyPortal = nothing
%>
</center></div>
</BODY>
</HTML>
