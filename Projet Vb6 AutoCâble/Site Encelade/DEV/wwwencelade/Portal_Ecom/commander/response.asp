<!--#include file="../lib/admin-lib.asp"-->
<%
	if cstr(request("PAIEMENT"))="ANNULATION - RETOUR A LA BOUTIQUE" then
		if session("id_commande")<>"" then
		set conn=server.createobject("adodb.connection")
		conn.open myDSN
		sql = "update t_commande set etat_commande=4 where id_commande="&session("id_commande")
		conn.execute(sql)
		conn.close
		set conn = nothing		
		end if
		response.redirect "index.asp"
	end if
%>

<html>
<head>

</head>
<%

sub liste_vars()
 Dim sessitem
 response.write("<div class=""description"">Variables de formulaires:")
 For Each v in Request.form
  Response.write("<li>" & v & " : " & request(v) & "<BR>")
 next
 response.write("</div><p>")
 
end sub
call liste_vars


'	liste_vars()

	data = request("data")
	if isEmpty(data) then
		%>
<font class="erreur">Une Erreur s'est produite...</font>
		<%
	end if
	
	set sips = server.createobject("paiementsips400.SIPS")
	sips.name = "reponse"
	sips.pathfile = "c:/cyberplus/payment/parm/pathfile"
	sips.data = cstr(data)
	sips.reponse_sips
	reponse = sips.code_reponse
	

	if (reponse<>"00") then
		%>
		<font class="titre2">Votre commande a bien été passée.</font><p><%=barre%><p>
		<%
	else
		if (reponse="17") then
		%>
			<font class="titre2" color="#FF0000">Vous avez annulé votre commande.</font><p><%=barre%><p>
		<%
		else
		%>
		<font class="titre2" color="#FF0000">Votre commande a été refusée par votre organisme bancaire.</font><p><%=barre%><p>
		<%
		end if
	end if
	
	%>
	<script language="JavaScript">
	function reload_page()
	{
		top.location = "http://www.cas-aviation.fr";
	}
	setTimeout("reload_page()",2000);
	</script>
	<font class="titre1">Vous allez être redirigé vers la page de garde dans 2 secondes, ou vous pouvez cliquer <a href="http://www.cas-aviation.fr">ici</a></font>
<body>
</body>
</html>