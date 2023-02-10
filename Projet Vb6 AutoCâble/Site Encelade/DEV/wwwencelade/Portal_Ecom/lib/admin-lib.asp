<!--include file="style.css"-->
<%
'variables globales
'	myDSN="DSN=casaviation;uid=casaviation;pwd=gyiodkbm"
	myDSN="Flyway"
	LOCAL_PATH = "E:\webprod\wwwflywayonline\portal_ecom\"
	HTTP_PATH = "http://portal.flywayonline.com/portal_ecom/"
	CRLF = chr(13) & chr(10)
	etoile="<font color=""#FF0000"">*</font>"

'fonctions
'--------------------------------------------------------------------------------------------------
sub barre()
	response.write ("<p><hr noshade size=1 width=""80%"">")
end sub

'--------------------------------------------------------------------------------------------------
function noquote(s0)
	noquote = replace(cstr(s0),"'","''")
end function

function noquote2(s0)
	s0 = s0
	if (cstr(s0) = "NULL") or isNull(s0) or (len(cstr(s0)) = 0) then
		noquote2 = "NULL"
	else
		noquote2 = "'" & noquote(s0) & "'"
	end if
end function

'--------------------------------------------------------------------------------------------------
function fmt(v)
	w=v
	if isNull(w) then
		fmt = "0.00"
	else
		w = round(cdbl(w),2)
		fmt = replace(cstr(w),",",".")
	end if
end function


function ff2euro(v)
	w=v
	if isNull(w) then
		ff2euro = "0.00"
	else
		ff2euro = round(cdbl(w)/6.55957,2)
	end if
end function

function Euro2ff(v)
	w=v
	if isNull(w) then
		Euro2ff = "0.00"
	else
		Euro2ff = round(cdbl(w)*6.55957,2)
	end if
end function


function inc(i)
	if i<>"" then
		inc = i + 1
	else
		inc = 1
	end if
end function

'--------------------------------------------------------------------------------------------------
function selected(valeur1,valeur2)
	ret = ""
	if (not isNull(valeur1)) and (not isNull(valeur2)) then
		if cstr(valeur1) <> cstr(valeur2) then
			ret = "     "
		else
			ret = " SELECTED "
		end if
	end if
	selected = ret
end function

'--------------------------------------------------------------------------------------------------
function checked(valeur)
	valeur = valeur
	if isNull (valeur) then
		valeur = ""
	end if
	ret = ""
	if cstr(valeur)="1" then
		ret = " CHECKED "
	else
		ret = "     "
	end if
	checked = ret
end function

'--------------------------------------------------------------------------------------------------
function toNull(i)
	j = i
	if isNull(j) then
		j = ""
	end if
	if cstr(j) <> "" then
		toNull = cstr(j)
	else
		toNull = "NULL"
	end if
end function

'--------------------------------------------------------------------------------------------------
function toZS(v)
	v = v		' Ah, Microsoft... le pire c'est que ca marche plus sans cette ligne !
	if isNull(v) then
		toZS = ""
	else
		toZS = cStr(v)
	end if
end function
'--------------------------------------------------------------------------------------------------
function ifNotNull (v, t)		' if (v<>NULL) then (return t) else (return NULL)
	if isNull(v) or (cstr(v) = "NULL") then
		ifNull = ""
	else
		ifNull = cstr(t)
	end if
end function

'--------------------------------------------------------------------------------------------------
function nl2br (s)
	nl2br = replace(cstr(s),CRLF,"<br>")
end function

'--------------------------------------------------------------------------------------------------
function br2nl (s)
	if not isNull(s) then
		br2nl = replace(cstr(s),"<br>",CRLF)
	else
		br2nl = ""
	end if
end function

'--------------------------------------------------------------------------------------------------
function ddmmyyyy(d)
	if not isNull(d) then
		dd = cstr(day(d))
		mm = cstr(month(d))
		yyyy = cstr(year(d))
		if len(dd)<2 then
			dd = "0" & dd
		end if
		if len(mm)<2 then
			mm = "0" & mm
		end if
		ddmmyyyy = dd & "/" & mm & "/" & yyyy
	else
		ddmmyyyy = null
	end if
end function


'--------------------------------------------------------------------------------------------------
function jsalert(s)
	response.write ("<script language=""JavaScript"">")
	response.write ("alert (""" & replace(s,"""","""""") & """);")
	response.write ("</script>")
end function
%>
