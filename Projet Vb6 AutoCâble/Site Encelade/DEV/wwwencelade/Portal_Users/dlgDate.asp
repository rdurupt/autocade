<% 'Copyright 1999 Sam Hurdowar   sam_hurdowar@yahoo.com %>
<html>
<head>
	<title>Calendar</title>
</head>
<SCRIPT Language="Javascript">
<!--
function javIt(d){
	window.opener.document.frm.HireDate.value = d;
}	
function javCal(m,y){
	location.href = "dlgDate.asp?calMonth=" + m + "&calYear=" + y
}
//-->
</SCRIPT>
<body bgcolor="silver">
<center>
<%
Function GetLastDay(intMonthNum, intYearNum)
    Dim dNextStart
    If CInt(intMonthNum) = 12 Then
        dNextStart = CDate("1/1/" & intYearNum)
    Else
        dNextStart = CDate(intMonthNum + 1 & "/1/" & intYearNum)
    End If
    GetLastDay = Day(dNextStart - 1)
End Function
Sub Write_TD1(sValue, sColor)
    Response.Write "<TD BGCOLOR='" & sColor & "'> " & sValue & "</TD>" & vbCrLf
End Sub

Const cSUN = 1, cMON = 2, cTUE = 3, cWED = 4, cTHU = 5, cFRI = 6, cSAT = 7

intThisDay = Day(Date)
datToday = Date

If Request("calMonth") = "" Then
  intThisMonth = month(datToday)
Else
  intThisMonth = CInt(Request("calMonth"))
End If

If IsEmpty(Request("calYear")) Or Not IsNumeric(Request("calYear")) Then
  datToday = Date
  intThisYear = Year(datToday)
Else
  intThisYear = CInt(Request("calYear"))
End If

strMonthName = MonthName(intThisMonth)
datFirstDay = DateSerial(intThisYear, intThisMonth, 1)
intFirstWeekDay = WeekDay(datFirstDay, vbSunday)
intLastDay = GetLastDay(intThisMonth, intThisYear)

IntPrevMonth = intThisMonth - 1
If IntPrevMonth = 0 Then
    IntPrevMonth = 12
    intPrevYear = intThisYear - 1
Else
    intPrevYear = intThisYear
End If

IntNextMonth = intThisMonth + 1
If IntNextMonth > 12 Then
    IntNextMonth = 1
    intNextYear = intThisYear + 1
Else
    intNextYear = intThisYear
End If

LastMonthDate = GetLastDay(intLastMonth, intPrevYear) - intFirstWeekDay + 2
NextMonthDate = 1
intPrintDay = 1

dFirstDay = intThisMonth & "/1/" & intThisYear
dLastDay = intThisMonth & "/" & intLastDay & "/" & intThisYear
%>

<table bgcolor="gray" cellspacing="1">
<tr>
<td><a HREF="javascript:javCal(<%= IntPrevMonth %>,<%= intPrevYear %>)">&lt;</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
<form name="frmCal" action="dlgDate.asp">
<td colspan="5">
<select name="calMonth">
<% For i = 1 To 12 %>
    <% Mon = MonthName(i) %>
    <% If i = intThisMonth Then %>
        <option value="<%= i %>" selected><%= Mon %> 
    <% Else %>
        <option value="<%= i %>"><%= Mon %> 
    <% End If %>
<% Next %> 
</select>

<% a = Year(Date) - 10 %>
<% b = Year(Date) + 10 %>
&nbsp;<select name="calYear">
<% For i = a To b %>
    <% If i = intThisYear Then %>
        <option value="<%= i %>" selected><%= i %> 
    <% Else %>
        <option value="<%= i %>"><%= i %> 
    <% End If %>
<% Next %>
</select>
<input type="submit" name="subAction" value="Go!">
</td>

<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a HREF="javascript:javCal(<%= IntNextMonth %>,<%= intNextYear %>)">&gt;</a></td>
</tr>

<% '*****************************  Day Label  ********************* %>
<tr>
   <td ALIGN="LEFT" VALIGN="TOP">S</td>
   <td ALIGN="LEFT" VALIGN="TOP">M</td>
   <td ALIGN="LEFT" VALIGN="TOP">T</td>
   <td ALIGN="LEFT" VALIGN="TOP">W</td>
   <td ALIGN="LEFT" VALIGN="TOP">T</td>
   <td ALIGN="LEFT" VALIGN="TOP">F</td>
   <td ALIGN="LEFT" VALIGN="TOP">S</td>
</tr>

<% '*****************************  Days  ********************* %>
<% EndRows = False %>
<% Do While EndRows = False %>
    <TR>
<%
   For intLoopDay = cSUN To cSAT
        If intFirstWeekDay > cSUN Then
            Call Write_TD1(LastMonthDate, "SILVER")
            LastMonthDate = LastMonthDate + 1
            intFirstWeekDay = intFirstWeekDay - 1
        Else

            If intPrintDay > intLastDay Then
                Call Write_TD1(NextMonthDate, "SILVER")
                NextMonthDate = NextMonthDate + 1
                EndRows = True
            Else
                If intPrintDay = intLastDay Then
                    EndRows = True
                End If
				strDate = intThisMonth & "/" & intPrintDay & "/" & intThisYear
				strLink = "<a href=""javascript:javIt('" & strDate & "')"">" & intPrintDay & "</a>"
                Call Write_TD1(strLink, "#FEB4A9")
            End If

            intPrintDay = intPrintDay + 1
        End If
    
    Next 
	%>
    </TR>
<% Loop %>
</table>
<br>
<input type="button" value="Close" onClick="window.close();">
</form>
</center>
</body>
</html>
