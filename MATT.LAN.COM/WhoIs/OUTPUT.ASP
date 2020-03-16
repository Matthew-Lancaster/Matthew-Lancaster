<h2 align="center">Matthewlan.com Stats - Who Is Online?</h2>
<BR>
<!--#include virtual="include.asp"--><center>

<table border="1">
<tr><td width="50%">Session ID</td><td>Looking where?</td><td>They got here when?</td></tr><!--Header for the table-->
<%
    For Each variable in Application.Contents <!--Plain english :)-->
    		if left(variable,4)="yoyo" then <!--Filter out only those with a "yoyo" to start-->
					if Application.Contents(variable)<>"" then%> <!--If it contains some information output it-->
<tr><td><%=mid(variable,5,999)%><!--Output all but the "yoyo"--></td><td><%=Application.Contents(variable)%><!--Output the rest (url/time they got here)</td></tr>
<%end if
Next%></table>
<!--Visit Beckerist.com :)-->
