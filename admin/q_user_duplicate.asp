<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/include_admin.asp" -->
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz duplicate username. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
</HEAD>
<BODY>
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> Quiz duplicate username</td>
  </tr>
  <tr> 
    <td class="text" align="left" valign="bottom"><br><br><br>The username that you have entered (<%=request.querystring("requsername")%>) is already in use. Please create a unique username.<br><br><br>
        <input type="button" name="redirect" value="Go back and try again" class="quiz_button" onClick="history.go(-1)">
      
    </td>
  </tr>
</table>
<p>&nbsp;</p></BODY>
</HTML>

<%
call log_the_page ("Quiz Duplicate User: " & (request.querystring("requsername")))
%>
