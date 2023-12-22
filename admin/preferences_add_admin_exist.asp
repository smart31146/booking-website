<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/include_admin.asp" -->
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Administrator uplicate username. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
</HEAD>
<BODY>
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> BBP administrator duplicate 
      username</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <p>&nbsp;</p>
      <p>&nbsp;</p>
      <p>Sorry, the username you want to use (<%=request.querystring("requsername")%>) is already in use.</p>
      <p>Please create a new one.</p>
      <p> 
        <input type="button" name="redirect" value="Go back and try again" class="quiz_button" onClick="history.go(-1)">
      </p>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>
<p>&nbsp;</p></BODY>
</HTML>

<%
call log_the_page ("BBG Duplicate Administrator: " & (request.querystring("requsername")))
%>
