<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/include_admin.asp" -->
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Error. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
</HEAD>
<BODY>
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> Error!</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <p>&nbsp;</p>
      <table width="400" border="1" cellspacing="0" cellpadding="5" align="center">
        <tr> 
          <td align="center" bgcolor="#FF0000" > 
            <p><font color="#FFFFFF">Unfortunately, the Admin interface page you 
              are requesting is not available in the system any more or there 
              seems to be an error on this page.</font></p>
            <p><font color="#FFFFFF">If you get to this page through your bookmark 
              or through self generated URL, it is possible that your login session 
              had expired or that the URL structure has changed and page you are 
              requesting is available under a new URL. Please <a href="index.asp">login 
              and try again</a>.</font></p>
            <p><font color="#FFFFFF">If you are continuously getting this error 
              message while trying to reach some particular link anywhere within 
              Admin interface, please contact your BBP provider.</font></p>
            <p><font color="#FFFFFF">Please do not forget to copy following information:</font></p>
            <p><font color="#FFFFFF">Querry string: <%=Request.QueryString %><br>
              Administrator: <%=Session("MM_Username_admin")%><br>
              Referer: <%= Request.ServerVariables("HTTP_REFERER") %></font></p>
          </td>
        </tr>
      </table>
      <p>&nbsp;</p>
      </td>
  </tr>
</table>
</BODY>
</HTML>
<%
call log_the_page ("ERROR PAGE: " & Request.QueryString)
%>
