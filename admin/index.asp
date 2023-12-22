<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include.asp" -->
<!--#include file="sha256.asp"-->
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString<>"" Then MM_LoginAction = MM_LoginAction + "?" + Request.QueryString
MM_valUsername=CStr(Request.Form("username"))
If MM_valUsername <> "" Then
  MM_fldUserAuthorization="admin_level"
  MM_redirectLoginSuccess="frames.html"
  MM_redirectLoginFailed="index.asp?message=Wrong username orpassword, please try again."
  
  
	
	Dim password
	password=CStr(Request.Form("password"))
	Dim salt
	salt = MM_valUsername
	password=password&salt
	password=sha256(password)
  
  
  'MM_flag="ADODB.Recordset"
  'set MM_rsUser = Server.CreateObject(MM_flag)
  'MM_rsUser.ActiveConnection = Connect
  'MM_rsUser.Source = "SELECT * "
  'If MM_fldUserAuthorization <> "" Then MM_rsUser.Source = MM_rsUser.Source & "," & MM_fldUserAuthorization
 'MM_rsUser.Source = MM_rsUser.Source & " FROM admin WHERE admin_name='" & MM_valUsername &"' AND admin_pwd='" & password & "'"
  'MM_rsUser.CursorType = 0
  'MM_rsUser.CursorLocation = 3
  'MM_rsUser.LockType = 3
  'MM_rsUser.Open
   SQL= "SELECT TOP 1 * FROM admin WHERE (admin_name) =? and (admin_pwd) =?"
set objCommand = Server.CreateObject("ADODB.Command") 
objCommand.ActiveConnection = Connect
objCommand.CommandText = SQL
objCommand.Parameters(0).value = Replace(MM_valUsername, "'", "''")
objCommand.Parameters(1).value = password
Set MM_rsUser = objCommand.Execute()
  
  
  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then
    ' username and password match - this is a valid user
    Session("MM_Username_admin") = MM_valUsername
    Session("MM_id_admin") = MM_rsUser.Fields.Item("id_admin").Value
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization_admin") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization_admin") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And true Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If

		MM_editConnection = Connect
		MM_editTable = "admin"
		MM_editQuery = "update " & MM_editTable & " set admin_IP = '" & Request.ServerVariables("REMOTE_ADDR") & "', admin_change = '" & cDateSql(Now()) & "' WHERE ID_admin = " & (MM_rsUser.Fields.Item("ID_admin").Value) & ";"
		Set MM_editCmd = Server.CreateObject("ADODB.Command")
		MM_editCmd.ActiveConnection = MM_editConnection
		MM_editCmd.CommandText = MM_editQuery
		MM_editCmd.Execute
		MM_editCmd.ActiveConnection.Close

    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP - Administration log-in</TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
<script type="text/javascript" language="JavaScript">
		if (window.location.protocol != "https:")
    window.location.href = "https:" + window.location.href.substring(window.location.protocol.length);

var d = new String(window.location.host); 
var p = new String(window.location.pathname); 
var u = "http://" + d + p; 



if (u.indexOf("lawofthejungle.com.au") >= 0 || u.indexOf("lotj.com") >= 0 || u.indexOf("lawofthejungle.co.nz") >= 0 || u.indexOf("lawofthejungle.co.nz") >= 0 || u.indexOf("lawofthejungle.net") >= 0 || u.indexOf("lawofthejungle.net.au")  >= 0 || u.indexOf("lotj.co.uk") >= 0 || u.indexOf("lotj.co.nz") >= 0  )
{ 

d = "www.lawofthejungle.com";

if (window.location.protocol == "https:")
var u = "https://" + d + p;
else
var u = "http://" + d + p; 


window.location = u; 
} 


		
		 msieversion();
		function msieversion() {
		
            var ua = window.navigator.userAgent;
            var msie = ua.indexOf("MSIE ");

            //if (msie > 0)      // If Internet Explorer, return version number
             //   alert(parseInt(ua.substring(msie + 5, ua.indexOf(".", msie))));
           if (msie <= 0)                 // If another browser, return 0
                //alert('otherbrowser');
			//	window.location.replace('http://www.lawofthejungle.com/acme32/nobrowse/');

            return false;
        }
		</script>


</HEAD>
<BODY onLoad="document.forms[0].username.focus();">
<table>
  <tr>
    <td align="left" valign="top" height="30"><img src="images/loj.gif" width="227" height="33">
    </td>
  </tr>
</table>
<table>
  <tr>
    <td width="87" height="52">&nbsp;</td>
    <td height="52" align="left" valign="bottom" width="576" class="heading">
      BBP - administration interface login</td>
  </tr>
</table>
<table>
  <tr>
    <td width="87" height="52">&nbsp;</td>
    <td height="52" align="left" valign="bottom" width="576">
      <p>&nbsp;</p>
      <font color="#FF0000"><b><%= CStr(Request.QueryString("message"))%> </b></font>
      <p>Please log in with your administration username and password</p>
      <form name="login" method="post" action="<%=MM_LoginAction%>">
        <table>

		  <br><br><br>



		  <tr>
            <td >Username</td>
          </tr>

		  <tr>
            <td>
              <input type="text" name="username" value = "">
            </td>
          </tr>
          <tr>
            <td >Password</td>
          </tr>
          <tr>
            <td>
              <input type="password" name="password" value = "">
            </td>
          </tr>
          <tr>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td>
              <input type="submit" name="Submit" value="Log me in to the admin interface" class="quiz_button">
            </td>
          </tr>
        </table>
        <p>&nbsp;</p>
        <p>&nbsp; </p>
      </form>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>

</BODY>
</HTML>
<%
call log_the_page ("admin", "0", "n/a", "0", "n/a", "0", "n/a", "Login")
%>

