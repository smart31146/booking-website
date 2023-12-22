<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->

<%
IF request.querystring("alt")="update" THEN
  MM_editConnection = Connect
  MM_editQuery = "UPDATE preferences SET welcome_note='"& request.form("welcome_note") &"' WHERE pref_active = 1"
'response.write ( MM_editQuery)
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
	MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
    call log_the_page ("Updated Welcome Message")
End If

set welcome = Server.CreateObject("ADODB.Recordset")
welcome.ActiveConnection = Connect
welcome.Source ="Select welcome_note FROM preferences WHERE pref_active = 1 ORDER BY preferences.pref_date DESC;"
welcome.CursorType = 0
welcome.CursorLocation = 3
welcome.LockType = 3
welcome.Open()
welcome_numRows = 0
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz new user. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">

<script language="JavaScript">
function exitpage()
{
	if (change==true)
	{
		if (confirm("You have changed at least one field on this page.\rBefore exiting this page, do you want to save those changes first?"))
		{
		return trySubmit();
		}
	}
	return true;
}
//-->
</script>
<script type="text/javascript" src="ckeditor/ckeditor.js?v=bbp34"></script>
</HEAD>
<BODY>
	<form action="?alt=update" METHOD="POST" name="welcome">
		<table>
			  <tr class="table_normal" >
				<td class="text" align="left" valign="top" width="120">Welcome Message:</td>
				<td class="text" align="left" valign="top" colspan="3">
				 <textarea name="welcome_note" ><%=welcome(0)%></textarea>
				 <script type="text/javascript">
				//<![CDATA[

					CKEDITOR.replace( 'welcome_note',
						{
						width: 500,
						height: 220
							});

				//]]>
				</script>
				</td>
			  </tr>
		</table>
		<input type="submit" name="Submit" value="Save" class="quiz_button" <%call IsEditOK%>>
		<input type="hidden" name="MM_insert" value="true">
    </form>
    </td>
  </tr>

</table>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
call log_the_page ("Edit Welcome Message")
%>


