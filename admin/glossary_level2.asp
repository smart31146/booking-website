<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
' *** Edit Operations: declare variables

MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>



<%
Dim glossary
If cInt(Request.QueryString("glossary") <> 0) Then 
glossary = cInt(Request.QueryString("glossary"))
Else 
Response.Redirect("error.asp?" & request.QueryString) 
End If
' *** Insert Record: construct a sql insert staatement and execute it

If (CStr(Request("MM_insert")) <> "") Then
dim glos
glos=Request("newglos")
 Set uobj = Server.CreateObject("ADODB.Command")
SQL="update glossary set description='"&glos&"' WHERE gid="+ Replace(glossary, "'", "''")
uobj.ActiveConnection = Connect
uobj.CommandText = SQL
uobj.Execute
uobj.ActiveConnection.Close

End If
%>
<%
numbers=1
%>

<%
set glossary2 = Server.CreateObject("ADODB.Recordset")
glossary2.ActiveConnection = Connect
glossary2.Source = "SELECT * FROM glossary where gid=" + Replace(glossary, "'", "''")
glossary2.CursorType = 0
glossary2.CursorLocation = 3
glossary2.LockType = 3
glossary2.Open()
glossary2_numRows = 0


Function Escape(sString)

    'Replace any Cr and Lf to <br>
    strReturn = Replace(sString , vbCrLf, "\n")     : 'visual basic carriage return line feed     ***********\
    strReturn = Replace(strReturn , vbCr , "\n")     : 'visual basic carriage return               *********These 3 are line breaks hence <BR>
    strReturn = Replace(strReturn , vbLf , "\n")     : 'visual basic line feed          ***********/
    strReturn = Replace(strReturn, "'", "''")          : 'Single quote changed to 2 single quotes ASP knows what to do
    Escape = strReturn
End Function

%>


<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: glossary site list. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
function trySubmit()
{
	if (document.forms[0].newbus.value.length<3)
	{
		alert("Sorry, you must enter a name for a new glossary site!\n(min. 3 characters)");
		return false;
	}
	if (confirm("Are you sure you want to add a new glossary site?"))	{	document.forms[0].submit();
	return false;
	}
return false;
}

//-->
</script>
</HEAD>
<BODY>
  
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> glossary site list</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form name="add_bus" method="POST" action="<%=MM_editAction%>" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">          <td class="text" width="10"><img src="../admin/images/back.gif" width="18" height="14"></td>
            <td class="text" colspan=2><a href="glossary_level1.asp">...go up one level 
              to list of Glossary words</a></td>
          </tr>
         <% 
While NOT glossary2.EOF
%>
<tr class="table_normal"> 
            <td class="text"><img src="images/new2.gif" width="11" height="13"></td>
            <td class="text" colspan=2> 
			<br>
			<strong>Description:</strong><br><br>
              <textarea  name="newglos"  cols="60" rows="10" class="formitem1"><%=Escape((glossary2.Fields.Item("description").Value))%></textarea> 
            </td>
          </tr>
<%
glossary2.MoveNext()
Wend
%>
          
          <tr> 
            <td class="text">&nbsp;</td>
            <td class="text"> 
          
              <input type="submit" name="Submit" value="updated description" class="quiz_button" <%call IsEditOK%>>
              <input type="hidden" name="id_info1" value="<%=glossary%>">
            </td>
          </tr>
        </table>
        <input type="hidden" name="MM_insert" value="true">
      </form>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
call log_the_page ("BBG Edit Info2: " & (glossary))
%>

<%
glossary2.Close()
%>


