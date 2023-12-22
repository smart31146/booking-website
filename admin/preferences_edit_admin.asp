<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<!--#include file="sha256.asp"-->
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
' *** Update Record: set variables

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = Connect
  MM_editTable = "admin"
  MM_editColumn = "id_admin"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "preferences.asp"
  MM_fieldsStr  = "admin_pwd|value|admin_level|value|admin_active|value|admin_info4|value"
  MM_columnsStr = "admin_pwd|',none,''|admin_level|',none,''|admin_active|none,1,0|admin_info4|',none,none"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")

  ' set the form values
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(i+1) = CStr(Request.Form(MM_fields(i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Update Record: construct a sql update staatement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update staatement
  MM_editQuery = "update " & MM_editTable & " set "
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    FormVal = MM_fields(i+1)
    MM_typeArray = Split(MM_columns(i+1),",")
    Delim = MM_typeArray(0)
    If (Delim = "none") Then Delim = ""
    AltVal = MM_typeArray(1)
    If (AltVal = "none") Then AltVal = ""
    EmptyVal = MM_typeArray(2)
    If (EmptyVal = "none") Then EmptyVal = ""
    If (FormVal = "") Then
      FormVal = EmptyVal
    Else
      If (AltVal <> "") Then
        FormVal = AltVal
      ElseIf (Delim = "'") Then  ' escape quotes
        FormVal = "'" & Replace(FormVal,"'","''") & "'"
      Else
        FormVal = Delim + FormVal + Delim
      End If
    End If
    If (i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(i) & " = " & FormVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    if Edit_OK = true then MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
	
	'update password
	Set obj = Server.CreateObject("ADODB.Recordset")
SQL="SELECT  * FROM  admin  where id_admin="&Request("MM_recordId")
obj.ActiveConnection = Connect
obj.Source = SQL 
obj.CursorType = 0
obj.CursorLocation = 3
obj.LockType = 3
obj.Open

If obj.EOF then
Response.write("The END")

Else

Do While Not obj.EOF
Dim salt
salt = obj("admin_name")
password=obj("admin_pwd")
password=password&salt
password=sha256(password)
If  Len(Request("admin_pwd"))<20 Then
Set uobj = Server.CreateObject("ADODB.Command")
SQL="update admin set admin_pwd='"&password&"' WHERE id_admin="&obj("id_admin")
uobj.ActiveConnection = Connect
uobj.CommandText = SQL
uobj.Execute
uobj.ActiveConnection.Close



end IF






obj.MoveNext
Loop
End If

obj.close  
	

	
	'end update
	
	
	
	
	
    call log_the_page ("BBG Execute - UPDATE Administrator: " & MM_recordId)
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim admins__MMColParam
admins__MMColParam = "1"
if (Request.QueryString("admin_id") <> "") then admins__MMColParam = Request.QueryString("admin_id")
%>
<%
set admins = Server.CreateObject("ADODB.Recordset")
admins.ActiveConnection = Connect
admins.Source = "SELECT * FROM admin inner join q_info4 on admin.admin_info4=q_info4.id_info4 WHERE id_admin = " + Replace(admins__MMColParam, "'", "''") + ""
admins.CursorType = 0
admins.CursorLocation = 3
admins.LockType = 3
admins.Open()
admins_numRows = 0
%>
<%
set admin_user = Server.CreateObject("ADODB.Recordset")
admin_user.ActiveConnection = Connect
admin_user.Source = "SELECT * FROM admin inner join q_info4 on admin.admin_info4=q_info4.id_info4 where admin.id_admin="&Session("MM_id_admin")&""
admin_user.CursorType = 0
admin_user.CursorLocation = 3
admin_user.LockType = 3
admin_user.Open()

set business = Server.CreateObject("ADODB.Recordset")
business.ActiveConnection = Connect
if admin_user.fields.item("info4_viewall")=1 then
	business.Source = "SELECT *  FROM q_info4   ORDER BY info4"
else
	business.Source = "SELECT *  FROM q_info4 WHERE id_info4="&admins.fields.item("id_info4")&"  ORDER BY info4"
end if
business.CursorType = 0
business.CursorLocation = 3
business.LockType = 3
business.Open()
business_numRows = 0
%>

<%
admin_edit_ok = false
if (lCase(admins.Fields.Item("admin_name").Value) = "administrator") then
	if (lCase(admins.Fields.Item("admin_name").Value) = admin_logged_in) then admin_edit_ok = true
else
	if (Edit_OK) or (lCase(admins.Fields.Item("admin_name").Value) = admin_logged_in) then admin_edit_ok = true
end if
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Preferences - admin edit. You are logged in as </TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
function trySubmit()
{
	if (document.forms[0].admin_pwd.value.length<4)
	{
		alert("Sorry, you must enter a password!\n(min. 4 characters)");
		return false;
	}
	if (confirm("Are you sure you want to update this administrator?"))	{	document.forms[0].submit();
	return false;
	}
return false;
}

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
</HEAD>
<BODY onLoad="change=false;" onUnload="<% if admin_edit_OK then response.write("return exitpage(); ") %>">
<table>
  <tr>
    <td align="left" valign="bottom" class="heading"> BBP administrator edit</td>
  </tr>
  <tr>
    <td align="left" valign="bottom">
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="add_subject" onSubmit="<% if admin_edit_OK then response.write("change=false; return trySubmit(0); ") else response.write("alert('"& admin_error_message &"'); return false; ") end if %>" onReset="<%call on_form_Reset%>">
        <table>
          <tr>
            <td width="99%" >Administrator name</td>
          </tr>
          <tr>
            <td width="99%" >
              <input type="text" onChange="change=true;" name="admin_name" size="60" class="formitem1" disabled="true" value="<%=(admins.Fields.Item("admin_name").Value)%>">
            </td>
          </tr>
          <tr>
            <td width="99%" >Password</td>
          </tr>
          <tr>
            <td width="99%" >
              <input type="text" onChange="change=true;" name="admin_pwd" size="60" <% if admin_user.fields.item("info4_viewall")=0 AND admin_user.fields.item("id_info4")<>admins.Fields.Item("id_info4").Value then%>disabled="true"<% end if %> class="formitem1" value="<%if admin_edit_ok then response.write(admins.Fields.Item("admin_pwd").Value)%>">
            </td>
          </tr>
          <tr>
            <td width="99%" >Administration level</td>
          </tr>
          <tr>
            <td width="99%" >
              <select <% if admin_user.fields.item("info4_viewall")=0 AND admin_user.fields.item("id_info4")<>admins.Fields.Item("id_info4").Value then%>disabled<% end if %> name="admin_level" <%if (lCase(admins.Fields.Item("admin_name").Value) = "administrator") or (NOT Edit_OK) then response.write("disabled='true'")%> class="formitem1">
                <option value="admin" <%if lCase(admins.Fields.Item("admin_level").Value) = "admin" then response.write("selected")%>>Administrator</option>
                <option value="other" <%if lCase(admins.Fields.Item("admin_level").Value) = "other" then response.write("selected")%>>Reviewer</option>
              </select>
              <%if (lCase(admins.Fields.Item("admin_name").Value) = "administrator") or (NOT Edit_OK) then response.write("<input type='hidden' name='admin_level' value='"& lCase(admins.Fields.Item("admin_level").Value) &"'>")%>
            </td>
          </tr>
          <tr>
            <td width="99%" >Company</td>
          </tr>
          <tr>
            <td width="99%" >
              <select name="admin_info4" class="formitem1" <% if admin_user.fields.item("info4_viewall")=0  AND admin_user.fields.item("id_info4")<>admins.Fields.Item("id_info4").Value then%>disabled<% end if%>>
			<% while not business.eof %>
				<option value="<%=business.fields.item("id_info4")%>" <% if admins.fields.item("admin_info4")=business.fields.item("id_info4") then %> SELECTED <% end if %>><%=business.fields.item("info4")%></option>
			<% business.movenext %>
			<% wend %>
              </select>
            </td>
          </tr>
          <tr>
            <td width="99%" >Last accessed</td>
          </tr>
          <tr>
            <td width="99%" >
              <input type="text" onChange="change=true;" name="admin_access" size="60" class="formitem1" disabled="true" value="<%=cDateSQL(admins.Fields.Item("admin_change").Value)%>">
            </td>
          </tr>
          <tr>
            <td width="99%" >Accessed from</td>
          </tr>
          <tr>
            <td width="99%" >
              <input type="text" onChange="change=true;" name="admin_ip" size="60" class="formitem1" disabled="true" value="<%=(admins.Fields.Item("admin_IP").Value)%>">
            </td>
          </tr>
          <tr>
            <td width="99%" >Account active?
              <input <% if admin_user.fields.item("info4_viewall")=0 AND admin_user.fields.item("id_info4")<>admins.Fields.Item("id_info4").Value then%>disabled="true"<% end if %>  <%If (abs(admins.Fields.Item("admin_active").Value) = 1) Then Response.Write("CHECKED")%> type="checkbox" name="admin_active" value="checkbox" <%if (lCase(admins.Fields.Item("admin_name").Value) = "administrator") or (NOT Edit_OK) then response.write("disabled='true'")%>>
              <%if (lCase(admins.Fields.Item("admin_name").Value) = "administrator") or (NOT Edit_OK) then response.write("<input type='hidden' name='admin_active' value='"& abs(admins.Fields.Item("admin_active").Value) &"'>")%>
            </td>
          </tr>
          <tr>
            <td width="99%" >&nbsp; </td>
          </tr>
          <tr>
            <td width="99%" >
              <input type="reset" name="Submit2" value="Reset the form" class="quiz_button">
              <input type="submit" name="Submit" value="Update this administrator" class="quiz_button" <% if admin_edit_ok AND (admin_user.fields.item("info4_viewall")=1 OR admin_user.fields.item("id_info4")=admins.Fields.Item("id_info4").Value) then response.write("") else response.write("disabled='true'")%>>
              or
              <input type="button" name="goback" value="Go back to preferences" class="quiz_button" onClick="document.location='preferences.asp'">
            </td>
          </tr>
        </table>
        <input type="hidden" name="MM_update" value="true">
        <input type="hidden" name="MM_recordId" value="<%= admins.Fields.Item("id_admin").Value %>">
      </form>
    </td>
  </tr>
</table>
<p>
<p>&nbsp; </p>
</BODY>
</HTML>

<%
call log_the_page ("BBG Edit Administrator: " & (admins.Fields.Item("id_admin").Value))
%>

<%
admins.Close()
admin_user.close()
%>
