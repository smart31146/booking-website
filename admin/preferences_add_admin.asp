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
' *** Redirect if username exists
MM_flag="MM_insert"
If (CStr(Request(MM_flag)) <> "") Then
  MM_dupKeyRedirect="preferences_add_admin_exist.asp"
  MM_rsKeyConnection=Connect
  MM_dupKeyUsernameValue = CStr(Request.Form("admin_name"))
  MM_dupKeySQL="SELECT admin_name FROM admin WHERE admin_name='" & MM_dupKeyUsernameValue & "'"
  MM_adodbRecordset="ADODB.Recordset"
  set MM_rsKey=Server.CreateObject(MM_adodbRecordset)
  MM_rsKey.ActiveConnection=MM_rsKeyConnection
  MM_rsKey.Source=MM_dupKeySQL
  MM_rsKey.CursorType=0
  MM_rsKey.CursorLocation=2
  MM_rsKey.LockType=3
  MM_rsKey.Open
  If Not MM_rsKey.EOF Or Not MM_rsKey.BOF Then
    ' the username was found - can not add the requested username
    MM_qsChar = "?"
    If (InStr(1,MM_dupKeyRedirect,"?") >= 1) Then MM_qsChar = "&"
    MM_dupKeyRedirect = MM_dupKeyRedirect & MM_qsChar & "requsername=" & MM_dupKeyUsernameValue
    Response.Redirect(MM_dupKeyRedirect)
  End If
  MM_rsKey.Close
End If
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then

  MM_editConnection = Connect
  MM_editTable = "admin"
  MM_editRedirectUrl = "preferences.asp"
  MM_fieldsStr  = "admin_name|value|admin_pwd|value|admin_level|value|admin_active|value|admin_ip|value|admin_access|value|admin_info4|value"
  MM_columnsStr = "admin_name|',none,''|admin_pwd|',none,''|admin_level|',none,''|admin_active|none,1,0|admin_IP|',none,''|admin_change|',none,NULL|admin_info4|',none,''"

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
' *** Insert Record: construct a sql insert staatement and execute it

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert staatement
  MM_tableValues = ""
  MM_dbValues = ""
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
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End if
    MM_tableValues = MM_tableValues & MM_columns(i)
    MM_dbValues = MM_dbValues & FormVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    if Edit_OK = true then MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
	
	
	'update password
	Set obj = Server.CreateObject("ADODB.Recordset")
SQL="SELECT TOP 1 * FROM  admin  ORDER BY id_admin DESC"
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
Set uobj = Server.CreateObject("ADODB.Command")
SQL="update admin set admin_pwd='"&password&"' WHERE id_admin="&obj("id_admin")
uobj.ActiveConnection = Connect
uobj.CommandText = SQL
uobj.Execute
uobj.ActiveConnection.Close



obj.MoveNext
Loop
End If

obj.close  
	
	'end update
	
	
	
    call log_the_page ("BBG Execute - INSERT Administrator")
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If
  
  

End If
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
	business.Source = "SELECT *  FROM q_info4 WHERE id_info4="&admin_user.fields.item("id_info4")&"  ORDER BY info4"
end if
business.CursorType = 0
business.CursorLocation = 3
business.LockType = 3
business.Open()
business_numRows = 0

admin_user.close()
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Preferences - admin add. You are logged in as </TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
function trySubmit()
{
	if (document.forms[0].admin_name.value.length<4)
	{
		alert("Sorry, you must enter a name administrator's name!\n(min. 4 characters)");
		return false;
	}
	if (document.forms[0].admin_pwd.value.length<4)
	{
		alert("Sorry, you must enter a password!\n(min. 4 characters)");
		return false;
	}
	if (confirm("Are you sure you want to add this new administrator?"))
	{	document.forms[0].submit();
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
<BODY onLoad="change=false;" onUnload="<% call on_page_unload %>">
<table>
  <tr>
    <td align="left" valign="bottom" class="heading"> BBP add a new administrator</td>
  </tr>
  <tr>
    <td align="left" valign="bottom">
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="add_subject" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr>
            <td width="99%" >Administrator name</td>
          </tr>
          <tr>
            <td width="99%" >
              <input type="text" onChange="change=true;" name="admin_name" size="60" class="formitem1">
            </td>
          </tr>
          <tr>
            <td width="99%" >Password</td>
          </tr>
          <tr>
            <td width="99%" >
              <input type="text" onChange="change=true;" name="admin_pwd" size="60" class="formitem1">
            </td>
          </tr>
          <tr>
            <td width="99%" >Administration level</td>
          </tr>
          <tr>
            <td width="99%" >
              <select name="admin_level" class="formitem1">
                <option value="admin">Administrator</option>
                <option value="other">Reviewer</option>
              </select>
            </td>
          </tr>
          <tr>
            <td width="99%" >Company</td>
          </tr>
          <tr>
            <td width="99%" >
              <select name="admin_info4" class="formitem1">
			<% while not business.eof %>
				<option value="<%=business.fields.item("id_info4")%>"><%=business.fields.item("info4")%></option>
			<% business.movenext %>
			<% wend %>
              </select>
            </td>
          </tr>
          <tr>
            <td width="99%" >
              <input type="hidden" name="admin_active" value="1">
              <input type="hidden" name="admin_ip" value="<%=Request.ServerVariables("REMOTE_ADDR")%>">
              <input type="hidden" name="admin_access" value="<%=cDateSql(Now())%>">
            </td>
          </tr>
          <tr>
            <td width="99%" >
              <input type="reset" name="Submit2" value="Reset the form" class="quiz_button">
              <input type="submit" name="Submit" value="Insert this new administrator" class="quiz_button" <%call IsEditOK%>>
              or
              <input type="button" name="goback" value="Go back to preferences" class="quiz_button" onClick="document.location='preferences.asp'">
            </td>
          </tr>
        </table>
        <input type="hidden" name="MM_insert" value="true">
      </form>
    </td>
  </tr>
</table>
<p>
<p>&nbsp; </p>
</BODY>
</HTML>


<%
business.close()
call log_the_page ("BBG Add a new Administrator")
%>
