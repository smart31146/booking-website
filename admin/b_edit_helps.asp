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
' *** Update Record: set variables

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = Connect
  MM_editTable = "b_hlp"
  MM_editColumn = "ID_hlp"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "b_list_of_help.asp"
  MM_fieldsStr  = "help_name|value|help_tab|value"
  MM_columnsStr = "hlp_name|',none,''|hlp_tab|',none,''"

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
    call log_the_page ("BBG Execute - UPDATE Help: " & MM_recordId)
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim helps__MMColParam
helps__MMColParam = "1"
if (Request.QueryString("help")    <> "") then helps__MMColParam = Request.QueryString("help")   
%>
<%
set helps = Server.CreateObject("ADODB.Recordset")
helps.ActiveConnection = Connect
helps.Source = "SELECT *  FROM b_hlp  WHERE b_hlp.ID_hlp = " + Replace(helps__MMColParam, "'", "''") + " ;"
helps.CursorType = 0
helps.CursorLocation = 3
helps.LockType = 3
helps.Open()
helps_numRows = 0
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP Workplace ADMIN: Help tab editor. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
function trySubmit()
{
	if (document.forms[0].help_name.value.length<3)
	{
		alert("Sorry, you must enter a valid name for a Help tab!\n(min. 2 characters)");
		return false;
	}
	if (confirm("Are you sure you want to update this Help tab?"))	{	document.forms[0].submit();
	return false;
	}
return false;
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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
    <td align="left" valign="bottom" class="heading"> BBG Help tab editor</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="add_subject" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr> 
            <td width="99%" >Help tab name</td>
          </tr>
          <tr> 
            <td width="99%" ><a href="../admin/q_list_of_topics.asp?subj="></a> 
              <input type="text" onChange="change=true;" name="help_name" size="60" class="formitem1" value="<%=(helps.Fields.Item("hlp_name").Value)%>">
            </td>
          </tr>
          <tr> 
            <td width="99%" >Help tab content <a href="javascript:" onClick="MM_openBrWindow('_editor.asp?field=help_tab','editor','width=520,height=400')"><img src="images/editor.gif" width="24" height="10" border="0"></a> 
            </td>
          </tr>
          <tr> 
            <td width="99%" > 
              <textarea name="help_tab" onChange="change=true;" cols="80" class="formitem1" rows="15"><%=(helps.Fields.Item("hlp_tab").Value)%></textarea>
            </td>
          </tr>
          <tr> 
            <td width="99%" > 
              <input type="reset" name="Submit2" value="Reset the form" class="quiz_button">
              <input type="submit" name="Submit" value="Update this Help tab" class="quiz_button" <%call IsEditOK%>>
              or 
              <input type="button" name="goback" value="Go back to help list" class="quiz_button" onClick="document.location='b_list_of_help.asp'">
            </td>
          </tr>
        </table>
        <input type="hidden" name="MM_recordId" value="<%= helps.Fields.Item("ID_hlp").Value %>">
        <input type="hidden" name="MM_update" value="true">
      </form>
    </td>
  </tr>
</table>
<p> 
<p>&nbsp; </p>
</BODY>
</HTML>

<%
call log_the_page ("BBG Edit Help: " & (helps.Fields.Item("ID_hlp").Value))
%>

<%
helps.Close()
%>

