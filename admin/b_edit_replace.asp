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
  MM_editTable = "toreplace"
  MM_editColumn = "id_replace"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "b_list_of_replace.asp"
  MM_fieldsStr  = "what|value|bywhat|value|active|value|bbg|value|training|value|quiz|value|search|value"
  MM_columnsStr = "repl_what|',none,''|repl_bywhat|',none,''|repl_active|none,1,0|repl_bbg|none,1,0|repl_tr|none,1,0|repl_q|none,1,0|repl_search|none,1,0"

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
    call log_the_page ("BBG Execute - UPDATE Replace: " & MM_recordId)	
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim replaceto__MMColParam
replaceto__MMColParam = "1"
if (Request.QueryString("id_replace") <> "") then replaceto__MMColParam = Request.QueryString("id_replace")
%>
<%
set replaceto = Server.CreateObject("ADODB.Recordset")
replaceto.ActiveConnection = Connect
replaceto.Source = "SELECT * FROM toreplace WHERE id_replace = " + Replace(replaceto__MMColParam, "'", "''") + ""
replaceto.CursorType = 0
replaceto.CursorLocation = 3
replaceto.LockType = 3
replaceto.Open()
replaceto_numRows = 0
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Replacements. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
function trySubmit()
{
	if (document.forms[0].what.value.length<2)
	{
		alert("Sorry, you must enter a string to be replaced!\n(min. 2 characters)");
		return false;
	}
	if (document.forms[0].bywhat.value.length<2)
	{
		alert("Sorry, you must enter a replacement string!\n(min. 2 characters)");
		return false;
	}
	if (confirm("Are you sure you want to update this replacement?"))	{	document.forms[0].submit();
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
    <td align="left" valign="bottom" class="heading"> Global replacements</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="add_subject" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr> 
            <td>String to be replaced</td>
          </tr>
          <tr> 
            <td> 
              <input type="text" name="what" onChange="change=true;" size="60" class="formitem1" value="<%=(replaceto.Fields.Item("repl_what").Value)%>">
            </td>
          </tr>
          <tr> 
            <td >Replacement of above string <a href="javascript:" onClick="MM_openBrWindow('_editor.asp?field=bywhat','editor','width=520,height=400')"><img src="images/editor.gif" width="24" height="10" border="0"></a> 
            </td>
          </tr>
          <tr> 
            <td> 
              <textarea name="bywhat" onChange="change=true;" cols="60" class="formitem1" rows="5"><%=(replaceto.Fields.Item("repl_bywhat").Value)%></textarea>
            </td>
          </tr>
          <tr> 
            <td>Replacement active? 
              <input onChange="change=true;" <%If (CStr(abs(replaceto.Fields.Item("repl_active").Value)) = CStr(1)) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="active" value="<%=(replaceto.Fields.Item("repl_active").Value)%>">
            </td>
          </tr>
          <tr> 
            <td>- effective for theBBG 
              <input onChange="change=true;" <%If (CStr(abs(replaceto.Fields.Item("repl_bbg").Value)) = CStr(1)) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="bbg" value="<%=(replaceto.Fields.Item("repl_bbg").Value)%>">
            </td>
          </tr>
          <tr> 
            <td>- effective for the Training 
              <input onChange="change=true;" <%If (CStr(abs(replaceto.Fields.Item("repl_tr").Value)) = CStr(1)) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="training" value="<%=(replaceto.Fields.Item("repl_tr").Value)%>">
            </td>
          </tr>
          <tr> 
            <td>- effective for the Quiz 
              <input onChange="change=true;" <%If (CStr(abs(replaceto.Fields.Item("repl_q").Value)) = CStr(1)) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="quiz" value="<%=(replaceto.Fields.Item("repl_q").Value)%>">
            </td>
          </tr>
          <tr> 
            <td>- effective for the BBG Search 
              page 
              <input onChange="change=true;" <%If (CStr(abs(replaceto.Fields.Item("repl_search").Value)) = CStr(1)) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="search" value="<%=(replaceto.Fields.Item("repl_search").Value)%>">
            </td>
          </tr>
          <tr> 
            <td width="99%" > 
              <input type="reset" name="Submit2" value="Reset the form" class="quiz_button">
              <input type="submit" name="Submit" value="Update this replacement" class="quiz_button" <%call IsEditOK%>>
              or 
              <input type="button" name="goback" value="Go back to replace list" class="quiz_button" onClick="document.location='b_list_of_replace.asp'">
            </td>
          </tr>
          <tr> 
            <td width="99%" >If you need to include a link to 
              a particular page within a BBG, use this <a href="javascript:" onClick="MM_openBrWindow('_link_generator.asp','linkgenerator','width=600,height=300')">LINK 
              GENERATOR </a></td>
          </tr>
        </table>
        <input type="hidden" name="MM_recordId" value="<%= replaceto.Fields.Item("id_replace").Value %>">
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
call log_the_page ("BBG Edit Replace: " & (replaceto.Fields.Item("id_replace").Value))
%>

<%
replaceto.Close()
%>
