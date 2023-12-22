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
  MM_editTable = "subjects"
  MM_editColumn = "ID_subject"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "t_list_of_subjects.asp"
  MM_fieldsStr  = "subject_name|value|subject_banner|value|monkey_file|value|subject_comment|value|active|value"
  MM_columnsStr = "subject_name|',none,''|subject_banner|',none,''|subject_monkey|',none,''|subject_comment|',none,''|subject_active_t|none,1,0"

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
    call log_the_page ("Training Execute - UPDATE Subject: " & MM_recordId)
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim subjects__MMColParam
subjects__MMColParam = "1"
if (Request.QueryString("subj") <> "") then subjects__MMColParam = Request.QueryString("subj")
%>
<%
set subjects = Server.CreateObject("ADODB.Recordset")
subjects.ActiveConnection = Connect
subjects.Source = "SELECT *  FROM subjects  WHERE subjects.ID_subject = " + Replace(subjects__MMColParam, "'", "''") + " ;"
subjects.CursorType = 0
subjects.CursorLocation = 3
subjects.LockType = 3
subjects.Open()
subjects_numRows = 0
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Training subject. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
function trySubmit()
{
	if (document.forms[0].subject_name.value.length<2)
	{
		alert("Sorry, you must enter a name for a subject!\n(min. 2 characters)");
		return false;
	}
	if (confirm("Are you sure you want to update this subject?"))	{	document.forms[0].submit();
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
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</HEAD>
<BODY onLoad="change=false;" onUnload="<% call on_page_unload %>">
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> Training subject edit</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="add_subject" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr> 
            <td width="99%" >Name of the subject</td>
          </tr>
          <tr> 
            <td width="99%" ><a href="../admin/q_list_of_topics.asp?subj="></a> 
              <input type="text" name="subject_name" onChange="change=true;" size="60" class="formitem1" value="<%=(subjects.Fields.Item("subject_name").Value)%>">
            </td>
          </tr>
          <tr> 
            <td width="99%" >Name of the file with subject training 
              banner </td>
          </tr>
          <tr> 
            <td  align="left" valign="bottom"> 
              <input type="text" name="subject_banner" onChange="change=true;" size="60" class="formitem1" value="<%=lCase(subjects.Fields.Item("subject_banner").Value)%>">
              <%
if (subjects.Fields.Item("subject_banner").Value) <> "" Then			  
	if fileexist("../client/training_labels/"& lCase(subjects.Fields.Item("subject_banner").Value)) = True Then
	Response.Write " <img src='../admin/images/ok.gif'> "
	Else
	Response.Write " <img src='../admin/images/miss.gif'> "
	End If 
Else
	Response.Write " <img src='../admin/images/no.gif'> "
End If
%>
              <%If (subjects.Fields.Item("subject_banner").Value) <> "" Then %>
              <a href="javascript:"><img src="images/eye.gif" width="16" height="16" border="0" onClick="MM_openBrWindow('../client/training_labels/<%=lCase(subjects.Fields.Item("subject_banner").Value)%>','preview','resizable=yes,width=200,height=200')"></a> 
              <%End If %>
              <a href="javascript:"  onClick="MM_openBrWindow('_banner_browse.asp','icobrowse','scrollbars=yes,width=400,height=400')"><img src="images/search.gif" width="16" height="16" border="0"></a> 
            </td>
          </tr>
          <tr> 
            <td width="99%" >Name of the file with subject training 
              monkey </td>
          </tr>
          <tr> 
            <td  align="left" valign="bottom"> 
              <input type="text" name="monkey_file" onChange="change=true;" size="60" class="formitem1" value="<%=lCase(subjects.Fields.Item("subject_monkey").Value)%>">
              <%
If (subjects.Fields.Item("subject_monkey").Value) <> "" then			  
	if fileexist("../client/training_monkeys/"& lCase(subjects.Fields.Item("subject_monkey").Value)) = True Then
	Response.Write " <img src='../admin/images/ok.gif'> "
	Else
	Response.Write " <img src='../admin/images/miss.gif'> "
	End If 
Else
	Response.Write " <img src='../admin/images/no.gif'> "
End If
%>
              <%if (subjects.Fields.Item("subject_monkey").Value) <> "" Then %>
              <a href="javascript:"><img src="images/eye.gif" width="16" height="16" border="0" onClick="MM_openBrWindow('../client/training_monkeys/<%=lCase(subjects.Fields.Item("subject_monkey").Value)%>','preview','resizable=yes,width=200,height=200')"></a> 
              <%end if %>
              <a href="javascript:"  onClick="MM_openBrWindow('_monkey_browse.asp','icobrowse','scrollbars=yes,width=550,height=400')"><img src="images/search.gif" width="16" height="16" border="0"></a> 
            </td>
          </tr>
          <tr> 
            <td  align="left" valign="bottom">Comment</td>
          </tr>
          <tr> 
            <td  align="left" valign="bottom"> 
              <input type="text" name="subject_comment" onChange="change=true;" size="60" class="formitem1" value="<%=(subjects.Fields.Item("subject_comment").Value)%>">
            </td>
          </tr>
          <tr> 
            <td width="99%" >Subject active? 
              <input onChange="change=true;" <%If (abs(subjects.Fields.Item("subject_active_t").Value) = 1) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="active" value="checkbox">
            </td>
          </tr>
          <tr> 
            <td width="99%" >&nbsp; </td>
          </tr>
          <tr> 
            <td width="99%" > 
              <input type="reset" name="Submit2" value="Reset the form" class="quiz_button">
              <input type="submit" name="Submit" value="Update this subject" class="quiz_button" <%call IsEditOK%>>
              or 
              <input type="button" name="goback" value="Go back to subject list" class="quiz_button" onClick="document.location='t_list_of_subjects.asp'">
            </td>
          </tr>
        </table>
        <input type="hidden" name="MM_update" value="true">
        <input type="hidden" name="MM_recordId" value="<%= subjects.Fields.Item("ID_subject").Value %>">
      </form>
    </td>
  </tr>
</table>
<p> 
<p>&nbsp; </p>
</BODY>
</HTML>

<%
call log_the_page ("Training Edit Subject: " & (subjects.Fields.Item("ID_subject").Value))
%>

<%
subjects.Close()
%>
