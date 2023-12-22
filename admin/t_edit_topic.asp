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
Dim topics__MMColParam
topics__MMColParam = "1"
if (Request.QueryString("topic") <> "") then topics__MMColParam = Request.QueryString("topic")
%>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = Connect
  MM_editTable = "tr_topics"
  MM_editColumn = "ID_topic"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "t_list_of_topics.asp?subj=" + request.querystring("subj")
  MM_fieldsStr  = "topic_name|value|topic_subject|value|active|value"
  MM_columnsStr = "topic_name|',none,''|topic_subject|none,none,NULL|topic_active|none,1,0"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(i+1) = CStr(Request.Form(MM_fields(i)))
  Next

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
    call log_the_page ("Training Execute - UPDATE Topic: " & MM_recordId)
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>

<%
set topics = Server.CreateObject("ADODB.Recordset")
topics.ActiveConnection = Connect
topics.Source = "SELECT *  FROM tr_topics  WHERE ID_topic = " + Replace(topics__MMColParam, "'", "''") + ""
topics.CursorType = 0
topics.CursorLocation = 3
topics.LockType = 3
topics.Open()
topics_numRows = 0
%>
<%
set subjects = Server.CreateObject("ADODB.Recordset")
subjects.ActiveConnection = Connect
subjects.Source = "SELECT subjects.ID_subject, subjects.subject_name FROM subjects GROUP BY subjects.ID_subject, subjects.subject_name;"
subjects.CursorType = 0
subjects.CursorLocation = 3
subjects.LockType = 3
subjects.Open()
subjects_numRows = 0
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Training topics. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
function trySubmit()
{
	if (document.forms[0].topic_name.value.length<3)
	{
		alert("Sorry, you must enter a name for a topic!\n(min. 3 characters)");
		return false;
	}
	if (confirm("Are you sure you want to update this topic?"))	{	document.forms[0].submit();
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
    <td align="left" valign="bottom" class="heading"> Training topic edit</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="add_subject" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr> 
            <td width="99%" > 
              <p>Name of the topic</p>
            </td>
          </tr>
          <tr> 
            <td width="99%" ><a href="t_list_of_topics.asp?subj="></a> 
              <input type="text" name="topic_name" onChange="change=true;" size="60" class="formitem1" value="<%=(topics.Fields.Item("topic_name").Value)%>">
            </td>
          </tr>
          <tr> 
            <td width="99%" >Name of the subject this topic 
              belongs to</td>
          </tr>
          <tr> 
            <td width="99%" > 
              <select name="topic_subject" onChange="change=true;" class="formitem1">
                <%
While (NOT subjects.EOF)
%>
                <option value="<%=(subjects.Fields.Item("ID_subject").Value)%>" <%if ((subjects.Fields.Item("ID_subject").Value) = (topics.Fields.Item("topic_subject").Value)) then Response.Write("SELECTED") : Response.Write("")%> ><%=(subjects.Fields.Item("subject_name").Value)%></option>
                <%
  subjects.MoveNext()
Wend
'If (subjects.CursorType > 0) Then
'  subjects.MoveFirst
'Else
  subjects.Requery
'End If
%>
              </select>
            </td>
          </tr>
          <tr> 
            <td width="99%" >Topic active? 
              <input onChange="change=true;" <%If (abs(topics.Fields.Item("topic_active").Value) = 1) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="active" value="checkbox">
            </td>
          </tr>
          <tr> 
            <td width="99%" >&nbsp; </td>
          </tr>
          <tr> 
            <td width="99%" > 
              <input type="reset" name="Submit2" value="Reset the form" class="quiz_button">
              <input type="submit" name="Submit" value="Update this topic" class="quiz_button" <%call IsEditOK%>>
              or 
              <input type="button" name="goback" value="Go back to topic list" class="quiz_button" onClick="document.location='t_list_of_topics.asp?subj=<%=(topics.Fields.Item("topic_subject").Value)%>'">
            </td>
          </tr>
        </table>
        <input type="hidden" name="MM_update" value="true">
        <input type="hidden" name="MM_recordId" value="<%= topics.Fields.Item("ID_topic").Value %>">
      </form>
    </td>
  </tr>
</table>
<p> 
<p>&nbsp; </p>
</BODY>
</HTML>

<%
call log_the_page ("Training Edit Topic: " & (topics.Fields.Item("ID_topic").Value))
%>

<%
topics.Close()
%>
<%
subjects.Close()
%>

