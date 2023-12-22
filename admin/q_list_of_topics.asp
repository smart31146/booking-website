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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then

  MM_editConnection = Connect
  MM_editTable = "q_topics"
  MM_editRedirectUrl = "q_list_of_topics.asp"
  MM_fieldsStr  = "newtopic|value|id_subject|value|topic_active|value|topic_ord|value|UID|value"
  MM_columnsStr = "topic_name|',none,''|topic_subject|none,none,NULL|topic_active|none,none,NULL|topic_ord|none,none,NULL|topic_UID|',none,''"

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
    call log_the_page ("Quiz Execute - INSERT Topic")
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
numbers=1
%>
<%
Dim subj
If (Request.QueryString("subj") <> "") Then 
subj = cInt(Request.QueryString("subj"))
Else 
Response.Redirect("error.asp?" & request.QueryString) 
End If
%>
<%
set topics = Server.CreateObject("ADODB.Recordset")
topics.ActiveConnection = Connect
topics.Source = "SELECT q_topics.ID_topic, q_topics.topic_name, q_topics.topic_subject, q_topics.topic_active  FROM q_topics  GROUP BY q_topics.ID_topic, q_topics.topic_name, q_topics.topic_subject, q_topics.topic_active, q_topics.topic_ord, q_topics.ID_topic  HAVING (((q_topics.topic_subject)=" + Replace(subj, "'", "''") + "))  ORDER BY q_topics.topic_ord, q_topics.ID_topic;"
topics.CursorType = 0
topics.CursorLocation = 3
topics.LockType = 3
topics.Open()
topics_numRows = 0
%>
<%
set subject = Server.CreateObject("ADODB.Recordset")
subject.ActiveConnection = Connect
subject.Source = "SELECT subject_name, ID_subject  FROM subjects  WHERE ID_subject = " + Replace(subj, "'", "''") + ""
subject.CursorType = 0
subject.CursorLocation = 3
subject.LockType = 3
subject.Open()
subject_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
topics_numRows = topics_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Repeat2__numRows = -1
Dim Repeat2__index
Repeat2__index = 0
topics_stats_numRows = topics_stats_numRows + Repeat2__numRows
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz topics. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
function trySubmit()
{
	if (document.forms[0].newtopic.value.length<3)
	{
		alert("Sorry, you must enter a name for a new Topic");
		return false;
	}
	if (confirm("Are you sure you want to add a new Topic to this Subject?"))	{	document.forms[0].submit();
	return false;
	}
return false;
}
function openwin(aspfile){
	win = window.open(aspfile,'display','left= 50, top=50, width=600, height=500,toolbar=no,  menubar=no,status=no,resizable=yes,scrollbars=yes');
	win.focus();
}
//-->
</script>
</HEAD>
<BODY>
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> Quiz topics list</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form name="add_topic" method="POST" action="<%=MM_editAction%>" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr> 
            <td colspan="7" class="subheads">Topics in <%=(subject.Fields.Item("subject_name").Value)%>:</td>
          </tr>
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">            <td class="text" width="10"><img src="../admin/images/back.gif" width="18" height="14"></td>
            <td class="text" colspan="7"><a href="../admin/q_list_of_subjects.asp">...go 
              up one level to list of Subjects</a></td>
          </tr>
          <% If Not topics.EOF Or Not topics.BOF Then %>
          <% 
While ((Repeat1__numRows <> 0) AND (NOT topics.EOF)) 
%>
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
            <td class="text" width="20"><%=numbers%></td>
            <td class="text" width="480"><a href="../admin/q_list_of_questions.asp?subj=<%=subj%>&topic=<%=(topics.Fields.Item("ID_topic").Value)%>"><%=(topics.Fields.Item("topic_name").Value)%></a></td>
            <td class="text" width="20"><a href="../admin/_export_quiz_questions.asp?subj=<%=subj%>&topic=<%=(topics.Fields.Item("ID_topic").Value)%>"><img src="images/export.gif" width="16" height="16" border="0"></a></td>
            <td width="20" class="text" align="right"> 
              <%if abs(topics.Fields.Item("topic_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
            </td>
            <td width="20" class="text" align="right"> 
              <%if allow_word_export then %>
              <a href="_export_quiz.asp?subj=<%=subj%>&topic=<%=(topics.Fields.Item("ID_topic").Value)%>" target="_blank"><img src="images/wrd.gif" width="16" height="16" border="0"></a> 
              <%end if %>
            </td>
            <td width="20"  align="right"><a href="javascript:"   onClick="openwin('matrix_display.asp?subj=<%=subj%>&topic=<%=topics("ID_topic").Value%>');"><img src="images/matrix.gif" width="16" height="16" border="0"></a></td>
            <td width="20" class="text" align="right"><a href="../admin/q_edit_topic.asp?topic=<%=(topics.Fields.Item("ID_topic").Value)%>&subj=<%=subj%>"><img src="images/edit.gif" width="16" height="15" border="0"></a> 
            </td>
            <td width="20" class="text" align="right"><a href="../admin/q_order_questions.asp?subj=<%=subj%>&topic=<%=(topics.Fields.Item("ID_topic").Value)%>"><img src="images/change.gif" width="15" height="15" border="0"></a></td>
            
          </tr>
          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  topics.MoveNext()
  numbers=numbers+1
Wend
%>
          <% End If ' end Not topics.EOF Or NOT topics.BOF %>
          <% If topics.EOF And topics.BOF Then %>
          <tr > 
            <td class="text">&nbsp;</td>
            <td  colspan="7">Sorry, 
              there are no topics in this subject currently. </td>
          </tr>
          <% End If ' end topics.EOF And topics.BOF %>
          <tr class="table_normal"> 
            <td class="text"><img src="images/new2.gif" width="11" height="13"></td>
            <td class="text" colspan="7"> 
              <input type="text" name="newtopic" size="85" class="formitem1">
            </td>
          </tr>
          <tr> 
            <td class="text"> 
              <input type="hidden" name="UID" value="<%=GetUniqueID("t_",20,"")%>">
            </td>
            <td class="text" colspan="7"> 
              <input type="reset" name="Submit2" value="Reset the form" class="quiz_button">
              <input type="submit" name="Submit" value="Add this new topic to current subject" class="quiz_button" <%call IsEditOK%>>
              <input type="hidden" name="id_subject" value="<%=subj%>">
              <input type="hidden" name="topic_active" value="1">
              <input type="hidden" name="topic_ord" value="999999">
            </td>
          </tr>
          <tr > 
				<td align="center" valign="bottom" class="subheads" colspan="7"> 
					<br><BR><a href="javascript:" onClick="window.open('q_stats_of_topics.asp?subj=<%=subj%>','statswindow','scrollbars=yes,resizable=yes,width=700,height=500,left=50,top=50')">Click to view topic statistics</a>
				</td>
			</tr>
        </table>
        <input type="hidden" name="MM_insert" value="true">
      </form>
      <p>&nbsp;</p>
    </td>
  </tr>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
call log_the_page ("Quiz List Topics: " & subj)
%>

<%
topics.Close()
%>
<%
subject.Close()
%>
