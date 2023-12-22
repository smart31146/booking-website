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
  MM_editTable = "subjects"
  MM_editRedirectUrl = "q_list_of_subjects.asp"
  MM_fieldsStr  = "newsubject|value|subject_active_q|value|subject_active_t|value|subject_active_b|value|subject_ord|value|UID|value"
  MM_columnsStr = "subject_name|',none,''|subject_active_q|none,none,NULL|subject_active_t|none,none,NULL|subject_active_b|none,none,NULL|subject_ord|none,none,NULL|subject_UID|',none,''"

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
' *** Insert Record: construct a sql insert statement and execute it

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
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
    call log_the_page ("Quiz Execute - INSERT Subject")
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
set subjects = Server.CreateObject("ADODB.Recordset")
subjects.ActiveConnection = Connect
subjects.Source = "SELECT subjects.ID_subject, subjects.subject_name, subjects.subject_active_q, subjects.ID_subject  FROM subjects  GROUP BY subjects.ID_subject, subjects.subject_name, subjects.subject_active_q, subjects.subject_ord, subjects.ID_subject   ORDER BY subjects.subject_ord, subjects.ID_subject;"
subjects.CursorType = 0
subjects.CursorLocation = 3
subjects.LockType = 3
subjects.Open()
subjects_numRows = 0
%>
<%
set subjects_stats = Server.CreateObject("ADODB.Recordset")
subjects_stats.ActiveConnection = Connect
subjects_stats.Source = "SELECT COUNT(logs.ID_log) AS ID_log, subjects.subject_name, subjects.ID_subject, MAX(logs.log_date) AS log_date  FROM logs INNER JOIN subjects ON logs.log_subjID = subjects.ID_subject  GROUP BY logs.log_module, subjects.subject_name, subjects.ID_subject  HAVING (logs.log_module LIKE 'quiz')  ORDER BY COUNT(logs.ID_log) DESC;"
subjects_stats.CursorType = 0
subjects_stats.CursorLocation = 3
subjects_stats.LockType = 3
'subjects_stats.Open()
subjects_stats_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
subjects_numRows = subjects_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Repeat2__numRows = -1
Dim Repeat2__index
Repeat2__index = 0
subjects_stats_numRows = subjects_stats_numRows + Repeat2__numRows
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz subjects. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="javascript" type="text/javascript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
function trySubmit()
{
	if (document.forms[0].newsubject.value.length<2)
	{
		alert("Sorry, you must enter a name for a new Subject!\n(min. 2 characters)");
		return false;
	}
	if (confirm("Are you sure you want to add a new subject?"))	{	document.forms[0].submit();
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
    <td align="left" valign="bottom" class="heading"> Training & Quiz pages</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form name="add_subject" method="POST" action="<%=MM_editAction%>" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr> 
            <td colspan="5" class="subheads">Subjects:</td>
            <td align="right" class="subheads" colspan="3">
              <%if allow_word_export then %>
				<a href="_export_quiz.asp" target="_blank"><img src="images/wrd.gif" width="16" height="16" border="0" alt="Export All Subjects to Word"></a>
			  <% end if %>
				<a href="q_order_subjects.asp"><img src="images/change.gif" width="15" height="15" border="0"></a>
			</td>
          </tr>
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
            <td class="text" width="10"><img src="../admin/images/back.gif" width="18" height="14"></td>
            <td colspan="7" class="text"><a href="../admin/main.asp">...Home page 
              </a></td>
          </tr>
          <% If Not subjects.EOF Or Not subjects.BOF Then %>
          <% 
While ((Repeat1__numRows <> 0) AND (NOT subjects.EOF)) 
%>
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')" > 
            <td class="text" width="20"><%=numbers%></td>
            <td width="500" class="text"><a href="q_training.asp?ID_subject=<%=(subjects.Fields.Item("ID_subject").Value)%>"><%=(subjects.Fields.Item("subject_name").Value)%></a></td>
            <td width="20"  align="right"> 
              <%if abs(subjects.Fields.Item("subject_active_q").Value) = 1 then response.write "<img src='images/1.gif' alt='Active'>" else response.write "<img src='images/0.gif' alt='Inactive'>"%>
            </td>
            <td width="20"  align="right"> 
              <%if allow_word_export then %>
              <a href="_export_trainingquiz.asp?subj=<%=(subjects.Fields.Item("ID_subject").Value)%>" target="_blank"><img src="images/wrd.gif" width="16" height="16" border="0" alt="Export to Word"></a> 
              <%end if %>
            </td>
            <td width="20"  align="right"><a href="question_stats.asp?subj=<%=(subjects.Fields.Item("ID_subject").Value)%>"><img src="images/statistics.gif" width="15" height="15" border="0" alt="Question Statistics"></a></td>
			<td width="20"  align="right"><a href="matrix.asp?subject=<%=(subjects.Fields.Item("ID_subject").Value)%>"><img src="images/matrix.gif" width="16" height="16" border="0" alt="Matrix for user results"></a></td>
            <td width="20"  align="right"><a href="../admin/q_edit_subject.asp?subj=<%=(subjects.Fields.Item("ID_subject").Value)%>"><img src="images/edit.gif" width="16" height="15" border="0" alt="Edit user details"></a></td>
			<td width="20"  align="right"><a href="q_order_topics.asp?subj=<%=(subjects.Fields.Item("ID_subject").Value)%>"><img src="images/change.gif" width="15" height="15" border="0" alt="Sort Subject Topics"></a></td>
          </tr>
          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  subjects.MoveNext()
  numbers=numbers+1
Wend
%>
          <% End If %>
          <% If subjects.EOF And subjects.BOF Then %>
          <tr> 
            <td >&nbsp;</td>
            <td colspan="7" >Sorry, 
              there are no subjects in the quiz currently.</td>
          </tr>
          <% End If %>
          <tr class="table_normal" colspan=7 >
            <td ><img src="images/new2.gif" width="11" height="13"></td>
            <td width="99%"  colspan="7"> 
              <input type="text" name="newsubject" size="85" class="formitem1">
            </td>
          </tr>
          <tr> 
            <td> 
              <input type="hidden" name="UID" value="<%=GetUniqueID("s_",20,"")%>">
            </td>
            <td width="99%"  colspan="5"> 
              <input type="reset" name="Submit2" value="Reset the form" class="quiz_button">
              <input type="submit" name="Submit" value="Add this new subject" class="quiz_button" <%call IsEditOK%>>
              <input type="hidden" name="subject_active_q" value="1">
              <input type="hidden" name="subject_active_t" value="0">
              <input type="hidden" name="subject_active_b" value="0">
              <input type="hidden" name="subject_ord" value="999999">
            </td>
          </tr>
		   <tr > 
				<td align="center" valign="bottom" class="subheads" colspan="7"> 
					<br><BR><a href="javascript:" onClick="window.open('q_stats_of_subjects.asp','statswindow','scrollbars=yes,resizable=yes,width=700,height=500,left=50,top=50')">Click to view subject statistics</a>
				</td>
			</tr>
        </table>
        <input type="hidden" name="MM_insert" value="true">
      </form>
      <p>&nbsp;</p>
    </td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      
    </td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
call log_the_page ("Quiz List Subjects")
%>

<%
subjects.Close()
%>

