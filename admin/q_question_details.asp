<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->

<%
Dim qid
If (Request.QueryString("qid") <> "") Then 
qid = cInt(Request.QueryString("qid"))
Else 
Response.Redirect("error.asp?" & request.QueryString) 
End If
%>
<% 
sess_id = Request.QueryString("sess_id")
date_end_v1 = CDate("30/10/2012")

set sessions = Server.CreateObject("ADODB.Recordset")
sessions.ActiveConnection = Connect
sessions.Source = "SELECT q_session.Session_finish  FROM q_session WHERE id_Session = " + sess_id + ""
sessions.CursorType = 0
sessions.CursorLocation = 3
sessions.LockType = 3
sessions.Open()
sessions_numRows = 0

set question = Server.CreateObject("ADODB.Recordset")
question.ActiveConnection = Connect
IF date_end_v1 < sessions.Fields.item("Session_finish").Value THEN
	question.Source = "SELECT q_question.*, subjects.ID_subject, subjects.subject_name, new_subjects.s_topic AS topic_name  FROM subjects INNER JOIN (new_subjects INNER JOIN q_question ON new_subjects.s_ID = q_question.question_topic) ON subjects.ID_subject = new_subjects.s_qID  WHERE q_question.ID_question =" + Replace(qid, "'", "''") + ";"
ELSE
	question.Source = "SELECT q_question_v1.*, subjects.ID_subject, subjects.subject_name, q_topics.topic_name  FROM subjects INNER JOIN (q_topics INNER JOIN q_question_v1 ON q_topics.ID_topic = q_question_v1.question_topic) ON subjects.ID_subject = q_topics.topic_subject  WHERE q_question_v1.ID_question =" + Replace(qid, "'", "''") + ";"
END IF
question.CursorType = 0
question.CursorLocation = 3
question.LockType = 3
question.Open()
question_numRows = 0
%>
<%
Dim choices__qID
choices__qID = "1"
if ((question.Fields.Item("ID_question").Value) <> "") then choices__qID = cInt((question.Fields.Item("ID_question").Value))
%>
<%
set choices = Server.CreateObject("ADODB.Recordset")
choices.ActiveConnection = Connect
IF date_end_v1 < sessions.Fields.item("Session_finish").Value THEN
	choices.Source = "SELECT q_choice.*, q_choice.choice_question  FROM q_choice  WHERE (((q_choice.choice_question)=" + Replace(choices__qID, "'", "''") + "));"
ELSE
	choices.Source = "SELECT q_choice_v1.*, q_choice_v1.choice_question  FROM q_choice_v1  WHERE (((q_choice_v1.choice_question)=" + Replace(choices__qID, "'", "''") + "));"
END IF
choices.CursorType = 0
choices.CursorLocation = 3
choices.LockType = 3
choices.Open()
choices_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
choices_numRows = choices_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
choices_total = choices.RecordCount

' set the number of rows displayed on this page
If (choices_numRows < 0) Then
  choices_numRows = choices_total
Elseif (choices_numRows = 0) Then
  choices_numRows = 1
End If

' set the first and last displayed record
choices_first = 1
choices_last  = choices_first + choices_numRows - 1

' if we have the correct record count, check the other stats
If (choices_total <> -1) Then
  If (choices_first > choices_total) Then choices_first = choices_total
  If (choices_last > choices_total) Then choices_last = choices_total
  If (choices_numRows > choices_total) Then choices_numRows = choices_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (choices_total = -1) Then

  ' count the total records by iterating through the recordset
  choices_total=0
  While (Not choices.EOF)
    choices_total = choices_total + 1
    choices.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (choices.CursorType > 0) Then
'    choices.MoveFirst
'  Else
    choices.Requery
  End If

  ' set the number of rows displayed on this page
  If (choices_numRows < 0 Or choices_numRows > choices_total) Then
    choices_numRows = choices_total
  End If

  ' set the first and last displayed record
  choices_first = 1
  choices_last = choices_first + choices_numRows - 1
  If (choices_first > choices_total) Then choices_first = choices_total
  If (choices_last > choices_total) Then choices_last = choices_total

End If
%>
<%
session("choices_total") = choices_total
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz question info. You are logged in as <%=Session("MM_Username_admin")%> </TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_changeProp(objName,x,theProp,theValue) { //v3.0
  var obj = MM_findObj(objName);
  if (obj && (theProp.indexOf("style.")==-1 || obj.style)) eval("obj."+theProp+"='"+theValue+"'");
}
//-->
</script>
</HEAD>
<BODY BGCOLOR=#FFFF99 TEXT=#000000 VLINK=#000000 LINK=#000000 leftmargin="10" topmargin="10">
<table border="0" cellspacing="1" cellpadding="2" width="400">
  <tr align="left" valign="top"> 
    <td class="text_table" width="200">Subject</td>
    <td class="text_table" width="200">Topic</td>
  </tr>
  <tr align="left" valign="top"> 
    <td class="formitem1"><%=(question.Fields.Item("subject_name").Value)%></td>
    <td class="formitem1"><%=(question.Fields.Item("topic_name").Value)%></td>
  </tr>
  <tr align="left" valign="top"> 
    <td class="text_table" colspan="2">Question body</td>
  </tr>
  <tr align="left" valign="top"> 
    <td class="formitem1" colspan="2"><%=(question.Fields.Item("question_body").Value)%>&nbsp;</td>
  </tr>
  <tr align="left" valign="top"> 
    <td class="text_table" colspan="2"> 
      <table border="0" cellspacing="1" cellpadding="0">
        <% 
While ((Repeat1__numRows <> 0) AND (NOT choices.EOF)) 
%>
        <tr align="left" valign="middle"> 
          <td class="text_table" width="10"> 
            <%if abs(choices.Fields.Item("choice_cor").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
          </td>
          <td class="formitem2" width="30"><%=(choices.Fields.Item("choice_label").Value)%></td>
          <td class="formitem2" width="550"><%=(choices.Fields.Item("choice_body").Value)%></td>
        </tr>
        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  choices.MoveNext()
Wend
%>
      </table>
    </td>
  </tr>
  <tr align="left" valign="top"> 
    <td class="text_table">General <font color="#006600"><b>correct</b> </font> 
      feedback</td>
    <td class="text_table">General <font color="#FF0000"><b>incorrect</b></font><font color="#006600"> 
      </font> feedback</td>
  </tr>
  <tr align="left" valign="top"> 
    <td class="formitem_cor"><%=(question.Fields.Item("question_fb_cor").Value)%>&nbsp;</td>
    <td class="formitem_inc"><%=(question.Fields.Item("question_fb_inc").Value)%>&nbsp;</td>
  </tr>
  <tr align="left" valign="top"> 
    <td class="text_table" colspan="2">More information about the question</td>
  </tr>
  <tr align="left" valign="top"> 
    <td class="formitem1" colspan="2"><%=(question.Fields.Item("question_more").Value)%>&nbsp;</td>
  </tr>
  <tr align="left" valign="top"> 
    <td class="text_table">Question active? 
      <%if abs(question.Fields.Item("question_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
    </td>
    <td class="text_table" align="right">
<input type="button" name="close" value="Close this window" class="quiz_button" onClick="window.close()">
    </td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</HTML>

<%
call log_the_page ("Quiz Question details: " & qid)
question.Close()
choices.Close()
%>
