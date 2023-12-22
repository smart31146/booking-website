<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
Dim subj
subj = ""
if (Request.QueryString("subj") <> "") then subj = Request.QueryString("subj")
%>
<%
Dim topic
topic = "*"
if (Request.QueryString("topic") <> "") then topic = Request.QueryString("topic")
%>
<%
set training = Server.CreateObject("ADODB.Recordset")
training.ActiveConnection = Connect

if subj <> "" then
  if topic <> "*" then
    training.Source = "SELECT tr_pages.page_subject, tr_pages.page_topic, tr_pages.ID_page, subjects.subject_name, subjects.subject_active_t, subjects.subject_ord, tr_topics.topic_name, tr_topics.topic_active, tr_pages.page_monkey, tr_monkeys.monkey_name, tr_pages.page_title, tr_pages.page_photo, tr_pages.page_text, tr_pages.page_scenario, tr_pages.page_note, tr_pages_1.page_title AS scenario_page_title  FROM (subjects INNER JOIN tr_topics ON subjects.ID_subject = tr_topics.topic_subject) INNER JOIN (tr_monkeys INNER JOIN (tr_pages LEFT JOIN tr_pages AS tr_pages_1 ON tr_pages.page_scenario = tr_pages_1.ID_page) ON tr_monkeys.ID_monkey = tr_pages.page_monkey) ON tr_topics.ID_topic = tr_pages.page_topic  WHERE (((tr_pages.page_subject)= " + Replace(subj, "'", "''") + ") AND ((tr_pages.page_topic)= " + Replace(topic, "'", "''") + ")  AND tr_pages.page_active = 1)  ORDER BY subjects.subject_ord, tr_topics.topic_ord, tr_topics.id_topic,tr_pages.page_ord, tr_pages.ID_page, tr_pages.page_subject, tr_pages.page_topic ;  "
  else
    training.Source = "SELECT tr_pages.page_subject, tr_pages.page_topic, tr_pages.ID_page, subjects.subject_name, subjects.subject_active_t, subjects.subject_ord, tr_topics.topic_name, tr_topics.topic_active, tr_pages.page_monkey, tr_monkeys.monkey_name, tr_pages.page_title, tr_pages.page_photo, tr_pages.page_text, tr_pages.page_scenario, tr_pages.page_note, tr_pages_1.page_title AS scenario_page_title  FROM (subjects INNER JOIN tr_topics ON subjects.ID_subject = tr_topics.topic_subject) INNER JOIN (tr_monkeys INNER JOIN (tr_pages LEFT JOIN tr_pages AS tr_pages_1 ON tr_pages.page_scenario = tr_pages_1.ID_page) ON tr_monkeys.ID_monkey = tr_pages.page_monkey) ON tr_topics.ID_topic = tr_pages.page_topic  WHERE (((tr_pages.page_subject)= " + Replace(subj, "'", "''") + ") AND (tr_topics.topic_active=1) AND tr_pages.page_active = 1)  ORDER BY  subjects.subject_ord,tr_topics.topic_ord ,tr_topics.id_topic,  tr_pages.page_ord, tr_pages.page_subject, tr_pages.page_topic,  tr_pages.ID_page ;  "
  end if
else
    ' Denis 2008.09.12 - Export all subjects
    training.Source = "SELECT tr_pages.page_subject, tr_pages.page_topic, tr_pages.ID_page, subjects.subject_name, subjects.subject_active_t, subjects.subject_ord, tr_topics.topic_name, tr_topics.topic_active, tr_pages.page_monkey, tr_monkeys.monkey_name, tr_pages.page_title, tr_pages.page_photo, tr_pages.page_text, tr_pages.page_scenario, tr_pages.page_note, tr_pages_1.page_title AS scenario_page_title  FROM (subjects INNER JOIN tr_topics ON subjects.ID_subject = tr_topics.topic_subject) INNER JOIN (tr_monkeys INNER JOIN (tr_pages LEFT JOIN tr_pages AS tr_pages_1 ON tr_pages.page_scenario = tr_pages_1.ID_page) ON tr_monkeys.ID_monkey = tr_pages.page_monkey) ON tr_topics.ID_topic = tr_pages.page_topic WHERE  (tr_topics.topic_active=1) AND (subjects.subject_active_t=1)  AND (tr_pages.page_active = 1) ORDER BY  subjects.subject_ord, tr_pages.page_subject, tr_pages.page_topic,  tr_pages.ID_page, tr_pages.page_ord ;  "
end if

training.CursorType = 0
training.CursorLocation = 3
training.LockType = 3
training.Open()
training_numRows = 0
%>

<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
training_numRows = training_numRows + Repeat1__numRows
%>
<%
set questions = Server.CreateObject("ADODB.Recordset")
questions.ActiveConnection = Connect
questions.CursorType = 0
questions.CursorLocation = 3
questions.LockType = 3
questions_numRows = 0
%>
<%
set feedbacks = Server.CreateObject("ADODB.Recordset")
feedbacks.ActiveConnection = Connect
feedbacks.CursorType = 0
feedbacks.CursorLocation = 3
feedbacks.LockType = 3
feedbacks_numRows = 0
%>


<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Training export. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
</HEAD>
<BODY BGCOLOR=#FFFFFF TEXT=#000000 VLINK=#000000 LINK=#000000 leftmargin="5" topmargin="5" >
<p class="subheads" align="center">Better Business Training - content export</p>
<p>
  <%if allow_word_export then %>
</p>
<%
' Denis - 2008.09.12 - Keep tracking changes of subject if exporting all subjects
dim previousHeading
previousHeading = ""

While ((Repeat1__numRows <> 0) AND (NOT training.EOF))
%>
<%
if cInt(training.Fields.Item("subject_active_t").Value) = 0 then subject_active = false else subject_active = true
if cInt(training.Fields.Item("topic_active").Value) = 0 then topic_active = false else topic_active = true

' Denis - 2008.09.12 -
' If we are exporting many subjects, and there's a new subject being displayed, make it more evident to the user
subject_name=(training.Fields.Item("subject_name").Value)

if ((subj = "") and (previousHeading <> subject_name)) then
	previousHeading = subject_name
	response.write("<h2>Subject: " + subject_name + "</h2>")
end if
%>
<table width="100%" border="1" cellspacing="0" cellpadding="0" >
  <tr align="left" valign="top">
    <td colspan="4" bgcolor="#000000">&nbsp;</td>
  </tr>
  <tr align="left" valign="top">
    <td width="10%" bgcolor="#CCCCCC">Subject:</td>
    <td width="40%" class="subheads" <%
if NOT subject_active then response.write("bgcolor='#FF0000'")
%>><%
response.write (subject_name)
%></td>
    <td width="10%" bgcolor="#CCCCCC">Date:</td>
    <td width="40%"><%=cDate(Now())%></td>
  </tr>
  <tr align="left" valign="top">
    <td width="10%" bgcolor="#CCCCCC">Topic:</td>
    <td width="40%" class="subheads" <%
if (NOT topic_active) or (NOT subject_active) then response.write("bgcolor='#FF0000'")
%>><%=(training.Fields.Item("topic_name").Value)%></td>
    <td width="10%" bgcolor="#CCCCCC">&nbsp;</td>
    <td width="40%">&nbsp;</td>
  </tr>
</table>
<br>
<table width="100%" border="1" cellspacing="0" cellpadding="0" >
  <tr align="left" valign="top">
    <td width="10%" bgcolor="#CCCCCC">Author:</td>
    <td width="40%">GSM</td>
    <td width="10%" bgcolor="#CCCCCC">Reviewer:</td>
    <td width="40%">&nbsp;</td>
  </tr>
  <tr align="left" valign="top">
    <td width="10%" bgcolor="#CCCCCC">Version:</td>
    <td width="40%">&nbsp;</td>
    <td width="10%" bgcolor="#CCCCCC">Review date:</td>
    <td width="40%">&nbsp;</td>
  </tr>
</table>
<br>
<%
image = (training.Fields.Item("page_photo").Value)
%>
<table width="100%" border="1" cellspacing="0" cellpadding="0" >
  <tr align="left" valign="top">
    <td width="10%" bgcolor="#FFFFCC">Title:</td>
    <td width="50%" class="subheads"><%=(training.Fields.Item("page_title").Value)%></td>
    <td bgcolor="#FFFFCC" width="40%">Comments / Feedback:</td>
  </tr>
  <tr align="left" valign="top">
    <td width="10%" bgcolor="#FFFFCC">Page body:</td>
    <td width="50%"><%=(training.Fields.Item("page_text").Value)%></td>
    <td rowspan="4">&nbsp;</td>
  </tr>
  <tr align="left" valign="top">
    <td width="10%" bgcolor="#FFFFCC">Image:</td>
    <td width="50%">
      <%if image <> "" then response.write(image) else response.write("none")%>
    </td>
  </tr>
  <tr align="left" valign="top">
    <td width="10%" bgcolor="#FFFFCC">Monkey:</td>
    <td width="50%"><%=(training.Fields.Item("monkey_name").Value)%></td>
  </tr>
  <tr align="left" valign="top">
    <td width="10%" bgcolor="#FFFFCC">Scenario:</td>
    <td width="50%">
      <%if (training.Fields.Item("page_scenario").Value) <> 0 then response.write(training.Fields.Item("scenario_page_title").Value) else response.write"none"%>
    </td>
  </tr>
</table>
<br>
<%
whichpage = cInt(training.Fields.Item("id_page").Value)
%>
<%
questions.Source = "SELECT tr_questions.question_text, tr_questions.question_ok, tr_feedback.feedback_head  FROM tr_questions INNER JOIN tr_feedback ON tr_questions.question_ok = tr_feedback.ID_feedback  WHERE (((tr_questions.question_ID_page)= " + Replace(whichpage, "'", "''") + "));"
questions.Open()
questions_numRows = 0
%>
<%
Dim Repeat2__numRows
Repeat2__numRows = -1
Dim Repeat2__index
Repeat2__index = 0
questions_numRows = questions_numRows + Repeat2__numRows
%>
<% If Not questions.EOF Or Not questions.BOF Then %>
<table width="100%" border="1" cellspacing="0" cellpadding="0" >
  <%
While ((Repeat2__numRows <> 0) AND (NOT questions.EOF))
%>
  <tr align="left" valign="top">
    <td width="10%" bgcolor="#CCFFFF">Answer:</td>
    <td width="40%" ><%=(questions.Fields.Item("question_text").Value)%></td>
    <td width="10%"  align="right"><b><font color="#FF0000"><%= "(" & (questions.Fields.Item("question_ok").Value) & ")"%></font></b></td>
    <td width="40%">&nbsp;</td>
  </tr>
  <%
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  questions.MoveNext()
Wend
%>
</table>
<% End If ' end Not questions.EOF Or NOT questions.BOF %>
  <%
questions.Close()
%>
<br>
<%
feedbacks.Source = "SELECT tr_feedback.feedback_head, tr_feedback.feedback_text, tr_feedback.ID_feedback FROM tr_feedback WHERE (((tr_feedback.feedback_ID_page)= " + Replace(whichpage, "'", "''") + "));"
feedbacks.Open()
feedbacks_numRows = 0
%>
<%
Dim Repeat3__numRows
Repeat3__numRows = -1
Dim Repeat3__index
Repeat3__index = 0
feedbacks_numRows = feedbacks_numRows + Repeat3__numRows
%>
<% If Not feedbacks.EOF Or Not feedbacks.BOF Then %>
<table width="100%" border="1" cellspacing="0" cellpadding="0" >
  <%
While ((Repeat3__numRows <> 0) AND (NOT feedbacks.EOF))
%>
  <tr align="left" valign="top">
    <% headers = (feedbacks.Fields.Item("feedback_head").Value)
if headers = "" then headers = "none"
%>
    <td width="10%" bgcolor="#FFFF66">FB
      header:</td>
    <td class="subheads"><%=(feedbacks.Fields.Item("ID_feedback").Value)%> - <%=headers%></td>
    <td rowspan="2" width="40%">&nbsp;</td>
  </tr>
  <tr align="left" valign="top">
    <td width="10%" bgcolor="#FFFF66">FB
      Body:</td>
    <td width="50%" ><%=(feedbacks.Fields.Item("feedback_text").Value)%></td>
  </tr>
  <%
  Repeat3__index=Repeat3__index+1
  Repeat3__numRows=Repeat3__numRows-1
  feedbacks.MoveNext()
Wend
%>
</table>
<% End If ' end Not feedbacks.EOF Or NOT feedbacks.BOF %>
  <%
feedbacks.Close()
%>
<br>
<%page_note=(training.Fields.Item("page_note").Value)
if page_note <>"" then
%>
<table width="100%" border="1" cellspacing="0" cellpadding="0" >
  <tr>
    <td width="10%" bgcolor="#FFFFCC">Bottom line:</td>
    <td><%=page_note%></td>
  </tr>
</table>
<%end if %>
<%
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  training.MoveNext()
Wend
%>
<p>
  <% else %>
  Sorry, the export function is not available at the moment.
  <% end if %>
</p>
<hr>
<p align="center">strictly confidential | &copy; copyright 2002-<%= Year(Now_BBP()) %> Law of the Jungle
  Pty Limited<br>
  not to be used or disclosed otherwise than for a purpose expressly permitted
  by Law of the Jungle | all rights reserved</p>
</BODY>
</HTML>

<%
call log_the_page ("Training Word Export: " & subject_name)
%>

<%
training.Close()
%>

