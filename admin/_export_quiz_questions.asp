<%@LANGUAGE="VBSCRIPT"%>
<% Response.Buffer="true" %>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
q_body = ""
q_count = 0
q_correct = 0
s_count = 0
Dim subj
subj = "1"
if (Request.QueryString("subj") <> "") then subj = Request.QueryString("subj")
Dim topic
topic = "1"
if (Request.QueryString("topic") <> "") then topic = Request.QueryString("topic")
set quiz = Server.CreateObject("ADODB.Recordset")
quiz.ActiveConnection = Connect
quiz.Source = "SELECT subjects.subject_name, subjects.subject_active_q, q_topics.topic_name, q_topics.topic_active, q_question.ID_question, q_question.question_body, q_question.question_active FROM subjects INNER JOIN (q_topics INNER JOIN q_question ON q_topics.ID_topic = q_question.question_topic) ON subjects.ID_subject = q_topics.topic_subject  WHERE (((subjects.ID_subject)=" + Replace(subj, "'", "''") + ") AND ((q_topics.ID_topic)=" + Replace(topic, "'", "''") + "))  ORDER BY q_topics.topic_ord, q_question.question_topic, q_question.question_ord, q_question.ID_question;"
quiz.CursorType = 0
quiz.CursorLocation = 3
quiz.LockType = 3
quiz.Open()
quiz_numRows = 0
%>
<%
Response.Clear()
Response.AddHeader "Content-Disposition","inline; filename=quiz - " & (quiz.Fields.Item("subject_name").Value) & " - " & (quiz.Fields.Item("topic_name").Value) & " - "& day(now()) & "." & month(now()) & "." & year(now()) & ".txt"
Response.ContentType = "application/unknown"
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
quiz_numRows = quiz_numRows + Repeat1__numRows
While ((Repeat1__numRows <> 0) AND (NOT quiz.EOF)) 
if cInt(quiz.Fields.Item("question_active").Value) = 0 or cInt(quiz.Fields.Item("subject_active_q").Value) = 0 or cInt(quiz.Fields.Item("topic_active").Value) = 0 then page_active = false else page_active = true
q_body = replace((quiz.Fields.Item("question_body").Value),vbcrlf,"") + vbcrlf
whichone = cInt(quiz.Fields.Item("id_question").Value)
s_count = s_count + 1
set choices = Server.CreateObject("ADODB.Recordset")
choices.ActiveConnection = Connect
choices.Source = "SELECT q_choice.choice_label, q_choice.choice_body, q_choice.choice_cor FROM q_choice WHERE (((q_choice.choice_question)=" + Replace(whichone, "'", "''") + "));"
choices.CursorType = 0
choices.CursorLocation = 3
choices.LockType = 3
choices.Open()
choices_numRows = 0
q_count = 0
q_correct = 0
Dim Repeat2__numRows
Repeat2__numRows = -1
choices_numRows = choices_numRows + Repeat2__numRows
While ((Repeat2__numRows <> 0) AND (NOT choices.EOF)) 
if (choices.Fields.Item("choice_label").Value) <> "" then q_body = q_body + (choices.Fields.Item("choice_label").Value) + " "
correct = abs(choices.Fields.Item("choice_cor").Value)
q_count = q_count + 1
if correct = 1 then q_correct = q_count
if (choices.Fields.Item("choice_body").Value) <> "" then q_body = q_body + (choices.Fields.Item("choice_body").Value) 
q_body = q_body + vbcrlf
  Repeat2__numRows=Repeat2__numRows-1
  choices.MoveNext()
Wend
s_help = s_count mod noqis
if s_help = 0 then s_help = noqis
	Select Case s_help
      Case 1  mm_str = "MM"
      Case 2  mm_str = "AM"
      Case 3  mm_str = "MA"
      Case 4  mm_str = "-M"
      Case Else  mm_str = "M-"
   	End Select
q_body = mm_str + cStr(q_count) + cStr(q_correct) + vbcrlf + q_body + vbcrlf
response.write(q_body)
choices.Close()
  Repeat1__numRows=Repeat1__numRows-1
  quiz.MoveNext()
Wend
response.write("strictly confidential | (c) Copyright 2007 Law of the Jungle Pty Limited | not to be used or disclosed otherwise than for a purpose expressly permitted by Law of the Jungle | all rights reserved")
call log_the_page ("Quiz Export for upload")
quiz.Close()
Response.Flush()
Response.End()
%>
