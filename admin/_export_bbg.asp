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
set BBG = Server.CreateObject("ADODB.Recordset")
BBG.ActiveConnection = Connect

if subj <> "" then
  if topic <> "*" then
    BBG.Source = "SELECT subjects.subject_name, subjects.subject_active_b, subjects.subject_ord, b_topics.topic_name, b_topics.topic_active, b_topics.topic_title, b_topics.topic_keyp, b_topics.topic_exmp, b_topics.topic_training, b_topics.topic_qanda, b_pages.page_title, b_pages.page_active, b_pages.page_header, b_pages.page_text, b_pages.page_icon  FROM subjects INNER JOIN (b_topics INNER JOIN b_pages ON b_topics.ID_topic = b_pages.page_topic) ON subjects.ID_subject = b_topics.topic_subject  WHERE (((subjects.ID_subject)=" + Replace(subj, "'", "''") + ") AND ((b_topics.ID_topic)=" + Replace(topic, "'", "''") + "))  AND (b_pages.page_active = 1)  ORDER BY subjects.subject_ord, subjects.ID_subject, b_topics.topic_ord, b_topics.ID_topic, b_pages.page_ord, b_pages.ID_page;"
  else
    BBG.Source = "SELECT subjects.subject_name, subjects.subject_active_b, subjects.subject_ord, b_topics.topic_name, b_topics.topic_active, b_topics.topic_title, b_topics.topic_keyp, b_topics.topic_exmp, b_topics.topic_training, b_topics.topic_qanda, b_pages.page_title, b_pages.page_active, b_pages.page_header, b_pages.page_text, b_pages.page_icon  FROM subjects INNER JOIN (b_topics INNER JOIN b_pages ON b_topics.ID_topic = b_pages.page_topic) ON subjects.ID_subject = b_topics.topic_subject  WHERE (((subjects.ID_subject)=" + Replace(subj, "'", "''") + ") AND (b_topics.topic_active=1))  AND (b_pages.page_active = 1)  ORDER BY subjects.subject_ord, subjects.ID_subject, b_topics.topic_ord, b_topics.ID_topic, b_pages.page_ord, b_pages.ID_page;"
  end if
else
  ' Denis 2008.09.12 - Export all subjects
  BBG.Source = "SELECT subjects.subject_name, subjects.subject_active_b, subjects.subject_ord, b_topics.topic_name, b_topics.topic_active, b_topics.topic_title, b_topics.topic_keyp, b_topics.topic_exmp, b_topics.topic_training, b_topics.topic_qanda, b_pages.page_title, b_pages.page_active, b_pages.page_header, b_pages.page_text, b_pages.page_icon  FROM subjects INNER JOIN (b_topics INNER JOIN b_pages ON b_topics.ID_topic = b_pages.page_topic) ON subjects.ID_subject = b_topics.topic_subject WHERE (b_topics.topic_active=1) AND (subjects.subject_active_b=1)  AND (b_pages.page_active = 1) ORDER BY subjects.subject_ord, subjects.ID_subject, b_topics.topic_ord, b_topics.ID_topic, b_pages.page_ord, b_pages.ID_page;"
end if

BBG.CursorType = 0
BBG.CursorLocation = 3
BBG.LockType = 3
BBG.Open()
BBG_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
BBG_numRows = BBG_numRows + Repeat1__numRows
%>
<%
last_topic = ""
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: BBG export. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
</HEAD>
<BODY BGCOLOR=#FFFFFF TEXT=#000000 VLINK=#000000 LINK=#000000 leftmargin="5" topmargin="5" >
<p class="subheads" align="center">Better Business Guide - content export</p>
<p>
  <%if allow_word_export then %>
</p>
<%
' Denis - 2008.09.12 - Keep tracking changes of subject if exporting all subjects
dim previousHeading
previousHeading = ""

While ((Repeat1__numRows <> 0) AND (NOT BBG.EOF))
%>
<%
subject_name = (BBG.Fields.Item("subject_name").Value)
topic_title = (BBG.Fields.Item("topic_title").Value)
topic_name = (BBG.Fields.Item("topic_name").Value)
if cInt(BBG.Fields.Item("subject_active_b").Value) = 0 then subject_active = false else subject_active = true
if cInt(BBG.Fields.Item("topic_active").Value) = 0 then topic_active = false else topic_active = true

' Denis - 2008.09.12 -
' If we are exporting many subjects, and there's a new subject being displayed, make it more evident to the user
if ((subj = "") and (previousHeading <> subject_name)) then
	previousHeading = subject_name
	response.write("<h2>Subject: " + subject_name + "</h2>")
end if
%>
<%
if last_topic <> topic_name then
%>
<table width="100%" border="1" cellspacing="0" cellpadding="0" >
  <tr align="left" valign="top">
    <td colspan="4" bgcolor="#000000">&nbsp;</td>
  </tr>
  <tr align="left" valign="top">
    <td width="10%" bgcolor="#CCCCCC">Subject:</td>
    <td width="40%" class="subheads" <%
if NOT subject_active then response.write("bgcolor='#FF0000'")
%>><%=subject_name%></td>
    <td width="10%" bgcolor="#CCCCCC">Date:</td>
    <td width="40%"><%=cDate(Now_BBP())%></td>
  </tr>
  <tr align="left" valign="top">
    <td width="10%" bgcolor="#CCCCCC">Topic:</td>
    <td width="40%" class="subheads" <%
if (NOT topic_active) or (NOT subject_active) then response.write("bgcolor='#FF0000'")
%>><%=topic_name%></td>
    <td width="10%" bgcolor="#CCCCCC">Topic title:</td>
    <td width="40%"><%=topic_title%></td>
  </tr>
  <tr align="left" valign="top">
    <td width="10%" bgcolor="#CCCCCC">Keypoints:</td>
    <td width="40%"><%=ReplaceStrBBG((BBG.Fields.Item("topic_keyp").Value))%></td>
    <td width="10%" bgcolor="#CCCCCC">Examples:</td>
    <td width="40%"><%=ReplaceStrBBG((BBG.Fields.Item("topic_exmp").Value))%></td>
  </tr>
  <tr align="left" valign="top">
    <td width="10%" bgcolor="#CCCCCC">Training link:</td>
    <td width="40%">
      <%if (BBG.Fields.Item("topic_training").Value) <> 0 then response.write"Yes" else response.write"No"%>
    </td>
    <td width="10%" bgcolor="#CCCCCC">Q and A link:</td>
    <td width="40%">
      <%if (BBG.Fields.Item("topic_qanda").Value) <> 0 then response.write"Yes" else response.write"No"%>
    </td>
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
end if
%>
<%
page_title = (BBG.Fields.Item("page_title").Value)
if page_title ="" then page_title = "none"
page_header = (BBG.Fields.Item("page_header").Value)
if page_header ="" then page_header = "none"
page_text = (BBG.Fields.Item("page_text").Value)
if page_text ="" then page_text = "none"
page_icon = (BBG.Fields.Item("page_icon").Value)
if cInt(BBG.Fields.Item("page_active").Value) = 0 or cInt(BBG.Fields.Item("subject_active_b").Value) = 0 or cInt(BBG.Fields.Item("topic_active").Value) = 0 then page_active = false else page_active = true
%>
<table width="100%" border="1" cellspacing="0" cellpadding="0" >
  <tr align="left" valign="top" <%
if NOT page_active then response.write("bgcolor='#FF0000'")
%>>
    <td width="10%" bgcolor="#FFFFCC">Title:</td>
    <td width="50%" class="subheads"><%= ReplaceStrBBG(page_title)%></td>
    <td bgcolor="#FFFFCC" width="40%" colspan="2">Comments / Feedback:</td>
  </tr>
  <tr align="left" valign="top">
    <td width="10%" bgcolor="#FFFFCC">Header:</td>
    <td width="50%"><%= ReplaceStrBBG(page_header) %></td>
    <td rowspan="2" width="35%" align="left">&nbsp;</td>
    <td rowspan="2" width="5%" align="right">
      <%if page_icon <> "" then response.write("<img src='../client/bbg_icons/" & page_icon & "'>") else response.write("no icon")%>
    </td>
  </tr>
  <tr align="left" valign="top">
    <td width="10%" bgcolor="#FFFFCC">Body:</td>
    <td width="50%"><%= ReplaceStrBBG(page_text) %></td>
  </tr>
</table>
<br>
<%
last_topic = topic_name
%>
<%
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  BBG.MoveNext()
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
call log_the_page ("BBG Word Export: " & subject_name)
%>

<%
BBG.Close()
%>
