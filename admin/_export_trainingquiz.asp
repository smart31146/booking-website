<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
qset = 0
%>
<%
Dim subj
subj = ""
if (Request.QueryString("subj") <> "") then sID = Request.QueryString("subj")

set question = Server.CreateObject("ADODB.Recordset")
set obj = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT * "
SQL = SQL & " FROM new_subjects s1,subjects WHERE s1.s_qiD = ID_subject  AND ABS([s_active]) = 1  AND s_qID = "&fixstr(clng(sID))&" ORDER BY s_order ASC"

obj.Open SQL, Connect,3,3
objAntal = obj.RecordCount
'response.write SQL
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Training/Quiz export. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<!--<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">-->
<style>
* {font-size: 100%; font-family: Arial, Verdana, Geneva, Helvetica, sans-serif;}
.heading {font-size: 16px; font-weight: bold; font-style: normal; color: #FFFFFF}
.subheads {font-size: 12px; font-weight: bold}
</style>
</HEAD>
<BODY BGCOLOR=#FFFFFF TEXT=#000000 VLINK=#000000 LINK=#000000 leftmargin="5" topmargin="5" >
<p class="subheads" align="center" style="font-size:16px;">Better Business Training/Quiz - content export</p>

<% x=0
do until obj.eof
x=x+1%>
<% ' s_typ 1 = Traning, 2 = Quiz 
IF clng(obj("s_typ")) = 1 THEN
set question = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT q_id,q_title,q_div_info,q_order FROM new_questions WHERE q_tID = "&fixstr(clng(obj("s_id")))&" ORDER BY q_order"
'response.write SQL & "<br>"
question.Open SQL, Connect,3,3
if not question.eof then
	showQuestion = true
	QArr = question.GetRows 
ELSE
	showQuestion = False
	QArr = 0
END IF
question.close %>
<table width="100%" border="0" cellspacing="2" cellpadding="4"  bgcolor="#CCCCCC">
<TR bgcolor="#000000">
	<TD colspan="4"><br></TD>
</TR>
<TR bgcolor="#FFFFFF">
	<TD width="5%"><strong>Subject</strong></TD>
	<TD><% =obj("subject_name")%></TD>
	<td width="5%">Number</td>
	<td width="30%"><%=x %> of <% =objAntal%></td>
</TR>
<TR bgcolor="#FFFFFF">
	<TD><strong>Topic</strong></TD>
	<TD><% =obj("s_topic")%></TD>
	<TD>Training</TD>
	<TD>ID: <% =obj("s_id")%></TD>
</TR>

<TR bgcolor="#FFFFFF">
	<TD><strong>Title</strong></TD>
	<TD><% =ReplaceStrTR(obj("s_title"))%></TD>
	<TD colspan="2">Comments / feedback</TD>
</TR>
<TR bgcolor="#FFFFFF" valign="top">
	<TD><strong>Body</strong></TD>
	<TD><% =ReplaceStrTR(obj("s_body"))%>
	<%
			IF showQuestion = true THEN
			response.write "<br><br>"
				If Ubound(QArr,2) > -1 Then
				xi = 0
					 For i=0 to ubound(QArr,2)
					 xi=xi+1 %>
					<table>
			  		<td style="font-family: Arial, Verdana, Helvetica, sans-serif; font-size: 11px; border:1px solid #333333; background-color:#CCFF99; padding:5px;margin:5px; width: 700px"><strong><% =ReplaceStrTR(QArr(1,i))%></strong><br>
					<table>
					<td style="font-family: Arial, Verdana, Helvetica, sans-serif; font-size: 11px; border:1px solid #333333; background-color:#FFFF99; padding:5px;margin:5px; width: 700px"><% =ReplaceStrTR(QArr(2,i))%></td>
					</table>
					</td>
					</table>
					<% IF xi=2 THEN
					xi=0
					response.write "<div class=""clear""></div>"
					END IF%>
			<%		Next
				END IF
			END IF%>
	</TD>
	<TD colspan="2">&nbsp;</TD>
</TR>
<br clear="all" style="page-break-before:always" />
</table>
<% ELSEIF clng(obj("s_typ")) = 2 THEN
	SQL = "SELECT * FROM q_question,q_choice WHERE question_topic = "&fixstr(clng(obj("s_id")))&" AND ABS(question_active) = 1 AND ABS(choice_active) = 1 AND choice_question = ID_question ORDER BY ID_question, choice_label"
	'response.write SQL
	question.Open SQL, Connect,3,3

	
	%>
<table width="100%" border="0" cellspacing="2" cellpadding="4"  bgcolor="#CCCCCC">		
<TR bgcolor="#000000">
	<TD colspan="4"><br></TD>
</TR>
<TR bgcolor="#FFFFFF">
	<TD width="5%"><strong>Subject</strong></TD>
	<TD><% =obj("subject_name")%></TD>
	<td width="5%">Number</td>
	<td width="30%"><%=x %> of <% =objAntal%></td>
</TR>
<TR bgcolor="#FFFFFF">
	<TD><strong>Topic</strong></TD>
	<TD><% =obj("s_topic")%></TD>
	<TD>Quiz</TD>
	<TD>ID: <% =obj("s_id")%></TD>
</TR>
<!--<TR valign="top" bgcolor="#FFFFFF">
	<TD bgcolor="#CCFF99"><strong>Correct</strong></TD>
	<TD><% =ReplaceStrQuiz(question("question_fb_cor"))%></TD>
	<TD bgcolor="#FF3333">Incorrect</TD>
	<TD><% =ReplaceStrQuiz(question("question_fb_inc"))%></TD>
</TR>-->
<TR valign="top" bgcolor="#FFFFFF">
	<TD colspan="2">
	<% do until question.eof
	if prevCat <> question("question_body") THEN
	'response.write "<br>"&ReplaceStrQuiz(replace(question("question_body"),"<br />",""))&"<br>"
	%>
	<table >
		<TR valign="top">
			<TD bgcolor="#DDDDDD"><strong>Question</strong></TD>
			<TD><% =ReplaceStrQuiz(question("question_body"))%></TD>
		</TR>
		<TR valign="top">	
			<TD bgcolor="#CCFF99"><strong>Correct</strong></TD>
			<TD><% =ReplaceStrQuiz(question("question_fb_cor"))%></TD>
		</TR>
		<TR valign="top">		
			<TD bgcolor="#FF9999"><strong>Incorrect</strong></TD>
			<TD><% =ReplaceStrQuiz(question("question_fb_inc"))%></TD>
		</TR>
	</table>
	<%
	END IF 
	prevCat = question("question_body")%>
	<table >
		<TR valign="top">
			<TD width="30"><% IF cbool(question("choice_cor")) THEN %><img height="20" src="../images/icon_true.gif" alt=""><% END IF%>&nbsp;</TD>
			<TD><strong><% =question("choice_label")%></strong></TD>
			<TD width="200"><% =question("choice_body")%></TD>
		</TR>
	</table>
			 <% question.movenext
	loop%>
	
	<TD colspan="2">Comments / feedback</TD>
</TR>

<br clear="all" style="page-break-before:always" />
<% question.close
END IF%>
<!--<TR bgcolor="#FFFFFF">
	<TD colspan="4"><br></TD>
	</TR>
	<TR bgcolor="#000000">
	<TD colspan="4"><br></TD>
</TR>-->
<% obj.movenext
loop%>
</table>
</p>
<hr>
<p align="center">strictly confidential | &copy; copyright 2002-<%= Year(Now_BBP()) %> Law of the Jungle
  Pty Limited<br>
  not to be used or disclosed otherwise than for a purpose expressly permitted
  by Law of the Jungle | all rights reserved</p>
</BODY>
</HTML>

<%
'call log_the_page ("Quiz Word Export: " & subject_name)
%>

<%
obj.Close()
%>
