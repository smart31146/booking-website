<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->

<%
Dim user__MMColParam
user__MMColParam = "1"
if (Request.QueryString("user") <> "") then user__MMColParam = Request.QueryString("user")
%>
<%
numbers=1
%>
<%
set user = Server.CreateObject("ADODB.Recordset")
user.ActiveConnection = Connect
user.Source = "SELECT * FROM q_user WHERE ID_user = " + Replace(user__MMColParam, "'", "''") + ""
user.CursorType = 0
user.CursorLocation = 3
user.LockType = 3
user.Open()
user_numRows = 0
%>
<%
Dim sessions__MMColParam
sessions__MMColParam = "1"
if (Request.QueryString("user") <> "") then sessions__MMColParam = Request.QueryString("user")
%>
<%
Dim answers__MMColParam
answers__MMColParam = "1"
if (Request.QueryString("user_session") <> "") then answers__MMColParam = Request.QueryString("user_session")
%>
<%
set sessions = Server.CreateObject("ADODB.Recordset")
sessions.ActiveConnection = Connect
sessions.Source = "SELECT q_session.ID_Session, q_session.Session_date, subjects.subject_name, q_session.Session_total, q_session.Session_done, q_session.Session_correct, q_session.Session_finish  FROM q_user INNER JOIN (q_session INNER JOIN subjects ON q_session.Session_subject = subjects.ID_subject) ON q_user.ID_user = q_session.Session_users  WHERE id_Session = " + Replace(answers__MMColParam, "'", "''") + "  ORDER BY q_session.Session_date DESC , subjects.subject_name DESC;"
sessions.CursorType = 0
sessions.CursorLocation = 3
sessions.LockType = 3
sessions.Open()
sessions_numRows = 0




subj = cint(request("subject"))
set subject = Server.CreateObject("ADODB.Recordset")
subject.ActiveConnection = Connect
subject.Source = "SELECT * from subjects where id_subject=" & subj
subject.CursorType = 0
subject.CursorLocation = 3
subject.LockType = 3
subject.Open()
subject_numRows = 0
subject_name = subject("subject_name")

passrate = default_passrate
if passrate_type = 1 then
	set preferences = Server.CreateObject("ADODB.Recordset")
	preferences.ActiveConnection = Connect
	preferences.Source = "SELECT * FROM q_certification where q_certification.q_session = " + Replace(answers__MMColParam, "'", "''") + ""
	preferences.CursorType = 0
	preferences.CursorLocation = 3
	preferences.LockType = 3
	preferences.Open()
	preferences_numRows = 0
	if (not preferences.eof) then
		passrate = preferences.Fields.Item("percentage_required").Value
	else
		passrate = subject.fields.item("subject_passmark").value
	end if
	preferences.Close()
elseif passrate_type=2 then
	set preferences = Server.CreateObject("ADODB.Recordset")
	preferences.ActiveConnection = Connect
	preferences.Source = "SELECT * FROM q_certification where q_certification.q_session = " + Replace(answers__MMColParam, "'", "''") + ""
	preferences.CursorType = 0
	preferences.CursorLocation = 3
	preferences.LockType = 3
	preferences.Open()
	preferences_numRows = 0
	if (not preferences.eof) then
		passrate = preferences.Fields.Item("percentage_required").Value
	else
		set preferences2 = Server.CreateObject("ADODB.Recordset")
		preferences2.ActiveConnection = Connect
		preferences2.Source = "SELECT pass_rate FROM pass_rates where subject="&subj&" and q_info1="&user.Fields.Item("user_info1").Value&" and q_info3="&user.Fields.Item("user_info3").Value&";"
		preferences2.CursorType = 0
		preferences2.CursorLocation = 3
		preferences2.LockType = 3
		preferences2.Open()
		if (not preferences2.eof) then
			passrate = preferences2.Fields.Item("pass_rate").Value
		end if
		preferences2.Close()
	end if
	preferences.Close()
elseif passrate_type=3 then
	set preferences = Server.CreateObject("ADODB.Recordset")
	preferences.ActiveConnection = Connect
	preferences.Source = "SELECT pass_rate FROM pass_rates where subject="&subj&" and q_info1="&user.Fields.Item("user_info1").Value&" and q_info3="&user.Fields.Item("user_info3").Value&";"
	preferences.CursorType = 0
	preferences.CursorLocation = 3
	preferences.LockType = 3
	preferences.Open()
	preferences_numRows = 0
	if (not preferences.eof) then
		passrate = preferences.Fields.Item("pass_rate").Value
	end if
	preferences.Close()
else
	passrate = subject.fields.item("subject_passmark").value
end if
subject.Close
%>


<%
set answers = Server.CreateObject("ADODB.Recordset")
answers.ActiveConnection = Connect
' SET 2 SOURCES: 1 FOR NEW RESULTS AND THE OTHER FOR OLD RESULTS AND RETRIEVE THE SOURCES BASED UPON THE datevariable AND session_finish
IF date_end_v1 < sessions.Fields.item("Session_finish").Value THEN
	answers.Source = "SELECT q_result.ID_result, q_question.ID_question, q_question.question_body, new_subjects.s_topic AS topic_name, new_subjects.s_qID, q_choice.choice_label, q_choice.choice_body, q_choice.choice_cor  FROM new_subjects INNER JOIN ((q_result INNER JOIN q_question ON q_result.result_question = q_question.ID_question) INNER JOIN q_choice ON (q_question.ID_question = q_choice.choice_question) AND (q_result.result_answer = q_choice.ID_choice)) ON new_subjects.s_ID = q_question.question_topic  WHERE result_session = " + Replace(answers__MMColParam, "'", "''") + " ORDER BY new_subjects.s_order, new_subjects.s_ID, q_result.ID_result;"
ELSE
	answers.Source = "SELECT q_result.ID_result, q_question_v1.ID_question, q_question_v1.question_body, q_topics.topic_name, q_choice_v1.choice_label, q_choice_v1.choice_body, q_choice_v1.choice_cor  FROM q_topics INNER JOIN ((q_result INNER JOIN q_question_v1 ON q_result.result_question = q_question_v1.ID_question) INNER JOIN q_choice_v1 ON (q_question_v1.ID_question = q_choice_v1.choice_question) AND (q_result.result_answer = q_choice_v1.ID_choice)) ON q_topics.ID_topic = q_question_v1.question_topic  WHERE result_session = " + Replace(answers__MMColParam, "'", "''") + " ORDER BY q_topics.topic_ord, q_topics.ID_topic, q_result.ID_result;"
END IF
answers.CursorType = 0
answers.CursorLocation = 3
answers.LockType = 3
answers.Open()
answers_numRows = 0
%>



<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz user session details. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--

function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</HEAD>
<BODY BGCOLOR=#FFCC00 TEXT=#000000 VLINK=#000000 LINK=#000000 leftmargin="0" topmargin="0">

<table width="100%" border="0" cellspacing="3" cellpadding="0">
  <tr>
    <td align="left" valign="bottom" class="headers"> Quiz user session details</td>
  </tr>
  <tr>
    <td align="left" valign="bottom">
      <form name="add_subject">
        <table border="0" cellspacing="2" cellpadding="3" width="600">
          <tr>
            <td colspan="4" class="subheads" align="left" valign="top"><%=(user.Fields.Item("user_firstname").Value)%>&nbsp;<%=(user.Fields.Item("user_lastname").Value)%>'s session in <%=subject_name%>, <%=(sessions.Fields.Item("Session_date").Value)%>:</td>
            <td align="right" class="subheads" valign="top"><a href="q_export_quiz_users_results.asp?user_session=<%=answers__MMColParam%>&user=<%=request("user")%>&passrate=<%=passrate%>"><img src="../admin/images/xls.gif" width="16" height="16" border="0"></a></td>
          </tr>
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">
            <td class="text" width="10"><img src="../admin/images/back.gif" width="18" height="14"></td>
            <td colspan="4" class="text"><a href="../admin/<% if request("combined") = 1 then %>q_comp_list_of_users_results.asp<% else %>q_user_sessions.asp<% end if%>?user=<%=(user.Fields.Item("ID_user").Value)%>&filter_username=<%=request("filter_username")%>&filter_info1=<%=request("filter_info1")%>&filter_info3=<%=request("filter_info3")%>&fromdate=<%=request("fromdate")%>&todate=<%=request("todate")%>&active=<%=request("active")%>&results=<%=request("results")%>&passrate=<%=passrate%>&mths=<%=request("mths")%>&noquiz=0&completedquiz=<%=request("completedquiz")%>">...go
              up one level to user's session list</a></td>
          </tr>
          <% If Not answers.EOF Or Not answers.BOF Then %>
          <tr>
            <td class="text">&nbsp;</td>
            <td class="text">Question</td>
            <td class="text">Topic</td>
            <td class="text">Answer</td>
            <td class="text">Cor.</td>
          </tr>
          <%
user_all = 0
user_correct = 0
%>
          <%
While (NOT answers.EOF)
%>
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">
            <td class="text" width="10"><%=numbers%></td>
            <td width="450" class="text">
			<!-- ADDED sess_id TO GET FOR q_question_details.asp -->
              <a href="javascript:" onClick="MM_openBrWindow('q_question_details.asp?qid=<%=(answers.Fields.Item("ID_question").Value)%>&sess_id=<%=(sessions.Fields.Item("ID_session").Value)%>','Questiondetails','scrollbars=yes,width=450,height=350')">
              <% =(CropSentence((answers.Fields.Item("question_body").Value), 50, "...")) %>
              </a> </td>
            <td width="100" class="text"><%=(answers.Fields.Item("topic_name").Value)%></td>
            <td width="30" class="text"><%=(answers.Fields.Item("choice_label").Value)%></td>
            <td width="20" class="text">
              <%
			if abs(answers.Fields.Item("choice_cor").Value) = 1 then
				response.write "<img src='images/1.gif'>"
				user_correct = user_correct +1
			else
				response.write "<img src='images/0.gif'>"
			end if
			user_all = user_all +1
			%>
            </td>
          </tr>
          <%
  answers.MoveNext()
  numbers=numbers+1
Wend
%>
          <tr>
            <td class="text" colspan="5">
              <hr>
            </td>
          </tr>
          <tr class="table_normal">
            <td class="text" width="10">&nbsp;</td>
            <td width="450" class="text">Overall
              Pass &amp; Rate:</td>
            <td width="100" class="text">Correct:</td>
            <td colspan="2" class="text">Incorrect:</td>
          </tr>
          <tr class="table_normal">
            <td class="text" width="10">&nbsp;</td>
            <td width="450" class="text">
            <% if cstr(sessions.fields.item("session_done").value) = "True" then
            	if (user_correct/user_all*100) >= cint(passrate) then
            		%><font color=green>PASSED - <%
                else
                	%><font color=red>FAILED - <%
            	end if
            end if %>
            <% =FormatNumber(user_correct/user_all*100, 2) %>%
            <% if cstr(sessions.fields.item("session_done").value) = "True" then%></font><% end if %></td>
            <td width="100" class="text"><%="<font color=green>" & user_correct & "</font>"%></td>
            <td colspan="2" class="text"><%="<font color=red>" & user_all-user_correct & "</font>"%></td>
          </tr>
          <tr>
            <td class="text_table" colspan="5"><i>Pass
              rate is: <%=passrate%> %</i></td>
          </tr>
          <% End If ' end Not answers.EOF Or NOT answers.BOF %>
          <% If answers.EOF And answers.BOF Then %>
          <tr>
            <td class="text_table">&nbsp;</td>
            <td colspan="4" class="text_table">Sorry,
              there are no user's answers in this session in the quiz currently.</td>
          </tr>
          <% End If ' end answers.EOF And answers.BOF %>
        </table>
      </form>
      <p>&nbsp;</p>
      </td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
call log_the_page ("Quiz Session details: " & sessions__MMColParam)
%>

<%
user.Close()
%>
<%
sessions.Close()
%>
<%
answers.Close()
%>
