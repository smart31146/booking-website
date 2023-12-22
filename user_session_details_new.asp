<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="connections/bbg_conn.asp" -->
<!--#include file="connections/include.asp" -->

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
set sessions = Server.CreateObject("ADODB.Recordset")
sessions.ActiveConnection = Connect
sessions.Source = "SELECT q_session.ID_Session, q_session.Session_date, subjects.subject_name, q_session.Session_total, q_session.Session_done, q_session.Session_correct, q_certification.percentage_required  FROM q_user INNER JOIN (q_session INNER JOIN subjects ON q_session.Session_subject = subjects.ID_subject) ON q_user.ID_user = q_session.Session_users INNER JOIN q_certification ON q_certification.q_session=q_session.id_session WHERE Session_users = " + Replace(sessions__MMColParam, "'", "''") + "  ORDER BY q_session.Session_date DESC , subjects.subject_name DESC;"
sessions.CursorType = 0
sessions.CursorLocation = 3
sessions.LockType = 3
sessions.Open()
sessions_numRows = 0
%>

<%
subj = cint(request("subject"))
set subject = Server.CreateObject("ADODB.Recordset")
subject.ActiveConnection = Connect
subject.Source = "SELECT subject_name from subjects where id_subject=" & subj
subject.CursorType = 0
subject.CursorLocation = 3
subject.LockType = 3
subject.Open()
subject_numRows = 0
subject_name = subject("subject_name")
subject.Close
%>

<%
Dim answers__MMColParam
answers__MMColParam = "1"
if (Request.QueryString("user_session") <> "") then answers__MMColParam = Request.QueryString("user_session")
%>
<%
set answers = Server.CreateObject("ADODB.Recordset")
answers.ActiveConnection = Connect
answers.Source = "SELECT q_result.ID_result, q_question.ID_question, q_question.question_body, q_choice.choice_cor from q_question INNER JOIN q_result ON q_result.result_question = q_question.ID_question  INNER JOIN q_choice ON q_choice.choice_question=q_question.ID_question AND q_result.result_answer = q_choice.ID_choice WHERE result_session =" + Replace(answers__MMColParam, "'", "''")  + " ORDER BY  q_result.ID_result;"

'answers.Source = "SELECT q_result.ID_result, q_question.ID_question, q_question.question_body, q_topics.topic_name, q_choice.choice_label, q_choice.choice_body, q_choice.choice_cor  FROM q_topics INNER JOIN ((q_result INNER JOIN q_question ON q_result.result_question = q_question.ID_question) INNER JOIN q_choice ON (q_question.ID_question = q_choice.choice_question) AND (q_result.result_answer = q_choice.ID_choice)) ON q_topics.ID_topic = q_question.question_topic  WHERE result_session =2504 ORDER BY q_topics.topic_ord, q_topics.ID_topic, q_result.ID_result;"
answers.CursorType = 0
answers.CursorLocation = 3
answers.LockType = 3
answers.Open()
answers_numRows = 0
%>
<!doctype html>
<HTML>
<HEAD>

<TITLE>BBP ADMIN: Quiz user session details. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<style>
body    {overflow-x:scroll;}
body    {overflow-y:scroll;}
body    {overflow:scroll;}
</style>
<link rel="stylesheet" href="style/bbp_acme34.css" type="text/css">
<link rel="stylesheet" href="style/bbp_style_acme34.css" type="text/css">

<script >
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
<BODY  class=bodyresults>
<%
	'if Request.Cookies("passrate")<> "" then
	'	passrate= cint(Request.Cookies("passrate"))
	'else
	'	passrate=50
	'end if

	passrate=sessions.Fields.Item("percentage_required").Value
%>

<table >
	<tr>
		<!--<td width="190">&nbsp;</td>-->
		<td>
		
		<table >
		  <tr>
			<td  style="background:#fff; border:#fff 0px solid">
			  <form name="add_subject">
			  <div class="CSSTableGenerator">
				<table  >
				  <tr>
					<td colspan="3" class="subheadsuserdb"  ><%=(user.Fields.Item("user_firstname").Value)%>&nbsp;<%=(user.Fields.Item("user_lastname").Value)%>'s session in <%=subject_name%>, <%=(sessions.Fields.Item("Session_date").Value)%>:</td>
					<!--<td align="right" class="subheadsuserdb" ><br><br></td>-->
				  </tr>
				  <tr class="table_normaluserdb" onMouseOver="pviiClassNew(this,'table_hluserdb')" onMouseOut="pviiClassNew(this,'table_normaluserdb')">
					<td class="textuserdb" ><a class="quiz" href="user_sessions_new.asp?user=<%=(user.Fields.Item("ID_user").Value)%>&amp;latest=<%=request("latest")%>"><img src="images/return.png" alt=""></a></td>
					<td colspan="2" class="textuserdb" ><a class="quiz" href="user_sessions_new.asp?user=<%=(user.Fields.Item("ID_user").Value)%>&amp;latest=<%=request("latest")%>">...GO BACK</a></td>
				  </tr>
				  <% If Not answers.EOF Or Not answers.BOF Then %>
				  <tr>
					<td class="textuserdb">&nbsp;</td>
					<td class="textuserdb" ><b>Question</b></td>
					<!--<td class="textuserdb"><b>Topic</b></td>-->
					<td class="textuserdb"><b>Correct</b></td>
				  </tr>
				  <%
		user_all = 0
		user_correct = 0
		%>
				  <%
		While (NOT answers.EOF)
		%>
				  <tr  class="table_normaluserdb" onMouseOver="pviiClassNew(this,'table_hluserdb')" onMouseOut="pviiClassNew(this,'table_normaluserdb')">
					<td class="textuserdb" ><%=numbers%></td>
					<td  class="textuserdb" >
					  <% =((answers.Fields.Item("question_body").Value)) %>
					  </td>
					
					<td  class="textuserdb">
					  <%
					if abs(answers.Fields.Item("choice_cor").Value) = 1 then
						response.write "<img src='images/yes.png' alt=''>"
					user_correct = user_correct +1
					else
						response.write "<img src='images/no.png' alt=''>"
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
					<td class="textuserdb" colspan="3">
					  <hr>
					</td>
				  </tr>
				  <tr class="table_normaluserdb">
				  <td class="textuserdb" colspan=3   > 
				  <table >
				  <tr >
					
					<td class="textuserdb" >Result:</td>
					<td  class="textuserdb">Correct:</td>
					<td class="textuserdb">Incorrect:</td>
					</tr>
					<tr>
					
					<td class="textuserdb"  style="background:#fff;"><%if (user_correct/user_all*100) >= cint(passrate) then response.write("<div style='color:green'>PASSED - " & FormatNumber(user_correct/user_all*100, 2) & "%</div>") else response.write("<div style='color:red'>FAILED - " & FormatNumber(user_correct/user_all*100, 2) & "%</div>")%></td>
					<td  class="textuserdb"  style="background:#fff;"><%="<div style='color:green;'>" & user_correct & "</div>"%></td>
					<td class="textuserdb"  style="background:#fff;"><%="<div style='color:red;'>" & user_all-user_correct & "</div>"%></td>
				 </tr>
					</table>
					</td>
				  </tr>
				  
				  <!--<tr>
					<td class="text_tableuserdb" colspan="4">&nbsp;</td>
				  </tr>-->
				  <% End If ' end Not answers.EOF Or NOT answers.BOF %>
				  <% If answers.EOF And answers.BOF Then %>
				  <tr>
					<!--<td class="text_tableuserdb">&nbsp;</td>-->
					<td colspan="4" class="text_tableuserdb">Sorry,
					  there are no user's answers in this session in the quiz currently.</td>
				  </tr>
				  <% End If ' end answers.EOF And answers.BOF %>
				</table>
				</div>
			  </form>
			  
			  </td>
		  </tr>
		</table>
		
		</td>
	</tr>
</table>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
'call log_the_page ("Quiz Session details: " & sessions__MMColParam)
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
