<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
numbers=1

	
	if Request.Cookies("show_lines1")<> "" then
		show_lines1= cint(Request.Cookies("show_lines1"))
	else
		show_lines1=15
	end if
	%>

<%
'Dim top
'If (Request.QueryString("topic") <> "") Then 
'top = cInt(Request.QueryString("topic"))
'Else 
'Response.Redirect("error.asp?" & request.QueryString) 
'End If
%>
<%
Dim subj
if cStr(Request.Querystring("show_lines1")) <> "" then show_lines1 = cInt(Request.Querystring("show_lines1"))
If (Request.QueryString("subj") <> "") Then 
subj =cInt(Request.QueryString("subj"))
Else 
subj = 0
'Response.Redirect("error.asp?" & request.QueryString) 
End If
%>
<%
set questions = Server.CreateObject("ADODB.Recordset")
questions.ActiveConnection = connect
'questions.Source = "SELECT q_question.ID_question, q_question.question_body, q_question.question_ord, q_question.question_active  FROM subjects INNER JOIN (q_topics INNER JOIN q_question ON q_topics.ID_topic = q_question.question_topic) ON subjects.ID_subject = q_topics.topic_subject  WHERE q_question.question_topic =" + Replace(top, "'", "''") + "  ORDER BY q_question.question_ord, q_question.ID_question;"
questions.Source= "select * from q_question a , subjects b, new_subjects c where c.s_qid = b.id_subject and c.s_id = a.question_topic and c.s_qid="&subj&";"
questions.CursorType = 0
questions.CursorLocation = 3
questions.LockType = 3
questions.Open()
questions_numRows = 0
%>
<%
set subject = Server.CreateObject("ADODB.Recordset")
subject.ActiveConnection = connect
subject.Source = "SELECT subject_name  FROM subjects  WHERE ID_subject = " + Replace(subj, "'", "''") + ";"
subject.CursorType = 0
subject.CursorLocation = 3
subject.LockType = 3
subject.Open()
subject_numRows = 0
%>

<%
set topic = Server.CreateObject("ADODB.Recordset")
topic.ActiveConnection = connect
'topic.Source = "SELECT topic_name  FROM q_topics  WHERE ID_topic = " + Replace(top, "'", "''") + ";"
topic.Source = "SELECT id_topic  FROM q_topics  WHERE topic_subject="&subj&";"
topic.CursorType = 0
topic.CursorLocation = 3
topic.LockType = 3
topic.Open()
topic_numRows = 0

%>


<%
set pages_suc_all = Server.CreateObject("ADODB.Recordset")
pages_suc_all.ActiveConnection = connect
pages_suc_all.Source = "SELECT q_result.result_question, COUNT(q_result.ID_result) AS ID_result,q_question.question_topic FROM q_result INNER JOIN q_question ON q_result.result_question = q_question.ID_question GROUP BY q_question.question_topic, q_result.result_question, q_question.ID_question HAVING (q_question.question_topic in (SELECT s_id  FROM new_subjects  WHERE s_qid="&subj&"))  ORDER BY q_result.result_question;"
'pages_suc_all.Source = "select * from q_question a , subjects b, q_topics c where c.topic_subject = b.id_subject and c.id_topic = a.question_topic and c.topic_subject=1"
'Response.Write pages_suc_all.Source
pages_suc_all.CursorType = 0
pages_suc_all.CursorLocation = 3
pages_suc_all.LockType = 3
pages_suc_all.Open()
pages_suc_all_numRows = 0
%>


<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="content-type" content="text/html; charset=UTF-8">
<TITLE>BBP ADMIN: Quiz questions. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
var MyCookie = {
    Write:function(name,value,days) {
        var D = new Date();
        D.setTime(D.getTime()+86400000*days)
        document.cookie = escape(name)+"="+escape(value)+
            ((days == null)?"":(";expires="+D.toGMTString()))
        return (this.Read(name) == value);
    },
    Read:function(name) {
        var EN=escape(name)
        var F=' '+document.cookie+';', S=F.indexOf(' '+EN);
        return S==-1 ? null : unescape(F.substring(EN=S+EN.length+2,F.indexOf(';',EN)));
    }
} 
function check()
{
	show_lines1 = MyCookie.Read('show_lines1')
	if (show_lines1 != null) {
		document.forms[0].show_lines1.value=show_lines1;
	}
	else {
		document.forms[0].show_lines1.value=15;
	}
}

function show()
{
	if (isNaN(document.questions.show_lines1.value)){
		alert('Invalid number');
		document.questions.show_lines1.focus();
		return false;
	}
	else
	{
		AddCookieId("show_lines1",document.questions.show_lines1.value);
		show_lines1 = MyCookie.Read('show_lines1')
		document.forms[0].submit();
		return true;
	}
}

function AddCookieId(cn,id) {
        MyCookie.Write(cn,id,7);
}

function DelCookieId(cn,id) {
        MyCookie.Write(cn,id,-1);
}
//-->
</script>
</HEAD>
<BODY BGCOLOR=#FFFFF TEXT=#000000 VLINK=#000000 LINK=#000000 leftmargin="0" topmargin="0" onload="check();">
<%
Response.Clear()

Response.AddHeader "Content-Disposition","inline; filename=quiz_user" & day(now()) & "_" & month(now()) & "_" & year(now()) & ".csv"

Response.ContentType = "application/vnd.ms-excel"
%>
"Questions stats"<%=vbcrlf%><%=vbcrlf%>
<%=subject.Fields.item("subject_name") %>

"ID","Correct","Incorrect","%Correct","Active","Topic Name","Question","Answers"<%=vbcrlf%>"------------------------------------------------------------------------------------------------------------------------------------------------------------"<%=vbcrlf%>



    	    <% 
While (NOT questions.EOF) 
%>
            <% If Not questions.EOF Or Not questions.BOF Then 
			overall_correct = 0
overall_incorrect = 0
			%>
			
			
				 
				    <%
set pages_suc_ok = Server.CreateObject("ADODB.Recordset")
pages_suc_ok.ActiveConnection = connect
pages_suc_ok.Source = "SELECT COUNT(q_result.ID_result) AS ID_result FROM q_result INNER JOIN q_choice ON q_result.result_answer = q_choice.ID_choice GROUP BY q_result.result_question, q_choice.choice_cor HAVING (q_result.result_question = " & cInt(questions.Fields.Item("ID_question").Value) & ") AND (q_choice.choice_cor = 1);"
'Response.Write pages_suc_ok.Source
pages_suc_ok.CursorType = 0
pages_suc_ok.CursorLocation = 3
pages_suc_ok.LockType = 3
pages_suc_ok.Open()
pages_suc_ok_numRows = 0

set pages_incorrect = Server.CreateObject("ADODB.Recordset")
pages_incorrect.ActiveConnection = connect
pages_incorrect.Source = "SELECT COUNT(q_result.ID_result) AS ID_result FROM q_result INNER JOIN q_choice ON q_result.result_answer = q_choice.ID_choice GROUP BY q_result.result_question, q_choice.choice_cor HAVING (q_result.result_question = " & cInt(questions.Fields.Item("ID_question").Value) & ") AND (q_choice.choice_cor = 0);"
'Response.Write pages_incorrect.Source
pages_incorrect.CursorType = 0
pages_incorrect.CursorLocation = 3
pages_incorrect.LockType = 3
pages_incorrect.Open()
pages_incorrect_numRows = 0
%>
<%if not pages_suc_ok.EOF and Not pages_suc_ok.BOF and not pages_incorrect.EOF then 
pages_correct = (pages_suc_ok.Fields.Item("ID_result").Value) 
pages_no=(pages_incorrect.Fields.Item("ID_result").Value) 
dim perc

perc=Round((pages_correct/(pages_correct+pages_no)) *100)

 
 %>
 
<%end if %>
              <%
pages_suc_ok.Close()
pages_incorrect.Close()
%>
<%=(questions.Fields.Item("ID_question").Value)%>,<%=pages_correct %>,<%=pages_no %>, <%=perc %> %, <% if questions.Fields.Item("s_active")  then if questions.Fields.Item("question_active") then response.write "YES" else response.write "NO"  end if else response.write "NO" end if%>,<%=(questions.Fields.Item("s_topic").Value)%>, <% = (questions.Fields.Item("question_body").Value) %>,"",<%
set choice = Server.CreateObject("ADODB.Recordset")
choice.ActiveConnection = connect
choice.Source = "SELECT *  FROM q_choice  WHERE choice_question ="&questions.Fields.Item("ID_question").Value&" and choice_active=1 order by choice_label;"
choice.CursorType = 0
choice.CursorLocation = 3
choice.LockType = 3
choice.Open()
%><%While (NOT choice.EOF) %><%=vbcrlf%>"","","","","","",<%=(choice.Fields.Item("choice_label").Value)%><% = (choice.Fields.Item("choice_body").Value) %> 
        <% 
		choice.MoveNext()
		Wend %>
		
<%End If %>
<%=vbcrlf%>
        <% 
		
  questions.MoveNext()
  Wend
%>

        <% If questions.EOF And questions.BOF Then %>
        
          
          Sorry, there are no questions in this topic currently.</td>
        
        <% End If ' end questions.EOF And questions.BOF %>
       
   <%=vbcrlf%>"-------------------------------------------------------------------------------------------------------------------------"<%=vbcrlf%>"Generated on:","<%=Now()%>"<%=vbcrlf%><%=vbcrlf%>"Copyright 2002 - 2011 (c) Law of the Jungle Pty Limited"

</BODY>
</HTML>

<%
call log_the_page ("Quiz List Questions: " & subj& ", " & top)
%>

<%
questions.Close()
%>
<%
subject.Close()
%>
<%
topic.Close()
%>


