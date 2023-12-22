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
set questions_count = Server.CreateObject("ADODB.Recordset")
questions_count.ActiveConnection = connect
'questions.Source = "SELECT q_question.ID_question, q_question.question_body, q_question.question_ord, q_question.question_active  FROM subjects INNER JOIN (q_topics INNER JOIN q_question ON q_topics.ID_topic = q_question.question_topic) ON subjects.ID_subject = q_topics.topic_subject  WHERE q_question.question_topic =" + Replace(top, "'", "''") + "  ORDER BY q_question.question_ord, q_question.ID_question;"
questions_count.Source= "select  COUNT(*) as total from q_question a , subjects b, new_subjects c where c.s_qid = b.id_subject and c.s_id = a.question_topic and c.s_qid="&subj&";"
questions_count.CursorType = 0
questions_count.CursorLocation = 3
questions_count.LockType = 3
questions_count.Open()

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
<TITLE>BBP ADMIN: Quiz questions. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<!--[if lt IE 9]>
<script src="https://ie7-js.googlecode.com/svn/version/2.1(beta4)/IE9.js?v=bbp34"></script>
<![endif]-->


<!--<style>
#main_tb tr:nth-child(even) {
    background-color: #fae79b;
	
}

#sub_tb tr:nth-child(even) {
    background-color: transparent;
}
#sub_tb tr:nth-child(odd) {
    background-color: transparent;
}
</style>-->

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
<BODY BGCOLOR=#FFCC00 TEXT=#000000 VLINK=#000000 LINK=#000000 leftmargin="0" topmargin="0" onLoad="check();">
<table width=30%>
 <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')" onClick="document.location ='q_list_of_subjects.asp'"> 
            <td class="text" width="10"><img src="../admin/images/back.gif" width="18" height="14"></td>
            <td class="text" colspan="6"><a href="../admin/q_list_of_subjects_stats.asp">...go 
              up one level to list of Subjects</a></td>
          </tr>
</table>
<form name="questions">
<br>
<div style="width:600px">
<table border=0  >
<tr>
<td>

<div style="font-size:28px; width:400" ><%=subject.Fields.item("subject_name") %></div>
<input type="hidden" name="subj" value=<%=request("subj")%>>
</td>
<td   align="right" width="300" ><a style="text-decoration:none;" target="_blank" href='question_stats_report_export.asp?subj=<%=request("subj")%>' ><img src="../admin/images/xls.gif" align="top" width="16" height="16" border="0"> Export to CSV format</a>
</td>
</tr>
</table>
</div>
<div style="width:200px">
<table border="1" width="150" cellpadding="2" >
<tr><td colspan="2">% Correct</td></tr>
<tr><td height="20" width="20" bgcolor="#FF8D71"></td><td>0-49 </td></tr>
<tr><td height="20" width="20" bgcolor="#E6E65C"> </td><td>50-69 </td></tr>
<tr><td height="20" width="20" bgcolor="#85E085"></td><td>70-100 </td></tr>
</table>
</div>
<br>
Total questions: <%=questions_count.Fields.Item("total") %>
<br>
<br>
<div style="width:95%;margin-left:5px;">
<table width="100%" border="1" cellspacing="0" cellpadding="2" class="table_normal" id="main_tb" >
<tr>

<td class="text" > <strong>ID</strong></td><td class="text" ><font color="green"><strong>Correct </strong></font></td><td class="text" ><font color="red"><strong>Incorrect </strong></font></td><td width=75 class="text" ><font color="blue"><strong>% Correct </strong></font></td><td style="font-weight:bold;" class="text" align="center">Active</td><td style="font-weight:bold;" class="text" >Questions</td><td style="font-weight:bold;"class="text" >Answers</td></tr>
  
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
<tr  <% if perc>69 then response.write "bgcolor=#85E085" else if perc>49 and perc<70 then response.write "bgcolor=#E6E65C"  else if perc>-1 and perc<50 then response.write "bgcolor=#FF8D71" end if end if end if  %>   > 
<td class="text" ><%=(questions.Fields.Item("ID_question").Value)%></td>

<td class="text" align="center"><%=pages_correct %></td><td class="text" align="center"><%=pages_no %></td>
<td class="text" align="center" >

 <%=perc %> % 
 </td>
 <td align="center" width="75"><% if questions.Fields.Item("s_active")  then if questions.Fields.Item("question_active") then response.write "<img src='images/1.gif' alt='Active'>" else response.write "<img src='images/0.gif' alt='Inactive'>"  end if else response.write "<img src='images/0.gif' alt='Inactive'>" end if%></td>
				 <td width="560" class="text">
				 <strong><%=(questions.Fields.Item("s_topic").Value)%></strong>
				 <br><br>
				   <% = (questions.Fields.Item("question_body").Value) %>
				<br><br>
				
				 </td>
				 
				 <%
set choice = Server.CreateObject("ADODB.Recordset")
choice.ActiveConnection = connect
choice.Source = "SELECT *  FROM q_choice  WHERE choice_question ="&questions.Fields.Item("ID_question").Value&" and choice_active=1 order by choice_label;"
choice.CursorType = 0
choice.CursorLocation = 3
choice.LockType = 3
choice.Open()
%>
<td>
<table id="sub_tb">

<%
While (NOT choice.EOF) 
%>
<tr>
   <td class="text"><% if choice.Fields.Item("choice_cor") then %><strong style="color:#006600;"><%=(choice.Fields.Item("choice_label").Value)%></strong> <% else %><strong style="color:red;"><%=(choice.Fields.Item("choice_label").Value)%></strong> <% end if %>  <% = (choice.Fields.Item("choice_body").Value) %> </td> 
   </tr>
        <% 
		choice.MoveNext()
		Wend %>
</table>
</td>
</tr>		
<%End If %>
        <% 
  questions.MoveNext()
  Wend
%>
        <% If questions.EOF And questions.BOF Then %>
        <tr> 
          
          <td>Sorry, there are no questions in this topic currently.</td>
        </tr>
        <% End If ' end questions.EOF And questions.BOF %>
       
      </table>
	  </div>
</form>
</BODY>
</HTML>

<%
call log_the_page ("Quiz List Questions: " & subj& ", " & top)
%>

<%
questions_count.Close()
questions.Close()
%>
<%
subject.Close()
%>



