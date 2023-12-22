<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
numbers=1
%>
<%
Dim top
If (Request.QueryString("topic") <> "") Then 
top = cInt(Request.QueryString("topic"))
Else 
'Response.Redirect("error.asp?" & request.QueryString) 
End If
%>
<%
Dim subj
If (Request.QueryString("subj") <> "") Then 
subj = cInt(Request.QueryString("subj"))
Else 
Response.Redirect("error.asp?" & request.QueryString) 
End If
%>
<%
set questions = Server.CreateObject("ADODB.Recordset")
questions.ActiveConnection = Connect
questions.Source = "SELECT q_question.ID_question, q_question.question_body, q_question.question_ord, q_question.question_active  FROM subjects INNER JOIN (q_topics INNER JOIN q_question ON q_topics.ID_topic = q_question.question_topic) ON subjects.ID_subject = q_topics.topic_subject  WHERE q_question.question_topic =" + Replace(top, "'", "''") + "  ORDER BY q_question.question_ord, q_question.ID_question;"
questions.CursorType = 0
questions.CursorLocation = 3
questions.LockType = 3
questions.Open()
questions_numRows = 0
'Response.Write questions.Source

%>
<%
set subject = Server.CreateObject("ADODB.Recordset")
subject.ActiveConnection = Connect
subject.Source = "SELECT subject_name  FROM subjects  WHERE ID_subject = " + Replace(subj, "'", "''") + ";"
subject.CursorType = 0
subject.CursorLocation = 3
subject.LockType = 3
subject.Open()
subject_numRows = 0
%>
<%
set topic = Server.CreateObject("ADODB.Recordset")
topic.ActiveConnection = Connect
topic.Source = "SELECT topic_name  FROM q_topics  WHERE ID_topic = " + Replace(top, "'", "''") + ";"
topic.CursorType = 0
topic.CursorLocation = 3
topic.LockType = 3
topic.Open()
topic_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
questions_numRows = questions_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Repeat2__numRows = -1
Dim Repeat2__index
Repeat2__index = 0
pages_stats_numRows = pages_stats_numRows + Repeat2__numRows
%>
<%
Dim Repeat3__numRows
Repeat3__numRows = -1
Dim Repeat3__index
Repeat3__index = 0
pages_suc_all_numRows = pages_suc_all_numRows + Repeat3__numRows
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz questions. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
//-->
</script>
</HEAD>
<BODY>
<table>
  <tr> 
    <td align="left" valign="bottom" class="headers"> Quiz questions list</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <table>
        <tr> 
          <td colspan="3" class="subheads">Questions in <%=(subject.Fields.Item("subject_name").Value)%> / <%=(topic.Fields.Item("topic_name").Value)%>:</td>
        </tr>
        <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')" onClick="document.location ='q_list_of_topics.asp?subj=<%=subj%>&topic=<%=top%>'"> 
          <td class="text" width="10"><img src="../admin/images/back.gif" width="18" height="14"></td>
          <td colspan="2" class="text"><a href="../admin/q_list_of_topics.asp?subj=<%=subj%>&topic=<%=top%>">...go 
            up one level to list of Topics</a></td>
        </tr>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT questions.EOF)) 
%>
        <% If Not questions.EOF Or Not questions.BOF Then %>
		
		<% 'if highlight_q true, then display different styled line
			If (Cstr(Request.QueryString("highlight_q")) = Cstr(questions.Fields.Item("ID_question").Value)) Then  %>		
		
				<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')" style="background-color: #FFFF33"> 
				  <td class="text" width="20"><%=numbers%></td>
				  <td width="560" class="text"><a href="../admin/q_question_edit.asp?qid=<%=(questions.Fields.Item("ID_question").Value)%>">(ID: <%=(questions.Fields.Item("ID_question").Value)%>) <% =(CropSentence((questions.Fields.Item("question_body").Value), 50, "...")) %></a></td>
				  <td width="20" class="text"> 
					<%if abs(questions.Fields.Item("question_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
				  </td>
				</tr>
		
		<%else%>	
		
				<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
			  <td class="text" width="20"><%=numbers%></td>
			  <td width="560" class="text"><a href="../admin/q_question_edit.asp?qid=<%=(questions.Fields.Item("ID_question").Value)%>">(ID: <%=(questions.Fields.Item("ID_question").Value)%>) <% =(CropSentence((questions.Fields.Item("question_body").Value), 50, "...")) %></a></td>
			  <td width="20" class="text"> 
				<%if abs(questions.Fields.Item("question_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
			  </td>
			</tr>		
		
		<%End if%>  
		
        <% End If ' end Not questions.EOF Or NOT questions.BOF %>
        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  questions.MoveNext()
  numbers=numbers+1
Wend
%>
        <% If questions.EOF And questions.BOF Then %>
        <tr> 
          <td class="text">&nbsp;</td>
          <td width="99%"  colspan="2">Sorry, 
            there are no questions in this topic currently.</td>
        </tr>
        <% End If ' end questions.EOF And questions.BOF %>
        <tr> 
          <td class="text"><img src="images/new2.gif" width="11" height="13"></td>
          <td width="99%" class="text" colspan="2"> 
            <input type="button" name="Button" value="Add a new question" onClick="document.location='q_question_add.asp?subj=<%=subj%>&topic=<%=top%>';" class="quiz_button">
          </td>
        </tr>
      </table>
      <p>&nbsp;</p>
    </td>
  </tr>

 <tr> 
   <td align="center" valign="bottom" class="subheads" colspan="7">
		<br><BR><a href="javascript:" onClick="window.open('q_stats_of_questions.asp?subj=<%=subj%>&topic=<%=top%>','statswindow','scrollbars=yes,resizable=yes,width=700,height=500,left=50,top=50')">Click to view question statistics</a>
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
