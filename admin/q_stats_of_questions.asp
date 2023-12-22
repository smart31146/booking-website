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
set pages_stats = Server.CreateObject("ADODB.Recordset")
pages_stats.ActiveConnection = Connect
pages_stats.Source = "SELECT COUNT(logs.ID_log) AS ID_log, MAX(logs.log_date) AS log_date, q_question.ID_question, q_question.ID_question FROM logs INNER JOIN (q_topics INNER JOIN q_question ON q_topics.ID_topic = q_question.question_topic) ON logs.log_pageID = q_question.ID_question GROUP BY logs.log_module, q_question.ID_question, q_topics.ID_topic, q_question.ID_question HAVING (logs.log_module LIKE 'quiz') AND (q_topics.ID_topic = " + Replace(top, "'", "''") + ") ORDER BY COUNT(logs.ID_log) DESC;"
'SQL: "SELECT COUNT(logs.ID_log) AS ID_log, MAX(logs.log_date) AS log_date, q_question.ID_question, q_question.ID_question FROM logs INNER JOIN q_topics INNER JOIN q_question ON q_topics.ID_topic = q_question.question_topic ON logs.log_pageID = q_question.ID_question GROUP BY logs.log_module, q_question.ID_question, q_topics.ID_topic, q_question.ID_question HAVING (logs.log_module LIKE 'quiz') AND (q_topics.ID_topic = " + Replace(top, "'", "''") + ") ORDER BY COUNT(logs.ID_log) DESC;"
pages_stats.CursorType = 0
pages_stats.CursorLocation = 3
pages_stats.LockType = 3
pages_stats.Open()
pages_stats_numRows = 0
%>
<%
set pages_suc_all = Server.CreateObject("ADODB.Recordset")
pages_suc_all.ActiveConnection = Connect
pages_suc_all.Source = "SELECT q_result.result_question, COUNT(q_result.ID_result) AS ID_result FROM q_result INNER JOIN q_question ON q_result.result_question = q_question.ID_question GROUP BY q_result.result_question, q_question.question_topic HAVING (q_question.question_topic = " + Replace(top, "'", "''") + ") ORDER BY q_result.result_question;"
pages_suc_all.CursorType = 0
pages_suc_all.CursorLocation = 3
pages_suc_all.LockType = 3
pages_suc_all.Open()
pages_suc_all_numRows = 0
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
      <table>
        <tr> 
          <td colspan="3" class="subheads">Questions statistics:</td>
        </tr>
        <% If Not pages_stats.EOF Or Not pages_stats.BOF Then %>
        <tr> 
          <td width="99%"  colspan="3"> 
            <table>
              <tr valign="middle"> 
                <td width="30%" align="right"><b>Question ID</b></td>
                <td align="left"><b></b></td>
                <td width="130" align="left"><b>last visit</b></td>
              </tr>
              <%
bar_max_length = 0
%>
              <% 
While ((Repeat2__numRows <> 0) AND (NOT pages_stats.EOF)) 
%>
              <%
bar_current_length = (pages_stats.Fields.Item("ID_log").Value)
if cint(bar_current_length) > cint(bar_max_length) then bar_max_length = cint(bar_current_length)
%>
              <tr valign="middle"> 
                <td width="30%" align="right" class="table_normal"><%=(pages_stats.Fields.Item("ID_question").Value)%> - </td>
                <td align="left" class="table_normal">&nbsp;<img src="images/bar0.gif"><img src="images/bar1.gif" height ="9" width="<%=cInt(bar_current_length/bar_max_length*stat_bar_length)%>" ><img src="images/bar2.gif"> 
                  (<%=bar_current_length%>)</td>
                <td width="130" align="left" class="table_normal"><%=(pages_stats.Fields.Item("log_date").Value)%></td>
              </tr>
              <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  pages_stats.MoveNext()
Wend
%>
            </table>
          </td>
        </tr>
        <% End If ' end Not pages_stats.EOF Or NOT pages_stats.BOF %>
        <tr> 
          <% If pages_stats.EOF And pages_stats.BOF Then %>
          <td width="99%"  colspan="3">Sorry, 
            there are no statistics available yet...</td>
          <% End If ' end pages_stats.EOF And pages_stats.BOF %>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <table>
        <tr> 
          <td class="subheads">Inorrect/correct ratio</td>
        </tr>
        <% If Not pages_suc_all.EOF Or Not pages_suc_all.BOF Then %>
        <tr> 
          <td> 
            <table>
              <%
overall_correct = 0
overall_incorrect = 0
%>
              <% 
While ((Repeat3__numRows <> 0) AND (NOT pages_suc_all.EOF)) 
%>
              <%
pages_all = (pages_suc_all.Fields.Item("ID_result").Value)
%>
              <%
set pages_suc_ok = Server.CreateObject("ADODB.Recordset")
pages_suc_ok.ActiveConnection = Connect
pages_suc_ok.Source = "SELECT COUNT(q_result.ID_result) AS ID_result FROM q_result INNER JOIN q_choice ON q_result.result_answer = q_choice.ID_choice GROUP BY q_result.result_question, q_choice.choice_cor HAVING (q_result.result_question = " & cInt(pages_suc_all.Fields.Item("result_question").Value) & ") AND (q_choice.choice_cor = 1);"
pages_suc_ok.CursorType = 0
pages_suc_ok.CursorLocation = 3
pages_suc_ok.LockType = 3
pages_suc_ok.Open()
pages_suc_ok_numRows = 0
%>
              <%
if not pages_suc_ok.EOF or Not pages_suc_ok.BOF Then pages_correct = (pages_suc_ok.Fields.Item("ID_result").Value) else pages_correct = 0
pages_incorrect = cint(pages_all) - cint(pages_correct)
overall_correct = cint(overall_correct) + cint(pages_correct)
overall_incorrect = cint(overall_incorrect) + cint(pages_incorrect)
%>
              <tr valign="middle" class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')" onClick="document.location ='q_question_edit.asp?qid=<%=(pages_suc_all.Fields.Item("result_question").Value)%>'"> 
                <td width="280" align="right">(<%=pages_incorrect%>)<img src="images/bari2.gif"><img src="images/bari1.gif" width="<%=cInt(pages_incorrect/pages_all*stat_bar_length)%>" height="9"><img src="images/bari0.gif"></td>
                <td width="40" align="center">-&nbsp;<%=(pages_suc_all.Fields.Item("result_question").Value)%>&nbsp;-</td>
                <td width="280" align="left"><img src="images/barc0.gif"><img src="images/barc1.gif" width="<%=cInt(pages_correct/pages_all*stat_bar_length)%>" height="9"><img src="images/barc2.gif">(<%=pages_correct%>)</td>
              </tr>
              <%
pages_suc_ok.Close()
%>
              <% 
  Repeat3__index=Repeat3__index+1
  Repeat3__numRows=Repeat3__numRows-1
  pages_suc_all.MoveNext()
Wend
%>
            </table>
          </td>
        </tr>
        <tr> 
          <td>
            <table>
              <tr valign="middle"> 
                <td width="280"  align="right">(<%=overall_incorrect%>)<img src="images/bari2.gif"><img src="images/bari1.gif" width="<%=cInt(overall_incorrect/(overall_incorrect+overall_correct)*stat_bar_length)%>" height="9"><img src="images/bari0.gif"></td>
                <td width="40"  align="center">-Total-</td>
                <td width="280"  align="left"><img src="images/barc1.gif" width="<%=cInt(overall_correct/(overall_incorrect+overall_correct)*stat_bar_length)%>" height="9"><img src="images/barc2.gif">(<%=overall_correct%>)</td>
              </tr>
            </table>
          </td>
        </tr>
        <% End If ' end Not pages_suc_all.EOF Or NOT pages_suc_all.BOF %>
        <tr> 
          <% If pages_suc_all.EOF And pages_suc_all.BOF Then %>
          <td >Sorry, 
            there are no results available yet...</td>
          <% End If ' end pages_suc_all.EOF And pages_suc_all.BOF %>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
call log_the_page ("Quiz Question Stats: " & subj& ", " & top)
%>

<%
pages_stats.Close()
%>
<%
pages_suc_all.Close()
%>

