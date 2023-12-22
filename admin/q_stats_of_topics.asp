<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
' *** Edit Operations: declare variables

MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
numbers=1
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
set topics_stats = Server.CreateObject("ADODB.Recordset")
topics_stats.ActiveConnection = Connect
topics_stats.Source = "SELECT COUNT(logs.ID_log) AS ID_log, MAX(logs.log_date) AS log_date, q_topics.ID_topic, q_topics.topic_name FROM logs INNER JOIN (subjects INNER JOIN q_topics ON subjects.ID_subject = q_topics.topic_subject) ON logs.log_topicID = q_topics.ID_topic GROUP BY logs.log_module, subjects.ID_subject, q_topics.ID_topic, q_topics.topic_name HAVING (logs.log_module LIKE 'quiz') AND (subjects.ID_subject = " + Replace(subj, "'", "''") + ") ORDER BY COUNT(logs.ID_log) DESC;"
'SQL: "SELECT COUNT(logs.ID_log) AS ID_log, MAX(logs.log_date) AS log_date, q_topics.ID_topic, q_topics.topic_name FROM subjects INNER JOIN q_topics ON subjects.ID_subject = q_topics.topic_subject INNER JOIN logs ON q_topics.ID_topic = logs.log_topicID GROUP BY logs.log_module, subjects.ID_subject, q_topics.ID_topic, q_topics.topic_name HAVING (logs.log_module LIKE 'quiz') AND (subjects.ID_subject = " + Replace(subj, "'", "''") + ") ORDER BY COUNT(logs.ID_log) DESC;"
topics_stats.CursorType = 0
topics_stats.CursorLocation = 3
topics_stats.LockType = 3
topics_stats.Open()
topics_stats_numRows = 0
%>

<%
Dim Repeat2__numRows
Repeat2__numRows = -1
Dim Repeat2__index
Repeat2__index = 0
topics_stats_numRows = topics_stats_numRows + Repeat2__numRows
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz topics. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
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
    <td align="left" valign="bottom"> 
      <table>
        <tr> 
          <td colspan="3" class="subheads">Topics statistics:</td>
        </tr>
        <% If Not topics_stats.EOF Or Not topics_stats.BOF Then %>
        <tr> 
          <td width="99%"  colspan="3"> 
            <table>
              <tr valign="middle"> 
                <td width="30%" align="right"><b>Topic name</b></td>
                <td align="left"><b></b></td>
                <td width="130" align="left"><b>last visit</b></td>
              </tr>
              <%
bar_max_length = 0
%>
              <% 
While ((Repeat2__numRows <> 0) AND (NOT topics_stats.EOF)) 
%>
              <%
bar_current_length = (topics_stats.Fields.Item("ID_log").Value)
if cint(bar_current_length) > cint(bar_max_length) then bar_max_length = cint(bar_current_length)
%>
              <tr valign="middle"> 
                <td width="30%" align="right"><%=(topics_stats.Fields.Item("topic_name").Value)%> - </td>
                <td align="left">&nbsp;<img src="images/bar0.gif"><img src="images/bar1.gif" height ="9" width="<%=cInt(bar_current_length/bar_max_length*stat_bar_length)%>" ><img src="images/bar2.gif"> 
                  (<%=bar_current_length%>)</td>
                <td width="130" align="left"><%=(topics_stats.Fields.Item("log_date").Value)%></td>
              </tr>
              <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  topics_stats.MoveNext()
Wend
%>
            </table>
          </td>
        </tr>
        <% End If ' end Not topics_stats.EOF Or NOT topics_stats.BOF %>
        <tr> 
          <% If topics_stats.EOF And topics_stats.BOF Then %>
          <td width="99%"  colspan="3">Sorry, 
            there are no statistics available yet...</td>
          <% End If ' end topics_stats.EOF And topics_stats.BOF %>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
call log_the_page ("Quiz Topics Statistics " & subj)
%>


<%
topics_stats.Close()
%>
