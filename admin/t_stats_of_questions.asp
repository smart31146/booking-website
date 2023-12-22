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
Response.Redirect("error.asp?" & request.QueryString) 
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
pages_stats.Source = "SELECT COUNT(logs.ID_log) AS ID_log, MAX(logs.log_date) AS log_date, tr_pages.ID_page, tr_pages.page_title, tr_topics.ID_topic FROM logs INNER JOIN (tr_topics INNER JOIN tr_pages ON tr_topics.ID_topic = tr_pages.page_topic) ON logs.log_pageID = tr_pages.ID_page GROUP BY logs.log_module, tr_pages.ID_page, tr_pages.page_title, tr_topics.ID_topic, tr_pages.page_active HAVING (logs.log_module LIKE 'training') AND (tr_topics.ID_topic = " + Replace(top, "'", "''") + ") AND tr_pages.page_active = 1ORDER BY COUNT(logs.ID_log) DESC;"
'SQL: "SELECT COUNT(logs.ID_log) AS ID_log, MAX(logs.log_date) AS log_date, tr_pages.ID_page, tr_pages.page_title, tr_topics.ID_topic FROM tr_pages INNER JOIN logs ON tr_pages.ID_page = logs.log_pageID INNER JOIN tr_topics ON tr_pages.page_topic = tr_topics.ID_topic GROUP BY logs.log_module, tr_pages.ID_page, tr_pages.page_title, tr_topics.ID_topic, tr_pages.page_active HAVING (logs.log_module LIKE 'training') AND (tr_topics.ID_topic = " + Replace(top, "'", "''") + ")  AND tr_pages.page_active = 1 ORDER BY COUNT(logs.ID_log) DESC;"
pages_stats.CursorType = 0
pages_stats.CursorLocation = 3
pages_stats.LockType = 3
pages_stats.Open()
pages_stats_numRows = 0
%>
<%
Dim Repeat2__numRows
Repeat2__numRows = -1
Dim Repeat2__index
Repeat2__index = 0
pages_stats_numRows = pages_stats_numRows + Repeat2__numRows
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Training pages. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
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
          <td colspan="3" class="subheads">Pages statistics:</td>
        </tr>
        <% If Not pages_stats.EOF Or Not pages_stats.BOF Then %>
        <tr> 
          <td width="99%"  colspan="3"> 
            <table>
              <tr valign="middle"> 
                <td width="30%" align="right"><b>Page name</b></td>
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
                <td width="30%" align="right"><%=(pages_stats.Fields.Item("page_title").Value)%> - </td>
                <td align="left">&nbsp;<img src="images/bar0.gif"><img src="images/bar1.gif" height ="9" width="<%=cInt(bar_current_length/bar_max_length*stat_bar_length)%>" ><img src="images/bar2.gif"> 
                  (<%=bar_current_length%>)</td>
                <td width="130" align="left"><%=(pages_stats.Fields.Item("log_date").Value)%></td>
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
</table>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
call log_the_page ("Training Pages Stats: " & subj & ", " & top)
%>


<%
pages_stats.Close()
%>
