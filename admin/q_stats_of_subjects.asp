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
set subjects_stats = Server.CreateObject("ADODB.Recordset")
subjects_stats.ActiveConnection = Connect
subjects_stats.Source = "SELECT COUNT(logs.ID_log) AS ID_log, subjects.subject_name, subjects.ID_subject, MAX(logs.log_date) AS log_date  FROM logs INNER JOIN subjects ON logs.log_subjID = subjects.ID_subject  GROUP BY logs.log_module, subjects.subject_name, subjects.ID_subject  HAVING (logs.log_module LIKE 'quiz')  ORDER BY COUNT(logs.ID_log) DESC;"
subjects_stats.CursorType = 0
subjects_stats.CursorLocation = 3
subjects_stats.LockType = 3
subjects_stats.Open()
subjects_stats_numRows = 0
%>

<%
Dim Repeat2__numRows
Repeat2__numRows = -1
Dim Repeat2__index
Repeat2__index = 0
subjects_stats_numRows = subjects_stats_numRows + Repeat2__numRows
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz subjects. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="javascript" type="text/javascript">
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
          <td colspan="3" class="subheads">Subjects statistics:</td>
        </tr>
        <% If Not subjects_stats.EOF Or Not subjects_stats.BOF Then %>
        <tr> 
          <td width="99%"  colspan="3"> 
            <table>
              <tr valign="middle"> 
                <td width="30%" align="right"><b>Subject name</b></td>
                <td align="left"><b></b></td>
                <td width="130" align="left"><b>last visit</b></td>
              </tr>
              <%
bar_max_length = 0
%>
              <% 
While ((Repeat2__numRows <> 0) AND (NOT subjects_stats.EOF)) 
%>
              <%
bar_current_length = (subjects_stats.Fields.Item("ID_log").Value)
if Int(bar_current_length) > Int(bar_max_length) then bar_max_length = Int(bar_current_length)
%>
              <tr valign="middle"> 
                <td width="30%" align="right"><%=(subjects_stats.Fields.Item("subject_name").Value)%> - </td>
                <td align="left">&nbsp;<img src="images/bar0.gif"><img src="images/bar1.gif" height ="9" width="<%=cInt(bar_current_length/bar_max_length*stat_bar_length)%>" ><img src="images/bar2.gif"> 
                  (<%=bar_current_length%>)</td>
                <td width="130" align="left"><%=(subjects_stats.Fields.Item("log_date").Value)%></td>
              </tr>
              <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  subjects_stats.MoveNext()
Wend
%>
            </table>
          </td>
        </tr>
        <% End If ' end Not subjects_stats.EOF Or NOT subjects_stats.BOF %>
        <tr> 
          <% If subjects_stats.EOF And subjects_stats.BOF Then %>
          <td width="99%"  colspan="3">Sorry, 
            there are no statistics available yet...</td>
          <% End If ' end subjects_stats.EOF And subjects_stats.BOF %>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
call log_the_page ("Quiz Subject Statistics")
%>


<%
subjects_stats.Close()
%>
