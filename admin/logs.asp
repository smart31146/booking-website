<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
buss =  request("info1")
set subjects = Server.CreateObject("ADODB.Recordset")
subjects.ActiveConnection = Connect
subjects.Source = "SELECT ID_subject, subject_name FROM subjects"
subjects.CursorType = 0
subjects.CursorLocation = 3
subjects.LockType = 3
subjects.Open()
subjects_numRows = 0

set info1 = Server.CreateObject("ADODB.Recordset")
info1.ActiveConnection = Connect
info1.Source = "SELECT * from q_info1 order by info1"
info1.CursorType = 0
info1.CursorLocation = 3
info1.LockType = 3
info1.Open()
info1_numRows = 0

if buss <> "" then
set info2 = Server.CreateObject("ADODB.Recordset")
info2.ActiveConnection = Connect
info2.Source = "SELECT * from q_info2 where info2_info1 =" & buss &" order by info2" 
info2.CursorType = 0
info2.CursorLocation = 3
info2.LockType = 3
info2.Open()
info2_numRows = 0
'Response.Write info2.Source
end if

set info3 = Server.CreateObject("ADODB.Recordset")
info3.ActiveConnection = Connect
info3.Source = "SELECT * from q_info3 order by info3"
info3.CursorType = 0
info3.CursorLocation = 3
info3.LockType = 3
info3.Open()
info3_numRows = 0
%>
<script>
function checkform() {
	document.searchlogs.action="logs.asp"
	document.searchlogs.target="_self"
	document.searchlogs.submit()
}
</script>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Logs. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
</HEAD>
<BODY>
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> BBP LOG files</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom" class="text"> 
      <form name="searchlogs" method="post" action="log_results.asp" target="_blank">
        <p>Filter all log entries with following keys (nothing for all records).</p>
        <table>
          <tr> 
            <td width="120">Start date:</td>
            <td > 
              <input type="text" name="fromdate" maxlength="19" class="formitem1" onDblClick="this.value='<%=cDateSQL(Now()-1)%>';"size="30" value="<%=request("fromdate")%>">
              (yyyy-mm-dd hh:mm:ss), doubleclick = Today - 1 day</td>
          </tr>
          <tr> 
            <td>End date:</td>
            <td > 
              <input type="text" name="todate" maxlength="19" class="formitem1" onDblClick="this.value='<%=cDateSQL(Now())%>';" size="30" value="<%=request("todate")%>">
              (yyyy-mm-dd hh:mm:ss), doubleclick = Today</td>
          </tr>
          <tr> 
            <td>Client IP address:</td>
            <td> 
              <input type="text" name="ipaddress" maxlength="15" class="formitem1" size="50" value="<%=request("ipaddress")%>">
            </td>
          </tr>
          <tr> 
            <td>Module:</td>
            <td> 
              <select name="module" class="formitem1">
                <option value="all">All</option>
                <option value="home">Home page</option>
                <option value="Guide Index">Guide</option>
				<option value="Guide Key Points">Guide Key Points</option>
				<option value="Guide Examples">Guide Examples</option>
                <option value="Training and quiz">Training</option>
				 <option value="Certificate">Certificate</option>
                <option value="Guide Search">Search engine</option>
                <option value="Registration">User registration</option>
                <option value="password">Password reset</option>
				 <option value="Change password">Password change</option>
				 <option value="Help Page">Help page</option>
                <option value="admin">Admin interface</option>
              </select>
            </td>
          </tr>
          <tr> 
            <td>User name:</td>
            <td> 
              <input type="text" name="username" class="formitem1" size="50" value="<%=request("username")%>">
            </td>
          </tr>
          <tr> 
            <td>Subject:</td>
            <td> 
              <select name="subject" class="formitem1">
                <option value="0">All</option>
                <%
While (NOT subjects.EOF)
if cint(request("subject")) = cint(subjects.Fields.Item("ID_subject").Value) then
%>
     <option value="<%=(subjects.Fields.Item("ID_subject").Value)%>" selected><%=(subjects.Fields.Item("subject_name").Value)%></option>
          <%
         else%>
        <option value="<%=(subjects.Fields.Item("ID_subject").Value)%>" ><%=(subjects.Fields.Item("subject_name").Value)%></option>
<%        
end if
  subjects.MoveNext()
Wend
'If (subjects.CursorType > 0) Then
'  subjects.MoveFirst
'Else
  subjects.Requery
'End If
%>
              </select>
            </td>
          </tr>
         <!-- <tr> 
            <td>Comment</td>
            <td> 
              <input type="text" name="comment" class="formitem1" size="50" value="<%=request("comment")%>">
            </td>
          </tr>-->
          <tr> 
            <td>Business Group:</td>
            <td> 
              <select name="info1" class="formitem1" onchange=checkform();>
              <option value="0">--- All Business Groups --</option>
                <%
While (NOT info1.EOF)
if cint(buss) = cint(info1.Fields.Item("ID_info1").Value) then
%>
                <option value="<%=(info1.Fields.Item("ID_info1").Value)%>" selected><%=(info1.Fields.Item("info1").Value)%></option>
                <%
                else %>
                <option value="<%=(info1.Fields.Item("ID_info1").Value)%>"><%=(info1.Fields.Item("info1").Value)%></option>                         
  <%
  end if
  info1.MoveNext()
Wend%>
              </select>
            </td>
          </tr>
          <tr> 
            <td height=25>N/A:</td>
            <td> 
              <select name="info2" class="formitem1">
              <%if buss <> "" then%>
              <option value="0">--- All --</option>
                <%
While (NOT info2.EOF)
%>
                <option value="<%=(info2.Fields.Item("ID_info2").Value)%>"><%=(info2.Fields.Item("info2").Value)%></option>
                <%
  info2.MoveNext()
Wend
else%>
	<option value="0"></option>
<%end if
%>
              </select>
            </td>
          </tr>
          <tr> 
            <td><% =BBPinfo3 %>:</td>
            <td> 
              <select name="info3" class="formitem1">
              <option value="0">--- All <% =BBPinfo3s %> --</option>
                <%
While (NOT info3.EOF)
if cint(request("info3")) = cint(info3.Fields.Item("ID_info3").Value) then
%>
                <option value="<%=(info3.Fields.Item("ID_info3").Value)%>" selected><%=(info3.Fields.Item("info3").Value)%></option>
              
                <%
                else%>
                 <option value="<%=(info3.Fields.Item("ID_info3").Value)%>" ><%=(info3.Fields.Item("info3").Value)%></option>
<% 
end if       
info3.MoveNext()

Wend%>
              </select>
            </td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td> 
              <input type="reset" name="Reset" value="Reset the form" class="quiz_button">
              <input type="submit" name="Submit" value="View filtered LOG files" class="quiz_button">
            </td>
          </tr>
        </table>
        <p>&nbsp;</p>
      </form>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>
</BODY>
</HTML>
<%
subjects.Close()
info1.Close()
info3.Close()
call log_the_page ("BBG LOGS")
%>
