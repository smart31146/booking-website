<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->

<%
subject = Request("subject")
topic= Request("topic")
site = request("info2")
act = request("info3")
buss = Request("info1")
' ADDED 20 december 2006
fromdate=request("fromdate")
todate=request("todate")
sortby=request("sortby")
' END
if cint(buss) =0 then
	buss=0
end if

set subjects = Server.CreateObject("ADODB.Recordset")
subjects.ActiveConnection = Connect
subjects.Source = "SELECT * from subjects where subject_active_q=1"
subjects.CursorType = 0
subjects.CursorLocation = 3
subjects.LockType = 3
subjects.Open()
subjects_numRows = 0

set topics = Server.CreateObject("ADODB.Recordset")
topics.ActiveConnection = Connect
topics.Source = "SELECT * from new_subjects, q_question where s_active=1 AND s_ID=question_topic AND s_qID="&subject
topics.CursorType = 0
topics.CursorLocation = 3
topics.LockType = 3
topics.Open()
topics_numRows = 0

set info1 = Server.CreateObject("ADODB.Recordset")
info1.ActiveConnection = Connect
info1.Source = "SELECT * FROM q_info1 order by info1"
info1.CursorType = 0
info1.CursorLocation = 3
info1.LockType = 3
info1.Open()
info1_numRows = 0

set info2 = Server.CreateObject("ADODB.Recordset")
info2.ActiveConnection = Connect
info2.Source = "SELECT * FROM q_info2 where info2_info1="&buss & ""
info2.CursorType = 0
info2.CursorLocation = 3
info2.LockType = 3
info2.Open()
info2_numRows = 0

set info3 = Server.CreateObject("ADODB.Recordset")
info3.ActiveConnection = Connect
info3.Source = "SELECT * FROM q_info3 order by info3"
info3.CursorType = 0
info3.CursorLocation = 3
info3.LockType = 3
info3.Open()
info3_numRows = 0
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz topics. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}

var win=0;
function openwin(aspfile){
	win = window.open(aspfile,'display','left= 50, top=50, width=600, height=500,toolbar=no,  menubar=no,status=no,resizable=yes,scrollbars=yes');
	win.focus();
}

function filter_submit(){
	subj = document.add_topic.subject.value;
	topic = document.add_topic.topic.value;
	info1 = document.add_topic.info1.value;
	info2 = document.add_topic.info2.value;
	info3 = document.add_topic.info3.value;
	// ADDED 3 JAN 2007
	fromdate = document.add_topic.fromdate.value;
	todate = document.add_topic.todate.value;
	sortby = document.add_topic.sortby.value;
	openwin('matrix_display.asp?subj=' + subj + '&sortby=' + sortby + '&fromdate=' + fromdate + '&todate=' + todate + '&topic=' + topic + '&info1=' + info1 + '&info2=' + info2+ '&info3=' + info3)

}

</script>
</HEAD>
<BODY>
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading">Select Filter to view matrix</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form name="add_topic" method="GET">
   
        <table>
        
		  <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
          <td class="text">Subjects : </td>
          <td><select name="subject" class="text" onchange=add_topic.submit();>
          <%
		  while not subjects.EOF
		  if cint(subjects("ID_subject"))=cint(subject) then
		  %>
			<option value=<%=subjects("ID_subject").Value%> selected><%=subjects("subject_name").Value%></option>         
          <%
          else
          %>
			<option value=<%=subjects("ID_subject").Value%>><%=subjects("subject_name").Value%></option>
          <%
          end if
          subjects.MoveNext
          wend%>
          </select>
          </td>
          </tr>
          
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
          <td class="text">Topics : </td>
          <td><select name="topic" class="text">
          <%
          if topics.RecordCount > 1 then%>
			<option value=0>All Topics</option>         
		  <%
		  end if
		  while not topics.EOF
		  if cint(topics("s_ID").value)=cint(topic) then%>
			<option value=<%=topics("s_ID").Value%> selected><%=topics("s_topic").Value%></option>	
		  <% else %>
			<option value=<%=topics("s_ID").Value%>><%=topics("s_topic").Value%></option>         
          <%
          end if
          topics.MoveNext
          wend%>
          </select>
          </td>
          </tr>
          
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
          <td class="text">Business : </td>
          <td><select name="info1" class="text" onchange=add_topic.submit();>
          <option value=0> All </option>   
          <%
          while not info1.EOF
		   	if cint(buss) = cint(info1("ID_info1").Value) then
          %>
			<option value=<%=info1("ID_info1").Value%> selected><%=info1("info1").Value%></option>         
          <%
          else %>
          <option value=<%=info1("ID_info1").Value%>><%=info1("info1").Value%></option>         
          <%
          end if
          info1.MoveNext
          wend%>
          </select>
          </td>
          </tr>
          
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
          <td class="text">N/A : </td>
          <td><select name="info2" class="text">
           <option value=0> All </option>   
          <%
          while not info2.EOF
			if cint(site) = cint(info2("ID_info2").Value) then
          %>
          <option value=<%=info2("ID_info2").Value%> selected><%=info2("info2").Value%></option>
          <%
          else
          %>
           <option value=<%=info2("ID_info2").Value%>><%=info2("info2").Value%></option>    
			<%end if
          info2.MoveNext
          wend%>
          </select>
          </td>
          </tr>
          
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
          <td class="text"><% =BBPinfo3 %> : </td>
          <td><select name="info3" class="text">
          <option value=0> All </option>   
          <%
          while not info3.EOF
          if cint(act) = cint(info3("ID_info3").Value) then
          %>
			<option value=<%=info3("ID_info3").Value%> selected><%=info3("info3").Value%></option>         
          <%
          else%>
          <option value=<%=info3("ID_info3").Value%>><%=info3("info3").Value%></option>  
          <%end if
          info3.MoveNext
          wend%>
          </select>
          </td>
          </tr>
                 <tr class="table_normal">
			<td class="text" valign="top">Sessions between:</td>
			<td class="text">
				<input type="text" name="fromdate" maxlength="19" class="formitem1" onDblClick="this.value='<% =cDateSQL(Now()-1)%>'; " size="25" value="<%=fromdate%>" >&nbsp;<bR>(yyyy-mm-dd hh:mm:ss), doubleclick = TODAY - 1 day<br>
				&nbsp;&nbsp;&nbsp;&nbsp;<Br>AND<br><br>
				<input type="text" name="todate" maxlength="19" class="formitem1" onDblClick="this.value='<% =cDateSQL(Now())%>'; " size="25" value="<%=todate%>">
				<br>(yyyy-mm-dd hh:mm:ss), doubleclick = TODAY
             </td>
		</tr> 
		<tr class="table_normal">
			<td class="text" valign="top">Sort by:</td>
			<td class="text"><select name="sortby" class="text">
			 <% if sortby = "user_firstname" or sortby = "" then %>
			<option value="user_firstname" selected>First name</option>  
          <option value="user_lastname">Last name</option>   
          <option value="session_finish">Date</option>         
              <% elseif sortby = "user_lastname" then %>
			<option value="user_firstname">First name</option>  
          <option value="user_lastname" selected>Last name</option>   
          <option value="session_finish">Date</option>             
              <% elseif sortby = "session_finish" then %>
			<option value="user_firstname">First name</option>  
          <option value="user_lastname">Last name</option>   
          <option value="session_finish" selected>Date</option>
		  <%end if%>
          </select>
             </td>
		</tr> 
          <tr><td>
          <input type="button" name="Submit" value="&gt;&gt;&gt; Filter users &lt;&lt;&lt;" class="quiz_button" onclick="return filter_submit();">
          </td></tr>
 
</table>
</form>
</BODY>
</HTML>
<%
call log_the_page ("Quiz List Topics: " & subj)
%>

<%
topics.Close()
info1.Close()
info2.Close()
info3.close()
%>

