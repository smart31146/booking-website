<%@LANGUAGE="VBSCRIPT"%>

<% 
'Response buffer is used to buffer the output page. That means if any database exception occurs the contents can be cleared without processed any script to browser
 Response.Buffer = True
 
' "On Error Resume Next" method allows page to move to the next script even if any error present on page whcich will be caught after processing all asp script on page
 On Error Resume Next
 
'Changed by PR on 25.02.16
 %>

<!--#include file="connections/bbg_conn.asp" -->
<!--#include file="connections/include.asp" -->

<%
userid = request("user")
if Err.Number = 0 then
set user = Server.CreateObject("ADODB.Recordset")
user.ActiveConnection = Connect
user.Source = "SELECT * FROM q_user WHERE ID_user = "&userid
user.CursorType = 0
user.CursorLocation = 3
user.LockType = 3
user.Open()
user_numRows = 0
end if


'set preferences = Server.CreateObject("ADODB.Recordset")
'preferences.ActiveConnection = Connect
'preferences.Source = "SELECT pass_rate FROM preferences where pref_active=1"
'preferences.CursorType = 0
'preferences.CursorLocation = 3
'preferences.LockType = 3
'preferences.Open()
'preferences_numRows = 0
'
'while not preferences.eof
'	passrate = preferences.Fields.item("pass_rate").value
'	preferences.MoveNext
'wend
'preferences.Close()
if Err.Number = 0 then
set sessions = Server.CreateObject("ADODB.Recordset")
sessions.ActiveConnection = Connect
if request("latest")= "View latest quiz results" then
	sessions.Source = "SELECT q_session.ID_Session, q_session.session_users, q_session.Session_date, subjects.id_subject,subjects.subject_name, q_session.Session_total, q_session.Session_done, q_session.Session_correct, q_session.Session_stop, q_certification.percentage_required  FROM q_user INNER JOIN (q_session INNER JOIN subjects ON q_session.Session_subject = subjects.ID_subject) ON q_user.ID_user = q_session.Session_users INNER JOIN q_certification ON q_certification.q_session=q_session.id_session WHERE Session_users = "&userid + " AND (q_session.Session_done = 1) ORDER BY subjects.id_subject, q_session.Session_date desc ;"
else
	sessions.Source = "SELECT q_session.ID_Session, q_session.session_users, q_session.Session_date, subjects.id_subject,subjects.subject_name, q_session.Session_total, q_session.Session_done, q_session.Session_correct, q_session.Session_stop, q_certification.percentage_required  FROM q_user INNER JOIN (q_session INNER JOIN subjects ON q_session.Session_subject = subjects.ID_subject) ON q_user.ID_user = q_session.Session_users INNER JOIN q_certification ON q_certification.q_session=q_session.id_session WHERE Session_users = "&userid + " ORDER BY subjects.id_subject, q_session.Session_date desc ;"
end if

sessions.CursorType = 0
sessions.CursorLocation = 3
sessions.LockType = 3
sessions.Open()
sessions_numRows = 0
end if
%>



<!doctype html>
<HTML>

<head>
<TITLE><%=client_name_short%> Quiz results <%=Session("firstname") & " " & Session("lastname")%></TITLE>
<link rel="stylesheet" href="style/bbp_acme34.css" type="text/css">
<link rel="stylesheet" href="style/bbp_style_acme34.css" type="text/css">
 <script src="//ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js?v=bbp34"></script>
 <script >


function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

</script>
<style>
body    {overflow-x:hidden;}
body    {overflow-y:scroll;}

</style>
<script>
function change_style(id)
{
if(id=="all"){

$("#"+id).addClass("button");

$("#late").addClass("button_not_pressed");
}
}
</script>
</head>

<BODY  class=bodyprivacy>

            
<table  >
    <tr>
      <td ><div class="header_logo"></div></td>
      <td>&nbsp;</td>
      <td >
        
      </td>
    </tr>
</table>
<TABLE  >
  <TR>
    <!--<TD width=219 class="headersuserdb">&nbsp;</TD>-->
    <TD  class="headersuserdb" ><div style="color:#000"><h1>View previous quiz results</h1></div></TD>
  </TR>
</TABLE>

<table  >
<tr><td>
<!--<TD width=182>-->Below are the results for your previous attempts at the Quiz questions for each topic. Use this page to see which topics you have completed and which of those topics you have passed. You should revise those topics for which you have not yet passed the quiz<br></TD>
</tr><tr>
    <TD >
<b><br><%=(user.Fields.Item("user_firstname").Value)%>&nbsp;<%=(user.Fields.Item("user_lastname").Value)%>'s quiz results:</b>
</td>
</tr>
<tr>
<td>&nbsp;</td>
</tr>
<tr>
<!--<TD width=182>&nbsp;</TD>-->
    <TD   >
	
</td>
</tr>

</table>

<table class="centerposition"  >
	<tr>
		<!--<td width="175">&nbsp;</td>-->
		<td >
		<form name="results" action="user_sessions_new.asp">
	<table >
	<tr>
	<% if request("latest")="View latest quiz results" or request("latest")="1"  then %>
<td><input type="submit" name="latest" id="late" value="View latest quiz results" class="button"  style="width:170px;"  ></td>
<td><input type="submit" name="latest" id="all" value="View all quiz results" class="button_not_pressed" style="width:150px;" ></td>
<%else %>
<td><input type="submit" name="latest" id="late" value="View latest quiz results"  class="button_not_pressed" class="button"  style="width:170px;"  ></td>
<td><input type="submit" name="latest" id="all" value="View all quiz results"  class="button" style="width:150px;" ></td>
<% end if %>

</tr>

</table>
<input type="hidden" name="user" value=<%=userid%>>
</form>
		<div class="CSSTableGenerator" >
			<table  >
			<tr>
				
				<td class="subheadsuserdb">
					Subject
				</td>
				<td class="subheadsuserdb">
					Date
				</td>
				<td class="subheadsuserdb">
					Correct
				</td>
				<td class="subheadsuserdb">
					Total
				</td>
				<!--<td class="subheadsuserdb">
					Done
				</td>-->
				<td class="subheadsuserdb">
					Finished
				</td>
				<td class="subheadsuserdb">
					Score
				</td>
				<td class="subheadsuserdb">
					Passed
				</td>
			</tr>
			<%
			subid=0
			overall_rate = 0
			sum_tests = 0
			numbers=1
			While (NOT sessions.EOF)
			passrate=sessions.Fields.Item("percentage_required").Value
			if request("latest") <> "View all quiz results" then
				if cint(subid) <> (sessions.Fields.Item("id_subject").Value)then
				subid = (sessions.Fields.Item("id_subject").Value)
				user_rate = FormatNumber((sessions.Fields.Item("Session_correct").Value)/(sessions.Fields.Item("Session_total").Value)*100,2)
				if cInt(user_rate) >= cInt(passrate) then user_pass = 1 else user_pass = 0
				overall_rate = overall_rate + user_rate
				%>
						<tr class="table_normaluserdb" onMouseOver="pviiClassNew(this,'table_hluserdb')" onMouseOut="pviiClassNew(this,'table_normaluserdb')">
						  <td class="textuserdb" ><strong><a class="quiz" style="color:#000;text-decoration:underline;" href="user_session_details_new.asp?user_session=<%=(sessions.Fields.Item("ID_Session").Value)%>&amp;user=<%=(user.Fields.Item("ID_user").Value)%>&amp;subject=<%=(sessions.Fields.Item("id_subject").Value)%>&amp;latest=<%=request("latest")%>"><%=(sessions.Fields.Item("subject_name").Value)%></a></strong></td>
						  <td class="textuserdb"><%=(sessions.Fields.Item("Session_date").Value)%></td>
						  <td   class="textuserdb"><%=(sessions.Fields.Item("Session_correct").Value)%></td>
						  <td   class="textuserdb"><%=(sessions.Fields.Item("Session_total").Value)%></td>
						  <!--<td   class="textuserdb"><%=(sessions.Fields.Item("Session_stop").Value)%></td>-->
						  <td class="textuserdb" >
							<%
							if abs(sessions.Fields.Item("Session_done").Value) = 1 then
								response.write "<b>YES</b>"
								sum_tests = sum_tests+1
							else
								response.write "<b>NO</b>"
							end if
							%>
						  </td>
						  <td class="textuserdb" >
							<%if user_pass = 1 then response.write ("<div style='color:green'>" & Round(user_rate) & "%</div>") else response.write ("<div style='color:red'>" & Round(user_rate) & "% </div>")%></td>
						  <td class="textuserdb"   >
							<%if user_pass = 1 then response.write "<img src='images/yes.png' alt=''>" else response.write "<img src='images/no.png' alt=''>"%>
						  </td>
						</tr>
						<%
				  sessions.MoveNext()
				  numbers=numbers+1
			else
				sessions.MoveNext()
			end if
			else
				user_rate = FormatNumber((sessions.Fields.Item("Session_correct").Value)/(sessions.Fields.Item("Session_total").Value)*100,2)
				if cInt(user_rate) >= cInt(passrate) then user_pass = 1 else user_pass = 0
				overall_rate = overall_rate + user_rate
				%>
						<tr class="table_normaluserdb" onMouseOver="pviiClassNew(this,'table_hluserdb')" onMouseOut="pviiClassNew(this,'table_normaluserdb')">
						  <td class="textuserdb" ><strong><a class="quiz" style="color:#000;text-decoration:underline;" href="user_session_details_new.asp?user_session=<%=(sessions.Fields.Item("ID_Session").Value)%>&amp;user=<%=(user.Fields.Item("ID_user").Value)%>&amp;subject=<%=(sessions.Fields.Item("id_subject").Value)%>&amp;latest=<%=request("latest")%>"><%=(sessions.Fields.Item("subject_name").Value)%></a></strong></td>
						  <td class="textuserdb"><%=(sessions.Fields.Item("Session_date").Value)%></td>
						  <td   class="textuserdb"><%=(sessions.Fields.Item("Session_correct").Value)%></td>
						  <td   class="textuserdb"><%=(sessions.Fields.Item("Session_total").Value)%></td>
						 <!-- <td   class="textuserdb"><%=(sessions.Fields.Item("Session_stop").Value)%></td>-->
						  <td class="textuserdb"    >
							<%
							if abs(sessions.Fields.Item("Session_done").Value) = 1 then
								response.write "<b>YES</b>"
								sum_tests = sum_tests+1
							else
								response.write "<b>NO</b>"		
							end if
							%>
						  </td>
						  <td class="textuserdb"    >
							<%if user_pass = 1 then response.write ("<div style='color:green'>" & Round(user_rate) & "%</div>") else response.write ("<div style='color:red'>" & Round(user_rate) & "%</div>")%>
						  </td>
						  <td class="textuserdb"    >
							<%if user_pass = 1 then response.write "<img src='images/yes.png'>" else response.write "<img src='images/no.png'>"%>
						  </td>
						</tr>
						<%
				  sessions.MoveNext()
				  numbers=numbers+1
			end if
			Wend
			%>
			
			</table>
			</div>
		</td>
	</tr>
</table>

</BODY>

</HTML>
<!-- #include file = "errorhandler/index.asp"-->










