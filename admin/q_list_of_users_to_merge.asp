<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
Response.Buffer=true
Server.ScriptTimeout = 400
if cStr(Request.Querystring("show_lines")) <> "" then show_lines = cInt(Request.Querystring("show_lines"))
numbers=1
count = 1
SQL_having = ""
user_to_merge_into = 0

If cInt(Request.Querystring("user")) <> 0 then 
	user_to_merge_into = cInt(Request.Querystring("user"))

end if 

If cStr(Request.Querystring("filter_username")) <> "" then 
	SQL_having = " HAVING ((q_user.user_lastname) Like '%" + Replace(uCase(cStr(Request.Querystring("filter_username"))), "'", "''") + "%' OR  (q_user.user_firstname) Like '%" + Replace(uCase(cStr(Request.Querystring("filter_username"))), "'", "''") + "%') " 
Else
	'show no records
	SQL_having = " HAVING q_user.ID_user=-1 "
end if 

'PN 040811 if the user is requesting to merge a user into another
merge_user=0
If cInt(Request.Querystring("mergerecords")) = 1 then 
	merge_user = cInt(Request.Querystring("mergeuser"))
	'firstly set all Session_users in q_session to user_to_merge_into
	set merge_sessions  = Server.CreateObject("ADODB.Command")
	merge_sessions.ActiveConnection = Connect
	merge_sessions.CommandText = "UPDATE q_session set Session_users = '" & user_to_merge_into & "' WHERE Session_users = '" & merge_user & "';"
	merge_sessions.Execute
	merge_sessions.ActiveConnection.Close
	
	'then delete the merge_user record in q_user
	set delete_user  = Server.CreateObject("ADODB.Command")
	delete_user.ActiveConnection = Connect
	delete_user.CommandText = "delete from q_user where ID_user  = '" & merge_user & "';"
	delete_user.Execute
	delete_user.ActiveConnection.Close
end if 

set user_merge_with = Server.CreateObject("ADODB.Recordset")
user_merge_with.ActiveConnection = Connect
user_merge_with.Source = "SELECT q_user.ID_user, q_user.user_lastname, q_user.user_firstname,q_user.user_city, q_info1.info1, q_info2.info2, q_info3.info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3, COUNT(q_session.ID_session) AS session_count FROM (q_info3 RIGHT JOIN (q_info2 RIGHT JOIN (q_info1 RIGHT JOIN q_user ON q_info1.ID_info1 = q_user.user_info1) ON q_info2.ID_info2 = q_user.user_info2) ON q_info3.ID_info3 = q_user.user_info3) LEFT JOIN q_session ON q_user.ID_user = q_session.Session_users where q_user.ID_user='"& user_to_merge_into &"' GROUP BY q_user.user_lastname, q_user.user_firstname, q_user.ID_user, q_info1.info1, q_info2.info2, q_info3.info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3,q_user.user_city ORDER BY q_user.user_lastname, q_user.user_firstname;"

'Response.Write users.Source
user_merge_with.CursorType = 0
user_merge_with.CursorLocation = 3
user_merge_with.LockType = 3
user_merge_with.Open()



set users = Server.CreateObject("ADODB.Recordset")
users.ActiveConnection = Connect
users.Source = "SELECT q_user.ID_user, q_user.user_lastname, q_user.user_firstname,q_user.user_city, q_info1.info1, q_info2.info2, q_info3.info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3, COUNT(q_session.ID_session) AS session_count FROM (q_info3 RIGHT JOIN (q_info2 RIGHT JOIN (q_info1 RIGHT JOIN q_user ON q_info1.ID_info1 = q_user.user_info1) ON q_info2.ID_info2 = q_user.user_info2) ON q_info3.ID_info3 = q_user.user_info3) LEFT JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_user.ID_user <> " & user_to_merge_into & ") GROUP BY q_user.user_lastname, q_user.user_firstname, q_user.ID_user, q_info1.info1, q_info2.info2, q_info3.info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3,q_user.user_city " + SQL_having + " ORDER BY q_user.user_lastname, q_user.user_firstname;"

'Response.Write users.Source
users.CursorType = 0
users.CursorLocation = 3
users.LockType = 3
users.Open()
users_numRows = 0
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Merge user records. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--

function filter_submit() {
	document.filter_users.action="q_list_of_users_to_merge.asp"
	document.filter_users.target="_self"
	document.filter_users.submit()
}

function merge_user(fu,mu){

	var agree=confirm("Are you sure you want to merge the record for this user? This action cannot be undone.  Click OK to proceed, or Cancel if you are not sure.");
	if(agree){
		//alert("q_list_of_users_to_merge.asp?user=<%=(user_to_merge_into)%>&filter_username="+fu+"&mergerecords=1&mergeuser="+mu);
		
		location.href="q_list_of_users_to_merge.asp?user=<%=(user_to_merge_into)%>&filter_username="+fu+"&mergerecords=1&mergeuser="+mu;
		window.opener.location.reload();
	}
}

//-->
</script>
</HEAD>

<BODY>
	
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading">Merge User Records</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form name="filter_users">
	  <input type="hidden" name="user" value="<%=user_to_merge_into%>">
	 <table>
	  
	  <tr>
				
				<td colspan="7" >&nbsp; 
				 This form will find user records that you can merge into the record for
				<%If (Not user_merge_with.EOF Or Not user_merge_with.BOF) Then%>
					<%=(user_merge_with.Fields.Item("user_firstname").Value)%>&nbsp;<%=(user_merge_with.Fields.Item("user_lastname").Value)%>,&nbsp;<%=(user_merge_with.Fields.Item("user_city").Value)%><%
				End If%>.  Use this form to remove duplicate entries from the user records.
				</td>
					
	  </tr>
	  <tr> 
				<td  width="18">&nbsp;</td>
				<td colspan="5" >&nbsp; 
				  <table width="50%" align="center">
					 <tr class="table_normal"> 
						<td class="text" width="18">&nbsp;</td>
						<td class="text" valign="top" width="143">First OR Last name:</td>
						<td class="text" valign="top" > 
						<table><tr><td><input type="text" name="filter_username" value="<%=request.querystring("filter_username")%>" class="formitem1"></td><td></td></tr></table>
						</td>
					<td class="text" width="18" colspan="3">&nbsp;</td>
					
			  </tr>
			<tr class="table_normal"> 
			  <td colspan="6" align="center"class="text"> 
				<input type="button" name="Submit" value="&gt;&gt;&gt; Find users &lt;&lt;&lt;" class="quiz_button" onclick="return filter_submit();">
			  </td>
			</tr> 
	</table>
     
        <table>
			
         <tr> 
            <td >&nbsp;</td>
            <td >Last name &amp; First name</td>
			<td >City of birth</td>
            <td >Business &amp; <% =BBPinfo3 %></td>
            <td ><% =BBPinfo3 %></td>
           <!-- <td >Active</td>
            <td >Logs</td>
            <td >Sess.</td>-->
            <!--td >Rate</td-->
            <td >Merge to current user record</td>
          </tr>
          <% If Not users.EOF Or Not users.BOF Then
While (  (NOT users.EOF)) 



set user_details = Server.CreateObject("ADODB.Recordset")
user_details.ActiveConnection = Connect
user_details.Source = "SELECT q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.Session_done, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ")  order by session_subject,session_date desc"

user_details.Open()
user_details_numRows = 0
'Response.Write user_details.source
user_session_rate = 0
user_session_count = 0
user_total_rate = 0
subid = 0 

user_details.Close()
if user_session_count > 0 then 
	user_total_rate = (user_session_rate/user_session_count)
end if

if (cstr(noquiz)="1") then
	if cInt(users.Fields.Item("session_count").Value) = 0 then
%>
	<tr class="table_normal">
			
            <td class="text" width="20"><%=count%></td>
            <td width="200" class="text"> 
             <%=(users.Fields.Item("user_lastname").Value)%>&nbsp;<%=(users.Fields.Item("user_firstname").Value)%></td>
			 <td width="50" class="text"><%=(users.Fields.Item("user_city").Value)%> (<%=(users.Fields.Item("info2").Value)%>)</td>
            <td width="140" class="text"><%=(users.Fields.Item("info1").Value)%> (<%=(users.Fields.Item("info2").Value)%>)</td>
            <td width="140" class="text"><%=(users.Fields.Item("info3").Value)%></td>
            <td width="30" class="text" align=center> 
            <!--  <%if abs(users.Fields.Item("user_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
            </td>
            <td width="20" class="text"><%=(users.Fields.Item("user_logcount").Value)%>x</td>
            <td width="20" class="text"><%=(users.Fields.Item("session_count").Value)%></td>-->
            <!--td width="20" class="text"> 
              <%
			response.write("<font color = blue>N/A</font>") 
			count = count + 1
			%>
            </td-->
            <td  align="right" width="20"> 
             <a href="javascript:merge_user('<%=(Request.Querystring("filter_username"))%>',<%=users.Fields.Item("ID_user").Value%>)"><img src="images/merge.gif" alt="Merge this user with current user" width="16" height="15" border="0"></a>
            </td>
			<!--PN 040811 add merge facility so that users can be merged into this user to get rid of self reg duplicates-->
			<!--td  align="right" width="20"> 
			<a href="q_list_of_users_to_merge.asp?user=<%=(users.Fields.Item("ID_user").Value)%>"  target="_blank"><img src="images/late.gif" alt="View users yet to complete training" width="15" height="15" border="0"></a>
			</td-->
          </tr>
<%
end if
else

if cstr(results)="" or cstr(results)="2" then
%>
			<tr class="table_normal">
			
            <td class="text" width="20"><%=numbers%></td>
            <td width="200" class="text"> 
               <%=(users.Fields.Item("user_lastname").Value)%>&nbsp;<%=(users.Fields.Item("user_firstname").Value)%>
			   </td>
			<td width="50" class="text"><%=(users.Fields.Item("user_city").Value)%>
            <td width="140" class="text"><%=(users.Fields.Item("info1").Value)%> (<%=(users.Fields.Item("info2").Value)%>)</td>
            <td width="140" class="text"><%=(users.Fields.Item("info3").Value)%></td>
           <!-- <td width="30" class="text" align=center> 
              <%if abs(users.Fields.Item("user_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
            </td>
            <td width="20" class="text"><%=(users.Fields.Item("user_logcount").Value)%>x</td>
            <td width="20" class="text"><%=(users.Fields.Item("session_count").Value)%></td>-->
            <!--td width="20" class="text"> 
              <%
			if user_session_count = 0 then
				response.write("<font color = blue>N/A</font>") 
			elseif (user_total_rate) >= cint(passrate) then 
				response.write("<font color = green>" & FormatNumber(user_total_rate,2) & "%</font>") 
			else 
				response.write("<font color = red>" & FormatNumber(user_total_rate,2) & "%</font>")
			end if
			%>
            </td-->
            <td  align="right" width="20"> 
             <a href="javascript:merge_user('<%=(Request.Querystring("filter_username"))%>',<%=users.Fields.Item("ID_user").Value%>)"><img src="images/merge.gif" alt="Merge this user with current user" width="16" height="15" border="0"></a>
            </td>
          </tr>
  <%
else if (cstr(results)="1") and (user_total_rate >= cint(passrate)) then

%>

	<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">
            <td class="text" width="20"><%=count%></td>
            <td width="200" class="text"> 
              
              <%=(users.Fields.Item("user_lastname").Value)%>&nbsp;<%=(users.Fields.Item("user_firstname").Value)%>
			  </td>
			<td width="50" class="text"><%=(users.Fields.Item("user_city").Value)%>
            <td width="140" class="text"><%=(users.Fields.Item("info1").Value)%> (<%=(users.Fields.Item("info2").Value)%>)</td>
            <td width="140" class="text"><%=(users.Fields.Item("info3").Value)%></td>
           <!-- <td width="30" class="text" align=center> 
              <%if abs(users.Fields.Item("user_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
            </td>
            <td width="20" class="text"><%=(users.Fields.Item("user_logcount").Value)%>x</td>
            <td width="20" class="text"><%=(users.Fields.Item("session_count").Value)%></td>-->
            <!--td width="20" class="text"> 
              <%
			response.write("<font color = green>" & FormatNumber(user_total_rate,2) & "%</font>") 
			count = count + 1
			%>
            </td-->
            <td  align="right" width="20"> 
             <a href="javascript:merge_user('<%=(Request.Querystring("filter_username"))%>',<%=users.Fields.Item("ID_user").Value%>)"><img src="images/merge.gif" alt="Merge this user with current user" width="16" height="15" border="0"></a>
             </td>
          </tr>     
  <%
  
  else if (cstr(results)="0") and (user_total_rate <= cint(passrate)) and (user_session_count<>0) then%>
			<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
            <td class="text" width="20"><%=count%></td>
            <td width="200" class="text"> 
             
              <%=(users.Fields.Item("user_lastname").Value)%>&nbsp;<%=(users.Fields.Item("user_firstname").Value)%></td>
			<td width="50" class="text"><%=(users.Fields.Item("user_city").Value)%>
            <td width="140" class="text"><%=(users.Fields.Item("info1").Value)%> (<%=(users.Fields.Item("info2").Value)%>)</td>
            <td width="140" class="text"><%=(users.Fields.Item("info3").Value)%></td>
          <!--  <td width="30" class="text" align=center> 
              <%if abs(users.Fields.Item("user_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
            </td>
            <td width="20" class="text"><%=(users.Fields.Item("user_logcount").Value)%>x</td>
            <td width="20" class="text"><%=(users.Fields.Item("session_count").Value)%></td>-->
            <!--td width="20" class="text"> 
              <%
			response.write("<font color = red>" & FormatNumber(user_total_rate,2) & "%</font>")
			count = count + 1
			%>
            </td-->
            <td  align="right" width="20"> 
             <a href="javascript:merge_user('<%=(Request.Querystring("filter_username"))%>',<%=users.Fields.Item("ID_user").Value%>)"><img src="images/merge.gif" alt="Merge this user with current user" width="16" height="15" border="0"></a>
            </td>
          </tr>
  <%
  end if
  end if
  end if
end if
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  users.MoveNext()
  numbers=numbers+1
Wend
'If (users.CursorType > 0) Then
'  users.MoveFirst
'Else
  users.Requery
'End If
End If
'_______________________________________________________________________________
%>

				
              </table>
            </td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
</BODY>
</HTML>
<%
call log_the_page ("Quiz List Users Merge")
users.Close()

user_merge_with.Close()


Set users = Nothing
Set users = Nothing


%>

