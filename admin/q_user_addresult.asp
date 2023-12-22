<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%


'CREATED BY JOHAN BREDENHOLT, CARLJ WEBBYRÅ, NORRKÖPING SWEDEN 2011-01-24


Response.Buffer=true
Server.ScriptTimeout = 400

Set obj= Server.CreateObject("ADODB.RecordSet")

if Request.Querystring("user") <> "" THEN
	user = clng(Request.Querystring("user"))
	SQL = "SELECT * FROM q_user WHERE id_user = "&user&""
	obj.Open SQL, Connect, 3,3
	q_name = obj("user_firstname") & " " &  obj("user_lastname")
	obj.close
END IF



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
		//alert("q_list_of_users_to_merge.asp?user=<%=(user)%>&filter_username="+fu+"&mergerecords=1&mergeuser="+mu);
		
		location.href="q_list_of_users_to_merge.asp?user=<%=(user)%>&filter_username="+fu+"&mergerecords=1&mergeuser="+mu;
		window.opener.location.reload();
	}
}

function checkSubject() {
 if (document.frmSubject.session_subject.value.length == "" )
	  {
	   alert ("You have to choose a subject to proceed") ; return false;
	  }
}
function checkUser() {
 if (document.frmSubject.filter_username.value.length < 2 )
	  {
	   alert ("First or last name has to be at least 2 characters.") ; return false;
	  }
}

function checkForm() {
 if (document.frmSubject.Session_date.value.length == "" )
	  {
	   alert ("Date can't be empty") ; return false;
	  }
 if (document.frmSubject.session_version.value.length == "" )
	  {
	   alert ("Version can't be empty") ; return false;
	  }
 if (document.frmSubject.session_correct.value.length == "" || document.frmSubject.session_total.value.length == "")
	  {
	   alert ("Score can't be empty") ; return false;
	  }
 if (parseInt(document.frmSubject.session_correct.value) > parseInt(document.frmSubject.session_total.value))
	  {
	   alert ("User score can't be higher than total score") ; return false;
	  }
 if (document.frmSubject.percentage_required.value.length == "" )
	  {
	   alert ("Percent required can't be empty") ; return false;
	  }
	return confirm('The result will now be saved.\n\nProceed?');
}
function isNumberKey(evt) {
   var charCode = (evt.which) ? evt.which : event.keyCode
   if ((charCode > 31 && (charCode < 48 || charCode > 57)))
      return false;

   return true;
}
//-->
</script>
<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.3.2/jquery.min.js?v=bbp34"></script>
<!-- required plugins -->
<script type="text/javascript" src="styles/date.js?v=bbp34"></script>
<!--[if IE]><script type="text/javascript" src="styles/jquery.bgiframe.min.js?v=bbp34"></script><![endif]-->
<!-- jquery.datePicker.js?v=bbp34 -->
<script type="text/javascript" src="styles/jquery.datePicker.js?v=bbp34"></script>
<!-- datePicker required styles -->
<link rel="stylesheet" type="text/css" media="screen" href="styles/datePicker.css">
		
 <!-- page specific scripts -->
<script type="text/javascript" charset="utf-8">
    $(function()
        {
			$('.date-pick').datePicker({startDate:'01/01/1996',clickInput:true})
        });
</script>
<script language="JavaScript" type="text/javascript">
 function CloseAndRefresh() 
  {
     opener.location.reload(true);
     self.close();
  }
</script>

</HEAD>

<BODY>	
<table>
	<tr> 
		<td align="left" valign="bottom" class="heading">Add user results <% =q_name %><br><br></td>
	</tr>
	<tr> 
    <td align="left" valign="bottom"> 
	<% if request.querystring("alt")="search" THEN%>
	
	<form method="post" name="frmSubject" action="?alt=searchresult&user=<%=user%>" onsubmit="return(checkUser());">	  
		<table>
			<TD class="text" width="120">First OR Last name</TD>
			<TD class="text" width="108"><input type="text" name="filter_username" value="" class="formitem1" style="width:100px;"></TD>
			<TD class="text" width="108"><input type="Submit" name="fSubmit" id="fSubmit" value="Search user" class="quiz_button"></TD>
		</table>
	</form>
	</TR>
   
	 <% ELSEif request.querystring("alt")="searchresult" THEN
				SQL = "SELECT * FROM q_user WHERE user_firstname  LIKE '%"&trim(request.form("filter_username"))&"%' OR user_lastname LIKE '%"&trim(request.form("filter_username"))&"%'"
				obj.Open SQL, Connect, 3,3
				if obj.eof then
				response.write "Couldn't find any users that match the search criteria"
				else%>
				
		<table>
			<td >Name</td>
			<td  align="center">Status</td>
			<td  align="center">Add result</td>
			<% do until obj.eof %>
		<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">
            <td class="text"><% =obj("user_firstname") & " " & obj("user_lastname") %></td>
            <td class="text" align="center"><% if cbool(obj("user_status")) = false then response.write "<font color='green'>Online</font>" else if cbool(obj("user_status")) = true then response.write "<font color='red'>Offline</font>"%></td>
			<td class="text" align="center" width="60"><a href="?user=<% =obj("id_user")%>"><img src="images/addresults.gif" alt="Add results" border="0"></a></td>
        </tr>
				<% obj.movenext
				loop%>
		</table>
		<%	END IF
			obj.close%>
	<% ELSEif request.querystring("alt")="" THEN%>
      <form method="post" name="frmSubject" action="?alt=step1&user=<%=user%>" onsubmit="return(checkSubject());">	  

		<table>
			<TD class="text">Subject Completed</TD>
			<TD class="text"><select name="session_subject" style="width:180px;">
				<option value="">
				<%
					SQL = "SELECT * FROM subjects"
					obj.Open SQL, Connect, 3,3
					do until obj.eof %>
					<option value="<% =obj("id_subject")%>"> <% =obj("subject_name")%>
					<% obj.movenext
					loop
				obj.close%>
				</select>
			</TD>
			<TR>
				<TD colspan="2" align="right"><input type="Submit" name="Submit" value="Choose subject" class="quiz_button"></TD>
			</TR>
		</table>
		
      </form>
	  <% ELSEif request.querystring("alt")="step1" THEN
	  
				SQL = "SELECT * FROM subjects WHERE id_subject = "&clng(request.form("session_subject"))&""
				obj.Open SQL, Connect, 3,3
				if obj.eof then%>
				No subject found!
				<% ELSE%>
				<span class="table_hl"><% =obj("subject_name")%></span>
	    <form method="post" name="frmSubject" action="?alt=save&user=<%=user%>" onsubmit="return(checkForm());">	  
		<input name="session_subject" type="hidden" value="<% =request.form("session_subject")%>">
		<table>
		<TR>
			<TD class="text" width="120">Date</TD>
			<TD class="text"><input type="Text" name="Session_date" id="date1" class="date-pick" autocomplete="off"></TD>
		</TR>
		<TR>
			<TD class="text">Path</TD>
			<TD class="text"><select name="session_version" style="width:60px;">
			<option value="">
			<option value="1"> 1
			<option value="2"> 2
			<option value="3"> 3
			<option value="4"> 4
			<option value="5"> 5
			</select></TD>
		</TR>
		<TR>
			<TD class="text">Score</TD>
			<TD class="text"><input type="Text" name="session_correct" value="" style="width:40px;text-align:center;" onkeypress="return isNumberKey(event)"> of <input type="Text" name="session_total" value="" style="width:40px;text-align:center;" onkeypress="return isNumberKey(event)"></TD>
		</TR>
		<TR>
			<TD class="text">Percent required</TD>
			<TD class="text"><input type="Text" name="percentage_required" value="<% =obj("subject_passmark")%>" style="width:40px;text-align:center;" onkeypress="return isNumberKey(event)">%</TD>
		</TR>
		
		<TR>
			<TD class="text" colspan="2" align="right"><input type="Submit" name="Submit" value="Add result" class="quiz_button"></TD>
		</TR>
		</table>
		
      </form>
	  <% obj.close
	  END IF
	   ELSEif request.querystring("alt")="save" THEN 
on error resume next
SQL = "INSERT INTO q_session (Session_users,Session_subject,Session_date,Session_total,Session_done,Session_correct,Session_stop,Session_finish,Session_version) VALUES (" 
SQL = SQL & ""&trim(clng(user))&"," 
SQL = SQL & ""&trim(clng(request.form("session_subject")))&"," 
SQL = SQL & "'"&cDateSql(trim(request.form("Session_date")))&"'," 
SQL = SQL & ""&trim(request.form("session_total"))&"," 
SQL = SQL & "1," 
SQL = SQL & ""&trim(request.form("session_correct"))&"," 
SQL = SQL & ""&trim(request.form("session_total"))&"," 
SQL = SQL & "'"&cDateSql(trim(request.form("Session_date")))&"'," 
SQL = SQL & ""&trim(request.form("session_version"))&"" 
SQL = SQL & ") "
'response.write SQL
'response.end
obj.Open SQL, Connect, 3,3 

if err.number <> 0 then
		response.write "The data could not be saved! Please try again"
else
	
dim score_total, score_user,percent
percent = clng(request.form("percentage_required"))
score_total = clng(request.form("session_total"))
score_user = clng(request.form("session_correct"))

score_percent = 0
score_percent = (score_user/score_total)*100
IF score_percent > percent THEN passed = 1 ELSE passed = 0

' GET THE LAST ID USED
SQL = "SELECT TOP 1 MAX(id_session) FROM q_session"
obj.Open SQL, Connect, 3,3 
id_current = obj(0)
obj.close

SQL = "INSERT INTO q_certification (q_session,quiz_date,expiry_date,passed,percentage_achieved,percentage_required) VALUES (" 
SQL = SQL & ""&trim(clng(id_current))&"," 
SQL = SQL & "'"&cDateSql(trim(request.form("Session_date")))&"'," 
SQL = SQL & "'"&cDateSql(DateAdd("d",364,cdate(trim(request.form("Session_date")))))  &"'," 
SQL = SQL & ""&passed&"," 
SQL = SQL & ""&formatnumber(score_percent,0)&"," 
SQL = SQL & ""&percent&"" 
SQL = SQL & ") "
obj.Open SQL, Connect, 3,3 
response.redirect "?alt=slut&user="&user&""
END IF

	   %>

	   
	   <%
	   
 ELSEif request.querystring("alt")="slut" THEN %>	   

<span class="table_hl">The result has now been saved to the database</span><br><br>
<a href="?user=<% =user%>">Add another result</a>
&nbsp;&nbsp;
|
&nbsp;&nbsp;
<a href="#" onClick="CloseAndRefresh(); return false;">Close this window</A>
<br><br>
 <%
	  END IF%>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
</BODY>
</HTML>
<%
call log_the_page ("Quiz List Add results")
set obj = nothing


%>

