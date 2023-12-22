<%@LANGUAGE="VBSCRIPT"%>

<% 

'Response buffer is used to buffer the output page. That means if any database exception occurs the contents can be cleared without processed any script to browser
 Response.Buffer = True
 
' "On Error Resume Next" method allows page to move to the next script even if any error present on page whcich will be caught after processing all asp script on page
 On Error Resume Next
 
'Changed by PR on 25.02.16
 %>
 
<!--#include file="connections/bbg_conn.asp"-->
<!--#include file="connections/include.asp"-->
<!--#include file="sha256.asp"-->
<%
' SqlCheckInclude.asp file provides sanitisation of user input to provide protection agains SQL injection
%>
<!--#include file="SqlCheckInclude.asp"-->

<%
if request.querystring("alt")="registered" THEN

'Our Function generatePassword accepts one parameter 'passwordLength'
'passwordLength will obviously determine the password length.
'The aplhanumeric character set is assigned to the variable sDefaultChars
Function generatePassword(passwordLength, passwordType)
	'Declare variables
	Dim sDefaultChars
	Dim iCounter
	Dim sMyPassword
	Dim iPickedChar
	Dim iDefaultCharactersLength
	Dim iPasswordLength
	'Initialize variables
	if passwordType = 1 then
		sDefaultChars="ABCDEFGHIJKLMNOPQRSTUVXYZacdefghijklmnopqrstuvwyz0123456789"
	else
		sDefaultChars="0123456789"
	end if
	iPasswordLength=passwordLength
	iDefaultCharactersLength = Len(sDefaultChars)
	Randomize'initialize the random number generator
	'Loop for the number of characters password is to have
	For iCounter = 1 To iPasswordLength
		'Next pick a number from 1 to length of character set
		iPickedChar = Int((iDefaultCharactersLength * Rnd) + 1)
		'Next pick a character from the character set using the random number iPickedChar
		'and Mid function
		sMyPassword = sMyPassword & Mid(sDefaultChars,iPickedChar,1)
	Next
	generatePassword = sMyPassword
End Function
'end random password function

first_name = trim(Server.HTMLEncode(Request.form("first_name")))
last_name = trim(Server.HTMLEncode(Request.form("last_name")))
username = trim(Server.HTMLEncode(Request.form("username")))
'password = trim(Request.form("password"))
password = generatePassword(8,1)
email = trim(Server.HTMLEncode(Request.form("email")))
business = Request.form("info1")
site = Request.form("info2")
activities = Request.form("info3")
company = Request.form("info4")


'(+) 130501 CD (START)
'validation of the aphostropy mark in all fields
first_name = Replace(first_name,"'","''")
last_name = Replace(last_name,"'","''")
usernameforemail = username
username = Replace(username,"'","''")
email = Replace(email,"'","''")
business = Replace(business,"'","''")
site = Replace(site,"'","''")
activities = Replace(activities,"'","''")
company = Replace(company,"'","''")
'(+) 130501 CD (FINISH)



'Inserts new registered user
' Err.Number is a attribute of "On Error Resume Next" method
' It is used to terminate any database query or transaction to provide protection against data integrity
' Changed by PR 23.02.16
if Err.Number = 0 then
Set CheckDupe = Server.CreateObject("ADODB.Recordset")
	CheckDupe.ActiveConnection = connect
	CheckDupe.Source = "SELECT * FROM q_user WHERE user_username='"&username&"'OR user_email='"&email&"'"	
	'OR user_email='"&email&"'
	CheckDupe.CursorType = 0
	CheckDupe.CursorLocation = 3
	CheckDupe.LockType = 3
	CheckDupe.Open()
	CheckDupe_numRows = 0 
	
end if
if CheckDupe.EOF And CheckDupe.BOF then

'Inserts new registered user
if Err.Number = 0 then
Set MM_editCmd = Server.CreateObject("ADODB.Command")
	MM_editCmd.ActiveConnection = connect
	user_logcount = 1
	user_access= cDateSql(Now())
	user_IP=Request.ServerVariables("REMOTE_ADDR")
	Dim strSQL
	strSQL = "Set Nocount on "
	strSQL = strSQL + "INSERT INTO q_user(user_lastname,user_firstname,user_username,user_city,user_password,user_info1,user_info2,user_info3,user_info4,user_email,user_IP,user_access,user_logcount,user_active) values ('"&last_name&"','"&first_name&"','"&username&"','"&password&"','1','"&business&"','"&site&"','"&activities&"','"&company&"','"&email&"','"&user_ip&"',GETDATE(),"&user_logcount&",1)"
	strSQL = strSQL + " select IdentityInsert=@@identity"
	strSQL = strSQL + " set nocount off" 
end if

	
	MM_editCmd.CommandText = strSQL
	
	Dim uid	
	set rs=MM_editCmd.Execute
	
	uid = rs("IdentityInsert")
	
	
	MM_editCmd.ActiveConnection.Close
	
	set  MM_editCmd= nothing
	'if MM_editCmd.EOF or MM_editCmd.BOF then response.write("Username already exists!")
	
	
	
Dim pass
Dim salt
salt = email
pass=password&salt
pass=sha256(pass)


if Err.Number = 0 then
SQL="update q_user set user_city=? WHERE ID_User=?"
set objCommand = Server.CreateObject("ADODB.Command") 
objCommand.ActiveConnection = Connect
objCommand.CommandText = SQL 
objCommand.Parameters(0).value = pass
objCommand.Parameters(1).value=uid
objCommand.Execute()
end if
	
'Adds subject to user profile
'First we must get the user
if Err.Number = 0 then
set user = Server.CreateObject("ADODB.Recordset")
user.ActiveConnection = connect
user.Source = "SELECT * FROM q_user WHERE user_firstname='"&first_name&"' and user_lastname='"&last_name&"' and user_username='"&username&"' and user_city='"&password&"'"
user.CursorType = 0
user.CursorLocation = 3
user.LockType = 3
user.Open()
user_numRows = 0 
end if
'Then we add the subject to that user




set RSActiveSubjects= Server.CreateObject("ADODB.Recordset") :RSActiveSubjects.ActiveConnection = Connect
RSActiveSubjects.Source = "SELECT *  FROM subjects where subject_active_q=1"
RSActiveSubjects.CursorType = adOpenForwardOnly : RSActiveSubjects.CursorLocation = 3 : RSActiveSubjects.LockType = 3 : RSActiveSubjects.Open()  
Dim totalSubjects
totalSubjects = RSActiveSubjects.RecordCount
Set MM_subjects = Server.CreateObject("ADODB.Command")
MM_subjects.ActiveConnection = connect
if not RSActiveSubjects.eof then arrV = RSActiveSubjects.GetRows ELSE arrV = -1
RSActiveSubjects.close 
IF IsArray(arrV) THEN
	For i = 0 to ubound(arrV,2)
	
	
	
	MM_insertsub="insert into subject_user(ID_subject,ID_user) values ("& arrV(0,i) & ",'"&uid&"')"
	MM_subjects.CommandText = MM_insertsub
	MM_subjects.Execute
	'response.write(arrV(0,i))
	
	NEXT
	 END IF
			   
		 Erase arrV

MM_subjects.ActiveConnection.Close
user.close
 
	if email > "" Then
	'Emails the user

		strMessage = "<html><head>"
		strMessage = strMessage + "<style type=text/css>"
		strMessage = strMessage + "<!--"
		strMessage = strMessage + ".storbox {width: 580px !important;width /**/:580px;border:1px solid #cccccc; padding:6px; margin-bottom:0px; background-color:#ffffff; }"
		strMessage = strMessage + "td, tr, div {font-family: Arial, Helvetica, sans-serif;font-size: 12px;color: #4B4B4B;}"
		strMessage = strMessage + "body {font-family: Arial, Helvetica, sans-serif;font-size: 12px;color: #4B4B4B;}"
		strMessage = strMessage + "A:link, a:visited, a:active {text-decoration: underline;font-weight:bold;color: #4B4B4B}"
		strMessage = strMessage + "A:hover {text-decoration: none;}"
		strMessage = strMessage + ".small{font-size: 11px;}"
		strMessage = strMessage + ".smallgrey{font-size: 11px;color:#AAB7A9;}"
		strMessage = strMessage + "-->"
		strMessage = strMessage + "</style></head>"
		strMessage = strMessage + "<body><h2>"&client_name_long&"</h2>"
		strMessage = strMessage + "<h3>Better Business Program - Registration</h3><br><table width=100% border=0 cellpadding=1 cellspacing=1 bordercolor=#111111>"
		strMessage = strMessage + "<tr><td>Thank you for registering with the Better Business Program.  Please find your login details below:</td></tr>"
		strMessage = strMessage + "<tr><td>&nbsp;</td></tr>"
		strMessage = strMessage + "<tr><td>Username: " + username + "</td></tr>"
		strMessage = strMessage + "<tr><td>Password: " + password + "</td></tr>"
		strMessage = strMessage + "<tr><td>&nbsp;</td></tr>"
		strMessage = strMessage + "<tr><td>Click on this link to take you to the Better Business Program: <a href="& client_homepage &">"&client_homepage&"</a></td></tr>"
		strMessage = strMessage + "</table></body></html>"

	'(+) 130308 CD (START)
	'Const client_homepage        = _"http://www.lawofthejungle.com/acme33testa/index.asp"
	'(+) 130308 CD (FINISH)
	
	Const cdoSendUsingMethod        = _
		"http://schemas.microsoft.com/cdo/configuration/sendusing"
	Const cdoSendUsingPort          = 1
	Const cdoSMTPServer             = _
		"http://schemas.microsoft.com/cdo/configuration/smtpserver"
	Const cdoSMTPServerPort         = _
		"http://schemas.microsoft.com/cdo/configuration/smtpserverport"
	Const cdoSMTPConnectionTimeout  = _
		"http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
	Const cdoSMTPAuthenticate       = _
		"http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
	Const cdoBasic                  = 1
	Const cdoSendUserName           = _
		"http://schemas.microsoft.com/cdo/configuration/sendusername"
	Const cdoSendPassword           = _
		"http://schemas.microsoft.com/cdo/configuration/sendpassword"

    	Set cdoConfig = Server.CreateObject("CDO.Configuration")
			Set Fields = cdoConfig.Fields

			' Set config fields we care about
			With Fields
				.Item(cdoSendUsingMethod)       = cdoSendUsingPort
				.Item(cdoSMTPServer)            = "4493HVIRT"
				.Item(cdoSMTPServerPort)        = 25
				.Item(cdoSMTPConnectionTimeout) = 10
				.Item(cdoSMTPAuthenticate) = cdoBasic
				.Item(cdoSendUserName) = "lotjb-system@lotj.com" 'this is my name in smtp.163.com
				.Item(cdoSendPassword) = "lotjb#100" 'my password
				.Update
		End With

    	Set cdoMessage = CreateObject("CDO.Message")
    	With cdoMessage
    	    Set .Configuration = cdoConfig
    	    .From = "lotjb-system@lotj.com"
    	    .To = email
    	    .Subject = "Better Business Program - Registration"
    	    .HTMLBody = strMessage
    	    .Send
    	End With
	
	elseif CheckDupe.Fields.Item("user_email").Value = email or CheckDupe.Fields.Item("user_username").Value = username then 
	CheckDupe.Close
	end if
end if
end if

%>

<% 
if Err.Number = 0 then
set info2 = Server.CreateObject("ADODB.Recordset")
info2.ActiveConnection = Connect
'When the page is refreshed the code below will fetch the data from info2 table according which business was chosen. The SQL line below is what fetches the data from info2 table.
if request("info1")<> "" then
	info2_prm = request("info1")
else
	info2_prm = 0
end if
info2.Source = "SELECT * FROM q_info2 where info2_info1 =" & info2_prm &" and info2_active=1 order by info2"
info2.CursorType = 0
info2.CursorLocation = 3
info2.LockType = 3
info2.Open()
info2_numRows = 0



set info1 = Server.CreateObject("ADODB.Recordset")
info1.ActiveConnection = Connect
info1.Source = "SELECT * FROM q_info1 where info1_active=1 order by info1"
info1.CursorType = 0
info1.CursorLocation = 3
info1.LockType = 3
info1.Open()
info1_numRows = 0



set info3 = Server.CreateObject("ADODB.Recordset")
info3.ActiveConnection = connect
info3.Source = "SELECT * FROM q_info3 where info3_active=1 order by info3"
info3.CursorType = 0
info3.CursorLocation = 3
info3.LockType = 3
info3.Open()
info3_numRows = 0


set info4 = Server.CreateObject("ADODB.Recordset")
info4.ActiveConnection = connect
info4.Source = "SELECT * FROM q_info4 where info4_active=1 order by info4"
info4.CursorType = 0
info4.CursorLocation = 3
info4.LockType = 3
info4.Open()
info4_numRows = 0

end if


%>

<!doctype html>
<head>
<meta charset="UTF-8">
<title><%=client_name_short%> - Register</title>
<script src="jquery-1.11.1.js?v=bbp34"></script>
<link rel="stylesheet" href="style/bbp_acme34.css" type="text/css">
<link rel="stylesheet" href="//code.jquery.com/mobile/1.4.5/jquery.mobile-1.4.5.min.css" />
<link rel="stylesheet" type="text/css" href="js/sweet-alert.css">
    		  <script src="js/sweet-alert.min.js?v=bbp34"></script>
<style>
	p,h1,h2,h3,td,a {font-family: Arial}
	.a,.label {font-size: 0.8em}
	.label {width: 130px}
</style>
<script >
<!-- This is where the checkform is run when the business dropdown is changed.
function checkform() {
	document.forms[0].action="self-register.asp"
	document.forms[0].target="_self"
	document.forms[0].submit()
}
//-->
</script>

<script >
function trySubmit()
{

	//var y=document.forms["register"]["email"].value;
	//if (y > '')
	/* Email Validation */
	//{
	var x=document.forms["register"]["email"].value;
	//var atpos=x.indexOf("@gwf");
	//var dotpos=x.lastIndexOf(".com.au");
	//var gwf=x.indexOf("@gwf.com.au");
	var lotj=x.indexOf("@lotj.com");
		//if (atpos<1 || dotpos<atpos+2 || dotpos+2>=x.length)
		if (lotj<1)
		{
		
		  swal({   title: "Sorry, you must enter a valid e-mail address",   text: "",   type: "error",   confirmButtonText: "OK",html: true });
		  return false;
		}
	//}

	if (document.forms[0].first_name.value.length<2 || document.forms[0].first_name.value.length>50)
	{
		swal({   title: "Sorry, you must enter a first name!<br>(min. 2 characters and max. 50 characters)",   text: "",   type: "error",   confirmButtonText: "OK",html: true });
		return false;
	}
	
	if (document.forms[0].last_name.value.length<2 || document.forms[0].last_name.value.length>50)
	{
		
		swal({   title: "Sorry, you must enter a last name!<br>(min. 2 characters and max. 50 characters)",   text: "",   type: "error",   confirmButtonText: "OK",html: true });
		return false;
	}
	if (document.forms[0].username.value.length<2 || document.forms[0].username.value.length>50)
	{
		
		swal({   title: "Sorry, you must enter a username!<br>(min. 2 characters and max. 50 characters)",   text: "",   type: "error",   confirmButtonText: "OK",html: true });
		return false;
	}
	
		if (document.forms[0].email.value.length<2 || document.forms[0].email.value.length>70)
	{
		
		swal({   title: "Sorry, you must enter an email!<br>(max. 70 characters)",   text: "",   type: "error",   confirmButtonText: "OK",html: true });
		return false;
	}
	//if (document.forms[0].password.value.length<2)
	//{
	//	alert("Sorry, you must enter a password!\n(min. 2 characters)");
	//	return false;
	//}
	if (document.forms[0].info1.selectedIndex==0)
	{
		
		swal({   title: "Sorry, you must select a business",   text: "",   type: "error",   confirmButtonText: "OK",html: true });
		return false;
	}
	//if (document.forms[0].info2.selectedIndex==0)
	//{
	//	alert("Sorry, you must select a site");
	//	return false;
	//}
	if (document.forms[0].info3.selectedIndex==0)
	{
		
		swal({   title: "Sorry, you must select an activity",   text: "",   type: "error",   confirmButtonText: "OK",html: true });
		return false;
	}
	//if (document.forms[0].info4.selectedIndex==0)
	//{
	//	alert("Sorry, you must select a company");
	//	return false;
	//}
	else{
	swal({   title: "Are you sure you would like to proceed in registering?",   text: "",   type: "warning",   showCancelButton: true,   confirmButtonColor: "#DD6B55",   confirmButtonText: "Yes",   closeOnConfirm: false }, function(){  document.forms[0].submit(); return false; });
}
	/*if (confirm("Are you sure you would like to proceed in registering?"))	
	{	
	document.forms[0].submit();
	return false;
	}*/
return false;
}

</script>
<% 
if request.querystring("alt")="registered" THEN 
if CheckDupe.EOF And CheckDupe.BOF then
%>
<script > 
	<!-- 
	setTimeout("self.close();",5000) 
	//--> 
</script>
<% elseif CheckDupe.Fields.Item("user_email").Value = email  or CheckDupe.Fields.Item("user_username").Value = username then  %>

<% 
end if
end if 
%>
</head>

<body>
<div >
<form name="register" id="registration_form" method="post" action="self-register.asp?alt=registered" onsubmit="return trySubmit();">
	<table style="text-align:left; width:100%;" >
		
		<tr>
			<td style="background:#000; width:100%;" colspan=2><a href="index.asp" target="_top"><img src="images/news/bbp-page_01.gif" alt="" height=25 ></a></td>
		</tr>
<% if request.querystring("alt")="registered" THEN %>
		<tr>
			<td colspan=2><p style="text-align: center;">&nbsp;</p></td>
		</tr>
			<% if CheckDupe.EOF And CheckDupe.BOF then %>
				<tr>
					<td colspan=2><p style="text-align:center;font-weight:bold">Thank you for registering</p></td>
				</tr>
				<% if email > "" Then %>
				<tr>
					<td colspan=2><p style="text-align:center;font-weight:bold">You should receive an email shortly</p></td>
				</tr>
				<% end if %>
				<tr>
					<td colspan=2><p style="text-align:center;font-weight:bold">You may close this window.</p></td>
				</tr>
			<% elseif CheckDupe.Fields.Item("user_email").Value = email or CheckDupe.Fields.Item("user_username").Value = username then  response.write ("<script>swal({   title: ""Somebody already has this username or email address. Please try another."",   text: """",   type: ""error"",   confirmButtonText: ""OK"",html: true, closeOnConfirm: false }, function(){  window.location.href='self-register.asp'; return false; }); </script>") %>
			<tr>
				<td colspan=2><p style="text-align:center;color:red; font-weight:bold">Somebody already has this username or email address. Please try another.</p></td>
			</tr>	
			<% end if %>
<% else %>
		<tr>
			<td  colspan=2 style="text-align:center"><h1>Registration Form </h1></td>
		</tr>
		<tr>
		<!-- for each field you must have a request(field item) so when the business is selected the data will stay -->
			<td class="label">First Name:</td><td><input type="text" name="first_name" size="28" value="<%=request("first_name")%>"></td>
		</tr>
		<tr>
			<td class="label">Last Name:</td><td><input type="text" name="last_name" size="28" value="<%=request("last_name")%>"></td>
		</tr>
		<tr>
			<td class="label">Username:</td><td><input type="text" name="username" size="28" value="<%=request("username")%>"></td>
		</tr>
		<!--<tr>
			<td class="label">Password:</td><td><input type="text" name="password" size="28" ></td>
		</tr>-->
		<tr>
			<td class="label">Email Address:</td><td><input type="text" name="email" size="28" value="<%=request("email")%>"></td>
		</tr>
		<!--<tr>
			<%
			'pn 050720 pull out all active subjects, currently only in reference to guide
			set subjects = Server.CreateObject("ADODB.Recordset")
			subjects.ActiveConnection = Connect
			subjects.Source = "SELECT subjects.ID_subject, subjects.subject_name  FROM (subjects INNER JOIN b_topics ON subjects.ID_subject = b_topics.topic_subject) INNER JOIN b_pages ON b_topics.ID_topic = b_pages.page_topic  GROUP BY subjects.ID_subject, subjects.subject_name, subjects.subject_ord, subjects.ID_subject, Abs([subject_active_b]), Abs([topic_active]), Abs([page_active])  HAVING (((Abs([subject_active_b]))=1) AND ((Abs([topic_active]))=1) AND ((Abs([page_active]))=1))  ORDER BY subjects.subject_ord, subjects.ID_subject;"
			subjects.CursorType = 0
			subjects.CursorLocation = 2
			subjects.LockType = 3
			subjects.Open()
			subjects_numRows = 0

			While (NOT subjects.EOF)
			%>
			<td class="label"><%=subjects.Fields.Item("subject_name").Value%></td>
			<td class="label"><input type="checkbox" checked  name="user_subject|0|<%=subjects.Fields.Item("ID_subject").Value%>" /></td>
			<%subjects.MoveNext()
			Wend
			subjects.Close()
			%>
		</tr>-->
		<tr>
		<td class="label">Business:</td>
			<td>
			<select name="info1" onChange="checkform();" class="buttonz">
				<option value="0" >...Select a Business...</option>
				<% While (NOT info1.EOF) %>
				<!-- the if statement on the line below selects the data that was chosen before the page was refreshed -->
				<option value="<%=(info1.Fields.Item("ID_info1").Value)%>" <%if (CStr(info1.Fields.Item("ID_info1").Value) = CStr(request("info1"))) then Response.Write("SELECTED") : Response.Write("")%>><%=(info1.Fields.Item("info1").Value)%></option>
				<%
					info1.MoveNext()
					Wend
					info1.Requery
				%>
			</select>
			</td>
		</tr>
		<tr>
			<td class="label">Site:</td>
			<td>
				<select name="info2" onChange="change=true;" class="buttonz">
				<option value="0" selected>..Select Site...</option>
				<%While (NOT info2.EOF)%>
				<option value="<%=(info2.Fields.Item("ID_info2").Value)%>" <%if (CStr(info2.Fields.Item("ID_info2").Value) = CStr(request("info2"))) then Response.Write("SELECTED") : Response.Write("")%>><%=(info2.Fields.Item("info2").Value)%></option>
				<%
					info2.MoveNext()
					Wend
					info2.Requery
				%>
				</select>
			</td>
		</tr>
		<tr>
			<td class="label">Business Activity:</td>
			<td>
				<select name="info3" class="buttonz">
				<option value="0">...Select a Business Activity...</option>     
					<%
					While (NOT info3.EOF)
					%>
					<option value="<%=(info3.Fields.Item("ID_info3").Value)%>" <%if (CStr(info3.Fields.Item("ID_info3").Value) = CStr(request("info3"))) then Response.Write("SELECTED") : Response.Write("")%>><%=(info3.Fields.Item("info3").Value)%></option>
					<%
					info3.MoveNext()
					Wend
					info3.Requery
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td class="label">Company:</td>
			<td>
				<select name="info4" class="buttonz">
					<%
					While (NOT info4.EOF)
					%>
					<option value="<%=(info4.Fields.Item("ID_info4").Value)%>" <%if (CStr(info4.Fields.Item("ID_info4").Value) = CStr(request("info4"))) then Response.Write("SELECTED") : Response.Write("")%>><%=(info4.Fields.Item("info4").Value)%></option>
					<%
					info4.MoveNext()
					Wend
					info4.Requery
					%>
				</select>
			</td>
		</tr>

		<tr>
		<td class="label">&nbsp;</td>
			<td>
			<br>
				<input type="hidden" name="MM_insert" value="true">
				<input type="submit" class="rbutton" onClick="return CheckStringforSQL(str); trySubmit();" value="Register"> 
				
				<!--<a href="javascript:" onclick="window.open('privacy.asp','message','scrollbars=yes,resizable=yes,width=600, height=500, top=100, left=300')">Privacy Statement</a>-->
			</td>
		</tr>
<% END IF 
info1.Close()
info3.Close()
info4.Close()
%>
<% 
if request("username_req") <> "" then
	comment="Registered Username: "& request("username")
else
	comment="Registration"
end if
call log_the_page ("Registration", "0", "n/a", "0", "n/a", "0", "n/a", comment)
%>
		
	</table>
	</form>
	</div>
</body>
</html>
<!-- #include file = "errorhandler/index.asp"-->



