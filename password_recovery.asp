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

<%
Session("logo") = "client_logotype"
Session("font") = "arial"
not_found=-1


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
		sDefaultChars="ABCDEFGHIJKLMNOPQRSTUVXYZ0123456789"
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

%>

<% if request("username_req") <> "" then
if Err.Number = 0 then
Set obj = Server.CreateObject("ADODB.Recordset")
SQL="SELECT * FROM q_user WHERE upper(user_username)='"&UCASE(request("username_req"))&"'"
obj.ActiveConnection = Connect
obj.Source = SQL
obj.CursorType = 0
obj.CursorLocation = 3
obj.LockType = 3
obj.Open
end if

if not obj.eof then
	not_found=0
	if obj("user_email")<>"" then
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
		strMessage = strMessage + "<h3>BBP Password Recovery</h3><br><table width=100% border=0 cellpadding=1 cellspacing=1 bordercolor=#111111>"
		strMessage = strMessage + "<tr><td>You have requested to receive your username and password from the "&client_name_long&" BBP.  Please find these details below.</td></tr>"
		strMessage = strMessage + "<tr><td>Click on the link to take you to the BBP: <a href="&client_homepage&">"&client_homepage&"</a></td></tr>"
		strMessage = strMessage + "<tr><td>&nbsp;</td></tr>"
		strMessage = strMessage + "<tr><td>Username: " + obj("user_username") + "</td></tr>"

		' reset password to random
			    random_value = generatePassword(8,1)
				password_to_send=random_value
if Err.Number = 0 then
			    Set MM_editCmd = Server.CreateObject("ADODB.Command")
			    MM_editCmd.ActiveConnection = Connect
			    MM_editCmd.CommandText = "UPDATE q_user SET user_city = '"&password_to_send&"', user_password = 1 WHERE upper(user_username)='"&UCASE(obj("user_username"))&"'"
			    MM_editCmd.Execute
end if
    			MM_editCmd.ActiveConnection.Close
		
		'Select Case password_recovery_type
		'	Case 1   ' send current password
		'		password_to_send=obj("user_city")
		'	Case 2   ' reset password to null
		'	    Set MM_editCmd = Server.CreateObject("ADODB.Command")
		'	    MM_editCmd.ActiveConnection = Connect
		'	    MM_editCmd.CommandText = "UPDATE q_user SET user_city = null WHERE upper(user_username)='"&UCASE(obj("user_username"))&"'"
		'	    MM_editCmd.Execute
    	'		MM_editCmd.ActiveConnection.Close
		'		password_to_send="Your password is what you enter the next time you log in"
		'	Case 3   ' reset password to random
		'	    random_value = generatePassword(8,1)
		'		password_to_send=random_value
		'	    Set MM_editCmd = Server.CreateObject("ADODB.Command")
		'	    MM_editCmd.ActiveConnection = Connect
		'	    MM_editCmd.CommandText = "UPDATE q_user SET user_city = '"&password_to_send&"' WHERE upper(user_username)='"&UCASE(obj("user_username"))&"'"
		'	    MM_editCmd.Execute
    	'		MM_editCmd.ActiveConnection.Close
		'	Case 4   ' reset password to set pattern
		'		password_to_send=password_set_pattern
		'	    Set MM_editCmd = Server.CreateObject("ADODB.Command")
		'	    MM_editCmd.ActiveConnection = Connect
		'	    MM_editCmd.CommandText = "UPDATE q_user SET user_city = '"&password_to_send&"' WHERE upper(user_username)='"&UCASE(obj("user_username"))&"'"
		'	    MM_editCmd.Execute
    	'		MM_editCmd.ActiveConnection.Close
		'	Case 5   ' reset password to set pattern + random numbers
		'	    random_value = generatePassword(4,2)
		'		password_to_send=password_set_pattern&random_value
		'	    Set MM_editCmd = Server.CreateObject("ADODB.Command")
		'	    MM_editCmd.ActiveConnection = Connect
		'	    MM_editCmd.CommandText = "UPDATE q_user SET user_city = '"&password_to_send&"' WHERE upper(user_username)='"&UCASE(obj("user_username"))&"'"
		'	    MM_editCmd.Execute
    	'		MM_editCmd.ActiveConnection.Close
		'	Case 6   ' reset password to username
		'	    password_to_send=obj("user_username")
		'	    Set MM_editCmd = Server.CreateObject("ADODB.Command")
		'	    MM_editCmd.ActiveConnection = Connect
		'	    MM_editCmd.CommandText = "UPDATE q_user SET user_city = '"&password_to_send&"' WHERE upper(user_username)='"&UCASE(obj("user_username"))&"'"
		'	    MM_editCmd.Execute
    	'		MM_editCmd.ActiveConnection.Close
		'	Case 7   ' reset password to username + random numbers
		'	    random_value = generatePassword(4,2)
		'	    password_to_send=obj("user_username")&random_value
		'	    Set MM_editCmd = Server.CreateObject("ADODB.Command")
		'	    MM_editCmd.ActiveConnection = Connect
		'	    MM_editCmd.CommandText = "UPDATE q_user SET user_city = '"&password_to_send&"' WHERE upper(user_username)='"&UCASE(obj("user_username"))&"'"
		'	    MM_editCmd.Execute
    	'		MM_editCmd.ActiveConnection.Close
		'end Select



		strMessage = strMessage + "<tr><td>Password: " + password_to_send + "</td></tr>"
		strMessage = strMessage + "<tr><td>&nbsp;</td></tr>"
		strMessage = strMessage + "</table></body></html>"

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
    	    .To = obj("user_email")
    	    .Subject = "BBP Password Recovery"
    	    .HTMLBody = strMessage
    	    .Send
    	End With
 	%><%
 	   Set cdoMessage = Nothing
 	   Set cdoConfig = Nothing
 	   not_found=0
	else
		not_found=2
	end if
else
	not_found=1
end if

obj.close

end if %>

<html>
<head>
<title><%=client_name_short%> Better Business Program
<%if Session("MM_Username") <> "" then response.write(" - you are logged in as " & Session("firstname") & " " & Session("lastname"))%>
</title>
<link rel="stylesheet" href="styles/bbp_style_acme34.css" type="text/css">

</head>

<body class="none" style="background-image: none;">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
<td bgcolor="#000000" colspan=2><a href="index.asp" target="_top"><img src="images/news/bbp-page_01.gif" height=25 border=0></a></td>
</tr>
<tr>
<td height="10"  colspan=2><img src="images/news/spacer.gif" width=10 height=10></td>
</tr>
<tr>
<td width="10"><img src="images/news/spacer.gif" width=10 height=10></td>
<td align="center"><font class="home_heading">Password recovery</font>
</td>
</tr>
<tr>
<td height="10"  colspan=2><img src="images/news/spacer.gif" width=10 height=10></td>
</tr>
<form name="password_form" method="post" action="password_recovery.asp">
<tr>
<td width="10"><img src="images/news/spacer.gif" width=10 height=10></td>
<td><font class="focus_text">Please enter your username in the space below. Then click on the Submit button to recover your password.</font>
</td>
</tr>
<tr>
<td height="10"  colspan=2><img src="images/news/spacer.gif" width=10 height=10></td>
</tr>
<tr>
<td width="10"><img src="images/news/spacer.gif" width=10 height=10></td>
<td>

<font class="focus_text">Username:</font> <input type="text" name="username_req">
<input type="submit" name="submit" value="Submit">
</form>
</td>
</tr>
<tr>
<td height="10"  colspan=2><img src="images/news/spacer.gif" width=10 height=10></td>
</tr>
<tr>
<td width="10"><img src="images/news/spacer.gif" width=10 height=10></td>
<td>

<
<% if not_found=0 then%>
<font color="blue"><b>An email has been sent to your registered email address.</b></font><br><font class="focus_text">If you have not received an email, please check your Junk Mail folder.</font>
<% end if %>
<% if not_found=1 then%>
<font color="red"><b>The username has not been found in the system. Please try again.</b></font>
<% end if %>
<% if not_found=2 then%>
<font color="red"><b>A valid email address is not registered for this username.  Please try again.</b></font>
<% end if %>
</td></tr>
</table>
</body>
</html>
<%
if request("username_req") <> "" then
	comment="User: "& request("username_req")
else
	comment="Request password"
end if
call log_the_page ("password", "0", "n/a", "0", "n/a", "0", "n/a", comment)
%>
<!-- #include file = "errorhandler/index.asp"-->