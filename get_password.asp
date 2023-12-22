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
<% 
Function StringToHex(ByRef pstrString)
    	Dim llngIndex
    	Dim llngMaxIndex
    	Dim lstrHex
    	llngMaxIndex = Len(pstrString)
    	For llngIndex = 1 To llngMaxIndex
    		lstrHex = lstrHex & Right("0" & Hex(Asc(Mid(pstrString, llngIndex, 1))), 2)
    	Next
    	StringToHex = lstrHex
    End Function
    Function HexToString(ByRef pstrHex)
    	Dim llngIndex
    	Dim llngMaxIndex
    	Dim lstrString
    	llngMaxIndex = Len(pstrHex)
    	For llngIndex = 1 To llngMaxIndex Step 2
    		lstrString = lstrString & Chr("&h" & Mid(pstrHex, llngIndex, 2))
    	Next
    	HexToString = lstrString
    End Function
    Function URLDecode(str) 
        str = Replace(str, "+", " ") 
        For i = 1 To Len(str) 
            sT = Mid(str, i, 1) 
            If sT = "%" Then 
                If i+2 < Len(str) Then 
                    sR = sR & _ 
                        Chr(CLng("&H" & Mid(str, i+1, 2))) 
                    i = i+2 
                End If 
            Else 
                sR = sR & sT 
            End If 
        Next 
        URLDecode = sR 
    End Function 
 
    Function URLEncode(str) 
        URLEncode = Server.URLEncode(str) 
    End Function 
	
	
 
%>
<%
Private Function Encrypt(ByVal string)
    Dim x, i, tmp
    For i = 1 To Len( string )
        x = Mid( string, i, 1 )
        tmp = tmp & Chr( Asc( x ) + 1 )
    Next
    tmp = StrReverse( tmp )
    Encrypt = tmp
End Function
%>
<% if request("username_req") <> "" AND len(request("username_req")) < 50 then
Set obj = Server.CreateObject("ADODB.Recordset")
user_name = UCASE(request("username_req"))
SQL="SELECT * FROM q_user WHERE upper(user_username)='"& replace(user_name, "'", "''" ) &"' AND user_active=1"
obj.ActiveConnection = Connect
obj.Source = SQL
obj.CursorType = 0
obj.CursorLocation = 3
obj.LockType = 3
obj.Open

if not obj.eof then
	not_found=0
	if obj("user_email")<>"" then
		Dim sid
		sid = obj("ID_User")
		Dim all
		all=Encrypt(sid)
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
		strMessage = strMessage + "<h3>Better Business Program Password Recovery</h3><br><table width=100% border=0 cellpadding=1 cellspacing=1 bordercolor=#111111>"
		strMessage = strMessage + "<tr><td>You have requested to receive your username and password from the Better Business Program.  Please find these details below.</td></tr>"
		strMessage = strMessage + "<tr><td>Click on this link to reset your password <a href='"&client_homepage&"/reset_password.asp?uid="&URLEncode(StringToHex(all))&"'>Reset password</a></td></tr>"
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
'				.Item(cdoSMTPServer)            = "4493HVIRT"
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
    	    .Subject = "Better Business Program Password Recovery"
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
<% if request("email_req") <> "" AND len(request("email_req")) < 70  then
Set obj = Server.CreateObject("ADODB.Recordset")
email = UCASE(request("email_req"))
' Added fix to allow emails with apostrophe for username forgot functionality by PR 15.07.2016 JIRA:BBP-70
SQL="SELECT * FROM q_user WHERE upper(user_email)='"& replace(email, "'", "''" ) &"' AND user_active=1"
obj.ActiveConnection = Connect
obj.Source = SQL
obj.CursorType = 0
obj.CursorLocation = 3
obj.LockType = 3
obj.Open

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
		strMessage = strMessage + "<h3>Better Business Program Username Recovery</h3><br><table width=100% border=0 cellpadding=1 cellspacing=1 bordercolor=#111111>"
		strMessage = strMessage + "<tr><td>You have requested to receive your username from the Better Business Program.  Please find these details below.</td></tr>"
		strMessage = strMessage + "<tr><td>Click on the link to take you to the Better Business Program: <a href="&client_homepage&">"&client_homepage&"</a></td></tr>"
		strMessage = strMessage + "<tr><td>&nbsp;</td></tr>"
		strMessage = strMessage + "<tr><td><b>Username:</b> " + obj("user_username") + "</td></tr>"

		
		strMessage = strMessage + "<tr><td>&nbsp;</td></tr>"
		strMessage = strMessage + "</table></body></html>"

	

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
    	    .Subject = "Better Business Program Username Recovery"
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
	not_found=3
end if

obj.close

end if %>

<html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" >
<head>
<meta http-equiv="X-UA-Compatible" content="IE=9" />
<title><%=client_name_short%> - Username and Password Recovery </title>

 
 <!-- Latest compiled and minified CSS -->
/* <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.2/css/bootstrap.min.css"> */
/* <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-T3c6CoIi6uLrA9TneNEoa7RxnatzjcDSCmG1MXxSR1GAsXEV/Dwwykc2MPK8M2HN" crossorigin="anonymous">		 */
<link rel="stylesheet" href="style/slide.css" type="text/css">
   <script src="jquery-1.11.1.js?v=bbp34"></script>
 
<script type="text/javascript" src="js/jquery.accordion.js?v=bbp34"></script>
    <script type="text/javascript">
        $(document).ready(function() {
            $('.accordion').accordion(); //some_id section1 in demo
        });
    </script>
</head>

<body class="none" style="background-image: none;">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
	<td bgcolor="#000000" colspan=2><a href="index.asp" target="_top"><img src="images/news/bbp-page_01.gif" height=25 border=0></a></td>
</tr>
</table>
<div class="accordion" id="section1"><button class="btn btn-primary btn-default" id="passbut">Retrieve Password</button><span></span></div>
    <div class="container">
        <div class="content">
            <div class="table-responsive">
			<table  style="width:100%;"  border=0 cellspacing=2 cellpadding=2>
			<form name="password_form" method="post" action="get_password.asp">
		<tr>
			
			<td  ><font class="focus_text"><b>Please enter your username. </b></font><br><br></td>
		</tr>
		
		<tr>
			
			<td  >
				<font class="focus_text">Username:</font> <input type="text" name="username_req" size="12"> <input type="submit" name="submit" value="Submit">
			</td>
			
			
		</tr>
	</form>
	</table>
	</div>
            
        </div>
    </div>
    <div class="accordion" id="section2"><button class="btn btn-primary btn-default" id="userbut" >Retrieve Username</button><span></span></div>
    <div class="container">
        <div class="content">
            <div class="table-responsive">
			<table style="width:100%;" >
			<form name="password_form" method="post" action="get_password.asp">
		<tr>
			<td width="10"><img src="images/news/spacer.gif" width=10 height=10></td>
			<td><font class="focus_text"><b>Please enter your email address.</b></font></td>
		</tr>
		<tr>
			<td height="10"  colspan=2><img src="images/news/spacer.gif" width=10 height=10></td>
		</tr>
		<tr>
			<td width="10"><img src="images/news/spacer.gif" width=10 height=10></td>
			<td>
				<font class="focus_text">Email:</font> <input type="text" name="email_req">
				<input type="submit" name="submit" value="Submit">
			</td>
		</tr>
	</form>
	</table>
	</div>
            
        </div>
    </div>
	<div class="msg">
<% if not_found=0 then%>
<font color="blue"><b>An email has been sent to your registered email address.</b></font><br><font class="focus_text">If you have not received an email, please check your Junk Mail folder.</font>
<% end if %>
<% if not_found=1 then%>
<font color="red"><b>This username is not in the system or the length of the username is over the maximum character limit of 50 characters. Please try again</b></font>
<% end if %>
<% if not_found=2 then%>
<font color="red"><b>A valid email address is not registered for this username. Please try again.</b></font>
<% end if %>
<% if not_found=3 then%>
<font color="red"><b>This email address is not in the system or the length of the email address is over the maximum character limit of 70 characters. Please try again</b></font>
<% end if %>
</div>
</body>
</html>
<%
if request("username_req") <> "" then
	comment="User: "& request("username_req")
	
else if request("email_req") <> "" then

comment="Email: "& request("email_req")
else
	comment="Request password"
end if

end if
call log_the_page ("password", "0", "n/a", "0", "n/a", "0", "n/a", comment)
%>
<!-- #include file = "errorhandler/index.asp"-->