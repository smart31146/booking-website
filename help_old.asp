<%@LANGUAGE="VBSCRIPT"%>
<% 
'Response buffer is used to buffer the output page. That means if any database exception occurs the contents can be cleared without processed any script to browser
 Response.Buffer = True
 
' "On Error Resume Next" method allows page to move to the next script even if any error present on page whcich will be caught after processing all asp script on page
 On Error Resume Next
 
'Changed by PR on 25.02.16
%>

<!-- #include file = "connections/bbg_conn.asp" -->
<!-- #include file = "connections/include.asp"-->

<%
Session("qsave") = "" 
if Session("userID") = "" then response.redirect("error.asp?" & request.QueryString)
if NOT pref_quiz_avail then response.redirect("error.asp?" & request.QueryString)
%>

<!doctype html>
<head>

	<title><%=client_name_short%> - Better Business Program</title>
		<META name="DESCRIPTION"	content="">
		<!-- #include file = "inc_header.asp" -->
	<script type="text/javascript" src="admin/ckeditor/ckeditor.js?v=bbp34"></script>
	<script>
	function closeWin(thetime) {
		setTimeout("window.close()", thetime);
	}
	
' set maximum size of help box is set to 500 characters. PR-25.02.16	
function checkform() {
	var x = document.forms["feedback"]["details"].value;
	if (x == null || x == "" || x.length > 500){
		alert("Maximum 500 characters allowed for bug/feedback")
		return false;
	}
}
	</script>	
</head>

<body id="b" style="background: url('images/bg.jpg') no-repeat;width:auto;height:auto;position:relative;">
<div class="main_content" style="margin: 0 auto;">

		<div style="position:absolute;left:15px;top:15px;right:15px;color:#FFF;">

		<strong>LEGAL HELP</strong><br>
		</p>
		<br>
	
		<strong>TECHNICAL HELP</strong><br>
		Need technical support or assistance? Please contact:
		<p style="text-align: center;font-size:13px;" >
		
		</p>
		<br>
		
		<br>

		<strong>BUG/FEEDBACK</strong><br>
		If you have observed a bug or would like to comment on your experience, please use the form below.
		We welcome your feedback.
<%
if request("submit")="Submit" then
if err.Number = 0 then
 SQL="INSERT INTO  f_feedback (ID_user, details) values (?,?)"
	if Err.Number = 0 then
	set objCommand = Server.CreateObject("ADODB.Command") 
	objCommand.ActiveConnection = Connect
	objCommand.CommandText = SQL
	objCommand.Parameters(0).value = Session("userID") 
	objCommand.Parameters(1).value = Server.HTMLEncode(Request.Form("details"))

	Set obj = objCommand.Execute()
end if
	end if 
if (request("save_type") = "dbe") then
	

	set feedbackemail = Server.CreateObject("ADODB.Recordset")
	feedbackemail.ActiveConnection = Connect
	feedbackemail.Source = "SELECT f_address FROM f_email"
	feedbackemail.CursorType = 0
	feedbackemail.CursorLocation = 3
	feedbackemail.LockType = 3
	feedbackemail.Open()
	feedbackemail_numRows = 0

	subject = "Feedback Report"
	name = request("name")
	sender=request("email")
	receiver = feedbackemail.Fields.Item("f_address").Value

	hrtg = "<hr size='1' color='#330099' noshade>"

	strMessage = "<html></head>"
	strMessage = strMessage & "<BODY leftmargin='0' topmargin='0' marginwidth='0' marginheight='0' bgcolor='#FFFFFF' text='#000099'>"
	strMessage = strMessage & "<font face='Arial'>"
	strMessage = strMessage & "<b>" & UCase(Request.Form("title1")) & "</b><br><br>"
	strMessage = strMessage & "<b>Report from:</b> " & Session("firstname") & " " & Session("lastname") & "<br>"
	strMessage = strMessage & "<b>Report at:</b> " & now() & "<br><br>"

	strMessage = strMessage & "<br><b>Comments:</b> <br>"
	strMessage = strMessage & Server.HTMLEncode(Request.Form("details")) & "<br><br>"
	strMessage = strMessage & "</body></html>"


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

	Dim objConfig  ' As CDO.Configuration
	Dim objMessage ' As CDO.Message
	Dim Fields     ' As ADODB.Fields

	' Get a handle on the config object and it's fields
	Set objConfig = Server.CreateObject("CDO.Configuration")
	Set Fields = objConfig.Fields

	' Set config fields we care about
	With Fields
		.Item(cdoSendUsingMethod)       = cdoSendUsingPort
		.Item(cdoSMTPServer)            = "lotjfp01"
		.Item(cdoSMTPServerPort)        = 25
		.Item(cdoSMTPConnectionTimeout) = 10
		.Update
	End With

	Set objMessage = Server.CreateObject("CDO.Message")

	Set objMessage.Configuration = objConfig

	With objMessage
		.To       = receiver
		.CC = "rony@lotj.com"
		.From     = name & "<bbp_user@lotj.com>"
		.Subject  = client_name_short & " Feedback Report"
		.HTMLBody  = strMessage
		.Send
	End With

	Set Fields = Nothing
	Set objMessage = Nothing
	Set objConfig = Nothing
	
end if
	msg ="<br><br><b>Thank you for your feedback. Your feedback has been sent to the program support team. <br><br>You may close this window.</b>"

end if
%>

		<form name="feedback" method="post" action="help.asp" onsubmit="return checkform()" >
			<br>
			<%if msg="" then %>
			<br>

				<TEXTAREA class="counter" name="details" id="bugtext" style="width:350px; height:150px;"></TEXTAREA>
				<br>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;  Note: Max 500 characters allowed
				
				<br>
				<INPUT class="btn btn-default" type="submit" value="Submit" name=submit>
				<input type="hidden" NAME="save_type" VALUE="dbe">

			<%else%>
				<%=msg%>
				
			<%end if%>
		</form>

		</div>
	<div class="clear"></div>
</div>
</body>
</html>
<%
call log_the_page ("Help Page", "0", "n/a", "0", "n/a", "0", "n/a", comment)
%>

<!-- #include file = "errorhandler/index.asp"-->