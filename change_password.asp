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
<!--#include file="sha256.asp"-->
<%
Session("qsave") = "" 
if Session("userID") = "" then response.redirect("error.asp?" & request.QueryString)

if NOT pref_quiz_avail then response.redirect("error.asp?" & request.QueryString)

old_password = CStr(Request.form("bbp_old_password"))
new_password = CStr(Request.form("bbp_new_password"))
userid = cInt(Session("UserID"))

if request.querystring("alt")="change_pass" AND len(request.form("bbp_old_password"))>1 AND len(request.form("bbp_new_password"))>1 THEN
if Err.Number = 0 then
Set obj = Server.CreateObject("ADODB.Recordset")
	SQL="SELECT user_email FROM q_user WHERE ID_user = '" & userid & "'"
	obj.ActiveConnection = Connect
	obj.Source = SQL
	obj.CursorType = 0
	obj.CursorLocation = 3
	obj.LockType = 3
	obj.Open
end if
	if obj.EOF then
	response.redirect "index.asp?error=login"
	end if
	
	Dim salt
	salt = obj("user_email")
	
	old_password=old_password&salt
	old_password=sha256(old_password)

	'set MM_rsUser = Server.CreateObject("ADODB.Recordset")
		'MM_rsUser.ActiveConnection = Connect
		'MM_rsUser.Source = "SELECT TOP 1 * from q_user WHERE ID_user = '" & userid & "' AND  user_city = '" & old_password & "'"
		'MM_rsUser.CursorType = 0
		'MM_rsUser.CursorLocation = 3
		'MM_rsUser.LockType = 3
		'MM_rsUser.Open
		
		
		 SQL= "SELECT TOP 1 * from q_user WHERE (ID_user) =? AND  (user_city) = ?"
		 
' Err.Number is a attribute of "On Error Resume Next" method
' It is used to terminate any database query or transaction to provide protection against data integrity
' Changed by PR 23.02.16

if Err.Number = 0 then
set objCommand = Server.CreateObject("ADODB.Command") 
objCommand.ActiveConnection = Connect
objCommand.CommandText = SQL
objCommand.Parameters(0).value = userid
objCommand.Parameters(1).value = old_password
Set MM_rsUser = objCommand.Execute()
end if
		
	
	
	
	new_password=new_password&salt
	new_password=sha256(new_password)
	
if Err.Number = 0 then
	set MM_change = Server.CreateObject("ADODB.Command")
		SQL = "UPDATE q_user SET user_city=? WHERE ID_user='" & userid & "' AND  user_city =?"	
MM_change.ActiveConnection = Connect
MM_change.CommandText = SQL
MM_change.Parameters(0).value = new_password
MM_change.Parameters(1).value = old_password	
MM_change.Execute()		
end if
		'MM_change.Open SQL, Connect,3,3
		
		If ( MM_rsUser.EOF Or MM_rsUser.BOF) Then
			response.redirect "change_password.asp?error=change_password"
		else
			response.redirect "change_password.asp?ok=change_password"
		End If
end if
%>

<!doctype html>
<head>
	
	<title><%=client_name_short%> - Better Business Program</title>
		<META name="DESCRIPTION"	content="">
		
		<script src="jquery-1.11.1.js?v=bbp34"></script>
		<!-- <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-T3c6CoIi6uLrA9TneNEoa7RxnatzjcDSCmG1MXxSR1GAsXEV/Dwwykc2MPK8M2HN" crossorigin="anonymous"> -->
		<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-C6RzsynM9kWDrMNeT87bh95OGNyZPhcTNXj1NW7RuBCsyN/o0jlpcV8Qyq46cDfL" crossorigin="anonymous"></script>
		<!-- <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js?v=bbp34"></script>
		<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/css/bootstrap.min.css" /> -->
		<link rel="stylesheet" type="text/css" href="js/sweet-alert.css">
    		  <script src="js/sweet-alert.min.js?v=bbp34"></script>
		<!-- #include file = "inc_header.asp" -->
		<script >
		
			$(document).ready(function() {
		$("#bbp_old_password").focus();
		 
//alert($("*:focus").attr("id"));
setTimeout(function() {

$("#bbp_old_password").blur();

      // Do something after 2 seconds

}, 100);


 
  
	
});
function checkLogin() {
	var pass2 = document.getElementById('bbp_conf_password').value;
	var pass1 = document.getElementById('bbp_new_password').value;

	if (document.getElementById("bbp_old_password").value == '' )
	 {
		 
		swal({   title: "Old password is missing.",   text: "",   type: "error",   confirmButtonText: "OK" });
		return false;
	 }
	 else if (document.getElementById("bbp_new_password").value == '' )
	 {
	
		swal({   title: "New password is missing.",   text: "",   type: "error",   confirmButtonText: "OK" });
		return false;
	 }
	else if(pass1 != pass2) {
		
		swal({   title: "The new and confirmed passwords do not match!",   text: "",   type: "error",   confirmButtonText: "OK" });
		return false;
	}
	else {
	    return true;
	}
}
		</script>
		
		<!-- Masking a textbox with a date field 
		<script src="scriptlibrary/jquery.min.js?v=bbp34" type="text/javascript"></script>
		<script src="scriptlibrary/jquery.maskedinput-1.3.min.js?v=bbp34" type="text/javascript"></script>
		
		<script>
			jQuery(function($){
			   $("#bbp_password").mask("99/99/9999");
			});
		</script>-->
		
</head>
<body>
	<div class="page-content">
		<div class="white-container">
		  <!-- #include file = "partials/header.asp" -->
		  <div class="allcontent">
			<div class="allcontent_main">
<div class="main_content">
	<div style="background: url('images/start_main.jpg') no-repeat;width:600px;height:300px;position:relative;">
		<div style="position:absolute;left:205px;top:15px;color:#FFF;width:370px;"><strong>Change Your Password</strong>

	<form method="post" action="change_password.asp?alt=change_pass" onsubmit="return(checkLogin());">
		<br><br>
		Old Password:
		<input type="password" class="form-control" name="bbp_old_password" style="margin-bottom:10px;" id="bbp_old_password" >
		
		New Password:
		<input type="password" style="margin-bottom:10px;" class="form-control" name="bbp_new_password" id="bbp_new_password">
		
		Confirm New Password:
		<input type="password" style="margin-bottom:10px;" class="form-control" name="bbp_conf_password" id="bbp_conf_password">
		
		<input type="Submit" class="btn btn-default btn-light" style="width:180px;" value="Change Password">
		<% IF request.querystring("error")="change_password" THEN  %>
		<script>swal({   title: "The old password is incorrect",   text: "",   type: "error",   confirmButtonText: "OK", html:true });  $("#bbp_old_password").focus(); </script>
		<% end if%>
		<% IF request.querystring("ok")="change_password" THEN %>
		
		<script>swal({   title: "Your password has been changed successfully!",   text: "",   type: "success",   confirmButtonText: "OK", html:true },function(){window.location.href = 'index.asp';});  $("#bbp_old_password").focus(); </script>	
		
		<%set update_pass = Server.CreateObject("ADODB.Recordset")
			SQL = "UPDATE q_user SET user_password=0 WHERE id_user='"& fixstr(Session("UserID")) &"';"
			update_pass.Open SQL, Connect,3,3
		end if
		%>
	</form>

		</div>
	</div>
	
	<div class="box_blue">
		<div class="box_inside2">
			<br><br>
			<br><br>
		</div>
	</div>
	<div class="clear"></div>

</div>

<div class="menu_content" style="margin-left:20px;">
	<img src="images/start_right.jpg" width="300" height="300" alt=""><br>
	<div class="box_grey"><div class="box_inside2">
			<br><br>
			<br><br>
		</div>
		</div>
	</div>
	<div class="clear"></div>
</div>

<div class="clear"></div>

</div>


</div>
<!-- #include file = "partials/footer.asp" -->
</div>
</body>
<%
call log_the_page ("Change password", "0", "n/a", "0", "n/a", "0", "n/a", "Change password")
%>
<!-- #include file = "errorhandler/index.asp"-->