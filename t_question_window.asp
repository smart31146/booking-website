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
if NOT pref_quiz_avail then response.redirect("error.asp?" & request.QueryString)




' ID of users session
Dim sID
if (Session("ID") <> "") Then
	sID = Session("ID")
Else
	Response.Redirect("error.asp?" & request.QueryString)
End If



if Err.Number = 0 then
set subject = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM new_subjects WHERE s_id = "&fixstr(clng(request.querystring("s_id")))&""
subject.Open SQL, Connect,3,3
end if
if subject.EOF or subject.BOF then response.redirect("error.asp?" & request.QueryString)


		
		
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<title><%=client_name_short%> - Better Business Program</title>
		<META name="DESCRIPTION"	content="">
		<!-- #include file = "inc_header.asp" -->
		 <link href="perfect-scrollbar-0.4.8/src/perfect-scrollbar.css" rel="stylesheet">
       <script src="jquery-1.11.1.js?v=bbp34"></script>
      <script src="perfect-scrollbar-0.4.8/src/jquery.mousewheel.js?v=bbp34"></script>
      <script src="perfect-scrollbar-0.4.8/src/perfect-scrollbar.js?v=bbp34"></script>
	 <link rel="stylesheet" type="text/css" href="js/sweet-alert.css">
    		  <script src="js/sweet-alert.min.js?v=bbp34"></script>
			  <script>
			  $(document).ready(function() {
			  $(".inside_content").perfectScrollbar({suppressScrollX: true});
			  
			  });
			  </script>
</head>
<body>
		
		<div class="main_content">
		<div class="guide_blue" style="text-align:left;">
			<div class="box_inside"><h3><%=ReplaceStrQuiz(subject("s_title"))%></h3>
		
			
			
			<div class="box_text_blue">
			<div class="inside_content" style="height:350px;" > 
	        <p><%=ReplaceStrQuiz(subject("s_body"))%></p>
			</div>
			<div class="clear"></div>
				
			  <br><br><br>
			
	 
	<div class="clear"></div>
</div>
</div>
</body>
</html>
<% subject.close : Set subject = Nothing%>
<!-- #include file = "errorhandler/index.asp"-->