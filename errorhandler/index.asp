

<%
' This page is called when any error regarding to database found in program. Implemented for protection
' against SQLi injection. Developerd by PR 24.02.2016	

' Error Handler (Check if there is error in page)
 	If Err.Number <> 0 Then
 
' Clear response buffer (Cleares contents of page)
 		Response.Clear
%><!doctype html>

<head>
 <meta id="viewport" name='viewport' >
<!-- #include file = "inc_header.asp" -->
</head>
<body >


<img src="images/bg.jpg" alt="background image" id="bg">
<!-- #include file = "main_top.asp" -->
<div style="float: left; width: 100%;  background: rgb(1, 79, 147);  border-radius: 9px;" id="dd" >

	<div style="width:600px;height:250px;position:relative;" >
		<div style="position:absolute;left:35px;top:15px;color:#fff;width:870px;margin-top:20px;"><div align="center"><img src="images/logos.png"/></div>
			<div style="text-align:center;color:#fff;font-size:20px;margin-top:20px;"><strong>Page not found</strong>
</div>  
  </div> <!-- end message -->
		</div>
		<div style="padding:20px;text-align:center;color:#fff;font-size:16px;"> 
    There was an error in processing your request. Please click <a href="<%= homeURL&"?alt=change"%>" style="color:#fff;">here</a> to return to the homepage. If the problem persists, please contact your program administrator<!-- enter your details here -->
   <br>
   <br>
   <a href="<%= homeURL&"?alt=change"%>" style="color:#fff;">Home</a> 
</div> 
		</div>
</div>
<!-- #include file = "inc_bottom.asp" -->
</body>
</html>
<% end if %>