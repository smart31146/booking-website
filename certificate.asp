<%@LANGUAGE="VBSCRIPT"%>
<% 
'Response buffer is used to buffer the output page. That means if any database exception occurs the contents can be cleared without processed any script to browser
 Response.Buffer = True
 
' "On Error Resume Next" method allows page to move to the next script even if any error present on page whcich will be caught after processing all asp script on page
' On Error Resume Next
 
'Changed by PR on 25.02.16
 %>
<!-- #include file = "connections/bbg_conn.asp" -->
<!-- #include file = "connections/include.asp"-->

<%
if NOT pref_quiz_avail then response.redirect("error.asp?" & request.QueryString)


' lets create some general variables for certificate'
Dim Total_answered
Dim Total_correct
Total_answered = 0
Total_correct = 0
'CXS 26062007 - SumTotal specific'
Dim sumtotal_score
Dim sumtotal_totalquestions
Dim sumtotal_coursefolder
Dim sumtotal_returnURL


Dim SessionID
if (Session("SessionID") <> "") Then
	SessionID = CLng(Session("SessionID"))
Else
	Response.Redirect("error.asp?" & request.QueryString)
End If

Dim UserID
if (Session("UserID") <> "") Then
	UserID = CLng(Session("UserID"))
Else
	Response.Redirect("error.asp?" & request.QueryString)
End If

' ID of users session
Dim sID
if (Session("ID") <> "") Then
	sID = Session("ID")
Else
	Response.Redirect("error.asp?" & request.QueryString)
End If



set subject = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT *  FROM subjects WHERE ID_subject = "&fixstr(clng(sID))&" "
subject.Open SQL, Connect,3,3
subject_name = subject("subject_name")


if (Session("original_subject") <> "") then
    original_subject_id = Session("original_subject")
    set original_subject_details = Server.CreateObject("ADODB.Recordset")
    SQL = "SELECT *  FROM subjects WHERE ID_subject = "&fixstr(clng(original_subject_id))&" "
    original_subject_details.Open SQL, Connect,3,3
    subject_name = original_subject_details("subject_name")
    original_subject_details.close
end if


set results = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT (SELECT choice_cor FROM q_choice WHERE result_answer = ID_choice) qanswer, s_topic FROM q_result,q_question,new_subjects WHERE question_topic = s_id AND result_question = ID_question AND  result_session = " &fixstr(clng(SessionID))& " ORDER BY id_result"
'SQL = "SELECT (SELECT choice_cor FROM q_choice WHERE result_answer = ID_choice) qanswer,* FROM q_result,q_question,new_subjects WHERE question_topic = s_id AND result_question = ID_question AND  result_session = " &fixstr(clng(SessionID))& " ORDER BY id_result"
'response.write SQL & "<br>"
results.Open SQL, Connect,3,3




set userdetails = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT q_user.ID_user, q_user.user_lastname, q_user.user_firstname FROM q_user WHERE q_user.ID_user=" &fixstr(clng(Session("userID")))& ""
userdetails.Open SQL, Connect,3,3


Function certdate( _
 byVal whatdate _
 )

 dim monthid
 dim monthstr
 monthid = month(whatdate)
 select case monthid
 	case 1
		monthstr="January"
 	case 2
		monthstr="February"
 	case 3
		monthstr="March"
 	case 4
		monthstr="April"
 	case 5
		monthstr="May"
 	case 6
		monthstr="June"
 	case 7
		monthstr="July"
 	case 8
		monthstr="August"
 	case 9
		monthstr="September"
 	case 10
		monthstr="October"
 	case 11
		monthstr="November"
 	case 12
		monthstr="December"
 end select

 dim weekdayid
 dim weekdaystr
 weekdayid = weekday(whatdate)
 select case weekdayid
 	case 1
		weekdaystr="Sunday"
 	case 2
		weekdaystr="Monday"
 	case 3
		weekdaystr="Tuesday"
 	case 4
		weekdaystr="Wednesday"
 	case 5
		weekdaystr="Thursday"
 	case 6
		weekdaystr="Friday"
 	case 7
		weekdaystr="Saturday"
 end select

 dim dayid
 dayid = day(whatdate)

 dim yearid
 yearid = year(whatdate)

 dim ordstr
 select case dayid
 	case 1, 21, 31
		ordstr="st"
 	case 2, 22
		ordstr="nd"
 	case 3, 23
		ordstr="rd"
	case else
		ordstr="th"
 end select

certdate=weekdaystr & ", " & dayid & ordstr & " of " & monthstr & " " & yearid
End Function
%>


<!doctype html>
<head>
	
	<title><%=client_name_short%> - Better Business Program</title>
	<META name="DESCRIPTION" content="">
	<meta id="viewport" name='viewport'>
	<meta name="viewport" content="width=device-width">
        
	<link rel="stylesheet" href="js/960.css" type="text/css" media="screen">
	<link rel="stylesheet" href="js/screen.css" type="text/css" media="screen" />
	<link rel="stylesheet" href="js/print.css" type="text/css" media="print" />
	<link rel="stylesheet" href="js/src/css/print-preview.css" type="text/css" media="screen">
	<script src="https://cdnjs.cloudflare.com/ajax/libs/jquery-tools/1.2.7/jquery.tools.min.js?v=bbp34"></script>
	<script src="js/src/jquery.print-preview.js?v=bbp34" type="text/javascript" charset="utf-8"></script>

	<!-- <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-T3c6CoIi6uLrA9TneNEoa7RxnatzjcDSCmG1MXxSR1GAsXEV/Dwwykc2MPK8M2HN" crossorigin="anonymous"> -->
	<!-- #include file = "inc_header.asp" -->
    <link href="style/certificate-page-styles.css" rel="stylesheet" type="text/css">
	<link href="style/media/certificate-print.media.css" rel="stylesheet" type="text/css">
	<!-- <link
	rel="stylesheet"
	href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.3.0/css/all.min.css"
  /> -->
		
	<link rel="stylesheet" href="style/normalize.min.css">
	<!--[if lt IE 9]>
	<script src="js/html5-3.6-respond-1.1.0.min.js?v=bbp34"></script>
	<![endif]-->
<script>
    (function(doc) {
        var viewport = document.getElementById('viewport');
        if ( navigator.userAgent.match(/iPhone/i) || navigator.userAgent.match(/iPod/i)) {
            viewport.setAttribute("content", "initial-scale=0.3");
        } else if ( navigator.userAgent.match(/iPad/i) ) {
            viewport.setAttribute("content", "initial-scale=1.05");
        }
    }(document));
</script>

	<link href="perfect-scrollbar-0.4.8/src/perfect-scrollbar.css" rel="stylesheet">
    <script src="perfect-scrollbar-0.4.8/src/jquery.mousewheel.js?v=bbp34"></script>
	<script src="perfect-scrollbar-0.4.8/src/perfect-scrollbar.js?v=bbp34"></script>
	<link rel="stylesheet" type="text/css" href="js/sweet-alert.css">
	<script src="js/sweet-alert.min.js?v=bbp34"></script>
	<script src="js/jQuery.print.js?v=bbp34"></script>  

	<script type='text/javascript'>
		$(function() {
		var isiDevice = /ipad|iphone|ipod/i.test(navigator.userAgent.toLowerCase());
		if (isiDevice) {
			/*
			* Initialise example carousel
			*/
			$("#feature > div").scrollable({interval: 2000}).autoscroll();
		   
		   /*
			* Initialise print preview plugin
			*/
		   // Add link for print preview and intialise
			  // $('a.print-preview').printPreview();
			 $('a.print-link').printPreview();
		   
		   // Add keybinding (not recommended for production use)
		}
		else {
		  /* $("#ele2").find('.print-link').on('click', function() {
		   //Print ele2 with default options
		   $.print("#ele2");
		   });*/
		   // $( ".header_inside_image" ).hide();
		   $("#ele1").find('.print-link').on('click', function() {
			   //Print ele4 with custom options
			    $("#ele1").print({
				   	//Use Global styles
					// globalStyles : false,
					//Add link with attrbute media=print
					mediaPrint : false,
					//Custom stylesheet
					stylesheet : "//fonts.googleapis.com/css?family=Inconsolata",
					//Print in a hidden iframe
					iframe : true,
					//Don't print this
					noPrintSelector : ".no-add-to-print-mode",
					//Add this at top
					// prepend : "<div class='print-heading'><img class='print-heading__logo' src='images/logo_certficate.jpg' /><h4>BETTER BUSINESS PROGRAM CERTIFICATE</h4><h3> <% =subject("subject_name")%></h3></div>",
					//Add this on bottom
					append : "<br/>!",
					timeout: 2000
			    });
		   });
		   // Fork https://github.com/sathvikp/jQuery.print for the full list of options
		}

		  
	   });
	</script>
</head>
<body>
	<div class="page-content">
        <div class="white-container">
			<% if Session("LMS") = 1 then %>
			<!-- #include file = "main_top_certification.asp" -->
			<% else %>
			<!-- #include file = "partials/header.asp" -->
			<%end if%>

			<!-- #include file = "partials/sub-header-certificate.asp" -->
			<!-- #include file = "partials/certificate-page-content.asp" -->

	
	
			<!-- <div class="menu_content" >
			
				<% if(pass_or_fail=1) then %>
					<img src="vault_image/cert_pass.jpg" width="320" height="390" alt="">
				<% else %>
				
					<img  src="vault_image/images/handstacked_1.jpg" width="320" height="390" alt="">
					
				<% END IF%></div> -->
		
            </div>
		<!-- #include file = "partials/footer.asp" -->
	</div>

</body>
<%
'call log_the_page ("quiz", ID_subject_prm, ReplaceStrQuiz(subject.Fields.Item("subject_name").Value), "0", "n/a", "0", "Certificate", "Quiz certificate page")
%>
<%
	call log_the_page("Training and quiz", Session("ID"), "Certificate", 0, "Certificate", 0, qst, "Certificate")

subject.Close()
results.Close()
userdetails.Close()


if Session("LMS") = 1 then
	if(pass_or_fail=1) then %>


<script type="text/javascript">
var ct=document.getElementById("countdown");
var currentsecond=11;
var sd=document.getElementById("second");
function countredirect(){
if (currentsecond!=1){

currentsecond-=1
if(currentsecond==1)
sd.innerHTML="second"

ct.innerHTML=currentsecond
}


setTimeout("countredirect()",1000)
}

countredirect();
		window.setTimeout("document.location = 'lmscallback.asp?SessionBBP=<% =SessionID%>';", 20000);
		
		
	</script>
	
	
<%	end if 
end if

%>
<!-- #include file = "errorhandler/index.asp"-->