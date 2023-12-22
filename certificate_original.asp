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
	SessionID = cInt(Session("SessionID"))
Else
	Response.Redirect("error.asp?" & request.QueryString)
End If

Dim UserID
if (Session("UserID") <> "") Then
	UserID = cInt(Session("UserID"))
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
		<META name="DESCRIPTION"	content="">
		<meta id="viewport" name='viewport'>
				<script src="jquery-1.11.1.js?v=bbp34"></script>
				<meta name="viewport" content="width=device-width">
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
		<!-- #include file = "inc_header.asp" -->
		<link href="perfect-scrollbar-0.4.8/src/perfect-scrollbar.css" rel="stylesheet">
       <script src="jquery-1.11.1.js?v=bbp34"></script>
      <script src="perfect-scrollbar-0.4.8/src/jquery.mousewheel.js?v=bbp34"></script>
      <script src="perfect-scrollbar-0.4.8/src/perfect-scrollbar.js?v=bbp34"></script>
	 <link rel="stylesheet" type="text/css" href="js/sweet-alert.css">
    		  <script src="js/sweet-alert.min.js?v=bbp34"></script>
			   <script src="js/jQuery.print.js?v=bbp34"></script>
        <script type='text/javascript'>
                                    //<![CDATA[
                                    $(function() {
                                     /* $("#ele2").find('.print-link').on('click', function() {
                                            //Print ele2 with default options
                                            $.print("#ele2");
                                        });*/
                                        $("#ele1").find('.print-link').on('click', function() {
                                            //Print ele4 with custom options
                                            $("#ele1").print({
                                                //Use Global styles
                                                globalStyles : false,
                                                //Add link with attrbute media=print
                                                mediaPrint : false,
                                                //Custom stylesheet
                                                stylesheet : "//fonts.googleapis.com/css?family=Inconsolata",
                                                //Print in a hidden iframe
                                                iframe : true,
                                                //Don't print this
                                                noPrintSelector : ".avoid-this",
                                                //Add this at top
                                                prepend : "<div class='header_inside' style='text-align:center'><img  style='width:100px' src='images/logo_certificate.png' /><h4>BETTER BUSINESS PROGRAM CERTIFICATE</h4> <h3> <% =subject("subject_name")%></h3></div>",
                                                //Add this on bottom
                                                append : "<br/>Buh Bye!",
												timeout: 2000
                                            });
                                        });
                                        // Fork https://github.com/sathvikp/jQuery.print for the full list of options
                                    });
                                    //]]>
        </script>
</head>
<body>

<img src="images/bg.jpg" alt="background image" id="bg">
<!-- #include file = "main_top.asp" -->

	<div class="header_blue">
		
		
		<div class="header_inside">CERTIFICATE<br>
		<h3> <% =subject("subject_name")%></h3>
		</div>
	</div>
	
	<div class="clear"></div>
	
	<div class="main_content2">
		<div class="guide_blue">
			<div class="box_inside2" >
			 <div id="ele1" class="a">
			<h5 ><span  style="color:#0354AC">Date Completed: </span><span><%= FormatDateTime(Date, 1) %> at <%= FormatDateTime(Now, 3) %></span></h5>
			<h2><% =UCASE(userdetails("user_firstname"))%>&nbsp;<% =UCASE(userdetails("user_lastname"))%></h2>		
			 <table style="width:90%">
              <tr >
                <td ><i>&nbsp;</i></td>
                <td ><i>Quiz</i></td>
                <td ><i>Answer</i></td>
              </tr>
              <%
			  xi = 0
While (NOT results.EOF)
	count_of_correct = 0
	xi = xi + 1
	IF (results("qanswer")) = True THEN
		image = "<img src=""images/icon_true.gif"" alt="""">"
		Total_correct = Total_correct+1
	ELSE
		image = "<img src=""images/icon_false.png"" alt="""">"
	END IF
	%>
	<tr>
		<td  ><strong>Quiz <% =xi%></strong></td>
		<td ><% =results("s_topic")%></td>
		<td ><%=image%> </td>
	</tr>
	<%
  results.MoveNext()
Wend
%>
<%
 MM_editConnection = Connect
  MM_editTable = "q_session"
  MM_editQuery = "update " & MM_editTable & " set session_done = 1, session_correct = " & total_correct & ", session_finish = '" & cDateSql(Now())&"' " & "where ID_session = " & SessionID
  'Response.Write MM_editQuery
 	Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

	'pn 060127 work out percentage achieved and whether it was a pass or fail
	Dim percentage_achieved
	Dim subject_expiry
	Dim subject_passmark
	Dim pass_or_fail
	Dim add_to_expiry_date
	pass_or_fail=0
	add_to_expiry_date=0
	subject_expiry=subject("subject_expiry")
	subject_passmark=subject("subject_passmark")

	percentage_achieved=(total_correct/xi)*100

	if(percentage_achieved>=subject_passmark) then
			pass_or_fail=1
			add_to_expiry_date=subject_expiry
	end if

	IF Session("qsave") = "" THEN
		'pn 060127 add a certification for this person, can be 0 for failed or 1 for passed
		Set MM_SaveCertification = Server.CreateObject("ADODB.Command")
	    MM_SaveCertification.ActiveConnection = MM_editConnection
	    MM_SaveCertification.CommandText   = "insert into q_certification (q_session, quiz_date, expiry_date, passed, percentage_achieved, percentage_required) values ('" & SessionID & "' ,'" & cDateSql(Now())&"',DATEADD (week , "&add_to_expiry_date&", GETDATE()),"&pass_or_fail&","&percentage_achieved&","&subject_passmark&" );"
	    MM_SaveCertification.Execute
	    MM_SaveCertification.ActiveConnection.Close
		
	Session("qsave") = "yesbox"
	END IF
%>
              <tr>
                <td >&nbsp;</td>
                <td >&nbsp;</td>
                <td >&nbsp;</td>
              </tr>
              <tr>
                <td ><b>Total</b></td>
                <td >&nbsp;</td>
                <td ><b><%=total_correct%>/<%=xi%></b></td>
              </tr>
            </table>
            <br>The passmark for this subject is <%=subject_passmark%>%. <br><br>You achieved  <%=FormatNumber(percentage_achieved,2)%>% for this quiz which is a
			<% if(pass_or_fail=1) then %>
				<span style="color:green"><strong>PASS.</strong></span><br><br>
				<strong>Congratulations! You can print this certificate using the link below. To finish – you can either close the browser or click on the link below to return to the homepage.
</strong><br><br>
				 <% if Session("LMS") = 1 then %>
				<br><br>Click button below to save your results and return to the LMS.
				<% END IF%>
			<% else %>
				<span style="color:red"><strong>FAIL.</strong></span><br><br>
				<strong>You will need to redo this subject to become certified. Please click on the link below to return to the homepage to do the subject again.
</strong><br><br>
				
				<% if Session("LMS") = 1 then %>
				<br><br> <strong>You are required to complete this subject again.</strong> <br><br>Click button below to return to the LMS where you can start this subject again.
				<% END IF%>
			<% end if%>
			
         
        <% if Session("LMS") <> 1 then %>
		<div>
		  
			<% if(pass_or_fail=1) then %>
			<% END IF%>
			<%=client_name_long%> is committed to promoting practical knowledge of the law.<br><br>
             Thank you for participating in the Better Business Program.<br>
			
            </div>
	<%else%>
	<div >
		Otherwise, you will be returned to the LMS automatically in <span id="countdown" style="font-size:16px;color:red;">30</span> <span id="second">seconds</span>.
		<br>
		<br>
		
                          <a style="padding-left: 15px;" href="lmscallback.asp?SessionBBP=<% =SessionID%>" >  <div class="return_blue">Return to LMS</div></a>
                       
                    
		</div>
	<%end if%>
	
	<!-- added condition to remove print this page and home page buttons from LMS certificate page. by PR on 24.02.2016 -->
	
	<% if Session("LMS") <> 1 then %>
		
		    <div class="just_blue">
		<a href="javascript:void(0)" style="color:#fff;padding:0px;"  class="print-link avoid-this">
                Print this page
                </a>
				</div>
   &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<div class="just_blue avoid-this"><a href="<%= homeURL&"?alt=change"%>" style="color:#fff;">HomePage</a></div>
   </div>
   <% END IF%>
		
        <div>
		  <% if Session("LMS") = 1 then %>
			<% if(pass_or_fail=1) then %>
</strong>


			<% END IF%>
			<% END IF%>
           
			<br>
			
		</div>
			 
			</div>
			</div>
		</div>
	
	<div class="menu_content" >
	
		<% if(pass_or_fail=1) then %>
			<img src="vault_image/cert_pass.jpg" width="320" height="390" alt="">
		<% else %>
		
			<img  src="vault_image/images/handstacked_1.jpg" width="320" height="390" alt="">
			
		<% END IF%></div>
		
               

	<div class="clear"></div>
</div>
</div>
 
       
<!-- #include file = "inc_bottom.asp" -->
</html>
<%
'call log_the_page ("quiz", ID_subject_prm, ReplaceStrQuiz(subject.Fields.Item("subject_name").Value), "0", "n/a", "0", "Certificate", "Quiz certificate page")
%>
<%
	call log_the_page("Training and quiz", Session("ID"), "Certificate", 0, "Certificate", 0, qst, "Certificate")

subject.Close()
results.Close()
userdetails.Close()


if Session("LMS") = 1 then%>

<script type="text/javascript">
var ct=document.getElementById("countdown");
var currentsecond=31;
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
		window.setTimeout("document.location = 'lmscallback.asp?SessionBBP=<% =SessionID%>';", 30000);
		
		
	</script>
	
	
<%end if

%>
<!-- #include file = "errorhandler/index.asp"-->