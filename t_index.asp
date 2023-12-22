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

Dim selected_topic_query
if (cstr(Session("selected_topics")) <> "") Then
    selected_topic_query = cstr(Session("selected_topics"))
End If

Session("qsave") = "" 
 if Session("userID") = "" then response.redirect("error.asp?" & request.QueryString & "&message=NoUserIdFound")

 if NOT pref_quiz_avail then response.redirect("error.asp?" & request.QueryString & "&message=pref_quiz_avail not available")


' current position in the test
Dim position
position = 1

' PN 050515 fix random question generation. random definition of number of the first question in the set
Dim first_question
first_question=Int(rnd*(noqis))+1


Dim ID_subject_prm
ID_subject_prm = Session("id")
If ID_subject_prm = "" Then
    Response.Redirect("error.asp?" & request.QueryString)
End If

Dim userid
if (Session("UserID") <> "") Then
    userid = cInt(Session("UserID"))
Else
    Response.Redirect("error.asp?" & request.QueryString)
End If




set act_subj = Server.CreateObject("ADODB.Recordset")
if selected_topic_query <> "" then
	SQL = "SELECT subjects.subject_name,subjects.subject_passmark,  (SELECT count(s2.s_id) FROM new_subjects s2 WHERE s2.s_qID = "&fixstr(clng(ID_subject_prm))&" AND ( "&cstr(selected_topic_query)&" ) AND  ABS([s_active]) = 1 ) totAntal FROM subjects WHERE (subjects.ID_subject ="&fixstr(clng(ID_subject_prm))&") AND abs(subject_active_q) = 1;"
else
	SQL = "SELECT subjects.subject_name,subjects.subject_passmark,  (SELECT count(s2.s_id) FROM new_subjects s2 WHERE s2.s_qID = "&fixstr(clng(ID_subject_prm))&" AND  ABS([s_active]) = 1 ) totAntal FROM subjects WHERE (subjects.ID_subject ="&fixstr(clng(ID_subject_prm))&") AND abs(subject_active_q) = 1;"
end if

act_subj.Open SQL, Connect,3,3

if act_subj.EOF or act_subj.BOF then response.redirect("error.asp?" & request.QueryString) 
total_counter = act_subj("totAntal")

Dim sTopic
'sTopic=act_subj("sTop")


set startSubject = Server.CreateObject("ADODB.Recordset")
if selected_topic_query <> "" then
	SQL = "SELECT TOP 1 s_id, s_topic FROM new_subjects s2 WHERE  s2.s_qID = "&fixstr(clng(ID_subject_prm))&" AND ( "&cstr(selected_topic_query)&" ) AND ABS([s_active]) = 1 ORDER BY s_order"
else
	SQL = "SELECT TOP 1 s_id, s_topic FROM new_subjects WHERE  s_qID = "&fixstr(clng(ID_subject_prm))&" AND ABS([s_active]) = 1 ORDER BY s_order"
end if
startSubject.Open SQL, Connect,3,3
    sStart=startSubject(0)
	
startSubject.close

set user_check = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT session_current,q_session.ID_session, q_session.Session_users, q_session.Session_done, q_session.Session_stop, q_session.Session_date, q_session.Session_total  FROM q_session  WHERE (((q_session.Session_users)=" &fixstr(userID)& ") AND ((q_session.Session_done)=0) AND ((q_session.Session_subject)=" & fixstr(ID_subject_prm) & ") AND (Session_current IS NOT NULL)) ORDER BY q_session.Session_date DESC;"
'response.write SQL & "<br>"
user_check.Open SQL, Connect,3,3
Dim orderOn
If Not user_check.EOF Then
orderOn=user_check("Session_stop")
'Get the topics user is on
set topicOn = Server.CreateObject("ADODB.Recordset")

if selected_topic_query <> "" then
	SQL = "select distinct(s_topic) as dTopic from new_subjects s2 where s2.s_qId = "&fixstr(clng(ID_subject_prm))&" AND ( "&cstr(selected_topic_query)&" ) and s_active = 1 and s_order="&orderOn
else
	SQL = "select distinct(s_topic) as dTopic from new_subjects where s_qId = "&fixstr(clng(ID_subject_prm))&" and s_active = 1 and s_order="&orderOn
end if

topicOn.Open SQL, Connect, 3, 1 
Dim currentTopic
If Not topicOn.EOF Then
currentTopic=topicOn("dTopic")
end if
topicOn.Close
end if


'Get the total topics of the subject
set topicCount = Server.CreateObject("ADODB.Recordset")

if selected_topic_query <> "" then
	SQL = "select count(distinct(s_topic)) as cTopic from new_subjects s2 where s2.s_qId = "&fixstr(clng(ID_subject_prm))&" AND ( "&cstr(selected_topic_query)&" ) and s2.s_active = 1 "
else
	SQL = "select count(distinct(s_topic)) as cTopic from new_subjects where s_qId = "&fixstr(clng(ID_subject_prm))&" and s_active = 1 "
end if

topicCount.Open SQL, Connect, 3, 1 
Dim currentTopicCount

currentTopicCount=topicCount("cTopic")

topicCount.Close

set topicCount_total = Server.CreateObject("ADODB.Recordset")

SQL = "select count(distinct(s_topic)) as cTopic from new_subjects where s_qId = "&fixstr(clng(ID_subject_prm))&" and s_active = 1 "


topicCount_total.Open SQL, Connect, 3, 1 
Dim ctopicCount_total

ctopicCount_total=topicCount_total("cTopic")

topicCount_total.Close


'Get all topics and store in array
set topicsAll = Server.CreateObject("ADODB.Recordset")
if selected_topic_query <> "" then
	SQL = "SELECT s_topic as dTopic, MAX(s_order) as sorder FROM new_subjects s2 where s2.s_qid="&fixstr(clng(ID_subject_prm))&" AND ( "&cstr(selected_topic_query)&" ) and s2.s_active=1 GROUP BY s2.s_topic ORDER BY MAX(s2.s_order) Asc"
else
	SQL = "SELECT s_topic as dTopic, MAX(s_order) as sorder FROM new_subjects where s_qid="&fixstr(clng(ID_subject_prm))&" and s_active=1 GROUP BY s_topic ORDER BY MAX(s_order) Asc"
end if

topicsAll.Open SQL, Connect, 3, 1 
Dim myArray() 'Declaring a dynamic array
ReDim myArray(currentTopicCount) 'Re Declaring a dynamic array
Dim i
i=0
If Not topicsAll.EOF Then
    Do Until topicsAll.EOF
	myArray(i)=topicsAll("dTopic")
	i=i+1
		
    topicsAll.MoveNext
    Loop
	End If
topicsAll.Close

'find current order of the topic the user is on using the array
Dim counter
counter=1

for i=0 to UBound(myArray)
		
		if currentTopic=myArray(i) then
		topicPosition=counter
		Exit For
		else
		counter=counter+1
		end if
		
next

	



%>
<!DOCTYPE html>
<head>
  <meta
    name="viewport"
    content="width=device-width, initial-scale=1.0,  minimum-scale=.9"
  />

  <title><%=client_name_short%> - Better Business Program</title>
  <meta name="DESCRIPTION" content="" />
  <script src="jquery-1.11.1.js?v=bbp34"></script>
  <!-- <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-T3c6CoIi6uLrA9TneNEoa7RxnatzjcDSCmG1MXxSR1GAsXEV/Dwwykc2MPK8M2HN" crossorigin="anonymous"> -->
  <!-- #include file = "inc_header.asp" -->

  <link href="style/subject-welcome-page.css" rel="stylesheet" type="text/css">

  <!-- <link
    rel="stylesheet"
    href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.3.0/css/all.min.css"
  /> -->
  <link rel="stylesheet" type="text/css" href="js/sweet-alert.css" />
  <link
    href="perfect-scrollbar-0.4.8/src/perfect-scrollbar.css"
    rel="stylesheet"
  />


  <script src="js/modernizr-latest.js?v=bbp34"></script>
  <script src="js/sweet-alert.min.js?v=bbp34"></script>
  <script src="perfect-scrollbar-0.4.8/src/jquery.mousewheel.js?v=bbp34"></script>
  <script src="perfect-scrollbar-0.4.8/src/perfect-scrollbar.js?v=bbp34"></script>

  <script>
    // orientation code
    //adapt_to_orientation();
    $(document).ready(function () {
      if (window.innerHeight > window.innerWidth) {
        swal({
          title: "This site is best used in landscape view.",
          text: "",
          type: "warning",
          confirmButtonText: "OK",
        });
      }

      $("html, body").animate({ scrollTop: $(document).height() }, 10);
      /*var d = $("body");
		 var rotate = 90 - window.orientation;
		 d.css("transform", "rotate("+rotate+"deg)");
		window.addEventListener('orientationchange', function ()
		{
			//adapt_to_orientation();
   
				if(window.orientation > 0)
					rotate=0;
				else 
				rotate=90;
	
				d.css("transform", "rotate("+rotate+"deg)");
		});
	*/

      $(".inside_content").perfectScrollbar({ suppressScrollX: true });

      function adapt_to_orientation() {
        // For use within normal web clients
        var isiPad = navigator.userAgent.match(/iPad/i) != null;

        var content_width, screen_dimension;

        if (window.orientation == 0 || window.orientation == 180) {
          // portrait

          content_width = 900;
          screen_dimension = screen.width * 0.98; // fudge factor was necessary in my case
        } else if (window.orientation == 90 || window.orientation == -90) {
          // landscape
          content_width = 750;
          screen_dimension = screen.height;
        }

        var viewport_scale = screen_dimension / content_width;

        // resize viewport
        $("meta[name=viewport]").attr(
          "content",
          "width=" +
            content_width +
            "," +
            "initial-scale=" +
            viewport_scale +
            ", maximum-scale=" +
            viewport_scale
        );

        // resize viewport
        //$('meta[name=viewport]').attr('content','user-scalable=YES');
      }
    });
  </script>
</head>
<body>
  <div class="page-content">
    <div class="white-container">
      <!-- #include file = "partials/header.asp" -->
      <!-- #include file = "partials/subject-welcome-page-content.asp" -->
    </div>
      <!-- #include file = "partials/footer.asp" -->
  </div>
</body>
<% if Request.ServerVariables("HTTP_REFERER") <> "" then comment = Request.ServerVariables("HTTP_REFERER") else comment = "start quiz"
call log_the_page ("quiz start", "0", "n/a", "0", "n/a", "0", "n/a", comment)
%>

<!-- #include file = "errorhandler/index.asp"-->
