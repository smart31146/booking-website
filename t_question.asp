<%@LANGUAGE="VBSCRIPT"%>

<% 
'Response buffer is used to buffer the output page. That means if any database exception occurs the contents can be cleared without processed any script to browser
' Response.Buffer = True
 
' "On Error Resume Next" method allows page to move to the next script even if any error present on page whcich will be caught after processing all asp script on page
' On Error Resume Next
 
'Changed by PR on 25.02.16
%>
 
<% bbp_training = true%>
<!-- #include file = "connections/bbg_conn.asp" -->
<!-- #include file = "connections/include.asp"-->

<%
if NOT pref_quiz_avail then response.redirect("error.asp?" & request.QueryString & "&line=18")
'Used for bookmarking to check if the question has been answered or not
Session("answered")=0

Dim selected_topic_query
if (cstr(Session("selected_topics")) <> "") Then
    selected_topic_query = cstr(Session("selected_topics"))
End If

Dim selected_topic_query_last_query
selected_topic_query_last_query = "false"
if Request.querystring("nextID") = cstr("last") then
	selected_topic_query_last_query = "true"
end if

If  Request.querystring("quiz") <> "" Then

    
    response.write("question_ID")
    response.write("sessionID")
    response.write("question_time")

    ' Check if SessionID existx
        if (Session("sessionID") = "") OR (Session("question_ID") = "") Then Response.Redirect("error.asp?" & request.QueryString  & "&line=40")
        Session("answer") = Request.Form("answer")
        Session("currentID") = Request.Querystring("currentID")
        Dim question_time
        if (Session("question_time") <> "") Then		
            question_time = cDate(Session("question_time"))
        Else
            question_time = cDate(Now())
        End If
        question_time = abs(DateDiff("s",question_time, Now()))
		
' Err.Number is a attribute of "On Error Resume Next" method
' It is used to terminate any database query or transaction to provide protection against data integrity
' Changed by PR 23.02.16
if Err.Number = 0 then
        set INSERT = Server.CreateObject("ADODB.Recordset")
        SQL = "INSERT INTO q_result (result_session, result_question, result_answer, result_time) VALUES ("
        SQL = SQL & " "&fixstr(clng(Session("sessionID")))&","
        SQL = SQL & " "&fixstr(clng(Session("question_ID")))&","
        SQL = SQL & " "&fixstr(clng(Request.Form("answer")))&","
        SQL = SQL & " "&fixstr(question_time)&""
        SQL = SQL & " )"
        INSERT.Open SQL, Connect,3,3
end if
		'Set question to answered for bookmarking
		Session("answered")=1
		
        response.redirect "t_feedback.asp?" & request.querystring
        
        response.write cLng(Session("sessionID"))
END IF


' ID of users session
Dim sID
if (Session("ID") <> "") Then
    sID = Session("ID")
Else
    Response.Redirect("error.asp?" & request.QueryString  & "&line=78")
End If

if request.querystring("nextID")<>"" THEN
	if Request.querystring("nextID") = cstr("last") then
		set subject_selected_topic_query_last_query = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT TOP 1 *,(SELECT TOP 1 s2.s_order FROM new_subjects s2 WHERE s_typ = 2 AND s2.s_order > s1.s_order  AND ABS([s_active]) = 1 AND s2.s_qID = "&fixstr(clng(sID))&" ORDER BY s2.s_order ASC) NextQ,(SELECT TOP 1 s2.s_order FROM new_subjects s2 WHERE s_typ = 2 AND s2.s_order < s1.s_order  AND ABS([s_active]) = 1  AND s2.s_qID = "&fixstr(clng(sID))&" ORDER BY s2.s_order ASC) LastQ FROM new_subjects s1,subjects WHERE s1.s_qID = "&fixstr(clng(sID))&" AND s1.s_qiD = ID_subject  AND ABS([s_active]) = 1  ORDER BY s1.s_order DESC"
		subject_selected_topic_query_last_query.Open SQL, Connect,3,3
		currentID = cint(subject_selected_topic_query_last_query("s_id"))
		subject_selected_topic_query_last_query.close 
	else
		currentID = request.querystring("nextID")
	end if
END IF

showQuestion = False


set subject = Server.CreateObject("ADODB.Recordset")
if selected_topic_query <> "" then

	if selected_topic_query_last_query = "true" then
		SQL = "SELECT *,(SELECT TOP 1 s2.s_order FROM new_subjects s2 WHERE s_typ = 2 AND s2.s_order > s1.s_order  AND ABS([s_active]) = 1 AND s2.s_qID = "&fixstr(clng(sID))&"  ORDER BY s2.s_order ASC) NextQ,(SELECT TOP 1 s2.s_order FROM new_subjects s2 WHERE s_typ = 2 AND s2.s_order < s1.s_order  AND ABS([s_active]) = 1  AND s2.s_qID = "&fixstr(clng(sID))&" ORDER BY s2.s_order ASC) LastQ FROM new_subjects s1,subjects WHERE s1.s_id = "&fixstr(clng(currentID))&" AND s1.s_qiD = ID_subject AND ABS([s_active]) = 1 ORDER BY s1.s_order ASC"
		subject.Open SQL, Connect,3,3
	else

		if currentID="" THEN
			SQL = "SELECT *,(SELECT TOP 1 s2.s_order FROM new_subjects s2 WHERE s_typ = 2 AND ( "&cstr(selected_topic_query)&" ) AND s2.s_order > s1.s_order  AND ABS([s_active]) = 1 AND s2.s_qID = "&fixstr(clng(sID))&" ORDER BY s2.s_order ASC) NextQ,(SELECT TOP 1 s2.s_order FROM new_subjects s2 WHERE s_typ = 2 AND s2.s_order < s1.s_order  AND ABS([s_active]) = 1  AND s2.s_qID = "&fixstr(clng(sID))&" ORDER BY s2.s_order ASC) LastQ FROM new_subjects s1,subjects WHERE s1.s_qID = "&fixstr(clng(sID))&" AND s1.s_qiD = ID_subject AND ( "&replace(cstr(selected_topic_query),"s2","s1")&" ) AND ABS([s_active]) = 1  ORDER BY s1.s_order"
			subject.Open SQL, Connect,3,3
		ELSE
			SQL = "SELECT *,(SELECT TOP 1 s2.s_order FROM new_subjects s2 WHERE s_typ = 2 AND ( "&cstr(selected_topic_query)&" ) AND s2.s_order > s1.s_order  AND ABS([s_active]) = 1 AND s2.s_qID = "&fixstr(clng(sID))&"  ORDER BY s2.s_order ASC) NextQ,(SELECT TOP 1 s2.s_order FROM new_subjects s2 WHERE s_typ = 2 AND s2.s_order < s1.s_order  AND ABS([s_active]) = 1  AND s2.s_qID = "&fixstr(clng(sID))&" ORDER BY s2.s_order ASC) LastQ FROM new_subjects s1,subjects WHERE s1.s_id = "&fixstr(clng(currentID))&" AND s1.s_qiD = ID_subject AND ( "&replace(cstr(selected_topic_query),"s2","s1")&" ) AND ABS([s_active]) = 1 ORDER BY s1.s_order ASC"
			subject.Open SQL, Connect,3,3
		END IF
	
	end if

else

    if currentID="" THEN
        SQL = "SELECT *,(SELECT TOP 1 s2.s_order FROM new_subjects s2 WHERE s_typ = 2 AND s2.s_order > s1.s_order  AND ABS([s_active]) = 1 AND s2.s_qID = "&fixstr(clng(sID))&" ORDER BY s2.s_order ASC) NextQ,(SELECT TOP 1 s2.s_order FROM new_subjects s2 WHERE s_typ = 2 AND s2.s_order < s1.s_order  AND ABS([s_active]) = 1  AND s2.s_qID = "&fixstr(clng(sID))&" ORDER BY s2.s_order ASC) LastQ FROM new_subjects s1,subjects WHERE s1.s_qID = "&fixstr(clng(sID))&" AND s1.s_qiD = ID_subject  AND ABS([s_active]) = 1  ORDER BY s1.s_order"
        subject.Open SQL, Connect,3,3
    ELSE
        SQL = "SELECT *,(SELECT TOP 1 s2.s_order FROM new_subjects s2 WHERE s_typ = 2 AND s2.s_order > s1.s_order  AND ABS([s_active]) = 1 AND s2.s_qID = "&fixstr(clng(sID))&"  ORDER BY s2.s_order ASC) NextQ,(SELECT TOP 1 s2.s_order FROM new_subjects s2 WHERE s_typ = 2 AND s2.s_order < s1.s_order  AND ABS([s_active]) = 1  AND s2.s_qID = "&fixstr(clng(sID))&" ORDER BY s2.s_order ASC) LastQ FROM new_subjects s1,subjects WHERE s1.s_id = "&fixstr(clng(currentID))&" AND s1.s_qiD = ID_subject AND ABS([s_active]) = 1 ORDER BY s1.s_order ASC"
        subject.Open SQL, Connect,3,3
    END IF

end if 
'response.write SQL & "<br>"
if subject.EOF or subject.BOF then response.redirect("error.asp?" & request.QueryString  & "&line=111")


current_s_order = cint(subject("s_order"))
prev_s_order = (cint(subject("s_order")) - 1)
next_s_order = (cint(subject("s_order")) + 1)






' Finding the next ID in sequence
set objNext = Server.CreateObject("ADODB.Recordset")
if selected_topic_query <> "" then
    SQL = "SELECT TOP 1 s_topic,s_order,s_id,s_typ,(SELECT TOP 1 s2.s_typ FROM new_subjects s2 WHERE s2.s_order < s1.s_order AND ( "&cstr(selected_topic_query)&" ) AND ABS([s_active]) = 1 ) lastQuestion,(SELECT TOP 1 choice_cor FROM q_question,q_result,q_choice WHERE result_answer = ID_choice AND question_topic = s1.s_id AND result_question = ID_question AND result_session="&fixstr(clng(Session("sessionID")))&") qAnswer FROM new_subjects s1,subjects WHERE s1.s_qiD = ID_subject AND ( "&replace(cstr(selected_topic_query),"s2","s1")&" ) AND ABS([s_active]) = 1  AND s_qID = "&fixstr(clng(sID))&" AND s_order = "&fixstr(clng(next_s_order))&" ORDER BY s_order ASC"
Else
    SQL = "SELECT TOP 1 s_topic,s_order,s_id,s_typ,(SELECT TOP 1 s2.s_typ FROM new_subjects s2 WHERE s2.s_order < s1.s_order  AND ABS([s_active]) = 1 ) lastQuestion,(SELECT TOP 1 choice_cor FROM q_question,q_result,q_choice WHERE result_answer = ID_choice AND question_topic = s1.s_id AND result_question = ID_question AND result_session="&fixstr(clng(Session("sessionID")))&") qAnswer FROM new_subjects s1,subjects WHERE s1.s_qiD = ID_subject  AND ABS([s_active]) = 1  AND s_qID = "&fixstr(clng(sID))&" AND s_order = "&fixstr(clng(next_s_order))&" ORDER BY s_order ASC"
end if
'response.write SQL
objNext.Open SQL, Connect,3,3
nextAntal = objNext.RecordCount

do until objNext.eof
    IF clng(objNext("s_order")) > clng(subject("s_order")) THEN
        nextID = objNext("s_id")
    exit do
    ELSE
        nextID = 0
    END IF
objNext.movenext
loop

prevID = 0
prevQuiz = 0

if prev_s_order > 0 then
set objPrevious = Server.CreateObject("ADODB.Recordset")

if selected_topic_query <> "" then
    SQL = "SELECT TOP 1 s_topic,s_order,s_id,s_typ,(SELECT TOP 1 s2.s_typ FROM new_subjects s2 WHERE s2.s_order < s1.s_order AND ( "&cstr(selected_topic_query)&" ) AND ABS([s_active]) = 1 ) lastQuestion,(SELECT TOP 1 choice_cor FROM q_question,q_result,q_choice WHERE result_answer = ID_choice AND question_topic = s1.s_id AND result_question = ID_question AND result_session="&fixstr(clng(Session("sessionID")))&") qAnswer FROM new_subjects s1,subjects WHERE s1.s_qiD = ID_subject AND ( "&replace(cstr(selected_topic_query),"s2","s1")&" ) AND ABS([s_active]) = 1  AND s_qID = "&fixstr(clng(sID))&" AND s_order = "&fixstr(clng(prev_s_order))&" ORDER BY s_order ASC"
else
    SQL = "SELECT TOP 1 s_topic,s_order,s_id,s_typ,(SELECT TOP 1 s2.s_typ FROM new_subjects s2 WHERE s2.s_order < s1.s_order  AND ABS([s_active]) = 1 ) lastQuestion,(SELECT TOP 1 choice_cor FROM q_question,q_result,q_choice WHERE result_answer = ID_choice AND question_topic = s1.s_id AND result_question = ID_question AND result_session="&fixstr(clng(Session("sessionID")))&") qAnswer FROM new_subjects s1,subjects WHERE s1.s_qiD = ID_subject  AND ABS([s_active]) = 1  AND s_qID = "&fixstr(clng(sID))&" AND s_order = "&fixstr(clng(prev_s_order))&" ORDER BY s_order ASC"
end if
    'response.write SQL
objPrevious.Open SQL, Connect,3,3
previousAntal = objPrevious.RecordCount

do until objPrevious.eof 
    IF clng(objPrevious("s_order")) < clng(subject("s_order")) THEN
        IF clng(objPrevious("s_typ"))= 2 THEN
            prevID = 0
            prevQuiz = clng(objPrevious("s_order"))
        ELSE
            prevID = objPrevious("s_id")
        END IF
    END IF
objPrevious.movenext
loop

end if


IF clng(subject("s_typ")) = 1 THEN

set question = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT q_id,q_title,q_div_info,q_order FROM new_questions WHERE q_tID = "&fixstr(clng(subject("s_id")))&" ORDER BY q_order"
'response.write SQL & "<br>"
question.Open SQL, Connect,3,3
if not question.eof then
    showQuestion = true
    QArr = question.GetRows 
ELSE
    QArr = 0
END IF
question.close 


ELSEIF clng(subject("s_typ")) = 2 THEN

    set question = Server.CreateObject("ADODB.Recordset")
    SQL = "SELECT TOP 1 ID_question,question_body FROM q_question WHERE question_topic = "&fixstr(clng(subject("s_id")))&" AND ABS(question_active) = 1 ORDER BY newID()"
    'response.write SQL & "<br>"
    question.Open SQL, Connect,3,3
    QArr = question.GetRows 
    Session("question_ID")= QArr(0,0)
    Session("question_time") = now()
	

    set qchoice = Server.CreateObject("ADODB.Recordset")
    SQL = "SELECT id_choice,choice_label,choice_body,choice_cor FROM q_choice WHERE choice_question = "&fixstr(clng(QArr(0,0)))&" AND ABS(choice_active) = 1 ORDER BY choice_label"
    'response.write SQL & "<br>"
    qchoice.Open SQL, Connect,3,3
        showQuestion = true
    QchoiceArr = qchoice.GetRows 


    question.close : qchoice.close

END IF

' UPDATING CURRENT ID
if Err.Number = 0 then
set UPDATE = Server.CreateObject("ADODB.Recordset")
SQL = "UPDATE q_session SET session_current = "&fixstr(clng(subject("s_id")))&",session_stop = "&fixstr(clng(subject("s_order")))&" WHERE id_session = "&fixstr(clng(Session("sessionID")))&""	
UPDATE.Open SQL, Connect,3,3
end if

'**********************************************************************		
'GC: Modification to support Topic of and Page of Navigation 1-08-2013
'**********************************************************************
'Get the topics in the order they appear and the last ordinal for each topic
set topicOrdinals = Server.CreateObject("ADODB.Recordset")

if selected_topic_query <> "" then
SQL = "select s_topic, count(s_order) as topicCount, max(s_order) as lastTopicOrdinal from new_subjects s2 where s2.s_qId = "&fixstr(clng(sID))&" AND s2.s_active = 1 AND ( "&cstr(selected_topic_query)&" ) group by s_topic order by max(s_order)"
else
SQL = "select s_topic, count(s_order) as topicCount, max(s_order) as lastTopicOrdinal from new_subjects where s_qId = "&fixstr(clng(sID))&" and s_active = 1 group by s_topic order by max(s_order)"
end if

topicOrdinals.Open SQL, Connect, 3, 1 
'store a position for the current topic using a one-based counter and store 
'the previous topic last ordinal so we can work out where we are in the topic
Dim topicPosition, counter, previousTopicEnd, totalTopics, totalTopicQuestions
totalTopics = topicOrdinals.RecordCount
previousTopicEnd = 0
counter = 0
If Not topicOrdinals.EOF Then
    Do Until topicOrdinals.EOF
        counter = counter + 1
'		If selected_topic_query_last_query = "true" Then
'			previousTopicEnd = topicOrdinals("lastTopicOrdinal")
'			Exit Do
'		else
			If topicOrdinals("s_topic") = subject("s_topic") Then
				topicPosition = counter
				totalTopicQuestions = topicOrdinals("topicCount")
				Exit Do
			Else
				previousTopicEnd = topicOrdinals("lastTopicOrdinal")
			End If
'		end if
	
    topicOrdinals.MoveNext
    Loop
End If
topicOrdinals.Close
'redirect to the topic review if it is the first question of a new topic and we haven't come from the topic review page
If request.querystring("topicReviewed") = "" And subject("s_order") <> 1 And CInt(request.querystring("returning")) <> 1 Then
    If (subject("s_order") - previousTopicEnd) = 1 Then
        response.redirect("t_topic_review.asp?" & request.querystring)
    End If
End If
'**********************************************************************		
'GC: End of Modification to support Topic of and Page of Navigation 1-08-2013
'**********************************************************************





%>

<!doctype html>
<head>
   
    <title><%=client_name_short%> - Better Business Program</title>
        <META name="DESCRIPTION"	content="">
		<meta id="viewport" name='viewport'>
		<script src="jquery-1.11.1.js?v=bbp34"></script>
        <script>
            (function(doc) {
                if(window.innerHeight > window.innerWidth){
                swal({   title: "This site is best used in landscape view.",   text: "",   type: "warning",   confirmButtonText: "OK" }); 
                }
                if ( navigator.userAgent.match(/iPhone/i) || navigator.userAgent.match(/iPod/i)) {
                    viewport.setAttribute("content", "initial-scale=0.3");

                } else if ( navigator.userAgent.match(/iPad/i) ) {
                    viewport.setAttribute("content", "initial-scale=1.05");

                }

            }(document));
        </script>
	
        <!-- <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-T3c6CoIi6uLrA9TneNEoa7RxnatzjcDSCmG1MXxSR1GAsXEV/Dwwykc2MPK8M2HN" crossorigin="anonymous"> -->
        <!-- #include file = "inc_header.asp" -->
        <link href="style/question-page-styles.css" rel="stylesheet" type="text/css">
        <!-- <link
        rel="stylesheet"
        href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.3.0/css/all.min.css"
      /> -->
	 <link href="perfect-scrollbar-0.4.8/src/perfect-scrollbar.css" rel="stylesheet">
       <script src="jquery-1.11.1.js?v=bbp34"></script>
      <script src="perfect-scrollbar-0.4.8/src/jquery.mousewheel.js?v=bbp34"></script>
      <script src="perfect-scrollbar-0.4.8/src/perfect-scrollbar.js?v=bbp34"></script>
	 <link rel="stylesheet" type="text/css" href="js/sweet-alert.css">

    		  <script src="js/sweet-alert.min.js?v=bbp34"></script>
	  <script src="csspopup.js?v=bbp34"></script>

		<script>
			$(document).ready(function() {
			
			  $("html, body").animate({ scrollTop: $(document).height() }, 10);
			
			
            $(window).resize(function() {
                $("#popUpDiv").height($(window).height()/2);
				$("#popUpDiv").width($(window).width()/2);
				var window_width=$(window).width()/2-250;
				$("#popUpDiv").css("left", window_width + 'px');
            });
        
				//$("#blue_inside").perfectScrollbar({suppressScrollX: true});

				$(".inside_content").perfectScrollbar({suppressScrollX: true});
				
//SCROLL BAR CODE - This code is to make the scroll bar arrow disappear when scrolled to bottom
var $container = $(".inside_content");
$container.addClass("always-visible");
$container.scroll(function(e) {
  if($container.scrollTop() === 0) {
    // top
  }
  else if ($container.scrollTop() === $container.prop('scrollHeight') - $container.height()) {
    // end
	$('.ps-container .ps-scrollbar-y').css("background","#aaa");
    
  }
});

var $qcontainer = $("#question_name");
$qcontainer.scroll(function(e) {
  if($qcontainer.scrollTop() === 0) {
    // top
  }
  else if ($qcontainer.scrollTop() === $qcontainer.prop('scrollHeight') - $qcontainer.height()) {
    // end
	$('.ps-container .ps-scrollbar-y').css("background","#aaa");
    
  }
});
//------------------------------------------------------------------------------		
				
				// if ( $("#question_name").height() > 200 ){
				// $("#question_name").addClass( "inside_content2" );
				// $("#question_name").perfectScrollbar({suppressScrollX: true});
				// }
				// else
				// {
				// $("#question_name").removeClass( "inside_content2" );
				// }
				
				var real_height=0;
				$( "div[id^='button']" ).each(function( index ) {
				
				if($(this).height()>real_height)
				{
				real_height=$(this).height();
				}
  
				});
				
				$( "div[id^='button']" ).each(function( index ) {
				$(this).height(real_height);
				
  
				});
					$( "a[id^='link']" ).each(function( index ) {
				$(this).height(real_height);
				
  
				});
				
				
				
				});
				
				</script>
				
		
<script >



empty = true;
function trySubmit(){
    if (empty == true)
    {
        
		swal({   title: "You must make a selection before moving on.<br>Please click 'OK' and then select an answer before hitting 'Submit'.",   text: "",   type: "error", html:true,  confirmButtonText: "OK" });
        return false;
    }
        //This line below is nessessary to stop double form submissions from users double clicking
        document.getElementById('btnSubmit').disabled = true;
        return true;
}
function toggleRadio(thisField,thisValue){
   radioSet = eval("document.forms[0]."+thisField)
   for (i=0;i < radioSet.length;i++) {
      if (radioSet[i].value == thisValue)
         radioSet[i].checked = true
    }
}
b_click=true;
function gotonextpage(wheretogo)
{
if (b_click){swal({   title: "Please click on an information box to proceed.\n\nYou may need to scroll down the page to see the information boxes.",   text: "",   type: "error",   confirmButtonText: "OK" }); return} else {self.location=wheretogo}
}
function openwin(wname)
{
    self.open(wname, 'about', 'width=680,height=520,resizeable=yes,scrollbars=yes');
}

//CXS 03102007 - block use of LMS back controls
history.forward();

var shownLayer = null;
function showlayer(layer,layerbutton){
var ua = window.navigator.userAgent;
var msie = ua.indexOf("MSIE ");
var ie=false;
var ie10=false;
var other=false;

if (msie > 0 && msie<=7)      // If Internet Explorer, return version number
  ie=true;
		    //alert(parseInt(ua.substring(msie + 5, ua.indexOf(".", msie))));
else if(parseInt(ua.substring(msie + 5, ua.indexOf(".", msie))) > 7)
			ie10=true;
 else if (msie <= 0)                 // If another browser, return 0
 //alert('otherbrowser');
other=true;

           
<%
    IF showQuestion = true THEN
        If Ubound(QArr,2) > -1 Then
            For i=0 to ubound(QArr,2) %>
                document.getElementById("button<% =QArr(0,i)%>").className='t_choose'; 
                document.getElementById("button<% =QArr(0,i)%>").removeAttribute("disabled");
                var nodes = document.getElementById("button<% =QArr(0,i)%>").getElementsByTagName('*');

                for(var i = 0; i < nodes.length; i++)
                {
                    
                    if(ie10 || other) {
                        $("#"+'button<% =QArr(0,i)%>').children().attr("onclick","javascript:showlayer('layer<% =QArr(0,i)%>','button<% =QArr(0,i)%>'); b_click=false;");
                        $("#"+'button<% =QArr(0,i)%>').children().css("color","#0d4b77");
                        $("#"+'button<% =QArr(0,i)%>').css("background", "#fff")
                        $("#"+'button<% =QArr(0,i)%>').css("border-color", "transparent")
                        $("#"+'button<% =QArr(0,i)%>').children().css("font-weight","normal");
                    }
                    else {
                        nodes[i].disabled = false;
                    }
                }
            <%Next
        END IF
    END IF
%>



var myLayer = document.getElementById(layer).style.display.toLowerCase();
document.getElementById("layerfirst").style.display = 'none';
document.getElementById("blue_inside").style.display = 'block';

if(myLayer=="none"){
    document.getElementById(layer).style.display = 'block';
    if(shownLayer != null) document.getElementById(shownLayer).style.display = 'none';
   // document.getElementById(layerbutton).className='t_choose_active'; 
    shownLayer = layer;
	document.getElementById(layerbutton).setAttribute("disabled","disabled");
	var nodes = document.getElementById(layerbutton).getElementsByTagName('*');
	//nodes[0].innerHTML="test";
	
	if(ie10 || other)
	{
	$("#"+layerbutton).children().attr("onclick","javascript:void(0);");
    $("#"+layerbutton).css("background", "#C1D8E9")
    $("#"+layerbutton).css("border-color", "#0d4b77")
	$("#"+layerbutton).children().css("color","#0D4B77");
    $("#"+layerbutton).children().css("font-weight","700");
	}
else{ 

	
     nodes[0].disabled = true;
	 

}

} else {
document.getElementById(layer).style.display="none";
document.getElementById("layerfirst").style.display = 'block';
document.getElementById("blue_inside").style.display = 'none';
}
}

var message="Function Disabled!";

///////////////////////////////////
function clickIE4(){
if (event.button==2){
alert(message);
return false;
}
}

function clickNS4(e){
if (document.layers||document.getElementById&&!document.all){
if (e.which==2||e.which==3){
alert(message);
return false;
}
}
}

if (document.layers){
document.captureEvents(Event.MOUSEDOWN);
document.onmousedown=clickNS4;
}
else if (document.all&&!document.getElementById){
document.onmousedown=clickIE4;
}

document.oncontextmenu=new Function("return false;")

// --> 
</script>
</head>
<body>
    <div class="page-content">
        <div class="white-container">
            <!-- #include file = "partials/header.asp" -->

            <div id="blanket" style="display:none;"></div>
            <div id="popUpDiv" style="display:none;background-color:#FFF;border:#fff 10px solid;-moz-border-radius: 5px; border-radius: 5px;">
                <div style="background-color:#DBEEFF;height:50px;-moz-border-radius: 5px; border-radius: 5px;" >
                    &nbsp;<h3 style="float:left;padding:10px 40px 10px 10px; margin-left:12px;">Scenario</h3>
                    <div style="float:right;padding: 4px 4px 1px 2px;">
                        <a href="#" onclick="popup('popUpDiv')">
                            <div class="h_submit_blue" style="background-color:#0354AC;padding:10px;-moz-border-radius: 5px; border-radius: 5px;">CLOSE</div>
                        </a>
                    </div>
                </div>
                <div style="margin-right:0px;margin-top:5px;width:100%;height:87%;background-color:#DBEEFF;-moz-border-radius: 5px; border-radius: 5px;" id="popcontent"></div>
            </div>

            <!-- #include file = "partials/sub-header.asp" -->
            <!-- #include file = "partials/question-page-content.asp" -->
        </div>
      <!-- #include file = "partials/footer.asp" -->
    </div>

</body>
<%
IF clng(subject("s_typ")) = 1 THEN
  call log_the_page("Training and quiz", Session("ID"), (subject("s_topic")), (subject("s_id")), (subject("s_title")), 0, qst, "Training")
ELSE
    call log_the_page("Training and quiz", Session("ID"), "Quiz", (subject("s_id")), (subject("s_title")), 0, qst, "Quiz")
END IF
%>
<% 
objNext.close : Set objNext = Nothing
subject.close : Set subject = Nothing%>
<!-- #include file = "errorhandler/index.asp"-->