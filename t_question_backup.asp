<%@LANGUAGE="VBSCRIPT"%>

<% 
'Response buffer is used to buffer the output page. That means if any database exception occurs the contents can be cleared without processed any script to browser
 Response.Buffer = True
 
' "On Error Resume Next" method allows page to move to the next script even if any error present on page whcich will be caught after processing all asp script on page
 On Error Resume Next
 
'Changed by PR on 25.02.16
%>
 
<% bbp_training = true%>
<!-- #include file = "connections/bbg_conn.asp" -->
<!-- #include file = "connections/include.asp"-->

<%
if NOT pref_quiz_avail then response.redirect("error.asp?" & request.QueryString)
'Used for bookmarking to check if the question has been answered or not
Session("answered")=0

If  Request.querystring("quiz") <> "" Then

    
    response.write("question_ID")
    response.write("sessionID")
    response.write("question_time")

    ' Check if SessionID existx
        if (Session("sessionID") = "") OR (Session("question_ID") = "") Then Response.Redirect("error.asp?" & request.QueryString)
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
		
        response.redirect "t_feedback.asp"
        
        response.write cLng(Session("sessionID"))
END IF


' ID of users session
Dim sID
if (Session("ID") <> "") Then
    sID = Session("ID")
Else
    Response.Redirect("error.asp?" & request.QueryString)
End If

if request.querystring("nextID")<>"" THEN
    currentID = request.querystring("nextID")
END IF

showQuestion = False


set subject = Server.CreateObject("ADODB.Recordset")
if currentID="" THEN
    SQL = "SELECT *,(SELECT TOP 1 s2.s_order FROM new_subjects s2 WHERE s_typ = 2 AND s2.s_order > s1.s_order  AND ABS([s_active]) = 1 AND s2.s_qID = "&fixstr(clng(sID))&" ORDER BY s2.s_order ASC) NextQ,(SELECT TOP 1 s2.s_order FROM new_subjects s2 WHERE s_typ = 2 AND s2.s_order < s1.s_order  AND ABS([s_active]) = 1  AND s2.s_qID = "&fixstr(clng(sID))&" ORDER BY s2.s_order ASC) LastQ FROM new_subjects s1,subjects WHERE s1.s_qID = "&fixstr(clng(sID))&" AND s1.s_qiD = ID_subject  AND ABS([s_active]) = 1  ORDER BY s1.s_order"
    subject.Open SQL, Connect,3,3
ELSE
    SQL = "SELECT *,(SELECT TOP 1 s2.s_order FROM new_subjects s2 WHERE s_typ = 2 AND s2.s_order > s1.s_order  AND ABS([s_active]) = 1 AND s2.s_qID = "&fixstr(clng(sID))&"  ORDER BY s2.s_order ASC) NextQ,(SELECT TOP 1 s2.s_order FROM new_subjects s2 WHERE s_typ = 2 AND s2.s_order < s1.s_order  AND ABS([s_active]) = 1  AND s2.s_qID = "&fixstr(clng(sID))&" ORDER BY s2.s_order ASC) LastQ FROM new_subjects s1,subjects WHERE s1.s_id = "&fixstr(clng(currentID))&" AND s1.s_qiD = ID_subject AND ABS([s_active]) = 1 ORDER BY s1.s_order ASC"
    subject.Open SQL, Connect,3,3
END IF
'response.write SQL & "<br>"
if subject.EOF or subject.BOF then response.redirect("error.asp?" & request.QueryString)

prevID = 0
prevQuiz = 0
' Finding the next ID in sequence
set objNext = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT s_topic,s_order,s_id,s_typ,(SELECT TOP 1 s2.s_typ FROM new_subjects s2 WHERE s2.s_order < s1.s_order  AND ABS([s_active]) = 1 ) lastQuestion,(SELECT TOP 1 choice_cor FROM q_question,q_result,q_choice WHERE result_answer = ID_choice AND question_topic = s1.s_id AND result_question = ID_question AND result_session="&fixstr(clng(Session("sessionID")))&") qAnswer "
SQL = SQL & " FROM new_subjects s1,subjects WHERE s1.s_qiD = ID_subject  AND ABS([s_active]) = 1  AND s_qID = "&fixstr(clng(sID))&" ORDER BY s_order ASC"
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

    IF clng(objNext("s_order")) < clng(subject("s_order")) THEN
        IF clng(objNext("s_typ"))= 2 THEN
            prevID = 0
            prevQuiz = clng(objNext("s_order"))
        ELSE
            prevID = objNext("s_id")
        END IF
    END IF


objNext.movenext
loop

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
SQL = "select s_topic, count(s_order) as topicCount, max(s_order) as lastTopicOrdinal from new_subjects where s_qId = "&fixstr(clng(sID))&" and s_active = 1 group by s_topic order by max(s_order)"
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
        If topicOrdinals("s_topic") = subject("s_topic") Then
            topicPosition = counter
            totalTopicQuestions = topicOrdinals("topicCount")
            Exit Do
        Else
            previousTopicEnd = topicOrdinals("lastTopicOrdinal")
        End If
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
	
        <!-- #include file = "inc_header.asp" -->
     
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
				
				if ( $("#question_name").height() > 200 ){
				
			
				$("#question_name").addClass( "inside_content2" );
				$("#question_name").perfectScrollbar({suppressScrollX: true});
				}
				else
				{
				
				$("#question_name").removeClass( "inside_content2" );
				}
				
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
     
	 if(ie10 || other){
	$("#"+'button<% =QArr(0,i)%>').children().attr("onclick","javascript:showlayer('layer<% =QArr(0,i)%>','button<% =QArr(0,i)%>'); b_click=false;");
	$("#"+'button<% =QArr(0,i)%>').children().css("color","#1e89db");
	}
	else
	nodes[i].disabled = false;
}
<%Next
END IF
END IF%>



var myLayer = document.getElementById(layer).style.display.toLowerCase();
document.getElementById("layerfirst").style.display = 'none';
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
	$("#"+layerbutton).children().css("color","#939393");
	
	}
else{ 

	
     nodes[0].disabled = true;
	 

}

} else {
document.getElementById(layer).style.display="none";
document.getElementById("layerfirst").style.display = 'block';
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
<div id="blanket" style="display:none;"></div>
<div id="popUpDiv" style="display:none;background-color:#FFF;border:#fff 10px solid;-moz-border-radius: 5px; border-radius: 5px;">
<div style="background-color:#DBEEFF;height:50px;-moz-border-radius: 5px; border-radius: 5px;" >&nbsp;<h3 style="float:left;padding:10px 40px 10px 10px; margin-left:12px;">Scenario</h3>
<div style="float:right;padding: 4px 4px 1px 2px;">
<a href="#" onclick="popup('popUpDiv')">
<div class="h_submit_blue" style="background-color:#0354AC;padding:10px;-moz-border-radius: 5px; border-radius: 5px;">CLOSE</div></a></div></div>
<div style="margin-right:0px;margin-top:5px;width:100%;height:87%;background-color:#DBEEFF;-moz-border-radius: 5px; border-radius: 5px;" id="popcontent"></div>
</div>
<img src="images/bg.jpg" alt="background image" id="bg">

<!-- #include file = "main_top.asp" -->

    <div class="header_blue">
        
        <div class="header_progress" style="padding-top:28px;">
            <div class="topic_progress">Topic <div class="topicNumberCircle"><%=topicPosition%></div> of <%=totalTopics%></div>
            <div class="page_progress">Page <div class="pageNumberCircle"><%=subject("s_order") - previousTopicEnd%></div> of <%=totalTopicQuestions%> in this topic</></div>		
        </div>
        
        <div class="header_inside"><%=Ucase(ReplaceStrQuiz(subject("subject_name")))%><br>
        <h3><%=ReplaceStrQuiz(subject("s_topic"))%></h3>
        </div>
    </div>
    <a name="answer"></a>
    <div class="clear"></div>
    <div class="allcontent_content">
        <% ' s_typ 1 = Traning, 2 = Quiz 
        IF clng(subject("s_typ")) = 1 THEN%>
        <div class="main_content">
        <div class="guide_blue">
            <div class="box_inside"><h3><%=ReplaceStrQuiz(subject("s_title"))%></h3>
             <div class="box_text_blue">
			<div class="inside_content" > 
            <p ><%=ReplaceStrQuiz(subject("s_body"))%></p>
			
            <%
            IF showQuestion = true THEN
                If Ubound(QArr,2) > -1 Then
                xi = 0
                     For i=0 to ubound(QArr,2)
                     xi=xi+1 %>
                    <div class="t_choose" id="button<% =QArr(0,i)%>"><a id="link<% =QArr(0,i)%>" href="#answer" onclick="javascript:showlayer('layer<% =QArr(0,i)%>','button<% =QArr(0,i)%>'); b_click=false;"><% =ReplaceStrQuiz(QArr(1,i))%></a></div>
                    <% IF xi=2 THEN
                    xi=0
                    response.write "<div class=""clear""></div>"
                    END IF%>
            <%		Next
                END IF
            END IF%>
			<% IF clng(subject("s_goback"))>0 THEN%>
                <br>
				

                <div class="certificate_blue">
				
                    <div class="h_submit_blue"><!--<a onClick="window.open('t_question_window.asp?s_id=<% =subject("s_goback")%>','bppWindow','width=620,height=425,left=20,top=20,scrollbars=yes')" href="javascript:void(0)" class="box_link" style="padding-left:15px;">--><a class="box_link" style="padding-left:15px;" href="#" onclick="popup('popUpDiv','t_question_window.asp?s_id=<% =subject("s_goback")%>')">GO BACK AND SEE THIS SCENARIO AGAIN</a></div>
                </div>
                <div class="clear"></div>
                <% END IF%>
            <div class="clear"></div>
			</div>
                
                  <br>
              
              <% ' If last page on traning & quiz
                IF clng(NextID) = 0 THEN%>
                <img src="images/lyte_loading.gif" width="26" height="26" alt="" style="vertical-align:middle;"> Recording score...<br><br>
                You have now reached the end of the quiz.<br>Please wait while your score is recorded for this course.<br><br>
                If you are not redirected automatically please click the 'Certificate' button.<br>
                <br>
                <script type="text/javascript">
                     window.setTimeout("document.location = 'certificate.asp';", 7000);
                </script>
                
            <div class="div_button">
                <div class="next_blue" style="width:160px;" id="certend">
                    <div class="h_submit_blue"><a href="certificate.asp" class="box_link" style="padding-left:15px;">CERTIFICATE</a></div>
                </div>
            </div>
                <% ELSE%>
                
                <div class="div_button">
                <% IF clng(prevID)<>0 And subject("s_order") - previousTopicEnd <> 1 THen%>
                    <div class="back_blue">
                        <div class="h_submit_blue"><a href="t_question.asp?nextID=<% =prevID%>&returning=1" class="box_link" style="padding-left:65px;">BACK</a></div>
                    </div>
                <% end if%>
                    <div class="next_blue">
                <% IF showQuestion = true THEN%>
                    <div class="h_submit_blue" ><a href="javascript:gotonextpage('t_question.asp?nextID=<% =nextID%>')" onClick="closing=false" class="box_link" style="padding-left:15px;">NEXT</a></div>
                <% ELSE%>
                    <div class="h_submit_blue" ><a href="t_question.asp?nextID=<% =nextID%>" class="box_link" style="padding-left:15px;">NEXT</a></div>
                <% END IF%>
                </div>
            <% END IF %>
            </div>
    
            
                </div>
            </div>
        </div>
    </div>
    <div class="menu_content" id="layerfirst">
    <% IF subject("s_image")<>"" THEN%>
    <img src="vault_image/images/<% =subject("s_image")%>" alt="">
    <% ELSE%>
    <img src="vault_image/images/training.jpg" width="320" height="390" alt="">
    <% END IF%>
    </div>
	<div id="blue_inside">
    <% 
         IF showQuestion = true THEN
         For i=0 to ubound(QArr,2) %>
		 
        <div class="t_content_info" id="layer<% =QArr(0,i)%>" style="display:none;">
            <div class="t_content_info_inside"><% =ReplaceStrQuiz(QArr(2,i))%></div>
        </div>
		
         <%Next
         END IF%>
		</div> 
		 
    
    <% ' s_typ 1 = Traning, 2 = Quiz 
    ELSEIF clng(subject("s_typ")) = 2 THEN%>
    <div class="quiz_blue">
        <div class="box_inside">
            <h3 style="color:#FFF;">
                <img src="images/icon_quiz.gif" width="22" height="25" alt="" style="vertical-align:middle;margin-right:20px;"> Quiz</h3>
            </div>
        </div>
        <div class="clear"></div>
        
        <div class="main_content">
            <div class="quiz_question_blue">
                <div class="box_inside2">
				<div   id="question_name"	> 
                    <h1><%=ReplaceStrQuiz(QArr(1,0))%></h1></div>
                    <div class="box_text_blue" style="margin-top:15px;">
                        <form name="quiz" method="POST" onsubmit="return trySubmit();" action="t_question.asp?currentID=<% =subject("s_id")%>&quiz=yes">
                        <% 
                         IF showQuestion = true THEN
                             For i=0 to ubound(QchoiceArr,2) %>
                            <label style="cursor:pointer; display:block;" for="rbutton<% =QchoiceArr(0,i)%>">
                                <div style="margin:8px 0px 8px 0px;cursor:pointer;">
                                    <div style="float:left;width:30px;padding-top:6px;"><input type="radio" name="answer" value="<% =QchoiceArr(0,i)%>" onclick="empty = false" id='rbutton<% =QchoiceArr(0,i)%>'>
									</div>
                                    <div style="float:left;width:25px;padding-top:6px;"><strong><% =ReplaceStrQuiz(QchoiceArr(1,i))%></strong>
									</div>
                                    <div style="float:left;width:810px;"><div class="quiz_choose"><div class="quiz_choose_inside"><% =ReplaceStrQuiz(QchoiceArr(2,i))%>
									</div>
								</div>
								</div>
                                </div>
                            </label>
                            <div class="clear"></div>
                             <%Next
                         END IF%><br>
                           
							<div class="div_button" style="width:860px;text-align:right;margin:20px 0px;">
							<div class="next_blue">
							<input type="submit" class="start_blue" name="btnSubmit" id="btnSubmit" value="SUBMIT">
							</div>
							</div>
                        </form>
                    </div>
                </div>
            </div>
        </div>
        
    <% END IF%>
    <div class="clear"></div>
    </div>
        
     
    <div class="clear"></div>
</div>
</div>
<% 
 %>
<!-- #include file = "inc_bottom.asp" -->
</html>
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