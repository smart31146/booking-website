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

Dim selected_topic_query
if (cstr(Session("selected_topics")) <> "") Then
    selected_topic_query = cstr(Session("selected_topics"))
End If

if NOT pref_quiz_avail then response.redirect("error.asp?" & request.QueryString)
showQuestion = False

if (Session("currentID") = "") Then Response.Redirect("error.asp?" & request.QueryString)
' ID of users session
Dim SessionID
if (Session("ID") <> "") Then
    sID = Session("ID")
Else
    Response.Redirect("error.asp?" & request.QueryString)
End If


IF request.querystring("alt") = "startover" THEN

    ArrToDelete = split(request.QueryString("todelete"),"||")
    For xT =0 to UBound(ArrToDelete)
            MM_editConnection = Connect
            MM_editTable = "q_session"
            MM_editQuery = "DELETE FROM q_result WHERE ID_result = "&  ArrToDelete(xT) &" "
if Err.Number = 0 then
            Set MM_editCmd = Server.CreateObject("ADODB.Command")
            MM_editCmd.ActiveConnection = MM_editConnection
            MM_editCmd.CommandText = MM_editQuery
            MM_editCmd.Execute
end if
            'response.write MM_editQuery & "<br><br>"
    next
    'SQL = "DELETE FROM q_result WHERE ID_result "
    'response.end
    response.redirect "t_question.asp?nextID="& request.querystring("startID") & ""
END IF

'response.end
' the correct choice based on the ID of requested question
set correct = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT *  FROM q_choice  WHERE choice_question = "&fixstr(clng(Session("question_ID")))&" AND choice_cor = 1 AND  ABS([choice_active]) = 1  ORDER BY id_choice"
'response.write SQL & "<br>"
correct.Open SQL, Connect,3,3
Dim Correct_answer
if NOT correct.EOF or NOT correct.EOF then
    Correct_answer = (correct("ID_choice"))
else
    Correct_answer = 0
end if
correct.Close()




set subject = Server.CreateObject("ADODB.Recordset")

if selected_topic_query <> "" then
	SQL = "SELECT * FROM new_subjects s2,subjects WHERE s2.s_id = "&fixstr(clng(Session("currentID")))&" AND ( "&cstr(selected_topic_query)&" ) AND s2.s_qiD = ID_subject"
else
	SQL = "SELECT * FROM new_subjects s1,subjects WHERE s1.s_id = "&fixstr(clng(Session("currentID")))&" AND s1.s_qiD = ID_subject"
end if
'response.write SQL & "<br><br>"
subject.Open SQL, Connect,3,3
if subject.EOF or subject.BOF then response.redirect("error.asp?" & request.QueryString)


' Finding the next ID in sequence
set objNextID = Server.CreateObject("ADODB.Recordset")

if selected_topic_query <> "" then
	SQL = "SELECT  s_topic,s_order,s_id,s_typ FROM new_subjects s2,subjects WHERE s2.s_qiD = ID_subject AND  ABS([s_active]) = 1 AND ( "&cstr(selected_topic_query)&" ) AND s2.s_qID = "&fixstr(clng(sID))&"  ORDER BY s_order ASC"
else
	SQL = "SELECT  s_topic,s_order,s_id,s_typ FROM new_subjects s1,subjects WHERE s1.s_qiD = ID_subject AND  ABS([s_active]) = 1 AND s1.s_qID = "&fixstr(clng(sID))&"  ORDER BY s_order ASC"
end if

'response.write SQL& "<br>"
objNextID.Open SQL, Connect,3,3
do until objNextID.eof 
    IF clng(objNextID("s_order")) > clng(subject("s_order")) THEN
        nextTyp = objNextID("s_typ")
        nextID = objNextID("s_id")
    exit do
    END IF
objNextID.movenext
loop
objNextID.close

set question = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT TOP 1 ID_question,question_body,question_fb_cor,question_fb_inc FROM q_question WHERE ID_question = "&fixstr(clng(Session("question_ID")))&" "
question.Open SQL, Connect,3,3
QArr = question.GetRows 

set qchoice = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT id_choice,choice_label,choice_body,choice_cor FROM q_choice WHERE choice_question = "&fixstr(clng(QArr(0,0)))&" AND ABS(choice_active) = 1 ORDER BY choice_label"
qchoice.Open SQL, Connect,3,3
showQuestion = true
QchoiceArr = qchoice.GetRows 
question.close : qchoice.close


IF clng(Correct_answer) = clng(Session("answer")) THEN
    ' CORRECT ANSWER
    message = QArr(2,0)
    correct = true
ELSE
    ' INCORRECT ANSWER
    message = QArr(3,0)
    correct = false
END IF
'**********************************************************************		
'GC: Modification to support Topic of and Page of Navigation 10-08-2013
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
If request.querystring("topicReviewed") = "" And subject("s_order") <> 1 Then
    If (subject("s_order") - previousTopicEnd) = 1 Then
        'response.redirect("t_topic_review.asp?" & request.querystring)
    End If
End If
'**********************************************************************		
'GC: End of Modification to support Topic of and Page of Navigation 10-08-2013
'**********************************************************************
Session("topic_name")=ReplaceStrQuiz(subject("s_topic"))
%>

<!doctype html>
<head>
    <meta id="viewport" name='viewport'>
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
    <title><%=client_name_short%> - Building a Better Workplace</title>
    <META name="DESCRIPTION"	content="">
    <!-- #include file = "inc_header.asp" -->
    <link href="perfect-scrollbar-0.4.8/src/perfect-scrollbar.css" rel="stylesheet">
    <script src="//ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js?v=bbp34"></script>
    <script src="perfect-scrollbar-0.4.8/src/jquery.mousewheel.js?v=bbp34"></script>
    <script src="perfect-scrollbar-0.4.8/src/perfect-scrollbar.js?v=bbp34"></script>
    <link rel="stylesheet" type="text/css" href="js/sweet-alert.css">
    <script src="js/sweet-alert.min.js?v=bbp34"></script>
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
                    
            if ( $("#question_name").height() > 200 ) {
                $("#question_name").addClass( "inside_content2" );
                $("#question_name").perfectScrollbar({suppressScrollX: true});
            }
            else {
                $("#question_name").removeClass( "inside_content2" );
            }
            
            var real_height=0;
            $( "div[id^='button']" ).each(function( index ) {
            
                if($(this).height()>real_height) {
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
</head>
<body>
    <div class="page-content">
        <div class="white-container">
            <!-- #include file = "partials/header.asp" -->

            <div class="allcontent">
                <div class="allcontent_main">
                    <div class="header_blue">
                        
                        <div class="header_progress" style="padding-top:28px;">
                            <div class="topic_progress">Topic <div class="topicNumberCircle"><%=topicPosition%></div> of <%=totalTopics%></div>
                            <div class="page_progress">Page <div class="pageNumberCircle"><%=subject("s_order") - previousTopicEnd%></div> of <%=totalTopicQuestions%> in this topic</div>
                        </div>
                        
                        <div class="header_inside"><%=Ucase(ReplaceStrQuiz(subject("subject_name")))%><br>
                        <h3><%=ReplaceStrQuiz(subject("s_topic"))%></h3>
                        </div>
                    </div>
                
                    <div class="clear"></div>

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
                            <div id="question_name" > 
                                <h1><%=ReplaceStrQuiz(QArr(1,0))%></h1></div>
                                <div class="box_text_blue" style="margin-top:15px;">
                                    
                                    <% 
                                    IF showQuestion = true THEN
                                        For i=0 to ubound(QchoiceArr,2) %>
                                        
                                            <div style="margin-top:8px;cursor:pointer;">
                                                <div style="float:left;width:30px;padding-top:4px;"><% IF clng(Session("answer")) = clng(QchoiceArr(0,i)) AND correct = false THEN %><img src="images/icon_false.png" width="21" height="22" alt=""><% END IF%><% IF clng(Correct_answer) = clng(QchoiceArr(0,i)) THEN %><img src="images/icon_true.gif" alt=""><% END IF%></div>
                                                <div style="float:left;width:25px;padding-top:6px;"><strong><% =ReplaceStrQuiz(QchoiceArr(1,i))%></strong></div>
                                                <div style="float:left;width:810px;"><div class="quiz_choose"><div class="quiz_choose_inside"><% =ReplaceStrQuiz(QchoiceArr(2,i))%></div></div></div>
                                            </div>
                                        
                                        <div class="clear"></div>
                                        <%Next
                                    END IF%>
                                        <span class="answerfeed"><br><br><strong><% =ReplaceStrQuiz(message) %></strong><br></span>
                                        <div style="width:860px;text-align:right;">
                                        
                                        
                            <% 
                            
                            ' ADDED 2013-03-07 by Johan Bredenholt johan@america.se
                            ' This is if the user can proceed if the user get at least 1 incorrect answer on the questions
                            subject_proceed = 1
                            
                            set objP = Server.CreateObject("ADODB.Recordset")
                            IF clng(nextTyp) = 1 THEN
                            SQL = "SELECT subject_proceed,subject_name  FROM subjects  WHERE ID_subject = "&fixstr(clng(sID))&""
                            objP.Open SQL, Connect,3,3
                                subject_proceed = objP("subject_proceed")
                            objP.Close()
                            END IF
                            
                            ' Checking if there is more question in same topic

                            showNext = False
                            if selected_topic_query <> "" then
                                SQL = "SELECT s_topic,s_order,s_id,s_typ FROM new_subjects s2,subjects WHERE s2.s_qiD = ID_subject AND ( "&cstr(selected_topic_query)&" ) AND ABS([s_active]) = 1 AND s_qID = "&fixstr(clng(sID))&" AND s_typ = 2  AND s_order > "& clng(trim(subject("s_order"))) &" ORDER BY s_order DESC"
                            else
                                SQL = "SELECT s_topic,s_order,s_id,s_typ FROM new_subjects s1,subjects WHERE s1.s_qiD = ID_subject AND ABS([s_active]) = 1 AND s_qID = "&fixstr(clng(sID))&" AND s_typ = 2  AND s_order > "& clng(trim(subject("s_order"))) &" ORDER BY s_order DESC"
                            end if
                            
                            'response.write SQL
                            objP.Open SQL, Connect,3,3
                            IF objP.eof THEN showNext = True ELSE showNext = False
                            objP.Close()

                            
                            ' Set this to true if the check should be after each quiz in a topic
                            'showNext = True
                            IF showNext = True THEN
                            %>
                            
                            <%
                            if nextID = "" then
                                nextID = cstr("last")
                            end if
                            
                            %>
                            
                                
                                <%IF clng(subject_proceed) = 1 THEN %>		
                                <div class="next_blue" style="text-align:left;margin:20px 0px;">
                                    <div class="h_submit_blue"><a href="t_question.asp?nextID=<% =nextID%>" class="box_link" style="padding-left:15px;">NEXT</a></div>
                                </div>
                                <% END IF%>
                                
                                
                                <%IF clng(subject_proceed) = 2 THEN 
                                correctAnswer = True
                                set objS = Server.CreateObject("ADODB.Recordset")
                                if selected_topic_query <> "" then
                                    SQL = "SELECT s_topic,s_order,s_id,s_typ, (SELECT TOP 1 s_id FROM new_subjects s2 WHERE s_topic = '"& trim(subject("s_topic")) &"' AND s_typ = 1 AND ( "&cstr(selected_topic_query)&" ) ORDER BY s_order ASC) qStart, (SELECT TOP 1 choice_cor FROM q_question,q_result,q_choice WHERE result_answer = ID_choice AND question_topic = s1.s_id AND result_question = ID_question AND result_session="&fixstr(clng(Session("sessionID")))&") qAnswer, (SELECT TOP 1 ID_result FROM q_question,q_result,q_choice WHERE result_answer = ID_choice AND question_topic = s1.s_id AND result_question = ID_question AND result_session="&fixstr(clng(Session("sessionID")))&") qID FROM new_subjects s1,subjects WHERE s1.s_qiD = ID_subject AND ( "&replace(cstr(selected_topic_query),"s2","s1")&" ) ABS([s_active]) = 1 AND s_qID = "&fixstr(clng(sID))&" AND s_typ = 2 AND s_topic = '"& trim(subject("s_topic")) &"' ORDER BY s_order DESC"
                                else
                                    SQL = "SELECT s_topic,s_order,s_id,s_typ, (SELECT TOP 1 s_id FROM new_subjects WHERE s_topic = '"& trim(subject("s_topic")) &"' AND s_typ = 1 ORDER BY s_order ASC) qStart, (SELECT TOP 1 choice_cor FROM q_question,q_result,q_choice WHERE result_answer = ID_choice AND question_topic = s1.s_id AND result_question = ID_question AND result_session="&fixstr(clng(Session("sessionID")))&") qAnswer, (SELECT TOP 1 ID_result FROM q_question,q_result,q_choice WHERE result_answer = ID_choice AND question_topic = s1.s_id AND result_question = ID_question AND result_session="&fixstr(clng(Session("sessionID")))&") qID FROM new_subjects s1,subjects WHERE s1.s_qiD = ID_subject AND ABS([s_active]) = 1 AND s_qID = "&fixstr(clng(sID))&" AND s_typ = 2 AND s_topic = '"& trim(subject("s_topic")) &"' ORDER BY s_order DESC"
                                end if
                                
                                'response.write SQL & "<br>"
                                objS.Open SQL, Connect,3,3
                                IF not objS.eof then
                                    startProceed = objS("s_order")
                                    s_id = objS("qStart")
                                    x = 0
                                    todelete = ""
                                    do until objS.eof 
                                            'IF startProceed = clng(objS("s_order")) THEN
                                                's_id = objS("s_id")
                                                todelete = todelete & objS("qID") & "||"
                                                s_order = objS("s_order")
                                                IF correctAnswer = TRUE THEN
                                                    IF cbool(objS("qAnswer")) = False  THEN
                                                        correctAnswer = False
                                                    END IF
                                                END IF
                                            'ELSE
                                            '	exit do
                                            'END IF
                                            startProceed = startProceed - 1
                                            x = x + 1
                                    objS.movenext
                                    loop
                
                                END IF
                                objS.Close()
                                'response.write correctAnswer & "<br>"
                                'response.write x & "<br>"
                                'response.write s_order & "<br>"
                                'response.write s_id & "<br>"
                                'response.write startProceed & "<br>"
                                'response.write SQL
                                IF correctAnswer = True THEN%>		
                                    <div class="next_blue" style="text-align:left;margin:20px 0px;">
                                        <div class="h_submit_blue"><a href="t_question.asp?nextID=<% =nextID%>" class="box_link" style="padding-left:15px;">NEXT</a></div>
                                    </div>
                                <% ELSE%>
                                    <div class="clear"></div>
                                    <div style="float:right;">
                                    <div class="next_blue" style="text-align:left;margin:20px 0px;">
                                            <div class="h_submit_blue"><a href="t_question.asp?nextID=<% =nextID%>" class="box_link" style="padding-left:15px;">NEXT</a></div>
                                        </div>
                                        
                                        <div class="next_blue" style="text-align:left;margin:20px 0px;">
                                            <div class="h_submit_blue"><a href="t_feedback.asp?alt=startover&amp;nextID=<% =nextID%>&amp;startID=<% =s_id %>&amp;todelete=<% IF len(todelete) > 1 THEN response.write  left(todelete,len(todelete)-2)%>" class="box_link" style="padding-left:15px;">START OVER</a></div>
                                            
                                        </div>
                                        
                                    </div>
                                    <div style="float:right;margin-top:25px;margin-right:20px;">You did not answer all questions correct. </div>
                                    <div class="clear"></div>
                                <% END IF%>
                                <% END IF%>
                                
                                
                                
                                
                                <%
                                IF clng(subject_proceed) = 3 THEN 
                                
                                
                                
                                correctAnswer = True
                                set objS = Server.CreateObject("ADODB.Recordset")
                                if selected_topic_query <> "" then
                                    SQL = "SELECT s_topic,s_order,s_id,s_typ, (SELECT TOP 1 s_id FROM new_subjects s2 WHERE s_topic = '"& trim(subject("s_topic")) &"' AND ( "&cstr(selected_topic_query)&" ) s_qID = "&fixstr(clng(sID))& " AND s_typ = 1 ORDER BY s_order ASC) qStart, (SELECT TOP 1 choice_cor FROM q_question,q_result,q_choice WHERE result_answer = ID_choice AND question_topic = s1.s_id AND result_question = ID_question AND result_session="&fixstr(clng(Session("sessionID")))&") qAnswer, (SELECT TOP 1 ID_result FROM q_question,q_result,q_choice WHERE result_answer = ID_choice AND question_topic = s1.s_id AND result_question = ID_question AND result_session="&fixstr(clng(Session("sessionID")))&") qID FROM new_subjects s1,subjects WHERE s1.s_qiD = ID_subject AND ( "&replace(cstr(selected_topic_query),"s2","s1")&" ) ABS([s_active]) = 1 AND s_qID = "&fixstr(clng(sID))&" AND s_typ = 2 AND s_topic = '"& trim(subject("s_topic")) &"' ORDER BY s_order DESC"
                                else
                                    SQL = "SELECT s_topic,s_order,s_id,s_typ, (SELECT TOP 1 s_id FROM new_subjects WHERE s_topic = '"& trim(subject("s_topic")) &"' AND s_qID = "&fixstr(clng(sID))& " AND s_typ = 1 ORDER BY s_order ASC) qStart, (SELECT TOP 1 choice_cor FROM q_question,q_result,q_choice WHERE result_answer = ID_choice AND question_topic = s1.s_id AND result_question = ID_question AND result_session="&fixstr(clng(Session("sessionID")))&") qAnswer, (SELECT TOP 1 ID_result FROM q_question,q_result,q_choice WHERE result_answer = ID_choice AND question_topic = s1.s_id AND result_question = ID_question AND result_session="&fixstr(clng(Session("sessionID")))&") qID FROM new_subjects s1,subjects WHERE s1.s_qiD = ID_subject AND ABS([s_active]) = 1 AND s_qID = "&fixstr(clng(sID))&" AND s_typ = 2 AND s_topic = '"& trim(subject("s_topic")) &"' ORDER BY s_order DESC"
                                end if
                                
                                'response.write SQL & "<br>"
                                objS.Open SQL, Connect,3,3
                                IF not objS.eof then
                                    startProceed = objS("s_order")
                                    s_id = objS("qStart")
                                    x = 0
                                    todelete = ""
                                    do until objS.eof 
                                                todelete = todelete & objS("qID") & "||"
                                                s_order = objS("s_order")
                                                IF correctAnswer = TRUE THEN
                                                    IF cbool(objS("qAnswer")) = False  THEN
                                                        correctAnswer = False
                                                    END IF
                                                END IF
                                            startProceed = startProceed - 1
                                            x = x + 1
                                    objS.movenext
                                    loop
                
                                END IF
                                objS.Close()
                                IF correctAnswer = True THEN%>		
                                    <div class="next_blue" style="text-align:right;margin:20px 0px;">
                                        <div class="h_submit_blue"><a href="t_question.asp?nextID=<% =nextID%>" class="box_link" style="padding-left:15px;">NEXT</a></div>
                                    </div>
                                <% ELSE%>
                                    <div class="clear"></div>
                                    <div style="float:right;">
                                        <div class="next_blue" style="text-align:right;margin:20px 0px;">
                                            <div class="h_submit_blue"><a href="t_feedback.asp?alt=startover&amp;nextID=<% =nextID%>&amp;startID=<% =s_id %>&amp;todelete=<% IF len(todelete) > 1 THEN response.write  left(todelete,len(todelete)-2)%>" class="box_link" style="padding-left:15px;">START OVER</a></div>
                                        </div>
                                    </div>
                                    <div style="float:right;margin-top:25px;margin-right:20px;">You did not answer all questions correctly. To proceed start over.</div>
                                    <div class="clear"></div>
                                <% END IF%>
                                <% END IF%>
                            <% ELSE%>
                            
                            <div class="next_blue" style="text-align:left;margin:20px 0px;">
                                <div class="h_submit_blue"><a href="t_question.asp?nextID=<% =nextID%>" class="box_link" style="padding-left:15px;">NEXT</a></div>
                            </div>
                            <% END IF%>
                            
                            
                        </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="clear"></div>
                </div>
            </div>
        </div>
        <!-- #include file = "partials/footer.asp" -->
    </div>
</body>
<%
call log_the_page("Training and quiz", Session("ID"), "Feedback", (subject("s_id")), (subject("s_title")), 0, qst, "Feedback")
    
%>

<% subject.close : Set subject = Nothing%>
<!-- #include file = "errorhandler/index.asp"-->