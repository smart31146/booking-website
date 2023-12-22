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
Dim subject, current_topic, total_completed_questions, total_correct_questions, total_questions, ordered_topic_list, next_question_id, SQL, conn
Dim totalTopics, topicProgressCount, lastTopic, currentQuestionFound, movedPastTopic, sID, remaining_questions, topic_name
If Request.QueryString("nextID")<>"" Then
    next_question_id = CInt(request.querystring("nextID"))
Else
    'Response.Redirect("error.asp?" & request.QueryString)
    next_question_id = 824
End If
if (Session("topic_name") <> "") Then 
    topic_name = Session("topic_name")
End If
if (Session("sessionID") <> "") Then 
    sID = Session("sessionID")
Else
    Response.Redirect("error.asp?" & request.QueryString)
End If

Dim selected_topic_query
if (cstr(Session("selected_topics")) <> "") Then
    selected_topic_query = cstr(Session("selected_topics"))
End If

set ordered_topic_list = Server.CreateObject("System.Collections.ArrayList")
set conn = Server.CreateObject("ADODB.Connection")
conn.open Connect

if selected_topic_query <> "" then
	SQL = "select s_id, s_topic, s_order, s_typ, s_qID, subject_name from new_subjects s2 inner join subjects on ID_subject = s_qID where s_active = 1 AND ( "&cstr(selected_topic_query)&" ) and s_qID = (select s_qID from new_subjects s2 where s2.s_ID = " & next_question_id & " AND ( "&cstr(selected_topic_query)&" )) order by s_order"
else
	SQL = "select s_id, s_topic, s_order, s_typ, s_qID, subject_name from new_subjects inner join subjects on ID_subject = s_qID where s_active = 1 and s_qID = (select s_qID from new_subjects where s_ID = " & next_question_id & ") order by s_order"
end if

set subjectRecords = conn.Execute(SQL)
currentQuestionFound = False
total_completed_questions = 0
topicProgressCount = 0
If Not subjectRecords.EOF Then
    Do Until subjectRecords.EOF
        subject = subjectRecords("subject_name")
		subjectID=subjectRecords("s_qID")
		topicID=subjectRecords("s_id")
        If lastTopic <> subjectRecords("s_topic") Then
            lastTopic = subjectRecords("s_topic")
            ordered_topic_list.Add lastTopic
            If Not currentQuestionFound Then
                topicProgressCount = topicProgressCount + 1
            End If
        End If
        If subjectRecords("s_id") <> next_question_id Then
            If subjectRecords("s_typ") = 2 Then
                total_questions = total_questions + 1
                If Not currentQuestionFound Then
                    total_completed_questions = total_completed_questions + 1
                End If
            End If
        Else
            currentQuestionFound = True
            current_topic = lastTopic
        End If
        subjectRecords.MoveNext
    Loop
End If
remaining_questions = total_questions - total_completed_questions
totalTopics = ordered_topic_list.Count
'subtract one since the topic progress count includes the topic about to commence based on the question id
topicProgressCount = topicProgressCount - 1

'determine the number of correct answers
SQL = "select count(*) as correct from q_result inner join q_choice on q_result.result_answer = q_choice.ID_choice where result_session = " & sID & " and choice_cor = 1"
set correctCount = conn.Execute(SQL)
If Not correctCount.BOF Then
    total_correct_questions = correctCount("correct")
End If
conn.Close
%>

<!doctype html>
<html>
<head>
    <meta id="viewport" name='viewport'>
    <link href="perfect-scrollbar-0.4.8/src/perfect-scrollbar.css" rel="stylesheet">
    <script src="jquery-1.11.1.js?v=bbp34"></script>
    <script src="perfect-scrollbar-0.4.8/src/jquery.mousewheel.js?v=bbp34"></script>
    <script src="perfect-scrollbar-0.4.8/src/perfect-scrollbar.js?v=bbp34"></script>
    <link rel="stylesheet" type="text/css" href="js/sweet-alert.css">
    <script src="js/sweet-alert.min.js?v=bbp34"></script>
    <script>
        (function(doc) {
            var viewport = document.getElementById('viewport');
            if ( navigator.userAgent.match(/iPhone/i) || navigator.userAgent.match(/iPod/i)) {
                viewport.setAttribute("content", "initial-scale=0.3");
            } else if ( navigator.userAgent.match(/iPad/i) ) {
                viewport.setAttribute("content", "initial-scale=1.05");
            }
        }(document));
        
        $(document).ready(function() {
            $("html, body").animate({ scrollTop: $(document).height() }, 10);
        });
    </script>
    <title>ACME - Better Business Program</title>
    <META name="DESCRIPTION" content="">
    <META name="ROBOTS" content="index,follow">
    <META name="LANGUAGE" content="SE">
    <meta name="keywords" CONTENT="">
    <link href="style/bbp_acme34.css" rel="stylesheet" type="text/css">
    <!-- #include file = "inc_header.asp" -->

    <style media="all" type="text/css">
        img#bg {
          position:fixed;
          top:12000;
          left:0;
          width:100%;
          height:100%;
        
        }
        </style>
        <!--[if IE 6]>
        <style type="text/css">
        html { overflow-y: hidden; }
        body { overflow-y: auto; }
        img#bg { position:absolute; z-index:-1; }
        #content { position:static; }
        </style>
        <![endif]-->
		
		
		<!--[if IE 7]>
            <link href="style/bbp_ie7_acme34.css" rel="stylesheet" type="text/css">
        <![endif]-->
		
    <script >
        history.forward();
    </script>
</head>
<body>
    <div class="page-content">
        <div class="white-container">
            <!-- #include file = "partials/header.asp" -->

            <div class="allcontent">
                <div class="allcontent_main">
                    <div class="allcontent3">
    <div class="allcontent_main3">
            <div class="header_blue2">
                 <div class="header_progress" >
                    <div class="topic_progress">Topic <div class="topicNumberCircle"><%=topicProgressCount %></div> of <%=totalTopics %></div>
                    <div class="page_progress">Topic Review</div>		
                </div>
          <div class="header_inside"><%=subject %><br>
                <h3><%=topic_name %></h3>                                                                                                                                           
                </div>
            </div>
        
            <div class="allcontent_content_dynamic_height">
                     <div class="main_content" >
                    <div class="guide_blue">
					<div class="box_inside2"  >
                       <h3>Your <%=subject %> progress</h3>
                        <div class="box_text_blue"  >
						<div  style="height:287px;">
                            <p>
                                Well done! You have now completed the ticked topics in this subject.
                            <br />
                           <br>
                                Click on the 'Next' button below to proceed, or, click 'Exit' to stop and come back later.
                            </p>
                            <%
                            Dim colLength
                            colLength = ordered_topic_list.Count \ 2
                            'correct for uneven lists so that the first column is longer
                            If ordered_topic_list.Count Mod 2 = 1 Then
                                colLength = colLength + 1
                            End If
                                'write two at a time - the first is in the first column and the second is to it's right in the second column
                            For i = 0 To (colLength - 1)
                                If i < topicProgressCount Then
                                    response.write "<div class='t_topic completed'>" & ordered_topic_list.Item(i) & "</div>"
                                Else
                                    response.write "<div class='t_topic incomplete'>" & ordered_topic_list.Item(i) & "</div>"
                                End If
                                'ensure we don't go past the last index if the columns are uneven
                                If (i + colLength) <= (ordered_topic_list.Count - 1) Then
                                    If (i + colLength) < topicProgressCount Then
                                        response.write "<div class='t_topic completed'>" & ordered_topic_list.Item(i + colLength) & "</div>"
                                    Else
										response.write "<div class='t_topic incomplete'>" & ordered_topic_list.Item(i + colLength) & "</div>"
                                    End if
                                End If
                            Next
                            %>
                                <div class="clear"></div>
								</div>
                            <div class="div_button" >
                                <div class="up_blue" style="width:180px; position:relative;top:-3px;">
                                    <div class="h_submit_blue"><a href="index.asp?alt=logout" class="box_link" style="padding-left:40px;">SAVE & EXIT</a></div>
                                </div>
                                   <div class="next_blue" style="position:relative;top:-3px;">
                                    <div class="h_submit_blue" >
                                        <a href="t_question.asp?nextID=<%=next_question_id %>&amp;topicReviewed=<%=topicProgressCount %>" class="box_link" style="padding-left:15px;">NEXT</a>
                                    </div>
                            </div>
                        </div>
						
                            </div>
                        </div>
                    </div>
                   
            </div>
                <div class="t_progress_info">
                <% if total_completed_questions <> 0 then %>
                    <div class="t_progress_info_inside">
                        <h3 class="white_text">Your quiz score is</h3>
                        <br />
                        <div class="quiz_score white_text"><span class="quizScoreCircle"><%=total_correct_questions %></span> of <%=total_completed_questions %></div>
                        <%
                            If remaining_questions <> 0 Then
                                response.write "<div class='white_text' id='questionsRemaining'><strong>You have " & remaining_questions & " questions remaining!</strong></div>"
                            Else
                                response.write "<div class='white_text' id='questionsRemaining'><strong>All quiz questions completed!</strong></div>"
                            End If
                        %>
                    </div>
                    <% end if %>
                    <div class="white_text" style="margin:60px 0 20px 20px"><strong>Your next topic is</strong></div>
                    <div class="t_topic incomplete" style="margin:0 0 0 20px"><%=current_topic %></div>
                </div>
                <div class="clear"></div>
                    </div>
                    <div class="clear"></div>
                </div>
        
            </div>
	    </div>
	</div>
</div>
<!-- #include file = "partials/footer.asp" -->
</div>

<% call log_the_page("Training and quiz", subjectID, subject , topicID, current_topic , 0, 0, "Topic Review") %>

</body>
<!-- #include file = "errorhandler/index.asp"-->