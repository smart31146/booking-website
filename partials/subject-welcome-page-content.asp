<div class="subject-welcome-page">
    <div class="subject-welcome-page__content">
        <div class="subject-welcome-page__left">
            <div class="text__section">
                <% if act_subj.EOF or act_subj.BOF then Response.Redirect("error.asp?" & request.QueryString) %>
                <% If user_check.EOF Then %>
                <span class="page-title">Welcome</span>
                <div class="page-text">
                    <% if selected_topic_query <> "" then %>
                    <p>There are <strong><%=ctopicCount_total%></strong> topics on this subject, out of them you're required to complete <strong><%=currentTopicCount%></strong> topics</p>
                    <%
                        dim topics_to_complete
                        topics_to_complete = replace(selected_topic_query,"s2.s_topic = ","")
                        topics_to_complete = replace(topics_to_complete,"'","")
                        topics_to_complete = replace(topics_to_complete," OR ",", ")
                    %>
                    <p>Topics to complete: <strong><%=topics_to_complete%></strong></p>
                    <% else %>
                    <p>There are <strong><%=currentTopicCount%></strong> topics on this subject.</p>

                    <% end if %>
                    <p>You will be asked quiz questions on what you have learned. The pass mark is <strong><%=act_subj("subject_passmark")%>%.</strong></p>
                    <p>You can check your progress in the progress bar in the top right hand corner.</p>
                    <p>You can exit the training or quiz at any time. When you return, you will be given the choice to continue on or to start again.</p>
                </div>

                <a class="btn btn-info default-blue-btn start-btn" role="button" href="../t_ses_create.asp?ID_subject_prm=<%=ID_subject_prm%>&nextID=<%=sStart%>&total=<%=total_counter%>">Start</a>

                <%
                else
                Session("SessionID") = cStr(user_check("ID_session"))
                Dim currentPage
                Dim questionsTotal
                currentPage=user_check("Session_stop")
                questionsTotal=user_check("Session_total")
                %>

                <span class="page-title">You have a choice</span>
                <div class="page-text">
                    <%
                    '--------------------------------------------------------------------------- RS 14/08/15 ----------------------------------------
                    'The following code solves the issue of the bookmarking, where the user is brought back to the same question if they answered it and exited the training
                    '----------------------------------------------------------------------------------------------------------------------------------
                    'Checks to see if the question has been answered already by checking for available records
                    if (Session("sessionID") <> "") and (Session("question_ID") <> "") then
                    set answeredQuestion = Server.CreateObject("ADODB.Recordset")

                    if selected_topic_query <> "" then
                        SQL ="select distinct(ID_result) from q_result as res join q_question as ques on ques.ID_question=res.result_question Join new_subjects as s2 on s2.s_id="&user_check("session_current")&" AND ( "&cstr(selected_topic_query)&" ) where  res.result_session="&user_check("ID_session")&" and res.result_question="&Session("question_ID")&""
                    else
                        SQL ="select distinct(ID_result) from q_result as res join q_question as ques on ques.ID_question=res.result_question Join new_subjects as subj on subj.s_id="&user_check("session_current")&" where  res.result_session="&user_check("ID_session")&" and res.result_question="&Session("question_ID")&""
                    end if

                    answeredQuestion.Open SQL, Connect, 3, 1

                    'if user answered the question create a session variable
                    If Not answeredQuestion.EOF Then
                    Session("answered")=1
                    else
                    Session("answered")=0
                    end if
                    end if

                    'set bookmark to the next page/question if the question has been answered
                    set NextQuestion = Server.CreateObject("ADODB.Recordset")
                    if selected_topic_query <> "" then
                        SQL= "select Top 1 * from new_subjects s2 where s2.s_order >"&orderOn&" AND ( "&cstr(selected_topic_query)&" ) and s2.s_active=1 AND ( "&cstr(selected_topic_query)&" ) and s2.s_qid="&fixstr(clng(ID_subject_prm))& "order by s_order"
                    else
                        SQL= "select Top 1 * from new_subjects where s_order >"&orderOn&" and s_active=1 and s_qid="&fixstr(clng(ID_subject_prm))& "order by s_order"
                    end if

                    NextQuestion.Open SQL, Connect, 3, 1

                    set CurrentQuestion = Server.CreateObject("ADODB.Recordset")
                    if selected_topic_query <> "" then
                        SQL = "select * from new_subjects s2 where s2.s_order ="&orderOn&" AND ( "&cstr(selected_topic_query)&" ) AND s2.s_active=1 and s2.s_qid="&fixstr(clng(ID_subject_prm))
                    else
                        SQL = "select * from new_subjects where s_order ="&orderOn&" and s_active=1 and s_qid="&fixstr(clng(ID_subject_prm))
                    end if

                    CurrentQuestion.Open SQL, Connect, 3, 1


                    If Not NextQuestion.EOF Then
                    'if current bookmark is a question and question has been answered, go to next question or page
                    if CurrentQuestion("s_typ")=2 and Session("answered")=1 then
                    currentQuestion=NextQuestion("s_id")
                    'if next page is not a question and current page is a question and has been answered, go to the next question or page
                    else if NextQuestion("s_typ")=1 and CurrentQuestion("s_typ")=2 and Session("answered")=1 then
                    currentQuestion=NextQuestion("s_id")
                    'if next page is a question and current page is a question and the current question has been answered, go to the next question or page
                    else if NextQuestion("s_typ")=2 and CurrentQuestion("s_typ")=2 and Session("answered")=1 then
                    currentQuestion=NextQuestion("s_id")
                    'if above doesn't apply proceed as normal
                    else
                    currentQuestion=user_check("Session_current")
                    end if
                    end if
                    end if
                    else
                    'set nextID link variable
                    currentQuestion=user_check("Session_current")
                    end if
                    NextQuestion.Close

                    '--------------------------------------------------------------------------- END ----------------------------------------
                    '----------------------------------------------------------------------------------------------------------------------------------
                    '----------------------------------------------------------------------------------------------------------------------------------

                    %>
                    <p>When you last attempted <b><%=ReplaceStrQuiz(act_subj("subject_name"))%></b> on <b><%=cStr(user_check("Session_date")) %></b>, you did not finish the subject.</p>

                    <p>You were on Topic <strong><%=topicPosition%></strong> of <strong><%=currentTopicCount%></strong>, <strong><%=currentTopic%></strong></p>

                    <p>Click 'submit' to continue with your previous attempt.</p>
                    <p>If you wish to start from the beginning again, please click the check-box below before pressing 'submit'. This will discard your previous quiz.</p>
                </div>

                <form class="form-check discard-checkbox-form" name="discard" id="discard" method="post" action="t_ses_get.asp?nextID=<%=currentQuestion%>">
                    <div class="d-flex">
                        <input class="form-check-input discard-checkbox-form__checkbox" type="checkbox" name="discard" value="TRUE" id="flexCheckDefault">
                        <label class="form-check-label discard-checkbox-form__label" id="lbl_discard" for="flexCheckDefault">
                            Discard my earlier answers. I want to start again.
                        </label>
                    </div>
                    <button type="submit" class="btn btn-info default-blue-btn discard-checkbox-form__submit-btn">Submit</button>
                </form>
                <% end if%>
            </div>
        </div>
        <div class="picture-section">
            <img class="picture" src="../vault_image/images/training.jpg" alt="training.jpg">
        </div>
    </div>
</div>

