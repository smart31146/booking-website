<div class="question-page">
    <div class="question-page__content">
        <% ' s_typ 1 = Traning, 2 = Quiz 
        IF clng(subject("s_typ")) = 1 THEN%>
            <div class="question-page__left">
                <div class="text__section">
                    <a name="answer"></a>
                    <span class="question-title"><%=ReplaceStrQuiz(subject("s_title"))%></span>
                    <div class="page-text">
                        <p><%=ReplaceStrQuiz(subject("s_body"))%></p>
                        <% IF showQuestion = true THEN%>
                        <div class="d-flex t_choose__list">
                            <% If Ubound(QArr,2) > -1 Then
                                xi = 0
                                    For i=0 to ubound(QArr,2)
                                        xi=xi+1 %>
                                        <div class="t_choose" id="button<% =QArr(0,i)%>">
                                            <a class="t_choose__link" id="link<% =QArr(0,i)%>" href="#answer" onclick="javascript:showlayer('layer<% =QArr(0,i)%>','button<% =QArr(0,i)%>'); b_click=false;"><% =ReplaceStrQuiz(QArr(1,i))%></a>
                                        </div>
                                        <% IF xi=2 THEN
                                        xi=0
                                        response.write ""
                                        END IF%>
                                    <% Next
                            END IF%>
                        </div>
                        <% END IF%>
                        <% IF clng(subject("s_goback"))>0 THEN%>                        
                            <div class="certificate_blue">
                                <div class="h_submit_blue">
                                    <a class="box_link" style="padding-left:15px;" href="#" onclick="popup('popUpDiv','t_question_window.asp?s_id=<% =subject("s_goback")%>')">
                                        GO BACK AND SEE THIS SCENARIO AGAIN
                                    </a>
                                </div>
                            </div>
                        <% END IF%>
                    </div>
                </div>
                        
                <% ' If last page on traning & quiz
                IF clng(NextID) = 0 THEN%>
                    <div class="recording-result-section">
                        <div class="mb-3">
                            <img src="images/lyte_loading.gif" width="26" height="26" alt="" style="vertical-align:middle;"> Recording score...
                        </div>
                        <p>You have now reached the end of the quiz.</p>
                        <p>Please wait while your score is recorded for this course.</p>
                        <p>If you are not redirected automatically please click the 'Certificate' button.</p>
                        <script type="text/javascript">
                            window.setTimeout("document.location = 'certificate.asp';", 7000);
                        </script>
            

                        <a class="btn btn-info default-blue-btn certificate-btn" role="button" href="certificate.asp">
                            <span>Certificate</span>
                        </a>
                    </div>
                <% ELSE%>

                <div class="d-flex">
                    <% IF clng(prevID)<>0 And subject("s_order") - previousTopicEnd <> 1 THen%>
                    <a class="btn btn-primary back-btn" role="button" href="t_question.asp?nextID=<% =prevID%>&returning=1">
                        <i class="fa-solid fa-angle-left"></i>
                        <span>Back</span>
                    </a>
                    <% end if%>
                    <% IF showQuestion = true THEN%>
                    <a class="btn btn-primary next-btn" role="button" href="javascript:gotonextpage('t_question.asp?nextID=<% =nextID%>')" onClick="closing=false">
                        <span>Next</span>
                        <i class="fa-solid fa-angle-right"></i>
                    </a>
                    <% ELSE%>
                    <a class="btn btn-primary next-btn" role="button" href="t_question.asp?nextID=<% =nextID%>">
                        <span>Next</span>
                        <i class="fa-solid fa-angle-right"></i>
                    </a>
                    <% END IF%>
                </div>
                <% END IF %>

            </div>
            <div class="picture-section" id="layerfirst">
                <% IF subject("s_image")<>"" THEN%>
                <img class="picture" src="vault_image/images/<% =subject("s_image")%>" alt="">
                <% ELSE%>
                <img src="vault_image/images/training.jpg" width="320" height="390" alt="">
                <% END IF%>
            </div>

            <% IF showQuestion = true THEN%>
                <div class="blue-info-section" id="blue_inside" style="display:none;">
                    <%  For i=0 to ubound(QArr,2) %>
                        <div id="layer<% =QArr(0,i)%>" style="display:none;">
                            <div class="blue-info-section__text">
                                <% =ReplaceStrQuiz(QArr(2,i))%>
                            </div>
                        </div>
                    <%Next%>
                </div> 
            <% END IF%>

        <% ' s_typ 1 = Traning, 2 = Quiz 
        ELSEIF clng(subject("s_typ")) = 2 THEN%>
        <div class="question-page__left">
            <div class="quiz_section">
                <div class="quiz_section__content">
                    <div id="question_name"> 
                        <span class="quiz_section__question"><%=ReplaceStrQuiz(QArr(1,0))%></span>
                    </div>

                    <form class="quiz_section__form" name="quiz" method="POST" onsubmit="return trySubmit();" action="t_question.asp?currentID=<% =subject("s_id")%>&quiz=yes">
                        <% 
                        IF showQuestion = true THEN
                            For i=0 to ubound(QchoiceArr,2) %>
                                <label class="d-flex quiz_section__option" for="rbutton<% =QchoiceArr(0,i)%>">
                                    <input class="square-radio" type="radio" name="answer" value="<% =QchoiceArr(0,i)%>" onclick="empty = false" id='rbutton<% =QchoiceArr(0,i)%>'>
                                    <span class="quiz_section__letter"><% =ReplaceStrQuiz(QchoiceArr(1,i))%></span>
                                    <span class="quiz_section__value">
                                        <% =ReplaceStrQuiz(QchoiceArr(2,i))%>
                                    </span>
                                </label>
                            <%Next
                        END IF%>

                        <button type="submit" class="btn btn-info default-blue-btn submit-btn" name="btnSubmit" id="btnSubmit">Submit and continue</button>
                    </form>
                </div>
            </div>
        </div>
        <% END IF%>  
</div>
</div>