<div class="certificate-page">
    <div id="ele1" class="certificate-page__content d-flex">
        <!-- this container is added in print page mode only -->
        <div class='print-heading'>
            <img class='print-heading__logo' src='images/logo_certficate.jpg' />
            <h4>BETTER BUSINESS PROGRAM CERTIFICATE</h4>
            <h3> <% =subject("subject_name")%></h3>
        </div>

        <div class="certificate-page__left">
            <div class="completed-date">
                <span>
                    Date Completed: <%= FormatDateTime(Date, 1) %> at <%= FormatDateTime(Now, 3) %>
                </span>
            </div>

            <table class="table table-striped score-table">
                <thead>
                    <tr>
                        <td>Topic</td>
                        <td>Score</td>
                        <td>&nbsp;</td>
                    </tr>
                </thead>
            <%
            xi = 0
            While (NOT results.EOF)
                count_of_correct = 0
                xi = xi + 1
                IF (results("qanswer")) = True THEN
                    image = "<span class=""result pass""><i class=""fa-solid fa-check""></i> Pass</span>"
                    Total_correct = Total_correct+1
                ELSE
                    image = "<span class=""result fail""><i class=""fa-solid fa-xmark""></i> Incorrect</span>"
                END IF
                %>
                <tr>
                    <td>Quiz <% =xi%></td>
                    <td><% =results("s_topic")%></td>
                    <td><%=image%> </td>
                </tr>
                
            <%
            results.MoveNext()
            Wend
            %>
            <%
            session_completed = 1
            if Session("LMS") = 1 then
            session_completed = 0
            end if

            if (Session("erq_session") = 1) then

            MM_editConnection = Connect
            MM_editTable = "q_session"
            MM_editQuery = "update " & MM_editTable & " set Session_subject= "&original_subject_id&",session_done = 1, session_correct = " & total_correct & ", session_finish = '" & cDateSql(Now())&"' " & "where ID_session = " & SessionID
            'Response.Write MM_editQuery
            Set MM_editCmd = Server.CreateObject("ADODB.Command")
            MM_editCmd.ActiveConnection = MM_editConnection
            MM_editCmd.CommandText = MM_editQuery
            MM_editCmd.Execute
            MM_editCmd.ActiveConnection.Close

            else

            MM_editConnection = Connect
            MM_editTable = "q_session"
            MM_editQuery = "update " & MM_editTable & " set session_done = 1, session_correct = " & total_correct & ", session_finish = '" & cDateSql(Now())&"' " & "where ID_session = " & SessionID
            'Response.Write MM_editQuery
            Set MM_editCmd = Server.CreateObject("ADODB.Command")
            MM_editCmd.ActiveConnection = MM_editConnection
            MM_editCmd.CommandText = MM_editQuery
            MM_editCmd.Execute
            MM_editCmd.ActiveConnection.Close

            end if

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

            if (Session("erq_session") = 1) then 

                IF Session("qsave") = "" THEN
                    'pn 060127 add a certification for this person, can be 0 for failed or 1 for passed
                    Set MM_SaveCertification = Server.CreateObject("ADODB.Command")
                    MM_SaveCertification.ActiveConnection = MM_editConnection
                    MM_SaveCertification.CommandText   = "insert into q_certification (q_session, quiz_date, expiry_date, passed, percentage_achieved, percentage_required, isErq) values ('" & SessionID & "' ,'" & cDateSql(Now())&"',DATEADD (week , "&add_to_expiry_date&", GETDATE()),"&pass_or_fail&","&percentage_achieved&","&subject_passmark&",1);"
                    MM_SaveCertification.Execute
                    MM_SaveCertification.ActiveConnection.Close
                END IF

            else 
                IF Session("qsave") = "" THEN
                'pn 060127 add a certification for this person, can be 0 for failed or 1 for passed
                    Set MM_SaveCertification = Server.CreateObject("ADODB.Command")
                    MM_SaveCertification.ActiveConnection = MM_editConnection
                    MM_SaveCertification.CommandText   = "insert into q_certification (q_session, quiz_date, expiry_date, passed, percentage_achieved, percentage_required) values ('" & SessionID & "' ,'" & cDateSql(Now())&"',DATEADD (week , "&add_to_expiry_date&", GETDATE()),"&pass_or_fail&","&percentage_achieved&","&subject_passmark&" );"
                    MM_SaveCertification.Execute
                    MM_SaveCertification.ActiveConnection.Close
                END IF
            end if
            Session("qsave") = "yesbox"
            %>
            <!-- <tr>
                <td >&nbsp;</td>
                <td >&nbsp;</td>
                <td >&nbsp;</td>
            </tr>
            <tr>
                <td ><b>Total</b></td>
                <td >&nbsp;</td>
                <td ><b><%=total_correct%>/<%=xi%></b></td>
            </tr> -->
            </table>
        </div>

        <div class="certificate-page__right">
            <div class="certificate-page__right__score">
                <span class="score-label">Your score for this subject </span>
                <div class="percentage">
                    <% if(pass_or_fail=1) then %>
                        <i class="fa-solid fa-circle-check pass"></i>
                        <span class="percent pass"><%=FormatNumber(percentage_achieved,2)%>%</span>
                        <span class="mark pass">Pass</span>
                    <% else %>
                        <span class="percent fail"><%=FormatNumber(percentage_achieved,2)%>%</span>
                        <span class="mark fail">Below pass</span>
                    <% end if%>
                </div>
                <span class="score-label tiny">The pass mark for this subject is <b><%=subject_passmark%>%</b> of all questions regardless of topic.</span>
                <div class="hr separator"></div>
            </div>

            <div>             
                <% if(pass_or_fail=1) then %>
                    <span class="page-text">
                        Congratulations!
                        You successfully completed your training for <b><% =subject("subject_name")%></b>.
                    </span>
                <% else %>
                    <span class="page-text">
                        You must complete the full training for those topics. You will not be shown as having completed your compliance training for this subject until you have passed all topics.
                    </span>
                <% end if%>
            </div>

            <div class="certificate-page__right__buttons d-flex">
                <% if Session("LMS") = 1 then %>
                    <div class="no-add-to-print-mode">
                        <span class="page-text">
                            Otherwise, you will be returned to the LMS automatically in <span id="countdown" style="font-size:16px;color:red;">10</span> <span id="second">seconds</span>.
                        </span>
                        <div class="LMS-btns">
                            <% if(pass_or_fail=1) then %>
                                <a role="button" class="btn btn-info default-blue-btn" href="lmscallback.asp?SessionBBP=<%=SessionID%>">Return to LMS</a>
                            <% else
                                dim go_back_to_start_Session
                                if (Session("erq_session") = 1) then
                                    go_back_to_start_Session = "autolog.asp?cid=" & cstr(original_subject_id) & "uid=" & cstr(uid) & "callbackurl=" & cstr(Session("callbackurl"))
                                else
                                    go_back_to_start_Session = "autolog.asp?cid=" & cstr(Session("id")) & "uid=" & cstr(uid) & "callbackurl=" & cstr(Session("callbackurl"))
                                end if%>
                                <a role="button" class="btn btn-info default-blue-btn" href="go_back_to_start_Session">Start this session again</a>
                            <% end if %>

                            <a href="javascript:void(0)" role="button" class="btn btn-secondary default-blue-btn print-link print-btn">
                                <i class="fa-solid fa-print"></i>
                                <span>Print a certificate</span>
                            </a>
                        </div>          
                    </div>
                <% else %>
                    <div class="certificate_buttons no-add-to-print-mode">
                        <a href="javascript:void(0)" role="button" class="btn btn-secondary default-blue-btn print-link print-btn">
                            <i class="fa-solid fa-print"></i>
                            <span>Print a certificate</span>
                        </a>
                        <span class="page-text">or</span>
                        <a href="<%= homeURL&"?alt=change"%>" role="button" class="btn btn-info default-blue-btn">HomePage</a>

                        <!-- <a role="button" class="btn btn-info default-blue-btn" href="lmscallback.asp?SessionBBP=<%=SessionID%>">Return to LMS</a> -->
                    </div>
                <% END IF%>
            </div>
        </div>
    </div>
</div>