<%@LANGUAGE="VBSCRIPT"%>
<!-- #include file = "connections/bbg_conn.asp" -->
<!-- #include file = "connections/include.asp"-->
<%

Session.Contents.RemoveAll()

cid = CStr(Request.QueryString("cid"))

uid = CStr(Request.QueryString("uid"))

Session("uid") = uid

originurl = CStr(Request.QueryString("callbackurl"))

	
If (Len(originurl) > 8 And InStr(originurl, "connect.html") <> 0) Then
	
	Session("callbackurl") = Left(originurl,(InStr(originurl, "connect.html")-1)) + "return.html?course=" + cid
	
Else

	response.redirect("error.asp")

End If

'Does the user exist? If not, add them

	If Len(uid) > 0 Then 

		set rsUser = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * from q_user WHERE user_username = '" & uid & "'"
		rsUser.Open SQL, Connect,3,3
		SQL = ""

		If (rsUser.EOF Or rsUser.BOF) Then

			SQL = "INSERT INTO q_user (user_lastname, user_firstname, user_username) VALUES ('USER', 'LMS', '" & uid & "')"
			Set useraddCmd = Server.CreateObject("ADODB.Recordset")
			useraddCmd.Open SQL, Connect,3,3
			Set useraddCmd = nothing
			
			'Refresh page to requery with the added reference and the same querystring to that we can get back the BBP ID_user
			Response.redirect("autolog.asp?" & request.querystring)

		End If
		
		'Now we have a valid user, record the user ID and name

		user_id = rsUser.Fields.Item("ID_user").Value
		last_name = rsUser.Fields.Item("user_lastname").Value
		first_name = rsUser.Fields.Item("user_firstname").Value

        'Insert subject for user

        set rsAssignSubjectToUser = Server.CreateObject("ADODB.Recordset")
        SQLCheckIfSubjectAdded = "SELECT * from subject_user WHERE ID_subject = '"&cid&"' AND ID_user = '"&user_id&"'"
        rsAssignSubjectToUser.Open SQLCheckIfSubjectAdded, Connect,3,3
        Set SQLCheckIfSubjectAdded = nothing

        If (rsAssignSubjectToUser.EOF OR rsAssignSubjectToUser.BOF) Then
            InsertUserSubjectQuery = "INSERT INTO subject_user (ID_subject, ID_user) VALUES ('"&cid&"','"&user_id&"')"
            Set InsertUserSubjectQueryCmd = Server.CreateObject("ADODB.Recordset")
            InsertUserSubjectQueryCmd.Open InsertUserSubjectQuery, Connect,3,3
            Set InsertUserSubjectQuery = nothing
        End If
		Set rsUser = nothing

	Else
	
		response.redirect("error.asp")
	
	End If

Dim erq_subject_id
Session("UserID") = user_id
Session("id") = cid
Session("LMS") = 1
Session("name") = ""

SQL = "SELECT id_subject,subject_name,subject_erq FROM subjects WHERE id_subject=" & cid
obj.Open SQL, Connect,3,3
Session("name") = obj("subject_name")

if obj("subject_erq") > 0 then
    erq_subject_id = obj("subject_erq")
end if

obj.close

if erq_subject_id > 0 then
set erq_obj = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT id_subject,subject_name,subject_erq FROM subjects WHERE id_subject=" & erq_subject_id
erq_obj.Open SQL, Connect,3,3
' Session("name") = erq_obj("subject_name")

erq_obj.close

end if


if (erq_subject_id > 0) then 

'####################################################

' Check the latest session for this user
	dim sessionId
	dim sessionPassed	
	dim sessionErq

set rsSessionDetails = Server.CreateObject("ADODB.Recordset")
'sql="SELECT TOP 1 qc.q_session,qc.passed,qc.isErq FROM q_certification qc INNER JOIN q_session qs ON qs.ID_session=qc.q_session WHERE qs.session_subject=11 AND qs.session_users=2242 AND qc.passed=1 ORDER BY qc.id desc"
rsSessionDetails_query="SELECT TOP 1 qc.q_session,qc.passed,qc.isErq FROM q_certification as qc "&_
"INNER JOIN q_session as qs ON qc.q_session=qs.ID_session "&_
"WHERE qs.Session_subject='"&cid&"' AND qs.Session_users='"&user_id&"' ORDER BY qc.id desc"

rsSessionDetails.open rsSessionDetails_query, Connect,3,1

If (rsSessionDetails.EOF Or rsSessionDetails.BOF) Then
	
	rsSessionDetails.close
	Session("erq_session") = 0
	Session("id") = cid
	Response.redirect("t_index.asp?ID_subject_prm=" + cid)

else
'rsSessionDetails.open sql,3,3
	sessionId=rsSessionDetails.fields(0)
	sessionPassed=rsSessionDetails.fields(1)
	sessionErq=rsSessionDetails.fields(2)
	rsSessionDetails.close
	
	set rsSessionDetails_passed = Server.CreateObject("ADODB.Recordset")
	
	rsSessionDetails_query="SELECT TOP 1 qc.q_session,qc.passed,qc.isErq FROM q_certification as qc "&_
						   "INNER JOIN q_session as qs ON qc.q_session=qs.ID_session "&_
						   "WHERE qs.Session_subject='"&cid&"' AND qs.Session_users='"&user_id&"' ORDER BY qc.id desc"
	
	rsSessionDetails_passed.open rsSessionDetails_query, Connect,3,1
	
	sessionId_passed = rsSessionDetails_passed.fields(0)
	sessionErq_passed = rsSessionDetails_passed.fields(1)
    sessionIsErq_passed = rsSessionDetails_passed.fields(2)
	rsSessionDetails_passed.close

	if sessionId = sessionId_passed AND sessionErq_passed = 1 then

        'if sessionIsErq_passed = 1 then

            'Session("erq_session") = 0
		    'Session("id") = cid
		    'Response.redirect("t_index.asp?test=148")

        'else
		    Session("id") = erq_subject_id
		    Session("original_subject") = cid
		    Session("erq_session") = 1
		    Response.redirect("t_index.asp?test=154")
        'end if

	else
		dim selected_topics_conn,selected_topics_rs,selected_topics_query
		dim subjects_array()

        if sessionIsErq_passed = 1 then

            Set selected_topics_conn = Server.CreateObject("ADODB.Recordset")
            selected_topics_rs = "SELECT DISTINCT new_subjects.s_topic AS topic_name FROM new_subjects INNER JOIN ((q_result INNER JOIN q_question ON q_result.result_question = q_question.ID_question) INNER JOIN q_choice ON (q_question.ID_question = q_choice.choice_question) AND (q_result.result_answer = q_choice.ID_choice)) ON new_subjects.s_ID = q_question.question_topic WHERE result_session = '"&sessionId_passed&"' AND q_choice.choice_cor = 0"
            selected_topics_conn.Open selected_topics_rs, Connect,3,3
            
            
			redim subjects_array(-1)
			
			
            if not selected_topics_conn.EOF then 
                do until selected_topics_conn.EOF
				ReDim Preserve subjects_array(UBound(subjects_array) + 1)
				subjects_array(UBound(subjects_array)) = trim(cstr(selected_topics_conn.Fields.Item("topic_name").Value))
				selected_topics_conn.movenext
                loop
            end if
			
			selected_topic_query = ""
			
			For k=0 To UBound(subjects_array)
				selected_topic_query = selected_topic_query & "s2.s_topic = " & "'" & cstr(subjects_array(k)) & "' OR "
			Next
			
			selected_topic_query = selected_topic_query & "remove"
			selected_topic_query = replace(cstr(selected_topic_query),"OR remove","")
			
			
			Session("selected_topics") = selected_topic_query
            'Session("selected_topics") =  "s2.s_topic = " & "'" & "Introduction" & "'"
            
            selected_topics_conn.Close
            Set selected_topics_conn = nothing
			
			'Session("id") = erq_subject_id
		    'Session("original_subject") = cid
		    'Session("erq_session") = 1
		    'Response.redirect("t_index.asp?test=185&selected_topic_query=" & selected_topic_query)
			
			Session("id") = cid
		    Session("erq_session") = 0
			Response.redirect("t_index.asp?test=204")
			
			else 
            

            Set selected_topics_conn = Server.CreateObject("ADODB.Recordset")
            selected_topics_rs = "SELECT DISTINCT new_subjects.s_topic AS topic_name FROM new_subjects INNER JOIN ((q_result INNER JOIN q_question ON q_result.result_question = q_question.ID_question) INNER JOIN q_choice ON (q_question.ID_question = q_choice.choice_question) AND (q_result.result_answer = q_choice.ID_choice)) ON new_subjects.s_ID = q_question.question_topic WHERE result_session = '"&sessionId_passed&"' AND q_choice.choice_cor = 0"
            selected_topics_conn.Open selected_topics_rs, Connect,3,3
            
			redim subjects_array(-1)
			
			
            if not selected_topics_conn.EOF then 
                do until selected_topics_conn.EOF
				ReDim Preserve subjects_array(UBound(subjects_array) + 1)
				subjects_array(UBound(subjects_array)) = trim(cstr(selected_topics_conn.Fields.Item("topic_name").Value))
				selected_topics_conn.movenext
                loop
            end if
			
			selected_topic_query = ""
			
			For k=0 To UBound(subjects_array)
				selected_topic_query = selected_topic_query & "s2.s_topic = " & "'" & cstr(subjects_array(k)) & "' OR "
			Next
			
			selected_topic_query = selected_topic_query & "remove"
			selected_topic_query = replace(cstr(selected_topic_query),"OR remove","")
			
			
			Session("selected_topics") = selected_topic_query
            'Session("selected_topics") =  "s2.s_topic = " & "'" & "Introduction" & "'"
            
            selected_topics_conn.Close
            Set selected_topics_conn = nothing
			
			Session("id") = cid
		    Session("erq_session") = 0
		    Response.redirect("t_index.asp?ID_subject_prm=" & cid & "test=238&selected_topic_query=" & selected_topic_query)
        end if 		
	end if
end if

else 
	Session("id") = cid
	Session("erq_session") = 0
	Response.redirect("t_index.asp?ID_subject_prm=" + cid)
end if
%>