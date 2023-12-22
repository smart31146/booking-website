<%@LANGUAGE="VBSCRIPT"%>

<!-- #include file = "connections/bbg_conn.asp" -->
<!-- #include file = "connections/include.asp"-->

<%
if NOT pref_quiz_avail then response.redirect("error.asp?" & request.QueryString)

Dim Total_answered
Dim Total_correct
Total_answered = 0
Total_correct = 0

Dim SessionBBP
if (request.querystring("SessionBBP") <> "") Then
	SessionBBP = clng(request.querystring("SessionBBP"))
Else
	Response.Redirect("error.asp?" & request.QueryString)
End If

Dim UserID
if (Session("UserID") <> "") Then
	UserID = clng(Session("UserID"))
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
if Err.Number = 0 then
set subject = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT *  FROM subjects WHERE ID_subject = "&fixstr(clng(sID))&" "
subject.Open SQL, Connect,3,3
end if

if Err.Number = 0 then
set results = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT (SELECT choice_cor FROM q_choice WHERE result_answer = ID_choice AND ABS([choice_active]) = 1) qanswer, s_topic FROM q_result,q_question,new_subjects WHERE question_topic = s_id AND result_question = ID_question AND  result_session = " &fixstr(clng(SessionBBP))& " ORDER BY id_result"
'response.write SQL & "<br>"
results.Open SQL, Connect,3,3
end if

if Err.Number = 0 then
set userdetails = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT q_user.ID_user, q_user.user_lastname, q_user.user_firstname FROM q_user WHERE q_user.ID_user=" &fixstr(clng(Session("userID")))& ""
userdetails.Open SQL, Connect,3,3
end if

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
	
  results.MoveNext()
Wend

if (Session("original_subject") <> "") then
    original_subject_id = Session("original_subject")
end if

  MM_editConnection = Connect
  MM_editTable = "q_session"

  if (Session("erq_session") = 1) then
  MM_editQuery = "update " & MM_editTable & " set session_subject = " & original_subject_id & " ,session_done = 1, session_total = " & xi & ", session_correct = " & total_correct & ", session_finish = '" & cDateSql(Now())&"' " & "where ID_session = " & SessionBBP
  else
  MM_editQuery = "update " & MM_editTable & " set session_done = 1, session_total = " & xi & ", session_correct = " & total_correct & ", session_finish = '" & cDateSql(Now())&"' " & "where ID_session = " & SessionBBP
  end if
  
  'Response.Write MM_editQuery
if Err.Number = 0 then
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

	percentage_achieved=Cint((total_correct/xi)*100)

	if(percentage_achieved>=subject_passmark) then
		pass_or_fail=1
		add_to_expiry_date=subject_expiry
	end if

subject.Close()
results.Close()
userdetails.Close()

If pass_or_fail Then

	status = 1

Else

	status = 2

End if

callbackurl = Session("callbackurl") & "&status=" & status & "&score=" & percentage_achieved
	
call log_the_page("Training and quiz", SessionBBP, "lmscallback", 0, "limscallback", 0, qst, "callbackurl: " & callbackurl)

Response.Redirect(callbackurl)

%>
