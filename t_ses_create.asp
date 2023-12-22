<%@LANGUAGE="VBSCRIPT"%>

<% 
'Response buffer is used to buffer the output page. That means if any database exception occurs the contents can be cleared without processed any script to browser
 Response.Buffer = True
 
' "On Error Resume Next" method allows page to move to the next script even if any error present on page whcich will be caught after processing all asp script on page
 On Error Resume Next
 
'Changed by PR on 25.02.16
%>

<% bbp_training = true%>
<!--#include file="connections/bbg_conn.asp" -->
<!--#include file="connections/include.asp" -->

<%
if NOT pref_quiz_avail then response.redirect("error.asp?" & request.QueryString)


Dim ID_subject_prm
If (Request.QueryString("ID_subject_prm") <> "") Then
	ID_subject_prm = cInt(Request.QueryString("ID_subject_prm"))
Else
	Response.Redirect("error.asp?" & request.QueryString)
End If

Dim selected_topic_query
if (cstr(Session("selected_topics")) <> "") Then
    selected_topic_query = cstr(Session("selected_topics"))
End If

'Total is wrong, counts all pages instead of just questions, see totalq code below
Dim total
If (Request.QueryString("total") <> "") Then
	total = cInt(Request.QueryString("total"))
Else
	Response.Redirect("error.asp?" & request.QueryString)
End If

Dim userid
if (Session("UserID") <> "") Then
	userid = cInt(Session("UserID"))
Else
	Response.Redirect("error.asp?" & request.QueryString)
End If

' *** Edit Operations: declare variables

MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""


' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then

  MM_editConnection = Connect
  MM_editTable = "q_session"
  MM_editRedirectUrl = "t_question.asp"
  MM_fieldsStr  = "user|value|subject|value|date|value|date|value|total|value|correct|value|stop|value|done|value"
  MM_columnsStr = "Session_users|none,none,NULL|Session_subject|none,none,NULL|Session_date|',none,NULL|Session_finish|',none,NULL|Session_total|none,none,NULL|Session_correct|none,none,NULL|Session_stop|none,none,NULL|Session_done|none,none,NULL"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")

  ' set the form values
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(i+1) = CStr(Request.Form(MM_fields(i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If

' *** Insert Record: construct a sql insert statement and execute it

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    FormVal = MM_fields(i+1)
    MM_typeArray = Split(MM_columns(i+1),",")
    Delim = MM_typeArray(0)
    If (Delim = "none") Then Delim = ""
    AltVal = MM_typeArray(1)
    If (AltVal = "none") Then AltVal = ""
    EmptyVal = MM_typeArray(2)
    If (EmptyVal = "none") Then EmptyVal = ""
    If (FormVal = "") Then
      FormVal = EmptyVal
    Else
      If (AltVal <> "") Then
        FormVal = AltVal
      ElseIf (Delim = "'") Then  ' escape quotes
        FormVal = "'" & Replace(FormVal,"'","''") & "'"
      Else
        FormVal = Delim + FormVal + Delim
      End If
    End If
    If (i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End if
    MM_tableValues = MM_tableValues & MM_columns(i)
    MM_dbValues = MM_dbValues & FormVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"
	'response.write MM_editQuery
  If (Not MM_abortEdit) Then
    ' execute the insert
if Err.Number = 0 then
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
end if


' get the current session ID
	set user_check = Server.CreateObject("ADODB.Recordset")
	user_check.ActiveConnection = Connect
if Err.Number = 0 then
	user_check.Source = "SELECT q_session.ID_session, q_session.Session_users, q_session.Session_stop, q_session.Session_finish  FROM q_session  WHERE (((q_session.Session_users)=" + Replace(userID, "'", "''") + ") AND ((q_session.Session_done)=0) AND ((q_session.Session_subject)=" + Replace(ID_subject_prm, "'", "''") + "))  ORDER BY q_session.Session_date DESC;"
	user_check.CursorType = 0
	user_check.CursorLocation = 3
	user_check.LockType = 3
	user_check.Open()
	user_check_numRows = 0
end if

	Session("sessionID")=cStr(user_check.Fields.Item("ID_session").Value)

	user_check.Close()

' set the time for time measuring
	session("question_time") = Now()

    If (MM_editRedirectUrl <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "&sid=" & Session("sessionID")
      Response.Redirect("t_question.asp")
    End If
  End If

End If

'Total varable was wrong, it counted all pages instead of questions. This has been corrected and replaced with the following code below.
if Err.Number = 0 then
set totalq = Server.CreateObject("ADODB.Recordset")

if selected_topic_query <> "" then
	SQL = "SELECT COUNT(*) as qorder FROM new_subjects s2,subjects WHERE s2.s_qiD = ID_subject AND ABS([s_active]) = 1 AND ( "&cstr(selected_topic_query)&" ) AND s_qID = "& fixstr(ID_subject_prm)&" AND s_typ = 2"
else
	SQL = "SELECT COUNT(*) as qorder FROM new_subjects s1,subjects WHERE s1.s_qiD = ID_subject AND ABS([s_active]) = 1 AND s_qID = "& fixstr(ID_subject_prm)&" AND s_typ = 2"
end if

totalq.Open SQL, Connect,3,3
end if 

Dim totalq
if (totalq("qorder") <> "") Then
	totalq = cInt(totalq("qorder"))
ELSE
	totalq = 99
end if



%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<title><%=client_name_short%> - Better Business Program</title>
		<META name="DESCRIPTION"	content="">
		<!-- #include file = "inc_header.asp" -->
</head>
<body onload="javascript:document.forms[0].submit();" >
  <div class="page-content">
    <div class="white-container">
      <!-- #include file = "partials/header.asp" -->
      <div class="allcontent">
        <div class="allcontent_main">

          <div class="header_blue">
            
            <div class="header_progress">YOUR PROGRESS<br>
            </div>
            
            <div class="header_inside">TRAINING & QUIZ<br>
            <h3>Before you start</h3>
            </div>
          </div>
          
          <div class="clear"></div>
          
          <div class="main_content">
            <div class="guide_blue">
              <div class="box_inside"><h3>Welcome</h3>
                <div class="box_text_blue">
                  <form ACTION="<%=MM_editAction%>" name="session" method="POST">
                    <span class="heading"> Creating a session...</span>
                        <br>
                              <input type="hidden" name="user" value="<%=userID%>">
                              <input type="hidden" name="subject" value="<%=ID_subject_prm%>">
                              <input type="hidden" name="date" value="<%=cDateSql(Now())%>">
                              <input type="hidden" name="total" value="<%=totalq%>">
                              <input type="hidden" name="correct" value="0">
                              <input type="hidden" name="stop" value="0">
                              <input type="hidden" name="done" value="0">
                      Please wait while the subject is loading
                              for you....<% response.write(CStr(Request("MM_insert"))) %><br>
                          <br>
                          Thank you.
                      <input type="hidden" name="MM_insert" value="true">
                    </form>
                </div>
              </div>
            </div>
          </div>
          
          <div class="menu_content"><img src="vault_image/images/training.jpg" width="320" height="390" alt=""></div>

          <div class="clear"></div>
        </div>
      </div>
    </div>
    <!-- #include file = "partials/footer.asp" -->
  </div>
</body>
<!-- #include file = "errorhandler/index.asp"-->