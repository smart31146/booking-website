<!-- Begin ASP Source Code -->
  <%@ LANGUAGE="VBSCRIPT" %>
  <!--#include file="../connections/bbg_conn.asp" -->
  <!--#include file="../connections/include_admin.asp" -->
  <!--#include file="sha256.asp"-->
  
  <%
Response.Write(Server.ScriptTimeout)
%>
  <%
Server.ScriptTimeout=900
%>

<%
'***************************************************Function (1) Find Duplicate User Name************************************************
    Function FindDuplicateUserName()
       ' boolean to abort record edit
        MM_abortEdit = false
        ' query string to execute
            
			
            DuplicateUID = CStr(login_name)
            set ObjRS1 = Server.CreateObject("ADODB.Recordset")
            ObjRS1.Open "SELECT * FROM q_user WHERE user_username='" & Replace(DuplicateUID,"'","''") & "'", Connect
            
            'User name already exsits in the data base
            If Not ObjRS1.EOF Or Not ObjRS1.BOF Then
                FindDuplicateUserName = "True"
                Response.Write("<BR>User Name: <b>" & ObjRS1("user_username") & "</b> already exists. <BR>")
				userID=ObjRS1("ID_user")
             End If
             ObjRS1.Close
             Set ObjRS1 = nothing
       
    End Function
%>
     
<%
'***************************************************Function (2) Filling an Array with the info. of a user******************************
    Sub FillArrayWithUserInfo()

            MM_editConnection = Connect
            MM_editTable = "q_user"
            MM_editRedirectUrl = "q_list_of_users.asp"
            MM_fieldsStr  = "first_name|value|last_name|value|login_name|value|info1|value|info2|value|info3|value|info4|value|active|value|session|value|email|value|change_pass|value|reference|value"
            MM_columnsStr = "user_firstname|',none,''|user_lastname|',none,''|user_username|',none,''|user_info1|none,none,NULL|user_info2|none,none,NULL|user_info3|none,none,NULL|user_info4|none,none,NULL|user_active|none,1,0|user_new_session|',none,''|user_email|',none,''|user_password|none,1,0|user_reference|',none,''"

            ' create the MM_fields and MM_columns arrays
            MM_fields = Split(MM_fieldsStr, "|")
            MM_columns = Split(MM_columnsStr, "|")

            For i = LBound(MM_fields) To UBound(MM_fields) Step 2
            'Replacement of Request.Form(MM_fields(i)) with ReturnRequestForm(MM_fields(i))
                if not IsNull( ReturnRequestForm(MM_fields(i)) ) then
                    MM_fields(i+1) = CStr(ReturnRequestForm(MM_fields(i)))
                Else
                    MM_fields(i+1) = ""
                end if
            Next


%>

<%


' *** Insert Record: construct a sql insert staatement and execute it
'Hereeeeeeeeeeeeeeeeeee Set Request_MM_insert = True when the submit button is pressed

  ' create the sql insert staatement
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

    ' execute the insert
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    if Edit_OK = true then MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
    call log_the_page ("Quiz Execute - INSERT User")
    'create_membership
    'pn 050726 added means to connect a new user with a subject
    create_user_subjects


End Sub
 %>
 
 <%
   '***************************************************Function ReturnRequestForm************************************************ 
   Function ReturnRequestForm(Entry)
    
    
    set info11 = Server.CreateObject("ADODB.Recordset")
info11.ActiveConnection = Connect
info11.Source = "SELECT * FROM q_info1"
info11.CursorType = 0
info11.CursorLocation = 3
info11.LockType = 3
info11.Open()
info11_numRows = 0
%>
<%
set info22 = Server.CreateObject("ADODB.Recordset")
info22.ActiveConnection = Connect
info22.Source = "SELECT * FROM q_info2"
info22.CursorType = 0
info22.CursorLocation = 3
info22.LockType = 3
info22.Open()
info22_numRows = 0
%>
<%
set info33 = Server.CreateObject("ADODB.Recordset")
info33.ActiveConnection = Connect
info33.Source = "SELECT * FROM q_info3"
info33.CursorType = 0
info33.CursorLocation = 3
info33.LockType = 3
info33.Open()
info33_numRows = 0

set info44 = Server.CreateObject("ADODB.Recordset")
info44.ActiveConnection = Connect
info44.Source = "SELECT * FROM q_info4"
info44.CursorType = 0
info44.CursorLocation = 3
info44.LockType = 3
info44.Open()
info44_numRows = 0

Select Case Entry
	                Case "first_name"
		                ReturnRequestForm = first_name
	                Case "last_name"
		                ReturnRequestForm = last_name
		            Case "login_name"
		                ReturnRequestForm = login_name
					Case "login_pass"
	                    ReturnRequestForm = login_pass
		             Case "info1"
		                 While (NOT info11.EOF)
                            if UCASE(info11.Fields.Item("info1").Value) = UCASE(user_business) then
                                ReturnRequestForm = info11.Fields.Item("ID_info1").Value        
                            End if
                            info11.MoveNext()
                          Wend

                     Case "info2"
		                 While (NOT info22.EOF)
                            if UCASE(info22.Fields.Item("info2").Value) = UCASE(user_site) then
                                ReturnRequestForm = info22.Fields.Item("ID_info2").Value        
                            End if
                            info22.MoveNext()
                          Wend
                     
                     Case "info3"
		                 While (NOT info33.EOF)
                            if UCASE(info33.Fields.Item("info3").Value) = UCASE(user_activity) then
                                ReturnRequestForm = info33.Fields.Item("ID_info3").Value        
                            End if
                            info33.MoveNext()
                          Wend
					 
					 Case "info4"
		                 While (NOT info44.EOF)
                            if UCASE(info44.Fields.Item("info4").Value) = UCASE(user_company) then
                                ReturnRequestForm = info44.Fields.Item("ID_info4").Value        
                            End if
                            info44.MoveNext()
                          Wend  

	                 Case "active"
	                    ReturnRequestForm = 1
	                 Case "session"
	                    ReturnRequestForm = user_session	                    
		             Case "email"
	                    ReturnRequestForm = user_email	                    
		             Case "change_pass"
	                    ReturnRequestForm = user_password
					 Case "reference"
	                    ReturnRequestForm = user_employeereference
	 End Select
	 info11.Close
     Set info11 = Nothing
     info22.Close
     Set info22 = Nothing
     info33.Close
     Set info33 = Nothing
	 info44.Close
     Set info44 = Nothing
     
   End Function
%>
 
<% 
   '***************************************************Sub One BuildingRequestForm************************************************
Function GetSubjects()
    set subjects_b = Server.CreateObject("ADODB.Recordset")
			subjects_b.ActiveConnection = Connect
			subjects_b.Source = "SELECT subjects.ID_subject, subjects.subject_name  FROM (subjects INNER JOIN b_topics ON subjects.ID_subject = b_topics.topic_subject) INNER JOIN b_pages ON b_topics.ID_topic = b_pages.page_topic  GROUP BY subjects.ID_subject, subjects.subject_name, subjects.subject_ord, subjects.ID_subject, Abs([subject_active_b]), Abs([topic_active]), Abs([page_active])  HAVING (((Abs([subject_active_b]))=1) AND ((Abs([topic_active]))=1) AND ((Abs([page_active]))=1))  ORDER BY subjects.subject_ord, subjects.ID_subject;"
			subjects_b.CursorType = 0
			subjects_b.CursorLocation = 2
			subjects_b.LockType = 3
			subjects_b.Open()
			subjects_b_numRows = 0
			
		Dim AllSubjects
		AllSubjects = "" 
        While (NOT subjects_b.EOF)
			if ( instr(1, user_subjects, subjects_b.Fields.Item("subject_name").Value, 1) ) then
				AllSubjects = AllSubjects & "user_subject|0|" & subjects_b.Fields.Item("ID_subject").Value &"@"
			end if
			subjects_b.MoveNext()
		Wend
	subjects_b.close()
	set subjects_b = nothing
    GetSubjects = AllSubjects
End Function
  %>
  
<%
 '***************************************************Sub create_user_subjects************************************************
sub create_user_subjects
	MM_editConnection = Connect

	set new_user = Server.CreateObject("ADODB.Recordset")
	new_user.ActiveConnection = Connect
	
	new_user.Source = "SELECT * FROM q_user WHERE user_new_session = '" + ReturnRequestForm("session") + "';"
	
	 '" + RequestSession + "';"
	new_user.CursorType = 0
	new_user.CursorLocation = 3
	new_user.LockType = 3
	new_user.Open()
	new_user_numRows = 0
	new_user_id = (new_user.Fields.Item("ID_User").Value)
	new_user.Close()
    
	'PN 050720 Save the user subjects that have been submitted
	Dim updated_ok
	updated_ok=false

    'Replacement of request("send_user_email") with request_send_user_email
	if request_send_user_email = 1 then
		' insert email
		'Check the sql_date
		sql_date=cDateSql(now_bbp())

		MM_editConnection = Connect
		MM_editQuery = "insert into emails (q_user,date_to_send,type) values ('"&new_user_id&"','"&sql_date&"' , 1)"
	    Set MM_editCmd = Server.CreateObject("ADODB.Command")
	    MM_editCmd.ActiveConnection = MM_editConnection
	    MM_editCmd.CommandText = MM_editQuery
	    MM_editCmd.Execute
	    MM_editCmd.ActiveConnection.Close

		Dim last_email_inserted
		set last_id_insert = Server.CreateObject("ADODB.Recordset")
		last_id_insert.ActiveConnection = Connect
		last_id_insert.Source = "SELECT max (id) as idd from emails"
		last_id_insert.CursorType = 0
		last_id_insert.CursorLocation = 3
		last_id_insert.LockType = 3
		last_id_insert.Open()
		last_id_insert_numRows = 0
		last_email_inserted=last_id_insert.Fields.Item("idd").Value
		last_id_insert.Close()

		strSql = ""
	end if
 
	Dim Seps(1)
	Seps(0) = "|"
	'PN 050720 delete from the  subject user table

    'Make the Request.Form()
    ReturnSubjects = GetSubjects()
    SplitSubjects = Split(ReturnSubjects, "@", -1, 1)
    'Replacement of Request.Form() with RequestForm
	
	
	For Each q in SplitSubjects

		if (((InStr(q,"user_subject"))>0)=True) then

				Dim a
				a= Tokenize(q, Seps)
                
				'050720 do an insert to the subject_user table
				Set MM_editCmd = Server.CreateObject("ADODB.Command")
				MM_editCmd.ActiveConnection = Connect
				MM_editCmd.CommandText = "insert into subject_user (ID_subject, ID_user) values ("&a(2)&","&cdbl(new_user_id) &");"
				MM_editCmd.Execute
				if a(2) = 1 then
					MM_editCmd.CommandText = "insert into subject_user (ID_subject, ID_user) select 8,"&cdbl(new_user_id)&" union all select 9,"&cdbl(new_user_id)&" union all select 10,"&cdbl(new_user_id)&" union all select 11,"&cdbl(new_user_id)&";"
					MM_editCmd.Execute
				end if
				MM_editCmd.ActiveConnection.Close

 				'if request_change_user_password = 1 then
				 '   Set MM_insertPass = Server.CreateObject("ADODB.Command")
				  '  MM_insertPass.ActiveConnection = Connect
				'	strSql = "insert into q_user (user_password) values (1)"
				'	MM_insertPass.CommandText = strSql
				'	MM_insertPass.Execute
				'	MM_insertPass.ActiveConnection.Close
				'end if

				if request_send_user_email = 1 then
					set check_subject_quiz = Server.CreateObject("ADODB.Recordset")
					check_subject_quiz.ActiveConnection = Connect
					check_subject_quiz.Source = "SELECT subject_active_q from subjects where subject_active_q=1 and id_subject="&a(2)
					check_subject_quiz.CursorType = 0
					check_subject_quiz.CursorLocation = 3
					check_subject_quiz.LockType = 3
					check_subject_quiz.Open()

					if (not check_subject_quiz.eof) then
						'insert each subject for the emailer
						strSql = "Insert into subject_email (subject,email) values( '"&a(2)&"', '"& last_email_inserted&"')"

						'response.write ( MM_editQuery)
				    	Set MM_insertCmd = Server.CreateObject("ADODB.Command")
				    	MM_insertCmd.ActiveConnection = Connect
				    	MM_insertCmd.CommandText = strSql
				    	MM_insertCmd.Execute
				    	MM_insertCmd.ActiveConnection.Close
					end if
					if a(2)=1 then
						'insert each subject for the emailer
				    	Set MM_insertCmd = Server.CreateObject("ADODB.Command")
				    	MM_insertCmd.ActiveConnection = Connect
						strSql = "insert into subject_email (subject,email) select 8,"&last_email_inserted&" union all select 9,"&last_email_inserted&" union all select 10,"&last_email_inserted&" union all select 11,"&last_email_inserted&";"
						MM_insertCmd.CommandText = strSql
					   	MM_insertCmd.Execute
					   	MM_insertCmd.ActiveConnection.Close
					end if
					check_subject_quiz.Close()

				end if
		end if

	Next
end sub
   %>
   
   
  <%
 '***************************************************Sub create_user_subjects_exist************************************************
sub create_user_subjects_exist
	MM_editConnection = Connect

	set new_user = Server.CreateObject("ADODB.Recordset")
	new_user.ActiveConnection = Connect
	
	new_user.Source = "SELECT * FROM q_user WHERE user_username='" & CStr(login_name) & "'"
	
	 '" + RequestSession + "';"
	new_user.CursorType = 0
	new_user.CursorLocation = 3
	new_user.LockType = 3
	new_user.Open()
	new_user_numRows = 0
	new_user_id = (new_user.Fields.Item("ID_User").Value)
	new_user.Close()
    
	'PN 050720 Save the user subjects that have been submitted
	Dim updated_ok
	updated_ok=false

    'Replacement of request("send_user_email") with request_send_user_email
	if request_send_user_email = 1 then
		' insert email
		'Check the sql_date
		sql_date=cDateSql(now_bbp())

		MM_editConnection = Connect
		MM_editQuery = "insert into emails (q_user,date_to_send,type) values ('"&new_user_id&"','"&sql_date&"' , 1)"
	    Set MM_editCmd = Server.CreateObject("ADODB.Command")
	    MM_editCmd.ActiveConnection = MM_editConnection
	    MM_editCmd.CommandText = MM_editQuery
	    MM_editCmd.Execute
	    MM_editCmd.ActiveConnection.Close

		Dim last_email_inserted
		set last_id_insert = Server.CreateObject("ADODB.Recordset")
		last_id_insert.ActiveConnection = Connect
		last_id_insert.Source = "SELECT max (id) as idd from emails"
		last_id_insert.CursorType = 0
		last_id_insert.CursorLocation = 3
		last_id_insert.LockType = 3
		last_id_insert.Open()
		last_id_insert_numRows = 0
		last_email_inserted=last_id_insert.Fields.Item("idd").Value
		last_id_insert.Close()

		strSql = ""
	end if
 
	Dim Seps(1)
	Seps(0) = "|"
	'PN 050720 delete from the  subject user table

    'Make the Request.Form()
    ReturnSubjects = GetSubjects()
    SplitSubjects = Split(ReturnSubjects, "@", -1, 1)
    'Replacement of Request.Form() with RequestForm
	
	
	'Delete subjects from subject_user table
				    	Set MM_deleteCmd = Server.CreateObject("ADODB.Command")
				    	MM_deleteCmd.ActiveConnection = Connect
						strSql = "delete from subject_user where ID_user="&userID
						MM_deleteCmd.CommandText = strSql
					   	MM_deleteCmd.Execute
					   	MM_deleteCmd.ActiveConnection.Close
	
	For Each q in SplitSubjects

			if (((InStr(q,"user_subject"))>0)=True) then

				
				a= Tokenize(q, Seps)
                
				'050720 do an insert to the subject_user table
				Set MM_editCmd = Server.CreateObject("ADODB.Command")
				MM_editCmd.ActiveConnection = Connect
				MM_editCmd.CommandText = "insert into subject_user (ID_subject, ID_user) values ("&a(2)&","&userID&");"
				MM_editCmd.Execute
				if a(2) = 1 then
					MM_editCmd.CommandText = "insert into subject_user (ID_subject, ID_user) select 8,"&userID&" union all select 9,"&userID&" union all select 10,"&userID&" union all select 11,"&userID&";"
					MM_editCmd.Execute
				end if
				MM_editCmd.ActiveConnection.Close
				
				if request_send_user_email = 1 then
					set check_subject_quiz = Server.CreateObject("ADODB.Recordset")
					check_subject_quiz.ActiveConnection = Connect
					check_subject_quiz.Source = "SELECT subject_active_q from subjects where subject_active_q=1 and id_subject="&a(2)
					check_subject_quiz.CursorType = 0
					check_subject_quiz.CursorLocation = 3
					check_subject_quiz.LockType = 3
					check_subject_quiz.Open()

					if (not check_subject_quiz.eof) then
						'insert each subject for the emailer
						strSql = "Insert into subject_email (subject,email) values( '"&a(2)&"', '"& last_email_inserted&"')"

						'response.write ( MM_editQuery)
				    	Set MM_insertCmd = Server.CreateObject("ADODB.Command")
				    	MM_insertCmd.ActiveConnection = Connect
				    	MM_insertCmd.CommandText = strSql
				    	MM_insertCmd.Execute
				    	MM_insertCmd.ActiveConnection.Close
					end if
					if a(2)=1 then
						'insert each subject for the emailer
				    	Set MM_insertCmd = Server.CreateObject("ADODB.Command")
				    	MM_insertCmd.ActiveConnection = Connect
						strSql = "insert into subject_email (subject,email) select 8,"&last_email_inserted&" union all select 9,"&last_email_inserted&" union all select 10,"&last_email_inserted&" union all select 11,"&last_email_inserted&";"
						MM_insertCmd.CommandText = strSql
					   	MM_insertCmd.Execute
					   	MM_insertCmd.ActiveConnection.Close
					end if
					check_subject_quiz.Close()

				end if
			end if
				
	Next			
	

end sub
   %> 
   
 <%
 '***************************************************Function  Tokenize************************************************
  Function Tokenize(byVal TokenString, byRef TokenSeparators())
	Dim NumWords, a(), NumSeps
	NumWords = 0
	NumSeps = UBound(TokenSeparators)

	Do
		Dim SepIndex, SepPosition
		SepPosition = 0
		SepIndex    = -1

		for i = 0 to NumSeps-1
			' Find location of separator in the string
			Dim pos
			pos = InStr(TokenString, TokenSeparators(i))
			' Is the separator present, and is it closest to the beginning of the string?
			If pos > 0 and ( (SepPosition = 0) or (pos < SepPosition) ) Then
				SepPosition = pos
				SepIndex    = i
			End If
		Next

		' Did we find any separators?
		If SepIndex < 0 Then
			' None found - so the token is the remaining string
			redim preserve a(NumWords+1)
			a(NumWords) = TokenString
		Else
			' Found a token - pull out the substring
			Dim substr
			substr = Trim(Left(TokenString, SepPosition-1))
			' Add the token to the list
			redim preserve a(NumWords+1)
			a(NumWords) = substr
			' Cutoff the token we just found
			Dim TrimPosition
			TrimPosition = SepPosition+Len(TokenSeparators(SepIndex))
			TokenString = Trim(Mid(TokenString, TrimPosition))
		End If

		NumWords = NumWords + 1
	loop while (SepIndex >= 0)

	Tokenize = a
End Function
%>
<%
 '***************************************************Sub update Business, Site, Activity for existing users************************************************
sub update_fields

						'Update Business
				    	Set MM_updateCmd = Server.CreateObject("ADODB.Command")
				    	MM_updateCmd.ActiveConnection = Connect
						strSql = "update q_user set user_info1="&cint(ReturnRequestForm("info1"))&" where ID_user="&userID
						MM_updateCmd.CommandText = strSql
					   	MM_updateCmd.Execute
					   	MM_updateCmd.ActiveConnection.Close
						
						'Update Site
				    	Set MM_updateCmd = Server.CreateObject("ADODB.Command")
				    	MM_updateCmd.ActiveConnection = Connect
						strSql = "update q_user set user_info2="&cint(ReturnRequestForm("info2"))&" where ID_user="&userID
						MM_updateCmd.CommandText = strSql
					   	MM_updateCmd.Execute
					   	MM_updateCmd.ActiveConnection.Close
						
						'Update Activity
				    	Set MM_updateCmd = Server.CreateObject("ADODB.Command")
				    	MM_updateCmd.ActiveConnection = Connect
						strSql = "update q_user set user_info3="&cint(ReturnRequestForm("info3"))&" where ID_user="&userID
						MM_updateCmd.CommandText = strSql
					   	MM_updateCmd.Execute
					   	MM_updateCmd.ActiveConnection.Close
						
						'Update Email
				    	Set MM_updateCmd = Server.CreateObject("ADODB.Command")
				    	MM_updateCmd.ActiveConnection = Connect
						strSql = "update q_user set user_email='"&ReturnRequestForm("email")&"' where ID_user="&userID
						MM_updateCmd.CommandText = strSql
					   	MM_updateCmd.Execute
					   	MM_updateCmd.ActiveConnection.Close


end sub
%>
<%
 '***************************************************Checks the Validity of Business, Site, Activity, and Subjects************************************************
Function SubjectIsValid(Entry)
    set subjects_b = Server.CreateObject("ADODB.Recordset")
			subjects_b.ActiveConnection = Connect
			subjects_b.Source = "SELECT subjects.ID_subject, subjects.subject_name  FROM (subjects INNER JOIN b_topics ON subjects.ID_subject = b_topics.topic_subject) INNER JOIN b_pages ON b_topics.ID_topic = b_pages.page_topic  GROUP BY subjects.ID_subject, subjects.subject_name, subjects.subject_ord, subjects.ID_subject, Abs([subject_active_b]), Abs([topic_active]), Abs([page_active])  HAVING (((Abs([subject_active_b]))=1) AND ((Abs([topic_active]))=1) AND ((Abs([page_active]))=1))  ORDER BY subjects.subject_ord, subjects.ID_subject;"
			subjects_b.CursorType = 0
			subjects_b.CursorLocation = 2
			subjects_b.LockType = 3
			subjects_b.Open()
			subjects_b_numRows = 0
		
		SubjectIsValid = "False"
        While (NOT subjects_b.EOF)
			if ( instr(1, Entry, subjects_b.Fields.Item("subject_name").Value, 1) ) then
				Entry = Replace(Entry, subjects_b.Fields.Item("subject_name").Value, "")
			end if
			subjects_b.MoveNext()
		Wend
		Entry = Trim(Entry)

		While (instr(1, Entry, ";", 1) >= 1)
			Entry = Replace(Entry, ";", "")
			Entry = Trim(Entry)
		Wend
		While (instr(1, Entry, ",", 1) >= 1)
			Entry = Replace(Entry, ",", "")
			Entry = Trim(Entry)
		Wend		
		Entry=Trim(Entry)

		If Len(Entry) = 0 then
		    SubjectIsValid = "True"
		End if
		
		subjects_b.close()
		set subjects_b = nothing
		
End Function

Function BusinessIsValid(Entry)
    set info1 = Server.CreateObject("ADODB.Recordset")
	info1.ActiveConnection = Connect
	info1.Source = "SELECT * FROM q_info1"
	info1.CursorType = 0
	info1.CursorLocation = 3
	info1.LockType = 3
	info1.Open()
	
	BusinessIsValid = "False"
	
	   While (NOT info1.EOF)
            if UCASE(info1.Fields.Item("info1").Value) = UCASE(Entry) then
				BusinessIsValid = "True"
            End if
            info1.MoveNext()
        Wend
     info1.Close
     Set info1 = Nothing
 End Function

Function SiteIsValid(Entry)
    set info2 = Server.CreateObject("ADODB.Recordset")
	info2.ActiveConnection = Connect
	info2.Source = "SELECT * FROM q_info2"
	info2.CursorType = 0
	info2.CursorLocation = 3
	info2.LockType = 3
	info2.Open()
	
	SiteIsValid = "False"
	
	   While (NOT info2.EOF)
            if UCASE(info2.Fields.Item("info2").Value) = UCASE(Entry) then
				SiteIsValid = "True"
            End if
            info2.MoveNext()
        Wend
     info2.Close
     Set info2 = Nothing  
 End Function
 
 
 Function ActivityIsValid(Entry) 
    set info3 = Server.CreateObject("ADODB.Recordset")
	info3.ActiveConnection = Connect
	info3.Source = "SELECT * FROM q_info3"
	info3.CursorType = 0
	info3.CursorLocation = 3
	info3.LockType = 3
	info3.Open()
	
	ActivityIsValid = "False"
	
	   While (NOT info3.EOF)
            if UCASE(info3.Fields.Item("info3").Value) = UCASE(Entry) then
				ActivityIsValid = "True"
            End if
            info3.MoveNext()
        Wend
     info3.Close
     Set info3 = Nothing 
 End Function
 
  Function CompanyIsValid(Entry) 
    set info4 = Server.CreateObject("ADODB.Recordset")
	info4.ActiveConnection = Connect
	info4.Source = "SELECT * FROM q_info4"
	info4.CursorType = 0
	info4.CursorLocation = 3
	info4.LockType = 3
	info4.Open()
	
	CompanyIsValid = "False"
	
	   While (NOT info4.EOF)
            if UCASE(info4.Fields.Item("info4").Value) = UCASE(Entry) then
				CompanyIsValid = "True"
            End if
            info4.MoveNext()
        Wend
     info4.Close
     Set info4 = Nothing 
 End Function
%>

 <%
   '***************************************************Main Body of the Program************************************************
        Dim filePath, Range
       Dim userID
        Path = "C:\www\BBP\BBP_acme3.4_21-04-16\admin"
		filename = "myExcel.xls"
		filePath = Path & "\" & filename
        ExcelRange = "myRange1"
        'Check if the file exists or not
		dim fs
		set fs = Server.CreateObject("Scripting.FileSystemObject")
		if not fs.FileExists(filePath) = true then
			Response.Write(filePath & "<b> does not exist! </b>")
			set fs = nothing
			Response.end()
		end if
		set fs = nothing
		'end checking
        Set objConn = Server.CreateObject("ADODB.Connection")
        'Connection string
        objConn.Open "Driver={Microsoft Excel Driver (*.xls)}; DriverId=790; DBQ=" & filePath & ";"
        Response.Write("<BR><BR>Connecting to the file: <b>" & filePath & "</b> ...<BR>")
        Set objRS = Server.CreateObject("ADODB.Recordset")
        
        'Select from the specified range in the spreadsheet
        Response.Write("<BR><BR>Select from the specified range (should be named: " & ExcelRange & ") within the spreadsheet ... <BR>")
        objRS.Open "Select * from " & ExcelRange, objConn
        While Not objRS.EOF
            dim email_temp
            'Read the fileds of rows of the spreadsheet 
            For X = 0 To objRS.Fields.Count - 1
                Select Case objRS.Fields.Item(X).Name
	                Case "FirstName"
		                first_name = Trim(objRS.Fields.Item(X).Value)
	                Case "LastName"
		                last_name = Trim(objRS.Fields.Item(X).Value)
		            Case "UserName"
		                login_name = Trim(objRS.Fields.Item(X).Value)
		            Case "Password"					
					   login_pass=Trim(objRS.Fields.Item(X).Value)
		            Case "Business"
	                    user_business = Trim(objRS.Fields.Item(X).Value)
		            Case "Site"
	                    user_site = Trim(objRS.Fields.Item(X).Value)
		            Case "Activity"
	                    user_activity = Trim(objRS.Fields.Item(X).Value)
					Case "Company"
	                    user_company = Trim(objRS.Fields.Item(X).Value)
		            Case "Email"
	                    user_email = Trim(objRS.Fields.Item(X).Value)
						email_temp=Trim(objRS.Fields.Item(X).Value)
						                
		            Case "EmployeeReference"
	                    user_employeereference = Trim(objRS.Fields.Item(X).Value)
					Case "Subjects"
	                    user_subjects = Trim(objRS.Fields.Item(X).Value)
		            Case "SendEmail"
	                    if  objRS.Fields.Item(X).Value = "X" or objRS.Fields.Item(X).Value = "x" then
		                    request_send_user_email = 1
		                Else
		                    request_send_user_email = 0
		                End if
					Case "ChangePass"
	                    if  objRS.Fields.Item(X).Value = "X" or objRS.Fields.Item(X).Value = "x" then
		                    user_password = 1
		                Else
		                    user_password = 0
		                End if
	                'Case Else
		             '   Response.Write("Please check the Name of the Field; it is not in the list of data base fields")
                End Select
            Next
			
            user_session = getPassword(30, "", "true", "true", "true", "false", "true", "true", "true", "false")
            
            Dim ValidUser
            ValidUser = 1
            'Check that necessary field are filled out
            if IsNull(first_name) or IsNull(last_name) or IsNull(login_name) or IsNull(user_business) or IsNull(user_site) or IsNull(user_activity) or IsNull(user_email) then
                Response.Write("<BR>Either the range in spreadsheet contains empty row OR <BR> One of the mandatory fields for User: <B>" & login_name & "</B> in the spreadsheet is emtpy. <BR> Please fill in the field and run the program again <BR>")
                ValidUser = 0
            Else
            
            If ActivityIsValid(user_activity) = "False" then
				Response.Write("<BR>Activity <BR> for the User: <B>" & login_name & "</B> is not valid. Please change the activity <BR>")
				ValidUser = 0
			End If
			If BusinessIsValid(user_business) = "False" then
				Response.Write("<BR>Business <BR> for the User: <B>" & login_name & "</B> is not valid. Please change the business <BR>")
				ValidUser = 0
			End If
			If SiteIsValid(user_site) = "False" then
				Response.Write("<BR>Site <BR> for the User: <B>" & login_name & "</B> is not valid. Please change the site <BR>")
				ValidUser = 0
			End If
			Dim user_subjects_temp
			user_subjects_temp = user_subjects
			If SubjectIsValid(user_subjects_temp) = "False" then
				Response.Write("<BR>One or more subjects <BR> for the User: <B>" & login_name & "</B> are not valid. Please change the subjects <BR>")
				ValidUser = 0
			End If
            End if			
            'end of reading the first row of spreadsheet
            if ValidUser <> 0 then
                if (FindDuplicateUserName <> "True") then
					Dim salt
					salt = email_temp					
					login_pass=sha256(login_pass&salt)
                    call FillArrayWithUserInfo()
				else
				response.write(CStr(login_name))
				create_user_subjects_exist
				update_fields
				
                end if
            end if   
           
            'Go to the next row of the spreadsheet
            objRS.MoveNext          
        Wend
        
      objRS.Close
      Set objRS = Nothing

      objConn.Close
      Set objConn = Nothing
   %>
  
     
     

