<!-- Begin ASP Source Code -->
  <%@ LANGUAGE="VBSCRIPT" %>
  <!--#include file="../connections/bbg_conn.asp" -->
  <!--#include file="../connections/include_admin.asp" -->
  

<%
'***************************************************Function (1) Find Duplicate User Name************************************************
    Dim FindDuplicateActivity 
	FindDuplicateActivity=false
	Dim BizID
	Function FindDuplicateBusiness()
	
	FindDuplicateActivity = ""
       ' boolean to abort record edit
        MM_abortEdit = false
        ' query string to execute
            
            DuplicateUID = user_business
            set ObjRS1 = Server.CreateObject("ADODB.Recordset")
            ObjRS1.Open "SELECT * FROM q_info1 WHERE info1='" & Replace(DuplicateUID,"'","''") & "'", Connect
            
            'User name already exsits in the data base
            If Not ObjRS1.EOF Or Not ObjRS1.BOF Then
                FindDuplicateBusiness = "True"
                Response.Write("<BR>Business: <b>" & ObjRS1("info1") & "</b> already exists <br>")
			  BizID=ObjRS1("ID_info1") 
				
				else
				Response.Write("<BR>NEW Business<BR>")
             End If
             ObjRS1.Close
             Set ObjRS1 = nothing
       
    End Function
%>
     
<%
'***************************************************Function (2) Filling an Array with the info. of a user******************************
    Sub FillArrayWithUserInfo()

            MM_editConnection = Connect
            MM_editTable = "q_info1"
            MM_editRedirectUrl = "q_list_of_users.asp"
            MM_fieldsStr  = "info1|value"
            MM_columnsStr = "info1|',none,''"
Response.Write("M field " & MM_fieldsStr  & "<BR>")
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
				
				Response.Write("M field " & MM_fields(i) & "<BR>")
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
	Response.Write("<BR><BR>Connecting to the file: <b>" & MM_tableValues & "</b> ...<BR>")
    MM_tableValues = MM_tableValues & MM_columns(i)
    MM_dbValues = MM_dbValues & FormVal
	Response.Write("<BR><BR>values <b>" & MM_dbValues & "</b> ...<BR>")
  Next
 MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

    ' execute the insert
   Set MM_editCmd = Server.CreateObject("ADODB.Command")
   MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    if Edit_OK = true then MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
   ' call log_the_page ("Quiz Execute - INSERT Business")
    'create_membership
    'pn 050726 added means to connect a new user with a subject
	set info11 = Server.CreateObject("ADODB.Recordset")
info11.ActiveConnection = Connect
info11.Source = "SELECT * FROM q_info1 where info1=" & MM_dbValues 
info11.CursorType = 0
info11.CursorLocation = 3
info11.LockType = 3
info11.Open()
info11_numRows = 0
Response.Write("Site " & user_site & "<BR>")

While Not info11.EOF
	Response.Write("<BR><BR>business id is  <b>" & info11("ID_info1") & "</b> ...<BR>")
	
	Set MM_editCmd = Server.CreateObject("ADODB.Command")
					MM_editCmd.ActiveConnection = Connect
					MM_editCmd.CommandText = "insert into q_info2 (info2, info2_info1) values ('"&user_site&"',"&cInt(info11("ID_info1"))&");"
					MM_editCmd.Execute
					MM_editCmd.ActiveConnection.Close
	
	
	
info11.MoveNext()

WEND


   
'050720 do an insert to the subject_user table
					

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
		                 ReturnRequestForm = user_business        
                         
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
   '***************************************************Main Body of the Program************************************************
        Dim filePath, Range
       
        Path = "C:\www\BBP\BBP_acme3.4_21-04-16\admin"
		filename = "myExcelBiz.xls"
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
		                login_pass = Trim(objRS.Fields.Item(X).Value)
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
			
         if (FindDuplicateBusiness <> "True") then
                    call FillArrayWithUserInfo()
			else 
	
            
            'User name already exsits in the data base
            
			
			set info22 = Server.CreateObject("ADODB.Recordset")
				info22.ActiveConnection = Connect
				info22.Source = "SELECT * FROM q_info2 where info2_info1="&BizID & " and info2='"& user_site & "'"
				info22.CursorType = 0
				info22.CursorLocation = 3
				info22.LockType = 3
				info22.Open()
				info22_numRows = 0

			If Not info22.EOF then
				FindDuplicateActivity = true
					Response.Write("<BR>Site: <b>" & info22("info2") & "</b> Already Exists <br>")
			else
			
			Set MM_editCmd = Server.CreateObject("ADODB.Command")
					MM_editCmd.ActiveConnection = Connect
					MM_editCmd.CommandText = "insert into q_info2 (info2, info2_info1) values ('"&user_site&"',"&BizID&");"
					MM_editCmd.Execute
					MM_editCmd.ActiveConnection.Close
					
					Response.Write("<BR>Site: <b>" & user_site & "</b> has been Entered <br>")
					
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
  
     
     

