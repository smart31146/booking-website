<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/include_admin.asp" -->
<!--#include file="../connections/bbg_conn.asp" -->
<%
SQL_where = ""
'show_lines = 100

'if cStr(Request("show_lines")) <> "" then show_lines = cInt(Request("show_lines"))

fromdate=Replace(cStr(Replace(Request("fromdate"), ".", "/")), "'", "''")
fromdate=cDateSQL(fromdate)
If fromdate <> "" then 
	SQL_where = " WHERE (logs.log_date) >= " + database_date_string + fromdate + database_date_string 
end if 

todate = Replace(cStr(Replace(Request("todate"), ".", "/")), "'", "''")
todate=cDateSQL(todate)
If todate <> "" then 
	if SQL_where <> "" then
		SQL_where = SQL_where + " AND (logs.log_date) <= " + database_date_string + todate + database_date_string
	else
		SQL_where = " WHERE (logs.log_date) <= " + database_date_string + todate + database_date_string  
	end if 
end if 

If cStr(Request("ipaddress")) <> "" then 
	if SQL_where <> "" then
		SQL_where = SQL_where + " AND (logs.log_ip) Like '" + Replace(cStr(Request("ipaddress")), "'", "''") + "%' "
	else
		SQL_where = " WHERE (logs.log_ip) Like '" + Replace(cStr(Request("ipaddress")), "'", "''") + "%' " 
	end if 
end if 

If cStr(Request("module")) <> "" AND cStr(Request("module")) <> "all" then 
	if SQL_where <> "" then
		SQL_where = SQL_where + " AND (logs.log_module) Like '" + Replace(cStr(Request("module")), "'", "''") + "' "
	else
		SQL_where = " WHERE (logs.log_module) Like '" + Replace(cStr(Request("module")), "'", "''") + "' " 
	end if 
end if 

If cStr(Request("subject")) <> "" AND cStr(Request("subject")) <> "0" then 
	if SQL_where <> "" then
		SQL_where = SQL_where + " AND (logs.log_subjID) = " + (Request("subject")) + " "
	else
		SQL_where = " WHERE (logs.log_subjID) = " + (Request("subject")) + " " 
	end if 
end if 

If cStr(Request("username")) <> "" then 
	if SQL_where <> "" then
		SQL_where = SQL_where + " AND (logs.log_user) Like '%" + Replace(cStr(Request("username")), "'", "''") + "%' "
	else
		SQL_where = " WHERE (logs.log_user) Like '%" + Replace(cStr(Request("username")), "'", "''") + "%' " 
	end if 
end if 

If cStr(Request("comment")) <> "" then 
	if SQL_where <> "" then
		SQL_where = SQL_where + " AND (logs.log_comment) Like '%" + Replace(cStr(Request("comment")), "'", "''") + "%' "
	else
		SQL_where = " WHERE (logs.log_comment) Like '%" + Replace(cStr(Request("comment")), "'", "''") + "%' " 
	end if 
end if 

If cint(Request("info1")) <> 0 then 
	if SQL_where <> "" then
		SQL_where = SQL_where + " AND (q_user.user_info1) = " + (Request("info1")) + " "
	else
		SQL_where = " WHERE (q_user.user_info1) = " + (Request("info1")) + " " 
	end if 
end if 

If cint(Request("info2")) <> 0 then 
	if SQL_where <> "" then
		SQL_where = SQL_where + " AND (q_user.user_info2) = " + (Request("info2")) + " "
	else
		SQL_where = " WHERE (q_user.user_info2) = " + (Request("info2")) + " " 
	end if 
end if 

If cint(Request("info3")) <> 0 then 
	if SQL_where <> "" then
		SQL_where = SQL_where + " AND (q_user.user_info3) = " + (Request("info3")) + " "
	else
		SQL_where = " WHERE (q_user.user_info3) = " + (Request("info3")) + " " 
	end if 
end if 

If (cint(Request("info1")) <> 0 OR cint(Request("info2")) <> 0 OR cint(Request("info3")) <> 0) then 
	SQL_where = SQL_where + " AND logs.log_userID = q_user.ID_user "
end if

%>

<%
set logs = Server.CreateObject("ADODB.Recordset")
logs.ActiveConnection = Connect
If (cint(Request("info1")) <> 0 OR cint(Request("info2")) <> 0 OR cint(Request("info3")) <> 0) then 
	logs.Source = "SELECT * FROM logs, q_user " + SQL_where + " ORDER BY log_date DESC, ID_log DESC;"
else
	logs.Source = "SELECT * FROM logs " + SQL_where + " ORDER BY ID_log DESC;"
end if
logs.CursorType = 0
logs.CursorLocation = 3
logs.LockType = 3
logs.Open()
logs_numRows = 0
'Response.Write logs.Source

if cint(request("subject")) <> 0 then 
	set subjects = Server.CreateObject("ADODB.Recordset")
	subjects.ActiveConnection = Connect
	subjects.Source = "SELECT subject_name FROM subjects where ID_subject=" & request("subject")
	subjects.CursorType = 0
	subjects.CursorLocation = 3
	subjects.LockType = 3
	subjects.Open()
	subjects_numRows = 0
end if

if request("info1") <> 0 then 
	set info11 = Server.CreateObject("ADODB.Recordset")
	info11.ActiveConnection = Connect
	info11.Source = "SELECT * from q_info1 where ID_info1=" & request("info1")
	info11.CursorType = 0
	info11.CursorLocation = 3
	info11.LockType = 3
	info11.Open()
	info11_numRows = 0
end if

if request("info2") <> 0 then 
	set info22 = Server.CreateObject("ADODB.Recordset")
	info22.ActiveConnection = Connect
	info22.Source = "SELECT * from q_info2 where ID_info2=" & request("info2")
	info22.CursorType = 0
	info22.CursorLocation = 3
	info22.LockType = 3
	info22.Open()
	info22_numRows = 0
end if

if request("info3") <> 0	 then 
	set info33 = Server.CreateObject("ADODB.Recordset")
	info33.ActiveConnection = Connect
	info33.Source = "SELECT * from q_info3 where ID_info3=" & request("info3")
	info33.CursorType = 0
	info33.CursorLocation = 3
	info33.LockType = 3
	info33.Open()
	info33_numRows = 0
end if
%>

<%
'Dim Repeat1__numRows
'Repeat1__numRows = show_lines
'Dim Repeat1__index
'Repeat1__index = 0
'logs_numRows = logs_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
logs_total = logs.RecordCount

' set the number of rows displayed on this page
If (logs_numRows < 0) Then
  logs_numRows = logs_total
Elseif (logs_numRows = 0) Then
  logs_numRows = 1
End If

' set the first and last displayed record
logs_first = 1
logs_last  = logs_first + logs_numRows - 1

' if we have the correct record count, check the other stats
If (logs_total <> -1) Then
  If (logs_first > logs_total) Then logs_first = logs_total
  If (logs_last > logs_total) Then logs_last = logs_total
  If (logs_numRows > logs_total) Then logs_numRows = logs_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (logs_total = -1) Then

  ' count the total records by iterating through the recordset
  logs_total=0
  While (Not logs.EOF)
    logs_total = logs_total + 1
    logs.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (logs.CursorType > 0) Then
'    logs.MoveFirst
'  Else
    logs.Requery
  End If

  ' set the number of rows displayed on this page
  If (logs_numRows < 0 Or logs_numRows > logs_total) Then
    logs_numRows = logs_total
  End If

  ' set the first and last displayed record
  logs_first = 1
  logs_last = logs_first + logs_numRows - 1
  If (logs_first > logs_total) Then logs_first = logs_total
  If (logs_last > logs_total) Then logs_last = logs_total

End If
%>
<%
' *** Move To Record and Go To Record: declare variables

Set MM_rs    = logs
MM_rsCount   = logs_total
MM_size      = logs_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  r = Request.QueryString("index")
  If r = "" Then r = Request.QueryString("offset")
  If r <> "" Then MM_offset = Int(r)

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  i = 0
  While ((Not MM_rs.EOF) And (i < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    i = i + 1
  Wend
  If (MM_rs.EOF) Then MM_offset = i  ' set MM_offset to the last possible record

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  i = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or i < MM_offset + MM_size))
    MM_rs.MoveNext
    i = i + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = i
    If (MM_size < 0 Or MM_size > MM_rsCount) Then MM_size = MM_rsCount
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
'  If (MM_rs.CursorType > 0) Then
'    MM_rs.MoveFirst
'  Else
    MM_rs.Requery
'  End If

  ' move the cursor to the selected record
  i = 0
  While (Not MM_rs.EOF And i < MM_offset)
    MM_rs.MoveNext
    i = i + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
logs_first = MM_offset + 1
logs_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (logs_first > MM_rsCount) Then logs_first = MM_rsCount
  If (logs_last > MM_rsCount) Then logs_last = MM_rsCount
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then MM_removeList = MM_removeList & "&" & MM_paramName & "="
MM_keepURL="":MM_keepForm="":MM_keepBoth="":MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
if (MM_keepBoth <> "") Then MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
if (MM_keepURL <> "")  Then MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
if (MM_keepForm <> "") Then MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 0) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    params = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For i = 0 To UBound(params)
      nextItem = Left(params(i), InStr(params(i),"=") - 1)
      If (StrComp(nextItem,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & params(i)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then MM_keepMove = MM_keepMove & "&"
urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="
MM_moveFirst = urlStr & "0"
MM_moveLast  = urlStr & "-1"
MM_moveNext  = urlStr & Cstr(MM_offset + MM_size)
prev = MM_offset - MM_size
If (prev < 0) Then prev = 0
MM_movePrev  = urlStr & Cstr(prev)
%>

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<%
Response.Clear()
Response.AddHeader "Content-Disposition","attachment; filename=log_results_" & day(now()) & "_" & month(now()) & "_" & year(now()) & ".csv"
Response.ContentType="application/vnd.ms-excel"
%>

<%
todate=request("todate")
fromdate=request("fromdate")
module=request("module")
subject=request("subject")
username=request("username")
ipaddress=request("ipaddress")
comment=request("comment")
info1=request("info1")
info2=request("info2")
info3=request("info3")%>
 
Start date:, <%=fromdate%>
End date:, <%=todate%>
IP address:, <%=ipaddress%>
Module:, <%=module%>
User name:, <%=username%>
Subject:,  <%if subject <> 0 then %><%Response.Write subjects("subject_name")%><%subjects.close%><%end if%>
Comment:, <%=comment%>
Business Group:, <%if info1 <> 0 then%><%Response.Write info11.Fields.item("info1").value%><%info11.close%><%end if%>
Business <% =BBPinfo3 %>:, <%if info2 <> 0 then%><%Response.Write info22("info2")%><%info22.close%><%end if%>
<% =BBPinfo3 %>:, <%if info3 <> 0 then%><%Response.Write info33("info3")%><%info33.close%><%end if%>
    
"ID","DATE","IP ADDRESS","MODULE","USERNAME","SUBJECT","TOPIC","PAGE","URL","COMMENT"<%=vbcrlf%>"----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"<%=vbcrlf%>

<% 
While (NOT logs.EOF)
%>
<%=(logs.Fields.Item("ID_log").Value)%>,<%=(logs.Fields.Item("log_date").Value)%>,<%=(logs.Fields.Item("log_IP").Value)%>,<%=(logs.Fields.Item("log_module").Value)%>,<%=(logs.Fields.Item("log_user").Value)%>,<%=(logs.Fields.Item("log_subj").Value)%>,<%=(logs.Fields.Item("log_topic").Value)%>,<%=(logs.Fields.Item("log_page").Value)%>,<%=(logs.Fields.Item("log_url").Value)%>,<%=(logs.Fields.Item("log_comment").Value)%><%=vbcrlf%> 
<% 
Repeat1__index=Repeat1__index+1
Repeat1__numRows=Repeat1__numRows-1
logs.MoveNext()
Wend
%>
<% If logs.EOF And logs.BOF Then %>
Sorry, there are currently no LOG entries or your search criteria do not match with any loged records...
<% End If ' end logs.EOF And logs.BOF %>
<%=vbcrlf%>"-------------------------------------------------------------------------------------------------------------------------"<%=vbcrlf%>"Generated on:","<%=Now()%>"<%=vbcrlf%><%=vbcrlf%>"Copyright 2011 (c) Law of the Jungle Pty Limited" 
<%
logs.Close()
%>
