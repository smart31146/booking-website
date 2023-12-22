<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
' *** Edit Operations: declare variables
MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Update Record: set variables
If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = Connect
  MM_editTable = "f_email"
  MM_editColumn = "f_address"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "f_feedback_logs.asp"


  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

  
End If
%>
<%
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  'MM_editQuery = "update " & MM_editTable & " set f_address = '" & Request.form("email") & "'"
  'Allow email with apostrophe. JIRA issue BBP-59
	MM_editQuery = "update " & MM_editTable & " set f_address = '" & Replace(CStr(Request.Form("email")), "'", "''") & "'"

  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    if Edit_OK = true then MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
    call log_the_page ("Field feedback report Execute - UPDATE email: " & MM_recordId)
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim activity__MMColParam
activity__MMColParam = "1"
if (Request.QueryString("f_address") <> "") then activity__MMColParam = Request.QueryString("f_address")
%>
<%
set email = Server.CreateObject("ADODB.Recordset")
email.ActiveConnection = Connect
email.Source = "SELECT ID_f_address, f_address FROM f_email"
email.CursorType = 0
email.CursorLocation = 3
email.LockType = 3
email.Open()
email_numRows = 0
%>
<%
numbers=1
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Field feedback report email edit. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function replace(string,text,by)
{
    var strLength = string.length, txtLength = text.length;
    if ((strLength == 0) || (txtLength == 0)) return string;
    var i = string.indexOf(text);
    if ((!i) && (text != string.substring(0,txtLength))) return string;
    if (i == -1) return string;
    var newstr = string.substring(0,i) + by;
    if (i+txtLength < strLength)
        newstr += replace(string.substring(i+txtLength,strLength),text,by);
    return newstr;
}

function trySubmit()
{
	if (document.forms[0].email.value.length<2)
	{
		alert("Sorry, you must enter an email address!\n(min. 2 characters)");
		return false;
	}

	if (confirm("Are you sure you want to update this email address?"))	{	document.forms[0].submit();
	return false;
	}
return false;
}

function exitpage()
{
	if (change==true)
	{
		if (confirm("You have changed at least one field on this page.\rBefore exiting this page, do you want to save those changes first?"))
		{
		return trySubmit();
		}
	}
	return true;
}
//-->
</script>
</HEAD>

<table width="100%" border="0" cellspacing="3" cellpadding="0">
  <tr>
    <td align="left" valign="bottom" class="heading"> Field feedback report - email address edit</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="3" cellpadding="0">
  <tr>
    <td align="left" valign="bottom">
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="add_user" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="150">Email:(use &quot;;&quot;
              as delimiter)</td>
            <td class="text" align="left" valign="top" colspan="3">
              <input type="text" name="email" onChange="change=true;" size="70" class="formitem1" value="<%=(email.Fields.Item("f_address").Value)%>">
            </td>
			<td class="text_table" align="left" valign="top" colspan="3">
				<input type="reset" name="Submit3" value="Reset" class="quiz_button">
				<input type="submit" name="Submit" value="Update" class="quiz_button" <%call IsEditOK%>>
			</td>
          </tr>
         <tr>
            <td class="text_table" align="left" valign="top" width="100">
              <input type="hidden" name="session" value="<%=getPassword(30, "", "true", "true", "true", "false", "true", "true", "true", "false")%>">
              <input type="hidden" name="current_export">
            </td>
		 </tr>
        </table>
        <input type="hidden" name="MM_update" value="true">
        <input type="hidden" name="MM_recordId" value="<%= email.Fields.Item("f_address").Value %>">
      </form>
    </td>
</table>
<% If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then %>
Feedback Email(s) have been updated
<% End if %>
<% email.Close() %>
</BODY>
</HTML>
<!-- END EDIT EMAIL -->

<!-- START FEEDBACK REPORT -->
<%
buss =  request("info1")
set info1 = Server.CreateObject("ADODB.Recordset")
info1.ActiveConnection = Connect
info1.Source = "SELECT * from q_info1 order by info1"
info1.CursorType = 0
info1.CursorLocation = 3
info1.LockType = 3
info1.Open()
info1_numRows = 0

if buss <> "" then
set info2 = Server.CreateObject("ADODB.Recordset")
info2.ActiveConnection = Connect
info2.Source = "SELECT * from q_info2 where info2_info1 =" & buss &" order by info2"
info2.CursorType = 0
info2.CursorLocation = 3
info2.LockType = 3
info2.Open()
info2_numRows = 0
'Response.Write info2.Source
end if

set info3 = Server.CreateObject("ADODB.Recordset")
info3.ActiveConnection = Connect
info3.Source = "SELECT * from q_info3 order by info3"
info3.CursorType = 0
info3.CursorLocation = 3
info3.LockType = 3
info3.Open()
info3_numRows = 0
%>
<script>
function checkform() {
	document.searchlogs.action="f_feedback_logs.asp"
	document.searchlogs.target="_self"
	document.searchlogs.submit()
}
</script>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Field feedback reports. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
</HEAD>
<BODY BGCOLOR=#FFCC00 TEXT=#000000 VLINK=#000000 LINK=#000000 leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="3" cellpadding="0">
  <tr>
    <td align="left" valign="bottom" class="heading">Field feedback reports</td>
  </tr>
  <tr>
    <td align="left" valign="bottom" class="text">
      <form name="searchlogs" method="post" action="?alt=results">
        <p>Filter all log entries with following keys (nothing for all records).</p>
        <table width="600" border="0" cellspacing="0" cellpadding="0" class="text">
          <tr>
            <td width="150">Start date:</td>
            <td class="text_table">
              <input type="text" name="fromdate" maxlength="19" class="formitem1" onDblClick="this.value='<%=cDateSQL(Now()-1)%>';"size="30" value="<%=request("fromdate")%>">
              (yyyy/mm/dd hh:mm:ss), doubleclick = Today - 1 day</td>
          </tr>
          <tr>
            <td>End date:</td>
            <td class="text_table">
              <input type="text" name="todate" maxlength="19" class="formitem1" onDblClick="this.value='<%=cDateSQL(Now())%>';" size="30" value="<%=request("todate")%>">
              (yyyy/mm/dd hh:mm:ss), doubleclick = Today</td>
          </tr>
          <tr>
            <td>User name:</td>
            <td>
              <input type="text" name="username" class="formitem1" size="50" value="<%=request("username")%>">
            </td>
          </tr>

          <tr>
            <td>Business Group:</td>
            <td>
              <select name="info1" class="formitem1" onchange=checkform();>
              <option value="0">--- All Business Groups ---</option>
                <%
					While (NOT info1.EOF)
					if cint(buss) = cint(info1.Fields.Item("ID_info1").Value) then
					%>
						<option value="<%=(info1.Fields.Item("ID_info1").Value)%>" selected><%=(info1.Fields.Item("info1").Value)%></option>
						<% else %>
						<option value="<%=(info1.Fields.Item("ID_info1").Value)%>"><%=(info1.Fields.Item("info1").Value)%></option>
					<%
					  end if
					  info1.MoveNext()
					Wend%>
              </select>
            </td>
          </tr>
          <tr>
            <td height=25>Site:</td>
            <td>
              <select name="info2" class="formitem1">
              <%if buss <> "" then%>
              <option value="0">--- All Business Sites ---</option>
                <% While (NOT info2.EOF) %>
					<option value="<%=(info2.Fields.Item("ID_info2").Value)%>"><%=(info2.Fields.Item("info2").Value)%></option>
				<%
				  info2.MoveNext()
				Wend
				else%>
					<option value="0"></option>
				<% end if %>
              </select>
            </td>
          </tr>
          <tr>
            <td>Business activity:</td>
            <td>
              <select name="info3" class="formitem1">
              <option value="0">--- All Business Activities ---</option>
                <%
					While (NOT info3.EOF)
					if cint(request("info3")) = cint(info3.Fields.Item("ID_info3").Value) then
					%>
					<option value="<%=(info3.Fields.Item("ID_info3").Value)%>" selected><%=(info3.Fields.Item("info3").Value)%></option>
					<% else %>
					<option value="<%=(info3.Fields.Item("ID_info3").Value)%>" ><%=(info3.Fields.Item("info3").Value)%></option>
					<%
					end if
					info3.MoveNext()
					Wend%>
              </select>
            </td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td>
              <input type=RESET name="Reset" value="Reset the form" class="quiz_button">
              <input type="submit" name="Submit" value="View filtered field feedback reports" class="quiz_button">
            </td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table>

<!-- FEEDBACK TABLE BEGIN -->
<%
if request.querystring("alt")="results" THEN

SQL_where = ""
SQL_pivot = ""
show_lines = 100

if cStr(Request("show_lines")) <> "" then show_lines = cInt(Request("show_lines"))

fromdate=Replace(cStr(Replace(Request("fromdate"), ".", "/")), "'", "''")
fromdate=cDateSQL(fromdate)
If fromdate <> "" then
	SQL_where = " WHERE (f_feedback.f_date) >= " + database_date_string + fromdate + database_date_string
end if

todate = Replace(cStr(Replace(Request("todate"), ".", "/")), "'", "''")
todate=cDateSQL(todate)
If todate <> "" then
	if SQL_where <> "" then
		SQL_where = SQL_where + " AND (f_feedback.f_date) <= " + database_date_string + todate + database_date_string
	else
		SQL_where = " WHERE (f_feedback.f_date) <= " + database_date_string + todate + database_date_string
	end if
end if

If cStr(Request("username")) <> "" then
	if SQL_where <> "" then
		SQL_where = SQL_where + " AND ((q_user.user_username) Like '%" + Replace(cStr(Request("username")), "'", "''") + "%' OR (q_user.user_firstname) Like '%" + Replace(cStr(Request("username")), "'", "''") + "%' OR (q_user.user_lastname) Like '%" + Replace(cStr(Request("username")), "'", "''") + "%' )"
	else
		SQL_where = " WHERE ((q_user.user_username) Like '%" + Replace(cStr(Request("username")), "'", "''") + "%' OR (q_user.user_firstname) Like '%" + Replace(cStr(Request("username")), "'", "''") + "%' OR (q_user.user_lastname) Like '%" + Replace(cStr(Request("username")), "'", "''") + "%' )"
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

%>

<%
set logs = Server.CreateObject("ADODB.Recordset")
logs.ActiveConnection = Connect
logs.Source = "SELECT f_feedback.ID_feedback, f_feedback.f_date, f_feedback.title, f_feedback.company, f_feedback.details, q_user.user_lastname, " &_
			   "q_user.user_firstname " + SQL_pivot + "FROM f_feedback INNER JOIN q_user ON f_feedback.ID_user = q_user.ID_user" + SQL_where + " GROUP BY " &_
               "f_feedback.ID_feedback, f_feedback.f_date, f_feedback.title, f_feedback.company, f_feedback.details, q_user.user_lastname, q_user.user_firstname " &_
               "ORDER BY f_feedback.f_date DESC, f_feedback.ID_feedback DESC;"
logs.CursorType = 0
logs.CursorLocation = 3
logs.LockType = 3
logs.Open()
logs_numRows = 0

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
Dim Repeat1__numRows
Repeat1__numRows = show_lines
Dim Repeat1__index
Repeat1__index = 0
logs_numRows = logs_numRows + Repeat1__numRows
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

    MM_rs.Requery

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

<table>
<%
	todate=request("todate")
	fromdate=request("fromdate")
	username=request("username")
	info1=request("info1")
	info2=request("info2")
	info3=request("info3")
%>
<td align="right" class="subheads" valign="top"><br><font face="Arial" size="1">Click on the image for Excel format : <a href="f_export_results.asp?todate=<%=todate%>&fromdate=<%=fromdate%>&username=<%=username%>&info1=<%=info1%>&info2=<%=info2%>&info3=<%=info3%>"><img src="images/xls.gif" width="16" height="16" border="0"></a>
</td>
</tr>
</table>
<table width="300" border="0" cellspacing="2" cellpadding="0">
  <tr>
    <td width="120"><font face="Arial" size="1">Starting date:</font></td>
    <td><b><font face="Arial" size="1"><%=fromdate%></font></b></td>
  </tr>
  <tr>
    <td><font face="Arial" size="1">Ending date:</font></td>
    <td><b><font face="Arial" size="1"><%=todate%></font></b></td>
  </tr>
  <tr>
    <td><font face="Arial" size="1">User name:</font></td>
    <td><b><font face="Arial" size="1"><%=username%></font></b></td>
  </tr>
  <tr>
    <td><font face="Arial" size="1">Business Group:</font></td>
    <td><b><font face="Arial" size="1">
    <%
    if info1 <> 0 then
		Response.Write info11.Fields.item("info1").value
		info11.close
	else
		Response.Write ("All")
	end if
	%>
    </font></b></td>
  </tr>
  <tr>
    <td><font face="Arial" size="1">Business Site:</font></td>
    <td><b><font face="Arial" size="1">
    <%
    if info2 <> 0 then
		Response.Write info22("info2")
		info22.close
	else
		Response.Write ("All")
	end if
	%>
    </font></b></td>
  </tr>
  <tr>
    <td><font face="Arial" size="1">Business Activity:</font></td>
    <td><b><font face="Arial" size="1">
    <%
    if info3 <> 0 then
		Response.Write info33("info3")
		info33.close
	else
		Response.Write ("All")
	end if
	%>
    </font></b></td>
  </tr>
</table><br><br>
<table width="97%" border="0" cellspacing="0" cellpadding="0" bgcolor="#999999">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="1" cellpadding="1">
        <tr bgcolor="#333333">
          <td width="3%"><font color="#FFFFFF" face="Arial" size="1">ID</font></td>
          <td width="5%"><font color="#FFFFFF" face="Arial" size="1">Date</font></td>
          <td width="5%"><font color="#FFFFFF" face="Arial" size="1">Lodged By</font></td>
          <td width="10%"><font color="#FFFFFF" face="Arial" size="1">Details</font></td>
		</tr>
        <%
While ((Repeat1__numRows <> 0) AND (NOT logs.EOF))
%>
        <tr bgcolor="#FFFFFF">
          <td><font size="1"><%=logs.Fields.Item("ID_feedback").Value%></font></td>
          <td><font size="1"><%=logs.Fields.Item("f_date").Value%></font></td>
          <td><font size="1"><%=logs.Fields.Item("user_firstname").Value%>&nbsp;<%=logs.Fields.Item("user_lastname").Value%></font></td>
          <td><font size="1"><%=logs.Fields.Item("details").Value%></font></td>
        </tr>
        <%
		  Repeat1__index=Repeat1__index+1
		  Repeat1__numRows=Repeat1__numRows-1
		  logs.MoveNext()
		Wend
		%>
        <tr bgcolor="#FFFFFF">
          <% If logs.EOF And logs.BOF Then %>
          <td colspan="9"><font size="2">Sorry,
            there are currently no LOG entries or your search criteria do not
            match with any loged records...</font></td>
          <% End If %>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="3">
      <table border="0" width="50%" align="center" cellpadding="0" cellspacing="1">
        <tr bgcolor="#999999">
          <td width="25%" align="center"> <b><font face="Arial" size="2">
            <% If MM_offset <> 0 Then %>
            <a href="<%=MM_moveFirst%>">First</a>
            <% End If %>
            </font></b></td>
          <td width="25%" align="center"> <b><font face="Arial" size="2">
            <% If MM_offset <> 0 Then %>
            <a href="<%=MM_movePrev%>">Previous</a>
            <% End If %>
            </font></b></td>
          <td width="25%" align="center"> <b><font face="Arial" size="2">
            <% If Not MM_atTotal Then %>
            <a href="<%=MM_moveNext%>">Next</a>
            <% End If  %>
            </font></b></td>
          <td width="25%" align="center"> <b><font face="Arial" size="2">
            <% If Not MM_atTotal Then %>
            <a href="<%=MM_moveLast%>">Last</a>
            <% End If %>
            </font></b></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr align="center">
    <td colspan="3"><font face="Arial" size="2"><i>&nbsp;
      </i></font>

	  <form name="form1" method="post" action="">
        <font face="Arial" size="2"><i>
        <input type="hidden" name="fromdate" value="<%=Request("fromdate")%>">
        <input type="hidden" name="todate" value="<%=Request("todate")%>">
        <input type="hidden" name="username" value="<%=Request("username")%>">
        <input type="hidden" name="info1" value="<%=Request("info1")%>">
        <input type="hidden" name="info2" value="<%=Request("info2")%>">
        <input type="hidden" name="info3" value="<%=Request("info3")%>">

        Records <%=(logs_first)%> to <%=(logs_last)%> of <%=(logs_total)%> (Show
        <input type="text" name="show_lines" size="3" maxlength="3" value="<%=show_lines%>">
        lines per page - <a href="javascript:document.forms[0].submit();">recompose</a>)</i></font>
      </form>
      <font face="Arial" size="2"><i> </i></font></td>
  </tr>
  <tr align="center">
    <td width="33%" align="left"><font face="Arial" size="2">Log generated on: <%= Now()%></font></td>
    <td width="33%"><font face="Arial" size="2">Application name: <%=client_name_long%></font></td>
    <td width="33%" align="right"><font face="Arial" size="2">Generated from IP: <%=Request.ServerVariables("REMOTE_ADDR")%></font></td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
</body>
</html>
<%
logs.Close()
%>	

<%
	call log_the_page ("Feedback Reports - Results")
End if

call log_the_page ("Feedback Reports")
%>
