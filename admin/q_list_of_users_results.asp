<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
numbers=1
count = 1
SQL_having = ""
SQL_where = ""
if Request.Cookies("show_lines")<> "" then
	show_lines= cint(Request.Cookies("show_lines"))
else
	show_lines=15
end if
results=request("results")
fromdate=request("fromdate")
fromdate=cdatesql(fromdate)
todate=request("todate")
if len(todate) < 12 and todate <> "" then
	todate=todate&" 23:59:59"
end if
todate=cdatesql(todate)
active = request("active")

mths = request("mths")
if mths="" then
	mths=0
end if
noquiz = request("noquiz")

default_passrate=80

if cStr(Request.Querystring("show_lines")) <> "" then show_lines = cInt(Request.Querystring("show_lines"))

If cStr(Request.Querystring("filter_username")) <> "" then
	SQL_having = " HAVING ((q_user.user_username) Like '%" + Replace(uCase(cStr(Request.Querystring("filter_username"))), "'", "''") + "%' OR  (q_user.user_firstname) Like '%" + Replace(uCase(cStr(Request.Querystring("filter_username"))), "'", "''") + "%') "
end if

subject_prm = 0

If cInt(Request.Querystring("subject")) <> 0 then
	subject_prm = cInt(Request.Querystring("subject"))
	SQL_where = " WHERE (q_session.Session_subject = " + (Request.Querystring("subject")) + ") "
	if cstr(request("active"))="1" then
		SQL_WHERE=SQL_where + "and q_user.user_active=1"
	elseif cstr(request("active"))="0" then
		SQL_WHERE=SQL_where + "and q_user.user_active=0"
	end if
	if clng(request("status"))=1 then
		SQL_WHERE=SQL_where + "and q_user.user_status=0"
	elseif clng(request("status"))=2 then
		SQL_WHERE=SQL_where + "and q_user.user_status=1"
	end if
	
	sql_where_nodate = sql_where

	if cstr(fromdate) <> "" and cstr(todate) = "" then
		SQL_WHERE = sql_where + "and q_session.session_finish >='"&fromdate&"'"
	else if cstr(todate) <> "" and cstr(fromdate) = "" then
		SQL_WHERE = sql_where + "and q_session.session_finish <='"&todate&"'"
	else if cstr(fromdate) <> "" and cstr(todate) <> "" then
		SQL_WHERE = sql_where + "and q_session.session_finish between '"&fromdate&"' and '"&todate&"'"
	end if
	end if
	end if
else
	if cstr(request("active"))="1" then
		sql_where="where  q_user.user_active=1"
	elseif cstr(request("active"))="0" then
		sql_where="where  q_user.user_active=0"
	ELSE
		sql_where="where  1 = 1"
	end if
	
	if clng(request("status"))=1 then
		SQL_WHERE=SQL_where + "and q_user.user_status=0"
	elseif clng(request("status"))=2 then
		SQL_WHERE=SQL_where + "and q_user.user_status=1"
	end if

	sql_where_nodate = sql_where

	if cstr(fromdate) <> "" and cstr(todate)= "" then
		if request("active")<>"2" then
			SQL_WHERE = sql_where + "and session_finish >='"&fromdate&"'"
		else
			SQL_WHERE = "where q_session.session_finish >= '"&fromdate&"'"
		end if
	else if cstr(todate) <> "" and cstr(fromdate) ="" then
		if request("active")<>"2" then
			SQL_WHERE = sql_where + "and session_finish <= '"&todate&"'"
		else
			SQL_WHERE = "where  session_finish <='"&todate&"'"
		end if
	else if cstr(fromdate) <> "" and cstr(todate) <> "" then
		if request("active")<>"2" then
			SQL_WHERE = sql_where + "and q_session.session_finish <='"&todate&"' and session_finish >='"&fromdate&"'"
		else
			SQL_WHERE = "where q_session.session_finish between '"&fromdate&"' and '"&todate&"'"
		end if
	end if
	end if
	end if
end if

filter_info1_prm = 0
If cInt(Request.Querystring("filter_info1")) <> 0 then
	filter_info1_prm = cInt(Request.Querystring("filter_info1"))
	if SQL_having <> "" then
		SQL_having = SQL_having + " AND (q_user.user_info1)= " + (Request.Querystring("filter_info1")) + " "
	else
		SQL_having = " HAVING (q_user.user_info1)= " + (Request.Querystring("filter_info1")) + " "
	end if
end if

filter_info2_prm = 0
If cInt(Request.Querystring("filter_info2")) <> 0 then
	filter_info2_prm = cInt(Request.Querystring("filter_info2"))
	if SQL_having <> "" then
		SQL_having = SQL_having + " AND (q_user.user_info2)= " + (Request.Querystring("filter_info2")) + " "
	else
		SQL_having = " HAVING (q_user.user_info2)= " + (Request.Querystring("filter_info2")) + " "
	end if
end if

'default is 0 ie: all rights
admin_logged_in_info1=0
set admin_info1 = Server.CreateObject("ADODB.Recordset")
admin_info1.ActiveConnection = Connect
admin_info1.Source = "SELECT admin_info1 FROM admin where admin_name='"&cstr(Admin_logged_in)&"';"
admin_info1.CursorType = 0
admin_info1.CursorLocation = 3
admin_info1.LockType = 3
'admin_info1.Open()
admin_info1_numRows = 0

'While (NOT admin_info1.EOF)
'	admin_logged_in_info1=admin_info1.Fields.Item("admin_info1").Value
'	admin_info1.MoveNext()
'Wend

'admin_info1.Close()

set filter_info2 = Server.CreateObject("ADODB.Recordset")
filter_info2.ActiveConnection = Connect

if request("filter_info1")<> "" then
	info2_prm = request("filter_info1")
else
	info2_prm = admin_logged_in_info1
end if
filter_info2.Source = "SELECT * FROM q_info2 where info2_info1 =" & info2_prm &" and info2_active=1 order by info2"
filter_info2.CursorType = 0
filter_info2.CursorLocation = 3
filter_info2.LockType = 3
filter_info2.Open()
filter_info2_numRows = 0

filter_info3_prm = 0
If cInt(Request.Querystring("filter_info3")) <> 0 then
	filter_info3_prm = cInt(Request.Querystring("filter_info3"))
	if SQL_having <> "" then
		SQL_having = SQL_having + " AND (q_user.user_info3)= " + (Request.Querystring("filter_info3")) + " "
	else
		SQL_having = " HAVING (q_user.user_info3)= " + (Request.Querystring("filter_info3")) + " "
	end if
end if

if request("mths")=1 then
	session("mths") = 1
else
	session("mths")=""
end if

set users = Server.CreateObject("ADODB.Recordset")
users.ActiveConnection = Connect

set filter_info1 = Server.CreateObject("ADODB.Recordset")
filter_info1.ActiveConnection = Connect


'users.Source = "SELECT q_user.ID_user, q_user.user_lastname, q_user.user_firstname, q_user.user_username, q_info1.info1, q_info2.info2, q_info3.info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3, COUNT(q_session.ID_session) AS session_count FROM (q_info3 RIGHT JOIN (q_info2 RIGHT JOIN (q_info1 RIGHT JOIN q_user ON q_info1.ID_info1 = q_user.user_info1) ON q_info2.ID_info2 = q_user.user_info2) ON q_info3.ID_info3 = q_user.user_info3) LEFT JOIN q_session ON q_user.ID_user = q_session.Session_users " + SQL_where + " GROUP BY q_user.user_lastname, q_user.user_firstname, q_user.ID_user, q_info1.info1, q_info2.info2, q_info3.info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info2, q_user.user_info3 " + SQL_having + " ORDER BY q_user.user_lastname, q_user.user_firstname;"

if admin_logged_in_info1>0 then
	if SQL_having <> "" then
		SQL_having = SQL_having + " AND (q_user.user_info1)=  "&admin_logged_in_info1
	else
		SQL_having = " HAVING (q_user.user_info1)=  " &admin_logged_in_info1
	end if
	users.Source = "SELECT q_user.user_status, q_user.ID_user, q_user.user_lastname, q_user.user_firstname, q_user.user_username, q_info1.info1,q_info1.ID_info1, q_info2.info2, q_info3.info3,q_info3.ID_info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3, COUNT(q_session.ID_session) AS session_count FROM (q_info3 RIGHT JOIN (q_info2 RIGHT JOIN (q_info1 RIGHT JOIN q_user ON q_info1.ID_info1 = q_user.user_info1) ON q_info2.ID_info2 = q_user.user_info2) ON q_info3.ID_info3 = q_user.user_info3) LEFT JOIN q_session ON q_user.ID_user = q_session.Session_users " + SQL_where + " GROUP BY q_user.user_status, q_user.user_lastname, q_user.user_firstname, q_user.user_username, q_user.ID_user, q_info1.info1,q_info1.ID_info1, q_info2.info2, q_info3.info3,q_info3.ID_info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info2, q_user.user_info3 " + SQL_having + " ORDER BY q_user.user_lastname, q_user.user_firstname;"
	if (cstr(noquiz)="1") then
		users.Source = "SELECT q_user.user_status, q_user.ID_user, q_user.user_lastname, q_user.user_firstname, q_user.user_username, q_info1.info1,q_info1.ID_info1, q_info2.info2, q_info3.info3,q_info3.ID_info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3, 0 as session_count FROM (q_info3 RIGHT JOIN (q_info2 RIGHT JOIN (q_info1 RIGHT JOIN q_user ON q_info1.ID_info1 = q_user.user_info1) ON q_info2.ID_info2 = q_user.user_info2) ON q_info3.ID_info3 = q_user.user_info3)  " + sql_where_nodate + " GROUP BY q_user.user_status, q_user.user_lastname, q_user.user_firstname, q_user.ID_user, q_info1.info1,q_info1.ID_info1, q_info2.info2, q_info3.info3,q_info3.ID_info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info2, q_user.user_info3 " + SQL_having + " ORDER BY q_user.user_lastname, q_user.user_firstname;"
	end if
	filter_info1.Source = "SELECT * FROM q_info1 where info1_active=1 and id_info1="&admin_logged_in_info1 &" order by info1"
	bus_sel="selected"
else
	users.Source = "SELECT q_user.user_status, q_user.ID_user, q_user.user_lastname, q_user.user_firstname, q_user.user_username, q_info1.info1,q_info1.ID_info1, q_info2.info2, q_info3.info3,q_info3.ID_info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3, COUNT(q_session.ID_session) AS session_count FROM (q_info3 RIGHT JOIN (q_info2 RIGHT JOIN (q_info1 RIGHT JOIN q_user ON q_info1.ID_info1 = q_user.user_info1) ON q_info2.ID_info2 = q_user.user_info2) ON q_info3.ID_info3 = q_user.user_info3) LEFT JOIN q_session ON q_user.ID_user = q_session.Session_users " + SQL_where + " GROUP BY q_user.user_status, q_user.user_lastname, q_user.user_firstname, q_user.user_username, q_user.ID_user, q_info1.info1,q_info1.ID_info1, q_info2.info2, q_info3.info3,q_info3.ID_info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info2, q_user.user_info3 " + SQL_having + " ORDER BY q_user.user_lastname, q_user.user_firstname;"
	if (cstr(noquiz)="1") then
			users.Source = "SELECT q_user.user_status, q_user.ID_user, q_user.user_lastname, q_user.user_firstname, q_user.user_username, q_info1.info1,q_info1.ID_info1, q_info2.info2, q_info3.info3,q_info3.ID_info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3, 0 as session_count FROM (q_info3 RIGHT JOIN (q_info2 RIGHT JOIN (q_info1 RIGHT JOIN q_user ON q_info1.ID_info1 = q_user.user_info1) ON q_info2.ID_info2 = q_user.user_info2) ON q_info3.ID_info3 = q_user.user_info3)  " + sql_where_nodate + " GROUP BY q_user.user_status, q_user.user_lastname, q_user.user_firstname, q_user.user_username,  q_user.ID_user, q_info1.info1,q_info1.ID_info1, q_info2.info2, q_info3.info3,q_info3.ID_info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info2, q_user.user_info3 " + SQL_having + " ORDER BY q_user.user_lastname, q_user.user_firstname;"
	end if
	filter_info1.Source = "SELECT * FROM q_info1 where info1_active=1 order by info1"
end if


users.CursorType = 0
users.CursorLocation = 3
users.LockType = 3
users.Open()
users_numRows = 0

filter_info1.CursorType = 0
filter_info1.CursorLocation = 3
filter_info1.LockType = 3
filter_info1.Open()
filter_info1_numRows = 0

set filter_info3 = Server.CreateObject("ADODB.Recordset")
filter_info3.ActiveConnection = Connect
filter_info3.Source = "SELECT * FROM q_info3 where info3_active=1 order by info3"
filter_info3.CursorType = 0
filter_info3.CursorLocation = 3
filter_info3.LockType = 3
filter_info3.Open()
filter_info3_numRows = 0

set subjects = Server.CreateObject("ADODB.Recordset")
subjects.ActiveConnection = Connect
subjects.Source = "SELECT ID_subject, subject_name FROM subjects where subject_active_q <> 0"
subjects.CursorType = 0
subjects.CursorLocation = 3
subjects.LockType = 3
subjects.Open()
subjects_numRows = 0

set user_details = Server.CreateObject("ADODB.Recordset")
user_details.ActiveConnection = Connect
user_details.CursorType = 0
user_details.CursorLocation = 3
user_details.LockType = 3
user_details_numRows = 0

Dim Repeat1__numRows
Repeat1__numRows = show_lines
Dim Repeat1__index
Repeat1__index = 0
users_numRows = users_numRows + Repeat1__numRows

'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
users_total = users.RecordCount

' set the number of rows displayed on this page
If (users_numRows < 0) Then
  users_numRows = users_total
Elseif (users_numRows = 0) Then
  users_numRows = 1
End If

' set the first and last displayed record
users_first = 1
users_last  = users_first + users_numRows - 1

' if we have the correct record count, check the other stats
If (users_total <> -1) Then
  If (users_first > users_total) Then users_first = users_total
  If (users_last > users_total) Then users_last = users_total
  If (users_numRows > users_total) Then users_numRows = users_total
End If

' *** Recordset Stats: if we don't know the record count, manually count them

If (users_total = -1) Then

  ' count the total records by iterating through the recordset
  users_total=0
  While (Not users.EOF)
    users_total = users_total + 1
    users.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (users.CursorType > 0) Then
'    users.MoveFirst
'  Else
    users.Requery
  End If

  ' set the number of rows displayed on this page
  If (users_numRows < 0 Or users_numRows > users_total) Then
    users_numRows = users_total
  End If

  ' set the first and last displayed record
  users_first = 1
  users_last = users_first + users_numRows - 1
  If (users_first > users_total) Then users_first = users_total
  If (users_last > users_total) Then users_last = users_total

End If

' *** Move To Record and Go To Record: declare variables

Set MM_rs    = users
MM_rsCount   = users_total
MM_size      = users_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If

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

' *** Move To Record: update recordset stats

' set the first and last displayed record
users_first = MM_offset + 1
users_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (users_first > MM_rsCount) Then users_first = MM_rsCount
  If (users_last > MM_rsCount) Then users_last = MM_rsCount
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
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
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz users. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function checkmths()
{
	if (((document.filter_users.fromdate.value != "") && (document.filter_users.todate.value != "")) || ((document.filter_users.fromdate.value == "") && (document.filter_users.todate.value != "")) || ((document.filter_users.fromdate.value != "") && (document.filter_users.todate.value == "")))
	{
		document.filter_users.mths.checked = false;
		document.filter_users.mths.disabled=true;
		return;
	}
	else
	{
		document.filter_users.mths.disabled=false;
		return;
	}
}

var MyCookie = {
    Write:function(name,value,days) {
        var D = new Date();
        D.setTime(D.getTime()+86400000*days);
        document.cookie = escape(name)+"="+escape(value)+
            ((days == null)?"":(";expires="+D.toGMTString()))
        return (this.Read(name) == value);
    },
    Read:function(name) {
        var EN=escape(name)
        var F=' '+document.cookie+';', S=F.indexOf(' '+EN);
        return S==-1 ? null : unescape(F.substring(EN=S+EN.length+2,F.indexOf(';',EN)));
    }
}

function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}

function check()
{
	//PN 050506 pass rate will now come from the pass_rate table in the database
	/*passrate = MyCookie.Read('passrate')
	if (passrate != null) {
		document.forms[0].passrate.value=passrate;
	}
	else {
		document.forms[0].passrate.value=50;
	}
	*/
	show_lines = MyCookie.Read('show_lines');
	if (show_lines != null) {
		//alert('1');
		document.forms[0].show_lines.value=show_lines;
	}
	else {
		//alert('2');
		document.forms[0].show_lines.value=500;
	}
}

function AddCookieId(cn,id) {
        MyCookie.Write(cn,id,7);
}

function DelCookieId(cn,id) {
        MyCookie.Write(cn,id,-1);
}
// PN 050506 passrate will now come from the database table pass_rate
/*function pass_submit()
{
	if (isNaN(document.filter_users.passrate.value)){
		alert('Invalid pass rate');
		document.filter_users.passrate.focus();
		return false;
	}
	else
	{
		AddCookieId("passrate",document.filter_users.passrate.value);
		passrate = MyCookie.Read('passrate');
		document.forms[0].submit();
		return true;
	}
}
*/
function filter_submit()
{
	if (isNaN(document.filter_users.show_lines.value)){
		alert('Invalid number');
		document.filter_users.show_lines.focus();
		return false;
	}
	else
	{
		AddCookieId("show_lines",document.filter_users.show_lines.value);
		show_lines = MyCookie.Read('show_lines');
		document.forms[0].action="q_list_of_users_results.asp?filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&fromdate=<%=fromdate%>&todate=<%=todate%>&active=<%=active%>&results=<%=results%>&passrate=<%=passrate%>&mths=<%=mths%>&noquiz=<%=noquiz%>";
		document.forms[0].submit();
		return true;
	}
}
function clearform()
{
//document.forms[0].filter_username.value = "";
//document.forms[0].todate.value = "";
//document.forms[0].fromdate.value = "";
//document.forms[0].results.selectedIndex = 0;
//document.forms[0].subject.selectedIndex = 0;
//document.forms[0].active.selectedIndex = 0;
//document.forms[0].filter_info1.selectedIndex = 0;
//document.forms[0].filter_info3.selectedIndex = 0;
//document.forms[0].show_lines.value = "25";
document.forms[0].submit();
}

function checkform() {
	document.forms[0].action="q_list_of_users.asp";
	document.forms[0].target="_self";
	document.forms[0].submit();
}
//-->
</script>
</HEAD>

<BODY onload="check();">
	<%
	'PN 050506 passrate now comes from the database table pass_rate
	'if Request.Cookies("passrate")<> "" then
		'passrate= cint(Request.Cookies("passrate"))

	'else
		'passrate=50
	'end if

	if Request.Cookies("show_lines")<> "" then
		show_lines= cint(Request.Cookies("show_lines"))
	else
		show_lines=15
	end if
%>
<table>
  <tr>
    <td class="heading"> Quiz users</td>
  </tr>
  <tr>
    <td align="left" valign="bottom">
      <form name="filter_users" action="q_list_of_users_results.asp" method="post">
     <input type="hidden" name="hiddenmths" value="false">
        <table>
          <tr>
            <td class="subheads" align="left" valign="top">Users:</td>
              <td align="right" class="subheads" valign="top"><a href="q_export_quiz_users.asp?filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&fromdate=<%=fromdate%>&todate=<%=todate%>&active=<%=active%>&results=<%=results%>&passrate=<%=passrate%>&mths=<%=mths%>&noquiz=<%=noquiz%>&status=<%=request("status")%>"><img src="images/xls.gif" width="16" height="16" border="0"> export to Excel</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="q_export_quiz_users_summary.asp?filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&fromdate=<%=fromdate%>&todate=<%=todate%>&active=<%=active%>&results=<%=results%>&passrate=<%=passrate%>&mths=<%=mths%>&noquiz=<%=noquiz%>&status=<%=request("status")%>"><img src="images/summary.gif" width="16" height="16" border="0"> export summary</a></td>
          </tr>
		</table>
		<table>
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">
            <td class="text" width="18" ><img src="images/back.gif" width="18" height="14"></td>
            <td class="text" colspan="12"><a href="q_list_of_users.asp?filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&fromdate=<%=request("fromdate")%>&todate=<%=request("todate")%>&active=<%=request("active")%>&results=<%=request("results")%>&mths=<%=request("mths")%>&noquiz=<%=request("noquiz")%>&status=<%=request("status")%>">...Filters</a></td>
          </tr>
          <tr>
            <td >&nbsp;</td>
            <td >Username (email) &amp; First name</td>
            <td >Business &amp; Site</td>
            <td >Activity</td>
            <td >Active</td>
            <!--<td >Logs</td>-->
            <td >Sess.</td>
            <td >Avg Rate</td>
			<td >Passes</td>
			<td >Fails</td>
            <td >Edit</td>
            <!-- CXS 061122: emailer functionality -->
			<td >Email</td>
			<td >Merge</td>
			<td >Add</td>
          </tr>
          <%
	 If Not users.EOF Or Not users.BOF Then
While ((Repeat1__numRows <> 0) AND (NOT users.EOF))

if subject_prm <> 0 then
	subj_prm ="and (q_session.Session_subject ="&subject_prm&")"
else
	subj_prm =""
end if

if cstr(fromdate)="" and cstr(todate) <> "" then
	user_details.Source = "SELECT q_session.id_session, q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.Session_done, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) and (session_finish <= '"&todate&"') "&subj_prm&" order by session_subject,session_date desc"
else if cstr(todate)="" and cstr(fromdate) <> "" then
	user_details.Source = "SELECT q_session.id_session, q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.Session_done, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) and (session_finish >= '"&fromdate&"') "&subj_prm&" order by session_subject,session_date desc"
else if (cstr(todate)="" and cstr(fromdate)="") then
	user_details.Source = "SELECT q_session.id_session, q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.Session_done,q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) "&subj_prm&" order by session_subject,session_date desc"
else
	user_details.Source = "SELECT q_session.id_session, q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.Session_done,q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) and ((session_finish >= '"&fromdate&"') and (session_finish <= '"&todate&"')) "&subj_prm&" order by session_subject,session_date desc"
end if
end if
end if
user_details.Open()
user_details_numRows = 0
'Response.Write user_details.source
user_session_rate = 0
user_session_count = 0
user_total_rate = 0
subid = 0
'PN 050507 variables to store user passes and fails
user_passes=0
user_fails=0
user_session_percentage=0
While (NOT user_details.EOF)
'PN050506 set up the variable that stores the passrate for the session
subject_pass_rate_percentage=default_passrate
if session("mths")="" then
	user_session_correct = cInt(user_details.Fields.Item("session_correct").Value)
	user_session_total = cInt(user_details.Fields.Item("session_total").Value)
	user_session_rate = user_session_rate + (user_session_correct / user_session_total * 100)
	user_session_percentage=(user_session_correct / user_session_total * 100)
	user_session_count = user_session_count + 1
else if cint(subid) <> cInt(user_details.Fields.Item("session_subject").Value) then
	subid = cInt(user_details.Fields.Item("session_subject").Value)
	user_session_correct = cInt(user_details.Fields.Item("session_correct").Value)
	user_session_total = cInt(user_details.Fields.Item("session_total").Value)
	user_session_rate = user_session_rate + (user_session_correct / user_session_total * 100)
	user_session_percentage=(user_session_correct / user_session_total * 100)
	user_session_count = user_session_count + 1
end if
end if
	'PN 050506 work out what the passrate will be for this users session
	'ID_info1 and ID_info3 are newly selected columns in the above query to get the IDs


	subject_pass_rate_percentage = default_passrate
	if passrate_type = 1 then
		set pass_rate_query = Server.CreateObject("ADODB.Recordset")
		pass_rate_query.ActiveConnection = Connect
		pass_rate_query.Source = "SELECT * FROM q_certification where q_certification.q_session = "&user_details.Fields.Item("ID_Session").Value&""
		pass_rate_query.CursorType = 0
		pass_rate_query.CursorLocation = 3
		pass_rate_query.LockType = 3
		pass_rate_query.Open()
		if (not pass_rate_query.eof) then
			subject_pass_rate_percentage = pass_rate_query.Fields.Item("percentage_required").Value
		else
			set subject = Server.CreateObject("ADODB.Recordset")
			subject.ActiveConnection = Connect
			subject.Source = "SELECT * FROM subjects where subjects.id_subject = "&user_details.Fields.Item("session_subject").Value&""
			subject.CursorType = 0
			subject.CursorLocation = 3
			subject.LockType = 3
			subject.Open()
			subject_pass_rate_percentage = subject.fields.item("subject_passmark").value
			subject.Close()
		end if
		pass_rate_query.Close()
	elseif passrate_type=2 then
		set pass_rate_query = Server.CreateObject("ADODB.Recordset")
		pass_rate_query.ActiveConnection = Connect
		pass_rate_query.Source = "SELECT * FROM q_certification where q_certification.q_session = "&user_details.Fields.Item("ID_Session").Value&""
		pass_rate_query.CursorType = 0
		pass_rate_query.CursorLocation = 3
		pass_rate_query.LockType = 3
		pass_rate_query.Open()
		if (not pass_rate_query.eof) then
			subject_pass_rate_percentage = pass_rate_query.Fields.Item("percentage_required").Value
		else
			set pass_rate_query2 = Server.CreateObject("ADODB.Recordset")
			pass_rate_query2.ActiveConnection = Connect
			pass_rate_query2.Source = "SELECT pass_rate FROM pass_rates where subject="&user_details.Fields.Item("session_subject").Value&" and q_info1="&users.Fields.Item("user_info1").Value&" and q_info3="&users.Fields.Item("user_info3").Value&";"
			pass_rate_query2.CursorType = 0
			pass_rate_query2.CursorLocation = 3
			pass_rate_query2.LockType = 3
			pass_rate_query2.Open()
			if (not pass_rate_query2.eof) then
				subject_pass_rate_percentage = pass_rate_query2.Fields.Item("pass_rate").Value
			end if
			pass_rate_query2.Close()
		end if
		pass_rate_query.Close()
	elseif passrate_type=3 then
		set pass_rate_query = Server.CreateObject("ADODB.Recordset")
		pass_rate_query.ActiveConnection = Connect
		pass_rate_query.Source = "SELECT pass_rate FROM pass_rates where subject="&user_details.Fields.Item("session_subject").Value&" and q_info1="&users.Fields.Item("user_info1").Value&" and q_info3="&users.Fields.Item("user_info3").Value&";"
		pass_rate_query.CursorType = 0
		pass_rate_query.CursorLocation = 3
		pass_rate_query.LockType = 3
		pass_rate_query.Open()
		if (not pass_rate_query.eof) then
			subject_pass_rate_percentage = pass_rate_query.Fields.Item("pass_rate").Value
		end if
		pass_rate_query.Close()
	else
		set subject = Server.CreateObject("ADODB.Recordset")
		subject.ActiveConnection = Connect
		subject.Source = "SELECT * FROM subjects where subjects.id_subject = "&user_details.Fields.Item("session_subject").Value&""
		subject.CursorType = 0
		subject.CursorLocation = 3
		subject.LockType = 3
		subject.Open()
		subject_pass_rate_percentage = subject.fields.item("subject_passmark").value
		subject.Close()
	end if

	'PN050506 calulate if this session is a pass or a fail
	'Response.Write("the user_session_rate is  "&user_session_percentage&"   and the passrate is "&subject_pass_rate_percentage&"<br><br>")
	if(cInt(user_session_percentage)>=cInt(subject_pass_rate_percentage)) then
		'it was a pass
		user_passes = user_passes + 1
		total_passes=total_passes+1
	else
		'it was a fail
		user_fails = user_fails + 1
		total_fails=total_fails+1
	end if

user_details.MoveNext()
Wend

user_details.Close()
if user_session_count > 0 then
	user_total_rate = (user_session_rate/user_session_count)
end if

if (cstr(noquiz)="1") then
	if user_session_count = 0 then

	'if cInt(users.Fields.Item("session_count").Value) = 0 then
%>
	<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">

            <td class="text" width="20"><%=count%></td>
            <td width="200" class="text">
             <%=(users.Fields.Item("user_username").Value)%>&nbsp;<%=(users.Fields.Item("user_firstname").Value)%></td>
            <td width="140" class="text"><%=(users.Fields.Item("info1").Value)%> (<%=(users.Fields.Item("info2").Value)%>)</td>
            <td width="140" class="text"><%=(users.Fields.Item("info3").Value)%></td>
            <td width="30" class="text" align=center>
              <%if abs(users.Fields.Item("user_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
            </td>
            <!--<td width="20" class="text"><%=(users.Fields.Item("user_logcount").Value)%>x</td>-->
            <td width="20" class="text"><% if cInt(users.Fields.Item("session_count").Value)<>"" then%> <%=(users.Fields.Item("session_count").Value)%><%else%>0<%end if%></td>
            <td width="20" class="text">
              <%
			response.write("<font color = blue>N/A</font>")
			count = count + 1
			%>
            </td>
			<td width="20" class="text"><font color = green><%=user_passes%></font></td>
			<td width="20" class="text"><font color = red><%=user_fails%></font></td>
            <td  align="right" width="20">
              <a href="q_user_edit.asp?user=<%=(users.Fields.Item("ID_user").Value)%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&show_lines=<%=show_lines%>"><img src="images/edit.gif" width="16" height="15" border="0"></a>
            </td>
            <!-- CXS 061122: emailer functionality -->
				<td  align="right" width="18">
				<a href="send_individual_email.asp?individual=<%=(users.Fields.Item("ID_user").Value)%>"><img src="images/unread.gif" alt="Send reminder email to this user" border="0"></a></td>
				<!--PN 040811 add merge facility so that users can be merged into this user to get rid of self reg duplicates-->
				<td  align="right" width="18">
				<a href="#" onclick="var wintoopen=window.open('q_list_of_users_to_merge.asp?user=<%=(users.Fields.Item("ID_user").Value)%>','merge','toolbar=0, scrollbars=yes,resizable=1,width=700, height=500');wintoopen.focus();"><img src="images/merge.gif" alt="Merge users with this user" width="15" height="15" border="0"></a>
			</td>
				<td  align="right" width="18">
				<a href="#" onclick="var wintoopen=window.open('q_user_addresult.asp?user=<%=(users.Fields.Item("ID_user").Value)%>','merge','toolbar=0, scrollbars=yes,resizable=1,width=700, height=500');wintoopen.focus();"><img src="images/addresults.gif" alt="Add results" border="0"></a>
			</td>
          </tr>
<%
end if
else

if cstr(results)="" or cstr(results)="2" then
%>
			<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">

            <td class="text" width="20"><%=numbers%></td>
            <td width="200" class="text">
              <%if cInt(users.Fields.Item("session_count").Value) > 0 then%>
              <a href="q_user_sessions.asp?user=<%=(users.Fields.Item("ID_user").Value)%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&fromdate=<%=fromdate%>&todate=<%=todate%>&active=<%=active%>&results=<%=results%>&passrate=<%=passrate%>&mths=<%=mths%>&noquiz=<%=noquiz%>&status=<%=request("status")%>">
              <%end if%>
              <%=(users.Fields.Item("user_username").Value)%>&nbsp;<%=(users.Fields.Item("user_firstname").Value)%></a></td>
            <td width="140" class="text"><%=(users.Fields.Item("info1").Value)%> (<%=(users.Fields.Item("info2").Value)%>)</td>
            <td width="140" class="text"><%=(users.Fields.Item("info3").Value)%></td>
            <td width="30" class="text" align=center>
              <%if abs(users.Fields.Item("user_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
            </td>
            <!--<td width="20" class="text"><%=(users.Fields.Item("user_logcount").Value)%>x</td>-->
            <td width="20" class="text"><%=(users.Fields.Item("session_count").Value)%></td>
            <td width="20" class="text">
              <%
			if user_session_count = 0 then
				response.write("<font color = blue>N/A</font>")
			elseif (user_total_rate) >= passrate then
				response.write("" & FormatNumber(user_total_rate,2) & "%")
			else
				response.write("" & FormatNumber(user_total_rate,2) & "%")
			end if
			%>
            </td>
			<td width="20" class="text"><font color = green><%=user_passes%></font></td>
			<td width="20" class="text"><font color = red><%=user_fails%></font></td>
            <td  align="right" width="20">
              <a href="q_user_edit.asp?user=<%=(users.Fields.Item("ID_user").Value)%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&show_lines=<%=show_lines%>"><img src="images/edit.gif" width="16" height="15" border="0"></a>
            </td>
            <!-- CXS 061122: emailer functionality -->
				<td  align="right" width="18">
				<a href="send_individual_email.asp?individual=<%=(users.Fields.Item("ID_user").Value)%>"><img src="images/unread.gif" alt="Send reminder email to this user" border="0"></a></td>
				<!--PN 040811 add merge facility so that users can be merged into this user to get rid of self reg duplicates-->
				<td  align="right" width="18">
				<a href="#" onclick="var wintoopen=window.open('q_list_of_users_to_merge.asp?user=<%=(users.Fields.Item("ID_user").Value)%>','merge','toolbar=0, scrollbars=yes,resizable=1,width=700, height=500');wintoopen.focus();"><img src="images/merge.gif" alt="Merge users with this user" width="15" height="15" border="0"></a>
			</td>
				<td  align="right" width="18">
				<a href="#" onclick="var wintoopen=window.open('q_user_addresult.asp?user=<%=(users.Fields.Item("ID_user").Value)%>','merge','toolbar=0, scrollbars=yes,resizable=1,width=700, height=500');wintoopen.focus();"><img src="images/addresults.gif" alt="Add results" border="0"></a>
			</td>
          </tr>
  <%
else if (cstr(results)="1") and (user_total_rate >= passrate) then

%>

	<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">
            <td class="text" width="20"><%=count%></td>
            <td width="200" class="text">
              <%if cInt(users.Fields.Item("session_count").Value) > 0 then%>
             <a href="q_user_sessions.asp?user=<%=(users.Fields.Item("ID_user").Value)%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&fromdate=<%=fromdate%>&todate=<%=todate%>&active=<%=active%>&results=<%=results%>&passrate=<%=passrate%>&mths=<%=mths%>&noquiz=<%=noquiz%>&status=<%=request("status")%>">
              <%end if%>
              <%=(users.Fields.Item("user_username").Value)%>&nbsp;<%=(users.Fields.Item("user_firstname").Value)%></a></td>
            <td width="140" class="text"><%=(users.Fields.Item("info1").Value)%> (<%=(users.Fields.Item("info2").Value)%>)</td>
            <td width="140" class="text"><%=(users.Fields.Item("info3").Value)%></td>
            <td width="30" class="text" align=center>
              <%if abs(users.Fields.Item("user_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
            </td>
            <!--<td width="20" class="text"><%=(users.Fields.Item("user_logcount").Value)%>x</td>-->
            <td width="20" class="text"><%=(users.Fields.Item("session_count").Value)%></td>
            <td width="20" class="text">
              <%
			response.write("" & FormatNumber(user_total_rate,2) & "%")
			count = count + 1
			%>
            </td>
			<td width="20" class="text"><font color = green><%=user_passes%></font></td>
			<td width="20" class="text"><font color = red><%=user_fails%></font></td>
            <td  align="right" width="20">
              <a href="q_user_edit.asp?user=<%=(users.Fields.Item("ID_user").Value)%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&show_lines=<%=show_lines%>"><img src="images/edit.gif" width="16" height="15" border="0"></a>
            </td>
            <!-- CXS 061122: emailer functionality -->
				<td  align="right" width="18">
				<a href="send_individual_email.asp?individual=<%=(users.Fields.Item("ID_user").Value)%>"><img src="images/unread.gif" alt="Send reminder email to this user" border="0"></a></td>
				<!--PN 040811 add merge facility so that users can be merged into this user to get rid of self reg duplicates-->
				<td  align="right" width="18">
				<a href="#" onclick="var wintoopen=window.open('q_list_of_users_to_merge.asp?user=<%=(users.Fields.Item("ID_user").Value)%>','merge','toolbar=0, scrollbars=yes,resizable=1,width=700, height=500');wintoopen.focus();"><img src="images/merge.gif" alt="Merge users with this user" width="15" height="15" border="0"></a>
			</td>
				<td  align="right" width="18">
				<a href="#" onclick="var wintoopen=window.open('q_user_addresult.asp?user=<%=(users.Fields.Item("ID_user").Value)%>','merge','toolbar=0, scrollbars=yes,resizable=1,width=700, height=500');wintoopen.focus();"><img src="images/addresults.gif" alt="Add results" border="0"></a>
			</td>
          </tr>
  <%

  else if (cstr(results)="0") and (user_total_rate <= passrate) and (user_session_count<>0) then%>
			<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">
            <td class="text" width="20"><%=count%></td>
            <td width="200" class="text">
              <%if cInt(users.Fields.Item("session_count").Value) > 0 then%>
              <a href="q_user_sessions.asp?user=<%=(users.Fields.Item("ID_user").Value)%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&fromdate=<%=fromdate%>&todate=<%=todate%>&show_lines=<%=show_lines%>">
              <%end if%>
              <%=(users.Fields.Item("user_username").Value)%>&nbsp;<%=(users.Fields.Item("user_firstname").Value)%></a></td>
            <td width="140" class="text"><%=(users.Fields.Item("info1").Value)%> (<%=(users.Fields.Item("info2").Value)%>)</td>
            <td width="140" class="text"><%=(users.Fields.Item("info3").Value)%></td>
            <td width="30" class="text" align=center>
              <%if abs(users.Fields.Item("user_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
            </td>
            <!--<td width="20" class="text"><%=(users.Fields.Item("user_logcount").Value)%>x</td>-->
            <td width="20" class="text"><%=(users.Fields.Item("session_count").Value)%></td>
            <td width="20" class="text">
              <%
			response.write("" & FormatNumber(user_total_rate,2) & "%")
			count = count + 1
			%>
            </td>
			<td width="20" class="text"><font color = green><%=user_passes%></font></td>
			<td width="20" class="text"><font color = red><%=user_fails%></font></td>
            <td  align="right" width="20">
              <a href="q_user_edit.asp?user=<%=(users.Fields.Item("ID_user").Value)%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&show_lines=<%=show_lines%>"><img src="images/edit.gif" width="16" height="15" border="0"></a>
            </td>
            <!-- CXS 061122: emailer functionality -->
				<td  align="right" width="18">
				<a href="send_individual_email.asp?individual=<%=(users.Fields.Item("ID_user").Value)%>"><img src="images/unread.gif" alt="Send reminder email to this user" border="0"></a></td>
				<!--PN 040811 add merge facility so that users can be merged into this user to get rid of self reg duplicates-->
				<td  align="right" width="18">
				<a href="#" onclick="var wintoopen=window.open('q_list_of_users_to_merge.asp?user=<%=(users.Fields.Item("ID_user").Value)%>','merge','toolbar=0, scrollbars=yes,resizable=1,width=700, height=500');wintoopen.focus();"><img src="images/merge.gif" alt="Merge users with this user" width="15" height="15" border="0"></a>
			</td>
				<td  align="right" width="18">
				<a href="#" onclick="var wintoopen=window.open('q_user_addresult.asp?user=<%=(users.Fields.Item("ID_user").Value)%>','merge','toolbar=0, scrollbars=yes,resizable=1,width=700, height=500');wintoopen.focus();"><img src="images/addresults.gif" alt="Add results" border="0"></a>
			</td>
          </tr>
  <%
  end if
  end if
  end if
end if
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  users.MoveNext()
  numbers=numbers+1
Wend
'If (users.CursorType > 0) Then
'  users.MoveFirst
'Else
  users.Requery
'End If

'_______________________________________________________________________________

'PN 050507 variables to store user passes and fails
total_passes=0
total_fails=0

overall_session_rate = 0
overall_session_count = 0
overall_session_passed = 0
overall_user_sessions = 0
numbers=1
While (NOT users.EOF)

if subject_prm <> 0 then
	subj_prm ="and (q_session.Session_subject ="&subject_prm&")"
else
	subj_prm =""
end if

if cstr(fromdate)="" and cstr(todate) <> "" then
	user_details.Source = "SELECT q_session.id_session, q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) and (session_finish <= '"&todate&"') "&subj_prm&" order by session_date desc,session_subject"
else if cstr(todate)="" and cstr(fromdate) <> "" then
	user_details.Source = "SELECT q_session.id_session, q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) and (session_finish >= '"&fromdate&"') "&subj_prm&" order by session_date desc,session_subject"
else if (cstr(todate)="" and cstr(fromdate)="") then
	user_details.Source = "SELECT q_session.id_session, q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) "&subj_prm&" order by  session_date desc,session_subject"
else
	user_details.Source = "SELECT q_session.id_session, q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) and ((session_finish >= '"&fromdate&"') and (session_finish <= '"&todate&"')) "&subj_prm&" order by session_date desc,session_subject"
end if
end if
end if
user_details.Open()
user_details_numRows = 0
user_session_percentage=0
user_session_rate = 0
user_session_count = 0
user_total_rate = 0
subid =0
subject_pass_rate_percentage = default_passrate

While (NOT user_details.EOF)
if session("mths")="" then
	user_session_correct = cInt(user_details.Fields.Item("session_correct").Value)
	user_session_total = cInt(user_details.Fields.Item("session_total").Value)
	user_session_rate = user_session_rate + (user_session_correct / user_session_total * 100)
	user_session_percentage=(user_session_correct / user_session_total * 100)
	user_session_count = user_session_count + 1
	overall_user_sessions=overall_user_sessions + 1
else if cint(subid) <> cInt(user_details.Fields.Item("session_subject").Value) then
	subid = cInt(user_details.Fields.Item("session_subject").Value)
	user_session_correct = cInt(user_details.Fields.Item("session_correct").Value)
	user_session_total = cInt(user_details.Fields.Item("session_total").Value)
	user_session_rate = user_session_rate + (user_session_correct / user_session_total * 100)
	user_session_percentage=(user_session_correct / user_session_total * 100)
	user_session_count = user_session_count + 1
	overall_user_sessions=overall_user_sessions + 1
end if
end if

	subject_pass_rate_percentage = default_passrate
	if passrate_type = 1 then
		set pass_rate_query = Server.CreateObject("ADODB.Recordset")
		pass_rate_query.ActiveConnection = Connect
		pass_rate_query.Source = "SELECT * FROM q_certification where q_certification.q_session = "&user_details.Fields.Item("ID_Session").Value&""
		pass_rate_query.CursorType = 0
		pass_rate_query.CursorLocation = 3
		pass_rate_query.LockType = 3
		pass_rate_query.Open()
		if (not pass_rate_query.eof) then
			subject_pass_rate_percentage = pass_rate_query.Fields.Item("percentage_required").Value
		else
			set subject = Server.CreateObject("ADODB.Recordset")
			subject.ActiveConnection = Connect
			subject.Source = "SELECT * FROM subjects where subjects.id_subject = "&user_details.Fields.Item("session_subject").Value&""
			subject.CursorType = 0
			subject.CursorLocation = 3
			subject.LockType = 3
			subject.Open()
			subject_pass_rate_percentage = subject.fields.item("subject_passmark").value
			subject.Close()
		end if
		pass_rate_query.Close()
	elseif passrate_type=2 then
		set pass_rate_query = Server.CreateObject("ADODB.Recordset")
		pass_rate_query.ActiveConnection = Connect
		pass_rate_query.Source = "SELECT * FROM q_certification where q_certification.q_session = "&user_details.Fields.Item("ID_Session").Value&""
		pass_rate_query.CursorType = 0
		pass_rate_query.CursorLocation = 3
		pass_rate_query.LockType = 3
		pass_rate_query.Open()
		if (not pass_rate_query.eof) then
			subject_pass_rate_percentage = pass_rate_query.Fields.Item("percentage_required").Value
		else
			set pass_rate_query2 = Server.CreateObject("ADODB.Recordset")
			pass_rate_query2.ActiveConnection = Connect
			pass_rate_query2.Source = "SELECT pass_rate FROM pass_rates where subject="&user_details.Fields.Item("session_subject").Value&" and q_info1="&users.Fields.Item("user_info1").Value&" and q_info3="&users.Fields.Item("user_info3").Value&";"
			pass_rate_query2.CursorType = 0
			pass_rate_query2.CursorLocation = 3
			pass_rate_query2.LockType = 3
			pass_rate_query2.Open()
			if (not pass_rate_query2.eof) then
				subject_pass_rate_percentage = pass_rate_query2.Fields.Item("pass_rate").Value
			end if
			pass_rate_query2.Close()
		end if
		pass_rate_query.Close()
	elseif passrate_type=3 then
		set pass_rate_query = Server.CreateObject("ADODB.Recordset")
		pass_rate_query.ActiveConnection = Connect
		pass_rate_query.Source = "SELECT pass_rate FROM pass_rates where subject="&user_details.Fields.Item("session_subject").Value&" and q_info1="&users.Fields.Item("user_info1").Value&" and q_info3="&users.Fields.Item("user_info3").Value&";"
		pass_rate_query.CursorType = 0
		pass_rate_query.CursorLocation = 3
		pass_rate_query.LockType = 3
		pass_rate_query.Open()
		if (not pass_rate_query.eof) then
			subject_pass_rate_percentage = pass_rate_query.Fields.Item("pass_rate").Value
		end if
		pass_rate_query.Close()
	else
		set subject = Server.CreateObject("ADODB.Recordset")
		subject.ActiveConnection = Connect
		subject.Source = "SELECT * FROM subjects where subjects.id_subject = "&user_details.Fields.Item("session_subject").Value&""
		subject.CursorType = 0
		subject.CursorLocation = 3
		subject.LockType = 3
		subject.Open()
		subject_pass_rate_percentage = subject.fields.item("subject_passmark").value
		subject.Close()
	end if


'PN050506 calulate if this session is a pass or a fail
'Response.Write("the user_session_rate is  "&user_session_percentage&"   and the passrate is "&subject_pass_rate_percentage&"<br><br>")
if(cInt(user_session_percentage)>=cInt(subject_pass_rate_percentage)) then
	'it was a pass
	total_passes=total_passes+1
else
	'it was a fail
	total_fails=total_fails+1
end if
user_details.MoveNext()
Wend
user_details.Close()

if user_session_count > 0 then
	user_total_rate = (user_session_rate/user_session_count)
	if cstr(results)="1" then
		if user_total_rate >= passrate then
			overall_session_rate = overall_session_rate + user_total_rate
			overall_session_count = overall_session_count + 1
			overall_session_passed = overall_session_passed + 1
		end if
	else if cstr(results)="0" then
		if user_total_rate <= passrate then
			overall_session_rate = overall_session_rate + user_total_rate
			overall_session_count = overall_session_count + 1
		end if
	else
		overall_session_rate = overall_session_rate + user_total_rate
		overall_session_count = overall_session_count + 1
		if user_total_rate >= passrate then overall_session_passed = overall_session_passed + 1
	end if
	end if
end if
users.MoveNext()
if (cstr(noquiz)="1") then
	if user_session_count = 0 then
		numbers=numbers+1
	end if
else
	numbers=numbers+1
end if
Wend
if overall_session_count > 0 then overll_pass_rate = overall_session_rate/overall_session_count
'________________________________________________________________________________
%>
          <tr>
            <td class="text" colspan="12">
              <hr>
            </td>
          </tr>
          <tr>
            <td class="subheads" colspan="12">Overall results of filtered users : </td>
          </tr>
          <tr class="table_normal">
            <td class="text" colspan="2" align="left" valign="top">Number of users
              in selection :</td>
            <td colspan="2" class="text" align="left" valign="top">Users
              with at least 1 finished session :<br>
             </font> </td>
            <td colspan="4" class="text" align="left" valign="top">Average score (%) of users with at least 1 finished session :</td>

             <td class="text" colspan="4" align="left" valign="top">Percentage of users who have completed a quiz:</td>
             <%
             if cint(cnt)= 0 then cnt =1
             ' Response.Write overall_session_count
            'Response.Write overall_session_passed
             %>


          </tr>
          <tr class="table_normal">
			<%
			if (cstr(results)<> "2" and  cstr(results) <> "") or cstr(noquiz)= "1" then
			%>
            <td class="text" colspan="2"><%'=count-1%><%=numbers-1%><%cnt = count -1
			%></td>
            <%else%>
             <td class="text" colspan="2"><%=numbers-1%><%cnt = numbers-1%></td>
             <%end if%>

            <%if cstr(noquiz)="1" then
            overall_session_count = 0
            overall_session_passed = 0
            end if
          '  Response.Write overall_session_count
           ' Response.Write overall_session_passed
            %>

            <td colspan="2" class="text"><%=overall_session_count%></td>
            <td colspan="4" class="text">
            <%
			if overall_session_count = 0 then
				response.write("<font color = blue>Nothing to rate</font>")
			else
				response.write(" " & FormatNumber(overll_pass_rate,2) & "% </font>")
			'else
				'response.write("<font color = red>" & FormatNumber(overll_pass_rate,2) & "% - FAILED</font>")
			end if
			%>
            </td>
			<td colspan="4" class="text"><%
			if overall_session_count = 0 or cnt = 0 then
				response.write("0%")
			else
				response.write(FormatNumber((100 * (overall_session_count))/(cnt),2) &"%")
			end if %></td>
          </tr>

          <tr>
          </td>
          <tr class="table_normal">
            <td class="text" colspan=12>
			<%  if (cstr(noquiz)<>"1") then%>
            <table>
				<tr>
				  <td class="text">Details for all user sessions:</td>
				  <td class="text"></td>
				</tr>
				<tr>
				 <td class="text"> Total / <font color = green>Passed (%)</font> / <font color = red>Failed (%)</td>
				 <td class="text"></td>
				</tr>
				<tr>
				<td class="text">
				  <% if (total_passes+total_fails)<>0  then %>
				  <% =overall_user_sessions & " / <font color = green>" & total_passes &" ("&FormatNumber((100 * (total_passes))/(total_passes+total_fails),2)&"%) </font> / <font color = red>" & total_fails&" ("&FormatNumber((100 * (total_fails))/(total_passes+total_fails),2)&"%) </font>"%>
				  <% end if %>
				  </td>
				  <td class="text"></td>
				</tr>
            </table> 
			<% end if%>
            </td>
          </tr>

        <tr>
			<td  colspan="12">Please bear in mind, that the OVERALL results above reflect the whole
              selection of filtered users. If there is more than one page of users identified as
              matching your selection, the overall results will capture all the users on all the pages. i.e. if only 50 out
              of 65 users are shown on the screen, the results cover all 65 users, including those on the next page.</i>
			</td>
		</tr>
          <% End If %>
          <% If users.EOF And users.BOF Then %>
          <tr>
            <td  width="18">&nbsp;<input type="hidden" class="formitem1" name="passrate"  size=3 value=""></td>
            <td colspan="8" >Sorry,
              there are no users in the quiz currently or no user match your filter
              criteria.
			</td>
          </tr>
          <% End If %>
          <tr>
            <td  width="18"><img src="images/new2.gif" width="11" height="13"></td>
            <td colspan="8" >
              <input type="button" name="Button" value="Add a new user" onClick="document.location='q_user_add.asp';" class="quiz_button">
              <input type="button" name="Button2" value="Import users from TXT file" onClick="document.location='q_user_import.asp';" class="quiz_button">
            </td>
          </tr>
          <tr>
            <td  width="18">&nbsp;</td>
            <td colspan="12" >&nbsp;
              <table width="40%" >
                <tr class="table_normal">
                  <td width="25%" align="center" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">
                    <% If MM_offset <> 0 Then %>
                    <a href="<%=MM_moveFirst%>"><img src="images/first.gif" border=0></a>
                    <% End If %>
                  </td>
                  <td width="25%" align="center" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">
                    <% If MM_offset <> 0 Then %>
                    <a href="<%=MM_movePrev%>"><img src="images/previous.gif" border=0></a>
                    <% End If %>
                  </td>
                  <td width="25%" align="center" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">
                    <% If Not MM_atTotal Then %>
                    <a href="<%=MM_moveNext%>"><img src="images/next.gif" border=0></a>
                    <% End If %>
                  </td>
                  <td width="25%" align="center" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">
                    <% If Not MM_atTotal Then %>
                    <a href="<%=MM_moveLast%>"><img src="images/last.gif" border=0></a>
                    <% End If %>
                  </td>
                </tr>
                <tr class="table_normal">
                  <td colspan="4" align="center"class="text">&nbsp; Users <b><%=(users_first)%></b> to <b><%=(users_last)%></b> of <b><i><%=(users_total)%> </i></b>
                    <%if (SQL_having <> "") or (SQL_where <> "") then response.write("<font color='#FF0000'>(filtered - <a href='javascript:clearform();'>clear filter</a>)</font>")%>
                  </td>
                </tr>
                <tr class="table_normal">
                  <td colspan="4" align="center"class="text">Show
                    <input type="text" name="show_lines" class="formitem1" size="3" maxlength="3" value="">
                    users per page</td>
                </tr>
                <tr class="table_normal">
                  <td colspan="4" align="center"class="text">
                    <input type="button" name="Submit" value="&gt;&gt;&gt; Filter users &lt;&lt;&lt;" class="quiz_button" onclick="return filter_submit();">
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
</BODY>
</HTML>
<%
call log_the_page ("Quiz List Users")
users.Close()
Set users = Nothing
filter_info1.Close()
filter_info3.Close()
subjects.Close()
%>

