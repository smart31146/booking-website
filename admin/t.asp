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
todate=request("todate")
active = request("active")
mths = request("mths")
noquiz = request("noquiz")

if cStr(Request.Querystring("show_lines")) <> "" then show_lines = cInt(Request.Querystring("show_lines"))

If cStr(Request.Querystring("filter_username")) <> "" then 
	SQL_having = " HAVING ((q_user.user_lastname) Like '%" + Replace(uCase(cStr(Request.Querystring("filter_username"))), "'", "''") + "%' OR  (q_user.user_firstname) Like '%" + Replace(uCase(cStr(Request.Querystring("filter_username"))), "'", "''") + "%') " 
end if 

subject_prm = 0

If cInt(Request.Querystring("subject")) <> 0 then 
	subject_prm = cInt(Request.Querystring("subject"))
	SQL_where = " WHERE (q_session.Session_subject = " + (Request.Querystring("subject")) + ") " 
	if cstr(request("active"))="1" then
		SQL_WHERE=SQL_where + "and q_user.user_active=1"
	else if cstr(request("active"))="0" then
		SQL_WHERE=SQL_where + "and q_user.user_active=0"
	end if
	end if
	
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
	else if cstr(request("active"))="0" then
		sql_where="where  q_user.user_active=0"
	end if
	end if
	
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
	now_date = now()
	back_date = (DateAdd("m", -12, now_date))
	back_date = formatdatetime( back_date,2)
	back_date = back_date & " 00:00:00 AM"
	if sql_where <> "" then
		SQL_where = sql_where + "and  q_session.session_date between '"&back_date&"' and '"&now_date&"'"
	else
		sql_where ="where  q_session.session_date between '"&back_date&"' and '"&now_date&"'"
	end if
else
	session("mths")=""
end if
set users = Server.CreateObject("ADODB.Recordset")
users.ActiveConnection = Connect
users.Source = "SELECT q_user.ID_user, q_user.user_lastname, q_user.user_firstname, q_info1.info1, q_info2.info2, q_info3.info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3, COUNT(q_session.ID_session) AS session_count FROM (q_info3 RIGHT JOIN (q_info2 RIGHT JOIN (q_info1 RIGHT JOIN q_user ON q_info1.ID_info1 = q_user.user_info1) ON q_info2.ID_info2 = q_user.user_info2) ON q_info3.ID_info3 = q_user.user_info3) LEFT JOIN q_session ON q_user.ID_user = q_session.Session_users " + SQL_where + " GROUP BY q_user.user_lastname, q_user.user_firstname, q_user.ID_user, q_info1.info1, q_info2.info2, q_info3.info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3 " + SQL_having + " ORDER BY q_user.user_lastname, q_user.user_firstname;"
'SQL: "SELECT q_user.ID_user, q_user.user_lastname, q_user.user_firstname, q_info1.info1, q_info2.info2, q_info3.info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3, COUNT(q_session.ID_session) AS session_count FROM q_info2 RIGHT OUTER JOIN q_info1 RIGHT OUTER JOIN q_session RIGHT OUTER JOIN q_user ON q_session.Session_users = q_user.ID_user ON q_info1.ID_info1 = q_user.user_info1 ON  q_info2.ID_info2 = q_user.user_info2 LEFT OUTER JOIN q_info3 ON q_user.user_info3 = q_info3.ID_info3 " + SQL_where + " GROUP BY q_user.user_lastname, q_user.user_firstname, q_user.ID_user, q_info1.info1, q_info2.info2, q_info3.info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3 " + SQL_having + " ORDER BY q_user.user_lastname, q_user.user_firstname;"
'Access: "SELECT q_user.ID_user, q_user.user_lastname, q_user.user_firstname, q_info1.info1, q_info2.info2, q_info3.info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3, COUNT(q_session.ID_session) AS session_count FROM (q_info3 RIGHT JOIN (q_info2 RIGHT JOIN (q_info1 RIGHT JOIN q_user ON q_info1.ID_info1 = q_user.user_info1) ON q_info2.ID_info2 = q_user.user_info2) ON q_info3.ID_info3 = q_user.user_info3) LEFT JOIN q_session ON q_user.ID_user = q_session.Session_users " + SQL_where + " GROUP BY q_user.user_lastname, q_user.user_firstname, q_user.ID_user, q_info1.info1, q_info2.info2, q_info3.info3, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3 " + SQL_having + " ORDER BY q_user.user_lastname, q_user.user_firstname;"
'Response.Write users.Source
users.CursorType = 0
users.CursorLocation = 3
users.LockType = 3
users.Open()
users_numRows = 0

set filter_info1 = Server.CreateObject("ADODB.Recordset")
filter_info1.ActiveConnection = Connect
filter_info1.Source = "SELECT * FROM q_info1 order by info1"
filter_info1.CursorType = 0
filter_info1.CursorLocation = 3
filter_info1.LockType = 3
filter_info1.Open()
filter_info1_numRows = 0

set filter_info3 = Server.CreateObject("ADODB.Recordset")
filter_info3.ActiveConnection = Connect
filter_info3.Source = "SELECT * FROM q_info3 order by info3"
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
MM_rs.Requery


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
		document.filter_users.mths.checked = false
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
        D.setTime(D.getTime()+86400000*days)
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
	passrate = MyCookie.Read('passrate')
	if (passrate != null) {
		document.forms[0].passrate.value=passrate;
	}
	else {
		document.forms[0].passrate.value=50;
	}
		
	show_lines = MyCookie.Read('show_lines')
	if (show_lines != null) {
		document.forms[0].show_lines.value=show_lines;
	}
	else {
		document.forms[0].show_lines.value=25;
	}
}

function AddCookieId(cn,id) {
        MyCookie.Write(cn,id,7);
}

function DelCookieId(cn,id) {
        MyCookie.Write(cn,id,-1);
}
function pass_submit()
{
	if (isNaN(document.filter_users.passrate.value)){
		alert('Invalid pass rate');
		document.filter_users.passrate.focus();
		return false;
	}
	else
	{
		AddCookieId("passrate",document.filter_users.passrate.value);
		passrate = MyCookie.Read('passrate')
		document.forms[0].submit();
		return true;
	}
}

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
		show_lines = MyCookie.Read('show_lines')
		document.forms[0].submit();
		return true;
	}
}
function clearform()
{
document.forms[0].filter_username.value = "";
document.forms[0].todate.value = "";
document.forms[0].fromdate.value = "";
document.forms[0].results.selectedIndex = 0;
document.forms[0].subject.selectedIndex = 0;
document.forms[0].active.selectedIndex = 0;
document.forms[0].filter_info1.selectedIndex = 0;
document.forms[0].filter_info3.selectedIndex = 0;
document.forms[0].submit();
}
//-->
</script>
</HEAD>

<BODY onload="check();">
	<%
	if Request.Cookies("passrate")<> "" then
		passrate= cint(Request.Cookies("passrate"))
		'MyFile = Server.MapPath("passrate.txt")
		'Response.Write myfile
		'set fso = server.CreateObject("Scripting.FileSystemObject")
		'Set TSO = FSO.OpenTextFile(myfile,2,create)
		'TSO.write passrate
		'TSO.close
		'set TSO = nothing
		'set FSO = nothing
	else
		passrate=50
	end if
	
	if Request.Cookies("show_lines")<> "" then
		show_lines= cint(Request.Cookies("show_lines"))
	else
		show_lines=15
	end if

%>
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> Quiz users</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form name="filter_users">
     <input type="hidden" name="hiddenmths" value="false">
        <table>
          <tr> 
            <td colspan="8" class="subheads" align="left" valign="top">Users:</td>
           <td align="right" class="subheads" valign="top" width="20"><a href="q_export_quiz_users.asp?filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&fromdate=<%=fromdate%>&todate=<%=todate%>&active=<%=active%>&results=<%=results%>&passrate=<%=passrate%>&mths=<%=mths%>&noquiz=<%=noquiz%>"><img src="images/xls.gif" width="16" height="16" border="0"></a></td>
          </tr>
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
            <td class="text" width="18"><img src="images/back.gif" width="18" height="14"></td>
            <td class="text" colspan="8"><a href="main.asp">...Home page 
              </a> </td>
          </tr>
          <tr> 
            <td class="subheads" colspan="9">Filter users by:</td>
          </tr>
          <tr class="table_normal">
          <td class="text" width="18">&nbsp;</td>
          <td class="text" valign="top" width="143">User Type:</td>
           <td class="text" valign="top" colspan="7">
           <%'=request("active")%>
           <select name="active" class="formitem1">
              <%if cstr(request("active"))="" or cstr(request("active")="2") then%>
              <option value="2" selected>All Users</option>
              <option value="1">Active Users</option>
              <option value="0">Inactive Users</option>
              <%else if cstr(request("active"))="1" then%>
              <option value="2" >All Users</option>
              <option value="1" selected>Active Users</option>
              <option value="0">Inactive Users</option>
              <%else if cstr(request("active"))="0" then%>
              <option value="2"> All Users</option>
              <option value="1">Active Users</option>
              <option value="0" selected>Inactive Users</option>
              <%end if
              end if
              end if
              %>
              </select>
              </td>
		</tr>
		<tr class="table_normal">
          <td class="text" width="18">&nbsp;</td>
          <td class="text" valign="top" width="143">Sessions between:</td>
           <td  valign="top" colspan="7">
           <input type="text" name="fromdate" maxlength="19" class="formitem1" onDblClick="this.value='<%=cDateSQL(Now()-1)%>'; document.filter_users.mths.checked=false; document.filter_users.mths.disabled=true" onchange="return checkmths();" size="25" value="<%=fromdate%>" >&nbsp;(yyyy/mm/dd hh:mm:ss), doubleclick = TODAY - 1 day<br>
           &nbsp;&nbsp;&nbsp;&nbsp;and <br>
           <input type="text" name="todate" maxlength="19" class="formitem1" onDblClick="this.value='<%=cDateSQL(Now())%>'; document.filter_users.mths.checked=false; document.filter_users.mths.disabled=true" onchange="return checkmths();" size="25" value="<%=todate%>">
              (yyyy/mm/dd hh:mm:ss), doubleclick = TODAY
              </td>
		</tr>
		
          <tr class="table_normal"> 
            <td class="text" width="18">&nbsp;</td>
            <td class="text" valign="top" width="143">First OR Last name:</td>
            <td class="text" valign="top" colspan="7"> 
            <table><tr><td><input type="text" name="filter_username" value="<%=request.querystring("filter_username")%>" class="formitem1"></td><td><!--<a href="javascript:onclick=filter_users.submit();" target='_self'><img src="images/go.gif" border=0></a>--></td></tr></table>
            </td>
            
          </tr>
          <tr class="table_normal">
          <td class="text" width="18">&nbsp;</td>
          <td class="text" valign="top" width="143">Results:</td>
           <td class="text" valign="top" colspan="7">
           <select name="results" class="formitem1">
              <%if cstr(request("results"))="" or cstr(request("results")="2") then%>
              <option value="2" selected>All Users</option>
              <option value="1">Passed</option>
              <option value="0">Failed</option>
              <%else if cstr(request("results"))="1" then%>
              <option value="2" >All Users</option>
              <option value="1" selected>Passed</option>
              <option value="0">Failed</option>
              <%else if cstr(request("results"))="0" then%>
              <option value="2">All Users</option>
              <option value="1">Passed</option>
              <option value="0" selected>Failed</option>
              <%end if
              end if
              end if
              %>
              </select>
              </td>
              
              
             </tr>
          <tr class="table_normal"> 
            <td class="text" width="18">&nbsp;</td>
            <td class="text" valign="top" width="143">Subject:</td>
            <td class="text" valign="top" colspan="8"> 
              <select name="subject" class="formitem1">
                <option value="0">--- select a subject ---</option>
                <%
While (NOT subjects.EOF)
%>
                <option value="<%=(subjects.Fields.Item("ID_subject").Value)%>" <%if (CStr(subjects.Fields.Item("ID_subject").Value) = CStr(request.querystring("subject"))) then Response.Write("SELECTED") : Response.Write("")%>><%=(subjects.Fields.Item("subject_name").Value)%></option>
                <%
  subjects.MoveNext()
  
Wend
  subjects.Requery
%>
              </select>
            </td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="18">&nbsp;</td>
            <td class="text" valign="top" width="143">Business:</td>
            <td class="text" valign="top" colspan="8"> 
              <select name="filter_info1" class="formitem1">
                <option value="0">--- select a business ---</option>
                <%
While (NOT filter_info1.EOF)
%>
                <option value="<%=(filter_info1.Fields.Item("ID_info1").Value)%>" <%if (CStr(filter_info1.Fields.Item("ID_info1").Value) = CStr(request.querystring("filter_info1"))) then Response.Write("SELECTED") : Response.Write("")%>><%=(filter_info1.Fields.Item("info1").Value)%></option>
                <%
  filter_info1.MoveNext()
Wend
filter_info1.Requery
%>
              </select>
            </td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="18">&nbsp;</td>
            <td class="text" valign="top" width="143"><% =BBPinfo3 %>:</td>
            <td class="text" valign="top" colspan="8"> 
              <select name="filter_info3" class="formitem1">
                <option value="0">--- select a <% =BBPinfo3 %> ---</option>
                <%
While (NOT filter_info3.EOF)
%>
                <option value="<%=(filter_info3.Fields.Item("ID_info3").Value)%>" <%if (CStr(filter_info3.Fields.Item("ID_info3").Value) = CStr(request.querystring("filter_info3"))) then Response.Write("SELECTED") : Response.Write("")%>><%=(filter_info3.Fields.Item("info3").Value)%></option>
                <%
  filter_info3.MoveNext()
Wend
filter_info3.Requery
%>
              </select>
            </td>
          </tr>
	<tr class="table_normal"> 
	            <td class="text" width="18">&nbsp;</td>
	            <td class="text" valign="top" width="143">User sessions in the past 12 months:</td>
	            <td class="text" valign="top" colspan="8"> 
	              <%
	              if request("mths")=1 then%>
	              <input type="checkbox" name="mths" value=1 checked>
	              <%else if request("fromdate") <> "" or request("todate") <> "" then%>
	               <input type="checkbox" name="mths" value=1 disabled=true>
	              <%else%>
	              <input type="checkbox" name="mths" value=1>
	              <%
	              end if
	              end if%>
	            </td>
	            
	          </tr>       
	<tr class="table_normal"> 
	            <td class="text" width="18">&nbsp;</td>
	            <td class="text" valign="top" width="143">Users not attempted quiz:</td>
	            <td class="text" valign="top" colspan="8"> 
	            <%if request("noquiz") = 1 then %>
	            <input type="checkbox" name="noquiz" value=1 checked>
	            <%else%>
	            <input type="checkbox" name="noquiz" value=1>
	            <%end if%>
	            </td>
	            
	          </tr>	             
          <tr class="table_normal"> 
                  <td colspan="9" align="center"class="text"> 
                    <input type="button" name="Submit" value="&gt;&gt;&gt; Filter users &lt;&lt;&lt;" class="quiz_button" onclick="return filter_submit();">
                  </td>
                </tr>
          <tr> 
            <td >&nbsp;</td>
            <td >Last name &amp; First name</td>
            <td >Business &amp; <% =BBPinfo3 %></td>
            <td ><% =BBPinfo3 %></td>
            <td >Active</td>
            <td >Logs</td>
            <td >Sess.</td>
            <td >Rate</td>
            <td >Edit</td>
          </tr>
          <% If Not users.EOF Or Not users.BOF Then %>
          <% 
While ((Repeat1__numRows <> 0) AND (NOT users.EOF)) 
%>
          <%
if subject_prm <> 0 then
	subj_prm ="and (q_session.Session_subject ="&subject_prm&")"
else 
	subj_prm =""
end if

'if session("mths") <> "" then
	' mths_prm = "and  q_session.session_date between '"&back_date&"' and '"&now_date&"'"
'else
	'mths_prm = ""
'end if	


if cstr(fromdate)="" and cstr(todate) <> "" then
	user_details.Source = "SELECT q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) and (session_finish <= '"&todate&"') "&subj_prm&" "&mths_prm&" order by session_users, session_subject, session_date"
else if cstr(todate)="" and cstr(fromdate) <> "" then
	user_details.Source = "SELECT q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) and (session_finish >= '"&fromdate&"') "&subj_prm&" "&mths_prm&" order by session_users, session_subject, session_date"
else if (cstr(todate)="" and cstr(fromdate)="") then
	user_details.Source = "SELECT q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) "&subj_prm&" "&mths_prm&" order by session_users, session_subject, session_date"
else
	user_details.Source = "SELECT q_session.session_subject, q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) and ((session_finish >= '"&fromdate&"') and (session_finish <= '"&todate&"')) "&subj_prm&" "&mths_prm&" order by session_users,session_subject, session_date"
end if 
end if
end if		
user_details.Open()

user_details_numRows = 0
%>
          <%
user_session_rate = 0
user_session_count = 0
user_total_rate = 0
userid = 0
subjid=0

While (NOT user_details.EOF)

if session("mths") = "" then
	user_session_correct = cInt(user_details.Fields.Item("session_correct").Value)
	user_session_total = cInt(user_details.Fields.Item("session_total").Value)
	user_session_rate = user_session_rate + (user_session_correct / user_session_total * 100)
	user_session_count = user_session_count + 1
else
	if userid <> user_details.Fields.Item("session_users").Value then
		userid = user_details.Fields.Item("session_users").Value
		user_session_correct = cInt(user_details.Fields.Item("session_correct").Value)
		user_session_total = cInt(user_details.Fields.Item("session_total").Value)
		user_session_rate = user_session_rate + (user_session_correct / user_session_total * 100)
		user_session_count = user_session_count + 1
	else if userid = user_details.Fields.Item("session_users").Value then
		if subjid <> user_details.Fields.Item("session_subject").Value then
			Response.Write user_details.Fields.Item("session_subject").Value
			subjid = user_details.Fields.Item("session_subject").Value
			user_session_correct = cInt(user_details.Fields.Item("session_correct").Value)
			user_session_total = cInt(user_details.Fields.Item("session_total").Value)
			Response.Write user_session_rate
			user_session_rate = user_session_rate + (user_session_correct / user_session_total * 100)
			user_session_count = user_session_count + 1
		end if
	end if
	end if
end if
user_details.MoveNext()
Wend

%>
<%

user_details.Close()
%>
          <%
if user_session_count > 0 then 
	user_total_rate = (user_session_rate/user_session_count)
end if

if (cstr(noquiz)="1") then
	if cInt(users.Fields.Item("session_count").Value) = 0 then
%>
	<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">
			
            <td class="text" width="20"><%=count%></td>
            <td width="200" class="text"> 
             <%=(users.Fields.Item("user_lastname").Value)%>&nbsp;<%=(users.Fields.Item("user_firstname").Value)%></td>
            <td width="140" class="text"><%=(users.Fields.Item("info1").Value)%> (<%=(users.Fields.Item("info2").Value)%>)</td>
            <td width="140" class="text"><%=(users.Fields.Item("info3").Value)%></td>
            <td width="30" class="text" align=center> 
              <%if abs(users.Fields.Item("user_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
            </td>
            <td width="20" class="text"><%=(users.Fields.Item("user_logcount").Value)%>x</td>
            <td width="20" class="text"><%=(users.Fields.Item("session_count").Value)%></td>
            <td width="20" class="text"> 
              <%
			response.write("<font color = blue>N/A</font>") 
			count = count + 1
			%>
            </td>
            <td  align="right" width="20"> 
              <a href="q_user_edit.asp?user=<%=(users.Fields.Item("ID_user").Value)%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&show_lines=<%=show_lines%>"><img src="images/edit.gif" width="16" height="15" border="0"></a> 
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
              <a href="q_user_sessions.asp?user=<%=(users.Fields.Item("ID_user").Value)%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&fromdate=<%=fromdate%>&todate=<%=todate%>&show_lines=<%=show_lines%>"> 
              <%
               end if%>
              <%=(users.Fields.Item("user_lastname").Value)%>&nbsp;<%=(users.Fields.Item("user_firstname").Value)%></a></td>
            <td width="140" class="text"><%=(users.Fields.Item("info1").Value)%> (<%=(users.Fields.Item("info2").Value)%>)</td>
            <td width="140" class="text"><%=(users.Fields.Item("info3").Value)%></td>
            <td width="30" class="text" align=center> 
              <%if abs(users.Fields.Item("user_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
            </td>
            <td width="20" class="text"><%=(users.Fields.Item("user_logcount").Value)%>x</td>
            <td width="20" class="text"><%=(users.Fields.Item("session_count").Value)%></td>
            <td width="20" class="text"> 
              <%
			if user_session_count = 0 then
				response.write("<font color = blue>N/A</font>") 
			elseif (user_total_rate) >= passrate then 
				response.write("<font color = green>" & FormatNumber(user_total_rate,2) & "%</font>") 
			else 
				response.write("<font color = red>" & FormatNumber(user_total_rate,2) & "%</font>")
			end if
			%>
            </td>
            <td  align="right" width="20"> 
              <a href="q_user_edit.asp?user=<%=(users.Fields.Item("ID_user").Value)%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&show_lines=<%=show_lines%>"><img src="images/edit.gif" width="16" height="15" border="0"></a> 
            </td>
          </tr>
  <%
else if (cstr(results)="1") and (user_total_rate >= passrate) then

%>

	<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">
            <td class="text" width="20"><%=count%></td>
            <td width="200" class="text"> 
              <%if cInt(users.Fields.Item("session_count").Value) > 0 then%>
             <a href="q_user_sessions.asp?user=<%=(users.Fields.Item("ID_user").Value)%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&fromdate=<%=fromdate%>&todate=<%=todate%>&show_lines=<%=show_lines%>">
              <%end if%>
              <%=(users.Fields.Item("user_lastname").Value)%>&nbsp;<%=(users.Fields.Item("user_firstname").Value)%></a></td>
            <td width="140" class="text"><%=(users.Fields.Item("info1").Value)%> (<%=(users.Fields.Item("info2").Value)%>)</td>
            <td width="140" class="text"><%=(users.Fields.Item("info3").Value)%></td>
            <td width="30" class="text" align=center> 
              <%if abs(users.Fields.Item("user_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
            </td>
            <td width="20" class="text"><%=(users.Fields.Item("user_logcount").Value)%>x</td>
            <td width="20" class="text"><%=(users.Fields.Item("session_count").Value)%></td>
            <td width="20" class="text"> 
              <%
			response.write("<font color = green>" & FormatNumber(user_total_rate,2) & "%</font>") 
			count = count + 1
			%>
            </td>
            <td  align="right" width="20"> 
              <a href="q_user_edit.asp?user=<%=(users.Fields.Item("ID_user").Value)%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&show_lines=<%=show_lines%>"><img src="images/edit.gif" width="16" height="15" border="0"></a> 
            </td>
          </tr>     
  <%
  
  else if (cstr(results)="0") and (user_total_rate <= passrate) and (user_session_count<>0) then%>
			<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
            <td class="text" width="20"><%=count%></td>
            <td width="200" class="text"> 
              <%if cInt(users.Fields.Item("session_count").Value) > 0 then%>
              <a href="q_user_sessions.asp?user=<%=(users.Fields.Item("ID_user").Value)%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&fromdate=<%=fromdate%>&todate=<%=todate%>&show_lines=<%=show_lines%>">
              <%
              end if%>
              <%=(users.Fields.Item("user_lastname").Value)%>&nbsp;<%=(users.Fields.Item("user_firstname").Value)%></a></td>
            <td width="140" class="text"><%=(users.Fields.Item("info1").Value)%> (<%=(users.Fields.Item("info2").Value)%>)</td>
            <td width="140" class="text"><%=(users.Fields.Item("info3").Value)%></td>
            <td width="30" class="text" align=center> 
              <%if abs(users.Fields.Item("user_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
            </td>
            <td width="20" class="text"><%=(users.Fields.Item("user_logcount").Value)%>x</td>
            <td width="20" class="text"><%=(users.Fields.Item("session_count").Value)%></td>
            <td width="20" class="text"> 
              <%
			response.write("<font color = red>" & FormatNumber(user_total_rate,2) & "%</font>")
			count = count + 1
			%>
            </td>
            <td  align="right" width="20"> 
              <a href="q_user_edit.asp?user=<%=(users.Fields.Item("ID_user").Value)%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&show_lines=<%=show_lines%>"><img src="images/edit.gif" width="16" height="15" border="0"></a> 
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
users.Requery
%>
<%
'_______________________________________________________________________________
overall_session_rate = 0
overall_session_count = 0
overall_session_passed = 0
numbers=1
%>
          <% 
While (NOT users.EOF)
%>
          <%
if subject_prm <> 0 then
	subj_prm ="and (q_session.Session_subject ="&subject_prm&")"
else 
	subj_prm =""
end if

if session("mths") <> "" then
	 mths_prm = "and  q_session.session_date between '"&back_date&"' and '"&now_date&"'"
else
	mths_prm = ""
end if	

if session("mths") <> "" then
	 mths_prm = "and  q_session.session_date between '"&back_date&"' and '"&now_date&"'"
else
	mths_prm = ""
end if	

if cstr(fromdate)="" and cstr(todate) <> "" then
	user_details.Source = "SELECT q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) and (session_finish <= '"&todate&"') "&subj_prm&" "&mths_prm&""
else if cstr(todate)="" and cstr(fromdate) <> "" then
	user_details.Source = "SELECT q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) and (session_finish >= '"&fromdate&"') "&subj_prm&" "&mths_prm&" "
else if (cstr(todate)="" and cstr(fromdate)="") then
	user_details.Source = "SELECT q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) "&subj_prm&" "&mths_prm&""
else
	user_details.Source = "SELECT q_session.session_users,q_session.Session_total, q_session.Session_correct, q_session.session_finish FROM q_user INNER JOIN q_session ON q_user.ID_user = q_session.Session_users WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") AND (q_session.Session_done = 1) and ((session_finish >= '"&fromdate&"') and (session_finish <= '"&todate&"')) "&subj_prm&" "&mths_prm&""
end if 
end if
end if		
user_details.Open()
user_details_numRows = 0
%>
          <%
user_session_rate = 0
user_session_count = 0
user_total_rate = 0
While (NOT user_details.EOF)
%>
          <%
user_session_correct = cInt(user_details.Fields.Item("session_correct").Value)
user_session_total = cInt(user_details.Fields.Item("session_total").Value)
user_session_rate = user_session_rate + (user_session_correct / user_session_total * 100)
user_session_count = user_session_count + 1
%>
          <%
  user_details.MoveNext()

Wend
%>
          <%
user_details.Close()
%>
          <%
         
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
%>
          <% 
         
  users.MoveNext()
  numbers=numbers+1
Wend
%>
          <%
if overall_session_count > 0 then overll_pass_rate = overall_session_rate/overall_session_count
'________________________________________________________________________________
%>          
          <tr> 
            <td class="text" colspan="9"> 
              <hr>
            </td>
          </tr>
          <tr> 
            <td class="subheads" colspan="9">Overall results of filtered users : </td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" colspan="2" align="left" valign="top">Number of users 
              in selection :</td>
            <td colspan="2" class="text" align="left" valign="top">Users 
              with at least 1 finished session :<br>
              Total / <font color = green>passed</font> / <font color = red>failed</font> </td>
            <td colspan="5" class="text" align="left" valign="top">Avg. % of users with at least 1 finished session :</td>
          </tr>
          <tr class="table_normal"> 
			<% 
			if (cstr(results)<> "2" and  cstr(results) <> "") or cstr(noquiz)= "1" then
			%>
            <td class="text" colspan="2"><%=count-1%><%cnt = count -1%></td>
            <%else
            %>
             <td class="text" colspan="2"><%=numbers-1%><%cnt = numbers-1%></td>
             <%end if%>
            
            <%if cstr(noquiz)="1" then
            overall_session_count = 0
            overall_session_passed = 0
            end if
            
            %>
            
            <td colspan="2" class="text"><%=overall_session_count & " / <font color = green>" & overall_session_passed & "</font> / <font color = red>" & (overall_session_count - overall_session_passed) & "</font>"%></td>
            <td colspan="5" class="text"> 
              <%
			
			if overall_session_count = 0 then
				response.write("<font color = blue>Nothing to rate</font>") 
			elseif overll_pass_rate >= passrate then 
				response.write("<font color = green>" & FormatNumber(overll_pass_rate,2) & "% - PASSED</font>") 
			else 
				response.write("<font color = red>" & FormatNumber(overll_pass_rate,2) & "% - FAILED</font>")
			end if
			%>
            
          </tr>
          <tr>
          </td>
          <tr class="table_normal">
            <td class="text" colspan=9>
           <table>
            <tr>
             <td class="text">Percentage of users who have completed a quiz</td>
             <%
             if cint(cnt)= 0 then cnt =1
             %>
             <td class="text">:<b> <%=FormatNumber((100 * (overall_session_count))/(cnt),2) %>%</b></td>
            </tr>
            <tr>
              <td class="text">Percentage of users who have successfully passed</td>
              <td class="text">:<b><font color = green> <%=FormatNumber((100 * (overall_session_passed))/(cnt),2) %>%</font></b></td></tr>
            </tr>
            <tr>
              <td class="text">Percentage of users who have failed</td>
              <td class="text">:<b><font color = red> <%=FormatNumber((100 * (overall_session_count - overall_session_passed))/(cnt),2) %>%</font></b></td></tr>
            </tr>
            </table>
            </td>
            <%=no_quiz%>
          </tr>
          <tr> 
            <td  colspan="9"><i>Pass 
              rate is currently: <%'=passrate%><input type="text" class="formitem1" name="passrate" size=3 maxlength=3 value="">%  <input type="button" name="pass" value="Change Pass Rate" class="quiz_button" onclick="return pass_submit();"><br>
              Please bear in mind, that above OVERALL figures reflect the whole 
              selection, NOT the page as shown on screen. I.e. if only 50 out 
              of 65 users are shown on the screen, the Pass &amp; Rate covers 
              all 65 users, including those onnext page!!!</i></td>
          </tr>
          <% End If %>
          <% If users.EOF And users.BOF Then %>
          <tr> 
            <td  width="18">&nbsp;<input type="hidden" class="formitem1" name="passrate"  size=3 value=""></td>
            <td colspan="8" >Sorry, 
              there are no users in the quiz currently or no user match your filter 
              criteria.</td>
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
            <td colspan="8" >&nbsp; 
              <table width="50%" align="center">
                <tr class="table_normal"> 
                  <td width="25%" align="center" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
                    <% If MM_offset <> 0 Then %>
                    <a href="<%=MM_moveFirst%>"><img src="images/first.gif" border=0></a> 
                    <% End If %>
                  </td>
                  <td width="25%" align="center" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
                    <% If MM_offset <> 0 Then %>
                    <a href="<%=MM_movePrev%>"><img src="images/previous.gif" border=0></a> 
                    <% End If%>
                  </td>
                  <td width="25%" align="center" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
                    <% If Not MM_atTotal Then %>
                    <a href="<%=MM_moveNext%>"><img src="images/next.gif" border=0></a> 
                    <% End If%>
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
%>
<%
users.Close()
Set users = Nothing
%>
<%
filter_info1.Close()
%>
<%
filter_info3.Close()
%>
<%
subjects.Close()
%>

