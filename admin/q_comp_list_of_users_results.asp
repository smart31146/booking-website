<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
numbers=1
count = 0 'User count
noquizcount = 1 ' user count who have done no Quiz session
rowcount = 1 'Display row count
SQL_having = ""
SQL_where = ""
if Request.Cookies("show_lines")<> "" then
	show_lines= cint(Request.Cookies("show_lines"))
else
	show_lines=5
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
'passrate =  request("passrate")
if mths="" then
	mths=0
end if
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

filter_info4_prm = 0
If cInt(Request.Querystring("filter_info4")) <> 0 then
	filter_info4_prm = cInt(Request.Querystring("filter_info4"))
	if SQL_having <> "" then
		SQL_having = SQL_having + " AND (q_user.user_info4)= " + (Request.Querystring("filter_info4")) + " "
	else
		SQL_having = " HAVING (q_user.user_info4)= " + (Request.Querystring("filter_info4")) + " "
	end if
end if

if request("mths")="1" then
	session("mths") = "1"
else
	session("mths")=""
end if
set users = Server.CreateObject("ADODB.Recordset")
users.ActiveConnection = Connect
users.Source = "SELECT q_user.ID_user, q_user.user_lastname, q_user.user_firstname, q_info1.info1, q_info2.info2, q_info3.info3, q_info4.info4, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3, q_user.user_info4, COUNT(q_session.ID_session) AS session_count FROM (q_info4 RIGHT JOIN (q_info3 RIGHT JOIN (q_info2 RIGHT JOIN (q_info1 RIGHT JOIN q_user ON q_info1.ID_info1 = q_user.user_info1) ON q_info2.ID_info2 = q_user.user_info2) ON q_info3.ID_info3 = q_user.user_info3) ON q_info4.ID_info4 = q_user.user_info4) LEFT JOIN q_session ON q_user.ID_user = q_session.Session_users " + SQL_where + " GROUP BY q_user.user_lastname, q_user.user_firstname, q_user.ID_user, q_info1.info1, q_info2.info2, q_info3.info3, q_info4.info4, q_user.user_active, q_user.user_logcount, q_user.user_info1, q_user.user_info3, q_user.user_info4 " + SQL_having + " ORDER BY q_user.user_lastname, q_user.user_firstname;"
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

set filter_info4 = Server.CreateObject("ADODB.Recordset")
filter_info4.ActiveConnection = Connect
filter_info4.Source = "SELECT * FROM q_info4 order by info4"
filter_info4.CursorType = 0
filter_info4.CursorLocation = 3
filter_info4.LockType = 3
filter_info4.Open()
filter_info4_numRows = 0

if subject_prm <> 0 then
	subj_prm ="and (ID_subject ="&subject_prm&")"
else
	subj_prm =""
end if

set subjects = Server.CreateObject("ADODB.Recordset")
subjects.ActiveConnection = Connect
subjects.Source = "SELECT ID_subject, subject_name FROM subjects WHERE subject_active_q <> 0 "&subj_prm
subjects.CursorType = 0
subjects.CursorLocation = 3
subjects.LockType = 3
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
	//if (((document.filter_users.fromdate.value != "") && (document.filter_users.todate.value != "")) || ((document.filter_users.fromdate.value == "") && (document.filter_users.todate.value != "")) || ((document.filter_users.fromdate.value != "") && (document.filter_users.todate.value == "")))
	//{
	//	document.filter_users.mths.checked = false
	//	document.filter_users.mths.disabled=true;
	//	return;
	//}
	//else
	//{
	//	document.filter_users.mths.disabled=false;
	//	return;
//	}
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
	show_lines = MyCookie.Read('show_lines')
	if (show_lines != null) {
		//alert('1');
		document.forms[0].show_lines.value=show_lines;
	}
	else {
		//alert('2');
		document.forms[0].show_lines.value=5;
	}
}

function AddCookieId(cn,id) {
        MyCookie.Write(cn,id,7);
}

function DelCookieId(cn,id) {
        MyCookie.Write(cn,id,-1);
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
document.forms[0].filter_info4.selectedIndex = 0;
//document.forms[0].show_lines.value = "25";
document.forms[0].submit();
}
//-->
</script>
</HEAD>

<BODY onload="check();">
<%
	if Request.Cookies("show_lines")<> "" then
		show_lines= cint(Request.Cookies("show_lines"))
	else
		show_lines=5
	end if
%>
<table>
  <tr>
    <td align="left" valign="bottom" class="heading"> Quiz users - combined results</td>
  </tr>
  <tr>
    <td align="left" valign="bottom">
      <form name="filter_users">
     <input type="hidden" name="hiddenmths" value="false">
        <table>
          <tr>
			  <td colspan="6" class="subheads" align="left" valign="top">Users:</td>
              <td align="right" class="subheads" valign="top" colspan="9"><a href="q_comp_export_quiz_users.asp?filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&filter_info4=<%=filter_info4_prm%>&fromdate=<%=fromdate%>&todate=<%=todate%>&active=<%=active%>&results=<%=results%>&passrate=<%=passrate%>&mths=<%=mths%>&noquiz=<%=noquiz%>"><img src="images/xls.gif" width="16" height="16" border="0"></a>&nbsp;Export this screen to Excel</td>
			  <!-- <td align="right" class="subheads" valign="top" width="6"><a href="q_export_quiz_users_summary.asp?filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&fromdate=<%=fromdate%>&todate=<%=todate%>&active=<%=active%>&results=<%=results%>&passrate=<%=passrate%>&mths=<%=mths%>&noquiz=<%=noquiz%>"><img src="images/summary.gif" width="16" height="16" border="0"></a></td> -->
          </tr>
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">
            <td class="text" width="18"><img src="images/back.gif" width="18" height="14"></td>
            <td class="text" colspan="14"><a href="q_comp_list_of_users.asp">...Filters
              </a> </td>
          </tr>

          <tr>
            <td >&nbsp;</td>
            <td >Last name &amp; First name</td>
            <td >Business</td>
            <td ><% =BBPinfo3 %></td>
            <td >Company</td>
           <!-- <td >Active</td>-->
            <td >Logs</td>
            <td >Sess.</td>
            <td >Edit</td>
            <td >Subject</td>
            <td >Date</td>
            <td >Corr.</td>
            <td >Total</td>
            <td >Done</td>
            <td >Fin.</td>
            <td >Rate</td>
            <td >Pass</td>
			<!-- <td >Merge</td> SS 050707: merge not required-->
          </tr>
          <%
		  If Not users.EOF Or Not users.BOF Then
			While ((Repeat1__numRows <> 0) AND (NOT users.EOF))
				count = count + 1
				'if (cstr(noquiz)="1") then
				if cInt(users.Fields.Item("session_count").Value) = 0 then
					'SS 050708: Display all subjects
					subjects.Source = "SELECT subjects.ID_subject, subjects.subject_name FROM subjects inner join subject_user on subjects.id_subject=subject_user.id_subject where subject_user.id_user="&users.Fields.Item("ID_user").Value&" and subjects.subject_active_q <> 0 "&subj_prm
					subjects.Open()
					While (NOT subjects.EOF)
							'pn 050812 ensure that subjects do not appear if they are not this users subjects
							Dim user_has_subject
							user_has_subject=false
							set user_subject = Server.CreateObject("ADODB.Recordset")
							user_subject.ActiveConnection = Connect
							user_subject.Source = "SELECT * FROM subject_user where Id_Subject="&subjects.Fields.Item("ID_subject")&" and ID_user="&users.Fields.Item("ID_user").Value&";"
							user_subject.CursorType = 0
							user_subject.CursorLocation = 3
							user_subject.LockType = 3
							user_subject.Open()
							While (NOT user_subject.EOF)
								user_has_subject=true
								user_subject.MoveNext()
							Wend
							user_subject.Close()
						'only show this subject if the user has it
						if user_has_subject=true then
		  %>
						<!-- SS 050707: Users who have not done a single quiz -->
						<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">

							<td class="text"><%=rowcount%></td>
							<td class="text">
							 <%=(users.Fields.Item("user_lastname").Value)%>&nbsp;<%=(users.Fields.Item("user_firstname").Value)%></td>
							<td class="text"><%=(users.Fields.Item("info1").Value)%> (<%=(users.Fields.Item("info2").Value)%>)</td>
							<td class="text"><%=(users.Fields.Item("info3").Value)%></td>
							<td class="text"><%=(users.Fields.Item("info4").Value)%></td>
							<!--  <td class="text" align=center>
							<%if abs(users.Fields.Item("user_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
							</td>-->
							<td class="text"><%=(users.Fields.Item("user_logcount").Value)%>x</td>
							<td class="text"><%=(users.Fields.Item("session_count").Value)%></td>
							<td class="text" align="right">
							  <a href="q_user_edit.asp?user=<%=(users.Fields.Item("ID_user").Value)%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&filter_info4=<%=filter_info4_prm%>&show_lines=<%=show_lines%>"><img src="images/edit.gif" width="16" height="15" border="0"></a>
							</td>
							<td class="text" align="left"><%=subjects.Fields.Item("subject_name")%></td>
							<td class="text" align="center"><font color = blue>-</font></td>
							<td class="text" align="center"><font color = blue>-</font></td>
							<td class="text" align="center"><font color = blue>-</font></td>
							<td class="text" align="center"><font color = blue>-</font></td>
							<td class="text" align="center"><font color = blue>-</font></td>
							<td class="text" align="center"><font color = blue>-</font></td>
							<td class="text" align="center"><font color = blue>-</font></td>
							<!--PN 040811 add merge facility so that users can be merged into this user to get rid of self reg duplicates-->
							<!-- SS 050707: Merge functionality not required
								<td  align="right" width="18">
								<a href="#" onclick="var wintoopen=window.open('q_list_of_users_to_merge.asp?user=<%=(users.Fields.Item("ID_user").Value)%>','merge','toolbar=0, scrollbars=yes,resizable=1,width=700, height=500');wintoopen.focus();"><img src="images/merge.gif" alt="Merge users with this user" width="15" height="15" border="0"></a>
							</td>
							-->
						</tr>

		  <%
				  end if
						subjects.MoveNext()
						rowcount = rowcount + 1
					Wend
					subjects.Close()
					'noquizcount = noquizcount + 1
					'End if
				else

					'SS 050708: Go through subjects.
					subjects.Source = "SELECT subjects.ID_subject, subjects.subject_name FROM subjects inner join subject_user on subjects.id_subject=subject_user.id_subject where subject_user.id_user="&users.Fields.Item("ID_user").Value&" and subjects.subject_active_q <> 0 "&subj_prm
					subjects.Open()
					While (NOT subjects.EOF)
					
					'Gets the passrate based upon each subject
					currentSubjectID = subjects.Fields.Item("ID_subject")
					currentSubjectIDstr = "and (q_session.Session_subject ="&currentSubjectID&")"
					set preferences = Server.CreateObject("ADODB.Recordset")
							preferences.ActiveConnection = Connect
							preferences.Source = "SELECT subject_passmark FROM subjects WHERE ID_subject="&currentSubjectID&""
							preferences.CursorType = 0
							preferences.CursorLocation = 3
							preferences.LockType = 3
							preferences.Open()
							preferences_numRows = 0
							passrate = preferences.Fields.Item("subject_passmark").Value
							preferences.close()

							
						'SS 050706: Display all latest session information at this level....no drill down
						if cstr(fromdate)="" and cstr(todate) <> "" then
							user_details.Source = "SELECT TOP 1 q_session.ID_Session, q_session.session_subject, q_session.Session_date, subjects.id_subject,subjects.subject_name, q_session.Session_total, q_session.Session_done, q_session.Session_correct, q_session.Session_stop, q_session.session_finish  FROM q_user INNER JOIN (q_session INNER JOIN subjects ON q_session.Session_subject = subjects.ID_subject) ON q_user.ID_user = q_session.Session_users  WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") and (session_finish <= '"&todate&"') "&currentSubjectIDstr&" AND session_done = 1 order by session_date desc"
						else if cstr(todate)="" and cstr(fromdate) <> "" then
							user_details.Source = "SELECT TOP 1 q_session.ID_Session, q_session.session_subject, q_session.Session_date, subjects.id_subject,subjects.subject_name, q_session.Session_total, q_session.Session_done, q_session.Session_correct, q_session.Session_stop, q_session.session_finish  FROM q_user INNER JOIN (q_session INNER JOIN subjects ON q_session.Session_subject = subjects.ID_subject) ON q_user.ID_user = q_session.Session_users  WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") and (session_finish >= '"&fromdate&"') "&currentSubjectIDstr&" AND session_done = 1 order by session_date desc"
						else if (cstr(todate)="" and cstr(fromdate)="") then
							user_details.Source = "SELECT TOP 1 q_session.ID_Session, q_session.session_subject, q_session.Session_date, subjects.id_subject,subjects.subject_name, q_session.Session_total, q_session.Session_done, q_session.Session_correct, q_session.Session_stop, q_session.session_finish  FROM q_user INNER JOIN (q_session INNER JOIN subjects ON q_session.Session_subject = subjects.ID_subject) ON q_user.ID_user = q_session.Session_users  WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") "&currentSubjectIDstr&" AND session_done = 1 order by session_date desc"
						else
							user_details.Source = "SELECT TOP 1 q_session.ID_Session, q_session.session_subject, q_session.Session_date, subjects.id_subject,subjects.subject_name, q_session.Session_total, q_session.Session_done, q_session.Session_correct, q_session.Session_stop, q_session.session_finish  FROM q_user INNER JOIN (q_session INNER JOIN subjects ON q_session.Session_subject = subjects.ID_subject) ON q_user.ID_user = q_session.Session_users  WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") and ((session_finish >= '"&fromdate&"') and (session_finish <= '"&todate&"')) "&currentSubjectIDstr&" AND session_done = 1 order by session_date desc"
						end if
						end if
						end if
						'response.write user_details.Source
						user_details.Open()
						user_details_numRows = 0

						'SS 050708: Display all subjects
						If Not user_details.EOF Or Not user_details.BOF Then
							While (NOT user_details.EOF)

								session_done = abs(user_details.Fields.Item("Session_done").Value)
								user_session_rate = 0
								user_total_rate = 0
								user_pass = 0
								subid = 0
								if (session_done = 1) then
									if session("mths")="" then
										user_session_correct = cInt(user_details.Fields.Item("session_correct").Value)
										user_session_total = cInt(user_details.Fields.Item("session_total").Value)
										'user_session_rate = user_session_rate + (user_session_correct / user_session_total * 100)
										user_session_rate = FormatNumber((user_session_correct / user_session_total * 100),2)
										'user_session_count = user_session_count + 1
									else if cint(subid) <> cInt(user_details.Fields.Item("session_subject").Value) then
										subid = cInt(user_details.Fields.Item("session_subject").Value)
										user_session_correct = cInt(user_details.Fields.Item("session_correct").Value)
										user_session_total = cInt(user_details.Fields.Item("session_total").Value)
										'user_session_rate = user_session_rate + (user_session_correct / user_session_total * 100)
										user_session_rate = FormatNumber((user_session_correct / user_session_total * 100),2)
										'user_session_count = user_session_count + 1
									end if
									end if
									if cInt(user_session_rate) >= cInt(passrate) then user_pass = 1 else user_pass = 0
								end if

								%>
									<!-- SS 050707: html to display records -->
									<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">

										<td class="text"><%=rowcount%></td>
										<td class="text">
										 <%=(users.Fields.Item("user_lastname").Value)%>&nbsp;<%=(users.Fields.Item("user_firstname").Value)%></td>
										<td class="text"><%=(users.Fields.Item("info1").Value)%> (<%=(users.Fields.Item("info2").Value)%>)</td>
										<td class="text"><%=(users.Fields.Item("info3").Value)%></td>
										<td class="text"><%=(users.Fields.Item("info4").Value)%></td>
										<td class="text"><%=(users.Fields.Item("user_logcount").Value)%>x</td>
										<td class="text"><%=(users.Fields.Item("session_count").Value)%></td>
										<td class="text" align="right">
										  <a href="q_user_edit.asp?user=<%=(users.Fields.Item("ID_user").Value)%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&filter_info4=<%=filter_info4_prm%>&show_lines=<%=show_lines%>"><img src="images/edit.gif" width="16" height="15" border="0"></a>
										</td>
										<td class="text" align="left"><a href="q_session_details.asp?user_session=<%=(user_details.Fields.Item("ID_Session").Value)%>&user=<%=(users.Fields.Item("ID_user").Value)%>&subject=<%=(user_details.Fields.Item("id_subject").Value)%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&filter_info4=<%=filter_info4_prm%>&todate=<%=todate%>&active=<%=active%>&results=<%=results%>&passrate=<%=passrate%>&mths=<%=request("mths")%>&backsubject=<%=subject_prm%>&noquiz=0""><%=(user_details.Fields.Item("subject_name").Value)%></a></td>
										<td class="text"><%=(user_details.Fields.Item("Session_finish").Value)%></td>
										<td class="text"><%=(user_details.Fields.Item("Session_correct").Value)%></td>
										<td class="text"><%=(user_details.Fields.Item("Session_total").Value)%></td>
										<td class="text"><%=(user_details.Fields.Item("Session_stop").Value)%></td>
										<td class="text" align="center">
										<%
											if session_done = 1 then
												response.write "<img src='images/1.gif'>"
											else
												response.write "<img src='images/0.gif'>"
											end if
										%>
										</td>
										<td class="text" align="center">
										<%
											if session_done = 1 then
												if user_pass = 1 then response.write ("<font color=green>" & user_session_rate & "%") else response.write ("<font color=red>" & user_session_rate & "%")
											else
												response.write("<font color = blue>-</font>")
											end if
										%>
										</td>
										<td class="text" align="center">
										<%
											if session_done = 1 then
												if user_pass = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"
											else
												response.write("<font color = blue>-</font>")
											end if
										%>
										</td>
										<!--PN 040811 add merge facility so that users can be merged into this user to get rid of self reg duplicates-->
										<!-- SS 050707: Merge functionality not required
											<td  align="right" width="18">
											<a href="#" onclick="var wintoopen=window.open('q_list_of_users_to_merge.asp?user=<%=(users.Fields.Item("ID_user").Value)%>','merge','toolbar=0, scrollbars=yes,resizable=1,width=700, height=500');wintoopen.focus();"><img src="images/merge.gif" alt="Merge users with this user" width="15" height="15" border="0"></a>
										</td>
										-->
									</tr>

								<%
								user_details.MoveNext()
							Wend
						else
						%>

							<!-- SS 050708: Also show subjects for which no session -->
							<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">

								<td class="text"><%=rowcount%></td>
								<td class="text">
								 <%=(users.Fields.Item("user_lastname").Value)%>&nbsp;<%=(users.Fields.Item("user_firstname").Value)%></td>
								<td class="text"><%=(users.Fields.Item("info1").Value)%> (<%=(users.Fields.Item("info2").Value)%>)</td>
								<td class="text"><%=(users.Fields.Item("info3").Value)%></td>
								<td class="text"><%=(users.Fields.Item("info4").Value)%></td>

								<td class="text"><%=(users.Fields.Item("user_logcount").Value)%>x</td>
								<td class="text"><%=(users.Fields.Item("session_count").Value)%></td>
								<td class="text" align="right">
								  <a href="q_user_edit.asp?user=<%=(users.Fields.Item("ID_user").Value)%>&filter_username=<%=(Request.Querystring("filter_username"))%>&subject=<%=subject_prm%>&filter_info1=<%=filter_info1_prm%>&filter_info3=<%=filter_info3_prm%>&show_lines=<%=show_lines%>"><img src="images/edit.gif" width="16" height="15" border="0"></a>
								</td>
								<td class="text" align="left"><%=subjects.Fields.Item("subject_name")%></td>
								<td class="text" align="center"><font color = blue>-</font></td>
								<td class="text" align="center"><font color = blue>-</font></td>
								<td class="text" align="center"><font color = blue>-</font></td>
								<td class="text" align="center"><font color = blue>-</font></td>
								<td class="text" align="center"><font color = blue>-</font></td>
								<td class="text" align="center"><font color = blue>-</font></td>
								<td class="text" align="center"><font color = blue>-</font></td>
								<!--PN 040811 add merge facility so that users can be merged into this user to get rid of self reg duplicates-->
								<!-- SS 050707: Merge functionality not required
									<td  align="right" width="18">
									<a href="#" onclick="var wintoopen=window.open('q_list_of_users_to_merge.asp?user=<%=(users.Fields.Item("ID_user").Value)%>','merge','toolbar=0, scrollbars=yes,resizable=1,width=700, height=500');wintoopen.focus();"><img src="images/merge.gif" alt="Merge users with this user" width="15" height="15" border="0"></a>
								</td>
								-->
							</tr>
						<%
						end if
						user_details.Close()
						subjects.MoveNext()
						rowcount = rowcount + 1
					Wend 'end of subject loop
					subjects.Close()
				End if 'end of noquiz if loop
				Repeat1__index=Repeat1__index+1
				Repeat1__numRows=Repeat1__numRows-1
				users.MoveNext()
				'numbers=numbers+1
			Wend

' SS 050707: END - From here down comment the code


'If (users.CursorType > 0) Then
'  users.MoveFirst
'Else
  users.Requery
'End If

'_______________________________________________________________________________
overall_session_rate = 0
overall_session_count = 0
overall_session_passed = 0
numbers=1
While (NOT users.EOF)

	'SS 050708: Go through subjects.
	subjects.Open()
	While (NOT subjects.EOF)
		currentSubjectID = subjects.Fields.Item("ID_subject")
		currentSubjectIDstr = "and (q_session.Session_subject ="&currentSubjectID&")"

		if cstr(fromdate)="" and cstr(todate) <> "" then
			user_details.Source = "SELECT TOP 1 q_session.ID_Session, q_session.session_subject, q_session.Session_date, subjects.id_subject,subjects.subject_name, q_session.Session_total, q_session.Session_done, q_session.Session_correct, q_session.Session_stop, q_session.session_finish  FROM q_user INNER JOIN (q_session INNER JOIN subjects ON q_session.Session_subject = subjects.ID_subject) ON q_user.ID_user = q_session.Session_users  WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") and (session_finish <= '"&todate&"') "&currentSubjectIDstr&" order by session_date desc"
		else if cstr(todate)="" and cstr(fromdate) <> "" then
			user_details.Source = "SELECT TOP 1 q_session.ID_Session, q_session.session_subject, q_session.Session_date, subjects.id_subject,subjects.subject_name, q_session.Session_total, q_session.Session_done, q_session.Session_correct, q_session.Session_stop, q_session.session_finish  FROM q_user INNER JOIN (q_session INNER JOIN subjects ON q_session.Session_subject = subjects.ID_subject) ON q_user.ID_user = q_session.Session_users  WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") and (session_finish >= '"&fromdate&"') "&currentSubjectIDstr&" order by session_date desc"
		else if (cstr(todate)="" and cstr(fromdate)="") then
			user_details.Source = "SELECT TOP 1 q_session.ID_Session, q_session.session_subject, q_session.Session_date, subjects.id_subject,subjects.subject_name, q_session.Session_total, q_session.Session_done, q_session.Session_correct, q_session.Session_stop, q_session.session_finish  FROM q_user INNER JOIN (q_session INNER JOIN subjects ON q_session.Session_subject = subjects.ID_subject) ON q_user.ID_user = q_session.Session_users  WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") "&currentSubjectIDstr&" order by session_date desc"
		else
			user_details.Source = "SELECT TOP 1 q_session.ID_Session, q_session.session_subject, q_session.Session_date, subjects.id_subject,subjects.subject_name, q_session.Session_total, q_session.Session_done, q_session.Session_correct, q_session.Session_stop, q_session.session_finish  FROM q_user INNER JOIN (q_session INNER JOIN subjects ON q_session.Session_subject = subjects.ID_subject) ON q_user.ID_user = q_session.Session_users  WHERE (q_session.Session_users = " & (users.Fields.Item("ID_user").Value) & ") and ((session_finish >= '"&fromdate&"') and (session_finish <= '"&todate&"')) "&currentSubjectIDstr&" order by session_date desc"
		end if
		end if
		end if
		user_details.Open()
		user_details_numRows = 0

		user_session_rate = 0
		user_session_count = 0
		user_total_rate = 0
		subid =0

		If Not user_details.EOF Or Not user_details.BOF Then
			While (NOT user_details.EOF)
				session_done = 0
				session_done = abs(user_details.Fields.Item("Session_done").Value)
				user_session_rate = 0
				user_pass = 0
				if (session_done = 1) then
					overall_session_count = overall_session_count + 1
					if session("mths")="" then
						user_session_correct = cInt(user_details.Fields.Item("session_correct").Value)
						user_session_total = cInt(user_details.Fields.Item("session_total").Value)
						user_session_rate = FormatNumber((user_session_correct / user_session_total * 100),2)
					else if cint(subid) <> cInt(user_details.Fields.Item("session_subject").Value) then
						user_session_correct = cInt(user_details.Fields.Item("session_correct").Value)
						user_session_total = cInt(user_details.Fields.Item("session_total").Value)
						user_session_rate = FormatNumber((user_session_correct / user_session_total * 100),2)
					end if
					end if
					if cInt(user_session_rate) >= cInt(passrate) then user_pass = 1 else user_pass = 0
					if user_pass = 1 then
						overall_session_passed = overall_session_passed + 1
					end if
				end if
				user_details.MoveNext()
			Wend
		end if
		user_details.Close()
		subjects.MoveNext()
	Wend
	subjects.Close()
	users.MoveNext()
	numbers=numbers+1
Wend
if overall_session_count > 0 then overll_pass_rate = overall_session_rate/overall_session_count
'________________________________________________________________________________
%>
          <tr>
            <td class="text" colspan="15">
              <hr>
            </td>
          </tr>
          <tr>
            <td class="subheads" colspan="15">Overall results of filtered users : </td>
          </tr>
          <tr class="table_normal">
            <td class="text" colspan="2" align="left" valign="top">Number of users in selection :</td>
            <td colspan="6" class="text" align="left" valign="top">Number of completed quizes in selection:<br>
              Total / <font color = green>passed</font> / <font color = red>failed</font> </td>
          </tr>
          <tr class="table_normal">
            <%
			if cstr(noquiz)="1" then
				overall_session_count = 0
				overall_session_passed = 0
			%>
				<td class="text" colspan="2"><%=noquizcount-1%><%cnt = noquizcount -1%></td>
			<%
			else
				if (cstr(results)<> "2" and  cstr(results) <> "") then
			%>
					<td class="text" colspan="2"><%=count%><%cnt = count%></td>
            <%
				else
            %>
					<td class="text" colspan="2"><%=numbers-1%><%cnt = numbers-1%></td>
            <%
				end if
            end if
            %>


            <td colspan="6" class="text"><%=overall_session_count & " / <font color = green>" & overall_session_passed & "</font> / <font color = red>" & (overall_session_count - overall_session_passed) & "</font>"%></td>
          </tr>
		  <tr>
            <td  colspan="15"><i>
              <br><br> Please bear in mind, that the OVERALL results above reflect the whole
              selection of filtered users. If there is more than one page of users identified as
              matching your selection, the overall results will capture all the users on all the pages. i.e. if only 50 out
              of 65 users are shown on the screen, the results cover all 65 users, including those on the next page.</i></td>
          </tr>
          <% End If %>
          <% If users.EOF And users.BOF Then %>
          <tr>
            <td  width="18">&nbsp;<input type="hidden" class="formitem1" name="passrate"  size=3 value=""></td>
            <td colspan="14" >Sorry,
              there are no users in the quiz currently or no user match your filter
              criteria.</td>
          </tr>
          <% End If %>

			<tr>
            <td  width="36">&nbsp;</td>
            <td colspan="10" >&nbsp;
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
filter_info4.Close()
%>

