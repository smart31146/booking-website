<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<!--#include file="sha256.asp"-->
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
Function Tokenize(byVal TokenString, byRef TokenSeparators())

	Dim NumWords, a()
	NumWords = 0

	Dim NumSeps
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
' *** Redirect if username exists

MM_flag="MM_update"
If (CStr(Request(MM_flag)) <> "") Then
	Response.Write CStr(Request(MM_flag))
  current_user_id = cInt(request("MM_recordId"))
  MM_dupKeyRedirect="q_user_duplicate.asp"
  MM_rsKeyConnection=Connect
  MM_dupKeyUsernameValue = Replace(CStr(Request.Form("login_name")), "'", "''")
  MM_dupKeySQL="SELECT * FROM q_user WHERE user_username='" & MM_dupKeyUsernameValue & "' AND ID_user <> " & current_user_id & ";"

 MM_adodbRecordset="ADODB.Recordset"
  set MM_rsKey=Server.CreateObject(MM_adodbRecordset)
  MM_rsKey.ActiveConnection=MM_rsKeyConnection
  MM_rsKey.Source=MM_dupKeySQL
  MM_rsKey.CursorType=0
  MM_rsKey.CursorLocation=2
  MM_rsKey.LockType=3
  MM_rsKey.Open
  If  Not MM_rsKey.EOF Or Not MM_rsKey.BOF Then
    ' the username was found - can not add the requested username
    MM_qsChar = "?"
    If (InStr(1,MM_dupKeyRedirect,"?") >= 1) Then MM_qsChar = "&"
    MM_dupKeyRedirect = MM_dupKeyRedirect & MM_qsChar & "requsername=" & MM_dupKeyUsernameValue
    Response.Redirect(MM_dupKeyRedirect)
  End If
  MM_rsKey.Close
End If
%>
<%
' *** Redirect if firstname, lastname, city and month of birth exists
'pn 050816 remove this check to prevent duplicate names being entered
			'MM_flag="MM_update"
			'If (CStr(Request(MM_flag)) <> "") Then
			  'current_user_id = cInt(request("MM_recordId"))
			 ' MM_dupKeyRedirect="q_user_duplicate2.asp"
			 ' MM_rsKeyConnection=Connect
			 ' MM_dupKeyUsernameValue1 = CStr(Request.Form("first_name"))
			  'MM_dupKeyUsernameValue2 = CStr(Request.Form("last_name"))
			  'MM_dupKeyUsernameValue3 = CStr(Request.Form("month"))
			  'MM_dupKeyUsernameValue4 = CStr(Request.Form("city"))
			'  MM_dupKeySQL="SELECT * FROM q_user WHERE user_firstname='" & MM_dupKeyUsernameValue1 & "' AND user_lastname='" & MM_dupKeyUsernameValue2 & "' AND ID_user <> " & current_user_id & ";"
			  'AND user_month=" & MM_dupKeyUsernameValue3 & " AND user_city='" & MM_dupKeyUsernameValue4 & "'
			'  MM_adodbRecordset="ADODB.Recordset"
			 ' set MM_rsKey=Server.CreateObject(MM_adodbRecordset)
			  'MM_rsKey.ActiveConnection=MM_rsKeyConnection
			 ' MM_rsKey.Source=MM_dupKeySQL
			 ' MM_rsKey.CursorType=0
			 ' MM_rsKey.CursorLocation=2
			 ' MM_rsKey.LockType=3
			 ' MM_rsKey.Open
			 ' If Not MM_rsKey.EOF Or Not MM_rsKey.BOF Then
				' the username was found - can not add the requested username
				'MM_qsChar = "?"
				'If (InStr(1,MM_dupKeyRedirect,"?") >= 1) Then MM_qsChar = "&"
				'MM_dupKeyRedirect = MM_dupKeyRedirect & MM_qsChar & "reqfirstname=" & MM_dupKeyUsernameValue1 & "&reqlastname=" & MM_dupKeyUsernameValue2 & "&reqmonth=" & MM_dupKeyUsernameValue3 & "&reqcity=" & MM_dupKeyUsernameValue4
				'Response.Redirect(MM_dupKeyRedirect)
			 ' End If
			'  MM_rsKey.Close
			'End If

%>


<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

	

  MM_editConnection = Connect
  MM_editTable = "q_user"
  MM_editColumn = "ID_user"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "q_user_edit.asp"
  MM_fieldsStr  = "first_name|value|last_name|value|login_name|value|password|value|expires|value|city|value|info1|value|info2|value|info3|value|info4|value|email|value|active|value|status|value|comment|value|IP|value|access|value|logcount|value|reference|value"
  MM_columnsStr = "user_firstname|',none,''|user_lastname|',none,''|user_username|',none,''|user_password|',none,''|user_expires|',none,NULL|user_city|',none,NULL|user_info1|none,none,NULL|user_info2|none,none,NULL|user_info3|none,none,NULL|user_info4|none,none,NULL|user_email|',none,''|user_active|none,1,0|user_status|none,1,0|user_comment|',none,''|user_IP|',none,''|user_access|',none,NULL|user_logcount|none,none,NULL|user_reference|none,none,NULL"

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
%>
<%
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
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
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(i) & " = " & FormVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId


  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    if Edit_OK = true then MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
	
	
	'update password
	Set obj = Server.CreateObject("ADODB.Recordset")
SQL="SELECT  * FROM  q_user  where ID_user="&Request("MM_recordId")
obj.ActiveConnection = Connect
obj.Source = SQL 
obj.CursorType = 0
obj.CursorLocation = 3
obj.LockType = 3
obj.Open

If obj.EOF then
Response.write("The END")

Else

Do While Not obj.EOF
Dim salt
salt = obj("user_email")
password=obj("user_city")
password=password&salt
password=sha256(password)
If  Len(Request("city"))<20 Then
Set uobj = Server.CreateObject("ADODB.Command")
SQL="update q_user set user_city='"&password&"', user_updated='"&cDateSql(Now())&"' WHERE ID_User="&obj("ID_USER")
uobj.ActiveConnection = Connect
uobj.CommandText = SQL
uobj.Execute
uobj.ActiveConnection.Close
else
Set uobj = Server.CreateObject("ADODB.Command")
SQL="update q_user set user_updated='"&cDateSql(Now())&"' WHERE ID_User="&obj("ID_USER")
uobj.ActiveConnection = Connect
uobj.CommandText = SQL
uobj.Execute
uobj.ActiveConnection.Close 


end IF






obj.MoveNext
Loop
End If

obj.close  
	

	
	'end update

	'PN 050720 Save the user subjects that have been submitted
		Dim updated_ok
		updated_ok=false
		Dim Seps(1)
		Seps(0) = "|"
		'PN 050720 delete this users entries from the  subject user table

		Set MM_editCmd = Server.CreateObject("ADODB.Command")
		MM_editCmd.ActiveConnection = Connect
		MM_editCmd.CommandText = "delete from subject_user where ID_user="&cInt(request("MM_recordId"))&";"
		MM_editCmd.Execute
		MM_editCmd.ActiveConnection.Close


		For Each q in Request.Form()

			if (((InStr(q,"user_subject"))>0)=True) then

					Dim a
					a= Tokenize(q, Seps)

					'050720 do an insert to the subject_user table
					Set MM_editCmd = Server.CreateObject("ADODB.Command")
					MM_editCmd.ActiveConnection = Connect
					MM_editCmd.CommandText = "insert into subject_user (ID_subject, ID_user) values ("&a(2)&","&cInt(request("MM_recordId"))&");"
					MM_editCmd.Execute
					MM_editCmd.ActiveConnection.Close


			end if


		Next


    call log_the_page ("Quiz Execute - UPDATE User: " & MM_recordId)
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>

<%
Dim user__MMColParam
user__MMColParam = "1"
if (Request.QueryString("user") <> "") then user__MMColParam = Request.QueryString("user")
%>
<%
set user = Server.CreateObject("ADODB.Recordset")
user.ActiveConnection = Connect
user.Source = "SELECT * FROM q_user WHERE ID_user = " + Replace(user__MMColParam, "'", "''") + ""
user.CursorType = 0
user.CursorLocation = 3
user.LockType = 3
user.Open()
user_numRows = 0
%>
<%
set info1 = Server.CreateObject("ADODB.Recordset")
info1.ActiveConnection = Connect
info1.Source = "SELECT * FROM q_info1 where info1_active=1 order by info1"
info1.CursorType = 0
info1.CursorLocation = 3
info1.LockType = 3
info1.Open()
info1_numRows = 0

info2_prm = user.fields.item("user_info1").value

set info2 = Server.CreateObject("ADODB.Recordset")
info2.ActiveConnection = Connect
'When the page is refreshed the code below will fetch the data from info2 table according which business was chosen. The SQL line below is what fetches the data from info2 table.
'if request("info1")<> "" then
'	info2_prm = request("info1")
'else
'	info2_prm = 0
'end if
info2.Source = "SELECT * FROM q_info2 where info2_info1 =" & info2_prm &" and info2_active=1 order by info2"
info2.CursorType = 0
info2.CursorLocation = 3
info2.LockType = 3
info2.Open()
info2_numRows = 0

'set info1 = Server.CreateObject("ADODB.Recordset")
'info1.ActiveConnection = Connect
'info1.Source = "SELECT * FROM q_info1 where info1_active=1 order by info1"
'info1.CursorType = 0
'info1.CursorLocation = 3
'info1.LockType = 3
'info1.Open()
'info1_numRows = 0
%>
<%
set info3 = Server.CreateObject("ADODB.Recordset")
info3.ActiveConnection = Connect
info3.Source = "SELECT * FROM q_info3"
info3.CursorType = 0
info3.CursorLocation = 3
info3.LockType = 3
info3.Open()
info3_numRows = 0

set info4 = Server.CreateObject("ADODB.Recordset")
info4.ActiveConnection = Connect
info4.Source = "SELECT * FROM q_info4 order by info4"
info4.CursorType = 0
info4.CursorLocation = 3
info4.LockType = 3
info4.Open()
info4_numRows = 0

set group = Server.CreateObject("ADODB.Recordset")
group.ActiveConnection = Connect
group.Source = "SELECT ID_usergroup, usergroup_name FROM q_user_group"
group.CursorType = 0
group.CursorLocation = 3
group.LockType = 3
group.Open()
group_numRows = 0
%>
<%
numbers=1

set admin_user = Server.CreateObject("ADODB.Recordset")
admin_user.ActiveConnection = Connect
admin_user.Source = "SELECT * FROM admin inner join q_info4 on admin.admin_info4=q_info4.id_info4 where admin.id_admin="&Session("MM_id_admin")&""
admin_user.CursorType = 0
admin_user.CursorLocation = 3
admin_user.LockType = 3
admin_user.Open()
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz user edit. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!-- This is where the checkform is run when the business dropdown is changed.
function checkform() {
	document.forms[0].action="q_user_edit.asp?user=<% =Request.QueryString("user") %>"
	document.forms[0].target="_self"
	document.forms[0].submit()
}
//-->
</script>
<script language="JavaScript">
<!--
function emailCheck (emailStr) {
var emailPat=/^(.+)@(.+)$/
var specialChars="\\(\\)<>@,;:\\\\\\\"\\.\\[\\]"
var validChars="\[^\\s" + specialChars + "\]"
var quotedUser="(\"[^\"]*\")"
var ipDomainPat=/^\[(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\]$/
var atom=validChars + '+'
var word="(" + atom + "|" + quotedUser + ")"
var userPat=new RegExp("^" + word + "(\\." + word + ")*$")
var domainPat=new RegExp("^" + atom + "(\\." + atom +")*$")
var matchArray=emailStr.match(emailPat)
if (matchArray==null) {
	alert("Email address seems incorrect (check @ and .'s)")
	return false
}
var user=matchArray[1]
var domain=matchArray[2]
if (user.match(userPat)==null) {
    alert("The username doesn't seem to be valid.")
    return false
}
var IPArray=domain.match(ipDomainPat)
if (IPArray!=null) {
	  for (var i=1;i<=4;i++) {
	    if (IPArray[i]>255) {
	        alert("Destination IP address is invalid!")
		return false
	    }
    }
    return true
}
var domainArray=domain.match(domainPat)
if (domainArray==null) {
	alert("The domain name doesn't seem to be valid.")
    return false
}
var atomPat=new RegExp(atom,"g")
var domArr=domain.match(atomPat)
var len=domArr.length
if (domArr[domArr.length-1].length<2 ||
    domArr[domArr.length-1].length>3) {
   alert("The address must end in a three-letter domain, or two letter country.")
   return false
}
if (len<2) {
   var errStr="This address is missing a hostname!"
   alert(errStr)
   return false
}
return true;
}

function isEmail(str)
{
  var supported = 0;
  if (window.RegExp) {
    var tempStr = "a";
    var tempReg = new RegExp(tempStr);
    if (tempReg.test(tempStr)) supported = 1;
  }
  if (!supported)
    return (str.indexOf(".") > 2) && (str.indexOf("@") > 0);
  var r1 = new RegExp("(@.*@)|(\\.\\.)|(@\\.)|(^\\.)");
  var r2 = new RegExp("^.+\\@(\\[?)[a-zA-Z0-9\\-\\.]+\\.([a-zA-Z]{2,3}|[0-9]{1,3})(\\]?)$");
  return (!r1.test(str) && r2.test(str));
}

function trySubmit()
{
	//document.forms[0].first_name.value = document.forms[0].first_name.value.toUpperCase();
	//document.forms[0].last_name.value = document.forms[0].last_name.value.toUpperCase();
	document.forms[0].city.value = document.forms[0].city.value.toUpperCase();
	if document.forms[0].city.value = ""
		document.forms[0].city.value=NULL;
	document.forms[0].login_name.value = replace(document.forms[0].login_name.value.toUpperCase(),' ','');
	//document.forms[0].password.value = replace(document.forms[0].password.value.toUpperCase(),' ','');
	//if (isEmail(document.forms[0].email.value))
	//{
	//document.forms[0].email.value = "";
	//return false;
	//}

	if (document.forms[0].first_name.value.length<2)
	{
		alert("Sorry, you must enter a first name!\n(min. 2 characters)");
		return false;
	}
	if (document.forms[0].last_name.value.length<2)
	{
		alert("Sorry, you must enter a last name!\n(min. 2 characters)");
		return false;
	}
	if (document.forms[0].login_name.value.length<2)
	{
		alert("Sorry, you must enter a LOGIN name!\n(min. 2 characters)");
		return false;
	}
	if (document.forms[0].info1.selectedIndex==0)
	{
		alert("Sorry, you must select a division");
		return false;
	}
	if (document.forms[0].info3.selectedIndex==0)
	{
		alert("Sorry, you must select a <% =location%>");
		return false;
	}

	return emailCheck (document.forms[0].email.value);

	if (confirm("Are you sure you want to update this user?"))	{	document.forms[0].submit();
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
<BODY BGCOLOR=#FFCC00 TEXT=#000000 VLINK=#000000 LINK=#000000 leftmargin="0" topmargin="0" onload="change=false; " onUnload="<% call on_page_unload %>">
<!-- WA_FilterAndPopulateSubList(info2_WAJA,MM_findObj('info1'),MM_findObj('info2'),0,0,false,': '); -->
<table width="100%" border="0" cellspacing="3" cellpadding="0">
  <tr>
    <td align="left" valign="bottom" class="heading"> Quiz user edit</td>
  </tr>
  <tr>
    <td align="left" valign="bottom">
      <!--<form ACTION="<%=MM_editAction%>" METHOD="POST" name="add_user" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">-->
	  <form ACTION="<%=MM_editAction%>" METHOD="POST" name="add_user" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table border="0" cellspacing="2" cellpadding="3" width="600">
          <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="100">First name:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <input type="text" name="first_name" onChange="change=true;" size="70" class="formitem1" value="<%=(user.Fields.Item("user_firstname").Value)%>">
            </td>
          </tr>
          <tr>
            <td class="text" align="left" valign="top" width="100">Last name:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <input type="text" name="last_name" onChange="change=true;" size="70" class="formitem1" value="<%=(user.Fields.Item("user_lastname").Value)%>">
            </td>
          </tr>
          <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="100">User name:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <input type="text" name="login_name" onChange="change=true;" size="70" class="formitem1" value="<%=(user.Fields.Item("user_username").Value)%>">
            </td>
          </tr>
         <tr>
            <td class="text" align="left" valign="top" width="100">Password:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <input type="text" name="city" onChange="change=true;" size="70" class="formitem1" value="<%=(user.Fields.Item("user_city").Value)%>">
            </td>
          </tr>
         <!--  <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="100">Expires (yyyy/mm/dd):</td>
            <td class="text" align="left" valign="top" colspan="3">
              <input type="text" name="expires" onChange="change=true;" size="70" class="formitem1" value="<%=cDateSQL(user.Fields.Item("user_expires").Value)%>">
            </td>
          </tr> -->

		 <!--<tr class="table_normal" >
            <td class="text" align="left" valign="top" width="100">City of birth:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <input type="text" name="city" onChange="change=true;" size="70" class="formitem1" value="<%=(user.Fields.Item("user_city").Value)%>">
            </td>
          </tr> -->

          <tr>
            <td class="text" align="left" valign="top" width="100">Business:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <select name="info1" class="formitem1" onChange="checkform();" >
			  <!-- onChange="change=true;WA_FilterAndPopulateSubList(info2_WAJA,MM_findObj('info1'),MM_findObj('info2'),0,0,false,': ')" -->
                <option value="0">---please select---</option>
                <% While (NOT info1.EOF) %>
				<option value="<%=(info1.Fields.Item("ID_info1").Value)%>"<%if (user.Fields.Item("user_info1").Value <> "") then
				if cint((info1.Fields.Item("ID_info1").Value) = cint(user.Fields.Item("user_info1").Value)) then Response.Write("SELECTED") : Response.Write("")
					%>> <%=(info1.Fields.Item("info1").Value)%>
					<%
						else
					%>> <%=(info1.Fields.Item("info1").Value)%>
					<%
						end if
					%></option>
                <%
  info1.MoveNext()
Wend
'If (info1.CursorType > 0) Then
'  info1.MoveFirst
'Else
  info1.Requery
'End If
%>
              </select>
			  <br><i>When changing the business, any other changes will be saved. Else click Update this user </i>
            </td>
           </tr>

		  <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="100">Site:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <select name="info2" onChange="change=true;" class="formitem1">
                <option value="0">---please select---</option>
				<% While (NOT info2.EOF) %>
				<option value="<%=(info2.Fields.Item("ID_info2").Value)%>"
				<%
				if (user.Fields.Item("user_info2").Value <> "") then
					if (cint((info2.Fields.Item("ID_info2").Value)) = cint((user.Fields.Item("user_info2").Value))) then Response.Write("SELECTED") : Response.Write("")
				%>
				><%=(info2.Fields.Item("info2").Value)%>
				<% else %>> <%=(info2.Fields.Item("info2").Value)%>
				<% end if %>
				</option>
                <%
				  info2.MoveNext()
				Wend
				'If (info2.CursorType > 0) Then
				'  info2.MoveFirst
				'Else
				  info2.Requery
				'End If
				%>
		  </select>
            </td>
          </tr>
      <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="100"><% =BBPinfo3 %>:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <select name="info3" onChange="change=true;" class="formitem1">
                <option value="0">---please select---</option>
				<% While (NOT info3.EOF) %> 
				<option value="<%=(info3.Fields.Item("ID_info3").Value)%>"  <%if (info3.Fields.Item("ID_info3").Value) = (user.Fields.Item("user_info3").Value) then Response.Write("SELECTED") : Response.Write("")%> > <%=(info3.Fields.Item("info3").Value)%> </option><%
				 info3.MoveNext()
				Wend
				'If (info3.CursorType > 0) Then
				' info3.MoveFirst
				'Else
				'  info3.Requery
				'End If
				%>
              </select>
            </td>
          </tr>

<!-- 		  <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="100">E-mail:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <input type="text" name="email" onChange="change=true;" size="70" class="formitem1" value="<%=(user.Fields.Item("user_email").Value)%>">
            </td>
          </tr>
 -->
          <tr class="table_normal">
            <td class="text" valign="top" width="143">Company:</td>
            <td class="text" valign="top" colspan="3">
              <select name="info4" class="formitem1">
                	<option value="0">--- select a company ---</option>
                <%
While (NOT info4.EOF)
	if admin_user.fields.item("info4_viewall").value=1 OR admin_user.fields.item("id_info4").value=info4.fields.item("id_info4").value then
%>
                <option value="<%=(info4.Fields.Item("ID_info4").Value)%>" <%if (info4.Fields.Item("ID_info4").Value) = (user.Fields.Item("user_info4").Value) then Response.Write("SELECTED") : Response.Write("")%>><%=(info4.Fields.Item("info4").Value)%></option>
                <%
	end if
	info4.MoveNext()
Wend
  info4.Requery

%>
              </select>
            </td>
          </tr>
		<tr>

           <td class="text" align="left" valign="top">E-mail:</td>
		   <td class="text" align="left" valign="top" colspan="3">
              <input type="text" name="email" onChange="change=true;" size="70" class="formitem1" value="<%=(user.Fields.Item("user_email").Value)%>">
            </td>
		   </tr>
		   <tr >

           <td class="text" align="left" valign="top">Employee Reference:</td>
		   <td class="text" align="left" valign="top" colspan="3">
              <input type="text" name="reference" onChange="change=true;" size="20" class="formitem1" value="<%=(user.Fields.Item("user_reference").Value)%>">
            </td>
		   </tr>
		  <tr class="table_normal">
            <td class="text" align="left" valign="top" width="100">Active account:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <input onChange="change=true;" <%If (abs(user.Fields.Item("user_active").Value) = 1) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="active" value="1">
            </td>
          </tr>
		  <% IF pref_offline THEN %>
		  <tr class="table_normal">
            <td class="text" align="left" valign="top" width="100">Status:</td>
            <td class="text" align="left" valign="top" colspan="3">
			<input type="radio" name="status" value="" onChange="change=true;" <%If (abs(user.Fields.Item("user_status").Value) = 0) Then Response.Write("CHECKED") : Response.Write("")%>> Online
			<input type="radio" name="status" value="1" onChange="change=true;"<%If (abs(user.Fields.Item("user_status").Value) = 1) Then Response.Write("CHECKED") : Response.Write("")%>> Offline
            </td>
          </tr>
		  <% END IF %>
		   <tr  >
            <td class="text" align="left" valign="center" width="100">Subjects:</td>
			<td class="text" align="left" valign="top" colspan="3">
		  	<table>


							<tr >

									<%
										'pn 050720 pull out all active subjects, currently only in reference to guide

										set subjects_b = Server.CreateObject("ADODB.Recordset")
										subjects_b.ActiveConnection = Connect
										subjects_b.Source = "SELECT subjects.ID_subject, subjects.subject_name  FROM (subjects INNER JOIN b_topics ON subjects.ID_subject = b_topics.topic_subject) INNER JOIN b_pages ON b_topics.ID_topic = b_pages.page_topic  GROUP BY subjects.ID_subject, subjects.subject_name, subjects.subject_ord, subjects.ID_subject, Abs([subject_active_b]), Abs([topic_active]), Abs([page_active])  HAVING (((Abs([subject_active_b]))=1) AND ((Abs([topic_active]))=1) AND ((Abs([page_active]))=1))  ORDER BY subjects.subject_ord, subjects.ID_subject;"
										subjects_b.CursorType = 0
										subjects_b.CursorLocation = 2
										subjects_b.LockType = 3
										subjects_b.Open()
										subjects_b_numRows = 0

										While (NOT subjects_b.EOF)

													show_this_subject=""

															set user_bbp_subject = Server.CreateObject("ADODB.Recordset")
															user_bbp_subject.ActiveConnection = Connect
															user_bbp_subject.Source = "SELECT *  FROM subject_user where ID_user="& user.Fields.Item("ID_user").Value &" and ID_subject="&subjects_b.Fields.Item("ID_subject").Value&";"
															user_bbp_subject.CursorType = 0
															user_bbp_subject.CursorLocation = 2
															user_bbp_subject.LockType = 3
															user_bbp_subject.Open()
																While ((NOT user_bbp_subject.EOF))
																	show_this_subject="checked"
																	user_bbp_subject.MoveNext()
																Wend

															user_bbp_subject.Close()
														%>
											<td  class="text" width="100">
												<%=subjects_b.Fields.Item("subject_name").Value%>
											</td>
											<td  class="text" width="50">

												<input type="checkbox" <%=show_this_subject%> name="user_subject|0|<%=subjects_b.Fields.Item("ID_subject").Value%>" />
											</td>

										<%subjects_b.MoveNext()
												Wend
												subjects_b.Close()

								%>
								</tr>

				</table>
			 </td>
          </tr>
          <!--tr class="table_normal" >
            <td class="text" align="left" valign="top" width="100">More info:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <textarea name="comment" onChange="change=true;" cols="53" rows="3" class="formitem1"><%=(user.Fields.Item("user_comment").Value)%></textarea>
            </td>
          </tr-->
          <tr>
            <td class="text" align="left" valign="top" width="100">Last IP:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <!--input type="text" name="IP" onChange="change=true;" size="70" class="formitem1" value="--><%=(user.Fields.Item("user_IP").Value)%><!--" disabled="true"-->
            </td>
          </tr>
          <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="100">Last login:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <!--input type="text" name="access" onChange="change=true;" size="70" class="formitem1" value="--><%=cDateSQL(user.Fields.Item("user_access").Value)%><!--" disabled="true"-->
            </td>
          </tr>
          <tr>
            <td class="text" align="left" valign="top" width="100">Login counter:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <!--input type="text" name="logcount" onChange="change=true;" size="70" class="formitem1" value="--><%=(user.Fields.Item("user_logcount").Value)%><!--"-->
            </td>
          </tr>
		  <tr>
            <td class="text" align="left" valign="top" width="100">User added on:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <!--input type="text" name="logcount" onChange="change=true;" size="70" class="formitem1" value="--><%=(user.Fields.Item("user_added").Value)%><!--"-->
            </td>
          </tr>
		  <tr>
            <td class="text" align="left" valign="top" width="100">User last updated on:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <!--input type="text" name="logcount" onChange="change=true;" size="70" class="formitem1" value="--><%=(user.Fields.Item("user_updated").Value)%><!--"-->
            </td>
          </tr>
          <tr>
            <td class="text" align="left" valign="top" width="100">&nbsp;</td>
            <td class="text" align="left" valign="top" colspan="3">&nbsp;</td>
          </tr>
		  <!--PN 050506 commented out as evaluators do not need to see this functionality-->
          <!--tr class="table_normal" >
            <td class="text" align="left" valign="top" width="100">Member of:</td>
            <td class="text" align="left" valign="top" width="30%">
              <select name="available" onChange="change=true;" size="10" multiple class="formitem1">
                <option value="0">--&gt; Available groups &lt;--</option>
                <%
While (NOT group.EOF)
%>
                <option value="<%=(group.Fields.Item("ID_usergroup").Value)%>" ><%=(group.Fields.Item("usergroup_name").Value)%></option>
                <%
  group.MoveNext()
Wend
'If (group.CursorType > 0) Then
'  group.MoveFirst
'Else
  group.Requery
'End If
%>
              </select>
            </td>
            <td class="text" align="center" valign="top" width="10%">
              <p>&nbsp; </p>
              <p>&nbsp;</p>
              <p>
                <input type="button" name="Submit2" value="&gt;&gt;&gt;" onClick="WA_AddSubToSelected(MM_findObj('available'),MM_findObj('current'),false,true,false,1,0,'0','')" class="quiz_button">
                <br>
                <input type="button" name="Submit22" value="&lt;&lt;&lt;" onClick="WA_RemoveSelectedFromList(MM_findObj('current'),'|WA|0|WA|',0,'0','')" class="quiz_button">
              </p>
            </td>
            <td class="text" align="right" valign="top" width="30%">
              <select name="current" onChange="change=true;" size="10" multiple class="formitem1">
                <option value="0">--&gt; Selected groups &lt;--</option>
              </select>
            </td>
          </tr-->
          <tr>
            <td class="text_table" align="left" valign="top" width="100">
              <input type="hidden" name="session" value="<%=getPassword(30, "", "true", "true", "true", "false", "true", "true", "true", "false")%>">
              <input type="hidden" name="current_export">
            </td>
            <td class="text_table" align="left" valign="top" colspan="3">
              <input type="reset" name="Submit3" value="Reset the form" class="quiz_button">
              <input type="submit" name="Submit" value="Update this user" class="quiz_button" <%call IsEditOK%>>
              or
              <input type="button" name="goback" value="Go back to user list" class="quiz_button" onClick="document.location='q_list_of_users.asp?<%=(request.querystring)%>'">
            </td>
          </tr>
        </table>
        <input type="hidden" name="MM_update" value="true">
        <input type="hidden" name="MM_recordId" value="<%= user.Fields.Item("ID_user").Value %>">
      </form>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
call log_the_page ("Quiz Edit User: " & (user.Fields.Item("ID_user").Value))
%>

<%
info1.Close()
info4.Close()
admin_user.Close()
%>
<%
'info2.Close()
%>
<%
'info3.Close()
%>
<%
group.Close()
%>
<%
user.Close()
%>
