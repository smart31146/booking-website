<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<!--#include file="sha256.asp"-->

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


' Create relationships between users and usergroups
sub create_membership
	MM_editConnection = Connect
	membership_all = request.form("current_export")
	membership_array = Split(membership_all, ",")
	membership_count = UBound(membership_array)

	set new_user = Server.CreateObject("ADODB.Recordset")
	new_user.ActiveConnection = Connect
' Code has been replaced to make user_username unique string to get user ID_user from database by PR on 24.02.2016
'	new_user.Source = "SELECT * FROM q_user WHERE user_new_session = '" + request("session") + "';"
	new_user.Source = "SELECT * FROM q_user WHERE user_username = '" + request("login_name") + "';"
	new_user.CursorType = 0
	new_user.CursorLocation = 3
	new_user.LockType = 3
	new_user.Open()
	new_user_numRows = 0
	new_user_id = (new_user.Fields.Item("ID_User").Value)
	new_user.Close()

if membership_count > -1 then
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
	for iii = LBound(membership_array) to membership_count
 	MM_editQuery = "insert into membership (id_user, id_usergroup) values (" & cInt(new_user_id) & "," & cInt(membership_array(iii)) & ");"
    MM_editCmd.CommandText = MM_editQuery
    if Edit_OK = true then MM_editCmd.Execute
	next
    MM_editCmd.ActiveConnection.Close

end if
end sub
'pn 050726 add means to link a user with subjects
sub create_user_subjects
	MM_editConnection = Connect


	set new_user = Server.CreateObject("ADODB.Recordset")
	new_user.ActiveConnection = Connect
'   Code has been replaced to make user_username unique string to get user ID_user from database by PR on 24.02.2016
'	new_user.Source = "SELECT * FROM q_user WHERE user_new_session = '" + request("session") + "';"
	new_user.Source = "SELECT * FROM q_user WHERE user_username = '" + request("login_name") + "';"
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
	Dim Seps(1)
	Seps(0) = "|"
	'PN 050720 delete from the  subject user table

	if request("send_user_email")=1 then
		' insert email
		sql_date=cDateSql(now_bbp())

		MM_editConnection = Connect
		MM_editQuery = "insert into emails (q_user,date_to_send,type) values ('"&new_user_id&"','"&sql_date&"' , 1)"
		'response.write ( MM_editQuery)
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
	
	For Each q in Request.Form()

		if (((InStr(q,"user_subject"))>0)=True) then

				Dim a
				a= Tokenize(q, Seps)

				'050720 do an insert to the subject_user table
				Set MM_editCmd = Server.CreateObject("ADODB.Command")
				MM_editCmd.ActiveConnection = Connect
				MM_editCmd.CommandText = "insert into subject_user (ID_subject, ID_user) values ("&a(2)&","&cInt(new_user_id) &");"
				MM_editCmd.Execute
				MM_editCmd.ActiveConnection.Close


				'Response.Write "<li>Keyword " &a(1)&"   "&a(2)&"    "&a(3)& "   "&Request.Form(q)&"</li>"

				'Response.Write "starts with " & q & "<br>"

				if request("send_user_email")=1 then
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

					check_subject_quiz.Close()

				end if
		end if


	Next

end sub
%>
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
' *** Redirect if username exists
MM_flag="MM_insert"
If (CStr(Request(MM_flag)) <> "") Then
  MM_dupKeyRedirect="q_user_duplicate.asp"
  MM_rsKeyConnection=Connect
  MM_dupKeyUsernameValue = Replace(CStr(Request.Form("login_name")), "'", "''")
  MM_dupKeySQL="SELECT * FROM q_user WHERE user_username='" & MM_dupKeyUsernameValue & "'"
  MM_adodbRecordset="ADODB.Recordset"
  set MM_rsKey=Server.CreateObject(MM_adodbRecordset)
  MM_rsKey.ActiveConnection=MM_rsKeyConnection
  MM_rsKey.Source=MM_dupKeySQL
  MM_rsKey.CursorType=0
  MM_rsKey.CursorLocation=2
  MM_rsKey.LockType=3
  MM_rsKey.Open
  If Not MM_rsKey.EOF Or Not MM_rsKey.BOF Then
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
					'MM_flag="MM_insert"
					'If (CStr(Request(MM_flag)) <> "") Then
					'  MM_dupKeyRedirect="q_user_duplicate2.asp"
					'  MM_rsKeyConnection=Connect
					'  MM_dupKeyUsernameValue1 = CStr(Request.Form("first_name"))
					 ' MM_dupKeyUsernameValue2 = CStr(Request.Form("last_name"))

					 ' MM_dupKeySQL="SELECT * FROM q_user WHERE user_firstname='" & MM_dupKeyUsernameValue1 & "' AND user_lastname='" & MM_dupKeyUsernameValue2 & "';"
					 ' MM_adodbRecordset="ADODB.Recordset"
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
						'MM_dupKeyRedirect = MM_dupKeyRedirect & MM_qsChar & "reqfirstname=" & MM_dupKeyUsernameValue1 & "&reqlastname=" & MM_dupKeyUsernameValue2
						'Response.Redirect(MM_dupKeyRedirect)
					 ' End If
					 ' MM_rsKey.Close
					'End If
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then

  MM_editConnection = Connect
  MM_editTable = "q_user"
  MM_editRedirectUrl = "q_list_of_users.asp"
  
  MM_fieldsStr  = "first_name|value|last_name|value|login_name|value|info1|value|info2|value|info3|value|info4|value|active|value|session|value|email|value|reference|value"
  MM_columnsStr = "user_firstname|',none,''|user_lastname|',none,''|user_username|',none,''|user_info1|none,none,NULL|user_info2|none,none,NULL|user_info3|none,none,NULL|user_info4|none,none,NULL|user_active|none,1,0|user_new_session|',none,''|user_email|',none,''|user_reference|',none,''"
  
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
' *** Insert Record: construct a sql insert staatement and execute it

If (CStr(Request("MM_insert")) <> "") Then

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

  If (Not MM_abortEdit) Then
    ' execute the insert
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    if Edit_OK = true then MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
    call log_the_page ("Quiz Execute - INSERT User")
	
	'update password
	Set obj = Server.CreateObject("ADODB.Recordset")
SQL="SELECT TOP 1 * FROM  q_user  ORDER BY ID_user DESC"
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
if IsObject(password) then 
password=password&salt
password=sha256(password)
Set uobj = Server.CreateObject("ADODB.Command")
SQL="update q_user set user_city='"&password&"', user_added='"&cDateSql(Now())&"' WHERE ID_User="&obj("ID_USER")
uobj.ActiveConnection = Connect
uobj.CommandText = SQL
uobj.Execute
uobj.ActiveConnection.Close
End If


obj.MoveNext
Loop
End If

obj.close  
	
	'end update

create_membership
'pn 050726 added means to connect a new user with a subject
create_user_subjects

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If


'pn 050805 add in qinfo 1 2 and 3

set info1 = Server.CreateObject("ADODB.Recordset")
info1.ActiveConnection = Connect
info1.Source = "SELECT * FROM q_info1 WHERE info1_active = 1"
info1.CursorType = 0
info1.CursorLocation = 3
info1.LockType = 3
info1.Open()
info1_numRows = 0
%>
<%
set info2 = Server.CreateObject("ADODB.Recordset")
info2.ActiveConnection = Connect
info2.Source = "SELECT * FROM q_info2"
info2.CursorType = 0
info2.CursorLocation = 3
info2.LockType = 3
info2.Open()
info2_numRows = 0
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
%>
<%
set info4 = Server.CreateObject("ADODB.Recordset")
info4.ActiveConnection = Connect
info4.Source = "SELECT * FROM q_info4 order by info4"
info4.CursorType = 0
info4.CursorLocation = 3
info4.LockType = 3
info4.Open()
info4_numRows = 0

set admin_user = Server.CreateObject("ADODB.Recordset")
admin_user.ActiveConnection = Connect
admin_user.Source = "SELECT * FROM admin inner join q_info4 on admin.admin_info4=q_info4.id_info4 where admin.id_admin="&Session("MM_id_admin")&""
admin_user.CursorType = 0
admin_user.CursorLocation = 3
admin_user.LockType = 3
admin_user.Open()
%>
<%
function WA_VBreplace(thetext)
  if isNull(thetext) then thetext = ""
  newstring = Replace(cStr(thetext),"'","|WA|")
  newstring = Replace(newstring,"\","\\")
  WA_VBreplace = newstring
end function

if (NOT info2.EOF)     THEN

  Response.Write("<SC" & "RIPT>"&chr(10))
  Response.Write("var WAJA = new Array();"&chr(10))

  oldmainid = 0
  newmainid = info2.Fields("info2_info1").value
  if (oldmainid = newmainid)    THEN
    oldmainid = ""
  END IF
  n = 0
    while (NOT info2.EOF)
    if (oldmainid <> newmainid)     THEN
      Response.Write("WAJA[" & n & "] = new Array();"&chr(10))
      Response.Write("WAJA[" & n & "][0] = '" & WA_VBreplace(newmainid) & "';"&chr(10))
      m = 1
    END IF

    Response.Write("WAJA[" & n & "][" & m & "] = new Array();"&chr(10))
    Response.Write("WAJA[" & n & "][" & m & "][0] = " & "'" & WA_VBreplace(info2.Fields("ID_info2").value) & "'" & ";" &chr(10))
    Response.Write("WAJA[" & n & "][" & m & "][1] = " & "'" & WA_VBreplace(info2.Fields("info2").value) & "'" & ";" &chr(10))
    m=m+1
    if (cStr(oldmainid) = "0")      THEN
      oldmainid = newmainid
    END IF
    oldmainid = newmainid
    info2.MoveNext()
    if (NOT info2.EOF)     THEN
      newmainid = info2.Fields("info2_info1").value
    END IF
    if (oldmainid <> newmainid)     THEN
      n=n+1
    END IF
  WEND

  Response.Write("var info2_WAJA = WAJA;"&chr(10))
  Response.Write("WAJA = null;"&chr(10))
  Response.Write("</SC" & "RIPT>"&chr(10))
END IF
if (NOT info2.BOF)     THEN
  info2.MoveFirst()
END IF
%>






<%
function WA_VBreplace(thetext)
  if isNull(thetext) then thetext = ""
  newstring = Replace(cStr(thetext),"'","|WA|")
  newstring = Replace(newstring,"\","\\")
  WA_VBreplace = newstring
end function


%>
<%
numbers=1
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz new user. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function WA_ClientSideReplace(theval,findvar,repvar)     {
  var retval = "";
  while (theval.indexOf(findvar) >= 0)    {
    retval += theval.substring(0,theval.indexOf(findvar));
    retval += repvar;
    theval = theval.substring(theval.indexOf(findvar) + String(findvar).length);
  }
  if (retval == "" && theval.indexOf(findvar) < 0)    {
    retval = theval;
  }
  return retval;
}

function WA_UnloadList(thelist,leavevals,bottomnum)    {
  while (thelist.options.length > leavevals+bottomnum)     {
    if (thelist.options[leavevals])     {
      thelist.options[leavevals] = null;
    }
  }
  return leavevals;
}

function WA_FilterAndPopulateSubList(thearray,sourceselect,targetselect,leaveval,bottomleave,usesource,delimiter)     {
  if (bottomleave > 0)     {
    leaveArray = new Array(bottomleave);
    if (targetselect.options.length >= bottomleave)     {
      for (var m=0; m<bottomleave; m++)     {
        leavetext = targetselect.options[(targetselect.options.length - bottomleave + m)].text;
        leavevalue  = targetselect.options[(targetselect.options.length - bottomleave + m)].value;
        leaveArray[m] = new Array(leavevalue,leavetext);
      }
    }
    else     {
      for (var m=0; m<bottomleave; m++)     {
        leavetext = "";
        leavevalue  = "";
        leaveArray[m] = new Array(leavevalue,leavetext);
      }
    }
  }
  startid = WA_UnloadList(targetselect,leaveval,0);
  mainids = new Array();
  if (usesource)    maintext = new Array();
  for (var j=0; j<sourceselect.options.length; j++)     {
    if (sourceselect.options[j].selected)     {
      mainids[mainids.length] = sourceselect.options[j].value;
      if (usesource)     maintext[maintext.length] = sourceselect.options[j].text + delimiter;
    }
  }
  for (var i=0; i<thearray.length; i++)     {
    goodid = false;
    for (var h=0; h<mainids.length; h++)     {
      if (thearray[i][0] == mainids[h])     {
        goodid = true;
        break;
      }
    }
    if (goodid)     {
      theBox = targetselect;
      theLength = parseInt(theBox.options.length);
      theServices = thearray[i].length + startid;
      var l=1;
      for (var k=startid; k<theServices; k++)     {
        if (l == thearray[i].length)     break;
        theBox.options[k] = new Option();
        theBox.options[k].value = thearray[i][l][0];
        if (usesource)     theBox.options[k].text = maintext[h] + WA_ClientSideReplace(thearray[i][l][1],"|WA|","'");
        else               theBox.options[k].text = WA_ClientSideReplace(thearray[i][l][1],"|WA|","'");
        l++;
      }
      startid = k;
    }
  }
  if (bottomleave > 0)     {
    for (var n=0; n<leaveArray.length; n++)     {
      targetselect.options[startid+n] = new Option();
      targetselect.options[startid+n].value = leaveArray[n][0];
      targetselect.options[startid+n].text  = leaveArray[n][1];
    }
  }
  for (var l=0; l < targetselect.options.length; l++)    {
    targetselect.options[l].selected = false;
  }
  if (targetselect.options.length > 0)     {
    targetselect.options[0].selected = true;
  }
}

function WA_AddValueToList(ListObj,TextString,ValString,Position)  {
  if (isNaN(parseInt(Position)))   {
    Position = ListObj.options.length;
  }
  else  {
    Position = parseInt(Position);
  }
  if (ListObj.length > Position)  {
  ListObj.options[Position].text=TextString;
  if (ValString != "")  {
    ListObj.options[Position].value = ValString;
  }
    else  {
      ListObj.options[Position].value=TextString;
    }
  }
  else  {
    var LastOption = new Option();
    var OptionPosition = ListObj.options.length;
    ListObj.options[OptionPosition] = LastOption;
    ListObj.options[OptionPosition].text = TextString;
    if (ValString != "")  {
      ListObj.options[OptionPosition].value = ValString;
    }
    else  {
      ListObj.options[OptionPosition].value=TextString;
    }
  }
}

function WA_subAwithBinC(a,b,c)
{

	var i = c.indexOf(a);
	var l = b.length;

	while (i != -1)	{
		c = c.substring(0,i) + b + c.substring(i + a.length,c.length);  //replace all valid a values with b values in the selected string c.
  i += l
		i = c.indexOf(a,i);
	}
	return c;

}

function WA_RemoveSelectedFromList(theBox,nottoremove,noneselectedoption,noneselectedvalue,noneselectedtext)     {
  var n=0;
  var selectedArray = new Array();
  for (var j=0; j<theBox.options.length; j++)     {
    if (!theBox.options[j].selected || nottoremove.indexOf("|WA|" + theBox.options[j].value + "|WA|") >= 0)     {
      theBox.options[n].value = theBox.options[j].value;
      theBox.options[n].text = theBox.options[j].text;
      n++;
    }
    else {
	    selectedArray[selectedArray.length] = j;
    }
  }
  for (var k=0; k<selectedArray.length; k++)  {
    theBox.options[selectedArray[k]].selected = false;
  }
  m = n;
  while (m<=j)     {
    theBox.options[n] = null;
    m++;
  }
  if (theBox.options.length == noneselectedoption && noneselectedtext != "")     {
    noneselectedvalue = WA_subAwithBinC("|WA|",",",noneselectedvalue);
    noneselectedtext = WA_subAwithBinC("|WA|",",",noneselectedtext);
    WA_AddValueToList(theBox,noneselectedtext,noneselectedvalue,0);
  }
  for (var l=0; l < theBox.options.length; l++)    {
    theBox.options[l].selected = false;
  }
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function WA_AddSubToSelected(sublist,targetlist,repeatvalues,leavetop,leavebottom,noseltop,noselbot,topval,toptext)     {
  for (var j=0; j<noseltop; j++)     {
    sublist.options[j].selected = false;
  }
  for (var k=0; k<noselbot; k++)     {
    sublist.options[sublist.options.length-(k+1)].selected = false;
  }
  if (sublist.selectedIndex >= 0)      {
   if (leavebottom)     {
      botrec = new Array(2);
      botrec[0] = targetlist.options[targetlist.options.length-1].value;
      botrec[1] = targetlist.options[targetlist.options.length-1].text;
      targetlist.options[targetlist.options.length-1] = null;
    }
    if (!leavetop && targetlist.options.length > 0)     {
      if (targetlist.options[0].value == topval)     {
        targetlist.options[0] = null;
      }
    }
    else     {
      if (leavetop && toptext != '')     {
        targetlist.options[0].value = topval;
        targetlist.options[0].text = toptext;
      }
    }
    for (var o=0; o<sublist.options.length; o++)     {
      if (sublist.options[o].selected && o >= noseltop && o < sublist.options.length - noselbot)     {
        theText = sublist.options[o].text;
        theValue = sublist.options[o].value;
        addvalue = true;
        if (!repeatvalues)      {
          for (var p=0; p<targetlist.options.length; p++)     {
            if (theValue == targetlist.options[p].value)      {
              addvalue = false;
            }
          }
        }
        if (addvalue)  WA_AddValueToList(targetlist,theText,theValue,targetlist.options.length);
      }
    }
    if (leavebottom)     {
      WA_AddValueToList(targetlist,botrec[1],botrec[0],targetlist.options.length);
    }
  }
  for (var l=0; l < targetlist.options.length; l++)    {
    targetlist.options[l].selected = false;
  }
}
function SaveMe() {
var strValues = "";
var boxLength = document.forms[0].current.length;
var count = 0;
if (boxLength > 0) {
for (i = 1; i < boxLength; i++) {
if (count == 0) {
strValues = document.forms[0].current.options[i].value;
}
else {
strValues = strValues + "," + document.forms[0].current.options[i].value;
}
count++;
   }
}
if (strValues.length == 0) {
document.forms[0].current_export.value = "";
}
else {
document.forms[0].current_export.value = strValues;
   }
}

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

function check_date(field){
var checkstr = "0123456789";
var DateField = field;
var Datevalue = "";
var DateTemp = "";
var seperator = ".";
var day;
var month;
var year;
var leap = 0;
var err = 0;
var i;
   err = 0;
   DateValue = DateField.value;
   for (i = 0; i < DateValue.length; i++) {
	  if (checkstr.indexOf(DateValue.substr(i,1)) >= 0) {
	     DateTemp = DateTemp + DateValue.substr(i,1);
	  }
   }
   DateValue = DateTemp;
   if (DateValue.length == 6) {
      DateValue = DateValue.substr(0,4) + '20' + DateValue.substr(4,2); }
   if (DateValue.length != 8) {
      err = 19;}
   year = DateValue.substr(4,4);
   if (year == 0) {
      err = 20;
   }
   month = DateValue.substr(2,2);
   if ((month < 1) || (month > 12)) {
      err = 21;
   }
   day = DateValue.substr(0,2);
   if (day < 1) {
     err = 22;
   }
   if ((year % 4 == 0) || (year % 100 == 0) || (year % 400 == 0)) {
      leap = 1;
   }
   if ((month == 2) && (leap == 1) && (day > 29)) {
      err = 23;
   }
   if ((month == 2) && (leap != 1) && (day > 28)) {
      err = 24;
   }
   if ((day > 31) && ((month == "01") || (month == "03") || (month == "05") || (month == "07") || (month == "08") || (month == "10") || (month == "12"))) {
      err = 25;
   }
   if ((day > 30) && ((month == "04") || (month == "06") || (month == "09") || (month == "11"))) {
      err = 26;
   }
   if ((day == 0) && (month == 0) && (year == 00)) {
      err = 0; day = ""; month = ""; year = ""; seperator = "";
   }
   if (err == 0) {
      DateField.value = day + seperator + month + seperator + year;
   }
   else {
      alert("Date is incorrect!");
   }
}

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

	document.forms[0].first_name.value = document.forms[0].first_name.value.toUpperCase();
	document.forms[0].last_name.value = document.forms[0].last_name.value.toUpperCase();

	document.forms[0].login_name.value = replace(document.forms[0].login_name.value.toUpperCase(),' ','');

	if (document.forms[0].first_name.value.length<2)

	{
				//All alert messages has been updated by PR on 23.03.2016 HDK #2084
				alert("To add an user you must enter a first name!\n(min. 2 characters)");
		return false;
	}
	if (document.forms[0].last_name.value.length<2)
	{
		alert("To add an user you must enter a last name!\n(min. 2 characters)");
		return false;
	}
	if (document.forms[0].login_name.value.length<2)
	{
		alert("To add an user you must enter a login name!\n(min. 2 characters)");
		return false;
	}
	if (document.forms[0].info1.selectedIndex==0)
	{
		alert("To add an user you must select a business");
		return false;
	}
	if (document.forms[0].info3.selectedIndex==0)
	{
		alert("To add an user you must select a <% =BBPinfo3 %>");
		return false;
	}
	if (document.forms[0].info4.selectedIndex==0)
	{
		alert("To add an user you must select a Company");
		return false;
	}
	if (document.forms[0].reference.selectedIndex<2)
	{
		alert("To add an user you must enter a reference!\n(min. 2 characters)");
		return false;
	}

	return emailCheck (document.forms[0].email.value);

	if (confirm("Are you sure you want to add a new user?"))	{	document.forms[0].submit();
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
<BODY onLoad="change=false;" onUnload="<% call on_page_unload %>">
<table>
  <tr>
    <td align="left" valign="bottom" class="heading"> Quiz add new user</td>
  </tr>
  <tr>
    <td align="left" valign="bottom">
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="add_user" onSubmit=" <%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="120">First name:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <input type="text" name="first_name" onChange="change=true;" size="70" class="formitem1">
            </td>
          </tr>
          <tr>
            <td class="text" align="left" valign="top">Last name:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <input type="text" name="last_name" onChange="change=true;" size="70" class="formitem1">
            </td>
          </tr>
          <tr class="table_normal" >
		  <!--Cameron replace this text please. form of username, instructions to user-->
            <td class="text" align="left" valign="top">User Name:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <input type="text" name="login_name" onChange="change=true;" size="70" class="formitem1">
            </td>
          </tr>
         <tr >
				<td class="text" align="left" valign="center">Subjects:</td>
				<td class="text" align="left" valign="top" colspan="3">
				<table>
					<tr>

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
								%>
								<tr>
									<td  class="text" width="190">
										<%=subjects_b.Fields.Item("subject_name").Value%>
									</td>
									<td  class="text" width="50" colspan=2>

										<input type="checkbox" checked  name="user_subject|0|<%=subjects_b.Fields.Item("ID_subject").Value%>" />
									</td>
								</tr>
								<%subjects_b.MoveNext()
									Wend
									subjects_b.Close()
								%>

					</tr>
				</table>
				</tr>
				 <tr>
            <td class="text" align="left" valign="top">Business:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <select name="info1" onChange="change=true; WA_FilterAndPopulateSubList(info2_WAJA,MM_findObj('info1'),MM_findObj('info2'),0,0,false,': ')" class="formitem1">
                <option value="0">---please select---</option>
                <%
While (NOT info1.EOF)
%>
                <option value="<%=(info1.Fields.Item("ID_info1").Value)%>" ><%=(info1.Fields.Item("info1").Value)%></option>
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
            </td>
          </tr>
        <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="100">Site:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <select name="info2" onChange="change=true;" class="formitem1">
                <option value="0">---please select---</option>
                <%
					While (NOT info2.EOF)
					%>
						<option value="<%=(info2.Fields.Item("ID_info2").Value)%>"><%=(info2.Fields.Item("info2").Value)%></option>
					<%
					  info2.MoveNext()
					Wend
				%>
              </select>
            </td>
          </tr>
          <tr class="table_normal">
            <td class="text" align="left" valign="top"><% =BBPinfo3 %>:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <select name="info3" class="formitem1">
                <option value="0">---please select---</option>
                <%
While (NOT info3.EOF)
%>
                <option value="<%=(info3.Fields.Item("ID_info3").Value)%>" ><%=(info3.Fields.Item("info3").Value)%></option>
                <%
  info3.MoveNext()
Wend
'If (info3.CursorType > 0) Then
'  info3.MoveFirst
'Else
  info3.Requery
'End If
%>
              </select>
            </td>
          </tr>
          <tr class="table_normal">
            <td class="text" valign="top" width="143">Company:</td>
            <td class="text" valign="top" colspan="3">
              <select name="info4" class="formitem1">
                	<option value="0">--- select a company ---</option>
                <%
While (NOT info4.EOF)
	if admin_user.fields.item("info4_viewall").value=1 OR admin_user.fields.item("id_info4").value=info4.fields.item("id_info4").value then
%>
                <option value="<%=(info4.Fields.Item("ID_info4").Value)%>" <%if (CStr(info4.Fields.Item("ID_info4").Value) = CStr(request.querystring("info4"))) then Response.Write("SELECTED") : Response.Write("")%>><%=(info4.Fields.Item("info4").Value)%></option>
                <%
	end if
	info4.MoveNext()
Wend
  info4.Requery

%>
              </select>
            </td>
          </tr>
          <tr >

           <td class="text" align="left" valign="top">E-mail:</td>
		   <td class="text" align="left" valign="top" colspan="3">
              <input type="text" name="email" onChange="change=true;" size="70" class="formitem1" value="">
            </td>
		   </tr>
  		   <tr  class="table_normal" >
		     <td class="text" align="left" valign="top">Start induction:</td>
		     <td class="text" align="left" valign="top" colspan="3">
			     <input onChange="change=true;" type="checkbox" name="send_user_email" value="1" CHECKED>
		     </td>
		   </tr>
			<tr >

           <td class="text" align="left" valign="top">Employee Reference:</td>
		   <td class="text" align="left" valign="top" colspan="3">
              <input type="text" name="reference" onChange="change=true;" size="20" class="formitem1" value="">
            </td>
		   </tr>
          <tr>
            <td class="text" align="left" valign="top"></td>
            <td class="text" align="left" valign="top" colspan="3">
              <input type="hidden" name="active"  value="1" >
            </td>
          </tr>

          <tr>
            <td  align="left" valign="top">
              <input type="hidden" name="session" value="<%=getPassword(30, "", "true", "true", "true", "false", "true", "true", "true", "false")%>">
              <input type="hidden" name="current_export">
			  <input type="hidden" name="password"  value="cement" >
            </td>
            <td  align="left" valign="top" colspan="3">
              <input type="reset" name="Submit3" value="Reset the form" class="quiz_button">
              <input type="submit" name="Submit" value="Insert this new user" class="quiz_button" <%call IsEditOK%>>
              or
              <input type="button" name="goback" value="Go back to user list" class="quiz_button" onClick="document.location='q_list_of_users.asp'">
            </td>
          </tr>
        </table>
        <input type="hidden" name="MM_insert" value="true">
      </form>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
call log_the_page ("Quiz Add a new User")
%>


