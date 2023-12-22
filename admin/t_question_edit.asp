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
  MM_editTable = "tr_pages"
  MM_editColumn = "ID_page"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_fieldsStr  = "subject|value|topic|value|title|value|monkey|value|q_body|value|previous|value|next|value|scenario|value|foto|value|more|value|active|value"
  MM_columnsStr = "page_subject|none,none,NULL|page_topic|none,none,NULL|page_title|',none,''|page_monkey|none,none,NULL|page_text|',none,''|page_previous|none,none,NULL|page_next|none,none,NULL|page_scenario|none,none,NULL|page_photo|',none,''|page_note|',none,''|page_active|none,1,0|,''"

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
' *** Update Record: construct a sql update staatement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update staatement
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
    call log_the_page ("Training Execute - UPDATE Page: " & MM_recordId)
  End If

End If
%>

<%
' *** Update feedbacks

If (CStr(Request("MM_update")) <> "") Then

for iii = 1 to session("feedbacks_total")

  MM_editConnection = Connect
  MM_editTable = "tr_feedback"
  MM_editColumn = "ID_feedback"
  MM_recordId = "" + Request.Form("recordId2_" & iii) + "" 
  MM_fieldsStr  = "fb_header" & iii & "|value|fb_content" & iii & "|value"
  MM_columnsStr = "feedback_head|',none,''|feedback_text|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")

  ' set the form values
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(i+1) = CStr(Request.Form(MM_fields(i)))
  Next

  ' create the sql update staatement
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

  End If
next

End If
%>


<%
' *** Update questions

If (CStr(Request("MM_update")) <> "") Then

for iii = 1 to session("questions_total")

  MM_editConnection = Connect
  MM_editTable = "tr_questions"
  MM_editColumn = "ID_question"
  MM_recordId = "" + Request.Form("recordId1_" & iii) + "" 
  MM_editRedirectUrl = "t_question_edit.asp?qid=" & Request.Form("MM_recordId")
  MM_fieldsStr  = "answer" & iii & "|value|answer_ok" & iii & "|value"
  MM_columnsStr = "question_text|',none,''|question_ok|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")

  ' set the form values
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(i+1) = CStr(Request.Form(MM_fields(i)))
  Next

  ' create the sql update staatement
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

  End If
next
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
End If
%>

<%
Dim qid
If (Request.QueryString("qid") <> "") Then 
qid = cInt(Request.QueryString("qid"))
Else 
Response.Redirect("error.asp?" & request.QueryString) 
End If
%>
<%
set subjects = Server.CreateObject("ADODB.Recordset")
subjects.ActiveConnection = Connect
subjects.Source = "SELECT ID_subject, subject_name FROM subjects"
subjects.CursorType = 0
subjects.CursorLocation = 3
subjects.LockType = 3
subjects.Open()
subjects_numRows = 0
%>
<%
set topics = Server.CreateObject("ADODB.Recordset")
topics.ActiveConnection = Connect
topics.Source = "SELECT *  FROM tr_topics"
topics.CursorType = 0
topics.CursorLocation = 3
topics.LockType = 3
topics.Open()
topics_numRows = 0
%>
<%
set page_content = Server.CreateObject("ADODB.Recordset")
page_content.ActiveConnection = Connect
page_content.Source = "SELECT tr_pages.*, tr_topics.topic_name, subjects.subject_name, subjects.subject_monkey, subjects.subject_banner, tr_monkeys.monkey_file  FROM tr_topics RIGHT JOIN (subjects RIGHT JOIN (tr_monkeys RIGHT JOIN tr_pages ON tr_monkeys.ID_monkey = tr_pages.page_monkey) ON subjects.ID_subject = tr_pages.page_subject) ON tr_topics.ID_topic = tr_pages.page_topic  WHERE ID_page = " + Replace(qid, "'", "''") + ";"
page_content.CursorType = 0
page_content.CursorLocation = 3
page_content.LockType = 3
page_content.Open()
page_content_numRows = 0
%>
<%
set questions = Server.CreateObject("ADODB.Recordset")
questions.ActiveConnection = Connect
questions.Source = "SELECT *  FROM tr_questions  WHERE tr_questions.question_ID_page = " + Replace(qid, "'", "''") + ";"
questions.CursorType = 0
questions.CursorLocation = 3
questions.LockType = 3
questions.Open()
questions_numRows = 0
%>
<%
set feedbacks = Server.CreateObject("ADODB.Recordset")
feedbacks.ActiveConnection = Connect
feedbacks.Source = "SELECT tr_feedback.*  FROM tr_feedback  WHERE tr_feedback.feedback_ID_page = " + Replace(qid, "'", "''") + ";"
feedbacks.CursorType = 0
feedbacks.CursorLocation = 3
feedbacks.LockType = 3
feedbacks.Open()
feedbacks_numRows = 0
%>
<%
set monkeys = Server.CreateObject("ADODB.Recordset")
monkeys.ActiveConnection = Connect
monkeys.Source = "SELECT * FROM tr_monkeys"
monkeys.CursorType = 0
monkeys.CursorLocation = 3
monkeys.LockType = 3
monkeys.Open()
monkeys_numRows = 0
%>
<%
set allpages = Server.CreateObject("ADODB.Recordset")
allpages.ActiveConnection = Connect
allpages.Source = "SELECT tr_pages.ID_page, tr_pages.page_title, subjects.subject_name, tr_topics.topic_name FROM (subjects INNER JOIN tr_topics ON subjects.ID_subject = tr_topics.topic_subject) INNER JOIN tr_pages ON tr_topics.ID_topic = tr_pages.page_topic ORDER BY subjects.subject_ord, subjects.ID_subject, tr_topics.topic_ord, tr_topics.ID_topic, tr_pages.ID_page;"
allpages.CursorType = 0
allpages.CursorLocation = 3
allpages.LockType = 3
allpages.Open()
allpages_numRows = 0
%>
<%
set count_pages = Server.CreateObject("ADODB.Recordset")
count_pages.ActiveConnection = Connect
count_pages.Source = "SELECT Max(tr_pages.ID_page) AS MaxOfID_page  FROM tr_pages;"
count_pages.CursorType = 0
count_pages.CursorLocation = 3
count_pages.LockType = 3
count_pages.Open()
count_pages_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
questions_numRows = questions_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Repeat2__numRows = -1
Dim Repeat2__index
Repeat2__index = 0
feedbacks_numRows = feedbacks_numRows + Repeat2__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
feedbacks_total = feedbacks.RecordCount

' set the number of rows displayed on this page
If (feedbacks_numRows < 0) Then
  feedbacks_numRows = feedbacks_total
Elseif (feedbacks_numRows = 0) Then
  feedbacks_numRows = 1
End If

' set the first and last displayed record
feedbacks_first = 1
feedbacks_last  = feedbacks_first + feedbacks_numRows - 1

' if we have the correct record count, check the other stats
If (feedbacks_total <> -1) Then
  If (feedbacks_first > feedbacks_total) Then feedbacks_first = feedbacks_total
  If (feedbacks_last > feedbacks_total) Then feedbacks_last = feedbacks_total
  If (feedbacks_numRows > feedbacks_total) Then feedbacks_numRows = feedbacks_total
End If
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
questions_total = questions.RecordCount

' set the number of rows displayed on this page
If (questions_numRows < 0) Then
  questions_numRows = questions_total
Elseif (questions_numRows = 0) Then
  questions_numRows = 1
End If

' set the first and last displayed record
questions_first = 1
questions_last  = questions_first + questions_numRows - 1

' if we have the correct record count, check the other stats
If (questions_total <> -1) Then
  If (questions_first > questions_total) Then questions_first = questions_total
  If (questions_last > questions_total) Then questions_last = questions_total
  If (questions_numRows > questions_total) Then questions_numRows = questions_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (feedbacks_total = -1) Then

  ' count the total records by iterating through the recordset
  feedbacks_total=0
  While (Not feedbacks.EOF)
    feedbacks_total = feedbacks_total + 1
    feedbacks.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (feedbacks.CursorType > 0) Then
'    feedbacks.MoveFirst
'  Else
    feedbacks.Requery
  End If

  ' set the number of rows displayed on this page
  If (feedbacks_numRows < 0 Or feedbacks_numRows > feedbacks_total) Then
    feedbacks_numRows = feedbacks_total
  End If

  ' set the first and last displayed record
  feedbacks_first = 1
  feedbacks_last = feedbacks_first + feedbacks_numRows - 1
  If (feedbacks_first > feedbacks_total) Then feedbacks_first = feedbacks_total
  If (feedbacks_last > feedbacks_total) Then feedbacks_last = feedbacks_total

End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (questions_total = -1) Then

  ' count the total records by iterating through the recordset
  questions_total=0
  While (Not questions.EOF)
    questions_total = questions_total + 1
    questions.MoveNext
  Wend

  ' reset the cursor to the beginning
'  If (questions.CursorType > 0) Then
'    questions.MoveFirst
'  Else
    questions.Requery
'  End If

  ' set the number of rows displayed on this page
  If (questions_numRows < 0 Or questions_numRows > questions_total) Then
    questions_numRows = questions_total
  End If

  ' set the first and last displayed record
  questions_first = 1
  questions_last = questions_first + questions_numRows - 1
  If (questions_first > questions_total) Then questions_first = questions_total
  If (questions_last > questions_total) Then questions_last = questions_total

End If
%>
<% 
Session("questions_total") = questions_total
%>
<% 
Session("feedbacks_total") = feedbacks_total
%>
<%
function WA_VBreplace(thetext)
  if isNull(thetext) then thetext = ""
  newstring = Replace(cStr(thetext),"'","|WA|")
  newstring = Replace(newstring,"\","\\")
  WA_VBreplace = newstring
end function

if (NOT topics.EOF)     THEN

  Response.Write("<SC" & "RIPT>"&chr(10))
  Response.Write("var WAJA = new Array();"&chr(10))

  oldmainid = 0
  newmainid = topics.Fields("topic_subject").value
  if (oldmainid = newmainid)    THEN
    oldmainid = ""
  END IF
  n = 0
    while (NOT topics.EOF)
    if (oldmainid <> newmainid)     THEN
      Response.Write("WAJA[" & n & "] = new Array();"&chr(10))
      Response.Write("WAJA[" & n & "][0] = '" & WA_VBreplace(newmainid) & "';"&chr(10))
      m = 1
    END IF

    Response.Write("WAJA[" & n & "][" & m & "] = new Array();"&chr(10))
    Response.Write("WAJA[" & n & "][" & m & "][0] = " & "'" & WA_VBreplace(topics.Fields("ID_topic").value) & "'" & ";" &chr(10))
    Response.Write("WAJA[" & n & "][" & m & "][1] = " & "'" & WA_VBreplace(topics.Fields("topic_name").value) & "'" & ";" &chr(10))
    m=m+1
    if (cStr(oldmainid) = "0")      THEN
      oldmainid = newmainid
    END IF
    oldmainid = newmainid
    topics.MoveNext()
    if (NOT topics.EOF)     THEN
      newmainid = topics.Fields("topic_subject").value
    END IF
    if (oldmainid <> newmainid)     THEN
      n=n+1
    END IF
  WEND

  Response.Write("var topics_WAJA = WAJA;"&chr(10))
  Response.Write("WAJA = null;"&chr(10))
  Response.Write("</SC" & "RIPT>"&chr(10))
END IF
if (NOT topics.BOF)     THEN
  topics.MoveFirst()
END IF
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Training page edit. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function trySubmit(howmuchquestions)
	{
		if (document.forms[0].title.value.length<3)
		{
			alert("Sorry, you must enter a page title!\n(min. 3 characters)");
			return false;
		}
		if (document.forms[0].q_body.value.length<3)
		{
			alert("Sorry, you must fill in a page body!\n(min. 3 characters)");
			return false;
		}
		for (i = 1; i <= howmuchquestions; i++) 
		{
			answerindex = (MM_findObj('answer_ok'+i));
			if (answerindex.selectedIndex==0)
			{
				alert("Sorry, you must choose some feedback in the answer number: " + i );
				return false;
			}
			answerindex = (MM_findObj('answer'+i));
			if (answerindex.value.length<2)
			{
				alert("Sorry, you must fill in an answer number: " + i +"!\n(min. 2 characters)");
				return false;
			}

		}
	if (confirm("Are you sure you want to update this page?"))	{	document.forms[0].submit();
	return false;
	}
return false;
	}

function addnewanswer(whichanswer)
{
	if (change==true)
	{
		alert("You have changed at least one field on this page.\rBefore adding a new answer, save those changes first.");
		return false;
	}
	document.location=whichanswer;
}
function addnewfeedback(whichfeedback)
{
	if (change==true)
	{
		alert("You have changed at least one field on this page.\rBefore adding a new feedback, save those changes first.");
		return false;
	}
	document.location=whichfeedback;
}
function exitpage()
{
	if (change==true)
	{
		if (confirm("You have changed at least one field on this page.\rBefore exiting this page, do you want to save those changes first?"))
		{
		return trySubmit(<%=questions_total%>);
		}
	}
	return true;
}

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

function MM_changeProp(objName,x,theProp,theValue) { //v3.0
  var obj = MM_findObj(objName);
  if (obj && (theProp.indexOf("style.")==-1 || obj.style)) eval("obj."+theProp+"='"+theValue+"'");
}
function selecttopic(whichmenu,whichtopic){
 for (var i=0; i<whichmenu.options.length; i++)     {
 if (whichmenu.options[i].value == whichtopic)	{
 whichmenu.options[i].selected = true;
	 }
 }
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</HEAD>
<BODY onload="change=false; WA_FilterAndPopulateSubList(topics_WAJA,MM_findObj('subject'),MM_findObj('topic'),0,0,false,': '); selecttopic(MM_findObj('topic'),<%=(page_content.Fields.Item("page_topic").Value)%>);" onUnload="<% call on_page_unload %>">
<table>
  <tr> 
    <td align="left" valign="bottom" class="headers"> Training page edit</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom" > 
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="editq" onSubmit="<%call on_form_Submit(questions_total)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr align="left" valign="top"> 
            <td > 
              <table>
                <tr> 
                  <td  width="90">Subject</td>
                  <td > 
                    <select name="subject" onChange="WA_FilterAndPopulateSubList(topics_WAJA,MM_findObj('subject'),MM_findObj('topic'),0,0,false,': '); change=true;" class="formitem1">
                      <%
While (NOT subjects.EOF)
%>
                      <option value="<%=(subjects.Fields.Item("ID_subject").Value)%>" <%if ((subjects.Fields.Item("ID_subject").Value) = (page_content.Fields.Item("page_subject").Value)) then Response.Write("SELECTED") : Response.Write("")%>><%=(subjects.Fields.Item("subject_name").Value)%></option>
                      <%
  subjects.MoveNext()
Wend
'If (subjects.CursorType > 0) Then
'  subjects.MoveFirst
'Else
  subjects.Requery
'End If
%>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td  width="90">Topic</td>
                  <td > 
                    <select name="topic" class="formitem1" onChange="change=true;">
                      <%
While (NOT topics.EOF)
%>
                      <option value="<%=(topics.Fields.Item("ID_topic").Value)%>" ><%=(topics.Fields.Item("topic_name").Value)%></option>
                      <%
  topics.MoveNext()
Wend
'If (topics.CursorType > 0) Then
'  topics.MoveFirst
'Else
  topics.Requery
'End If
%>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td  width="90">Monkey</td>
                  <td > 
                    <select name="monkey" class="formitem1" onChange="change=true;" >
                      <%
While (NOT monkeys.EOF)
%>
                      <option value="<%=(monkeys.Fields.Item("ID_monkey").Value)%>" <%if ((monkeys.Fields.Item("ID_monkey").Value) = (page_content.Fields.Item("page_monkey").Value)) then Response.Write("SELECTED") : Response.Write("")%>><%=(monkeys.Fields.Item("monkey_name").Value) & ": " & (monkeys.Fields.Item("monkey_file").Value)%></option>
                      <%
  monkeys.MoveNext()
Wend
'If (monkeys.CursorType > 0) Then
'  monkeys.MoveFirst
'Else
  monkeys.Requery
'End If
%>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td  width="90">Title </td>
                  <td > 
                    <input type="text" size="49" class="formitem1" onChange="change=true;" name="title" value="<%=(page_content.Fields.Item("page_title").Value)%>">
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr align="left" valign="top"> 
            <td >Body <a href="javascript:" onClick="MM_openBrWindow('_editor.asp?field=q_body','editor','width=520,height=400')"><img src="images/editor.gif" width="24" height="10" border="0"></a> 
            </td>
          </tr>
          <tr align="left" valign="top"> 
            <td > 
              <textarea name="q_body" rows="12" class="formitem1" onChange="change=true;" cols="100"><%=(page_content.Fields.Item("page_text").Value)%></textarea>
            </td>
          </tr>
          <tr align="left" valign="top"> 
            <td > 
              <table>
                <tr> 
                  <td  width="90">Previous page</td>
                  <td > 
                    <select name="previous" class="formitem1" onChange="change=true;" >
                      <option value="-1">&lt;&lt;&lt; back to the menu</option>
                      <%
While (NOT allpages.EOF)
%>
                      <option value="<%=(allpages.Fields.Item("ID_page").Value)%>" <%if ((allpages.Fields.Item("ID_page").Value) = (page_content.Fields.Item("page_previous").Value)) then Response.Write("SELECTED") : Response.Write("")%>><%=(allpages.Fields.Item("subject_name").Value) & "-" & (allpages.Fields.Item("topic_name").Value) & ": " & (allpages.Fields.Item("page_title").Value)%></option>
                      <%
  allpages.MoveNext()
Wend
'If (allpages.CursorType > 0) Then
'  allpages.MoveFirst
'Else
  allpages.Requery
'End If
%>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td  width="90">Next page</td>
                  <td > 
                    <select name="next" class="formitem1" onChange="change=true;" >
                      <option value="-1">&gt;&gt;&gt; back to the menu</option>
                      <%
While (NOT allpages.EOF)
%>
                      <option value="<%
					  lastid=(allpages.Fields.Item("ID_page").Value)
					  response.write lastid
					  %>" <%if ((allpages.Fields.Item("ID_page").Value) = (page_content.Fields.Item("page_next").Value)) then Response.Write("SELECTED") : Response.Write("")%>><%=(allpages.Fields.Item("subject_name").Value) & "-" & (allpages.Fields.Item("topic_name").Value) & ": " & (allpages.Fields.Item("page_title").Value)%></option>
                      <%
  allpages.MoveNext()
Wend
'If (allpages.CursorType > 0) Then
'  allpages.MoveFirst
'Else
  allpages.Requery
'End If
%>
                      <option value="<%=(count_pages.Fields.Item("MaxOfID_page").Value) + 1%>" <%if (CStr((count_pages.Fields.Item("MaxOfID_page").Value) + 1) = CStr(page_content.Fields.Item("page_next").Value)) then Response.Write("SELECTED") : Response.Write("")%>>??? 
                      probably next ID - must be checked when new screen added!!!</option>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td  width="90">Scenario page</td>
                  <td > 
                    <select name="scenario" class="formitem1" onChange="change=true;" >
                      <option value="0">... no scenario</option>
                      <%
While (NOT allpages.EOF)
%>
                      <option value="<%=(allpages.Fields.Item("ID_page").Value)%>" <%if ((allpages.Fields.Item("ID_page").Value) = (page_content.Fields.Item("page_scenario").Value)) then Response.Write("SELECTED") : Response.Write("")%>><%=(allpages.Fields.Item("subject_name").Value) & "-" & (allpages.Fields.Item("topic_name").Value) & ": " & (allpages.Fields.Item("page_title").Value)%></option>
                      <%
  allpages.MoveNext()
Wend
'If (allpages.CursorType > 0) Then
'  allpages.MoveFirst
'Else
  allpages.Requery
'End If
%>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td  width="90">Photo</td>
                  <td > 
                    <input type="text" size="70" class="formitem1" onChange="change=true;" name="foto" value="<%=lCase(page_content.Fields.Item("page_photo").Value)%>">
                    <a href="javascript:"  onClick="MM_openBrWindow('_photo_browse.asp','photobrowse','scrollbars=yes,width=610,height=400')"><img src="images/search.gif" width="16" height="16" border="0"></a> 
                    <%
if (page_content.Fields.Item("page_photo").Value) <> "" Then
	if fileexist("../client/training_photos/"& lCase(page_content.Fields.Item("page_photo").Value)) = True Then
	Response.Write " <img src='../admin/images/ok.gif'> "
	Else
	Response.Write " <img src='../admin/images/miss.gif'> "
	End If 
Else
	Response.Write " <img src='../admin/images/no.gif'> "
End If
%>
                    <% if (page_content.Fields.Item("page_photo").Value) <> "" Then %>
                    <a href="javascript:"><img src="images/eye.gif" width="16" height="16" border="0" onClick="MM_openBrWindow('../client/training_photos/<%=lCase(page_content.Fields.Item("page_photo").Value)%>','preview','resizable=yes,width=200,height=200,top=250,left=250')"></a> 
                    <%End If %>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr align="left" valign="top"> 
            <td  bgcolor="#6699FF"> 
              <table>
                <tr valign="top" bgcolor="#66CCFF"> 
                  <td  align="left" colspan="4">Answers</td>
                </tr>
                <% 
ii=1
%>
                <% 
While ((Repeat1__numRows <> 0) AND (NOT questions.EOF)) 
%>
                <tr valign="top"> 
                  <td  align="left"> <a href="javascript:" onClick="MM_openBrWindow('_editor.asp?field=answer<%=ii%>','editor','width=520,height=400')"><img src="images/editor.gif" width="24" height="10" border="0"></a> 
                  </td>
                  <td  align="left"> 
                    <textarea cols="65" class="formitem2" onChange="change=true;" name="answer<%=ii%>" rows="5"><%=(questions.Fields.Item("question_text").Value)%></textarea>
                  </td>
                  <td  align="left"> 
                    <p>
                      <%if Edit_OK then %>
                      <a href="t_answer_del.asp?qid=<%=qid%>&aid=<%= (questions.Fields.Item("ID_question").Value)%>&comeback=<%=CStr(Request("URL"))%>"><img src="images/bin.gif" width="16" height="16" border="0" onClick="javascript:return (confirm('You are just about to delete this answer.\nAre you sure you want to do that?'));"></a>
                      <% end if %>
                    </p>
                  </td>
                  <td  align="right"> 
                    <select name="answer_ok<%=ii%>" class="formitem2" onChange="change=true;">
                      <option value="0">...no feedback</option>
                      <%
While (NOT feedbacks.EOF)
%>
                      <option value="<%=(feedbacks.Fields.Item("ID_feedback").Value)%>" <%if ((feedbacks.Fields.Item("ID_feedback").Value) = (questions.Fields.Item("question_ok").Value)) then Response.Write("SELECTED") : Response.Write("")%>><%=(feedbacks.Fields.Item("ID_feedback").Value) & ": " &(feedbacks.Fields.Item("feedback_head").Value)%></option>
                      <%
  feedbacks.MoveNext()
Wend
'If (feedbacks.CursorType > 0) Then
'  feedbacks.MoveFirst
'Else
  feedbacks.Requery
'End If
%>
                    </select>
                    <input type="hidden" name="recordId1_<%=ii%>" value="<%=(questions.Fields.Item("ID_question").Value)%>">
                  </td>
                </tr>
                <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  questions.MoveNext()
  ii = ii +1
Wend
%>
                <tr valign="top"> 
                  <td  colspan="4" align="left"> 
                    <input type="button" name="new_choice" value="Add a new answer" class="quiz_button" onClick="return addnewanswer('t_answer_add.asp?qid=<%=qid%>&comeback=<%=CStr(Request("URL"))%>');" <%call IsEditOK%>>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr align="left" valign="top" bgcolor="#33CC33"> 
            <td  bgcolor="#33CC33"> 
              <table>
                <tr valign="top" bgcolor="#66FF33"> 
                  <td  align="left" colspan="3">Feedbacks</td>
                </tr>
                <% 
ii=1
%>
                <% 
While ((Repeat2__numRows <> 0) AND (NOT feedbacks.EOF)) 
%>
                <tr valign="top"> 
                  <td  align="left"><%=(feedbacks.Fields.Item("ID_feedback").Value)%> 
                    <input type="text" size="26" class="formitem2" onChange="change=true;" name="fb_header<%=ii%>" value="<%=(feedbacks.Fields.Item("feedback_head").Value)%>">
                    <input type="hidden" name="recordId2_<%=ii%>" value="<%=(feedbacks.Fields.Item("ID_feedback").Value)%>">
                  </td>
                  <td  align="left">
                    <%if Edit_OK then %>
                    <a href="t_feedback_del.asp?qid=<%=qid%>&fid=<%= (feedbacks.Fields.Item("ID_feedback").Value) %>&comeback=<%=CStr(Request("URL"))%>"><img src="images/bin.gif" width="16" height="16" border="0" onClick="javascript:return (confirm('You are just about to delete this feedback.\nAre you sure you want to do that?'));">
                    <% end if %>
                    </a> </td>
                  <td  align="right" rowspan="3"> 
                    <textarea name="fb_content<%=ii%>" cols="63" rows="5" class="formitem2" onChange="change=true;"><%=(feedbacks.Fields.Item("feedback_text").Value)%></textarea>
                  </td>
                </tr>
                <tr valign="top"> 
                  <td  align="left" colspan="2"><a href="javascript:" onClick="MM_openBrWindow('_editor.asp?field=fb_content<%=ii%>','editor','width=520,height=400')"><img src="images/editor.gif" width="24" height="10" border="0"></a> 
                  </td>
                </tr>
                <tr valign="top"> 
                  <td  align="left" colspan="2"> 
                    <p><a href="javascript:"onClick="MM_changeProp('fb_header<%=ii%>','','value','Correct!','INPUT/TEXT')">Correct!</a>&nbsp;<a href="javascript:"onClick="MM_changeProp('fb_header<%=ii%>','','value','Incorrect','INPUT/TEXT')">Incorrect</a>&nbsp;<a href="javascript:"onClick="MM_changeProp('fb_header<%=ii%>','','value','Yes','INPUT/TEXT')">Yes</a>&nbsp;<a href="javascript:"onClick="MM_changeProp('fb_header<%=ii%>','','value','No','INPUT/TEXT')">No</a>&nbsp;<a href="javascript:"onClick="MM_changeProp('fb_header<%=ii%>','','value','Right','INPUT/TEXT')">Right</a> 
                    </p>
                  </td>
                </tr>
                <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  feedbacks.MoveNext()
  ii = ii + 1
Wend
%>
                <tr valign="top"> 
                  <td  colspan="3" align="left"> 
                    <input type="button" name="new_choice2" value="Add a new feedback" class="quiz_button" onClick="return addnewfeedback('t_feedback_add.asp?qid=<%=qid%>&comeback=<%=CStr(Request("URL"))%>');" <%call IsEditOK%>>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr align="left" valign="top"> 
            <td >Bottom note <a href="javascript:" onClick="MM_openBrWindow('_editor.asp?field=more','editor','width=520,height=400')"><img src="images/editor.gif" width="24" height="10" border="0"></a> 
            </td>
          </tr>
          <tr align="left" valign="top"> 
            <td > 
              <textarea name="more" cols="100" class="formitem1" onChange="change=true;" rows="2"><%=(page_content.Fields.Item("page_note").Value)%></textarea>
            </td>
          </tr>
		<tr> 
            <td width="99%" >Question active? 
              <input onChange="change=true;" <%If (abs(page_content.Fields.Item("page_active").Value) = 1) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="active" value="checkbox">
            </td>
         </tr>

		 </table>
        <p> 
          <input type="reset" name="Reset" value="Reset the form " class="quiz_button">
          <input type="submit" name="Submit" value="Update this page" class="quiz_button" <%call IsEditOK%>>
          <input type="button" name="New_question" value="Add a new page" onClick="document.location='t_question_add.asp';" class="quiz_button">
          or 
          <input type="button" name="goback" value="Go back to question list" class="quiz_button" onClick="document.location='t_list_of_questions.asp?subj=<%=(page_content.Fields.Item("page_subject").Value)%>&topic=<%=(page_content.Fields.Item("page_topic").Value)%>'">
        </p>
        <input type="hidden" name="MM_update" value="true">
        <input type="hidden" name="MM_recordId" value="<%= page_content.Fields.Item("ID_page").Value %>">
      </form>
    </td>
  </tr>
</table>
<p></p>
<p>&nbsp;</p>
</body>
</HTML>

<%
call log_the_page ("Training Edit Page: " & (page_content.Fields.Item("ID_page").Value))
%>

<%
subjects.Close()
%>
<%
topics.Close()
%>
<%
page_content.Close()
%>
<%
questions.Close()
%>
<%
feedbacks.Close()
%>
<%
monkeys.Close()
%>
<%
allpages.Close()
%>
<%
count_pages.Close()
%>
