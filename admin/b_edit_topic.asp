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
Dim topics__MMColParam
topics__MMColParam = "1"
if (Request.QueryString("topic") <> "") then topics__MMColParam = Request.QueryString("topic")
%>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = Connect
  MM_editTable = "b_topics"
  MM_editColumn = "ID_topic"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "b_list_of_topics.asp?subj=" + request.querystring("subj")
  MM_fieldsStr  = "topic_name|value|topic_title|value|topic_subject|value|topic_training|value|topic_keyp|value|topic_exmp|value|topic_help|value|topic_faq|value|active|value|topic_qanda|value"
  MM_columnsStr = "topic_name|',none,''|topic_title|',none,''|topic_subject|none,none,NULL|topic_training|none,none,NULL|topic_keyp|',none,''|topic_exmp|',none,''|topic_hlp|none,none,NULL|topic_faq|none,none,NULL|topic_active|none,1,0|topic_qanda|none,none,NULL"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(i+1) = CStr(Request.Form(MM_fields(i)))
  Next

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
    call log_the_page ("BBG Execute - UPDATE Topic: " & MM_recordId)	
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>

<%
set topics = Server.CreateObject("ADODB.Recordset")
topics.ActiveConnection = Connect
topics.Source = "SELECT *  FROM b_topics  WHERE ID_topic = " + Replace(topics__MMColParam, "'", "''") + ""
topics.CursorType = 0
topics.CursorLocation = 3
topics.LockType = 3
topics.Open()
topics_numRows = 0
%>
<%
set subjects = Server.CreateObject("ADODB.Recordset")
subjects.ActiveConnection = Connect
subjects.Source = "SELECT subjects.ID_subject, subjects.subject_name FROM subjects GROUP BY subjects.ID_subject, subjects.subject_name;"
subjects.CursorType = 0
subjects.CursorLocation = 3
subjects.LockType = 3
subjects.Open()
subjects_numRows = 0
%>
<%
set helps = Server.CreateObject("ADODB.Recordset")
helps.ActiveConnection = Connect
helps.Source = "SELECT * FROM b_hlp"
helps.CursorType = 0
helps.CursorLocation = 3
helps.LockType = 3
helps.Open()
helps_numRows = 0
%>
<%
set faqs = Server.CreateObject("ADODB.Recordset")
faqs.ActiveConnection = Connect
faqs.Source = "SELECT * FROM b_faq"
faqs.CursorType = 0
faqs.CursorLocation = 3
faqs.LockType = 3
faqs.Open()
faqs_numRows = 0
%>
<%
set qandas = Server.CreateObject("ADODB.Recordset")
qandas.ActiveConnection = Connect
qandas.Source = "SELECT q_topics.*, subjects.subject_name  FROM subjects INNER JOIN q_topics ON subjects.ID_subject = q_topics.topic_subject;"
qandas.CursorType = 0
qandas.CursorLocation = 3
qandas.LockType = 3
qandas.Open()
qandas_numRows = 0
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Topic edit. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
function trySubmit()
{
	if (document.forms[0].topic_name.value.length<3)
	{
		alert("Sorry, you must enter a name for the topic!\n(min. 3 characters)");
		return false;
	}
	if (document.forms[0].topic_keyp.value.length<3)
	{
		alert("Sorry, you must enter some key points (min. 3 characters) or click the 'No Key points' link.");
		return false;
	}
	if (document.forms[0].topic_exmp.value.length<3)
	{
		alert("Sorry, you must enter some examples (min. 3 characters) or click the 'No Examples' link.");
		return false;
	}	
	if (confirm("Are you sure you want to update this topic?"))	{	document.forms[0].submit();
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

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_changeProp(objName,x,theProp,theValue) { //v3.0
  var obj = MM_findObj(objName);
  if (obj && (theProp.indexOf("style.")==-1 || obj.style)) eval("obj."+theProp+"='"+theValue+"'");
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</HEAD>
<BODY onUnload="<% call on_page_unload %>" onLoad="change=false;">
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> BBG topic edit</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="add_subject" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr> 
            <td width="99%"  colspan="2">Name of the topic (as 
              on the menu)</td>
          </tr>
          <tr> 
            <td width="99%"  colspan="2"><a href="b_list_of_topics.asp?subj="></a> 
              <input type="text" onChange="change=true;" name="topic_name" size="60" class="formitem1" value="<%=(topics.Fields.Item("topic_name").Value)%>">
            </td>
          </tr>
          <tr> 
            <td width="99%"  colspan="2">Title of the page (as 
              at the top of the page)</td>
          </tr>
          <tr> 
            <td width="99%"  colspan="2"> 
              <input type="text" onChange="change=true;" name="topic_title" size="60" class="formitem1" value="<%=(topics.Fields.Item("topic_title").Value)%>">
            </td>
          </tr>
          <tr> 
            <td width="50%" >Name of the subject this topic 
              belongs to</td>
            <td width="50%" >Training link</td>
          </tr>
          <tr> 
            <td > 
              <select onChange="change=true;" name="topic_subject" class="formitem1">
                <%
While (NOT subjects.EOF)
%>
                <option value="<%=(subjects.Fields.Item("ID_subject").Value)%>" <%if ((subjects.Fields.Item("ID_subject").Value) = (topics.Fields.Item("topic_subject").Value)) then Response.Write("SELECTED") : Response.Write("")%> ><%=(subjects.Fields.Item("subject_name").Value)%></option>
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
            <td > 
              <select onChange="change=true;" name="topic_training" class="formitem1">
                <option value="0">---NO TRAINING---</option>
                <%
While (NOT subjects.EOF)
%>
                <option value="<%=(subjects.Fields.Item("ID_subject").Value)%>" <%if ((subjects.Fields.Item("ID_subject").Value) = (topics.Fields.Item("topic_training").Value)) then Response.Write("SELECTED") : Response.Write("")%>><%=(subjects.Fields.Item("subject_name").Value)%></option>
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
            <td  colspan="2">Key points tab (<a href="javascript:" onClick="MM_changeProp('topic_keyp','','value','<b>There are no key points on this topic.</b>','TEXTAREA')">No 
              Key points</a>) <a href="javascript:" onClick="MM_openBrWindow('_editor.asp?field=topic_keyp','editor','width=520,height=400')"><img src="images/editor.gif" width="24" height="10" border="0"></a> 
            </td>
          </tr>
          <tr> 
            <td width="99%"  colspan="2"> 
              <textarea name="topic_keyp" onChange="change=true;" cols="80" class="formitem1" rows="5"><%=(topics.Fields.Item("topic_keyp").Value)%></textarea>
            </td>
          </tr>
          <tr> 
            <td  colspan="2">Examples tab (<a href="javascript:" onClick="MM_changeProp('topic_exmp','','value','<b>There are no examples on this topic.</b>','TEXTAREA')">No 
              Examples</a>) <a href="javascript:" onClick="MM_openBrWindow('_editor.asp?field=topic_exmp','editor','width=520,height=400')"><img src="images/editor.gif" width="24" height="10" border="0"></a> 
            </td>
          </tr>
          <tr> 
            <td width="99%"  colspan="2"> 
              <textarea name="topic_exmp" onChange="change=true;" cols="80" class="formitem1" rows="10"><%=(topics.Fields.Item("topic_exmp").Value)%></textarea>
            </td>
          </tr>
          <tr> 
            <td >Help tab</td>
            <td >FAQ tab</td>
          </tr>
          <tr> 
            <td > 
              <select name="topic_help" onChange="change=true;" class="formitem1">
                <%
While (NOT helps.EOF)
%>
                <option value="<%=(helps.Fields.Item("ID_hlp").Value)%>" <%if ((helps.Fields.Item("ID_hlp").Value) = (topics.Fields.Item("topic_hlp").Value)) then Response.Write("SELECTED") : Response.Write("")%> ><%=(helps.Fields.Item("hlp_name").Value)%></option>
                <%
  helps.MoveNext()
Wend
'If (helps.CursorType > 0) Then
'  helps.MoveFirst
'Else
  helps.Requery
'End If
%>
              </select>
            </td>
            <td > 
              <select name="topic_faq" onChange="change=true;" class="formitem1">
                <%
While (NOT faqs.EOF)
%>
                <option value="<%=(faqs.Fields.Item("ID_faq").Value)%>" <%if ((faqs.Fields.Item("ID_faq").Value) = (topics.Fields.Item("topic_faq").Value)) then Response.Write("SELECTED") : Response.Write("")%> ><%=(faqs.Fields.Item("faq_name").Value)%></option>
                <%
  faqs.MoveNext()
Wend
'If (faqs.CursorType > 0) Then
'  faqs.MoveFirst
'Else
  faqs.Requery
'End If
%>
              </select>
            </td>
          </tr>
          <tr> 
            <td >Q&amp;A link</td>
            <td >Topic active? 
              <input onChange="change=true;" <%If (abs(topics.Fields.Item("topic_active").Value) = 1) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="active" value="checkbox">
            </td>
          </tr>
          <tr> 
            <td colspan="2" > 
              <select onChange="change=true;" name="topic_qanda" class="formitem1">
                <option value="0">---NO Q and A---</option>
                <%
While (NOT qandas.EOF)
%>
                <option value="<%=(qandas.Fields.Item("ID_topic").Value)%>" <%if ((qandas.Fields.Item("ID_topic").Value) = (topics.Fields.Item("topic_qanda").Value)) then Response.Write("SELECTED") : Response.Write("")%>><%="[" & (qandas.Fields.Item("subject_name").Value) & "] - " & (qandas.Fields.Item("topic_name").Value)%></option>
                <%
  qandas.MoveNext()
Wend
'If (qandas.CursorType > 0) Then
'  qandas.MoveFirst
'Else
  qandas.Requery
'End If
%>
              </select>
            </td>
          </tr>
          <tr> 
            <td width="99%"  colspan="2">&nbsp; </td>
          </tr>
          <tr> 
            <td width="99%"  colspan="2"> 
              <input type="reset" name="Submit2" value="Reset the form" class="quiz_button">
              <input type="submit" name="Submit" value="Update this topic" class="quiz_button" <%call IsEditOK%>>
              or 
              <input type="button" name="goback" value="Go back to topic list" class="quiz_button" onClick="document.location='b_list_of_topics.asp?subj=<%=(topics.Fields.Item("topic_subject").Value)%>'">
            </td>
          </tr>
        </table>
        <input type="hidden" name="MM_update" value="true">
        <input type="hidden" name="MM_recordId" value="<%= topics.Fields.Item("ID_topic").Value %>">
      </form>
    </td>
  </tr>
</table>
<p> 
<p>&nbsp; </p>
</BODY>
</HTML>

<%
call log_the_page ("BBG Edit Topic: " & (topics.Fields.Item("ID_topic").Value))
%>

<%
topics.Close()
%>
<%
subjects.Close()
%>
<%
helps.Close()
%>
<%
faqs.Close()
%>
<%
qandas.Close()
%>

