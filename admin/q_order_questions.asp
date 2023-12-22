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
Dim top
If (Request.QueryString("topic") <> "") Then 
top = cInt(Request.QueryString("topic"))
Else 
Response.Redirect("error.asp?" & request.QueryString) 
End If
%>
<%
Dim subj
If (Request.QueryString("subj") <> "") Then 
subj = cInt(Request.QueryString("subj"))
Else 
Response.Redirect("error.asp?" & request.QueryString) 
End If
%>

<%
' *** Update choices

If (CStr(Request("MM_update")) <> "") Then

for iii = 1 to cInt(Request.form("field_length"))

  MM_editConnection = Connect
  MM_editTable = "q_question"
  MM_editColumn = "ID_question"
  MM_recordId = "" + Request.Form("q_id" & iii) + "" 
  MM_editRedirectUrl = "q_list_of_topics.asp?subj=" & subj & "&topic=" & top
  MM_fieldsStr  = "q_ord" & iii & "|value"
  MM_columnsStr = "question_ord|',none,''"

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
    call log_the_page ("Quiz Execute - UPDATE Questions: " & MM_recordId)
  End If
next
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
End If
%>
<%
set questions = Server.CreateObject("ADODB.Recordset")
questions.ActiveConnection = Connect
questions.Source = "SELECT q_question.ID_question, q_question.question_body, q_question.question_ord  FROM subjects INNER JOIN (q_topics INNER JOIN q_question ON q_topics.ID_topic = q_question.question_topic) ON subjects.ID_subject = q_topics.topic_subject  WHERE ((q_question.question_topic =" + Replace(top, "'", "''") + ") AND (abs(q_question.question_active)=1))  ORDER BY q_question.question_ord, q_question.ID_question;"
questions.CursorType = 0
questions.CursorLocation = 3
questions.LockType = 3
questions.Open()
questions_numRows = 0
%>
<%
set subject = Server.CreateObject("ADODB.Recordset")
subject.ActiveConnection = Connect
subject.Source = "SELECT subject_name  FROM subjects  WHERE ID_subject = " + Replace(subj, "'", "''") + ""
subject.CursorType = 0
subject.CursorLocation = 3
subject.LockType = 3
subject.Open()
subject_numRows = 0
%>
<%
set topic = Server.CreateObject("ADODB.Recordset")
topic.ActiveConnection = Connect
topic.Source = "SELECT topic_name  FROM q_topics  WHERE ID_topic = " + Replace(top, "'", "''") + ""
topic.CursorType = 0
topic.CursorLocation = 3
topic.LockType = 3
topic.Open()
topic_numRows = 0
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz questions order. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
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

function WA_MoveSelectedInList(sourceselect,tomove,topnum,botnum)     {
  var selectedIndex = sourceselect.selectedIndex;
  if (selectedIndex > topnum && tomove == "0" && selectedIndex <= sourceselect.options.length-botnum-1)    {
    oldvals = new Array(sourceselect.options[selectedIndex - 1].value, sourceselect.options[selectedIndex - 1].text);
    sourceselect.options[selectedIndex-1].value = sourceselect.options[selectedIndex].value;
    sourceselect.options[selectedIndex-1].text  = sourceselect.options[selectedIndex].text;
    sourceselect.options[selectedIndex].value   = oldvals[0];
	sourceselect.options[selectedIndex].text    = oldvals[1];
    sourceselect.selectedIndex                  = selectedIndex-1;
  }
  if (selectedIndex < sourceselect.options.length-botnum-1 && tomove == "1" && selectedIndex >= topnum)     {
    oldvals = new Array(sourceselect.options[selectedIndex + 1].value, sourceselect.options[selectedIndex + 1].text);
    sourceselect.options[selectedIndex+1].value = sourceselect.options[selectedIndex].value;
    sourceselect.options[selectedIndex+1].text  = sourceselect.options[selectedIndex].text;
    sourceselect.options[selectedIndex].value   = oldvals[0];
	sourceselect.options[selectedIndex].text    = oldvals[1];
    sourceselect.selectedIndex                  = selectedIndex+1;
  }
}
function change_order(){
for(i=0; i<document.forms[0].order_box.length; i++){
which_id=(MM_findObj('q_id'+(i+1)));
which_id.value=document.forms[0].order_box[i].value;
}
}
function trySubmit(){
	document.forms[0].field_length.value = document.forms[0].order_box.length;
	return confirm("Do you realy want to save changes in questions order?");
}
//-->
</script>
</HEAD>
<BODY>
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> Quiz questions order</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <table>
        <tr> 
          <td class="subheads">Here is the list of available questions in subject 
            <b><%=(subject.Fields.Item("subject_name").Value)%></b>, topic <b><%=(topic.Fields.Item("topic_name").Value)%></b>:</td>
        </tr>
        <tr> 
          <td> 
            <form name="order_form" method="post" action="<%=MM_editAction%>" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
              <table>
                <tr> 
                  <td > 
                    <input type="button" name="up2" value="Move UP" onClick="WA_MoveSelectedInList(MM_findObj('order_box'),0,0,0); change_order()" class="quiz_button">
                    <input type="button" name="down2" value="Move DOWN" onClick="WA_MoveSelectedInList(MM_findObj('order_box'),1,0,0); change_order()" class="quiz_button">
                    <input type="submit" name="Submit2" value="Save changes" class="quiz_button" <%call IsEditOK%>>
                  </td>
                </tr>
                <tr> 
                  <td > 
                    <select name="order_box" size="20" class="formitem1">
                      <%
While (NOT questions.EOF)
%>
                      <option value="<%=(questions.Fields.Item("ID_question").Value)%>"><%=left((questions.Fields.Item("question_body").Value),50) & "..."%></option>
                      <%
  questions.MoveNext()
Wend
'If (questions.CursorType > 0) Then
'  questions.MoveFirst
'Else
  questions.Requery
'End If
%>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td > 
                    <input type="button" name="up" value="Move UP" onClick="WA_MoveSelectedInList(MM_findObj('order_box'),0,0,0); change_order()" class="quiz_button">
                    <input type="button" name="down" value="Move DOWN" onClick="WA_MoveSelectedInList(MM_findObj('order_box'),1,0,0); change_order()" class="quiz_button">
                    <input type="submit" name="Submit" value="Save changes" class="quiz_button" <%call IsEditOK%>>
                  </td>
                </tr>
                <tr> 
                  <td > 
                    <input type="hidden" name="field_length">
                    <%
ii=1
While (NOT questions.EOF)
%>
                    <input type="hidden" name="q_id<%=ii%>" value="<%=(questions.Fields.Item("ID_question").Value)%>">
                    <input type="hidden" name="q_ord<%=ii%>" value="<%=ii%>">
                    <%
ii=ii+1			  
  questions.MoveNext()
Wend
'If (questions.CursorType > 0) Then
'  questions.MoveFirst
'Else
  questions.Requery
'End If
%>
                    <input type="hidden" name="MM_update" value="true">
                    Select one item in the list and press 'Move UP' or 'Move DOWN' 
                    buttons to change the order.</td>
                </tr>
              </table>
            </form>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td align="left" valign="bottom" > 
      <input type="button" name="goback" value="Go back to question list" class="quiz_button" onClick="document.location='q_list_of_questions.asp?subj=<%=subj%>&topic=<%=top%>'">
      <input type="button" name="goback" value="Go back to topic list" class="quiz_button" onClick="document.location='q_list_of_topics.asp?subj=<%=subj%>'">
      or 
      <input type="button" name="goback" value="Go back to subject list" class="quiz_button" onClick="document.location='q_list_of_subjects.asp'">
    </td>
  </tr>
</table>
<p>&nbsp;</p></BODY>
</HTML>

<%
call log_the_page ("Quiz Reorder Questions " & subj & ", " & top)
%>

<%
questions.Close()
%>
<%
subject.Close()
%>
<%
topic.Close()
%>
