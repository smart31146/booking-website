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
' *** Update choices

If (CStr(Request("MM_update")) <> "") Then

for iii = 1 to cInt(Request.form("field_length"))

  MM_editConnection = Connect
  MM_editTable = "subjects"
  MM_editColumn = "ID_subject"
  MM_recordId = "" + Request.Form("q_id" & iii) + "" 
  MM_editRedirectUrl = "q_list_of_subjects.asp"
  MM_fieldsStr  = "q_ord" & iii & "|value"
  MM_columnsStr = "subject_ord|',none,''"

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
    call log_the_page ("Quiz Execute - UPDATE Subjects: " & MM_recordId)
  End If
next
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
End If
%>
<%
set subjects = Server.CreateObject("ADODB.Recordset")
subjects.ActiveConnection = Connect
subjects.Source = "SELECT subjects.ID_subject, subjects.subject_name, subjects.subject_active_q, subjects.ID_subject  FROM subjects  GROUP BY subjects.ID_subject, subjects.subject_name, subjects.subject_active_q, subjects.subject_ord, subjects.ID_subject  ORDER BY subjects.subject_ord, subjects.ID_subject;"
subjects.CursorType = 0
subjects.CursorLocation = 3
subjects.LockType = 3
subjects.Open()
subjects_numRows = 0
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz subjects order. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
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
	return confirm("Do you realy want to save changes in subjects order?");
}
//-->
</script>
</HEAD>
<BODY>
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> Quiz subjects order</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <table>
        <tr> 
          <td class="subheads">Here is the list of available subjects</td>
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
While (NOT subjects.EOF)
%>
                      <option value="<%=(subjects.Fields.Item("ID_subject").Value)%>"><%=(subjects.Fields.Item("subject_name").Value)%></option>
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
While (NOT subjects.EOF)
%>
                    <input type="hidden" name="q_id<%=ii%>" value="<%=(subjects.Fields.Item("ID_subject").Value)%>">
                    <input type="hidden" name="q_ord<%=ii%>" value="<%=ii%>">
                    <%
ii=ii+1			  
  subjects.MoveNext()
Wend
'If (subjects.CursorType > 0) Then
'  subjects.MoveFirst
'Else
  subjects.Requery
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
    <td align="left" valign="bottom">
      <input type="button" name="goback" value="Go back to subject list" class="quiz_button" onClick="document.location='q_list_of_subjects.asp'">
    </td>
  </tr>
</table>
<p>&nbsp;</p></BODY>
</HTML>

<%
call log_the_page ("Quiz Reorder Subjects")
%>

<%
subjects.Close()
%>

