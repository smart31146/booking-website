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
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then

  MM_editConnection = Connect
  MM_editTable = "b_hlp"
  MM_editRedirectUrl = "b_list_of_help.asp"
  MM_fieldsStr  = "newhelp|value|UID|value"
  MM_columnsStr = "hlp_name|',none,''|hlp_UID|',none,''"

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
    call log_the_page ("BBG Execute - INSERT Help")	
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
numbers=1
%>
<%
set helps = Server.CreateObject("ADODB.Recordset")
helps.ActiveConnection = Connect
helps.Source = "SELECT *  FROM b_hlp"
helps.CursorType = 0
helps.CursorLocation = 3
helps.LockType = 3
helps.Open()
helps_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
helps_numRows = helps_numRows + Repeat1__numRows
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Help tab list. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
function trySubmit()
{
	if (document.forms[0].newhelp.value.length<3)
	{
		alert("Sorry, you must enter a name for a new Help tab!\n(min. 3 characters)");
		return false;
	}
	if (confirm("Are you sure you want to add a new Help tab?"))	{	document.forms[0].submit();
	return false;
	}
return false;
}

//-->
</script>
</HEAD>
<BODY>
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> HELP tab list</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form name="add_subject" method="POST" action="<%=MM_editAction%>" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr> 
            <td colspan="2" class="subheads">Help tabs:</td>
          </tr>
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
            <td class="text" width="10"><img src="../admin/images/back.gif" width="18" height="14"></td>
            <td class="text"><a href="../admin/main.asp">...Home page </a></td>
          </tr>
          <% If Not helps.EOF Or Not helps.BOF Then %>
          <% 
While ((Repeat1__numRows <> 0) AND (NOT helps.EOF)) 
%>
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
            <td class="text" width="20"><%=numbers%></td>
            <td width="580" class="text"><a href="b_edit_helps.asp?help=<%=(helps.Fields.Item("ID_hlp").Value)%>"><%=(helps.Fields.Item("hlp_name").Value)%></a> </td>
          </tr>
          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  helps.MoveNext()
  numbers=numbers+1
Wend
%>
          <% End If ' end Not helps.EOF Or NOT helps.BOF %>
          <% If helps.EOF And helps.BOF Then %>
          <tr> 
            <td >&nbsp;</td>
            <td >Sorry, 
              there are no Help tabs in the BBP currently.</td>
          </tr>
          <% End If ' end helps.EOF And helps.BOF %>
          <tr class="table_normal"> 
            <td ><img src="images/new2.gif" width="11" height="13"></td>
            <td width="99%" > 
              <input type="text" name="newhelp" size="85" class="formitem1">
            </td>
          </tr>
          <tr> 
            <td>
              <input type="hidden" name="UID" value="<%=GetUniqueID("h_",20,"")%>">
            </td>
            <td width="99%" > 
              <input type="reset" name="Submit2" value="Reset the form" class="quiz_button">
              <input type="submit" name="Submit" value="Add this new help tab" class="quiz_button" <%call IsEditOK%>>
            </td>
          </tr>
        </table>
        <input type="hidden" name="MM_insert" value="true">
      </form>
      <p>&nbsp;</p>
      </td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
call log_the_page ("BBG List Help")
%>

<%
helps.Close()
%>
