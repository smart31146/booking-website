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
  MM_editTable = "glossary"
  MM_editRedirectUrl = "glossary_level1.asp"
  MM_fieldsStr  = "glossname|value"
  MM_columnsStr = "name|',none,''"

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
    call log_the_page ("BBG Execute - INSERT Info1")
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
set gloss = Server.CreateObject("ADODB.Recordset")
gloss.ActiveConnection = Connect
gloss.Source = "SELECT * FROM glossary"
gloss.CursorType = 0
gloss.CursorLocation = 3
gloss.LockType = 3
gloss.Open()
gloss_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
gloss_numRows = gloss_numRows + Repeat1__numRows
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: glossary list. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
function trySubmit()
{
	if (document.forms[0].glossname.value.length<2)
	{
		alert("Sorry, you must enter a name for a new glossary item!\n(min. 2 characters)");
		return false;
	}
	if (confirm("Are you sure you want to add a new item?"))	{	document.forms[0].submit();
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
    <td align="left" valign="bottom" class="heading"> Glossary list</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form name="add_bus" method="POST" action="<%=MM_editAction%>" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr> 
            <td colspan="2" class="subheads">Glossary:</td>
            <td align="right" class="subheads">&nbsp;</td>
          </tr>
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
            <td class="text" width="10"><img src="../admin/images/back.gif" width="18" height="14"></td>
            <td colspan="3" class="text"><a href="../admin/main.asp">...Home page 
              </a></td>
          </tr>
          <% If Not gloss.EOF Or Not gloss.BOF Then %>
          <% 
While ((Repeat1__numRows <> 0) AND (NOT gloss.EOF)) 
%>
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">
            <td class="text" width="20"><%=numbers%></td>
            <td width="560" class="text"><a href="../admin/glossary_level2.asp?glossary=<%=(gloss.Fields.Item("GID").Value)%>"><%=(gloss.Fields.Item("name").Value)%></a></td>
            <td width="20" class="text">
				<%if gloss.Fields.Item("Active").Value then%>
					<img src="images/1.gif">
                <%else%>					
					<img src="images/0.gif">
				<%end if%>					
            </td>
            <td width="20"  align="right"><a href="../admin/glossary_edit_level1.asp?glossary=<%=(gloss.Fields.Item("GID").Value)%>"><img src="images/edit.gif" width="16" height="15" border="0"></a></td>
          </tr>
          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  gloss.MoveNext()
  numbers=numbers+1
Wend
%>
          <% End If ' end Not gloss.EOF Or NOT gloss.BOF %>
          <% If gloss.EOF And gloss.BOF Then %>
          <tr> 
            <td >&nbsp;</td>
            <td colspan="3" >Sorry, 
              there is no glossary in the BBP currently.</td>
          </tr>
          <% End If ' end gloss.EOF And gloss.BOF %>
          <tr class="table_normal"> 
            <td ><img src="images/new2.gif" width="11" height="13"></td>
            <td width="99%"  colspan="3"> 
              <input type="text" name="glossname" size="85" class="formitem1">
            </td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td width="99%"  colspan="2"> 
              <input type="reset" name="Submit2" value="Reset the form" class="quiz_button">
              <input type="submit" name="Submit" value="Add this new glossary" class="quiz_button" <%call IsEditOK%>>
            </td>
          </tr>
        </table>
        <input type="hidden" name="MM_insert" value="true">
      </form>
      <p>&nbsp;</p>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
call log_the_page ("BBG List Info1")
%>

<%
gloss.Close()
%>
