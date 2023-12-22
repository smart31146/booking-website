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
  MM_editTable = "toreplace"
  MM_editRedirectUrl = "b_list_of_replace.asp"
  MM_fieldsStr  = "what|value|bywhat|value|bbg|value|training|value|quiz|value|search|value|repl_active|value|UID|value"
  MM_columnsStr = "repl_what|',none,''|repl_bywhat|',none,''|repl_bbg|none,1,0|repl_tr|none,1,0|repl_q|none,1,0|repl_search|none,1,0|repl_active|none,1,0|repl_UID|',none,''"

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
    call log_the_page ("BBG Execute - INSERT Replace")
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
set replaceto = Server.CreateObject("ADODB.Recordset")
replaceto.ActiveConnection = Connect
replaceto.Source = "SELECT * FROM toreplace"
replaceto.CursorType = 0
replaceto.CursorLocation = 3
replaceto.LockType = 3
replaceto.Open()
replaceto_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
replaceto_numRows = replaceto_numRows + Repeat1__numRows
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Replacements. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}


function trySubmit()
{
	if (document.forms[0].what.value.length<2)
	{
		alert("Sorry, you must enter a string of text to be replaced!\n(min. 2 characters)");
		return false;
	}
	if (document.forms[0].bywhat.value.length<2)
	{
		alert("Sorry, you must enter a replacements for the string you want to replace!\n(min. 2 characters)");
		return false;
	}
	if (confirm("Are you sure you want to add a new replacement?"))	{	document.forms[0].submit();
	return false;
	}
return false;
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</HEAD>
<BODY>
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> Global replacements list</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form name="add_repl" method="POST" action="<%=MM_editAction%>" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr> 
            <td colspan="6" class="subheads">Replacements:</td>
          </tr>
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
            <td class="text" width="10"><img src="../admin/images/back.gif" width="18" height="14"></td>
            <td colspan="5" class="text"><a href="../admin/main.asp">...Home page 
              </a></td>
          </tr>
          <% If Not replaceto.EOF Or Not replaceto.BOF Then %>
          <% 
While ((Repeat1__numRows <> 0) AND (NOT replaceto.EOF)) 
%>
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')" > 
            <td class="text" width="20"><%=numbers%></td>
            <td width="300" class="text" onClick="document.location ='../admin/b_edit_replace.asp?id_replace=<%=(replaceto.Fields.Item("id_replace").Value)%>'"><a href="../admin/b_edit_replace.asp?id_replace=<%=(replaceto.Fields.Item("id_replace").Value)%>"><%=(replaceto.Fields.Item("repl_what").Value)%></a></td>
            <td width="260" class="text"><%=ClearHTMLTags((replaceto.Fields.Item("repl_bywhat").Value),2)%></td>
            <td width="20"  align="right"> 
              <%if abs(replaceto.Fields.Item("repl_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
            </td>
            <td width="20"  align="right"> 
              <%if Edit_OK then %>
              <a href="b_replace_del.asp?id_replace=<%=(replaceto.Fields.Item("id_replace").Value)%>" onClick="javascript: return (confirm('You are just about to delete this replacement.\nAre you sure you want to do that?'));"><img src="images/bin.gif" width="16" height="16" border="0"></a> 
              <% end if %>
            </td>
          </tr>
          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  replaceto.MoveNext()
  numbers=numbers+1
Wend
%>
          <% End If ' end Not replaceto.EOF Or NOT replaceto.BOF %>
          <% If replaceto.EOF And replaceto.BOF Then %>
          <tr> 
            <td >&nbsp;</td>
            <td colspan="5" >Sorry, 
              there are no replacements in the BBP currently.</td>
          </tr>
          <% End If ' end replaceto.EOF And replaceto.BOF %>
          <tr class="table_normal"> 
            <td ><img src="images/new2.gif" width="11" height="13"></td>
            <td width="6%" > 
              <input type="text" name="what" size="20" class="formitem1">
            </td>
            <td colspan="4" > 
              <input type="text" name="bywhat" size="60" class="formitem1">
            </td>
          </tr>
          <tr class="table_normal"> 
            <td >&nbsp;</td>
            <td colspan="5" >Replacement active for: BBG 
              <input type="checkbox" name="bbg" value="1">
              , Training 
              <input type="checkbox" name="training" value="1">
              , Quiz 
              <input type="checkbox" name="quiz" value="1">
              BBG Search 
              <input type="checkbox" name="search" value="1" checked>
            </td>
          </tr>
          <tr> 
            <td> 
              <input type="hidden" name="UID" value="<%=GetUniqueID("f_",20,"")%>">
            </td>
            <td width="99%"  colspan="5"> 
              <input type="reset" name="Submit2" value="Reset this form" class="quiz_button">
              <input type="submit" name="Submit" value="Add a new replacement" class="quiz_button" <%call IsEditOK%>>
              <input type="hidden" name="repl_active" value="1">
            </td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td width="99%"  colspan="5">If you need to include 
              a link to a particular page within a BBG, use this <a href="javascript:" onClick="MM_openBrWindow('_link_generator.asp','linkgenerator','width=600,height=300')">LINK 
              GENERATOR </a></td>
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
call log_the_page ("BBG List Replace")
%>

<%
replaceto.Close()
%>
