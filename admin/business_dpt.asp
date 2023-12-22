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
  MM_editTable = "q_info3"
  MM_editRedirectUrl = "business_dpt.asp"
  MM_fieldsStr  = "newbus|value"
  MM_columnsStr = "info3|',none,''"

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
    call log_the_page ("BBG Execute - INSERT Info3")	
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
set business = Server.CreateObject("ADODB.Recordset")
business.ActiveConnection = Connect
business.Source = "SELECT *  FROM q_info3"
business.CursorType = 0
business.CursorLocation = 3
business.LockType = 3
business.Open()
business_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
business_numRows = business_numRows + Repeat1__numRows
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: <%=location%>. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
function trySubmit()
{
	if (document.forms[0].newbus.value.length<2)
	{
		alert("Sorry, you must enter a name for a new <% =BBPinfo3 %>!\n(min. 2 characters)");
		return false;
	}
	if (confirm("Are you sure you want to add a <% =BBPinfo3 %>?"))	{	document.forms[0].submit();
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
    <td align="left" valign="bottom" class="heading"> <% =BBPinfo3 %></td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form name="add_bus" method="POST" action="<%=MM_editAction%>" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr> 
            <td colspan="2" class="subheads"><% =BBPinfo3s %>:</td>
          </tr>
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">          <td class="text" width="10"><img src="../admin/images/back.gif" width="18" height="14"></td>
            <td class="text" colspan=2><a href="../admin/main.asp">...Home page </a></td>
          </tr>
          <% If Not business.EOF Or Not business.BOF Then %>
          <% 
While ((Repeat1__numRows <> 0) AND (NOT business.EOF)) 
%>
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
            <td class="text" width="20"><%=numbers%></td>
            <td width="580" class="text"><a href="../admin/business_edit_dpt.asp?business=<%=(business.Fields.Item("id_info3").Value)%>"><%=(business.Fields.Item("info3").Value)%></a></td>
            <td width="20" class="text">
				<%if business.Fields.Item("info3_active").Value then%>
					<img src="images/1.gif">
                <%else%>					
					<img src="images/0.gif">
				<%end if%>					
            </td>
          </tr>
          <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  business.MoveNext()
  numbers=numbers+1
Wend
%>
          <% End If ' end Not business.EOF Or NOT business.BOF %>
          <% If business.EOF And business.BOF Then %>
          <tr> 
            <td >&nbsp;</td>
            <td >Sorry, 
              there are no <% =BBPinfo3s %> in the BBP currently.</td>
          </tr>
          <% End If ' end business.EOF And business.BOF %>
          <tr class="table_normal"> 
            <td ><img src="images/new2.gif" width="11" height="13"></td>
            <td width="99%"  colspan=2> 
              <input type="text" name="newbus" size="85" class="formitem1">
            </td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td width="99%" > 
              <input type="reset" name="Submit2" value="Reset the form" class="quiz_button">
              <input type="submit" name="Submit" value="Add this <% =BBPinfo3 %>" class="quiz_button" <%call IsEditOK%>>
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
call log_the_page ("BBG List Info3")
%>

<%
business.Close()
%>
