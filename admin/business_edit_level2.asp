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
  MM_editTable = "q_info2"
  MM_editColumn = "ID_info2"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "business_level2.asp"
  MM_fieldsStr  = "business_name|value|business|value|active|value"
  MM_columnsStr = "info2|',none,''|info2_info1|none,none,NULL|info2_active|none,1,0"

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
    call log_the_page ("BBG Execute - UPDATE Info2: " & MM_recordId)
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim business__MMColParam
business__MMColParam = "1"
if (Request.QueryString("site") <> "") then business__MMColParam = Request.QueryString("site")
%>
<%
set business = Server.CreateObject("ADODB.Recordset")
business.ActiveConnection = Connect
business.Source = "SELECT *  FROM q_info2  WHERE ID_info2 = " + Replace(business__MMColParam, "'", "''") + ""
business.CursorType = 0
business.CursorLocation = 3
business.LockType = 3
business.Open()
business_numRows = 0
%>
<%
set business_1 = Server.CreateObject("ADODB.Recordset")
business_1.ActiveConnection = Connect
business_1.Source = "SELECT * FROM q_info1"
business_1.CursorType = 0
business_1.CursorLocation = 3
business_1.LockType = 3
business_1.Open()
business_1_numRows = 0
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Business site. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
function trySubmit()
{
	if (document.forms[0].business_name.value.length<2)
	{
		alert("Sorry, you must enter a name for a business site!\n(min. 2 characters)");
		return false;
	}
	if (confirm("Are you sure you want to update properties of this business site?"))	{	document.forms[0].submit();
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
<BODY onUnload="<% call on_page_unload %>" onLoad="change=false;">
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> Business site edit</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="add_business" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr> 
            <td width="99%" >Name of the business site</td>
          </tr>
          <tr> 
            <td width="99%" > 
              <input type="text" name="business_name" onChange="change=true;" size="60" class="formitem1" value="<%=(business.Fields.Item("info2").Value)%>">
            </td>
          </tr>
          <tr> 
            <td width="99%" >Name of the business this site 
              belongs to:</td>
          </tr>
          <tr> 
            <td width="99%" > 
              <select onChange="change=true;" name="business" class="formitem1">
                <%
While (NOT business_1.EOF)
%>
                <option value="<%=(business_1.Fields.Item("ID_info1").Value)%>" <%if (CStr(business_1.Fields.Item("ID_info1").Value) = CStr(business.Fields.Item("info2_info1").Value)) then Response.Write("SELECTED") : Response.Write("")%> ><%=(business_1.Fields.Item("info1").Value)%></option>
                <%
  business_1.MoveNext()
Wend
'If (business_1.CursorType > 0) Then
'  business_1.MoveFirst
'Else
  business_1.Requery
'End If
%>
              </select>
            </td>
          </tr>
           <tr> 
            <td width="99%" >Business active? 
              <input <%If (CStr(abs(business.Fields.Item("info2_active").Value)) = CStr(1)) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="active" value="checkbox">
            </td>
          </tr>
          <tr> 
            <td width="99%" >&nbsp; </td>
          </tr>
          <tr> 
            <td width="99%" > 
              <input type="reset" name="Submit2" value="Reset the form" class="quiz_button">
              <input type="submit" name="Submit" value="Update this business site" class="quiz_button" <%call IsEditOK%>>
              or 
              <input type="button" name="goback" value="Go back to business list" class="quiz_button" onClick="document.location='business_level1.asp?business=<%=(business.Fields.Item("info2_info1").Value)%>'">
            </td>
          </tr>
        </table>
        <input type="hidden" name="MM_update" value="true">
        <input type="hidden" name="MM_recordId" value="<%= business.Fields.Item("ID_info2").Value %>">
      </form>
    </td>
  </tr>
</table>
<p> 
<p>&nbsp; </p>
</BODY>
</HTML>

<%
call log_the_page ("BBG Edit Info2: " & (business.Fields.Item("ID_info2").Value))
%>

<%
business.Close()
%>
<%
business_1.Close()
%>

