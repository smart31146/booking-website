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
  MM_editTable = "b_faq"
  MM_editColumn = "ID_faq"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "b_list_of_faq.asp"
  MM_fieldsStr  = "faq_name|value|faq_tab|value"
  MM_columnsStr = "faq_name|',none,''|faq_tab|',none,''"

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
    call log_the_page ("BBG Execute - UPDATE FAQ: " & MM_recordId)
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim faqs__MMColParam
faqs__MMColParam = "1"
if (Request.QueryString("faq")   <> "") then faqs__MMColParam = Request.QueryString("faq")  
%>
<%
set faqs = Server.CreateObject("ADODB.Recordset")
faqs.ActiveConnection = Connect
faqs.Source = "SELECT *  FROM b_faq  WHERE b_faq.ID_faq = " + Replace(faqs__MMColParam, "'", "''") + " ;"
faqs.CursorType = 0
faqs.CursorLocation = 3
faqs.LockType = 3
faqs.Open()
faqs_numRows = 0
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: FAQ tab editor. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
function trySubmit()
{
	if (document.forms[0].faq_name.value.length<3)
	{
		alert("Sorry, you must enter a name for a new set of preferences!\n(min. 3 characters)");
		return false;
	}
	if (confirm("Are you sure you want to update this FAQ tab?"))	{	document.forms[0].submit();
	return false;
	}
return false;
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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
    <td align="left" valign="bottom" class="heading"> BBG FAQ tab editor</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="add_subject" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr> 
            <td width="99%" > FAQ tab name</td>
          </tr>
          <tr> 
            <td width="99%" ><a href="../admin/q_list_of_topics.asp?subj="></a> 
              <input type="text" name="faq_name" onChange="change=true;" size="60" class="formitem1" value="<%=(faqs.Fields.Item("faq_name").Value)%>">
            </td>
          </tr>
          <tr> 
            <td width="99%" >FAQ tab content <a href="javascript:" onClick="MM_openBrWindow('_editor.asp?field=faq_tab','editor','width=520,height=400')"><img src="images/editor.gif" width="24" height="10" border="0"></a> 
            </td>
          </tr>
          <tr> 
            <td width="99%" > 
              <textarea name="faq_tab" cols="80" class="formitem1" rows="15" onChange="change=true"><%=(faqs.Fields.Item("faq_tab").Value)%></textarea>
            </td>
          </tr>
          <tr> 
            <td width="99%" > 
              <input type="reset" name="Submit2" value="Reset the form" class="quiz_button">
              <input type="submit" name="Submit" value="Update this FAQ tab" class="quiz_button" <%call IsEditOK%>>
              or 
              <input type="button" name="goback" value="Go back to FAQ list" class="quiz_button" onClick="document.location='b_list_of_faq.asp'">
            </td>
          </tr>
        </table>
        <input type="hidden" name="MM_recordId" value="<%= faqs.Fields.Item("ID_faq").Value %>">
        <input type="hidden" name="MM_update" value="true">
      </form>
    </td>
  </tr>
</table>
<p> 
<p>&nbsp; </p>
</BODY>
</HTML>

<%
call log_the_page ("BBG Edit FAQ: " & (faqs.Fields.Item("ID_faq").Value))
%>

<%
faqs.Close()
%>

