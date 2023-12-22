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
  MM_editTable = "preferences"
  MM_editColumn = "ID_pref"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "preferences.asp"
  MM_fieldsStr  = "pref_name|value|pref_active|value|pref_bbg_avail|value|pref_quiz_avail|value|pref_IP|value|pref_IP_list|value|pref_IP_admin|value|pref_IP_admin_list|value|pref_forgot_pass|value|pref_change_pass|value|pref_self_reg|value|pref_offline|value|pref_date|value|enterAColor|value"
  MM_columnsStr = "pref_name|',none,''|pref_active|none,1,0|pref_bbg_avail|none,1,0|pref_quiz_avail|none,1,0|pref_IP_protect|none,1,0|pref_IP_list|',none,''|pref_admin_IP_protect|none,1,0|pref_admin_IP_list|',none,''|pref_forgot_pass|none,1,0|pref_change_pass|none,1,0|pref_self_reg|none,1,0|pref_offline|none,1,0|pref_date|',none,NULL|pref_admin_color|',none,''"

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
Response.Write MM_editQuery
  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    if Edit_OK = true then MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
    call log_the_page ("BBG Execute - UPDATE Preferences: " & MM_recordId)
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim pid
If (Request.QueryString("pid") <> "") Then 
pid = cInt(Request.QueryString("pid"))
Else 
Response.Redirect("error.asp?" & request.QueryString) 
End If
%>
<%
set preferences = Server.CreateObject("ADODB.Recordset")
preferences.ActiveConnection = Connect
preferences.Source = "SELECT *  FROM preferences  WHERE ID_pref = " + Replace(pid, "'", "''") + " ;"
preferences.CursorType = 0
preferences.CursorLocation = 3
preferences.LockType = 3
preferences.Open()
preferences_numRows = 0
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN:: Preferences edit. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script src="../jquery-1.11.1.js?v=bbp34"></script>
<script src='../js/spectrum.js?v=bbp34'></script>
<link rel='stylesheet' href='../style/spectrum.css' />

<script language="JavaScript">
<!--
function trySubmit()
{
	if (document.forms[0].pref_name.value.length<2)
	{
		alert("Sorry, you must enter a name for a set of preferences!\n(min. 2 characters)");
		return false;
	}
	if (confirm("Are you sure you want to update this set of preferences?"))	{	document.forms[0].submit();
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
$(document).ready(function() {
$(".basic").spectrum({
    color: "#f00",
    change: function(color) {
        $("#enterAColor").val(color.toHexString());
    }
});
$(".basic").spectrum("set", "<%=(preferences.Fields.Item("pref_admin_color").Value)%>");
$("#btnDefault").click(function(){
 $("#enterAColor").val("default");
});
});
</script>

</HEAD>
<BODY onLoad="change=false;" onUnload="<% call on_page_unload %>">
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> Preferences</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="pref" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr> 
            <td colspan="3" class="subheads">Saved profile:</td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" colspan="2"><a href="../admin/preferences_edit.asp?pid="></a> 
              Name of this profile<br>
              <input type="text" name="pref_name" onChange="change=true;" size="80" class="quiz_button" value="<%=(preferences.Fields.Item("pref_name").Value)%>">
            </td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="30"> 
              <input onChange="change=true;" <%If (Abs(preferences.Fields.Item("pref_active").Value) = 1) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="pref_active" value="1">
            </td>
            <td class="text" width="570">Activate this profile</td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="30"> 
              <input onChange="change=true;" <%If (Abs(preferences.Fields.Item("pref_bbg_avail").Value) = 1) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="pref_bbg_avail" value="1">
            </td>
            <td class="text" width="570">Activate Guide module</td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="30"> 
              <input onChange="change=true;" <%If (Abs(preferences.Fields.Item("pref_quiz_avail").Value) = 1) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="pref_quiz_avail" value="1">
            </td>
            <td class="text" width="570">Activate Training and Quiz module</td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="30"> 
              <input onChange="change=true;" <%If (Abs(preferences.Fields.Item("pref_IP_protect").Value) = 1) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="pref_IP" value="1">
            </td>
            <td class="text" width="570">Activate IP address protection for front-end</td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" colspan="2"> List of allowed IP addresses (use &quot;;&quot;  as delimiter and &quot;*&quot; as wildcard i.e. 192.168.*.*)<br>
              <input onChange="change=true;" value="<%=(preferences.Fields.Item("pref_IP_list").Value)%>" type="text" name="pref_IP_list" size="80" class="quiz_button">
            </td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="30"> 
              <input onChange="change=true;" <%If (Abs(preferences.Fields.Item("pref_admin_IP_protect").Value) = 1) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="pref_IP_admin" value="1">
            </td>
            <td class="text" width="570">Activate IP address protection for admin back-end </td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" colspan="2"> List of allowed IP addresses (use &quot;;&quot; as delimiter and &quot;*&quot; as wildcard i.e. 192.168.*.*)<br>
              <input onChange="change=true;" value="<%=(preferences.Fields.Item("pref_admin_IP_list").Value)%>" type="text" name="pref_IP_admin_list" size="80" class="quiz_button">
            </td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="30"> 
              <input onChange="change=true;" <%If (Abs(preferences.Fields.Item("pref_forgot_pass").Value) = 1) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="pref_forgot_pass" value="1">
            </td>
            <td class="text" width="570">Activate password recovery</td>
          </tr>
		  <tr class="table_normal"> 
            <td class="text" width="30"> 
              <input onChange="change=true;" <%If (Abs(preferences.Fields.Item("pref_change_pass").Value) = 1) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="pref_change_pass" value="1">
            </td>
            <td class="text" width="570">Activate change password</td>
          </tr>
		  <tr class="table_normal"> 
            <td class="text" width="30"> 
              <input onChange="change=true;" <%If (Abs(preferences.Fields.Item("pref_self_reg").Value) = 1) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="pref_self_reg" value="1">
            </td>
            <td class="text" width="570">Activate self registration</td>
          </tr>
		  <tr class="table_normal"> 
            <td class="text" width="30"> 
              <input onChange="change=true;" <%If (Abs(preferences.Fields.Item("pref_offline").Value) = 1) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="pref_offline" value="1">
            </td>
            <td class="text" width="570">Activate offline function</td>
          </tr>
		  <tr class="table_normal">
 		  
		  <td class="text" width="20">Background color</td>
           <td class="text" width="20"> 
              <input type='text' class="basic"/> <input id="enterAColor" name="enterAColor" type="text" size="5" value="<%=(preferences.Fields.Item("pref_admin_color").Value)%>" /> <span class="quiz_button" id="btnDefault" style="padding:5px;">Set default</span>
            </td>
            
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="30">&nbsp;</td>
            <td class="text" width="570">&nbsp;</td>
          </tr>
          <tr> 
            <td >&nbsp;</td>
            <td width="99%"  colspan="2"> 
              <input type="reset" name="Submit2" value="Reset the form" class="quiz_button">
              <input type="submit" name="Submit" value="Update this preference" class="quiz_button" <%call IsEditOK%>>
              or 
              <input type="button" name="goback" value="Go back to preferences" class="quiz_button" onClick="document.location='preferences.asp'">
              <input type="hidden" name="pref_date" value="<% =cDateSql(Now())%>">
            </td>
          </tr>
        </table>
        <p> 
          <input type="hidden" name="MM_update" value="true">
          <input type="hidden" name="MM_recordId" value="<%= preferences.Fields.Item("ID_pref").Value %>">
      </form>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
call log_the_page ("BBG Edit Preferences: " & (preferences.Fields.Item("ID_pref").Value))
%>

<%
preferences.Close()
%>
