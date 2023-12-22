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
  MM_editTable = "preferences"
  MM_editRedirectUrl = "preferences.asp"
  MM_fieldsStr  = "pref_name|value|pref_active|value|pref_bbg_avail|value|pref_training_avail|value|pref_quiz_avail|value|pref_qanda_avail|value|pref_IP|value|pref_IP_list|value|pref_IP_admin|value|pref_IP_admin_list|value|pref_training_force|value|pref_bbg_login|value|pref_quiz_free|value|pref_splash|value|pref_pdf|value|pref_date|value|pass_rate|value"
  MM_columnsStr = "pref_name|',none,''|pref_active|none,1,0|pref_bbg_avail|none,1,0|pref_training_avail|none,1,0|pref_quiz_avail|none,1,0|pref_qanda_avail|none,1,0|pref_IP_protect|none,1,0|pref_IP_list|',none,''|pref_admin_IP_protect|none,1,0|pref_admin_IP_list|',none,''|pref_training_force_a|none,1,0|pref_bbg_login|none,1,0|pref_quiz_free|none,1,0|pref_bbg_splash|none,1,0|pref_bbg_pdf|none,1,0|pref_date|',none,NULL|pass_rate|',none,''"

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
    call log_the_page ("BBG Execute - INSERT Preferences")	
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Preferences add. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function trySubmit()
{
	if (document.forms[0].pref_name.value.length<2)
	{
		alert("Sorry, you must enter a name for a new set of preferences!\n(min. 2 characters)");
		return false;
	}
	if (confirm("Are you sure you want to add a new set of preferences?"))	{	document.forms[0].submit();
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
<BODY onLoad="change=false;" onUnload="<% call on_page_unload %>">
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading">Preferences</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form name="pref" method="POST" action="<%=MM_editAction%>" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr> 
            <td colspan="3" class="subheads">Please, create a new set of preferences:</td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="10" align="left" valign="top">1</td>
            <td class="text" colspan="2"><a href="../admin/preferences_edit.asp?pid="></a> 
              Name of this profile<br>
              <input type="text" name="pref_name" onChange="change=true;" size="80" class="quiz_button">
            </td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="10" align="left" valign="top">2</td>
            <td class="text" width="30"> 
              <input type="checkbox" name="pref_active" onChange="change=true;" value="1" checked>
            </td>
            <td class="text" width="570">Activate this profile</td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="10" align="left" valign="top">3</td>
            <td class="text" width="30"> 
              <input type="checkbox" name="pref_bbg_avail" onChange="change=true;" value="1" checked>
            </td>
            <td class="text" width="570">Activate BBG module</td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="10" align="left" valign="top">4</td>
            <td class="text" width="30"> 
              <input type="checkbox" name="pref_training_avail" onChange="change=true;" value="1" checked>
            </td>
            <td class="text" width="570">Activate Training module</td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="10" align="left" valign="top">5</td>
            <td class="text" width="30"> 
              <input type="checkbox" name="pref_quiz_avail" onChange="change=true;" value="1" checked>
            </td>
            <td class="text" width="570">Activate Quiz module</td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="10" align="left" valign="top">6</td>
            <td class="text" width="30"> 
              <input type="checkbox" name="pref_qanda_avail" onChange="change=true;" value="1" checked>
            </td>
            <td class="text" width="570">Activate Q and A module</td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="10" align="left" valign="top">7</td>
            <td class="text" width="30"> 
              <input type="checkbox" name="pref_IP" onChange="change=true;" value="1">
            </td>
            <td class="text" width="570">Activate IP address protection for front-end</td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="10" align="left" valign="top">8</td>
            <td class="text" colspan="2"> List of allowed IP addresses (use &quot;;&quot; 
              as delimiter and &quot;*&quot; as wildcard i.e. 192.168.*.*)<br>
              <input type="text" name="pref_IP_list" onChange="change=true;" size="80" class="quiz_button">
            </td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="10" align="left" valign="top">9</td>
            <td class="text" width="30"> 
              <input type="checkbox" name="pref_IP_admin" onChange="change=true;" value="1">
            </td>
            <td class="text" width="570">Activate IP address protection for admin 
              back-end </td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="10" align="left" valign="top">10</td>
            <td class="text" colspan="2"> List of allowed IP addresses (use &quot;;&quot; 
              as delimiter and &quot;*&quot; as wildcard i.e. 192.168.*.*)<br>
              <input type="text" name="pref_IP_admin_list" onChange="change=true;" size="80" class="quiz_button">
            </td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="10" align="left" valign="top">11</td>
            <td class="text" width="30"> 
              <input type="checkbox" name="pref_training_force" onChange="change=true;" value="1" checked>
            </td>
            <td class="text" width="570">Force answers in training module</td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="10" align="left" valign="top">12</td>
            <td class="text" width="30"> 
              <input type="checkbox" name="pref_bbg_login" onChange="change=true;" value="1">
            </td>
            <td class="text" width="570">Require login or registration for BBG 
              module</td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="10" align="left" valign="top">13</td>
            <td class="text" width="30"> 
              <input type="checkbox" name="pref_quiz_free" onChange="change=true;" value="1" checked>
            </td>
            <td class="text" width="570">Allow free login to Quiz module</td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="10" align="left" valign="top">14</td>
            <td class="text" width="30"> 
              <input type="checkbox" name="pref_splash" onChange="change=true;" value="1" checked>
            </td>
            <td class="text" width="570">Allow free registration for BBP welcome 
              page (unchecked = LOGIN)</td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="10" align="left" valign="top">15</td>
            <td class="text" width="30"> 
              <input type="checkbox" name="pref_pdf" onChange="change=true;" value="1" checked>
            </td>
            <td class="text" width="570">Show PDF downloads on BBP welcome page</td>
          </tr>
          <tr class="table_normal"> 
            <td class="text" width="10" align="left" valign="top">16</td>
            <td class="text" colspan="2"> 
              Passrate<br>
			  <select name="pass_rate" class="formitem1">
			  <%
				passratecount = 1
				While (passratecount <= 100)
			  %>
			      <option value="<%=passratecount%>" <%if (CStr(50) = CStr(passratecount)) then Response.Write("SELECTED") : Response.Write("")%>>
				    <%=passratecount%>
				  </option>
			  <%
				  passratecount = passratecount + 1
				Wend
			  %>
			  </select>
            </td>
          </tr>

		  <tr class="table_normal"> 
            <td class="text" width="10" align="left" valign="top">&nbsp;</td>
            <td class="text" width="30">&nbsp;</td>
            <td class="text" width="570">&nbsp;</td>
          </tr>
          <tr> 
            <td >&nbsp;</td>
            <td width="99%"  colspan="2"> 
              <input type="reset" name="Submit2" value="Reset this form" class="quiz_button">
              <input type="submit" name="Submit" value="Add this preferences set" class="quiz_button" <%call IsEditOK%>>
              or 
              <input type="button" name="goback" value="Go back to preferences" class="quiz_button" onClick="document.location='preferences.asp'">
              <input type="hidden" name="pref_date" value="<%=cDateSql(Now())%>">
            </td>
          </tr>
        </table>
        <p> 
          <input type="hidden" name="MM_insert" value="true">
      </form>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
call log_the_page ("BBG Add a new Preferences")
%>
