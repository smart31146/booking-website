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
Dim comeback
If (Request.QueryString("comeback") <> "") Then 
comeback = cStr(Request.QueryString("comeback"))
Else 
Response.Redirect("error.asp?" & request.QueryString) 
End If
%>
<%
Dim qid
If (Request.QueryString("qid") <> "") Then 
qid = cInt(Request.QueryString("qid"))
Else 
Response.Redirect("error.asp?" & request.QueryString) 
End If
%>

<%
Dim lastchoice
lastchoice = cStr(Request.QueryString("lastchoice"))
if NOT (lastchoice) <> "" then
	lastchoice = "A."
elseif uCase(lastchoice) = "TRUE" then
	lastchoice = "False"
elseif uCase(lastchoice) = "FALSE" then
	lastchoice = "True"
elseif uCase(lastchoice) = "YES" then
	lastchoice = "No"
elseif uCase(lastchoice) = "NO" then
	lastchoice = "Yes"
else
for x = 97 to 122 
	if lcase(left(lastchoice,1)) = chr(x) then 
		lastchoice = ucase(chr(x+1) + ".")
		exit for
	end if
next
end if
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then

  MM_editConnection = Connect
  MM_editTable = "q_choice"
  MM_editRedirectUrl = "" + comeback + ""
  MM_fieldsStr  = "body|value|qid|value|label|value|UID|value"
  MM_columnsStr = "choice_body|',none,''|choice_question|none,none,NULL|choice_label|',none,''|choice_UID|',none,''"

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
    call log_the_page ("Quiz Execute - INSERT Choice")	
    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>Quiz admin </TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
</HEAD>
<BODY onload="javascript:document.forms[0].submit();">
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="form1">
  <input type="hidden" name="body">
  <input type="hidden" name="qid" value ="<%=qid%>">
  <input type="hidden" name="MM_insert" value="true">
  <input type="hidden" name="label" value ="<%=lastchoice%>">
  <input type="hidden" name="UID" value="<%=GetUniqueID("c_",20,"")%>">
</form>
</body>
</HTML>
