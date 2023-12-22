<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->

<%
' *** Edit Operations: declare variables

MM_editAction = CStr(Request("URL"))  '/2degrees/admin/send_individual_email.asp
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If


Dim email_sent_message
' query string to execute
MM_editQuery = ""
Dim array_of_subjects

' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then

  'get the array with the subjects ticked
	array_of_subjects = split(Request.Form ("subjects"), ",")
	'format the date
	'uk_date=Request("send_after")
	'sql_date= left(uk_date,2) &"-"& mid(uk_date,4,2) &"-"& right(uk_date,4)
	sql_date=cDateSql(Request("send_after"))
' Insert into DataBase

'Print array_of_subjects
'	for i = 0 to ubound(array_of_subjects)
'		response.write( array_of_subjects(i) )
'	next
'	response.end()

  MM_editConnection = Connect
  MM_editQuery = "insert into emails (q_user,date_to_send,type) values ('"+Request("individual")+"','"+sql_date+"' , "+request("email_select")+")"

    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

	Dim last_email_inserted

	set last_id_insert = Server.CreateObject("ADODB.Recordset")
	last_id_insert.ActiveConnection = Connect
	last_id_insert.Source = "SELECT max (id) as idd from emails"
	last_id_insert.CursorType = 0
	last_id_insert.CursorLocation = 3
	last_id_insert.LockType = 3
	last_id_insert.Open()
	last_id_insert_numRows = 0
		last_email_inserted=last_id_insert.Fields.Item("idd").Value
	last_id_insert.Close()

	strSql = ""


'loop through the checkboxes which were ticked
for i = 0 to ubound(array_of_subjects)

	strSql = "Insert into subject_email (subject,email) values( '"& array_of_subjects(i)&"', '"& last_email_inserted&"')"
	 'response.write ( MM_editQuery)
    Set MM_insertCmd = Server.CreateObject("ADODB.Command")
    MM_insertCmd.ActiveConnection = Connect
    MM_insertCmd.CommandText = strSql
    MM_insertCmd.Execute
    MM_insertCmd.ActiveConnection.Close
next


    email_sent_message="<br/><br/>That email will be sent on "&FormatDateTime(sql_date,2) &".  To send another please enter the details below"
	call log_the_page ("Sent Individual email")
End If


%>
<%
function WA_VBreplace(thetext)
  if isNull(thetext) then thetext = ""
  newstring = Replace(cStr(thetext),"'","|WA|")
  newstring = Replace(newstring,"\","\\")
  WA_VBreplace = newstring
end function

'Extracting User from q_user and asign the quser.ID_user to uid
set user1 = Server.CreateObject("ADODB.Recordset")
user1.ActiveConnection = Connect
user1.Source = "SELECT * from q_user where ID_user="&Request("individual")
user1.CursorType = 0
user1.CursorLocation = 3
user1.LockType = 3
user1.Open()
user1_numRows = 0


set subjects = Server.CreateObject("ADODB.Recordset")
subjects.ActiveConnection = Connect

'CXS 061123: limit the subjects shown to those assigned to the user so they can't accidentally be sent reminders for subjects they are not supposed to be doing
'subjects.Source = "SELECT ID_subject, subject_name FROM subjects where subject_active_q <> 0"

uid = user1.Fields.Item("ID_user").Value
subjects.Source = "SELECT subjects.ID_subject, subjects.subject_name FROM q_user INNER JOIN subject_user ON q_user.ID_user = subject_user.ID_user INNER JOIN subjects ON subject_user.ID_subject = subjects.ID_subject WHERE q_user.ID_user = "&uid

'Print Subjects

subjects.CursorType = 0
subjects.CursorLocation = 3
subjects.LockType = 3
subjects.Open()
subjects_numRows = 0

set email_info = Server.CreateObject("ADODB.Recordset")
email_info.ActiveConnection = Connect
email_info.Source = "SELECT * FROM email_info WHERE email_isadmin=0 and email_period=0 and email_active=1 and email_previous=0"
'To send the Final Reminder, comment out line above and uncomment line below.
'email_info.Source = "SELECT * FROM email_info WHERE email_isadmin=0 and email_period=14 and email_active=1 and email_previous=3"
email_info.CursorType = 0
email_info.CursorLocation = 3
email_info.LockType = 3
email_info.Open()


%>
<%
numbers=1
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<style type="text/css">@import url(jscalendar/calendar-win2k-1.css);</style>
<script type="text/javascript" src="jscalendar/calendar.js?v=bbp34"></script>
<script type="text/javascript" src="jscalendar/lang/calendar-en.js?v=bbp34"></script>
<script type="text/javascript" src="jscalendar/calendar-setup.js?v=bbp34"></script>

<script language="javascript" type="text/javascript" src="tiny_mce/tiny_mce.js?v=bbp34"></script>
<script language="javascript" type="text/javascript">
tinyMCE.init({
	theme : "advanced",
	mode : "textareas",
	theme_advanced_disable : "cut,copy,paste,undo,redo,image,cleanup,help,code,removeformat"


});
</script>
<script language="JavaScript">
<!--
function WA_ClientSideReplace(theval,findvar,repvar)     {
  var retval = "";
  while (theval.indexOf(findvar) >= 0)    {
    retval += theval.substring(0,theval.indexOf(findvar));
    retval += repvar;
    theval = theval.substring(theval.indexOf(findvar) + String(findvar).length);
  }
  if (retval == "" && theval.indexOf(findvar) < 0)    {
    retval = theval;
  }
  return retval;
}

function WA_UnloadList(thelist,leavevals,bottomnum)    {
  while (thelist.options.length > leavevals+bottomnum)     {
    if (thelist.options[leavevals])     {
      thelist.options[leavevals] = null;
    }
  }
  return leavevals;
}




function WA_subAwithBinC(a,b,c)
{

	var i = c.indexOf(a);
	var l = b.length;

	while (i != -1)	{
		c = c.substring(0,i) + b + c.substring(i + a.length,c.length);  //replace all valid a values with b values in the selected string c.
  i += l
		i = c.indexOf(a,i);
	}
	return c;

}



function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}




function replace(string,text,by)
{
    var strLength = string.length, txtLength = text.length;
    if ((strLength == 0) || (txtLength == 0)) return string;
    var i = string.indexOf(text);
    if ((!i) && (text != string.substring(0,txtLength))) return string;
    if (i == -1) return string;
    var newstr = string.substring(0,i) + by;
    if (i+txtLength < strLength)
        newstr += replace(string.substring(i+txtLength,strLength),text,by);
    return newstr;
}





function trySubmit()
{


	 if (document.forms[0].send_after.value.length<2)
	{
		alert("Sorry, you must enter a date to send the email after!\n");
		return false;
	}
	//check if at least one of the subject checkboxes are checked
	var one_checked=false;

	//if there is only one checkbox
	if(document.forms[0].subjects.checked){
		one_checked=true;
	}
	else{
		for (i=0;i<document.forms[0].subjects.length;i++){
				if(document.forms[0].subjects[i].checked==true){
					one_checked=true;
				}

		}
	}
	 if (one_checked==false)
	{
		alert("Sorry, you must enter a subject that you wish users to take!\n");
		return false;
	}



	return true;
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
<BODY>
<table>
  <tr>
    <td align="left" valign="bottom" class="headers"> Send a reminder email to an individual <%=email_sent_message%></td>
  </tr>
  <tr>
    <td align="left" valign="bottom">
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="send_email" onSubmit="return trySubmit();" >
        <table>
          <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="120">Employee:</td>
            <td class="text" align="left" valign="top" colspan="3">



                <%
					'While (NOT user1.EOF)
					%>
									<%=(user1.Fields.Item("user_firstname").Value)%>&nbsp; <%=(user1.Fields.Item("user_lastname").Value)%>
									<%
					'  user1.MoveNext()
					'Wend
					'If (filter_info1.CursorType > 0) Then
					'  filter_info1.MoveFirst
					'Else
					'  user1.Requery
					'End If
					%>

            </td>
          </tr>
		  <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="120">Email type:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <select name="email_select" >
			        <%
			        while not email_info.eof
			        	%><option value="<%=email_info.fields.item("email_id").value%>"><%=email_info.fields.item("email_name").value%></option><%
			        	email_info.movenext
			        Wend
			        %>
      			</select>
            </td>
          </tr>
		  <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="120">Send after:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <input type="text" name="send_after"  id="send_after" size="18" value="" class="formitem1" readonly="true"><button id="trigger">...</button>

            </td>
          </tr>
		  <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="120">Subjects</td>
            <td class="text" align="left" valign="top" colspan="3">


                <%
					While (NOT subjects.EOF)
					%>
									<input type="checkbox" name="subjects" id="subjects" value="<%=(subjects.Fields.Item("ID_subject").Value)%>" ></input><%=(subjects.Fields.Item("subject_name").Value)%><br/>
									<%

					  subjects.MoveNext()

					Wend
					'If (subjects.CursorType > 0) Then
					'  subjects.MoveFirst
					'Else
					  subjects.Requery
					'End If
					%>

            </td>
          </tr>



          <tr>
            <td class="text" align="left" valign="top"></td>
            <td class="text" align="left" valign="top" colspan="3">
              <input type="hidden" name="active"  value="1" >
			  <input type="hidden" name="individual"  value="<%=Request("individual")%>" >

            </td>
          </tr>

          <tr>
            <td  align="left" valign="top">
              <input type="hidden" name="session" value="<%=getPassword(30, "", "true", "true", "true", "false", "true", "true", "true", "false")%>">
              <input type="hidden" name="current_export">
            </td>
            <td  align="left" valign="top" colspan="3">

              <input type="submit" name="Submit" value="Send" class="quiz_button" >

            </td>
          </tr>
        </table>
        <input type="hidden" name="MM_insert" value="true">
      </form>
	  <script type="text/javascript">
		  Calendar.setup(
			{
			  inputField  : "send_after",         // ID of the input field
			  ifFormat    : "%d/%m/%Y",    // the date format
			  button      : "trigger"       // ID of the button
			}
		  );
		</script>

    </td>
  </tr>
</table>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
user1.Close()
subjects.Close()
email_info.close()
call log_the_page ("Send Individual email")
%>


