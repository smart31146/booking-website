<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->

<%
' *** Edit Operations: declare variables

MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If


Dim email_sent_message
' query string to execute
MM_editQuery = ""
Dim array_of_subjects

' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then

	'get all users in this business division
	set business_division_users = Server.CreateObject("ADODB.Recordset")
	business_division_users.ActiveConnection = Connect
	business_division_users.Source = "SELECT * from q_user where user_info1="&Request("filter_info1")
	business_division_users.CursorType = 0
	business_division_users.CursorLocation = 3
	business_division_users.LockType = 3
	business_division_users.Open()
	business_division_users_numRows = 0
	While (NOT business_division_users.EOF)


				'get the array with the subjects ticked
				array_of_subjects = split(Request.Form ("subjects"), ",")

				'format the date
				uk_date=Request("send_after")
				sql_date= left(uk_date,2)&"-"& mid(uk_date,4,2) &"-"&  right(uk_date,4)

				MM_editConnection = Connect
				MM_editQuery = "insert into emails (q_user,date_to_send,type) values ('"&business_division_users.Fields.Item("ID_user").Value&"','"+sql_date+"' , 100)"
				'response.write ( MM_editQuery)
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
	business_division_users.MoveNext()
	Wend

    email_sent_message="<br/><br/>Those emails will be sent after "&uk_date&".  To send another please enter the details below"
	call log_the_page ("Sent Business Division email")
End If


%>
<%
function WA_VBreplace(thetext)
  if isNull(thetext) then thetext = ""
  newstring = Replace(cStr(thetext),"'","|WA|")
  newstring = Replace(newstring,"\","\\")
  WA_VBreplace = newstring
end function

set filter_info1 = Server.CreateObject("ADODB.Recordset")
filter_info1.ActiveConnection = Connect
filter_info1.Source = "SELECT * FROM q_info1 where info1_active=1 order by info1"
filter_info1.CursorType = 0
filter_info1.CursorLocation = 3
filter_info1.LockType = 3
filter_info1.Open()
filter_info1_numRows = 0

set subjects = Server.CreateObject("ADODB.Recordset")
subjects.ActiveConnection = Connect
subjects.Source = "SELECT ID_subject, subject_name FROM subjects where subject_active_q <> 0"
subjects.CursorType = 0
subjects.CursorLocation = 3
subjects.LockType = 3
subjects.Open()
subjects_numRows = 0
%>
<%
numbers=1
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz new user. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
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
	theme_advanced_disable : "cut,copy,paste, undo, redo,image,cleanup,help, code,removeformat"


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


	if (document.forms[0].filter_info1.selectedIndex==0)
	{
		alert("Sorry, you must enter a Business Division");
		return false;
	}
	else if (document.forms[0].send_after.value.length<2)
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
    <td align="left" valign="bottom" class="heading"> Send a reminder email to a business division <%=email_sent_message%></td>
  </tr>
  <tr>
    <td align="left" valign="bottom">
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="send_email" onSubmit="return trySubmit();" >
        <table>
          <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="120">Business division:</td>
            <td class="text" align="left" valign="top" colspan="3">


               <select name="filter_info1" class="formitem1" >
                <option value="0">--- select a business ---</option>
                <%
					While (NOT filter_info1.EOF)
					%>
									<option value="<%=(filter_info1.Fields.Item("ID_info1").Value)%>" <%if (CStr(filter_info1.Fields.Item("ID_info1").Value) = CStr(request.querystring("filter_info1"))) then Response.Write("SELECTED") : Response.Write("")%>><%=(filter_info1.Fields.Item("info1").Value)%></option>
									<%
					  filter_info1.MoveNext()
					Wend
					'If (filter_info1.CursorType > 0) Then
					'  filter_info1.MoveFirst
					'Else
					  filter_info1.Requery
					'End If
					%>
              </select>
            </td>
          </tr>
		  <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="120">Send after:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <input type="text" name="send_after"  id="send_after" size="20" value="" class="formitem1" readonly="true"><button id="trigger">...</button>

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
filter_info1.Close()
subjects.Close()
call log_the_page ("Send Business Division email")
%>


