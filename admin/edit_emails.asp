<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->

<%
' *** Edit Operations: declare variables

MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

set edit_emails = Server.CreateObject("ADODB.Recordset")
edit_emails.ActiveConnection = Connect
edit_emails.Source ="Select * from email_info where email_active=1 AND email_admin=0 ORDER BY email_order, email_name"
edit_emails.CursorType = 0
edit_emails.CursorLocation = 3
edit_emails.LockType = 3
edit_emails.Open()
edit_emails_numRows = 0


' query string to execute
MM_editQuery = ""

' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then
Dim header_content
header_content=replace(Request("email_header"),"'","''")
header_content=LEFT(header_content, (LEN(header_content)-45))
header_content=header_content '&"<tr><td><td style='padding:5px; font-size:10pt;color:#000;'><div style='color:#000000;font-family:Arial'>"
header_content=replace(header_content,"'","''")


  MM_editConnection = Connect
  MM_editQuery = "UPDATE email_info SET email_subject='"+replace(Request("email_subject"),"'","''")+"',email_period='"+replace(Request("email_period"),"'","''")+"' , email_from= '"+replace(Request("email_from"),"'","''")+"', email_header= '"+header_content+"', email_footer= '"+replace(Request("email_footer"),"'","''")+"', email_cc='"+replace(Request("email_cc"),"'","''")+"', email_bcc='"+replace(Request("email_bcc"),"'","''")+"' WHERE email_id="+request("email_select")+""
'response.write ( MM_editQuery)
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
   MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
    call log_the_page ("Updated email properties")
End If

'load the properties from the database
Dim auto_subject, auto_from, auto_header, auto_footer, auto_cc, auto_bcc

if request("email_select") <> "" AND request("email_select") <> "0" then
	set props = Server.CreateObject("ADODB.Recordset")
	props.ActiveConnection = Connect
	props.Source ="Select * from email_info where email_id="&request("email_select")
	props.CursorType = 0
	props.CursorLocation = 3
	props.LockType = 3
	props.Open()
	props_numRows = 0
	if (props.Fields.Item("email_subject").Value) <> "" then
		email_subject=props.Fields.Item("email_subject").Value
	end if

	if (props.Fields.Item("email_from").Value) <> "" then
		email_from=props.Fields.Item("email_from").Value
	end if

	if (props.Fields.Item("email_header").Value) <> "" then
		email_header=props.Fields.Item("email_header").Value
	end if

	if (props.Fields.Item("email_footer").Value) <> "" then
		email_footer=props.Fields.Item("email_footer").Value
	end if
	if (props.Fields.Item("email_cc").Value) <> "" then
		email_cc=props.Fields.Item("email_cc").Value
	end if
	if (props.Fields.Item("email_bcc").Value) <> "" then
		email_bcc=props.Fields.Item("email_bcc").Value
	end if
	if (props.Fields.Item("email_name").Value) <> "" then
		email_name=props.Fields.Item("email_name").Value
	end if
	if (props.Fields.Item("email_period").Value) <> "" then
		email_period=props.Fields.Item("email_period").Value
	end if
	props.Close()
end if
%>
<%
function WA_VBreplace(thetext)
  if isNull(thetext) then thetext = ""
  newstring = Replace(cStr(thetext),"'","|WA|")
  newstring = Replace(newstring,"\","\\")
  WA_VBreplace = newstring
end function


%>
<%
numbers=1
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate" />
<meta http-equiv="Pragma" content="no-cache" />
<meta http-equiv="Expires" content="0" />
<TITLE>BBP ADMIN: Quiz new user. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">

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


function emailCheck (emailStr) {
var emailPat=/^(.+)@(.+)$/
var specialChars="\\(\\)<>@,;:\\\\\\\"\\.\\[\\]"
var validChars="\[^\\s" + specialChars + "\]"
var quotedUser="(\"[^\"]*\")"
var ipDomainPat=/^\[(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\]$/
var atom=validChars + '+'
var word="(" + atom + "|" + quotedUser + ")"
var userPat=new RegExp("^" + word + "(\\." + word + ")*$")
var domainPat=new RegExp("^" + atom + "(\\." + atom +")*$")
var matchArray=emailStr.match(emailPat)
if (matchArray==null) {
	alert("Email address seems incorrect (check @ and .'s)")
	return false
}
var user=matchArray[1]
var domain=matchArray[2]
if (user.match(userPat)==null) {
    alert("The username doesn't seem to be valid.")
    return false
}
var IPArray=domain.match(ipDomainPat)
if (IPArray!=null) {
	  for (var i=1;i<=4;i++) {
	    if (IPArray[i]>255) {
	        alert("Destination IP address is invalid!")
		return false
	    }
    }
    return true
}
var domainArray=domain.match(domainPat)
if (domainArray==null) {
	alert("The domain name doesn't seem to be valid.")
    return false
}
var atomPat=new RegExp(atom,"g")
var domArr=domain.match(atomPat)
var len=domArr.length
if (domArr[domArr.length-1].length<2 ||
    domArr[domArr.length-1].length>3) {
   alert("The address must end in a three-letter domain, or two letter country.")
   return false
}
if (len<2) {
   var errStr="This address is missing a hostname!"
   alert(errStr)
   return false
}
return true;
}

function isEmail(str)
{
  var supported = 0;
  if (window.RegExp) {
    var tempStr = "a";
    var tempReg = new RegExp(tempStr);
    if (tempReg.test(tempStr)) supported = 1;
  }
  if (!supported)
    return (str.indexOf(".") > 2) && (str.indexOf("@") > 0);
  var r1 = new RegExp("(@.*@)|(\\.\\.)|(@\\.)|(^\\.)");
  var r2 = new RegExp("^.+\\@(\\[?)[a-zA-Z0-9\\-\\.]+\\.([a-zA-Z]{2,3}|[0-9]{1,3})(\\]?)$");
  return (!r1.test(str) && r2.test(str));
}



function trySubmit()
{


	if (document.email_edit.email_subject.value.length<2)
	{
		alert("Sorry, you must enter a subject!\n(min. 2 characters)");
		return false;
	}
	if (document.email_edit.email_subject.value.length>200)
	{
		alert("Sorry, you must enter a shorter subject!\n(max. 200 characters)");
		return false;
	}
	if (document.email_edit.email_header.value.length>8000)
	{
		alert("Sorry, you must enter a shorter header!\n(max. 8000 characters)");
		return false;
	}
	if (document.email_edit.email_footer.value.length>4000)
	{
		alert("Sorry, you must enter a shorter footer!\n(max. 4000 characters)");
		return false;
	}
	if (document.email_edit.email_from.value.length<2)
	{
		alert("Sorry, you must enter a from address!\n(min. 2 characters)");
		return false;
	}
	if (document.email_edit.email_cc.value.length>0)
	{
		return emailCheck (document.email_edit.email_cc.value);
	}
	if (document.email_edit.email_bcc.value.length>0)
	{
		return emailCheck (document.email_edit.email_bcc.value);
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
<script src="styles/lytebox.js?v=bbp34" type="text/javascript"></script>
<link rel="STYLESHEET" type="text/css" href="styles/lytebox.css">
<script type="text/javascript" src="ckeditor/ckeditor.js?v=bbp34"></script>
</HEAD>
<BODY>
<table>
	  <tr>
		  <td align="left" valign="bottom" class="heading"> Edit emails </td>
	  </tr>
	  <tr>
		  <td align="left" valign="bottom" class="subheads"> Select an email to edit:</td>
	  </tr>
	  <tr>
		  <td align="left" valign="bottom" >
		  <form name="select_email" method="get" action="edit_emails.asp">
		  <select name="email_select" onchange="document.select_email.submit();">
		  <option value="0">Select an email to edit</option>
		  <%
		  while not edit_emails.eof
			%><option value="<%=edit_emails.fields.item("email_id").value%>" <% if int(request("email_select"))=int(edit_emails.fields.item("email_id").value) then%>SELECTED<%end if%>><%=edit_emails.fields.item("email_name").value%></option><%
			edit_emails.movenext
		  Wend
		  %>
		  </select>
		  </form>
		  </td>
	  </tr>
		<% if request("email_select") <> "" and request("email_select") <> "0" then	%>
	  <tr>
		  <td align="left" valign="bottom" class="heading"> &nbsp;</td>
	  </tr>
  <tr>
    <td align="left" valign="bottom">
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="email_edit" onSubmit="return trySubmit();" >
        <table>
		  <tr>
			<td align="left" valign="bottom" class="subheads"> Edit <%=email_name%> Email</td>
		  </tr>
          <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="120">Subject:</td>
            <td class="text" align="left" valign="top" colspan="3">


              <input type="text" name="email_subject"  size="70" value="<%=email_subject%>" class="formitem1">
              <input type="hidden" name="email_from"  size="70" value="<%=email_from%>" class="formitem1">
            </td>
          </tr>
		  <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="120">From Address:</td>
            <td class="text" align="left" valign="top" colspan="3">
              <input type="text" name="auto_from"  size="70" value="<%=email_from%>" class="formitem1">
            </td>
          </tr>
		  <tr class="table_normal" >
			 <td class="text" align="left" valign="top" width="120">CC Address:</td>
			 <td class="text" align="left" valign="top" colspan="3">
			   <input type="text" name="email_cc"  size="70" value="<%=email_cc%>" class="formitem1">
			 </td>
          </tr>
          <tr class="table_normal" >
		      <td class="text" align="left" valign="top" width="120">BCC Address:</td>
		      <td class="text" align="left" valign="top" colspan="3">
		       <input type="text" name="email_bcc"  size="70" value="<%=email_bcc%>" class="formitem1">
		      </td>
          </tr>

          <tr>
            <td class="text" align="left" valign="top"></td>
            <td class="text" align="left" valign="top" colspan="3">
              <input type="hidden" name="active"  value="1" >
            </td>
          </tr>

		  <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="120">Header:</td>
            <td class="text" align="left" valign="top" colspan="3">
             <textarea name="email_header"  rows="10" wrap="on" cols="80" ><%=email_header%></textarea>
			 <script type="text/javascript">
			//<![CDATA[

				CKEDITOR.replace( 'email_header',
					{
					width: 900,
					height: 320
						});

			//]]>
			</script>

            </td>
          </tr>
		  <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="120">Footer:</td>
            <td class="text" align="left" valign="top" colspan="3">
            <textarea name="email_footer"  rows="10" wrap="on" cols="80"><%=email_footer%></textarea>
			<!--<script type="text/javascript">
			//<![CDATA[

				CKEDITOR.replace( 'email_footer',
					{
					width: 900,
					height: 120
						});

			//]]>
			</script>-->
            </td>
          </tr>
		  <tr class="table_normal" >
            <td class="text" align="left" valign="top" width="120">Timeframe (in days):</td>
            <td class="text" align="left" valign="top" colspan="3">
            <input type="text" name="email_period" size="4" value="<%=email_period%>"></input>
            </td>
          </tr>

          <tr>
            <td  align="left" valign="top">
              <input type="hidden" name="session" value="<%=getPassword(30, "", "true", "true", "true", "false", "true", "true", "true", "false")%>">
              <input type="hidden" name="current_export">
            </td>
            <td  align="left" valign="top" colspan="3">

              <input type="submit" name="Submit" value="Save" class="quiz_button" <%call IsEditOK%>>

            </td>
          </tr>
        </table>
        <input type="hidden" name="MM_insert" value="true">
      </form>
    </td>
  </tr>
  <% end if %>
</table>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
call log_the_page ("Edit auto email")
%>


