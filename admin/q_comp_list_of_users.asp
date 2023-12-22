<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%

results=request("results")
fromdate=request("fromdate")
fromdate=cdatesql(fromdate)
todate=request("todate")
if len(todate) < 12 and todate <> "" then
	todate=todate&" 23:59:59"
end if
todate=cdatesql(todate)

active = request("active")
mths = request("mths")
if mths="" then
	mths=0
end if
noquiz = request("noquiz")


set filter_info1 = Server.CreateObject("ADODB.Recordset")
filter_info1.ActiveConnection = Connect
filter_info1.Source = "SELECT * FROM q_info1 order by info1"
filter_info1.CursorType = 0
filter_info1.CursorLocation = 3
filter_info1.LockType = 3
filter_info1.Open()
filter_info1_numRows = 0

set filter_info3 = Server.CreateObject("ADODB.Recordset")
filter_info3.ActiveConnection = Connect
filter_info3.Source = "SELECT * FROM q_info3 order by info3"
filter_info3.CursorType = 0
filter_info3.CursorLocation = 3
filter_info3.LockType = 3
filter_info3.Open()
filter_info3_numRows = 0

set filter_info4 = Server.CreateObject("ADODB.Recordset")
filter_info4.ActiveConnection = Connect
filter_info4.Source = "SELECT * FROM q_info4 order by info4"
filter_info4.CursorType = 0
filter_info4.CursorLocation = 3
filter_info4.LockType = 3
filter_info4.Open()
filter_info4_numRows = 0

set subjects = Server.CreateObject("ADODB.Recordset")
subjects.ActiveConnection = Connect
subjects.Source = "SELECT ID_subject, subject_name FROM subjects where subject_active_q <> 0"
subjects.CursorType = 0
subjects.CursorLocation = 3
subjects.LockType = 3
subjects.Open()
subjects_numRows = 0

set admin_user = Server.CreateObject("ADODB.Recordset")
admin_user.ActiveConnection = Connect
admin_user.Source = "SELECT * FROM admin inner join q_info4 on admin.admin_info4=q_info4.id_info4 where admin.id_admin="&Session("MM_id_admin")&""
admin_user.CursorType = 0
admin_user.CursorLocation = 3
admin_user.LockType = 3
admin_user.Open()


%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz users. You are logged in as <%=Session("MM_Username_admin") %></TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function checkmths()
{
	//if (((document.filter_users.fromdate.value != "") && (document.filter_users.todate.value != "")) || ((document.filter_users.fromdate.value == "") && (document.filter_users.todate.value != "")) || ((document.filter_users.fromdate.value != "") && (document.filter_users.todate.value == "")))
	//{
	//	document.filter_users.mths.checked = false
	//	document.filter_users.mths.disabled=true;
	//	return;
	//}
	//else
	//{
	//	document.filter_users.mths.disabled=false;
	//	return;
//	}
}

var MyCookie = {
    Write:function(name,value,days) {
        var D = new Date();
        D.setTime(D.getTime()+86400000*days)
        document.cookie = escape(name)+"="+escape(value)+
            ((days == null)?"":(";expires="+D.toGMTString()))
        return (this.Read(name) == value);
    },
    Read:function(name) {
        var EN=escape(name)
        var F=' '+document.cookie+';', S=F.indexOf(' '+EN);
        return S==-1 ? null : unescape(F.substring(EN=S+EN.length+2,F.indexOf(';',EN)));
    }
}

function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}



function AddCookieId(cn,id) {
        MyCookie.Write(cn,id,7);
}

function DelCookieId(cn,id) {
        MyCookie.Write(cn,id,-1);
}


function filter_submit()
{
	//if (isNaN(document.filter_users.show_lines.value)){
		//alert('Invalid number');
		//document.filter_users.show_lines.focus();
		//return false;
	//}
	//else
	//{
		//AddCookieId("show_lines",document.filter_users.show_lines.value);
		//show_lines = MyCookie.Read('show_lines')
		document.forms[0].submit();
		return true;
	//}
}
function clearform()
{
document.forms[0].filter_username.value = "";
document.forms[0].todate.value = "";
document.forms[0].fromdate.value = "";
document.forms[0].results.selectedIndex = 0;
document.forms[0].subject.selectedIndex = 0;
document.forms[0].active.selectedIndex = 0;
document.forms[0].filter_info1.selectedIndex = 0;
document.forms[0].filter_info3.selectedIndex = 0;
document.forms[0].filter_info4.selectedIndex = 0;
//document.forms[0].show_lines.value = "25";
document.forms[0].submit();
}
//-->
</script>
</HEAD>

<BODY>
	<%
	if Request.Cookies("show_lines")<> "" then
		show_lines= cint(Request.Cookies("show_lines"))
	else
		show_lines=15
	end if
	%>
<table>
  <tr>
    <td align="left" valign="bottom" class="heading"> Quiz users - combined results</td>
  </tr>
  <tr>
    <td align="left" valign="bottom">
      <form name="filter_users" action="q_comp_list_of_users_results.asp">
     <input type="hidden" name="hiddenmths" value="false">
        <table>
          <tr>
            <td colspan="8" class="subheads" align="left" valign="top">Users:</td>
          </tr>
          <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">
            <td class="text" width="18"><img src="images/back.gif" width="18" height="14"></td>
            <td class="text" colspan="8"><a href="main.asp">...Home page
              </a> </td>
          </tr>
          <tr>
            <td class="subheads" colspan="9">Filter users by:</td>
          </tr>
          <tr class="table_normal">
          <td class="text" width="18">&nbsp;</td>
          <td class="text" valign="top" width="143">User Type:</td>
           <td class="text" valign="top" colspan="7">

           <select name="active" class="formitem1">
              <%if cstr(request("active"))="" or cstr(request("active")="2") then%>
              <option value="2" selected>All Users</option>
              <option value="1">Active Users</option>
              <option value="0">Inactive Users</option>
              <%else if cstr(request("active"))="1" then%>
              <option value="2" >All Users</option>
              <option value="1" selected>Active Users</option>
              <option value="0">Inactive Users</option>
              <%else if cstr(request("active"))="0" then%>
              <option value="2"> All Users</option>
              <option value="1">Active Users</option>
              <option value="0" selected>Inactive Users</option>
              <%end if
              end if
              end if
              %>
              </select>
              </td>
		</tr>
		<tr class="table_normal">
          <td class="text" width="18">&nbsp;</td>
          <td class="text" valign="top" width="143">Sessions between:</td>
           <td  valign="top" colspan="7">
           <input type="text" name="fromdate" maxlength="19" class="formitem1" onDblClick="this.value='<%=cDateSQL(Now()-1)%>'; " size="25" value="<%=fromdate%>" >&nbsp;(yyyy-mm-dd hh:mm:ss), doubleclick = TODAY - 1 day<br>
           &nbsp;&nbsp;&nbsp;&nbsp;and <br>
           <input type="text" name="todate" maxlength="19" class="formitem1" onDblClick="this.value='<%=cDateSQL(Now())%>'; " size="25" value="<%=todate%>">
              (yyyy-mm-dd hh:mm:ss), doubleclick = TODAY
              </td>
		</tr>

          <tr class="table_normal">
            <td class="text" width="18">&nbsp;</td>
            <td class="text" valign="top" width="143">First OR Last name:</td>
            <td class="text" valign="top" colspan="7">
				<input type="text" name="filter_username" value="<%=request.querystring("filter_username")%>" class="formitem1"></td><td><!--<a href="javascript:onclick=filter_users.submit();" target='_self'><img src="images/go.gif" border=0></a>-->
            </td>

          </tr>
          <tr class="table_normal">
          <td class="text" width="18">&nbsp;</td>
          <td class="text" valign="top" width="143">Results:</td>
           <td class="text" valign="top" colspan="7">
           <select name="results" class="formitem1">
              <%if cstr(request("results"))="" or cstr(request("results")="2") then%>
              <option value="2" selected>All Users</option>
              <option value="1">Passed</option>
              <option value="0">Failed</option>
              <%else if cstr(request("results"))="1" then%>
              <option value="2" >All Users</option>
              <option value="1" selected>Passed</option>
              <option value="0">Failed</option>
              <%else if cstr(request("results"))="0" then%>
              <option value="2">All Users</option>
              <option value="1">Passed</option>
              <option value="0" selected>Failed</option>
              <%end if
              end if
              end if
              %>
              </select>
            </td>
          </tr>

		<input type="hidden" name="subject" value="0"/>

		  <tr class="table_normal">
            <td class="text" width="18">&nbsp;</td>
            <td class="text" valign="top" width="143">Business:</td>
            <td class="text" valign="top" colspan="8">
              <select name="filter_info1" class="formitem1">
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
          <tr class="table_normal">
            <td class="text" width="18">&nbsp;</td>
            <td class="text" valign="top" width="143"><% =BBPinfo3 %>:</td>
            <td class="text" valign="top" colspan="8">
              <select name="filter_info3" class="formitem1">
                <option value="0">--- select a <% =BBPinfo3 %> ---</option>
                <%
While (NOT filter_info3.EOF)
%>
                <option value="<%=(filter_info3.Fields.Item("ID_info3").Value)%>" <%if (CStr(filter_info3.Fields.Item("ID_info3").Value) = CStr(request.querystring("filter_info3"))) then Response.Write("SELECTED") : Response.Write("")%>><%=(filter_info3.Fields.Item("info3").Value)%></option>
                <%
  filter_info3.MoveNext()
Wend
'If (filter_info3.CursorType > 0) Then
'  filter_info3.MoveFirst
'Else
  filter_info3.Requery
'End If
%>
              </select>
            </td>
          </tr>
          <tr class="table_normal">
            <td class="text" width="18">&nbsp;</td>
            <td class="text" valign="top" width="143">Company:</td>
            <td class="text" valign="top" colspan="8">
              <select name="filter_info4" class="formitem1">
                <% if admin_user.fields.item("info4_viewall").value=1 then%>
                	<option value="0">--- select a company ---</option>
                <% end if %>
                <%
While (NOT filter_info4.EOF)
	if admin_user.fields.item("info4_viewall").value=1 OR admin_user.fields.item("id_info4").value=filter_info4.fields.item("id_info4").value then
%>
                <option value="<%=(filter_info4.Fields.Item("ID_info4").Value)%>" <%if (CStr(filter_info4.Fields.Item("ID_info4").Value) = CStr(request.querystring("filter_info4"))) then Response.Write("SELECTED") : Response.Write("")%>><%=(filter_info4.Fields.Item("info4").Value)%></option>
                <%
	end if
	filter_info4.MoveNext()
Wend
  filter_info4.Requery

%>
              </select>
            </td>
          </tr>
        <tr class="table_normal">
          <td colspan="10" align="center"class="text">
            <input type="button" name="Submit" value="&gt;&gt;&gt; Filter users &lt;&lt;&lt;" class="quiz_button" onclick="return filter_submit();">
          </td>
		</tr>

		<tr class="table_normal">
                  <td colspan="10" class="text"> <br><BR>
                    <b><font color="red">PLEASE NOTE:</font></b> 
					There are a large number of quiz user records.  You should set the filter options above to return only the specific user records that you
					want. Performing an unfiltered query is not recommended due to the number of records that will be returned.
                  </td>
                </tr>
 </table>


</BODY>
</HTML>
<%
call log_the_page ("Quiz List Users")
Set users = Nothing
filter_info1.Close()
filter_info3.Close()
filter_info4.Close()
admin_user.Close()
subjects.Close()
%>

