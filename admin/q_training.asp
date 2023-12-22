<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->

<%
IF request.querystring("ID_subject")<>"" THEN
	Session("ID_subject")=request.querystring("ID_subject")
END IF

IF request.querystring("alt")= "updatesorting" THEN
	set UPDATE = Server.CreateObject("ADODB.Recordset")
	For each record in request.form("s_id")
		SQL = "UPDATE new_subjects SET s_order = "&fixstr(clng(request.form("s_order"&record&"")))&",s_topic = '"&fixstr(trim(request.form("s_topic"&record&"")))&"'  WHERE s_id = "&fixstr(clng(record))&""	
		UPDATE.Open SQL, Connect,3,3
	next
	response.redirect "?"
END IF

IF request.querystring("alt")= "reorder" THEN
set subject = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM new_subjects s1,subjects WHERE s1.s_qID = "&fixstr(clng(Session("ID_subject")))&" AND s1.s_qiD = ID_subject AND s_active = 1 ORDER BY s1.s_order"
subject.Open SQL, Connect,3,3
x=0
do until subject.eof 
x=x+1
	set UPDATE = Server.CreateObject("ADODB.Recordset")
	SQL = "UPDATE new_subjects SET s_order = "&fixstr(clng(x))&" WHERE s_id = "&fixstr(clng(subject("s_id")))&""	
	UPDATE.Open SQL, Connect,3,3
subject.movenext
loop
subject.close
response.redirect "?alt=edit"
END IF


IF request.querystring("alt")= "deletepage" THEN

	set UPDATE = Server.CreateObject("ADODB.Recordset")
	SQL = "UPDATE new_subjects SET s_active = 0 WHERE s_id = "&fixstr(clng(request.querystring("s_id")))&""	
	UPDATE.Open SQL, Connect,3,3
	response.redirect "?alt=edit"
END IF

%>
<script language="javascript" type="text/javascript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
//-->
</script>
<script src="styles/lytebox.js?v=bbp34" type="text/javascript"></script>
<link rel="STYLESHEET" type="text/css" href="styles/lytebox.css">

<html>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz subjects. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
</script>
</HEAD>
<BODY>
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> Edit pages</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
	<form action="q_training.asp?alt=updatesorting" method="post" name="orderform">
	<table>
	<TR valign="top">
	<TD height="30"><img src="../admin/images/back.gif" width="18" height="14"></TD>
	<td colspan="4"><a href="q_list_of_subjects.asp">...Back to subject list</a></td>
	</TR>
	</table><% 
set subject = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT *,(SELECT Count(id_question) FROM q_question WHERE s_id = question_topic) qCounter FROM new_subjects s1,subjects WHERE s1.s_qID = "&fixstr(clng(Session("ID_subject")))&" AND s1.s_qiD = ID_subject AND s_active = 1 ORDER BY s1.s_order"
subject.Open SQL, Connect,3,3
sAntal = subject.RecordCOunt
IF subject.eof then
antals = 0
else%>

	<input type="Submit" name="Submit2" value="Update order & topic" class="quiz_button" style="cursor:pointer;">
<a href="q_training.asp?alt=reorder"  class="quiz_button" style="padding:1px 8px;text-decoration:none;margin-left:20px;">Re-order pages</A>
<%
antals = subject.RecordCount
END IF
%>
<a href="q_training_addpage.asp?s_order=<% =sAntal+2%>"  class="quiz_button" style="padding:1px 8px;text-decoration:none;margin-left:100px;">Add page</A><br>
<br>
<table>
      <% 
do until subject.eof
s_order = subject("s_order")
%>
<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')" >
<td  align="center" width="50"><input type="hidden" value="<% =subject("s_id")%>" name="s_id"><select name="s_order<% =subject("s_id")%>">
<% for s = 1 to 250%>
<option value="<% =s%>"<% IF clng(subject("s_order")) = s THEN response.write " SELECTED"%>> <% =s%>
<% next%>
</select></td>
<td  align="left" width="130"><input type="text" style="width:120px;" value="<% =subject("s_topic")%>" name="s_topic<% =subject("s_id")%>"></td>
<td  align="left" width="70"><% =gTyp(subject("s_typ"))%> </td>
<td  align="left"><% =subject("s_title")%><% IF clng(subject("qCounter"))>0 THEN response.write " "& subject("qCounter") & " random questions"%></td>
<td  align="left"><% IF len(subject("s_body")) > 0 THEN response.write left(replace(subject("s_body"),"<li>",""),35) & ".."%></td>
<td  align="center" width="60"><% IF clng(subject("s_typ")) = 1 THEN%>
<a href="q_training_addpage.asp?alt=edit&s_id=<% =subject("s_id")%>"  class="quiz_button" style="padding:1px 8px;text-decoration:none;">Edit</A>
	<% ELSE%>
<a href="q_training_addpage.asp?alt=edit&s_id=<% =subject("s_id")%>"  class="quiz_button" style="padding:1px 8px;text-decoration:none;">View</A>
<% END IF%></td>
<td  align="center"><% IF subject("s_image") <> "" THEN%><img src="images/icon_jpg.gif" width="25" height="18" alt=""><% END IF%>&nbsp;</td>
<td  align="center" width="60"><% IF clng(subject("s_typ")) = 1 THEN%><a href="q_training_lyte.asp?alt=image&s_id=<% =subject("s_id")%>"   rel="lyteframe" rev="width: 650px; height: 570px; scrolling: auto;" title="" class="quiz_button" style="padding:1px 8px;text-decoration:none;">Image</A><% END IF%>&nbsp;</td>
<td  align="right" width="140">
<a href="q_training.asp?alt=deletepage&s_id=<% =subject("s_id")%>" onclick="return confirm('This page will now be deleted.\n\nAre you sure?')"   class="quiz_button" style="padding:1px 8px;text-decoration:none;">Deactivate</A></td>
</TR>
<% subject.movenext
loop%></table><br>
</form>
<p>&nbsp;</p>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
'call log_the_page ("Quiz List Subjects")
subject.close

Function gTyp(str)
	select case str
	case 1 : gTyp = "Training"
	case 2 : gTyp = "Quiz"
	end select
End Function

%>


