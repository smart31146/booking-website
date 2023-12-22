<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->


<html>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz subjects. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css"></script>
<script language="javascript" type="text/javascript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}

function trySubmit()
{
	if (document.orderform.s_title.value.length<2)
	{
		alert("Sorry, you must enter a name for the title!\n(min. 2 characters)");
		return false;
	}
	if (confirm("Are you sure you want to update?"))	{	document.orderform.submit();
	return false;
	}
return false;
}
function exitpage()
{
	if (change==true)
	{
		if (confirm("You have changed at least one field on this page.\nBefore exiting this page, do you want to save those changes first?"))
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


	
	<% IF request.querystring("alt")= "" THEN%>
	<BODY>
<span class="heading"> Add page</span><br><br>

	Remember. This page will be saved as the last page. You can change the sorting when you have saved the page.<br><br>
What type of page do you want to create?<br><br>


<a href="q_training_addpage.asp?alt=save&s_typ=1&s_order=<% =request.querystring("s_order")%>" onclick="return confirm('This page will now be created.\n\nAre you sure?')"  class="quiz_button" style="padding:1px 8px;text-decoration:none;">Training page</A>
<a href="q_training_addpage.asp?alt=save&s_typ=2&s_order=<% =request.querystring("s_order")%>" onclick="return confirm('This page will now be created.\n\nAre you sure?')" class="quiz_button" style="padding:1px 8px;text-decoration:none;margin-left:20px;">Quiz page</A>
<% ELSEIF request.querystring("alt")= "save" THEN
set INSERT = Server.CreateObject("ADODB.Recordset")
		SQL = "INSERT INTO new_subjects (s_order,s_typ,s_qID) VALUES ("
		SQL = SQL & " "&fixstr(clng(Request.querystring("s_order")))&","
		SQL = SQL & " "&fixstr(clng(Request.querystring("s_typ")))&","
		SQL = SQL & " "&fixstr(clng(Session("ID_subject")))&""
		SQL = SQL & " )"
		INSERT.Open SQL, Connect,3,3
		
		set subject = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT TOP 1 s_id FROM new_subjects ORDER BY s_id desc"
		subject.Open SQL, Connect,3,3
		s_id = subject(0)
		subject.close
		response.redirect "q_training_addpage.asp?alt=edit&s_id="&s_id&""


ELSEIF request.querystring("alt")= "edit" THEN%>
<BODY onload="change=false;" onUnload="return exitpage();"><%

IF request.querystring("do")= "deletebox" THEN
	set DELETE = Server.CreateObject("ADODB.Recordset")
	SQL = "DELETE FROM new_questions WHERE q_id = "&fixstr(clng(request.querystring("q_id")))&""
	DELETE.Open SQL, Connect,3,3
	response.redirect "?alt=edit"
END IF

IF request.querystring("do")= "sortingbox" THEN

	set UPDATE = Server.CreateObject("ADODB.Recordset")
	For each record in request.form("q_id")
		SQL = "UPDATE new_questions SET q_order = "&fixstr(clng(request.form("q_order"&record&"")))&" WHERE q_id = "&fixstr(clng(record))&""	
		UPDATE.Open SQL, Connect,3,3
	next
	response.redirect "?alt=edit"
END IF

IF request.querystring("do")= "editsave" THEN
set UPDATE = Server.CreateObject("ADODB.Recordset")
	SQL = "UPDATE new_subjects SET s_goback = "&fixstr(request.form("s_goback"))&", s_title = '"&fixstr(trim(request.form("s_title")))&"', s_body = '"&fixstr(trim(request.form("s_body")))&"' WHERE s_id = "&fixstr(clng(request.form("s_id")))&""
	UPDATE.Open SQL, Connect,3,3
	response.redirect "q_training.asp"
END IF

IF request.querystring("do")= "deletequestion" THEN
	set UPDATE = Server.CreateObject("ADODB.Recordset")
	SQL = "UPDATE q_question SET question_active = 0 WHERE ID_question = "&fixstr(clng(request.querystring("ID_question")))&""
	UPDATE.Open SQL, Connect,3,3
	response.redirect "?alt=edit"
END IF

IF request.querystring("do")= "activatequestion" THEN
	set UPDATE = Server.CreateObject("ADODB.Recordset")
	SQL = "UPDATE q_question SET question_active = 1 WHERE ID_question = "&fixstr(clng(request.querystring("ID_question")))&""
	UPDATE.Open SQL, Connect,3,3
	response.redirect "?alt=edit"
END IF



IF request.querystring("s_id")<>"" THEN
	Session("s_id")=request.querystring("s_id")
END IF
IF Session("s_id") = "" THEN response.redirect "q_list_of_subjects.asp"


set subject = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM new_subjects WHERE s_id = "&fixstr(clng(Session("s_id")))&""
subject.Open SQL, Connect,3,3




IF clng(subject("s_typ")) = 1 THEN

set question = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT q_id,q_title,q_div_info,q_order FROM new_questions WHERE q_tID = "&fixstr(clng(subject("s_id")))&" ORDER BY q_order"
question.Open SQL, Connect,3,3
if not question.eof then
	showQuestion = true
	QArr = question.GetRows 
ELSE
	QArr = 0
END IF
question.close %>



<span class="heading"> Edit page</span><br>
<form action="q_training_addpage.asp?alt=edit&do=editsave" method="post" name="orderform" onSubmit=" change=false; return trySubmit(0); ">
<table>
	<TR valign="top">
	<TD height="30" width="30"><img src="../admin/images/back.gif" width="18" height="14"></TD>
	<td><input type="button" name="goback" value="...Back to pages" class="quiz_button" onClick="document.location='q_training.asp'"></td>
	</TR>
</table>

<input type="Hidden" name="s_id" value="<% =subject("s_id")%>">
Title<br>
<textarea onChange="change=true;" name="s_title" rows="4" cols="5" style="width:500px;height:50px;"><%=trim(subject("s_title"))%></textarea><br><br>

Information<br>
<textarea onChange="change=true;" name="s_body"><%=trim(subject("s_body"))%></textarea>
<script type="text/javascript">
			//<![CDATA[

				CKEDITOR.replace( 's_body',
					{
					width: 500,
					height: 220
						});

			//]]>
			</script>

<br>
<% set goback = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT s_id,s_title,s_body,s_topic FROM new_subjects WHERE s_qID = "&fixstr(clng(Session("ID_subject")))&" AND s_typ = 1 AND s_active = 1 ORDER BY s_order"
goback.Open SQL, Connect,3,3
if not goback.eof then
	QArrGoback = goback.GetRows 
ELSE
	QArrGoback = 0
END IF
goback.close %>
<span class="heading"> Go back and see this scenario again</span><br>
<select onChange="change=true;" name="s_goback">
<option value="0">
<% For i=0 to ubound(QArrGoback,2) %>
<option value="<% =QArrGoback(0,i)%>"<% IF clng(subject("s_goback")) = clng(QArrGoback(0,i)) THEN response.write " SELECTED"%>> <% =QArrGoback(3,i)%> /  <% =QArrGoback(1,i)%> / <% =left(QArrGoback(2,i),15)%>
<% next%>
</select><br><br><input type="Submit" name="Submit2" value="Update page" class="quiz_button">
</form>

<br><br>
<span class="heading"> Information boxes</span><br>
<%
			q_order = 0
			IF showQuestion = true THEN %>
			<form action="q_training_addpage.asp?alt=edit&do=sortingbox" method="post" name="sortingboxform">
			<table><%
				If Ubound(QArr,2) > -1 Then
					 For i=0 to ubound(QArr,2) 
					 q_order = q_order+1%>
					 <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')" >
					 <td  align="center" width="50"><input type="hidden" value="<% =QArr(0,i)%>" name="q_id">
<select name="q_order<% =QArr(0,i)%>">
<% for s = 1 to 250%>
<option value="<% =s%>"<% IF clng(QArr(3,i)) = clng(s) THEN response.write " SELECTED"%>> <% =s%>
<% next%>
</select></td>
					 <td  align="left"><% =QArr(1,i)%></td>
					 <td  align="center" width="50"> <a href="q_training_lyte.asp?alt=editbox&q_id=<% =QArr(0,i)%>"  rel="lyteframe" rev="width: 650px; height: 550px; scrolling: auto;" title="" class="quiz_button" style="padding:1px 8px;text-decoration:none;">Edit</a><br></td>
					 <td  align="center" width="60"> <a href="q_training_addpage.asp?alt=edit&do=deletebox&q_id=<% =QArr(0,i)%>"  onclick="return confirm('This information box will now be deleted.\n\nAre you sure?')" class="quiz_button" style="padding:1px 8px;text-decoration:none;">Delete</a><br></td>
					 </TR>
			  		
			<%		Next
				END IF
				response.write "</table>"
			END IF%><br><input type="Submit" name="Submit2" value="Update order" class="quiz_button" style="cursor:pointer;"><a href="q_training_lyte.asp?alt=addbox&q_order=<% =q_order+1%>"  rel="lyteframe" rev="width: 650px; height: 550px; scrolling: auto;" title="" class="quiz_button" style="padding:1px 8px;text-decoration:none;margin-left:20px;">Add information box</A>
			</form><br>

<%
 subject.close

ELSEIF clng(subject("s_typ")) = 2 THEN


	showTopic = true
	set question = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT ID_question,question_body,question_active FROM q_question WHERE question_topic = "&fixstr(clng(subject("s_id")))&" ORDER BY question_active desc,ID_question"
	question.Open SQL, Connect,3,3
	if not question.eof then
	QArr = question.GetRows 
	ELSE
	showTopic = false
	END IF

%>
<table>
	<TR valign="top">
	<TD height="30"><img src="../admin/images/back.gif" width="18" height="14"></TD>
	<td><a href="q_training.asp">...Back to page list</a></td>
	</TR>
	</table>

<form name="quiz" method="POST" onsubmit="return trySubmit();" action="t_question.asp?currentID=<% =subject("s_id")%>&quiz=yes">
				<% IF showTopic = true THEN %>
				<%
					 For q=0 to ubound(QArr,2) %>
					 <hr width="800" align="left">
					 <table>
					 <TR>
					 <td  align="left" colspan="4"><% IF not cbool(QArr(2,q)) THEN%><strike><% END IF%><% =QArr(1,q)%></strike> <a href="q_training_lyte.asp?alt=editquiz&ID_question=<% =QArr(0,q)%>"  rel="lyteframe" rev="width: 650px; height: 550px; scrolling: auto;" title="" class="quiz_button" style="padding:1px 8px;text-decoration:none;">Edit quiz</A>
					<% IF not cbool(QArr(2,q)) THEN%> 
					<a href="q_training_addpage.asp?alt=edit&do=activatequestion&ID_question=<% =QArr(0,q)%>"  onclick="return confirm('This question will now be activated.\n\nAre you sure?')" class="quiz_button" style="padding:1px 8px;text-decoration:none;">Activate quiz</A>
					<% ELSE %>
					<a href="q_training_addpage.asp?alt=edit&do=deletequestion&ID_question=<% =QArr(0,q)%>"  onclick="return confirm('This question will now be deleted.\n\nAre you sure?')" class="quiz_button" style="padding:1px 8px;text-decoration:none;">Inactivate quiz</A>
					<% END IF%>
					</td>
					 </TR><%
						set qchoice = Server.CreateObject("ADODB.Recordset")
						SQL = "SELECT id_choice,choice_label,choice_body,choice_cor FROM q_choice WHERE choice_question = "&fixstr(clng(QArr(0,q)))&" AND ABS(choice_active) = 1 ORDER BY choice_label"
						qchoice.Open SQL, Connect,3,3
						do until qchoice.eof
							 %>
							 <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')" >
							<td  align="center" width="30"><% IF cbool(qchoice("choice_cor")) = True THEN %><img src="../images/icon_true.gif" alt=""><% END IF%>&nbsp;</td>
							<td  align="center" width="30"><% =qchoice("choice_label")%></td>
							<td  align="left"><% =qchoice("choice_body")%></td>
							<td  align="left"></td>
							</TR>
						 <% qchoice.movenext
							 loop
							 qchoice.close%>
							</table><br>
						 <%Next %>
							<%
END IF%>
						 <br><a href="q_training_lyte.asp?alt=addquiz&s_id=<% =request.querystring("s_id")%>"  rel="lyteframe" rev="width: 650px; height: 550px; scrolling: auto;" title="" class="quiz_button" style="padding:1px 8px;text-decoration:none;">Add quiz</A><%
						 END IF%>
<%
END IF%>
<p>&nbsp;</p>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
'call log_the_page ("Traning Add page")


Function gTyp(str)
	select case str
	case 1 : gTyp = "Training"
	case 2 : gTyp = "Quiz"
	end select
End Function

%>



