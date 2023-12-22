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


//-->
</script>

<script language="JavaScript" type="text/javascript">
//window.parent.myLytebox.parent.document.location.href='a_omraden.asp?alt=visa';
window.parent.redirectFromLogin = true;
window.parent.myLytebox.doPageRefresh = true;
</script>

<script src="//code.jquery.com/jquery-2.1.3.js?v=bbp34"></script>
<!-- <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/js/bootstrap.min.js?v=bbp34"></script> -->
<script src="https://cdn.js?v=bbp34delivr.net/jquery.formvalidation/0.6.1/js/formValidation.min.js?v=bbp34"></script>
<!-- <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js?v=bbp34"></script> -->
<script src="//cdn.ckeditor.com/4.4.3/basic/ckeditor.js?v=bbp34"></script>
<script src="//cdn.ckeditor.com/4.4.3/basic/adapters/jquery.js?v=bbp34"></script>
</HEAD>
<BODY>
<% IF request.querystring("alt")= "image" THEN
set subjects = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM new_subjects WHERE s_id = "&fixstr(clng(request.querystring("s_id")))&""
subjects.Open SQL, Connect,3,3%>
<script src="ckeditor/ckfinder/ckfinder.js?v=bbp34" type="text/javascript"></script>
<script type="text/javascript">

function BrowseServer( startupPath, functionData )
{
	// You can use the "CKFinder" class to render CKFinder in a page:
	var finder = new CKFinder();

	// The path for the installation of CKFinder (default = "/ckfinder/").
	finder.basePath = '/ckfinder/';

	//Startup path in a form: "Type:/path/to/directory/"
	finder.startupPath = startupPath;

	// Name of a function which is called when a file is selected in CKFinder.
	finder.selectActionFunction = SetFileField;

	// Additional data to be passed to the selectActionFunction in a second argument.
	// We'll use this feature to pass the Id of a field that will be updated.
	finder.selectActionData = functionData;

	// Name of a function which is called when a thumbnail is selected in CKFinder.
	finder.selectThumbnailActionFunction = ShowThumbnails;

	// Launch CKFinder
	finder.popup();
}

// This is a sample function which is called when a file is selected in CKFinder.
function SetFileField( fileUrl, data )
{
	document.getElementById( data["selectActionData"] ).value = fileUrl;
}

// This is a sample function which is called when a thumbnail is selected in CKFinder.
function ShowThumbnails( fileUrl, data )
{
	// this = CKFinderAPI
	var sFileName = this.getSelectedFile().name;
	document.getElementById( 'thumbnails' ).innerHTML +=
			'<div class="thumb">' +
				'<img src="' + fileUrl + '" />' +
				'<div class="caption">' +
					'<a href="' + data["fileUrl"] + '" target="_blank">' + sFileName + '</a> (' + data["fileSize"] + 'KB)' +
				'</div>' +
			'</div>';

	document.getElementById( 'preview' ).style.display = "";
	// It is not required to return any value.
	// When false is returned, CKFinder will not close automatically.
	return false;
}
	</script>
	<br>
	
	<% IF subjects("s_image")<>"" THEN%>
	<img src="../vault_image/images/<% =subjects("s_image")%>" alt=""><br>
	<a href="q_training_lyte.asp?alt=imagedelete&s_id=<% =subjects("s_id")%>" onclick="return confirm('This image will now be deleted\nIt will not be erased from the server..\n\nAre you sure?')"   class="quiz_button" style="padding:1px 8px;text-decoration:none;">Delete image</A><br><br>
	<% END IF%>
	<form action="q_training_lyte.asp?alt=imagesave&s_id=<% =subjects("s_id")%>" method="post" name="orderform">
		<strong>Selected Image URL</strong><br/>
		<input id="xImagePath" name="ImagePath" type="text" size="60" />
		<input type="button" value="Browse Server" onclick="BrowseServer( 'Images:/', 'xImagePath' );" />
	<div id="preview" style="display:none">
		<strong>Selected Thumbnails</strong><br/>
		<div id="thumbnails"></div>
	</div><br>
	<input type="Submit" name="Submit2" value="Save image" class="quiz_button">
</form>

<% subjects.close
ELSEIF request.querystring("alt")= "imagesave" THEN

	sImage = request.form("ImagePath")
	sImage = Split(sImage,"/",7)
	'sImage = Split(sImage,"/") 
	For i=0 to UBound(sImage) 'the UBound function returns 3
	sImage2 = sImage(i)
	Next
	set UPDATE = Server.CreateObject("ADODB.Recordset")
	SQL = "UPDATE new_subjects SET s_image = '"&fixstr(sImage2)&"' WHERE s_id = "&fixstr(clng(request.querystring("s_id")))&""
	UPDATE.Open SQL, Connect,3,3
	response.redirect "q_training_lyte.asp?alt=image&s_id="&request.querystring("s_id")&""
	
ELSEIF request.querystring("alt")= "imagedelete" THEN
	set UPDATE = Server.CreateObject("ADODB.Recordset")
	SQL = "UPDATE new_subjects SET s_image = Null WHERE s_id = "&fixstr(clng(request.querystring("s_id")))&""
	UPDATE.Open SQL, Connect,3,3
	response.redirect "q_training_lyte.asp?alt=image&s_id="&request.querystring("s_id")&""%>
	
	
<% 
ELSEIF request.querystring("alt")= "editbox" THEN
set question = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM new_questions WHERE q_id = "&fixstr(clng(request.querystring("q_id")))&""
question.Open SQL, Connect,3,3%>
<span class="heading"> Edit information box</span><br><br>
<form action="q_training_lyte.asp?alt=editboxsave" method="post" name="orderform">
<input type="Hidden" name="q_id" value="<% =question("q_id")%>">
Title<br>
<input type="text" name="q_title" style="width:500px;" maxlength="249" class="formitem1" size="25" value="<% =question("q_title")%>"><br><br>
Information (When the user click on the Title)<br>
<textarea name="q_div_info"><% =question("q_div_info")%></textarea>
<script type="text/javascript">
			//<![CDATA[

				CKEDITOR.replace( 'q_div_info',
					{
					width: 500,
					height: 140
						});

			//]]>
			</script>

<br><input type="Submit" name="Submit2" value="Update information box" class="quiz_button">
</form><br><br>
<% question.close
ELSEIF request.querystring("alt")= "editboxsave" THEN
	set UPDATE = Server.CreateObject("ADODB.Recordset")
	SQL = "UPDATE new_questions SET q_title = '"&fixstr(request.form("q_title"))&"', q_div_info = '"&fixstr(request.form("q_div_info"))&"' WHERE q_id = "&fixstr(clng(request.form("q_id")))&""
	UPDATE.Open SQL, Connect,3,3%>
		<span class="heading">Information box updated</span><br><br>
		You can now close the window.	
		
<% ELSEIF request.querystring("alt")= "addquiz" THEN%>

<span class="heading"> Add quiz</span><br><br>
<script>

</script>
<script>

var counter= 1;
var limit= 5;

function addQuestion(divName){
     if (counter == limit)  {
          alert("You have reached the limit of adding " + counter + "options");
     }
     else {
		  var newDiv = document.createElement('div');
          newDiv.innerHTML = "<br><select name='choice_cor'><option value=" + '0' + "> - </option> <option value=" + '1' + "> Correct </option></select><input type='text' name='choice_label' maxlength=" + '49' +" style=" + 'width:50px;'+"> <input type='text' name='choice_body' maxlength=" + '999' +" style=" + 'width:450px;'+">";
          document.getElementById(divName).appendChild(newDiv);
          counter++;
     }
} 


function validate_form(){

var question_fb_cor = document.forms["orderform"]["question_body"].value
if (question_fb_cor== null || question_fb_cor == ""){
alert("Please enter question body");
return false;
}

var question_fb_cor = document.forms["orderform"]["question_fb_cor"].value
if (question_fb_cor== null || question_fb_cor == ""){
alert("Please insert correct answer");
return false;
}

var question_fb_inc = document.forms["orderform"]["question_fb_inc"].value
if (question_fb_inc == null || question_fb_inc == ""){
alert("Please insert incorrect answer");
return false;
}

var choice_cor_value = document.forms["orderform"]["choice_cor"].value
if (choice_cor_value == null || choice_cor_value == ""){
alert("Please choose options for answers");
return false;
}

var choice_cor_value = document.forms["orderform"]["choice_cor"].value
if (choice_cor_value != 1){
alert("Please choose correct options for answers");
return false;
}

$('#orderform')
        .formValidation({
            framework: 'bootstrap',
            excluded: [':disabled'],
            icon: {
                valid: 'glyphicon glyphicon-ok',
                invalid: 'glyphicon glyphicon-remove',
                validating: 'glyphicon glyphicon-refresh'
            },
            fields: {
                question_body: {
                    validators: {
                        notEmpty: {
                            message: 'The question body is required and cannot be empty'
                        }
                    },
					    callback: {
                            message: 'The question must be less than 200 characters long',
                            callback: function(value, validator, $field) {
                                if (value === '') {
                                    return true;
                                }
                                // Get the plain text without HTML
                                var div  = $('<div/>').html(value).get(0),
                                    text = div.textContent || div.innerText;

                                return text.length = 0;
                            }
                        }
                }
            }
        });
}


</script>

<form action="q_training_lyte.asp?alt=addquizsave&s_id=<% =request.querystring("s_id")%>" method="post" name="orderform">
Question<br>
<textarea id="question_body" name="question_body"></textarea>
<script type="text/javascript">
			//<![CDATA[

				CKEDITOR.replace( 'question_body',
					{
					width: 500,
					height: 140
						});

			//]]>
			</script>

<table>
<TR valign="top" >
<TD width="300">Correct answer
<textarea name="question_fb_cor" style="width:280px;height:50px;" rows="6" cols="80"></textarea></TD>
<TD>Incorrect answer
<textarea name="question_fb_inc" style="width:280px;height:50px;" rows="6" cols="80"></textarea></TD>
</TR>
</table>

Add Question & answer<br>
<div class="table_normal" id='questionTable'>
<div>
<select name="choice_cor">
<option value="0"> -
<option value="1"> Correct
</select>
<input type="Text" maxlength="49" style="width:50px;" value="" name="choice_label">
<input type="Text" maxlength="999" style="width:450px;" value="" name="choice_body">
</div>
</div>
<br>
<input type="button" class="quiz_button" value="Add another option" onClick="addQuestion('questionTable');">
<br><br><input type="Submit" name="Submit2" value="Add Quiz" class="quiz_button" onClick="return validate_form();">
</form><br><br>


<br><br>

<% ELSEIF request.querystring("alt")= "addquizsave" THEN

set INSERT = Server.CreateObject("ADODB.Recordset")
		SQL = "INSERT INTO q_question (question_ord,question_body,question_topic) VALUES ("
		SQL = SQL & " 999999,"
		SQL = SQL & " '"&fixstr(trim(Request.form("question_body")))&"',"
		SQL = SQL & " '"&fixstr(trim(Request.querystring("s_id")))&"'"
		SQL = SQL & " )"
		INSERT.Open SQL, Connect,3,3
		
		set subject = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT TOP 1 id_question FROM q_question ORDER BY id_question desc"
		subject.Open SQL, Connect,3,3
		id_question = subject(0)
		subject.close
		

For i=1 to Request.form("choice_cor").Count

set INSERT = Server.CreateObject("ADODB.Recordset")
	SQL = "INSERT INTO q_choice (choice_question,choice_label,choice_body,choice_cor) VALUES ("
	SQL = SQL & " "&fixstr(clng(id_question))&","
	SQL = SQL & " '"&fixstr(trim(Request.form("choice_label")(i)))&"',"
	SQL = SQL & " '"&fixstr(trim(Request.form("choice_body")(i)))&"',"
	SQL = SQL & " '"&fixstr(trim(Request.form("choice_cor")(i)))&"'"
	SQL = SQL & " )"
	INSERT.Open SQL, Connect,3,3
	
next		
		
	response.redirect "?alt=editquiz&id_question="&id_question&""
ELSEIF request.querystring("alt")= "editquiz" THEN

IF request.querystring("do")= "updatequiz" THEN
	set UPDATE = Server.CreateObject("ADODB.Recordset")
	SQL = "UPDATE q_question SET question_body = '"&fixstr(request.form("question_body"))&"',question_fb_cor = '"&fixstr(request.form("question_fb_cor"))&"',question_fb_inc = '"&fixstr(request.form("question_fb_inc"))&"' WHERE ID_question = "&fixstr(clng(request.form("ID_question")))&""
	UPDATE.Open SQL, Connect,3,3
	response.redirect "?alt=editquiz&ID_question="&request.form("ID_question")&""
END IF


IF request.querystring("do")= "deletequiz" THEN
	set UPDATE = Server.CreateObject("ADODB.Recordset")
	SQL = "UPDATE q_choice SET choice_active = 0 WHERE ID_choice = "&fixstr(clng(request.querystring("ID_choice")))&""

	UPDATE.Open SQL, Connect,3,3
	response.redirect "?alt=editquiz&ID_question="&request.querystring("ID_question")&""
END IF




IF request.querystring("do")= "updatequestions" THEN
	set UPDATE = Server.CreateObject("ADODB.Recordset")
	For each record in request.form("id_choice")
		SQL = "UPDATE q_choice SET choice_label = '"&fixstr(trim(request.form("choice_label"&record&"")))&"', choice_body = '"&fixstr(trim(request.form("choice_body"&record&"")))&"', choice_cor = "&fixstr(request.form("choice_cor"&record&""))&" WHERE ID_choice = "&fixstr(clng(record))&""	
		UPDATE.Open SQL, Connect,3,3
	next
	response.redirect "?alt=editquiz&ID_question="&request.form("ID_question")&""
END IF

IF request.querystring("do")= "addquestion" THEN
	set INSERT = Server.CreateObject("ADODB.Recordset")
	SQL = "INSERT INTO q_choice (choice_question,choice_label,choice_body,choice_cor) VALUES ("
	SQL = SQL & " "&fixstr(clng(Request.form("ID_question")))&","
	SQL = SQL & " '"&fixstr(trim(Request.form("choice_label")))&"',"
	SQL = SQL & " '"&fixstr(trim(Request.form("choice_body")))&"',"
	SQL = SQL & " '"&fixstr(trim(Request.form("choice_cor")))&"'"
	SQL = SQL & " )"
	INSERT.Open SQL, Connect,3,3
	response.redirect "?alt=editquiz&ID_question="&request.form("ID_question")&""
END IF

set question = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM q_question WHERE ID_question = "&fixstr(clng(request.querystring("ID_question")))&""
question.Open SQL, Connect,3,3



%>

<form action="q_training_lyte.asp?alt=editquiz&do=updatequiz" method="post" name="orderform">
<input type="Hidden" name="ID_question" value="<% =question("ID_question")%>">

Information<br>
<textarea name="question_body"><%=question("question_body")%></textarea>
<script type="text/javascript">
			//<![CDATA[

				CKEDITOR.replace( 'question_body',
					{
					width: 600,
					height: 80
						});

			//]]>
			</script>

<table>
<TR valign="top">
<TD width="300">Correct answer
<textarea name="question_fb_cor" style="width:280px;height:50px;" rows="6" cols="80"><%=question("question_fb_cor")%></textarea></TD>
<TD>Incorrect answer
<textarea name="question_fb_inc" style="width:280px;height:50px;" rows="6" cols="80"><%=question("question_fb_inc")%></textarea></TD>
</TR>
</table><input type="Submit" name="Submit2" value="Update quiz information" class="quiz_button">
</form><br>

 Questions & answers<br>
<form action="q_training_lyte.asp?alt=editquiz&do=updatequestions" method="post" name="sortingboxform">
<input type="Hidden" name="ID_question" value="<% =question("ID_question")%>">
			<table><%
						set qchoice = Server.CreateObject("ADODB.Recordset")
						SQL = "SELECT id_choice,choice_label,choice_body,choice_cor FROM q_choice WHERE choice_question = "&fixstr(clng(question("ID_question")))&" AND ABS(choice_active) = 1 ORDER BY choice_label"
						qchoice.Open SQL, Connect,3,3
						do until qchoice.eof
							 %>
							 <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')" >
								<td  align="center" width="30"><input type="Hidden" name="id_choice" value="<% =qchoice("id_choice")%>">
								<select name="choice_cor<% =qchoice("id_choice")%>">
								<option value="1" <% IF cbool(qchoice("choice_cor")) = True THEN response.write " SELECTED" %>> Correct
								<option value="0" <% IF cbool(qchoice("choice_cor")) = False THEN response.write " SELECTED" %>> -
								</select></td>
								<td  align="center" width="30"><input type="Text" maxlength="49" style="width:50px;" value="<% =qchoice("choice_label")%>" name="choice_label<% =qchoice("id_choice")%>"></td>
								<td  align="left"><input type="Text" maxlength="999" style="width:400px;" value="<% =qchoice("choice_body")%>" name="choice_body<% =qchoice("id_choice")%>"></td>
								<td  align="right"><a href="q_training_lyte.asp?alt=editquiz&do=deletequiz&ID_question=<% =question("ID_question")%>&id_choice=<% =qchoice("id_choice")%>"  class="quiz_button" style="padding:1px 8px;text-decoration:none;">Delete</A></TD>
							</TR>
						 <% qchoice.movenext
							 loop
							 qchoice.close%>
							</table><input type="Submit" name="Submit2" value="Update  Questions & answers" class="quiz_button"></form><br>
Add Question & answer<br>
<form action="q_training_lyte.asp?alt=editquiz&do=addquestion" method="post" name="sortingboxform">
<table>
<tr class="table_normal">
<td  align="left"><select name="choice_cor">
<option value="0"> -
<option value="1"> Correct
</select></td>
<td  align="left"><input type="Text" maxlength="49" style="width:50px;" value="" name="choice_label"></td>
<td  align="left"><input type="Text" maxlength="999" style="width:450px;" value="" name="choice_body"></td>
</TR>
</table>
<input type="Hidden" name="ID_question" value="<% =question("ID_question")%>">

<input type="Submit" name="Submit2" value="Add  Questions & answers" class="quiz_button">
</form>
<% question.close
ELSEIF request.querystring("alt")= "addbox" THEN%>
<span class="heading"> Add information box</span><br><br>
<form action="q_training_lyte.asp?alt=addboxsave&q_order=<% =request.querystring("q_order")%>&q_tID=<% =request.querystring("q_tID")%>" method="post" name="orderform">
Title<br>
<input type="text" name="q_title" style="width:500px;" maxlength="249" class="formitem1" size="25" value=""><br><br>
Information (When the user click on the Title)<br>
<textarea name="q_div_info"></textarea>
<script type="text/javascript">
			//<![CDATA[

				CKEDITOR.replace( 'q_div_info',
					{
					width: 500,
					height: 140
						});

			//]]>
			</script>

<br><input type="Submit" name="Submit2" value="Add information box" class="quiz_button">
</form><br><br>
<% ELSEIF request.querystring("alt")= "addboxsave" THEN
set INSERT = Server.CreateObject("ADODB.Recordset")
		SQL = "INSERT INTO new_questions (q_order,q_title,q_div_info,q_tID) VALUES ("
		SQL = SQL & " "&fixstr(clng(Request.querystring("q_order")))&","
		SQL = SQL & " '"&fixstr(trim(Request.form("q_title")))&"',"
		SQL = SQL & " '"&fixstr(trim(Request.form("q_div_info")))&"',"
		SQL = SQL & " "&fixstr(clng(Session("s_id")))&""
		SQL = SQL & " )"
		INSERT.Open SQL, Connect,3,3%>
		<span class="heading">Information box saved</span><br><br>
		You can now close the window.<%
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



