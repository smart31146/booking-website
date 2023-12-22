<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/include_admin.asp" -->
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>Quiz Upload users</TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
</HEAD>
<BODY>
<table>
  <tr> 
    <td class="heading"> Upload users</td>
  </tr>
  <tr> 
    <td align="left" valign="middle" class="text">
	<% if request.querystring("alt")="" THEN%>
 <form method="post" name="frmSubject" enctype="multipart/form-data" action="?alt=upload&user=<%=user%>">	  

		<table>
		<TR valign="bottom">
			<TD class="text">Please select the excel file on your hard drive that you want to upload.<br><br>
			Select file<br>
			<input type="file" name="importfil" id="importfil" value="" class="formitem1" style="width:400px;" size="60">
			<TD align="right"><input type="Submit" name="bSubmit" id="bSubmit" value="Upload users" class="quiz_button"></TD>
		</TR>
		</table>
		
      </form><br>
	  <br>
	  Use this file when you want to upload users <img src="images/xls.gif" style="vertical-align:middle;" alt=""> <a href="">Download excel-file &raquo;</a><br>
	  <br>
	  <br>
	  <strong>SUGGESTION</strong>
	  
	  <table>
          <tr>
            <td >First name</td>
            <td >Last name</td>
            <td >Username</td>
            <td >Status</td>
            <td >Workplace of Respect date</td>
            <td >Path</td>
            <td >User Score</td>
            <td >Total score</td>
            <td >Business</td>
            <td >Site</td>
            <td >Activity/Role</td>
            <td >Team Leader's Email</td>
            <td >SendEmail</td>
          </tr>
		  <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">

            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:80px;" size="20"></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:80px;" size="20"></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:80px;" size="20"></td>
            <td class="text"><select name="">
			<option value="">Online
			<option value="">Offline
			</select></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:80px;" size="20"></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:30px;" size="20"></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:30px;" size="20"></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:30px;" size="20"></td>
            <td class="text"><select name="">
			<option value="">Business
			</select></td>
            <td class="text"><select name="">
			<option value="">Site
			</select></td>
            <td class="text"><select name="">
			<option value="">Activity/Role
			</select></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:80px;" size="20"></td>
            <td class="text"><select name="">
			<option value="">Yes
			<option value="">No
			</select></td>
			</TR>
		  <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">

            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:80px;" size="20"></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:80px;" size="20"></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:80px;" size="20"></td>
            <td class="text"><select name="">
			<option value="">Online
			<option value="">Offline
			</select></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:80px;" size="20"></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:30px;" size="20"></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:30px;" size="20"></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:30px;" size="20"></td>
            <td class="text"><select name="">
			<option value="">Business
			</select></td>
            <td class="text"><select name="">
			<option value="">Site
			</select></td>
            <td class="text"><select name="">
			<option value="">Activity/Role
			</select></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:80px;" size="20"></td>
            <td class="text"><select name="">
			<option value="">Yes
			<option value="">No
			</select></td>
			</TR>
			
		  <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">

            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:80px;" size="20"></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:80px;" size="20"></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:80px;" size="20"></td>
            <td class="text"><select name="">
			<option value="">Online
			<option value="">Offline
			</select></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:80px;" size="20"></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:30px;" size="20"></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:30px;" size="20"></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:30px;" size="20"></td>
            <td class="text"><select name="">
			<option value="">Business
			</select></td>
            <td class="text"><select name="">
			<option value="">Site
			</select></td>
            <td class="text"><select name="">
			<option value="">Activity/Role
			</select></td>
            <td class="text"><input type="text" name=""  value="" class="formitem1" style="width:80px;" size="20"></td>
            <td class="text"><select name="">
			<option value="">Yes
			<option value="">No
			</select></td>
			</TR>
	  </table>
	  <% elseif request.querystring("alt")="upload" THEN
	  

'Set Upload = Server.CreateObject("Persits.Upload")
'Count = Upload.Save("d:\webhotel")

'Response.Write Count & " file(s) uploaded to c:\upload"
'response.end

Set Upload = Server.CreateObject("Persits.Upload")
path = "D:\webhotel\"
' Undvik felmeddelande från ASPUpload när sidan laddas första gången
'Upload.IgnoreNoPost = True

Upload.SetMaxSize 10000000, true
' Save to memory. Path parameter is omitted
Upload.Save
For Each File in Upload.Files
'Set File = Upload.Files(1)
	arrFilename = split(lcase(File.ExtractFileName),".")
	FileEnd = arrFilename(ubound(arrFilename))
	IF FileEnd = "xlsx" THEN
		
		randomize
		random_number=int(rnd*999)+1
		asptime = hour(now) & Minute(now) & Second(now)
		File.SaveAs path&random_number&asptime&".xlsx"
		Set objConn = Server.CreateObject("ADODB.Connection")
		strCnxn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="& path&random_number&asptime&".xlsx;Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
		objConn.Open strCnxn

		Set objRS = Server.CreateObject("ADODB.Recordset")
		objRS.ActiveConnection = objConn
		objRS.CursorType = 3 'Static cursor.
		objRS.LockType = 2 'Pessimistic Lock.

		while not objRS.EOF
			response.write "'"& objRS(0) &"'<br>"& vbCrLf
			objRS.MoveNext
		wend
		objRS.Close : Set objRS = Nothing

		'For each record in request.form("email_id")
		'	SQL = "INSERT INTO nyhetsbrev_email (namn,email_email,email_bef,gruppid) SELECT namn,email_email,email_bef,"&request.querystring("gruppid")&" WHERE email_id = "&record
		'	response.write sql & " test<br>"& vbCrLf
		'	'Connect.Execute SQL,,128
		'next
	elseif FileEnd = "xls" then
		randomize
		random_number=int(rnd*999)+1
		asptime = hour(now) & Minute(now) & Second(now)
		File.SaveAs path&random_number&asptime&".xls"
		
		Set objConn = Server.CreateObject("ADODB.Connection")
		With objConn
		.Provider = "Microsoft.Jet.OLEDB.4.0"
		.ConnectionString = _
		"Data Source="& path&random_number&asptime &".xls;" & _
		"Extended Properties=Excel 8.0;"
		.CursorLocation = 3
		.Open
		End With

		Set objRS = Server.CreateObject("ADODB.Recordset")
		objRS.ActiveConnection = objConn
		objRS.CursorType = 3 'Static cursor.
		objRS.LockType = 2 'Pessimistic Lock.

		sql = "SELECT * FROM [BSG$]"
		objRS.Source = sql
		objRS.Open
		For Each fld in objRS.Fields
			Response.Write fld.Name &"<br>"& vbCrLf
		Next
		objRS.Close : Set objRS = Nothing

		'File.Delete
end if
next
%>
	  
	  
	  <% END IF%>
      </td>
  </tr>
</table>
</BODY>
</HTML>

