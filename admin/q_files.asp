<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/include_admin.asp" -->
<% 
'response.write Request.ServerVariables("PATH_TRANSLATED")
'response.end
Set objFSO2 = Server.CreateObject("Scripting.FileSystemObject")
'Set objFolder = objFSO2.GetFolder("D:\webhotel\LOTJ\bbp_chh_bsg\vault\files\")
Set objFolder = objFSO2.GetFolder(Request.ServerVariables("APPL_PHYSICAL_PATH")&"vault_image\files\")

IF request.querystring("alt")= "delete" THEN
	 objFSO2.DeleteFile(objFolder&"\"&request.querystring("url"))
	response.redirect "q_files.asp"
END IF

%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>Quiz Upload users</TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
</HEAD>
<BODY BGCOLOR=#FFCC00 TEXT=#000000 VLINK=#000000 LINK=#000000 leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="3" cellpadding="0">
  <tr> 
    <td align="left" valign="bottom" class="headers"> Files</td>
  </tr>
  <tr> 
    <td align="left" valign="middle" class="text">
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
	<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
//-->
</script>
	<br>
	

<a href="q_files.asp" class="quiz_button" style="padding:1px 8px;text-decoration:none;">Refresh</A><br><br>
Right click on the file you want to save and choose "Save target as..."<br><br>

<table cellpadding="5" cellspacing="0" border="0" width="700" class="text_table">
<tr class="table_normal">
<td class="text"><b>File name</b></td>
<td class="text"><b>Size</b></td>
<td class="text"><b>Last changed</b></td>
<td class="text">&nbsp;</td>
</tr>
<% 
For Each objFile in objFolder.Files %>
<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">
<td class="text"><a target="_blank" href="../vault_image/files/<% =objFile.Name%>" target="_blank"><% =objFile.Name %></a>&nbsp;&nbsp;</td>
<td class="text" width="100"><nobr><% =formatnumber(objFile.Size/1000,0)%> kB</td>
<td class="text" width="170"><nobr><% =objFile.DateLastModified %></td>
<td class="text" width="50" align="center"><nobr><a href="q_files.asp?alt=delete&amp;url=<% =objFile.Name%>" onclick="return confirm('This file will now be deleted.\n\nAre you sure?')" class="quiz_button" style="padding:1px 8px;text-decoration:none;">Delete</A></td>
</tr>
<%Next
Set objFolder = Nothing
Set objFile = Nothing
Set objFSO2 = Nothing
%>
</table><br>
<br><br>Make sure to include the subject, version number and path reference in the document title. eg: Ethics v2 path A.<br><br>

Once you have uploaded a document, you will need to refresh the page to see the file.<br>

<br>


	<form action="q_files.asp" method="post" name="orderform">
		<input id="xImagePath" name="ImagePath" style="display:none;" type="text" size="60" />
		<input type="button" value="Upload file" onclick="BrowseServer( 'Files:/', 'xFilePath' );" />
	<div id="preview" style="display:none">
		<strong>Selected Thumbnails</strong><br/>
		<div id="thumbnails"></div>
	</div><br>
	<input type="Submit" name="Submit2" style="display:none;" value="Upload" class="quiz_button">
</form><br><br>
      </td>
  </tr>
</table>
</BODY>
</HTML>

