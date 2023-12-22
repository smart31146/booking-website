<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/include_admin.asp" -->
<HTML>
<HEAD>
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}

function trySubmit()
{
if (document.forms[0].file.value == '') { alert ("You must select a file first...");
return false;
} else {
return true;
}
}

function checkFileUpload(form,extensions,requireUpload,sizeLimit,minWidth,minHeight,maxWidth,maxHeight,saveWidth,saveHeight) { //v2.08
  document.MM_returnValue = true;
  for (var i = 0; i<form.elements.length; i++) {
    field = form.elements[i];
    if (field.type.toUpperCase() != 'FILE') continue;
    checkOneFileUpload(field,extensions,requireUpload,sizeLimit,minWidth,minHeight,maxWidth,maxHeight,saveWidth,saveHeight);
} }

function checkOneFileUpload(field,extensions,requireUpload,sizeLimit,minWidth,minHeight,maxWidth,maxHeight,saveWidth,saveHeight) { //v2.08
  document.MM_returnValue = true;
  if (extensions != '') var re = new RegExp("\.(" + extensions.replace(/,/gi,"|").replace(/s/gi,"") + ")$","i");
    if (field.value == '') {
      if (requireUpload) {alert('File is required!');document.MM_returnValue = false;field.focus();return;}
    } else {
      if(extensions != '' && !re.test(field.value)) {
        alert('This file type is not allowed for uploading.\nOnly the following file extensions are allowed: ' + extensions + '.\nPlease select another file and try again.');
        document.MM_returnValue = false;field.focus();return;
      }
    document.PU_uploadForm = field.form;
    re = new RegExp(".(gif|jpg|png|bmp|jpeg)$","i");
    if(re.test(field.value) && (sizeLimit != '' || minWidth != '' || minHeight != '' || maxWidth != '' || maxHeight != '' || saveWidth != '' || saveHeight != '')) {
      checkImageDimensions(field,sizeLimit,minWidth,minHeight,maxWidth,maxHeight,saveWidth,saveHeight);
    } }
}

function showImageDimensions(fieldImg) { //v2.08
  var isNS6 = (!document.all && document.getElementById ? true : false);
  var img = (fieldImg && !isNS6 ? fieldImg : this);
  if ((img.minWidth != '' && img.minWidth > img.width) || (img.minHeight != '' && img.minHeight > img.height)) {
    alert('Uploaded Image is too small!\nShould be at least ' + img.minWidth + ' x ' + img.minHeight); return;}
  if ((img.maxWidth != '' && img.width > img.maxWidth) || (img.maxHeight != '' && img.height > img.maxHeight)) {
    alert('Uploaded Image is too big!\nShould be max ' + img.maxWidth + ' x ' + img.maxHeight); return;}
  if (img.sizeLimit != '' && img.fileSize > img.sizeLimit) {
    alert('Uploaded Image File Size is too big!\nShould be max ' + (img.sizeLimit/1024) + ' KBytes'); return;}
  if (img.saveWidth != '') document.PU_uploadForm[img.saveWidth].value = img.width;
  if (img.saveHeight != '') document.PU_uploadForm[img.saveHeight].value = img.height;
  document.MM_returnValue = true;
}

function checkImageDimensions(field,sizeL,minW,minH,maxW,maxH,saveW,saveH) { //v2.08
  if (!document.layers) {
    var isNS6 = (!document.all && document.getElementById ? true : false);
    document.MM_returnValue = false; var imgURL = 'file:///' + field.value.replace(/\\/gi,'/');
    if (!field.gp_img || (field.gp_img && field.gp_img.src != imgURL) || isNS6) {field.gp_img = new Image();
		   with (field) {gp_img.sizeLimit = sizeL*1024; gp_img.minWidth = minW; gp_img.minHeight = minH; gp_img.maxWidth = maxW; gp_img.maxHeight = maxH;
  	   gp_img.saveWidth = saveW; gp_img.saveHeight = saveH; gp_img.onload = showImageDimensions; gp_img.src = imgURL; }
	 } else showImageDimensions(field.gp_img);}
}
//-->
</script>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Icon manager. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
</HEAD>
<BODY BGCOLOR=#FFCC00 TEXT=#000000 VLINK=#000000 LINK=#000000 leftmargin="5" topmargin="5" onLoad="self.focus();">
<form name="upload" enctype="multipart/form-data" method="post" action="_chili_soft_upload.asp?path=client/bbg_icons&redirect=_ico_browse.asp" onSubmit="<%call on_form_Submit(0)%>;checkFileUpload(this,'',true,10,'','','','','','')">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td align="left" valign="bottom" class="heading"> List of available icons...</td>
    </tr>
    <tr> 
      <td align="left" valign="bottom"> 
        <%
Dim strPath
Dim objFSO
Dim objFolder
Dim objItem
Dim formfldname

if request.querystring("formfldname") <> "" then formfldname = request.querystring("formfldname") else formfldname = "p_icon"

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
strPath = "../client/bbg_icons/"
	if cInt(ScriptEngineMinorVersion) >= Good_MS_VBScript_version then
		Set objFolder = objFSO.GetFolder(Server.MapPath(strPath))
	else
		Set objFolder = objFSO.GetFolder((strPath))
	end if

%>
        <table border="0" cellspacing="0" cellpadding="3" class="table_normal" width="100%">
          <%
if objFolder.Files.count > 0 then 
line_num =1
For Each objItem In objFolder.Files
	f_filename = objItem.Name
if line_num=1 then%>
          <tr align="center" valign="top"> 
            <%end if%>
            <td onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"><a href="#" onClick='opener.document.forms[0].<%=formfldname%>.value = "<%=f_filename%>"; window.close("icobrowse");'><img src="../client/bbg_icons/<%=f_filename%>" border="0"></a><br>
              <a href="#" onClick='opener.document.forms[0].p_icon.value = "<%=f_filename%>"; window.close("icobrowse");' class="text_table"><%=f_filename%></a> 
              <%if line_num=1 then%>
          </tr>
          <%end if
if line_num = 8 then line_num =1 else line_num = line_num +1
Next
else
response.write("<br><br><p class='text_table'>There are no files in the folder currently...</p>")
end if 

Set objItem = Nothing
Set objFolder = Nothing
Set objFSO = Nothing
%>
        </table>
      </td>
    </tr>
    <tr> 
      <td align="left" valign="bottom" class="heading" height="50">You can upload 
        a new icon:</td>
    </tr>
    <tr> 
      <td align="left" valign="bottom" class="text">1. Locate the image on your 
        harddrive (only GIF, JPEG or PNG, 58x58 points!!!)</td>
    </tr>
    <tr> 
      <td align="left" valign="bottom" class="text"> 2. 
        <input type="file" name="file" size="50" onChange="checkOneFileUpload(this,'GIF,JPG,JPEG,BMP,PNG',true,10,50,50,58,58,'','')" class="formitem1">
      </td>
    </tr>
    <tr> 
      <td align="left" valign="bottom" class="text">3. 
        <input type="submit" name="Submit" value="Upload the image" class="quiz_button" <%call IsEditOK%>>
        or 
        <input type="button" name="close" value="Close this window" class="quiz_button" onClick="window.close()">
      </td>
    </tr>
    <tr> 
      <td align="left" valign="bottom" class="text">&nbsp;</td>
    </tr>
  </table>
</form>
</BODY>
</HTML>

<%
call log_the_page ("BBG Icons browse")
%>
