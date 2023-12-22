<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/include_admin.asp" -->
<!--#include file="../scriptlibrary/incpureupload.asp" -->
<%
'*** Pure ASP File Upload -----------------------------------------------------
' Copyright (c) 2001-2002 George Petrov, www.UDzone.com
' Process the upload
' Version: 2.0.8
'------------------------------------------------------------------------------
'*** File Upload to: """../client/uploads""", Extensions: "TXT", Form: add_user, Redirect: "q_user_import2.asp", "file", "1000", "over", "true", "", "" , "", "", "", "", "600", "showprogress.htm", "300", "100"

Dim GP_redirectPage, RequestBin, UploadQueryString, GP_uploadAction, UploadRequest
PureUploadSetup

If (CStr(Request.QueryString("GP_upload")) <> "") Then
  on error resume next
  Dim reqPureUploadVersion, foundPureUploadVersion
  reqPureUploadVersion = 2.08
  foundPureUploadVersion = getPureUploadVersion()
  if err or reqPureUploadVersion > foundPureUploadVersion then
    Response.Write "<b>You don't have latest version of ScriptLibrary/incPureUpload.asp uploaded on the server.</b><br>"
    Response.Write "This library is required for the current page. It is fully backwards compatible so old pages will work as well.<br>"
    Response.End    
  end if
  on error goto 0
  GP_redirectPage = "q_user_import2.asp"
  Server.ScriptTimeout = 600
  
  RequestBin = Request.BinaryRead(Request.TotalBytes)
  Set UploadRequest = CreateObject("Scripting.Dictionary")  
  BuildUploadRequest RequestBin, """../client/uploads""", "file", "1000", "over"
  
  If (GP_redirectPage <> "" and not (CStr(UploadFormRequest("MM_insert")) <> "" or CStr(UploadFormRequest("MM_update")) <> "")) Then
    If (InStr(1, GP_redirectPage, "?", vbTextCompare) = 0 And UploadQueryString <> "") Then
      GP_redirectPage = GP_redirectPage & "?" & UploadQueryString
    End If
    Response.Redirect(GP_redirectPage)  
  end if  
else
  if UploadQueryString <> "" then
    UploadQueryString = UploadQueryString & "&GP_upload=true"
  else  
    UploadQueryString = "GP_upload=true"
  end if  
end if
' End Pure Upload
'------------------------------------------------------------------------------
%>
<HTML>
<HEAD>
<script language="JavaScript">
<!--

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

function showProgressWindow(progressFile,popWidth,popHeight) { //v2.08
  var showProgress = false, form, field;
  for (var f = 0; f<document.forms.length; f++) {
    form = document.forms[f];
    for (var i = 0; i<form.elements.length; i++) {
      field = form.elements[i];
      if (field.type.toUpperCase() != 'FILE') continue;
      if (field.value != '') {showProgress = true;break;}
  } }
  if (showProgress && document.MM_returnValue) {
    var w = 480, h = 340;
    if (document.all || document.layers || document.getElementById) {
      w = screen.availWidth; h = screen.availHeight;}
    var leftPos = (w-popWidth)/2, topPos = (h-popHeight)/2;
    document.progressWindow = window.open(progressFile,'ProgressWindow','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,width=' + popWidth + ',height='+popHeight);
    document.progressWindow.moveTo(leftPos, topPos);document.progressWindow.focus();
		window.onunload = function () {document.progressWindow.close();};
} }
//-->
</script>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz import users. You are logged in as </TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
</HEAD>
<BODY>
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> Quiz new users import</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form name="upload" onSubmit="<%call on_form_Submit(-1)%> return document.MM_returnValue;checkFileUpload(this,'TXT',true,1000,'','','','','','');showProgressWindow('showprogress.htm',300,100)" enctype="multipart/form-data" action="_chili_soft_upload.asp?path=client/uploads&redirect=q_user_import2.asp" method="post">

<!--<form name="upload" enctype="multipart/form-data" method="post" action="_chili_soft_upload.asp?path=client/uploads&redirect=q_user_import.asp" onSubmit="<%call on_form_Submit(0)%>;checkFileUpload(this,'',true,10,'','','','','','')">-->

        <table>
          <tr align="left" valign="top"> 
            <td width="20">1.</td>
            <td>Choose a text file with users*</td>
          </tr>
          <tr align="left" valign="top"> 
            <td width="20">&nbsp;</td>
            <td> 
              <input type="file" name="upload" class="quiz_button" onChange="checkOneFileUpload(this,'TXT',true,1000,'','','','','','')">
            </td>
          </tr>
          <tr align="left" valign="top"> 
            <td width="20">2.</td>
            <td>Upload the file to import users to a database</td>
          </tr>
          <tr align="left" valign="top"> 
            <td width="20">&nbsp;</td>
            <td> 
              <input type="submit" name="Submit" value="Upload above selected text file" class="quiz_button" <%call IsEditOK%>>
            </td>
          </tr>
          <tr align="left" valign="top"> 
            <td width="20">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr align="left" valign="top"> 
            <td width="20">*</td>
            <td> 
              <p>The format of the import text file MUST be as follows, otherwise 
                errors will occur.<br>
                a) Pure ASCII TXT file - no MS Word, No MS Excel, No HTML, just 
                ASCII TXT file!<br>
                b) Each user is on a new line ended by return &lt;crlf&gt;<br>
                c) ';' is used to separate each column<br>
                d) To enable user accounts overwriting insert line with ###overwrite### 
                to the upload file<br>
                e) Every entry after this line will be overwritten</p>
            </td>
          </tr>
          <tr align="left" valign="top"> 
            <td width="20">&nbsp;</td>
            <td> 
              <p><b><u>Column structure:<br>
                </u> </b> First name; Last name; Login name; Password; Month of 
                birth; City of birth; e-mail&lt;crlf&gt;</p>
              <p><b><u>Example:<br>
                </u> </b> <i>--------------------UPLOAD.TXT---------BEGIN-------------------------------------------------<br>
                John; Mc Smith; JMSmith; Smithspassword; 11; Melbourne; j.m.smith@domain.com.au</i><br>
                <i>Paul Patrick; Brown; PPBROWN; Mywifename; 6; Sri Lanka; paul_brown@domain.com.au<br>
                ###overwrite###<br>
                Bob; Black; BBlack; balckpassword; 11; Paris; bob.black@domain.com.au<br>
                George Petrov; petrovg; petrovpwd; 2; Moscow; gpetrov@domain.com.au<br>
                --------------------UPLOAD.TXT---------END---------------------------------------------------- 
                </i></p>
              <p><i>Above example will result in importing John and Paul. If either 
                username JMSmith or PPBrown will exist or any other combination 
                of John's or Patrick's initial and personal information, such 
                user will NOT be imported.<br>
                However Bob and George will be either inserted as new users or 
                their accounts will be updated in case the database contains any 
                of those user accounts already.</i></p>
              <p><i>Tip: create your upload files either so that new users will 
                be before ###overwrite### tag and the rest of the company will 
                be after this tag or for administrator's convenience use the ###overwrite### 
                tag at the top of the file so all accounts will be updated or 
                inserted as new for newly added users.</i></p>
              <p><i>Please bear in mind, that upload DOES NOT DELETE users!!! 
                They have to be deleted manually!</i></p>
            </td>
          </tr>
          <tr align="left" valign="top"> 
            <td width="20">&nbsp;</td>
            <td> 
              <input type="button" name="goback" value="Go back to user list" class="quiz_button" onClick="document.location='q_list_of_users.asp'">
            </td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
call log_the_page ("Quiz User import - file selection")
%>
