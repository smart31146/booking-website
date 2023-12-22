<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
Dim field
If (Request.QueryString("field") <> "") Then 
field = cStr(Request.QueryString("field"))
Else 
Response.Redirect("error.asp?" & request.QueryString) 
End If
%>
<script language="JavaScript">
<!--
  var errorString = "Sorry but this web page needs\nWindows95 and Internet Explorer 5 or above to view.";
  var Ok = "false";
  var name =  navigator.appName;
  var version =  parseFloat(navigator.appVersion);
  var platform = navigator.platform;

	if (platform == "Win32" && name == "Microsoft Internet Explorer" && version >= 4){
		Ok = "true";
	} else {
		Ok= "false";
	}

	if (Ok == "false") {
		alert(errorString);
	}

function initToolBar(ed) {
    
	var eb = document.all.editbar;
	if (ed!=null) {
		eb._editor = window.frames['myEditor'];
	}
}

function doFormat(what) {

	var eb = document.all.editbar;
		
	if(what == "FontName"){
		if(arguments[1] != 1){
			eb._editor.execCommand(what, arguments[1]);
			document.all.font.selectedIndex = 0;
		} 
	} else if(what == "FontSize"){
    if(arguments[1] != 1){
      eb._editor.execCommand(what, arguments[1]);
      document.all.size.selectedIndex = 0;
    } 
	} else {
	   eb._editor.execCommand(what, arguments[1]);
	}
}

function swapMode() {

	var eb = document.all.editbar._editor;
  eb.swapModes();
}

function makeUrl(){

	sUrl = document.all.what.value + document.all.url.value;
	doFormat('CreateLink',sUrl);
}

function copyValue() {

	var theHtml = "" + document.frames("myEditor").document.frames("textEdit").document.body.innerHTML + "";

	re = /<p>/gi;
   	theHtml = theHtml.replace(re, "");

	re = /<\/?p>/gi;
   	theHtml = theHtml.replace(re, "<br><br>");

	re = /<strong>/gi;
   	theHtml = theHtml.replace(re, "<b>");

	re = /<\/?strong>/gi;
   	theHtml = theHtml.replace(re, "</b>");

	re = /<em>/gi;
   	theHtml = theHtml.replace(re, "<i>");

	re = /<\/?em>/gi;
   	theHtml = theHtml.replace(re, "</i>");
	
	document.all.EditorValue.value = theHtml;
	opener.document.all.<%=field%>.value = theHtml;
	window.opener.change = true;
}

function SwapView_OnClick(){

  if(document.all.btnSwapView.value == "Show Html"){
		var sMes = "Show Wysiwyg";
    var sStatusBarMes = "Current View Html";
	} else {
		var sMes = "Show Html"
    var sStatusBarMes = "Current View Wysiwyg";
  }
	
	document.all.btnSwapView.value = sMes;
  window.status  = sStatusBarMes;
	swapMode();
}

function OnFormSubmit(){

    copyValue();
	window.close();
}
//-->
</script>
<html>
<head>
<title>BBP ADMIN: HTML WYSIWYG Editor. You are logged in as <%=Session("MM_Username_admin")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#FFCC33" onLoad="self.focus(); document.all.EditorValue.value = opener.document.all.<%=field%>.value; document.all.editbar._editor.initEditor();">
<form name="upravform">
        
  <table width="500" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr align="left" valign="top"> 
      <td colspan="2"> 
        <table cellspacing="0" cellpadding="0" width="100%" bordercolor="" border="0">
          <tr valign="top"> 
            <td> 
              <textarea name="EditorValue" style="display: none"></textarea>
            </td>
          </tr>
          <tr valign="top"> 
            <td> 
              <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
                <tr valign="top"> 
                  <td valign="top"> 
                    <div id=editbar > 
                      <table width="100%" border="0" cellpadding="0" cellspacing="0" align="left">
                        <tr> 
                          <td> 
                            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                              <tr> 
                                <td> <a href="javascript:"><img class='clsCursor' src="editor_images/cut.gif" width="16" height="16" border="0" onClick="doFormat('Cut')"></a> 
                                  <a href="javascript:"><img class='clsCursor' src="editor_images/copy.gif" width="16" height="16" border="0" onClick="doFormat('Copy')"></a>&nbsp<a href="javascript:"><img class='clsCursor' src="editor_images/paste.gif" border="0" onClick="doFormat('Paste')" width="16" height="16"></a> 
                                  <a href="javascript:"><img class='clsCursor' src="editor_images/para_bul.gif" width="16" height="16" border="0" onClick="doFormat('InsertUnorderedList');" ></a>&nbsp<a href="javascript:"><img class='clsCursor' src="editor_images/para_num.gif" width="16" height="16" border="0" onClick="doFormat('InsertOrderedList');" ></a>&nbsp<a href="javascript:"><img class='clsCursor' src="editor_images/indent.gif" width="20" height="16" onClick="doFormat('Indent')" border="0"></a>&nbsp<a href="javascript:"><img class='clsCursor' src="editor_images/outdent.gif" width="20" height="16" onClick="doFormat('Outdent')" border="0"></a> 
                                  <a href="javascript:"><img class='clsCursor' src="editor_images/bold.gif" width="16" height="16" border="0" onClick="doFormat('Bold')"></a> 
                                  <a href="javascript:"><img class='clsCursor' src="editor_images/italics.gif" width="16" height="16" border="0" onClick="doFormat('Italic')"></a> 
                                  <a href="javascript:"><img class='clsCursor' src="editor_images/underline.gif" width="16" height="16" border="0" onClick="doFormat('Underline')" ></a> 
                                  <a href="javascript:"><img class='clsCursor' src="editor_images/left.gif" width="16" height="16" border="0"  onClick="doFormat('JustifyLeft')"></a> 
                                  <a href="javascript:"><img class='clsCursor' src="editor_images/centre.gif" width="16" height="16" border="0" onClick="doFormat('JustifyCenter')"></a>&nbsp<a href="javascript:"><img class='clsCursor' src="editor_images/right.gif" width="16" height="16" border="0"  onClick="doFormat('JustifyRight')"></a> 
                                </td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                        <tr> 
                          <td> 
                            <table border="0" cellpadding="0" cellspacing="0" width="100%">
                              <tr> 
                                <td> 
                                  <select name="what" class="formitem1">
				    <option value="" selected>BBG direct link</option>
                                    <option value="http://">http://</option>
                                    <option value="mailto:">mailto:</option>
                                    <option value="ftp://">ftp://</option>
                                    <option value="https://">https://</option>
                                  </select>
                                  <input type="text" name="url" size="45" class="formitem1">
                                </td>
                                <td align="right"> 
                                  <input type="button" name="button2" value="Insert URL" onClick="makeUrl();" class="quiz_button">
                                </td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </table>
                    </div>
                  </td>
                </tr>
                <tr valign="top" align="left"> 
                  <td valign="top"> 
                    <table width="100%" border="0" height="100%">
                      <tr valign="top"> 
                        <td width="100%" height="100%"> 
                          <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
                            <tr valign="top"> 
                              <td bgcolor="#FFFFFF" height="300"><iframe id=myEditor src="pd_edit.htm" onFocus="initToolBar(this)" width=100% height=100%></iframe></td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <script>
  initToolBar("foo");
  window.status  = "Current View: Wysiwyg";
</script>
        </table>
      </td>
    </tr>
    <tr align="left" valign="top"> 
      <td > 
        <input type="button" name="Button" value="Transfer this code back to the form"  onClick="OnFormSubmit();" class="quiz_button">
        or <input type="button" name="Close" value="Close this window" class="quiz_button" onClick="window.close()"></td>
      <td align="right" > 
        <input type="button" name="btnSwapView" value="Show Html" onClick="SwapView_OnClick();" class="quiz_button">
      </td>
    </tr>
    <tr align="left" valign="top"> 
      <td colspan="2"  align="center">&nbsp;</td>
    </tr>
    <tr align="left" valign="top"> 
      <td colspan="2"  align="center">If you need to include 
        a link to a particular page within a BBG, use this <a href="javascript:" onClick="MM_openBrWindow('_link_generator.asp','linkgenerator','width=600,height=300')">LINK 
        GENERATOR </a></td>
    </tr>
  </table>
      </form>
</body>
</html>
<%
call log_the_page ("WYSIVYG editor")
%>
