<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
numbers=1
%>
<%
set preferences = Server.CreateObject("ADODB.Recordset")
preferences.ActiveConnection = Connect
preferences.Source = "SELECT preferences.*  FROM preferences  ORDER BY preferences.pref_date DESC;"
preferences.CursorType = 0
preferences.CursorLocation = 3
preferences.LockType = 3
preferences.Open()
preferences_numRows = 0
%>
<%
set admins = Server.CreateObject("ADODB.Recordset")
admins.ActiveConnection = Connect
admins.Source = "SELECT * FROM admin"
admins.CursorType = 0
admins.CursorLocation = 3
admins.LockType = 3
admins.Open()
admins_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
preferences_numRows = preferences_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Repeat2__numRows = -1
Dim Repeat2__index
Repeat2__index = 0
admins_numRows = admins_numRows + Repeat2__numRows
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Preferences. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</HEAD>
<BODY>
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> Preferences</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <table>
        <tr> 
          <td colspan="6" class="subheads">Saved preferences profiles. The first 
            ACTIVE in the list is the EFFECTIVE one!</td>
        </tr>
        <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">         <td class="text" width="10"><img src="../admin/images/back.gif" width="18" height="14"></td>
          <td class="text" colspan="5"><a href="main.asp">...Home page </a></td>
        </tr>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT preferences.EOF)) 
%>
        <% If Not preferences.EOF Or Not preferences.BOF Then %>
        <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')" > 
          <td class="text" width="20"><%=numbers%></td>
          <td class="text" width="540"><a href="../admin/preferences_edit.asp?pid=<%=(preferences.Fields.Item("ID_pref").Value)%>"><%=(preferences.Fields.Item("pref_name").Value)%></a> </td>
          <td class="text" width="20"> 
            <%if abs(preferences.Fields.Item("pref_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
          </td>
          <td class="text" width="20"> 
            <%if Edit_OK then %>
            <a href="preferences_del.asp?pid=<%=(preferences.Fields.Item("ID_pref").Value)%>" onClick="javascript: return (confirm('You are just about to delete this preference.\nAre you sure you want to do that?'));"><img src="images/bin.gif" width="16" height="16" border="0"></a> 
            <% end if %>
          </td>
        </tr>
        <% End If ' end Not preferences.EOF Or NOT preferences.BOF 
		%>
        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  preferences.MoveNext()
  numbers=numbers+1
Wend
%>
        <% If preferences.EOF And preferences.BOF Then %>
        <tr> 
          <td class="text">&nbsp;</td>
          <td width="99%"  colspan="5">Sorry, 
            there are no saved preferences profiles currently.</td>
        </tr>
        <% End If ' end preferences.EOF And preferences.BOF 
		%>
        <tr> 
          <td class="text"><img src="images/new2.gif" width="11" height="13"></td>
          <td width="99%" class="text" colspan="5"> 
            <input type="button" name="Button" value="Add a new preference profile" onClick="document.location='preferences_add.asp';" class="quiz_button">
          </td>
        </tr>
		<tr>
          <td class="text"><img src="images/home.gif" width="16" height="16"></td>
          <td width="99%" class="text" colspan="5"><a href="javascript:"  onClick="MM_openBrWindow('edit_welcome.asp','WelcomeMessage','scrollbars=yes,width=610,height=400')">Welcome Message</a></td>
        </tr>

		<tr>
        <td class="text"><img src="images/search.gif" width="16" height="16"></td>
		<td class="text">
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
<a style="cursor: pointer;" value="Browse Server" onclick="BrowseServer( 'Images:/', 'xImagePath' );" />Browse Server</a>
</td>
</tr>
		
		<tr> 
          <td class="text">&nbsp; 
            <%
numbers=1
%>
          </td>
          <td width="99%" class="text" colspan="5">&nbsp;</td>
        </tr>
        <tr> 
          <td class="subheads" colspan="6">Administrators and reviewers</td>
        </tr>
        <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')">        <td class="text" width="10"><img src="../admin/images/back.gif" width="18" height="14"></td>
          <td class="text" colspan="5"><a href="main.asp">...Home page </a></td>
        </tr>
        <% 
While ((Repeat2__numRows <> 0) AND (NOT admins.EOF)) 
%>
        <% If Not admins.EOF Or Not admins.BOF Then %>
        <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
          <td class="text" width="20"><%=numbers%></td>
          <td class="text" width="540"><a href="preferences_edit_admin.asp?admin_id=<%=(admins.Fields.Item("id_admin").Value)%>"><%=(admins.Fields.Item("admin_name").Value)%></a> </td>
          <td class="text" width="20"> 
            <%if abs(admins.Fields.Item("admin_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
          </td>
          <td class="text" width="20"> 
            <% if lCase(admins.Fields.Item("admin_level").Value) = "admin" then response.write "<img src='images/admin.gif'>" else response.write "<img src='images/review.gif'>"%>
          </td>
        </tr>
        <% End If ' end Not admins.EOF Or NOT admins.BOF 
		%>
        <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  admins.MoveNext()
  numbers=numbers+1
Wend
%>
        <% If admins.EOF And admins.BOF Then %>
        <tr> 
          <td class="text">&nbsp;</td>
          <td width="99%"  colspan="5">Sorry, 
            there are no administrators or reviewers of this application.</td>
        </tr>
        <% End If ' end admins.EOF And admins.BOF 
		%>
        <tr> 
          <td class="text"><img src="images/new2.gif" width="11" height="13"></td>
          <td width="99%" class="text" colspan="5"> 
            <input type="button" name="Button2" value="Add a new administrator or reviewer" onClick="document.location='preferences_add_admin.asp';" class="quiz_button">
          </td>
        </tr>
        <tr> 
          <td class="text">&nbsp;</td>
          <td width="99%" class="text" colspan="5">&nbsp;</td>
        </tr>
      </table>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
</BODY>
</HTML>

<%
call log_the_page ("BBG List Preferences")
%>

<%
preferences.Close()
%>
<%
admins.Close()
%>

