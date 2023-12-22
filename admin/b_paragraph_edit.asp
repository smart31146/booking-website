<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
' *** Edit Operations: declare variables

MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = Connect
  MM_editTable = "b_pages"
  MM_editColumn = "ID_page"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = ""
  MM_fieldsStr  = "topic|value|p_title|value|p_header|value|p_body|value|p_icon|value|active|value"
  MM_columnsStr = "page_topic|none,none,NULL|page_title|',none,''|page_header|',none,''|page_text|',none,''|page_icon|',none,''|page_active|none,1,0"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(i+1) = CStr(Request.Form(MM_fields(i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Update Record: construct a sql update staatement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update staatement
  MM_editQuery = "update " & MM_editTable & " set "
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    FormVal = MM_fields(i+1)
    MM_typeArray = Split(MM_columns(i+1),",")
    Delim = MM_typeArray(0)
    If (Delim = "none") Then Delim = ""
    AltVal = MM_typeArray(1)
    If (AltVal = "none") Then AltVal = ""
    EmptyVal = MM_typeArray(2)
    If (EmptyVal = "none") Then EmptyVal = ""
    If (FormVal = "") Then
      FormVal = EmptyVal
    Else
      If (AltVal <> "") Then
        FormVal = AltVal
      ElseIf (Delim = "'") Then  ' escape quotes
        FormVal = "'" & Replace(FormVal,"'","''") & "'"
      Else
        FormVal = Delim + FormVal + Delim
      End If
    End If
    If (i <> LBound(MM_fields)) Then
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(i) & " = " & FormVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    if Edit_OK = true then MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
    call log_the_page ("BBG Execute - UPDATE Paragraph: " & MM_recordId)	
	'make sure redirect called

	
	'now we redirect back to the list of questions page. but 1st we need to know what subj and topic this question belongs to
	set qlistinfo = Server.CreateObject("ADODB.Recordset")
	qlistinfo.ActiveConnection = Connect
	qlistinfo.Source = "SELECT b_topics.topic_subject AS subject, b_pages.page_topic AS topic FROM b_pages INNER JOIN b_topics ON b_pages.page_topic = b_topics.ID_topic WHERE b_pages.ID_page = " & Request.QueryString("pid")
	qlistinfo.CursorType = 0
	qlistinfo.CursorLocation = 3
	qlistinfo.LockType = 3
	qlistinfo.Open()

	While (NOT qlistinfo.EOF)
		MM_editRedirectUrl = "b_list_of_paragraphs.asp?subj=" & (qlistinfo.Fields.Item("subject").Value) & "&topic=" & (qlistinfo.Fields.Item("topic").Value) & "&highlight_q=" & Request.QueryString("pid")
		qlistinfo.MoveNext()
	Wend
	
	'response.write("redirect: '"&MM_editRedirectUrl&"'")
	'response.end	
    Response.Redirect(MM_editRedirectUrl)

  End If

End If
%>

<%
Dim pid
If (Request.QueryString("pid") <> "") Then 
pid = cInt(Request.QueryString("pid"))
Else 
Response.Redirect("error.asp?" & request.QueryString) 
End If
%>
<%
set paragraph = Server.CreateObject("ADODB.Recordset")
paragraph.ActiveConnection = Connect
paragraph.Source = "SELECT b_pages.*, subjects.ID_subject  FROM subjects INNER JOIN (b_topics INNER JOIN b_pages ON b_topics.ID_topic = b_pages.page_topic) ON subjects.ID_subject = b_topics.topic_subject  WHERE (((b_pages.ID_page)=" + Replace(pid, "'", "''") + "));"
paragraph.CursorType = 0
paragraph.CursorLocation = 3
paragraph.LockType = 3
paragraph.Open()
paragraph_numRows = 0
%>
<%
set subjects = Server.CreateObject("ADODB.Recordset")
subjects.ActiveConnection = Connect
subjects.Source = "SELECT ID_subject, subject_name FROM subjects"
subjects.CursorType = 0
subjects.CursorLocation = 3
subjects.LockType = 3
subjects.Open()
subjects_numRows = 0
%>
<%
set topics = Server.CreateObject("ADODB.Recordset")
topics.ActiveConnection = Connect
topics.Source = "SELECT *  FROM b_topics"
topics.CursorType = 0
topics.CursorLocation = 3
topics.LockType = 3
topics.Open()
topics_numRows = 0
%>
<%
function WA_VBreplace(thetext)
  if isNull(thetext) then thetext = ""
  newstring = Replace(cStr(thetext),"'","|WA|")
  newstring = Replace(newstring,"\","\\")
  WA_VBreplace = newstring
end function

if (NOT topics.EOF)     THEN

  Response.Write("<SC" & "RIPT>"&chr(10))
  Response.Write("var WAJA = new Array();"&chr(10))

  oldmainid = 0
  newmainid = topics.Fields("topic_subject").value
  if (oldmainid = newmainid)    THEN
    oldmainid = ""
  END IF
  n = 0
    while (NOT topics.EOF)
    if (oldmainid <> newmainid)     THEN
      Response.Write("WAJA[" & n & "] = new Array();"&chr(10))
      Response.Write("WAJA[" & n & "][0] = '" & WA_VBreplace(newmainid) & "';"&chr(10))
      m = 1
    END IF

    Response.Write("WAJA[" & n & "][" & m & "] = new Array();"&chr(10))
    Response.Write("WAJA[" & n & "][" & m & "][0] = " & "'" & WA_VBreplace(topics.Fields("ID_topic").value) & "'" & ";" &chr(10))
    Response.Write("WAJA[" & n & "][" & m & "][1] = " & "'" & WA_VBreplace(topics.Fields("topic_name").value) & "'" & ";" &chr(10))
    m=m+1
    if (cStr(oldmainid) = "0")      THEN
      oldmainid = newmainid
    END IF
    oldmainid = newmainid
    topics.MoveNext()
    if (NOT topics.EOF)     THEN
      newmainid = topics.Fields("topic_subject").value
    END IF
    if (oldmainid <> newmainid)     THEN
      n=n+1
    END IF
  WEND

  Response.Write("var topics_WAJA = WAJA;"&chr(10))
  Response.Write("WAJA = null;"&chr(10))
  Response.Write("</SC" & "RIPT>"&chr(10))
END IF
if (NOT topics.BOF)     THEN
  topics.MoveFirst()
END IF
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Reference edit paragraph. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function trySubmit()
	{
		if (document.forms[0].p_header.value.length<3 && document.forms[0].p_body.value.length<3)
		{
			alert("Sorry, you must fill in either header or body.\n(min. 3 characters)");
			return false;
		}
		if (document.forms[0].p_header.value.length<3 && document.forms[0].p_title.value.length<3)
		{
			alert("Sorry, when leaving header blank, you must fill in a title.\n(min. 3 characters)");
			return false;
		}
	document.forms[0].submit();
	return false;

	}

function exitpage()
{
	if (change==true)
	{
		if (confirm("You have changed at least one field on this page.\rBefore exiting this page, do you want to save those changes first?"))
		{
		return trySubmit();
		}
	}
	return true;
}

function WA_ClientSideReplace(theval,findvar,repvar)     {
  var retval = "";
  while (theval.indexOf(findvar) >= 0)    {
    retval += theval.substring(0,theval.indexOf(findvar));
    retval += repvar;
    theval = theval.substring(theval.indexOf(findvar) + String(findvar).length);
  }
  if (retval == "" && theval.indexOf(findvar) < 0)    {
    retval = theval;
  }
  return retval;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function WA_UnloadList(thelist,leavevals,bottomnum)    {
  while (thelist.options.length > leavevals+bottomnum)     {
    if (thelist.options[leavevals])     {
      thelist.options[leavevals] = null;
    }
  }
  return leavevals;
}

function WA_FilterAndPopulateSubList(thearray,sourceselect,targetselect,leaveval,bottomleave,usesource,delimiter)     {
  if (bottomleave > 0)     {
    leaveArray = new Array(bottomleave);
    if (targetselect.options.length >= bottomleave)     {
      for (var m=0; m<bottomleave; m++)     {
        leavetext = targetselect.options[(targetselect.options.length - bottomleave + m)].text;
        leavevalue  = targetselect.options[(targetselect.options.length - bottomleave + m)].value;
        leaveArray[m] = new Array(leavevalue,leavetext);
      }
    }
    else     {
      for (var m=0; m<bottomleave; m++)     {
        leavetext = "";
        leavevalue  = "";
        leaveArray[m] = new Array(leavevalue,leavetext);
      }
    }
  }  
  startid = WA_UnloadList(targetselect,leaveval,0);
  mainids = new Array();
  if (usesource)    maintext = new Array();
  for (var j=0; j<sourceselect.options.length; j++)     {
    if (sourceselect.options[j].selected)     {
      mainids[mainids.length] = sourceselect.options[j].value;
      if (usesource)     maintext[maintext.length] = sourceselect.options[j].text + delimiter;
    }
  }
  for (var i=0; i<thearray.length; i++)     {
    goodid = false;
    for (var h=0; h<mainids.length; h++)     {
      if (thearray[i][0] == mainids[h])     {
        goodid = true;
        break;
      }
    }
    if (goodid)     {
      theBox = targetselect;
      theLength = parseInt(theBox.options.length);
      theServices = thearray[i].length + startid;
      var l=1;
      for (var k=startid; k<theServices; k++)     {
        if (l == thearray[i].length)     break;
        theBox.options[k] = new Option();
        theBox.options[k].value = thearray[i][l][0];
        if (usesource)     theBox.options[k].text = maintext[h] + WA_ClientSideReplace(thearray[i][l][1],"|WA|","'");
        else               theBox.options[k].text = WA_ClientSideReplace(thearray[i][l][1],"|WA|","'");
        l++;
      }
      startid = k;
    }
  }
  if (bottomleave > 0)     {
    for (var n=0; n<leaveArray.length; n++)     {
      targetselect.options[startid+n] = new Option();
      targetselect.options[startid+n].value = leaveArray[n][0];
      targetselect.options[startid+n].text  = leaveArray[n][1];
    }
  }
  for (var l=0; l < targetselect.options.length; l++)    {
    targetselect.options[l].selected = false;
  }
  if (targetselect.options.length > 0)     {
    targetselect.options[0].selected = true;
  }
}

function MM_changeProp(objName,x,theProp,theValue) { //v3.0
  var obj = MM_findObj(objName);
  if (obj && (theProp.indexOf("style.")==-1 || obj.style)) eval("obj."+theProp+"='"+theValue+"'");
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
function selecttopic(whichmenu,whichtopic){
 for (var i=0; i<whichmenu.options.length; i++)     {
 if (whichmenu.options[i].value == whichtopic)	{
 whichmenu.options[i].selected = true;
	 }
 }
}
//-->
</script>
<script type="text/javascript" src="ckeditor/ckeditor.js?v=bbp34"></script>
</HEAD>
<BODY onload="change=false;WA_FilterAndPopulateSubList(topics_WAJA,MM_findObj('subject'),MM_findObj('topic'),0,0,false,': '); selecttopic(MM_findObj('topic'),<%=(paragraph.Fields.Item("page_topic").Value)%>);" onUnload="<% call on_page_unload %>">
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> BBG Edit Paragraph </td>
  </tr>
  <tr> 
    <td align="left" valign="bottom" > 
      <form ACTION="<%=MM_editAction%>" METHOD="POST" name="editb" onSubmit="<%call on_form_Submit(0)%>" onReset="<%call on_form_Reset%>">
        <table>
          <tr> 
            <td  colspan="5" align="left" valign="top"> 
              <table>
                <tr> 
                  <td  width="50">Subject</td>
                  <td > 
                    <select name="subject" onChange="WA_FilterAndPopulateSubList(topics_WAJA,MM_findObj('subject'),MM_findObj('topic'),0,0,false,': '); change=true;" class="formitem1">
                      <%
While (NOT subjects.EOF)
%>
                      <option value="<%=(subjects.Fields.Item("ID_subject").Value)%>" <%if ((subjects.Fields.Item("ID_subject").Value) = (paragraph.Fields.Item("ID_subject").Value)) then Response.Write("SELECTED") : Response.Write("")%>><%=(subjects.Fields.Item("subject_name").Value)%></option>
                      <%
  subjects.MoveNext()
Wend
' If (subjects.CursorType > 0) Then
'   subjects.MoveFirst
' Else
  subjects.Requery
' End If
%>
                    </select>
                  </td>
                </tr>
                <tr> 
                  <td  width="50">Topic </td>
                  <td > 
                    <select name="topic" class="formitem1" onChange="change=true;">
                      <%
While (NOT topics.EOF)
%>
                      <option value="<%=(topics.Fields.Item("ID_topic").Value)%>" ><%=(topics.Fields.Item("topic_name").Value)%></option>
                      <%
  topics.MoveNext()
Wend
'If (topics.CursorType > 0) Then
'  topics.MoveFirst
'Else
  topics.Requery
'End If
%>
                    </select>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td  align="left" valign="top">Title</td>
            <td  align="left" valign="top" colspan="4"> 
              <input type="text" name="p_title" size="80" class="formitem1" onChange="change=true;" value="<%=(paragraph.Fields.Item("page_title").Value)%>">
            </td>
          </tr>
          <tr> 
            <td  colspan="5" align="left" valign="top">Header </td>
          </tr>
          <tr> 
            <td  colspan="5" align="left" valign="top"> 
              <textarea name="p_header" rows="5" class="formitem1" onChange="change=true;" cols="100"><%=(paragraph.Fields.Item("page_header").Value)%></textarea>
			    <script type="text/javascript">
			//<![CDATA[

				CKEDITOR.replace( 'p_header',
					{
					width: 600,
					height: 120
						});

			//]]>
			</script>
            </td>
          </tr>
          <tr> 
            <td  colspan="5" align="left" valign="top">Body </td>
          </tr>
          <tr> 
            <td  colspan="5" align="left" valign="top"> 
              <textarea name="p_body" rows="15" class="formitem1" onChange="change=true;" cols="100"><%=(paragraph.Fields.Item("page_text").Value)%></textarea>
			  <script type="text/javascript">
			//<![CDATA[

				CKEDITOR.replace( 'p_body',
					{
					width: 600,
					height: 200
						});

			//]]>
			</script>
            </td>
          </tr>
          <tr> 
            <td >Icon file</td>
            <td  align="left" valign="top"> 
              <input type="text" name="p_icon" size="75" class="formitem1" onChange="change=true;" value="<%=(paragraph.Fields.Item("page_icon").Value)%>">
            </td>
            <td  align="left" valign="middle" width="20"><a href="javascript:"  onClick="MM_openBrWindow('_ico_browse.asp?formfldname=p_icon','icobrowse','scrollbars=yes,width=610,height=400')"><img src="images/search.gif" width="16" height="16" border="0"></a></td>
            <td  align="left" valign="middle" width="20"> 
              <%
if (paragraph.Fields.Item("page_icon").Value) <> "" Then
	if fileexist("../client/bbg_icons/"& (paragraph.Fields.Item("page_icon").Value)) = True Then
	Response.Write " <img src='../admin/images/ok.gif'> "
	Else
	Response.Write " <img src='../admin/images/miss.gif'> "
	End If 
Else
	Response.Write " <img src='../admin/images/no.gif'> "
End If
%>
            </td>
            <td  align="left" valign="middle" width="50"> 
              <% if (paragraph.Fields.Item("page_icon").Value) <> "" Then %>
              <a href="javascript:"><img src="images/eye.gif" width="16" height="16" border="0" onClick="MM_openBrWindow('../client/bbg_icons/<%=(paragraph.Fields.Item("page_icon").Value)%>','preview','resizable=yes,width=200,height=200,top=250,left=250')"></a> 
              <%End If %>
            </td>
          </tr>
          <tr> 
            <td  colspan="5" align="left" valign="top">Active? 
              <input <%If (abs(paragraph.Fields.Item("page_active").Value) = 1) Then Response.Write("CHECKED") : Response.Write("")%> type="checkbox" name="active" value="checkbox">
            </td>
          </tr>
        </table>
        <p> 
          <input type="reset" name="Reset" value="Reset the form" class="quiz_button">
          <input type="submit" name="Submit" value="Update this paragraph" class="quiz_button" <%call IsEditOK%>>
          <input type="button" name="New_question" value="Add a new paragraph" onClick="document.location='b_paragraph_add.asp';" class="quiz_button">
          or 
          <input type="button" name="goback" value="Go back to paragraphs list" class="quiz_button" onClick="document.location='b_list_of_paragraphs.asp?subj=<%=(paragraph.Fields.Item("ID_subject").Value)%>&topic=<%=(paragraph.Fields.Item("page_topic").Value)%>'">
        </p>
        <input type="hidden" name="MM_update" value="true">
        <input type="hidden" name="MM_recordId" value="<%= paragraph.Fields.Item("ID_page").Value %>">
      </form>
    </td>
  </tr>
</table>
<p>&nbsp;</p><p>&nbsp;</p>
</body>
</HTML>

<%
call log_the_page ("BBG Edit Paragraph: " & (paragraph.Fields.Item("ID_page").Value))
%>

<%
paragraph.Close()
%>
<%
subjects.Close()
%>
<%
topics.Close()
%>
