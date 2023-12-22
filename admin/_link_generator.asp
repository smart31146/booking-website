<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
set subjects = Server.CreateObject("ADODB.Recordset")
subjects.ActiveConnection = Connect
subjects.Source = "SELECT subjects.ID_subject, subjects.subject_name FROM (subjects INNER JOIN b_topics ON subjects.ID_subject = b_topics.topic_subject) INNER JOIN b_pages ON b_topics.ID_topic = b_pages.page_topic GROUP BY subjects.ID_subject, subjects.subject_name, subjects.subject_ord, subjects.ID_subject, Abs([subject_active_b]), Abs([topic_active]), Abs([page_active]) HAVING (((Abs([subject_active_b]))=1) AND ((Abs([topic_active]))=1) AND ((Abs([page_active]))=1)) ORDER BY subjects.subject_ord, subjects.ID_subject;"
subjects.CursorType = 0
subjects.CursorLocation = 3
subjects.LockType = 3
subjects.Open()
subjects_numRows = 0
%>
<%
set topics = Server.CreateObject("ADODB.Recordset")
topics.ActiveConnection = Connect
topics.Source = "SELECT b_topics.ID_topic, b_topics.topic_name, b_topics.topic_subject  FROM subjects INNER JOIN (b_topics INNER JOIN b_pages ON b_topics.ID_topic = b_pages.page_topic) ON subjects.ID_subject = b_topics.topic_subject  GROUP BY b_topics.ID_topic, b_topics.topic_name, b_topics.topic_subject, b_topics.topic_ord, b_topics.ID_topic, Abs([subject_active_b]), Abs([topic_active]), Abs([page_active])  HAVING (((Abs([subject_active_b]))=1) AND ((Abs([topic_active]))=1) AND ((Abs([page_active]))=1))  ORDER BY b_topics.topic_ord, b_topics.ID_topic;"
topics.CursorType = 0
topics.CursorLocation = 3
topics.LockType = 3
topics.Open()
topics_numRows = 0
%>
<%
set paragraphs = Server.CreateObject("ADODB.Recordset")
paragraphs.ActiveConnection = Connect
paragraphs.Source = "SELECT b_pages.ID_page, b_pages.page_title, b_pages.page_topic  FROM subjects INNER JOIN (b_topics INNER JOIN b_pages ON b_topics.ID_topic = b_pages.page_topic) ON subjects.ID_subject = b_topics.topic_subject  GROUP BY b_pages.ID_page, b_pages.page_title, b_pages.page_topic, b_pages.page_ord, b_pages.ID_page, Abs([subject_active_b]), Abs([topic_active]), Abs([page_active])  HAVING (((Abs([subject_active_b]))=1) AND ((Abs([topic_active]))=1) AND ((Abs([page_active]))=1))  ORDER BY b_pages.page_ord, b_pages.ID_page;"
paragraphs.CursorType = 0
paragraphs.CursorLocation = 3
paragraphs.LockType = 3
paragraphs.Open()
paragraphs_numRows = 0
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
<%
function WA_VBreplace(thetext)
  if isNull(thetext) then thetext = ""
  newstring = Replace(cStr(thetext),"'","|WA|")
  newstring = Replace(newstring,"\","\\")
  WA_VBreplace = newstring
end function

if (NOT paragraphs.EOF)     THEN

  Response.Write("<SC" & "RIPT>"&chr(10))
  Response.Write("var WAJA = new Array();"&chr(10))

  oldmainid = 0
  newmainid = paragraphs.Fields("page_topic").value
  if (oldmainid = newmainid)    THEN
    oldmainid = ""
  END IF
  n = 0
    while (NOT paragraphs.EOF)
    if (oldmainid <> newmainid)     THEN
      Response.Write("WAJA[" & n & "] = new Array();"&chr(10))
      Response.Write("WAJA[" & n & "][0] = '" & WA_VBreplace(newmainid) & "';"&chr(10))
      m = 1
    END IF

    Response.Write("WAJA[" & n & "][" & m & "] = new Array();"&chr(10))
    Response.Write("WAJA[" & n & "][" & m & "][0] = " & "'" & WA_VBreplace(paragraphs.Fields("ID_page").value) & "'" & ";" &chr(10))
    Response.Write("WAJA[" & n & "][" & m & "][1] = " & "'" & WA_VBreplace(paragraphs.Fields("page_title").value) & "'" & ";" &chr(10))
    m=m+1
    if (cStr(oldmainid) = "0")      THEN
      oldmainid = newmainid
    END IF
    oldmainid = newmainid
    paragraphs.MoveNext()
    if (NOT paragraphs.EOF)     THEN
      newmainid = paragraphs.Fields("page_topic").value
    END IF
    if (oldmainid <> newmainid)     THEN
      n=n+1
    END IF
  WEND

  Response.Write("var paragraphs_WAJA = WAJA;"&chr(10))
  Response.Write("WAJA = null;"&chr(10))
  Response.Write("</SC" & "RIPT>"&chr(10))
END IF
if (NOT paragraphs.BOF)     THEN
  paragraphs.MoveFirst()
END IF
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Link generator. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}

function generate_url()
{
document.linkgenerator.urlstr.value = "../_bbg/index.asp?ID_subject_prm=" + document.linkgenerator.subjects.value + "&ID_topic_prm=" + document.linkgenerator.topics.value + "#" + document.linkgenerator.paragraphs.value;
document.linkgenerator.linkstr.value = "<a href='../_bbg/index.asp?ID_subject_prm=" + document.linkgenerator.subjects.value + "&ID_topic_prm=" + document.linkgenerator.topics.value + "#" + document.linkgenerator.paragraphs.value + "' target='_self'>Click here</a>";
return true;
}
function generate_link()
{
document.filegenerator.fileurl.value = "<a href='" + document.filegenerator.filelink.value + "' target='_self'>click here</a>";
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

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</HEAD>
<BODY onload="self.focus();">
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> BBG Link generator</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <form name="linkgenerator" method="post" action="">
        <table>
          <tr> 
            <td colspan="2" class="subheads">&nbsp;</td>
          </tr>
          <tr> 
            <td class="text" width="130">Subject</td>
            <td class="text"> 
              <select name="subjects" class="formitem1"  onChange="WA_FilterAndPopulateSubList(topics_WAJA,MM_findObj('subjects'),MM_findObj('topics'),0,0,false,': ')">
                <option>--- select a topic ---</option>
                <%
While (NOT subjects.EOF)
%>
                <option value="<%=(subjects.Fields.Item("ID_subject").Value)%>" ><%=(subjects.Fields.Item("subject_name").Value)%></option>
                <%
  subjects.MoveNext()
Wend
'If (subjects.CursorType > 0) Then
'  subjects.MoveFirst
'Else
  subjects.Requery
'End If
%>
              </select>
            </td>
          </tr>
          <tr> 
            <td class="text"> Topic</td>
            <td class="text"> 
              <select name="topics" class="formitem1" onChange="WA_FilterAndPopulateSubList(paragraphs_WAJA,MM_findObj('topics'),MM_findObj('paragraphs'),0,0,false,': ')">
              </select>
            </td>
          </tr>
          <tr> 
            <td class="text"> Paragraph</td>
            <td class="text"> 
              <select name="paragraphs" class="formitem1">
              </select>
            </td>
          </tr>
          <tr> 
            <td class="text" colspan="2"> 
              <input type="button" name="generateurl" value="Generate a link &amp; URL" class="quiz_button" onClick="generate_url();">
              &amp; copy to clipboard &amp; 
              <input type="button" name="close" value="Close this window" class="quiz_button" onClick="window.close()"></td>
          </tr>
          <tr> 
            <td class="text">URL</td>
            <td class="text"> 
              <input type="text" name="urlstr" size="90" class="formitem1" onClick="this.select();">
            </td>
          </tr>
          <tr> 
            <td class="text">Link </td>
            <td class="text"> 
              <input type="text" name="linkstr" size="90" class="formitem1"  onClick="this.select();">
            </td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
  <tr> 
    <td align="left" valign="bottom">&nbsp;</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom" class="heading">File link generator</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom">
      <form name="filegenerator" method="post" action="">
        <table>
          <tr> 
            <td class="text" width="130">Browse</td>
            <td class="text"><a href="javascript:"  onClick="MM_openBrWindow('_file_browse.asp?path=../client/bbg_files','filebrowse','scrollbars=yes,width=610,height=400')">File 
              manager</a></td>
          </tr>
          <tr> 
            <td class="text" width="130">URL</td>
            <td class="text"> 
              <input type="text" name="filelink" size="90" class="formitem1"  onClick="this.select();">
            </td>
          </tr>
          <tr> 
            <td class="text" colspan="2">
              <input type="button" name="generatelink" value="Generate a link from above URL" class="quiz_button" onClick="generate_link();">
              &amp; copy to clipboard &amp;
<input type="button" name="close2" value="Close this window" class="quiz_button" onClick="window.close()">
            </td>
          </tr>
          <tr> 
            <td class="text" width="130">Link</td>
            <td class="text">
              <input type="text" name="fileurl" size="90" class="formitem1"  onClick="this.select();">
            </td>
          </tr>
        </table>
		</form>
    </td>
  </tr>
</table>
</BODY>
</HTML>

<%
call log_the_page ("Link generator")
%>

<%
subjects.Close()
%>
<%
topics.Close()
%>
<%
paragraphs.Close()
%>

