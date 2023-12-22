<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
numbers=1

	
	if Request.Cookies("show_lines1")<> "" then
		show_lines1= cint(Request.Cookies("show_lines1"))
	else
		show_lines1=15
	end if
	%>

<%
'Dim top
'If (Request.QueryString("topic") <> "") Then 
'top = cInt(Request.QueryString("topic"))
'Else 
'Response.Redirect("error.asp?" & request.QueryString) 
'End If
%>
<%
Dim subj
if cStr(Request.Querystring("show_lines1")) <> "" then show_lines1 = cInt(Request.Querystring("show_lines1"))
If (Request.QueryString("subj") <> "") Then 
subj = cInt(Request.QueryString("subj"))
Else 
'Response.Redirect("error.asp?" & request.QueryString) 
End If
%>
<%
set questions = Server.CreateObject("ADODB.Recordset")
questions.ActiveConnection = Connect
'questions.Source = "SELECT q_question.ID_question, q_question.question_body, q_question.question_ord, q_question.question_active  FROM subjects INNER JOIN (q_topics INNER JOIN q_question ON q_topics.ID_topic = q_question.question_topic) ON subjects.ID_subject = q_topics.topic_subject  WHERE q_question.question_topic =" + Replace(top, "'", "''") + "  ORDER BY q_question.question_ord, q_question.ID_question;"
questions.Source= "SELECT * FROM new_subjects,q_question WHERE s_typ = 2 AND s_qID = "&subj&"  AND s_active = 1 AND question_topic = s_id"
'response.write questions.Source
questions.CursorType = 0
questions.CursorLocation = 3
questions.LockType = 3
questions.Open()
questions_numRows = 0
%>
<%
set subject = Server.CreateObject("ADODB.Recordset")
subject.ActiveConnection = Connect
subject.Source = "SELECT subject_name  FROM subjects  WHERE ID_subject = " + Replace(subj, "'", "''") + ";"
subject.CursorType = 0
subject.CursorLocation = 3
subject.LockType = 3
subject.Open()
subject_numRows = 0
%>
<%
set topic = Server.CreateObject("ADODB.Recordset")
topic.ActiveConnection = Connect
'topic.Source = "SELECT topic_name  FROM q_topics  WHERE ID_topic = " + Replace(top, "'", "''") + ";"
topic.Source = "SELECT id_topic  FROM q_topics  WHERE topic_subject="&subj&";"
topic.CursorType = 0
topic.CursorLocation = 3
topic.LockType = 3
topic.Open()
topic_numRows = 0

%>

<%
set pages_suc_all = Server.CreateObject("ADODB.Recordset")
pages_suc_all.ActiveConnection = Connect
'pages_suc_all.Source = "SELECT q_result.result_question, COUNT(q_result.ID_result) AS ID_result FROM q_result INNER JOIN q_question ON q_result.result_question = q_question.ID_question GROUP BY q_result.result_question, q_question.question_topic HAVING (q_question.question_topic in (SELECT id_topic  FROM q_topics  WHERE topic_subject="&subj&" )) ORDER BY q_result.result_question;"
'pages_suc_all.Source = "select * from q_question a , subjects b, q_topics c where c.topic_subject = b.id_subject and c.id_topic = a.question_topic and c.topic_subject=1"
pages_suc_all.Source= "SELECT q_result.result_question, COUNT(q_result.ID_result) AS ID_result FROM q_choice,q_result,q_question,new_subjects WHERE  result_answer = id_choice  AND result_question = id_question  AND question_topic = s_id AND s_typ = 2 AND s_qID = "&subj&"  AND s_active = 1 GROUP BY q_result.result_question"
'Response.Write pages_suc_all.Source
pages_suc_all.CursorType = 0
pages_suc_all.CursorLocation = 3
pages_suc_all.LockType = 3
pages_suc_all.Open()
pages_suc_all_numRows = 0
%>

<%
Dim Repeat1__numRows
Repeat1__numRows = show_lines1
Dim Repeat1__index
Repeat1__index = 0
questions_numRows = questions_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Repeat2__numRows = -1
Dim Repeat2__index
Repeat2__index = 0
pages_stats_numRows = pages_stats_numRows + Repeat2__numRows
%>
<%
Dim Repeat3__numRows
Repeat3__numRows = -1
Dim Repeat3__index
Repeat3__index = 0
pages_suc_all_numRows = pages_suc_all_numRows + Repeat3__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
questions_total = questions.RecordCount

' set the number of rows displayed on this page
If (questions_numRows < 0) Then
  questions_numRows = questions_total
Elseif (questions_numRows = 0) Then
  questions_numRows = 1
End If

' set the first and last displayed record
questions_first = 1
questions_last  = questions_first + questions_numRows - 1

' if we have the correct record count, check the other stats
If (questions_total <> -1) Then
  If (questions_first > questions_total) Then questions_first = questions_total
  If (questions_last > questions_total) Then questions_last = questions_total
  If (questions_numRows > questions_total) Then questions_numRows = questions_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (questions_total = -1) Then

  ' count the total records by iterating through the recordset
  questions_total=0
  While (Not questions.EOF)
    questions_total = questions_total + 1
    questions.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (questions.CursorType > 0) Then
'    questions.MoveFirst
'  Else
    questions.Requery
  End If

  ' set the number of rows displayed on this page
  If (questions_numRows < 0 Or questions_numRows > questions_total) Then
    questions_numRows = questions_total
  End If

  ' set the first and last displayed record
  questions_first = 1
  questions_last = questions_first + questions_numRows - 1
  If (questions_first > questions_total) Then questions_first = questions_total
  If (questions_last > questions_total) Then questions_last = questions_total

End If
%>
<%
' *** Move To Record and Go To Record: declare variables

Set MM_rs    = questions
MM_rsCount   = questions_total
MM_size      = questions_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  r = Request.QueryString("index")
  If r = "" Then r = Request.QueryString("offset")
  If r <> "" Then MM_offset = Int(r)

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  i = 0
  While ((Not MM_rs.EOF) And (i < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    i = i + 1
  Wend
  If (MM_rs.EOF) Then MM_offset = i  ' set MM_offset to the last possible record

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  i = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or i < MM_offset + MM_size))
    MM_rs.MoveNext
    i = i + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = i
    If (MM_size < 0 Or MM_size > MM_rsCount) Then MM_size = MM_rsCount
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
'  If (MM_rs.CursorType > 0) Then
'    MM_rs.MoveFirst
'  Else
    MM_rs.Requery
'  End If

  ' move the cursor to the selected record
  i = 0
  While (Not MM_rs.EOF And i < MM_offset)
    MM_rs.MoveNext
    i = i + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
questions_first = MM_offset + 1
questions_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (questions_first > MM_rsCount) Then questions_first = MM_rsCount
  If (questions_last > MM_rsCount) Then questions_last = MM_rsCount
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then MM_removeList = MM_removeList & "&" & MM_paramName & "="
MM_keepURL="":MM_keepForm="":MM_keepBoth="":MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
if (MM_keepBoth <> "") Then MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
if (MM_keepURL <> "")  Then MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
if (MM_keepForm <> "") Then MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 0) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    params = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For i = 0 To UBound(params)
      nextItem = Left(params(i), InStr(params(i),"=") - 1)
      If (StrComp(nextItem,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & params(i)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then MM_keepMove = MM_keepMove & "&"
urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="
MM_moveFirst = urlStr & "0"
MM_moveLast  = urlStr & "-1"
MM_moveNext  = urlStr & Cstr(MM_offset + MM_size)
prev = MM_offset - MM_size
If (prev < 0) Then prev = 0
MM_movePrev  = urlStr & Cstr(prev)
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz questions. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
var MyCookie = {
    Write:function(name,value,days) {
        var D = new Date();
        D.setTime(D.getTime()+86400000*days)
        document.cookie = escape(name)+"="+escape(value)+
            ((days == null)?"":(";expires="+D.toGMTString()))
        return (this.Read(name) == value);
    },
    Read:function(name) {
        var EN=escape(name)
        var F=' '+document.cookie+';', S=F.indexOf(' '+EN);
        return S==-1 ? null : unescape(F.substring(EN=S+EN.length+2,F.indexOf(';',EN)));
    }
} 
function check()
{
	show_lines1 = MyCookie.Read('show_lines1')
	if (show_lines1 != null) {
		document.forms[0].show_lines1.value=show_lines1;
	}
	else {
		document.forms[0].show_lines1.value=15;
	}
}

function show()
{
	if (isNaN(document.questions.show_lines1.value)){
		alert('Invalid number');
		document.questions.show_lines1.focus();
		return false;
	}
	else
	{
		AddCookieId("show_lines1",document.questions.show_lines1.value);
		show_lines1 = MyCookie.Read('show_lines1')
		document.forms[0].submit();
		return true;
	}
}

function AddCookieId(cn,id) {
        MyCookie.Write(cn,id,7);
}

function DelCookieId(cn,id) {
        MyCookie.Write(cn,id,-1);
}
//-->
</script>
<script src="styles/lytebox.js?v=bbp34" type="text/javascript"></script>
<link rel="STYLESHEET" type="text/css" href="styles/lytebox.css">
</HEAD>
<BODY onload="check();">
<form name="questions">

<input type="hidden" name="subj" value=<%=request("subj")%>>
<table>
    <tr> 
    <td align="left" valign="bottom" valign=top> 
      <table>
        <tr> 
          <td colspan="3" class="subheads">Questions in <%=(subject.Fields.Item("subject_name").Value)%> </td>
        </tr>
        <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')" onClick="document.location ='q_list_of_subjects.asp'"> 
            <td class="text" width="10"><img src="../admin/images/back.gif" width="18" height="14"></td>
            <td class="text" colspan="6"><a href="../admin/q_list_of_subjects.asp">...go 
              up one level to list of Subjects</a></td>
          </tr>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT questions.EOF)) 
%>
        <% If Not questions.EOF Or Not questions.BOF Then %>
        <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')" onClick="document.location ='q_question_edit.asp?qid=<%=(questions.Fields.Item("ID_question").Value)%>'"> 
          <td class="text" width="20"><%=numbers%></td>
		  <td class="text"><%=(questions.Fields.Item("s_topic").Value)%></td>
		  <td class="text" width="60" align="center">ID: <%=(questions.Fields.Item("ID_question").Value)%> </td>
          <td class="text"> 
           <a href="q_training_lyte.asp?alt=editquiz&ID_question=<%=(questions.Fields.Item("ID_question").Value)%>" title=""><% =(CropSentence((questions.Fields.Item("question_body").Value), 120, "...")) %></a>
            </td>
          <td width="20" class="text"> 
            <%if abs(questions.Fields.Item("question_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
          </td>
        </tr>
        <% End If%>
        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  questions.MoveNext()
  numbers=numbers+1
Wend
%>
        <% If questions.EOF And questions.BOF Then %>
        <tr> 
          <td class="text">&nbsp;</td>
          <td width="99%"  colspan="2">Sorry, 
            there are no questions in this topic currently.</td>
        </tr>
        <% End If %>
        <tr> 
          <!--<td class="text"><img src="images/new2.gif" width="11" height="13"></td>
          <td width="99%" class="text" colspan="2"> 
            <input type="button" name="Button" value="Add a new question" onClick="document.location='q_question_add.asp?subj=<%=subj%>&topic=<%=top%>';" class="quiz_button">
          </td>-->
          
        </tr>
        
      </table></td></tr><tr><td>
      <table>
                <tr class="table_normal"> 
                 
                    <% If MM_offset <> 0 Then %>
                    <td align="center" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
                    <a href="<%=MM_moveFirst%>"><img src="images/first.gif" border=0></a> 
                    <%else%>
                    <td align="center"> 
                    <img src="images/first.gif" border=0> 
                    <% End If %>
                  </td>
                  
                    <% If MM_offset <> 0 Then %>
                    <td align="center" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
                    <a href="<%=MM_movePrev%>"><img src="images/previous.gif" border=0></a> 
                    <%else %>
                    <td align="center"> 
                    <img src="images/previous.gif" border=0> 
                    <%End If %>
                  </td>
                  
                    <% If Not MM_atTotal Then %>
					<td align="center" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
                    <a href="<%=MM_moveNext%>"><img src="images/next.gif" border=0></a> 
                    <%else%>
                    <td align="center"> 
                    <img src="images/next.gif" border=0> 
                    <% End If %>
                  </td>
                  
                    <% If Not MM_atTotal Then %>
                    <td align="center" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
                    <a href="<%=MM_moveLast%>"><img src="images/last.gif" border=0></a> 
                    <% else %>
                    <td align="center"> 
                    <img src="images/last.gif" border=0> 
                    <%End If %>
                  </td>
                </tr>
                <tr class="table_normal"> 
                  <td colspan="4" align="center"class="text">&nbsp; Questions <b><%=(questions_first)%></b> to <b><%=(questions_last)%></b> of <b><i><%=(questions_total)%> </i></b> 
                    <%if (SQL_having <> "") or (SQL_where <> "") then response.write("<font color='#FF0000'>(filtered - <a href='javascript:clearform();'>clear filter</a>)</font>")%>
                  </td>
                </tr>
                <tr class="table_normal"> 
                  <td colspan="4" align="center"class="text">Show 
                    <input type="text" name="show_lines1" class="formitem1" size="3" maxlength="3" value="">
                    Questions per page</td>
                </tr>
                <tr class="table_normal"> 
                  <td colspan="4" align="center"class="text">
                <input type="button" name="Submit" value="&gt;&gt;&gt; Show &lt;&lt;&lt;" class="quiz_button" onclick="show();">
                </td>
                </tr>
              </table>
     
      <p>&nbsp;</p>
    </td>

  <tr> 
    <td align="left" valign="bottom"> 
      <table>
        <tr> 
          <td class="subheads">Incorrect/correct ratio</td>
        </tr>
        <% If Not pages_suc_all.EOF Or Not pages_suc_all.BOF Then %>
        <tr> 
          <td> 
            <table>
              <%
overall_correct = 0
overall_incorrect = 0
%>
              <% 
While ((Repeat3__numRows <> 0) AND (NOT pages_suc_all.EOF)) 
%>
              <%
pages_all = (pages_suc_all.Fields.Item("ID_result").Value)
%>
              <%
set pages_suc_ok = Server.CreateObject("ADODB.Recordset")
pages_suc_ok.ActiveConnection = Connect
pages_suc_ok.Source = "SELECT COUNT(q_result.ID_result) AS ID_result FROM q_result INNER JOIN q_choice ON q_result.result_answer = q_choice.ID_choice GROUP BY q_result.result_question, q_choice.choice_cor HAVING (q_result.result_question = " & cInt(pages_suc_all.Fields.Item("result_question").Value) & ") AND (q_choice.choice_cor = 1);"
'Response.Write pages_suc_ok.Source
pages_suc_ok.CursorType = 0
pages_suc_ok.CursorLocation = 3
pages_suc_ok.LockType = 3
pages_suc_ok.Open()
pages_suc_ok_numRows = 0
%>
              <%
if not pages_suc_ok.EOF or Not pages_suc_ok.BOF Then pages_correct = (pages_suc_ok.Fields.Item("ID_result").Value) else pages_correct = 0
pages_incorrect = pages_all - pages_correct
overall_correct = overall_correct + pages_correct
overall_incorrect = overall_incorrect + pages_incorrect
%>
              <tr valign="middle" class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')" onClick="document.location =' q_training_lyte.asp?alt=editquiz&ID_question=<%=(pages_suc_all.Fields.Item("result_question").Value)%>'"> 
                <td width="280" align="right">(<%=pages_incorrect%>)<img src="images/bari2.gif"><img src="images/bari1.gif" width="<%=cInt(pages_incorrect/pages_all*stat_bar_length)%>" height="9"><img src="images/bari0.gif"></td>
                <td width="40" align="center">-&nbsp;<%=(pages_suc_all.Fields.Item("result_question").Value)%>&nbsp;-</td>
                <td width="280" align="left"><img src="images/barc0.gif"><img src="images/barc1.gif" width="<%=cInt(pages_correct/pages_all*stat_bar_length)%>" height="9"><img src="images/barc2.gif">(<%=pages_correct%>)</td>
              </tr>
              <%
pages_suc_ok.Close()
%>
              <% 
  Repeat3__index=Repeat3__index+1
  Repeat3__numRows=Repeat3__numRows-1
  pages_suc_all.MoveNext()
Wend
%>
            </table>
          </td>
        </tr>
        <tr> 
          <td>
            <table>
              <tr valign="middle"> 
                <td width="280"  align="right">(<%=overall_incorrect%>)<img src="images/bari2.gif"><img src="images/bari1.gif" width="<%=cInt(overall_incorrect/(overall_incorrect+overall_correct)*stat_bar_length)%>" height="9"><img src="images/bari0.gif"></td>
                <td width="40"  align="center">-Total-</td>
                <td width="280"  align="left"><img src="images/barc1.gif" width="<%=cInt(overall_correct/(overall_incorrect+overall_correct)*stat_bar_length)%>" height="9"><img src="images/barc2.gif">(<%=overall_correct%>)</td>
              </tr>
            </table>
          </td>
        </tr>
        <% End If %>
        <tr> 
          <% If pages_suc_all.EOF And pages_suc_all.BOF Then %>
          <td >Sorry, 
            there are no results available yet...</td>
          <% End If  %>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p>&nbsp;</p></form>
</BODY>
</HTML>

<%
call log_the_page ("Quiz List Questions: " & subj& ", " & top)
%>

<%
questions.Close()
%>
<%
subject.Close()
%>
<%
topic.Close()
%>

<%
pages_suc_all.Close()
%>

