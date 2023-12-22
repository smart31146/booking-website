<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
numbers=1
%>
<%
Dim top
If (Request.QueryString("topic") <> "") Then 
top = cInt(Request.QueryString("topic"))
Else 
Response.Redirect("error.asp?" & request.QueryString) 
End If
%>
<%
Dim subj
If (Request.QueryString("subj") <> "") Then 
subj = cInt(Request.QueryString("subj"))
Else 
Response.Redirect("error.asp?" & request.QueryString) 
End If
%>
<%
set pages = Server.CreateObject("ADODB.Recordset")
pages.ActiveConnection = Connect
pages.Source = "SELECT b_pages.ID_page, b_pages.page_title, b_pages.page_header, b_pages.page_ord, b_pages.page_active  FROM subjects INNER JOIN (b_topics INNER JOIN b_pages ON b_topics.ID_topic = b_pages.page_topic) ON subjects.ID_subject = b_topics.topic_subject  WHERE (((b_pages.page_topic)=" + Replace(top, "'", "''") + "))  ORDER BY b_pages.page_ord, b_pages.ID_page;"
pages.CursorType = 0
pages.CursorLocation = 3
pages.LockType = 3
pages.Open()
pages_numRows = 0
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
topic.Source = "SELECT topic_name  FROM b_topics  WHERE ID_topic = " + Replace(top, "'", "''") + ";"
topic.CursorType = 0
topic.CursorLocation = 3
topic.LockType = 3
topic.Open()
topic_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
pages_numRows = pages_numRows + Repeat1__numRows
%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Reference paragraphs. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="../admin/styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}
//-->
</script>
</HEAD>
<BODY>
<table>
  <tr> 
    <td align="left" valign="bottom" class="heading"> BBG paragraphs</td>
  </tr>
  <tr> 
    <td align="left" valign="bottom"> 
      <table>
        <tr> 
          <td colspan="4" class="subheads">Paragraphs on page <%=(subject.Fields.Item("subject_name").Value)%> / <%=(topic.Fields.Item("topic_name").Value)%>:</td>
        </tr>
        <tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
          <td class="text" width="10"><img src="../admin/images/back.gif" width="18" height="14"></td>
          <td colspan="3" class="text"><a href="../admin/b_list_of_topics.asp?subj=<%=subj%>&topic=<%=top%>">...go 
            up one level to list of Topics</a></td>
        </tr>
        <% 
While ((Repeat1__numRows <> 0) AND (NOT pages.EOF)) 
%>
    <% If Not pages.EOF Or Not pages.BOF Then %>
	
		<% If (Cstr(Request.QueryString("highlight_q")) = Cstr(pages.Fields.Item("ID_page").Value)) Then  %>

			<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')" style="background-color: #FFFF33"> 
				<td class="text" width="20"><%=numbers%></td>
				<td width="540" class="text"><a href="b_paragraph_edit.asp?pid=<%=(pages.Fields.Item("ID_page").Value)%>"> 
				<% ="[" & (pages.Fields.Item("page_title").Value) & "]: " & (CropSentence((pages.Fields.Item("page_header").Value), 50, "...")) %>
				</a></td>
				<td width="20" class="text"> 
				<%if abs(pages.Fields.Item("page_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
				</td>
				<td width="20" class="text"><%if Edit_OK then %><a href="b_paragraph_del.asp?pid=<%=(pages.Fields.Item("ID_page").Value)%>&subj=<%=subj%>&topic=<%=top%>" onClick="javascript: return (confirm('You are just about to delete this paragraph.\nAre you sure you want to do that?'));"><img src="images/bin.gif" width="16" height="16" border="0"></a><% end if %></td>
			</tr>
	
		<% else %>

			<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_hl')" onMouseOut="pviiClassNew(this,'table_normal')"> 
				<td class="text" width="20"><%=numbers%></td>
				<td width="540" class="text"><a href="b_paragraph_edit.asp?pid=<%=(pages.Fields.Item("ID_page").Value)%>"> 
				<% ="[" & (pages.Fields.Item("page_title").Value) & "]: " & (CropSentence((pages.Fields.Item("page_header").Value), 50, "...")) %>
				</a></td>
				<td width="20" class="text"> 
				<%if abs(pages.Fields.Item("page_active").Value) = 1 then response.write "<img src='images/1.gif'>" else response.write "<img src='images/0.gif'>"%>
				</td>
				<td width="20" class="text"><%if Edit_OK then %><a href="b_paragraph_del.asp?pid=<%=(pages.Fields.Item("ID_page").Value)%>&subj=<%=subj%>&topic=<%=top%>" onClick="javascript: return (confirm('You are just about to delete this paragraph.\nAre you sure you want to do that?'));"><img src="images/bin.gif" width="16" height="16" border="0"></a><% end if %></td>
			</tr>
			
		<% end if %>	
		
    <% End If ' end Not pages.EOF Or NOT pages.BOF 
	Repeat1__index=Repeat1__index+1
	Repeat1__numRows=Repeat1__numRows-1
	pages.MoveNext()
	numbers=numbers+1
Wend
%>
        <% If pages.EOF And pages.BOF Then %>
        <tr> 
          <td class="text">&nbsp;</td>
          <td width="99%"  colspan="3">Sorry, 
            there is no paragraph in this topic currently.</td>
        </tr>
        <% End If %>
        <tr> 
          <td class="text"><img src="images/new2.gif" width="11" height="13"></td>
          <td width="99%" class="text" colspan="3"> 
            <input type="button" name="Button" value="Add a new paragraph to this topic" onClick="document.location='b_paragraph_add.asp?subj=<%=subj%>&topic=<%=top%>';" class="quiz_button">
          </td>
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
call log_the_page ("BBG List Paragraphs: " & subj & ", " & top)
%>

<%
pages.Close()
%>
<%
subject.Close()
%>
<%
topic.Close()
%>
