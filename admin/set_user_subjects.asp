<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
'Function StringStartsWith(byVal strValue As String, _
 ' CheckFor As String, Optional CompareType As VbCompareMethod _
  ' = vbBinaryCompare) As Boolean
   
'Determines if a string starts with the same characters as 
'CheckFor string

'True if starts with CheckFor, false otherwise
'Case sensitive by default.  If you want non-case sensitive, set
'last parameter to vbTextCompare
    
    'Examples:
    'MsgBox StringStartsWith("Test", "TE") 'false
    'MsgBox StringStartsWith("Test", "TE", vbTextCompare) 'True
    
  'Dim sCompare As String
  'Dim lLen As Long
   
  'lLen = Len(CheckFor)
 ' If lLen > Len(strValue) Then Exit Function
  'sCompare = Left(strValue, lLen)
  'StringStartsWith = StrComp(sCompare, CheckFor, CompareType) = 0

'End Function


Function Tokenize(byVal TokenString, byRef TokenSeparators())

	Dim NumWords, a()
	NumWords = 0
	
	Dim NumSeps
	NumSeps = UBound(TokenSeparators)
	
	Do 
		Dim SepIndex, SepPosition
		SepPosition = 0
		SepIndex    = -1
		
		for i = 0 to NumSeps-1
		
			' Find location of separator in the string
			Dim pos
			pos = InStr(TokenString, TokenSeparators(i))
			
			' Is the separator present, and is it closest to the beginning of the string?
			If pos > 0 and ( (SepPosition = 0) or (pos < SepPosition) ) Then
				SepPosition = pos
				SepIndex    = i
			End If
			
		Next

		' Did we find any separators?	
		If SepIndex < 0 Then

			' None found - so the token is the remaining string
			redim preserve a(NumWords+1)
			a(NumWords) = TokenString
			
		Else

			' Found a token - pull out the substring		
			Dim substr
			substr = Trim(Left(TokenString, SepPosition-1))
	
			' Add the token to the list
			redim preserve a(NumWords+1)
			a(NumWords) = substr
		
			' Cutoff the token we just found
			Dim TrimPosition
			TrimPosition = SepPosition+Len(TokenSeparators(SepIndex))
			TokenString = Trim(Mid(TokenString, TrimPosition))
						
		End If	
		
		NumWords = NumWords + 1
	loop while (SepIndex >= 0)
	
	Tokenize = a
	
End Function




'PN 050720 Save the user subjects that have been submitted
Dim updated_ok
updated_ok=false
Dim Seps(1)
Seps(0) = "|"
'PN 050720 delete from the  subject user table
is_update=Request.Form("MM_insert")		
if(is_update<>"") then
	Set MM_editCmd = Server.CreateObject("ADODB.Command")
	MM_editCmd.ActiveConnection = Connect
	MM_editCmd.CommandText = "delete from subject_user;"
	MM_editCmd.Execute
	MM_editCmd.ActiveConnection.Close
end if

For Each q in Request.Form()
	
	if (((InStr(q,"user_subject"))>0)=True) then
		
			Dim a
			a= Tokenize(q, Seps)
			
			'050720 do an insert to the subject_user table
			Set MM_editCmd = Server.CreateObject("ADODB.Command")
			MM_editCmd.ActiveConnection = Connect
			MM_editCmd.CommandText = "insert into subject_user (ID_subject, ID_user) values ("&a(2)&","&a(1)&");"
			MM_editCmd.Execute
			MM_editCmd.ActiveConnection.Close
			
			
			'Response.Write "<li>Keyword " &a(1)&"   "&a(2)&"    "&a(3)& "   "&Request.Form(q)&"</li>"
			
			'Response.Write "starts with " & q & "<br>"
	end if
	

Next
if(is_update<>"") then
	updated_ok=true
end if
if updated_ok then
	Response.Write "<p><font color=red>The user subjects have been saved</font></p>"
end if

%>

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
' *** Insert Record: set variables


%>
<%

function WA_VBreplace(thetext)
  if isNull(thetext) then thetext = ""
  newstring = Replace(cStr(thetext),"'","|WA|")
  newstring = Replace(newstring,"\","\\")
  WA_VBreplace = newstring
end function


%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE> BBP ADMIN: Question add. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
var selected = false;

function selectAllOrNone()
	{
		//cycle through the passrates and check they are all a number
		var arrayOfCheckboxes= document.subject_user.elements;

		var isANumber=true;
		var temp="";
		if(selected){
			for(p=0;p<arrayOfCheckboxes.length;p++){
				if(((arrayOfCheckboxes[p].name).indexOf("user_subject"))==0){
						arrayOfCheckboxes[p].checked=false;
							
				}
			}
			selected=false;
		}
		else{
			for(p=0;p<arrayOfCheckboxes.length;p++){
				if(((arrayOfCheckboxes[p].name).indexOf("user_subject"))==0){
						arrayOfCheckboxes[p].checked=true;
							
				}
			}
			selected=true;
		}
					
				
			
		
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


function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</HEAD>
<BODY onLoad="change=false;" onUnload="<% call on_page_unload %>">
<table>
  <tr>
    <td align="left" valign="bottom" class="heading"> Set user subjects</td>
  </tr>
  <tr>
    <td align="left" valign="bottom" class="text"> 
      <form METHOD="POST" name="subject_user" action="<%=MM_editAction%>">
	  <input type="button" name="selectall" value="Select all/none" class="quiz_button" onClick=" selectAllOrNone()">
		 
       <%
           
			
			'pn 050720 pull out all users
			set users = Server.CreateObject("ADODB.Recordset")
			users.ActiveConnection = Connect
			users.Source = "SELECT * from q_user order by user_lastname asc;"
			users.CursorType = 0
			users.CursorLocation = 2
			users.LockType = 3
			users.Open()
			
			%>
			<table>
			<%
			While (NOT users.EOF)
			%>

							<tr class="table_normal">
									
									<td class="text" width="50">
									
										<%=users.Fields.Item("user_firstname").Value%> <%=users.Fields.Item("user_lastname").Value%>
										
									</td>
										<%
										'pn 050720 pull out all active subjects, currently only in reference to guide
										
										set subjects_b = Server.CreateObject("ADODB.Recordset")
										subjects_b.ActiveConnection = Connect
										subjects_b.Source = "SELECT subjects.ID_subject, subjects.subject_name  FROM (subjects INNER JOIN b_topics ON subjects.ID_subject = b_topics.topic_subject) INNER JOIN b_pages ON b_topics.ID_topic = b_pages.page_topic  GROUP BY subjects.ID_subject, subjects.subject_name, subjects.subject_ord, subjects.ID_subject, Abs([subject_active_b]), Abs([topic_active]), Abs([page_active])  HAVING (((Abs([subject_active_b]))=1) AND ((Abs([topic_active]))=1) AND ((Abs([page_active]))=1))  ORDER BY subjects.subject_ord, subjects.ID_subject;"
										subjects_b.CursorType = 0
										subjects_b.CursorLocation = 2
										subjects_b.LockType = 3
										subjects_b.Open()
										subjects_b_numRows = 0
										
										While (NOT subjects_b.EOF)
										
													show_this_subject=""
														
															set user_bbp_subject = Server.CreateObject("ADODB.Recordset")
															user_bbp_subject.ActiveConnection = Connect
															user_bbp_subject.Source = "SELECT *  FROM subject_user where ID_user="& users.Fields.Item("ID_user").Value &" and ID_subject="&subjects_b.Fields.Item("ID_subject").Value&";"
															user_bbp_subject.CursorType = 0
															user_bbp_subject.CursorLocation = 2
															user_bbp_subject.LockType = 3
															user_bbp_subject.Open()
																While ((NOT user_bbp_subject.EOF)) 
																	show_this_subject="checked"
																	user_bbp_subject.MoveNext()
																Wend

															user_bbp_subject.Close()
														%>
											<td  class="text" width="100">
												<%=subjects_b.Fields.Item("subject_name").Value%>
											</td>
											<td  class="text" width="50">
											
												<input type="checkbox" <%=show_this_subject%> name="user_subject|<%=users.Fields.Item("ID_user").Value%>|<%=subjects_b.Fields.Item("ID_subject").Value%>" />
											</td>
										
										<%subjects_b.MoveNext()
												Wend
												subjects_b.Close()
								users.MoveNext()
								%>
								</tr>
								<%
						wend
						users.Close()
						%>
				</table>

        <p> 
          
          <input type="submit" name="Submit" value="Save user subjects" class="quiz_button" <%call IsEditOK%>>
          or 
          <input type="button" name="goback" value="Go back" class="quiz_button" onClick="history.go(-1)">
		  
        </p>
        <input type="hidden" name="MM_insert" value="true">
       
        <input type="hidden" name="active" value="1">
        <input type="hidden" name="UID" value="<%=GetUniqueID("q_",20,"")%>">
      </form>
    </td>
  </tr>
</table>
<p>&nbsp;</p></BODY>
</HTML>

<%
call log_the_page ("Set pass rates ")
%>


