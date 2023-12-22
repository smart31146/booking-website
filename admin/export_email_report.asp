<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->

<%
' *** Edit Operations: declare variables

MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If


Dim email_sent_message
' query string to execute
MM_editQuery = ""
Dim array_of_subjects



function WA_VBreplace(thetext)
  if isNull(thetext) then thetext = ""
  newstring = Replace(cStr(thetext),"'","|WA|")
  newstring = Replace(newstring,"\","\\")
  WA_VBreplace = newstring
end function



Dim firstname_filter, lastname_filter
firstname_filter=""
lastname_filter=""

If (Request("user_firstname")<>"")  then 
	firstname_filter=" and user_firstname like '%"+Request("user_firstname")+"%' "
end if 

If (Request("user_lastname")<>"")  then 
	lastname_filter=" and user_lastname like '%"+Request("user_lastname")+"%' "
end if 



Dim bd_filter, location_filter
bd_filter=""
location_filter=""


filter_info1_prm = 0
If cInt(Request("filter_info1")) <> 0 then 
	filter_info1_prm = cInt(Request("filter_info1"))
	bd_filter = " and (q_user.user_info1)= " + (Request("filter_info1")) + " " 
	
end if 

filter_info2_prm = 0
If cInt(Request("filter_info2")) <> 0 then 
	filter_info2_prm = cInt(Request("filter_info2"))
	location_filter =  " AND (q_user.user_info2)= " + (Request("filter_info2")) + " " 
	
end if 

set filter_info2 = Server.CreateObject("ADODB.Recordset")
filter_info2.ActiveConnection = Connect
if request("filter_info1")<> "" then 
	info2_prm = request("filter_info1")
else
	info2_prm = 0
end if	
filter_info2.Source = "SELECT * FROM q_info2 where info2_info1 =" & info2_prm &" order by info2"
filter_info2.CursorType = 0
filter_info2.CursorLocation = 3
filter_info2.LockType = 3
filter_info2.Open()
filter_info2_numRows = 0



set filter_info1 = Server.CreateObject("ADODB.Recordset")
filter_info1.ActiveConnection = Connect
filter_info1.Source = "SELECT * FROM q_info1 where info1_active=1 order by info1"
filter_info1.CursorType = 0
filter_info1.CursorLocation = 3
filter_info1.LockType = 3
filter_info1.Open()
filter_info1_numRows = 0


Dim subject_filter
subject_filter=""

If (Cint(Request("subject"))<>0)  then 

	subject_filter=" and subject = "+Request("subject")+" "

end if 

Dim type_filter
type_filter=""

If (Cint(Request("type"))<>0)  then 

	type_filter=" and type = "+Request("type")+" "

end if 

Dim from_date_filter
from_date_filter=""

If ((Request("from_date"))<>"")  then 
	temp_from_date=Request("from_date")
	from_sql_date= right(temp_from_date,4)&"-"& mid(temp_from_date,4,2) &"-"& left(temp_from_date,2) 
	from_date_filter=" and date_to_send > '"+from_sql_date+"' "

end if 

Dim to_date_filter
to_date_filter=""

If ((Request("to_date"))<>"")  then 
	temp_to_date=Request("to_date")
	to_sql_date= right(temp_to_date,4)&"-"& mid(temp_to_date,4,2) &"-"& left(temp_to_date,2) 
	to_date_filter=" and date_to_send < '"+to_sql_date+"' "

end if 


set subjects = Server.CreateObject("ADODB.Recordset")
subjects.ActiveConnection = Connect
subjects.Source = "SELECT ID_subject, subject_name FROM subjects where subject_active_q <> 0"
subjects.CursorType = 0
subjects.CursorLocation = 3
subjects.LockType = 3
subjects.Open()
subjects_numRows = 0


set to_be_sent = Server.CreateObject("ADODB.Recordset")
to_be_sent.ActiveConnection = Connect
to_be_sent.Source = "SELECT * from emails  inner join q_user on emails.q_user=q_user.ID_user where status= 0 "+firstname_filter+lastname_filter+bd_filter+location_filter+type_filter+from_date_filter+to_date_filter+" order by date_to_send desc"
to_be_sent.CursorType = 0
to_be_sent.CursorLocation = 3
to_be_sent.LockType = 3
to_be_sent_numRows = 0

set sent = Server.CreateObject("ADODB.Recordset")
sent.ActiveConnection = Connect
sent.Source = "SELECT * from emails  inner join q_user on emails.q_user=q_user.ID_user where status<> 0 "+firstname_filter+lastname_filter+bd_filter+location_filter+type_filter+from_date_filter+to_date_filter+" order by date_to_send desc"
sent.CursorType = 0
sent.CursorLocation = 3
sent.LockType = 3
sent_numRows = 0

%>
<% Response.Clear()
Response.AddHeader "Content-Disposition","attachment; filename=Email_Report_" & day(now()) & "_" & month(now()) & "_" & year(now()) & ".xls"
Response.ContentType="application/vnd.ms-excel"

%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz new user. You are logged in as <%=Session("MM_Username_admin")%></TITLE>



<script language="JavaScript">
<!--
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

function WA_UnloadList(thelist,leavevals,bottomnum)    {
  while (thelist.options.length > leavevals+bottomnum)     {
    if (thelist.options[leavevals])     {
      thelist.options[leavevals] = null;
    }
  }
  return leavevals;
}




function WA_subAwithBinC(a,b,c)
{

	var i = c.indexOf(a);
	var l = b.length;

	while (i != -1)	{
		c = c.substring(0,i) + b + c.substring(i + a.length,c.length);  //replace all valid a values with b values in the selected string c.
  i += l
		i = c.indexOf(a,i);
	}
	return c;

}



function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}




function replace(string,text,by)
{
    var strLength = string.length, txtLength = text.length;
    if ((strLength == 0) || (txtLength == 0)) return string;
    var i = string.indexOf(text);
    if ((!i) && (text != string.substring(0,txtLength))) return string;
    if (i == -1) return string;
    var newstr = string.substring(0,i) + by;
    if (i+txtLength < strLength)
        newstr += replace(string.substring(i+txtLength,strLength),text,by);
    return newstr;
}

function checkform() {
	document.forms[0].action="email_report.asp"
	document.forms[0].target="_self"
	document.forms[0].submit()
}
//-->
</script>


</HEAD>
<BODY>
<table>
  <tr>
    <td align="left" valign="bottom" class="heading"> Email report</td>
  </tr>
  <tr>
    <td align="left" valign="bottom">
      
        <table>
			<tr> 
				<td class="subheads" colspan="5">To be sent:</td>
			</tr>
			<tr> 
				<td   width="100">Name</td>
				<td   width="120">Send date</td>
				<td  width="150">Requested subjects completed</td>
				<td  width="150">Subject scores</td>
				<td  width="130">Email type</td>
			</tr>
          <%		
					Dim tobeSent_SubjectIDArray()  
					Dim tobeSent_SubjectNameArray() 
					Dim tobeSent_SubjectScoreArray() 
					to_be_sent.Open()
					While (NOT to_be_sent.EOF)
						
						got_subject=false
						set filter_subject = Server.CreateObject("ADODB.Recordset")
						filter_subject.ActiveConnection = Connect
						filter_subject.Source = "SELECT  *  from subject_email  where email="&to_be_sent.Fields.Item("id").Value&"  "&subject_filter
						filter_subject.CursorType = 0
						filter_subject.CursorLocation = 3
						filter_subject.LockType = 3
						filter_subject_numRows = 0
						filter_subject.Open()
						While (NOT filter_subject.EOF)
							got_subject=true
							filter_subject.MoveNext()
						Wend
						filter_subject.Close()
						
						if(got_subject) then
								%>
								<tr class="table_normal" >
									<td class="text" align="left" valign="top">
												<%=to_be_sent.Fields.Item("user_firstname").Value%>&nbsp;<%=to_be_sent.Fields.Item("user_lastname").Value%>
									</td>
									<td class="text" align="left" valign="top">
												<%=to_be_sent.Fields.Item("date_to_send").Value%>
									</td>
												<%
									
									
									tobeSent_emailType=""
									if(to_be_sent.Fields.Item("type").Value=1) then
										tobeSent_emailType="Induction Welcome"
									elseif (to_be_sent.Fields.Item("type").Value=2) then
										tobeSent_emailType="Auto reminder"
									elseif (to_be_sent.Fields.Item("type").Value=3) then
										tobeSent_emailType="Escalation reminder"
									elseif (to_be_sent.Fields.Item("type").Value=4) then
										tobeSent_emailType="Final reminder"
									elseif (to_be_sent.Fields.Item("type").Value=5) then
										tobeSent_emailType="Inducation Escalation"
									elseif (to_be_sent.Fields.Item("type").Value=6) then
										tobeSent_emailType="Inducation Final"
									elseif (to_be_sent.Fields.Item("type").Value=7) then
										tobeSent_emailType="Manual reminder"	
									end if
									
									tobeSent_numberSubjects=0
									set tobeSent_number_of_subjects = Server.CreateObject("ADODB.Recordset")
									tobeSent_number_of_subjects.ActiveConnection = Connect
									tobeSent_number_of_subjects.Source = "SELECT  Count(*) as number_of_subjects from subject_email  where email="&to_be_sent.Fields.Item("id").Value&" "
									tobeSent_number_of_subjects.CursorType = 0
									tobeSent_number_of_subjects.CursorLocation = 3
									tobeSent_number_of_subjects.LockType = 3
									tobeSent_number_of_subjects_numRows = 0
									tobeSent_number_of_subjects.Open()
									While (NOT tobeSent_number_of_subjects.EOF)
										tobeSent_numberSubjects=tobeSent_number_of_subjects.Fields.Item("number_of_subjects").Value
										tobeSent_number_of_subjects.MoveNext()
									Wend
									tobeSent_number_of_subjects.Close()
									
									'Response.Write("the number of subjects is "&tobeSent_numberSubjects)
									
									ReDim tobeSent_SubjectIDArray(tobeSent_numberSubjects) 
									ReDim tobeSent_SubjectNameArray(tobeSent_numberSubjects) 
									ReDim tobeSent_SubjectScoreArray(tobeSent_numberSubjects) 
									'Response.Write("size is "&UBound(subjectIDArray))
									Dim tobeSent_ii
									tobeSent_ii=0
									
									set tobeSent_email_subjects = Server.CreateObject("ADODB.Recordset")
									tobeSent_email_subjects.ActiveConnection = Connect
									tobeSent_email_subjects.Source = "SELECT *  from subject_email inner join subjects on subject_email.subject=subjects.ID_subject where email="&to_be_sent.Fields.Item("id").Value&" order by subject_name"
									tobeSent_email_subjects.CursorType = 0
									tobeSent_email_subjects.CursorLocation = 3
									tobeSent_email_subjects.LockType = 3
									tobeSent_email_subjects_numRows = 0
									tobeSent_email_subjects.Open()
									While (NOT tobeSent_email_subjects.EOF)
										
										'tobeSent_tempSubject=tobeSent_email_subjects.Fields.Item("subject_name").Value
										tobeSent_SubjectIDArray(tobeSent_ii)=tobeSent_email_subjects.Fields.Item("ID_subject").Value
										tobeSent_SubjectNameArray(tobeSent_ii)=tobeSent_email_subjects.Fields.Item("subject_name").Value
										'response.Write("name is "&tobeSent_email_subjects.Fields.Item("subject_name").Value)
										tobeSent_SubjectScoreArray(tobeSent_ii)="0"
										tobeSent_ii=tobeSent_ii+1
										tobeSent_email_subjects.MoveNext()
									Wend
									tobeSent_email_subjects.Close()
									
									
									tobeSent_number_of_certs=0
									set tobeSent_certification = Server.CreateObject("ADODB.Recordset")
									tobeSent_certification.ActiveConnection = Connect
									tobeSent_certification.Source = "select session_users,session_subject,subject_name,* from q_certification inner join q_session on q_session.ID_Session=q_certification.q_session inner join subjects on q_session.session_subject=subjects.ID_subject where passed=1 and expiry_date>GETDATE() and session_users="&to_be_sent.Fields.Item("ID_user").Value&" order by session_subject,expiry_date desc"
									tobeSent_certification.CursorType = 0
									tobeSent_certification.CursorLocation = 3
									tobeSent_certification.LockType = 3
									tobeSent_certification_numRows = 0
									tobeSent_tempSubject=0
									tobeSent_certification.Open()
									While (NOT tobeSent_certification.EOF)
										'due to the ordering of the query we know that each time a subject changes it will be their latest cert expiry
										'tobeSent_latestCert=false
										if(tobeSent_tempSubject<>tobeSent_certification.Fields.Item("session_subject").Value) then
											tobeSent_tempSubject=tobeSent_certification.Fields.Item("session_subject").Value
											For JJ = LBound(tobeSent_SubjectIDArray) To UBound(tobeSent_SubjectIDArray)-1
												if (tobeSent_SubjectIDArray(JJ)=tobeSent_tempSubject) then
													tobeSent_SubjectScoreArray(JJ)=tobeSent_certification.Fields.Item("percentage_achieved").Value
													tobeSent_number_of_certs=tobeSent_number_of_certs+1
												end if
											Next 
											
										end if
										tobeSent_certification.MoveNext()
									Wend
									tobeSent_certification.Close()
									
									
									%>
									<td class="text" align="left" valign="top">
												<%=tobeSent_number_of_certs%> out of <%=UBound(tobeSent_SubjectNameArray)%>
									</td>
									<td class="text" align="left" valign="top">
												<%
												
												For II = LBound(tobeSent_SubjectNameArray) To UBound(tobeSent_SubjectNameArray)-1
													Response.Write(tobeSent_SubjectNameArray(II) &" - "&tobeSent_SubjectScoreArray(II)&"%<br/>")
					
												Next 

												
												
												%>
									</td>
									<td class="text" align="left" valign="top">
												<%=tobeSent_emailType%>
									</td>
						<%
					end if
					to_be_sent.MoveNext()%>
					 
					</tr>
					<%
					Wend
					to_be_sent.Close()
					
			%>
			<tr> 
				<td class="subheads" colspan="5">&nbsp;</td>
			</tr>
			<tr> 
				<td class="subheads" colspan="5">&nbsp;</td>
			</tr>
			<tr> 
				<td class="subheads" colspan="5">Sent:</td>
			</tr>
			<tr> 
				<td  >Name</td>
				<td  >Sent date</td>
				<td  >Requested subjects completed</td>
				<td  >Subject scores</td>
				<td  >Email type</td>
			</tr>
			
           <%
					Dim subjectIDArray()  
					Dim subjectNameArray() 
					Dim subjectScoreArray() 
					sent.Open()
					While (NOT sent.EOF)
						got_subject2=false
						set filter_subject2 = Server.CreateObject("ADODB.Recordset")
						filter_subject2.ActiveConnection = Connect
						filter_subject2.Source = "SELECT  *  from subject_email  where email="&sent.Fields.Item("id").Value&"  "&subject_filter
						filter_subject2.CursorType = 0
						filter_subject2.CursorLocation = 3
						filter_subject2.LockType = 3
						filter_subject2_numRows = 0
						filter_subject2.Open()
						While (NOT filter_subject2.EOF)
							got_subject2=true
							filter_subject2.MoveNext()
						Wend
						filter_subject2.Close()
						
						if(got_subject2) then
								%>
								<tr class="table_normal" >
									<td class="text" align="left" valign="top">
												<%=sent.Fields.Item("user_firstname").Value%>&nbsp;<%=sent.Fields.Item("user_lastname").Value%>
									</td>
									<td class="text" align="left" valign="top">
												<%=sent.Fields.Item("date_to_send").Value%>
									</td>
												<%
									
									
									emailType=""
									if(sent.Fields.Item("type").Value=1) then
										emailType="Induction Welcome"
									elseif (sent.Fields.Item("type").Value=2) then
										emailType="Auto reminder"
									elseif (sent.Fields.Item("type").Value=3) then
										emailType="Escalation reminder"
									elseif (sent.Fields.Item("type").Value=4) then
										emailType="Final reminder"
									elseif (sent.Fields.Item("type").Value=5) then
										emailType="Induction Escalation"
									elseif (sent.Fields.Item("type").Value=6) then
										emailType="Induction Final"
									elseif (sent.Fields.Item("type").Value=7) then
										emailType="Manual Reminder"	
									end if
									
									numberSubjects=0
									set number_of_subjects = Server.CreateObject("ADODB.Recordset")
									number_of_subjects.ActiveConnection = Connect
									number_of_subjects.Source = "SELECT  Count(*) as number_of_subjects from subject_email  where email="&sent.Fields.Item("id").Value&" "
									number_of_subjects.CursorType = 0
									number_of_subjects.CursorLocation = 3
									number_of_subjects.LockType = 3
									number_of_subjects_numRows = 0
									number_of_subjects.Open()
									While (NOT number_of_subjects.EOF)
										numberSubjects=number_of_subjects.Fields.Item("number_of_subjects").Value
										number_of_subjects.MoveNext()
									Wend
									number_of_subjects.Close()
									
									'Response.Write("the number of subjects is "&numberSubjects)
									
									ReDim subjectIDArray(numberSubjects) 
									ReDim subjectNameArray(numberSubjects) 
									ReDim subjectScoreArray(numberSubjects) 
									'Response.Write("size is "&UBound(subjectIDArray))
									Dim ii
									ii=0
									
									set email_subjects = Server.CreateObject("ADODB.Recordset")
									email_subjects.ActiveConnection = Connect
									email_subjects.Source = "SELECT *  from subject_email inner join subjects on subject_email.subject=subjects.ID_subject where email="&sent.Fields.Item("id").Value&" order by subject_name"
									email_subjects.CursorType = 0
									email_subjects.CursorLocation = 3
									email_subjects.LockType = 3
									email_subjects_numRows = 0
									email_subjects.Open()
									While (NOT email_subjects.EOF)
										
										tempSubject=email_subjects.Fields.Item("subject_name").Value
										subjectIDArray(ii)=email_subjects.Fields.Item("ID_subject").Value
										subjectNameArray(ii)=email_subjects.Fields.Item("subject_name").Value
										subjectScoreArray(ii)="0"
										ii=ii+1
										email_subjects.MoveNext()
									Wend
									email_subjects.Close()
									
									
									number_of_certs=0
									set certification = Server.CreateObject("ADODB.Recordset")
									certification.ActiveConnection = Connect
									certification.Source = "select session_users,session_subject,subject_name,* from q_certification inner join q_session on q_session.ID_Session=q_certification.q_session inner join subjects on q_session.session_subject=subjects.ID_subject where passed=1 and expiry_date>GETDATE() and session_users="&sent.Fields.Item("ID_user").Value&" order by session_subject,expiry_date desc"
									certification.CursorType = 0
									certification.CursorLocation = 3
									certification.LockType = 3
									certification_numRows = 0
									tempSubject=0
									certification.Open()
									While (NOT certification.EOF)
										'due to the ordering of the query we know that each time a subject changes it will be their latest cert expiry
										latestCert=false
										if(tempSubject<>certification.Fields.Item("session_subject").Value) then
											tempSubject=certification.Fields.Item("session_subject").Value
											For J = LBound(subjectIDArray) To UBound(subjectIDArray)-1
												if (subjectIDArray(J)=tempSubject) then
													subjectScoreArray(J)=certification.Fields.Item("percentage_achieved").Value
													number_of_certs=number_of_certs+1
												end if
											Next 
											
										end if
										certification.MoveNext()
									Wend
									certification.Close()
									
									
									%>
									<td class="text" align="left" valign="top" >
												<%=number_of_certs%> out of <%=UBound(subjectNameArray)%>
									</td>
									<td class="text" align="left" valign="top">
												<%
												
												For I = LBound(subjectNameArray) To UBound(subjectNameArray)-1
													Response.Write(subjectNameArray(I) &" - "&subjectScoreArray(I)&"%<br/>")
					
												Next 

												
												
												%>
									</td>
									<td class="text" align="left" valign="top">
												<%=emailType%>
									</td>
									<%
						end if
						
					sent.MoveNext()
					%>
					 
					</tr>
					<%
					Wend
					sent.Close()
					
					%>
              
            </td>
          </tr>
		 
         

          <tr>
            <td class="text" align="left" valign="top"></td>
            <td class="text" align="left" valign="top" colspan="3">
              
			  
            </td>
          </tr>

          <tr>
            <td  align="left" valign="top">
              <!--input type="hidden" name="session" value="<%=getPassword(30, "", "true", "true", "true", "false", "true", "true", "true", "false")%>">
              <input type="hidden" name="current_export"-->
            </td>
            <td  align="left" valign="top" colspan="3">
              
              
             
            </td>
          </tr>
        </table>
       
     
	  <!--script type="text/javascript">
		  Calendar.setup(
			{
			  inputField  : "from",         // ID of the input field
			  ifFormat    : "%d/%m/%Y",    // the date format
			  button      : "trigger"       // ID of the button
			}
		  );
		</script-->

    </td>
  </tr>
</table>
<p>&nbsp;</p>
</BODY>
</HTML>

<%

call log_the_page ("Email report")
%>


