<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%

' SET certifiation perion in days
cert_period = 365

filter_cert = 0
If cInt(Request.Querystring("filter_cert")) <> 0 then filter_cert = cInt(Request.Querystring("filter_cert"))

filter_name = ""
If trim(Request.Querystring("filter_name")) <> "" then filter_name = trim(Request.Querystring("filter_name"))

filter_info1_prm = 0
If cInt(Request.Querystring("filter_info1")) <> 0 then	filter_info1_prm = cInt(Request.Querystring("filter_info1"))

filter_info2_prm = 0
If cInt(Request.Querystring("filter_info2")) <> 0 then 	filter_info2_prm = cInt(Request.Querystring("filter_info2"))

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
set filter_info3 = Server.CreateObject("ADODB.Recordset")

%><HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">

</HEAD>
<body>
<% Response.Clear()
Response.AddHeader "Content-Disposition","attachment; filename=Certification_Report_" & day(now()) & "_" & month(now()) & "_" & year(now()) & ".xls"
Response.ContentType="application/vnd.ms-excel"

%>
<style media="all" type="text/css">
table, tr, td {
	font-size:11px;
	font-family:arial;
}

</style>
<TABLE BORDER="1" width="95%">
	<tr>
		<td valign="top"FILTER=ALL>Reference</td>
		<td valign="top"FILTER=ALL>Last name &amp; First name</td>
		<td valign="top"FILTER=ALL>Certified?</td>
		<td valign="top"FILTER=ALL>Earliest subject expires</td>
		<td valign="top"FILTER=ALL>Outstanding subjects</td>
		<td valign="top"FILTER=ALL>Business</td>
		<td valign="top"FILTER=ALL>Site</td>
		<td valign="top"FILTER=ALL>Activity</td>
	</tr> 
   
<%   
resFilters = " "
if IsNumeric(request("filter_info1")) AND Cint(request("filter_info1")) <> 0 then    resFilters = resFilters & " AND user_info1 = "& request("filter_info1") &" "
if IsNumeric(request("filter_info1")) AND Cint(request("filter_info2")) <> 0 then    resFilters = resFilters & " AND user_info2 = "& request("filter_info2") &" "
if IsNumeric(request("filter_info1")) AND Cint(request("filter_info3")) <> 0 then    resFilters = resFilters & " AND user_info3 = "& request("filter_info3") &" "
if filter_name <> "" THEN    resFilters = resFilters & " AND (user_lastname LIKE '"& filter_name &"%' OR user_firstname LIKE '"& filter_name &"%') "

'get id, name and passmark of all active subjects

set RSActiveUsers = Server.CreateObject("ADODB.Recordset") :RSActiveUsers.ActiveConnection = Connect
RSActiveUsers.Source = "SELECT ID_user, user_lastname, user_firstname, info1, info2, info3, user_reference  FROM dbo.q_user LEFT JOIN  q_info1 ON ID_info1 = user_info1 LEFT JOIN  q_info2 ON ID_info2 = user_info2 LEFT JOIN  q_info3 ON ID_info3 = user_info3 WHERE dbo.q_user.user_active = 1 "&resFilters&" ORDER BY user_lastname "
'response.write RSActiveUsers.Source 
RSActiveUsers.CursorType = adOpenForwardOnly : RSActiveUsers.CursorLocation = 3 : RSActiveUsers.LockType = 3 : RSActiveUsers.Open()  
totalUsers = RSActiveUsers.RecordCount
totalCertified = 0
if not RSActiveUsers.eof then arrV = RSActiveUsers.GetRows ELSE arrV = -1
RSActiveUsers.close 


'set RSActiveSubjects = Server.CreateObject("ADODB.Recordset")	: RSActiveSubjects.ActiveConnection = Connect
'RSActiveSubjects.Source = "SELECT subjects.ID_subject, subject_name, subject_passmark, subject_expiry  FROM subjects WHERE subjects.subject_active_q = 1"
'RSActiveSubjects.CursorType = adOpenForwardOnly : RSActiveSubjects.CursorLocation = 3 : RSActiveSubjects.LockType = 3 : RSActiveSubjects.Open() 
'if not RSActiveSubjects.eof then arrS = RSActiveSubjects.GetRows ELSE arrS = -1
'RSActiveSubjects.close 
				
counterI = 1
IF IsArray(arrV) THEN
	For i = 0 to ubound(arrV,2)


    certifiedSubjects = 0
    earliestExpires = DateAdd("yyyy", 10, Date())
    outstandingSubjects = ""
    isCertified = "Yes"
    userId = arrV(0,i)
	prevCat = ""
	prevCorrect = 0
	currentCorrect = 1
	'response.write RSActiveUsers("info2")
	'response.end
    IF arrV(3,i)<>"" THEN tblBusiness = arrV(3,i) ELSE tblBusiness = "n/a"
    IF arrV(4,i)<>"" THEN tblSite = arrV(4,i) ELSE tblSite = "n/a"
    IF arrV(5,i)<>"" THEN tblActivity = arrV(5,i) ELSE tblSite = "n/a"

				'set RSActiveSubjects = Server.CreateObject("ADODB.Recordset")	: RSActiveSubjects.ActiveConnection = Connect
				'RSActiveSubjects.Source = "SELECT subjects.ID_subject, subject_name, subject_passmark, subject_expiry  FROM subject_user,subjects WHERE id_user = "&clng(userId)&" AND subject_user.id_subject = subjects.id_subject AND subjects.subject_active_q = 1 ORDER BY subject_name"
				'RSActiveSubjects.CursorType = adOpenForwardOnly : RSActiveSubjects.CursorLocation = 3 : RSActiveSubjects.LockType = 3 : RSActiveSubjects.Open() 
				'if not RSActiveSubjects.eof then arrS = RSActiveSubjects.GetRows ELSE arrS = -1
				'RSActiveSubjects.close 
				
				set RSActiveSubjects = Server.CreateObject("ADODB.Recordset")	: RSActiveSubjects.ActiveConnection = Connect
				RSActiveSubjects.Source = "SELECT subjects.ID_subject, subject_name, subject_passmark, subject_expiry,q_session.Session_finish, session_correct, session_total   " & _
										"FROM subjects,subject_user " & _
										"LEFT JOIN q_session ON (session_subject = id_subject AND Session_users = "&userId&" AND session_finish BETWEEN '"&cDateSql(now()-cert_period)&"' AND '"&cDateSql(now())&"' )" & _
										"WHERE subjects.id_subject = subject_user.id_subject AND subject_user.id_user = "&userId&" AND subjects.subject_active_q = 1 ORDER BY subject_name,session_correct desc,Session_finish desc"
				'response.write RSActiveSubjects.Source & "<br>"
				RSActiveSubjects.CursorType = adOpenForwardOnly : RSActiveSubjects.CursorLocation = 3 : RSActiveSubjects.LockType = 3 : RSActiveSubjects.Open() 
				if not RSActiveSubjects.eof then arrS = RSActiveSubjects.GetRows ELSE arrS = -1
				RSActiveSubjects.close 
			    
				IF IsArray(arrS) THEN
				
					For y = 0 to ubound(arrS,2)
	
						currentPassmark = arrS(2,y)
				        currentSubject = arrS(0,y)
				        currentSubjectName = arrS(1,y)
						currentCorrect = arrS(5,y)
	  
						'set RSrecentQuiz = Server.CreateObject("ADODB.Recordset")
				        'RSrecentQuiz.ActiveConnection = Connect
				        'RSrecentQuiz.Source = "SELECT TOP 1 q_session.Session_finish, session_correct, session_total FROM q_session WHERE session_finish BETWEEN '"&cDateSql(now()-365)&"' AND '"&cDateSql(now())&"' AND Session_users = "&userId&" AND session_subject = "&currentSubject&" ORDER BY Session_correct DESC"
						'response.write RSrecentQuiz.Source & "<br>"
					'	response.end
				       ' RSrecentQuiz.CursorType = adOpenForwardOnly : RSrecentQuiz.CursorLocation = 3 : RSrecentQuiz.LockType = 3 : RSrecentQuiz.Open()
						'if not RSrecentQuiz.eof then arrQ = RSrecentQuiz.GetRows ELSE arrQ = -1
						'RSrecentQuiz.close 
				if prevCat <> currentSubjectName THEN
					prevCat = currentSubjectName
					prevCorrect = currentCorrect
						IF arrS(4,y) <> "" THEN
			                IF ((clng(arrS(5,y)) / clng(arrS(6,y)) )*100) >= clng(currentPassmark) THEN
			           			finDate = CDate(arrS(4,y))
			           			quizExpires = DateAdd("d", (arrS(3,y) * 7), finDate)
			                    earliestExpires = quizExpires
							ELSE
			                	'subject out of date
			              	  outstandingSubjects = outstandingSubjects & currentSubjectName & vbcrlf & vbcrlf
							END IF
						ELSE
							 outstandingSubjects = outstandingSubjects & currentSubjectName & vbcrlf & vbcrlf
						END IF
				END IF
						
				   	NEXT
					
					Erase arrS
				END IF
						' arrS next
				    If (outstandingSubjects <> "") then      
				        isCertified = "No"
				        earliestExpires = "Not certified"
				        cutTo = Len(outstandingSubjects) - 2
				        'chop off " ," if exists         
				        outstandingSubjects = Mid(outstandingSubjects , 1 , cutTo )
				    else
				        isCertified = "Yes"
				        outstandingSubjects = "None"
				        'format date to look nicer
				        earliestExpires = FormatDateTime(earliestExpires,2)
				        totalCertified = totalCertified + 1
				    end if
					
		showUser = False
					
		'response.write isCertified
		IF filter_cert = 0 THEN showUser = True
		IF isCertified = "Yes" AND filter_cert = 1 THEN showUser = True
		IF isCertified = "No" AND filter_cert = 2 THEN showUser = True
		IF showUser = True THEN
    %>    
	
		<tr>
            <td valign="top" align="left"><%=arrV(6,i)%></td>
			<td valign="top" align="left"><%=arrV(1,i) & ", " & arrV(2,i)%></td>
			<td valign="top" align="left"><% if isCertified = "Yes" then response.write("YES") else response.write("NO") %></td>
            <td valign="top" align="left"><%=earliestExpires%></td>
            <td valign="top" align="left"><%=outstandingSubjects%></td>
            <td valign="top" align="left"><%=tblBusiness%></td>
            <td valign="top" align="left"><%=tblSite%></td>
            <td valign="top" align="left"><%=tblActivity%></td>
        </tr>
		
		

<% 

counterI = counterI+1
 END IF
   NEXT
   Erase arrV
%>
</table>



				<% 
   END IF%>
		
		

<%
'call log_the_page ("Certification List Users")

 Set Connect = Nothing
%>

