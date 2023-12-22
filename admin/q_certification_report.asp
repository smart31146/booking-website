<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->
<%
'fixes the all upper case names entered into the database
function upfix(a)
    upfix = ucase(left(a,1)) & lcase(mid(a,2))
end function

' SET certifiation perion in days
'cert_period = 365
' 1 year by Jhan Bredenholt
'cert_periodsql = cDateSql(dateadd("yyyy",-1,now()))


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


SQL = "SELECT * FROM q_info1 where info1_active=1 order by info1"
filter_info1.Open SQL, Connect, 3,3

SQL = "SELECT * FROM q_info3 where info3_active=1 order by info3"
filter_info3.Open SQL, Connect, 3,3

%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<TITLE>BBP ADMIN: Quiz users. You are logged in as <%=Session("MM_Username_admin")%></TITLE>
<link rel="stylesheet" href="styles/adminquizstyle.css" type="text/css">
<script language="JavaScript">
<!--
function pviiClassNew(obj, new_style) {
    obj.className = new_style;
}

function filter_submit()
{
		//alert(document.filter_users.filter_name.value.length);
  		//if (document.filter_users.filter_name.value.length < 3 && document.filter_users.filter_name.value.length > 0 )
  		//	{
  		 //alert ("First or Last Name must be at least 3 character") ;
		 //submitbutton.style.display='block';loadingr.style.display= 'none';
  		 // return false;
  		//	}
		document.forms[0].submit();
		return true;
}
function checkform() {
	document.forms[0].action="q_certification_report.asp"
	document.forms[0].target="_self"
	document.forms[0].submit()
}
//-->
</script>
<style media="all" type="text/css">
.table_normal {
	background: #FFCC66;
	font-size:11px;
}
.table_normal td {
	border-right: 1px solid #F2C162;
	border-bottom: 1px solid #F2C162;
	padding: 3px 3px 3px 6px;
	color: #000000;
	font-size:11px;
	font-family:arial;
}
.table_normal_over {
	background: #FFE867;
	font-size:11px;
}
.table_normal_over td {
	border-right: 1px solid #F2C162;
	border-bottom: 1px solid #F2C162;
	padding: 3px 3px 3px 6px;
	color: #000000;
	font-size:11px;
	font-family:arial;
}

</style>
</HEAD>

<BODY BGCOLOR=#FFCC00 TEXT=#000000 VLINK=#000000 LINK=#000000 leftmargin="0" topmargin="0">
<span class="heading">Certification report</span><br>
	<table>
	<table>
	<form name="filter_users" method="get">
        <table>
			
				<tr>
					<td class="subheads" align="left" valign="top" width="200">Users:</td>
					<td>&nbsp;</td>
					<td align="right" class="subheads" valign="top" style="text-align: right" ><a href="q_export_certification.asp?filter_info1=<% =trim(request("filter_info1"))%>&amp;filter_info2=<% =trim(request("filter_info2"))%>&amp;filter_info3=<% =trim(request("filter_info3"))%>&amp;filter_cert=<% =trim(request("filter_cert"))%>&filter_name=<% =trim(request("filter_name"))%>"><img src="images/xls.gif" width="16" height="16" border="0" style="vertical-align:middle;"> Export to excel</a></td>
				</tr>
				<tr class="table_normal">
					<td><img src="images/back.gif" width="18" height="14"> <a href="main.asp">...Home page</a></td>
					<td colspan="2"><a href="q_certification_report.asp">... Clear certification report</a></td>
				</tr>
				<tr>
					<td class="subheads" colspan="3">Filter users by:</td>
				</tr>
				<tr class="table_normal">
					<td valign="top">Business:</td>
					<td valign="top" colspan="2">
					<select name="filter_info1" class="formitem1" onchange=checkform();>
						<option value="0">--- select a business ---</option>
						<% While (NOT filter_info1.EOF)%>
						<option value="<%=(filter_info1.Fields.Item("ID_info1").Value)%>" <%if (CStr(filter_info1.Fields.Item("ID_info1").Value) = CStr(request.querystring("filter_info1"))) then Response.Write("SELECTED") : Response.Write("")%>><%=(filter_info1.Fields.Item("info1").Value)%></option>
						<%  filter_info1.MoveNext()
						Wend
						filter_info1.Requery
						%>
					</select>
					</td>
				</tr>
				<tr class="table_normal">
					<td valign="top">Site:</td>
					<td valign="top" colspan="2">
					  <select name="filter_info2" class="formitem1">
						<option value="0">--- select a business site---</option>
						<% While (NOT filter_info2.EOF)	%>
						<option value="<%=(filter_info2.Fields.Item("ID_info2").Value)%>" <%if (CStr(filter_info2.Fields.Item("ID_info2").Value) = CStr(request.querystring("filter_info2"))) then Response.Write("SELECTED") : Response.Write("")%>><%=(filter_info2.Fields.Item("info2").Value)%></option>
						<%
						filter_info2.MoveNext()
						Wend
						
						filter_info2.Requery
						%>
					  </select>
					</td>
				</tr>
				<tr class="table_normal">
					<td valign="top">Activity:</td>
					<td valign="top" colspan="2">
					  <select name="filter_info3" class="formitem1">
						<option value="0">--- select a business activity ---</option><%
						While (NOT filter_info3.EOF)
						%><option value="<%=(filter_info3("ID_info3"))%>" <%if (CStr(filter_info3.Fields.Item("ID_info3").Value) = CStr(request.querystring("filter_info3"))) then Response.Write("SELECTED") : Response.Write("")%>><%=(filter_info3.Fields.Item("info3").Value)%></option><%
						  filter_info3.MoveNext()
						Wend
						filter_info3.Requery %>
					  </select>
					</td>
				</tr>
				<tr class="table_normal">
					<td valign="top">Certification:</td>
					<td valign="top" colspan="2">
						<select name="filter_cert" class="formitem1">
							<option value="0"<% if filter_cert = 0 THEN response.write " SELECTED"%>> All
							<option value="1"<% if filter_cert = 1 THEN response.write " SELECTED"%>> Yes
							<option value="2"<% if filter_cert = 2 THEN response.write " SELECTED"%>> No
						</select>
					</td>
				</tr>
				<tr class="table_normal">
					<td valign="top">First or Last name:</td>
					<td valign="top" colspan="2">
					<input type="text" name="filter_name" style="width:200px;" size="60" class="formitem1" value="<% =filter_name%>">
					</td>
				</tr>
				<tr>	
					<TD height="30">&nbsp;</TD>
					<td height="30" align="left" class="text" colspan="2">
					<input type="Submit" name="Submit" value="&gt;&gt;&gt; Filter users &lt;&lt;&lt;" class="quiz_button" id="submitbutton" onclick="this.style.display='none';loadingr.style.display= 'block';return filter_submit();">
					<img  src="images/loading.gif" alt="" width="24"  border="0"  name ="loadingr"  id ="loadingr" style ="display: none;"></td>
				</tr>
			</table>
			<br>
			<% 
			filter_info1.Close()
			filter_info3.Close()
			IF  request("filter_info1") <> "" OR filter_name <> "" THEN
			'IF (request("filter_info1")<>"0" AND request("submit")<>"")   OR filter_name <> "" THen
			%>
			<table>
				<tr>
					<td class="head">&nbsp;</td>
					<td class="head">Reference</td>
					<td class="head">Last name &amp; First name</td>
					<td class="head">Certified?</td>
					<td class="head">Earliest subject expires</td>
					<td class="head">Outstanding subjects</td>
					<td class="head">Business</td>
					<td class="head">Site</td>
					<td class="head">Activity</td>
				</tr> 

<% 
resFilters = " "
if IsNumeric(request("filter_info1")) AND Cint(request("filter_info1")) <> 0 then    resFilters = resFilters & " AND user_info1 = "& request("filter_info1") &" "
if IsNumeric(request("filter_info1")) AND Cint(request("filter_info2")) <> 0 then    resFilters = resFilters & " AND user_info2 = "& request("filter_info2") &" "
if IsNumeric(request("filter_info1")) AND Cint(request("filter_info3")) <> 0 then    resFilters = resFilters & " AND user_info3 = "& request("filter_info3") &" "
if filter_name <> "" THEN    resFilters = resFilters & " AND (user_lastname LIKE '%"& filter_name &"%' OR user_firstname LIKE '%"& filter_name &"%' ) "

'get id, name and passmark of all active subjects

set RSActiveUsers = Server.CreateObject("ADODB.Recordset") :RSActiveUsers.ActiveConnection = Connect
RSActiveUsers.Source = "SELECT ID_user, user_lastname, user_firstname, info1, info2, info3, user_reference    FROM dbo.q_user LEFT JOIN  q_info1 ON ID_info1 = user_info1 LEFT JOIN  q_info2 ON ID_info2 = user_info2 LEFT JOIN  q_info3 ON ID_info3 = user_info3 WHERE dbo.q_user.user_active = 1 "&resFilters&" ORDER BY user_lastname "
RSActiveUsers.CursorType = adOpenForwardOnly : RSActiveUsers.CursorLocation = 3 : RSActiveUsers.LockType = 3 : RSActiveUsers.Open()  
totalUsers = RSActiveUsers.RecordCount
totalCertified = 0
if not RSActiveUsers.eof then arrV = RSActiveUsers.GetRows ELSE arrV = -1
RSActiveUsers.close 

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
	
    IF arrV(3,i)<>"" THEN tblBusiness = arrV(3,i) ELSE tblBusiness = "n/a"
    IF arrV(4,i)<>"" THEN tblSite = arrV(4,i) ELSE tblSite = "n/a"
    IF arrV(5,i)<>"" THEN tblActivity = arrV(5,i) ELSE tblSite = "n/a"

				set RSActiveSubjects = Server.CreateObject("ADODB.Recordset")	: RSActiveSubjects.ActiveConnection = Connect
				
'				RSActiveSubjects.Source = "SELECT subjects.ID_subject, subject_name, subject_passmark, subject_expiry, q_session.Session_finish, session_correct, session_total, subject_user.ID_user  " & _
'										"FROM subject_user,subjects " & _
'										"LEFT JOIN q_session ON (session_subject = id_subject AND Session_users = "&userId&" AND session_finish BETWEEN DATEADD([week], - subject_expiry, GETDATE()) AND GETDATE() )" & _
'										"WHERE id_user = "&userId&" AND subject_user.id_subject = subjects.id_subject AND subjects.subject_active_q = 1 ORDER BY subject_name,Session_finish desc"
				
				RSActiveSubjects.Source = "SELECT subjects.ID_subject, subject_name, subject_passmark, subject_expiry, q_session.Session_finish, session_correct, session_total, subject_user.ID_user  " & _
										"FROM subject_user,subjects " & _
										"LEFT JOIN q_session ON (session_subject = id_subject AND Session_users = "&userId&" AND session_finish BETWEEN DATEADD([week], - subject_expiry, GETDATE()) AND GETDATE() and ((session_correct / CAST(session_total AS DECIMAL(10, 2))) * 100) >= subject_passmark ) "  & _
										"WHERE id_user = "&userId&" AND subject_user.id_subject = subjects.id_subject AND subjects.subject_active_q = 1 ORDER BY subject_name,Session_finish"
				
				RSActiveSubjects.CursorType = adOpenForwardOnly : RSActiveSubjects.CursorLocation = 3 : RSActiveSubjects.LockType = 3 : RSActiveSubjects.Open() 
								
				if not RSActiveSubjects.eof then arrS = RSActiveSubjects.GetRows ELSE arrS = -1
				RSActiveSubjects.close
			    PrevDate = cdate("4/06/2030")
				prevCat = ""
				prevCorrect = 0
			    latestExpire = cdate("4/06/2100")
				IF IsArray(arrS) THEN
				
					For y = 0 to ubound(arrS,2)
					takeThis = false
					currentCorrect = arrS(5,y)
					IF arrS(4,y) <> "" THEN
					IF prevCat = currentSubjectName AND PrevDate < cdate(arrS(4,y)) AND ((clng(arrS(5,y)) / clng(arrS(6,y)) )*100) >= clng(currentPassmark)    then
						takeThis = true
					END IF
					END IF
					currentPassmark = arrS(2,y)
			        'currentSubject = arrS(0,y)
			        currentSubjectName = arrS(1,y)
					'currentDate = arrS(4,y)
					Subject_userID = arrS(7,y)
						  
				if  takeThis = true OR prevCat <> currentSubjectName THEN
					prevCat = currentSubjectName
					prevCorrect = currentCorrect
						IF arrS(4,y) <> "" THEN
			                IF ((clng(arrS(5,y)) / clng(arrS(6,y)) )*100) >= clng(currentPassmark) THEN
			           			finDate = CDate(arrS(4,y))
								subject_expiry = arrS(3,y)
			           			'quizExpires = DateAdd("yyyy",1, finDate)
								quizExpires = DateAdd("ww",subject_expiry,finDate)
								
								IF cdate(latestExpire) > cdate(arrS(4,y)) AND PrevDate < cdate(arrS(4,y)) THEN
									latestExpire = cdate(arrS(4,y))
									else
									earliestExpires = quizExpires
								END IF
			                   	
							ELSE
			                	'subject out of date
			              	  outstandingSubjects = outstandingSubjects & currentSubjectName & "<br>"
							END IF
						ELSE
							 outstandingSubjects = outstandingSubjects & currentSubjectName & "<br>"
						END IF
				END IF
						
					PrevDate = arrS(4,y)
				   	NEXT
					
					Erase arrS
					
					
				END IF
						' arrS next
				    If (outstandingSubjects <> "" or userId<>Subject_userID) then      
				        isCertified = "No"
				        earliestExpires = "Not certified"
				        cutTo = Len(outstandingSubjects) - 2
				        'chop off " ," if exists         
				        'if userId=Subject_userID then
						'	outstandingSubjects = Mid(outstandingSubjects , 1 , cutTo )
						'end if
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
			<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_normal_over')" onMouseOut="pviiClassNew(this,'table_normal')">
				<td width="20"><% =counterI %></td>
				<td width="30"><% =arrV(6,i)%></td>
				<td><a href="q_user_sessions.asp?user=<%=userId%>&subject=0&filter_info1=0&filter_info3=0&active=2&results=2&passrate=75&mths=1"><%=upfix(arrV(1,i)) & ", " & upfix(arrV(2,i))%></a></td>
				<td><% if isCertified = "Yes" then response.write("<img src=""images/cert_yes.png"" alt=""Yes"">") else response.write("<img src=""images/cert_no.png"" alt=""No"">") %></td>
				<td><%=earliestExpires%></td>
				<td><%=outstandingSubjects%></td>
				<td><%=tblBusiness%></td>
				<td width="100"><%=tblSite%></td>
				<td width="140"><%=tblActivity%></td>
			</tr> 
			<%counterI = counterI+1
			 END IF
			   NEXT
			   Erase arrV
			%>
			</table>
			<br>
			<% IF filter_cert = 0 THEN %>
			<table>
				<tr>
					<td class="subheads" colspan="9">Overall results of users on this page: </td>
				</tr>
				<tr class="table_normal" valign="top">
					<td align="left">Total number of active users</td>
					<td align="left">Users who are</td>
					<td align="left"> <font color = green>certified</font></td>
					<td align="left"><font color = red>uncertified</font></td>
				</tr>
				<tr class="table_normal">
					<td><%=totalUsers%></td>
					<td>&nbsp;</td>
					<td><font color = green><%=totalCertified %></font></td>
					<td><font color = red><%=(totalUsers - totalCertified) %></font></td>
				</tr>
					<% If totalUsers = 0 Then %>
				<tr>
					<td  width="18">&nbsp;<input type="hidden" class="formitem1" name="passrate"  size=3 value=""></td>
					<td colspan="8" >Sorry,
						there are no users in the quiz currently or no user match your filter
						criteria.
					</td>
				</tr>
					  <% 'End If %>
			</table>
			<br>
			<% End If %>
			<table>
				<tr><td width=150 class="head">Subject</td><td class="head">Expiry period <i>(in weeks)</i></td><td class="head">Pass rate</td></tr>
				<% set subjects = Server.CreateObject("ADODB.Recordset")
				SQL = "SELECT ID_subject, subject_name, subject_passmark, subject_expiry FROM subjects where subject_active_q <> 0"
				subjects.Open SQL, Connect, 3,3
				While (NOT subjects.EOF) %>
				<tr class="table_normal" onMouseOver="pviiClassNew(this,'table_normal_over')" onMouseOut="pviiClassNew(this,'table_normal')">
					<td><%=(subjects("subject_name"))%></td>
					<td><%=(subjects("subject_expiry"))%></td>
					<td><%=(subjects("subject_passmark"))%></td>
				</tr>
				<% subjects.MoveNext()
				Wend 
				subjects.Close() 
				
				END IF%>
			</table>
			</td>
          </tr>
    </form>
</table>
</table>
<% end if
END IF%>
<p>&nbsp;</p>
</BODY>
</HTML>
<%
'call log_the_page ("Certification List Users")
 Set Connect = Nothing
%>

