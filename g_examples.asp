<%@LANGUAGE="VBSCRIPT"%>
<% 
 Response.Buffer = True
 On Error Resume Next
 %>
<!-- #include file = "connections/bbg_conn.asp" -->
<!-- #include file = "connections/include.asp"-->


<%
if pref_bbg_login AND Session("userID") = "" then
session("aa") = Request.ServerVariables("URL") & "?" & Request.QueryString
	response.redirect("index.asp")
end if

IF NOT PREF_BBG_AVAIL THEN RESPONSE.REDIRECT("ERROR.ASP?" & REQUEST.QUERYSTRING)

ID_subject_prm = Session("id")


'TOPICS


SET topics = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT b_topics.ID_topic, b_topics.topic_name FROM subjects INNER JOIN ((b_hlp INNER JOIN (b_faq INNER JOIN b_topics ON b_faq.ID_faq = b_topics.topic_faq) ON b_hlp.ID_hlp = b_topics.topic_hlp) INNER JOIN b_pages ON b_topics.ID_topic = b_pages.page_topic) ON subjects.ID_subject = b_topics.topic_subject GROUP BY b_topics.ID_topic, b_topics.topic_ord, b_topics.topic_name, b_topics.topic_subject, Abs([subject_active_b]), Abs([topic_active]), Abs([page_active]) HAVING (((b_topics.topic_subject)=" + Replace(ID_subject_prm, "'", "''") + ") AND ((Abs([subject_active_b]))=1) AND ((Abs([topic_active]))=1) AND ((Abs([page_active]))=1)) order by b_topics.topic_ord;"
topics.Open SQL, Connect,3,3
if topics.EOF or topics.BOF then
	response.redirect("error.asp?" & request.QueryString)
ELSE
	IF Request.QueryString("ID_topic_prm")<>"" THEN
		ID_topic_prm = cInt(Request.QueryString("ID_topic_prm"))
	ELSE
		ID_topic_prm = cInt(topics("id_topic"))
	END IF
END IF

' INFORMATION ABOUT TOPIC
SET header = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT b_topics.ID_topic, b_topics.topic_name, b_topics.topic_title, b_topics.topic_keyp, b_topics.topic_exmp, b_faq.faq_tab, b_hlp.hlp_tab, b_topics.topic_training, b_topics.topic_qanda, b_topics.topic_subject, subjects.subject_name  FROM subjects INNER JOIN (b_hlp INNER JOIN (b_faq INNER JOIN b_topics ON b_faq.ID_faq = b_topics.topic_faq) ON b_hlp.ID_hlp = b_topics.topic_hlp) ON subjects.ID_subject = b_topics.topic_subject  WHERE (((b_topics.ID_topic)=" + Replace(ID_topic_prm, "'", "''") + ") AND ((Abs([topic_active]))=1));"
header.Open SQL, Connect,3,3
if header.EOF or header.BOF then response.redirect("error.asp?" & request.QueryString)

Dim exmessage
exmessage=Replace(RemoveHTML( header("topic_exmp"))," ","")
Dim exm
exm="Therearenoexamplesonthistopic."


Dim keymessage
keymessage=Replace(RemoveHTML( header("topic_keyp"))," ","")
Dim keym
keym="Therearenokeypointsonthistopic."
Function RemoveHTML( strText )
	Dim RegEx

	Set RegEx = New RegExp

	RegEx.Pattern = "<[^>]*>"
	RegEx.Global = True

	RemoveHTML = RegEx.Replace(strText, "")
End Function

%>
<!doctype html>
<head>
	
	<title><%=client_name_short%> - Guide - Better Business Program</title>
		<META name="DESCRIPTION"	content="">
		<!-- #include file = "inc_header.asp" -->
		<link href="perfect-scrollbar-0.4.8/src/perfect-scrollbar.css" rel="stylesheet">
       <script src="jquery-1.11.1.js?v=bbp34"></script>
      <script src="perfect-scrollbar-0.4.8/src/jquery.mousewheel.js?v=bbp34"></script>
      <script src="perfect-scrollbar-0.4.8/src/perfect-scrollbar.js?v=bbp34"></script>
	 <link rel="stylesheet" type="text/css" href="js/sweet-alert.css">
    		  <script src="js/sweet-alert.min.js?v=bbp34"></script>
</head>
<body>
	<div class="page-content">
		<div class="white-container">
			<!-- #include file = "partials/header.asp" -->
			<div class="allcontent h-100">
				<div class="allcontent_main h-100 m-0 w-100">
					<div class="menu_content" style="width: 230px;max-height: 100%;overflow-y: overlay;">

						<div class="header_blue" style="width:190px;height:50px;">
							<div class="header_inside" style="padding-left: 10px;padding-right: 10px;">SUBJECT<br>
								<h3 style="font-size:16px;">	<% =session("name")%></h3>
							</div>
						</div>
								<%WHILE NOT topics.EOF
								if clng(topics("ID_topic")) = clng(ID_topic_prm) THEN
								%>
								<div class="t_menu_blue_active" style="width: 190px;">
								<% ELSE%>
								<div class="t_menu_blue" style="width: 190px;">
								<% END IF%>
									<a  style="padding-left: 8px" href="g_index.asp?ID_topic_prm=<%=(topics("ID_topic"))%>"><%=ReplaceStrBBG(topics("topic_name"))%></a>
								</div>
								<%topics.MoveNext()
								Wend %>		
						<% if  Instr(exmessage, exm) = 0  or Instr(keymessage, keym) = 0  then  %>

						<div class="box_grey" style="width: 190px;">
						
							<div class="box_inside" style="padding: 8px;"><h2>Topic Highlights!</h2>
							<% if Instr(exmessage, exm) = 0 then %>		
							Click on the buttons below to see some real case examples of how this works in practice.
							<% else if Instr(keymessage, keym) = 0 then %>
							Click on the buttons below to see some key points on this topic.
							<% else %>		
							Click on the buttons below to see some real case examples of how this works in practice and the key points on this topic.
							<% end if %>
							<% end if %>
							
							<br>
					<% if  Instr(exmessage, exm) = 0 then  %>
								<div class="submit_grey">
									<div class="h_submit_grey"><a href="g_examples.asp?ID_topic_prm=<%=ID_topic_prm%>" class="box_link">EXAMPLES</a></div>
								</div>
					<% end if %>
					<% if  Instr(keymessage, keym) = 0 then  %>	
								<div class="submit_grey">
									<div class="h_submit_grey"><a href="g_keypoints.asp?ID_topic_prm=<%=ID_topic_prm%>" class="box_link">KEY POINTS</a></div>
								</div>
					<% end if %>
							</div>
						</div>
					<% end if %>
						<div class="clear"></div>
						</div>
						<div class="main_content mt-0 mb-0" style="width: 700px;margin-left: 20px;max-height: 100%;overflow-y: overlay;">
							<h1>Examples <%=ReplaceStrBBG(header("topic_title"))%></h1><br>
							<%IF header("topic_exmp") <> "" THEN response.write Highlight(ReplaceStrBBG(header("topic_exmp"))) ELSE response.write "There are no examples on this topic"%>
						
							<br><br>
								<div class="back_blue">
									<div class="h_submit_blue"><a href="javascript:history.back()" class="box_link" style="padding-left:65px;">BACK</a></div>
								</div>
						</div>
						<div class="clear"></div>
					</div>
				</div>
				<%
				call log_the_page ("Guide Examples", Session("id"), session("name"), "0", "n/a", "0", "n/a", "Guide Examples")
				%>
			</div>
			<!-- #include file = "partials/footer.asp" -->
		</div>
</body>
<!-- #include file = "errorhandler/index.asp"-->