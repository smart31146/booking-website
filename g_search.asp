<%@LANGUAGE="VBSCRIPT"%>
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



' INFORMATION
SET content = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT *  FROM b_pages,b_topics, subjects  WHERE page_topic = id_topic AND ID_subject=topic_subject AND subject_active_b=1 AND (((b_pages.page_header) NOT LIKE '') AND ( replace(cast(dbo.b_pages.page_title as nvarchar(4000)),'_client_short_name_','"&client_name_short&"')  LIKE ? OR replace(cast(dbo.b_pages.page_header as nvarchar(4000)),'_client_short_name_','"&client_name_short&"') LIKE ? OR  b_pages.page_text LIKE ?) AND ((Abs([page_active]))=1))  ORDER BY b_pages.page_ord, b_pages.ID_page;"
	set objCommand = Server.CreateObject("ADODB.Command") 
	objCommand.ActiveConnection = Connect
	objCommand.CommandText = SQL
	objCommand.Parameters(0).value = "%"&Session("bbp_search")&"%"
	objCommand.Parameters(1).value = "%"&Session("bbp_search")&"%"
	objCommand.Parameters(2).value = "%"&Session("bbp_search")&"%"
	Set content= objCommand.Execute()

%>
<!doctype html>
<head>
		<title><%=client_name_short%> - Guide - Building a Better Workplace</title>
		<script src="jquery-1.11.1.js?v=bbp34"></script>
			<script src="js/freewall.js?v=bbp34"></script>
		<script src="js/modernizr-latest.js?v=bbp34"></script>
			<link rel="stylesheet" type="text/css" href="js/sweet-alert.css">
    		  <script src="js/sweet-alert.min.js?v=bbp34"></script>
		<!-- <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-T3c6CoIi6uLrA9TneNEoa7RxnatzjcDSCmG1MXxSR1GAsXEV/Dwwykc2MPK8M2HN" crossorigin="anonymous"> -->
		<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-C6RzsynM9kWDrMNeT87bh95OGNyZPhcTNXj1NW7RuBCsyN/o0jlpcV8Qyq46cDfL" crossorigin="anonymous"></script>
		
		<!-- <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js?v=bbp34"></script>
		<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/css/bootstrap.min.css" /> -->
		<META name="DESCRIPTION"	content="">
		<!-- #include file = "inc_header.asp" -->
		<style type="text/css">
			.searchHitWord {font-size: 14px;background: #C4E0F2;font-weight: bold;text-decoration:underline;}
			.gray { color: #999999; font-size: 14px; font-weight: bold; }
		</style>
</head>
<body>
	<div class="page-content">
		<div class="white-container">
			<!-- #include file = "partials/header.asp" -->
			<div class="allcontent h-100">
				<div class="allcontent_main h-100">
					<div class="menu_content" style="width: 200px;">
						<div class="header_blue" style="width:200px;">
							<div class="header_inside"><h3 style="font-size:18px;">	Search result for<br>
							'<% =Server.HTMLEncode(Session("bbp_search"))%>'</h3>
							</div>
						</div>
					</div>
					<div class="main_content mt-0 mb-0" style="width: 700px;margin-left: 20px;max-height: 100%;overflow-y: overlay;">
					<% IF content.eof then%><br><h1>Your search has returned no results</h1>
					<%ELSE
					do until content.EOF%>
						<h1><%=content("topic_name")%> / <%if (content("page_title")) <> "" then response.write(Highlight(ReplaceStrSearch(content("page_title"))))%></h1>
						
						<%=HighlightWord(ReplaceStrSearch(content("page_header"))) & ""%>
						<br><br><%=HighlightWord(ReplaceStrSearch(content("page_text"))) & ""%><br><img src="images/g_line.png" width="600" height="28" alt=""><br>
						
						
					<%
					content.MoveNext()
					loop
					END IF
					%>
					</div>
				</div>
			</div>
		</div>
		<!-- #include file = "partials/footer.asp" -->
		<span style="color: white">
			<%
				call log_the_page ("Guide Search", Session("id"), session("name"), "0", "n/a", "0", "n/a", "Guide Search")
			%>
		</span>
	</div>
</body>

<% Function HighlightWord(ord)
    Dim b
    for b = 0 to arrSokordLength
        ord = replace(ord,ReplaceStrSearch(Session("bbp_search")),"<span class=""searchHitWord"">"& ReplaceStrSearch(Ucase(Session("bbp_search"))) &"</span>",1,-1,1)
    next
    HighlightWord = ord
End Function%>