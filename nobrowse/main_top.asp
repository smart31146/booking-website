	<div class="header">
		<div class="header_main">
		<div class="header_logo">
		<% IF not bbp_training = True THEN%>
		<% IF Session("userID") <> "" AND Session("id") <> "" THEN%>
			<div style="position:absolute;right:30px;top:10px;color:#fff;">Search guide&nbsp;&nbsp;<form action="g_search.asp" method="post" name="searchform" onsubmit="return checkSearchField();"><input type="text" onfocus="if (this.value == '<% =Session("bbp_search")%>') {this.value = '';}" onblur="if (this.value.length == 0) { this.value = '<% =Session("bbp_search")%>' }" style="width:130px;vertical-align:middle;" value="<% =Session("bbp_search")%>" name="bbp_search"> <input type="Image" src="images/search_button.png" value="Search" style="vertical-align:middle;"></form></div>
			<!--<div style="position:absolute;right:70px;top:72px;color:#777; font-size: 1.3em; font-weight: bold; "><i>Better Business Program</i></div>-->
		<% END IF %>
		<% END IF %>
		</div>
		
		<div id="bbp_menu">
		
		<ul class="menu">
		<% IF Session("userID") <> "" AND Session("id") <> "" THEN%>
			<li  style="float:left;margin-right:0px;"><a href="t_index.asp" title="TRANING">TRAINING</a></li>
			<li style="margin-right: -10px;"><a href="g_index.asp" title="GUIDE">GUIDE</a></li>
			<li style="margin-right:0px;width: 120px;"><a href="javascript:" style="width:100px" onClick="window.name='bbgwindow'; window.open('help.asp','feedbackwindow','scrollbars=yes,resizable=yes,top=100,left=100,width=700,height=800')">HELP/FEEDBACK</a></li>
			<!--<li style="float:left;margin-left: 20px;width: 155px;"><a href="#" onClick="window.open('user_sessions_new.asp?user=<%=Session("userID")%>&latest=1','newwindow','scrollbars=yes,resizable=yes, width=920,height=624,left=50,top=50')">PREVIOUS RESULTS</a></li>-->
		<% END IF%>
		<% IF bbp_training = True THEN%>
			<li style="float:right;margin-right:30px;"><a href="index.asp" onclick="return confirm('Are you sure you want to exit the training?\n\nYour records will be saved and you can\ncontinue from this place later on.')">EXIT TRAINING</a></li>
			
		<%ELSE%>
		<%IF Session("userID") <> "" THEN %>
				<li style="float:left;margin-left: 0px;"><a href="#" onClick="window.open('user_sessions_new.asp?user=<%=Session("userID")%>&latest=1','newwindow','scrollbars=yes,resizable=yes, width=920,height=624,left=50,top=50')">PREVIOUS RESULTS</a></li>
			
		<%END IF%>
		

			<% IF Session("userID") <> "" AND Session("id") <> "" THEN%>
			<li style="float:right;margin-right:0px;"><a href="index.asp?alt=logout" onclick="return confirm('Are you sure you want to log out?')">LOG OUT</a></li>
			
			<% SQL = "SELECT subject_user.ID_subject,subject_name FROM subject_user,subjects WHERE subject_user.id_subject = subjects.id_subject AND id_user='"& fixstr(Session("UserID")) &"' AND subjects.subject_active_q=1;"
			obj.Open SQL, Connect,1,1%>
			<% if obj.recordcount > 1 Then %>
			<li style="float:right;margin-right:-10px;"><a href="index.asp?alt=change">CHANGE SUBJECT</a></li>
			<% end if %>
			
			<% IF pref_change_pass THEN %>
			<li style="float:right;margin-right:0px;"><a style="padding-left:0px;" href="change_password.asp">CHANGE PASSWORD</a></li>
			<% END IF %>
			
			<!--<li style="float:right;margin-right:0px;"><a href="javascript:" onClick="window.open('fieldreport.asp','LogdeFeedbackReport','resizable=yes,scrollbars=yes,width=568,height=550,left=50,top=50')">HELPDESK</a></li>-->
			<% END IF%>
		<% END IF%>
		
		</ul>
		</div>
	</div>
</div>
<div class="allcontent">
	<div class="allcontent_main">