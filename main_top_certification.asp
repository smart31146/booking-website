<script src="js/jquery.magnific-popup.min.js?v=bbp34"></script>
<link rel="stylesheet" type="text/css"href="style/magnific-popup.css">
<script>
$(document).ready(function() {
$('.pop1').magnificPopup({
  type: 'iframe',
  
  iframe: {
    markup: '<div style="width:500px; height:700px;maring-top:80px;">'+
    '<div class="mfp-iframe-scaler" >'+
            '<div class="mfp-close"></div>'+
		
    '<iframe class="mfp-iframe" frameborder="0" allowfullscreen></iframe>'+
    '</div></div>'
  }
});





window.addEventListener('orientationchange', function ()
		{
		if(window.orientation == 0){
				swal({   title: "This site is best used in landscape view.",   text: "",   type: "warning",   confirmButtonText: "OK" }); 
				
				}
				else{
$('meta[name=viewport]').attr('content', "width=600, initial-scale=1.0");
				}
				 // resize viewport
					//
				
					
		});



});

function checkSearchField() {
	var x = document.forms["searchform"]["bbp_search"].value;
	if (x == null || x == "" || x.length > 30){
		alert("Maximum 30 characters allowed for search guide")
		return false;
	}
}

</script>
<%
function curPageName()
 dim pagename

 pagename = Request.ServerVariables("SCRIPT_NAME") 

  if inStr(pagename, "/") > 0 then 
    pagename = Right(pagename, Len(pagename) - instrRev(pagename, "/")) 
  end if 

 curPageName = pagename
end function 


%>

	<div class="header">
		<div class="header_main">
		<div class="header_logo">
		<% IF not bbp_training = True THEN%>
		<% IF Session("userID") <> "" AND Session("id") <> "" THEN%>
			<!--<div style="position:absolute;right:10px;top:1px;color:#fff;"><div style="float:left;color: #fff;font-weight:bold;margin:9px;font-size:16px;">Search guide </div><div style="float:left;"><form  style="float:left;display:inline-box; display:-webkit-inline-box; display:-ms-inline-flexbox" action="g_search.asp" method="post"  id="searchform"  name="searchform"onsubmit="return checkSearchField();"><input type="text" class="form-control" onfocus="if (this.value == '<% =Session("bbp_search")%>') {this.value = '';}" onblur="if (this.value.length == 0) { this.value = '<% =Session("bbp_search")%>' }" style="width:130px;vertical-align:middle;" value="<% =Session("bbp_search")%>" name="bbp_search"> <button type="submit" value="Search" style="vertical-align:middle;background:url(images/search_button.png) no-repeat; width:30px;height:30px;border:0px;cursor:pointer;" ></button></form></div></div>-->
			<!--<div style="position:absolute;right:70px;top:72px;color:#777; font-size: 1.3em; font-weight: bold; "><i>Better Business Program</i></div>-->
		<% END IF %>
		<% END IF %>
		</div>
		
		<div id="bbp_menu">
		
		<ul class="menu">
		<% IF Session("userID") <> "" AND Session("id") <> "" THEN%>

			<% IF Session("LMS") <> "1" THEN %>
			<li style="margin-right:0px;width: 120px;" id="help"><a class="pop1" href="help.asp">HELP/FEEDBACK</a></li>
			<% end if %>
			
			<!-- Updated code to remove user ID from URL by PR on 23.02.2016-->
			<% IF Session("LMS") <> "1" THEN %>
			<li style="float:left;margin-left: -10px;width: 200px; padding-left:35px;" id="results">
			<form name='myForm' action='user_sessions_new.asp' method='post'>
			<input type="hidden" name="latest" value="1"/>
			<a href="javascript: previousResults();">PREVIOUS RESULTS</a>
			</form>
			
			<script type="text/javascript">
			previousResults = function()
 			{
    		window.open('', 'popup', 'width=920,height=624,left=50,top=50,resizeable=yes, scrollbars=yes');
    		document.forms["myForm"].setAttribute('target', 'popup');
    		document.forms["myForm"].setAttribute('onsubmit', '');
    		document.forms["myForm"].submit();
			};
			</script>
			</li>
			
			<% END IF %>
			
			<% END IF%>
		<% IF bbp_training = True THEN%>
			<% IF Session("LMS") <> "1" THEN %><li style="float:right;margin-right:30px;" id="exit"><a href="#" onclick="swal({   title: 'Are you sure you want to exit the training?',   text: 'Your records will be saved and you can continue from this place later on.',   type: 'warning',   showCancelButton: true,   confirmButtonColor: '#DD6B55',   confirmButtonText: 'Yes',   closeOnConfirm: false }, function(isConfirm){  if(isConfirm) {window.location.href='index.asp'; return false;} else { $('html, body').animate({ scrollTop: $(document).height() }, 10); } });">EXIT TRAINING</a></li><% END IF %>
			
		<%ELSE%>
		<%IF Session("userID") <> ""  THEN %>
				<!--<li style="float:left;margin-left: 0px;"><a href="#" onClick="window.open('user_sessions_new.asp?user=<%=Session("userID")%>&amp;latest=1','newwindow','scrollbars=yes,resizable=yes, width=920,height=624,left=50,top=50')">PREVIOUS RESULTS</a></li>-->
			
		<%END IF%>
		
		
		
			

			<% IF Session("userID") <> "" AND Session("id") <> "" THEN%>
			<% IF pref_change_pass and Session("LMS") <> "1"  THEN %>
			
			<% END IF %>
			
			
			<!--<li style="float:right;margin-right:0px;"><a href="javascript:" onClick="window.open('fieldreport.asp','LogdeFeedbackReport','resizable=yes,scrollbars=yes,width=568,height=550,left=50,top=50')">HELPDESK</a></li>-->
			<% END IF%>
			<% IF Session("userID") <> "" then %>
			<% ' removed logout link from scorm if session is LMS by PR on 29.02.16
			IF Session("LMS") <> "1"  THEN 
			%> <li style="float:left;margin-left:-10px;"><a style="padding-left:40px;" href="change_password.asp" id="changepass" >CHANGE PASSWORD</a></li>
			<li style="float:right;margin-left:0px; padding-right:5px;" id="logout"><a style="padding-left:10px;" href="#" onclick="swal({   title: 'Do you want to logout?',   text: '',   type: 'warning',   showCancelButton: true,   confirmButtonColor: '#DD6B55',   confirmButtonText: 'Yes',   closeOnConfirm: false }, function(){  window.location.href='index.asp?alt=logout'; return false; });">LOG OUT</a></li>
			<% end if %>
			<% 'if curPageName = "g_search.asp" THEN 
			%>
				<!--<li style="float:left;margin-right:15px;" id="home"><a style="padding-left:10px;" href="index.asp">HOME</a></li>-->
				<%' end if 
				%>
			<%else%>
			
			
			<% end if %>
			
		<% END IF%>
		
		</ul>
		</div>
	</div>
</div>

<div class="allcontent">
	<div class="allcontent_main">