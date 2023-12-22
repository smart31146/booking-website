<%@LANGUAGE="VBSCRIPT"%>
<% 

'Response buffer is used to buffer the output page. That means if any database exception occurs the contents can be cleared without processed any script to browser
 Response.Buffer = True
 
' "On Error Resume Next" method allows page to move to the next script even if any error present on page whcich will be caught after processing all asp script on page
 On Error Resume Next
 
'Changed by PR on 25.02.16
 %>
 
<% response.buffer = true%>
<% Response.Expires         = -1 %>
<% Response.ExpiresAbsolute = Now() - 1 %>
<% Response.CacheControl    = "no-cache; private; no-store; must-revalidate; max-stale=0; post-check=0; pre-check=0; max-age=0" %>
<% Response.AddHeader         "Cache-Control", "no-cache; private; no-store; must-revalidate; max-stale=0; post-check=0; pre-check=0; max-age=0" %>
<% Response.AddHeader         "Pragma", "no-cache" %>
<% Response.AddHeader         "Expires", "-1" %>
#include file = "connections/bbg_conn.asp"
#include file = "connections/include.asp"
#include file="sha256.asp"
<%
  
  
if request.querystring("alt")="change" THEN

Session("id") = ""
Session("name") = ""
response.redirect "index.asp"
END IF

if request.querystring("alt")="logout" THEN
  Session.Contents.RemoveAll() 
  Session.abandon
  response.redirect "index.asp"
  
END IF
'is there an autolog query string parameter?

' Err.Number is a attribute of "On Error Resume Next" method
' It is used to terminate any database query or transaction to provide protection against data integrity
' Changed by PR 23.02.16

if err.Number = 0 then
set preferences = Server.CreateObject("ADODB.Recordset")
  preferences.ActiveConnection = Connect
  preferences.Source = "SELECT preferences.*  FROM preferences ORDER BY preferences.pref_date DESC;"
  preferences.CursorType = 0
  preferences.CursorLocation = 3
  preferences.LockType = 3
  preferences.Open()
  preferences_numRows = 0
end if  
autolog_str = CStr(Request.QueryString("autolog"))

if request.querystring("alt")="login" AND len(request.form("bbp_username"))>1 AND len(request.form("bbp_password"))>1 THEN

  Session("userID") = ""
  Session("firstname") = ""
  Session("lastname") = ""
  Session("id") = ""
  Session("name") = ""
  Session.Contents.RemoveAll() 
  session.Timeout=1440
  username = Replace(CStr(Request.form("bbp_username")), "'", "''")
  password = CStr(Request.form("bbp_password"))

' Condition to terminate following query if any error in code. Changed by PR 23.02.16 
if Err.Number = 0 then
  Set obj = Server.CreateObject("ADODB.Recordset")
  SQL="SELECT user_email FROM q_user WHERE user_username='"&username&"'"
  obj.ActiveConnection = Connect
  obj.Source = SQL
  obj.CursorType = 0
  obj.CursorLocation = 3
  obj.LockType = 3
  obj.Open
  
  if obj.EOF then
  response.redirect "index.asp?error=login"
  end if
end if  
  Dim pass
  Dim salt
  salt = obj("user_email")
  password=password&salt
  password=sha256(password)
  
  'Logging user logincount, date and IP
  'Condition to terminate following query if any error in code. Changed by PR 23.02.16  
if Err.Number = 0 then
  Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = connect
    user_logcount = 1
    user_IP=Request.ServerVariables("REMOTE_ADDR")
      MM_insert="UPDATE q_user SET user_IP='"&user_ip&"', user_access=GETDATE(), user_logcount='"&user_logcount&"' WHERE user_username='"&username&"'"
      MM_editCmd.CommandText = MM_insert
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close
end if
 
  'END log
' Condition to terminate following query if any error in code. Changed by PR 23.02.16
if Err.Number = 0 then
    'pn 051021 check to see if this users password in database is null , if it is then set it as another one
        dim setPassword
        setPassword=false
              
          SQL= "SELECT user_city from  q_user  WHERE (user_username) =?"
        set objCommand = Server.CreateObject("ADODB.Command") 
        objCommand.ActiveConnection = Connect
        objCommand.CommandText = SQL
        objCommand.Parameters(0).value = username


      Set MM_rsPassword = objCommand.Execute()
      
      
      
        If (Not MM_rsPassword.EOF Or Not MM_rsPassword.BOF) Then

              if (isNull(MM_rsPassword.Fields.Item("user_city").Value)) then
                setPassword=true
              end if

        end if
        MM_rsPassword.Close

        if setPassword then

                MM_passwordeditQuery = "update q_user set user_city=? WHERE user_username =?"
                Set MM_passwordeditCmd = Server.CreateObject("ADODB.Command")
                MM_passwordeditCmd.ActiveConnection = connect
                MM_passwordeditCmd.CommandText = MM_passwordeditQuery
                
                MM_passwordeditCmd.Parameters(0).value = password
                MM_passwordeditCmd.Parameters(1).value = username
                
                MM_passwordeditCmd.Execute
                MM_passwordeditCmd.ActiveConnection.Close
         end if
 
end if 
' Login code
SQL= "SELECT TOP 1 * FROM q_user WHERE (user_username) =? and (user_city) =? and user_active=1"
set objCommand = Server.CreateObject("ADODB.Command") 
objCommand.ActiveConnection = Connect
objCommand.CommandText = SQL
'objCommand.Parameters(0).value = CStr(Request.form("bbp_username"))
objCommand.Parameters(0).value = username
objCommand.Parameters(1).value = password
Set MM_rsUser = objCommand.Execute()

  'set MM_rsUser = Server.CreateObject("ADODB.Recordset")
  ' MM_rsUser.ActiveConnection = Connect
    'MM_rsUser.Source = "SELECT TOP 1 * from q_user WHERE user_username = '" & username & "' AND  user_city = '" & password & "' AND  user_active = 1 "
    'MM_rsUser.CursorType = 0
    'MM_rsUser.CursorLocation = 3
    'MM_rsUser.LockType = 3
    'MM_rsUser.Open
    
    If (MM_rsUser.EOF Or MM_rsUser.BOF) Then
      response.redirect "index.asp?error=login"
          
    End If

    Session("userID") = MM_rsUser("ID_user")
    Session("firstname") = MM_rsUser("user_firstname")
    Session("lastname") = MM_rsUser("user_lastname")

    MM_rsUser.Close
      response.redirect "index.asp"     
END IF

    'Set that user's subject to the requested course ID only
    ' Condition to terminate following query if any error in code. Changed by PR 23.02.16 
if Err.Number = 0 then
if request.querystring("alt")="choose" AND request.querystring("id_subject") <> "" THEN

    ' Getting the current subject that the user has loged in to
    SQL = "SELECT id_subject,subject_name FROM subjects WHERE id_subject="& fixstr(clng(request.querystring("id_subject"))) &""
    obj.Open SQL, Connect,3,3
      Session("id") = obj("id_subject")
      Session("name") = obj("subject_name")
    obj.close
    
    response.redirect "index.asp"   
End If
end if

%>
<!doctype html>

<head>
 <meta id="viewport" name="viewport">
<script>

</script>
  <title><%=client_name_short%> - Better Business Program</title>
    <META name="DESCRIPTION" content="">
    <script src="jquery-1.11.1.js?v=bbp34"></script>
      <script src="js/freewall.js?v=bbp34"></script>
    <script src="js/modernizr-latest.js?v=bbp34"></script>
      <link rel="stylesheet" type="text/css" href="js/sweet-alert.css">
    <script src="js/sweet-alert.min.js?v=bbp34"></script>
    <script src="perfect-scrollbar-0.4.8/src/perfect-scrollbar.js?v=bbp34"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-C6RzsynM9kWDrMNeT87bh95OGNyZPhcTNXj1NW7RuBCsyN/o0jlpcV8Qyq46cDfL" crossorigin="anonymous"></script>
      <!-- #include file = "inc_header.asp" -->
  
    <script >
      
  
  
    $(document).ready(function() {
	
	if(window.innerHeight > window.innerWidth){
   swal({   title: "This site is best used in landscape view.",   text: "",   type: "warning",   confirmButtonText: "OK" }); 
 
}
	


    		  $("html, body").animate({ scrollTop: $(document).height() }, 10);
		/*var d = $("body");
		 var rotate = 90 - window.orientation;
		 d.css("transform", "rotate("+rotate+"deg)");
		window.addEventListener('orientationchange', function ()
		{
			//adapt_to_orientation();
   
				if(window.orientation > 0)
					rotate=0;
				else 
				rotate=90;
	
				d.css("transform", "rotate("+rotate+"deg)");
		});
	*/
		 	
      
	  $(".inside_content").perfectScrollbar({suppressScrollX: true});
	  
	  
     $("#registerLink").click(function(){
     $("#regmodal").attr("src","self-register.asp");
     
     });
     $("#forgot").click(function(){
     $("#getpassmodal").attr("src","get_password.asp");
     
     });
	 

    

   
              var wall = new freewall("#container");
      wall.fitWidth();

    if (window.location.protocol != "https:")
    window.location.href = "https:" + window.location.href.substring(window.location.protocol.length);

var d = new String(window.location.host); 
var p = new String(window.location.pathname); 
var u = "http://" + d + p; 



if (u.indexOf("lawofthejungle.com.au") >= 0 || u.indexOf("lotj.com") >= 0 || u.indexOf("lawofthejungle.co.nz") >= 0 || u.indexOf("lawofthejungle.co.nz") >= 0 || u.indexOf("lawofthejungle.net") >= 0 || u.indexOf("lawofthejungle.net.au")  >= 0 || u.indexOf("lotj.co.uk") >= 0 || u.indexOf("lotj.co.nz") >= 0  )
{ 

d = "www.lawofthejungle.com";

if (window.location.protocol == "https:")
var u = "https://" + d + p;
else
var u = "http://" + d + p; 


window.location = u; 
} 

    var version=getInternetExplorerVersion();
    if (version <0)
    {
    //window.location.replace("https://www.lawofthejungle.com/acme33_release2/nobrowse/");
    }
	

});


	// orientation code
    //adapt_to_orientation();
    var d = $("body");
     var rotate = 90 - window.orientation;
     d.css("transform", "rotate("+rotate+"deg)");
    window.addEventListener('orientationchange', function ()
    {
      //adapt_to_orientation();
		
		
        if(window.orientation > 0)
          rotate=0;
        else 
        rotate=90;
  
        d.css("transform", "rotate("+rotate+"deg)");
    });
    

    function getInternetExplorerVersion()
{
  var rv = -1;
  if (navigator.appName == 'Microsoft Internet Explorer')
  {
    var ua = navigator.userAgent;
    var re  = new RegExp("MSIE ([0-9]{1,}[\.0-9]{0,})");
    if (re.exec(ua) != null)
      rv = parseFloat( RegExp.$1 );
  }
  else if (navigator.appName == 'Netscape')
  {
    var ua = navigator.userAgent;
    var re  = new RegExp("Trident/.*rv:([0-9]{1,}[\.0-9]{0,})");
    if (re.exec(ua) != null)
      rv = parseFloat( RegExp.$1 );
  }
  return rv;
}
 

function checkLogin() {

  if (document.getElementById("bbp_username").value == '' )
   {
    swal({   title: "Username is missing!",   text: "",   type: "error",   confirmButtonText: "OK" }); return false;
   }else if (document.getElementById("bbp_password").value == '' )
   {
    swal({   title: "Password is missing!",   text: "",   type: "error",   confirmButtonText: "OK" }); return false;
   }
    else {
    
     return true;
  }
}

function adapt_to_orientation() {
// For use within normal web clients 
var isiPad = navigator.userAgent.match(/iPad/i) != null;

      var content_width, screen_dimension;

      if (window.orientation == 0 || window.orientation == 180) {
        // portrait

        content_width = 900;
        screen_dimension = screen.width * 0.98; // fudge factor was necessary in my case
      } else if (window.orientation == 90 || window.orientation == -90) {
        // landscape
        content_width = 750;
        screen_dimension = screen.height;
      }

      var viewport_scale = screen_dimension / content_width;

      // resize viewport
      $('meta[name=viewport]').attr('content',
        'width=' + content_width + ',' +
        'initial-scale=' + viewport_scale + ', maximum-scale=' + viewport_scale);
    
    // resize viewport
      //$('meta[name=viewport]').attr('content','user-scalable=YES');
       
    }
    </script>
</head>
<body>
  <div class="page-content">
    <div class="white-container">
      <!-- #include file = "partials/header.asp" -->
      <div class="allcontent">
        <div class="allcontent_main">
          <div class="main_content" id="dd" >

            <div style="background: url('images/start_main.jpg') no-repeat;width:600px;height:300px;position:relative;" >
              <div style="position:absolute;left:205px;top:10px;color:#fff;width: 418px;max-height:285px;overflow-y:overlay;overflow-x: hidden;">
                <strong>Welcome to the Better Business Program<%= username%></strong>
                  <% IF Session("userID") <> "" THEN%>
              <% IF Session("id") = "" THEN
              SQL = "SELECT subject_user.ID_subject,subject_name FROM subject_user,subjects WHERE subject_user.id_subject = subjects.id_subject AND id_user='"& fixstr(Session("UserID")) &"' AND (subjects.subject_active_q = 1 OR subjects.subject_active_b = 1);"
              obj.Open SQL, Connect,3,3%><br>

              <br>
              Please choose which subject you would like to work with<br>
              <br>
              <table width = 100% border = '0' cellspacing = '2' cellpadding = '0' class="item">
              
              <% 
              Dim position
              position=1
              Dim over
              over=0
              if obj.recordcount = 1 Then %>
                <meta http-equiv="refresh" content="0;URL='index.asp?alt=choose&amp;id_subject=<% =obj("id_subject")%>'">
              <% else %>
              <% do until obj.eof
              if position=1 then 
              over=0
              response.write("<tr style='height:20px;'>")
              end if
              if over=1 then
              response.write("<td  style='padding-left:10px;'>")
              
              else
              response.write("<td >")
              end if
            
              %>
                  
                <a id="subject" style="color:#000; padding-left:0px; padding-top:15px; color:#0B66BD; font-size:11px;"  href="index.asp?alt=choose&amp;id_subject=<% =obj("id_subject")%>"><div  id="button" style="color:#000; padding-left:0px; color:#0B66BD; font-size:11px;" ><% =ReplaceStrBBG(obj("subject_name"))%></div></a>
              <% 
              response.write("</td>")
              over=1
                
              if position=2 then 
              response.write("</tr>")
              position=1
              
              else
              position=position+1
              
              end if
              
              obj.movenext
              loop 
              
                  response.write(theend & "</table>") 
              %>
              <% obj.Close%>
              <% end if %>
              <% ELSE%>
              
              
          <br><span style="font-style:italic;font-size:14px;line-height:15px;">
            <%  
            'Welcome Message
            response.write (preferences.Fields.Item("welcome_note").value)
            %>
          </span>
              <% END IF%><% ELSE%><br>
          <br>
          <strong>Test site login</strong> <br><div style="font-size:11px;color:#fff;">
          If this is your first visit to this test site, please self-select a password by entering it in the password field below
          </div>
          <br>

          <form role="form" method="post" action="index.asp?alt=login" onsubmit="return(checkLogin());">
              <div class="mb-3">
                <input type="text" class="form-control" name="bbp_username" id="bbp_username"  placeholder="Enter your username">
              </div>
              <div class="mb-3">
                <input type="password" class="form-control" name="bbp_password" id="bbp_password" placeholder="Enter your password">
              </div>
              
            <% IF request.querystring("error")="login" THEN response.write "<script>swal({   title: ""The Username or Password you entered is incorrect."",   text: ""Your username is  your <b>firstname.lastname</b>  (e.g. JOHN.SMITH )<br><br>If you have forgotten your password, you can click 'Forgotten Password' below to reset your password."",   type: ""error"",   confirmButtonText: ""OK"",html: true });</script>" %>

              <button id="login" type="Submit" class="btn btn-light" value="Log in" >Log in</button>
          
          <br>
          <br>
          <% IF pref_self_reg THEN %>
                
                <a id="registerLink" href="#"  data-bs-toggle="modal" data-bs-target="#myModal" >I'm a new user</a> | 
                <% END IF %>
          <% IF pref_forgot_pass THEN %>
          <a style="color: #fff"  id="forgot" href="#" data-bs-toggle="modal" data-bs-target="#passModal" >Forgot your username or password?</a>
          <% END IF %>

          </form>

          <% END IF%>
              </div>
            </div>
            
            <div class="box_blue">
              <div class="box_inside2">
              <% IF Session("userID") <> "" AND Session("id") <> "" THEN%>
                <h1>The training is the place to start </h1>
              <!--  You will be taken through a number of topics on this subject.<br>
                At the end of each topic, your understanding will be tested by a couple of quiz<br>
                questions that relate to that topic. You can leave the training at any time<br>and you will be taken back to the start of the topic you were working on.<br>-->

                  <div class="submit_blue">
                    <div class="h_submit_blue">
                      <a class="box_link" href="t_index.asp?ID_subject_prm=<%=Session("id")%>" style="padding-left:15px;">
                        START TRAINING (<%=UCASE(ReplaceStrBBG(Session("name")))%>)
                      </a>
                    </div>
                  </div>
                </div>
              <% ELSE%>
                <br><br>
                <br><br>
              </div>
              <% END IF%>
              
          </div>
          <div class="clear"></div>
        </div>

        <div class="menu_content" style="margin-left:20px;">
          <img src="images/start_right_1.jpg" width="300" height="300" alt=""><br>
          <div class="box_grey"><div class="box_inside2">
            <% IF Session("userID") <> "" AND Session("id") <> "" THEN%>
              <% IF Session("id") <> 3 THEN %>
            <!--<h2>Useful information</h2>--><!--This useful guide is a quick<br>
              search and easy reference for all<br>
              your training subjects in the <br>
              'Better Business Program'.<br>-->
              <!-- <div class="submit_grey">
                  <!--<div class="h_submit_grey"><a  class="box_link" href="g_index.asp"> QUICK GUIDE</a></div>
                <!--</div> -->
            <br/><br/><br/><br/>
              <% ELSE%>
              <br><br>
              <br><br>
              <br><br>
              <br><br>
              <% END IF%>
            <% ELSE%>
              <br><br>
              <br><br>
            <% END IF%>
            </div>
            </div>
          </div>
          <div class="clear"></div>
        </div>

      <div class="clear"></div>

      <!-- Modal -->
      <div class="modal fade" id="myModal" data-bs-backdrop="static" data-bs-keyboard="false" tabindex="-1" aria-labelledby="myModalLabel" aria-hidden="true">
        <div class="modal-dialog">
          <div class="modal-content">
            <div class="modal-header">
              <button type="button" class="close btn-close" data-bs-dismiss="modal" aria-label="Close">X</button>
            </div>
            <div class="modal-body">
                <iframe  id="regmodal" src="" width="500" height="400" frameborder="0" allowtransparency="true"></iframe>  
            </div>
            
          </div>
          <!-- /.modal-content -->
        </div>
        <!-- /.modal-dialog -->
      </div>
      <!-- /.modal -->

      <!-- Modal -->
      <div class="modal fade" id="passModal" data-bs-backdrop="static" data-bs-keyboard="false" tabindex="-1" aria-labelledby="myModalLabel" aria-hidden="true">
        <div class="modal-dialog">
          <div class="modal-content">
            <div class="modal-header">
              <button type="button" class="close btn-close" data-bs-dismiss="modal" aria-label="Close">X</button>
              &nbsp;
            </div>
            <div class="modal-body">              
                <iframe id="getpassmodal" src="" width="100%" height="400" frameborder="0" allowtransparency="true"></iframe>  
            </div>
            
          </div>
          <!-- /.modal-content -->
        </div>
        <!-- /.modal-dialog -->
      </div>
      <!-- /.modal -->
    </div>
    </div>
    <!-- #include file = "partials/footer.asp" -->
  </div>
<% if request.querystring("alt")="login" AND  len(request.form("bbp_password"))<2 then response.write "<script>swal({   title: ""You must enter a password"",   text: """",   type: ""error"",   confirmButtonText: ""OK"",html: true });</script>" end if %>

<% if request.querystring("alt")="login" AND len(request.form("bbp_username"))<1 AND len(request.form("bbp_password"))<1 then response.write "<script>swal({   title: ""You must enter a username"",   text: """",   type: ""error"",   confirmButtonText: ""OK"",html: true });</script>" end if %>
</body>
<% if Request.ServerVariables("HTTP_REFERER") <> "" then comment = Request.ServerVariables("HTTP_REFERER") else comment = "new page"
call log_the_page ("home", "0", "n/a", "0", "n/a", "0", "n/a", comment)
%>
<script src="js/ios-orientationchange-fix.js?v=bbp34"></script>
<%
' Page 'errorhandler/index.asp' will be called if there is any database exception present in Script
%>

<!-- #include file = "errorhandler/index.asp"-->