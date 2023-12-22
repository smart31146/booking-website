<%@LANGUAGE="VBSCRIPT"%>

<% 
'Response buffer is used to buffer the output page. That means if any database exception occurs the contents can be cleared without processed any script to browser
 Response.Buffer = True
 
' "On Error Resume Next" method allows page to move to the next script even if any error present on page whcich will be caught after processing all asp script on page
 On Error Resume Next
 
'Changed by PR on 25.02.16
%>

<!--#include file="connections/bbg_conn.asp"-->
<!--#include file="connections/include.asp"-->
<!--#include file="sha256.asp"-->

<%
Session("logo") = "client_logotype"
Session("font") = "arial"
not_found=-1
Dim pval

Dim pval2
Dim miserror



%>
<%
Private Function Decrypt(ByVal encryptedstring)
    Dim x, i, tmp
    encryptedstring = StrReverse( encryptedstring )
    For i = 1 To Len( encryptedstring )
        x = Mid( encryptedstring, i, 1 )
        tmp = tmp & Chr( Asc( x ) - 1 )
    Next
    Decrypt = tmp
End Function
%>
<% if request.querystring("uid") <> "" then
Dim userid

userid=URLDecode(request.querystring("uid"))
'userid=EnDeCrypt(userid,salt)
'userid=StringToHex(userid)



'Set obj = Server.CreateObject("ADODB.Recordset")
'SQL="SELECT * FROM q_user WHERE ID_User="&request.querystring("uid")
'obj.ActiveConnection = Connect
'obj.Source = SQL
'obj.CursorType = 0
'obj.CursorLocation = 3
'obj.LockType = 3
'obj.Open



'obj.close



end if %>

<% 
 Function StringToHex(ByRef pstrString)
    	Dim llngIndex
    	Dim llngMaxIndex
    	Dim lstrHex
    	llngMaxIndex = Len(pstrString)
    	For llngIndex = 1 To llngMaxIndex
    		lstrHex = lstrHex & Right("0" & Hex(Asc(Mid(pstrString, llngIndex, 1))), 2)
    	Next
    	StringToHex = lstrHex
    End Function
    Function HexToString(ByRef pstrHex)
    	Dim llngIndex
    	Dim llngMaxIndex
    	Dim lstrString
    	llngMaxIndex = Len(pstrHex)
    	For llngIndex = 1 To llngMaxIndex Step 2
    		lstrString = lstrString & Chr("&h" & Mid(pstrHex, llngIndex, 2))
    	Next
    	HexToString = lstrString
    End Function
    Function URLDecode(str) 
        str = Replace(str, "+", " ") 
        For i = 1 To Len(str) 
            sT = Mid(str, i, 1) 
            If sT = "%" Then 
                If i+2 < Len(str) Then 
                    sR = sR & _ 
                        Chr(CLng("&H" & Mid(str, i+1, 2))) 
                    i = i+2 
                End If 
            Else 
                sR = sR & sT 
            End If 
        Next 
        URLDecode = sR 
    End Function 
 
    Function URLEncode(str) 
        URLEncode = Server.URLEncode(str) 
    End Function 
 
%>
<% if request("pass_req")= "" AND request("submit")<>"" then

userid=request("puserid")
Dim req1
req1="* Field Required"
else

pval=request("pass_req")

req1=""
end if
%>
<% if request("repass_req")= "" AND request("submit")<>"" then

userid=request("puserid")
Dim req2
req2="* Field Required"
else

pval2=request("repass_req")
req2=""

end if
%>
<% 



if request("pass_req") <>""  AND len(request("pass_req")) > 5 AND request("submit")<>"" AND request("repass_req")<> "" AND len(request("repass_req"))> 5 then

userid=request("puserid")

if pval<>pval2 then

miserror="<br><span style='color:red' >- Passwords don't match</span>"
else



Dim uid
userid=HexToString(userid)
dim all
all=Decrypt(userid)
uid=all
'uid=EnDeCrypt(userid,salt)

if Err.Number = 0 then
Set obj = Server.CreateObject("ADODB.Recordset")
SQL="SELECT user_email FROM q_user WHERE ID_User="&uid
obj.ActiveConnection = Connect
obj.Source = SQL
obj.CursorType = 0
obj.CursorLocation = 3
obj.LockType = 3
obj.Open
end if

Dim pass
Dim salt
salt = obj("user_email")
pass=pval&salt
pass=sha256(pass)
obj.close

if Err.Number = 0 then
SQL="update q_user set user_city=? WHERE ID_User=?"
set objCommand = Server.CreateObject("ADODB.Command") 
objCommand.ActiveConnection = Connect
objCommand.CommandText = SQL 
objCommand.Parameters(0).value = pass
objCommand.Parameters(1).value=uid
objCommand.Execute()
end if

miserror="<br><span style='color:green' >Your password has been reset.</span><br><br><div style='font-size:14px;'>To sign in go to <a href='http://www.lawofthejungle.com/"&client_name_short&"/' >login</a> page</div>"


end if

else if len(request("pass_req")) <= 5 AND len(request("pass_req")) > 0  AND request("submit")<>"" OR len(request("repass_req"))<= 5 AND len(request("repass_req")) > 0   then 
if pval<>pval2 then

miserror="<br><span style='color:red' >- Passwords don't match</span>"
end if

miserror=miserror&"<br><span style='color:red' >- Your password must be 6 characters or more</span>"
userid=request("puserid")

end if






end if
%>
<!doctype html>
<head>
<script src="jquery-1.11.1.js?v=bbp34"></script>
			<script src="js/freewall.js?v=bbp34"></script>
		<script src="js/modernizr-latest.js?v=bbp34"></script>
			<link rel="stylesheet" type="text/css" href="js/sweet-alert.css">
    		  <script src="js/sweet-alert.min.js?v=bbp34"></script>
<script>
//var parameter = getUrlParameters("uid", "", true);
//alert(parameter);
	$(document).ready(function() {
		
		 
//alert($("*:focus").attr("id"));
setTimeout(function() {
$("#pass_req").blur();

      // Do something after 2 seconds

}, 100);


 
  
	
});

function getActive(){
   return $(document.activeElement).is('input') || $(document.activeElement).is('textarea');
}

function getUrlParameters(parameter, staticURL, decode){
   /*
    Function: getUrlParameters
    Description: Get the value of URL parameters either from 
                 current URL or static URL
    Author: Tirumal
    URL: www.code-tricks.com
   */
   var currLocation = (staticURL.length)? staticURL : window.location.search,
       parArr = currLocation.split("?")[1].split("&"),
       returnBool = true;
   
   for(var i = 0; i < parArr.length; i++){
        parr = parArr[i].split("=");
        if(parr[0] == parameter){
            return (decode) ? decodeURIComponent(parr[1]) : parr[1];
            returnBool = true;
        }else{
            returnBool = false;            
        }
   }
   
   if(!returnBool) return false;  
}
</script>

<title><%=client_name_short%> Better Business Program
<%if Session("MM_Username") <> "" then response.write(" - you are logged in as " & Session("firstname") & " " & Session("lastname"))%>
</title>
<link rel="stylesheet" href="style/bbp_reset_password_acme34.css" type="text/css">

</head>

<body class="login">

<div class="loginbox radius">
<a href="index.asp" target="_top"><h4 style="color:#FFF; text-align:center"></h4></a>
	<div class="loginboxinner radius">
    	<div class="loginheader">
    		<h2 class="title">Reset Password</h2>
        	<div class="info">(Your password must be 6 characters or more)</div>
    	</div><!--loginheader-->
        <div class="loginform">
        	<form name="password_form" method="post" action="reset_password.asp?focus=false">
            	<p>
                <label for="username" class="bebas">Enter new password:</label>
             <input  type="password" class="radius2" name="pass_req" id="pass_req" value='<% response.write(pval)%>' />
                </p>
                <p>
                <label for="password" class="bebas">Re-enter new password:</label>
                <input  type="password"    class="radius2" name="repass_req" id="repass_req" value='<% response.write(pval2) %>'  />
                </p>
                <p>
               <button  type="submit" id="sbmt" class="radius title" value="Submit" name="submit">Submit</button>
                </p>
				<p>
				<script>
				
				function hideKeyboard() {
					document.activeElement.blur();
					var inputs = document.querySelectorAll('input');
						for(var i=0; i < inputs.length; i++) {
					inputs[i].blur();
					}
					
				}
			
				//swal({   title: "Username is missing!",   text: "<%= miserror %>",   type: "error",   confirmButtonText: "OK", html:true });
				
				</script>
				<% if request.querystring("focus")="false" then %>
				<script>swal({   title: "",   text: "<%= miserror %>",   type: "warning",   confirmButtonText: "OK", html:true }); $("#pass_req").focus(); </script>
				<% end if %>
				
				</p>
				 <input type="hidden" value="<%= userid%>" name="puserid"/>
            </form>
        </div><!--loginform-->
    </div><!--loginboxinner-->
</div><!--loginbox-->








</body>












</html>
<%
if Request.QueryString("uid") <> "" then
	comment="User: "& request("username_req")
else
	comment="Request password"
end if
call log_the_page ("password", "0", "n/a", "0", "n/a", "0", "n/a", comment)
%>
<!-- #include file = "errorhandler/index.asp"-->