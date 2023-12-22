<html>
<head>
<title>Error</title>
<link rel="stylesheet" href="styles/bbp_style_acme34.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000"><br>

<DIV id="topindex" style="LEFT: 40%; POSITION: absolute; TOP: 150px">
<table width="400" border="1" cellspacing="0" cellpadding="10" align="center">
  <tr> 
    <td align="center" bgcolor="#FFFFFF"> 
           <p class="text"><b>Access to the Better Business Program is restricted to computers connected to the Asahi Beverages network.<br><br>

If you are accessing the program from within the your network, please contact your IT helpdesk and mention the following IP address of your computer:</font></b></p>
      <p class="text"><b><font size="5" color="red"><%=Request.ServerVariables("REMOTE_ADDR")%></font></b></p>
      </td>
  </tr>
</table></DIV>
</body>
</html>
