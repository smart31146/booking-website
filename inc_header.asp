<% Response.Expires         = -1 %>
<% Response.ExpiresAbsolute = Now() - 1 %>
<% Response.CacheControl    = "no-cache; private; no-store; must-revalidate; max-stale=0; post-check=0; pre-check=0; max-age=0" %>
<% Response.AddHeader         "Cache-Control", "no-cache; private; no-store; must-revalidate; max-stale=0; post-check=0; pre-check=0; max-age=0" %>
<% Response.AddHeader         "Pragma", "no-cache" %>
<% Response.AddHeader         "Expires", "-1" %>

		<META name="ROBOTS"		content="index,follow">
		<META name="LANGUAGE"	content="SE">
		<meta name="keywords" CONTENT="">
		<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-T3c6CoIi6uLrA9TneNEoa7RxnatzjcDSCmG1MXxSR1GAsXEV/Dwwykc2MPK8M2HN" crossorigin="anonymous">
		<link href="style/bbp_acme34.css" rel="stylesheet" type="text/css">
		<link href="style/header-styles.css" rel="stylesheet" type="text/css">
		<link href="style/footer.css" rel="stylesheet" type="text/css">
		<link href="style/global-styles.css" rel="stylesheet" type="text/css">
		<link
		rel="stylesheet"
		href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.3.0/css/all.min.css"
	  	/>
		<script src="js/javascript.js?v=bbp34" type="text/javascript"></script>
		<style media="all" type="text/css">
		img#bg {
		  position:fixed;
		  top:12000;
		  left:0;
		  width:100%;
		  height:100%;
		
		}
		</style>
		<!--[if IE 6]>
		<style type="text/css">
		html { overflow-y: hidden; }
		body { overflow-y: auto; }
		img#bg { position:absolute; z-index:-1; }
		#content { position:static; }
		</style>
		
		<![endif]-->
		<!--[if IE 7]>
<link href="style/bbp_ie7_acme34.css" rel="stylesheet" type="text/css">
<![endif]-->