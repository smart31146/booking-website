<%@LANGUAGE="VBSCRIPT"%>
<!-- #include file = "connections/bbg_conn.asp" -->
<!-- #include file = "connections/include.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
  <title><%=client_name_short%> - Better Business Program</title>
  <meta name="DESCRIPTION" content="" />
  <!-- <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-T3c6CoIi6uLrA9TneNEoa7RxnatzjcDSCmG1MXxSR1GAsXEV/Dwwykc2MPK8M2HN" crossorigin="anonymous"> -->
  <!-- #include file = "inc_header.asp" -->
  <link rel="stylesheet" href="style/error-page.css">
</head>
<body>
  <div class="page-content">
    <div class="white-container">
      <!-- #include file = "partials/header.asp" -->
      <div class="allcontent">
        <div class="allcontent_main">
          <div class="header_blue">
            <div class="header_inside">
              404<br />
              <h3>Page not found</h3>
            </div>
          </div>

          <div class="clear"></div>

          <div class="main_content">
            <div class="guide_blue">
              <div class="box_inside">
                <h3>Page not found</h3>
                <div class="box_text_blue">
                  <% strQueryString= Request.ServerVariables("QUERY_STRING")
                  'Response.Write("uid=" & uid & "<br />")
                  'Response.Write("cid=" & cid & "<br />")
                  'Response.Write("callbackurl=" & originurl & "<br />")
                  Response.Write(strQueryString & "<br />")
                  Response.Write("uid=" & Session("UserID") & "<br />")
                  Response.Write("id=" & Session("id") & "<br />")
                  Response.Write("LMS=" & Session("LMS") & "<br />") %> There
                  was an error in processing your request. Please click the Home
                  button below to return to the homepage. If the problem
                  persists, please contact your program administrator.
                </div>
                <!-- <div> -->
                  <a
                    class="btn btn-info default-blue-btn p-3 h-auto home-btn"
                    role="button"
                    href="index.asp"
                  >
                    Home
                  </a>
                <!-- </div> -->
              </div>
            </div>
          </div>

          <div class="menu_content">
            <img
              src="vault_image/images/training.jpg"
              width="320"
              height="390"
              alt=""
            />
          </div>

          <div class="clear"></div>
        </div>
      </div>
    </div>
    <!-- #include file = "partials/footer.asp" -->
  </div>
</body>
