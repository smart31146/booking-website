<script src="js/jquery.magnific-popup.min.js?v=bbp34"></script>
<link rel="stylesheet" type="text/css" href="style/magnific-popup.css" />
<script>
  $(document).ready(function () {
    $(".pop1").magnificPopup({
      type: "iframe",

      iframe: {
        markup:
          '<div style="width:500px; height:700px;maring-top:80px;">' +
          '<div class="mfp-iframe-scaler" >' +
          '<div class="mfp-close"></div>' +
          '<iframe class="mfp-iframe" frameborder="0" allowfullscreen></iframe>' +
          "</div></div>",
      },
    });

    window.addEventListener("orientationchange", function () {
      if (window.orientation == 0) {
        swal({
          title: "This site is best used in landscape view.",
          text: "",
          type: "warning",
          confirmButtonText: "OK",
        });
      } else {
        $("meta[name=viewport]").attr(
          "content",
          "width=600, initial-scale=1.0"
        );
      }
      // resize viewport
      //
    });
  });

  function checkSearchField() {
    var x = document.forms["searchform"]["bbp_search"].value;
    if (x == null || x == "" || x.length > 30) {
      alert("Maximum 30 characters allowed for search guide");
      return false;
    }
  }

  function previousResults() {
    window.open(
      "",
      "popup",
      "width=920,height=624,left=50,top=50,resizeable=yes, scrollbars=yes"
    );
    document.forms["myForm"].setAttribute("target", "popup");
    document.forms["myForm"].setAttribute("onsubmit", "");
    document.forms["myForm"].submit();
  }
</script>

<div class="header-section">
  <img src="../images/new_design/svg/twe-logo.svg" alt="twe-logo" />

  <div class="header-section__navigation">
    <% IF Session("userID") <> "" then %>
    <a class="header-section__link" href="t_index.asp">
      <i class="fa-solid fa-house"></i>
      <span>Training Home</span>
    </a>
    <a class="pop1 header-section__link" href="help.asp">
      <i class="fa-solid fa-headset"></i>
      <span>Support</span>
    </a>
    <a href="index.asp?alt=change" class="header-section__link"
      >Change subject</a
    >
    <!-- <div style="position:absolute;right:10px;top:1px;color:#fff;">
      <div style="float:left;color: #fff;font-weight:bold;margin:9px;font-size:16px;">
        Search guide
      </div>
      <div style="float:left;">
        <form  style="float:left;display:inline-box; display:-webkit-inline-box; display:-ms-inline-flexbox" action="g_search.asp" method="post"  id="searchform"  name="searchform"onsubmit="return checkSearchField();">
        <input type="text" class="form-control" onfocus="if (this.value == '<% =Session("bbp_search")%>') {this.value = '';}" onblur="if (this.value.length == 0) { this.value = '<% =Session("bbp_search")%>' }" style="width:130px;vertical-align:middle;" value="<% =Session("bbp_search")%>" name="bbp_search"> 
        <button type="submit" value="Search" style="vertical-align:middle;background:url(images/search_button.png) no-repeat; width:30px;height:30px;border:0px;cursor:pointer;" ></button>
      </form>
    </div>
  </div> -->

    <!-- <a href="change_password.asp" class="header-section__link">
      Change password
    </a> -->
    <!-- <div class="d-block m-auto">
      <form name="myForm" action="user_sessions_new.asp" method="post">
        <input type="hidden" name="user" value="<%=Session("userID")%>"/>
        <input type="hidden" name="latest" value="1" />
        <a href="javascript: previousResults();" class="header-section__link"
          >Previous results</a
        >
      </form>
    </div> -->
    <a
      href="#"
      class="header-section__link"
      onclick="swal({   title: 'Do you want to logout?',   text: '',   type: 'warning',   showCancelButton: true,   confirmButtonColor: '#DD6B55',   confirmButtonText: 'Yes',   closeOnConfirm: false }, function(){  window.location.href='index.asp?alt=logout'; return false; });"
    >
      Log out
    </a>
    <% END IF%>
  </div>
</div>