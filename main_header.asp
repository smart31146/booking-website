<div style="position:relative;margin-bottom:2px;height:140px;width:"><a href="index.asp" target="_top"><img src="bilder/logotype.gif" width="690" height="140" alt=""></A></td>

<div style="position:absolute;top:10px;left:400px;">
<form action="i_search.asp" method="post"><strong>SÖK VARA</strong><br>
<input autofocus type="text" name="ica_search" style="width:140px;" class="finput_small" value="<% =Session("search_string")%>">
<input type="image" src="bilder/b_search.gif" align="middle"><br></form>
</div>

<div style="position:absolute;top:0px;right:0px;width:250px;height:140px;background-color:#d22728">
<br><A href="ica_konto.asp"  class="menu"><img src="bilder/b_skapa.gif" height="26" alt="Skapa konto och börja handla idag!"></a>&nbsp;&nbsp;&nbsp;&nbsp;
<img src="bilder/b_loggain.gif" width="79" height="26" alt="">&nbsp;
<FORM name="form" method="post" action="sida_kontroll.asp?alt=index"><br>
<input type="text" class="finput_small" name="anvnamn" style="width:80px;margin-bottom:10px;">&nbsp;
<input type="password" class="finput_small" name="losenord" style="width:80px;margin-bottom:10px;"><br>
<label for="remember"><input type="checkbox" name="remember" value="ja"> Kom ihåg mig!</label>&nbsp;&nbsp;
<input type="submit" value="Logga in" style="vertical-align:middle" title="Logga in" ><br></form>
</div>
</div><!--<A href="ica_konto.asp" rel="lyteframe" title="" rev="width: 760px; height: 600px; scrolling: no;top: 5px;" class="menu" title="ÅF LOGIN"><img src="bilder/top_logo.gif" width="950" height="146" alt=""></A><br>-->
