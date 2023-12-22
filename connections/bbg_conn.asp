<%
' FileName="Connection_odbc_conn_dsn.htm"
' Type="ADO"
' HTTP="false"
' Catalog=""
' Schema=""
 <!-- Connect = "Provider=sqloledb;Network Library=DBMSSOCN; Data Source=localhost,1433; Initial Catalog=BBP_twe_live_erq; User ID=sa; Password=123456789;" -->
 Connect = "Provider=sqloledb;Network Library=DBMSSOCN; Data Source=91.221.66.87,1433; Initial Catalog=BBP_twe_live_erq; User ID=sa; Password=123456789;"

set obj = Server.CreateObject("ADODB.Recordset")
session.Timeout=1440

BBPinfo3 = "Activity"
BBPinfo3s = "Activities"

If Session("userID") <> "" THEN
	Session("userID") = Session("userID")
	Session("firstname") = Session("firstname")
	Session("lastname") = Session("lastname")
	Session("id") = Session("id") 
	Session("name") = Session("name")
END IF

if  request("bbp_search")<>"" then
	Session("bbp_search") = request("bbp_search")
ELSE
	if not isEmpty(Session("bbp_search")) then
		Session("bbp_search")=Session("bbp_search")
	elseif Session("bbp_search") <> "" then
		Session("bbp_search")=Session("bbp_search")	
	else
		Session("bbp_search")="Enter a keyword"
	end if
END IF
%>
