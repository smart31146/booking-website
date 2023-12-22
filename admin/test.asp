<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../connections/bbg_conn.asp" -->
<!--#include file="../connections/include_admin.asp" -->

<%

set filter_info1 = Server.CreateObject("ADODB.Recordset")
filter_info1.ActiveConnection = Connect
filter_info1.Source = "SELECT * FROM logs where id_log between 103391 and 105048 order by id_log desc"
filter_info1.CursorType = 0
filter_info1.CursorLocation = 3
filter_info1.LockType = 3
filter_info1.Open()
filter_info1_numRows = 0
  MM_editConnection = Connect
Set MM_editCmd = Server.CreateObject("ADODB.Command")
MM_editCmd.ActiveConnection = MM_editConnection
	

While (NOT filter_info1.EOF)
	tmp = filter_info1.Fields.Item("id_log").Value
	MM_editQuery = "UPDATE logs SET log_date = (SELECT REPLACE ((SELECT log_date FROM logs WHERE id_log = '"&tmp&"') , '/', '-') AS Expr1) WHERE (id_log = '"&tmp&"')"
	MM_editCmd.CommandText = MM_editQuery
	response.write MM_editQuery
    MM_editCmd.Execute            
filter_info1.MoveNext()
Wend

filter_info1.Close()
%>
   