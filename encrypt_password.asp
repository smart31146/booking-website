<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="connections/bbg_conn.asp"-->
<!--#include file="connections/include.asp"-->
<!--#include file="sha256.asp"-->


<% 
Set obj = Server.CreateObject("ADODB.Recordset")
SQL="SELECT * FROM q_user"
obj.ActiveConnection = Connect
obj.Source = SQL 
obj.CursorType = 0
obj.CursorLocation = 3
obj.LockType = 3
obj.Open

If obj.EOF then
Response.write("The END")

Else

Do While Not obj.EOF
Dim salt
salt = obj("user_email")
password=obj("user_city")
password=password&salt
password=sha256(password)
Set uobj = Server.CreateObject("ADODB.Command")
SQL="update q_user set user_city='"&password&"' WHERE ID_User="&obj("ID_USER")
uobj.ActiveConnection = Connect
uobj.CommandText = SQL
uobj.Execute
uobj.ActiveConnection.Close

response.write(obj("user_firstname") & "<br>")

obj.MoveNext
Loop
End If

obj.close

%>

