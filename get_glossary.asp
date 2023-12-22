<%@LANGUAGE="VBSCRIPT"%>

<!--#include file="connections/bbg_conn.asp" -->


<%
Dim Products
Dim Products_cmd
Dim Products_numRows

Set Products_cmd = Server.CreateObject ("ADODB.Command")
Products_cmd.ActiveConnection = Connect
Products_cmd.CommandText = "SELECT gid, name, description, active FROM glossary" 
Products_cmd.Prepared = true

Set Products = Products_cmd.Execute

Dim arrProducts

arrProducts = Products.GetRows()



dim i, strQueryString
strQueryString ="["
For i = 0 to ubound(arrProducts,2)

'strQueryString ='{"a":1,"b":2}'
'"','" & Products(1).Name & "':'" & trim(arrProducts(1,i)) & "'
'strQueryString  = strQueryString  & """" & trim(arrProducts(1,i)) & ""","

 strQueryString = strQueryString &  "{" & """"&Products(1).Name&"""" &":""" & trim(arrProducts(1,i)) &"""" & "," & """"&Products(2).Name&"""" & ":""" & Escape(trim(arrProducts(2,i))) &"""" & "}"
     strQueryString = strQueryString & ","
	  
next
'Cut off the last character
   strQueryString = Left(strQueryString, Len(strQueryString) - 1)
strQueryString = strQueryString & "]"
response.write  strQueryString
'dim i
'response.write "<table>"
'For i = 0 to ubound(arrProducts, 2)
   'response.write "<tr>"
   
  ' response.write("<td>" + trim(arrProducts(0,i)))
   'response.write("<td>" + trim(arrProducts(1,i)))
  ' response.write("<td>" + trim(arrProducts(2,i)))
    'response.write("<td>" + trim(arrProducts(3,i)))
'next
'response.write "</table>"
Function Escape(sString)

    'Replace any Cr and Lf to <br>
    strReturn = Replace(sString , vbCrLf, "\n")     : 'visual basic carriage return line feed     ***********\
    strReturn = Replace(strReturn , vbCr , "\n")     : 'visual basic carriage return               *********These 3 are line breaks hence <BR>
    strReturn = Replace(strReturn , vbLf , "\n")     : 'visual basic line feed          ***********/
    strReturn = Replace(strReturn, "'", "''")          : 'Single quote changed to 2 single quotes ASP knows what to do
    Escape = strReturn
End Function

%>

