<% 
'  SqlCheckInclude.asp
'  This is the include file to use with your asp pages to 
'  validate input for SQL injection.
'  Developed by PR 25.02.16

Response.Buffer = True
Dim BlackList, ErrorPage, st



BlackList = Array("--", ";", "/*", "*/", "@@",_
                  "char", "nchar", "varchar", "nvarchar",_
                  "alter", "begin", "cast", "create", "cursor",_
                  "declare", "delete", "drop", "end", "exec",_
                  "execute", "fetch", "insert", "kill", "open",_
                  "select", "sys", "sysobjects", "syscolumns",_
                  "table", "update")

'  Populate the error page you want to redirect to in case the 
'  check fails.

ErrorPage = "errorhandler/input_error.asp"
               

Function CheckStringForSQL(str) 
  On Error Resume Next 
  
  Dim lstr 
  
  ' If the string is empty, return true
  If ( IsEmpty(str) ) Then
    CheckStringForSQL = false
    Exit Function
  ElseIf ( StrComp(str, "") = 0 ) Then
    CheckStringForSQL = false
    Exit Function
  End If
  
  lstr = LCase(str)
  
  ' Check if the string contains any patterns in our
  ' black list
  For Each st in BlackList
  
    If ( InStr (lstr, st) <> 0 ) Then
      CheckStringForSQL = true
      Exit Function
    End If
  
  Next
  
  CheckStringForSQL = false
  
End Function 


'''''''''''''''''''''''''''''''''''''''''''''''''''
'  Check forms data
'''''''''''''''''''''''''''''''''''''''''''''''''''

For Each st in Request.Form
  If ( CheckStringForSQL(Request.Form(st)) ) Then
  
    ' Redirect to an error page
	Response.Clear
	Response.Redirect(ErrorPage)
  
  End If
Next



'''''''''''''''''''''''''''''''''''''''''''''''''''
'  Check query string
'''''''''''''''''''''''''''''''''''''''''''''''''''

For Each st in Request.QueryString
  If ( CheckStringForSQL(Request.QueryString(st)) ) Then
  
    ' Redirect to error page
	Response.Clear
	Response.Redirect(ErrorPage)

    End If

  
Next


'''''''''''''''''''''''''''''''''''''''''''''''''''
'  Check cookies
'''''''''''''''''''''''''''''''''''''''''''''''''''

For Each st in Request.Cookies
  If ( CheckStringForSQL(Request.Cookies(st)) ) Then
  
    ' Redirect to error page
	Response.Clear
	Response.Redirect(ErrorPage)

  End If
  
Next
%>