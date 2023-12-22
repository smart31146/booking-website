<!--#include file="../connections/bbg_conn.asp" -->

<%buffer=true%>
<%response.buffer=true%>
<%

Function FixStr(str)
    If Not IsNull(str) Then
        str=Replace(str, "'", "''")
        str=Replace(str, "\", "\\")
    End If
    FixStr=str
End Function 

Function PCase(strInput)
    Dim iPosition
    Dim iSpace
    Dim strOutput
    iPosition = 1
    Do While InStr(iPosition, strInput, " ", 1) <> 0
        iSpace = InStr(iPosition, strInput, " ", 1)
        strOutput = strOutput & UCase(Mid(strInput, iPosition, 1))
        strOutput = strOutput & LCase(Mid(strInput, iPosition + 1, iSpace - iPosition))
        iPosition = iSpace + 1
    Loop
    strOutput = strOutput & UCase(Mid(strInput, iPosition, 1))
    strOutput = strOutput & LCase(Mid(strInput, iPosition + 1))
    PCase = strOutput
End Function



' Settings for color
bgcolor		=	"#FFFFFF"
headercolor	=	"#FF6633"
textcolor	=	"#FFFFFF"

' On Error Resume Next
' this is general include file with all variables and frequently used functions
Session.LCID = 3081							' setting Australien date format
randomize()									' to make sure that the rundom number in quiz will be random

' Global variables
Dim Good_MS_VBScript_version				' General:	version of the best VB script for regexp etc
Dim client_name_short						' General: 	name of the client - short version
Dim client_name_long						' general: 	name of the client - long version
dim intranet_homepage						' General:	link back to client's legal intranet homepage
Dim noqis									' QUIZ: 	number of questions in a set
Dim barlng								    ' QUIZ: 	length of the navigation bar in Quiz
Dim client_IP								' General:	actual IP address of a visitor
Dim access_IP_array						    ' General:	array with all allowed IP addresses from preferences
Dim restricted							    ' General:	boolean with result of IP address protection test
Dim pref_bbg_avail						    ' Pref:		BBG available
Dim pref_training_avail					    ' Pref:		Training available
Dim pref_quiz_avail						    ' Pref:		Quiz available
Dim pref_qanda_avail					    ' Pref:		Q and A available
Dim pref_IP_list						    ' Pref:		list of IP addresses
Dim pref_IP_protect						    ' Pref:		the status of BBP IP address protection - True/False
Dim pref_admin_IP_list				    	' Pref:		list of Admin IP addresses
Dim pref_admin_IP_protect					' Pref:		the status of BBP Admin IP address protection - True/False
Dim pref_admin_color					    ' Pref:		Background Color
'Dim pref_training_force_a					' Pref: 	the status of forced answers in Training
'Dim pref_bbg_login					    	' Pref:		login required for BBG
'Dim pref_quiz_free						    ' Pref:		free login / password login allowed for Quiz
'Dim pref_bbg_splash						' Pref:		BBG splash page available
'Dim pref_bbg_pdf						    ' Pref:		PDF list available
Dim pref_forgot_pass					    ' Pref:		Enable/Disable forgot password function
Dim pref_change_pass					    ' Pref:		Enable/Disable change password function
Dim pref_self_reg					    	' Pref:		Enable/Disable self registration function
dim whatstr_full							' General:	string replacement - what to replace
dim bywhatstr_full							' General:	string replacement - by what to replace
dim module_bbg_full							' General:	string replacement - BBG replace  - True/False
dim module_tr_full							' General:	string replacement - Training replace - True/False
dim module_q_full							' General:	string replacement - Quiz replace - True/False
dim module_search_full						' General:	string replacement - Search page replace - True/False
dim whatstr									' General:	Array - string replacement - what to replace
dim bywhatstr								' General:	Array - string replacement - by what to replace
dim module_bbg								' General:	Array - string replacement - BBG replace - True/False
dim module_tr								' General:	Array - string replacement - Training replace - True/False
dim module_q								' General:	Array - string replacement - Quiz replace - True/False
dim module_search							' General:	Array - string replacement - Search page replace - True/False
dim countofreplace							' General:	number of all replacements
Dim database_date_string					' General:  the date specific string for different database
'Dim delim_welcome_start						' Admin:	text file deliminator welcome start
'Dim delim_welcome_end						' Admin:	text file deliminator welcome end
'Dim delim_welcome1_start					' Admin:	text file deliminator welcome start
'Dim delim_welcome1_end						' Admin:	text file deliminator welcome end
'Dim delim_welcome2_start					' Admin:	text file deliminator welcome start
'Dim delim_welcome2_end						' Admin:	text file deliminator welcome end
'Dim delim_f1_start							' Admin:	text file deliminator focus1 start
'Dim delim_f1_end							' Admin:	text file deliminator focus1 end
'Dim delim_f1_icon_start						' Admin:	text file deliminator focus1 icon start
'Dim delim_f1_icon_end						' Admin:	text file deliminator focus1 icon end
'Dim delim_f1_url_start						' Admin:	text file deliminator focus1 url start
'Dim delim_f1_url_end						' Admin:	text file deliminator focus1 url end
'Dim delim_f2_start							' Admin:	text file deliminator focus2 start
'Dim delim_f2_end							' Admin:	text file deliminator focus2 end
'Dim delim_f2_icon_start						' Admin:	text file deliminator focus2 icon start
'Dim delim_f2_icon_end						' Admin:	text file deliminator focus2 icon end
'Dim delim_f2_url_start						' Admin:	text file deliminator focus2 url start
'Dim delim_f2_url_end						' Admin:	text file deliminator focus2 url end

'Dim delim_f3_start							' Admin:	text file deliminator focus2 start
'Dim delim_f3_end							' Admin:	text file deliminator focus2 end
'Dim delim_f3_icon_start						' Admin:	text file deliminator focus2 icon start
'Dim delim_f3_icon_end						' Admin:	text file deliminator focus2 icon end
'Dim delim_f3_url_start						' Admin:	text file deliminator focus2 url start
'Dim delim_f3_url_end
'Dim homepagetextfile						' Admin:	name of the file to store the HomePage information


' Local variables
Dim tested_IP							    ' Local: 	runtime variable
Dim bbp_version								' Local:	current version

' preset variables
database_date_string = "#" ' for MS Access
database_date_string = "'" ' for MS SQL server 2000

'HomePageTextFile = "client/homepage.txt"
bbp_version = "version 1.4 final SQL - 14.5.2002"
Good_MS_VBScript_version = 5
noqis = 3
barlng = 150
client_name_short = "TWE"
client_name_long = "TWE"
intranet_homepage = "http://"
client_homepage = "https://www.lawofthejungle.com/twe"
restricted = True
client_IP = "127.0.0.1"
whatstr_full = ""
bywhatstr_full = ""
module_bbg_full = ""
module_tr_full = ""
module_q_full = ""
module_search_full = ""
countofreplace = 0
Dim homeURL
homeURL="index.asp"
'delim_welcome_start = "###welcome_start###"
'delim_welcome_end = "###welcome_end###"
'delim_welcome1_start = "###welcome1_start###"
'delim_welcome1_end = "###welcome1_end###"
'delim_welcome2_start = "###welcome2_start###"
'delim_welcome2_end = "###welcome2_end###"
'delim_f1_start = "###focus1_start###"
'delim_f1_end = "###focus1_end###"
'delim_f1_icon_start = "###focus1_icon_start###"
'delim_f1_icon_end = "###focus1_icon_end###"
'delim_f1_url_start = "###focus1_url_start###"
'delim_f1_url_end = "###focus1_url_end###"
'delim_f2_start = "###focus2_start###"
'delim_f2_end = "###focus2_end###"
'delim_f2_icon_start = "###focus2_icon_start###"
'delim_f2_icon_end = "###focus2_icon_end###"
'delim_f2_url_start = "###focus2_url_start###"
'delim_f2_url_end = "###focus2_url_end###"

'delim_f3_start = "###focus3_start###"
'delim_f3_end = "###focus3_end###"
'delim_f3_icon_start = "###focus3_icon_start###"
'delim_f3_icon_end = "###focus3_icon_end###"
'delim_f3_url_start = "###focus3_url_start###"
'delim_f3_url_end = "###focus3_url_end###"

pref_bbg_avail = True
pref_training_avail = True
pref_quiz_avail = True
pref_qanda_avail = True
pref_IP_list = ""
pref_IP_protect = False
pref_admin_IP_list = ""
pref_admin_IP_protect = False
pref_forgot_pass = True
pref_change_pass = True
pref_self_reg = True
pref_admin_color=""
'pref_training_force_a = True
'pref_bbg_login = False
'pref_quiz_free = True
'pref_bbg_splash = True
'pref_bbg_pdf = True

'this function adjusts the local time to different time zone
FUNCTION Now_BBP
    Now_BBP = cDate(Year(Now()) & "/" & Month(Now()) & "/" & Day(Now()) & " " & Hour(Now()) & ":" & Minute(Now()) & ":" &Second(Now()))
END FUNCTION


' Set up preferences + security test
' this test is to protect the BBP agains unwanted visitors in case the IP address protection
' is turned on in preferrences
' the most recent active profile is the effective one

set preferences = Server.CreateObject("ADODB.Recordset")
preferences.ActiveConnection = Connect
preferences.Source = "SELECT * FROM preferences WHERE (((preferences.pref_active)=1)) ORDER BY preferences.pref_date DESC;"
preferences.CursorType = 0
preferences.CursorLocation = 3
preferences.LockType = 3
preferences.Open()
preferences_numRows = 0


' if there is a saved profile lets get it
if NOT preferences.EOF then

    pref_bbg_avail = CBool(preferences.Fields.Item("pref_bbg_avail").Value)
    'pref_training_avail = CBool(preferences.Fields.Item("pref_training_avail").Value)
    pref_quiz_avail = CBool(preferences.Fields.Item("pref_quiz_avail").Value)
    'pref_qanda_avail = CBool(preferences.Fields.Item("pref_qanda_avail").Value)
    pref_IP_list = CStr(preferences.Fields.Item("pref_IP_list").Value)
    pref_IP_protect = CBool(preferences.Fields.Item("pref_IP_protect").Value)
    pref_admin_IP_list = CStr(preferences.Fields.Item("pref_admin_IP_list").Value)
    pref_admin_IP_protect = CBool(preferences.Fields.Item("pref_admin_IP_protect").Value)
    pref_forgot_pass = CBool(preferences.Fields.Item("pref_forgot_pass").Value)
    pref_change_pass = CBool(preferences.Fields.Item("pref_change_pass").Value)
    pref_self_reg = CBool(preferences.Fields.Item("pref_self_reg").Value)
	'pref_admin_color=CStr(preferences.Fields.Item("pref_admin_color").Value)
    'pref_training_force_a = CBool(preferences.Fields.Item("pref_training_force_a").Value)
    'pref_bbg_login = CBool(preferences.Fields.Item("pref_bbg_login").Value)
    'pref_quiz_free = CBool(preferences.Fields.Item("pref_quiz_free").Value)
    'pref_bbg_splash = CBool(preferences.Fields.Item("pref_bbg_splash").Value)
    'pref_bbg_pdf = CBool(preferences.Fields.Item("pref_bbg_pdf").Value)

    adminurl=Request.ServerVariables("PATH_INFO")
    'Response.Write adminurl
    isadmin=instr(adminurl,"admin")
    'Response.Write isadmin

    if cint(isadmin) <> 0 then
        if pref_admin_IP_protect = True then
        client_IP = Request.ServerVariables("REMOTE_ADDR")
        if pref_admin_IP_list <> "" then
            pref_admin_IP_list = replace(pref_admin_IP_list," ","")
            access_IP_array = Split(pref_admin_IP_list, ";")
            for iii = 0 to uBound(access_IP_array)
                tested_IP = access_IP_array(iii)
                if (InStr(tested_IP,"*")) <> 0 then
                    tested_IP = Left(tested_IP,(InStr(tested_IP,"*")-1))
                    if inStr(client_IP,tested_IP) > 0 then restricted = False
                end if
                if (client_IP = tested_IP) then restricted = False

            next
            if restricted then response.redirect("ip_error.asp")
        end if
        end if
    else if pref_IP_protect = True then
        client_IP = Request.ServerVariables("REMOTE_ADDR")
        if pref_IP_list <> "" then
            pref_IP_list = replace(pref_IP_list," ","")
            access_IP_array = Split(pref_IP_list, ";")
            for iii = 0 to uBound(access_IP_array)
                tested_IP = access_IP_array(iii)
                if (InStr(tested_IP,"*")) <> 0 then
                    tested_IP = Left(tested_IP,(InStr(tested_IP,"*")-1))
                    if inStr(client_IP,tested_IP) > 0 then restricted = False
                end if
                if (client_IP = tested_IP) then restricted = False

            next
            if restricted then response.redirect("ip_error.asp")
        end if
    end if
end if
end if

preferences.Close()
Set preferences = Nothing
' end of IP address protection test

' myregexp - to keep all regular expressions on one spot
Function myregexp(thestring, thepattern, thereplacement)
    if cInt(ScriptEngineMinorVersion) >= Good_MS_VBScript_version then
        dim regEx
        set regEx = New RegExp
' 		set regEx = Server.CreateObject("RE")
        regEx.IgnoreCase = True
        regEx.Global = True
        regEx.Pattern = thepattern
        myregexp = regEx.Replace(thestring, thereplacement)
' 		myregexp = regEx.subst(thestring,thereplacement,thepattern)
        set regEx = nothing
    else
        myregexp = thestring
    end if
End Function


' Highlighting search results
function Highlight(theField)
    Set objRegExp= New RegExp
' 	set regEx = Server.CreateObject("RE")
    if result_str <> "" then
        if (search_mode = "and") or (search_mode = "or") then
            sarySearchWord = Split(Trim(result_str), " ")
            Dim intLoopCounter
            intLoopCounter = 0
            For intLoopCounter = 0 To UBound(sarySearchWord)
                objRegExp.Pattern="(" & sarySearchWord(intLoopCounter) & ")"
                objRegExp.IgnoreCase=True
                objRegExp.Global=True
                'theField=objRegExp.Replace(theField,"<span class='search_hl'>$1</span>")
'		 		theField=regEx.subst(theField,"<span class='search_hl'>" & sarySearchWord(intLoopCounter) & "</span>","(" & sarySearchWord(intLoopCounter) & ")")
            Next
            Highlight = theField
     else
            objRegExp.Pattern="(" & result_str & ")"
            objRegExp.IgnoreCase=True
            objRegExp.Global=True
            Highlight=objRegExp.Replace(theField,"$1")
'			Highlight=regEx.subst(theField,"<span class='search_hl'>" & result_str & "</span>","(" & result_str & ")")

        end if
    else
    Highlight = theField
    end if
end Function


' function which tests wheather a file exist
Function fileexist(whichfile)
    fileexist = false
    Dim objFSO
    Dim checkpath
    checkpath = Trim(whichfile)
    if cInt(ScriptEngineMinorVersion) >= Good_MS_VBScript_version then
        if Left(checkpath,1) = "/" then
            checkpath = ".." & checkpath
        elseif (Left(checkpath,1) <> "/") and (Left(checkpath,2) <> "..") then
            checkpath = "../" & checkpath
        end if
        checkpath = Trim(Server.mappath(checkpath))
    end if
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(checkpath) Then fileexist = True
    Set objFSO = Nothing
End Function
' end of fileexist function

' function which removes all HTML tags
function ClearHTMLTags(strHTML, intWorkFlow)
'	intWorkFlow: An integer that if equals to 0 runs only the RegExp filter
'              .. 1 runs only the HTML source render filter
'              .. 2 runs both the RegExp and the HTML source render
'              .. >2 defaults to 0
    strTagLess = strHTML
      if intWorkFlow <> 1 then
        strTagLess = myregexp(strTagLess, "<[^>]*>", "")
      end if

      if intWorkFlow > 0 and intWorkFlow < 3 then
        strTagLess = myregexp(strTagLess, "[<]", "<")
        strTagLess = myregexp(strTagLess, "[>]", ">")
      end if

      ClearHTMLTags = strTagLess
end function
' end of strip HTML tags function

' converts date to SQL friendly format (yyyy/mm/dd)
Function cDateSql(ByVal Date_prm)
    if isdate(Date_prm) then
        cDateSql = Year(Date_prm)  & "/" & Right("0" & Month(Date_prm), 2) & "/" & Right("0" & Day(Date_prm), 2)& " " & Hour(Date_prm) & ":" & Minute(Date_prm) & ":" & Second(Date_prm)
        'cDateSql = Right("0" & Day(Date_prm), 2)  & "/" & Right("0" & Month(Date_prm), 2) & "/" & Year(Date_prm)& " " & Hour(Date_prm) & ":" & Minute(Date_prm) & ":" & Second(Date_prm)
    else
        Date_prm = ""
    end if
End Function


' general replace tool
set toreplace = Server.CreateObject("ADODB.Recordset")
toreplace.ActiveConnection = Connect
toreplace.Source = "SELECT * FROM toreplace  WHERE (Abs([repl_active]))=1;"
toreplace.CursorType = 0
toreplace.CursorLocation = 3
toreplace.LockType = 3
toreplace.Open()


' create replace arrays
While (NOT toreplace.EOF) AND (NOT toreplace.BOF)
if whatstr_full = "" then whatstr_full = cStr(toreplace.Fields.Item("repl_what").Value) else whatstr_full = whatstr_full & "|" & cStr(toreplace.Fields.Item("repl_what").Value)
if bywhatstr_full = "" then bywhatstr_full = cStr(toreplace.Fields.Item("repl_bywhat").Value) else bywhatstr_full = bywhatstr_full & "|" & cStr(toreplace.Fields.Item("repl_bywhat").Value)
if module_bbg_full = "" then module_bbg_full = abs(cInt(toreplace.Fields.Item("repl_bbg").Value)) else module_bbg_full = module_bbg_full & "|" & abs(cInt(toreplace.Fields.Item("repl_bbg").Value))
if module_tr_full = "" then module_tr_full = abs(cInt(toreplace.Fields.Item("repl_tr").Value)) else module_tr_full = module_tr_full & "|" & abs(cInt(toreplace.Fields.Item("repl_tr").Value))
if module_q_full = "" then module_q_full = abs(cInt(toreplace.Fields.Item("repl_q").Value)) else module_q_full = module_q_full & "|" & abs(cInt(toreplace.Fields.Item("repl_q").Value))
if module_search_full = "" then module_search_full = abs(cInt(toreplace.Fields.Item("repl_search").Value)) else module_search_full = module_search_full & "|" & abs(cInt(toreplace.Fields.Item("repl_search").Value))
countofreplace = countofreplace +1
toreplace.MoveNext()
Wend

toreplace.Close()
Set toreplace = Nothing

' string to array
whatstr = Split(whatstr_full, "|")
bywhatstr = Split(bywhatstr_full, "|")
module_bbg = Split(module_bbg_full, "|")
module_tr = Split(module_tr_full, "|")
module_q = Split(module_q_full, "|")
module_search = Split(module_search_full, "|")


' replacement for BBG pages
function ReplaceStrBBG(theField)
    ReplaceStrBBG = theField
    if theField <> "" then
        if whatstr_full <> "" then
            for iii = 1 to countofreplace
                if (module_bbg(iii-1)) = "1"  then ReplaceStrBBG=myregexp(ReplaceStrBBG, "(" & whatstr(iii-1) & ")", bywhatstr(iii-1))
            next
        end if
    end if
end function

' replacement for Training pages
function ReplaceStrTR(theField)
    ReplaceStrTR = theField
    if theField <> "" then
        if whatstr_full <> "" then
            for iii = 1 to countofreplace
                if (module_tr(iii-1)) = "1"  then ReplaceStrTR=myregexp(ReplaceStrTR, "(" & whatstr(iii-1) & ")", bywhatstr(iii-1))
            next
        end if
    end if
end function

' replacement for Quiz pages
function ReplaceStrQuiz(theField)
    ReplaceStrQuiz = theField
    if theField <> "" then
        if whatstr_full <> "" then
            for iii = 1 to countofreplace
                if (module_q(iii-1)) = "1"  then ReplaceStrQuiz=myregexp(ReplaceStrQuiz, "(" & whatstr(iii-1) & ")", bywhatstr(iii-1))
            next
        end if
    end if
end function

' replacement for Search pages
function ReplaceStrSearch(theField)
    ReplaceStrSearch = theField
    if theField <> "" then
        if whatstr_full <> "" then
            for iii = 1 to countofreplace
                if (module_search(iii-1)) = "1"  then ReplaceStrSearch=myregexp(ReplaceStrSearch, "(" & whatstr(iii-1) & ")", bywhatstr(iii-1))
            next
        end if
    end if
end function

' reverse replacement
function RevReplaceStr(theField)
    RevReplaceStr = theField
    if theField <> "" then
        if bywhatstr_full <> "" then
            for iii = 1 to countofreplace
                if (module_search(iii-1)) = "1" then RevReplaceStr=myregexp(RevReplaceStr, "(" & bywhatstr(iii-1) & ")", whatstr(iii-1))
            next
        end if
    end if
end function


' logging function
Function log_the_page(log_module, log_subjID, log_subj, log_topicID, log_topic, log_pageID, log_page, log_comment)
    MM_editConnection = Connect
    MM_editTable = "logs"

    log_date = cDateSql(Now_BBP())
    log_ip = Left(Request.ServerVariables("REMOTE_ADDR"),15)
    log_session = right(Request.ServerVariables("HTTP_COOKIE"),45)
    if log_session = "" then log_session = "not yet available"
    log_url = Left(Request.ServerVariables("URL") & "?" & Request.QueryString,100)
    log_agent = Left(Request.ServerVariables("HTTP_USER_AGENT"),100)
    if log_comment = "" then log_comment = "n/a"
    if Session("UserID") <> "" then log_userID = Session("UserID") else log_userID = 0
    if Session("lastname") <> "" then log_user = Replace(Session("firstname"),"'","''") & " " & Replace(Session("lastname"),"'","''") else log_user = "n/a"

    '***************************************************************************
    'GC - Added fixstr to ensure that single and double quotes in content don't cause errors 10/08/2013
    '***************************************************************************
	
	' Added fix to allow emails with apostrophe for username forgot functionality by PR 15.07.2016 JIRA:BBP-70
    MM_tableValues = "log_session, log_date, log_ip, log_module, log_userID, log_user, log_subjID, log_subj, log_topicID, log_topic, log_pageID, log_page, log_url, log_agent, log_comment"
    MM_dbValues = "'" & log_session & "', '"& log_date & "', '" & log_ip & "', '" & log_module & "', " & log_userID & ", '" & log_user & "', " & log_subjID & ", '" & fixstr(log_subj) & "', " & log_topicID & ", '" & fixstr(log_topic) & "', " & log_pageID & ", '" & log_page & "', '" & log_url & "', '" & log_agent & "', '"& replace(log_comment, "'", "''" ) & "'"

    MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ");"
    'Response.Write MM_editQuery
    'response.end

    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    Set MM_editCmd = Nothing
end Function


'Generate password
Function getRandomNum(lbound, ubound)
For j = 1 To (250 - ubound)
    Randomize
    getRandomNum = Int(((ubound - lbound) * Rnd) + 1)
Next
End Function

Function getRandomChar(number, lower, upper, other, extra)
numberChars = "0123456789"
lowerChars = "abcdefghijklmnopqrstuvwxyz"
upperChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
otherChars = "`~!@#$%^&*()-_=+[{]}\\|;:"""'\,<.>/? "
charSet = extra
    if (number = "true") Then charSet = charSet + numberChars
    if (lower = "true") Then charSet = charSet + lowerChars
    if (upper = "true") Then charSet = charSet + upperChars
    if (other = "true") Then charSet = charSet + otherChars
jmi = Len(charSet)
getRandomChar = Mid(charSet, getRandomNum(1, jmi), 1)
End Function

Function getPassword(length, extraChars, firstNumber, firstLower, firstUpper, firstOther, latterNumber, latterLower, latterUpper, latterOther)
rc = ""
    If (length > 0) Then
        rc = rc + getRandomChar(firstNumber, firstLower, firstUpper, firstOther, extraChars)
    End If

    For idx = 1 To length - 1
        rc = rc + getRandomChar(latterNumber, latterLower, latterUpper, latterOther, extraChars)
    Next
getPassword = rc
End Function

' generate unique ID for upload/download
Function GetUniqueID(prefix,idlength,suffix)
    prefix_length = len(prefix)
    suffix_length = len(suffix)
    originaldate = cDate("1/1/1900")
    numofseconds = cStr(abs(DateDiff("s",Now,originaldate)))
    time_length = len(numofseconds)
    if idlength < (prefix_length + suffix_length + time_length) then
        idlength = 3
    else
        idlength = idlength - prefix_length - suffix_length - time_length
    end if
    GetUniqueID = cStr(prefix) & numofseconds & getPassword(idlength, "", "true", "true", "true", "false", "true", "true", "true", "false") & cStr(suffix)
End Function


FUNCTION CropSentence(strText, intLength, strTrial)
  Dim wsCount
  Dim intTempSize
  Dim intTotalLen
  Dim strTemp

  wsCount = 0
  intTempSize = 0
  intTotalLen = 0
  intLength = intLength - Len(strTrial)
  strTemp = ""

  IF Len(strText) > intLength THEN
    arrTemp = Split(strText, " ")
    FOR EACH x IN arrTemp
      IF Len(strTemp) <= intLength THEN
        strTemp = strTemp & x & " "
      END IF
    NEXT
      CropSentence = Left(strTemp, Len(strTemp) - 1) & strTrial
  ELSE
    CropSentence = strText
  END IF
END FUNCTION
%>
