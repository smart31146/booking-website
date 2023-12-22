<%

Function EnCrypt(strCryptThis)
Dim strChar, iKeyChar, iStringChar, i,g_Key
g_Key = mid(ReadKeyFromFile(),1,Len(strCryptThis))
for i = 1 to Len(strCryptThis)
iKeyChar = Asc(mid(g_Key,i,1))
iStringChar = Asc(mid(strCryptThis,i,1))
iCryptChar = iKeyChar Xor iStringChar 
iCryptCharHex = Hex(iCryptChar)
iCryptCharHexStr = cstr(iCryptCharHex)
if len(iCryptCharHexStr)=1 then iCryptCharHexStr = "0" & iCryptCharHexStr
strEncrypted = strEncrypted & iCryptCharHexStr
next
EnCrypt = strEncrypted
End Function


Function DeCrypt(strEncrypted)
Dim strChar, iKeyChar, iStringChar, i,g_Key
g_Key = mid(ReadKeyFromFile(),1,Len(strEncrypted)/2)
for i = 1 to Len(strEncrypted) /2
iKeyChar = (Asc(mid(g_Key,i,1)))
iStringChar2 = mid(strEncrypted,(i*2)-1,2)
iStringChar = CLng("&H" & iStringChar2)
iDeCryptChar = iKeyChar Xor iStringChar
strDecrypted = strDecrypted & chr(iDeCryptChar)
next
DeCrypt = strDecrypted
End Function
Function ReadKeyFromFile()

'Const strFileName = "C:\acdelar.txt" 'Change this string 
Const strFileName = "D:\webhotel\database\acdelar.txt" 'Change this string 
Dim keyFile, fso, f
set fso = Server.CreateObject("Scripting.FileSystemObject") 
set f = fso.GetFile(strFileName) 
set ts = f.OpenAsTextStream(1, -2)
Do While not ts.AtEndOfStream
keyFile = keyFile & ts.ReadLine
Loop 
ReadKeyFromFile = keyFile
End Function

%>