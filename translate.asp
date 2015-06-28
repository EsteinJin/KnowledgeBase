<html>
<head>
<title>在線翻譯</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>

<body>
<%
'on error resume next
' 如果網速很慢的話，可以調整以下時間。單位秒
Server.ScriptTimeout = 999999
'========================================================
'字符編碼函數
'========================================================
Function BytesToBstr(body,code) 
dim objstream 
set objstream = Server.CreateObject("adodb.stream") 
objstream.Type = 1 
objstream.Mode =3 
objstream.Open 
objstream.Write body 
objstream.Position = 0 
objstream.Type = 2 
objstream.Charset =code
BytesToBstr = objstream.ReadText 
objstream.Close 
set objstream = nothing 
End Function 

'取行字符串在另一字符串中的出現位置
Function Newstring(wstr,strng) 
Newstring=Instr(lcase(wstr),lcase(strng)) 
if Newstring<=0 then Newstring=Len(wstr) 
End Function 
'替換字符串函數
function ReplaceStr(ori,str1,str2)
ReplaceStr=replace(ori,str1,str2)
end function
'=====================================================
function ReadXml(url,code,start,ends)
set oSend=createobject("Microsoft.XMLHTTP")
SourceCode = oSend.open ("GET",url,false) 
oSend.send()
ReadXml=BytesToBstr(oSend.responseBody,code )
if(start="" or ends="") then
else
start=Newstring(ReadXml,start)
ReadXml=mid(ReadXml,start)
ends=Newstring(ReadXml,ends)
ReadXml=left(ReadXml,ends-1)
end if
end function
dim urlpage,lan
urlpage=request("urls")
lan=request("lan")
%>
<form method="post" action="translate.asp">
<input type="text" name="urls" size="150" value="<%=urlpage%>">
<input type="hidden" name="lan" value="<%=lan%>">
<input type="submit" value="submit">
</form>
<%
dim transURL
transURL="http://216.239.39.104/translate_c?hl=zh-CN&ie=UTF-8&oe=UTF-8&langpair="&server.URLEncode(lan)&"&u="&urlpage&"&prev=/language_tools"
if(len(urlpage)>3) then
getcont=ReadXml(transURL,"gb2312","","")
response.Write(getcont)
end if

%>
</body>
</html>