<%@ LANGUAGE=VBScript CodePage=65001%> 
<%Option Explicit 
Class GoogleTranslator 
sub Class_Initialize() 
RURI="http://translate.google.com/translate_t?langpair={0}&text={1}" 
End Sub 
Private Opt_ ' 
Property Get Opt 
Opt=Opt_ 
End Property 
Property Let Opt(Opt_s) 
Opt_=Opt_s 
End Property 
Private RURI 
Function AnalyzeChild(patrn,texts,IPos) 
Dim regEx, Match, Matches 
Set regEx = New RegExp 
regEx.IgnoreCase = true 
regEx.Global = True 
regEx.Pattern = patrn 
regEx.Multiline = True 
Dim RetStr 
Set Matches = regEx.Execute(texts) 
If(Matches.Count > 0)Then RetStr= Matches(0).SubMatches(IPos) 
AnalyzeChild=RetStr 
Set regEx =Nothing 
End Function 
Function getHTTPPage(url) 
dim objXML 
set objXML=server.createobject("MSXML2.XMLHTTP")'定义 
objXML.open "GET",url,false'打开 
objXML.send()'发送 
If objXML.readystate<>4 then 
exit function 
End If 
getHTTPPage=BytesToBstr(objXML.responseBody) 
set objXML=nothing'关闭 
if err.number<>0 then err.Clear 
End Function 
Function BytesToBstr(body) 
dim objstream 
set objstream = Server.CreateObject("adodb.stream") 
objstream.Type = 1 
objstream.Mode =3 
objstream.Open 
objstream.Write body 
objstream.Position = 0 
objstream.Type = 2 
objstream.Charset = "utf-8" 
'转换原来默认的UTF-8编码转换成GB2312编码，否则直接用XMLHTTP调用有中文字符的网页得到的将是乱码 
BytesToBstr = objstream.ReadText 
objstream.Close 
set objstream = nothing 
End Function 
Public Function GetText(str) 
If(isempty(str)) Then Exit Function 
Dim newUrl,Rs 
newUrl=Replace(Replace(RURI,"{0}",Server.URLEncode(Opt)),"{1}",Server.URLEncode(str)) 
Rs=getHTTPPage(newUrl) 
GetText = AnalyzeChild("(<div id=result_box dir=""ltr"">)([?:\s\S]*?)(</div>)",Rs,1) 
End Function 
Sub class_Terminate 
End Sub 
End Class 
%>

<%
Dim Obj 
Set Obj = new GoogleTranslator 
Obj.Opt="zh-CN|en" 
response.write(Obj.GetText("我们")) 
%>