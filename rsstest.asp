<%@codepage = 936%>
<% @language="VBScript"%>

<%
Function readrss(xmlseed)
dim xmlDoc 
dim http
Set http=Server.CreateObject("Microsoft.XMLHTTP") 
http.Open "GET",xmlseed,False 
http.send 
Set xmlDoc=Server.CreateObject("Microsoft.XMLDOM") 
xmlDoc.Async=False 
xmlDoc.ValidateOnParse=False 
xmlDoc.Load(http.ResponseXML)
Set item=xmlDoc.getElementsByTagName("item")
if item.Length<=10 then
%>
<script language="javascript">
alert("对不起,该新闻条数已经少于10条新闻条数!");
</script>
<%
else
For i=0 To (item.Length-1)
Set title=item.Item(i).getElementsByTagName("title")
Set link=item.Item(i).getElementsByTagName("link")
Response.Write("<a href="""& link.Item(0).Text &""" target='_blank'>"& title.Item(0).Text &"</a><br>")
Next
end if
End Function
%>
<html>
<head>
<title>远程读取XML文件</title>
</head>
<body>
<%
call readrss("http://www.itlearner.com/article/feed.asp")
%>
<br><br>

</body>
</html>

