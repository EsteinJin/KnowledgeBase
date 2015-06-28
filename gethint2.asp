<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->

<%

dim rs,sql,keyword
keyword=request.QueryString("q")
set rs=server.CreateObject("adodb.recordset")
sql = "select * from CMS_Article where CMS_Title like '%"&keyword&"%' or CMS_Keyword like '%"&keyword&"%' or CMS_Content like '%"&keyword&"%' or CMS_Tag like '%"&keyword&"%'"
rs.open sql,conn,1,1
if not rs.eof then
do while not rs.eof 
%>
<a href="detail.asp?ShowId=<%=rs("CMS_ID")%>" onclick="updateview(<%=rs("CMS_ID")%>)" target="_self"><%=rs("CMS_Title")%></a><br>

<%
rs.movenext
loop

else 
response.Write("No Record Found")

end if 
%>