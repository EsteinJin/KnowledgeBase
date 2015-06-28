<%@codepage =936%>
<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->

<%
	dim showid
	showid = request.querystring("Category")
%>



<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" href="style/basic.css" />
<title>Back Office Mgmt System</title>
</head>
<body>

<!--#include file="header.asp"-->

<div id="myIndexList" style=" overflow:scroll;border:1px solid #0000CC;">

<%
	'非法操作
	set rs = server.createobject("adodb.recordset")
	sql = "select * from ToolsName where ToolsCategory='"&showid&"'"
	rs.open sql,conn,1,1
	
	if rs.eof then
	call errorHistoryBack("Not Exist Category")
	else
%>

<h1><%=rs("ToolsCategory")%> Tools Link List</h1>
<table>
<th>项目名称</th><th>工具名称</th><th>工具链接</th><th>查看工具信息</th>
<%
  do while not rs.eof 
%>
<tr style="border:1px solid maroon">
<td><%=rs("ProjectName")%></td><td><%=rs("ToolsName")%></td><td><a href="<%=rs("ToolsLink")%>">点击进入网站 </a></td>
<td><a href="ToolsLinkDetail.asp?ShowId=<%=rs("Item")%>">点击查看</a></td>
</tr>


<%
   rs.movenext
   loop		
	end if
%>
</table>	

</div>



</body>
</html>
<%
	call close_rs
	call close_conn
%>