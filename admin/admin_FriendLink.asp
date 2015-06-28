<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%

	if session("Admin") = "" then
		call sussLoctionHref("非法操作","admin_login.asp")
	end if
%>
<!--#include file="../include/function.asp"-->
<!--#include file="conn.asp"-->
<%
   dim rs, sql
set rs = server.createobject("adodb.recordset")
sql = "select * from FriendLink "
rs.open sql,conn,1,1


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Back Office Mgmt System--后台管理页面</title>
<link rel="stylesheet" type="text/css" href="style/admin.css" />
</head>
<body>
	
	
	<table id="votename" cellspacing="1">
		<tr><th>连接名称</th><th>连接地址</th><th>连接信息</th><th>操作</th></tr>
  <%
   do while not rs.eof 
  %>
        <tr>
         <td><%=rs("LinkName")%></td><td>
		 <a href="<%=rs("LinkAddress")%>" >点击 </a></td><td><%=rs("LinkInfo")%></td><td><a href="admin_FriendLink_mof.asp?showid=<%=rs("ID")%>">修改</a> | <a onclick="return confirm('您确定进行删除吗？')" href="admin_FriendLink_del.asp?del=ok&showid=<%=rs("ID")%>">删除</a>
         </td>
        </tr>



<%
 rs.movenext
 loop
 
%>


	</table>

	<p style="text-align:center;"><a href="admin_FriendLink_Add.asp">[添加链接]</a></p>
	
</body>
</html>
<%
	call close_rs
	call close_conn
%>