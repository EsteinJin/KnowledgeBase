<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%

	if session("Admin") = "" then
		call sussLoctionHref("error occured!","admin_login.asp")
	end if
%>
<!--#include file="../include/function.asp"-->
<!--#include file="conn.asp"-->
<%
   dim rs, sql
set rs = server.createobject("adodb.recordset")
sql = "select * from ToolsName order by Item desc"
rs.open sql,conn,1,1


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>KB - BACK END </title>
<link rel="stylesheet" type="text/css" href="style/admin.css" />
</head>
<body>
	
	
	<table id="votename" cellspacing="1">
		<tr><th>ITEM</th><th>TEAM</th><th>工具名称</th><th>工具链接</th><th>操作</th></tr>
  <%
   do while not rs.eof 
  %>
        <tr>
         <td><%=rs("ProjectName")%></td><td><%=rs("ToolsCategory")%></td><td><%=rs("ToolsName")%></td><td>
		 <a href="<%=rs("ToolsLink")%>" ><%=rs("ToolsName")%> </a></td><td><a href="admin_Tools_mof.asp?ShowId=<%=rs("Item")%>">修改</a> | <a onclick="return confirm('您确定进行删除吗？')" href="admin_Tools_del.asp?del=ok&ShowId=<%=rs("Item")%>">删除</a>
         </td>
        </tr>
<%
 rs.movenext
 loop
 
%>


	</table>

	<p style="text-align:center;"><a href="admin_Tools_List_Add.asp">[添加名称]</a></p>
	
</body>
</html>
<%
	call close_rs
	call close_conn
%>