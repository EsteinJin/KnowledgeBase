<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%

	if session("Admin") = "" then
		call sussLoctionHref("非法操作","admin_login.asp")
	end if
%>
<!--#include file="../include/function.asp"-->
<!--#include file="conn.asp"-->
<%
	dim rs,sql
	
	if request.querystring("first") = "ok" then
		'第一种，先查找目前首选，然后变成否，再将你点击的变成是
		'第二种，将所有的都变成否，然后再将你点击的变成是
		
		'这两句将所有的改成0,sql是可以批处理的，没有做任何的where筛选，就是所有
		sql = "Update CMS_Vote Set CMS_Level=0"
		conn.execute(sql)
		
		'再将你要选择的那个标题改成首选
		sql = "Update CMS_Vote Set CMS_Level=1 where CMS_ID="&request.querystring("ShowId")
		conn.execute(sql)
		
	end if
	
	
	
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Vote where CMS_VoteSid=0 order by CMS_Date desc"
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
		<tr><th>编号</th><th>标题</th><th>是否首页显示</th><th>操作</th></tr>
		<%
			do while not rs.eof
				dim levelstr
				if rs("CMS_Level") = 0 then
					levelstr = "否"
				elseif rs("CMS_Level") = 1 then
					levelstr = "<span style='color:red;font-weight:bold;'>是</span>"
				end if
		%>
		<tr><td class="id"><%=rs("CMS_ID")%></td><td><a href="admin_vote_x.asp?ShowId=<%=rs("CMS_ID")%>"><%=rs("CMS_VoteName")%></a></td><td><%=levelstr%></td><td><a href="admin_vote_name.asp?first=ok&ShowId=<%=rs("CMS_ID")%>">确定首选</a> | <a href="admin_vote_mof.asp?ShowId=<%=rs("CMS_ID")%>">修改</a> | <a onclick="return confirm('您确定进行删除吗？')" href="admin_vote_del.asp?del=ok&ShowId=<%=rs("CMS_ID")%>">删除</a></td></tr>
        
        
		<%
				rs.movenext
			loop
		%>
	</table>

	<p style="text-align:center;"><a href="admin_vote_name_add.asp">[添加名称]</a></p>
	
</body>
</html>
<%
	call close_rs
	call close_conn
%>