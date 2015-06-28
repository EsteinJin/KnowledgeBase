<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../include/function.asp"-->
<!--#include file="conn.asp"-->
<%

	if session("Admin") = "" then
		call sussLoctionHref("非法操作","admin_login.asp")
	end if

	dim showid
	showid = request.querystring("ShowId")
	'判断showid有效
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("非法操作")
	end if
	
	'判断showid这个栏目是否存在
	dim rs,sql,votename
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Vote where CMS_ID="&showid
	rs.open sql,conn,1,1
	
	if rs.eof then
		call close_rs
		call close_conn
		call errorHistoryBack("不存在此标题")
	else
		'有数据
		votename = rs("CMS_VoteName")
	end if
	
	call close_rs
	
	'开始增加项目
	if request.form("send") = "增加" then  '判断按钮是否点过
		dim votex,xrs,xsql
		votex = request.form("votex")
		
		if len(votex) < 2 then
			call errorHistoryBack("项目名不得少于2位！")
		end if
		
		
		set xrs = server.createobject("adodb.recordset")
		xsql = "select * from CMS_Vote where CMS_VoteSid="&showid
		xrs.open xsql,conn,1,1
		'项目新增,要做个判断，如果有4个项目名了，就不能新新增了
		
		'还有一个重名的判断
		
		'recordcount
		if xrs.recordcount >=10 then
			call errorHistoryBack("项目的数量已经封顶")
		else
			sql = "Insert into CMS_Vote (CMS_VoteName,CMS_VoteSid,CMS_Date) values ('"&votex&"',"&showid&",now())"
			conn.execute(sql)
		end if
		
		xrs.close
		set xrs = nothing
		
		'因为新增是在提取数据之后完成的，所以，要跳转刷新一下，才能得到数据
		response.redirect "admin_vote_x.asp?ShowId="&showid
		
	end if
	
	
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Vote where CMS_VoteSid="&showid
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
		<tr><td colspan="4" class="title"><%=votename%></td></tr>
		<tr><th>编号</th><th>项目名称</th><th>得票数</th><th>操作</th></tr>
		<%
			do while not rs.eof
		%>
		<tr><td><%=rs("CMS_ID")%></td><td><%=rs("CMS_VoteName")%></td><td><%=rs("CMS_VoteCount")%></td><td><a onclick="return confirm('您确定进行删除吗？')" href="admin_vote_x_del.asp?del=ok&ShowId=<%=rs("CMS_ID")%>">删除</a> <a href="admin_vote_x_mof.asp?ShowId=<%=rs("CMS_ID")%>">修改</a></td></tr>
 
		<%
				rs.movenext
			loop
		%>
	</table>
	<form method="post" action="admin_vote_x.asp?ShowId=<%=showid%>">
		<dl style="width:250px;margin:auto;">
			<dt>请增加项目：</dt>
			<dd>项目名：<input type="text" name="votex" /> <input type="submit" value="增加" name="send" /></dd>
		</dl>
	</form>

</body>
</html>
<%
		call close_rs
		call close_conn
%>