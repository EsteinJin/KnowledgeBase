<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../include/function.asp"-->
<!--#include file="conn.asp"-->
<%

	if session("Admin") = "" then
		call sussLoctionHref("非法操作","admin_login.asp")
	end if
	
		dim rs,sql
	set rs = server.createobject("adodb.recordset")
	sql = "SELECT * FROM EscalationLog order by EscalatedDate desc"
	rs.open sql,conn,1,1
	rs.pagesize=2
	if isnumeric(request.querystring("page")) then
		if request.querystring("page") = "" or cint(request.querystring("page"))<1 then
			rs.absolutepage = 1
		elseif cint(request.querystring("page"))>rs.pagecount then
			rs.absolutepage = rs.pagecount
		else
			rs.absolutepage = request.querystring("page")
		end if
	else
		rs.absolutepage = 1
	end if 
	
	
	'删除模块
	if request.querystring("del")="ok" then
		dim showid,delrs,delsql,delsql2
		
		showid = request.querystring("ShowId")
		'非法操作
		if showid = "" or not isnumeric(showid) then
			call errorHistoryBack("非法操作")
		end if
		
		'判断数据是否存在
		set delrs = server.createobject("adodb.recordset")
		delsql = "select * from EscalationLog where ID="&showid
		delrs.open delsql,conn,1,1
		
		'如果数据不存在
		if delrs.eof then
			call close_rs
			call close_conn
			delrs.close
			set delrs = nothing
			call errorHistoryBack("你要删除的数据不存在")
		else
		    call ConfirmDel()
			'执行删除命令
			delsql2 = "delete from EscalationLog where ID="&showid
			conn.execute(delsql2)
			call sussLoctionHref("删除成功","admin_Escalation.asp")
		end if
		
		delrs.close
		set delrs = nothing
		
	end if

%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" href="style/admin.css" />
<title>无标题文档</title>
</head>

<body>
<p style=" font-size:12px; text-align:center; margin-top:10px;"><a href="admin_Escalation_add.asp">点击这里</a>添加内容</p>     
<table border="1" id="listcontent">
<tr><th>升级人员</th><th>问题类型</th><th>处理状况</th><th>负责人员</th><th>问题描述</th><th>操作</th></tr>
<%
for i=1 to rs.pagesize
if rs.eof then exit for	
Summary = rs("IssueSummary")
if len(Summary) > 35 then
Summary = left(Summary,35)
Summary = Summary & "..."
end if
%>
<tr><td><%=rs("EscalatedBy")%></td><td><%=rs("EscalationType")%></td><td><%=rs("StatusTrack")%></td><td><%=rs("ResponsibleBy")%></td><td><%=Summary%></td><td class="d"><a href="admin_Escalation_mof.asp?ShowId=<%=rs("ID")%>">修改</a> | <a  onclick="return confirm('您确定进行删除吗？')" href="admin_Escalation.asp?del=ok&ShowId=<%=rs("ID")%>">删除</a></td></tr>



<%
rs.movenext
next	
%> 

</table>
<p style="text-align:center;padding:10px;">
<%
	for i = 1 to rs.pagecount
		response.write "<a href='admin_Escalation.asp?page="&i&"'>" & i & "</a> | "
	next
%>
   
     </p>

     
</body>
</html>
<%
	call close_rs
	call close_conn
%>
