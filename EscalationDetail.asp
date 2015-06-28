<%@codepage =936%>
<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->
<%
	dim showid

	showid = request.querystring("ShowId")
	'非法操作
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("非法操作")
	end if
	
	dim title,content
	set rs = server.createobject("adodb.recordset")
	sql = "select * from EscalationLog where ID="&showid
	rs.open sql,conn,1,1
	
	if rs.eof then
		call errorHistoryBack("不存在此内容")
	else 
	EscalatedDate=rs("EscalatedDate")
	EscalatedBy=rs("EscalatedBy")
	EscalationType=rs("EscalationType")
	TicketNumber=rs("TicketNumber")
	Category=rs("Category")
	IssueSummary=rs("IssueSummary")
	IssueDetails=rs("IssueDetails")
	ResponsibleBy=rs("ResponsibleBy")
	EscalationLog=rs("EscalationLog")
	StatusTrack=rs("StatusTrack")	
	
	end if
	
	call close_rs
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />

<!--上面表示采用简体中文显示-->
<!--

<!--#include file="fckeditor/fckeditor.asp"-->
<!--->
<title>Back Office Mgmt System</title>
<link rel="stylesheet" type="text/css" href="style/basic.css" />
</head>
<style type="text/css">
.contents{border:1px dashed #999900; margin-top:20px;}
</style>
<script type="text/javascript">
function ExpandThis()
{
 document.getElementById('detail').style.height="100%"
}

function CollapsThis()
{
document.getElementById('detail').style.height="300px"
}

</script>

<body>
<span style="position:absolute; right:120px; top:170px;">&nbsp;&nbsp;<a href="javascript:ExpandThis();void(0);">展开</a>&nbsp;&nbsp;<a href="javascript:CollapsThis();void(0);">收缩</a></span>
<!--#include file="header.asp"-->


<div id="detail">
	<h3><%=IssueSummary%></h3>
	<p class="d">升级日期：<span style="color:red;"><%=EscalatedDate%></span>| 升级人员：<span style="color:red;"><%=EscalatedBy%></span>| 问题类型：<span style="color:red;"><%=EscalationType%></span>| 问题单号：<span style="color:red;"><%=TicketNumber%></span>| 
    类别：<span style="color:red;"><%=Category%></span>|
    
     负责人员：<span style="color:red;"><%=ResponsibleBy%></span>|状态跟踪：<span style="color:red;"><%=StatusTrack%></span>|<a href="admin/admin_Escalation_Front_mof.asp?ShowId=<%=showid%>">修改</a> </p>
    
<p class="contents"><%=IssueDetails%></p>
<p class="contents"><%=EscalationLog%></p>

</div>
<div id="MyComment">
  <h1>评论内容 </h1>
  <%
	dim IpAddressInfo,CommentTime,CommentContent,NewsId,rs2,sql2
	set rs2 = server.createobject("adodb.recordset")
	sql2 = "select * from MyComment where NewsId="&showid
	rs2.open sql2,conn,1,1
	do while not rs2.eof 
%>
  <p class="d">评论者：N/A |   评论者IP地址： <%=rs2("IpAddressInfo")%>|   评论时间：<%=rs2("CommentTime")%> </p>
  <p><%=rs2("CommentContent")%>&nbsp;</p>
  <hr style="height:1px;border:none;border-top:1px dashed #0066CC;" />
  <%
rs2.movenext
loop
rs2.close
set rs2 = nothing
%>
</div>
<div id="CommentInput">
  <form action="myCommentEscalation.asp" method="post">
    <%
				Dim oFCKeditor
				Set oFCKeditor = New FCKeditor '创建一个编辑器的实例
				oFCKeditor.BasePath = "fckeditor/" '配置编辑器的路径，我站点根目录下的一个目录
				oFCKeditor.ToolbarSet = "Basic" '完整和简化.Basic
				oFCKeditor.Width = "100%" '编辑器的长度
				oFCKeditor.Height = "400" '编辑器的高度
				oFCKeditor.Value = "" '这个是给编辑器初始值
				oFCKeditor.Create "content" '以后编辑器里的内容都是由这个content 取得，
			%>
    <label for="yzm">验证码：
    <input type="text" name="yzm" id="yzm" class="text yzm" />
    <img src="../include/code.asp" onclick="javascript:this.src='../include/code.asp?tm='+Math.random()" style="cursor:pointer" alt="验证码" /></label>
    <input type="hidden" name="newsId" value="<%=showid%>" />
    <input type="submit" value="添加评论" name="send" class="submit" />
  </form>
</div>
</body>
</html>