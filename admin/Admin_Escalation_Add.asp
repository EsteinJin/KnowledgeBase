<%@codepage = 936%>
<!--上面表示采用简体中文显示-->
<!
<!--#include file="../include/function.asp"-->
<!--#include file="fckeditor/fckeditor.asp"-->
<!--#include file="conn.asp"-->
<%
if session("Admin") = "" then
		call sussLoctionHref("非法操作","admin_login.asp")
	end if
	if request.Form("send")="添加内容" then
	dim EscalationType,TicketNumber,Category,IssueSummary,IssueDetails,ResponsibleBy,HandleStatus,EscalationLog,addsql
	EscalationType=replace(request.form("EscalationType"),"'","")
	TicketNumber=replace(request.form("TicketNumber"),"'","")
	Category=replace(request.form("Category"),"'","")
	IssueSummary=replace(request.form("IssueSummary"),"'","")
	IssueDetails=replace(request.form("IssueDetails"),"'","")
	ResponsibleBy=replace(request.form("ResponsibleBy"),"'","")
	StatusTrack=replace(request.form("HandleStatus"),"'","")
	EscalationLog=replace(request.form("EscalationLog"),"'","")
	EscalatedBy=replace(request.form("EscalatedBy"),"'","")
addsql="Insert into EscalationLog(EscalatedDate,EscalationType,TicketNumber,Category,IssueSummary,IssueDetails,ResponsibleBy,EscalationLog,StatusTrack,EscalatedBy) values (now(),'"&EscalationType&"','"&TicketNumber&"','"&Category&"','"&IssueSummary&"','"&IssueDetails&"','"&ResponsibleBy&"','"&EscalationLog&"','"&StatusTrack&"','"&EscalatedBy&"')"
	conn.execute(addsql)
	call sussLoctionHref("内容新增成功","/Escalation.asp")	
	end if 
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理</title>
<link rel="stylesheet" type="text/css" href="style/admin.css" />
<script type="text/javascript" src="js/content.js"></script>
</head>
<body>
<form name="add" id="articleadd" method="post" action="admin_Escalation_add.asp">
<dl>
<dt>请发布文章</dt>

<dd>
提交人员:
<input type="text" name="EscalatedBy"  />
问题类型：
<input type="radio" name="EscalationType" value="ToolsIssue" checked="checked"/>工具问题
<input type="radio" name="EscalationType" value="ProcessIssue"/>流程问题
&nbsp;&nbsp;&nbsp;&nbsp;单号编号:
<input type="text" name="TicketNumber"  />
类型分类:
<input type="text" name="Category"  />
</dd>
<dd>
<dd>
负责人员:
<input type="radio" name="ResponsibleBy" value="LiuYang" checked="checked"/>LiuYang
<input type="radio" name="ResponsibleBy" value="JiangZhiMin"/>JiangZhiMin
<input type="radio" name="ResponsibleBy" value="ChenQiang"/>ChenQiang

进展情况:
<select  name="HandleStatus">
<option value="Logged">Logged</option>
<option value="Pending">Pending</option>
<option value="Assigned">Assigned</option>
<option value="Resolved">Resolved</option>
<option value="Closed">Closed</option>
</select>
</dd>

<dd>
问题描述:
<textarea rows="2"  style="width:100%;"  name="IssueSummary"></textarea>
</dd>
<dd>
<%
	Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '创建一个编辑器的实例
	oFCKeditor.BasePath = "fckeditor/" '配置编辑器的路径，我站点根目录下的一个目录
	oFCKeditor.ToolbarSet = "Default" '完整和简化.Basic
	oFCKeditor.Width = "100%" '编辑器的长度
	oFCKeditor.Height = "250" '编辑器的高度
	oFCKeditor.Value = "这里填写Issue Details" '这个是给编辑器初始值
	oFCKeditor.Create "IssueDetails" '以后编辑器里的内容都是由这个content 取得，
%>
</dd>

<dd>
<%
	
	Set oFCKeditor = New FCKeditor '创建一个编辑器的实例
	oFCKeditor.BasePath = "fckeditor/" '配置编辑器的路径，我站点根目录下的一个目录
	oFCKeditor.ToolbarSet = "Default" '完整和简化.Basic
	oFCKeditor.Width = "100%" '编辑器的长度
	oFCKeditor.Height = "250" '编辑器的高度
	oFCKeditor.Value = "这里填写EscalationLog" '这个是给编辑器初始值
	oFCKeditor.Create "EscalationLog" '以后编辑器里的内容都是由这个content 取得，
%>
</dd>

</dl>
<dd><input type="submit" onclick="return Morecheck();" name="send" value="添加内容" /></dd>
</form>

</body>
</html>
