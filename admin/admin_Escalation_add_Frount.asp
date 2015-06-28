<%@codepage = 936%>
<!--上面表示采用简体中文显示-->
<!
<!--#include file="../include/function.asp"-->
<!--#include file="fckeditor/fckeditor.asp"-->
<!--#include file="conn.asp"-->
<%

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
	
		if EscalatedBy = "" then
			call errorHistoryBack("提交人不得为空")
		end if	
	if len(IssueSummary) < 4 or len(IssueSummary) > 100 then
			call errorHistoryBack("标题不小于4位，或者大于100位")
		end if
	
addsql="Insert into EscalationLog(EscalatedDate,EscalationType,TicketNumber,Category,IssueSummary,IssueDetails,ResponsibleBy,EscalationLog,StatusTrack,EscalatedBy) values (now(),'"&EscalationType&"','"&TicketNumber&"','"&Category&"','"&IssueSummary&"','"&IssueDetails&"','"&ResponsibleBy&"','"&EscalationLog&"','"&StatusTrack&"','"&EscalatedBy&"')"
application.lock()
conn.execute(addsql)
set newrs=conn.execute("SELECT TOP 1 ID FROM EscalationLog ORDER BY ID DESC")
dim NewID
NewID=newrs("ID")
application.unlock() 
set msg = Server.CreateOBject( "JMail.Message" )
msg.Logging = true
msg.Charset = "utf-8"
msg.ContentTransferEncoding = "base64"
msg.ContentType = "text/html"  
msg.From = "RGCNSISGOEUSBASFBackOffice@internal.siemens.com"
msg.FromName = "No-Reply-IssueID:"&NewID
set rs = server.createobject("adodb.recordset")
sql="select * from CMS_Agent"
rs.open sql,conn,1,1
do while not rs.eof 
msg.AddRecipient rs("Agent_MailAddress"),rs("Agent_Name")
rs.movenext
loop

msg.Subject = "Issue Summary:"&IssueSummary
msg.Body = "提交人员:"&EscalatedBy&"<br />发生时间："&now()&"<br />问题分类:"&EscalationType&"<br />Ticket号码："&TicketNumber&"<br />问题分类："&Category&"<br />负责人："&ResponsibleBy
msg.appendText "<br /><br /><br />"
msg.appendText "<br />"&replace(IssueDetails,"/upFile/","http://"&Request.ServerVariables("server_name")&"/upFile/")
msg.appendText "<br />"&EscalationLog
msg.Send( "apac.internal.siemens-it-solutions.com" )
	call sussLoctionHref("内容新增成功","/EscalationDetail.asp?ShowId="&NewID)	
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
<form name="add" id="articleadd" method="post" action="admin_Escalation_add_Frount.asp">
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
<select  name="HandleStatus" disabled="disabled">
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
	oFCKeditor.Value = "" '这个是给编辑器初始值
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
oFCKeditor.Value = "提示：<br>IVR问题，第一时间升级ICM并录入Genesis Ticket<br>其他问题按流程升级至ICM,同时要向Local IT升级<br>" '这个是给编辑器初始值
	oFCKeditor.Create "EscalationLog" '以后编辑器里的内容都是由这个content 取得，
%>
</dd>

</dl>
<dd><input type="submit" onclick="return Morecheck();" name="send" value="添加内容" /></dd>
</form>

</body>
</html>
