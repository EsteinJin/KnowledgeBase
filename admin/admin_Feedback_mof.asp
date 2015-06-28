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
	if request.Form("send")="提交修改" then
	dim CMS_MonitorDate,CMS_Source,CMS_TicketType,CMS_TicketNumber,CMS_TicketSummary,CMS_AgentName,CMS_TicketStatus,CMS_IsComplaint,CMS_CompliantSource,CMS_ComlaintBy,CMS_RootCause,CMS_HandleStatus,CMS_PAHighLight,CMS_CoachingLog,CMS_InternalUAT,CMS_RealCase,CMS_Point,addsql,RootCause,PAHighLight,CoachingLog,InternalUAT,RealCase
     id = request.form("id")
	CMS_Source=replace(request.form("FeedbackSource"),"'","")
	CMS_OSDNumber=replace(request.form("OSDNumber"),"'","")
	CMS_GAHDNumber=replace(request.form("GAHDNumber"),"'","")
	CMS_RequestID=replace(request.form("RequestNumber"),"'","")
	CMS_IMACNumber=replace(request.form("IMACNumber"),"'","")
	CMS_BTRequestNumber=replace(request.form("BTRequestNumber"),"'","")
	CMS_BTIncidentNumber=replace(request.form("BTIncidentNumber"),"'","")
	
	CMS_TicketSummary=replace(request.Form("IssueSummary"),"'","")
	CMS_AgentName=replace(request.Form("AgentName"),"'","")
	CMS_TicketStatus=replace(request.Form("TicketStatus"),"'","")
	CMS_IsComplaint=replace(request.Form("IsComplaint"),"'","")
	CMS_CompliantSource=replace(request.Form("CompliantSource"),"'","")
	CMS_ComlaintBy=replace(request.Form("ComplaintBy"),"'","")
	CMS_RootCause=replace(request.Form("RootCause"),"'","")
	
	CMS_HandleStatus=replace(request.Form("HandlingStatus"),"'","")
	CMS_PAHighLight=replace(request.Form("PAHighLight"),"'","")
	CMS_CoachingLog=replace(request.Form("CoachingLog"),"'","")
	CMS_InternalUAT=replace(request.Form("InternalUAT"),"'","")
	CMS_RealCase=replace(request.Form("RealCase"),"'","")
	CMS_Point=replace(request.Form("AgentKPI"),"'","")



		'开始执行修改命令
		updatesql = "update CMS_Complaint set CMS_Source='"&CMS_Source&"',CMS_OSDNumber='"&CMS_OSDNumber&"',CMS_GAHDNumber='"&CMS_GAHDNumber&"',CMS_RequestID='"&CMS_RequestID&"',CMS_IMACNumber='"&CMS_IMACNumber&"',CMS_BTRequestNumber='"&CMS_BTRequestNumber&"',CMS_BTIncidentNumber='"&CMS_BTIncidentNumber&"',CMS_TicketSummary='"&CMS_TicketSummary&"',CMS_AgentName='"&CMS_AgentName&"',CMS_TicketStatus='"&CMS_TicketStatus&"',CMS_IsComplaint='"&CMS_IsComplaint&"',CMS_CompliantSource='"&CMS_CompliantSource&"',CMS_ComlaintBy='"&CMS_ComlaintBy&"',CMS_RootCause='"&CMS_RootCause&"',CMS_HandleStatus='"&CMS_HandleStatus&"',CMS_PAHighLight='"&CMS_PAHighLight&"',CMS_CoachingLog='"&CMS_CoachingLog&"',CMS_InternalUAT='"&CMS_InternalUAT&"',CMS_RealCase='"&CMS_RealCase&"',CMS_Point='"&CMS_Point&"' where CMS_ID="&id
		conn.execute(updatesql)
		call sussLoctionHref("内容修改完成","admin_Feedback.asp")
	end if

	showid = request.querystring("ShowId")

	'非法操作
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("非法操作")
	end if
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Complaint where CMS_ID="&showid
	rs.open sql,conn,1,1	
	if rs.eof then
		call close_rs
		call close_conn
		call errorHistoryBack("不存在此数据")
		'以上全部是验证数据，已经成功
	else
	RootCause=rs("CMS_RootCause")
	PAHighLight=rs("CMS_PAHighLight")
	CoachingLog=rs("CMS_CoachingLog")
	InternalUAT=rs("CMS_InternalUAT")
	RealCase=rs("CMS_RealCase")
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
<form name="add" id="articleadd" method="post" action="admin_Feedback_mof.asp">
<input type="hidden" value="<%=showid%>" name="id" />
<dl>
<dt>请发布文章</dt>
<dd>投诉来源:<font style="color:red"><%=rs("CMS_Source")%> </font>
<select  name="FeedbackSource">
<option value="Call">Call</option>
<option value="Ticket">Ticket</option>
<option value="eMail">eMail</option>

</select>

OSD单号:
<input type="text" name="OSDNumber" value="<%=rs("CMS_OSDNumber")%>"  />
GAHD单号:
<input type="text" name="GAHDNumber" value="<%=rs("CMS_GAHDNumber")%>"  />
RequestID:
<input type="text" name="RequestNumber"  value="<%=rs("CMS_RequestID")%>" />
</dd>
<dd>
IMAC ID:
<input type="text" name="IMACNumber"  value="<%=rs("CMS_IMACNumber")%>"   />
BTRequestID:
<input type="text" name="BTRequestNumber"  value="<%=rs("CMS_BTRequestNumber")%>"  />
BTIncidentID:
<input type="text" name="BTIncidentNumber"  value="<%=rs("CMS_BTIncidentNumber")%>"   />
</dd>

<dd>Agent名称：
<input type="text" name="AgentName" value="<%=rs("CMS_AgentName")%>" />

问题状态:<font style="color:red"><%=rs("CMS_TicketStatus")%> </font>
<select  name="TicketStatus">
<option value="Pending">Pending</option>
<option value="Assigned">Assgined</option>
<option value="Resolved">Resolved</option>
<option value="ReOpened">ReOpened</option>
<option value="Closed">Closed</option>

</select>
投诉来源:<font style="color:red"><%=rs("CMS_CompliantSource")%> </font>
<select  name="CompliantSource">
<option value="Global">Global</option>
<option value="EndUser">End User</option>
<option value="Atos">Atos</option>
<option value="LocalIM">Local IM</option>
<option value="QC">QC</option>
</select>


</dd>

<dd>是否投诉：<font style="color:red"><%=rs("CMS_IsComplaint")%> </font>
<input type="radio" name="IsComplaint" value="Yes" checked="checked"/>Yes
<input type="radio" name="IsComplaint" value="No"/>No
投诉人员：
<input type="text" name="ComplaintBy" value="<%=rs("CMS_ComlaintBy")%>"  />
投诉追踪:<font style="color:red"><%=rs("CMS_HandleStatus")%> </font>
<select  name="HandlingStatus">
<option value="Logged">Logged</option>
<option value="Tracking">Tracking</option>
<option value="Investigation">Investigation</option>
<option value="Coaching">Coaching</option>
<option value="RealCase">Real Case</option>
<option value="HighLight">High Light</option>
<option value="HighLight">High Light</option>
<option value="Sign Off">Sign Off</option>
<option value="Fixed">Fixed</option>
</select>

</dd>
<dd>
KPI Point:<font style="color:red"><%=rs("CMS_Point")%> </font>
<input type="radio" name="AgentKPI" value="-3" />-3
<input type="radio" name="AgentKPI" value="-2" />-2
<input type="radio" name="AgentKPI" value="-1" />-1
<input type="radio" name="AgentKPI" value="0" />0
<font style="color:red; font-weight:bold;">请选择扣除的KPI分数</font>
</dd>


<dd>问题描述：
<textarea rows="2" style="width:100%;"  name="IssueSummary"><%=rs("CMS_TicketSummary")%></textarea>
</dd>


<dd>

<%
	Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '创建一个编辑器的实例
	oFCKeditor.BasePath = "fckeditor/" '配置编辑器的路径，我站点根目录下的一个目录
	oFCKeditor.ToolbarSet = "Default" '完整和简化.Basic
	oFCKeditor.Width = "50%" '编辑器的长度
	oFCKeditor.Height = "250" '编辑器的高度
	oFCKeditor.Value = RootCause '这个是给编辑器初始值
	oFCKeditor.Create "RootCause" '以后编辑器里的内容都是由这个content 取得，
%>


<%
	'Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '创建一个编辑器的实例
	oFCKeditor.BasePath = "fckeditor/" '配置编辑器的路径，我站点根目录下的一个目录
	oFCKeditor.ToolbarSet = "Default" '完整和简化.Basic
	oFCKeditor.Width = "50%" '编辑器的长度
	oFCKeditor.Height = "250" '编辑器的高度
	oFCKeditor.Value = PAHighLight '这个是给编辑器初始值
	oFCKeditor.Create "PAHighLight" '以后编辑器里的内容都是由这个content 取得，
%>
</dd>
<dd>

<%
	'Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '创建一个编辑器的实例
	oFCKeditor.BasePath = "fckeditor/" '配置编辑器的路径，我站点根目录下的一个目录
	oFCKeditor.ToolbarSet = "Default" '完整和简化.Basic
	oFCKeditor.Width = "50%" '编辑器的长度
	oFCKeditor.Height = "250" '编辑器的高度
	oFCKeditor.Value = CoachingLog '这个是给编辑器初始值
	oFCKeditor.Create "CoachingLog" '以后编辑器里的内容都是由这个content 取得，
%>

<%
	'Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '创建一个编辑器的实例
	oFCKeditor.BasePath = "fckeditor/" '配置编辑器的路径，我站点根目录下的一个目录
	oFCKeditor.ToolbarSet = "Default" '完整和简化.Basic
	oFCKeditor.Width = "50%" '编辑器的长度
	oFCKeditor.Height = "250" '编辑器的高度
	oFCKeditor.Value = InternalUAT '这个是给编辑器初始值
	oFCKeditor.Create "InternalUAT" '以后编辑器里的内容都是由这个content 取得，
%>

</dd>
<dd>



<%
	'Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '创建一个编辑器的实例
	oFCKeditor.BasePath = "fckeditor/" '配置编辑器的路径，我站点根目录下的一个目录
	oFCKeditor.ToolbarSet = "Default" '完整和简化.Basic
	oFCKeditor.Width = "100%" '编辑器的长度
	oFCKeditor.Height = "250" '编辑器的高度
	oFCKeditor.Value = RealCase '这个是给编辑器初始值
	oFCKeditor.Create "RealCase" '以后编辑器里的内容都是由这个content 取得，
%>

</dd>
<dd><input type="submit" onclick="return check();" name="send" value="提交修改" /></dd>
</dl>

</form>
</body>
</html>
<%
call close_rs
%>
	