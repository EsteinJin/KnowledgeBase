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
	dim CMS_Agent
	
	CMS_Agent=request.Form("CMS_Agent")
    CMS_Title=replace(request.form("CMS_Title"),"'","")
	CMS_PraisedBy=replace(request.form("CMS_PraisedBy"),"'","")
	CMS_Type=replace(request.form("CMS_Type"),"'","")
	CMS_Evidence=replace(request.form("CMS_Evidence"),"'","")
	CMS_QAComment=replace(request.form("CMS_QAComment"),"'","")	
	CMS_Learnd=replace(request.form("CMS_Learnd"),"'","")
	CMS_KPI=replace(request.form("CMS_KPI"),"'","")	
	
	addsql="Insert into CMS_Compliment(CMS_Date,CMS_Agent,CMS_Title,CMS_PraisedBy,CMS_Type,CMS_Evidence,CMS_QAComment,CMS_Learnd,CMS_KPI) values(now(),'"&CMS_Agent&"','"&CMS_Title&"','"&CMS_PraisedBy&"','"&CMS_Type&"','"&CMS_Evidence&"','"&CMS_QAComment&"','"&CMS_Learnd&"','"&CMS_KPI&"')"
	conn.execute(addsql)
	call sussLoctionHref("内容新增成功","admin_compliment.asp")	
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
<form name="add" id="articleadd" method="post" action="admin_compliment_add.asp">
<dl>
<dt>请发布文章</dt>
<dd>Agent名称:

<select  name="CMS_Agent">
<%
set rs = server.createobject("adodb.recordset")
sql="select * from CMS_Agent"
	rs.open sql,conn,1,1
	do while not rs.eof 
%>
<option value="<%=rs("Agent_Name")%>"><%=rs("Agent_Name")%></option>
<%
 rs.movenext
 loop
%>
</select>

表扬自：
<input type="text" name="CMS_PraisedBy" />

表扬来源:
<select  name="CMS_Type">
<option value="Call">Call</option>
<option value="Ticket">Ticket</option>
<option value="Email">Email</option>
<option value="Sametime">Sametime</option>
<option value="Remote">Remote</option>
</select>
</dd>
<dd>
KPI Point:
<input type="radio" name="CMS_KPI" value="3" />3
<input type="radio" name="CMS_KPI" value="2" />2
<input type="radio" name="CMS_KPI" value="1" />1
<input type="radio" name="CMS_KPI" value="0" />0
<font style="color:red; font-weight:bold;">请选择表扬的KPI分数</font>
</select>
</dd>
<dd>

<%
	Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '创建一个编辑器的实例
	oFCKeditor.BasePath = "fckeditor/" '配置编辑器的路径，我站点根目录下的一个目录
	oFCKeditor.ToolbarSet = "Default" '完整和简化.Basic
	oFCKeditor.Width = "50%" '编辑器的长度
	oFCKeditor.Height = "250" '编辑器的高度
	oFCKeditor.Value = "这里描述表扬的内容" '这个是给编辑器初始值
	oFCKeditor.Create "CMS_Title" '以后编辑器里的内容都是由这个content 取得，
%>


<%
	'Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '创建一个编辑器的实例
	oFCKeditor.BasePath = "fckeditor/" '配置编辑器的路径，我站点根目录下的一个目录
	oFCKeditor.ToolbarSet = "Default" '完整和简化.Basic
	oFCKeditor.Width = "50%" '编辑器的长度
	oFCKeditor.Height = "250" '编辑器的高度
	oFCKeditor.Value = "这里记录QA的Comment" '这个是给编辑器初始值
	oFCKeditor.Create "CMS_QAComment" '以后编辑器里的内容都是由这个content 取得，
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
	oFCKeditor.Value = "这里记录相关Raw Data的链接或截图" '这个是给编辑器初始值
	oFCKeditor.Create "CMS_Evidence" '以后编辑器里的内容都是由这个content 取得，
%>

<%
	'Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '创建一个编辑器的实例
	oFCKeditor.BasePath = "fckeditor/" '配置编辑器的路径，我站点根目录下的一个目录
	oFCKeditor.ToolbarSet = "Default" '完整和简化.Basic
	oFCKeditor.Width = "50%" '编辑器的长度
	oFCKeditor.Height = "250" '编辑器的高度
	oFCKeditor.Value = "这里记录Agent反馈的心得内容" '这个是给编辑器初始值
	oFCKeditor.Create "CMS_Learnd" '以后编辑器里的内容都是由这个content 取得，
%>

</dd>

<dd><input type="submit" onclick="return Morecheck();" name="send" value="添加内容" /></dd>
</dl>

</form>
</body>
</html>