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
	if request.Form("send")="修改内容" then
	id = request.form("id")
	CMS_Agent=request.Form("CMS_Agent")
    CMS_Title=replace(request.form("CMS_Title"),"'","")
	CMS_PraisedBy=replace(request.form("CMS_PraisedBy"),"'","")
	CMS_Type=replace(request.form("CMS_Type"),"'","")
	CMS_Evidence=replace(request.form("CMS_Evidence"),"'","")
	CMS_QAComment=replace(request.form("CMS_QAComment"),"'","")	
	CMS_Learnd=replace(request.form("CMS_Learnd"),"'","")
	CMS_KPI=replace(request.form("CMS_KPI"),"'","")	
	
	updatesql="update CMS_Compliment set CMS_Agent='"&CMS_Agent&"',CMS_Title='"&CMS_Title&"',CMS_PraisedBy='"&CMS_PraisedBy&"',CMS_Type='"&CMS_Type&"',CMS_Evidence='"&CMS_Evidence&"',CMS_QAComment='"&CMS_QAComment&"',CMS_Learnd='"&CMS_Learnd&"',CMS_KPI='"&CMS_KPI&"' where CMS_ID="&id
	conn.execute(updatesql)
	call sussLoctionHref("内容修改成功","admin_compliment.asp")	
	end if 
	
	showid = request.querystring("ShowId")

	'非法操作
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("非法操作")
	end if
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Compliment where CMS_ID="&showid
	rs.open sql,conn,1,1	
	if rs.eof then
		call close_rs
		call close_conn
		call errorHistoryBack("不存在此数据")
		'以上全部是验证数据，已经成功
	else
	CMS_Agent=rs("CMS_Agent")
    CMS_Title= rs("CMS_Title")
	CMS_PraisedBy= rs("CMS_PraisedBy")
	CMS_Type= rs("CMS_Type")
	CMS_Evidence= rs("CMS_Evidence")
	CMS_QAComment= rs("CMS_QAComment")
	CMS_Learnd= rs("CMS_Learnd")
	CMS_KPI= 	rs("CMS_KPI")

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
<form name="add" id="articleadd" method="post" action="admin_compliment_mof.asp">
<input type="hidden" value="<%=showid%>" name="id" />
<dl>
<dt>请发布文章</dt>
<dd>Agent名称:

<select  name="CMS_Agent">
<%
set rs2 = server.createobject("adodb.recordset")
sql2="select * from CMS_Agent"
	rs2.open sql2,conn,1,1
	do while not rs2.eof 
%>
<option value="<%=rs2("Agent_Name")%>"><%=rs2("Agent_Name")%></option>

<%
 rs2.movenext
 loop
%>

</select>
<font style="color:red; font-weight:bold;">当前为:<%=CMS_Agent%></font>
表扬自：
<input type="text" name="CMS_PraisedBy" value="<%=CMS_PraisedBy%>" />
</dd>
<dd>
表扬来源:
<select  name="CMS_Type">
<option value="Call">Call</option>
<option value="Ticket">Ticket</option>
<option value="Email">Email</option>
<option value="Sametime">Sametime</option>
<option value="Remote">Remote</option>
</select>
<font style="color:red; font-weight:bold;">当前为:<%=CMS_Type%></font>
KPI Point:
<input type="radio" name="CMS_KPI" value="3" />3
<input type="radio" name="CMS_KPI" value="2" />2
<input type="radio" name="CMS_KPI" value="1" />1
<input type="radio" name="CMS_KPI" value="0" />0
<font style="color:red; font-weight:bold;">请选择表扬的KPI分数</font>
</select>
<font style="color:red; font-weight:bold;">当前为:<%=CMS_KPI%></font>
</dd>
<dd>

<%
	Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '创建一个编辑器的实例
	oFCKeditor.BasePath = "fckeditor/" '配置编辑器的路径，我站点根目录下的一个目录
	oFCKeditor.ToolbarSet = "Default" '完整和简化.Basic
	oFCKeditor.Width = "50%" '编辑器的长度
	oFCKeditor.Height = "250" '编辑器的高度
	oFCKeditor.Value = CMS_Title '这个是给编辑器初始值
	oFCKeditor.Create "CMS_Title" '以后编辑器里的内容都是由这个content 取得，
%>


<%
	'Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '创建一个编辑器的实例
	oFCKeditor.BasePath = "fckeditor/" '配置编辑器的路径，我站点根目录下的一个目录
	oFCKeditor.ToolbarSet = "Default" '完整和简化.Basic
	oFCKeditor.Width = "50%" '编辑器的长度
	oFCKeditor.Height = "250" '编辑器的高度
	oFCKeditor.Value = CMS_QAComment '这个是给编辑器初始值
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
	oFCKeditor.Value = CMS_Evidence '这个是给编辑器初始值
	oFCKeditor.Create "CMS_Evidence" '以后编辑器里的内容都是由这个content 取得，
%>

<%
	'Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '创建一个编辑器的实例
	oFCKeditor.BasePath = "fckeditor/" '配置编辑器的路径，我站点根目录下的一个目录
	oFCKeditor.ToolbarSet = "Default" '完整和简化.Basic
	oFCKeditor.Width = "50%" '编辑器的长度
	oFCKeditor.Height = "250" '编辑器的高度
	oFCKeditor.Value = CMS_Learnd '这个是给编辑器初始值
	oFCKeditor.Create "CMS_Learnd" '以后编辑器里的内容都是由这个content 取得，
%>

</dd>

<dd><input type="submit" onclick="return Morecheck();" name="send" value="修改内容" /></dd>
</dl>

</form>
</body>
</html>