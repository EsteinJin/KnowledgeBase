<%@codepage =936%>
<%
	dim showid

	showid = request.querystring("ShowId")
	'非法操作
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("非法操作")
	end if
	
	dim title,content
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Article where CMS_ID="&showid
	rs.open sql,conn,1,1
	
	if rs.eof then
		call errorHistoryBack("不存在此内容")
	else 
		title = rs("CMS_Title")
		content = rs("CMS_Content")
		info = rs("CMS_Info")
		tag = rs("CMS_Tag")
		keyword = rs("CMS_Keyword")
		name = rs("CMS_Name")
		fdate = rs("CMS_Date")
	end if
	
	call close_rs
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />

<!--上面表示采用简体中文显示-->
<!--
<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="fckeditor/fckeditor.asp"-->
-->
<title>Back Office Mgmt System</title>
<link rel="stylesheet" type="text/css" href="style/basic.css" />
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
</head>
<body>

<!--#include file="header.asp"-->
<span style=" position:absolute; right:10px; top:10px;">&nbsp;&nbsp;<a href="javascript:ExpandThis();void(0);">展开</a>&nbsp;&nbsp;<a href="javascript:CollapsThis();void(0);">收缩</a></span>

<div id="detail">

	<h3><%=title%> </h3>
	<p class="d">TAG标签：<%=tag%> | 搜索关键字：<%=keyword%> | 发布者：<%=name%> | 发布时间：<%=FormatDateTime(fdate,2)%><a href="admin/admin_article_Front_mof.asp?ShowId=<%=showid%>">修改</a> </p>
	<!--<p class="info"><%=info%></p>-->
	<%=content%>	
</div>
<div id="MyComment">
<h1>评论内容 </h1>

<p class="d">评论者： |   评论者IP地址： |   评论时间：  </p>
<div><strong>New Employee</strong></div>
<div><em>Contractor/Temporary</em></div>
<div>BGD ID to be created by SR, request   should come from BASF Internal Employee</div>
<div>All other request should be submitted   via AccessIT</div>
<p>&nbsp;</p>
<hr style="height:1px;border:none;border-top:1px dashed #0066CC;" />

<p class="d">评论者： |   评论者IP地址： |   评论时间：  </p>
<div><strong>New Employee</strong></div>
<div><em>Contractor/Temporary</em></div>
<div>BGD ID to be created by SR, request   should come from BASF Internal Employee</div>
<div>All other request should be submitted   via AccessIT</div>
<p>&nbsp;</p>

</div>
<div id="CommentInput">
<form action="detail.asp" method="post">
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


</form>
</div>

	
</body>
</html>