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
	sql = "select * from CMS_Complaint where CMS_ID="&showid
	rs.open sql,conn,1,1
	
	if rs.eof then
		call errorHistoryBack("不存在此内容")
	else 
	CMS_MonitorDate=rs("CMS_MonitorDate")
	CMS_Source=rs("CMS_Source")
	CMS_OSDNumber=rs("CMS_OSDNumber")
	CMS_GAHDNumber=rs("CMS_GAHDNumber")
	CMS_RequestID=rs("CMS_RequestID")
	CMS_IMACNumber=rs("CMS_IMACNumber")
	CMS_BTRequestNumber=rs("CMS_BTRequestNumber")
	CMS_BTIncidentNumber=rs("CMS_BTIncidentNumber")
	
	CMS_AgentName=rs("CMS_AgentName")
	CMS_TicketStatus=rs("CMS_TicketStatus")
	CMS_CompliantSource=rs("CMS_CompliantSource")
	CMS_IsComplaint=rs("CMS_IsComplaint")
	CMS_ComlaintBy=rs("CMS_ComlaintBy")
	CMS_HandleStatus=rs("CMS_HandleStatus")
	CMS_Point=rs("CMS_Point")
	
	CMS_TicketSummary=rs("CMS_TicketSummary")
	CMS_RootCause=rs("CMS_RootCause")
	CMS_PAHighLight=rs("CMS_PAHighLight")
	CMS_CoachingLog=rs("CMS_CoachingLog")
	CMS_InternalUAT=rs("CMS_InternalUAT")
	CMS_RealCase=rs("CMS_RealCase")
		
	
		 
	end if
	
	call close_rs
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
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
<body>

<!--#include file="header.asp"-->


<div id="detail">
<span style=" position:absolute; right:120px; top:170px;">&nbsp;&nbsp;<a href="javascript:ExpandThis();void(0);">展开</a>&nbsp;&nbsp;<a href="javascript:CollapsThis();void(0);">收缩</a></span>
	<h3><%=CMS_TicketSummary%></h3>
	<p class="d">监控来源：<span style="color:red;"><%=CMS_Source%></span>| OSD单号：<span style="color:red;"><%=CMS_OSDNumber%></span>| GAHD单号：<span style="color:red;"><%=CMS_GAHDNumber%></span>| Request单号：<span style="color:red;"><%=CMS_RequestID%></span>| IMAC单号：<span style="color:red;"><%=CMS_IMACNumber%></span>| BTRequest单号：<span style="color:red;"><%=CMS_BTRequestNumber%></span>|BTIncident单号：<span style="color:red;"><%=CMS_BTIncidentNumber%></span>
    <p class="d">
    
     情况跟踪：<span style="color:red;"><%=CMS_HandleStatus%></span>|是否投诉：<span style="color:red;"><%=CMS_IsComplaint%></span>|Feedback来自：<span style="color:red;"><%=CMS_CompliantSource%></span>|投诉人员：<span style="color:red;"><%=CMS_ComlaintBy%></span>|  Agent名称：<span style="color:red;"><%=CMS_AgentName%></span>|单号状态：<span style="color:red;"><%=CMS_TicketStatus%></span>|记录时间：<%=CMS_MonitorDate%> <a href="admin/admin_Feedback_Front_mof.asp?ShowId=<%=showid%>">修改</a> </p>
    
<p class="contents"><%=CMS_RootCause%></p>
<p class="contents"><%=CMS_PAHighLight%></p>
<p class="contents"><%=CMS_CoachingLog%></p>
<p class="contents"><%=CMS_InternalUAT%></p>
<p class="contents"><%=CMS_RealCase%></p>

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
  <form action="myCommentFeedback.asp" method="post">
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