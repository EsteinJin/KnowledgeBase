<%@codepage =936%>
<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->
<%
	dim showid

	showid = request.querystring("ShowId")
	'�Ƿ�����
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("�Ƿ�����")
	end if
	
	dim title,content
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Complaint where CMS_ID="&showid
	rs.open sql,conn,1,1
	
	if rs.eof then
		call errorHistoryBack("�����ڴ�����")
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

<!--�����ʾ���ü���������ʾ-->
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
<span style=" position:absolute; right:120px; top:170px;">&nbsp;&nbsp;<a href="javascript:ExpandThis();void(0);">չ��</a>&nbsp;&nbsp;<a href="javascript:CollapsThis();void(0);">����</a></span>
	<h3><%=CMS_TicketSummary%></h3>
	<p class="d">�����Դ��<span style="color:red;"><%=CMS_Source%></span>| OSD���ţ�<span style="color:red;"><%=CMS_OSDNumber%></span>| GAHD���ţ�<span style="color:red;"><%=CMS_GAHDNumber%></span>| Request���ţ�<span style="color:red;"><%=CMS_RequestID%></span>| IMAC���ţ�<span style="color:red;"><%=CMS_IMACNumber%></span>| BTRequest���ţ�<span style="color:red;"><%=CMS_BTRequestNumber%></span>|BTIncident���ţ�<span style="color:red;"><%=CMS_BTIncidentNumber%></span>
    <p class="d">
    
     ������٣�<span style="color:red;"><%=CMS_HandleStatus%></span>|�Ƿ�Ͷ�ߣ�<span style="color:red;"><%=CMS_IsComplaint%></span>|Feedback���ԣ�<span style="color:red;"><%=CMS_CompliantSource%></span>|Ͷ����Ա��<span style="color:red;"><%=CMS_ComlaintBy%></span>|  Agent���ƣ�<span style="color:red;"><%=CMS_AgentName%></span>|����״̬��<span style="color:red;"><%=CMS_TicketStatus%></span>|��¼ʱ�䣺<%=CMS_MonitorDate%> <a href="admin/admin_Feedback_Front_mof.asp?ShowId=<%=showid%>">�޸�</a> </p>
    
<p class="contents"><%=CMS_RootCause%></p>
<p class="contents"><%=CMS_PAHighLight%></p>
<p class="contents"><%=CMS_CoachingLog%></p>
<p class="contents"><%=CMS_InternalUAT%></p>
<p class="contents"><%=CMS_RealCase%></p>

</div>
<div id="MyComment">
  <h1>�������� </h1>
  <%
	dim IpAddressInfo,CommentTime,CommentContent,NewsId,rs2,sql2
	set rs2 = server.createobject("adodb.recordset")
	sql2 = "select * from MyComment where NewsId="&showid
	rs2.open sql2,conn,1,1
	do while not rs2.eof 
%>
  <p class="d">�����ߣ�N/A |   ������IP��ַ�� <%=rs2("IpAddressInfo")%>|   ����ʱ�䣺<%=rs2("CommentTime")%> </p>
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
				Set oFCKeditor = New FCKeditor '����һ���༭����ʵ��
				oFCKeditor.BasePath = "fckeditor/" '���ñ༭����·������վ���Ŀ¼�µ�һ��Ŀ¼
				oFCKeditor.ToolbarSet = "Basic" '�����ͼ�.Basic
				oFCKeditor.Width = "100%" '�༭���ĳ���
				oFCKeditor.Height = "400" '�༭���ĸ߶�
				oFCKeditor.Value = "" '����Ǹ��༭����ʼֵ
				oFCKeditor.Create "content" '�Ժ�༭��������ݶ��������content ȡ�ã�
			%>
    <label for="yzm">��֤�룺
    <input type="text" name="yzm" id="yzm" class="text yzm" />
    <img src="../include/code.asp" onclick="javascript:this.src='../include/code.asp?tm='+Math.random()" style="cursor:pointer" alt="��֤��" /></label>
    <input type="hidden" name="newsId" value="<%=showid%>" />
    <input type="submit" value="�������" name="send" class="submit" />
  </form>
</div>
</body>
</html>