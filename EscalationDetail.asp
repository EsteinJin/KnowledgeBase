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
	sql = "select * from EscalationLog where ID="&showid
	rs.open sql,conn,1,1
	
	if rs.eof then
		call errorHistoryBack("�����ڴ�����")
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
<span style="position:absolute; right:120px; top:170px;">&nbsp;&nbsp;<a href="javascript:ExpandThis();void(0);">չ��</a>&nbsp;&nbsp;<a href="javascript:CollapsThis();void(0);">����</a></span>
<!--#include file="header.asp"-->


<div id="detail">
	<h3><%=IssueSummary%></h3>
	<p class="d">�������ڣ�<span style="color:red;"><%=EscalatedDate%></span>| ������Ա��<span style="color:red;"><%=EscalatedBy%></span>| �������ͣ�<span style="color:red;"><%=EscalationType%></span>| ���ⵥ�ţ�<span style="color:red;"><%=TicketNumber%></span>| 
    ���<span style="color:red;"><%=Category%></span>|
    
     ������Ա��<span style="color:red;"><%=ResponsibleBy%></span>|״̬���٣�<span style="color:red;"><%=StatusTrack%></span>|<a href="admin/admin_Escalation_Front_mof.asp?ShowId=<%=showid%>">�޸�</a> </p>
    
<p class="contents"><%=IssueDetails%></p>
<p class="contents"><%=EscalationLog%></p>

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
  <form action="myCommentEscalation.asp" method="post">
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