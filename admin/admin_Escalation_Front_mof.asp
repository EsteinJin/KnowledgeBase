<%@codepage = 936%>
<!--�����ʾ���ü���������ʾ-->
<!
<!--#include file="../include/function.asp"-->
<!--#include file="fckeditor/fckeditor.asp"-->
<!--#include file="conn.asp"-->
<%

	if request.Form("send")="�޸�����" then
	dim EscalationType,TicketNumber,Category,IssueSummary,IssueDetails,ResponsibleBy,HandleStatus,EscalationLog,addsql
     id = request.form("id")
	EscalationType=replace(request.form("EscalationType"),"'","")
	TicketNumber=replace(request.form("TicketNumber"),"'","")
	Category=replace(request.form("Category"),"'","")
	IssueSummary=replace(request.form("IssueSummary"),"'","")
	IssueDetails=replace(request.form("IssueDetails"),"'","")
	ResponsibleBy=replace(request.form("ResponsibleBy"),"'","")
	StatusTrack=replace(request.form("HandleStatus"),"'","")
	EscalationLog=replace(request.form("EscalationLog"),"'","")
	EscalatedBy=replace(request.form("EscalatedBy"),"'","")

  updatesql="update EscalationLog set EscalationType='"&EscalationType&"',TicketNumber='"&TicketNumber&"',Category='"&Category&"',IssueSummary='"&IssueSummary&"',IssueDetails='"&IssueDetails&"',ResponsibleBy='"&ResponsibleBy&"',StatusTrack='"&StatusTrack&"',EscalationLog='"&EscalationLog&"',EscalatedBy='"&EscalatedBy&"' where ID="&id

	conn.execute(updatesql)
	call sussLoctionHref("���ݸ��³ɹ�","/EscalationDetail.asp?ShowId="&id)	
	end if 
	showid = request.querystring("ShowId")

	'�Ƿ�����
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("�Ƿ�����")
	end if
	set rs = server.createobject("adodb.recordset")
	sql = "select * from EscalationLog where ID="&showid
	rs.open sql,conn,1,1	
	if rs.eof then
		call close_rs
		call close_conn
		call errorHistoryBack("�����ڴ�����")
		'����ȫ������֤���ݣ��Ѿ��ɹ�
	else
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
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��̨����</title>
<link rel="stylesheet" type="text/css" href="style/admin.css" />
<script type="text/javascript" src="js/content.js"></script>
</head>
<body>
<form name="add" id="articleadd" method="post" action="admin_Escalation_Front_mof.asp">
<input type="hidden" value="<%=showid%>" name="id" />
<dl>
<dt>�뷢������</dt>

<dd>
�ύ��Ա:
<input type="text" name="EscalatedBy"  value="<%=EscalatedBy%>" />
�������ͣ�<font style="color:red"><%=EscalationType%></font>
<input type="radio" name="EscalationType" value="ToolsIssue" checked="checked"/>��������
<input type="radio" name="EscalationType" value="ProcessIssue"/>��������
&nbsp;&nbsp;&nbsp;&nbsp;���ű��:
<input type="text" name="TicketNumber" value="<%=TicketNumber%>"  />
���ͷ���:
<input type="text" name="Category" value="<%=Category%>"   />
</dd>
<dd>
<dd>
������Ա:<font style="color:red"><%=ResponsibleBy%></font>
<input type="radio" name="ResponsibleBy" value="LiuYang" checked="checked"/>LiuYang
<input type="radio" name="ResponsibleBy" value="JiangZhiMin"/>JiangZhiMin
<input type="radio" name="ResponsibleBy" value="ChenQiang"/>ChenQiang

��չ���:<font style="color:red"><%=StatusTrack%></font>
<select  name="HandleStatus">
<option value="Logged">Logged</option>
<option value="Pending">Pending</option>
<option value="Assigned">Assigned</option>
<option value="Resolved">Resolved</option>
<option value="Closed">Closed</option>
</select>
</dd>

<dd>
��������:
<textarea rows="2"  style="width:100%;"  name="IssueSummary"><%=IssueSummary%></textarea>
</dd>
<dd>
<%
	Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '����һ���༭����ʵ��
	oFCKeditor.BasePath = "fckeditor/" '���ñ༭����·������վ���Ŀ¼�µ�һ��Ŀ¼
	oFCKeditor.ToolbarSet = "Default" '�����ͼ�.Basic
	oFCKeditor.Width = "100%" '�༭���ĳ���
	oFCKeditor.Height = "250" '�༭���ĸ߶�
	oFCKeditor.Value = IssueDetails '����Ǹ��༭����ʼֵ
	oFCKeditor.Create "IssueDetails" '�Ժ�༭��������ݶ��������content ȡ�ã�
%>
</dd>

<dd>
<%
	
	Set oFCKeditor = New FCKeditor '����һ���༭����ʵ��
	oFCKeditor.BasePath = "fckeditor/" '���ñ༭����·������վ���Ŀ¼�µ�һ��Ŀ¼
	oFCKeditor.ToolbarSet = "Default" '�����ͼ�.Basic
	oFCKeditor.Width = "100%" '�༭���ĳ���
	oFCKeditor.Height = "250" '�༭���ĸ߶�
	oFCKeditor.Value = EscalationLog '����Ǹ��༭����ʼֵ
	oFCKeditor.Create "EscalationLog" '�Ժ�༭��������ݶ��������content ȡ�ã�
%>
</dd>

</dl>
<dd><input type="submit" onclick="return Morecheck();" name="send" value="�޸�����" /></dd>
</form>

</body>
</html>
