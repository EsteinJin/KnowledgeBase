<%@codepage = 936%>
<!--�����ʾ���ü���������ʾ-->
<!
<!--#include file="../include/function.asp"-->
<!--#include file="fckeditor/fckeditor.asp"-->
<!--#include file="conn.asp"-->
<%

	if request.Form("send")="�������" then
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
			call errorHistoryBack("�ύ�˲���Ϊ��")
		end if	
	if len(IssueSummary) < 4 or len(IssueSummary) > 100 then
			call errorHistoryBack("���ⲻС��4λ�����ߴ���100λ")
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
msg.Body = "�ύ��Ա:"&EscalatedBy&"<br />����ʱ�䣺"&now()&"<br />�������:"&EscalationType&"<br />Ticket���룺"&TicketNumber&"<br />������ࣺ"&Category&"<br />�����ˣ�"&ResponsibleBy
msg.appendText "<br /><br /><br />"
msg.appendText "<br />"&replace(IssueDetails,"/upFile/","http://"&Request.ServerVariables("server_name")&"/upFile/")
msg.appendText "<br />"&EscalationLog
msg.Send( "apac.internal.siemens-it-solutions.com" )
	call sussLoctionHref("���������ɹ�","/EscalationDetail.asp?ShowId="&NewID)	
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
<form name="add" id="articleadd" method="post" action="admin_Escalation_add_Frount.asp">
<dl>
<dt>�뷢������</dt>

<dd>
�ύ��Ա:
<input type="text" name="EscalatedBy"  />
�������ͣ�
<input type="radio" name="EscalationType" value="ToolsIssue" checked="checked"/>��������
<input type="radio" name="EscalationType" value="ProcessIssue"/>��������
&nbsp;&nbsp;&nbsp;&nbsp;���ű��:
<input type="text" name="TicketNumber"  />
���ͷ���:
<input type="text" name="Category"  />
</dd>
<dd>
<dd>
������Ա:
<input type="radio" name="ResponsibleBy" value="LiuYang" checked="checked"/>LiuYang
<input type="radio" name="ResponsibleBy" value="JiangZhiMin"/>JiangZhiMin
<input type="radio" name="ResponsibleBy" value="ChenQiang"/>ChenQiang

��չ���:
<select  name="HandleStatus" disabled="disabled">
<option value="Logged">Logged</option>
<option value="Pending">Pending</option>
<option value="Assigned">Assigned</option>
<option value="Resolved">Resolved</option>
<option value="Closed">Closed</option>
</select>
</dd>

<dd>
��������:
<textarea rows="2"  style="width:100%;"  name="IssueSummary"></textarea>
</dd>
<dd>
<%
	Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '����һ���༭����ʵ��
	oFCKeditor.BasePath = "fckeditor/" '���ñ༭����·������վ���Ŀ¼�µ�һ��Ŀ¼
	oFCKeditor.ToolbarSet = "Default" '�����ͼ�.Basic
	oFCKeditor.Width = "100%" '�༭���ĳ���
	oFCKeditor.Height = "250" '�༭���ĸ߶�
	oFCKeditor.Value = "" '����Ǹ��༭����ʼֵ
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
oFCKeditor.Value = "��ʾ��<br>IVR���⣬��һʱ������ICM��¼��Genesis Ticket<br>�������ⰴ����������ICM,ͬʱҪ��Local IT����<br>" '����Ǹ��༭����ʼֵ
	oFCKeditor.Create "EscalationLog" '�Ժ�༭��������ݶ��������content ȡ�ã�
%>
</dd>

</dl>
<dd><input type="submit" onclick="return Morecheck();" name="send" value="�������" /></dd>
</form>

</body>
</html>
