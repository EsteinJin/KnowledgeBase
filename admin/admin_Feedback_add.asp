<%@codepage = 936%>
<!--�����ʾ���ü���������ʾ-->
<!
<!--#include file="../include/function.asp"-->
<!--#include file="fckeditor/fckeditor.asp"-->
<!--#include file="conn.asp"-->
<%
	if session("Admin") = "" then
		call sussLoctionHref("�Ƿ�����","admin_login.asp")
	end if
	if request.Form("send")="�������" then
	dim CMS_MonitorDate,CMS_Source,CMS_TicketType,CMS_TicketNumber,CMS_TicketSummary,CMS_AgentName,CMS_TicketStatus,CMS_IsComplaint,CMS_CompliantSource,CMS_ComlaintBy,CMS_RootCause,CMS_HandleStatus,CMS_PAHighLight,CMS_CoachingLog,CMS_InternalUAT,CMS_RealCase,CMS_Point,addsql

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
	addsql="Insert into CMS_Complaint(CMS_MonitorDate,CMS_Source,CMS_OSDNumber,CMS_GAHDNumber,CMS_RequestID,CMS_IMACNumber,CMS_BTRequestNumber,CMS_BTIncidentNumber,CMS_TicketSummary,CMS_AgentName,CMS_TicketStatus,CMS_IsComplaint,CMS_CompliantSource,CMS_ComlaintBy,CMS_RootCause,CMS_HandleStatus,CMS_PAHighLight,CMS_CoachingLog,CMS_InternalUAT,CMS_RealCase,CMS_Point) values(now(),'"&CMS_Source&"','"&CMS_OSDNumber&"','"&CMS_GAHDNumber&"','"&CMS_RequestID&"','"&CMS_IMACNumber&"','"&CMS_BTRequestNumber&"','"&CMS_BTIncidentNumber&"','"&CMS_TicketSummary&"','"&CMS_AgentName&"','"&CMS_TicketStatus&"','"&CMS_IsComplaint&"','"&CMS_CompliantSource&"','"&CMS_ComlaintBy&"','"&CMS_RootCause&"','"&CMS_HandleStatus&"','"&CMS_PAHighLight&"','"&CMS_CoachingLog&"','"&CMS_InternalUAT&"','"&CMS_RealCase&"','"&CMS_Point&"') "
	conn.execute(addsql)
	call sussLoctionHref("���������ɹ�","admin_Feedback.asp")	
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
<form name="add" id="articleadd" method="post" action="admin_Feedback_add.asp">
<dl>
<dt>�뷢������</dt>
<dd>Ͷ����Դ:
<select  name="FeedbackSource">
<option value="Call">Call</option>
<option value="Ticket">Ticket</option>
<option value="eMail">eMail</option>

</select>

OSD����:
<input type="text" name="OSDNumber"  />
GAHD����:
<input type="text" name="GAHDNumber"  />
RequestID:
<input type="text" name="RequestNumber"  />
</dd>
<dd>
IMAC ID:
<input type="text" name="IMACNumber"  />
BTRequestID:
<input type="text" name="BTRequestNumber"  />
BTIncidentID:
<input type="text" name="BTIncidentNumber"  />
</dd>

<dd>Agent���ƣ�
<input type="text" name="AgentName" />

����״̬:
<select  name="TicketStatus">
<option value="Pending">Pending</option>
<option value="Assigned">Assgined</option>
<option value="Resolved">Resolved</option>
<option value="ReOpened">ReOpened</option>
<option value="Closed">Closed</option>

</select>
Ͷ����Դ:
<select  name="CompliantSource">
<option value="Global">Global</option>
<option value="EndUser">End User</option>
<option value="Atos">Atos</option>
<option value="LocalIM">Local IM</option>
<option value="QC">QC</option>
</select>


</dd>

<dd>�Ƿ�Ͷ�ߣ�
<input type="radio" name="IsComplaint" value="Yes" checked="checked"/>Yes
<input type="radio" name="IsComplaint" value="No"/>No
Ͷ����Ա��
<input type="text" name="ComplaintBy"  />
Ͷ��׷��:
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
KPI Point:
<input type="radio" name="AgentKPI" value="-3" />-3
<input type="radio" name="AgentKPI" value="-2" />-2
<input type="radio" name="AgentKPI" value="-1" />-1
<input type="radio" name="AgentKPI" value="0" />0
<font style="color:red; font-weight:bold;">��ѡ��۳���KPI����</font>
</dd>


<dd>����������
<textarea rows="2" style="width:100%;"  name="IssueSummary"></textarea>
</dd>


<dd>

<%
	Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '����һ���༭����ʵ��
	oFCKeditor.BasePath = "fckeditor/" '���ñ༭����·������վ���Ŀ¼�µ�һ��Ŀ¼
	oFCKeditor.ToolbarSet = "Default" '�����ͼ�.Basic
	oFCKeditor.Width = "50%" '�༭���ĳ���
	oFCKeditor.Height = "250" '�༭���ĸ߶�
	oFCKeditor.Value = "������дRoot Cause" '����Ǹ��༭����ʼֵ
	oFCKeditor.Create "RootCause" '�Ժ�༭��������ݶ��������content ȡ�ã�
%>


<%
	'Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '����һ���༭����ʵ��
	oFCKeditor.BasePath = "fckeditor/" '���ñ༭����·������վ���Ŀ¼�µ�һ��Ŀ¼
	oFCKeditor.ToolbarSet = "Default" '�����ͼ�.Basic
	oFCKeditor.Width = "50%" '�༭���ĳ���
	oFCKeditor.Height = "250" '�༭���ĸ߶�
	oFCKeditor.Value = "������дPA High Light����" '����Ǹ��༭����ʼֵ
	oFCKeditor.Create "PAHighLight" '�Ժ�༭��������ݶ��������content ȡ�ã�
%>
</dd>
<dd>

<%
	'Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '����һ���༭����ʵ��
	oFCKeditor.BasePath = "fckeditor/" '���ñ༭����·������վ���Ŀ¼�µ�һ��Ŀ¼
	oFCKeditor.ToolbarSet = "Default" '�����ͼ�.Basic
	oFCKeditor.Width = "50%" '�༭���ĳ���
	oFCKeditor.Height = "250" '�༭���ĸ߶�
	oFCKeditor.Value = "������дTrainer Coaching����" '����Ǹ��༭����ʼֵ
	oFCKeditor.Create "CoachingLog" '�Ժ�༭��������ݶ��������content ȡ�ã�
%>

<%
	'Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '����һ���༭����ʵ��
	oFCKeditor.BasePath = "fckeditor/" '���ñ༭����·������վ���Ŀ¼�µ�һ��Ŀ¼
	oFCKeditor.ToolbarSet = "Default" '�����ͼ�.Basic
	oFCKeditor.Width = "50%" '�༭���ĳ���
	oFCKeditor.Height = "250" '�༭���ĸ߶�
	oFCKeditor.Value = "����ճ�����ϴ�Internal UAT����" '����Ǹ��༭����ʼֵ
	oFCKeditor.Create "InternalUAT" '�Ժ�༭��������ݶ��������content ȡ�ã�
%>

</dd>
<dd>



<%
	'Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '����һ���༭����ʵ��
	oFCKeditor.BasePath = "fckeditor/" '���ñ༭����·������վ���Ŀ¼�µ�һ��Ŀ¼
	oFCKeditor.ToolbarSet = "Default" '�����ͼ�.Basic
	oFCKeditor.Width = "100%" '�༭���ĳ���
	oFCKeditor.Height = "250" '�༭���ĸ߶�
	oFCKeditor.Value = "����ճ�����ϴ�Real Case�ĵ�������" '����Ǹ��༭����ʼֵ
	oFCKeditor.Create "RealCase" '�Ժ�༭��������ݶ��������content ȡ�ã�
%>

</dd>
<dd><input type="submit" onclick="return Morecheck();" name="send" value="�������" /></dd>
</dl>

</form>
</body>
</html>