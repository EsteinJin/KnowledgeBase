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
	call sussLoctionHref("���������ɹ�","admin_compliment.asp")	
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
<form name="add" id="articleadd" method="post" action="admin_compliment_add.asp">
<dl>
<dt>�뷢������</dt>
<dd>Agent����:

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

�����ԣ�
<input type="text" name="CMS_PraisedBy" />

������Դ:
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
<font style="color:red; font-weight:bold;">��ѡ������KPI����</font>
</select>
</dd>
<dd>

<%
	Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '����һ���༭����ʵ��
	oFCKeditor.BasePath = "fckeditor/" '���ñ༭����·������վ���Ŀ¼�µ�һ��Ŀ¼
	oFCKeditor.ToolbarSet = "Default" '�����ͼ�.Basic
	oFCKeditor.Width = "50%" '�༭���ĳ���
	oFCKeditor.Height = "250" '�༭���ĸ߶�
	oFCKeditor.Value = "�����������������" '����Ǹ��༭����ʼֵ
	oFCKeditor.Create "CMS_Title" '�Ժ�༭��������ݶ��������content ȡ�ã�
%>


<%
	'Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '����һ���༭����ʵ��
	oFCKeditor.BasePath = "fckeditor/" '���ñ༭����·������վ���Ŀ¼�µ�һ��Ŀ¼
	oFCKeditor.ToolbarSet = "Default" '�����ͼ�.Basic
	oFCKeditor.Width = "50%" '�༭���ĳ���
	oFCKeditor.Height = "250" '�༭���ĸ߶�
	oFCKeditor.Value = "�����¼QA��Comment" '����Ǹ��༭����ʼֵ
	oFCKeditor.Create "CMS_QAComment" '�Ժ�༭��������ݶ��������content ȡ�ã�
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
	oFCKeditor.Value = "�����¼���Raw Data�����ӻ��ͼ" '����Ǹ��༭����ʼֵ
	oFCKeditor.Create "CMS_Evidence" '�Ժ�༭��������ݶ��������content ȡ�ã�
%>

<%
	'Dim oFCKeditor
	Set oFCKeditor = New FCKeditor '����һ���༭����ʵ��
	oFCKeditor.BasePath = "fckeditor/" '���ñ༭����·������վ���Ŀ¼�µ�һ��Ŀ¼
	oFCKeditor.ToolbarSet = "Default" '�����ͼ�.Basic
	oFCKeditor.Width = "50%" '�༭���ĳ���
	oFCKeditor.Height = "250" '�༭���ĸ߶�
	oFCKeditor.Value = "�����¼Agent�������ĵ�����" '����Ǹ��༭����ʼֵ
	oFCKeditor.Create "CMS_Learnd" '�Ժ�༭��������ݶ��������content ȡ�ã�
%>

</dd>

<dd><input type="submit" onclick="return Morecheck();" name="send" value="�������" /></dd>
</dl>

</form>
</body>
</html>