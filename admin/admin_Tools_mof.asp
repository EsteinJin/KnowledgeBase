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
	
	'������ӵ�����
	if request.form("send") = "�޸�����" then
    dim ProjectName,Category,ToolsName,ToolsLink,ToolsHowTo,KnownIssue,EscalationStory
		ProjectName = replace(request.form("ProjectName"),"'","")
		Category = replace(request.form("Category"),"'","")
		ToolsName = replace(request.form("ToolsName"),"'","")
		ToolsLink = replace(request.form("ToolsLink"),"'","")
		ToolsHowTo = replace(request.form("ToolsHowTo"),"'","")
		KnownIssue = replace(request.form("KnownIssue"),"'","")
		EscalationHistory =  replace(request.form("EscalationHistory"),"'","")
		ItemId= request.form("ItemId")
		if len(ToolsName) < 2 or len(ToolsName) > 100 then
			call errorHistoryBack("�������Ʋ�С��2λ�����ߴ���100λ")
		end if	
		if len(ToolsLink) < 2 or len(ToolsLink) > 100 then
			call errorHistoryBack("�������Ӳ�С��2λ�����ߴ���100λ")
		end if				

		if len(ToolsHowTo) < 2  then
			call errorHistoryBack("����Ϊ��")
		end if				
		if len(KnownIssue) < 2  then
			call errorHistoryBack("����Ϊ��")
		end if				
		if len(EscalationHistory) < 2  then
			call errorHistoryBack("����Ϊ��")
		end if	

		'��������,�����ɹ�����ת�����ݹ���ҳ��
	updatesql = "update ToolsName set ProjectName='"&ProjectName&"',ToolsCategory='"&Category&"',ToolsName='"&ToolsName&"',ToolsLink='"&ToolsLink&"',ToolsHowTo='"&ToolsHowTo&"',KnownIssue='"&KnownIssue&"' ,EscalationHistory='"&EscalationHistory&"' where Item="&ItemId
		conn.execute(updatesql)
		call sussLoctionHref("�����޸����","admin_Tools_List.asp")
end if 		



	dim showid
	showid = request.querystring("ShowId")
	
	'�ж�showid��Ч
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("error Occured!")
	end if
	'�ж�showid�����Ŀ�Ƿ����	
	dim rs,sql,nToolsName,nToolsLink,nToolsHowTo,nKnownIssue,nEscalationHistory
	set rs = server.createobject("adodb.recordset")
	sql = "select * from ToolsName where Item="&showid
	rs.open sql,conn,1,1
	
	if rs.eof then
		call close_rs
		call close_conn
		call errorHistoryBack("Not Existing Data!")
	else
		nToolsName=rs("ToolsName")
		nToolsLink=rs("ToolsLink")
		nToolsHowTo=rs("ToolsHowTo")
		nKnownIssue=rs("KnownIssue")
		nEscalationHistory=rs("EscalationHistory")
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

<form name="add" id="articleadd" method="post" action="admin_Tools_mof.asp">
	<dl>
		<dt>�뷢������</dt>
		<dd>������Ŀ:
			<select name="ProjectName">
                    <option value="BASF">----BASF</option>					
					<option value="TowerWaton">----Tower Watson</option>					
					<option value="Coke">----Coke</option>					
					<option value="Nike">----Nike</option>					                                        
			</select>
		</dd>
		<dd>�������:
        			<select name="Category">
					<option value="ATOS">----ATOS</option>					
					<option value="Customer">----Customer</option>					
			</select>
        </dd>
		<dd>
				�������ƣ�
                <input type="text" name="ToolsName" value="<%=nToolsName%>" class="text" /> 
		</dd>

		<dd>�������ӣ�
                <textarea rows="2" name="ToolsLink"  ><%=nToolsLink%></textarea>
       
		</dd>
		<dd>
			<%
				Dim oFCKeditor
				Set oFCKeditor = New FCKeditor '����һ���༭����ʵ��
				oFCKeditor.BasePath = "fckeditor/" '���ñ༭����·������վ���Ŀ¼�µ�һ��Ŀ¼
				oFCKeditor.ToolbarSet = "Default" '�����ͼ�.Basic
				oFCKeditor.Width = "100%" '�༭���ĳ���
				oFCKeditor.Height = "300" '�༭���ĸ߶�
				oFCKeditor.Value = nToolsHowTo '����Ǹ��༭����ʼֵ
				oFCKeditor.Create "ToolsHowTo" '�Ժ�༭��������ݶ��������content ȡ�ã�
			%>
		</dd>

		<dd>
			<%
				
				Set oFCKeditor = New FCKeditor '����һ���༭����ʵ��
				oFCKeditor.BasePath = "fckeditor/" '���ñ༭����·������վ���Ŀ¼�µ�һ��Ŀ¼
				oFCKeditor.ToolbarSet = "Default" '�����ͼ�.Basic
				oFCKeditor.Width = "100%" '�༭���ĳ���
				oFCKeditor.Height = "300" '�༭���ĸ߶�
				oFCKeditor.Value = nKnownIssue '����Ǹ��༭����ʼֵ
				oFCKeditor.Create "KnownIssue" '�Ժ�༭��������ݶ��������content ȡ�ã�
			%>
			<%
				
				Set oFCKeditor = New FCKeditor '����һ���༭����ʵ��
				oFCKeditor.BasePath = "fckeditor/" '���ñ༭����·������վ���Ŀ¼�µ�һ��Ŀ¼
				oFCKeditor.ToolbarSet = "Default" '�����ͼ�.Basic
				oFCKeditor.Width = "100%" '�༭���ĳ���
				oFCKeditor.Height = "300" '�༭���ĸ߶�
				oFCKeditor.Value = nEscalationHistory '����Ǹ��༭����ʼֵ
				oFCKeditor.Create "EscalationHistory" '�Ժ�༭��������ݶ��������content ȡ�ã�
			%>


		</dd>
		<input type="hidden" name="ItemId" value="<%=showid%>" />
		<dd><input type="submit" onclick="return check();" name="send" value="�޸�����" /></dd>
	</dl>
</form>

</body>
</html>