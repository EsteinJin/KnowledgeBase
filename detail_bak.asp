<%@codepage =936%>
<%
	dim showid

	showid = request.querystring("ShowId")
	'�Ƿ�����
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("�Ƿ�����")
	end if
	
	dim title,content
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Article where CMS_ID="&showid
	rs.open sql,conn,1,1
	
	if rs.eof then
		call errorHistoryBack("�����ڴ�����")
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

<!--�����ʾ���ü���������ʾ-->
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
<span style=" position:absolute; right:10px; top:10px;">&nbsp;&nbsp;<a href="javascript:ExpandThis();void(0);">չ��</a>&nbsp;&nbsp;<a href="javascript:CollapsThis();void(0);">����</a></span>

<div id="detail">

	<h3><%=title%> </h3>
	<p class="d">TAG��ǩ��<%=tag%> | �����ؼ��֣�<%=keyword%> | �����ߣ�<%=name%> | ����ʱ�䣺<%=FormatDateTime(fdate,2)%><a href="admin/admin_article_Front_mof.asp?ShowId=<%=showid%>">�޸�</a> </p>
	<!--<p class="info"><%=info%></p>-->
	<%=content%>	
</div>
<div id="MyComment">
<h1>�������� </h1>

<p class="d">�����ߣ� |   ������IP��ַ�� |   ����ʱ�䣺  </p>
<div><strong>New Employee</strong></div>
<div><em>Contractor/Temporary</em></div>
<div>BGD ID to be created by SR, request   should come from BASF Internal Employee</div>
<div>All other request should be submitted   via AccessIT</div>
<p>&nbsp;</p>
<hr style="height:1px;border:none;border-top:1px dashed #0066CC;" />

<p class="d">�����ߣ� |   ������IP��ַ�� |   ����ʱ�䣺  </p>
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
				Set oFCKeditor = New FCKeditor '����һ���༭����ʵ��
				oFCKeditor.BasePath = "fckeditor/" '���ñ༭����·������վ���Ŀ¼�µ�һ��Ŀ¼
				oFCKeditor.ToolbarSet = "Basic" '�����ͼ�.Basic
				oFCKeditor.Width = "100%" '�༭���ĳ���
				oFCKeditor.Height = "400" '�༭���ĸ߶�
				oFCKeditor.Value = "" '����Ǹ��༭����ʼֵ
				oFCKeditor.Create "content" '�Ժ�༭��������ݶ��������content ȡ�ã�
			%>


</form>
</div>

	
</body>
</html>