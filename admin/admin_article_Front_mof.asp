<%@codepage = 936%>

<!--#include file="../include/function.asp"-->
<!
<!--#include file="fckeditor/fckeditor.asp"-->
<!--#include file="conn.asp"-->
<%

	'�ж��Ƿ�Ϊ����Ա
	if session("Admin") = "" then
		call sussLoctionHref("�Ƿ�����","admin_login.asp")
	end if
	
	dim id,showid,sort,title,top,rmp,pic,bold,color,tag,keyword,name,info,content
	
	'�����޸ĵ�����
if request.form("send") = "�޸�����" then
id = request.form("id")
sort = replace(request.form("sort"),"'","")
title = replace(request.form("title"),"'","")
top = replace(request.form("top"),"'","")
rmp = replace(request.form("rmp"),"'","")
pic = replace(request.form("pic"),"'","")
bold = replace(request.form("bold"),"'","")
color =  replace(request.form("color"),"'","")
tag =  replace(request.form("tag"),"'","")
keyword =  replace(request.form("keyword"),"'","")
info =  replace(request.form("info"),"'","")
name =  replace(request.form("name"),"'","")
content =  replace(request.form("content"),"'","")
		
		'��ʼʹ��VBScript���ֶν�����֤
		
		'����=>����Ϊ�ղ��Ҳ�������4λ�����ܴ���100λ
		if len(title) < 4 or len(title) > 100 then
			call errorHistoryBack("���ⲻС��4λ�����ߴ���100λ")
		end if
		
		'�����߲���Ϊ��
		if name = "" then
			call errorHistoryBack("�����˲���Ϊ��")
		end if
		
		'���ݼ��
		if len(info) < 10 or len(info) > 200 then
			call errorHistoryBack("���ݼ�鲻��С��10λ������200λ")
		end if 
		
		'��Ҫ����
		if len(content) < 10 then
			call errorHistoryBack("���ݲ�������10λ")
		end if
		
		
		if top = "" then 
			top =0
		end if
		
		if rmp = "" then 
			rmp =0
		end if
		
		if pic =  "" then
			pic = 0
		end if
		
		if bold = "" then
			bold = 0
		end if
		
		
		'��ʼִ���޸�����
		updatesql = "update CMS_Article set CMS_Top="&top&",CMS_Pic="&pic&",CMS_Bold="&bold&",CMS_Rmp="&rmp&",CMS_Title='"&title&"',CMS_Sort="&sort&",CMS_Name='"&name&"',CMS_Info='"&info&"',CMS_Color='"&color&"',CMS_Tag='"&tag&"',CMS_Keyword='"&keyword&"',CMS_Content='"&content&"' where CMS_ID="&id
		conn.execute(updatesql)
		call sussLoctionHref("�����޸����","/sap/detail.asp?ShowId="&id)
	end if
	
	
	showid = request.querystring("ShowId")

	'�Ƿ�����
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("�Ƿ�����")
	end if
	
	'�ж������Ƿ����
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Article where CMS_ID="&showid
	rs.open sql,conn,1,1
	
	if rs.eof then
		call close_rs
		call close_conn
		call errorHistoryBack("�����ڴ�����")
		'����ȫ������֤���ݣ��Ѿ��ɹ�
	else
		sort = rs("CMS_Sort")
		top = rs("CMS_Top")
		rmp = rs("CMS_Rmp")
		pic = rs("CMS_Pic")
		bold = rs("CMS_Bold")
		color = rs("CMS_Color")
		title = rs("CMS_Title")
		tag = rs("CMS_Tag")
		keyword = rs("CMS_Keyword")
		name = rs("CMS_Name")
		info = rs("CMS_Info")
		content = rs("CMS_Content")
	end if
	
	call close_rs
	
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

<form name="add" id="articleadd" method="post" action="admin_article_Front_mof.asp">
	<input type="hidden" value="<%=showid%>" name="id" />
	<dl>
		<dt>���޸�����</dt>
		<dd>��ѡ�����
			<select name="sort">
				<%
					dim rs,sql,rs2,sql2
					set rs = server.createobject("adodb.recordset")
					sql = "select * from CMS_Nav where CMS_Sid=0 order by CMS_Sort asc"
					rs.open sql,conn,1,1
					do while not rs.eof
				%>
				
								
				<optgroup label="<%=rs("CMS_NavName")%>">
					
					<%
						set rs2 = server.createobject("adodb.recordset")
						sql2 = "select * from CMS_Nav where CMS_Sid<>0 and CMS_Sid="&rs("CMS_ID")
						rs2.open sql2,conn,1,1
						do while not rs2.eof
					%>
					<option value="<%=rs2("CMS_ID")%>"<%
							'����㴫������ID�е�sort���޸���ѭ����sortһ�µĻ�����ô��selected
							if rs2("CMS_ID") = sort then
								response.write "selected='selected'"
							end if
						%>>----<%=rs2("CMS_NavName")%></option>					
					<%
							rs2.movenext
						loop
					%>
				</optgroup>
				
				
				<%
						rs.movenext
					loop
				%>
			</select>
		</dd>
		<dd>���±��⣺<input type="text" name="title" value="<%=title%>" class="text" /> <span>*</span></dd>
		<dd>
				�Զ������ԣ�
				<input type="checkbox" <%if top = 1 then response.write "checked='checked'"%> class="radio" name="top" value="1" /> ͷ��
				<input type="checkbox" <%if rmp = 1 then response.write "checked='checked'"%> class="radio" name="rmp" value="1" /> �Ƽ�
				<input type="checkbox" <%if pic = 1 then response.write "checked='checked'"%> class="radio" name="pic" value="1" /> ��ͼ
				<input type="checkbox" <%if bold = 1 then response.write "checked='checked'"%> class="radio" name="bold" value="1" /> �Ӵ�
		</dd>
		<dd>������ɫ��
			<input type="radio" name="color" value="black" <%if color="black" then response.write "checked='checked'"%> /> ��ɫ
			<input type="radio" name="color" value="red" <%if color="red" then response.write "checked='checked'"%> /> ��ɫ
			<input type="radio"  name="color" value="green" <%if color="green" then response.write "checked='checked'"%> /> ��ɫ
		</dd>
		<dd>TAG ��ǩ��<input type="text" value="<%=tag%>" name="tag" class="text" /> (�����ǩ��ʹ��Ӣ�Ķ���","����)</dd>
		<dd>�� �� �֣�<input type="text" value="<%=keyword%>" name="keyword" class="text" /> (����ؼ�����ʹ��Ӣ�Ķ���","����)</dd>
		<dd>�� �� �ߣ�<input type="text" name="name" value="<%=name%>" class="text" />  <span>*</span></dd>
		<dd>
				���¼�飺<textarea cols="30" rows="2" name="info"><%=info%></textarea>	  <span>*</span>
		</dd>
		<dd>
			<%
				Dim oFCKeditor
				Set oFCKeditor = New FCKeditor '����һ���༭����ʵ��
				oFCKeditor.BasePath = "fckeditor/" '���ñ༭����·������վ���Ŀ¼�µ�һ��Ŀ¼
				oFCKeditor.ToolbarSet = "Default" '�����ͼ�.Basic
				oFCKeditor.Width = "100%" '�༭���ĳ���
				oFCKeditor.Height = "400" '�༭���ĸ߶�
				oFCKeditor.Value = content '����Ǹ��༭����ʼֵ
				oFCKeditor.Create "content" '�Ժ�༭��������ݶ��������content ȡ�ã�
			%>
		</dd>
		<dd><input type="submit" onclick="return check();" name="send" value="�޸�����" /></dd>
	</dl>
</form>

</body>
</html>