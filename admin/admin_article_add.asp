<%@codepage = 936%>


<!--#include file="../include/function.asp"-->
<!--#include file="fckeditor/fckeditor.asp"-->
<!--#include file="conn.asp"-->
<%

	if session("Admin") = "" then
		call sussLoctionHref("Error Occured!","admin_login.asp")
	end if
	
	'������ӵ�����
	if request.form("send") = "Submit" then
		dim addsql,sort,title,top,rmp,pic,bold,color,tag,keyword,info,name
		sort = request.form("sort")
		title = replace(request.form("title"),"'","")
		top = request.form("top")
		rmp = request.form("rmp")
		pic = request.form("pic")
		bold = request.form("bold")
		color =  request.form("color")
		tag =  request.form("tag")
		keyword =  request.form("keyword")
		info =  replace(request.form("info"),"'","")
		name =  request.form("name")
		content =  replace(request.form("content"),"'","")
		
		'��ʼʹ��VBScript���ֶν�����֤
		
		'����=>����Ϊ�ղ��Ҳ�������4λ�����ܴ���100λ
		if len(title) < 4 or len(title) > 100 then
			call errorHistoryBack("No less than 4��no More than 100")
		end if
		
		'�����߲���Ϊ��
		if name = "" then
			call errorHistoryBack("This Value can not be null")
		end if
		
		'���ݼ��
		if len(info) < 10 or len(info) > 200 then
			call errorHistoryBack("KB Info Can not be less then 10, no more than  200")
		end if 
		
		'��Ҫ����
		if len(content) < 10 then
			call errorHistoryBack("Contents no less than 10 ")
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
		
		'�������֮ǰ���Ƚ���Tag��ǩ����ӻ����
		dim tagArr,i,tagrs,tagsql
		tagArr = split(tag,",")  'ͨ��split��������Tag��ǩ����
		
		'ͨ��ȡ����������±��������Tag��ǩ���в���
		for i = lbound(tagArr) to ubound(tagArr)
			'ȡ�ñ�ǩ�����ж��Ƿ������ݿ����������ǩ
			'����У����ۼ�+1
			'���û�У�������һ��Tag��ǩ
			set tagrs = server.createobject("adodb.recordset")
			tagsql = "select * from CMS_Tag where CMS_TagName='"&tagArr(i)&"'"
			tagrs.open tagsql,conn,1,1
			
			if not tagrs.eof then
				'�����ݣ��ۼ�
				sql = "Update CMS_Tag Set CMS_TagCount=CMS_TagCount+1 where CMS_TagName='"&tagArr(i)&"'"
				conn.execute(sql)
			else
				'û���ݣ�����	
				sql = "Insert into CMS_Tag (CMS_TagName,CMS_Date) values ('"&tagArr(i)&"',now())"			
				conn.execute(sql)
			end if
			tagrs.close
			set tagrs = nothing
		next
	'	response.Write sort
'		response.Write title
'		response.Write top
'		response.Write rmp
'		response.Write pic
'		response.Write color
'		response.Write tag
'		response.Write keyword
'		response.Write name
'		response.Write info
'		response.Write content
		
		'��������,�����ɹ�����ת�����ݹ���ҳ��
		addsql = "Insert into CMS_Article (CMS_Sort,CMS_Title,CMS_Top,CMS_Rmp,CMS_Pic,CMS_Bold,CMS_Color,CMS_Tag,CMS_Keyword,CMS_Name,CMS_Info,CMS_Content,CMS_Date) values ("&sort&",'"&title&"',"&top&","&rmp&","&pic&","&bold&",'"&color&"','"&tag&"','"&keyword&"','"&name&"','"&info&"','"&content&"',now())"
		conn.execute(addsql)
		call sussLoctionHref("Added Successfully!","admin_article.asp")
		
		
		
	end if
	

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>POST KB</title>
<link rel="stylesheet" type="text/css" href="style/admin.css" />
<script type="text/javascript" src="js/content.js"></script>
</head>
<body>

<form name="add" id="articleadd" method="post" action="admin_article_add.asp">
	<dl>
		<dt>KB POSTING PAGE</dt>
		<dd>Type:&nbsp;&nbsp;
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
					<option value="<%=rs2("CMS_ID")%>">----<%=rs2("CMS_NavName")%></option>					
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
		<dd>Caption��<input type="text" name="title" class="text" /></dd>
		<dd>
				Atribute��
				<input type="checkbox" class="radio" name="top" value="1" /> <span  style="color:red; font-weight:bold;">Top Article</span>
				<input type="checkbox" class="radio" name="rmp" value="1" /> <span  style="color:red; font-weight:bold;">Cascade Article</span>
			    <input type="checkbox" class="radio" name="bold" value="1" /> Processes
				<input type="checkbox" class="radio" name="pic" disabled="disabled" value="1" /> High Light                
		</dd>
		<dd>Color��
			<input type="radio" name="color" value="black" checked="checked" /> Black
			<input type="radio" name="color" value="red" /> Red
			<input type="radio"  name="color" value="green" /> Green
		</dd>
		<dd>TAG ��&nbsp;&nbsp;&nbsp;<input type="text" name="tag" class="text" /> (Using , to split )</dd>
		<dd>Key Word��<input type="text" name="keyword" class="text" /> (Using , to split )</dd>
		<dd>Poster��<input type="text" name="name" value="<%=session("Admin")%>" class="text" /></dd>
		<dd>
				Info��<textarea cols="30" rows="2" name="info"></textarea>	
		</dd>
		<dd>
			<%
				Dim oFCKeditor
				Set oFCKeditor = New FCKeditor '����һ���༭����ʵ��
				oFCKeditor.BasePath = "fckeditor/" '���ñ༭����·������վ���Ŀ¼�µ�һ��Ŀ¼
				oFCKeditor.ToolbarSet = "Default" '�����ͼ�.Basic
				oFCKeditor.Width = "100%" '�༭���ĳ���
				oFCKeditor.Height = "400" '�༭���ĸ߶�
				oFCKeditor.Value = "" '����Ǹ��༭����ʼֵ
				oFCKeditor.Create "content" '�Ժ�༭��������ݶ��������content ȡ�ã�
			%>
		</dd>
		<dd><input type="submit" onclick="return check();" name="send" value="Submit" /></dd>
	</dl>
</form>

</body>
</html>