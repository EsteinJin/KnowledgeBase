<%@codepage = 936%>


<!--#include file="../include/function.asp"-->
<!--#include file="fckeditor/fckeditor.asp"-->
<!--#include file="conn.asp"-->
<%

	if session("Admin") = "" then
		call sussLoctionHref("Error Occured!","admin_login.asp")
	end if
	
	'接收添加的内容
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
		
		'开始使用VBScript对字段进行验证
		
		'标题=>不能为空并且不能少于4位，不能大于100位
		if len(title) < 4 or len(title) > 100 then
			call errorHistoryBack("No less than 4，no More than 100")
		end if
		
		'发布者不能为空
		if name = "" then
			call errorHistoryBack("This Value can not be null")
		end if
		
		'内容简介
		if len(info) < 10 or len(info) > 200 then
			call errorHistoryBack("KB Info Can not be less then 10, no more than  200")
		end if 
		
		'主要内容
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
		
		'添加数据之前，先进行Tag标签的添加或积累
		dim tagArr,i,tagrs,tagsql
		tagArr = split(tag,",")  '通过split函数，将Tag标签分组
		
		'通过取得数组的上下标来逐个对Tag标签进行操作
		for i = lbound(tagArr) to ubound(tagArr)
			'取得标签，先判断是否在数据库里有这个标签
			'如果有，就累计+1
			'如果没有，就新增一个Tag标签
			set tagrs = server.createobject("adodb.recordset")
			tagsql = "select * from CMS_Tag where CMS_TagName='"&tagArr(i)&"'"
			tagrs.open tagsql,conn,1,1
			
			if not tagrs.eof then
				'有数据，累计
				sql = "Update CMS_Tag Set CMS_TagCount=CMS_TagCount+1 where CMS_TagName='"&tagArr(i)&"'"
				conn.execute(sql)
			else
				'没数据，新增	
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
		
		'新增数据,发布成功后跳转到内容管理页面
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
		<dd>Caption：<input type="text" name="title" class="text" /></dd>
		<dd>
				Atribute：
				<input type="checkbox" class="radio" name="top" value="1" /> <span  style="color:red; font-weight:bold;">Top Article</span>
				<input type="checkbox" class="radio" name="rmp" value="1" /> <span  style="color:red; font-weight:bold;">Cascade Article</span>
			    <input type="checkbox" class="radio" name="bold" value="1" /> Processes
				<input type="checkbox" class="radio" name="pic" disabled="disabled" value="1" /> High Light                
		</dd>
		<dd>Color：
			<input type="radio" name="color" value="black" checked="checked" /> Black
			<input type="radio" name="color" value="red" /> Red
			<input type="radio"  name="color" value="green" /> Green
		</dd>
		<dd>TAG ：&nbsp;&nbsp;&nbsp;<input type="text" name="tag" class="text" /> (Using , to split )</dd>
		<dd>Key Word：<input type="text" name="keyword" class="text" /> (Using , to split )</dd>
		<dd>Poster：<input type="text" name="name" value="<%=session("Admin")%>" class="text" /></dd>
		<dd>
				Info：<textarea cols="30" rows="2" name="info"></textarea>	
		</dd>
		<dd>
			<%
				Dim oFCKeditor
				Set oFCKeditor = New FCKeditor '创建一个编辑器的实例
				oFCKeditor.BasePath = "fckeditor/" '配置编辑器的路径，我站点根目录下的一个目录
				oFCKeditor.ToolbarSet = "Default" '完整和简化.Basic
				oFCKeditor.Width = "100%" '编辑器的长度
				oFCKeditor.Height = "400" '编辑器的高度
				oFCKeditor.Value = "" '这个是给编辑器初始值
				oFCKeditor.Create "content" '以后编辑器里的内容都是由这个content 取得，
			%>
		</dd>
		<dd><input type="submit" onclick="return check();" name="send" value="Submit" /></dd>
	</dl>
</form>

</body>
</html>