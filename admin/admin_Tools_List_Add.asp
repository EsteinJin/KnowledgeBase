<%@codepage = 936%>
<!--上面表示采用简体中文显示-->
<!
<!--#include file="../include/function.asp"-->
<!--#include file="fckeditor/fckeditor.asp"-->
<!--#include file="conn.asp"-->
<%

	if session("Admin") = "" then
		call sussLoctionHref("非法操作","admin_login.asp")
	end if
	
	'接收添加的内容
	if request.form("send") = "添加内容" then
    dim ProjectName,Category,ToolsName,ToolsLink,ToolsHowTo,KnownIssue,EscalationStory
		ProjectName = replace(request.form("ProjectName"),"'","")
		Category = replace(request.form("Category"),"'","")
		ToolsName = replace(request.form("ToolsName"),"'","")
		ToolsLink = replace(request.form("ToolsLink"),"'","")
		ToolsHowTo = replace(request.form("ToolsHowTo"),"'","")
		KnownIssue = replace(request.form("KnownIssue"),"'","")
		EscalationHistory =  replace(request.form("EscalationHistory"),"'","")
		
		if len(ToolsName) < 2 or len(ToolsName) > 100 then
			call errorHistoryBack("工具名称不小于2位，或者大于100位")
		end if	
		if len(ToolsLink) < 2 or len(ToolsLink) > 300 then
			call errorHistoryBack("工具链接不小于2位，或者大于100位")
		end if				

		if len(ToolsHowTo) < 2  then
			call errorHistoryBack("不能为空")
		end if				
		if len(KnownIssue) < 2  then
			call errorHistoryBack("不能为空")
		end if				
		if len(EscalationHistory) < 2  then
			call errorHistoryBack("不能为空")
		end if	

		'新增数据,发布成功后跳转到内容管理页面
		addsql = "Insert into ToolsName (ProjectName,ToolsCategory,ToolsName,ToolsLink,ToolsHowTo,KnownIssue,EscalationHiStory) values ('"&ProjectName&"','"&Category&"','"&ToolsName&"','"&ToolsLink&"','"&ToolsHowTo&"','"&KnownIssue&"','"&EscalationHistory&"')"
		conn.execute(addsql)
		call sussLoctionHref("内容新增成功","admin_Tools_List.asp")




	end if
	

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>后台管理</title>
<link rel="stylesheet" type="text/css" href="style/admin.css" />
<script type="text/javascript" src="js/content.js"></script>
</head>
<body>

<form name="add" id="articleadd" method="post" action="admin_Tools_List_Add.asp">
	<dl>
		<dt>请发布文章</dt>
		<dd>所属项目:
			<select name="ProjectName">
					<option value="BASF">----BASF</option>					
					<option value="TowerWaton">----Tower Watson</option>					
					<option value="Coke">----Coke</option>					
					<option value="Nike">----Nike</option>					                                        
			</select>
		</dd>
		<dd>所属类别:
        			<select name="Category">
					<option value="ATOS">----ATOS</option>					
					<option value="Customer">----Customer</option>					
			</select>
        </dd>
		<dd>
				工具名称：
                <input type="text" name="ToolsName" class="text" /> 
		</dd>

		<dd>工具链接：
                <textarea rows="2" name="ToolsLink" ></textarea>
       
		</dd>
		<dd>
			<%
				Dim oFCKeditor
				Set oFCKeditor = New FCKeditor '创建一个编辑器的实例
				oFCKeditor.BasePath = "fckeditor/" '配置编辑器的路径，我站点根目录下的一个目录
				oFCKeditor.ToolbarSet = "Default" '完整和简化.Basic
				oFCKeditor.Width = "100%" '编辑器的长度
				oFCKeditor.Height = "300" '编辑器的高度
				oFCKeditor.Value = "请输入工具使用教程文档" '这个是给编辑器初始值
				oFCKeditor.Create "ToolsHowTo" '以后编辑器里的内容都是由这个content 取得，
			%>
		</dd>

		<dd>
			<%
				
				Set oFCKeditor = New FCKeditor '创建一个编辑器的实例
				oFCKeditor.BasePath = "fckeditor/" '配置编辑器的路径，我站点根目录下的一个目录
				oFCKeditor.ToolbarSet = "Default" '完整和简化.Basic
				oFCKeditor.Width = "100%" '编辑器的长度
				oFCKeditor.Height = "300" '编辑器的高度
				oFCKeditor.Value = "请输入工具的Known Issue" '这个是给编辑器初始值
				oFCKeditor.Create "KnownIssue" '以后编辑器里的内容都是由这个content 取得，
			%>
			<%
				
				Set oFCKeditor = New FCKeditor '创建一个编辑器的实例
				oFCKeditor.BasePath = "fckeditor/" '配置编辑器的路径，我站点根目录下的一个目录
				oFCKeditor.ToolbarSet = "Default" '完整和简化.Basic
				oFCKeditor.Width = "100%" '编辑器的长度
				oFCKeditor.Height = "300" '编辑器的高度
				oFCKeditor.Value = "请输入Issue Escalation Log" '这个是给编辑器初始值
				oFCKeditor.Create "EscalationHistory" '以后编辑器里的内容都是由这个content 取得，
			%>


		</dd>

		<dd><input type="submit" onclick="return check();" name="send" value="添加内容" /></dd>
	</dl>
</form>

</body>
</html>