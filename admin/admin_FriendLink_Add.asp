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
    dim LinkName,LinkAddress,LinkInfo,rs
		LinkName = request.form("LinkName")
		LinkAddress = request.form("LinkAddress")
		LinkInfo = request.form("LinkInfo")
		
		if len(LinkName) < 2 or len(LinkName) > 100 then
			call errorHistoryBack("友情链接不小于2位，或者大于100位")
		end if	
		if len(LinkAddress) < 2  then
			call errorHistoryBack("不能为空！")
		end if				

		if len(LinkInfo) < 2  then
			call errorHistoryBack("不能为空")
		end if				


		'新增数据,发布成功后跳转到内容管理页面
		addsql = "Insert into FriendLink(LinkName,LinkAddress,LinkInfo) values ('"&LinkName&"','"&LinkAddress&"','"&LinkInfo&"')"
		conn.execute(addsql)
		call sussLoctionHref("内容新增成功","admin_FriendLink.asp")




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

<form name="add" id="articleadd" method="post" action="admin_FriendLink_Add.asp">
	<dl>
		<dt>请发布文章</dt>
		<dd>链接名称:&nbsp;
                <input type="text" name="LinkName" class="text" /> 
		</dd>
		<dd>链接地址:&nbsp;
                <input type="text" name="LinkAddress" class="text" />
        </dd>
		<dd>链接信息：
                <textarea rows="2" name="LinkInfo" ></textarea>
		</dd>
		<dd><input type="submit" onclick="return check();" name="send" value="添加内容" /></dd>
	</dl>
</form>

</body>
</html>