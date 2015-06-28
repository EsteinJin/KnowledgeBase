<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="../include/function.asp"-->
<%

	if session("Admin") = "" then
		call sussLoctionHref("非法操作","admin_login.asp")
	end if
	
	if request.form("send") = "添加" then
		dim sql
		sql = "Insert into CMS_Vote (CMS_VoteName,CMS_Date) values ('"&request.form("votename")&"',now())"
		conn.execute(sql)
		call close_conn
		
		response.write "<script>alert('添加成功！');location.href='admin_vote_name.asp'</script>"
	end if
	
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Back Office Mgmt System--后台管理页面</title>
<link rel="stylesheet" type="text/css" href="style/admin.css" />
</head>
<body>
	
	
	<form method="post" action="admin_vote_name_add.asp">
		<dl id="voteadd">
			<dt>请添加一个投票的标题：</dt>
			<dd><input type="text" name="votename" /> <input type="submit" name="send" value="添加" /></dd>
		</dl>
	</form>


</body>
</html>