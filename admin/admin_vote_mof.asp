<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="../include/function.asp"-->
<%

	if session("Admin") = "" then
		call sussLoctionHref("非法操作","admin_login.asp")
	end if
	
	if request.form("send") = "修改" then
	dim voteId,sql2,voteName2
	voteId=request.Form("voteId")
	voteName2=request.Form("voteName")
	sql2="update CMS_Vote set CMS_VoteName='"&voteName2&"' where CMS_ID="&voteId
	conn.execute(sql2)
		'跳转
		call sussLoctionHref("修改成功，请查看投票项目与该内容是否符合！","admin_vote_name.asp")
	end if 

	dim showid
	showid = request.querystring("ShowId")
	
	'判断showid有效
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("error Occured!")
	end if	
    dim rs,sql
	set rs= server.createobject("adodb.recordset")
	sql = "select * from CMS_Vote where CMS_ID="&showid
	rs.open sql,conn,1,1
	if rs.eof then
		call close_rs
		call close_conn
		call errorHistoryBack("Not Existing Data!")
	else
		voteName = rs("CMS_VoteName")
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
	
	
	<form method="post" action="admin_vote_mof.asp">
		<dl id="voteadd">
			<dt>请添加一个投票的标题：</dt>
		<input type="hidden" name="voteId" value="<%=showid%>" />
		<p>[<a href="admin_vote_name_add.asp?sid=<%=showid%>">返回</a>]</p>            
			<dd><input type="text" name="voteName" value="<%=voteName%>" /> <input type="submit" name="send" value="修改" /></dd>
		</dl>
	</form>


</body>
</html>
<%

%>