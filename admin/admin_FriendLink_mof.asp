<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="../include/function.asp"-->
<%

	if session("Admin") = "" then
		call sussLoctionHref("�Ƿ�����","admin_login.asp")
	end if

	
	
	
	if request.form("send") = "�޸�" then
	dim LinkId,sql2,LinkName,LinkAddress,LinkInfo
	LinkId=request.Form("LinkId")
	LinkName=request.Form("LinkName")
	LinkAddress=request.Form("LinkAddress")
	LinkInfo=request.Form("LinkInfo")
	sql2="update FriendLink set LinkName='"&LinkName&"',LinkAddress='"&LinkAddress&"',LinkInfo='"&LinkInfo&"' where ID="&LinkId
	conn.execute(sql2)
		'��ת
		call sussLoctionHref("�޸ĳɹ���","admin_FriendLink.asp")
	end if 

	dim showid
	showid = request.querystring("ShowId")
	'�ж�showid��Ч
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("error Occured!")
	end if		

    dim rs,sql
	set rs= server.createobject("adodb.recordset")
	sql = "select * from FriendLink where ID="&showid
	rs.open sql,conn,1,1
	if rs.eof then
		call close_rs
		call close_conn
		call errorHistoryBack("Not Existing Data!")
	else
		LinkName=rs("LinkName")
		LinkAddress=rs("LinkAddress")
		LinkInfo=rs("LinkInfo")
		
	end if	
	
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Back Office Mgmt System--��̨����ҳ��</title>
<link rel="stylesheet" type="text/css" href="style/admin.css" />
</head>
<body>
	
<form name="add" id="articleadd" method="post" action="admin_FriendLink_mof.asp">
	<dl>
		<dt>�뷢������</dt>
		<dd>��������:&nbsp;
        <input type="text" name="LinkName" class="text" value="<%=LinkName%>" />
		<input type="hidden" name="LinkId" value="<%=showid%>" />                 
		</dd>
		<dd>���ӵ�ַ:&nbsp;
                <input type="text" name="LinkAddress" class="text"  value="<%=LinkAddress%>" />
        </dd>
		<dd>������Ϣ��
                <textarea rows="2" name="LinkInfo" ><%=LinkInfo%></textarea>
		</dd>
		<dd><input type="submit" onclick="return check();" name="send" value="�޸�" /></dd>
	</dl>
</form>

</body>
</html>
<%

%>