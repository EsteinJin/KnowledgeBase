<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../include/function.asp"-->
<!--#include file="conn.asp"-->
<%

	if session("Admin") = "" then
		call sussLoctionHref("�Ƿ�����","admin_login.asp")
	end if

	dim showid
	showid = request.querystring("ShowId")
	'�ж�showid��Ч
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("�Ƿ�����")
	end if
	
	'�ж�showid�����Ŀ�Ƿ����
	dim rs,sql,votename
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Vote where CMS_ID="&showid
	rs.open sql,conn,1,1
	
	if rs.eof then
		call close_rs
		call close_conn
		call errorHistoryBack("�����ڴ˱���")
	else
		'������
		votename = rs("CMS_VoteName")
	end if
	
	call close_rs
	
	'��ʼ������Ŀ
	if request.form("send") = "����" then  '�жϰ�ť�Ƿ���
		dim votex,xrs,xsql
		votex = request.form("votex")
		
		if len(votex) < 2 then
			call errorHistoryBack("��Ŀ����������2λ��")
		end if
		
		
		set xrs = server.createobject("adodb.recordset")
		xsql = "select * from CMS_Vote where CMS_VoteSid="&showid
		xrs.open xsql,conn,1,1
		'��Ŀ����,Ҫ�����жϣ������4����Ŀ���ˣ��Ͳ�����������
		
		'����һ���������ж�
		
		'recordcount
		if xrs.recordcount >=10 then
			call errorHistoryBack("��Ŀ�������Ѿ��ⶥ")
		else
			sql = "Insert into CMS_Vote (CMS_VoteName,CMS_VoteSid,CMS_Date) values ('"&votex&"',"&showid&",now())"
			conn.execute(sql)
		end if
		
		xrs.close
		set xrs = nothing
		
		'��Ϊ����������ȡ����֮����ɵģ����ԣ�Ҫ��תˢ��һ�£����ܵõ�����
		response.redirect "admin_vote_x.asp?ShowId="&showid
		
	end if
	
	
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Vote where CMS_VoteSid="&showid
	rs.open sql,conn,1,1
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Back Office Mgmt System--��̨����ҳ��</title>
<link rel="stylesheet" type="text/css" href="style/admin.css" />
</head>
<body>

	<table id="votename" cellspacing="1">
		<tr><td colspan="4" class="title"><%=votename%></td></tr>
		<tr><th>���</th><th>��Ŀ����</th><th>��Ʊ��</th><th>����</th></tr>
		<%
			do while not rs.eof
		%>
		<tr><td><%=rs("CMS_ID")%></td><td><%=rs("CMS_VoteName")%></td><td><%=rs("CMS_VoteCount")%></td><td><a onclick="return confirm('��ȷ������ɾ����')" href="admin_vote_x_del.asp?del=ok&ShowId=<%=rs("CMS_ID")%>">ɾ��</a> <a href="admin_vote_x_mof.asp?ShowId=<%=rs("CMS_ID")%>">�޸�</a></td></tr>
 
		<%
				rs.movenext
			loop
		%>
	</table>
	<form method="post" action="admin_vote_x.asp?ShowId=<%=showid%>">
		<dl style="width:250px;margin:auto;">
			<dt>��������Ŀ��</dt>
			<dd>��Ŀ����<input type="text" name="votex" /> <input type="submit" value="����" name="send" /></dd>
		</dl>
	</form>

</body>
</html>
<%
		call close_rs
		call close_conn
%>