<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%

	if session("Admin") = "" then
		call sussLoctionHref("�Ƿ�����","admin_login.asp")
	end if
%>
<!--#include file="../include/function.asp"-->
<!--#include file="conn.asp"-->
<%
	dim rs,sql
	
	if request.querystring("first") = "ok" then
		'��һ�֣��Ȳ���Ŀǰ��ѡ��Ȼ���ɷ��ٽ������ı����
		'�ڶ��֣������еĶ���ɷ�Ȼ���ٽ������ı����
		
		'�����佫���еĸĳ�0,sql�ǿ���������ģ�û�����κε�whereɸѡ����������
		sql = "Update CMS_Vote Set CMS_Level=0"
		conn.execute(sql)
		
		'�ٽ���Ҫѡ����Ǹ�����ĳ���ѡ
		sql = "Update CMS_Vote Set CMS_Level=1 where CMS_ID="&request.querystring("ShowId")
		conn.execute(sql)
		
	end if
	
	
	
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Vote where CMS_VoteSid=0 order by CMS_Date desc"
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
		<tr><th>���</th><th>����</th><th>�Ƿ���ҳ��ʾ</th><th>����</th></tr>
		<%
			do while not rs.eof
				dim levelstr
				if rs("CMS_Level") = 0 then
					levelstr = "��"
				elseif rs("CMS_Level") = 1 then
					levelstr = "<span style='color:red;font-weight:bold;'>��</span>"
				end if
		%>
		<tr><td class="id"><%=rs("CMS_ID")%></td><td><a href="admin_vote_x.asp?ShowId=<%=rs("CMS_ID")%>"><%=rs("CMS_VoteName")%></a></td><td><%=levelstr%></td><td><a href="admin_vote_name.asp?first=ok&ShowId=<%=rs("CMS_ID")%>">ȷ����ѡ</a> | <a href="admin_vote_mof.asp?ShowId=<%=rs("CMS_ID")%>">�޸�</a> | <a onclick="return confirm('��ȷ������ɾ����')" href="admin_vote_del.asp?del=ok&ShowId=<%=rs("CMS_ID")%>">ɾ��</a></td></tr>
        
        
		<%
				rs.movenext
			loop
		%>
	</table>

	<p style="text-align:center;"><a href="admin_vote_name_add.asp">[�������]</a></p>
	
</body>
</html>
<%
	call close_rs
	call close_conn
%>