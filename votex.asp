<%@codepage =936%>
<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->
<%
	dim rs,sql,votetitle,votesid
	'��ȡ�ı���
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Vote where CMS_Level=1"
	rs.open sql,conn,1,1
	
	if not rs.eof then
		votetitle = rs("CMS_VoteName")
		votesid = rs("CMS_ID")
	end if
	
	call close_rs
	
	'��ȡ��Ŀ��
	dim countsum,countavg,countavg2,i
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Vote where CMS_VoteSid="&votesid
	rs.open sql,conn,1,1
	
	do while not rs.eof 
		countsum = countsum + rs("CMS_VoteCount")
		rs.movenext
	loop
	
	'��ָ�뷵�ص���һ��λ����
	rs.movefirst
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Back Office Mgmt System</title>
<link rel="stylesheet" type="text/css" href="style/basic.css" />
</head>
<body>
	
	<h1 class="votex"><%=votetitle%></h1>
	<table id="votex" border="1">
		<tr><th>��Ŀ��</th><th>��״����</th><th>��Ʊ��</th><th>�ٷֱ�</th></tr>
		<%
			i = 1
			do while not rs.eof
				if countsum <> 0 then
				countavg = rs("CMS_VoteCount")/countsum*100
				countavg2 = int(rs("CMS_VoteCount")/countsum*100)
				end if 
		%>
		<tr><td class="name"><%=rs("CMS_VoteName")%></td><td><img src="image/b<%=i%>.jpg" width="<%=countavg2*3%>" height="21" alt="�ٷֱ�" /></td><td><%=rs("CMS_VoteCount")%></td><td><%=FormatNumber(countavg,2)%>%</td></tr>
		<%
				rs.movenext
				i = i+1
			loop
		%>
		<tr><td class="name" colspan="4">һ��Ͷ�ˣ�<strong><%=countsum%></strong>Ʊ</td></tr>
	</table>
	<input type="button" onclick="javascript:window.close();" value="�ر�" style="margin-top:10px;margin-left:240px;" />

</body>
</html>
<%
		call close_rs
		call close_conn
%>