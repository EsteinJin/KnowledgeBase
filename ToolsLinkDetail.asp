<%@codepage =936%>
<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->
<%
	dim showid

	showid = request.querystring("ShowId")
	'�Ƿ�����
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("�Ƿ�����")
	end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Back Office Mgmt System</title>
<link rel="stylesheet" type="text/css" href="style/basic.css" />
</head>
<body>

<!--#include file="header.asp"-->
<%

	set rs = server.createobject("adodb.recordset")
	sql = "select * from ToolsName where Item="&showid
	rs.open sql,conn,1,1
	
	if rs.eof then
	call errorHistoryBack("Not Exist Category")
	else
	
%>

<div id="detail">
	<h3> </h3>
	<p class="d">������Ŀ��<%=rs("ProjectName")%> | ������˾��<%=rs("ToolsCategory")%> | �������ƣ�<%=rs("ToolsName")%>  | �������ӣ�<a href="<%=rs("ToolsLink")%>">��� </a> </p>
	<p><%=rs("ToolsHowTo")%> </p>
    <p><%=rs("KnownIssue")%> </p>
    <p><%=rs("EscalationHistory")%> </p>	 
</div>
	
</body>
</html>

<%
end if 
	call close_rs
	call close_conn

%>
