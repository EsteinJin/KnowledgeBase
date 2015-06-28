<%@codepage =936%>
<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" href="style/basic.css" />
<title>Back Office Mgmt System</title>
</head>
<body>



<div id="myIndexList" style=" overflow:scroll;border:1px solid #0000CC; height:400px; ">

<%
	'非法操作
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_LogRate_Store  order by CMS_Date desc "
	rs.open sql,conn,1,1
	
	if rs.eof then
	call errorHistoryBack("Not Exist List")
	else
	
%>

<h1>Daily LogRate List<span ><a style=" margin-left:850px; color:white" href="LogRateAdd.asp">Click to Log Rate</a></span></h1>
<table style="overflow:scroll;">
<th style="width:100px;">Log Date</th><th style=" width:100px;">Ticket Number</th><th style=" width:100px;">Agent Name</th><th style=" width:100px;">Langguage</th><th style=" width:100px;">Is Remoted?</th><th style=" width:100px;">Ticket Compliance</th><th style=" width:100px;">Handled Time</th><th style=" width:100px;">Category</th><th style=" width:100px;">Type</th><th style=" width:100px;">Item</th><th style=" width:50px;">CallType</th>
<%
do while not rs.eof
%>
<tr style="border:1px solid maroon">
<td><%=rs("CMS_Date")%></td><td><%=rs("CMS_TicketNumber")%></td><td><%=rs("CMS_AgentName")%></td>
<td><%=rs("CMS_Language")%></td><td><%=rs("CMS_Remote")%></td><td><%=rs("CMS_Compliance")%></td><td><%=rs("CMS_HandleTime")%></td>
<td><%=rs("CMS_Category")%></td><td><%=rs("CMS_Type")%></td><td><%=rs("CMS_Item")%></td><td><%=rs("CMS_CallType")%></td><td class="d"></td>
</tr>


<%
   rs.movenext
   loop		
	end if
%>
</table>	

</div>



</body>
</html>
<%
	call close_rs
	call close_conn
%>