<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../include/function.asp"-->
<%

	if session("Admin") = "" then
		call sussLoctionHref("Please Login First","admin_login.asp")
	end if
%>
<!--#include file="conn.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=GB2312" />
<title>KB -BACK END</title>
<link rel="stylesheet" type="text/css" href="style/admin.css" />
</head>
<body style="margin:20px;font-size:14px;">

<h4 style="margin:10px 0;color:green;text-align:center;">NAV MGMT</h4>

<div id="nav_show">
	<%
		dim rs,sql,rs2,sql2
		set rs = server.createobject("adodb.recordset")
		sql = "select * from CMS_Nav where CMS_Sid=0 order by CMS_Sort asc"
		rs.open sql,conn,1,1
		
		
		'ѭ����Ŀ
		do while not rs.eof
	%>
	<dl>
		<dt><a href="admin_nav_sort.asp?ShowId=<%=rs("CMS_ID")%>&aa=left">��</a> <a href="admin_nav_add.asp?sid=<%=rs("CMS_ID")%>"><%=rs("CMS_NavName")%></a> <a href="admin_nav_sort.asp?ShowId=<%=rs("CMS_ID")%>&aa=right">��</a></dt>
		<dd>
        
				<%
					set rs2 = server.createobject("adodb.recordset")
					sql2 = "select * from CMS_Nav where CMS_Sid="&rs("CMS_ID")
					rs2.open sql2,conn,1,1
					do while not rs2.eof
				
					response.write "<a href='admin_nav_add.asp?sid="&rs2("CMS_ID")&"'>" & rs2("CMS_NavName") & "</a> "
				
						rs2.movenext
					loop
				%>
		</dd>
	</dl>
	<%
			rs.movenext
		loop
	%>
	
</div>


<div id="nav_add">
	[<a href="admin_nav_add.asp?sid=0">Add Main Nav</a>]
</div>

</body>
</html>