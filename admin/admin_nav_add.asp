<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<!--#include file="../include/function.asp"-->
<%

	if session("Admin") = "" then
		call sussLoctionHref("Please Login First","admin_login.asp")
	end if
	'������ӵ���?
	dim sid
	sid = request.querystring("sid")	
%>

<!--#include file="conn.asp"-->
<%
'�����Ŀ
	if request.form("send") = "Add Menu" then
		dim nav,rs,sql
		nav = request.form("nav")
		sid = request.form("sid")
		
		if len(nav) < 2 or len(nav) > 40 then
			call errorHistoryBack("Not Less than 2 and more than 40")
		end if
		
		'����֮ǰ��֤�Ƿ��ظ�
		set rs = server.createobject("adodb.recordset")
		sql = "select * from CMS_Nav where CMS_NavName='"&nav&"'"
		rs.open sql,conn,1,1
		
		if not rs.eof then
			call errorHistoryBack("Already Exist")
			call close_rs
			call close_conn
		end if
		
		call close_rs  '�ر����ٱ�
		
		'����,������õ���SQL������
		'�������SQL��������ô����ID��ֵ�Ƚ�����
		'���ԣ����ǽ����ﻻ��recordset������
		'sql = "insert into CMS_Nav (CMS_NavName,CMS_Sid,CMS_Date) 
		'values ('"&nav&"',"&sid&",now())"
		'conn.execute(sql)
		
		'����recordset������1,3
		set rs = server.createobject("adodb.recordset")
		sql = "select * from CMS_Nav"
		rs.open sql,conn,1,3   '1,1��ʾֻ����1,3��ʾ��д
		
		rs.addnew
		rs("CMS_NavName") = nav
		rs("CMS_Sid") = sid
		rs("CMS_Date") = now()
		rs("CMS_Sort") = rs("CMS_ID") '��仰���ո����������ݵ�ID��ֵ��sort�ֶ�
		rs.update
		
		call close_rs '�ر�
		
		
		call close_conn '�ر��������ݿ�
		
		call sussLoctionHref("Successfully Added!","admin_nav.asp")
		
	end if
	
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>KB -BACK END</title>
<link rel="stylesheet" type="text/css" href="style/admin.css" />
</head>
<body style="margin:20px;font-size:14px;">

<h4 style="margin:10px 0;color:green;text-align:center;">Add Nav</h4>
<p style="margin:10px 0;"><a href="admin_nav.asp">Back to Nav Mgmt</a></p>


<%
	'dim navname
	'set rs = server.createobject("adodb.recordset")
	'sql = "select * from CMS_Nav where CMS_ID="&sid
	'rs.open sql,conn,1,1
	
	'if not rs.eof then
	'	navname = rs("CMS_NavName")
	'end if
	
	'call close_rs
	
	dim location,navname,sid2,navname2,id
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Nav where CMS_ID="&sid
	rs.open sql,conn,1,1
	
	if not rs.eof then
		navname = rs("CMS_NavName")
		sid2 = rs("CMS_Sid")
		location = navname
		
		if sid2 <> 0 then   '���������������£�ȥִ��
			
			'������ѭ�����������е��ϼ�����ѭ������
			do while true
				if sid2 = 0 then exit do  'ֻҪѭ�������ϼ������Ҳ��������
				dim rs2,sql2
				set rs2 = server.createobject("adodb.recordset")
				sql2 = "select * from CMS_Nav where CMS_ID="&sid2
				rs2.open sql2,conn,1,1
				
				if not rs2.eof then
					navname2 = rs2("CMS_NavName")
					sid2 = rs2("CMS_Sid")
					id = rs2("CMS_ID")
					location = "<a href='admin_nav_add.asp?sid="&id&"'>" & navname2 & "</a> >> " & location
				end if
				
				rs2.close
				set rs2 = nothing	
			
			loop
			
		end if
		
		
		
	end if
	
	call close_rs
%>

<p class="nav_h">
	<a href="admin_nav.asp">Top</a> >> <%=location%>
    [<a href="admin_nav_mof.asp?ShowId=<%=sid%>">Update</a>]
	[<a onclick="return confirm('Are you Sure?')" href="admin_nav_del.asp?ShowId=<%=sid%>">Delete</a>]
</p>



<p class="nav_h">
	
	<%
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Nav where CMS_Sid="&sid
	rs.open sql,conn,1,1
	
	if not rs.eof then
		do while not rs.eof
	
			response.write "<a href='admin_nav_add.asp?sid="&rs("CMS_ID")&"'>" & rs("CMS_NavName") & "</a> | "
			
			rs.movenext
		loop
	else
		response.write "No sub Nav"
	end if
	
	
	%>
	
</p>




<form id="form_nav_add" method="post" action="admin_nav_add.asp">
	<input type="hidden" value="<%=sid%>" name="sid" />
	Nav Name��<input type="text" name="nav" /> <input type="submit" name="send" value="Add Menu" />
</form>

</body>
</html>
<%
	call close_conn
%>