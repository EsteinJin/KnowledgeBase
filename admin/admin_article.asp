<%@codepage = 936%>
<!--#include file="../include/function.asp"-->
<!--#include file="conn.asp"-->
<%

	if session("Admin") = "" then
		call sussLoctionHref("�Ƿ�����","admin_login.asp")
	end if



	dim rs,sql,title,i
	set rs = server.createobject("adodb.recordset")
	sql = "SELECT CMS_Article.CMS_ID as CMS_ID, CMS_Article.CMS_Title as CMS_Title,CMS_Nav.CMS_NavName as CMS_NavName,CMS_Article.CMS_Name as CMS_Name,CMS_Article.CMS_Date as CMS_Date  FROM CMS_Article INNER JOIN CMS_Nav ON CMS_Article.CMS_Sort = CMS_Nav.CMS_ID  order by CMS_NavName"
	rs.open sql,conn,1,1
	rs.pagesize=10
	if isnumeric(request.querystring("page")) then
		if request.querystring("page") = "" or cint(request.querystring("page"))<1 then
			rs.absolutepage = 1
		elseif cint(request.querystring("page"))>rs.pagecount then
			rs.absolutepage = rs.pagecount
		else
			rs.absolutepage = request.querystring("page")
		end if
	else
		rs.absolutepage = 1
	end if 
	
	'ɾ��ģ��
	if request.querystring("del")="ok" then
		dim showid,delrs,delsql,delsql2
		
		showid = request.querystring("ShowId")
		'�Ƿ�����
		if showid = "" or not isnumeric(showid) then
			call errorHistoryBack("�Ƿ�����")
		end if
		
		'�ж������Ƿ����
		set delrs = server.createobject("adodb.recordset")
		delsql = "select * from CMS_Article where CMS_ID="&showid
		delrs.open delsql,conn,1,1
		
		'������ݲ�����
		if delrs.eof then
			call close_rs
			call close_conn
			delrs.close
			set delrs = nothing
			call errorHistoryBack("��Ҫɾ�������ݲ�����")
		else
		    call ConfirmDel()
			'ִ��ɾ������
			delsql2 = "delete from CMS_Article where CMS_ID="&showid
			conn.execute(delsql2)
			call sussLoctionHref("ɾ���ɹ�","admin_article.asp")
		end if
		
		delrs.close
		set delrs = nothing
		
	end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��̨����</title>
<link rel="stylesheet" type="text/css" href="style/admin.css" />
</head>
<body>
	

<table border="1" id="listcontent">
	<tr><th>���</th><th>����</th><th>�ĵ�����</th><th>������</th><th>����ʱ��</th><th>����</th></tr>
	<%
		for i=1 to rs.pagesize
			if rs.eof then exit for	
			title = rs("CMS_Title")
			if len(title) >40 then
				title = left(title,40)
				title = title & "..."
			end if
	%>
	<tr><td><%=rs("CMS_ID")%></td><td><%=title%></td><td><%=rs("CMS_NavName")%></td><td><%=rs("CMS_Name")%></td><td><%=rs("CMS_Date")%></td><td class="d"><a href="admin_article_mof.asp?ShowId=<%=rs("CMS_ID")%>">�޸�</a> | <a  onclick="return confirm('��ȷ������ɾ����')" href="admin_article.asp?del=ok&ShowId=<%=rs("CMS_ID")%>">ɾ��</a></td></tr>
	<%
	rs.movenext
	next	
	%>
</table>

	<p style="text-align:center;padding:10px;">
    <%
	for i = 1 to rs.pagecount
		response.write "<a href='admin_article.asp?page="&i&"'>" & i & "</a> | "
	next
%>
    
    </p>
	
	
</body>
</html>
<%
	call close_rs
	call close_conn
%>