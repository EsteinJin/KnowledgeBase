<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../include/function.asp"-->
<!--#include file="conn.asp"-->
<%

	if session("Admin") = "" then
		call sussLoctionHref("Please login first!","admin_login.asp")
	end if
	
	dim showid
	showid = request.querystring("ShowId")
	
	'�ж�showid��Ч
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("error Occured!")
	end if
	
	'�ж�showid�����Ŀ�Ƿ����
	dim rs,sql,navid
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Nav where CMS_ID="&showid
	rs.open sql,conn,1,1
	
	if rs.eof then
		call close_rs
		call close_conn
		call errorHistoryBack("Not exsting Data")
	else
		'ȡ���������ID
		navid = rs("CMS_ID")
	end if
	
	call close_rs
	
	'ȥ��ѯ��������û������
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Nav where CMS_Sid="&navid
	rs.open sql,conn,1,1
	
	
	if not rs.eof then
		call close_rs
		call close_conn
		call errorHistoryBack("Already existing Sub category, Please delete Sub category First!")
	else
		dim delsql
		delsql = "delete from CMS_Nav where CMS_ID="&navid
		conn.execute(delsql)
		call close_rs
		call close_conn
		call sussLoctionHref("Successfully Deleted!","admin_nav.asp")
	end if
	
	
	call close_rs
	call close_conn
%>