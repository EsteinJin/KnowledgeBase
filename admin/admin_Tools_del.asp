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
	sql = "select * from ToolsName where Item="&showid
	rs.open sql,conn,1,1
	
	if rs.eof then
		call close_rs
		call close_conn
		call errorHistoryBack("Not exsting Data")
	else
		'ȡ���������ID
		ItemId= rs("Item")
			dim delsql
		delsql = "delete from ToolsName where Item="&ItemId
		conn.execute(delsql)
		call close_rs
		call close_conn
		call sussLoctionHref("Successfully Deleted!","admin_Tools_List.asp")
	end if
	
	call close_rs
	call close_conn
%>