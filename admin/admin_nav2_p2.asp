<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../include/function.asp"-->
<!--#include file="conn.asp"-->
<%

	if session("Admin") = "" then
		call sussLoctionHref("�Ƿ���¼","admin_login.asp")
	end if
	
	'ȷ����ѡ֮ǰ���ж�һ�£��Ƿ��Ѿ�����6��
	dim countrs,countsql
	set countrs = server.createobject("adodb.recordset")
	countsql = "select * from CMS_Nav where CMS_Level=true"
	countrs.open countsql,conn,1,1
	
	if countrs.recordcount >=6 then
		countrs.close
		set countrs = nothing
		call errorHistoryBack("��ѡ��Ŀ�Ѿ��ﵽ�������6��\n��ȡ����������Ŀ������ȷ������Ŀ")
		response.end
	end if
	
	countrs.close
	set countrs = nothing
	
	showid = request.querystring("ShowId")
	'�Ƿ�����
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("�Ƿ�����")
	end if
	
	dim rs,sql,rs2,sql2,count2
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Nav where CMS_ID="&showid
	rs.open sql,conn,1,1
	
	if rs.eof then
		call close_rs
		call close_conn
		call errorHistoryBack("�����������Ŀ")
	else 		
		sql = "update CMS_Nav Set CMS_Level=true where CMS_ID="&showid
		conn.execute(sql)
		call close_rs
		call close_conn
		call sussLoctionHref(count2&"ȷ������ѡ","admin_nav2.asp")
	end if
%>

