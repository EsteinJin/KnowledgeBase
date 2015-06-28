<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->
<%


	
	dim showid
	showid = request.querystring("ShowId")
	
	'判断showid有效
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("error Occured!")
	end if
	
	'判断showid这个栏目是否存在
	dim rs,sql,LinkId
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_LogRate where ID="&showid
	rs.open sql,conn,1,1
	
	if rs.eof then
		call close_rs
		call close_conn
		call errorHistoryBack("Not exsting Data")
	else
		'取出这个类别的ID
		CID = rs("ID")

	

		dim delsql
		delsql = "delete from CMS_LogRate where ID="&CID
		conn.execute(delsql)
		call close_rs
		call close_conn
		call sussLoctionHref("Successfully Deleted!","LogRateList.asp")

	end if	
	
	call close_rs
	call close_conn
%>