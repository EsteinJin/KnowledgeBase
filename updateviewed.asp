<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->
<%
dim showid

	showid = request.querystring("ShowId")
	'非法操作
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("error occured!")
	end if
dim title,content
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Article where CMS_ID="&showid
	rs.open sql,conn,1,1
	
	if rs.eof then
		call errorHistoryBack("Data Not Exist!")
	
		end if 
		
		info = rs("CMS_Info")
		tag = rs("CMS_Tag")
		keyword = rs("CMS_Keyword")
		name = rs("CMS_Name")
		fdate = rs("CMS_Date")
		viewed=cint(rs("CMS_Viewed"))+1
		updatesql="update CMS_Article set CMS_Viewed="&viewed&" where CMS_ID="&showid
		conn.execute(updatesql)

	
	call close_rs
%>