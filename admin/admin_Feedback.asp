<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../include/function.asp"-->
<!--#include file="conn.asp"-->
<%

	if session("Admin") = "" then
		call sussLoctionHref("�Ƿ�����","admin_login.asp")
	end if
	
		dim rs,sql
	set rs = server.createobject("adodb.recordset")
	sql = "SELECT * FROM CMS_Complaint order by CMS_MonitorDate desc"
	rs.open sql,conn,1,1
	rs.pagesize=2
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
		delsql = "select * from CMS_Complaint where CMS_ID="&showid
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
			delsql2 = "delete from CMS_Complaint where CMS_ID="&showid
			conn.execute(delsql2)
			call sussLoctionHref("ɾ���ɹ�","admin_Feedback.asp")
		end if
		
		delrs.close
		set delrs = nothing
		
	end if

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" href="style/admin.css" />
<title>�ޱ����ĵ�</title>
</head>

<body>
<p style=" font-size:12px; text-align:center; margin-top:10px;"><a href="admin_Feedback_add.asp">�������</a>�������</p>     
<table border="1" id="listcontent">
<tr><th>�������</th><th>Ͷ����Դ</th><th>׷��״��</th><th>Agent����</th><th>��������</th><th>����</th></tr>
<%
for i=1 to rs.pagesize
if rs.eof then exit for	
Summary = rs("CMS_TicketSummary")
if len(Summary) > 35 then
Summary = left(Summary,35)
Summary = Summary & "..."
end if
%>
<tr><td><%=rs("CMS_MonitorDate")%></td><td><%=rs("CMS_CompliantSource")%></td><td><%=rs("CMS_HandleStatus")%></td><td><%=rs("CMS_AgentName")%></td><td><%=Summary%></td><td class="d"><a href="admin_Feedback_mof.asp?ShowId=<%=rs("CMS_ID")%>">�޸�</a> | <a  onclick="return confirm('��ȷ������ɾ����')" href="admin_Feedback.asp?del=ok&ShowId=<%=rs("CMS_ID")%>">ɾ��</a></td></tr>



<%
rs.movenext
next	
%> 

</table>
<p style="text-align:center;padding:10px;">
<%
	for i = 1 to rs.pagecount
		response.write "<a href='admin_Feedback.asp?page="&i&"'>" & i & "</a> | "
	next
%>
   
     </p>

     
</body>
</html>
<%
	call close_rs
	call close_conn
%>
