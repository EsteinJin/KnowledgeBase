<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../include/function.asp"-->
<!--#include file="conn.asp"-->
<%

	if session("Admin") = "" then
		call sussLoctionHref("�Ƿ�����","admin_login.asp")
	end if
	
		dim rs,sql
	set rs = server.createobject("adodb.recordset")
	sql = "SELECT * FROM EscalationLog order by EscalatedDate desc"
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
		delsql = "select * from EscalationLog where ID="&showid
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
			delsql2 = "delete from EscalationLog where ID="&showid
			conn.execute(delsql2)
			call sussLoctionHref("ɾ���ɹ�","admin_Escalation.asp")
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
<p style=" font-size:12px; text-align:center; margin-top:10px;"><a href="admin_Escalation_add.asp">�������</a>�������</p>     
<table border="1" id="listcontent">
<tr><th>������Ա</th><th>��������</th><th>����״��</th><th>������Ա</th><th>��������</th><th>����</th></tr>
<%
for i=1 to rs.pagesize
if rs.eof then exit for	
Summary = rs("IssueSummary")
if len(Summary) > 35 then
Summary = left(Summary,35)
Summary = Summary & "..."
end if
%>
<tr><td><%=rs("EscalatedBy")%></td><td><%=rs("EscalationType")%></td><td><%=rs("StatusTrack")%></td><td><%=rs("ResponsibleBy")%></td><td><%=Summary%></td><td class="d"><a href="admin_Escalation_mof.asp?ShowId=<%=rs("ID")%>">�޸�</a> | <a  onclick="return confirm('��ȷ������ɾ����')" href="admin_Escalation.asp?del=ok&ShowId=<%=rs("ID")%>">ɾ��</a></td></tr>



<%
rs.movenext
next	
%> 

</table>
<p style="text-align:center;padding:10px;">
<%
	for i = 1 to rs.pagecount
		response.write "<a href='admin_Escalation.asp?page="&i&"'>" & i & "</a> | "
	next
%>
   
     </p>

     
</body>
</html>
<%
	call close_rs
	call close_conn
%>
