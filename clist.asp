<%@codepage =936%>
<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->
<%
	dim showid
	


	showid = request.querystring("ShowId")
	'�Ƿ�����
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("error Occured!")
	end if
	
	dim title
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Nav where CMS_ID="&showid
	rs.open sql,conn,1,1

	if rs.eof then
		call errorHistoryBack("�����ڴ���Ŀ")
	else 
		name = rs("CMS_NavName")
	end if
	
	call close_rs
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" href="style/basic.css" />
<title>Back Office Mgmt System</title>
</head>
<body>

<!--#include file="header.asp"-->


<div id="clistmyIndexList">
<div id="cListmyIndexLeft">
	<h1><%=name%></h1>
	<ul>
		<%
			dim color,bold,rmp,pic
			set rs = server.createobject("adodb.recordset")
            sql = "select  * from CMS_Article where  CMS_Sort="&showid&" order by CMS_Date desc"
			rs.open sql,conn,1,1
			if not rs.eof then
			rs.pagesize=50	
	'��������ҳ��
	'���յ���ֵΪ�ַ���������ת���������Ƚ�
	'cint(����)�����ԱȽ���
	'���ж��Ƿ�Ϊ�ַ���������ǵĻ�����rs.absolutepage = 1
	'������ǣ����ж��Ƿ�Ϊ�գ��Ƿ�ΪС��1����������ҳ��
	
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
			
					
			for i=1 to rs.pagesize
			if rs.eof then exit for			
			
			
				title = rs("CMS_Title")
				if len(title) > 100 then
					title = left(title,100)
					title = title & "..."
				end if
				
				color = rs("CMS_Color")
				
				if rs("CMS_Bold") = 1 then
					bold = " bold"
				else
					bold = ""
				end if
				
				if rs("CMS_Rmp") = 1 then
					rmp = "<span style='color:blue'>[�Ƽ�]</span>"
				else
					rmp = ""
				end if
				
				if rs("CMS_Pic") = 1 then
					pic = "<img src='image/pic.gif' alt='��ͼ' />"
				else
					pic = ""
				end if
				
		%>

		<li><em>[<%=FormatDateTime(rs("CMS_Date"),2)%>]</em> ��<%=rmp%><%=pic%> <a href="detail.asp?ShowId=<%=rs("CMS_ID")%>"><span class="<%=color%> <%=bold%>"><%=title%></span></a></li>
		<%
				rs.movenext
			next	
end if 			
		%>
	</ul>
	<p style="text-align:center;padding:10px;">
    <%
	for i = 1 to rs.pagecount
		response.write "<a href='clist.asp?ShowId="&showid&"&page="&i&"'>" & i & "</a> | "
	next
%>
    
    </p>
</div>


<div id="clistmyIndexRight" >
<!--#include file="myRight.asp"-->
</div>

</body>
</html>
