<%@codepage =936%>
<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->
<%

	
	dim tagName
	tagName = request.querystring("tag")
	
	if tagName = "" then
		call errorHistoryBack("�Ƿ�����")
	end if
	

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title> Back Office Mgmt System</title>
<link rel="stylesheet" type="text/css" href="style/basic.css" />
</head>
<body>
	


<!--#include file="header.asp"-->



<div id="SearchResult" >
	<h1>����ҳ</h1>
	<ul>
		<%
			dim i,tag,errorstr


			
			set rs = server.createobject("adodb.recordset")
			sql = "select * from CMS_Article where CMS_Tag like '%"&tagName&"%'"
			rs.open sql,conn,1,1
			i = 1
			
			if rs.eof then
				errorstr = "û�ҵ���ر�ǩ������"
			else

				rs.pagesize=40
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
end if 			
			
			for j=1 to rs.pagesize
			if rs.eof then exit for			

				tag = rs("CMS_tag")
				tag = replace(tag,tagName,"<span style='color:red'>"&tagName&"</span>")
				CMSID=rs("CMS_ID")

		%>
		<li>
			<dl>
				<dd class="title"><%=i%>. <%=rs("CMS_Title")%>&nbsp;&nbsp;&nbsp;<a style="font-size:12px; color:red;" href="detail.asp?ShowId=<%=CMSID%>">����Ķ�</a></dd>
				<dd class="info"><%=rs("CMS_Info")%><dd>
				<dd class="tag">TAG��ǩ��<%=tag%> �ؼ��֣�<%=rs("CMS_Keyword")%></dd>
				<dd class="tag">�����ˣ�<%=rs("CMS_Name")%> ����ʱ�䣺<%=rs("CMS_Date")%></dd>
			</dl>
		</li>
		<%
				i = i+1
				rs.movenext
			next
		%>
		
	</ul>
	<p style="text-align:center; margin-top:10px;"><%=errorstr%></p>
	<p style="text-align:center;padding:10px;">
    <%
	for i = 1 to rs.pagecount
		response.write "<a href='tag.asp?tag="&tagName&"&page="&i&"'>" & i & "</a> | "
	next
%>
    
    </p>
</div>




	
</body>
</html>
<%
	call close_rs
	call close_conn
%>