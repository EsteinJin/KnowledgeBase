<%@codepage = 936%>
<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" type="text/css" href="style/basic.css" />
<script type="text/javascript" src="Script/jquery-latest.js"></script>
<script type="text/javascript" src="Script/thickbox.js"></script>
<link rel="stylesheet" href="Common/thickbox.css" type="text/css" media="screen" />
<title>Back Office Mgmt System</title>
</head>
<!--#include file="header.asp"-->
<div id="clistmyIndexList">
  <div id="cListmyIndexLeft">
    <h1>列表页</h1>
    <ul>
      <%
		
		set rs = server.createobject("adodb.recordset")
		sql = "select  * from CMS_Complaint order by CMS_MonitorDate desc"
		rs.open sql,conn,1,1
		if not rs.eof then
		rs.pagesize=30	
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
			Summary = rs("CMS_TicketSummary")
			if len(Summary) > 100 then
			Summary = left(Summary,100)
			Summary = Summary & "..."
			end if				    
	
	  %>
      <li><em>[<%=rs("CMS_MonitorDate")%>]</em> [<%=rs("CMS_CompliantSource")%>]--[<%=rs("CMS_AgentName")%>]-- <a href="FeedbackDetail.asp?ShowId=<%=rs("CMS_ID")%>"><span><%=Summary%></span></a></li>
      <%
				rs.movenext
			next	
end if 

		%>
    </ul>
    <p style="text-align:center;padding:10px;">
<p style="text-align:center;padding:10px;">
    <%
	for i = 1 to rs.pagecount
		response.write "<a href='QCFeedback.asp?page="&i&"'>" & i & "</a> | "
	next
%>
    
    </p>
    </p>
    
    
  </div>
  <div id="clistmyIndexRight" >
    <!--#include file="myRight.asp"-->
</div>
</body></html>
