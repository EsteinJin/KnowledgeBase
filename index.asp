<%@codepage = 65001%>
<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>

<!--#include file="navleft.asp"-->

   
    <td valign="top"  id="IndexBody"><table border="0" width="100%" height="100%" cellspacing="5" cellpadding="0">
      <tr>
        <td valign="top"><img src="image/header-grey.jpg" border="0" width="600" height="100" vspace="0" hspace="0" alt=""><br>
          <br>
          <div id="win">

          <div id="">
          <p>
    <br>
    <%
	 set artrs=server.CreateObject("adodb.recordset")
	 artsql="select *  from CMS_Article"
	 artrs.open artsql,conn,1,1
	 if not artrs.eof then
	 artcount=artrs.recordcount
	 do while not artrs.eof 
	 artviewed=artviewed+artrs("CMS_Viewed")
	 artrs.movenext
	 loop
	 end if 
	
	%>
    <span style="color:red;">Total Article:<%=artcount%> vs Current Viewed:<%=artviewed%></span><br><br>
           

 <%

	
	'提取项目名

	set poprs = server.createobject("adodb.recordset")
	popsql = "select top 5 * from CMS_Article where CMS_Viewed >2 order by CMS_Viewed desc"
	poprs.open popsql,conn,1,1
	
	do while not poprs.eof 
		countsum = countsum + poprs("CMS_Viewed")
		poprs.movenext
	loop
	
	'将指针返回到第一个位置上
	if not poprs.eof then 
	poprs.movefirst
	end if 
%>

	
	<h1 class="votex" style="font-size:12px; color:red; font-weight:bold;">Populer Articles</h1>
	<TABLE class=search2 border=0 cellSpacing=1 cellPadding=2 
                  width="100%">
		<tr><th>Title</th><th>Graph</th><th>Count</th><th>Percent</th></tr>
		<%
			i = 1
			do while not poprs.eof
				if countsum <> 0 then
				countavg = poprs("CMS_Viewed")/countsum*100
				countavg2 = int(poprs("CMS_Viewed")/countsum*100)
				end if 
		%>
		 <TR style="BACKGROUND-COLOR: #e9e9e9" class=result-title 
                    onmouseover='this.style.background = "#AAAAAA"' 
                    onmouseout='this.style.background = "#E9E9E9"'>
        
        <td class="name"><a href="detail.asp?ShowId=<%=poprs("CMS_ID")%>"><%=poprs("CMS_Title")%></a></td><td><img src="image/b<%=i%>.jpg" width="<%=countavg2*3%>" height="21" alt="Pop Rate" /></td><td><%=poprs("CMS_Viewed")%></td><td><%=FormatNumber(countavg,2)%>%</td></tr>
		<%
				poprs.movenext
				i = i+1
			loop
		%>
		<tr></tr>
	</table>  
            
	<h1 class="votex" style="font-size:12px; color:red; font-weight:bold;">Top 5 Articles</h1>
               <TABLE class=search2 border=0 cellSpacing=1 cellPadding=2 
                  width="100%">
                  <TBODY>
                    <TR>
                      <TD class=result-header width=15>No</TD>
                      <TD class=result-header>Title</TD>
                      <TD class=result-header>Log Date</TD>
                      <TD class=result-header>Article Tag</TD>

                      <TD class=result-header width=35>Created By</TD>
                    </TR>
                    <%

			set rs = server.createobject("adodb.recordset")
            sql = "select  top 5 * from CMS_Article where CMS_Top=1 order by CMS_Date desc"
			rs.open sql,conn,1,1
			if rs.eof then 
			response.Write("<tr><td colspan=5 style=""color:green;"">No Top Articles were found!</td></tr>")
			end if 
			if not rs.eof then
			
			do while not rs.eof 
				title = rs("CMS_Title")
				cms_info=rs("CMS_Info")
				if len(cms_info)>50 then
					cms_info = left(cms_info,50)
					cms_info = cms_info & "..."
				end if 
				
			
			%>
                    <TR style="BACKGROUND-COLOR:#FFFF99" class=result-title 
                    onmouseover='this.style.background = "#AAAAAA"' 
                    onmouseout='this.style.background = "#FFFF99"'>
                      <TD align=middle><%=i%>.</TD>
                      <TD><B><A class=site-nav 
                        href=""><U><a href="detail.asp?ShowId=<%=rs("CMS_ID")%>"><span class="<%=color%> <%=bold%>"><%=title%></span></a></U></A></B></TD>
                      <TD>[<%=FormatDateTime(rs("CMS_Date"),2)%>]</TD>
                      <TD><%=rs("CMS_Tag")%></TD>

                      <TD align=middle><%=rs("CMS_Name")%></TD>
                    </TR>
                    <TR style="BACKGROUND-COLOR: #e9e9e9">
                      <TD style="BACKGROUND-COLOR: #fff">&nbsp;</TD>
                      <TD colSpan=6><%=rs("CMS_Info")%></TD>
                      <%
				rs.movenext
			loop
			%>
            <a style="float:right; color:red; font-weight:bold;" href="TopArticleList.asp">View More</a>
            <%	
end if 			
		%>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE border=0 cellSpacing=0 cellPadding=0 width=250>
                  <TBODY>       

	<h1 class="votex" style="font-size:12px; color:red; font-weight:bold;">Cascade Related Articles</h1>

               <TABLE class=search2 border=0 cellSpacing=1 cellPadding=2 
                  width="100%">
                  <TBODY>
                    <TR>
                      <TD class=result-header width=15>No</TD>
                      <TD class=result-header>Title</TD>
                      <TD class=result-header>Log Date</TD>
                      <TD class=result-header>Article Tag</TD>
                      <TD class=result-header width=35>Created By</TD>
                      
                    </TR>
                    <%

			set rs = server.createobject("adodb.recordset")
            sql = "select  top 5 * from CMS_Article where CMS_Rmp=1 order by CMS_Date desc"
			rs.open sql,conn,1,1
			if rs.eof then 
			response.Write("<tr><td colspan=5 style=""color:green;"">No Cascade related  Articles were found!</td></tr>")
			end if 
			if not rs.eof then
			
			do while not rs.eof 
				title = rs("CMS_Title")
				cms_info=rs("CMS_Info")
				if len(cms_info)>50 then
					cms_info = left(cms_info,50)
					cms_info = cms_info & "..."
				end if 
				
			
			%>
                    <TR style="BACKGROUND-COLOR:#CCFF66" class=result-title 
                    onmouseover='this.style.background = "#AAAAAA"' 
                    onmouseout='this.style.background = "#FFFF99"'>
                      <TD align=middle><%=i%>.</TD>
                      <TD><B><A class=site-nav 
                        href=""><U><a href="detail.asp?ShowId=<%=rs("CMS_ID")%>"><span class="<%=color%> <%=bold%>"><%=title%></span></a></U></A></B></TD>
                      <TD>[<%=FormatDateTime(rs("CMS_Date"),2)%>]</TD>
                      <TD><%=rs("CMS_Tag")%></TD>

                      <TD align=middle><%=rs("CMS_Name")%></TD>
                    </TR>
                    <TR style="BACKGROUND-COLOR: #e9e9e9">
                      <TD style="BACKGROUND-COLOR: #fff">&nbsp;</TD>
                      <TD colSpan=6><%=rs("CMS_Info")%></TD>
                      <%
				rs.movenext
			loop	
			%>
            <a style="float:right; color:red; font-weight:bold;" href="cascadelist.asp">View More</a>
            <%
end if 			
		%>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE border=0 cellSpacing=0 cellPadding=0 width=250>
                  <TBODY>       


            
            <h1 class="votex" style="font-size:12px; color:red; font-weight:bold;">New Articles</h1>
                <TABLE class=search2 border=0 cellSpacing=1 cellPadding=2 
                  width="100%">
                  <TBODY>
                    <TR>
                      <TD class=result-header width=15>No</TD>
                      <TD class=result-header>Title</TD>
                      <TD class=result-header>Log Date</TD>
                      <TD class=result-header>Article Tag</TD>

                      <TD class=result-header width=35>Created By</TD>
                    </TR>
                    <%
			dim color,bold,rmp,pic
			set rs = server.createobject("adodb.recordset")
            sql = "select  top 10 * from CMS_Article order by CMS_Date desc"
			rs.open sql,conn,1,1
			if not rs.eof then
			
			do while not rs.eof 
				title = rs("CMS_Title")
				cms_info=rs("CMS_Info")
				if len(cms_info)>50 then
					cms_info = left(cms_info,50)
					cms_info = cms_info & "..."
				end if 
				
			
			%>
                    <TR style="BACKGROUND-COLOR: #e9e9e9" class=result-title 
                    onmouseover='this.style.background = "#AAAAAA"' 
                    onmouseout='this.style.background = "#E9E9E9"'>
                      <TD align=middle><%=i%>.</TD>
                      <TD><B><A class=site-nav 
                        href=""><U><a href="detail.asp?ShowId=<%=rs("CMS_ID")%>"><span class="<%=color%> <%=bold%>"><%=title%></span></a></U></A></B></TD>
                      <TD>[<%=FormatDateTime(rs("CMS_Date"),2)%>]</TD>
                      <TD><%=rs("CMS_Tag")%></TD>

                      <TD align=middle><%=rs("CMS_Name")%></TD>
                    </TR>
                    <TR style="BACKGROUND-COLOR: #e9e9e9">
                      <TD style="BACKGROUND-COLOR: #fff">&nbsp;</TD>
                      <TD colSpan=6><%=rs("CMS_Info")%></TD>
                      <%
				rs.movenext
			loop	
end if 			
		%>
                    </TR>
                  </TBODY>
                </TABLE>
                <TABLE border=0 cellSpacing=0 cellPadding=0 width=250>
                  <TBODY>          
          </p>
    <br>
 

    </div>
    </div>
    <!-- latest article & latest news BEGIN -->
    <table cellpadding="0" cellspacing="0" border="0" width="100%">
      <tr>
        <td width="300" valign="top" ><!-- latest article BEGIN -->
        </td>
        <td valign="top"></td>
      </tr>
    </table>
    </div>
    </td>
  </tr>
  <tr>
    <td class="siemens_footer"><hr width="100%">
      &copy;&nbsp;&nbsp;2012&nbsp;|&nbsp;For Internal Use Only </td>
  </tr>
</table>
</td>
<td valign="top" class="color1">
<div style="padding-top: 10px; padding-left: 7px;">
<div style="border: 1px solid #757575; width: 300px;">
  <div style="font-weight: bold; color: #FFFFFF; background: #757575; padding: 2px 0 2px 2px;">
    <table border="0" cellspacing="0" cellpadding="0" width="100%" style="color: #ffffff; font-weight: bold;">
      <tr>
        <td></td>
        <td width="30" align="right">&nbsp;</td>
      </tr>
    </table>
  </div>
</div>
<!--#include file="navright.asp"-->
</body>
</html>
