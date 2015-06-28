<div id="myHeader">
<div id="MyLeftHeader">
<a>System Announcement</a>
<marquee direction="up" height="100" scrollamount="2" onmouseover="this.stop()" onmouseout="this.start()">
<%
		dim toprs,topsql
		set toprs = server.createobject("adodb.recordset")
		topsql = "select * from CMS_Article where CMS_Top=1 order by CMS_Date desc"
		toprs.open topsql,conn,1,1
		
		for i = 1 to 2
			if toprs.eof then exit for
				
				info = toprs("CMS_Info")
				if len(info) > 40 then
					info = left(info,40)
					info = info & "..."
				end if
	%>
	
	<dl class="top">
		<dt><a href="detail.asp?ShowId=<%=toprs("CMS_ID")%>"><%=toprs("CMS_Title")%></a></dt>
		<dd><%=info%></dd>
	</dl>
	<%
			toprs.movenext
		next

	%>
</marquee>   
</div>
<div id="MyMiddleHeader">
<a>Major Issue Handling Status Announcement</a>

<marquee direction="up" height="100" scrollamount="2" onmouseover="this.stop()" onmouseout="this.start()">
<%
		dim Majorrs,Majorsql
		set Majorrs = server.createobject("adodb.recordset")
		Majorsql = "select * from CMS_Article where CMS_Rmp=1 order by CMS_Date desc"
		Majorrs.open Majorsql,conn,1,1
		
		for i = 1 to 2
			if Majorrs.eof then exit for
				
				info = Majorrs("CMS_Info")
				if len(info) > 40 then
					info = left(info,40)
					info = info & "..."
				end if
	%>
	
	<dl class="top">
		<dt><a href="detail.asp?ShowId=<%=Majorrs("CMS_ID")%>"><%=Majorrs("CMS_Title")%></a></dt>
		<dd><%=info%></dd>
	</dl>
	<%
			Majorrs.movenext
		next

	%>
</marquee>  

</div>
<div id="MyRightHeader">
<a>Important Update/Change Weekly HL/Sign Off</a>
<marquee direction="up" height="100" scrollamount="2" onmouseover="this.stop()" onmouseout="this.start()">
<%
dim HLrs, HLsql
 set HLrs=server.createobject("adodb.recordset")
		HLsql = "select * from CMS_Article where CMS_Bold=1 order by CMS_Date desc"
		HLrs.open HLsql,conn,1,1
		
		for i = 1 to 2
			if HLrs.eof then exit for
				
				info = HLrs("CMS_Info")
				if len(info) > 40 then
					info = left(info,40)
					info = info & "..."
				end if
	%>
	
	<dl class="top">
		<dt><a href="detail.asp?ShowId=<%=HLrs("CMS_ID")%>"><%=HLrs("CMS_Title")%></a></dt>
		<dd><%=info%></dd>
	</dl>
	<%
			HLrs.movenext
		next
	%>

</marquee>  
</div>
</div>
<div id="MyNav">
<ul>
<li><a href="index.asp">Home</a>  
<ul>
<li><a href="clist.asp?ShowId=174">Process</a></li>
</ul>
 </li>
<%
			dim rs,sql
			set rs = server.createobject("adodb.recordset")
			sql = "select  * from CMS_Nav where CMS_Sid=0 order by CMS_Sort asc"
			rs.open sql,conn,1,1
			'循环栏目
			do while not rs.eof
			CMS_id=rs("CMS_ID")
		%>
		<li> 
          <a href="list.asp?ShowId=<%=rs("CMS_ID")%>"><%=rs("CMS_NavName")%></a>
          <ul>
		<%
		set rs3=server.createobject("adodb.recordset")
		newsql="select * from CMS_Nav where CMS_Sid="&CMS_id
		rs3.open newsql,conn,1,1
		do while not rs3.eof 
		%>
          
<li><a href="clist.asp?ShowId=<%=rs3("CMS_ID")%>"><%=rs3("CMS_NavName")%></a></li>
<%
rs3.movenext
loop
%>           

</ul> 
       </li>
		<%
				rs.movenext
			loop
		%>
        <li><a href="QCFeedback.asp">Complaint</a>  </li>
        <li><a href="QCCompliment.asp">Compliment</a>   </li>
        <li><a href="Escalation.asp">ViewIssue</a>  </li>
        <li><a href="ViewIssue.asp">LogIssue</a> </li>

</ul>
</div>

<div id="MySearch">
<form class="search" method="post" action="search.asp" style=" padding-left:2px; margin-top:2px;">
		<select name="kind">
			<option selected="selected" value="1">By Title</option>
			<option value="2">By Keyword</option>
			<option value="3">Potrion Match Search</option>
		</select>
		<input type="text" name="keyword" />
<input id="MySearchBtn" type="submit" value="Search"  name="send" />
	</form>

</form>
</div>
<div id="MyTag">

<span>Favorite Tags:</span>

		<%
			set rs = server.createobject("adodb.recordset")
			sql = "select * from CMS_Tag order by CMS_TagCount desc"
			rs.open sql,conn,1,1
		
			for i = 1 to 10
				if rs.eof then exit for
		%>
		<a href="tag.asp?tag=<%=rs("CMS_TagName")%>"><%=rs("CMS_TagName")%>(<%=rs("CMS_TagCount")%>)</a>
		<%
				rs.movenext
			next
		%>


</div>
</div>