<script type="text/javascript">
function showHint(str)
{

var xmlhttp;
if (str.length==0)
  { 
  document.getElementById("txtHint").innerHTML="";
  return;
  }
if (window.XMLHttpRequest)
  {// code for IE7+, Firefox, Chrome, Opera, Safari
  xmlhttp=new XMLHttpRequest();
  }
else
  {// code for IE6, IE5
  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
  }
xmlhttp.onreadystatechange=function()
  {
  if (xmlhttp.readyState==4 && xmlhttp.status==200)
    {
	document.getElementById("txtHint").style.display="";
    document.getElementById("txtHint").innerHTML=xmlhttp.responseText;
    }
  }
xmlhttp.open("GET","../gethint.asp?q="+str,true);
xmlhttp.send();


}

function updateview(str)
{
var xmlhttp;
if (str.length==0)
  { 
  
  return;
  }
if (window.XMLHttpRequest)
  {// code for IE7+, Firefox, Chrome, Opera, Safari
  xmlhttp=new XMLHttpRequest();
  }
else
  {// code for IE6, IE5
  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
  }
xmlhttp.onreadystatechange=function()
  {
  if (xmlhttp.readyState==4 && xmlhttp.status==200)
    {
	
    }
  }
xmlhttp.open("GET","/updateviewed.asp?showid="+str,true);
xmlhttp.send();

}

</script>
<div class="mySubIndexRight">
<%
dim Atoolrs,Atoolsql
set Atoolrs=server.createobject("adodb.recordset")
Atoolsql= "select * from ToolsName where ToolsCategory='ATOS'"
Atoolrs.open Atoolsql,conn,1,1
if not Atoolrs.eof then
		
%>

<h1><em><a href="ToolsLinkList.asp?Category=ATOS">MORE</a></em>ATOS Tools Link</h1>
<ul>
<%
do while not Atoolrs.eof 
%>
<li style="float:left; width:160px; " ><a href="<%=Atoolrs("ToolsLink")%>" target="_blank"><%=Atoolrs("ToolsName")%></a>&nbsp;|&nbsp;</li>
<%
Atoolrs.movenext
loop
end if 
	Atoolrs.close
	set Atoolrs = nothing

%>
</ul>
</div>
<div class="mySubIndexRight">
<%
dim Ctoolrs,Ctoolsql
set Ctoolrs=server.createobject("adodb.recordset")
Ctoolsql= "select * from ToolsName where ToolsCategory='Customer'"
Ctoolrs.open Ctoolsql,conn,1,1
if not Ctoolrs.eof then
		
%>

<h1><em><a href="ToolsLinkList.asp?Category=Customer">MORE</a></em>ATOS Tools Link</h1>
<ul>
<%
do while not Ctoolrs.eof 
%>
<li  style="float:left; width:160px; "><a href="<%=Ctoolrs("ToolsLink")%>" target="_blank"><%=Ctoolrs("ToolsName")%></a>&nbsp;|&nbsp;</li>
<%
Ctoolrs.movenext
loop
end if 
	Ctoolrs.close
	set Ctoolrs = nothing

%>
</ul>
</div>
<%
	dim voters,votesql,votetitle,xrs,xsql,votesid
	'recordset是需要建立的，不是单纯变量，而是对象
	set voters = server.createobject("adodb.recordset")
	votesql = "select * from CMS_Vote where CMS_Level=1"
	voters.open votesql,conn,1,1
	
	if not voters.eof then
		votesid = voters("CMS_ID")
		votetitle = voters("CMS_VoteName")
	else
		votetitle = "error!"
	end if
	
	voters.close
	set voters = nothing
	
	set xrs = server.createobject("adodb.recordset")
	xsql = "select * from CMS_Vote where CMS_VoteSid="&votesid
	xrs.open xsql,conn,1,1
%>

<div class="mySubIndexRight">
	<h1>Online Vote</h1>
	<form method="post" action="vote.asp">
	<dl class="vote">
		<dt><%=votetitle%></dt>
		<%
			do while not xrs.eof
		%>
		<dd><input type="radio" name="vote" value="<%=xrs("CMS_VoteName")%>" /> <%=xrs("CMS_VoteName")%></dd>
		<%
				xrs.movenext
			loop
	xrs.close
	set xrs = nothing
		%>
		<dd><input type="submit" value="Vote" /> <input type="button" onclick="javascript:window.open('votex.asp','votex','width=500,height=500')" value="Check" /></dd>
	</dl>
	</form>
    </div>
</div>
</div>
</div>



</div>