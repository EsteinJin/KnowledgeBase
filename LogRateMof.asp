<%@codepage = 936%>

<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->
<%
if request.Form("send")="修改内容" then
dim CMS_TicketNumber,CMS_AgentName,CMS_Language,CMS_Remote,CMS_Compliance,CMS_HandleTime,CMS_Category,CMS_Type,CMS_Item,CMS_Summary,CID
CID=request.Form("id")
CMS_TicketNumber=request.Form("CMS_TicketNumber")
CMS_AgentName=request.Form("CMS_AgentName")
CMS_Language=request.Form("CMS_Language")
CMS_Remote=request.Form("CMS_Remote")
CMS_Compliance=request.Form("CMS_Compliance")
CMS_HandleTime=request.Form("CMS_HandleTime")
'CMS_Category=request.Form("CMS_Category")
'CMS_Type=request.Form("CMS_Type")
'CMS_Item=request.Form("CMS_Item")
'CMS_Summary=request.Form("CMS_Summary")

if CMS_AgentName = "" then
		call errorHistoryBack("Agent名称不能为空！")
	end if
if CMS_HandleTime = "" then
		call errorHistoryBack("Handle Time不能为空！")
	end if	
'if CMS_Summary = "" then
'		call errorHistoryBack("Ticket Summary不能为空！")
'	end if	
if CMS_Compliance="Yes" then
if CMS_TicketNumber = "" then
		call errorHistoryBack("单号不能为空！")
end if 
end if 
  updatesql="update CMS_LogRate set CMS_TicketNumber='"&CMS_TicketNumber&"',CMS_AgentName='"&CMS_AgentName&"',CMS_Language='"&CMS_Language&"',CMS_Remote='"&CMS_Remote&"',CMS_Compliance='"&CMS_Compliance&"',CMS_HandleTime='"&CMS_HandleTime&"' where ID="&CID
 conn.execute(updatesql)
	call sussLoctionHref("内容修改成功","/LogRateList_Raw.asp")	
	end if 




	'if session("Admin") = "" then
	'	call sussLoctionHref("非法操作","admin_login.asp")
	'end if


%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>LogRate记录</title>
<link rel="stylesheet" type="text/css" href="style/basic.css" />
<style type="text/css">
body{}

</style>
<script type="text/javascript" src="js/content.js"></script>
<script type="text/javascript">
/**
* 删除左右两端的空格
*/
function trim(str)
{
     return str.replace(/(^\s*)(\s*$)/g,"");
}


function IsTime()
{
var str = trim(document.getElementById("str").value)
if(str.length==0)
{
alert("时间不能为空！")
document.getElementById("str").focus();
}
else if(str.length!=0)
{
reg=/^((20|21|22|23|[0-1]\d)\:[0-5][0-9])(\:[0-5][0-9])?$/;
if(!reg.test(str)){    
            alert("对不起，您输入的日期格式不正确!");//请将“日期”改成你需要验证的属性名称!    
			document.getElementById("str").focus();
}
else if(str=="00:00:00")  
{
 alert("时间不能为0秒！")
 document.getElementById("str").focus();
} 

}
   


}

</script>
</head>
<body>
<%
	showid = request.querystring("ShowId")

	'非法操作
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("非法操作")
	end if
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_LogRate where ID="&showid
	rs.open sql,conn,1,1	
	if rs.eof then
		call close_rs
		call close_conn
		call errorHistoryBack("不存在此数据")
		'以上全部是验证数据，已经成功
	else
	CMS_TicketNumber=rs("CMS_TicketNumber")
	CMS_AgentName=rs("CMS_AgentName")
	CMS_Language=rs("CMS_Language")
	CMS_Remote=rs("CMS_Remote")
	CMS_Compliance=rs("CMS_Compliance")
	CMS_HandleTime=rs("CMS_HandleTime")
	CMS_Category=rs("CMS_Category")
	CMS_Type=rs("CMS_Type")
	CMS_Item=rs("CMS_Item")
	CMS_Summary=rs("CMS_Summary")
	end if 
%>
<form name="add" id="articleadd" method="post" action="LogRateMof.asp">
<dl>
<dt>LogRate Update</dt>


<dd>
Agent  Name：
<select name="CMS_AgentName">
<%
set agentrs=server.createobject("adodb.recordset")
agentsql="select * from CMS_Agent"
agentrs.open agentsql,conn,1,1
do while not agentrs.eof 
%>
<option  value="<%=agentrs("Agent_Name")%>" ><%=agentrs("Agent_Name")%></option>

<%
agentrs.movenext
loop
%>
</select>

<span style="color:red; font-weight:bold;">Current Value：<%=CMS_AgentName%></span>

</dd>
<input type="hidden" value="<%=showid%>" name="id" />
<dd>
花费时间：
<input type="text" name="CMS_HandleTime"  id="str" onblur="IsTime()"  value="<%=CMS_HandleTime%>"/> <span>e.g: HH:MM:SS</span>
</dd>
<dd>
标题内容：
<input type="text" name="CMS_Summary"  value="<%=CMS_Summary%>"  /> 
</dd>


<dd>
支持语言:
<select  name="CMS_Language">
<option value="Mandarin">Mandarin</option>
<option value="Japanese">Japanese</option>
<option value="Korean">Korean</option>
</select>
<span style="color:red; font-weight:bold;">Current Value：<%=CMS_Language%></span>
</dd>
<dd>
是否远程:
<select  name="CMS_Remote">
<option value="Yes">Yes</option>
<option value="NO">NO</option>
</select>
<span style="color:red; font-weight:bold;">Current Value：<%=CMS_Remote%></span>
</dd>
<dd>
有无单号:
<select  name="CMS_Compliance">
<option value="Yes">Yes</option>
<option value="NO">NO</option>
</select>
<span style="color:red; font-weight:bold;">Current Value：<%=CMS_Compliance%></span>
</dd>
<dd>
记录单号 :
<input type="text" name="CMS_TicketNumber"  value="<%=CMS_TicketNumber%>" />
</dd>

<dd>
Category:
<select  name="CMS_Category">
<%
set rs = server.createobject("adodb.recordset")
sql = "select CMS_Category from CMS_Category "
rs.open sql,conn,1,1
do while not rs.eof
%>
<option value="<%=rs("CMS_Category")%>"><%=rs("CMS_Category")%></option>
<%
rs.movenext
loop
%>
</select>
<span style="color:red; font-weight:bold;">Current Value：<%=CMS_Category%></span>
</dd>

<dd>
Type   :&nbsp;&nbsp;&nbsp;

<select  name="CMS_Type">
<%
set typers = server.createobject("adodb.recordset")
typesql = "select CMS_Type from CMS_Type "
typers.open typesql,conn,1,1
do while not typers.eof
%>
<option value="<%=typers("CMS_Type")%>"><%=typers("CMS_Type")%></option>
<%
typers.movenext
loop
%>

</select>
<span style="color:red; font-weight:bold;">Current Value：<%=CMS_Type%></span>
</dd>
<dd>
Item   :&nbsp;&nbsp;&nbsp;
<select  name="CMS_Item">
<%
set Itemrs = server.createobject("adodb.recordset")
Itemsql = "select CMS_Item from CMS_Item "
Itemrs.open Itemsql,conn,1,1
do while not Itemrs.eof
%>
<option value="<%=Itemrs("CMS_Item")%>"><%=Itemrs("CMS_Item")%></option>
<%
Itemrs.movenext
loop
%>

</select>
<span style="color:red; font-weight:bold;">Current Value：<%=CMS_Item%></span>
</dd>
<dd><input type="submit"  name="send" value="修改内容" /></dd>

</form>

</body>
</html>