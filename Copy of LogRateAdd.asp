<%@codepage = 936%>

<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->
<%
if request.Form("send")="Submit Log" then
dim CMS_TicketNumber,CMS_AgentName,CMS_Language,CMS_Remote,CMS_Compliance,CMS_HandleTime,CMS_Category,CMS_Type,CMS_Item,CMS_Summary
CMS_TicketNumber=request.Form("CMS_TicketNumber")
CMS_AgentName=request.Form("CMS_AgentName")
CMS_Language=request.Form("CMS_Language")
CMS_Remote=request.Form("CMS_Remote")
CMS_Compliance=request.Form("CMS_Compliance")
CMS_HandleTime=request.Form("CMS_HandleTime")
CMS_Category=request.Form("CMS_Category")
CMS_Type=request.Form("CMS_Type")
CMS_Item=request.Form("CMS_Item")
CMS_Summary=request.Form("CMS_Summary")


if CMS_HandleTime = "" then
		call errorHistoryBack("Handle Time不能为空！")
	end if	
if CMS_Summary = "" then
		call errorHistoryBack("Ticket Summary不能为空！")
	end if	
if CMS_Compliance="Yes" then
if CMS_TicketNumber = "" then
		call errorHistoryBack("单号不能为空！")
end if 
end if 
addsql="Insert into CMS_LogRate(CMS_Date,CMS_TicketNumber,CMS_AgentName,CMS_Language,CMS_Remote,CMS_Compliance,CMS_HandleTime,CMS_Category,CMS_Type,CMS_Item,CMS_Summary) values (now(),'"&CMS_TicketNumber&"','"&CMS_AgentName&"','"&CMS_Language&"','"&CMS_Remote&"','"&CMS_Compliance&"','"&CMS_HandleTime&"','"&CMS_Category&"','"&CMS_Type&"','"&CMS_Item&"','"&CMS_Summary&"')"
	conn.execute(addsql)
	call sussLoctionHref("内容新增成功","/LogRateList.asp")	
	end if 




	'if session("Admin") = "" then
	'	call sussLoctionHref("非法操作","admin_login.asp")
	'end if


%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>LogRate Logging Page</title>
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
<form name="add" id="articleadd" method="post" action="LogRateAdd.asp">
<dl>
<dt>Daily LogRate &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="LogRateList.asp">Check Log Rate</a></dt>
<dd>
Team:
<select  name="CMS_Team">
<option value="BASF">BASF</option>
<option value="TW">TW</option>
<option value="Novozyme">Novozyme</option>
</select>
</dd>
</dd>

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
</dd>
<dd>
Handled Time：
<input type="text" name="CMS_HandleTime"  id="str" onblur="IsTime()" value="00:00:00" /> <span>eg: HH:MM:SS</span>
</dd>
<dd>
Summary：
<input type="text" name="CMS_Summary"   /> 
</dd>


<dd>
Language:
<select  name="CMS_Language">
<option value="Mandarin">Mandarin</option>
<option value="Japanese">Japanese</option>
<option value="Korean">Korean</option>
</select>
</dd>
<dd>
Is Remoted?:
<select  name="CMS_Remote">
<option value="Yes">Yes</option>
<option value="NO">NO</option>
</select>
</dd>
<dd>
Ticket Compliance:
<select  name="CMS_Compliance">
<option value="Yes">Yes</option>
<option value="NO">NO</option>
</select>
</dd>
<dd>
Ticket Number :
<input type="text" name="CMS_TicketNumber"  />
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
</dd>
<dd><input type="submit"  name="send" value="Submit Log" /></dd>

</form>
<br />

Quick Links:
<select name="QuickLinks"  onchange="window.open(this.options[this.selectedIndex].value,'_blank')" >
<option  value="">Select from List</option>
<option  value="http://sharepoint-ph.it-solutions.myatos.net/sites/SBS/BASF/BASF_APAC/Process/default.aspx">Process Page</option>
<option  value="http://www.computerhope.com/">Online FAQ</option>
<option  value="http://support.microsoft.com/fixit/en-us">Microsof Fix It</option>
<option  value="http://sap01089.de.it-solutions.myatos.net/pkilogin/">Web Based OSD</option>
<option  value="https://knox.it-solutions.atos.net/c/portal/layout?p_l_id=1">KNOX</option>
<option  value="http://www.google.com/webhp?domains=www.google.com&hl=en">GOOGLE</option>
<option  value="http://support.microsoft.com/search/?adv=">Microsoft Tech</option>


</select>

</body>
</html>