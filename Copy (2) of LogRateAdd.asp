<%@codepage = 936%>

<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->

<!-- 三级联动菜单 开始 --> 

<script language="JavaScript"> 
<!-- 
<% 
'二级数据保存到数组 
Dim count2,rsClass2,sqlClass2 
set rsClass2=server.createobject("adodb.recordset") 
sqlClass2="select * from aa" 
rsClass2.open sqlClass2,conn,1,1 
%> 
var subval2 = new Array(); 
//数组结构：一级根值,二级根值,二级显示值 
<% 
count2 = 0 
do while not rsClass2.eof 
%> 
subval2[<%=count2%>] = new Array('<%=rsClass2("aID")%>','<%=rsClass2("ID")%>','<%=rsClass2("Name")%>') 
<% 
count2 = count2 + 1 
rsClass2.movenext 
loop 
rsClass2.close 
%> 
<% 
'三级数据保存到数组 
Dim count3,rsClass3,sqlClass3 
set rsClass3=server.createobject("adodb.recordset") 
sqlClass3="select * from aaa" 
rsClass3.open sqlClass3,conn,1,1 
%> 
var subval3 = new Array(); 
//数组结构：二级根值,三级根值,三级显示值 
<% 
count3 = 0 
do while not rsClass3.eof 
%> 
subval3[<%=count3%>] = new Array('<%=rsClass3("aaID")%>','<%=rsClass3("ID")%>','<%=rsClass3("Name")%>') 
<% 
count3 = count3 + 1 
rsClass3.movenext 
loop 
rsClass3.close 
%> 
function changeselect1(locationid) 
{ 
document.form1.s2.length = 0; 
document.form1.s2.options[0] = new Option('==Type==',''); 
document.form1.s3.length = 0; 
document.form1.s3.options[0] = new Option('==Item==',''); 
for (i=0; i<subval2.length; i++) 
{ 
if (subval2[i][0] == locationid) 
{document.form1.s2.options[document.form1.s2.length] = new Option(subval2[i][2],subval2[i][1]);} 
} 
} 
function changeselect2(locationid) 
{ 
document.form1.s3.length = 0; 
document.form1.s3.options[0] = new Option('==Item==',''); 
for (i=0; i<subval3.length; i++) 
{ 
if (subval3[i][0] == locationid) 
{document.form1.s3.options[document.form1.s3.length] = new Option(subval3[i][2],subval3[i][1]);} 
} 
} 
//--> 
</script>

<%
if request.Form("send")="Submit Log" then
dim CMS_TicketNumber,CMS_AgentName,CMS_Language,CMS_Remote,CMS_Compliance,CMS_HandleTime,CMS_Category,CMS_Type,CMS_Item,CMS_Summary
CMS_TicketNumber=request.Form("CMS_TicketNumber")
CMS_AgentName=request.Form("CMS_AgentName")
CMS_Language=request.Form("CMS_Language")
CMS_Remote=request.Form("CMS_Remote")
CMS_Compliance=request.Form("CMS_Compliance")
CMS_HandleTime=request.Form("CMS_HandleTime")
's1=request.Form("s1")
'set s1rs = server.createobject("adodb.recordset")
's1sql="select * from a where ID="&s1
's1rs.open s1sql,conn,1,1
'CMS_Category=s1rs("Name")
's2=request.Form("s2")
'set s2rs = server.createobject("adodb.recordset")
's2sql="select * from aa where ID="&s2
's2rs.open s2sql,conn,1,1
'CMS_Type=s2rs("Name")

's3=request.Form("s3")
'set s3rs = server.createobject("adodb.recordset")
's3sql="select * from aaa where ID="&s3
's3rs.open s3sql,conn,1,1
'CMS_Item=s3rs("Name")
'CMS_Summary=request.Form("CMS_Summary")


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
addsql="Insert into CMS_LogRate(CMS_Date,CMS_TicketNumber,CMS_AgentName,CMS_Language,CMS_Remote,CMS_Compliance,CMS_HandleTime,CMS_Summary) values (now(),'"&CMS_TicketNumber&"','"&CMS_AgentName&"','"&CMS_Language&"','"&CMS_Remote&"','"&CMS_Compliance&"','"&CMS_HandleTime&"','"&CMS_Summary&"')"
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
<body style="padding:15px;">
<form name="form1" id="articleadd" method="post" action="LogRateAdd.asp">
<dl style="background:#CCFFCC; padding:15px; font-size:12px">
<dt style="font-weight:bold; font-family:Arial; font-size:18px; padding-bottom:10px;">Daily LogRate &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="LogRateList.asp" style="font-family:Arial; float:right">Check Log Rate</a></dt><br/><br/>

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
<input type="text" name="CMS_HandleTime"  id="str" onBlur="IsTime()" value="00:00:00" /> <span>eg: HH:MM:SS</span>
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
<input type="text" name="CMS_TicketNumber"  /><span style="color:red; font-weight:bold;">Input Order ID</span>
</dd>

<dd>
<% 
Dim count1,rsClass1,sqlClass1 
set rsClass1=server.createobject("adodb.recordset") 
sqlClass1="select * from a" 
rsClass1.open sqlClass1,conn,1,1 
%> 
<select name="s1" onChange="changeselect1(this.value)"> 
<option>==Category==</option> 
<% 
count1 = 0 
do while not rsClass1.eof 
response.write"<option value="&rsClass1("ID")&">"&rsClass1("Name")&"</option>" 
count1 = count1 + 1 
rsClass1.movenext 
loop 
rsClass1.close 
%> 
</select> 
<select name="s2" onChange="changeselect2(this.value)"> 
<option>==Type==</option> 
</select> 
<select name="s3"> 
<option>==Item==</option> 
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
<option  value="https://workspace.it-solutions.myatos.net/content/10002278/BASF/Lists/JobAids/All%20Items.aspx">JobAids</option>
<option  value="https://workspace.it-solutions.myatos.net/content/10002278/BASF/Lists/BASF%20Service%20Desk%20Case%20Share/Expand.aspx">Case Share</option>

</select>

Korean Radio:
<iframe id="Radio" height="30" width="100" frameborder="0" scrolling="no" src="SSKorea.asp"></iframe>
</body>
</html>