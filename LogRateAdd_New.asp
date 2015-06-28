<%@codepage = 936%>
<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->
<!-- 三级联动菜单 开始 -->
<%
if request.Form("send")="Submit Log" then
dim CMS_TicketNumber,CMS_AgentName,CMS_Language,CMS_HandleTime,CMS_CallType,CMS_Source,CMS_Topic,CMS_Team
CMS_TicketNumber=request.Form("CMS_TicketNumber")
CMS_AgentName=request.Form("CMS_AgentName")
CMS_HandleTime=request.Form("CMS_HandleTime")
CMS_CallType=request.Form("CMS_CallType")
CMS_Language=request.Form("CMS_Language")
CMS_Source=request.Form("CMS_Source")
CMS_Topic=request.Form("CMS_Topic")
CMS_Team=request.Form("CMS_Team")
CMS_Exist=request.Form("CMS_Exist")

CMS_Site=request.Form("CMS_Site")
CMS_Remote=request.Form("CMS_Remote")
CMS_Rule=request.Form("CMS_Rule")


if CMS_HandleTime = "" then
		call errorHistoryBack("Handle Time Can't be null!")
	end if	
'if CMS_Summary = "" then
'		call errorHistoryBack("Ticket Summary不能为空！")
'	end if	


addsql="Insert into CMS_New_LogRate(CMS_Date,CMS_TicketNumber,CMS_AgentName,CMS_Language,CMS_HandleTime,CMS_CallType,CMS_Source,CMS_Topic,CMS_Team,CMS_Exist,CMS_Site,CMS_Remote,CMS_Rule) values (now(),'"&CMS_TicketNumber&"','"&CMS_AgentName&"','"&CMS_Language&"','"&CMS_HandleTime&"','"&CMS_CallType&"','"&CMS_Source&"','"&CMS_Topic&"','"&CMS_Team&"','"&CMS_Exist&"','"&CMS_Site&"','"&CMS_Remote&"','"&CMS_Rule&"')"
	conn.execute(addsql)
	call sussLoctionHref("Data Logged!","/LogRateAdd_New.asp")	
	end if 





%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>LogRate Logging Page</title>
<link rel="stylesheet" type="text/css" href="style/basic.css" />
<script type="text/javascript" src="script/content.js"></script>
</head>
<body style="padding:15px;">
<form name="form1" id="articleadd" method="post" action="LogRateAdd_New.asp">
  <dl style="background:#CCFFCC; padding:15px; font-size:12px">
  <dt style="font-weight:bold; font-family:Arial; font-size:18px; padding-bottom:10px;">Daily LogRate &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</dt>
  <br/>
  <br/>
  </dt>
  <br/>
  <br/>
  <dd> </dd>
  </dd>
  <dd> Agent  Name：
    <%
ip=Request.ServerVariables("REMOTE_HOST")

set agentrs=server.createobject("adodb.recordset")
agentsql="select * from CMS_New_Agent where Agent_IP like '%"&ip&"%'"
agentrs.open agentsql,conn,1,1
if not agentrs.eof then

%>
    <input type="text" readonly="readonly" name="CMS_AgentName" value="<%=agentrs("Agent_Name")%>" />
    <%
else 

%>
    <select name="CMS_AgentName">
      <%
set newagentrs=server.createobject("adodb.recordset")
newagentsql="select * from CMS_New_Agent"
newagentrs.open newagentsql,conn,1,1
do while not newagentrs.eof 
%>
      <option  value="<%=newagentrs("Agent_Name")%>" ><%=newagentrs("Agent_Name")%></option>
      <%
newagentrs.movenext
loop
end if 
%>
    </select>
  </dd>
  <dd> Handled Time：
    <input type="text" name="CMS_HandleTime"  id="str" onBlur="IsTime()" value="" />
    <span>eg: 只需输入分数</span> </dd>
  <dd>Support Language?:
    <select  name="CMS_Language">
      <option value="Mandarin">Mandarin</option>
      <option value="Cangtonese">Cangtonese</option>
      <option value="Japanese">Japanese</option>
      <option value="Korean">Korean</option>
    </select>
  </dd>
  <dd>Support Team:
    <select  name="CMS_Team">
      <option value="Coke">Coke</option>
      <option value="TW">TW</option>
    </select>
  </dd>

  <dd> Call Type:
    <select name="CMS_CallType">
      <option value="Inbound">Inbound</option>
      <option value="Outbound">Outbound</option>
      <option value="TransferCall">TransferCall</option>
      <option value="JunkCall">JunkCall</option>
      <option value="TestCall">TestCall</option>
      <option value="Queue Monitor">Queue Monitor</option>

    </select>
  </dd>

  <dd> New/Exist Ticket?:
    <select name="CMS_Exist">
      <option value="New Ticket">New Ticket</option>
      <option value="Exist Ticket">Exist Ticket</option>
    </select>
  </dd>
    
    <dd> Support Site:
    <select  name="CMS_Site">
      <%
	set caters=server.CreateObject("adodb.recordset")
	catesql="select * from CMS_Site"
	caters.open catesql,conn,1,1
	if not caters.eof then
	do while not caters.eof
	%>
      <option value="<%=caters("CMS_Site")%>"><%=caters("CMS_Site")%></option>
      <%
	caters.movenext
	loop
	end if
	%>
    </select>
  </dd>

  
  <dd> Ticket Number :
    <input type="text" name="CMS_TicketNumber"  />
    <span style="color:red; font-weight:bold;">Input Ticket Number</span> </dd>

  <dd> Call Source:
    <select  name="CMS_Source">
      <option value="Email">Email</option>
      <option value="Phone">Phone</option>
      <option value="Chat">Chat</option>
    </select>
  </dd>
  <dd> Issue Category:
    <select  name="CMS_Topic">
      <%
	set caters=server.CreateObject("adodb.recordset")
	catesql="select * from CMS_Category"
	caters.open catesql,conn,1,1
	if not caters.eof then
	do while not caters.eof
	%>
      <option value="<%=caters("CMS_Category")%>"><%=caters("CMS_Category")%></option>
      <%
	caters.movenext
	loop
	end if
	%>
    </select>
  </dd>
    <dd> Indicate Remote Is Done?:
    <select name="CMS_Remote">
      <option value="Yes">Yes</option>
      <option value="No">No</option>
    </select>
  </dd>
  
  </dd>
    <dd> Ticket Chasing?:
    <select name="CMS_Rule">
      <option value="Yes">Yes</option>
      <option value="No">No</option>
    </select>
  </dd>
  
  
  <dd> <br />
    <input type="submit"  name="send" value="Submit Log" />
  </dd>
</form>
<br />
</body>
</html>
