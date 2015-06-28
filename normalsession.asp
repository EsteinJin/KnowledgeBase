<!--#include file="conn.asp"-->
<%
if session("AgentName") = "" then 
call sussLoctionHref("Login First!","login.asp")
else 
set adrs=server.CreateObject("adodb.recordset")
adsql="select * from CMS_Agent where CMS_AdminName like '%"session("AgentName")&"%' "
adrs.open adsql,conn,1,1
if  not  adrs.eof then
CMS_Role=adrs("CMS_Role")
CMS_Team=adrs("CMS_Team")
end if 
end if 
%>