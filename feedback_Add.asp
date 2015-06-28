<%@codepage = 936%>
<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->


<%
	
	if request.Form("send")="send" then
	SID=request.Form("FID")
	content=replace(request.Form("content"),"'","")
	IpAddressInfo=trim(request.ServerVariables("REMOTE_ADDR"))
	set agentrs=server.CreateObject("adodb.recordset")
	agentsql="select * from CMS_Agent where Agent_IP like '%"&IpAddressInfo&"%'"


	agentrs.open agentsql,conn,1,1
	if not agentrs.eof then
	commentby=agentrs("Agent_Name")
	commentbyadd=agentrs("Agent_MailAddress")
	else 
    commentby=IpAddressInfo
	end if  
	addsql="insert into MyComment(CommentBy,CommentTime,CommentContent,NewsId) values('"&commentby&"',now(),'"&content&"',"&SID&")"
	conn.execute(addsql)
set msg = Server.CreateOBject( "JMail.Message" )
msg.Logging = true
msg.Charset = "utf-8"
msg.ContentTransferEncoding = "base64"
msg.ContentType = "text/html"  
msg.From = "RGCNSISGOEUSBASFBackOffice@internal.siemens.com"
msg.FromName = "Feedback Article Added By :"&commentby
msg.AddRecipientBCC "shangxue.jin@atos.net","BT Colleage"
msg.AddRecipient commentbyadd,commentby
msg.Subject = "Feedback Article:"&title
msg.Body = content
msg.Send( "apac.internal.siemens-it-solutions.com" )
	call sussLoctionHref("Thanks for your Feedback,Will check for you soon !","/detail.asp?ShowId="&SID)	
	end if 
%>
   
  
