<%@codepage = 936%>
<!--#include file="/include/function.asp"-->
<!--#include file="conn.asp"-->
<%
if request.Form("send")= "�������"  then
dim yzm,content,newsId,IpAddressInfo,addsql
yzm = request.form("yzm")
content=replace(request.form("content"),"'","")
newsId=request.form("newsId")
IpAddressInfo=request.ServerVariables("REMOTE_ADDR")

if len(yzm) <> 4 then
	call errorHistoryBack("Digit Match with 4 Charactors")
end if		
if not isnumeric(yzm) then
	call errorHistoryBack("Digit Match Only with Numbers")
end if
if cint(yzm) <> Session("CheckCode") then
	call errorHistoryBack("Digit Does not Match")
end if
if len(content)>5000 then
    call errorHistoryBack("�ַ��Ѿ�����5000�֣�����ϵ����Ա��")
end if 
addsql="Insert into MyComment(IpAddressInfo,CommentTime,CommentContent,NewsId) values('"&IpAddressInfo&"',now(),'"&content&"','"&newsId&"')"

conn.execute(addsql)
call sussLoctionHref("��л���ۣ�","/FeedbackDetail.asp?ShowId="&newsId)

end if 
%>



<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
</head>

<body>
</body>
</html>
