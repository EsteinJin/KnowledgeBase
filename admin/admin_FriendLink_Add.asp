<%@codepage = 936%>
<!--�����ʾ���ü���������ʾ-->
<!
<!--#include file="../include/function.asp"-->
<!--#include file="fckeditor/fckeditor.asp"-->
<!--#include file="conn.asp"-->
<%

	if session("Admin") = "" then
		call sussLoctionHref("�Ƿ�����","admin_login.asp")
	end if
	
	'������ӵ�����
	if request.form("send") = "�������" then
    dim LinkName,LinkAddress,LinkInfo,rs
		LinkName = request.form("LinkName")
		LinkAddress = request.form("LinkAddress")
		LinkInfo = request.form("LinkInfo")
		
		if len(LinkName) < 2 or len(LinkName) > 100 then
			call errorHistoryBack("�������Ӳ�С��2λ�����ߴ���100λ")
		end if	
		if len(LinkAddress) < 2  then
			call errorHistoryBack("����Ϊ�գ�")
		end if				

		if len(LinkInfo) < 2  then
			call errorHistoryBack("����Ϊ��")
		end if				


		'��������,�����ɹ�����ת�����ݹ���ҳ��
		addsql = "Insert into FriendLink(LinkName,LinkAddress,LinkInfo) values ('"&LinkName&"','"&LinkAddress&"','"&LinkInfo&"')"
		conn.execute(addsql)
		call sussLoctionHref("���������ɹ�","admin_FriendLink.asp")




	end if
	

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>��̨����</title>
<link rel="stylesheet" type="text/css" href="style/admin.css" />
<script type="text/javascript" src="js/content.js"></script>
</head>
<body>

<form name="add" id="articleadd" method="post" action="admin_FriendLink_Add.asp">
	<dl>
		<dt>�뷢������</dt>
		<dd>��������:&nbsp;
                <input type="text" name="LinkName" class="text" /> 
		</dd>
		<dd>���ӵ�ַ:&nbsp;
                <input type="text" name="LinkAddress" class="text" />
        </dd>
		<dd>������Ϣ��
                <textarea rows="2" name="LinkInfo" ></textarea>
		</dd>
		<dd><input type="submit" onclick="return check();" name="send" value="�������" /></dd>
	</dl>
</form>

</body>
</html>