<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<script type="text/javascript">
function Test(Src)
{
 //document.getElementById("RadioResource").src=Src;
<!-- 
window.open ('playerTest.asp?ShowId='+Src,'newwindow','height=40,width=330,bottom=0,right=0,toolbar=no,menubar=no,scrollbars=no, resizable=no,location=no, status=no') 

//写成一行 
--> 

 
}

</script>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>无标题文档</title>
</head>

<body style="background:#CCFFCC;">
<iframe id="RadioResource"  height="0" width="0" frameborder="0" scrolling="no" src=""></iframe>
<select name="QuickLinks"  onchange="Test(this.options[this.selectedIndex].value)" >

<option value="">选择即播放</option>
<%
dim i 
for i=1 to 31
%>
<option  value="<%=i%>">5月<%=i%>日Radio</option>
<%
next
%>
</select>
</body>
</html>
