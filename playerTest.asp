<%
	dim ShowId
	ShowId=request.QueryString("ShowId")
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Media Player--By Kin</title>
<link href="http://world.kbs.co.kr/english/css/rki_e.css" rel="stylesheet" type="text/css">
<script src='http://world.kbs.co.kr/jss/Js_Basic.js'></script>
</head>
<body topmargin="0" leftmargin="0">
<table width="333" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="310" height="66">
    <%
	if ShowId<10 then
	response.Write(	"<script language=""javascript"">")
	response.Write("DisplayAOD('http://world.kbs.co.kr/src/asx/rki_asx.php?m_name=realkorea&file_name=k12050"&ShowId&".wma&title=2012-05-"&ShowId&"&date=2012-05-"&ShowId&"&lang=k&starttime=&endtime=&info=', '310', '100%','false');")
	response.Write("</script>")

elseif ShowId>=10 and ShowId<=31 then 
response.Write(	"<script language=""javascript"">")
	response.Write("DisplayAOD('http://world.kbs.co.kr/src/asx/rki_asx.php?m_name=realkorea&file_name=k1205"&ShowId&".wma&title=2012-05-"&ShowId&"&date=2012-05-"&ShowId&"&lang=k&starttime=&endtime=&info=', '310', '100%','false');")
	response.Write("</script>")

end if 
	%>


  </tr>   
</table>

</body>
</html>


