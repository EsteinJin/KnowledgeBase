<title>Knowledge Base</title>
<link href="styles/grey.css" rel="stylesheet" type="text/css" media="screen" />
<link rel="shortcut icon" href="images/favicon.ico" >
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" type="text/css" href="style/grey.css" />
<script type="text/javascript" src="Script/jquery-latest.js"></script>
<script language="JavaScript" type="text/JavaScript">
function change_view(obj_name)
                {
				
                   var aa=document.getElementById(obj_name);
                   if(aa.style.display=="")
                         {
                            aa.style.display="none";
                         }
                   else
                         {
                            aa.style.display="";
                         }
                }
function showhide(id){ 
if (document.getElementById){ 
obj = document.getElementById(id); 
if (obj.style.display == "none"){ 
obj.style.display = ""; 
} else { 
obj.style.display = "none"; 
} 
} 
}
	function document.onkeydown()
	{
	if(event.keyCode==13)
	{
	
	event.returnValue=false;
	 event.cancel = true;
	 document.getElementById("MySearchBtn").click();
	 return false;   
	}
	}
 

</script>
<link rel="stylesheet" href="Common/thickbox.css" type="text/css" media="screen" />
</head><body>
<table border="0" margin="0" width="100%" height="100%" cellspacing="0" cellpadding="0">
<tr>
  <td width="144" height="90" rowspan="2" class="keyvisual"></td>
  <td width="612" height="54"><table border="0" cellspacing="0" cellpadding="0" width="100%">
      <tr>
        <td></td>
        <td align="right"><div id="padding-right">&nbsp;</div></td>
      </tr>
    </table></td>
  <td>&nbsp;</td>
</tr>
<tr>
  <td height="36" class="color1" align="right">&nbsp;</td>
  <td class="color2"><div id="padding-left">
  </td>
</tr>
<tr>
  <td height="54" class="color1key"> Wellcome!</td>
  <td class="color2"><!-- top menu BEGIN -->
    <table border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="15">&nbsp;</td>
        <td><b> <a href="index.asp" class="site-nav">Home</a> &nbsp;|&nbsp;<a href="/index.asp " target="" class="site-nav">CasCade</a></td>
      </tr>
    </table></td>
  <td class="color3"></td>
</tr>
<tr>
  <td valign="top" class="color2" width="144" align="left"><div style="padding-top: 10px; padding-bottom: 5px; padding-left: 3px;"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Collaps & Expand</b></div>
    <div id="menuleft" class="leftmenucontainer">
      <%
			dim rs,sql
			set rs = server.createobject("adodb.recordset")
			sql = "select  * from CMS_Nav where CMS_Sid=0 order by CMS_Sort asc"
			rs.open sql,conn,1,1
			'ѭ����Ŀ
			'do while not rs.eof
			for i= 1 to rs.recordcount
			CMS_id=rs("CMS_ID")
		%>
      <ul >
        <li onClick="change_view('a<%=i%>')"><a href="#"><%=rs("CMS_NavName")%></a></li>
      </ul>
      <div id="leftmenu2" >
        <ul id="a<%=i%>" >
          <%

		set rs3=server.createobject("adodb.recordset")
		newsql="select * from CMS_Nav where CMS_Sid="&CMS_id
		rs3.open newsql,conn,1,1
		for j=1 to rs3.recordcount
		'do while not rs3.eof 
		%>
          <li><a href="list.asp?ShowId=<%=rs3("CMS_ID")%>">--<%=rs3("CMS_NavName")%></a></li>
          <%
rs3.movenext
next
%>
        </ul>
      </div>
      <%
				rs.movenext
	next
		%>
    </div></td>
  </td>
