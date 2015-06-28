<title>Knowledge Base</title>
<link href="styles/grey.css" rel="stylesheet" type="text/css" media="screen" />
<link rel="shortcut icon" href="images/favicon.ico" >
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
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

function showHint(str)
{


var xmlhttp;
if (str.length==0)
  {
  document.getElementById("txtHint").value="";
  return;
  }
if (window.XMLHttpRequest)
  {// code for IE7+, Firefox, Chrome, Opera, Safari
  xmlhttp=new XMLHttpRequest();
  }
else
  {// code for IE6, IE5
  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
  }
xmlhttp.onreadystatechange=function()
  {
  if (xmlhttp.readyState==4 && xmlhttp.status==200)
    {
    document.getElementById("txtHint").value=xmlhttp.responseText;
    }
  }
xmlhttp.open("GET","gethint.asp?q="+str,true);
xmlhttp.send();
}
 

</script>
<link rel="stylesheet" href="Common/thickbox.css" type="text/css" media="screen" />
</head>
<body>
<table border="0" margin="0" width="100%" height="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="144" height="90" rowspan="2" class="keyvisual"><img src="image/logo.png" border="0" width="144" height="90" vspace="0" hspace="0" alt=""></td>
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
          <td><b> <a href="index.asp" class="site-nav">Home</a> &nbsp;|&nbsp; <a href=" " class="site-nav">About Me</a> &nbsp;|&nbsp; </td>
        </tr>
      </table></td>
    <td class="color3"></td>
  </tr>
  <tr>
    <td valign="top" class="color2" width="144" align="left">
    <form method="post" action="" style="padding-bottom: 5px; padding-top: 15px;">
    <div style="padding-top: 10px; padding-bottom: 5px; padding-left: 3px;"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Collaps & Expand</b></div>
   
     </td>
    </td>