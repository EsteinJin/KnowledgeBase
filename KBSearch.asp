<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<style type="text/css">
body{background:#cccccc;}

</style>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

<script type="text/javascript">
function showHint(str)
{

var xmlhttp;
if (str.length==0)
  { 
  document.getElementById("txtHint").innerHTML="";
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
	document.getElementById("txtHint").style.display="";
    document.getElementById("txtHint").innerHTML=xmlhttp.responseText;
    }
  }
xmlhttp.open("GET","gethint.asp?q="+str,true);
xmlhttp.send();


}

function updateview(str)
{
var xmlhttp;
if (str.length==0)
  { 
  
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
	
    }
  }
xmlhttp.open("GET","/updateviewed.asp?showid="+str,true);
xmlhttp.send();

}

</script>

<span style="color:red; font-weight:bold;">KB Search</span><input type="text" id="txt1" style=" width:200px;" onKeyUp="showHint(this.value)" />
<div id="txtHint" style="height:400px; width:800px; overflow:scroll; display:none;">
</body>
</html>