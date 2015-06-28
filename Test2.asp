<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>无标题文档</title>
<script type="text/javascript">
//获取起始日期   
var startDate=document.all.startdate.value;   
//转换为日期格式   
startDate=startDate.replace(/-/g,"/");   
    
//获取结束日期   
 var endDate=document.all.enddate.value;   
 endDate=endDate.replace(/-/g,"/");   
 //如果起始日期大于结束日期   
 if(Date.parse(startDate)-Date.parse(endDate)>0){   
  alert("起始日期要在结束日期之前!");   
  //返回false   
  return false;   
 }  


</script>

</head>

<body>
</body>
</html>
