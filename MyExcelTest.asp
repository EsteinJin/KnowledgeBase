<%@codepage = 936%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>IE6 Excel Example</title>
<script type="text/javascript">
function startXL(strFile) {
  var myApp = new ActiveXObject("Excel.Application");
  if (myApp != null) {
     myApp.visible = true;
     myApp.workbooks.open(strFile);
  }
  return false
}
</script>
</head>
<body>
<a href="#" onclick="return startXL('/excel/CustomerFeedbackManagement.xls')">Pipeline Quality Report</a>
<iframe src="/excel/CustomerFeedbackManagement.xls" width="100%" height="500"></iframe>


</body>
</html>