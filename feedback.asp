<%@codepage = 936%>
<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="fckeditor/fckeditor.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<!--#include file="navleft.asp"-->
<%
	dim showid
	showid = cint(request.querystring("FeedbackId"))
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("error Occured!")
	end if
	
	dim title
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Article where CMS_ID="&showid
	rs.open sql,conn,1,1

	if rs.eof then
		call errorHistoryBack("No Menu Exists")
	else 
	//Comment
	title=rs("CMS_Title")
	FID=rs("CMS_ID")
		
	end if
	
	call close_rs
	
	
%>
   
    <td valign="top"><table border="0" width="100%" height="100%" cellspacing="5" cellpadding="0">
      <tr>
        <td valign="top"><img src="image/header-grey.jpg" border="0" width="600" height="100" vspace="0" hspace="0" alt=""><br>
          <br>
          <div id="win">
          <div id="win-header">Log your Comment Here!</div>
          <div id="win-body">
           <form method="post" action="feedback_Add.asp">
           <span><strong style="color:red">TITLE:</strong><%=title%></span>
           <%
				Dim oFCKeditor
				Set oFCKeditor = New FCKeditor 
				oFCKeditor.BasePath = "fckeditor/"  
				oFCKeditor.ToolbarSet = "Basic" 
				oFCKeditor.Width = "100%"  
				oFCKeditor.Height = "400"  
				oFCKeditor.Value = ""  
				oFCKeditor.Create "content"  
			%>
            <input type="hidden" value="<%=FID%>" name="FID">
            <input type="submit" name="send" value="send">
           </form>

    </div>
    </div>
    <!-- latest article & latest news BEGIN -->
    <table cellpadding="0" cellspacing="0" border="0" width="100%">
      <tr>
        <td width="300" valign="top" ><!-- latest article BEGIN -->
        </td>
        <td valign="top"></td>
      </tr>
    </table>
    </div>
    </td>
  </tr>
  <tr>
    <td class="siemens_footer"><hr width="100%">
      &copy;&nbsp;&nbsp;2012&nbsp;|&nbsp;For Internal Use Only </td>
  </tr>
</table>
</td>
<td valign="top" class="color1">
<div style="padding-top: 10px; padding-left: 7px;">
<div style="border: 1px solid #757575; width: 300px;">
  <div style="font-weight: bold; color: #FFFFFF; background: #757575; padding: 2px 0 2px 2px;">
    <table border="0" cellspacing="0" cellpadding="0" width="100%" style="color: #ffffff; font-weight: bold;">
      <tr>
        <td></td>
        <td width="30" align="right">&nbsp;</td>
      </tr>
    </table>
  </div>
  <div id="win-body"></div>
</div>
<!--#include file="navright.asp"-->
</body>
</html>
