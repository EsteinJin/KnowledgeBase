<%@ CODEPAGE=65001 %> 
<% Response.CodePage=65001%> 
<% Response.Charset="UTF-8" %> 

<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->
<%
	dim showid

	showid = request.querystring("ShowId")
	'非法操作
	if showid = "" or not isnumeric(showid) then
		call errorHistoryBack("error occured!")
	end if
	
	dim title,content
	set rs = server.createobject("adodb.recordset")
	sql = "select * from CMS_Article where CMS_ID="&showid
	rs.open sql,conn,1,1
	
	if rs.eof then
		call errorHistoryBack("Data Not Exist!")
	else 
		
		title = rs("CMS_Title")
		content = rs("CMS_Content")
		if instr(content,"upFile") = 1 then
		content= replace(trim(content)," ","%20")
		content= "<a href="&content&">"&content&"</a>"
		
		end if 
		
		info = rs("CMS_Info")
		tag = rs("CMS_Tag")
		keyword = rs("CMS_Keyword")
		name = rs("CMS_Name")
		fdate = rs("CMS_Date")
		viewed=cint(rs("CMS_Viewed"))+1
		updatesql="update CMS_Article set CMS_Viewed="&viewed&" where CMS_ID="&showid
		conn.execute(updatesql)

	end if
	
	call close_rs
	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
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

</script>
<link rel="stylesheet" href="Common/thickbox.css" type="text/css" media="screen" />
</head>
<body>
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
          <td><b> <a href="index.asp" class="site-nav">Home</a> &nbsp;|&nbsp; <a href="/index.asp " target="_blank" class="site-nav">Man Portal</a> &nbsp;|&nbsp;  </td>
        </tr>
      </table></td>
    <td class="color3"></td>
  </tr>
  <tr>
  
  <td valign="top" class="color2" width="144" align="left">

    <div style="padding-top: 10px; padding-bottom: 5px; padding-left: 3px;"><b>Click to Collaps & Expand</b></div>
    <div id="menuleft" class="leftmenucontainer">
      <%
			dim rs,sql
			set rs = server.createobject("adodb.recordset")
			sql = "select  * from CMS_Nav where CMS_Sid=0 order by CMS_Sort asc"
			rs.open sql,conn,1,1
			'循环栏目
			'do while not rs.eof
			for i= 1 to rs.recordcount
			CMS_id=rs("CMS_ID")
		%>
      <ul >
        <li onClick="change_view('a<%=i%>')"><a href="#"><%=rs("CMS_NavName")%></a></li>
      </ul>
      <div id="leftmenu2" >
        <ul id="a<%=i%>" style="display:none;">
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
  
  <TD vAlign=top>
  
  <TABLE border=0 cellSpacing=5 cellPadding=0 width="100%" height="100%">
    <TBODY>
      <TR>
        <TD vAlign=top><H3></H3>
          <TABLE border=0 cellSpacing=0 cellPadding=0 width=100%>
            <TBODY>
              <TR>
                <TD width=65 align=left><A onclick=history.go(-1) 
                  href="javascript:history.go(1);"><IMG 
                  border=0 hspace=5 
src="detail_files/arrow-back.gif">Back</A></TD>
                <TD width=65 align=left><A 
                  href="#"><IMG 
                  style="MARGIN-RIGHT: 5px" border=0 
                  src="detail_files/search.gif">Search</A></TD>
                <TD>&nbsp;</TD>
                <TD>&nbsp;</TD>
                <TD>&nbsp;</TD>
                <TD width=130><A 
                  href="Feedback.asp?FeedbackId=<%=showid%>"><IMG 
                  border=0 hspace=5 src="detail_files/comment.gif" width=12 
                  height=12>Comment / Feedback</A></TD>
              </TR>
            </TBODY>
          </TABLE>
          <DIV id=win>
            <DIV id=win-header>Article</DIV>
            <DIV id=win-body>
              <TABLE class=view_info border=0 cellSpacing=0 cellPadding=3 
            width="100%">
                <TBODY>
                  <TR class=result-title>
                    <TD class=view-padding-title>Title</TD>
                  </TR>
                  <TR>
                    <TD class=view-padding><STRONG><%=title%></STRONG></TD>
                  </TR>
                  <TR class=result-title>
                    <TD class=view-padding-title>Information</TD>
                  </TR>
                  <TR>
                    <TD class=view-padding><%=info%></TD>
                  </TR>
                  <span style="float:left; color:red; font-weight:bold;">Rate:&nbsp;&nbsp;<%=viewed%></a></span>
                  <span style="float:left; color:red; font-weight:bold; margin-left:30px;"><a href="admin/admin_article.asp?del=ok&ShowId=<%=showid%>">Delete</a></span>
                  
                  <span style="float:right; color:red; font-weight:bold;"><a href="admin/admin_article_Front_mof.asp?ShowId=<%=showid%>">Update</a></span>
                  <TR class=result-title>
                    <TD class=view-padding-title>Contents</TD>
                  </TR>
                  <TR>
                    <TD class=view-padding><%=content%></TD>
                  </TR>
                  <TR class=result-title>
                    <TD class=view-padding-title>Article details</TD>
                  </TR>
                  <TR>
                    <TD class=view-padding><TABLE border=0 cellSpacing=0 cellPadding=0 width="100%">
                        <TBODY>
                          <TR>
                            <TD width=100>Creation date:</TD>
                            <TD width=200><STRONG><%=fdate%></STRONG></TD>
                            <TD width=100>Created By:</TD>
                            <TD><STRONG><%=name%></STRONG></TD>
                          </TR>
                          <TR>
                            <TD>Tag:</TD>
                            <TD><STRONG><%=tag%></STRONG></TD>
                            <TD>Keyword:</TD>
                            <TD><STRONG><%=keyword%></STRONG></TD>
                          </TR>
                        </TBODY>
                      </TABLE></TD>
                  </TR>
                </TBODY>
              </TABLE>
            </DIV>
          </DIV>
          <TABLE border=0 cellSpacing=0 cellPadding=0 width=600>
            <TBODY>
              <TR>
                <TD width=65 align=left></TD>
                <TD width=65 align=left></TD>
                <TD>&nbsp;</TD>
                <TD>&nbsp;</TD>
                <TD>&nbsp;</TD>
                <TD width=130></TD>
              </TR>
            </TBODY>
          </TABLE></TD>
      </TR>
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
            <td>Search KB</td>
            <td width="30" align="right">&nbsp;</td>
          </tr>
        </table>
      </div>
     
    </div>
    <div style="display: block; height: 0; font-size:0; clear: both; visibility:hidden;"></div>
<br />
 <form class="search" method="post" action="search.asp" style=" padding-left:2px; margin-top:2px;">
		
        <select name="kind" style="width:60px;">
			
            <option selected="selected" value="1">Title</option>
			<option value="2">Keyword</option>
			<option value="3">Potrion </option>
		</select>

		<input type="text" name="keyword" style="width:90px; height:19px;" />
<input id="MySearchBtn" type="submit" value="search"  name="send" style=" height:19px;" />
	</form>
    
    
<FORM method="GET" action="http://www.google.com/search" target="_blank">
  <table>
    <tr>
      <td>Search Google<br>
        <input type="hidden" name="domains" value="www.google.com">
        <INPUT TYPE="text" name="q" size="23" maxlength="255" value="" style="background-image: url(images/google2.jpg); background-repeat: no-repeat; background-position: 100% 0%;">
        <input type="hidden" name="hl" value="en">
        <INPUT type="submit" name="sa" VALUE=" go! ">
      </td>
    </tr>
  </table>
  <br>
    </FORM>
    <FORM method="GET" action="http://www.google.com/search" target="_blank">
      <table>
        <tr>
          <td>Search Google<br>
            <input type="hidden" name="domains" value="www.google.com">
            <INPUT TYPE="text" name="q" size="23" maxlength="255" value="" style="background-image: url(images/google2.jpg); background-repeat: no-repeat; background-position: 100% 0%;">
            <input type="hidden" name="hl" value="en">
            <INPUT type="submit" name="sa" VALUE=" go! ">
          </td>
        </tr>
      </table>
      <br>
    </FORM>
    <form name="microsoft" action="http://support.microsoft.com/search/?adv=&query=" method="get" target="_blank">
      <table>
        <tr>
          <td>Search Microsoft Knowledge Base<br>
            <input type="text" name="query" size="23" style="background-image: url(images/microsoft2.jpg); background-repeat: no-repeat; background-position: 100% 0%;">
            <input type="submit" value=" go! ">
        </tr>
        </td>
        
      </table>
      <br>
    </form>
    <br />
    <div style="background: none; padding-left: 3px;">
      <table border="0" cellspacing="0" cellpadding="5" width="200">
        <tr style="background: #e8e8e8;">
          <td colspan="2"><select name="quick_links" onChange="window.open(this.options[this.selectedIndex].value,'_blank')">
              <option value="">--- Quick links ---</option>
              <optgroup label="Desktop/Clients">
              <option value="http://support.microsoft.com/gp/cp_fixit_main">Microsoft Fix It</option>
              <option value="http://www.computerhope.com/">Online IT FAQ</option>
              <optgroup label="General"> </optgroup>
            </select>
          </td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
      </table>
      <table border="0" cellspacing="2" cellpadding="2" width="200">
        <tr style="background: #e8e8e8;">
          <td><a href="#" onClick="showhide('phalph'); return(false);"><img src="image/arrow.gif" hspace="5" border="0" />Phonetic alphabet</a> </td>
        </tr>
      </table>
      <div  id="phalph">
        <table border="0" cellspacing="1" cellpadding="3" width="200">
          <tr style="background-color: #dedede" onmouseover='this.style.background = "#AAAAAA"' onmouseout='this.style.background = "#dedede"'>
            <td width="10"><strong>A</strong></td>
            <td>Alpha</td>
            <td width="10"><strong>N</strong></td>
            <td>November</td>
          </tr>
          <tr style="background-color: #dedede" onmouseover='this.style.background = "#AAAAAA"' onmouseout='this.style.background = "#dedede"'>
            <td><strong>B</strong></td>
            <td>Bravo</td>
            <td><strong>O</strong></td>
            <td>Oscar</td>
          </tr>
          <tr style="background-color: #dedede" onmouseover='this.style.background = "#AAAAAA"' onmouseout='this.style.background = "#dedede"'>
            <td><strong>C</strong></td>
            <td>Charlie</td>
            <td><strong>P</strong></td>
            <td>Papa</td>
          </tr>
          <tr style="background-color: #dedede" onmouseover='this.style.background = "#AAAAAA"' onmouseout='this.style.background = "#dedede"'>
            <td><strong>D</strong></td>
            <td>Delta</td>
            <td><strong>Q</strong></td>
            <td>Quebec</td>
          </tr>
          <tr style="background-color: #dedede" onmouseover='this.style.background = "#AAAAAA"' onmouseout='this.style.background = "#dedede"'>
            <td><strong>E</strong></td>
            <td>Echo</td>
            <td><strong>R</strong></td>
            <td>Romeo</td>
          </tr>
          <tr style="background-color: #dedede" onmouseover='this.style.background = "#AAAAAA"' onmouseout='this.style.background = "#dedede"'>
            <td><strong>F</strong></td>
            <td>Foxtrot</td>
            <td><strong>S</strong></td>
            <td>Sierra</td>
          </tr>
          <tr style="background-color: #dedede" onmouseover='this.style.background = "#AAAAAA"' onmouseout='this.style.background = "#dedede"'>
            <td><strong>G</strong></td>
            <td>Golf</td>
            <td><strong>T</strong></td>
            <td>Tango</td>
          </tr>
          <tr style="background-color: #dedede" onmouseover='this.style.background = "#AAAAAA"' onmouseout='this.style.background = "#dedede"'>
            <td><strong>H</strong></td>
            <td>Hotel</td>
            <td><strong>U</strong></td>
            <td>Uniform</td>
          </tr>
          <tr style="background-color: #dedede" onmouseover='this.style.background = "#AAAAAA"' onmouseout='this.style.background = "#dedede"'>
            <td><strong>I</strong></td>
            <td>India</td>
            <td><strong>V</strong></td>
            <td>Victor</td>
          </tr>
          <tr style="background-color: #dedede" onmouseover='this.style.background = "#AAAAAA"' onmouseout='this.style.background = "#dedede"'>
            <td><strong>J</strong></td>
            <td>Juliet</td>
            <td><strong>W</strong></td>
            <td>Whiskey</td>
          </tr>
          <tr style="background-color: #dedede" onmouseover='this.style.background = "#AAAAAA"' onmouseout='this.style.background = "#dedede"'>
            <td><strong>K</strong></td>
            <td>Kilo</td>
            <td><strong>X</strong></td>
            <td>X-Ray</td>
          </tr>
          <tr style="background-color: #dedede" onmouseover='this.style.background = "#AAAAAA"' onmouseout='this.style.background = "#dedede"'>
            <td><strong>L</strong></td>
            <td>Lima</td>
            <td><strong>Y</strong></td>
            <td>Yankee</td>
          </tr>
          <tr style="background-color: #dedede" onmouseover='this.style.background = "#AAAAAA"' onmouseout='this.style.background = "#dedede"'>
            <td><strong>M</strong></td>
            <td>Mike</td>
            <td><strong>Z</strong></td>
            <td>Zulu</td>
          </tr>
        </table>
      </div>
    </div>
    </td>
    </tr>
</table>
</body>
</html>
