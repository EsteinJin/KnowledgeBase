<%@codepage =936%>
<!--
******************************
*直接改一下server.mappath("execl.xls")里的execl路径
*默认execl内的第一行数据为字段名
*我设置的列数为 3 列 ，可根据数据多少添加列就可以了（有不当的请留言）
****************************** 
-->
<style type="text/css">
th { 
font: bold 11px "Trebuchet MS", Verdana, Arial, Helvetica, sans-serif; 
color: #4f6b72; 
border-right: 1px solid #C1DAD7; 
border-bottom: 1px solid #C1DAD7; 
border-top: 1px solid #C1DAD7; 
letter-spacing: 2px; 
text-transform: uppercase; 
text-align: left; 
padding: 6px 6px 6px 12px; 
background: #CAE8EA no-repeat; 
} 
/*power by www.winshell.cn*/ 
th.nobg { 
border-top: 0; 
border-left: 0; 
border-right: 1px solid #C1DAD7; 
background: none; 
} 

td { 
border-right: 1px solid #C1DAD7; 
border-bottom: 1px solid #C1DAD7; 
background: #fff; 
font-size:11px; 
padding: 6px 6px 6px 12px; 
color: #4f6b72; 
} 
</style>
<%
dim conn
set conn=server.createobject("adodb.connection")
conn.open "driver={Microsoft Excel Driver (*.xls)};DBQ="&server.mappath("/excel/BASFVoiceandNetworkinfo.xls")

dim rs
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [Info$]",conn,1,1
%>

<table width="284" height="57" border="1" cellpadding="0" cellspacing="0">
<tr>
<%
for i=0 to rs.fields.count-1
%>
    <td align="center"><%=rs(i).name%></td>

<%
next
%>
</tr>
<%
do while not rs.eof
%>
</p>

<tr>
<%for j=0  to rs.fields.count-1 %>
    <td align="center"><%=rs(j).value%></td>
<%next%>
</tr>

<%
rs.movenext
loop
rs.close
%>
</table>