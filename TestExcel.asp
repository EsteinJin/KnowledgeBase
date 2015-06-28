<%@codepage =936%>
<!--
******************************
*直接改一下server.mappath("execl.xls")里的execl路径
*默认execl内的第一行数据为字段名
*我设置的列数为 3 列 ，可根据数据多少添加列就可以了（有不当的请留言）
****************************** 
-->
<style type="text/css">

</style>
<%
dim conn
set conn=server.createobject("adodb.connection")
conn.open "driver={Microsoft Excel Driver (*.xls)};DBQ="&server.mappath("excel/CustomerFeedbackManagement.xls")

dim rs
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from [Sheet5$]",conn,1,1
%>

<table width="" height="" border="1" cellpadding="0" cellspacing="0">
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