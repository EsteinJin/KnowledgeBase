<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp"-->
<!--#include file="include/function.asp"-->
<%
set rs=server.CreateObject("adodb.recordset")
sql="select * from CMS_New_LogRate"
rs.open sql,conn,1,1
if not (rs.eof and rs.bof) then 
dim ttxt, file, filepath , writefile
ttxt="LogRate.csv"
set file= createobject("scripting.FileSystemObject")
Application.Lock()
filepath=server.MapPath("/"&ttxt)
set writefile=file.createTextFile(filepath,true)
writefile.WriteLine "CMS_Date,CMS_TicketNumber,CMS_AgentName,CMS_Language,CMS_HandleTime,CMS_CallType,CMS_Source,CMS_Topic,CMS_Team,CMS_Exist,CMS_Site,CMS_Remote,CMS_Rule"
do while not rs.eof 
writefile.WriteLine rs("CMS_Date")&","&rs("CMS_TicketNumber")&","&rs("CMS_AgentName")&","&rs("CMS_Language")&","&rs("CMS_HandleTime")&","&rs("CMS_CallType")&","&rs("CMS_Source")&","&rs("CMS_Topic")&","&rs("CMS_Team")&","&rs("CMS_Exist")&","&rs("CMS_Site")&","&rs("CMS_Remote")&","&rs("CMS_Rule")
rs.movenext
loop
writefile.close
Application.UnLock()
rs.close
set rs=nothing
call sussLoctionHref("Hold...","/"&ttxt)	
end if 

function HTMLEncode(fString)
 if not isnull(fString) then
    fString = Replace(fString,",", "，")
    fString = Replace(fString,chr(10), "，")
    fString = Replace(fString,chr(13), " ")
    fString = Replace(fString,"<br>", "，")
    fString = Replace(fString,"&nbsp;", " ")
    HTMLEncode2 = fString
 end if
end function


%>


















