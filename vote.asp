<%@codepage =65001%>
<!--#include file="include/function.asp"-->
<!--#include file="conn.asp"-->
<%

	dim votename,rs,sql
	votename = request.form("vote")
	

	if votename = "" then
		errorHistoryBack("Please input your VoteName")
	end if
	
	if isDate(request.cookies("vote")) then
		if DateDiff("s",request.cookies("vote"),now()) < 600000000 then
			call errorHistoryBack("You Alredy Voted!")
		end if
	end if
	
	
	application.lock
	sql = "Update CMS_Vote Set CMS_VoteCount=CMS_VoteCount+1 where CMS_VoteName='"&votename&"'"
	conn.execute(sql)
	application.unlock
	
	
	
	response.cookies("vote") = now()
	call sussLoctionHref("Thank you !","index.asp")
	
			call close_rs
		call close_conn
%>