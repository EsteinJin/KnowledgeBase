<%@codepage = 936%>
<script type="text/javascript" src="Script/jquery-latest.js"></script>
<script type="text/javascript" src="Script/thickbox.js"></script>
<link rel="stylesheet" href="Common/thickbox.css" type="text/css" media="screen" />
<link href="Common.css" type="text/css" rel="stylesheet" />


<%
dim oXML,Entry,OrderID
Set oXML=Server.CreateObject("Microsoft.XMLDOM")
oXML.load(Server.MapPath("StatusQueryTicket.xml"))
Set oXMLRoot=oXML.documentElement
Set Instance=oXMLRoot.selectSingleNode("//instance")

response.Write("<div id=""Contents"">")
response.Write("<table>")
response.Write("<tr>")
response.Write("<th>Ticket ID+</th>")
response.Write("<th>Submiter</th>")
response.Write("<th>Status</th>")
response.Write("<th>Description</th>")
response.Write("</tr>")

for i = 1 to Instance.childnodes.length-1
response.Write("<tr>")
response.Write("<td>")
response.Write(Instance.ChildNodes.item(i).ChildNodes.item(0).text)
response.Write("</td>")
response.Write("<td>")
response.Write(Instance.ChildNodes.item(i).ChildNodes.item(1).text)
response.Write("</td>")
response.Write("<td>")
select case Instance.ChildNodes.item(i).ChildNodes.item(2).text
case 4
response.Write("Resolved")
case 3
response.Write("Pending")
case 2
response.Write("WIP")
case 1
response.Write("Assgined")
end select
response.Write("</td>")
response.Write("<td>")
response.Write("<p style="" overflow:scroll; height:100px; width:300px;"">"&replace(Instance.ChildNodes.item(i).ChildNodes.item(3).text,"\r\n","<br />")&"</p>")
response.Write("</td>")
next 
response.Write("</tr>")
response.Write("</div>")
Set Instance=nothing
set oXMLRoot=nothing
set xml=nothing

%>
