<%@codepage =936%>
<script type="text/javascript" src="Script/jquery-latest.js"></script>
<script type="text/javascript" src="Script/thickbox.js"></script>
<link rel="stylesheet" href="Common/thickbox.css" type="text/css" media="screen" />
<link href="Common.css" type="text/css" rel="stylesheet" />

<%
dim oXML,Entry,OrderID
Set oXML=Server.CreateObject("Microsoft.XMLDOM")
oXML.load(Server.MapPath("SDPendingTicketTrack.xml"))
Set oXMLRoot=oXML.documentElement
Set Instance=oXMLRoot.selectSingleNode("//instance")

response.Write("<div id=""Contents"">")
response.Write("<table>")
response.Write("<tr>")
response.Write("<th>OrderID</th>")
response.Write("<th>Submiter</th>")
response.Write("<th>Requester Name</th>")
response.Write("<th>Receiver Name</th>")
response.Write("<th>Category</th>")
response.Write("<th>Type</th>")
response.Write("<th>Item</th>")
response.Write("<th>Status</th>")
response.Write("<th>Assginee++</th>")
response.Write("<th>ExternalID</th>")
response.Write("<th>Group+</th>")
response.Write("<th>Modified Time</th>")
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
response.Write(Instance.ChildNodes.item(i).ChildNodes.item(2).text)
response.Write("</td>")
response.Write("<td>")
response.Write(Instance.ChildNodes.item(i).ChildNodes.item(3).text)
response.Write("</td>")
response.Write("<td>")
response.Write(Instance.ChildNodes.item(i).ChildNodes.item(4).text)
response.Write("</td>")
response.Write("<td>")
response.Write(Instance.ChildNodes.item(i).ChildNodes.item(5).text)
response.Write("</td>")
response.Write("<td>")
response.Write(Instance.ChildNodes.item(i).ChildNodes.item(6).text)
response.Write("</td>")
response.Write("<td>")
response.Write(Instance.ChildNodes.item(i).ChildNodes.item(7).text)
response.Write("</td>")
response.Write("<td>")
response.Write(Instance.ChildNodes.item(i).ChildNodes.item(8).text)
response.Write("</td>")
response.Write("<td>")
response.Write(Instance.ChildNodes.item(i).ChildNodes.item(9).text)
response.Write("</td>")
response.Write("<td>")
response.Write(Instance.ChildNodes.item(i).ChildNodes.item(10).text)
response.Write("</td>")
response.Write("<td>")
response.Write(Instance.ChildNodes.item(i).ChildNodes.item(11).text)
response.Write("</td>")

response.Write("<td><input id=""ShowResult"" alt=""#TB_inline?height=500&width=800&inlineId=myOnPageContent"&i&""" title=""�鿴����:<b>"&Instance.ChildNodes.item(i).ChildNodes.item(0).text&"</b>"" class=""thickbox"" type=""button"" value=""Show"" /> </td>")
response.Write("<div id=""myOnPageContent"&i&""" class=""Descritpion"">")
response.Write("<p>"&replace(Instance.ChildNodes.item(i).ChildNodes.item(12).text,"\r\n","<br />")&"</p>")
response.Write("</div>")
next 
response.Write("</tr>")
response.Write("</div>")
Set Instance=nothing
set oXMLRoot=nothing
set xml=nothing

%>
