<%
	Set list = CreateObject("System.Collections.ArrayList")
	list.Add "Banana"
	list.Add "Apple"
	list.Add 123
	
	Response.Write(list.Count)
	For i = 0 To list.Count - 1
		Response.Write(list(i))
	Next
%>