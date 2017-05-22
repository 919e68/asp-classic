<%
	Response.Write("<h3>Dictionary</h3>") 
	
	Set dict = Server.CreateObject("Scripting.Dictionary")
	dict.Add "wil", "Wilson"
	dict.Add "yel", "Arielle"
	
	dict.Item("wil") = "Anciro"
	
	For i = 0 To dict.Count - 1
		Response.Write( dict.Keys()(i) & " : " & dict.Items()(i) & "<br>")
	Next
%>