<%
	Response.Write("<h3>Dynamic Array</h3>")
	
	Dim arr()
	ReDim arr(-1)
	
	ReDim Preserve arr(UBound(arr) + 1)
	arr(0) = "Wilson"
	
	ReDim Preserve arr(UBound(arr) + 1)
	arr(1) = "Tolentino"
	
	ReDim Preserve arr(UBound(arr) + 1)
	arr(2) = "Anciro"
	
	For i = 0 To UBound(arr)
		Response.Write(arr(i) & "<br>")
	Next
%>