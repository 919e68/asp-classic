<%
	Response.Write("<h3>Database Recordset</h3>")
	
	conn.Open
	Set rs = Server.CreateObject("ADODB.recordset")
	rs.Open "SELECT * FROM Users", conn
	
	Do Until rs.EOF
		Response.Write(rs.fields("Username").value & "<br>")
		rs.Movenext
	Loop

	conn.Close
%>