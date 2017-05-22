<%
	Set conn = Server.CreateObject("ADODB.Connection")
	'conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("db/forums.mdb")
  conn.ConnectionString = "Driver={SQL Server}; Server=RED; Database=sample_forums; User Id=sa; Password=ABC12abc"
	
%>