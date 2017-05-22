<!--#include file="lib/init.asp" -->
<!--#include file="lib/ToJSON.function.asp" -->
<% 
	Response.AddHeader "Content-Type", "application/json" 

	

	Set simpleArray = CreateObject("System.Collections.ArrayList")
	simpleArray.Add "Apple"
	simpleArray.Add "Banana"
	simpleArray.Add 98
	simpleArray.Add 99

	Set object = Server.CreateObject("Scripting.Dictionary")
	object.add "firstName", "Wilson"
	object.add "lastName", "Anciro"
	object.add "array", simpleArray
	
	Set data = Server.CreateObject("Scripting.Dictionary")
	data.Add "string", "Simple String"
	data.Add "integer", 25
	data.Add "simpleArray", simpleArray
	data.Add "object", object
	
	' Insert User
	' UserInsert = DB.Query("INSERT INTO Users(FirstName) VALUES('Zeki')")
	' If UserInsert Then
		' Set MaxID = DB.Query("SELECT MAX(ID) as ID FROM Users")
		' Set User = DB.Query("SELECT * FROM Users WHERE ID = " & MaxID("ID"))
		
		' Response.Write(ToJSON(User))
	' End If
	
	' Delete User
	' DeleteUser = DB.Query("DELETE FROM Users WHERE ID = 30")
	' If DeleteUser Then
		' Response.Write(DeleteUser)
	' End If
	
%>
