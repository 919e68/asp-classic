 <%
	Class Database
		Public Connection
		Public ConnectionString
		
		Private Sub Class_Initialize()
			Set Connection = Server.CreateObject("ADODB.Connection")
			'Connection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("db/forums.mdb")
			Connection.ConnectionString = "Driver={SQL Server}; Server=RED; Database=sample_forums; User Id=sa; Password=ABC12abc"
		End Sub
		
		Public Sub Open
			Connection.Open
		End Sub
		
		Public Sub Close
			Connection.Close
		End Sub
	
		Public Function Query(sqlString)
			Dim result

			Connection.Open
			
			Set rows = CreateObject("System.Collections.ArrayList")
			
			Dim sqlType
			sqlType = Split(sqlString, " ")
			If UBound(sqlType) > 0 Then
				If LCase(sqlType(0)) = "select" Then
					Set rs = Server.CreateObject("ADODB.recordset")
					rs.Open sqlString, Connection
					
					Do Until rs.EOF
						Set row = Server.CreateObject("Scripting.Dictionary")

						For Each field in rs.fields
							row.Add field.Name, field.Value
						Next

						rows.Add(row)
						rs.Movenext
					Loop
					Set rs = Nothing
					
					If rows.Count = 1 Then
						Set result = rows(0)
					Else 
						Set result = rows
					End If
					
				ElseIf LCase(sqlType(0)) = "insert" Or LCase(sqlType(0)) = "update" Or LCase(sqlType(0)) = "delete" Then
					Connection.Execute sqlString
					result = true
					
				End If
			End If
			
			Connection.Close

			If Not IsEmpty(result) Then
				If IsObject(result) Then
					Set Query = result
				Else
					Query = result
				End If
			End If
		End Function
	End Class
%>