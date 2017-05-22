<%
	Class JSON
		Private items

		Private Sub Class_Initialize()  
			Set items = Server.CreateObject("Scripting.Dictionary")
		End Sub  

		Public Sub Add (key, val)
			items.Add key, val
		End Sub
		
		Public Function Data
			Dim result
			result = "{"
			
			If items.Count > 0 Then
				For i = 0 To items.Count - 1
					dim delimeter, key, value
					If i = items.Count - 1 Then 
						delimeter = ""
					Else 
						delimeter = ","
					End If
					
					key = """" & items.Keys()(i) & """"
					value = """" & items.Items()(i) & """"

					' Integer Type Value
					If TypeName(items.Items()(i)) = "Integer" Then
						value = items.Items()(i)
					End If
					
					
					result = result & key & ":" & value & delimeter
				Next
			End If

			result = result & "}"
			Data = result
		End Function

	End Class
%>