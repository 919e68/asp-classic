<%
	Function ToJSON(data)
		'Response.Write(TypeName(data) & "<br>")
		Dim result
		result = ""
		
		dim delimeter
		
		' Dictionary Type
		If TypeName(data) = "Dictionary" Then
			result = "{"
			
			' iterate dictionary
			For i = 0 To data.Count - 1
				If i = data.Count - 1 Then 
					delimeter = ""
				Else 
					delimeter = ","
				End If
				
				key = """" & data.Keys()(i) & """"
				value = ToJSON(data.Items()(i))
				
				result = result & key & ":" & value & delimeter
			Next
			' end iterate dictionary
			
			result = result & "}"
			
		' Array Type
		ElseIf TypeName(data) = "ArrayList" Then
			result = "["
			For i = 0 To data.Count - 1
				If i = data.Count - 1 Then 
					delimeter = ""
				Else 
					delimeter = ","
				End If
				
				result = result & ToJSON(data(i)) & delimeter
			Next
			result = result & "]"

		' String Type
		ElseIf TypeName(data) = "String" Then
			result = result & """" & data & """" 


		' Integer Type
		ElseIf TypeName(data) = "Integer" Or TypeName(data) = "Long" Then
			result = data

		
		' Boolean Type
		ElseIf TypeName(data) = "Boolean" Then
			If data = True Then
				result = "true"
			Else
				result = "false"
			End If
			
			
		' Null Type
		ElseIf TypeName(data) = "Null" Then
			result = "null"
		Else 
			result = "null"
		End If

		ToJSON = result
	End Function
%>