Option Explicit

Dim objExampleGenerator: Set objExampleGenerator = CreateObject("ASPRequestImpersonator.ExampleGenerator")

Dim objRequest: Set objRequest = objExampleGenerator.GetExample1()
WScript.Echo objRequest.Item("Key1")

DisplayData objExampleGenerator.GetExample1()

Function DisplayData(ByVal objRequest)

	Dim strKey, strValue

	WScript.Echo "Querystring: " & objRequest.Querystring
	WScript.Echo objRequest.Querystring.Count & " item(s)"
	For Each strKey in objRequest.Querystring
		WScript.Echo strKey & " (" & objRequest.Querystring(strKey) & ")"
		For Each strValue in objRequest.Querystring(strKey)
			WScript.Echo "- " & strValue
		Next
	Next

	WScript.Echo ""

	WScript.Echo "Form: " & objRequest.Form
	WScript.Echo objRequest.Form.Count & " item(s)"
	For Each strKey in objRequest.Form
		WScript.Echo strKey & " (" & objRequest.Form(strKey) & ")"
		For Each strValue in objRequest.Form(strKey)
			WScript.Echo "- " & strValue
		Next
	Next

	WScript.Echo ""

	WScript.Echo "ServerVariables: " & objRequest.ServerVariables
	WScript.Echo objRequest.ServerVariables.Count & " item(s)"
	For Each strKey in objRequest.ServerVariables
		WScript.Echo strKey & " (" & objRequest.ServerVariables(strKey) & ")"
		For Each strValue in objRequest.ServerVariables(strKey)
			WScript.Echo "- " & strValue
		Next
	Next

	WScript.Echo ""

	Dim objDictAllKeys: Set objDictAllKeys = CreateObject("Scripting.Dictionary")
	For Each strKey in objRequest.Querystring
		If Not objDictAllKeys.Exists(strKey) Then
			objDictAllKeys.Add strKey, ""
		End If
	Next
	For Each strKey in objRequest.Form
		If Not objDictAllKeys.Exists(strKey) Then
			objDictAllKeys.Add strKey, ""
		End If
	Next	
	For Each strKey in objRequest.ServerVariables
		If Not objDictAllKeys.Exists(strKey) Then
			objDictAllKeys.Add strKey, ""
		End If
	Next	

	WScript.Echo "Combined"
	WScript.Echo objDictAllKeys.Count & " item(s)"
	For Each strKey in objDictAllKeys
		WScript.Echo strKey
		For Each strValue in objRequest(strKey)
			WScript.Echo "- " & strValue
		Next
	Next

End Function