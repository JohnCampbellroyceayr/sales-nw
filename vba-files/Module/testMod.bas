Attribute VB_Name = "TestMod"


Sub test()
    Dim path As String, obj As Object
    
    path = buildPath("/quotes/test.json")

    obj = ParseJson(readData(path))
End Sub


Function ParseJson(jsonString As String) As Object
    Dim scriptControl As Object
    Set scriptControl = CreateObject("MSScriptControl.ScriptControl")
    scriptControl.Language = "JScript"
    Set ParseJson = scriptControl.Eval("(" + jsonString + ")")
End Function

Sub TestJsonParsing()
    Dim jsonString As String
    Dim jsonObject As Object
    Dim propertyName As Variant
    

    jsonString = "{""name"":""John"",""age"":30,""city"":""New York""}"
    Set jsonObject = ParseJson(jsonString)

    For Each propertyName In jsonObject
        Debug.Print propertyName & ": " & jsonObject(propertyName)
    Next propertyName
End Sub