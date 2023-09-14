Function RegX(strInput As String, regexPattern As Variant) As String
    ' Check if regexPattern is not empty
    If Not IsEmpty(regexPattern) Then
        ' Convert regexPattern to a string
        Dim pattern As String
        pattern = CStr(regexPattern)
        
        ' Create a RegExp object
        Dim regEx As Object
        Set regEx = CreateObject("VBScript.RegExp")
        
        ' Configure the RegExp object
        With regEx
            .Global = True
            .IgnoreCase = False
            .Pattern = pattern
        End With
        
        ' Test the input string against the regex pattern
        If regEx.Test(strInput) Then
            ' If a match is found, return the matched value
            Set matches = regEx.Execute(strInput)
            RegX = matches(0).Value
        Else
            ' If no match is found, return "not matched"
            RegX = "not matched"
        End If
    Else
        ' If regexPattern is empty, return an error message
        RegX = "Regex pattern is empty"
    End If
End Function
