'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Code to make a user defined regex function in excel. Must turn on Regex in Visual Basic -> Tools -> References. Check Microsoft 'VBScript Regular Expressions 1.0 and Microsoft VBScript Regular Expressions 5.5
'-----------------------------------------
'code used from: https://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops
'syntax listed below
'
'
'=regex(cell number, "regex")
'----------------------------------------------------------------
'creating fucntion name and variables
'strInput = cell you want to run the regex command on
'regexPattern = rexex pattern
Function RegX(strInput As String, regexPattern As String) As String
    'creating regEx variable as object of RegExp
    Dim regEx As New RegExp
    'adding with statement to execute series of statements to a single object
    With regEx
    'setting this function to module
        .Global = True
        '.MultiLine = True
    'making pattern case sensitive
        .IgnoreCase = False
    'returning pattern that matches regexPattern
        .Pattern = regexPattern
    'ending with statement
    End With

    If regEx.Test(strInput) Then
        Set matches = regEx.Execute(strInput)
        RegX = matches(0).Value
    Else
        RegX = "not matched"
    End If
End Function

