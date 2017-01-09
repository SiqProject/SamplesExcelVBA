Attribute VB_Name = "Module1"
Function AutoVersionUp() As Boolean
    
    Dim rc As Boolean
    rc = True
    
    AutoVersionUp = rc
    
End Function

Function IsCsvFile(ByVal filepath As String) As Boolean

    Dim rc As Boolean
    rc = True
    
    IsCsvFile = rc

End Function

Function GetDictionary(ByVal filepath As String) As String

    GetDictionary = Left(filepath, InStrRev(filepath, "\"))

End Function

Function GetFilename(ByVal filepath As String) As String

    GetFilename = Mid(filepath, InStrRev(filepath, "\") + 1)

End Function
