VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SysParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-- パラメータシート
Private Param As Object

Public Sub Load()

    Dim ws As Worksheet
    '-- Set ws = ThisWorkbook.Worksheets(CONT_WS_PARM)
    Set ws = ThisWorkbook.Worksheets("パラメータ")
    
    Dim row_s, row_index As Long
    Dim col_s As Long
    
    row_s = 1
    col_s = 1
    row_index = 0

    Set Param = Nothing
    Set Param = CreateObject("Scripting.Dictionary")
    Do Until (ws.Cells(row_s + row_index, col_s).value = "")
    
        Dim tmpA As String
        Dim tmpB As String
        
        tmpA = ws.Cells(row_s + row_index, col_s).value
        tmpB = ws.Cells(row_s + row_index, col_s).Offset(0, 1).value
        If Param.Exists(tmpA) = False Then
            Param.Add tmpA, tmpB
        End If
        
        row_index = row_index + 1
        
    Loop

End Sub

Public Function GetValue(ByVal key As String) As String

    Dim rc As String
    rc = ""
    
    If Param.Exists(key) Then
        rc = Param.item(key)
    End If
    
    GetValue = rc

End Function
