VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_objParam As SysParameter
Private p_objInput As InputData
Private p_objOutput As OutputData

'-- Constructor
Private Sub Class_Initialize()

End Sub

'-- Refer Parameter
Public Function SetParam(ByVal Para As SysParameter)

    Set p_objParam = Para

End Function

'-- Refer Input and Output
Public Function SetRef(ByVal InDt As InputData, ByVal OutDt As OutputData)

    Set p_objInput = InDt
    Set p_objOutput = OutDt

End Function

'-- Main Process
Public Function CreateOutput()

    

End Function
