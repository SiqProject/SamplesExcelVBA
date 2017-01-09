VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InputData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_objHeaderData As Variant
Private m_objCsvFileList As Variant
Private m_strErrorMessage As String

'-- Constructor
Private Sub Class_Initialize()

    Set m_objHeaderData = CreateObject("Scripting.Dictionary")
    Set m_objCsvFileList = CreateObject("Scripting.Dictionary")

End Sub

'-- Set Header Data
Public Function SetHeaderData(ByVal lw As ListView)

    m_objHeaderData.RemoveAll
    
    Dim i As Long
    For i = 1 To lw.ListItems.Count
        m_objHeaderData.Add lw.ListItems(i).Text, lw.ListItems(i).SubItems(1)
    Next

End Function

'-- CSV FileList
Public Function UpdateCsvFileList(ByVal list As Collection, ByVal bFlag As Boolean)

    Dim file As String
    For Each file In list
    
        '-- bFlag : True=Add, False=Delete
        If m_objCsvFileList.Exists(file) <> bFlag Then
            m_objCsvFileList.Add file, file
        Else
            m_objCsvFileList.Remove file
        End If
        
    Next

End Function

'-- ListView : CSV FileList
Public Sub UpdateLw_CsvList(ByVal lw As ListView)

    lw.ListItems.Clear

    Dim file As String
    For Each file In m_objCsvFileList.Keys
        With lw.ListItems.Add
            .Text = GetFilename(file)
            .SubItems(1) = GetDictionary(file)
        End With
    Next

End Sub

'-- Check Input Data before Creating Output
Public Function PreCheck() As Boolean

    Dim rc As Boolean
    Dim key As String
    
    rc = True
    m_strErrorMessage = ""
    
    '-- Header
    For Each key In m_objHeaderData.Keys
        If m_objHeaderData.item(key) = "" Then
            rc = False
            m_strErrorMessage = ""
        End If
    Next
    
    '-- Context
    If rc = True And m_objCsvFileList.Count = 0 Then
        rc = False
        m_strErrorMessage = ""
    End If
    
    PreCheck = rc

End Function
