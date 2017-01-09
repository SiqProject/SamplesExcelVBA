VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmMain 
   Caption         =   "F001"
   ClientHeight    =   8904
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   6096
   OleObjectBlob   =   "FrmMain.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_objParam As SysParameter
Private m_objProcess As Process
Private m_objInput As InputData
Private m_objOutput As OutputData

'-- Close Button
Private Sub BtnClose_Click()

    '-- Object Release
    Set m_objParam = Nothing
    Set m_objProcess = Nothing
    Set m_objInput = Nothing
    Set m_objOutput = Nothing

    '-- Application Close
    If Workbooks.Count = 1 Then
        Application.Quit
    End If
    ThisWorkbook.Close
    
End Sub

'-- Create Button
Private Sub BtnCreate_Click()

    '-- Create
    If m_objInput.PreCheck Then
        m_objProcess.CreateOutput
    Else
        '-- Error Message
    End If

End Sub

'-- Delete Button
Private Sub BtnCsvDelete_Click()

    Dim i As Long
    Dim s_item As ListItem
    
    For i = LstCsvList.ListItems.Count To 1 Step -1
        Set s_item = LstCsvList.ListItems(i)
        If s_item.Selected Then
            LstCsvList.ListItems.Remove s_item.Index
        End If
    Next
    
    If LstCsvList.ListItems.Count = 0 Then
        TxtOutput.Text = ""
    End If

End Sub

'-- Reference Button
Private Sub BtnReference_Click()

    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            TxtOutput.Text = .SelectedItems(1) & "\out.xlsx"
        End If
    End With

End Sub

'-- Drop & Drag CSV File
Private Sub LstCsvList_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim FileList As Collection
    Dim ErrList As Collection
    Set FileList = New Collection
    Set ErrList = New Collection
    
    Dim file As Variant
    For Each file In Data.Files
        If IsCsvFile(file) Then
        
            FileList.Add file
            
            '-- Setting Default Output File
            If TxtOutput.Text = "" Then
                TxtOutput.Text = GetDictionary(file) & "\out.xlsx"
            End If
            
        Else
            ErrList.Add file
        End If
    Next
    
    '-- Setting CSV File List
    m_objInput.UpdateCsvFileList FileList, True
    m_objInput.UpdateLw_CsvList LstCsvList
    
    '-- Error : Not CSV Files
    If ErrList.Count > 0 Then
        MsgBox ""
    End If

End Sub

'-- Header Item Click
Private Sub LstHeader_ItemClick(ByVal item As MSComctlLib.ListItem)

    With LstHeader
        FrmHeader.TxtItem = .ListItems(.SelectedItem.key).Text
        FrmHeader.TxtValue = .ListItems(.SelectedItem.key).SubItems(1)
    End With
    
    FrmHeader.Show
    
    m_objInput.SetHeaderData LstHeader

End Sub

'-- Change Type
Private Sub OptTypeA_Change()
    
    With LstHeader
    
        .ListItems.Clear
    
        If OptTypeA.value = True Then
            '-- TypeA
            .ListItems.Add , "xA", "A"
            .ListItems("xA").SubItems(1) = ""
            .ListItems.Add , "xB", "B"
            .ListItems("xB").SubItems(1) = ""
        Else
            '-- TypeB
            .ListItems.Add , "xC", "C"
            .ListItems("xC").SubItems(1) = ""
            .ListItems.Add , "xD", "D"
            .ListItems("xD").SubItems(1) = ""
        End If

    End With
    
End Sub

'-- Form Load
Private Sub UserForm_Initialize()

    '-- Auto VersionUp
    If AutoVersionUp = False Then
        Exit Sub
    End If
    
    '-- Object Create
    Set m_objParam = New SysParameter
    Set m_objProcess = New Process
    Set m_objInput = New InputData
    Set m_objOutput = New OutputData
    
    m_objProcess.SetParam m_objParam
    m_objProcess.SetRef m_objInput, m_objOutput
    
    '-- Load System Parameter
    m_objParam.Load
    
    '-- Setting Header
    LstCsvList.ColumnHeaders.Add , "xFile", "File"
    LstCsvList.ColumnHeaders.Add , "xPath", "Path"
    LstHeader.ColumnHeaders.Add , "xItem", "Item"
    LstHeader.ColumnHeaders.Add , "xValue", "Value"
    
    '-- Initialize of RadioButton
    OptTypeA.value = True
    OptLangJpn.value = True
    OptRegOkz.value = True
    
End Sub
