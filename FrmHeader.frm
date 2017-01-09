VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmHeader 
   Caption         =   "F002"
   ClientHeight    =   1620
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   4560
   OleObjectBlob   =   "FrmHeader.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FrmHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnCancel_Click()

    FrmHeader.Hide
    
End Sub

Private Sub BtnUpdate_Click()

    With FrmMain.LstHeader
        .ListItems(.SelectedItem.key).Text = TxtItem.Text
        .ListItems(.SelectedItem.key).SubItems(1) = TxtValue.Text
    End With
    
    FrmHeader.Hide
    
End Sub
