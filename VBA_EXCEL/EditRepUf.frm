VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EditRepUf 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   84
   ClientTop       =   360
   ClientWidth     =   3660
   OleObjectBlob   =   "EditRepUf.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "EditRepUf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents ListReport As MSForms.ListBox
Dim WithEvents OKButton As MSForms.CommandButton
Dim WithEvents CancelButton As MSForms.CommandButton
Dim WithEvents AddReport As MSForms.CommandButton
Dim WithEvents Frame As MSForms.Frame
Attribute Frame.VB_VarHelpID = -1
Dim WithEvents RegReport As MSForms.CommandButton
Dim WithEvents UnRegReport As MSForms.CommandButton
Dim myTable As Range

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub OKButton_Click()
    myTable.Value = ListReport.List
    Unload Me
End Sub

Private Sub RegReport_Click()
    Call SelectedItemValue("完了")
End Sub

Private Sub UnRegReport_Click()
     Call SelectedItemValue("未完了")
End Sub

Private Sub UserForm_Initialize()
    Set myTable = Sheets(1).ListObjects(1).DataBodyRange
    Me.Width = 380
    Me.Height = 230
    Set ListReport = Me.Controls.Add("Forms.ListBox.1", "ListReport", True)
    Call ListReport_Init
    Set OKButton = Me.Controls.Add("Forms.CommandButton.1", "ListReport", True)
    Call OKButton_Init
    Set CancelButton = Me.Controls.Add("Forms.CommandButton.1", "CancelButton", True)
    Call CancelButton_Init
    Set Frame = Me.Controls.Add("Forms.Frame.1", "Frame", True)
    Call Frame_Init
    Set RegReport = Frame.Controls.Add("Forms.CommandButton.1", "RegReport", True)
    Call RegReport_Init
    Set UnRegReport = Frame.Controls.Add("Forms.CommandButton.1", "UnRegReport", True)
    Call UnRegReport_Init
    
End Sub

Private Sub ListReport_Init()
    With ListReport
        .Top = 10
        .Left = 10
        .Height = Me.Height - 50
        .Width = 200
        .ColumnCount = 3
        .ColumnWidths = "100;60;30"
        .MultiSelect = fmMultiSelectExtended
        .List = myTable.Value
        .ListIndex = 0
    End With
End Sub
Private Sub OKButton_Init()
    With OKButton
        .Top = ListReport.Top
        .Left = ListReport.Width + ListReport.Left + 20
        .Width = 110
        .Height = 30
        .Caption = "OK"
    End With
End Sub

Private Sub CancelButton_Init()
    With CancelButton
        .Top = OKButton.Top + OKButton.Height + 10
        .Left = OKButton.Left
        .Width = OKButton.Width
        .Height = OKButton.Height
        .Caption = "キャンセル(C)"
    End With
End Sub

Private Sub Frame_Init()
    With Frame
        .Top = CancelButton.Top + CancelButton.Height + 10
        .Left = OKButton.Left
        .Width = OKButton.Width
        .Height = OKButton.Height * 3
        .Caption = "課題の編集"
    End With
End Sub

Private Sub RegReport_Init()
    With RegReport
        .Top = 10
        .Left = 10
        .Width = Frame.Width * 0.8
        .Height = OKButton.Height * 0.8
        .Caption = "課題 ""完了"""
        .BackColor = RGB(200, 233, 240)
    End With
End Sub

Private Sub UnRegReport_Init()
    With UnRegReport
        .Top = RegReport.Top + RegReport.Height + 10
        .Left = RegReport.Left
        .Width = RegReport.Width
        .Height = RegReport.Height
        .Caption = "課題 ""未完了"""
        .BackColor = RGB(233, 200, 240)
    End With
End Sub

Private Sub SelectedItemValue(txt As String)
    Dim i As Long
    Dim j As Long
    With ListReport
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                .List(i, 2) = txt
            End If
        Next i
    End With
End Sub
