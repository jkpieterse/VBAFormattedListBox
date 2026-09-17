VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufDemo 
   Caption         =   "Formatted ListBox Demo"
   ClientHeight    =   5304
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10104
   OleObjectBlob   =   "ufDemo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Build 001
'***************************************************************************
'
' Author:  JKP Application Development Services, info@jkp-ads.com, https://www.jkp-ads.com
'
' (c)2026, all rights reserved to the authors
'
' You are free to use and adapt the code in these modules for
' your own purposes and to distribute as part of your overall project.
' However all headers and copyright notices should remain intact
'
' You may not publish the code in these modules, for example on a web site,
' without the explicit consent of the authors
'***************************************************************************
'---------------------------------------------------------------------------------------
' File   : ufDemo
' Author : Jan Karel Pieterse
' (c)    : Copyright JKP Application Development Services, all rights reserved
' Date   : 15-Sep-26
' Purpose: Demonstrates the use of the Formatted List Box
'---------------------------------------------------------------------------------------
Option Explicit

' The form owns one list object. The list object owns the dynamic labels and
' focus proxies that are drawn inside frListControl.
Private WithEvents mcList As clsFormattedListBox
Attribute mcList.VB_VarHelpID = -1
Private msAppName As String

Public Property Let AppName(ByVal sValue As String)
    msAppName = sValue
End Property

Private Sub UserForm_Initialize()
    ' Configure the ordinary design-time Frame before attaching the list.
    ApplyReferenceFont
    Set mcList = New clsFormattedListBox
    Set mcList.HostFrame = Me.frListControl
    mcList.ColumnCount = 3
    mcList.ColumnWidths = "145;65;100"
    mcList.RowHeight = 20
    mcList.Padding = 3
    mcList.CellPadding = 3
    mcList.BackColor = Me.frListControl.BackColor
    mcList.ForeColor = Me.frListControl.ForeColor
    mcList.MultiSelect = fmMultiSelectExtended
    PopulateList
End Sub

Private Sub ApplyReferenceFont()
    Dim oControl As MSForms.Control
    Me.Font.Name = "MS Sans Serif"
    Me.Font.Size = 8.25
    For Each oControl In Me.Controls
        oControl.Font.Name = "MS Sans Serif"
        oControl.Font.Size = 8.25
    Next oControl
End Sub

Private Sub PopulateList()
    Dim oItem As clsFormattedListItem
    Dim lIndex As Long
    Dim sName As String
    mcList.BeginUpdate
    mcList.Clear
    For lIndex = 1 To 8
        sName = Choose(lIndex, "Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot", "Golf", "Hotel")
        Set oItem = mcList.AddItem
        oItem.Text(0) = sName
        oItem.Text(1) = "Group " & CStr(((lIndex - 1) Mod 3) + 1)
        oItem.Text(2) = "Value " & Format$(lIndex * 10, "00")
        oItem.Tag = "ID-" & Format$(lIndex, "00")
        ' Characters is zero-independent like Excel: its first character is 1.
        oItem.Characters(1, Len(sName), 0).Bold = (lIndex Mod 2 = 1)
        oItem.Characters(1, Len(oItem.Text(2)), 2).ForeColor = RGB(31, 95, 146)
    Next lIndex
    mcList.EndUpdate
    mcList.ListIndex = 0
    UpdateInfo
End Sub

Private Sub UserForm_Activate()
    ' The list creates a tiny hidden TextBox as its keyboard focus proxy.
    If Not mcList Is Nothing Then mcList.SetFocus
End Sub

Private Sub cmdPopulate_Click()
    PopulateList
    mcList.SetFocus
End Sub

Private Sub cmdAdd_Click()
    Dim oItem As clsFormattedListItem
    Set oItem = mcList.AddItem
    oItem.Text(0) = "New item"
    oItem.Text(1) = "Added"
    oItem.Text(2) = "Value 90"
    oItem.Characters(1, 3, 0).Italic = True
    mcList.ListIndex = mcList.ListCount - 1
    UpdateInfo
    mcList.SetFocus
End Sub

Private Sub cmdSelect_Click()
    If mcList.ListCount >= 3 Then mcList.ListIndex = 2
    UpdateInfo
    mcList.SetFocus
End Sub

Private Sub cmdClear_Click()
    mcList.Clear
    UpdateInfo
    mcList.SetFocus
End Sub

Private Sub cmdBold_Click()
    Dim oItem As clsFormattedListItem
    If mcList.ListIndex < 0 Then Exit Sub
    Set oItem = mcList.Item(mcList.ListIndex)
    oItem.Characters(1, Len(oItem.Text(0)), 0).Bold = Not CBool(oItem.Characters(1, Len(oItem.Text(0)), 0).Bold)
End Sub

Private Sub mcList_Click(ByVal Index As Long)
    UpdateInfo
    mcList.SetFocus
End Sub

Private Sub mcList_Change(ByVal Index As Long)
    UpdateInfo
End Sub

Private Sub mcList_DblClick(ByVal Index As Long)
    MsgBox "Double-clicked " & mcList.Item(Index).Text(0) & " (" & mcList.Item(Index).Tag & ")", , msAppName
End Sub

Private Sub mcList_KeyDown(ByVal Index As Long, ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    UpdateInfo
End Sub

Private Sub frListControl_Resize()
    If Not mcList Is Nothing Then mcList.Resize
End Sub

Private Sub UpdateInfo()
    If mcList Is Nothing Or mcList.ListIndex < 0 Then
        Me.labInfo.Caption = "No current row."
    Else
        Me.labInfo.Caption = "Index: " & mcList.ListIndex & vbCrLf & _
            "Value: " & CStr(mcList.Value) & vbCrLf & _
            "Selected: " & CStr(mcList.Selected(mcList.ListIndex))
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Terminate removes the generated controls and breaks the parent/child references.
    If Not mcList Is Nothing Then mcList.Terminate
    Set mcList = Nothing
End Sub

