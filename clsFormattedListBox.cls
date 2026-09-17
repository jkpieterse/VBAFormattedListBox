VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormattedListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' File   : clsFormattedListBox
' Author : Jan Karel Pieterse
' (c)    : Copyright JKP Application Development Services, all rights reserved
' Date   : 15-Sep-26
' Purpose: Provides a ListBox-style control inside an MSForms Frame, with character formatting.
'---------------------------------------------------------------------------------------
Option Explicit

'Provides a ListBox-style control inside an MSForms Frame, with character formatting.
Private Const MC_SHIFT_MASK As Integer = 1
Private Const MC_CTRL_MASK As Integer = 2

Private WithEvents moHostFrame As MSForms.Frame
Attribute moHostFrame.VB_VarHelpID = -1
Private mcolItems As Collection
Private mlColumnCount As Long
Private msColumnWidths As String
Private msngColumnWidths() As Single
Private mlBoundColumn As Long
Private mlListIndex As Long
Private mlAnchorIndex As Long
Private mlMultiSelect As Long
Private mlUpdateDepth As Long
Private mbRefreshPending As Boolean
Private msngRowHeight As Single
Private msngPadding As Single
Private msngCellPadding As Single
Private mlBackColor As Long
Private mlForeColor As Long
Private mlSelectedBackColor As Long
Private mlSelectedForeColor As Long
Private msControlPrefix As String
Private msngItemTops() As Single
Private msngItemHeights() As Single

'Raised after the user clicks a row; Index is zero-based.
Public Event Click(ByVal Index As Long)
'Raised after the user double-clicks a row; Index is zero-based.
Public Event DblClick(ByVal Index As Long)
'Raised whenever the current row or selected rows change.
Public Event Change(ByVal Index As Long)
'Raised after the control handles a key press for the current row.
Public Event KeyDown(ByVal Index As Long, ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'Raised when a key is released while the formatted list has focus.
Public Event KeyUp(ByVal Index As Long, ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

'Creates an empty single-column list and sets familiar Windows ListBox defaults.
Private Sub Class_Initialize()
    Static lInstance As Long

    lInstance = lInstance + 1
    msControlPrefix = "FList" & CStr(lInstance)
    Set mcolItems = New Collection
    mlColumnCount = 1
    ReDim msngColumnWidths(0 To 0)
    msngColumnWidths(0) = 120
    mlBoundColumn = 1
    mlListIndex = -1
    mlAnchorIndex = -1
    mlMultiSelect = fmMultiSelectSingle
    msngRowHeight = 15
    msngPadding = 2
    msngCellPadding = 2
    mlBackColor = vbWindowBackground
    mlForeColor = vbWindowText
    mlSelectedBackColor = vbHighlight
    mlSelectedForeColor = vbHighlightText
End Sub

'Attaches the formatted list to the Frame that will contain its dynamic controls.
Public Property Set HostFrame(ByVal oFrame As MSForms.Frame)
    If Not moHostFrame Is Nothing Then ClearRenderedControls
    Set moHostFrame = oFrame
    If Not moHostFrame Is Nothing Then
        mlBackColor = moHostFrame.BackColor
        mlForeColor = moHostFrame.ForeColor
        If Len(msColumnWidths) = 0 Then msngColumnWidths(0) = MaxSingle(20, moHostFrame.InsideWidth - 2 * msngPadding)
        RequestRefresh
    End If
End Property

'Returns the Frame currently used to display the list.
Public Property Get HostFrame() As MSForms.Frame
    Set HostFrame = moHostFrame
End Property

'Returns the unique prefix used for this instance's dynamic control names.
Public Property Get ControlPrefix() As String
    ControlPrefix = msControlPrefix
End Property

'Returns the number of items in the list.
Public Property Get ListCount() As Long
    ListCount = mcolItems.Count
End Property

'Returns the number of data columns held by each item.
Public Property Get ColumnCount() As Long
    ColumnCount = mlColumnCount
End Property

'Changes the number of columns and resizes every existing item's storage.
Public Property Let ColumnCount(ByVal lCount As Long)
    Dim oItem As clsFormattedListItem

    If lCount < 1 Then Err.Raise 5, TypeName(Me), "ColumnCount must be at least 1"
    If lCount = mlColumnCount Then Exit Property
    mlColumnCount = lCount
    ReDim msngColumnWidths(0 To mlColumnCount - 1)
    ParseColumnWidths
    For Each oItem In mcolItems
        oItem.SetColumnCount mlColumnCount
    Next oItem
    If mlBoundColumn > mlColumnCount Then mlBoundColumn = mlColumnCount
    RequestRefresh
End Property

'Returns the semicolon-separated column-width setting.
Public Property Get ColumnWidths() As String
    ColumnWidths = msColumnWidths
End Property

'Sets column widths in points; use zero to keep a data column hidden.
Public Property Let ColumnWidths(ByVal sWidths As String)
    msColumnWidths = sWidths
    ParseColumnWidths
    RequestRefresh
End Property

'Returns one parsed column width to an item while it draws itself.
Friend Property Get ColumnWidth(ByVal lColumn As Long) As Single
    If lColumn < 0 Or lColumn >= mlColumnCount Then Err.Raise 9, TypeName(Me), "Column index out of range"
    ColumnWidth = msngColumnWidths(lColumn)
End Property

'Returns the 1-based column used by the Value property.
Public Property Get BoundColumn() As Long
    BoundColumn = mlBoundColumn
End Property

'Chooses the 1-based column returned by Value; zero returns the row number.
Public Property Let BoundColumn(ByVal lColumn As Long)
    If lColumn < 0 Or lColumn > mlColumnCount Then Err.Raise 5, TypeName(Me), "BoundColumn is outside the available columns"
    mlBoundColumn = lColumn
End Property

'Returns the zero-based current row, or -1 when no row is current.
Public Property Get ListIndex() As Long
    ListIndex = mlListIndex
End Property

'Makes one row current, selects it and scrolls it into view.
Public Property Let ListIndex(ByVal lIndex As Long)
    If lIndex < -1 Or lIndex >= ListCount Then Err.Raise 9, TypeName(Me), "ListIndex out of range"
    SelectOnly lIndex
    mlListIndex = lIndex
    mlAnchorIndex = lIndex
    UpdateSelectionAppearance
    If lIndex >= 0 Then EnsureVisible lIndex
    RaiseEvent Change(lIndex)
End Property

'Returns the current MSForms single, multi or extended selection mode.
Public Property Get MultiSelect() As Long
    MultiSelect = mlMultiSelect
End Property

'Sets the MSForms selection mode and reduces the selection when switching to single.
Public Property Let MultiSelect(ByVal lMode As Long)
    If lMode < fmMultiSelectSingle Or lMode > fmMultiSelectExtended Then
        Err.Raise 5, TypeName(Me), "Invalid MultiSelect value"
    End If
    If mlMultiSelect <> lMode Then
        mlMultiSelect = lMode
        If lMode = fmMultiSelectSingle And mlListIndex >= 0 Then SelectOnly mlListIndex
        RequestRefresh
    End If
End Property

'Returns whether the requested zero-based row is selected.
Public Property Get Selected(ByVal lIndex As Long) As Boolean
    Selected = Item(lIndex).Selected
End Property

'Selects or clears one row while applying the current MultiSelect rules.
Public Property Let Selected(ByVal lIndex As Long, ByVal bSelected As Boolean)
    SetItemSelected lIndex, bSelected
End Property

'Returns the current row's bound value, or Null when there is no current row.
Public Property Get Value() As Variant
    If mlListIndex < 0 Then
        Value = Null
    ElseIf mlBoundColumn = 0 Then
        Value = mlListIndex
    Else
        Value = Item(mlListIndex).Value(mlBoundColumn - 1)
    End If
End Property

'Finds and selects the first row whose bound value matches the supplied value.
Public Property Let Value(ByVal vValue As Variant)
    Dim lColumn As Long
    Dim lIndex As Long

    If mlBoundColumn = 0 Then
        Me.ListIndex = CLng(vValue)
        Exit Property
    End If
    lColumn = mlBoundColumn - 1
    For lIndex = 0 To ListCount - 1
        If StrComp(CStr(Item(lIndex).Value(lColumn)), CStr(vValue), vbBinaryCompare) = 0 Then
            Me.ListIndex = lIndex
            Exit Property
        End If
    Next lIndex
    Me.ListIndex = -1
End Property

'Returns a value using zero-based row and column numbers, like an MSForms ListBox.
Public Property Get List(ByVal lRow As Long, Optional ByVal lColumn As Long = 0) As Variant
    List = Item(lRow).Value(lColumn)
End Property

'Stores a value using zero-based row and column numbers and redraws the list.
Public Property Let List(ByVal lRow As Long, Optional ByVal lColumn As Long = 0, ByVal vValue As Variant)
    Item(lRow).SetText lColumn, CStr(vValue)
End Property

'Returns an item object so its Text, Tag, Selected or Characters properties can be used.
Public Property Get Item(ByVal lIndex As Long) As clsFormattedListItem
    If lIndex < 0 Or lIndex >= ListCount Then Err.Raise 9, TypeName(Me), "Item index out of range"
    Set Item = mcolItems(lIndex + 1)
End Property

'Adds a row at the end or at a requested zero-based position and returns the new item.
Public Function AddItem(Optional ByVal vItem As Variant, Optional ByVal vIndex As Variant) As clsFormattedListItem
    Dim oItem As clsFormattedListItem
    Dim lIndex As Long

    Set oItem = New clsFormattedListItem
    If IsMissing(vIndex) Then
        lIndex = ListCount
        mcolItems.Add oItem
    Else
        lIndex = CLng(vIndex)
        If lIndex < 0 Or lIndex > ListCount Then Err.Raise 9, TypeName(Me), "Item index out of range"
        If lIndex = ListCount Then
            mcolItems.Add oItem
        Else
            mcolItems.Add oItem, Before:=lIndex + 1
        End If
    End If
    oItem.Initialize Me, lIndex, mlColumnCount
    If Not IsMissing(vItem) Then oItem.SetText 0, CStr(vItem)
    ReindexItems
    RequestRefresh
    Set AddItem = oItem
End Function

'Removes one row, its controls and its selection state from the list.
Public Sub RemoveItem(ByVal lIndex As Long)
    Dim oItem As clsFormattedListItem

    Set oItem = Item(lIndex)
    oItem.Terminate
    mcolItems.Remove lIndex + 1
    ReindexItems
    If mlListIndex = lIndex Then mlListIndex = -1
    If mlListIndex > lIndex Then mlListIndex = mlListIndex - 1
    If mlAnchorIndex >= ListCount Then mlAnchorIndex = ListCount - 1
    RequestRefresh
End Sub

'Removes every item and all dynamic controls from the host Frame.
Public Sub Clear()
    Dim lIndex As Long
    Dim oItem As clsFormattedListItem

    If mcolItems Is Nothing Then Exit Sub
    For lIndex = mcolItems.Count To 1 Step -1
        Set oItem = mcolItems(lIndex)
        oItem.Terminate
        mcolItems.Remove lIndex
    Next lIndex
    mlListIndex = -1
    mlAnchorIndex = -1
    RequestRefresh
End Sub

'Postpones redraws while several list values or settings are being changed.
Public Sub BeginUpdate()
    mlUpdateDepth = mlUpdateDepth + 1
End Sub

'Ends one postponed-update level and performs the pending redraw when ready.
Public Sub EndUpdate()
    If mlUpdateDepth > 0 Then mlUpdateDepth = mlUpdateDepth - 1
    If mlUpdateDepth = 0 And mbRefreshPending Then Refresh
End Sub

'Rebuilds all visible row controls and recalculates their positions and scrollbars.
Public Sub Refresh()
    Dim oItem As clsFormattedListItem
    Dim sngTop As Single
    Dim lIndex As Long

    If moHostFrame Is Nothing Then Exit Sub
    If mlUpdateDepth > 0 Then
        mbRefreshPending = True
        Exit Sub
    End If

    ClearRenderedControls
    If ListCount > 0 Then
        ReDim msngItemTops(0 To ListCount - 1)
        ReDim msngItemHeights(0 To ListCount - 1)
    Else
        Erase msngItemTops
        Erase msngItemHeights
    End If

    sngTop = msngPadding
    lIndex = 0
    For Each oItem In mcolItems
        msngItemTops(lIndex) = sngTop
        msngItemHeights(lIndex) = oItem.Render(sngTop)
        sngTop = sngTop + msngItemHeights(lIndex)
        lIndex = lIndex + 1
    Next oItem
    SetScrollbars sngTop + msngPadding
    mbRefreshPending = False
End Sub

'Redraws the list after the host Frame changes size.
Public Sub Resize()
    If moHostFrame Is Nothing Then Exit Sub
    If Len(msColumnWidths) = 0 And mlColumnCount = 1 Then
        msngColumnWidths(0) = MaxSingle(20, moHostFrame.InsideWidth - 2 * msngPadding)
    End If
    Refresh
End Sub

'Moves keyboard focus to the current row, selecting the first row when necessary.
Public Sub SetFocus()
    If ListCount = 0 Then Exit Sub
    If mlListIndex < 0 Then Me.ListIndex = 0
    Item(mlListIndex).ActivateFocus
End Sub

'Adjusts vertical scrolling until the requested row is fully visible.
Public Sub EnsureVisible(ByVal lIndex As Long)
    Dim sngTop As Single
    Dim sngBottom As Single

    If moHostFrame Is Nothing Or lIndex < 0 Or lIndex >= ListCount Then Exit Sub
    If mlUpdateDepth > 0 Or mbRefreshPending Then Exit Sub
    sngTop = msngItemTops(lIndex)
    sngBottom = sngTop + msngItemHeights(lIndex)
    If sngTop < moHostFrame.ScrollTop Then
        moHostFrame.ScrollTop = sngTop
    ElseIf sngBottom > moHostFrame.ScrollTop + moHostFrame.InsideHeight Then
        moHostFrame.ScrollTop = MaxSingle(0, sngBottom - moHostFrame.InsideHeight)
    End If
End Sub

'Removes generated controls and breaks references so the UserForm can unload cleanly.
Public Sub Terminate()
    Clear
    Set mcolItems = Nothing
    Set moHostFrame = Nothing
End Sub

'Returns the minimum height used for each list row.
Public Property Get RowHeight() As Single
    RowHeight = msngRowHeight
End Property

'Sets the minimum row height; larger fonts can still make a row taller.
Public Property Let RowHeight(ByVal sngHeight As Single)
    If sngHeight <= 0 Then Err.Raise 5, TypeName(Me), "RowHeight must be greater than zero"
    msngRowHeight = sngHeight
    RequestRefresh
End Property

'Returns the space between the list content and the Frame edges.
Public Property Get Padding() As Single
    Padding = msngPadding
End Property

'Sets the space between the list content and the Frame edges.
Public Property Let Padding(ByVal sngPadding As Single)
    If sngPadding < 0 Then Err.Raise 5, TypeName(Me), "Padding cannot be negative"
    msngPadding = sngPadding
    RequestRefresh
End Property

'Returns the left-hand space inside each visible cell.
Public Property Get CellPadding() As Single
    CellPadding = msngCellPadding
End Property

'Sets the left-hand space inside each visible cell.
Public Property Let CellPadding(ByVal sngPadding As Single)
    If sngPadding < 0 Then Err.Raise 5, TypeName(Me), "CellPadding cannot be negative"
    msngCellPadding = sngPadding
    RequestRefresh
End Property

'Returns the normal row and Frame background color.
Public Property Get BackColor() As Long
    BackColor = mlBackColor
End Property

'Sets the normal row and Frame background color.
Public Property Let BackColor(ByVal lColor As Long)
    mlBackColor = lColor
    If Not moHostFrame Is Nothing Then moHostFrame.BackColor = lColor
    RequestRefresh
End Property

'Returns the default text color for unselected rows.
Public Property Get ForeColor() As Long
    ForeColor = mlForeColor
End Property

'Sets the default text color for unselected rows.
Public Property Let ForeColor(ByVal lColor As Long)
    mlForeColor = lColor
    RequestRefresh
End Property

'Returns the background color used for selected rows.
Public Property Get SelectedBackColor() As Long
    SelectedBackColor = mlSelectedBackColor
End Property

'Sets the background color used for selected rows.
Public Property Let SelectedBackColor(ByVal lColor As Long)
    mlSelectedBackColor = lColor
    RequestRefresh
End Property

'Returns the text color used for selected rows.
Public Property Get SelectedForeColor() As Long
    SelectedForeColor = mlSelectedForeColor
End Property

'Sets the text color used for selected rows.
Public Property Let SelectedForeColor(ByVal lColor As Long)
    mlSelectedForeColor = lColor
    RequestRefresh
End Property

'Returns the total drawable row width, never narrower than the host Frame.
Friend Property Get ContentWidth() As Single
    Dim lColumn As Long
    Dim sngWidth As Single

    For lColumn = 0 To mlColumnCount - 1
        sngWidth = sngWidth + msngColumnWidths(lColumn)
    Next lColumn
    If Not moHostFrame Is Nothing Then sngWidth = MaxSingle(sngWidth, moHostFrame.InsideWidth - 2 * msngPadding)
    ContentWidth = sngWidth
End Property

'Requests a redraw after an item changes its text or character formatting.
Friend Sub ItemChanged(ByVal oItem As clsFormattedListItem)
    RequestRefresh
End Sub

'Applies a selection requested through an item's Selected property and raises Change.
Friend Sub SetItemSelected(ByVal lIndex As Long, ByVal bSelected As Boolean)
    If mlMultiSelect = fmMultiSelectSingle And bSelected Then SelectOnly lIndex
    Item(lIndex).SelectedInternal = bSelected
    If bSelected Then mlListIndex = lIndex
    If Not bSelected And mlMultiSelect = fmMultiSelectSingle And mlListIndex = lIndex Then mlListIndex = -1
    UpdateSelectionAppearance
    RaiseEvent Change(mlListIndex)
End Sub

'Applies native-style click, Ctrl-click and Shift-click selection behavior.
Friend Sub ItemClick(ByVal oItem As clsFormattedListItem, ByVal iShift As Integer)
    Dim lIndex As Long

    lIndex = oItem.Index
    Select Case mlMultiSelect
    Case fmMultiSelectSingle
        SelectOnly lIndex
    Case fmMultiSelectMulti
        oItem.SelectedInternal = Not oItem.Selected
    Case fmMultiSelectExtended
        If (iShift And MC_SHIFT_MASK) <> 0 Then
            If mlAnchorIndex < 0 Then mlAnchorIndex = lIndex
            SelectRange mlAnchorIndex, lIndex, (iShift And MC_CTRL_MASK) <> 0
        ElseIf (iShift And MC_CTRL_MASK) <> 0 Then
            oItem.SelectedInternal = Not oItem.Selected
            mlAnchorIndex = lIndex
        Else
            SelectOnly lIndex
            mlAnchorIndex = lIndex
        End If
    End Select
    mlListIndex = lIndex
    UpdateSelectionAppearance
    EnsureVisible lIndex
    RaiseEvent Click(lIndex)
    RaiseEvent Change(lIndex)
End Sub

'Makes the double-clicked item current and raises the public DblClick event.
Friend Sub ItemDblClick(ByVal oItem As clsFormattedListItem)
    mlListIndex = oItem.Index
    RaiseEvent DblClick(mlListIndex)
End Sub

'Handles navigation and selection keys before raising the public KeyDown event.
Friend Sub ItemKeyDown(ByVal oItem As clsFormattedListItem, _
                       ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim lTarget As Long
    Dim lPageSize As Long
    Dim bMove As Boolean

    lTarget = oItem.Index
    lPageSize = MaxLong(1, Int(moHostFrame.InsideHeight / msngRowHeight) - 1)
    Select Case KeyCode
    Case vbKeyUp
        lTarget = MaxLong(0, lTarget - 1): bMove = True
    Case vbKeyDown
        lTarget = MinLong(ListCount - 1, lTarget + 1): bMove = True
    Case vbKeyHome
        lTarget = 0: bMove = True
    Case vbKeyEnd
        lTarget = ListCount - 1: bMove = True
    Case vbKeyPageUp
        lTarget = MaxLong(0, lTarget - lPageSize): bMove = True
    Case vbKeyPageDown
        lTarget = MinLong(ListCount - 1, lTarget + lPageSize): bMove = True
    Case vbKeySpace
        ApplyKeyboardSelection oItem.Index, Shift, True
    End Select

    If bMove Then
        mlListIndex = lTarget
        ApplyKeyboardSelection lTarget, Shift, False
        EnsureVisible lTarget
        Item(lTarget).ActivateFocus
        KeyCode = 0
    End If
    RaiseEvent KeyDown(mlListIndex, KeyCode, Shift)
End Sub

'Raises the public KeyUp event for the item that currently owns keyboard focus.
Friend Sub ItemKeyUp(ByVal oItem As clsFormattedListItem, _
                     ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    RaiseEvent KeyUp(oItem.Index, KeyCode, Shift)
End Sub

'Updates selected rows after keyboard movement or Space, including Ctrl and Shift.
Private Sub ApplyKeyboardSelection(ByVal lTarget As Long, ByVal iShift As Integer, _
                                   ByVal bSpace As Boolean)
    If mlMultiSelect = fmMultiSelectSingle Then
        SelectOnly lTarget
    ElseIf mlMultiSelect = fmMultiSelectMulti Then
        If bSpace Then Item(lTarget).SelectedInternal = Not Item(lTarget).Selected
    ElseIf (iShift And MC_SHIFT_MASK) <> 0 Then
        If mlAnchorIndex < 0 Then mlAnchorIndex = lTarget
        SelectRange mlAnchorIndex, lTarget, (iShift And MC_CTRL_MASK) <> 0
    ElseIf (iShift And MC_CTRL_MASK) <> 0 Then
        If bSpace Then Item(lTarget).SelectedInternal = Not Item(lTarget).Selected
    Else
        SelectOnly lTarget
        mlAnchorIndex = lTarget
    End If
    UpdateSelectionAppearance
    RaiseEvent Change(lTarget)
End Sub

'Clears the existing selection and leaves only the requested row selected.
Private Sub SelectOnly(ByVal lIndex As Long)
    Dim oItem As clsFormattedListItem

    For Each oItem In mcolItems
        oItem.SelectedInternal = (oItem.Index = lIndex)
    Next oItem
End Sub

'Selects every row between two indexes, optionally keeping earlier selections.
Private Sub SelectRange(ByVal lFirst As Long, ByVal lLast As Long, ByVal bAdd As Boolean)
    Dim oItem As clsFormattedListItem
    Dim lLow As Long
    Dim lHigh As Long

    lLow = MinLong(lFirst, lLast)
    lHigh = MaxLong(lFirst, lLast)
    For Each oItem In mcolItems
        If Not bAdd Then oItem.SelectedInternal = False
        If oItem.Index >= lLow And oItem.Index <= lHigh Then oItem.SelectedInternal = True
    Next oItem
End Sub

'Renumbers item objects after a row is inserted or removed.
Private Sub ReindexItems()
    Dim oItem As clsFormattedListItem
    Dim lIndex As Long

    For Each oItem In mcolItems
        oItem.Index = lIndex
        lIndex = lIndex + 1
    Next oItem
End Sub

'Converts the semicolon-separated width string into numeric widths for drawing.
Private Sub ParseColumnWidths()
    Dim vWidths As Variant
    Dim lColumn As Long
    Dim sWidth As String
    Dim sngDefault As Single

    ReDim msngColumnWidths(0 To mlColumnCount - 1)
    sngDefault = 100
    If Not moHostFrame Is Nothing Then sngDefault = MaxSingle(20, moHostFrame.InsideWidth / mlColumnCount)
    If Len(Trim$(msColumnWidths)) = 0 Then
        For lColumn = 0 To mlColumnCount - 1
            msngColumnWidths(lColumn) = sngDefault
        Next lColumn
        Exit Sub
    End If

    vWidths = Split(msColumnWidths, ";")
    For lColumn = 0 To mlColumnCount - 1
        If lColumn <= UBound(vWidths) Then
            sWidth = Trim$(CStr(vWidths(lColumn)))
            msngColumnWidths(lColumn) = CSng(Val(sWidth))
        Else
            msngColumnWidths(lColumn) = sngDefault
        End If
    Next lColumn
End Sub

'Sets the Frame's scroll area and enables only the scrollbars that are needed.
Private Sub SetScrollbars(ByVal sngContentHeight As Single)
    Dim lPass As Long
    Dim lBars As Long

    With moHostFrame
        .ScrollWidth = ContentWidth + 2 * msngPadding
        .ScrollHeight = sngContentHeight
        For lPass = 1 To 2
            lBars = fmScrollBarsNone
            If .InsideWidth < .ScrollWidth Then lBars = fmScrollBarsHorizontal
            If .InsideHeight < .ScrollHeight Then lBars = lBars + fmScrollBarsVertical
            If .ScrollBars = lBars Then Exit For
            .ScrollBars = lBars
        Next lPass
    End With
End Sub

'Asks every item to remove its current dynamic controls before a full redraw.
Private Sub ClearRenderedControls()
    Dim oItem As clsFormattedListItem

    If mcolItems Is Nothing Then Exit Sub
    For Each oItem In mcolItems
        oItem.RemoveControls
    Next oItem
End Sub

'Repaints row selection and focus without deleting controls during their own events.
Private Sub UpdateSelectionAppearance()
    Dim oItem As clsFormattedListItem

    If mlUpdateDepth > 0 Or mbRefreshPending Then
        RequestRefresh
        Exit Sub
    End If
    For Each oItem In mcolItems
        oItem.UpdateSelectionAppearance
    Next oItem
End Sub

'Redraws immediately or records a pending redraw while updates are postponed.
Private Sub RequestRefresh()
    If mlUpdateDepth > 0 Or moHostFrame Is Nothing Then
        mbRefreshPending = True
    Else
        Refresh
    End If
End Sub

'Returns the smaller of two Long values.
Private Function MinLong(ByVal lLeft As Long, ByVal lRight As Long) As Long
    If lLeft < lRight Then MinLong = lLeft Else MinLong = lRight
End Function

'Returns the larger of two Long values.
Private Function MaxLong(ByVal lLeft As Long, ByVal lRight As Long) As Long
    If lLeft > lRight Then MaxLong = lLeft Else MaxLong = lRight
End Function

'Returns the larger of two Single values used in layout and scrolling.
Private Function MaxSingle(ByVal sngLeft As Single, ByVal sngRight As Single) As Single
    If sngLeft > sngRight Then MaxSingle = sngLeft Else MaxSingle = sngRight
End Function

'Performs best-effort cleanup if the host did not call Terminate explicitly.
Private Sub Class_Terminate()
    On Error Resume Next
    Terminate
    On Error GoTo 0
End Sub


