VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormattedListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Stores one list row and creates the MSForms controls that display and operate it.
Private Const MC_FORMAT_COUNT As Long = 10
Private Const MC_FONT_NAME As Long = 0
Private Const MC_FONT_SIZE As Long = 1
Private Const MC_BOLD As Long = 2
Private Const MC_ITALIC As Long = 3
Private Const MC_UNDERLINE As Long = 4
Private Const MC_STRIKETHROUGH As Long = 5
Private Const MC_WEIGHT As Long = 6
Private Const MC_CHARSET As Long = 7
Private Const MC_FORE_COLOR As Long = 8
Private Const MC_BACK_COLOR As Long = 9

Private moList As clsFormattedListBox
Private mlIndex As Long
Private mbSelected As Boolean
Private mvTag As Variant
Private msValues() As String
Private mvFormats() As Variant
Private mlColumnCount As Long
Private mlFormatColumns As Long
Private mlFormatLength As Long
Private mcolControlNames As Collection
Private WithEvents moHitTarget As MSForms.Label
Attribute moHitTarget.VB_VarHelpID = -1
Private WithEvents moFocusProxy As MSForms.TextBox
Attribute moFocusProxy.VB_VarHelpID = -1
Private miMouseShift As Integer
Private msngTop As Single

'Connects a new item to its list and prepares storage for all columns.
Friend Sub Initialize(ByVal oList As clsFormattedListBox, ByVal lIndex As Long, _
                      ByVal lColumnCount As Long)
    Set moList = oList
    mlIndex = lIndex
    SetColumnCount lColumnCount
End Sub

'Returns this item's zero-based row number.
Public Property Get Index() As Long
    Index = mlIndex
End Property

'Updates the row number after items are inserted or removed.
Friend Property Let Index(ByVal lIndex As Long)
    mlIndex = lIndex
End Property

'Returns whether this item is part of the current selection.
Public Property Get Selected() As Boolean
    Selected = mbSelected
End Property

'Selects or clears this item through the parent list's selection rules.
Public Property Let Selected(ByVal bSelected As Boolean)
    If moList Is Nothing Then
        mbSelected = bSelected
    Else
        moList.SetItemSelected mlIndex, bSelected
    End If
End Property

'Changes the stored selection state without raising events or applying selection rules.
Friend Property Let SelectedInternal(ByVal bSelected As Boolean)
    mbSelected = bSelected
End Property

'Returns optional caller data associated with this item.
Public Property Get Tag() As Variant
    Tag = mvTag
End Property

'Stores optional caller data without displaying it in the list.
Public Property Let Tag(ByVal vTag As Variant)
    mvTag = vTag
End Property

'Returns the displayed text from the requested zero-based column.
Public Property Get Text(Optional ByVal lColumn As Long = 0) As String
    ValidateColumn lColumn
    Text = msValues(lColumn)
End Property

'Replaces a column's text and clears formatting that belonged to its old text.
Public Property Let Text(Optional ByVal lColumn As Long = 0, ByVal sText As String)
    SetText lColumn, sText
End Property

'Stores column text, resets that column's character formatting and requests a redraw.
Friend Sub SetText(ByVal lColumn As Long, ByVal sText As String)
    Dim lPosition As Long

    ValidateColumn lColumn
    msValues(lColumn) = sText
    EnsureFormatStorage mlColumnCount, MaxLong(1, Len(sText))
    For lPosition = 1 To mlFormatLength
        mvFormats(lColumn, lPosition) = Empty
    Next lPosition
    If Not moList Is Nothing Then moList.ItemChanged Me
End Sub

'Returns a column value to the parent list without exposing internal arrays.
Friend Property Get Value(ByVal lColumn As Long) As Variant
    ValidateColumn lColumn
    Value = msValues(lColumn)
End Property

'Returns an Excel-style object for formatting part or all of one cell's text.
Public Function Characters(Optional ByVal lStart As Long = 1, _
                           Optional ByVal lLength As Long = -1, _
                           Optional ByVal lColumn As Long = 0) As clsFormattedCharacters
    Dim oCharacters As clsFormattedCharacters

    ValidateRange lColumn, lStart, lLength
    If lLength = -1 Then lLength = Len(msValues(lColumn)) - lStart + 1
    Set oCharacters = New clsFormattedCharacters
    oCharacters.Initialize Me, lStart, lLength, lColumn
    Set Characters = oCharacters
End Function

'Returns one formatting property for a range, or Null when that property is mixed.
Friend Function FormatValue(ByVal lColumn As Long, ByVal lStart As Long, _
                            ByVal lLength As Long, ByVal sProperty As String) As Variant
    Dim lPosition As Long
    Dim lProperty As Long
    Dim vFirst As Variant
    Dim vCurrent As Variant

    ValidateRange lColumn, lStart, lLength
    lProperty = FormatPropertyIndex(sProperty)
    vFirst = CharacterFormatValue(lColumn, lStart, lProperty)
    For lPosition = lStart + 1 To lStart + lLength - 1
        vCurrent = CharacterFormatValue(lColumn, lPosition, lProperty)
        If Not ValuesEqual(vFirst, vCurrent) Then
            FormatValue = Null
            Exit Function
        End If
    Next lPosition
    FormatValue = vFirst
End Function

'Applies one formatting property to each character in the requested range.
Friend Sub SetFormatValue(ByVal lColumn As Long, ByVal lStart As Long, _
                          ByVal lLength As Long, ByVal sProperty As String, _
                          ByVal vValue As Variant)
    Dim lPosition As Long
    Dim lProperty As Long
    Dim vFormat As Variant

    ValidateRange lColumn, lStart, lLength
    lProperty = FormatPropertyIndex(sProperty)
    EnsureFormatStorage mlColumnCount, MaxLong(1, Len(msValues(lColumn)))
    For lPosition = lStart To lStart + lLength - 1
        vFormat = CharacterFormat(lColumn, lPosition)
        vFormat(lProperty) = vValue
        mvFormats(lColumn, lPosition) = vFormat
    Next lPosition
    If Not moList Is Nothing Then moList.ItemChanged Me
End Sub

'Clears all custom formatting from the requested character range.
Friend Sub ClearFormat(ByVal lColumn As Long, ByVal lStart As Long, ByVal lLength As Long)
    Dim lPosition As Long

    ValidateRange lColumn, lStart, lLength
    For lPosition = lStart To lStart + lLength - 1
        mvFormats(lColumn, lPosition) = Empty
    Next lPosition
    If Not moList Is Nothing Then moList.ItemChanged Me
End Sub

'Resizes this item's value and formatting storage to match the parent list.
Friend Sub SetColumnCount(ByVal lColumnCount As Long)
    Dim lOldCount As Long

    If lColumnCount < 1 Then Err.Raise 5, TypeName(Me), "ColumnCount must be at least 1"
    lOldCount = mlColumnCount
    mlColumnCount = lColumnCount
    If lOldCount = 0 Then
        ReDim msValues(0 To mlColumnCount - 1)
    ElseIf lOldCount <> mlColumnCount Then
        ReDim Preserve msValues(0 To mlColumnCount - 1)
    End If
    EnsureFormatStorage mlColumnCount, MaxLong(1, mlFormatLength)
End Sub

'Builds the row's background, formatted text labels and mouse surface in the host Frame.
Friend Function Render(ByVal sngTop As Single) As Single
    Dim oHost As MSForms.Frame
    Dim oBack As MSForms.Label
    Dim oRun As MSForms.Label
    Dim lColumn As Long
    Dim lStart As Long
    Dim lEnd As Long
    Dim sngColumnLeft As Single
    Dim sngColumnWidth As Single
    Dim sngRunLeft As Single
    Dim sngRowHeight As Single
    Dim sngMeasuredWidth As Single
    Dim sName As String
    Dim sText As String

    msngTop = sngTop
    RemoveControls
    Set oHost = moList.HostFrame
    Set mcolControlNames = New Collection
    sngRowHeight = moList.RowHeight

    sName = ControlName("Back")
    Set oBack = oHost.Controls.Add("Forms.Label.1", sName, False)
    TrackControl sName
    With oBack
        .Caption = vbNullString
        .Left = moList.Padding
        .Top = sngTop
        .Width = moList.ContentWidth
        .Height = sngRowHeight
        If mlIndex = moList.ListIndex Then
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = vbButtonShadow
        Else
            .BorderStyle = fmBorderStyleNone
        End If
        .BackStyle = fmBackStyleOpaque
        .BackColor = IIf(mbSelected, moList.SelectedBackColor, moList.BackColor)
        .Visible = True
        .ZOrder 1
    End With

    sngColumnLeft = moList.Padding
    For lColumn = 0 To mlColumnCount - 1
        sngColumnWidth = moList.ColumnWidth(lColumn)
        If sngColumnWidth > 0 Then
            sText = msValues(lColumn)
            sngRunLeft = sngColumnLeft + moList.CellPadding
            lStart = 1
            Do While lStart <= Len(sText)
                lEnd = lStart
                Do While lEnd < Len(sText)
                    If Not SameCharacterFormat(lColumn, lStart, lEnd + 1) Then Exit Do
                    lEnd = lEnd + 1
                Loop

                sName = ControlName("Run" & CStr(lColumn) & "_" & CStr(lStart))
                Set oRun = oHost.Controls.Add("Forms.Label.1", sName, False)
                TrackControl sName
                With oRun
                    .Caption = Mid$(sText, lStart, lEnd - lStart + 1)
                    .WordWrap = False
                    ApplyRunFormat oRun, lColumn, lStart
                    .AutoSize = True
                    .Left = sngRunLeft
                    .Top = sngTop
                    sngMeasuredWidth = .Width
                    If sngRunLeft + .Width > sngColumnLeft + sngColumnWidth Then
                        .Width = MaxSingle(0.1, sngColumnLeft + sngColumnWidth - sngRunLeft)
                        .AutoSize = False
                    End If
                    If .Height > sngRowHeight Then sngRowHeight = .Height
                    .Visible = True
                    .ZOrder 0
                End With
                sngRunLeft = sngRunLeft + sngMeasuredWidth
                If sngRunLeft >= sngColumnLeft + sngColumnWidth Then Exit Do
                lStart = lEnd + 1
            Loop
        End If
        sngColumnLeft = sngColumnLeft + sngColumnWidth
    Next lColumn

    oBack.Height = sngRowHeight
    CenterRunLabels sngTop, sngRowHeight
    CreateHitTarget sngTop, sngRowHeight
    Render = sngRowHeight
End Function

'Creates the hidden keyboard control when needed and moves focus into this item.
Friend Sub ActivateFocus()
    Dim oHost As MSForms.Frame
    Dim sName As String

    If moFocusProxy Is Nothing Then
        Set oHost = moList.HostFrame
        sName = ControlName("Focus")
        Set moFocusProxy = oHost.Controls.Add("Forms.TextBox.1", sName, False)
        TrackControl sName
        With moFocusProxy
            .Left = moList.Padding
            .Top = msngTop
            .Width = 0.1
            .Height = 0.1
            .BorderStyle = fmBorderStyleNone
            .Locked = True
            .TabStop = False
            .Visible = True
            .ZOrder 1
        End With
    End If
    ' Keep the proxy aligned with the row so SetFocus doesn't force the Frame to scroll to Top=0.
    moFocusProxy.Top = msngTop
    On Error Resume Next
    moFocusProxy.SetFocus
    On Error GoTo 0
End Sub

'Repaints selection colors and the current-row border without rebuilding the row.
Friend Sub UpdateSelectionAppearance()
    Dim oControl As MSForms.control
    Dim lColumn As Long
    Dim lPosition As Long
    Dim lRunMarker As Long
    Dim lIndex As Long
    Dim sName As String
    Dim sRun As String
    Dim vParts As Variant

    If mcolControlNames Is Nothing Then Exit Sub
    For lIndex = 1 To mcolControlNames.Count
        sName = CStr(mcolControlNames(lIndex))
        Set oControl = moList.HostFrame.Controls(sName)
        If Right$(sName, 5) = "_Back" Then
            oControl.BackColor = IIf(mbSelected, moList.SelectedBackColor, moList.BackColor)
            If mlIndex = moList.ListIndex Then
                oControl.BorderStyle = fmBorderStyleSingle
                oControl.BorderColor = vbButtonShadow
            Else
                oControl.BorderStyle = fmBorderStyleNone
            End If
        Else
            lRunMarker = InStr(1, sName, "_Run", vbBinaryCompare)
            If lRunMarker > 0 Then
                sRun = Mid$(sName, lRunMarker + 4)
                vParts = Split(sRun, "_")
                lColumn = CLng(vParts(0))
                lPosition = CLng(vParts(1))
                ApplyRunFormat oControl, lColumn, lPosition
            End If
        End If
    Next lIndex
End Sub

'Removes every dynamic control created for this item from the host Frame.
Friend Sub RemoveControls()
    Dim oHost As MSForms.Frame
    Dim lIndex As Long

    Set moHitTarget = Nothing
    Set moFocusProxy = Nothing
    If mcolControlNames Is Nothing Then Exit Sub
    If moList Is Nothing Then Exit Sub
    Set oHost = moList.HostFrame
    If oHost Is Nothing Then Exit Sub
    On Error Resume Next
    For lIndex = mcolControlNames.Count To 1 Step -1
        oHost.Controls.Remove CStr(mcolControlNames(lIndex))
    Next lIndex
    On Error GoTo 0
    Set mcolControlNames = Nothing
End Sub

'Removes the row controls and breaks the item's reference back to the parent list.
Friend Sub Terminate()
    RemoveControls
    Set moList = Nothing
End Sub

'Adds one transparent label over the row so all formatted runs act as one list item.
Private Sub CreateHitTarget(ByVal sngTop As Single, ByVal sngHeight As Single)
    Dim sName As String

    sName = ControlName("Hit")
    Set moHitTarget = moList.HostFrame.Controls.Add("Forms.Label.1", sName, False)
    TrackControl sName
    With moHitTarget
        .Caption = vbNullString
        .Left = moList.Padding
        .Top = sngTop
        .Width = moList.ContentWidth
        .Height = sngHeight
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .ControlTipText = msValues(0)
        .Visible = True
        .ZOrder 0
    End With
End Sub

'Applies the stored character style or the current selection colors to one run label.
Private Sub ApplyRunFormat(ByVal oRun As MSForms.Label, ByVal lColumn As Long, _
                           ByVal lPosition As Long)
    Dim vFormat As Variant

    vFormat = CharacterFormat(lColumn, lPosition)
    With oRun
        .Font.Name = moList.HostFrame.Font.Name
        .Font.Size = moList.HostFrame.Font.Size
        If Not IsEmpty(vFormat(MC_FONT_NAME)) Then .Font.Name = CStr(vFormat(MC_FONT_NAME))
        If Not IsEmpty(vFormat(MC_FONT_SIZE)) Then .Font.Size = CSng(vFormat(MC_FONT_SIZE))
        If Not IsEmpty(vFormat(MC_BOLD)) Then .Font.Bold = CBool(vFormat(MC_BOLD))
        If Not IsEmpty(vFormat(MC_ITALIC)) Then .Font.Italic = CBool(vFormat(MC_ITALIC))
        If Not IsEmpty(vFormat(MC_UNDERLINE)) Then .Font.Underline = CBool(vFormat(MC_UNDERLINE))
        If Not IsEmpty(vFormat(MC_STRIKETHROUGH)) Then .Font.Strikethrough = CBool(vFormat(MC_STRIKETHROUGH))
        If Not IsEmpty(vFormat(MC_WEIGHT)) Then .Font.Weight = CInt(vFormat(MC_WEIGHT))
        If Not IsEmpty(vFormat(MC_CHARSET)) Then .Font.Charset = CInt(vFormat(MC_CHARSET))
        .ForeColor = IIf(mbSelected, moList.SelectedForeColor, moList.ForeColor)
        If Not mbSelected And Not IsEmpty(vFormat(MC_FORE_COLOR)) Then .ForeColor = CLng(vFormat(MC_FORE_COLOR))
        If Not mbSelected And Not IsEmpty(vFormat(MC_BACK_COLOR)) Then
            .BackStyle = fmBackStyleOpaque
            .BackColor = CLng(vFormat(MC_BACK_COLOR))
        Else
            .BackStyle = fmBackStyleTransparent
        End If
    End With
End Sub

'Vertically centers every formatted text label within the final row height.
Private Sub CenterRunLabels(ByVal sngTop As Single, ByVal sngHeight As Single)
    Dim oControl As MSForms.control
    Dim lIndex As Long
    Dim sName As String

    For lIndex = 1 To mcolControlNames.Count
        sName = CStr(mcolControlNames(lIndex))
        If InStr(1, sName, "_Run", vbBinaryCompare) > 0 Then
            Set oControl = moList.HostFrame.Controls(sName)
            oControl.Top = sngTop + (sngHeight - oControl.Height) / 2
        End If
    Next lIndex
End Sub

'Returns the complete style array for one character, using empty defaults when unformatted.
Private Function CharacterFormat(ByVal lColumn As Long, ByVal lPosition As Long) As Variant
    Dim vFormat As Variant

    If IsEmpty(mvFormats(lColumn, lPosition)) Then
        vFormat = Array(Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty)
    Else
        vFormat = mvFormats(lColumn, lPosition)
    End If
    CharacterFormat = vFormat
End Function

'Returns one property from a character's complete style array.
Private Function CharacterFormatValue(ByVal lColumn As Long, ByVal lPosition As Long, _
                                      ByVal lProperty As Long) As Variant
    Dim vFormat As Variant

    vFormat = CharacterFormat(lColumn, lPosition)
    CharacterFormatValue = vFormat(lProperty)
End Function

'Checks whether two characters can be combined into the same display label.
Private Function SameCharacterFormat(ByVal lColumn As Long, ByVal lFirst As Long, _
                                     ByVal lSecond As Long) As Boolean
    Dim lProperty As Long
    Dim vFirst As Variant
    Dim vSecond As Variant

    vFirst = CharacterFormat(lColumn, lFirst)
    vSecond = CharacterFormat(lColumn, lSecond)
    For lProperty = 0 To MC_FORMAT_COUNT - 1
        If Not ValuesEqual(vFirst(lProperty), vSecond(lProperty)) Then Exit Function
    Next lProperty
    SameCharacterFormat = True
End Function

'Compares formatting values safely, including Empty and Null values.
Private Function ValuesEqual(ByVal vLeft As Variant, ByVal vRight As Variant) As Boolean
    If IsEmpty(vLeft) Or IsEmpty(vRight) Then
        ValuesEqual = IsEmpty(vLeft) And IsEmpty(vRight)
    ElseIf IsNull(vLeft) Or IsNull(vRight) Then
        ValuesEqual = IsNull(vLeft) And IsNull(vRight)
    Else
        ValuesEqual = (vLeft = vRight)
    End If
End Function

'Translates a public formatting property name to its position in the style array.
Private Function FormatPropertyIndex(ByVal sProperty As String) As Long
    Select Case LCase$(sProperty)
    Case "fontname": FormatPropertyIndex = MC_FONT_NAME
    Case "fontsize": FormatPropertyIndex = MC_FONT_SIZE
    Case "bold": FormatPropertyIndex = MC_BOLD
    Case "italic": FormatPropertyIndex = MC_ITALIC
    Case "underline": FormatPropertyIndex = MC_UNDERLINE
    Case "strikethrough": FormatPropertyIndex = MC_STRIKETHROUGH
    Case "weight": FormatPropertyIndex = MC_WEIGHT
    Case "charset": FormatPropertyIndex = MC_CHARSET
    Case "forecolor": FormatPropertyIndex = MC_FORE_COLOR
    Case "backcolor": FormatPropertyIndex = MC_BACK_COLOR
    Case Else: Err.Raise 5, TypeName(Me), "Unknown formatting property: " & sProperty
    End Select
End Function

'Grows character-format storage while preserving formatting already assigned.
Private Sub EnsureFormatStorage(ByVal lColumns As Long, ByVal lLength As Long)
    Dim vNewFormats() As Variant
    Dim lColumn As Long
    Dim lPosition As Long
    Dim lCopyColumns As Long
    Dim lCopyLength As Long

    If lColumns <= mlFormatColumns And lLength <= mlFormatLength And mlFormatLength > 0 Then Exit Sub
    If lLength < 1 Then lLength = 1
    ReDim vNewFormats(0 To lColumns - 1, 1 To lLength)
    If mlFormatLength > 0 Then
        lCopyColumns = MinLong(mlFormatColumns, lColumns)
        lCopyLength = MinLong(mlFormatLength, lLength)
        For lColumn = 0 To lCopyColumns - 1
            For lPosition = 1 To lCopyLength
                vNewFormats(lColumn, lPosition) = mvFormats(lColumn, lPosition)
            Next lPosition
        Next lColumn
    End If
    mvFormats = vNewFormats
    mlFormatColumns = lColumns
    mlFormatLength = lLength
End Sub

'Raises a clear error when a caller requests a column that does not exist.
Private Sub ValidateColumn(ByVal lColumn As Long)
    If lColumn < 0 Or lColumn >= mlColumnCount Then
        Err.Raise 9, TypeName(Me), "Column index out of range"
    End If
End Sub

'Raises a clear error when a character range falls outside the cell text.
Private Sub ValidateRange(ByVal lColumn As Long, ByVal lStart As Long, ByVal lLength As Long)
    ValidateColumn lColumn
    If lLength = -1 Then lLength = Len(msValues(lColumn)) - lStart + 1
    If lStart < 1 Or lLength < 1 Or lStart + lLength - 1 > Len(msValues(lColumn)) Then
        Err.Raise 5, TypeName(Me), "Character range is outside the cell text"
    End If
End Sub

'Records a generated control name so the item can remove it later.
Private Sub TrackControl(ByVal sName As String)
    If mcolControlNames Is Nothing Then Set mcolControlNames = New Collection
    mcolControlNames.Add sName
End Sub

'Creates a unique, readable name for one dynamic control in this row.
Private Function ControlName(ByVal sPart As String) As String
    ControlName = moList.ControlPrefix & "_" & CStr(mlIndex) & "_" & sPart
End Function

'Returns the smaller of two Long values.
Private Function MinLong(ByVal lLeft As Long, ByVal lRight As Long) As Long
    If lLeft < lRight Then MinLong = lLeft Else MinLong = lRight
End Function

'Returns the larger of two Long values.
Private Function MaxLong(ByVal lLeft As Long, ByVal lRight As Long) As Long
    If lLeft > lRight Then MaxLong = lLeft Else MaxLong = lRight
End Function

'Returns the larger of two Single values used in control layout.
Private Function MaxSingle(ByVal sngLeft As Single, ByVal sngRight As Single) As Single
    If sngLeft > sngRight Then MaxSingle = sngLeft Else MaxSingle = sngRight
End Function

'Remembers Ctrl and Shift before the following Click event selects the row.
Private Sub moHitTarget_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                  ByVal x As Single, ByVal y As Single)
    miMouseShift = Shift
End Sub

'Passes a row click and its modifier keys to the parent list.
Private Sub moHitTarget_Click()
    moList.ItemClick Me, miMouseShift
End Sub

'Passes a row double-click to the parent list.
Private Sub moHitTarget_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    moList.ItemDblClick Me
End Sub

'Passes keyboard navigation and selection keys to the parent list.
Private Sub moFocusProxy_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    moList.ItemKeyDown Me, KeyCode, Shift
End Sub

'Passes KeyUp to the parent list after its own keyboard handling finishes.
Private Sub moFocusProxy_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    moList.ItemKeyUp Me, KeyCode, Shift
End Sub

'Releases event controls and the parent reference if explicit cleanup was missed.
Private Sub Class_Terminate()
    Set moHitTarget = Nothing
    Set moFocusProxy = Nothing
    Set moList = Nothing
End Sub
