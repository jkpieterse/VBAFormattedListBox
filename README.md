# FormattedListBox

(c)2026, Copyright JKP Application Development Services, all rights reserved

An all-VBA list box control for MSForms `UserForm`s that supports per-character text
formatting (bold, italic, underline, strikethrough, font, size, color) similar to
Excel's `Range.Characters` object — something the native `MSForms.ListBox` cannot do.

The control is built entirely from three class modules; there is no external
dependency, ActiveX control, or compiled binary. It renders each row using dynamically
created `Label` controls inside an `MSForms.Frame` that you provide.

For background and more information, see the accompanying article:
[jkp-ads.com/articles/vba-listbox-control.aspx](https://jkp-ads.com/articles/vba-listbox-control.aspx)

## Files

| File | Purpose |
|---|---|
| `clsFormattedListBox.cls` | The list control itself. Owns items, columns, selection, scrolling and keyboard/mouse handling. |
| `clsFormattedListItem.cls` | One row. Stores per-column text and per-character formatting, and renders its own controls. |
| `clsFormattedCharacters.cls` | An Excel-`Characters`-like object returned by `Item.Characters(...)` for reading/writing formatting on a character range. |
| `Formatted ListBox 1.0.xlsm` | Demo workbook. Run `modDemo.DemoNow` to see the control in action. |

## Getting started

1. Import `clsFormattedCharacters.cls`, `clsFormattedListItem.cls` and `clsFormattedListBox.cls`
   into your VBA project (all three are required).
2. Add an `MSForms.Frame` to a `UserForm` — this is where the list will be drawn.
3. Create the control and attach it to the frame:

```vb
Private WithEvents moList As clsFormattedListBox

Private Sub UserForm_Initialize()
    Set moList = New clsFormattedListBox
    Set moList.HostFrame = fraList
    moList.ColumnCount = 3
    moList.ColumnWidths = "100;60;40"
    moList.AddItem "First row"
End Sub
```

4. Resize the list when the frame resizes:

```vb
Private Sub UserForm_Resize()
    moList.Resize
End Sub
```

## Populating items

```vb
Dim oItem As clsFormattedListItem
Set oItem = moList.AddItem("Row text")   ' adds to column 0
oItem.Text(1) = "Second column"
oItem.Tag = 42                           ' any caller-defined data
```

Use `moList.BeginUpdate` / `moList.EndUpdate` around bulk changes to avoid redrawing
after every call.

## Formatting characters

`Item.Characters` returns an object with the same shape as Excel's `Characters`,
exposing `FontName`, `FontSize`, `Bold`, `Italic`, `Underline`, `Strikethrough`,
`Weight`, `Charset`, `ForeColor` and `BackColor`. Each property returns `Null` when the
requested range has mixed formatting.

```vb
With oItem.Characters(1, 5)     ' first 5 characters of column 0
    .Bold = True
    .ForeColor = vbRed
End With

oItem.Characters(, , 1).Italic = True   ' whole of column 1
oItem.Characters().Clear                ' remove all custom formatting from column 0
```

`Characters([Start], [Length], [Column])` — `Start` defaults to 1, `Length` defaults to
"rest of the text", `Column` defaults to 0 (zero-based).

## Key members of `clsFormattedListBox`

- **Items**: `AddItem`, `RemoveItem`, `Clear`, `Item(Index)`, `ListCount`
- **Columns**: `ColumnCount`, `ColumnWidths` (semicolon-separated points; `0` hides a
  column), `BoundColumn` (column returned by `Value`; `0` returns the row number)
- **Selection**: `ListIndex`, `Selected(Index)`, `Value`, `MultiSelect`
  (`fmMultiSelectSingle` / `MultiSelectMulti` / `MultiSelectExtended`)
- **Appearance**: `RowHeight`, `Padding`, `BackColor`, `ForeColor`, `SelectedBackColor`,
  `SelectedForeColor`
- **Lifecycle**: `BeginUpdate` / `EndUpdate`, `Refresh`, `Resize`, `SetFocus`,
  `EnsureVisible`, `Terminate` (call from `UserForm_Terminate` to release dynamic
  controls)
- **Events**: `Click`, `DblClick`, `Change`, `KeyDown`, `KeyUp` — all pass the
  zero-based row `Index`

## Demo

Open `Formatted ListBox 1.0.xlsm` and run `modDemo.DemoNow` to open a form that
creates a `clsFormattedListBox` inside an `MSForms.Frame`, fills it with eight
three-column rows, applies per-character bold and color formatting, and reports
selection changes. Use the buttons to repopulate, add, select, clear, or toggle bold
formatting. Double-click a row to see its `Tag` value.

## Limitations

- Designed for `UserForm`/MSForms hosts only (no ribbon task pane or worksheet
  support).
- Rows are rendered with real controls, so very large lists (thousands of rows) will
  be slower than a native `ListBox`.
