VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormattedCharacters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Represents one character range in one list cell, similar to Excel's Characters object.
Private moItem As clsFormattedListItem
Private mlStart As Long
Private mlLength As Long
Private mlColumn As Long

'Connects this range to its item, column, start position and length.
Friend Sub Initialize(ByVal oItem As clsFormattedListItem, ByVal lStart As Long, _
                      ByVal lLength As Long, ByVal lColumn As Long)
    Set moItem = oItem
    mlStart = lStart
    mlLength = lLength
    mlColumn = lColumn
End Sub

'Returns the font name, or Null when the range contains different font names.
Public Property Get FontName() As Variant
    FontName = moItem.FormatValue(mlColumn, mlStart, mlLength, "FontName")
End Property

'Applies one font name to every character in the range.
Public Property Let FontName(ByVal vValue As Variant)
    moItem.SetFormatValue mlColumn, mlStart, mlLength, "FontName", vValue
End Property

'Returns the font size, or Null when the range contains different sizes.
Public Property Get FontSize() As Variant
    FontSize = moItem.FormatValue(mlColumn, mlStart, mlLength, "FontSize")
End Property

'Applies one font size to every character in the range.
Public Property Let FontSize(ByVal vValue As Variant)
    moItem.SetFormatValue mlColumn, mlStart, mlLength, "FontSize", vValue
End Property

'Returns the bold setting, or Null when bold is mixed in the range.
Public Property Get Bold() As Variant
    Bold = moItem.FormatValue(mlColumn, mlStart, mlLength, "Bold")
End Property

'Makes every character in the range bold or not bold.
Public Property Let Bold(ByVal vValue As Variant)
    moItem.SetFormatValue mlColumn, mlStart, mlLength, "Bold", vValue
End Property

'Returns the italic setting, or Null when italic is mixed in the range.
Public Property Get Italic() As Variant
    Italic = moItem.FormatValue(mlColumn, mlStart, mlLength, "Italic")
End Property

'Makes every character in the range italic or not italic.
Public Property Let Italic(ByVal vValue As Variant)
    moItem.SetFormatValue mlColumn, mlStart, mlLength, "Italic", vValue
End Property

'Returns the underline setting, or Null when underline is mixed in the range.
Public Property Get Underline() As Variant
    Underline = moItem.FormatValue(mlColumn, mlStart, mlLength, "Underline")
End Property

'Underlines or removes the underline from every character in the range.
Public Property Let Underline(ByVal vValue As Variant)
    moItem.SetFormatValue mlColumn, mlStart, mlLength, "Underline", vValue
End Property

'Returns the strikethrough setting, or Null when it is mixed in the range.
Public Property Get Strikethrough() As Variant
    Strikethrough = moItem.FormatValue(mlColumn, mlStart, mlLength, "Strikethrough")
End Property

'Applies or removes strikethrough for every character in the range.
Public Property Let Strikethrough(ByVal vValue As Variant)
    moItem.SetFormatValue mlColumn, mlStart, mlLength, "Strikethrough", vValue
End Property

'Returns the numeric font weight, or Null when the range has mixed weights.
Public Property Get Weight() As Variant
    Weight = moItem.FormatValue(mlColumn, mlStart, mlLength, "Weight")
End Property

'Applies one numeric font weight to every character in the range.
Public Property Let Weight(ByVal vValue As Variant)
    moItem.SetFormatValue mlColumn, mlStart, mlLength, "Weight", vValue
End Property

'Returns the font character set, or Null when the range has mixed character sets.
Public Property Get Charset() As Variant
    Charset = moItem.FormatValue(mlColumn, mlStart, mlLength, "Charset")
End Property

'Applies one font character set to every character in the range.
Public Property Let Charset(ByVal vValue As Variant)
    moItem.SetFormatValue mlColumn, mlStart, mlLength, "Charset", vValue
End Property

'Returns the text color, or Null when the range contains different text colors.
Public Property Get ForeColor() As Variant
    ForeColor = moItem.FormatValue(mlColumn, mlStart, mlLength, "ForeColor")
End Property

'Applies one text color to every character in the range.
Public Property Let ForeColor(ByVal vValue As Variant)
    moItem.SetFormatValue mlColumn, mlStart, mlLength, "ForeColor", vValue
End Property

'Returns the character background color, or Null when the colors are mixed.
Public Property Get BackColor() As Variant
    BackColor = moItem.FormatValue(mlColumn, mlStart, mlLength, "BackColor")
End Property

'Applies one background color to every character in the range.
Public Property Let BackColor(ByVal vValue As Variant)
    moItem.SetFormatValue mlColumn, mlStart, mlLength, "BackColor", vValue
End Property

'Removes custom formatting from the range so the list defaults are used again.
Public Sub Clear()
    moItem.ClearFormat mlColumn, mlStart, mlLength
End Sub

'Releases the owning item when this temporary range object goes out of scope.
Private Sub Class_Terminate()
    Set moItem = Nothing
End Sub
