Attribute VB_Name = "modDemo"
Option Explicit

Public Const GCSAPPNAME As String = "Formatted ListBox Demo"
Public Const GCSBUILD As String = "1.0"
Public gbCellsChanged As Boolean

Public Sub DemoNow()
    Dim ufForm As ufDemo
    Set ufForm = New ufDemo
    ufForm.AppName = GCSAPPNAME
    ufForm.Show
    Unload ufForm
    Set ufForm = Nothing
End Sub
