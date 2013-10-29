Imports ExcelDna.Integration
Imports Microsoft.Office.Interop.Excel
Imports ExcelButter.Utils


Public Module ToggleHighlightRow
    ' Toggle highlight the row of the current selection
    ' Keyboard Shortcut set in excel-AddIn.dna to Ctrl+Shift+H

    '<ExcelFunction(Category:="ToggleHighlightRow", Description:="Highlight the current row for easier exploration", IsMacroType:=True)> _
    Public Function ToggleHighlightRow(<ExcelArgument(AllowReference:=True)> ByVal RangeToHighlight As Object) As Double

        Dim rng As Object,
            sel As Object

        ExcelDnaUtil.Application.ScreenUpdating = False
        rng = ReferenceToRange(RangeToHighlight)
        sel = ExcelDnaUtil.Application.Selection

        rng.Select()
        With Selection.Interior
            If .Pattern = Constants.xlNone Then
                ' light green
                .Pattern = Constants.xlSolid
                .ThemeColor = XlThemeColor.xlThemeColorAccent3
                .TintAndShade = 0.599993896298105
            Else
                .Pattern = Constants.xlNone
                .TintAndShade = 0
            End If
        End With

        GoTo Exitproc

OnError:
        Select Case ErrorHandler()
            Case vbYes, vbRetry : Stop : Resume
            Case vbNo, vbIgnore : Resume Next
            Case Else : Resume Exitproc ' vbCancel
        End Select

Exitproc:
        On Error GoTo 0 ' restore any screenupdating etc
        ExcelDnaUtil.Application.Calculation = XlCalculation.xlCalculationAutomatic
        ExcelDnaUtil.Application.Cursor = XlMousePointer.xlDefault
        ExcelDnaUtil.Application.ScreenUpdating = True

    End Function

End Module
