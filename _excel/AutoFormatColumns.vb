Imports ExcelDna.Integration
Imports Microsoft.Office.Interop.Excel  ' XlLookAt, xlSearchOrder
Imports ExcelButter.Utils


Public Module AutoformatColumns
    ' Automatically format columns based on contained data and column name
    ' Keyboard Shortcut set in excel-AddIn.dna to Ctrl+Shift+A

    Const FORMAT_NUMBER_0_PLACES = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Const FORMAT_NUMBER_2_PLACES = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    Const FORMAT_NUMBER_PERCENT = "0.00%"
    Const FORMAT_DATE = "m/d/yy"
    Const FORMAT_TEXT = "@"

    '<ExcelFunction(Category:="AutoformatColumns", Description:="Sets some clean default formatting for numerics", IsMacroType:=True)> _
    Public Function AutoFormatColumns(<ExcelArgument(AllowReference:=True)> ByVal RangeToFormat As Object) As Double

        Dim c As Integer,
            rng As Object,
            sel As Object,
            fmt As String

        ExcelDnaUtil.Application.ScreenUpdating = False
        rng = ReferenceToRange(RangeToFormat)
        sel = ExcelDnaUtil.Application.Selection

        ' clean values from database that come into range as NULL
        rng.Select()
        sel.Replace(What:="NULL", Replacement:="", LookAt:=XlLookAt.xlPart, _
                    SearchOrder:=XlSearchOrder.xlByRows, MatchCase:=False, SearchFormat:=False, _
                    ReplaceFormat:=False)

        rng = ExcelDnaUtil.Application.ActiveSheet.UsedRange
        For c = 1 To rng.Columns.Count
            fmt = GuessFormat(rng, c)
            If fmt <> "" Then
                With rng.resize(rng.Cells(2, c), rng.Cells(rng.Rows.Count, c))
                    .NumberFormat = fmt
                End With
            End If
        Next c

        ' make headers bold to logical end of excel file
        With rng.Resize(1, -4161)
            .Font.Bold = True
        End With

        ' resize columns
        rng.Select()
        rng.EntireColumn.AutoFit()

        ExcelDnaUtil.Application.ScreenUpdating = False
        sel.Select()

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

    Private Function GuessFormat(rng As Range, col As Integer) As String
        Dim hdr As String,
            typ As Integer,
            r As Long,
            lo As Double,
            hi As Double,
            val As Object

        hdr = rng.Cells(1, col).Value
        val = rng.Cells(2, col).Value
        typ = GetDataType(rng, col)
        Select Case typ
            ' if this is numeric, use a format with decimal places based on range of values
            Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
                lo = val
                hi = val
                For r = 2 To rng.Rows.Count
                    val = rng.Cells(r, col).Value
                    If Not IsEmpty(val) Then
                        If val < lo Then lo = val
                        If val > hi Then hi = val
                    End If
                Next r
                If InStr(LCase(hdr), "pct") > 0 Or InStr(hdr, "%") > 0 Then
                    GuessFormat = FORMAT_NUMBER_PERCENT
                ElseIf hi - lo > 1000 Then
                    GuessFormat = FORMAT_NUMBER_0_PLACES
                Else
                    GuessFormat = FORMAT_NUMBER_2_PLACES
                End If
            Case vbDate
                ' format as m/d/y
                GuessFormat = FORMAT_DATE
            Case Else
                GuessFormat = FORMAT_TEXT
        End Select
    End Function

    Private Function GetDataType(rng As Range, col As Integer) As Integer
        Dim r As Long,
            val As Object,
            typ As Integer,
            isDateVal As Boolean

        ' Strings are the most general so if we ever find
        ' one, we just return that type. Dates are also
        ' specific values, so if we find that type we will
        ' return it. Otherwise we go through keeping track
        ' of whether the values are all valid numeric representations
        ' of dates and return a date if that is the case, otherwise
        ' we return a number.

        typ = vbEmpty
        isDateVal = True
        For r = 2 To rng.Rows.Count
            On Error Resume Next
            val = rng.Cells(r, col).Value
            If Err.Number <> 0 Then
                val = rng.Cells(r, col).Value2
            End If
            typ = VarType(val)
            If typ = vbString Or typ = vbDate Then
                GetDataType = typ
                Exit Function
            End If
            If typ = vbInteger Or typ = vbLong Or typ = vbDouble Or typ = vbCurrency Or typ = vbDecimal Then
                ' 1/1/2000 is 36526 and 1/1/2100 is 73051 so if all values are in this range
                ' assume it is a date
                isDateVal = isDateVal And val >= 36526 And val <= 73051
            End If
        Next r
        If isDateVal Then
            GetDataType = vbDate
        Else
            GetDataType = vbDouble
        End If
    End Function

End Module

