Option Explicit On

Imports ExcelDna.Integration  ' IExcelAddin, ExcelDnaUtil, xlCall
Imports ExcelDna.Integration.CustomUI  ' needed for IRibbonControl
Imports ExcelDna.Integration.ExcelDnaUtil  ' Application etc. since no netoffice
Imports System.Runtime.InteropServices  ' needed for <ComVisible(True)> 
Imports System.IO


' This module contains <global> shortcuts to the Application members.
' via Patrick O'Beirne, http://www.sysmod.com/

Public Module Utils

    Public Function Array(ByVal ParamArray items() As Object) As Array
        Array = items
    End Function

    Public Function IsEmpty(ByVal p1 As Object) As Boolean
        Return IsNothing(p1)
    End Function

    Public Function IsNull(ByVal p1 As Object) As Boolean
        Return IsDBNull(p1)
    End Function

    Public Function IsObject(ByVal p1 As Object) As Boolean
        Return TypeOf (p1) Is Object
    End Function

    Public Function AppendToFile(ByVal FileName As String, ByVal txt As String) As String
        Dim ErrInfo As String = vbNullString
        SaveTextToFile(txt, FileName, ErrInfo, True)
        Return ErrInfo
    End Function

    Public Function GetFileContents(ByVal FullPath As String, _
            Optional ByRef ErrInfo As String = "") As String

        Dim strContents As String = vbNullString
        Dim objReader As StreamReader
        Try
            objReader = New StreamReader(FullPath)
            strContents = objReader.ReadToEnd()
            objReader.Close()
        Catch Ex As Exception
            ErrInfo = Ex.Message
        End Try
        Return strContents
    End Function

    Public Function SaveTextToFile(ByVal strData As String, _
            ByVal FullPath As String, _
            Optional ByVal ErrInfo As String = "", _
            Optional ByVal append As Boolean = False) As Boolean

        Dim bAns As Boolean = False
        Dim objReader As StreamWriter
        Try
            objReader = New StreamWriter(FullPath, append)
            objReader.Write(strData)
            objReader.Close()
            bAns = True
        Catch Ex As Exception
            ErrInfo = Ex.Message

        End Try
        Return bAns
    End Function

    Public Function Round(ByVal arg1 As Double, ByVal arg2 As Double) As Double
        Return Application.WorksheetFunction.Round(arg1, arg2)
    End Function

    Function QuotedName(ByVal sName As String) As String
        ' return a name properly quoted
        ' >> Dec'08 --> 'Dec''08'
        ' >> MyBudget --> 'My Budget'
        QuotedName = "'" & Replace(sName, "'", "'" & "'") & "'"
    End Function

    ' Used to get around NetOffice errors about "not a member of 'LateBindingApi.Core.COMObject'
    ' but 1.5.0 now allows .Parent etc

    ' wrapper to help get .Name etc
    Public Function ParentName(ByVal obj As Object) As String
        On Error Resume Next
        ParentName = vbNullString
        ParentName = obj.Parent.Name       ' Return
    End Function

    Public Function ObjectCount(ByVal obj As Object) As Long
        On Error Resume Next
        ObjectCount = 0
        ObjectCount = obj.count   ' Return
    End Function

    ' Get obj.collection(item) by name or Index posn to get around
    'Error	62	Class 'LateBindingApi.Core.COMObject' cannot be indexed because it has no default property.	
    Public Function ObjectItem(ByVal oOwner As Object, ByVal Item As Object) As Object
        ObjectItem = oOwner(Item)
    End Function

    ReadOnly Property Selection As Object
        Get
            Return Application.Selection
        End Get
    End Property

    Public Const vbRed As Long = 255

    Public Function ReferenceToRange(ByVal xlRef As ExcelReference) As Object
        'converts full range references "[TestFuncs.xlsx]Sheet1!$A$1:$A$6" to xlf C calls
        Dim strAddress As String = XlCall.Excel(XlCall.xlfReftext, xlRef, True)
        ReferenceToRange = ExcelDnaUtil.Application.Range(strAddress)
    End Function

    Public Function ErrorHandler()
        Dim sErrMsg As String
        sErrMsg = "Error " & Err.Number & IIf(Erl() = 0, "", " at line " & Erl()) & " " &
        Err.Description
        Debug.Print(sErrMsg)
        ErrorHandler = MsgBox(sErrMsg, vbAbortRetryIgnore, "Error")
    End Function

End Module

'The ExcelRibbon-derived class must also be marked as ComVisible(True),
' or in the project properties, advanced options, the ComVisible option must be checked.
' (Note that this is not the ‘Register for COM Interop’ option, which must never be used with Excel-DNA)

<ComVisible(True)> _
Public Class Ribbon
    Inherits ExcelRibbon
    'ExcelDna provides a feature to use onAction="RunTagMacro" which will run a VBA sub named in tag="MyVBAMacro"
    ' Only methods in this class are visible to onAction :

    Sub RunIDMacro(ByVal ctl As IRibbonControl)
        Application.Run(ctl.Id)
    End Sub

    Sub RunIDMacroWithTag(ByVal ctl As IRibbonControl)
        Application.Run(ctl.Id, ctl.Tag)
    End Sub

End Class
