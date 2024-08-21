Attribute VB_Name = "F_semicolon"
Option Explicit

Function semicolon_create(first_cell As Range)
    Dim last_cell As Range
    Set last_cell = Cells(Rows.Count, first_cell.Resize(1, 1).Column).End(xlUp)
    Dim C_range As Range, c As Range
    If first_cell.Resize(1, 1).Address <> last_cell.Address Then
        Set C_range = Range(first_cell.Address & ":" & last_cell.Address)
    Else
        Set C_range = first_cell
    End If
    Dim i As Long
    Dim A(), buf As String
    ReDim A(C_range.Count - 1, 0)
    For Each c In C_range
        buf = c.Value
        Do While Left(buf, 1) = " "
            buf = Mid(buf, 2)
        Loop
        If buf <> "" And Left(buf, 1) <> "%" And Left(buf, 8) <> "function" Then
            A(i, 0) = c.Value & ";"
        ElseIf buf <> "" Then
            A(i, 0) = buf
        Else
            A(i, 0) = ""
        End If
        i = i + 1
    Next
    semicolon_create = A
End Function
Function semicolon_delete(first_cell As Range)
    Dim last_cell As Range
    Set last_cell = Cells(Rows.Count, first_cell.Resize(1, 1).Column).End(xlUp)
    Dim C_range As Range, c As Range
    If first_cell.Resize(1, 1).Address <> last_cell.Address Then
        Set C_range = Range(first_cell.Address & ":" & last_cell.Address)
    Else
        Set C_range = first_cell
    End If
    Dim i As Long
    Dim A()
    ReDim A(C_range.Count - 1, 0)
    For Each c In C_range
        If c.Value <> "" And Right(c.Value, 1) = ";" Then
            A(i, 0) = Left(c.Value, Len(c.Value) - 1)
        ElseIf c.Value <> "" Then
            A(i, 0) = c.Value
        Else
            A(i, 0) = ""
        End If
        i = i + 1
    Next
    semicolon_delete = A
End Function

