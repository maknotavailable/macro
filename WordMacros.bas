Attribute VB_Name = "NewMacros"
Sub ResizeAllTables()
    Dim oTbl As Table
    For Each oTbl In ActiveDocument.Tables
        oTbl.AutoFitBehavior wdAutoFitFixed
        With ActiveDocument.PageSetup
            oTbl.PreferredWidth = .PageWidth - .LeftMargin - .RightMargin
        End With
    Next oTbl
End Sub
Sub Resize4ColTables()
    Dim t As Table
    For Each t In ActiveDocument.Tables
        t.Select
        CurPage = Selection.Information(wdActiveEndPageNumber)
        If CurPage < 231 Then
            t.AutoFitBehavior wdAutoFitFixed
            With ActiveDocument.PageSetup
                t.PreferredWidth = .PageWidth - .LeftMargin - .RightMargin
            End With
            If t.Columns.Count = 4 Then
                t.Columns(1).Width = InchesToPoints(0.9)
                t.Columns(2).Width = InchesToPoints(2.87)
                t.Columns(3).Width = InchesToPoints(0.85)
                t.Columns(3).Width = InchesToPoints(1.69)
            End If
        End If
    Next t
End Sub
