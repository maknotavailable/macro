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
