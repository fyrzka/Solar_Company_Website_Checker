Sub ExtractHyperlinks()
    Dim HL As Hyperlink
    For Each HL In ActiveSheet.Hyperlinks
        HL.Range.Offset(0, 1).Value = HL.Address
    Next
End Sub
