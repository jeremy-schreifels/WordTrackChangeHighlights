# WordTrackChangeHighlights
When making changes to a Microsoft Word document with track changes enabled, the deleted text is formatted with a different color and strikethrough and added text is foratted with a different color and underline. This VBA code, which can be added to the document as a macro, will format the deleted and added text with different color highlighting. 

# VBA code
Sub HighlightChanges()
'
' Highlight all deletions with grey highlight and all additions with yellow highlight
'

' Get current state of tracking changes -- on or off
tempState = ActiveDocument.TrackRevisions
Application.ScreenUpdating = False

' Turn off track changes
ActiveDocument.TrackRevisions = False
    
' Loop through changes and highlight deletions with grey highlighter
For Each Revision In ActiveDocument.Revisions
    If Revision.Type = wdRevisionDelete Then
        Set myRange = Revision.Range
        myRange.HighlightColorIndex = wdGray25
    ElseIf Revision.Type = wdRevisionInsert Then
        Set myRange = Revision.Range
        myRange.HighlightColorIndex = wdYellow
    Else
    End If
Next
    
Application.ScreenUpdating = True
ActiveDocument.TrackRevisions = tempState

End Sub
