'PLACE IN A STANDARD MODULE
Sub LoopThroughRows()
Dim i As Long, lastrow As Long
Dim pctdone As Single
lastrow = Range("A" & Rows.Count).End(xlUp).Row

'(Step 1) Display your Progress Bar
ufProgress.LabelProgress.Width = 0
ufProgress.Show
For i = 1 To lastrow
'(Step 2) Periodically update progress bar
    pctdone = i / lastrow
    With ufProgress
        .LabelCaption.Caption = "Processing Row " & i & " of " & lastrow
        .LabelProgress.Width = pctdone * (.FrameProgress.Width)
    End With
    DoEvents
        '--------------------------------------
        'the rest of your macro goes below here
        '
        '
        '--------------------------------------
'(Step 3) Close the progress bar when you're done
    If i = lastrow Then Unload ufProgress
Next i
End Sub


'ufProgress.LabelProgress.Width = 0
'ufProgress.Show

'pctdone = i / lastrow
'    With ufProgress
'        .LabelCaption.Caption = "Processing Row " & i & " of " & lastrow
'        .LabelProgress.Width = pctdone * (.FrameProgress.Width)
'    End With
'    DoEvents

'If i = lastrow Then Unload ufProgress
'@voir https://wellsr.com/vba/2017/excel/beautiful-vba-progress-bar-with-step-by-step-instructions/
