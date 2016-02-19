# X-ile
## VBA Code
## fills data into specified number of buckets of appx. equal volume
## should not be used for small datasets or datasets with extreme outliers

Function xile(Decile_range As Range, Value As Double, Num_of_buckets As Integer)
    
    If Num_of_buckets < 1 Then
        xile = CVErr(xlErrValue)
    
    ElseIf Value = 0 Then
        xile = 0
    Else
        xile = WorksheetFunction.RoundUp(WorksheetFunction.SumIfs(Decile_range, ">0", Decile_range, "<=" & Value) / (WorksheetFunction.SumIfs(Decile_range, ">0") / Num_of_buckets), 0)
    End If
 
End Function
