
Function xile(Decile_range As Range, Value As Double, Num_of_buckets As Integer, Optional Order As Boolean)


' Set default conditions

If IsMissing(Order) = True Then
        Order = "False"
        
End If

' Run conditions for errors and zero or negative volume values

If Num_of_buckets < 1 Then
        xile = CVErr(xlErrValue)
        Exit Function
    
    ElseIf Value <= 0 Then
        xile = 0
        Exit Function
        
End If


' Establish bucket size

Bucket_Size = WorksheetFunction.SumIfs(Decile_range, Decile_range, ">0") / Num_of_buckets

' Establish volume above selected value

Rolling_Size = WorksheetFunction.SumIfs(Decile_range, Decile_range, ">0", Decile_range, ">=" & Value)

' Establish number of buckets above selected value

Bucket = Rolling_Size / Bucket_Size

' Bucket if Order is False

Bucket_False = Num_of_buckets - WorksheetFunction.RoundDown(Bucket, 0)

' Bucket if Order is True

Bucket_True = 1 + WorksheetFunction.RoundDown(Bucket, 0)


' Calculate if Order is False

If Order = False And Bucket >= Num_of_buckets Then
        xile = 1
        Exit Function
        
    ElseIf Order = False And Bucket < Num_of_buckets Then
        xile = Bucket_False
        Exit Function

'Calculate if Order is True
    
    ElseIf Order = True And Bucket >= Num_of_buckets Then
        xile = Num_of_buckets
        Exit Function
    
    Else
        xile = Bucket_True
        Exit Function
    
    End If


End Function



