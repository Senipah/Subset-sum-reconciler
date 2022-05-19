Attribute VB_Name = "SubsetSum"
'@IgnoreModule ModuleWithoutFolder
'@IgnoreModule FunctionReturnValueDiscarded
Option Explicit

Private MemoizedPaths() As Boolean
Private IsFloat As Boolean

'@Description("Converts Ints to Floats and vice versa")
Private Function ConvertIfFloat(ByVal source As Variant, Optional ByVal Increase As Boolean) As Variant
Attribute ConvertIfFloat.VB_Description = "Converts Ints to Floats and vice versa"
    Const Order As Long = 2
    If IsFloat Then
        If Increase Then
            ConvertIfFloat = CLng(source * (10 ^ Order))
        Else
            ConvertIfFloat = source / (10 ^ Order)
        End If
    Else
        ConvertIfFloat = CLng(source)
    End If
End Function

'@Description("A recursive function to add all subsets")
Private Sub FindPaths( _
    ByRef Numbers() As Variant, _
    ByVal Index As Long, _
    ByVal Target As Long, _
    ByRef Paths As BetterArray, _
    Optional ByRef Path As BetterArray _
)
Attribute FindPaths.VB_Description = "A recursive function to add all subsets"
    If Path Is Nothing Then
        Set Path = New BetterArray
    End If
        
    ' If we reached end and sum is non-zero.
    ' We store Path only if numbers(0) is equal to Target OR MemoizedPaths(0, target) is true
    If Index = 0 And Target > 0 And MemoizedPaths(0, Target) Then
        Path.Push ConvertIfFloat(Numbers(Index))
        Path.Unshift "Combination " & Paths.Length + 1
        Paths.Push Path.Items
        Exit Sub
    End If

    ' If Target becomes 0
    If Index = 0 And Target = 0 Then
        Path.Unshift "Combination " & Paths.Length + 1
        Paths.Push Path.Items
        Exit Sub
    End If
    
    ' If given target can be achieved after ignoring current element.
    If MemoizedPaths(Index - 1, Target) Then
        ' Create a new array to store path
        Dim Fork As BetterArray
        If Path.Length Then
            Set Fork = Path.Clone
        Else
            Set Fork = New BetterArray
        End If
        FindPaths Numbers, Index - 1, Target, Paths, Fork
    End If
  
    ' If given sum can be achieved after considering current element.
    Dim NextIndex As Long
    NextIndex = Target - Numbers(Index)
    If NextIndex >= 0 Then
        If Target >= Numbers(Index) And MemoizedPaths(Index - 1, NextIndex) Then
            Path.Push ConvertIfFloat(Numbers(Index))
            FindPaths Numbers, Index - 1, NextIndex, Paths, Path
        End If
    End If
    
End Sub

'@Description("Populates the MemoizedPaths truth table array")
Private Sub MemoizePaths(ByRef Numbers() As Variant, ByRef Length As Long, ByVal Target As Long)
Attribute MemoizePaths.VB_Description = "Populates the MemoizedPaths truth table array"
    Dim i As Long
    Dim j As Long
    
    ReDim MemoizedPaths(Length, Target + 1)
    
    For i = 0 To Length
        ' Target 0 can always be achieved with 0 elements
        MemoizedPaths(i, 0) = True
    Next
    
    ' Target Numbers(0) can be achieved with single element
    If Numbers(0) <= Target Then
       MemoizedPaths(0, Numbers(0)) = True
    End If
    
    ' Fill rest of the entries in MemoizedPaths
    For i = 1 To Length - 1
        For j = 0 To Target
            If Numbers(i) <= j Then
                MemoizedPaths(i, j) = _
                    MemoizedPaths(i - 1, j) Or _
                    MemoizedPaths(i - 1, j - Numbers(i)) Or _
                    Numbers(i) = Target
            Else
                MemoizedPaths(i, j) = _
                    MemoizedPaths(i - 1, j) Or _
                    Numbers(i) = Target
            End If
        Next
    Next
End Sub

'@Description("Outputs the results to a new workbook")
Private Sub WritePaths(ByVal Paths As BetterArray, ByVal Target As Variant)
Attribute WritePaths.VB_Description = "Outputs the results to a new workbook"
    ' create new workbook to store results
    Dim DestFile As Workbook
    Dim DestSheet As Worksheet
    Dim FileName As String
    
    FileName = "Elements that sum to " & CStr(Target) & " - " & Format$(Now, "yyyy-mm-dd hh-mm-ss")
    Set DestFile = Workbooks.Add
    DestFile.SaveAs Environ$("temp") & Application.PathSeparator & FileName
    Set DestSheet = DestFile.Sheets.[_Default](1)
    DestSheet.Name = "Output"
    Paths.ToExcelRange DestSheet.Range("A1")
    DestSheet.Columns.AutoFit
    DestFile.Activate
End Sub

'@Description("Prints all subsets of numbers with sum 0.")
Private Sub FindAllSubsets(ByRef Numbers() As Variant, ByRef Length As Long, ByVal Target As Long)
Attribute FindAllSubsets.VB_Description = "Prints all subsets of numbers with sum 0."
    If Length = 0 Or Target < 0 Then
        Exit Sub
    End If
    
    MemoizePaths Numbers, Length, Target
    
    If MemoizedPaths(Length - 1, Target) = False Then
        MsgBox "There are no subsets with sum " & CStr(Target), vbCritical
        Exit Sub
    End If
        
    Dim Paths As BetterArray
    Set Paths = New BetterArray
    FindPaths Numbers, Length - 1, Target, Paths
    WritePaths Paths, ConvertIfFloat(Target)
End Sub

'@Description("Retrieves the data required for calculation from the worksheet")
Private Sub GetInputData(ByRef Numbers() As Variant, ByRef Target As Long)
Attribute GetInputData.VB_Description = "Retrieves the data required for calculation from the worksheet"
    Dim LocalValues As BetterArray
    Set LocalValues = New BetterArray
    
    LocalValues.LowerBound = 0
    
    With Sheet1
        IsFloat = CBool(.Range("IsFloat"))
        Target = ConvertIfFloat(.Range("Target"), True)
        LocalValues.FromExcelRange .Range("DataStart"), True
    End With
    
    ' convert values if necessary
    Dim i As Long
    For i = LocalValues.LowerBound To LocalValues.UpperBound
        LocalValues.Item(i) = ConvertIfFloat(LocalValues.Item(i), True)
    Next
    Numbers = LocalValues.Items
End Sub

'@Description("Entry point")
Public Sub Main()
Attribute Main.VB_Description = "Entry point"
    Dim Numbers() As Variant
    Dim Length As Long
    Dim Target As Long
    
    GetInputData Numbers, Target
    Length = UBound(Numbers) - LBound(Numbers) + 1
    FindAllSubsets Numbers, Length, Target
End Sub
