Module PancakeSorting
    Sub Main()
        ' Define an unsorted array
        Dim pancakes As Integer() = {3, 6, 1, 5, 9, 8, 2}
        
        Console.WriteLine("Unsorted Pancakes: " & String.Join(", ", pancakes))
        
        ' Perform pancake sorting
        PancakeSort(pancakes)
        
        Console.WriteLine("Sorted Pancakes: " & String.Join(", ", pancakes))
    End Sub

    ' Pancake sort implementation
    Sub PancakeSort(ByRef arr As Integer())
        Dim n As Integer = arr.Length
        
        ' Start sorting from the last element down to the first
        For size As Integer = n To 2 Step -1
            ' Find the largest pancake in the current range
            Dim maxIndex As Integer = FindMaxIndex(arr, size)
            
            ' If the largest pancake is not already in its correct position
            If maxIndex <> size - 1 Then
                ' Flip it to the top if it's not already there
                If maxIndex > 0 Then
                    Flip(arr, maxIndex)
                End If
                
                ' Flip it to its correct position
                Flip(arr, size - 1)
            End If
        Next
    End Sub

    ' Find the index of the largest element in the range 0 to size - 1
    Function FindMaxIndex(ByVal arr As Integer(), ByVal size As Integer) As Integer
        Dim maxIndex As Integer = 0
        For i As Integer = 1 To size - 1
            If arr(i) > arr(maxIndex) Then
                maxIndex = i
            End If
        Next
        Return maxIndex
    End Function

    ' Flip the array up to the specified index
    Sub Flip(ByRef arr As Integer(), ByVal k As Integer)
        Dim start As Integer = 0
        Dim temp As Integer
        
        While start < k
            temp = arr(start)
            arr(start) = arr(k)
            arr(k) = temp
            start += 1
            k -= 1
        End While
    End Sub
End Module
