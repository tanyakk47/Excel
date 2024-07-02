Attribute VB_Name = "Module1"
Function SumTwoNumbers(myArray() As Variant) As String
    Dim i As Integer
    Dim result As String
    result = 0
    For i = LBound(myArray) To UBound(myArray)
        result = result + myArray(i) * 10 ^ (i - 1)
    Next i
    SumTwoNumbers = result
End Function
Function MultTwoNumbers(ByVal num1 As String, ByVal num2 As String) As String
    Dim result As String
    If Len(num2) = 1 Then
        result = num1 * num2
    End If
    MultTwoNumbers = result
End Function

Function Mult(ByVal num1 As String, ByVal num2 As String) As String
    Dim length As Integer
    length = Len(num2)
    Dim myArray() As Variant
    ReDim myArray(1 To length)
    Dim Iteration As Integer
    For Iteration = 1 To Len(num2) Step 1
        myArray(Iteration) = MultTwoNumbers(num1, Mid(num2, length - Iteration + 1, 1))
    Next
    Mult = SumTwoNumbers(myArray())
End Function


