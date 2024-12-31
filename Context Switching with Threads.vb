Imports System.Threading

Module ThreadExample
    Sub Main()
        Dim thread1 As New Thread(AddressOf Task1)
        Dim thread2 As New Thread(AddressOf Task2)

        thread1.Start()
        thread2.Start()

        thread1.Join()
        thread2.Join()
    End Sub

    Sub Task1()
        For i As Integer = 1 To 5
            Console.WriteLine($"Task1 - Count {i} on Thread {Thread.CurrentThread.ManagedThreadId}")
            Thread.Sleep(500) ' Simulates work
        Next
    End Sub

    Sub Task2()
        For i As Integer = 1 To 5
            Console.WriteLine($"Task2 - Count {i} on Thread {Thread.CurrentThread.ManagedThreadId}")
            Thread.Sleep(500) ' Simulates work
        Next
    End Sub
End Module
