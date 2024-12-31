Imports System.Threading
Imports System.Threading.Tasks

Module ContextSwitchExample
    Sub Main()
        ' Start an async operation
        PerformAsyncOperation().Wait()
    End Sub

    Async Function PerformAsyncOperation() As Task
        Console.WriteLine($"Before Await: Thread ID = {Thread.CurrentThread.ManagedThreadId}")

        ' Simulate asynchronous work
        Await Task.Delay(1000)

        Console.WriteLine($"After Await: Thread ID = {Thread.CurrentThread.ManagedThreadId}")
    End Function
End Module
