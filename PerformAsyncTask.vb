Async Function PerformAsyncTask() As Task
    Console.WriteLine($"Before Await: Thread ID = {Thread.CurrentThread.ManagedThreadId}")

    ' ConfigureAwait(False) avoids switching back to the original context
    Await Task.Delay(1000).ConfigureAwait(False)

    Console.WriteLine($"After Await: Thread ID = {Thread.CurrentThread.ManagedThreadId}")
End Function
