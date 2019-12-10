Imports System.Threading

Public Class TasksRunner

    Private RunningTasks As List(Of RunningTask) = New List(Of RunningTask)

    Public Function RunMethodEveryXMinutes(ByVal XMinutes As Integer, MethodToRun As Func(Of Task)) As ULong
        Dim TokenSource As CancellationTokenSource = New CancellationTokenSource()
        Dim NewTask = New RunningTask(Task.Run(action:=Async Sub() Await PeriodicAsyncTask(TimeSpan.FromMinutes(XMinutes), TokenSource.Token, MethodToRun)), Date.Now.Ticks, TokenSource)
        RunningTasks.Add(NewTask)
        Return NewTask.TaskID
    End Function

    Public Sub StopRunning(TaskID As ULong)
        Dim TheTask = RunningTasks.FirstOrDefault(Function(x) x.TaskID = TaskID)
        If (TheTask Is Nothing) Then
            Return
        End If
        TheTask.StopRunning()
        RunningTasks.Remove(TheTask)
    End Sub
    Public Sub StopRunningAll()
        If (RunningTasks.Count = 0) Then
            Return
        End If
        RunningTasks.ForEach(Sub(x) x.StopRunning())
        RunningTasks.Clear()
    End Sub

    Private Async Function PeriodicAsyncTask(ByVal interval As TimeSpan, ByVal cancellationToken As CancellationToken, MethodToRun As Func(Of Task)) As Task
        While True And Not cancellationToken.IsCancellationRequested
            Await MethodToRun()
            Await Task.Delay(interval)
        End While
    End Function


    Private Class RunningTask
        Private InnerTask As Task
        Public Property TaskID As ULong
        Public Property TokenSource As CancellationTokenSource

        Public Sub New(innerTask As Task, taskID As ULong, tokenSource As CancellationTokenSource)
            Me.InnerTask = innerTask
            Me.TaskID = taskID
            Me.TokenSource = tokenSource
        End Sub

        Public Sub StopRunning()
            Me.TokenSource.Cancel()
        End Sub

    End Class

End Class
