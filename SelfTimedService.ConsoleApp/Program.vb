Imports System.Configuration
Imports System.IO
Imports System.Threading
Imports System.Timers

Module Module1

    Dim ConditionCounter = 1

    Sub Main()

        Dim Runner = New TasksRunner()


        Dim TimeInMinutes = Integer.Parse(ConfigurationManager.AppSettings("TimeInMinutes"))


        'Adding Tasks:
        Dim FirstTaskID = Runner.RunMethodEveryXMinutes(1, AddressOf ExampleMethodReadingFromFile)

        Dim SecondTaskID = Runner.RunMethodEveryXMinutes(2, AddressOf ExampleMethodDoingNothing)



        'Looping forever, checking if there's a reason to stop running a single or every task every 60 seconds.
        While True
            If (SomeConditionToStopAll()) Then
                Runner.StopRunningAll()
                Exit While
            End If
            If (SomeConditionToStopFirstTask()) Then
                Runner.StopRunning(FirstTaskID)
            End If
            Thread.Sleep(60000)
        End While


    End Sub

    Function SomeConditionToStopAll() As Boolean
        ConditionCounter = ConditionCounter + 1
        Return ConditionCounter > 10
    End Function

    Function SomeConditionToStopFirstTask() As Boolean
        ConditionCounter = ConditionCounter + 1
        Return ConditionCounter > 5
    End Function


    Async Function ExampleMethodReadingFromFile() As Task

        Dim Mailing = New MailHelper("SomeMail@gmail.com", "SomePassword", New String() {"SomeOther@Mail.com"})

        Using reader = File.OpenText("SomeTextFile.txt")
            Dim Result = Await reader.ReadToEndAsync()
            If (Result = "Alert") Then
                Mailing.Send("Alert.")
            End If
        End Using

    End Function

    Async Function ExampleMethodDoingNothing() As Task

        Await Task.Delay(100)

    End Function



End Module
