
Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports System.Security.Permissions
Imports System.Threading

Namespace TeamAgile.Threading

''' <summary>
''' Replaces the standard BackgroundWorker Component in .NET 2.0 Winforms
''' To :
'''		- support the ability of aborting the thread the worker is using.
'''		- support fast propogation of ProgressChanged events
''' </summary>
''' <remarks></remarks>     
Public Class BackgroundWorkerEx
        Inherits Component

        ' Events
        Public Custom Event DoWork As DoWorkEventHandler
            AddHandler(ByVal value As DoWorkEventHandler)
                MyBase.Events.AddHandler(BackgroundWorkerEx.doWorkKey, value)
            End AddHandler

            RemoveHandler(ByVal value As DoWorkEventHandler)
                MyBase.Events.RemoveHandler(BackgroundWorkerEx.doWorkKey, value)
            End RemoveHandler

            RaiseEvent(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs)

            End RaiseEvent
        End Event
        Public Custom Event ProgressChanged As ProgressChangedEventHandler
            AddHandler(ByVal value As ProgressChangedEventHandler)
                MyBase.Events.AddHandler(BackgroundWorkerEx.progressChangedKey, value)
            End AddHandler

            RemoveHandler(ByVal value As ProgressChangedEventHandler)
                MyBase.Events.RemoveHandler(BackgroundWorkerEx.progressChangedKey, value)
            End RemoveHandler

            RaiseEvent(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs)

            End RaiseEvent
        End Event
        Public Custom Event RunWorkerCompleted As RunWorkerCompletedEventHandler
            AddHandler(ByVal value As RunWorkerCompletedEventHandler)
                MyBase.Events.AddHandler(BackgroundWorkerEx.runWorkerCompletedKey, value)
            End AddHandler

            RemoveHandler(ByVal value As RunWorkerCompletedEventHandler)
                MyBase.Events.RemoveHandler(BackgroundWorkerEx.runWorkerCompletedKey, value)
            End RemoveHandler

            RaiseEvent(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)

            End RaiseEvent
        End Event

        ' Fields
        Private asyncOperation As AsyncOperation
        Private canCancelWorker As Boolean = True
        Private mCancellationPending As Boolean
        Private Shared ReadOnly doWorkKey As Object
        Private isRunning As Boolean
        Private ReadOnly operationCompleted As SendOrPostCallback
        Private Shared ReadOnly progressChangedKey As Object
        Private ReadOnly progressReporter As SendOrPostCallback
        Private Shared ReadOnly runWorkerCompletedKey As Object
        Private ReadOnly threadStart As WorkerThreadStartDelegate
        Private mWorkerReportsProgress As Boolean = True
        Private mRunningThread As Thread
        ' Nested Types
        Private Delegate Sub WorkerThreadStartDelegate(ByVal argument As Object)


        Private threadSleepProgressMs As Integer = 10
        Public Property DelayAfterReportProgressMillisec() As Integer
            Get
                Return threadSleepProgressMs
            End Get
            Set(ByVal value As Integer)
                threadSleepProgressMs = value
            End Set
        End Property

        ' Methods
        Shared Sub New()
            BackgroundWorkerEx.doWorkKey = New Object
            BackgroundWorkerEx.runWorkerCompletedKey = New Object
            BackgroundWorkerEx.progressChangedKey = New Object
        End Sub

        Public Sub CancelImmediately()
            If (Not IsBusy) Or (mRunningThread Is Nothing) Then
                Return
            End If


            mRunningThread.Abort()
            'no need to catch a ThreadAbortException since we will 
            'be swallowing and resetting it directly inside the OnDoWork method

            'we use AndAlso to short circuit the evaluation of the second check
            'so if isrunning is false, the check after andalso will not happen
            If isRunning AndAlso (Me.asyncOperation IsNot Nothing) Then
                Dim workCompletedArgs As New RunWorkerCompletedEventArgs(Nothing, Nothing, True)
                Me.asyncOperation.PostOperationCompleted(Me.operationCompleted, workCompletedArgs)
            End If

        End Sub

        Public Sub New()
            Me.threadStart = New WorkerThreadStartDelegate(AddressOf Me.WorkerThreadStart)
            Me.operationCompleted = New SendOrPostCallback(AddressOf Me.AsyncOperationCompleted)
            Me.progressReporter = New SendOrPostCallback(AddressOf Me.ProgressReporterSub)
        End Sub

        Private Sub AsyncOperationCompleted(ByVal arg As Object)
            Me.isRunning = False
            Me.mCancellationPending = False
            Me.OnRunWorkerCompleted(CType(arg, RunWorkerCompletedEventArgs))
        End Sub

        Public Sub CancelAsync()
            If Not Me.WorkerSupportsCancellation Then
                Throw New InvalidOperationException("BackgroundWorker_WorkerDoesntSupportCancellation")
            End If
            Me.mCancellationPending = True
        End Sub


        Protected Overridable Sub OnDoWork(ByVal e As DoWorkEventArgs)
            mRunningThread = Thread.CurrentThread
            Dim doWorkDelegate As DoWorkEventHandler = CType(MyBase.Events.Item(BackgroundWorkerEx.doWorkKey), DoWorkEventHandler)
            If (Not doWorkDelegate Is Nothing) Then

                Try
                    doWorkDelegate.Invoke(Me, e)
                Catch ex As ThreadAbortException
                    'prevent the exception from propogating further in any other thread or caller
                    'if the code in the DoWork method (by the developer) catches exceptions, 
                    'it should also catch and ignore a threadabortexception. (but not do a resetabort)
                    Thread.ResetAbort()
                End Try

            End If
        End Sub

        Protected Overridable Sub OnProgressChanged(ByVal e As ProgressChangedEventArgs)
            Dim progressChangedDelegate As ProgressChangedEventHandler = CType(MyBase.Events.Item(BackgroundWorkerEx.progressChangedKey), ProgressChangedEventHandler)
            If (Not progressChangedDelegate Is Nothing) Then
                progressChangedDelegate.Invoke(Me, e)
            End If
        End Sub

        Protected Overridable Sub OnRunWorkerCompleted(ByVal e As RunWorkerCompletedEventArgs)
            Dim workCompletedDelegate As RunWorkerCompletedEventHandler = CType(MyBase.Events.Item(BackgroundWorkerEx.runWorkerCompletedKey), RunWorkerCompletedEventHandler)
            If (Not workCompletedDelegate Is Nothing) Then
                workCompletedDelegate.Invoke(Me, e)
            End If
        End Sub

        Private Sub ProgressReporterSub(ByVal arg As Object)
            Me.OnProgressChanged(CType(arg, ProgressChangedEventArgs))
        End Sub

        Public Sub ReportProgress(ByVal percentProgress As Integer)
            Me.ReportProgress(percentProgress, Nothing)
        End Sub

        Public Sub ReportProgress(ByVal percentProgress As Integer, ByVal userState As Object)
            If Not Me.WorkerReportsProgress Then
                Throw New InvalidOperationException("BackgroundWorker_WorkerDoesntReportProgress")
            End If
            Dim changeEventArgs As New ProgressChangedEventArgs(percentProgress, userState)
            If (Not Me.asyncOperation Is Nothing) Then
                Me.asyncOperation.Post(Me.progressReporter, changeEventArgs)
            Else
                Me.progressReporter.Invoke(changeEventArgs)
            End If
            Thread.Sleep(Me.threadSleepProgressMs)
        End Sub

        Public Sub RunWorkerAsync()
            Me.RunWorkerAsync(Nothing)
        End Sub

        Public Sub RunWorkerAsync(ByVal argument As Object)
            If Me.isRunning Then
                Throw New InvalidOperationException("BackgroundWorker_WorkerAlreadyRunning")
            End If
            Me.isRunning = True
            Me.mCancellationPending = False
            Me.asyncOperation = AsyncOperationManager.CreateOperation(Nothing)
            Me.threadStart.BeginInvoke(argument, Nothing, Nothing)
        End Sub

        Private Sub WorkerThreadStart(ByVal argument As Object)
            Dim userState As Object = Nothing
            Dim doWorkException As Exception = Nothing
            Dim cancel As Boolean = False
            Try
                Dim workEventArgs As New DoWorkEventArgs(argument)
                Me.OnDoWork(workEventArgs)
                If workEventArgs.Cancel Then
                    cancel = True
                Else
                    userState = workEventArgs.Result
                End If
            Catch ex As Exception
                doWorkException = ex
            End Try
            If Me.asyncOperation IsNot Nothing Then
                Dim workCompletedArgs As New RunWorkerCompletedEventArgs(userState, doWorkException, cancel)
                Me.asyncOperation.PostOperationCompleted(Me.operationCompleted, workCompletedArgs)
            End If
        End Sub


        ' Properties
        Public ReadOnly Property CancellationPending() As Boolean
            Get
                Return Me.mCancellationPending
            End Get
        End Property

        <Browsable(False)> _
        Public ReadOnly Property IsBusy() As Boolean
            Get
                Return Me.isRunning
            End Get
        End Property

        <DefaultValue(True)> _
        Public Property WorkerReportsProgress() As Boolean
            Get
                Return Me.mWorkerReportsProgress
            End Get
            Set(ByVal value As Boolean)
                Me.mWorkerReportsProgress = value
            End Set
        End Property

        <DefaultValue(True)> _
        Public Property WorkerSupportsCancellation() As Boolean
            Get
                Return Me.canCancelWorker
            End Get
            Set(ByVal value As Boolean)
                Me.canCancelWorker = value
            End Set
        End Property


    End Class
End Namespace


