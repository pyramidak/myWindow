#Region " Window "

Class myWindow

#Region " PPI Screen Conversion "

    Public Shared Function PPItoPixel(PPI As Point, Always As Boolean) As Point
        If Application.Current.MainWindow Is Nothing Then Return New Point(0, 0)
        Dim transform As Matrix = PresentationSource.FromVisual(Application.Current.MainWindow).CompositionTarget.TransformToDevice
        If Always Then
            Return transform.Transform(PPI)
        Else
            Dim Height As Double = SystemParameters.PrimaryScreenHeight
            If Height = 720 Or Height = 768 Or Height = 800 Or Height = 900 Or Height = 1024 Or Height = 1050 Or Height = 1080 Or Height = 1200 Or Height = 1440 Or Height = 1600 Or Height = 2160 Or Height = 4320 Then
                Return PPI
            Else
                Return transform.Transform(PPI)
            End If
        End If
    End Function

    Public Shared Function PixelToPPI(PIX As Point, Always As Boolean) As Point
        Dim transform As Matrix = PresentationSource.FromVisual(Application.Current.MainWindow).CompositionTarget.TransformFromDevice
        If Always Then
            Return transform.Transform(PIX)
        Else
            Dim Height As Double = SystemParameters.PrimaryScreenHeight
            If Height = 720 Or Height = 768 Or Height = 800 Or Height = 900 Or Height = 1024 Or Height = 1050 Or Height = 1080 Or Height = 1200 Or Height = 1440 Or Height = 1600 Or Height = 2160 Or Height = 4320 Then
                Return transform.Transform(PIX)
            Else
                Return PIX
            End If
        End If
    End Function

#End Region

#Region " Mouse Position "

    <DllImport("user32.dll")>
    Private Shared Function GetCursorPos(ByRef pt As Win32Point) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Private Structure Win32Point
        Public X As Int32
        Public Y As Int32
    End Structure

    Public Shared Function GetMousePosition() As System.Windows.Point
        Dim w32Mouse As New Win32Point()
        GetCursorPos(w32Mouse)
        Return New System.Windows.Point(w32Mouse.X, w32Mouse.Y)
    End Function

#End Region

#Region " Do Events "

    Public Shared Sub DoEvents()
        WaitForPriority(System.Windows.Threading.DispatcherPriority.Background)
    End Sub

    Private Shared Sub WaitForPriority(ByVal priority As System.Windows.Threading.DispatcherPriority)
        Dim frame As New System.Windows.Threading.DispatcherFrame()
        Dim dispatcherOperation As System.Windows.Threading.DispatcherOperation = System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(priority, New System.Windows.Threading.DispatcherOperationCallback(AddressOf ExitFrameOperation), frame)
        System.Windows.Threading.Dispatcher.PushFrame(frame)
        If dispatcherOperation.Status <> System.Windows.Threading.DispatcherOperationStatus.Completed Then
            dispatcherOperation.Abort()
        End If
    End Sub

    Private Shared Function ExitFrameOperation(ByVal obj As Object) As Object
        DirectCast(obj, System.Windows.Threading.DispatcherFrame).Continue = False
        Return Nothing
    End Function
#End Region

#Region " Move Form "

    Public Declare Sub ReleaseCapture Lib "User32" ()
    Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Int32) As Integer

    Public Shared Sub Drag(ByVal wnd As Window)
        Dim mainWindowPtr As IntPtr = New System.Windows.Interop.WindowInteropHelper(wnd).Handle
        Dim mainWindowSrc As System.Windows.Interop.HwndSource = System.Windows.Interop.HwndSource.FromHwnd(mainWindowPtr)
        SendMessage(CInt(mainWindowSrc.Handle), &HA1S, 2, 0) '&HA1S=WM_NCLBUTTONDOWN; 2=HTCAPTION
    End Sub
#End Region

End Class

#End Region