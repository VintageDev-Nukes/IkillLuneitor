Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports System.Windows.Shapes
Imports System.Diagnostics
Imports System.Windows.Interop
Imports System.Runtime.InteropServices

Partial Public Class WebBrowserOverlay
    Inherits Window
    Private _placementTarget As FrameworkElement

    Public ReadOnly Property WebBrowser() As WebBrowser
        Get
            Return _wb
        End Get
    End Property

    Public Sub New(placementTarget As FrameworkElement)
        InitializeComponent()

        _placementTarget = placementTarget
        Dim owner__1 As Window = Window.GetWindow(placementTarget)
        Debug.Assert(owner__1 IsNot Nothing)

        'owner.SizeChanged += delegate { OnSizeLocationChanged(); };
        AddHandler owner__1.LocationChanged, AddressOf OnSizeLocationChanged
        AddHandler _placementTarget.SizeChanged, AddressOf OnSizeLocationChanged

        If owner__1.IsVisible Then
            Owner = owner__1
            Show()
        Else
            AddHandler Owner.IsVisibleChanged, Sub()
                                                   If Owner.IsVisible Then
                                                       Me.Owner = Owner
                                                       Me.Show()
                                                   End If
                                               End Sub

            'owner.LayoutUpdated += new EventHandler(OnOwnerLayoutUpdated);
        End If
    End Sub

    Private Shadows Sub Load() Handles MyBase.Loaded

        'Auto calibrate position
        OnSizeLocationChanged()

    End Sub

    Protected Overrides Sub OnClosing(e As System.ComponentModel.CancelEventArgs)
        MyBase.OnClosing(e)
        If Not e.Cancel Then
            ' Delayed call to avoid crash due to Window bug.
            Dispatcher.BeginInvoke(DirectCast(Sub() Owner.Close(), Action))
        End If
    End Sub

    Private Sub OnSizeLocationChanged()

        Dim offset As Point = _placementTarget.TranslatePoint(New Point(), Owner)
        Dim size As New Point(_placementTarget.ActualWidth, _placementTarget.ActualHeight)
        Dim hwndSource__1 As HwndSource = DirectCast(HwndSource.FromVisual(Owner), HwndSource)
        Dim ct As CompositionTarget = hwndSource__1.CompositionTarget

        offset = ct.TransformToDevice.Transform(offset)
        size = ct.TransformToDevice.Transform(size)

        Dim screenLocation As New Win32.POINT(offset)
        Win32.ClientToScreen(hwndSource__1.Handle, screenLocation)
        Dim screenSize As New Win32.POINT(size)

        Win32.MoveWindow(DirectCast(HwndSource.FromVisual(Me), HwndSource).Handle, screenLocation.X, screenLocation.Y, screenSize.X, screenSize.Y, True)

    End Sub

End Class
