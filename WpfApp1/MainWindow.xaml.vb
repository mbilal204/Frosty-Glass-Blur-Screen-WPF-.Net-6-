Imports System.Windows.Interop
Imports System.Runtime.InteropServices
Imports System.Windows.Media.Effects
Imports System.Windows.Media
Imports AForge.Imaging.Filters


Class MainWindow
    Inherits Window

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub MainWindow_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
        ' Apply the "frosty window" effect using the Accent Policy API
        Dim windowHelper = New WindowInteropHelper(Me)
        Dim accent = New AccentPolicy()
        accent.AccentState = AccentState.ACCENT_ENABLE_BLURBEHIND
        Dim accentStructSize = Marshal.SizeOf(accent)
        Dim accentPtr = Marshal.AllocHGlobal(accentStructSize)
        Marshal.StructureToPtr(accent, accentPtr, False)

        Dim data = New WindowCompositionAttributeData()
        data.Attribute = WindowCompositionAttribute.WCA_ACCENT_POLICY
        data.SizeOfData = accentStructSize
        data.Data = accentPtr

        SetWindowCompositionAttribute(windowHelper.Handle, data)

        Marshal.FreeHGlobal(accentPtr)
    End Sub

    <DllImport("user32.dll")>
    Private Shared Function SetWindowCompositionAttribute(hwnd As IntPtr, ByRef data As WindowCompositionAttributeData) As Integer
    End Function

    ' ... (AccentState, AccentPolicy, WindowCompositionAttribute, SetWindowCompositionAttribute definitions)

End Class

Public Enum AccentState
    ACCENT_DISABLED = 0
    ACCENT_ENABLE_GRADIENT = 1
    ACCENT_ENABLE_TRANSPARENTGRADIENT = 2
    ACCENT_ENABLE_BLURBEHIND = 3
    ACCENT_ENABLE_ACRYLICBLURBEHIND = 4 ' Windows 10 April 2018 Update (1803) or newer
    ACCENT_ENABLE_HOSTBACKDROP = 5 ' Windows 10 April 2019 Update (1903) or newer
    ACCENT_INVALID_STATE = 6
End Enum

<StructLayout(LayoutKind.Sequential)>
Public Structure AccentPolicy
    Public AccentState As Integer
    Public AccentFlags As Integer
    Public GradientColor As Integer
    Public AnimationId As Integer
End Structure

<StructLayout(LayoutKind.Sequential)>
Public Structure WindowCompositionAttributeData
    Public Attribute As WindowCompositionAttribute
    Public Data As IntPtr
    Public SizeOfData As Integer
End Structure

Public Enum WindowCompositionAttribute
    WCA_ACCENT_POLICY = 19
End Enum
