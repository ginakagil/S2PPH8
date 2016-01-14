Imports System.Runtime.InteropServices
Public Class WindowsApi

    <DllImport("kernel32.dll", EntryPoint:="WTSGetActiveConsoleSessionId", SetLastError:=True)> _
    Public Shared Function WTSGetActiveConsoleSessionId() As UInteger
    End Function

    <DllImport("Wtsapi32.dll", EntryPoint:="WTSQueryUserToken", SetLastError:=True)> _
    Public Shared Function WTSQueryUserToken(ByVal SessionId As UInteger, ByRef phToken As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("kernel32.dll", EntryPoint:="CloseHandle", SetLastError:=True)> _
    Public Shared Function CloseHandle(<InAttribute()> ByVal hObject As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("advapi32.dll", EntryPoint:="CreateProcessAsUserW", SetLastError:=True)> _
    Public Shared Function CreateProcessAsUser(<InAttribute()> ByVal hToken As IntPtr, _
                                                    <InAttribute(), MarshalAs(UnmanagedType.LPWStr)> ByVal lpApplicationName As String, _
                                                    ByVal lpCommandLine As System.IntPtr, _
                                                    <InAttribute()> ByVal lpProcessAttributes As IntPtr, _
                                                    <InAttribute()> ByVal lpThreadAttributes As IntPtr, _
                                                    <MarshalAs(UnmanagedType.Bool)> ByVal bInheritHandles As Boolean, _
                                                    ByVal dwCreationFlags As UInteger, _
                                                    <InAttribute()> ByVal lpEnvironment As IntPtr, _
                                                    <InAttribute(), MarshalAsAttribute(UnmanagedType.LPWStr)> ByVal lpCurrentDirectory As String, _
                                                    <InAttribute()> ByRef lpStartupInfo As STARTUPINFOW, _
                                                    <OutAttribute()> ByRef lpProcessInformation As PROCESS_INFORMATION) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure SECURITY_ATTRIBUTES
        Public nLength As UInteger
        Public lpSecurityDescriptor As IntPtr
        <MarshalAs(UnmanagedType.Bool)> _
        Public bInheritHandle As Boolean
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure STARTUPINFOW
        Public cb As UInteger
        <MarshalAs(UnmanagedType.LPWStr)> _
        Public lpReserved As String
        <MarshalAs(UnmanagedType.LPWStr)> _
        Public lpDesktop As String
        <MarshalAs(UnmanagedType.LPWStr)> _
        Public lpTitle As String
        Public dwX As UInteger
        Public dwY As UInteger
        Public dwXSize As UInteger
        Public dwYSize As UInteger
        Public dwXCountChars As UInteger
        Public dwYCountChars As UInteger
        Public dwFillAttribute As UInteger
        Public dwFlags As UInteger
        Public wShowWindow As UShort
        Public cbReserved2 As UShort
        Public lpReserved2 As IntPtr
        Public hStdInput As IntPtr
        Public hStdOutput As IntPtr
        Public hStdError As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure PROCESS_INFORMATION
        Public hProcess As IntPtr
        Public hThread As IntPtr
        Public dwProcessId As UInteger
        Public dwThreadId As UInteger
    End Structure

End Class
