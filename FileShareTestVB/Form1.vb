Imports System.IO
Imports System.Runtime.CompilerServices.RuntimeHelpers
Imports System.Runtime.InteropServices
Imports System.Security.Principal

Public Class Form1
    Dim LOGON32_LOGON_INTERACTIVE As Integer = 2
    Dim LOGON32_PROVIDER_DEFAULT As Integer = 0
    Dim impersonationContext As WindowsImpersonationContext

    Declare Function LogonUserA Lib "advapi32.dll" (ByVal lpszUsername As String,
                        ByVal lpszDomain As String,
                        ByVal lpszPassword As String,
                        ByVal dwLogonType As Integer,
                        ByVal dwLogonProvider As Integer,
                        ByRef phToken As IntPtr) As Integer

    Declare Auto Function DuplicateToken Lib "advapi32.dll" (
                        ByVal ExistingTokenHandle As IntPtr,
                        ByVal ImpersonationLevel As Integer,
                        ByRef DuplicateTokenHandle As IntPtr) As Integer

    Declare Auto Function RevertToSelf Lib "advapi32.dll" () As Long
    Declare Auto Function CloseHandle Lib "kernel32.dll" (ByVal handle As IntPtr) As Long

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim userName As String = "<USERNAME>"   'Username Of the account To be impersonated
        Dim domain As String = "<DOMAIN>"       'Domain of the user account. Use localhost for local accounts
        Dim password As String = "<PASSWORD>"   'Account Password

        If impersonateValidUser(userName, domain, password) Then
            MessageBox.Show($"Successfully Impersonated as {WindowsIdentity.GetCurrent(False).Name}.{vbNewLine}Click OK to Copy the file.")

            'Insert your code that runs under the security context of a specific user here.
            Dim src As String = "<SOURCE FILE PATH>"            'Path of the file to be copied
            Dim dest As String = "<DESTINATION FILE PATH>"      'Destination path. For the Azure fileshare use the network path instead of the drive letter
            Try
                File.Copy(src, dest, True)
                MessageBox.Show($"Copied {src} to {dest} successfully.")

            Catch ex As Exception
                MessageBox.Show($"Failed to copy file with error: {ex}")
            End Try

            undoImpersonation()
        Else
            'Your impersonation failed. Therefore, include a fail-safe mechanism here.
            MessageBox.Show($"Failed to Impersonation user with error code: {Marshal.GetLastWin32Error()}")
        End If
    End Sub

    Private Function impersonateValidUser(ByVal userName As String, ByVal domain As String, ByVal password As String) As Boolean

        Dim tempWindowsIdentity As WindowsIdentity
        Dim token As IntPtr = IntPtr.Zero
        Dim tokenDuplicate As IntPtr = IntPtr.Zero
        impersonateValidUser = False

        If RevertToSelf() Then
            If LogonUserA(userName, domain, password, LOGON32_LOGON_INTERACTIVE,
                        LOGON32_PROVIDER_DEFAULT, token) <> 0 Then
                If DuplicateToken(token, 2, tokenDuplicate) <> 0 Then
                    tempWindowsIdentity = New WindowsIdentity(tokenDuplicate)
                    impersonationContext = tempWindowsIdentity.Impersonate()
                    If Not impersonationContext Is Nothing Then
                        impersonateValidUser = True
                    End If
                End If
            End If
        End If
        If Not tokenDuplicate.Equals(IntPtr.Zero) Then
            CloseHandle(tokenDuplicate)
        End If
        If Not token.Equals(IntPtr.Zero) Then
            CloseHandle(token)
        End If
    End Function

    Private Sub undoImpersonation()
        impersonationContext.Undo()
    End Sub
End Class
