Imports System.Configuration
Imports System.Configuration.Install
Imports System.Runtime.InteropServices
Imports System.IO
Public Class MSMQInstaller

    Inherits Installer

    <DllImport("kernel32")>
    Public Shared Function LoadLibrary(ByVal lpFileName As String) As IntPtr
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Public Shared Function FreeLibrary(ByVal hModule As IntPtr) As Boolean
    End Function

    Public Overrides Sub Install(stateSaver As IDictionary)
        MyBase.Install(stateSaver)

        Dim loaded As Boolean

        Try
            Dim handle As IntPtr = LoadLibrary("Mqrt.dll")

            If (handle = IntPtr.Zero Or handle.ToInt32() = 0) Then
                loaded = False
            Else
                loaded = True

                FreeLibrary(handle)
            End If
        Catch

            loaded = False
        End Try

        If loaded = False Then
                        {
            If Environment.OSVersion.Version.Major < 6 Then ' Windows Then XP Or earlier
                Dim fileName As String = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "MSMQAnswer.ans");

                Using writer As StreamWriter = New StreamWriter(fileName)
                    writer.WriteLine("[Version]")
                    writer.WriteLine("Signature = Windows NT")
                    writer.WriteLine()
                    writer.WriteLine("[Global]")
                    writer.WriteLine("FreshMode = Custom")
                    writer.WriteLine("MaintenanceMode = RemoveAll")
                    writer.WriteLine("UpgradeMode = UpgradeOnly")
                    writer.WriteLine()
                    writer.WriteLine("[Components]")
                    writer.WriteLine("msmq_Core = ON")
                    writer.WriteLine("msmq_LocalStorage = ON")
                End Using


                Using p As New System.Diagnostics.Process()
                    Dim start As ProcessStartInfo = New ProcessStartInfo("sysocmgr.exe", "/i:sysoc.inf /u:" + fileName + ""

                    p.StartInfo = start

                    p.Start();
                    p.WaitForExit();
                End Using
            Else  ' Vista Or later
                Using p As New System.Diagnostics.Process()
                    Dim start = New System.Diagnostics.ProcessStartInfo("ocsetup.exe", "MSMQ-Container;MSMQ-Server /passive")

                    p.StartInfo = start

                    p.Start()
                    p.WaitForExit()
                End Using
            End If
        End If
    End Sub

End Class
