Imports System.IO

Public Class Logger
    Private Shared ReadOnly LogFilePath As String = GetConfigValue("LogFilePath", "app.log")
    Private Shared ReadOnly EnableLogging As Boolean = GetConfigValue("EnableLogging", "true")

    Private Shared Function GetConfigValue(key As String, defaultValue As String) As String
        Try
            Dim value As String = ConfigurationManager.AppSettings(key)
            Return If(String.IsNullOrEmpty(value), defaultValue, value)
        Catch
            Return defaultValue
        End Try
    End Function

    Private Shared Function GetConfigValue(key As String, defaultValue As Boolean) As Boolean
        Try
            Dim value As String = ConfigurationManager.AppSettings(key)
            If String.IsNullOrEmpty(value) Then
                Return defaultValue
            End If
            Return Boolean.Parse(value)
        Catch
            Return defaultValue
        End Try
    End Function

    Public Shared Sub LogInfo(message As String)
        If EnableLogging Then
            WriteLog("INFO", message)
        End If
    End Sub

    Public Shared Sub LogWarning(message As String)
        If EnableLogging Then
            WriteLog("WARNING", message)
        End If
    End Sub

    Public Shared Sub LogError(message As String, Optional ex As Exception = Nothing)
        If EnableLogging Then
            Dim fullMessage As String = message
            If ex IsNot Nothing Then
                fullMessage &= $" | Exception: {ex.Message} | StackTrace: {ex.StackTrace}"
            End If
            WriteLog("ERROR", fullMessage)
        End If
    End Sub

    Private Shared Sub WriteLog(level As String, message As String)
        Try
            Dim logMessage As String = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] {message}"
            File.AppendAllText(LogFilePath, logMessage & Environment.NewLine)
        Catch
            ' Ignore logging errors to prevent application crashes
        End Try
    End Sub
End Class
