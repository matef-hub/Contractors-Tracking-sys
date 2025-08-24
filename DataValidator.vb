Imports System.Text.RegularExpressions

Public Class DataValidator
    Private Shared ReadOnly EmailPattern As String = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    Private Shared ReadOnly PhonePattern As String = "^[\+]?[0-9\s\-\(\)]{10,}$"

    Public Shared Function ValidateContractorName(name As String) As ValidationResult
        If String.IsNullOrWhiteSpace(name) Then
            Return New ValidationResult(False, "اسم المقاول لا يمكن أن يكون فارغ")
        End If

        If name.Length < 2 Then
            Return New ValidationResult(False, "اسم المقاول يجب أن يكون على الأقل حرفين")
        End If

        If name.Length > 100 Then
            Return New ValidationResult(False, "اسم المقاول لا يمكن أن يتجاوز 100 حرف")
        End If

        Return New ValidationResult(True, "")
    End Function

    Public Shared Function ValidateWorkDescription(work As String) As ValidationResult
        If String.IsNullOrWhiteSpace(work) Then
            Return New ValidationResult(False, "وصف الأعمال لا يمكن أن يكون فارغ")
        End If

        If work.Length < 5 Then
            Return New ValidationResult(False, "وصف الأعمال يجب أن يكون على الأقل 5 أحرف")
        End If

        If work.Length > 500 Then
            Return New ValidationResult(False, "وصف الأعمال لا يمكن أن يتجاوز 500 حرف")
        End If

        Return New ValidationResult(True, "")
    End Function

    Public Shared Function ValidateDate(dateValue As DateTime) As ValidationResult
        If dateValue < DateTime.Now.AddYears(-50) Then
            Return New ValidationResult(False, "التاريخ لا يمكن أن يكون أكثر من 50 سنة في الماضي")
        End If

        If dateValue > DateTime.Now.AddYears(10) Then
            Return New ValidationResult(False, "التاريخ لا يمكن أن يكون أكثر من 10 سنوات في المستقبل")
        End If

        Return New ValidationResult(True, "")
    End Function

    Public Shared Function ValidateEmail(email As String) As ValidationResult
        If String.IsNullOrWhiteSpace(email) Then
            Return New ValidationResult(True, "") ' Email is optional
        End If

        If Not Regex.IsMatch(email, EmailPattern) Then
            Return New ValidationResult(False, "صيغة البريد الإلكتروني غير صحيحة")
        End If

        Return New ValidationResult(True, "")
    End Function

    Public Shared Function ValidatePhoneNumber(phone As String) As ValidationResult
        If String.IsNullOrWhiteSpace(phone) Then
            Return New ValidationResult(True, "") ' Phone is optional
        End If

        If Not Regex.IsMatch(phone, PhonePattern) Then
            Return New ValidationResult(False, "صيغة رقم الهاتف غير صحيحة")
        End If

        Return New ValidationResult(True, "")
    End Function
End Class

Public Class ValidationResult
    Public Property IsValid As Boolean
    Public Property ErrorMessage As String

    Public Sub New(isValid As Boolean, errorMessage As String)
        Me.IsValid = isValid
        Me.ErrorMessage = errorMessage
    End Sub
End Class
