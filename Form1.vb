Imports System.IO
Imports System.Runtime.InteropServices

Public Class Form1

    ' Excel COM Objects
    Dim xlApp As Object ' Excel.Application
    Dim xlWorkbook As Object ' Excel.Workbook  
    Dim xlWorksheet As Object ' Excel.Worksheet

    ' Excel Constants
    Private Const xlUp As Integer = -4162
    Private Const xlLandscape As Integer = 2
    Private Const xlCenter As Integer = -4108
    Private Const xlContinuous As Integer = 1

    ' Configuration
    Dim filePath As String = Path.Combine(Application.StartupPath, "Contracts.xlsx")

    ' Contractor names history for combobox
    Private contractorHistory As New List(Of String)()
    Private Const MaxHistoryItems As Integer = 50


    ' --------------------------------------
    ' Safe Save Workbook Function
    ' --------------------------------------
    Private Sub TrySaveWorkbook()
        Try
            If xlWorkbook Is Nothing Then Return

            xlWorkbook.Save()
            SafeLogInfo("Workbook saved successfully")

        Catch ex As Exception
            SafeLogError("Error saving workbook", ex)

            ' Try SaveAs with a backup name if save fails
            Try
                Dim backupPath As String = Path.Combine(Path.GetDirectoryName(filePath),
                                                       Path.GetFileNameWithoutExtension(filePath) & "_backup_" &
                                                        DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx")
                xlWorkbook.SaveAs(backupPath)
                SafeLogWarning($"Saved as backup: {backupPath}")
                MessageBox.Show($"تم حفظ البيانات في ملف احتياطي: {Path.GetFileName(backupPath)}",
                               "تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Catch backupEx As Exception
                SafeLogError("Error creating backup", backupEx)
                Throw New Exception($"فشل في الحفظ: {ex.Message}")
            End Try
        End Try
    End Sub

    ' --------------------------------------
    ' Global Exception Handler
    ' --------------------------------------
    Private Sub Application_ThreadException(sender As Object, e As Threading.ThreadExceptionEventArgs)
        Try
            SafeLogError("Unhandled application exception", e.Exception)
            MessageBox.Show($"حدث خطأ غير متوقع في التطبيق:" & Environment.NewLine & e.Exception.Message,
                           "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch
            ' Last resort - show basic error
            MessageBox.Show("حدث خطأ خطير في التطبيق.", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' --------------------------------------
    ' Contractor History Management
    ' --------------------------------------
    Private Sub LoadContractorHistory()
        Try
            Dim historyFile As String = Path.Combine(Application.StartupPath, "contractor_history.txt")
            If File.Exists(historyFile) Then
                Dim lines() As String = File.ReadAllLines(historyFile)
                contractorHistory.Clear()
                For Each line As String In lines
                    If Not String.IsNullOrWhiteSpace(line) AndAlso Not contractorHistory.Contains(line.Trim()) Then
                        contractorHistory.Add(line.Trim())
                    End If
                Next
                SafeLogInfo($"Loaded {contractorHistory.Count} contractor names from history")
            End If
        Catch ex As Exception
            SafeLogWarning($"Failed to load contractor history: {ex.Message}")
        End Try
    End Sub

    Private Sub SaveContractorHistory()
        Try
            Dim historyFile As String = Path.Combine(Application.StartupPath, "contractor_history.txt")
            File.WriteAllLines(historyFile, contractorHistory.ToArray())
            SafeLogInfo($"Saved {contractorHistory.Count} contractor names to history")
        Catch ex As Exception
            SafeLogWarning($"Failed to save contractor history: {ex.Message}")
        End Try
    End Sub

    Private Sub AddContractorToHistory(contractorName As String)
        Try
            If String.IsNullOrWhiteSpace(contractorName) Then Return

            Dim trimmedName As String = contractorName.Trim()
            If String.IsNullOrWhiteSpace(trimmedName) Then Return

            ' Remove if already exists (to move to top)
            If contractorHistory.Contains(trimmedName) Then
                contractorHistory.Remove(trimmedName)
            End If

            ' Add to beginning of list
            contractorHistory.Insert(0, trimmedName)

            ' Keep only the most recent items
            If contractorHistory.Count > MaxHistoryItems Then
                contractorHistory.RemoveRange(MaxHistoryItems, contractorHistory.Count - MaxHistoryItems)
            End If

            ' Update combobox items
            UpdateContractorComboBox()

            ' Save to file
            SaveContractorHistory()

            SafeLogInfo($"Added contractor '{trimmedName}' to history")
        Catch ex As Exception
            SafeLogWarning($"Failed to add contractor to history: {ex.Message}")
        End Try
    End Sub

    Private Sub UpdateContractorComboBox()
        Try
            If cmbContractor Is Nothing Then Return

            ' Store current selection
            Dim currentText As String = cmbContractor.Text

            ' Clear and reload items
            cmbContractor.Items.Clear()
            cmbContractor.Items.AddRange(contractorHistory.ToArray())

            ' Restore current text if it exists
            If Not String.IsNullOrWhiteSpace(currentText) Then
                cmbContractor.Text = currentText
            End If
        Catch ex As Exception
            SafeLogWarning($"Failed to update contractor combobox: {ex.Message}")
        End Try
    End Sub

    Private Sub SetupContractorComboBoxAutoComplete()
        Try
            If cmbContractor Is Nothing Then Return

            ' Enable auto-complete features
            cmbContractor.AutoCompleteMode = AutoCompleteMode.SuggestAppend
            cmbContractor.AutoCompleteSource = AutoCompleteSource.ListItems
            cmbContractor.DropDownStyle = ComboBoxStyle.DropDown

            ' Add custom text changed event for better filtering
            AddHandler cmbContractor.TextChanged, AddressOf ContractorComboBox_TextChanged

            SafeLogInfo("Contractor combobox auto-complete configured")
        Catch ex As Exception
            SafeLogWarning($"Failed to setup contractor combobox auto-complete: {ex.Message}")
        End Try
    End Sub

    Private Sub ContractorComboBox_TextChanged(sender As Object, e As EventArgs)
        Try
            If cmbContractor Is Nothing Then Return

            Dim searchText As String = cmbContractor.Text.Trim().ToLower()

            ' If text is empty, show all items
            If String.IsNullOrEmpty(searchText) Then
                cmbContractor.DroppedDown = False
                Return
            End If

            ' Filter items based on search text
            Dim filteredItems As New List(Of String)()
            For Each item As String In contractorHistory
                If item.ToLower().Contains(searchText) Then
                    filteredItems.Add(item)
                End If
            Next

            ' Update dropdown items
            If filteredItems.Count > 0 Then
                cmbContractor.Items.Clear()
                cmbContractor.Items.AddRange(filteredItems.ToArray())
                cmbContractor.DroppedDown = True
            Else
                cmbContractor.DroppedDown = False
            End If

        Catch ex As Exception
            SafeLogWarning($"Error in contractor combobox text changed: {ex.Message}")
        End Try
    End Sub

    Private Sub LoadExistingContractorsFromExcel()
        Try
            If xlWorkbook Is Nothing Then Return

            Dim existingContractors As New List(Of String)()

            ' Search in all sheets for existing contractor names
            For Each sheet As Object In xlWorkbook.Sheets
                Try
                    Dim currentSheet As Object = sheet
                    Dim row As Integer = 2
                    While currentSheet.Cells(row, 1).Value IsNot Nothing
                        Dim contractor As String = GetCellValue(currentSheet.Cells(row, 2).Value)
                        If Not String.IsNullOrWhiteSpace(contractor) AndAlso Not existingContractors.Contains(contractor) Then
                            existingContractors.Add(contractor)
                        End If
                        row += 1
                    End While
                Catch ex As Exception
                    SafeLogWarning($"Error loading contractors from sheet {sheet.Name}: {ex.Message}")
                End Try
            Next

            ' Add existing contractors to history if not already present
            For Each contractor As String In existingContractors
                If Not contractorHistory.Contains(contractor) Then
                    contractorHistory.Add(contractor)
                End If
            Next

            SafeLogInfo($"Loaded {existingContractors.Count} existing contractor names from Excel")
        Catch ex As Exception
            SafeLogWarning($"Failed to load existing contractors from Excel: {ex.Message}")
        End Try
    End Sub

    Private Sub SetupContractorComboBoxContextMenu()
        Try
            If cmbContractor Is Nothing Then Return

            ' Create context menu
            Dim contextMenu As New ContextMenuStrip()

            ' Add menu items
            Dim clearHistoryItem As New ToolStripMenuItem("مسح التاريخ", Nothing, AddressOf ClearContractorHistory)
            Dim refreshHistoryItem As New ToolStripMenuItem("تحديث من Excel", Nothing, AddressOf RefreshContractorHistory)

            contextMenu.Items.Add(clearHistoryItem)
            contextMenu.Items.Add(refreshHistoryItem)

            ' Assign to combobox
            cmbContractor.ContextMenuStrip = contextMenu

            SafeLogInfo("Contractor combobox context menu configured")
        Catch ex As Exception
            SafeLogWarning($"Failed to setup contractor combobox context menu: {ex.Message}")
        End Try
    End Sub

    Private Sub ClearContractorHistory(sender As Object, e As EventArgs)
        Try
            Dim result As DialogResult = MessageBox.Show("هل أنت متأكد من مسح تاريخ أسماء المقاولين؟",
                                                        "تأكيد مسح التاريخ",
                                                        MessageBoxButtons.YesNo,
                                                        MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                contractorHistory.Clear()
                UpdateContractorComboBox()
                SaveContractorHistory()
                MessageBox.Show("تم مسح تاريخ أسماء المقاولين بنجاح", "تم المسح", MessageBoxButtons.OK, MessageBoxIcon.Information)
                SafeLogInfo("Contractor history cleared by user")
            End If
        Catch ex As Exception
            SafeLogError("Failed to clear contractor history", ex)
            MessageBox.Show($"خطأ في مسح التاريخ: {ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RefreshContractorHistory(sender As Object, e As EventArgs)
        Try
            LoadExistingContractorsFromExcel()
            UpdateContractorComboBox()
            SaveContractorHistory()
            MessageBox.Show("تم تحديث قائمة أسماء المقاولين من ملف Excel", "تم التحديث", MessageBoxButtons.OK, MessageBoxIcon.Information)
            SafeLogInfo("Contractor history refreshed from Excel")
        Catch ex As Exception
            SafeLogError("Failed to refresh contractor history", ex)
            MessageBox.Show($"خطأ في تحديث التاريخ: {ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' --------------------------------------
    ' Safe Logging Functions
    ' --------------------------------------
    Private Sub SafeLogInfo(message As String)
        Try
            Logger.LogInfo(message)
        Catch
            ' Fallback to debug output if Logger fails
            Debug.WriteLine($"[INFO] {message}")
        End Try
    End Sub

    Private Sub SafeLogWarning(message As String)
        Try
            Logger.LogWarning(message)
        Catch
            ' Fallback to debug output if Logger fails
            Debug.WriteLine($"[WARNING] {message}")
        End Try
    End Sub


    Private Sub StyleDataGrid()
        With DataGridView1
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
            .EnableHeadersVisualStyles = False

            ' الهيدر
            Dim headerStyle As New DataGridViewCellStyle()
            headerStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            headerStyle.BackColor = Color.FromArgb(64, 64, 128)
            headerStyle.ForeColor = Color.White
            headerStyle.Font = New Font("Segoe UI", 10, FontStyle.Bold)
            headerStyle.WrapMode = DataGridViewTriState.True
            .ColumnHeadersDefaultCellStyle = headerStyle

            ' الصفوف
            Dim rowStyle As New DataGridViewCellStyle()
            rowStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            rowStyle.BackColor = Color.White
            rowStyle.Font = New Font("Segoe UI", 9.75F, FontStyle.Regular)
            rowStyle.ForeColor = Color.Black
            rowStyle.SelectionBackColor = Color.LightSteelBlue
            rowStyle.SelectionForeColor = Color.Black
            .DefaultCellStyle = rowStyle

            ' صفوف متناوبة
            .AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240)

            ' RowHeader
            Dim rowHeaderStyle As New DataGridViewCellStyle()
            rowHeaderStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            rowHeaderStyle.BackColor = Color.FromArgb(90, 90, 200)
            rowHeaderStyle.ForeColor = Color.White
            rowHeaderStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            rowHeaderStyle.SelectionBackColor = Color.SteelBlue
            rowHeaderStyle.SelectionForeColor = Color.White
            .RowHeadersDefaultCellStyle = rowHeaderStyle

            ' خصائص إضافية
            .GridColor = Color.Black
            .BackgroundColor = Color.White
            .BorderStyle = BorderStyle.Fixed3D
            .CellBorderStyle = DataGridViewCellBorderStyle.Single
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        End With
    End Sub

    Private Sub SafeLogError(message As String, Optional ex As Exception = Nothing)
        Try
            Logger.LogError(message, ex)
        Catch
            ' Fallback to debug output if Logger fails
            Dim errorMsg As String = message
            If ex IsNot Nothing Then
                errorMsg &= $" | Exception: {ex.Message}"
            End If
            Debug.WriteLine($"[ERROR] {errorMsg}")
        End Try
    End Sub

    ' --------------------------------------
    ' Form Load Event
    ' --------------------------------------
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ' Initialize logging safely first
            SafeLogInfo("Application started")

            ' Set up global exception handling after basic initialization
            Try
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException)
                AddHandler Application.ThreadException, AddressOf Application_ThreadException
                SafeLogInfo("Global exception handler configured")
            Catch ex As Exception
                SafeLogWarning($"Failed to set up global exception handler: {ex.Message}")
                ' Continue without global exception handler - this is not critical
            End Try

            ' Validate file path
            If String.IsNullOrEmpty(filePath) Then
                MessageBox.Show("مسار الملف غير صحيح.", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()
                Return
            End If

            ' Ensure directory exists
            Try
                Dim directoryPath As String = Path.GetDirectoryName(filePath)
                If Not Directory.Exists(directoryPath) Then
                    Directory.CreateDirectory(directoryPath)
                    SafeLogInfo($"Created directory: {directoryPath}")
                End If
            Catch ex As Exception
                SafeLogError("Failed to create directory", ex)
                MessageBox.Show($"فشل في إنشاء المجلد: {ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()
                Return
            End Try



            Try
                SafeLogInfo("Attempting to create Excel application")
                xlApp = CreateObject("Excel.Application")

                If xlApp Is Nothing Then
                    Throw New Exception("Excel application object is null")
                End If

                xlApp.Visible = False ' Hide Excel
                xlApp.DisplayAlerts = False ' Disable alerts
                xlApp.ScreenUpdating = False ' Improve performance
                SafeLogInfo("Excel application created successfully")
            Catch ex As Exception
                SafeLogError("Failed to create Excel application", ex)
                MessageBox.Show($"فشل في إنشاء تطبيق Excel. تأكد من تثبيت Microsoft Excel على النظام." & Environment.NewLine & $"خطأ: {ex.Message}",
                               "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Me.Close()
                Return
            End Try

            If File.Exists(filePath) Then
                Try
                    xlWorkbook = xlApp.Workbooks.Open(filePath, [ReadOnly]:=False)
                    SafeLogInfo($"Opened existing Excel file: {filePath}")
                Catch ex As Exception
                    SafeLogWarning($"Failed to open file in read-write mode: {ex.Message}")
                    Try
                        ' Try opening in read-only mode if exclusive access fails
                        xlWorkbook = xlApp.Workbooks.Open(filePath, [ReadOnly]:=True)
                        SafeLogWarning($"Opened Excel file in read-only mode: {filePath}")
                        MessageBox.Show("تم فتح الملف في وضع القراءة فقط. لن تتمكن من الحفظ.", "تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Catch readOnlyEx As Exception
                        SafeLogError($"Failed to open file even in read-only mode: {readOnlyEx.Message}")
                        MessageBox.Show($"فشل في فتح الملف: {readOnlyEx.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        CleanupExcel()
                        Return
                    End Try
                End Try
            Else
                Try
                    xlWorkbook = xlApp.Workbooks.Add()
                    xlWorkbook.SaveAs(filePath)
                    SafeLogInfo($"Created new Excel file: {filePath}")
                Catch ex As Exception
                    SafeLogError($"Failed to create new Excel file: {ex.Message}")
                    MessageBox.Show($"فشل في إنشاء ملف Excel جديد: {ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    CleanupExcel()
                    Return
                End Try
            End If

            ' Load sheet names
            RefreshSheetsList()
            StyleDataGrid()

            ' Load contractor history and existing contractors from Excel
            LoadContractorHistory()
            LoadExistingContractorsFromExcel()
            UpdateContractorComboBox()

            ' Setup contractor combobox features
            SetupContractorComboBoxContextMenu()
            SetupContractorComboBoxAutoComplete()

        Catch ex As Exception
            SafeLogError("Error loading Excel file", ex)
            MessageBox.Show($"خطأ في تحميل ملف Excel: {ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CleanupExcel()
        End Try
    End Sub

    ' --------------------------------------
    ' Refresh Sheets List
    ' --------------------------------------
    Private Sub RefreshSheetsList()
        Try
            If cmbSheets Is Nothing Then Return

            cmbSheets.Items.Clear()
            If xlWorkbook IsNot Nothing Then
                For Each sh As Object In xlWorkbook.Sheets
                    cmbSheets.Items.Add(sh.Name)
                Next

                ' Select first sheet by default
                If cmbSheets.Items.Count > 0 Then
                    cmbSheets.SelectedIndex = 0
                End If
            End If
        Catch ex As Exception
            MessageBox.Show($"خطأ في تحديث قائمة الشيتات: {ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' --------------------------------------
    ' Sheet Selection Changed
    ' --------------------------------------
    Private Sub cmbSheets_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSheets.SelectedIndexChanged
        Try
            If cmbSheets.SelectedItem IsNot Nothing AndAlso xlWorkbook IsNot Nothing Then
                xlWorksheet = xlWorkbook.Sheets(cmbSheets.SelectedItem.ToString())
                LoadData()
            End If
        Catch ex As Exception
            MessageBox.Show($"خطأ في اختيار الشيت: {ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' --------------------------------------
    ' Load Data to DataGridView
    ' --------------------------------------
    Private Sub LoadData()
        Try
            If DataGridView1 Is Nothing Then Return

            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()

            ' إضافة الأعمدة
            DataGridView1.Columns.Add("ID", "مسلسل")
            DataGridView1.Columns.Add("Contractor", "المقاول")
            DataGridView1.Columns.Add("Work", "الأعمال")
            DataGridView1.Columns.Add("ContractDate", "تاريخ العقد")
            DataGridView1.Columns.Add("EsnadDate", "تاريخ الإسناد")
            DataGridView1.Columns.Add("ContractorSign", "توقيع المقاول")
            DataGridView1.Columns.Add("GeneralSign", "توقيع عام")
            DataGridView1.Columns.Add("Delivery", "تسليم")

            If xlWorksheet Is Nothing Then Return

            ' Start from row 2 (after header)
            Dim row As Integer = 2
            While xlWorksheet.Cells(row, 1).Value IsNot Nothing
                ' Ensure values are not empty
                Dim id As String = GetCellValue(xlWorksheet.Cells(row, 1).Value)
                Dim contractor As String = GetCellValue(xlWorksheet.Cells(row, 2).Value)
                Dim work As String = GetCellValue(xlWorksheet.Cells(row, 3).Value)
                Dim contractDate As String = GetCellValue(xlWorksheet.Cells(row, 4).Value)
                Dim esnadDate As String = GetCellValue(xlWorksheet.Cells(row, 5).Value)
                Dim contractorSign As Boolean = GetBooleanValue(xlWorksheet.Cells(row, 6).Value)
                Dim generalSign As Boolean = GetBooleanValue(xlWorksheet.Cells(row, 7).Value)
                Dim delivery As Boolean = GetBooleanValue(xlWorksheet.Cells(row, 8).Value)

                ' Convert boolean values to checkmarks/X marks
                Dim contractorSignDisplay As String = GetCellValue(xlWorksheet.Cells(row, 6).Value)
                Dim generalSignDisplay As String = GetCellValue(xlWorksheet.Cells(row, 7).Value)
                Dim deliveryDisplay As String = GetCellValue(xlWorksheet.Cells(row, 8).Value)

                DataGridView1.Rows.Add(id, contractor, work, contractDate, esnadDate, contractorSignDisplay, generalSignDisplay, deliveryDisplay)
                row += 1
            End While

            ' تحديث عداد السجلات
            If lblRecordCount IsNot Nothing Then
                lblRecordCount.Text = $"عدد السجلات: {row - 2}"
            End If

        Catch ex As Exception
            MessageBox.Show($"خطأ في تحميل البيانات: {ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' --------------------------------------
    ' Helper Functions for Value Handling
    ' --------------------------------------
    Private Function GetCellValue(cellValue As Object) As String
        If cellValue Is Nothing Then
            Return ""
        Else
            ' Check if the value is a date and format it properly
            Dim dateValue As DateTime
            If DateTime.TryParse(cellValue.ToString(), dateValue) Then
                Return dateValue.ToString("yyyy-MM-dd")
            Else
                Return cellValue.ToString()
            End If
        End If
    End Function

    Private Function GetBooleanValue(cellValue As Object) As Boolean
        If cellValue Is Nothing Then
            Return False
        Else
            Dim result As Boolean
            Boolean.TryParse(cellValue.ToString(), result)
            Return result
        End If
    End Function

    Private Function GetControlText(control As Control) As String
        If control Is Nothing Then
            Return ""
        Else
            Return control.Text
        End If
    End Function

    Private Function GetDateValue(dateControl As DateTimePicker) As DateTime
        If dateControl Is Nothing Then
            Return DateTime.Now
        Else
            Return dateControl.Value
        End If
    End Function

    Private Function GetCheckBoxValue(checkBox As CheckBox) As Boolean
        If checkBox Is Nothing Then
            Return False
        Else
            Return checkBox.Checked
        End If
    End Function

    ' --------------------------------------
    ' Data Validation Before Save
    ' --------------------------------------
    Private Function ValidateData() As Boolean
        ' التحقق من اسم المقاول
        Dim contractorValidation = DataValidator.ValidateContractorName(GetControlText(cmbContractor))
        If Not contractorValidation.IsValid Then
            MessageBox.Show(contractorValidation.ErrorMessage, "خطأ في التحقق", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            If cmbContractor IsNot Nothing Then cmbContractor.Focus()
            SafeLogWarning($"Validation failed for contractor: {contractorValidation.ErrorMessage}")
            Return False
        End If

        ' التحقق من وصف الأعمال
        Dim workValidation = DataValidator.ValidateWorkDescription(GetControlText(con_work))
        If Not workValidation.IsValid Then
            MessageBox.Show(workValidation.ErrorMessage, "خطأ في التحقق", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            If con_work IsNot Nothing Then con_work.Focus()
            SafeLogWarning($"Validation failed for work description: {workValidation.ErrorMessage}")
            Return False
        End If

        ' التحقق من تاريخ العقد
        If contract_date IsNot Nothing Then
            Dim contractDateValidation = DataValidator.ValidateDate(contract_date.Value)
            If Not contractDateValidation.IsValid Then
                MessageBox.Show(contractDateValidation.ErrorMessage, "خطأ في التحقق", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                contract_date.Focus()
                SafeLogWarning($"Validation failed for contract date: {contractDateValidation.ErrorMessage}")
                Return False
            End If
        End If

        ' التحقق من تاريخ الإسناد
        If esnad_date IsNot Nothing Then
            Dim esnadDateValidation = DataValidator.ValidateDate(esnad_date.Value)
            If Not esnadDateValidation.IsValid Then
                MessageBox.Show(esnadDateValidation.ErrorMessage, "خطأ في التحقق", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                esnad_date.Focus()
                SafeLogWarning($"Validation failed for assignment date: {esnadDateValidation.ErrorMessage}")
                Return False
            End If
        End If

        SafeLogInfo("Data validation passed successfully")
        Return True
    End Function

    ' --------------------------------------
    ' Save New Data
    ' --------------------------------------
    Private Sub save_data_Click(sender As Object, e As EventArgs) Handles save_data.Click
        Try
            If xlWorksheet Is Nothing Then
                MessageBox.Show("يرجى اختيار شيت أولاً.", "تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Check if workbook is read-only
            If xlWorkbook IsNot Nothing AndAlso xlWorkbook.ReadOnly Then
                MessageBox.Show("الملف مفتوح في وضع القراءة فقط. لا يمكن الحفظ." & Environment.NewLine &
                               "يرجى إغلاق الملف من Excel وإعادة تشغيل التطبيق.",
                               "خطأ في الحفظ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                SafeLogWarning("Attempted to save to read-only workbook")
                Return
            End If

            If Not ValidateData() Then Return

            ' الحصول على آخر صف باستخدام الثابت
            Dim lastRow As Integer = xlWorksheet.Cells(xlWorksheet.Rows.Count, 1).End(xlUp).Row + 1

            ' إضافة البيانات مع التعامل المتسق مع القيم المنطقية
            xlWorksheet.Cells(lastRow, 1).Value = lastRow - 1
            xlWorksheet.Cells(lastRow, 2).Value = GetControlText(cmbContractor)
            xlWorksheet.Cells(lastRow, 3).Value = GetControlText(con_work)
            xlWorksheet.Cells(lastRow, 4).Value = GetDateValue(contract_date)
            xlWorksheet.Cells(lastRow, 5).Value = GetDateValue(esnad_date)
            xlWorksheet.Cells(lastRow, 6).Value = If(GetCheckBoxValue(contr_sign_checkbox), "✔", "✖")
            xlWorksheet.Cells(lastRow, 7).Value = If(GetCheckBoxValue(general_sign_checkbox), "✔", "✖")
            xlWorksheet.Cells(lastRow, 8).Value = If(GetCheckBoxValue(con_deliver), "✔", "✖")

            ' Add contractor to history
            AddContractorToHistory(GetControlText(cmbContractor))

            ' Attempt to save with better error handling
            TrySaveWorkbook()
            LoadData()
            ReorderIDs()

            SafeLogInfo($"New record saved successfully. Contractor: {GetControlText(cmbContractor)}, Work: {GetControlText(con_work)}")
            MessageBox.Show("تم حفظ البيانات بنجاح", "نجح", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            SafeLogError("Error saving data", ex)
            MessageBox.Show($"خطأ في حفظ البيانات: {ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' --------------------------------------
    ' Edit Data
    ' --------------------------------------
    Private Sub edit_data_Click(sender As Object, e As EventArgs) Handles edit_data.Click
        Try
            If xlWorksheet Is Nothing Then
                MessageBox.Show("يرجى اختيار شيت أولاً.", "تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If DataGridView1 Is Nothing OrElse DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("يرجى اختيار سجل للتعديل", "تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Check if workbook is read-only
            If xlWorkbook IsNot Nothing AndAlso xlWorkbook.ReadOnly Then
                MessageBox.Show("الملف مفتوح في وضع القراءة فقط. لا يمكن التعديل." & Environment.NewLine &
                               "يرجى إغلاق الملف من Excel وإعادة تشغيل التطبيق.",
                               "خطأ في التعديل", MessageBoxButtons.OK, MessageBoxIcon.Error)
                SafeLogWarning("Attempted to edit in read-only workbook")
                Return
            End If

            If Not ValidateData() Then Return

            Dim rowIndex As Integer = DataGridView1.SelectedRows(0).Index + 2

            ' تحديث البيانات مع التعامل المتسق مع القيم المنطقية
            xlWorksheet.Cells(rowIndex, 2).Value = GetControlText(cmbContractor)
            xlWorksheet.Cells(rowIndex, 3).Value = GetControlText(con_work)
            xlWorksheet.Cells(rowIndex, 4).Value = GetDateValue(contract_date)
            xlWorksheet.Cells(rowIndex, 5).Value = GetDateValue(esnad_date)
            xlWorksheet.Cells(rowIndex, 6).Value = If(GetCheckBoxValue(contr_sign_checkbox), "✔", "✖")
            xlWorksheet.Cells(rowIndex, 7).Value = If(GetCheckBoxValue(general_sign_checkbox), "✔", "✖")
            xlWorksheet.Cells(rowIndex, 8).Value = If(GetCheckBoxValue(con_deliver), "✔", "✖")

            TrySaveWorkbook()
            LoadData()
            MessageBox.Show("تم تعديل البيانات بنجاح", "نجح", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"خطأ في تعديل البيانات: {ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' --------------------------------------
    ' Delete Data
    ' --------------------------------------
    Private Sub delete_data_Click(sender As Object, e As EventArgs) Handles delete_data.Click
        Try
            If xlWorksheet Is Nothing Then
                MessageBox.Show("يرجى اختيار شيت أولاً.", "تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If DataGridView1 Is Nothing OrElse DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("يرجى اختيار سجل للحذف", "تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Check if workbook is read-only
            If xlWorkbook IsNot Nothing AndAlso xlWorkbook.ReadOnly Then
                MessageBox.Show("الملف مفتوح في وضع القراءة فقط. لا يمكن الحذف." & Environment.NewLine &
                               "يرجى إغلاق الملف من Excel وإعادة تشغيل التطبيق.",
                               "خطأ في الحذف", MessageBoxButtons.OK, MessageBoxIcon.Error)
                SafeLogWarning("Attempted to delete from read-only workbook")
                Return
            End If

            ' تأكيد الحذف
            Dim result As DialogResult = MessageBox.Show("هل أنت متأكد من حذف هذا السجل؟", "تأكيد الحذف", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.No Then Return

            Dim rowIndex As Integer = DataGridView1.SelectedRows(0).Index + 2
            xlWorksheet.Rows(rowIndex).Delete()
            TrySaveWorkbook()
            LoadData()
            MessageBox.Show("تم حذف السجل بنجاح", "نجح", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"خطأ في حذف البيانات: {ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' --------------------------------------
    ' Search Functionality - Search in All Sheets
    ' --------------------------------------
    Private Sub Search_box_TextChanged(sender As Object, e As EventArgs) Handles Search_box.TextChanged
        Try
            If DataGridView1 Is Nothing OrElse Search_box Is Nothing OrElse xlWorkbook Is Nothing Then Return

            Dim searchText As String = Search_box.Text.Trim().ToLower()

            ' Clear current data
            DataGridView1.Rows.Clear()
            DataGridView1.Columns.Clear()

            ' Add columns
            DataGridView1.Columns.Add("ID", "مسلسل")
            DataGridView1.Columns.Add("Contractor", "المقاول")
            DataGridView1.Columns.Add("Work", "الأعمال")
            DataGridView1.Columns.Add("ContractDate", "تاريخ العقد")
            DataGridView1.Columns.Add("EsnadDate", "تاريخ الإسناد")
            DataGridView1.Columns.Add("ContractorSign", "توقيع المقاول")
            DataGridView1.Columns.Add("GeneralSign", "توقيع عام")
            DataGridView1.Columns.Add("Delivery", "تسليم")
            DataGridView1.Columns.Add("SheetName", "اسم الشيت")

            If String.IsNullOrEmpty(searchText) Then
                ' If search is empty, show current sheet data
                LoadData()
                ' Update record counter to show normal count
                If lblRecordCount IsNot Nothing Then
                    lblRecordCount.Text = "عدد السجلات: " & (DataGridView1.Rows.Count - 1)
                End If
                Return
            End If

            Dim totalRecords As Integer = 0

            ' Search in all sheets
            For Each sheet As Object In xlWorkbook.Sheets
                Try
                    Dim currentSheet As Object = sheet
                    Dim sheetName As String = currentSheet.Name

                    ' Start from row 2 (after header)
                    Dim row As Integer = 2
                    While currentSheet.Cells(row, 1).Value IsNot Nothing
                        ' Get all cell values for this row
                        Dim id As String = GetCellValue(currentSheet.Cells(row, 1).Value)
                        Dim contractor As String = GetCellValue(currentSheet.Cells(row, 2).Value)
                        Dim work As String = GetCellValue(currentSheet.Cells(row, 3).Value)
                        Dim contractDate As String = GetCellValue(currentSheet.Cells(row, 4).Value)
                        Dim esnadDate As String = GetCellValue(currentSheet.Cells(row, 5).Value)
                        Dim contractorSign As Boolean = GetBooleanValue(currentSheet.Cells(row, 6).Value)
                        Dim generalSign As Boolean = GetBooleanValue(currentSheet.Cells(row, 7).Value)
                        Dim delivery As Boolean = GetBooleanValue(currentSheet.Cells(row, 8).Value)

                        ' Check if any field contains the search text
                        Dim found As Boolean = False
                        If contractor.ToLower().Contains(searchText) OrElse
                           work.ToLower().Contains(searchText) OrElse
                           contractDate.ToLower().Contains(searchText) OrElse
                           esnadDate.ToLower().Contains(searchText) OrElse
                           id.ToLower().Contains(searchText) OrElse
                           sheetName.ToLower().Contains(searchText) Then
                            found = True
                        End If

                        ' If found, add to DataGridView
                        If found Then
                            ' Convert boolean values to checkmarks/X marks
                            Dim contractorSignDisplay As String = If(contractorSign, "✔", "✖")
                            Dim generalSignDisplay As String = If(generalSign, "✔", "✖")
                            Dim deliveryDisplay As String = If(delivery, "✔", "✖")

                            DataGridView1.Rows.Add(id, contractor, work, contractDate, esnadDate, contractorSignDisplay, generalSignDisplay, deliveryDisplay, sheetName)
                            totalRecords += 1
                        End If

                        row += 1
                    End While
                Catch ex As Exception
                    ' Continue with next sheet if there's an error
                    SafeLogWarning($"Error searching in sheet {sheet.Name}: {ex.Message}")
                End Try
            Next

            ' Update record counter
            If lblRecordCount IsNot Nothing Then
                lblRecordCount.Text = $"نتائج البحث: {totalRecords} سجل"
            End If

        Catch ex As Exception
            MessageBox.Show($"خطأ في البحث: {ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' --------------------------------------
    ' Create New Sheet
    ' --------------------------------------
    Private Sub create_sheet_Click(sender As Object, e As EventArgs) Handles create_sheet.Click
        Try
            Dim sheetName As String = InputBox("اكتب اسم الشيت الجديد:", "إنشاء شيت جديد")

            If String.IsNullOrWhiteSpace(sheetName) Then
                MessageBox.Show("اسم الشيت لا يمكن أن يكون فارغ")
                Return
            End If

            Dim newSheet As Object = xlWorkbook.Sheets.Add()
            newSheet.Name = sheetName

            newSheet.DisplayRightToLeft = True
            newSheet.PageSetup.Orientation = xlLandscape
            newSheet.Rows("2:2").Select()
            newSheet.Application.ActiveWindow.FreezePanes = True

            ' إضافة رؤوس الأعمدة
            Dim headers() As String = {"مسلسل", "اسم المقاول", "الأعمال", "تاريخ العقد", "تاريخ أمر الإسناد",
                                  "توقيع المقاول", "توقيع المدير العام", "التسليم"}
            For i As Integer = 0 To headers.Length - 1
                newSheet.Cells(1, i + 1).Value = headers(i)
            Next

            ' تنسيق الهيدر
            With newSheet.Range("A1:H1")
                .Font.Bold = True
                .Font.Size = 12
                .Interior.Color = RGB(0, 102, 204)
                .Font.Color = RGB(255, 255, 255)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .RowHeight = 25
                .Borders.LineStyle = xlContinuous
            End With

            ' تنسيق الأعمدة
            ' Column A (ID) - Number format
            newSheet.Columns("A:A").NumberFormat = "0"

            ' Column D (Contract Date) - Date format
            newSheet.Columns("D:D").NumberFormat = "yyyy-mm-dd"

            ' Column E (Assignment Date) - Date format
            newSheet.Columns("E:E").NumberFormat = "yyyy-mm-dd"

            ' Columns F, G, H (Boolean) - Custom format for checkmarks
            newSheet.Columns("F:F").NumberFormat = """✔"";""✖"""
            newSheet.Columns("G:G").NumberFormat = """✔"";""✖"""
            newSheet.Columns("H:H").NumberFormat = """✔"";""✖"""

            ' توسيع الأعمدة
            newSheet.Columns("A:H").ColumnWidth = 20
            newSheet.Columns("A:A").ColumnWidth = 6
            newSheet.Columns("C:C").ColumnWidth = 30

            ' إضافة Borders + تظليل صفوف بالتبادل
            Dim lastRow As Long = newSheet.Cells(newSheet.Rows.Count, 1).End(xlUp).Row
            If lastRow < 2 Then lastRow = 50 ' نجهز 50 صف مبدئي للتنسيق

            With newSheet.Range(newSheet.Cells(2, 1), newSheet.Cells(lastRow, 8))
                .Borders.LineStyle = xlContinuous
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                ' تظليل الصفوف الزوجية
                For i As Long = 2 To lastRow Step 2
                    newSheet.Rows(i).Interior.Color = RGB(240, 248, 255)
                Next i
            End With

            ' حفظ الملف
            TrySaveWorkbook()
            MessageBox.Show($"تم إنشاء الشيت '{sheetName}' بنجاح.")

            ' تحديث ComboBox
            RefreshSheetsList()
        Catch ex As Exception
            MessageBox.Show($"خطأ أثناء إنشاء الشيت: {ex.Message}")
        End Try
    End Sub

    ' --------------------------------------
    ' Reorder ID Column
    ' --------------------------------------
    Private Sub ReorderIDs()
        Try
            If xlWorksheet Is Nothing Then Return
            Dim row As Integer = 2
            Dim counter As Integer = 1
            While xlWorksheet.Cells(row, 1).Value IsNot Nothing
                xlWorksheet.Cells(row, 1).Value = counter
                row += 1
                counter += 1
            End While
            TrySaveWorkbook()
        Catch ex As Exception
            MessageBox.Show($"خطأ في إعادة ترتيب الأرقام: {ex.Message}")
        End Try
    End Sub

    ' --------------------------------------
    ' Check Sheet Existence
    ' --------------------------------------
    Private Function SheetExists(sheetName As String) As Boolean
        Try
            If xlWorkbook Is Nothing Then Return False

            For Each sheet As Object In xlWorkbook.Sheets
                If String.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase) Then
                    Return True
                End If
            Next
            Return False
        Catch
            Return False
        End Try
    End Function

    ' --------------------------------------
    ' Clear Fields
    ' --------------------------------------
    Private Sub clear_fields_Click(sender As Object, e As EventArgs) Handles clear_fields.Click
        Try
            If cmbContractor IsNot Nothing Then cmbContractor.Text = ""
            If con_work IsNot Nothing Then con_work.Text = ""
            If contract_date IsNot Nothing Then contract_date.Value = DateTime.Now
            If esnad_date IsNot Nothing Then esnad_date.Value = DateTime.Now
            If contr_sign_checkbox IsNot Nothing Then contr_sign_checkbox.Checked = False
            If general_sign_checkbox IsNot Nothing Then general_sign_checkbox.Checked = False
            If con_deliver IsNot Nothing Then con_deliver.Checked = False
        Catch ex As Exception
            MessageBox.Show($"خطأ في مسح الحقول: {ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' --------------------------------------
    ' DataGridView Selection Changed
    ' --------------------------------------
    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        Try
            If DataGridView1 Is Nothing OrElse DataGridView1.SelectedRows.Count = 0 Then Return
            If DataGridView1.SelectedRows(0).IsNewRow Then Return

            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            ' Check if this is a search result (has SheetName column)
            Try
                Dim sheetNameCell As DataGridViewCell = selectedRow.Cells("SheetName")
                If sheetNameCell IsNot Nothing AndAlso sheetNameCell.Value IsNot Nothing Then
                    Dim sheetName As String = GetCellValue(sheetNameCell.Value)

                    ' Switch to the sheet that contains this record
                    If cmbSheets IsNot Nothing AndAlso cmbSheets.Items.Contains(sheetName) Then
                        cmbSheets.SelectedItem = sheetName
                    End If
                End If
            Catch ex As Exception
                ' SheetName column doesn't exist, this is normal view
            End Try

            If cmbContractor IsNot Nothing Then
                Dim contractorValue As Object = selectedRow.Cells("Contractor").Value
                cmbContractor.Text = GetCellValue(contractorValue)
            End If

            If con_work IsNot Nothing Then
                Dim workValue As Object = selectedRow.Cells("Work").Value
                con_work.Text = GetCellValue(workValue)
            End If

            ' Parse date safely
            If contract_date IsNot Nothing Then
                Dim contractDateValue As String = GetCellValue(selectedRow.Cells("ContractDate").Value)
                Dim parsedDate As DateTime
                If DateTime.TryParse(contractDateValue, parsedDate) Then
                    contract_date.Value = parsedDate
                End If
            End If

            If esnad_date IsNot Nothing Then
                Dim esnadDateValue As String = GetCellValue(selectedRow.Cells("EsnadDate").Value)
                Dim parsedDate As DateTime
                If DateTime.TryParse(esnadDateValue, parsedDate) Then
                    esnad_date.Value = parsedDate
                End If
            End If

            If contr_sign_checkbox IsNot Nothing Then
                Dim contractorSignValue As String = GetCellValue(selectedRow.Cells("ContractorSign").Value)
                contr_sign_checkbox.Checked = (contractorSignValue = "✔")
            End If

            If general_sign_checkbox IsNot Nothing Then
                Dim generalSignValue As String = GetCellValue(selectedRow.Cells("GeneralSign").Value)
                general_sign_checkbox.Checked = (generalSignValue = "✔")
            End If

            If con_deliver IsNot Nothing Then
                Dim deliveryValue As String = GetCellValue(selectedRow.Cells("Delivery").Value)
                con_deliver.Checked = (deliveryValue = "✔")
            End If

        Catch ex As Exception
            ' Ignore errors here as they may occur during data loading
        End Try
    End Sub

    ' --------------------------------------
    ' Cleanup Excel Resources
    ' --------------------------------------
    Private Sub CleanupExcel()
        Try
            SafeLogInfo("Starting Excel cleanup")

            If xlWorksheet IsNot Nothing Then
                Try
                    Marshal.ReleaseComObject(xlWorksheet)
                    SafeLogInfo("Worksheet object released")
                Catch ex As Exception
                    SafeLogWarning($"Error releasing worksheet: {ex.Message}")
                End Try
                xlWorksheet = Nothing
            End If

            If xlWorkbook IsNot Nothing Then
                Try
                    xlWorkbook.Close(False)
                    SafeLogInfo("Workbook closed")
                Catch ex As Exception
                    SafeLogWarning($"Error closing workbook: {ex.Message}")
                End Try
                Try
                    Marshal.ReleaseComObject(xlWorkbook)
                    SafeLogInfo("Workbook object released")
                Catch ex As Exception
                    SafeLogWarning($"Error releasing workbook: {ex.Message}")
                End Try
                xlWorkbook = Nothing
            End If

            If xlApp IsNot Nothing Then
                Try
                    xlApp.Quit()
                    SafeLogInfo("Excel application quit")
                Catch ex As Exception
                    SafeLogWarning($"Error quitting Excel: {ex.Message}")
                End Try
                Try
                    Marshal.ReleaseComObject(xlApp)
                    SafeLogInfo("Excel application object released")
                Catch ex As Exception
                    SafeLogWarning($"Error releasing Excel app: {ex.Message}")
                End Try
                xlApp = Nothing
            End If

            ' جمع الموارد مرتين للتأكيد
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            GC.WaitForPendingFinalizers()

            ' قفل أي Excel يتيم
            For Each proc As Process In Process.GetProcessesByName("EXCEL")
                Try
                    If String.IsNullOrEmpty(proc.MainWindowTitle) Then
                        proc.Kill()
                        SafeLogWarning("Orphan Excel process killed")
                    End If
                Catch ex As Exception
                    SafeLogWarning($"Error killing orphan Excel: {ex.Message}")
                End Try
            Next

            SafeLogInfo("Excel cleanup completed")

        Catch ex As Exception
            SafeLogError("Error during Excel cleanup", ex)
        End Try
    End Sub


    ' --------------------------------------
    ' Form Closing Event
    ' --------------------------------------
    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        Try
            CleanupExcel()
            MyBase.OnFormClosing(e)
        Catch ex As Exception
            MessageBox.Show($"خطأ في إغلاق Excel: {ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MyBase.OnFormClosing(e)
        End Try
    End Sub


End Class