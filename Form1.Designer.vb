<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As DataGridViewCellStyle = New DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As DataGridViewCellStyle = New DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        GroupBox1 = New GroupBox()
        GroupBox4 = New GroupBox()
        Search_box = New TextBox()
        lblRecordCount = New Label()
        create_sheet = New Button()
        cmbSheets = New ComboBox()
        Label1 = New Label()
        GroupBox2 = New GroupBox()
        con_deliver = New CheckBox()
        general_sign_checkbox = New CheckBox()
        contr_sign_checkbox = New CheckBox()
        contract_date = New DateTimePicker()
        esnad_date = New DateTimePicker()
        Label6 = New Label()
        Label5 = New Label()
        con_work = New TextBox()
        cmbContractor = New ComboBox()
        Label4 = New Label()
        Label3 = New Label()
        GroupBox3 = New GroupBox()
        delete_data = New Button()
        clear_fields = New Button()
        save_data = New Button()
        edit_data = New Button()
        DataGridView1 = New DataGridView()
        GroupBox1.SuspendLayout()
        GroupBox4.SuspendLayout()
        GroupBox2.SuspendLayout()
        GroupBox3.SuspendLayout()
        CType(DataGridView1, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' GroupBox1
        ' 
        GroupBox1.BackColor = Color.LightSteelBlue
        GroupBox1.Controls.Add(GroupBox4)
        GroupBox1.Controls.Add(lblRecordCount)
        GroupBox1.Controls.Add(create_sheet)
        GroupBox1.Controls.Add(cmbSheets)
        GroupBox1.Controls.Add(Label1)
        GroupBox1.Location = New Point(12, 12)
        GroupBox1.Name = "GroupBox1"
        GroupBox1.Size = New Size(963, 104)
        GroupBox1.TabIndex = 0
        GroupBox1.TabStop = False
        GroupBox1.Text = "ملفات الأكسيل"
        ' 
        ' GroupBox4
        ' 
        GroupBox4.Controls.Add(Search_box)
        GroupBox4.Location = New Point(6, 13)
        GroupBox4.Name = "GroupBox4"
        GroupBox4.Size = New Size(355, 77)
        GroupBox4.TabIndex = 5
        GroupBox4.TabStop = False
        GroupBox4.Text = "شريط البحث"
        ' 
        ' Search_box
        ' 
        Search_box.Font = New Font("Tahoma", 10F)
        Search_box.Location = New Point(18, 28)
        Search_box.Name = "Search_box"
        Search_box.Size = New Size(318, 24)
        Search_box.TabIndex = 4
        ' 
        ' lblRecordCount
        ' 
        lblRecordCount.AutoSize = True
        lblRecordCount.Font = New Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        lblRecordCount.ForeColor = Color.Maroon
        lblRecordCount.Location = New Point(544, 45)
        lblRecordCount.Name = "lblRecordCount"
        lblRecordCount.Size = New Size(92, 20)
        lblRecordCount.TabIndex = 3
        lblRecordCount.Text = "عدد السجلات"
        ' 
        ' create_sheet
        ' 
        create_sheet.BackColor = Color.DarkOliveGreen
        create_sheet.FlatAppearance.BorderSize = 0
        create_sheet.FlatStyle = FlatStyle.Flat
        create_sheet.ForeColor = Color.White
        create_sheet.Location = New Point(388, 43)
        create_sheet.Name = "create_sheet"
        create_sheet.Size = New Size(119, 26)
        create_sheet.TabIndex = 2
        create_sheet.Text = "انشاء sheet"
        create_sheet.UseVisualStyleBackColor = False
        ' 
        ' cmbSheets
        ' 
        cmbSheets.Font = New Font("Tahoma", 10F)
        cmbSheets.FormattingEnabled = True
        cmbSheets.Location = New Point(642, 45)
        cmbSheets.Name = "cmbSheets"
        cmbSheets.Size = New Size(237, 24)
        cmbSheets.TabIndex = 1
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(885, 50)
        Label1.Name = "Label1"
        Label1.Size = New Size(72, 15)
        Label1.TabIndex = 0
        Label1.Text = "ملف الأكسيل"
        ' 
        ' GroupBox2
        ' 
        GroupBox2.BackColor = Color.LightSteelBlue
        GroupBox2.Controls.Add(con_deliver)
        GroupBox2.Controls.Add(general_sign_checkbox)
        GroupBox2.Controls.Add(contr_sign_checkbox)
        GroupBox2.Controls.Add(contract_date)
        GroupBox2.Controls.Add(esnad_date)
        GroupBox2.Controls.Add(Label6)
        GroupBox2.Controls.Add(Label5)
        GroupBox2.Controls.Add(con_work)
        GroupBox2.Controls.Add(cmbContractor)
        GroupBox2.Controls.Add(Label4)
        GroupBox2.Controls.Add(Label3)
        GroupBox2.Location = New Point(12, 122)
        GroupBox2.Name = "GroupBox2"
        GroupBox2.Size = New Size(963, 141)
        GroupBox2.TabIndex = 1
        GroupBox2.TabStop = False
        GroupBox2.Text = "ادخال / عرض البيانات"
        ' 
        ' con_deliver
        ' 
        con_deliver.AutoSize = True
        con_deliver.ImeMode = ImeMode.NoControl
        con_deliver.Location = New Point(95, 105)
        con_deliver.Name = "con_deliver"
        con_deliver.RightToLeft = RightToLeft.No
        con_deliver.Size = New Size(97, 19)
        con_deliver.TabIndex = 10
        con_deliver.Text = "استلام المقاول"
        con_deliver.UseVisualStyleBackColor = True
        ' 
        ' general_sign_checkbox
        ' 
        general_sign_checkbox.AutoSize = True
        general_sign_checkbox.ImeMode = ImeMode.NoControl
        general_sign_checkbox.Location = New Point(457, 105)
        general_sign_checkbox.Name = "general_sign_checkbox"
        general_sign_checkbox.RightToLeft = RightToLeft.No
        general_sign_checkbox.Size = New Size(112, 19)
        general_sign_checkbox.TabIndex = 9
        general_sign_checkbox.Text = "توقيع المدير العام"
        general_sign_checkbox.UseVisualStyleBackColor = True
        ' 
        ' contr_sign_checkbox
        ' 
        contr_sign_checkbox.AutoSize = True
        contr_sign_checkbox.Location = New Point(757, 105)
        contr_sign_checkbox.Name = "contr_sign_checkbox"
        contr_sign_checkbox.RightToLeft = RightToLeft.No
        contr_sign_checkbox.Size = New Size(96, 19)
        contr_sign_checkbox.TabIndex = 8
        contr_sign_checkbox.Text = "توقيع  المقاول"
        contr_sign_checkbox.UseVisualStyleBackColor = True
        ' 
        ' contract_date
        ' 
        contract_date.Location = New Point(24, 64)
        contract_date.Name = "contract_date"
        contract_date.Size = New Size(209, 23)
        contract_date.TabIndex = 7
        ' 
        ' esnad_date
        ' 
        esnad_date.Location = New Point(24, 24)
        esnad_date.Name = "esnad_date"
        esnad_date.Size = New Size(209, 23)
        esnad_date.TabIndex = 6
        ' 
        ' Label6
        ' 
        Label6.AutoSize = True
        Label6.ImeMode = ImeMode.NoControl
        Label6.Location = New Point(250, 70)
        Label6.Name = "Label6"
        Label6.Size = New Size(60, 15)
        Label6.TabIndex = 5
        Label6.Text = "تاريخ العقد"
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.ImeMode = ImeMode.NoControl
        Label5.Location = New Point(250, 25)
        Label5.Name = "Label5"
        Label5.Size = New Size(85, 15)
        Label5.TabIndex = 4
        Label5.Text = "تاريخ امر الاسناد"
        ' 
        ' con_work
        ' 
        con_work.Font = New Font("Tahoma", 10F)
        con_work.Location = New Point(367, 62)
        con_work.Name = "con_work"
        con_work.Size = New Size(470, 24)
        con_work.TabIndex = 3
        ' 
        ' cmbContractor
        ' 
        cmbContractor.Font = New Font("Tahoma", 10F)
        cmbContractor.FormattingEnabled = True
        cmbContractor.Location = New Point(367, 22)
        cmbContractor.Name = "cmbContractor"
        cmbContractor.Size = New Size(470, 24)
        cmbContractor.TabIndex = 2
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.ImeMode = ImeMode.NoControl
        Label4.Location = New Point(860, 65)
        Label4.Name = "Label4"
        Label4.Size = New Size(97, 15)
        Label4.TabIndex = 1
        Label4.Text = "الأعمال المكلف بها"
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(860, 30)
        Label3.Name = "Label3"
        Label3.Size = New Size(67, 15)
        Label3.TabIndex = 0
        Label3.Text = "اسم المقاول"
        ' 
        ' GroupBox3
        ' 
        GroupBox3.BackColor = Color.LightSteelBlue
        GroupBox3.Controls.Add(delete_data)
        GroupBox3.Controls.Add(clear_fields)
        GroupBox3.Controls.Add(save_data)
        GroupBox3.Controls.Add(edit_data)
        GroupBox3.Location = New Point(12, 268)
        GroupBox3.Name = "GroupBox3"
        GroupBox3.Size = New Size(963, 69)
        GroupBox3.TabIndex = 2
        GroupBox3.TabStop = False
        GroupBox3.Text = "التحكم"
        ' 
        ' delete_data
        ' 
        delete_data.BackColor = Color.Firebrick
        delete_data.FlatAppearance.BorderSize = 0
        delete_data.FlatStyle = FlatStyle.Flat
        delete_data.Font = New Font("Segoe UI", 10F, FontStyle.Bold)
        delete_data.ForeColor = Color.White
        delete_data.Image = My.Resources.Resources.icons8_delete_24
        delete_data.ImageAlign = ContentAlignment.MiddleRight
        delete_data.Location = New Point(799, 22)
        delete_data.Name = "delete_data"
        delete_data.Size = New Size(128, 34)
        delete_data.TabIndex = 7
        delete_data.Text = "حذف البيانات"
        delete_data.TextAlign = ContentAlignment.MiddleLeft
        delete_data.UseVisualStyleBackColor = False
        ' 
        ' clear_fields
        ' 
        clear_fields.BackColor = Color.DarkOrange
        clear_fields.FlatAppearance.BorderSize = 0
        clear_fields.FlatStyle = FlatStyle.Flat
        clear_fields.Font = New Font("Segoe UI", 10F, FontStyle.Bold)
        clear_fields.ForeColor = Color.White
        clear_fields.Image = My.Resources.Resources.icons8_clear_24
        clear_fields.ImageAlign = ContentAlignment.MiddleRight
        clear_fields.Location = New Point(642, 22)
        clear_fields.Name = "clear_fields"
        clear_fields.Size = New Size(128, 34)
        clear_fields.TabIndex = 7
        clear_fields.Text = "مسح الحقول"
        clear_fields.TextAlign = ContentAlignment.MiddleLeft
        clear_fields.UseVisualStyleBackColor = False
        ' 
        ' save_data
        ' 
        save_data.BackColor = Color.SeaGreen
        save_data.FlatAppearance.BorderSize = 0
        save_data.FlatStyle = FlatStyle.Flat
        save_data.Font = New Font("Segoe UI", 10F, FontStyle.Bold)
        save_data.ForeColor = Color.White
        save_data.Image = My.Resources.Resources.icons8_save_504
        save_data.ImageAlign = ContentAlignment.MiddleRight
        save_data.Location = New Point(24, 22)
        save_data.Name = "save_data"
        save_data.Size = New Size(128, 34)
        save_data.TabIndex = 6
        save_data.Text = "حفظ البيانات"
        save_data.TextAlign = ContentAlignment.MiddleLeft
        save_data.UseVisualStyleBackColor = False
        ' 
        ' edit_data
        ' 
        edit_data.BackColor = Color.RoyalBlue
        edit_data.FlatAppearance.BorderSize = 0
        edit_data.FlatStyle = FlatStyle.Flat
        edit_data.Font = New Font("Segoe UI", 10F, FontStyle.Bold)
        edit_data.ForeColor = Color.White
        edit_data.Image = My.Resources.Resources.icons8_edit_24
        edit_data.ImageAlign = ContentAlignment.MiddleRight
        edit_data.Location = New Point(182, 22)
        edit_data.Name = "edit_data"
        edit_data.Size = New Size(128, 34)
        edit_data.TabIndex = 5
        edit_data.Text = "تعديل البيانات"
        edit_data.TextAlign = ContentAlignment.MiddleLeft
        edit_data.UseVisualStyleBackColor = False
        ' 
        ' DataGridView1
        ' 
        DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = Color.WhiteSmoke
        DataGridViewCellStyle1.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        DataGridViewCellStyle1.ForeColor = SystemColors.ControlText
        DataGridViewCellStyle1.SelectionBackColor = SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = DataGridViewTriState.False
        DataGridView1.DefaultCellStyle = DataGridViewCellStyle1
        DataGridView1.GridColor = Color.Black
        DataGridView1.Location = New Point(12, 343)
        DataGridView1.Name = "DataGridView1"
        DataGridViewCellStyle2.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = Color.FromArgb(CByte(149), CByte(149), CByte(255))
        DataGridViewCellStyle2.Font = New Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        DataGridViewCellStyle2.ForeColor = SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = SystemColors.ActiveCaption
        DataGridViewCellStyle2.SelectionForeColor = SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = DataGridViewTriState.True
        DataGridView1.RowHeadersDefaultCellStyle = DataGridViewCellStyle2
        DataGridView1.Size = New Size(963, 225)
        DataGridView1.TabIndex = 3
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        AutoSize = True
        BackColor = Color.LightSteelBlue
        ClientSize = New Size(984, 574)
        Controls.Add(DataGridView1)
        Controls.Add(GroupBox3)
        Controls.Add(GroupBox2)
        Controls.Add(GroupBox1)
        FormBorderStyle = FormBorderStyle.FixedSingle
        Icon = CType(resources.GetObject("$this.Icon"), Icon)
        MaximizeBox = False
        Name = "Form1"
        RightToLeft = RightToLeft.Yes
        RightToLeftLayout = True
        StartPosition = FormStartPosition.CenterScreen
        Text = "شاشة بيانات المقاولين- Mohamed Atef"
        GroupBox1.ResumeLayout(False)
        GroupBox1.PerformLayout()
        GroupBox4.ResumeLayout(False)
        GroupBox4.PerformLayout()
        GroupBox2.ResumeLayout(False)
        GroupBox2.PerformLayout()
        GroupBox3.ResumeLayout(False)
        CType(DataGridView1, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents lblRecordCount As Label
    Friend WithEvents create_sheet As Button
    Friend WithEvents cmbSheets As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents con_work As TextBox
    Friend WithEvents cmbContractor As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents contract_date As DateTimePicker
    Friend WithEvents esnad_date As DateTimePicker
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents contr_sign_checkbox As CheckBox
    Friend WithEvents general_sign_checkbox As CheckBox
    Friend WithEvents con_deliver As CheckBox
    Friend WithEvents edit_data As Button
    Friend WithEvents Search_box As TextBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents clear_fields As Button
    Friend WithEvents delete_data As Button
    Friend WithEvents save_data As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents GroupBox4 As GroupBox

End Class
