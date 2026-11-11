<div align="center"><img src="https://raw.githubusercontent.com/matef-hub/atefdodo/main/assets/research-signal.svg" width="100%" alt="Legal research · verified data · applied AI"/></div>

# Contractors Tracking System | متابعة عقود المقاولين

**VB.NET · Windows Forms · Excel data storage**

A contractor contracts and work-assignment tracker with an Arabic-facing workflow. This is a desktop application rather than the separate Laravel or Android law-office projects.

```mermaid
flowchart LR
  A["Contractor"] --> B["Work assignment"]
  B --> C["Excel workbook"]
  C --> D["Search / print / review"]
```

## Technical documentation


A VB.NET Windows Forms application for managing contractor contracts and work assignments using Excel as the backend database.

## Features

- **Excel Integration**: Direct integration with Microsoft Excel for data storage
- **CRUD Operations**: Create, Read, Update, and Delete contractor records
- **Multi-Sheet Support**: Create and manage multiple Excel sheets
- **Search Functionality**: Real-time search across all data fields
- **Data Validation**: Comprehensive input validation with user-friendly error messages
- **Logging System**: Detailed logging for debugging and audit trails
- **Configuration Management**: Flexible configuration through App.config

## Recent Improvements

### 🔧 **Fixed Critical Issues**

1. **Hardcoded File Path**: 
   - **Before**: `"G:\hadeer\Contracts.xlsx"`
   - **After**: `Path.Combine(Application.StartupPath, "Contracts.xlsx")`
   - **Benefit**: Application now works on any machine

2. **Inconsistent Boolean Handling**:
   - **Before**: Mixed string ("✔"/"✖") and boolean values
   - **After**: Consistent boolean values throughout
   - **Benefit**: Data integrity and easier data processing

3. **Magic Numbers**:
   - **Before**: Hardcoded Excel constants like `-4162`
   - **After**: Named constants (`xlUp`, `xlLandscape`, etc.)
   - **Benefit**: Better code readability and maintainability

4. **Resource Management**:
   - **Before**: Potential COM object leaks
   - **After**: Proper cleanup with `Marshal.ReleaseComObject()`
   - **Benefit**: Prevents memory leaks and Excel process hanging

5. **Arabic Interface Restored**:
   - **Feature**: All user interface texts and messages restored to Arabic
   - **Benefit**: Maintains original Arabic user experience while keeping all technical improvements

### 🚀 **New Features**

1. **Comprehensive Logging System**:
   - Logs all operations (INFO, WARNING, ERROR)
   - Configurable through App.config
   - Helps with debugging and audit trails

2. **Advanced Data Validation**:
   - Contractor name validation (2-100 characters)
   - Work description validation (5-500 characters)
   - Date validation (reasonable date ranges)
   - Email and phone number format validation

3. **Configuration Management**:
   - Excel file name configurable
   - Logging settings configurable
   - Maximum records per sheet configurable

4. **Improved Error Handling**:
   - Detailed error messages
   - Proper exception logging
   - Graceful error recovery

## Project Structure

```
contarctors/
├── Form1.vb                 # Main application form
├── Form1.Designer.vb        # Form design code
├── Logger.vb                # Logging utility class
├── DataValidator.vb         # Data validation utility class
├── App.config              # Application configuration
├── contarctors.vbproj      # Project file
└── README.md               # This file
```

## Configuration

Edit `App.config` to customize application behavior:

```xml
<appSettings>
    <add key="ExcelFileName" value="Contracts.xlsx" />
    <add key="DefaultSheetName" value="Contracts" />
    <add key="MaxRecordsPerSheet" value="1000" />
    <add key="EnableLogging" value="true" />
    <add key="LogFilePath" value="app.log" />
</appSettings>
```

## Usage

1. **Start the Application**: Run the executable
2. **Select/Create Sheet**: Choose an existing sheet or create a new one
3. **Add Records**: Fill in contractor details and click "Save"
4. **Edit Records**: Select a record from the grid and modify fields
5. **Delete Records**: Select a record and click "Delete"
6. **Search**: Use the search box to filter records

## Data Fields

- **مسلسل (ID)**: Auto-generated sequential number
- **المقاول (Contractor)**: Contractor name (2-100 characters)
- **الأعمال (Work)**: Work description (5-500 characters)
- **تاريخ العقد (Contract Date)**: Date of contract signing
- **تاريخ الإسناد (Assignment Date)**: Date of work assignment
- **توقيع المقاول (Contractor Sign)**: Boolean flag for contractor signature
- **توقيع عام (General Sign)**: Boolean flag for general manager signature
- **تسليم (Delivery)**: Boolean flag for work delivery status

## Requirements

- .NET 8.0 or later
- Microsoft Excel (for COM automation)
- Windows operating system

## Building the Project

1. Open the solution in Visual Studio
2. Restore NuGet packages if needed
3. Build the solution (Ctrl+Shift+B)
4. Run the application (F5)

## Logging

The application creates detailed logs in `app.log` (configurable location):

```
2024-01-15 10:30:15 [INFO] Application started
2024-01-15 10:30:16 [INFO] Opened existing Excel file: C:\...\Contracts.xlsx
2024-01-15 10:30:20 [INFO] New record saved successfully. Contractor: ABC Corp, Work: Building Construction
```

## Troubleshooting

1. **Excel not opening**: Ensure Microsoft Excel is installed
2. **Permission errors**: Run as administrator if needed
3. **File access issues**: Check file permissions and antivirus settings
4. **COM errors**: Restart the application and ensure Excel is not running

## Future Enhancements

- [ ] Database backend option (SQL Server, SQLite)
- [ ] Export to PDF functionality
- [ ] Email notifications
- [ ] Multi-user support
- [ ] Advanced reporting
- [ ] Data backup/restore
- [ ] User authentication
- [ ] Audit trail
- [ ] Mobile app companion

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.
