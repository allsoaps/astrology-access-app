# frmEventNew Development - Lessons Learned

## Swiss Ephemeris Integration
- **Critical Fix**: Initialize error strings with `vbNullString` instead of `Space$(256)` to prevent hard crashes
- **DLL Management**: Use proper 64-bit vs 32-bit DLL declarations with PtrSafe attributes
- **Environment Variables**: Swiss Ephemeris can use system environment variables for ephemeris path configuration
- **Parameter Handling**: Arrays must be properly dimensioned and passed by reference

## Database Design Patterns
- **Relationship Structure**: Events → Charts → ChartPositions provides clean data separation
- **EventID vs PersonID**: Added EventID field to tblCharts for semantic clarity instead of overloading PersonID
- **Duplicate Prevention**: Always check for existing records before inserting to prevent duplicates
- **Module Variables**: Use form-level variables (mEventID) to track saved record state

## Error Handling Best Practices
- **DAO vs ADO**: DAO recordsets don't have State property - use proper error handling when closing
- **Consistent Cleanup**: Always clean up database objects in error handlers
- **User Feedback**: Provide clear messages about what succeeded vs failed
- **Debug Information**: Use Debug.Print statements for troubleshooting complex operations

## Form Design Principles
- **Conditional UI**: Show/hide controls based on user selections (sports teams for sports events)
- **Save Without Close**: Keep forms open after saving for better user experience
- **Validation First**: Always validate required fields before attempting database operations
- **Progress Indication**: Inform users about long-running operations like chart generation

## API Integration Lessons
- **Caching Strategy**: Check local database before making external API calls
- **Error Handling**: Handle API failures gracefully with user-friendly messages
- **Data Validation**: Validate API responses before using the data
- **Performance**: Minimize external calls through intelligent caching

## Code Maintenance Issues
- **Version Control**: Changes to working code often reintroduce previously fixed bugs
- **Incremental Updates**: Add new functionality without modifying working code sections
- **Testing**: Test each change independently before adding more features
- **Documentation**: Keep track of what works to avoid regression

## Data Type Management
- **Null Handling**: Use proper SQL NULL syntax in dynamic queries
- **Date Formatting**: Use consistent date formats for SQL statements
- **String Escaping**: Always escape single quotes in SQL strings
- **Type Conversion**: Handle variant types carefully when dealing with optional fields

## User Experience Insights
- **Duplicate Detection**: Ask users before overwriting existing records
- **Clear Messaging**: Distinguish between "save new" vs "update existing" operations
- **Visual Feedback**: Use checkboxes and status indicators to show completion
- **Error Recovery**: Provide options when operations fail rather than just error messages