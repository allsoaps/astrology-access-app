# Swiss Ephemeris Deployment and Troubleshooting Guide

## Problem Overview
When deploying an MS Access application that uses Swiss Ephemeris, users may encounter errors like:
- "Failed to initialize Swiss Ephemeris: File not found: swedll64.dll"
- "Bad file name or number" (Error 52)
- "Ambiguous name detected: DateToJulianDay"

## Step-by-Step Deployment Solution

### 1. Create Required Folder Structure
In the same folder where your Access database (.accdb) file is located, create this exact folder structure:

```
[Database Folder]
├── YourDatabase.accdb
└── Resources/
    ├── swedll32.dll
    ├── swedll64.dll
    └── ephe/
        ├── sepl_18.se1
        ├── semo_18.se1
        ├── sepl_06.se1
        └── [other .se1 ephemeris files]
```

### 2. Download Swiss Ephemeris Files
1. Go to: https://www.astro.com/swisseph/
2. Download the following:
   - Swiss Ephemeris DLL files (swedll32.dll and swedll64.dll)
   - Ephemeris data files (.se1 files)
3. Copy DLL files to `Resources` folder
4. Copy .se1 files to `Resources\ephe` folder

### 3. Fix Common Compilation Errors

#### Error: "Ambiguous name detected: DateToJulianDay"
**Solution:** You have duplicate functions with the same name
1. Press Ctrl+F in VBA editor
2. Search for "DateToJulianDay" across all modules
3. Delete duplicate function definitions
4. Keep only one version (preferably from the main Swiss Ephemeris module)

#### Error: "Method or data member not found" (rs.State)
**Solution:** DAO recordsets don't have a State property
Replace this code:
```vba
If rs.State = 1 Then rs.Close
```
With this:
```vba
rs.Close
```

### 4. Debugging File Path Issues

If you get "Bad file name or number" errors, add these debug lines to check paths:

```vba
Debug.Print "Database Path: " & CurrentProject.Path
Debug.Print "Expected DLL: " & GetExpectedDLLPath()
Debug.Print "Ephe Path: " & m_EphePath
Debug.Print "DLL Exists: " & (Dir(GetExpectedDLLPath()) <> "")
Debug.Print "Ephe Folder Exists: " & (Dir(m_EphePath, vbDirectory) <> "")
```

### 5. Test the Installation

After setting up files, test in VBA Immediate Window (Ctrl+G):
```vba
? InitSwissEph()
```
Should return `True` if successful.

### 6. Database Settings Table

The code expects a table called `tblSwissEphSettings` with these fields:
- SettingID (Primary Key)
- EphePath (Text) - Path to ephemeris data folder
- DLLPath (Text) - Path to DLL files
- DefaultHouseSystem (Number)
- DefaultOrbs (Yes/No)
- IncludeAsteroids (Yes/No)
- UseTopocentric (Yes/No)
- UseTrue (Yes/No)

The code will auto-create this table with default values if it doesn't exist.

## Common Error Messages and Solutions

### "File not found: swedll64.dll"
**Cause:** DLL files are missing or in wrong location
**Solution:** Ensure swedll32.dll and swedll64.dll are in the Resources folder

### "Bad file name or number" (Error 52)
**Cause:** 
- DLL files missing
- Ephemeris data files missing  
- Invalid file paths
- Permission issues
**Solution:** 
- Verify all files exist in correct locations
- Check folder permissions
- Use debug code above to verify paths

### "Ambiguous name detected"
**Cause:** Duplicate function names across modules
**Solution:** Search for and remove duplicate function definitions

### "Method or data member not found"
**Cause:** Using ADO syntax with DAO objects
**Solution:** Remove State property checks for DAO recordsets

## File Checklist for Deployment

Before deploying, ensure these files exist:

**Required DLL Files:**
- [ ] Resources/swedll32.dll
- [ ] Resources/swedll64.dll

**Required Ephemeris Files (minimum):**
- [ ] Resources/ephe/sepl_18.se1 (planets)
- [ ] Resources/ephe/semo_18.se1 (moon)
- [ ] Resources/ephe/sepl_06.se1 (older planets data)

**Database Components:**
- [ ] tblSwissEphSettings table exists
- [ ] Swiss Ephemeris VBA module compiled without errors
- [ ] No duplicate function names across modules

## Verification Commands

Run these in VBA Immediate Window to verify setup:

```vba
' Test initialization
? InitSwissEph()

' Check paths
? CurrentProject.Path

' Test planet calculation
? GetPlanetPosition(0, Now)  ' Sun position now

' Verify settings table
? DCount("*", "tblSwissEphSettings")
```

## Support Notes

- The code automatically detects 32-bit vs 64-bit systems
- Default paths are relative to database location
- All file paths can be customized via tblSwissEphSettings
- Error messages now provide specific guidance on missing files
- Code includes automatic fallback to default paths if database settings fail

## Final Deployment Checklist

1. [ ] Database and Resources folder in same directory
2. [ ] Both 32-bit and 64-bit DLLs present
3. [ ] Ephemeris data files in ephe subfolder
4. [ ] VBA code compiles without errors
5. [ ] InitSwissEph() returns True
6. [ ] Basic planet calculation works
7. [ ] No duplicate function names
8. [ ] tblSwissEphSettings table exists and populated

Following this guide should resolve all common Swiss Ephemeris deployment issues in MS Access applications.