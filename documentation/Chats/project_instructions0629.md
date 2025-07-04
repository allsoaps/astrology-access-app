# Project Instructions - Moving Forward

## Code Documentation Standards
**Since in-code comments are the ONLY documentation the next developer will have:**

### Required Comments for Every Function/Subroutine
```vba
' PURPOSE: [One line description of what this does]
' CALLED FROM: [Form/module that calls this]
' PARAMETERS: [Brief description of inputs]
' RETURNS: [What it returns, if anything]
' DEPENDENCIES: [Swiss Ephemeris, LocationIQ, specific tables]
' LAST MODIFIED: [Date and brief change description]
```

### Critical System Comments
- **Swiss Ephemeris functions**: Comment the DLL calls, file paths, error handling
- **LocationIQ integration**: Comment API endpoints, error scenarios
- **Database operations**: Comment complex SQL, relationship dependencies
- **Chart calculations**: Comment the astronomical logic and formulas

## Consistency Patterns to Maintain

### MS Access Database Standards
- **use DAO** for database operations
- **use SQL** for complex queries
- **use VBA** for complex logic
- **use VBA** for form events
- **use VBA** for form navigation
- **use VBA** for form validation
- **do not rewrite existing code unless explicitly asked to do so**

### Form Design Standards
- **Cancel buttons** on ALL forms that modify data
- **"Lose changes" warnings** when closing forms with unsaved data
- **Required field indicators** (consistent visual approach)
- **Error message format**: Use MsgBox with consistent tone and helpful guidance

### User Feedback Standards  
- **Success confirmations**: "Chart generated successfully" 
- **Clear error messages**: Tell user what went wrong AND what to do about it
- **Progress indicators**: Hourglass cursor for long operations
- **Validation messages**: Specific field requirements

### Database Operation Patterns
- **Soft deletes preferred** over hard deletes (mark as inactive/deleted)
- **Foreign key validation** before allowing deletions
- **Transaction wrapping** for multi-table operations
- **Consistent error handling** with proper resource cleanup

## One-Workflow-at-a-Time Checklist

### Before Declaring Any Workflow "Complete"
1. **End-to-end test** the entire workflow from start to finish
2. **Test error scenarios** (missing data, API failures, invalid inputs)
3. **Verify user feedback** (success messages, error handling, progress indicators)
4. **Check consistency** (Cancel buttons, change warnings, visual patterns)
5. **Add code comments** for any complex logic
6. **Test with tech-phobic mindset** (could user get confused or stuck?)

## Critical Technical Standards

### LocationIQ Error Handling
```vba
' ALWAYS provide fallback messaging for API failures
' NEVER leave user stuck without coordinates
' Example: "Unable to retrieve coordinates. Please check internet connection 
'          and try again, or contact support if problem persists."
```

### Swiss Ephemeris Error Handling
```vba
' ALWAYS check for successful DLL initialization
' ALWAYS provide specific error messages about missing files
' Example: "Swiss Ephemeris files not found. Please run the installer 
'          or contact support."
```

### Form State Management
- Track **saved vs unsaved state** consistently
- Use **module-level variables** (like mEventID) to track record state
- **Validate required data** before enabling dependent operations

## User Experience Guidelines

### For Tech-Phobic User
- **Multiple entry points** to workflows (if no Events exist, guide to Event creation)
- **Clear success confirmation** for every major operation
- **No technical jargon** in user-facing messages
- **Predictable behavior** across similar operations

### Message Box Standards
```vba
' SUCCESS: Green checkmark feeling
MsgBox "Chart generated successfully!", vbInformation, "Success"

' ERROR: Helpful guidance, not just error codes  
MsgBox "Unable to generate chart. Please ensure birth date and location are entered.", vbExclamation, "Chart Generation"

' WARNING: Clear consequences
MsgBox "Closing this form will lose your unsaved changes. Continue?", vbYesNo + vbQuestion, "Unsaved Changes"
```

## Development Workflow

### Testing Requirements
1. **Test on the clean laptop** for any significant changes
2. **Test Swiss Ephemeris connectivity** after any DLL-related changes
3. **Test LocationIQ** with various address formats
4. **End-to-end workflow test** after completing each major feature

### Version Control Best Practices
- **Backup .accdb file** before major changes
- **Test incrementally** - don't change multiple things at once
- **Document breaking changes** in code comments

## Delivery Readiness Criteria

### Must-Have Before Delivery
- [ ] All critical fixes from end-to-end testing completed
- [ ] LocationIQ error handling graceful (no user stuck scenarios)
- [ ] Consistent Cancel buttons and change warnings
- [ ] Code comments on all complex functions
- [ ] Swiss Ephemeris installer tested
- [ ] Core workflows (Viewer→Event→Session→Impressions) solid

### Nice-to-Have If Time Permits
- [ ] Consistent soft delete patterns
- [ ] Additional user guidance messages
- [ ] Enhanced error recovery options

## Risk Mitigation

### Highest Risks
1. **Swiss Ephemeris complexity** - Heavily comment all astronomical calculations
2. **LocationIQ dependency** - Ensure graceful failures don't block users
3. **Database relationships** - Comment foreign key dependencies clearly
4. **Next developer learning curve** - Code comments are critical

### Safety Nets
- **30-day warranty period** for critical issues
- **User's existing backup habits** for data protection
- **Swiss Ephemeris files** easily replaceable from GitHub

---

*These instructions prioritize maintainability, user experience, and successful handoff within the project constraints.*