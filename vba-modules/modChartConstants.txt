Attribute VB_Name = "modChartConstants"
Option Compare Database
Option Explicit

' Chart Type Constants - centralized location for all modules to reference
Public Const CHART_TYPE_NATAL As String = "Natal"
Public Const CHART_TYPE_EVENT As String = "Event"
Public Const CHART_TYPE_SESSION As String = "Session"

' Swiss Ephemeris Planet Constants
Public Const SE_SUN As Long = 0
Public Const SE_MOON As Long = 1
Public Const SE_MERCURY As Long = 2
Public Const SE_VENUS As Long = 3
Public Const SE_MARS As Long = 4
Public Const SE_JUPITER As Long = 5
Public Const SE_SATURN As Long = 6
Public Const SE_URANUS As Long = 7
Public Const SE_NEPTUNE As Long = 8
Public Const SE_PLUTO As Long = 9

' Swiss Ephemeris Calculation Flags
Public Const SEFLG_SPEED As Long = 256        ' Return speed values
Public Const SEFLG_SWIEPH As Long = 2         ' Use Swiss Ephemeris
Public Const SEFLG_GEOCENTRIC As Long = 0     ' Default geocentric
Public Const SEFLG_HELIOCENTRIC As Long = 8   ' Heliocentric flag

' Calendar Constants
Public Const SE_GREG_CAL As Long = 1          ' Gregorian calendar

' House System Constants
Public Const SE_PLACIDUS As Long = 80         ' 'P' in ASCII
