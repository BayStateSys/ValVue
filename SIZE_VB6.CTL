VERSION 5.00
Begin VB.UserControl SizingCtrl 
   ClientHeight    =   165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   195
   ScaleHeight     =   165
   ScaleWidth      =   195
End
Attribute VB_Name = "SizingCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_TestProperty = 0
'Property Variables:
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
Dim m_TestProperty As Variant

Private g_str_DatabasePath As String
'Private wrk As Workspace
Private g_db_Prod As Database
Private g_db_Gensz As Database

Private Declare Sub GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String)
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private DbPath As String

#Const DebugVersion = "yes"
Const PI2 = 9.8696047
Const PI = 3.1415927
Const ANSIMESSAGE = "SPECIFIED PRESSURE EXCEEDS ANSI LIMIT"
Const STEMMESSAGE = "ACTUATOR OUTPUT EXCEEDS STEM LIMIT"
Const ACTUATORMESSAGEC = "ACTUATOR CAPABILITY EXCEEDED (CLOSED)"
Const ACTUATORMESSAGE = "ACTUATOR CAPABILITY EXCEEDED"
Const TRIMMESSAGE = "SPECIFIED PRESSURE EXCEEDS TRIM LIMIT"
Const LIMITEDDPMESSAGE = "WARNING - MAX INLET PRESSURE EXCEEDS ACTUATOR SHUTOFF LIMIT"
Const SUPPLYMESSAGE = "ACTUATOR CAPABILITY EXCEEDED (AIR SUPPLY PRESSURE)"
Const INVALIDACTUATOR = "INVALID ACTUATOR SELECTION"

Const ATO = 6
Const ATC = 7
Const FTO = 1
Const FTC = 2

Const dbl_Pl = 3.14159
Const dbl_Ps = 101325
Const dbl_Tc = 647.3
Const dbl_Pc = 221.2
Const dbl_Hc = 70.1204

Function GetStringValueFromINIFile(str_AppName As String, str_KeyName As String, str_FileName As String) As String
Dim str_ReturnValue As String

'------------------------------------------------------------------------------------------
'This function is used to obtain string input variables for frmTest.
'------------------------------------------------------------------------------------------

str_ReturnValue = Space(128)

GetPrivateProfileString str_AppName, str_KeyName, "Not found", str_ReturnValue, Len(str_ReturnValue), str_FileName
str_ReturnValue = Trim(str_ReturnValue)
GetStringValueFromINIFile = Left(str_ReturnValue, InStr(1, str_ReturnValue, Chr(0)) - 1)
End Function

Function GetWinDir() As String
Dim str_WinDir As String
Dim lng_StrLen As Long

DbPath = String(255, 32)
str_WinDir = String(255, 32)
lng_StrLen = GetWindowsDirectory(str_WinDir, Len(str_WinDir))
GetWinDir = Left(str_WinDir, lng_StrLen) + "\"
End Function


Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub

Private Sub UserControl_Initialize()
On Error GoTo Init_error
g_str_DatabasePath = GetStringValueFromINIFile("Database", "DbPath", "SizingOCX.INI")


'Set wrk = CreateWorkspace("DeepSpace", "Admin", "", dbUseODBC)
'Set g_db_Prod = wrk.OpenDatabase("GSTPrdDSN", , True, "ODBC;UID=Admin;PWD=;DSN=GSTPrdDSN")
'Set g_db_Gensz = wrk.OpenDatabase("GSTSizDSN", , True, "ODBC;UID=Admin;PWD=;DSN=GSTSizDSN")

MsgBox "Rev 6"

Set g_db_Prod = OpenDatabase(g_str_DatabasePath + "Product.mdb")
Set g_db_Gensz = OpenDatabase(g_str_DatabasePath + "Gensz.mdb")

Exit Sub
Init_error:
MsgBox ("Error opening databases = " + Error)
End Sub

'Public Property Get TestProperty() As Variant
'    TestProperty = m_TestProperty
'End Property
'
'Public Property Let TestProperty(ByVal New_TestProperty As Variant)
'    m_TestProperty = New_TestProperty
'    PropertyChanged "TestProperty"
'End Property


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
'    m_TestProperty = m_def_TestProperty
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
'    m_TestProperty = PropBag.ReadProperty("TestProperty", m_def_TestProperty)
End Sub

Private Sub UserControl_Terminate()
g_db_Prod.Close
g_db_Gensz.Close
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
'    Call PropBag.WriteProperty("TestProperty", m_TestProperty, m_def_TestProperty)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INCLUDE Actuator UL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetReducedFlag(long_TrimPackageID As Long, bln_ReducedArea As Boolean)
On Error GoTo GetReducedFlag_Error
'------------------------------------------------------------------------------------------
' This routine retrieves the reduced area flag fromt the trim package table.
'------------------------------------------------------------------------------------------
'11/25/98 Schoonover
'------------------------------------------------------------------------------------------
'validation
If (long_TrimPackageID <= 0) Then GoTo GetReducedFlag_Error

'declarations
Dim rs As Recordset
Dim str_Query As String
Dim str_Path As String

'process

str_Query = "select REDUCED_AREA from tblPackage where PACKAGE_ID=" + CStr(long_TrimPackageID)
Set rs = g_db_Prod.OpenRecordset(str_Query, dbReadOnly)
If (rs.RecordCount = 0) Then GoTo GetReducedFlag_Error
bln_ReducedArea = rs.Fields!REDUCED_AREA

rs.Close

Exit Sub
GetReducedFlag_Error:
MsgBox "Error getting Reduced Area Flag =" + CStr(long_TrimPackageID) + " = " + Error
End Sub

Public Sub GetTrimData(long_TrimPackageID As Long, _
    long_CharacteristicID As Long, dbl_RatedCv As Double, _
    dbl_StemTravel As Double, dbl_DynamicFactor As Double, _
    dbl_StemStressArea As Double, dbl_PackingHeight As Double, _
    dbl_StemSize As Double, dbl_Eccentricity As Double, _
    dbl_PlugDiameter As Double, dbl_AllowPlugDP _
    As Double, dbl_SeatRingDiameter As Double, dbl_MaxSupply As Double, _
    dbl_MaxTrimTorque As Double, dbl_MaxShaftDPOpen As Double, _
    dbl_MaxShaftDPClosed As Double, dbl_MaxSupplyStability As Double, _
    dbl_Kv As Double, dbl_Fcs As Double, dbl_Au As Double, _
    str_StemMaterial As String)
    
'--------------------------------------------------------------------------------------
'Helper routine to retrieve data from the database - this routine is called
'   by Get21000Data, Get41000Data, etc.
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID     package ID of trim
'
'Output Variables:
'long_CharacteristicID
'dbl_RatedCV            Rated Cv of the control valve
'dbl_StemTravel
'dbl_DynamicFactor      factor used to predict dynamic stability        inches
'dbl_StemStressArea
'dbl_PackingHeight      height of packing set                           inches
'dbl_StemSize           diameter of stem                                inches
'dbl_Eccentricity       (dist. frm ctr of shaft to ctr of seat ring)    inches
'dbl_PlugDiameter       diameter of plug seat                           inches
'dbl_AllowPlugDP        pressure limit of plug                          psi
'dbl_SeatRingDiameter   diameter of seat ring orifice                   inches
'dbl_MaxSupply          maximum allowable supply pressure               psi
'dbl_MaxTrimTorque
'dbl_MaxShaftDPOpen
'dbl_MaxShaftDPClosed
'dbl_MaxSupplyStability
'dbl_Kv                 unbalanced area rate of change                  in^2/in
'dbl_Fcs                force needed to compress plug pilot spring      lbs
'dbl_Au                 effective unbalanced area                       in^2
'str_StemMaterial
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/17 Larry Schoonover
'Modified 9/16/98 to match changes in database - Larry Schoonover
'-----------------------------------------------------------------------------------------

On Error GoTo GetTrimData_Error

'***************************
'Declarations
'***************************
Dim str_DatabasePath As String
Dim str_TrimQuery As String
Dim rset_Trim As Recordset

'***************************
'Parameter validation
'***************************
If (long_TrimPackageID <= 0) Then GoTo GetTrimData_Error

'****************************
'Begin calculations
'****************************

str_TrimQuery = _
      "SELECT tblGroupData.CHARACTERISTIC_ID, tblGroupData.RATED_CV, " _
    & "tblGroupData.TRAVEL, tblGroupData.DYNAMIC_FACTOR, " _
    & "tblGroupData.STRESS_AREA, tblGroupData.PACKING_HEIGHT, " _
    & "tblGroupData.STEM_SIZE, tblGroupData.ECCENTRICITY, tblGroupData.PLUG_DIAMETER, " _
    & "tblGroupData.PLUG_DP, tblGroupData.SEAT_RING_DIAMETER, tblGroupData.MAX_SUPPLY, " _
    & "tblGroupData.MAX_TRIM_TORQUE, tblGroupData.MAX_SHAFT_DP_OPEN, " _
    & "tblGroupData.MAX_SHAFT_DP_CLOSED, tblGroupData.MAX_SUPPLY_STABILITY, " _
    & "tblGroupData.KV, tblGroupData.FCS, tblGroupData.AU, tblMaterial.MATERIAL " _
    & "FROM (((tblMaterial INNER JOIN tblPart ON tblMaterial.MATERIAL_ID = " _
    & "tblPart.MATERIAL_ID) INNER JOIN tblTrim ON tblPart.PART_ID = tblTrim.PART_ID) " _
    & "INNER JOIN ((tblGroupData INNER JOIN tblPackage ON tblGroupData.GROUP_ID = " _
    & "tblPackage.GROUP_ID) INNER JOIN tblTrimPackage ON tblPackage.PACKAGE_ID = " _
    & "tblTrimPackage.PACKAGE_ID) ON tblTrim.PART_ID = tblTrimPackage.PART_ID) " _
    & "INNER JOIN tblTrimCompDesc ON tblTrim.TRIM_COMP_ID = tblTrimCompDesc.TRIM_COMP_ID " _
    & "WHERE (((tblPackage.PACKAGE_ID)=?) AND (((tblTrimCompDesc.TRIM_COMP)='?') OR ((tblTrimCompDesc.TRIM_COMP)='?')))"


str_TrimQuery = ReplaceStr(str_TrimQuery, "?", Str$(long_TrimPackageID), "Stem", "Shaft")

Set rset_Trim = g_db_Prod.OpenRecordset(str_TrimQuery, dbReadOnly)
If (rset_Trim.RecordCount <> 1) Then GoTo GetTrimData_Error
long_CharacteristicID = IIf(IsNull(rset_Trim.Fields(0)), 0#, rset_Trim.Fields(0))
dbl_RatedCv = IIf(IsNull(rset_Trim.Fields(1)), 0#, rset_Trim.Fields(1))
dbl_StemTravel = IIf(IsNull(rset_Trim.Fields(2)), 0#, rset_Trim.Fields(2))
dbl_DynamicFactor = IIf(IsNull(rset_Trim.Fields(3)), 0#, rset_Trim.Fields(3))
dbl_StemStressArea = IIf(IsNull(rset_Trim.Fields(4)), 0#, rset_Trim.Fields(4))
dbl_PackingHeight = IIf(IsNull(rset_Trim.Fields(5)), 0#, rset_Trim.Fields(5))
dbl_StemSize = IIf(IsNull(rset_Trim.Fields(6)), 0#, rset_Trim.Fields(6))
dbl_Eccentricity = IIf(IsNull(rset_Trim.Fields(7)), 0#, rset_Trim.Fields(7))
dbl_PlugDiameter = IIf(IsNull(rset_Trim.Fields(8)), 0#, rset_Trim.Fields(8))
dbl_AllowPlugDP = IIf(IsNull(rset_Trim.Fields(9)), 0#, rset_Trim.Fields(9))
dbl_SeatRingDiameter = IIf(IsNull(rset_Trim.Fields(10)), 0#, rset_Trim.Fields(10))
dbl_MaxSupply = IIf(IsNull(rset_Trim.Fields(11)), 0#, rset_Trim.Fields(11))
dbl_MaxTrimTorque = IIf(IsNull(rset_Trim.Fields(12)), 0#, rset_Trim.Fields(12))
dbl_MaxShaftDPOpen = IIf(IsNull(rset_Trim.Fields(13)), 0#, rset_Trim.Fields(13))
dbl_MaxShaftDPClosed = IIf(IsNull(rset_Trim.Fields(14)), 0#, rset_Trim.Fields(14))
dbl_MaxSupplyStability = IIf(IsNull(rset_Trim.Fields(15)), 0#, rset_Trim.Fields(15))
dbl_Kv = IIf(IsNull(rset_Trim.Fields(16)), 0#, rset_Trim.Fields(16))
dbl_Fcs = IIf(IsNull(rset_Trim.Fields(17)), 0#, rset_Trim.Fields(17))
dbl_Au = IIf(IsNull(rset_Trim.Fields(18)), 0#, rset_Trim.Fields(18))
str_StemMaterial = IIf(IsNull(rset_Trim.Fields(19)), 0#, rset_Trim.Fields(19))
'dbl_ValveSize = IIf(IsNull(rset_Trim.Fields(18)), 0#, rset_Trim.Fields(18))

rset_Trim.Close

Exit Sub

GetTrimData_Error:
'Placeholder.  How do we want to handle error messages?

End Sub
Public Sub Get2600Data(long_TrimPackageID As Long, dbl_DynamicFactor As Double, _
    dbl_PackingHeight As Double, _
    dbl_PlugDiameter As Double, dbl_SeatDiameter As Double, dbl_StemDiameter As Double, _
    dbl_AllowPlugDP As Double, str_StemMaterial As String, dbl_StemTravel As Double)
'--------------------------------------------------------------------------------------
'Reference:           Actuator sizing methods for reciprocating valves.
'Solutions reference: 21000.frm, subroutines: getStemDia, getSeatDia.
'---------------------------------------------------------------------------------------
'This subroutine is used to retrieve data for 21000 series valves.
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID        package ID of trim
'
'Output Variables:
'dbl_DynamicFactor         factor used to predict dynamic stability        inches
'dbl_PlugDiameter          diameter of plug seat                           inches
'dbl_StemDiameter          diameter of plug stem                           inches
'dbl_SeatDiameter          diameter of seat ring orifice                   inches
'dbl_AllowPlugDP           pressure limit of plug                          psi
'str_StemMaterial
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 12/29 Larry Schoonover
'-----------------------------------------------------------------------------------------

Dim dbl_NotUsed As Double
Dim str_NotUsed As String
Dim long_NotUsed As Long

On Error GoTo Get2600Data_Error

Call GetTrimData(long_TrimPackageID, long_NotUsed, dbl_NotUsed, _
    dbl_StemTravel, dbl_DynamicFactor, dbl_NotUsed, _
    dbl_PackingHeight, dbl_StemDiameter, dbl_NotUsed, dbl_PlugDiameter, _
    dbl_AllowPlugDP, dbl_SeatDiameter, dbl_NotUsed, dbl_NotUsed, dbl_NotUsed, _
    dbl_NotUsed, dbl_NotUsed, dbl_NotUsed, dbl_NotUsed, dbl_NotUsed, str_StemMaterial)

#If (DebugVersion = "yes") Then
MsgBox "Get2600Data" + Chr(13) + Chr(10) + _
    "dbl_DynamicFactor=" + Str(dbl_DynamicFactor) + Chr(13) + Chr(10) + _
    "dbl_PackingHeight=" + Str(dbl_PackingHeight) + Chr(13) + Chr(10) + _
    "dbl_StemDiameter=" + Str(dbl_StemDiameter) + Chr(13) + Chr(10) + _
    "dbl_PlugDiameter=" + Str(dbl_PlugDiameter) + Chr(13) + Chr(10) + _
    "dbl_AllowPlugDP=" + Str(dbl_AllowPlugDP) + Chr(13) + Chr(10) + _
    "dbl_SeatDiameter=" + Str(dbl_SeatDiameter) + Chr(13) + Chr(10) + _
    "str_StemMaterial=" + str_StemMaterial + Chr(13) + Chr(10) + _
    "dbl_StemTravel=" + Str(dbl_StemTravel) + Chr(13) + Chr(10)

#End If

Exit Sub
Get2600Data_Error:
End Sub

Public Sub Get21000Data(long_TrimPackageID As Long, dbl_DynamicFactor As Double, _
    dbl_PlugDiameter As Double, dbl_SeatDiameter As Double, _
    dbl_StemDiameter As Double, dbl_PackingHeight As Double, _
    dbl_AllowPlugDP As Double, dbl_StemStressArea As Double, _
    dbl_Travel As Double, str_StemMaterial As String)
    
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for reciprocating valves.
'Solutions reference: 21000.frm, subroutines: getStemDia, getSeatDia.
'---------------------------------------------------------------------------------------
'This subroutine is used to retrieve data for 21000 series valves.
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID         package ID of trim
'-----------------------------------------------------------------------------------------
'Output Variables:
'dbl_StemDiameter       diameter of plug stem                           inches
'dbl_SeatDiameter       diameter of seat ring orifice                   inches
'dbl_Kv                 unbalanced area rate of change                  in^2/in
'dbl_Fcs                force needed to compress plug pilot spring      lbs
'dbl_Au                 effective unbalanced area                       in^2
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/12 Larry Schoonover - not tested
'6/12 -Larry Schoonover -tested - passed
'9/16 Modified to match changes in data base - Larry Schoonover
'-----------------------------------------------------------------------------------------

On Error GoTo Get21000Data_Error

'***************************
'Declarations
'***************************
Dim dbl_NotUsed As Double
Dim long_NotUsed As Long

'***************************
'Parameter validation
'***************************
If (long_TrimPackageID <= 0) Then GoTo Get21000Data_Error

'****************************
'Begin calculations
'****************************
Call GetTrimData(long_TrimPackageID, long_NotUsed, dbl_NotUsed, _
    dbl_Travel, dbl_DynamicFactor, dbl_StemStressArea, _
    dbl_PackingHeight, dbl_StemDiameter, dbl_NotUsed, dbl_PlugDiameter, _
    dbl_AllowPlugDP, dbl_SeatDiameter, dbl_NotUsed, dbl_NotUsed, dbl_NotUsed, _
    dbl_NotUsed, dbl_NotUsed, dbl_NotUsed, dbl_NotUsed, dbl_NotUsed, str_StemMaterial)
    
#If (DebugVersion = "yes") Then
MsgBox "Get21000Data" + Chr(13) + Chr(10) + _
    "dbl_Travel=" + Str(dbl_Travel) + Chr(13) + Chr(10) + _
    "dbl_DynamicFactor=" + Str(dbl_DynamicFactor) + Chr(13) + Chr(10) + _
    "dbl_StemStressArea=" + Str(dbl_StemStressArea) + Chr(13) + Chr(10) + _
    "dbl_PackingHeight=" + Str(dbl_PackingHeight) + Chr(13) + Chr(10) + _
    "dbl_PlugDiameter, _=" + Str(dbl_PlugDiameter) + Chr(13) + Chr(10) + _
    "dbl_AllowPlugDP=" + Str(dbl_AllowPlugDP) + Chr(13) + Chr(10) + _
    "dbl_SeatDiameter=" + Str(dbl_SeatDiameter) + Chr(13) + Chr(10) + _
    "str_StemMaterial=" + str_StemMaterial + Chr(13) + Chr(10)
#End If
    

Exit Sub

Get21000Data_Error:
'Placeholder.  How do we want to handle error messages?
    
    
End Sub

Public Sub Get41000Data(long_TrimPackageID As Long, dbl_RatedCv As Double, _
    dbl_StemDiameter As Double, dbl_SeatDiameter As Double, dbl_Kv As Double, _
    dbl_Au As Double, dbl_PackingHeight As Double, _
    dbl_StemStressArea As Double, str_StemMaterial As String, dbl_ForcePilotSpring As Double, _
    dbl_Travel As Double, dbl_PlugSealDiameter As Double, dbl_PlugThrottlingDiameter As Double, _
    dbl_DynPlugDiameter As Double, dbl_SealRingCoef As Double, dbl_TECSealCoef As Double, _
    dbl_TECCoefJ As Double, dbl_TECCoefS As Double, dbl_CageInsideDiameter As Double, _
    dbl_SealRingWidth As Double, dbl_ForcePlugWt As Double, dbl_MaxAllowFlatSpringDP As Double, _
    dbl_SeatingDiameter As Double, dbl_PilotTravel As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for reciprocating valves.
'Solutions reference: 41000.frm, subroutines: getStemDia, getSeatDia.
'---------------------------------------------------------------------------------------
'This subroutine is used to retrieve data for 41000 series valves.
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID        package ID of trim
'
'Output Variables:
'dbl_RatedCv
'dbl_StemDiameter          diameter of plug stem                           inches
'dbl_SeatDiameter          diameter of seat ring orifice                   inches
'dbl_Kv                    unbalanced area rate of change                  in^2/in
'dbl_Au                    effective unbalanced area                       in^2
'dbl_PackingHeight
'dbl_StemStressArea
'str_StemMaterial
'dbl_Travel                rated travel of specified trim                  inches
'dbl_PlugSealDiameter      diameter of balanced plug next to               inches
'                          seal ring groove
'dbl_PlugThrottlingDiameter diameter at bottom of plug                     inches
'dbl_DynPlugDiameter       dynamically unaffected diameter                 inches
'dbl_SealRingCoef          balanced valve seal ring friction coef.
'
'dbl_TECsealCoef           balanced valve TEC seal friction coef.
'dbl_TECcoefJ              seal ring friction coef (area)
'dbl_TECcoefS              seal ring friction coef (spring)
'dbl_CageInsideDiameter    inside diameter of cage                         inches
'                          (seal ring contact dia.)
'dbl_SealRingWidth         width of balanced plug seal ring                inches
'dbl_SeatingDiameter       contact diameter (plug/seat ring)               inches
'dbl_PilotTravel           nominal travel of pilot plug (if applicable)    inches
'dbl_ForcePlugWt           weight of valve plug                            lbs
'dbl_MaxAllowFlatSpringDP  maximum allowable pressure drop
'                          across flat spring (cage seal)                  psi
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/12 Larry Schoonover - not tested
'6/12 -Larry Schoonover -tested - passed
'9/16 modified to match new database - Larry Schoonover
'11/9 modified for new 41000 methods
'-----------------------------------------------------------------------------------------

On Error GoTo Get41000Data_Error

'***************************
'Declarations
'***************************
Dim str_DatabasePath As String
Dim str_TrimQuery As String
Dim rset_Trim As Recordset

'***************************
'Parameter validation
'***************************
If (long_TrimPackageID <= 0) Then GoTo Get41000Data_Error

'****************************
'Begin calculations
'****************************

str_TrimQuery = _
      "SELECT tblGroupData.RATED_CV, tblGroupData.STEM_SIZE, tblGroupData.SEAT_RING_DIAMETER, " _
    & "tblGroupData.TRAVEL, tblGroupData.KV, " _
    & "tblGroupData.STRESS_AREA, tblGroupData.PACKING_HEIGHT, " _
    & "tblGroupData.AU, tblGroupData.FORCE_PLUG_SPRING, tblGroupData.PLUG_SEAL_DIAMETER, " _
    & "tblGroupData.PLUG_THROTTLE_DIAMETER, tblGroupData.DYN_PLUG_DIAMETER, tblGroupData.SEAL_RING_COEF, " _
    & "tblGroupData.TEC_SEAL_COEF, tblGroupData.TEC_COEF_J, " _
    & "tblGroupData.TEC_COEF_S, tblGroupData.CAGE_INSIDE_DIAMETER, " _
    & "tblGroupData.SEAL_WIDTH, tblGroupData.FORCE_PLUG_WEIGHT, tblGroupData.PILOT_TRAVEL, " _
    & "tblGroupData.FLAT_SPRING_MAX_DP, tblGroupData.SEATING_DIAMETER, tblMaterial.MATERIAL " _
    & "FROM (((tblMaterial INNER JOIN tblPart ON tblMaterial.MATERIAL_ID = " _
    & "tblPart.MATERIAL_ID) INNER JOIN tblTrim ON tblPart.PART_ID = tblTrim.PART_ID) " _
    & "INNER JOIN ((tblGroupData INNER JOIN tblPackage ON tblGroupData.GROUP_ID = " _
    & "tblPackage.GROUP_ID) INNER JOIN tblTrimPackage ON tblPackage.PACKAGE_ID = " _
    & "tblTrimPackage.PACKAGE_ID) ON tblTrim.PART_ID = tblTrimPackage.PART_ID) " _
    & "INNER JOIN tblTrimCompDesc ON tblTrim.TRIM_COMP_ID = tblTrimCompDesc.TRIM_COMP_ID " _
    & "WHERE (((tblPackage.PACKAGE_ID)=?) AND ((tblTrimCompDesc.TRIM_COMP)='?'))"



str_TrimQuery = ReplaceStr(str_TrimQuery, "?", Str$(long_TrimPackageID), "stem")

Set rset_Trim = g_db_Prod.OpenRecordset(str_TrimQuery, dbReadOnly)
If (rset_Trim.RecordCount <> 1) Then GoTo Get41000Data_Error

dbl_RatedCv = IIf(IsNull(rset_Trim.Fields!RATED_CV), 0#, rset_Trim.Fields!RATED_CV)
dbl_StemDiameter = IIf(IsNull(rset_Trim.Fields!STEM_SIZE), 0#, rset_Trim.Fields!STEM_SIZE)
dbl_SeatDiameter = IIf(IsNull(rset_Trim.Fields!SEAT_RING_DIAMETER), 0#, rset_Trim.Fields!SEAT_RING_DIAMETER)
dbl_Kv = IIf(IsNull(rset_Trim.Fields!KV), 0#, rset_Trim.Fields!KV)
dbl_Au = IIf(IsNull(rset_Trim.Fields!AU), 0#, rset_Trim.Fields!AU)
dbl_PackingHeight = IIf(IsNull(rset_Trim.Fields!PACKING_HEIGHT), 0#, rset_Trim.Fields!PACKING_HEIGHT)
dbl_StemStressArea = IIf(IsNull(rset_Trim.Fields!STRESS_AREA), 0#, rset_Trim.Fields!STRESS_AREA)
str_StemMaterial = IIf(IsNull(rset_Trim.Fields!MATERIAL), 0#, rset_Trim.Fields!MATERIAL)
dbl_ForcePilotSpring = IIf(IsNull(rset_Trim.Fields!FORCE_PLUG_SPRING), 0#, rset_Trim.Fields!FORCE_PLUG_SPRING)
dbl_Travel = IIf(IsNull(rset_Trim.Fields!TRAVEL), 0#, rset_Trim.Fields!TRAVEL)
dbl_PlugSealDiameter = IIf(IsNull(rset_Trim.Fields!PLUG_SEAL_DIAMETER), 0#, rset_Trim.Fields!PLUG_SEAL_DIAMETER)
dbl_PlugThrottlingDiameter = IIf(IsNull(rset_Trim.Fields!PLUG_THROTTLE_DIAMETER), 0#, rset_Trim.Fields!PLUG_THROTTLE_DIAMETER)
dbl_DynPlugDiameter = IIf(IsNull(rset_Trim.Fields!DYN_PLUG_DIAMETER), 0#, rset_Trim.Fields!DYN_PLUG_DIAMETER)
dbl_SealRingCoef = IIf(IsNull(rset_Trim.Fields!SEAL_RING_COEF), 0#, rset_Trim.Fields!SEAL_RING_COEF)
dbl_TECSealCoef = IIf(IsNull(rset_Trim.Fields!TEC_SEAL_COEF), 0#, rset_Trim.Fields!TEC_SEAL_COEF)
dbl_TECCoefJ = IIf(IsNull(rset_Trim.Fields!TEC_COEF_J), 0#, rset_Trim.Fields!TEC_COEF_J)
dbl_TECCoefS = IIf(IsNull(rset_Trim.Fields!TEC_COEF_S), 0#, rset_Trim.Fields!TEC_COEF_S)
dbl_CageInsideDiameter = IIf(IsNull(rset_Trim.Fields!CAGE_INSIDE_DIAMETER), 0#, rset_Trim.Fields!CAGE_INSIDE_DIAMETER)
dbl_SealRingWidth = IIf(IsNull(rset_Trim.Fields!SEAL_WIDTH), 0#, rset_Trim.Fields!SEAL_WIDTH)
dbl_ForcePlugWt = IIf(IsNull(rset_Trim.Fields!FORCE_PLUG_WEIGHT), 0#, rset_Trim.Fields!FORCE_PLUG_WEIGHT)
dbl_MaxAllowFlatSpringDP = IIf(IsNull(rset_Trim.Fields!FLAT_SPRING_MAX_DP), 0#, rset_Trim.Fields!FLAT_SPRING_MAX_DP)
dbl_SeatingDiameter = IIf(IsNull(rset_Trim.Fields!SEATING_DIAMETER), 0#, rset_Trim.Fields!SEATING_DIAMETER)
dbl_PilotTravel = IIf(IsNull(rset_Trim.Fields!PILOT_TRAVEL), 0#, rset_Trim.Fields!PILOT_TRAVEL)

rset_Trim.Close

Exit Sub

Get41000Data_Error:
'Placeholder.  How do we want to handle error messages?
End Sub

Public Sub Get30000Data(long_TrimPackageID As Long, _
    dbl_AllowShaftDPOpen As Double, dbl_AllowShaftDPClosed As Double, _
    dbl_AllowPlugDP As Double, dbl_AllowSupplyMax As Double, _
    dbl_AllowSupplyStability As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves, "Varimax allowable DP calculations" by C. Vlassoff Aug. 23, 1996.
'Solutions reference: 30000.frm, subroutines:    getAllowableShaftDP
'                        getTrimPressureLimit
'
'---------------------------------------------------------------------------------------
'This subroutine is used to retrieve data for 30000 series valves.
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID             package ID of trim
'---------------------------------------------------------------------------------------
'Output Variables:
'dbl_AllowShaftDPopen      allowable pressure with selected shaft       psi
'dbl_AllowShaftDPclosed    allowable pressure with selected shaft       psi
'dbl_AllowPlugDP           allowable pressure with selected trim        psi
'                          material Group
'dbl_AllowSupplyMax        maximum allowable supply                     psig
'dbl_AllowSupplyStability  maximum allowable supply - stability         psig
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/12 Larry Schoonover - not tested
'6/12 -Larry Schoonover -tested - passed
'9/12 Modified to match new database - Larry Schoonover
'-----------------------------------------------------------------------------------------

On Error GoTo Get30000Data_Error

'***************************
'Declarations
'***************************
Dim dbl_NotUsed As Double
Dim str_NotUsed As String
Dim long_NotUsed As Long

'***************************
'Parameter validation
'***************************
If (long_TrimPackageID <= 0) Then GoTo Get30000Data_Error

'****************************
'Begin calculations
'****************************
Call GetTrimData(long_TrimPackageID, long_NotUsed, dbl_NotUsed, _
    dbl_NotUsed, dbl_NotUsed, dbl_NotUsed, _
    dbl_NotUsed, dbl_NotUsed, dbl_NotUsed, dbl_NotUsed, _
    dbl_AllowPlugDP, dbl_NotUsed, dbl_AllowSupplyMax, dbl_NotUsed, dbl_AllowShaftDPOpen, _
    dbl_AllowShaftDPClosed, dbl_AllowSupplyStability, _
    dbl_NotUsed, dbl_NotUsed, dbl_NotUsed, str_NotUsed)

Exit Sub

Get30000Data_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub Get28000Data(long_TrimPackageID As Long, _
    dbl_AllowTrimDP As Double, dbl_PSupplyDP As Double)
'--------------------------------------------------------------------------------------
'Reference :Varipak Specification data CH4501 5/91.
'Solutions reference: 28000.frm, subroutines:    getSpringRange
'---------------------------------------------------------------------------------------
'This subroutine is used to retrieve data for 28000 series valves.
'Due to the small orifice diameters used in this product line there is very little variation in allowable pressure drop limits between trim sizes.  As a result, the trim limit and the ANSI limit may be used in place of actuator calculations.
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID     package ID of trim
'
'Output Variables:
'dbl_AllowTrimDP        allowable pressure drop for selected trim       psig
'dbl_PSupplyDP          supply pressure                                 psig
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/12 Larry Schoonover - not tested
'6/12 -Larry Schoonover -tested - passed
'9/16 Modified to match new database - Larry Schoonover
'-----------------------------------------------------------------------------------------

On Error GoTo Get28000Data_Error

'***************************
'Declarations
'***************************
Dim dbl_NotUsed As Double
Dim str_NotUsed As String
Dim long_NotUsed As Long

'***************************
'Parameter validation
'***************************
If (long_TrimPackageID <= 0) Then GoTo Get28000Data_Error

'****************************
'Begin calculations
'****************************
Call GetTrimData(long_TrimPackageID, long_NotUsed, dbl_NotUsed, _
    dbl_NotUsed, dbl_NotUsed, dbl_NotUsed, _
    dbl_NotUsed, dbl_NotUsed, dbl_NotUsed, dbl_NotUsed, _
    dbl_AllowTrimDP, dbl_NotUsed, dbl_PSupplyDP, dbl_NotUsed, dbl_NotUsed, _
    dbl_NotUsed, dbl_NotUsed, _
    dbl_NotUsed, dbl_NotUsed, dbl_NotUsed, str_NotUsed)

Exit Sub

Get28000Data_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub Get35000Data(long_TrimPackageID As Long, _
    dbl_shaftDiameter As Double, dbl_Eccentricity As Double, _
    dbl_MaxAllowValvePressure As Double, dbl_MaxAllowTrimTorque As Double, _
    dbl_SeatRingDiameter As Double, str_ShaftMaterial As String)
'--------------------------------------------------------------------------------------
'Reference : product release
'Solutions reference:    camact.bas, subroutine:     getCalcData.
'            camact.bas, subroutine:     getSeatRingArea  (note : diameter replaces area)
'
'---------------------------------------------------------------------------------------
'This subroutine is used to retrieve shaft diameter, eccentricity, lever length and max allowable pressure.
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID             package ID of trim
'
'Output Variables:
'dbl_ShaftDiameter              diameter of shaft                       inches
'dbl_Eccentricity               eccentricity of plug                    inches
'                               (distance from center of shaft
'                               to center of seat ring)
'dbl_MaxAllowValvePressure      maximum allowable pressure for          psi
'                               selected valve size
'dbl_MaxAllowTrimTorque         maximum torque limitation of trim       in-lbs
'dbl_SeatRingDiameter           seat ring diameter                      inches
'str_ShaftMaterial
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/12 Larry Schoonover - not tested
'9/16 modified to match new database - Larry Schoonover
'-----------------------------------------------------------------------------------------

On Error GoTo Get35000Data_Error

'***************************
'Declarations
'***************************
Dim dbl_NotUsed As Double
Dim str_NotUsed As String
Dim long_NotUsed As Long

'***************************
'Parameter validation
'***************************
If (long_TrimPackageID <= 0) Then GoTo Get35000Data_Error

'****************************
'Begin calculations
'****************************
Call GetTrimData(long_TrimPackageID, long_NotUsed, dbl_NotUsed, _
    dbl_NotUsed, dbl_NotUsed, dbl_NotUsed, _
    dbl_NotUsed, dbl_shaftDiameter, dbl_Eccentricity, dbl_NotUsed, _
    dbl_MaxAllowValvePressure, dbl_SeatRingDiameter, dbl_NotUsed, _
    dbl_MaxAllowTrimTorque, dbl_NotUsed, dbl_NotUsed, dbl_NotUsed, _
    dbl_NotUsed, dbl_NotUsed, dbl_NotUsed, str_ShaftMaterial)

Exit Sub

Get35000Data_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub GetActuatorData(long_ActuatorPartID As Long, dbl_AreaInitial As Double, _
    dbl_AreaFinal As Double, dbl_MaxSupplyPressure As Double, dbl_ActuatorFriction As Double, _
    dbl_PositionerLossCoef As Double, str_Model As String, long_AirActionID As Long, _
    int_QtySeals As Integer, dbl_StemSealWidth As Double, dbl_PistonRingWidth As Double, _
    dbl_VolOver As Double, dbl_ForceSpringClosed As Double, dbl_ActuatorStemDiameter As Double, _
    dbl_LeverLength As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for linear travel valves.
'Solutions reference :   21000.frm, subroutine: getActArea
'---------------------------------------------------------------------------------------
'This subroutine is used to define actuator data based on selected actuator parameters.
'---------------------------------------------------------------------------------------
'Input Variables:
'long_ActuatorPartID                Part ID of trim package
'
'Output Variables:
'dbl_AreaInitial            initial actuator area                       in^2
'dbl_AreaFinal              final actuator area                         in^2
'dbl_MaxSupplyPressure      maximum allowable supply pressure           psi
'dbl_ActuatorFriction       force necessary to overcome                 lb
'                           actuator friction
'dbl_PositionerLossCoef     coefficient used to describe effect of
'                           positioner pressure loss
'int_QtySeals
'dbl_StemSealWidth
'dbl_PistonRingWidth
'dbl_VolOver
'dbl_ForceSpringClosed
'dbl_ActuatorStemDiameter
'dbl_LeverLength As Double
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/12 Larry Schoonover - not tested
'Added 3 output parameters 11/9/98 Larry Schoonover
'-----------------------------------------------------------------------------------------

On Error GoTo GetActuatorData_Error

'***************************
'Declarations
'***************************
Dim str_DatabasePath As String
Dim str_ActuatorQuery As String
Dim rset_Actuator As Recordset

'***************************
'Parameter validation
'***************************
If (long_ActuatorPartID <= 0) Then GoTo GetActuatorData_Error

'****************************
'Begin calculations
'****************************
str_ActuatorQuery = "SELECT tblActuator.AREA_INITIAL, " _
    & "tblActuator.AREA_FINAL, " _
    & "tblActuator.MAXIMUM_SUPPLY_PRESSURE, " _
    & "tblActuator.ACTUATOR_FRICTION, " _
    & "tblActuator.POSITIONER_LOSS_COEF, " _
    & "tblActuator.MODEL, " _
    & "tblActuator.AIR_ACTION_ID, " _
    & "tblActuator.QTY_SEALS, " _
    & "tblActuator.STEM_SEAL_WIDTH, " _
    & "tblActuator.VOL_OVER, " _
    & "tblActuator.FORCE_SPRING_CLOSED, " _
    & "tblActuator.ACTUATOR_SEAL_DIAMETER, " _
    & "tblActuator.LEVER_LENGTH " _
    & "FROM tblActuator " _
    & "WHERE tblActuator.PART_ID=" & Str$(long_ActuatorPartID)

Set rset_Actuator = g_db_Prod.OpenRecordset(str_ActuatorQuery, dbReadOnly)
If (rset_Actuator.RecordCount <> 1) Then GoTo GetActuatorData_Error

dbl_AreaInitial = IIf(IsNull(rset_Actuator.Fields!AREA_INITIAL), 0#, rset_Actuator.Fields!AREA_INITIAL)
dbl_AreaFinal = IIf(IsNull(rset_Actuator.Fields!AREA_FINAL), 0#, rset_Actuator.Fields!AREA_FINAL)
dbl_MaxSupplyPressure = IIf(IsNull(rset_Actuator.Fields!MAXIMUM_SUPPLY_PRESSURE), 0#, rset_Actuator.Fields!MAXIMUM_SUPPLY_PRESSURE)
dbl_ActuatorFriction = IIf(IsNull(rset_Actuator.Fields!ACTUATOR_FRICTION), 0#, rset_Actuator.Fields!ACTUATOR_FRICTION)
dbl_PositionerLossCoef = IIf(IsNull(rset_Actuator.Fields!POSITIONER_LOSS_COEF), 0#, rset_Actuator.Fields!POSITIONER_LOSS_COEF)
str_Model = IIf(IsNull(rset_Actuator.Fields!MODEL), "", rset_Actuator.Fields!MODEL)
long_AirActionID = IIf(IsNull(rset_Actuator.Fields!AIR_ACTION_ID), "", rset_Actuator.Fields!AIR_ACTION_ID)
int_QtySeals = IIf(IsNull(rset_Actuator.Fields!QTY_SEALS), 0, rset_Actuator.Fields!QTY_SEALS)
dbl_StemSealWidth = IIf(IsNull(rset_Actuator.Fields!STEM_SEAL_WIDTH), 0#, rset_Actuator.Fields!STEM_SEAL_WIDTH)
dbl_PistonRingWidth = 0.25 'default
dbl_VolOver = IIf(IsNull(rset_Actuator.Fields!VOL_OVER), 0#, rset_Actuator.Fields!VOL_OVER)
dbl_ForceSpringClosed = IIf(IsNull(rset_Actuator.Fields!FORCE_SPRING_CLOSED), 0#, rset_Actuator.Fields!FORCE_SPRING_CLOSED)
dbl_ActuatorStemDiameter = IIf(IsNull(rset_Actuator.Fields!ACTUATOR_SEAL_DIAMETER), 0#, rset_Actuator.Fields!ACTUATOR_SEAL_DIAMETER)
dbl_LeverLength = IIf(IsNull(rset_Actuator.Fields!LEVER_LENGTH), 0#, rset_Actuator.Fields!LEVER_LENGTH)

rset_Actuator.Close

#If (DebugVersion = "yes") Then
MsgBox "GetActuatorData" + Chr(13) + Chr(10) + _
    "dbl_AreaInitial=" + Str(dbl_AreaInitial) + Chr(13) + Chr(10) + _
    "dbl_AreaFinal=" + Str(dbl_AreaFinal) + Chr(13) + Chr(10) + _
    "dbl_MaxSupplyPressure=" + Str(dbl_MaxSupplyPressure) + Chr(13) + Chr(10) + _
    "dbl_ActuatorFriction=" + Str(dbl_ActuatorFriction) + Chr(13) + Chr(10) + _
    "dbl_PositionerLossCoef=" + Str(dbl_PositionerLossCoef) + Chr(13) + Chr(10) + _
    "str_Model=" + str_Model + Chr(13) + Chr(10) + _
    "long_AirActionID=" + CStr(long_AirActionID) + Chr(13) + Chr(10) + _
    "dbl_VolOver=" + CStr(dbl_VolOver) + Chr(13) + Chr(10) + _
    "dbl_ForceSpringClosed=" + CStr(dbl_ForceSpringClosed) + Chr(13) + Chr(10) + _
    "dbl_ActuatorStemDiameter=" + CStr(dbl_ActuatorStemDiameter) + Chr(13) + Chr(10) + _
    "dbl_LeverLength=" + CStr(dbl_LeverLength)

#End If

Exit Sub

GetActuatorData_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub GetVarimaxActuatorData(long_ActuatorPartID As Long, dbl_ActuatorOutputTorque As Double, _
    dbl_OffBalOpenTorque As Double, dbl_OffBalClosedTorque As Double, _
    dbl_SeatingTorque As Double, dbl_LeverLength As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves, "Varimax allowable DP calculations" by C. Vlassoff Aug. 23, 1996.
'Solutions reference: 30000.frm, subroutines:    getAllowableShaftDP
'                        getTrimPressureLimit
'                        getOffbalTorque
'---------------------------------------------------------------------------------------
'This subroutine is used to retrieve data for 30000 series valves.
'---------------------------------------------------------------------------------------
'Input Variables:
'long_ActuatorPartID        Part ID of trim package
'
'Output Variables:
'dbl_ActuatorOutputTorque   output torque of selected actuator          daNm/psi
'dbl_OffBalOpenTorque       dynamic torque in open position             daNm/psi
'dbl_OffBalClosedTorque     dynamic torque in closed position           daNm/psi
'dbl_SeatingTorque          seating torque                              daNm/psi
'dbl_LeverLength
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/12 Larry Schoonover - not tested
'9/16 modified to match new database - Larry Schoonover
'-----------------------------------------------------------------------------------------

On Error GoTo GetVarimaxActuatorData_Error

'***************************
'Declarations
'***************************
Dim str_DatabasePath As String
Dim str_ActuatorQuery As String
Dim rset_Actuator As Recordset

'***************************
'Parameter validation
'***************************
If (long_ActuatorPartID <= 0) Then GoTo GetVarimaxActuatorData_Error

'****************************
'Begin calculations
'****************************
str_ActuatorQuery = "SELECT tblActuator.ACTUATOR_OUTPUT_TORQUE, " _
    & "tblActuator.OFF_BALANCE_OPEN_TORQUE , " _
    & "tblActuator.OFF_BALANCE_CLOSED_TORQUE, " _
    & "tblActuator.SEATING_TORQUE, tblActuator.LEVER_LENGTH " _
    & "FROM tblActuator WHERE tblActuator.PART_ID= " _
    & Str$(long_ActuatorPartID)
Set rset_Actuator = g_db_Prod.OpenRecordset(str_ActuatorQuery, dbReadOnly)
If (rset_Actuator.RecordCount <> 1) Then GoTo GetVarimaxActuatorData_Error

dbl_ActuatorOutputTorque = rset_Actuator.Fields(0)
dbl_OffBalOpenTorque = rset_Actuator.Fields(1)
dbl_OffBalClosedTorque = rset_Actuator.Fields(2)
dbl_SeatingTorque = rset_Actuator.Fields(3)
dbl_LeverLength = rset_Actuator.Fields(4)

rset_Actuator.Close
Exit Sub

GetVarimaxActuatorData_Error:
'Placeholder.  How do we want to handle error messages?

End Sub


Public Sub CalcDP_21000_1Stage_FTO_ATO(long_TrimPackageID As Long, _
    long_ActuatorPartID As Long, _
    dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, _
    dbl_InletPressure3 As Double, dbl_InletPressure4 As Double, _
    dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, _
    dbl_InletTemperature1 As Double, dbl_InletTemperature2 As Double, _
    dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, _
    dbl_ShutOffDP As Double, str_PressureUM As String, _
    str_TemperatureUM As String, str_CalibPressureUM As String, str_BodyMaterial As String, _
    str_Rating As String, str_PackingMaterial As String, _
    str_LeakageClass As String, str_SpringRange As String, dbl_Travel As Double, str_LengthUM As String, _
    dbl_ShowDP As Double, dbl_ShowDPSupplyPressure As Double, dbl_ShowMaxDP As Double, _
    dbl_ShowMaxDPSupplyPressure As Double, str_Message As String)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the selected valve and actuator shutoff capability and compare the result to defined requirements.
'This subroutine is specific to 21000 seies valves with single stage trim, flow to open, air to open.
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID
'long_ActuatorPartID
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff pressure drop
'str_PressureUM                 unit of measure associated with inlet and outlet pressures
'str_TemperatureUM              unit of measure associated with inlet temperature
'str_CalibPressureUM            unit of measure associated with spring range and supply pressure
'str_LengthUM                   unit of measure for length (Travel)
'str_BodyMaterial
'str_Rating
'str_PackingMaterial            selected packing material
'str_LeakageClass               selected leakage class
'str_SpringRange                selected actuator spring range
'dbl_Travel                     Actuator travel
'
'Output Variables:
'dbl_ShowDP                     specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure       calculated required supply pressure for specified
'dbl_ShowMaxDP                  calculated maximum pressure drop for selected valve and actuator combination
'dbl_ShowMaxDPsupplyPressure    calculated required supply pressure for maximum pressure drop
'str_Message                    warning message if pressure drop requirement exceeds equipment capacity
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/15 Larry Schoonover - not tested
'
'revised 7/8/98 John Haberlin - evaluation logic revised during testing
'                               test code remains as commented text
'revised 8/25/98 John Haberlin - duplicate stem calculation and evaluation removed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcDP_21000_1Stage_FTO_ATO_Error

#If (DebugVersion = "yes") Then
MsgBox "long_TrimPackageID= " + CStr(long_TrimPackageID) + Chr(13) + Chr(10) + "long_ActuatorPartID= " + CStr(long_ActuatorPartID) + Chr(13) + Chr(10) + "dbl_InletPressure1= " + CStr(dbl_InletPressure1) + Chr(13) + Chr(10) + _
    "dbl_InletPressure2= " + CStr(dbl_InletPressure2) + Chr(13) + Chr(10) + "dbl_InletPressure3= " + CStr(dbl_InletPressure3) + Chr(13) + Chr(10) + "dbl_InletPressure4= " + CStr(dbl_InletPressure4) + Chr(13) + Chr(10) + _
    "dbl_OutletPressure1= " + CStr(dbl_OutletPressure1) + Chr(13) + Chr(10) + "dbl_OutletPressure2= " + CStr(dbl_OutletPressure2) + Chr(13) + Chr(10) + "dbl_OutletPressure3= " + CStr(dbl_OutletPressure3) + Chr(13) + Chr(10) + _
    "dbl_OutletPressure4= " + CStr(dbl_OutletPressure4) + Chr(13) + Chr(10) + "dbl_InletTemperature1= " + CStr(dbl_InletTemperature1) + Chr(13) + Chr(10) + "dbl_InletTemperature2= " + CStr(dbl_InletTemperature2) + Chr(13) + Chr(10) + _
    "dbl_InletTemperature3= " + CStr(dbl_InletTemperature3) + "dbl_InletTemperature4= " + CStr(dbl_InletTemperature4) + Chr(13) + Chr(10) + "dbl_ShutoffDP= " + CStr(dbl_ShutOffDP) + Chr(13) + Chr(10) + "str_PressureUM= " + str_PressureUM + _
    Chr(13) + Chr(10) + "str_TemperatureUM= " + str_TemperatureUM + Chr(13) + Chr(10) + "str_CalibPressureUM= " + str_CalibPressureUM + Chr(13) + Chr(10) + "str_bodyMaterial= " + str_BodyMaterial + Chr(13) + Chr(10) + "str_Rating= " + str_Rating + _
    Chr(13) + Chr(10) + "str_PackingMaterial= " + str_PackingMaterial + Chr(13) + Chr(10) + "str_LeakageClass= " + str_LeakageClass + Chr(13) + Chr(10) + "str_SpringRange= " + str_SpringRange + Chr(13) + Chr(10) + "dbl_Travel= " + _
    CStr(dbl_Travel) + Chr(13) + Chr(10) + "str_LengthUM= " + str_LengthUM + Chr(13) + Chr(10) + "dbl_ShowDP= " + CStr(dbl_ShowDP) + Chr(13) + Chr(10) + "dbl_ShowDPSupplyPressure= " + CStr(dbl_ShowDPSupplyPressure) + Chr(13) + Chr(10) _
    + "dbl_ShowMaxDP= " + CStr(dbl_ShowMaxDP) + Chr(13) + Chr(10) + "dbl_ShowMaxDPSupplyPressure= " + CStr(dbl_ShowMaxDPSupplyPressure) + Chr(13) + Chr(10) + "str_Message= " + str_Message
#End If

'***************************
'Declarations
'***************************
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_AllowANSIPressure As Double
Dim dbl_DynamicFactor As Double
Dim dbl_PlugDiameter As Double
Dim dbl_SeatDiameter As Double
Dim dbl_StemDiameter As Double
Dim dbl_PackingHeight As Double
Dim dbl_AllowPlugDP As Double
Dim dbl_ForcePacking As Double
Dim dbl_ForceSeating As Double
Dim dbl_ForceUnbalancedClosed As Double
Dim dbl_ReqForceClosed As Double
Dim dbl_AreaInitial As Double
Dim dbl_AreaFinal As Double
Dim dbl_MaxSupplyPressure As Double
Dim dbl_ActuatorFriction As Double
Dim dbl_PositionerLossCoef As Double
Dim dbl_PInitial As Double
Dim dbl_PFinal As Double
Dim dbl_SpringRate As Double
Dim dbl_RequiredSupplyPressure As Double
Dim dbl_ActuatorForce As Double
Dim dbl_AllowDPClosed As Double
Dim dbl_StemStressArea As Double
Dim dbl_Sy As Double
Dim dbl_E As Double
Dim dbl_Sd As Double
Dim dbl_L As Double
Dim dbl_AllowableCompressiveForce As Double
Dim dbl_AllowableTensileForce As Double
Dim dbl_AllowableStemLoad As Double
Dim dbl_AllowDPStability  As Double
Dim dbl_AllowDPOpen As Double
Dim dbl_MaxAllowDP As Double
Dim dbl_PSupplyDP As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim str_StemMaterial As String

Dim dbl_InletPressure(4) As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double

Dim int_QtySeals As Integer
Dim dbl_StemSealWidth As Double
Dim dbl_PistonRingWidth As Double
Dim dbl_VolOver As Double
Dim dbl_ForceSpringClosed As Double
Dim dbl_ActuatorStemDiameter As Double
Dim dbl_LeverLength As Double

Dim str_NotUsed As String
Dim long_NotUsed As Long


'***************************
'Parameter validation
'***************************
'default for now
If (str_PackingMaterial = "") Then str_PackingMaterial = "PTFE"

If (str_PressureUM = "") Then GoTo CalcDP_21000_1Stage_FTO_ATO_Error
If (str_TemperatureUM = "") Then GoTo CalcDP_21000_1Stage_FTO_ATO_Error
If (str_CalibPressureUM = "") Then GoTo CalcDP_21000_1Stage_FTO_ATO_Error
If (str_PackingMaterial = "") Then GoTo CalcDP_21000_1Stage_FTO_ATO_Error
If (str_LeakageClass = "") Then GoTo CalcDP_21000_1Stage_FTO_ATO_Error
If (str_SpringRange = "") Then GoTo CalcDP_21000_1Stage_FTO_ATO_Error

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = IIf(dbl_InletTemperature1 = 0, -99999, dbl_InletTemperature1)
dbl_InletTemperature(2) = IIf(dbl_InletTemperature2 = 0, -99999, dbl_InletTemperature2)
dbl_InletTemperature(3) = IIf(dbl_InletTemperature3 = 0, -99999, dbl_InletTemperature3)
dbl_InletTemperature(4) = IIf(dbl_InletTemperature4 = 0, -99999, dbl_InletTemperature4)

'determine requirements
Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature(), dbl_MaxInletTemperature)
dbl_ShowDP = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, str_PressureUM)
dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, str_TemperatureUM)
dbl_ShowShutoffDP = dbl_ShutOffDP
If (dbl_ShutOffDP > 0) Then dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, str_PressureUM)

'determine ANSI limitation


    Call GetANSIPressure(str_BodyMaterial, dbl_MaxInletTemperature, str_Rating, _
    dbl_AllowANSIPressure)
    
    Call Get21000Data(long_TrimPackageID, dbl_DynamicFactor, _
        dbl_PlugDiameter, dbl_SeatDiameter, dbl_StemDiameter, dbl_PackingHeight, _
        dbl_AllowPlugDP, dbl_StemStressArea, dbl_Travel, str_StemMaterial)
    CnvrtLengthToMM dbl_Travel, "in", dbl_Travel
    CnvrtLengthFromMM dbl_Travel, str_LengthUM, dbl_Travel

    
'calculate required packing force
Call CalcPackingForce(dbl_MaxInletPressure, dbl_StemDiameter, str_PackingMaterial, _
    dbl_PackingHeight, dbl_ForcePacking)

'calculate required seating force
Call CalcSeatingForce(str_LeakageClass, dbl_SeatDiameter, dbl_ForceSeating)


'calculate required offbalance force
Call CalcSingleSeatUnbalancedForceFTO(dbl_MaxInletPressure, dbl_SeatDiameter, _
    dbl_ForceUnbalancedClosed)
dbl_ReqForceClosed = dbl_ForcePacking + dbl_ForceSeating + dbl_ForceUnbalancedClosed

'calculate actuator capability
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'If debug_on = False Then
    Call GetActuatorData(long_ActuatorPartID, dbl_AreaInitial, _
        dbl_AreaFinal, dbl_MaxSupplyPressure, dbl_ActuatorFriction, _
        dbl_PositionerLossCoef, str_NotUsed, long_NotUsed, _
        int_QtySeals, dbl_StemSealWidth, dbl_PistonRingWidth, _
        dbl_VolOver, dbl_ForceSpringClosed, dbl_ActuatorStemDiameter, dbl_LeverLength)


Call CalcSpringRate(str_SpringRange, dbl_AreaInitial, dbl_PInitial, _
    dbl_AreaFinal, dbl_PFinal, str_CalibPressureUM, dbl_Travel, str_LengthUM, _
    dbl_SpringRate)
    
Call CalcATOActuatorForce(dbl_AreaInitial, dbl_PInitial, dbl_PFinal, _
    dbl_ActuatorFriction, dbl_RequiredSupplyPressure, dbl_ActuatorForce)
    
'calculate allowable pressure drop
Call CalcAllowDPSingleStageUnbalFTO(dbl_ActuatorForce, dbl_ForceSeating, _
    dbl_ActuatorFriction, dbl_SeatDiameter, dbl_AllowDPClosed)

'calculate stem limitation
'Call GetStemStressArea(dbl_StemDiameter, dbl_StemStressArea) now returned by get21000data

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'If debug_on = False Then
    Call GetStemProperties(str_StemMaterial, dbl_MaxInletTemperature, dbl_Sy, _
        dbl_E, dbl_Sd)

'Else
'    Call screenStemData(dbl_E, dbl_Sd, dbl_Sy)
'End If
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

dbl_L = 6# 'default value

'***********************revision 8/25/98********************************************

Call CalcAllowStemForce(dbl_Sy, dbl_E, dbl_StemDiameter, dbl_L, _
    dbl_Sd, dbl_StemStressArea, dbl_AllowableStemLoad)
'Call CalcAllowStemTensileForce(dbl_Sd, dbl_stemStressArea, dbl_AllowableTensileForce)
'
'If (dbl_AllowableTensileForce < dbl_AllowableCompressiveForce) Then
'    dbl_AllowableStemLoad = dbl_AllowableTensileForce
'Else
'    dbl_AllowableStemLoad = dbl_AllowableCompressiveForce
'End If
'***********************************************************************************
'calculate actuator pressure drop capability
dbl_AllowDPStability = 0#
dbl_AllowDPOpen = 0#
Call CalcMaxDPAllowable(dbl_AllowDPStability, dbl_AllowDPClosed, dbl_AllowDPOpen, _
    dbl_AllowANSIPressure, dbl_MaxAllowDP)
dbl_PSupplyDP = dbl_PFinal + 5#



'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'If debug_on Then
'
'    form21000.txtAllowStemLoad.Text = dbl_AllowableStemLoad
'    form21000.txtAllowANSIpressure = dbl_AllowANSIPressure
'    form21000.txtPackingForce.Text = Format(dbl_ForcePacking, "#####")
'    form21000.txtSeatingForce = Format(dbl_ForceSeating, "#####")
'
'    form21000.txtUnbalForceClosed.Text = Format(dbl_ForceUnbalancedClosed, "#####")
'    form21000.txtUnbalForceOpen.Text = Format(dbl_ForceUnbalancedOpen, "#####")
'    form21000.txtReqForceClosed.Text = Format(dbl_ReqForceClosed, "#####")
'    form21000.txtReqForceOpen.Text = Format(dbl_ReqForceOpen, "#####")
'
'    form21000.txtActuatorForce.Text = dbl_ActuatorForce
'    form21000.txtPsupply.Text = dbl_PSupply
'
'End If
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

#If (DebugVersion = "yes") Then
MsgBox "Other Data" + Chr(13) + Chr(10) + _
    "dbl_ForcePacking" + Str(dbl_ForcePacking) + Chr(13) + Chr(10) + _
    "dbl_ForceSeating" + Str(dbl_ForceSeating) + Chr(13) + Chr(10) + _
    "dbl_ForceUnbalancedClosed" + Str(dbl_ForceUnbalancedClosed) + Chr(13) + Chr(10) + _
    "dbl_SpringRate" + Str(dbl_SpringRate) + Chr(13) + Chr(10) + _
    "dbl_ActuatorForce" + Str(dbl_ActuatorForce) + Chr(13) + Chr(10) + _
    "dbl_AllowDPClosed" + Str(dbl_AllowDPClosed) + Chr(13) + Chr(10) + _
    "dbl_Sy" + Str(dbl_Sy) + Chr(13) + Chr(10) + _
    "dbl_E" + Str(dbl_E) + Chr(13) + Chr(10) + _
    "dbl_Sd" + Str(dbl_Sd) + Chr(13) + Chr(10) + _
    "dbl_AllowableStemLoad" + Str(dbl_AllowableStemLoad) + Chr(13) + Chr(10) + _
    "dbl_MaxAllowDP" + Str(dbl_MaxAllowDP) + Chr(13) + Chr(10) + _
    "dbl_MaxInletPressure" + Str(dbl_MaxInletPressure) + Chr(13) + Chr(10) + _
    "dbl_AllowANSIPressure" + Str(dbl_AllowANSIPressure) + Chr(13) + Chr(10) + _
    "dbl_AllowPlugDP" + Str(dbl_AllowPlugDP) + Chr(13) + Chr(10) + _
    "dbl_ShutoffDP" + Str(dbl_ShutOffDP) + Chr(13) + Chr(10)
#End If

str_Message = ""
dbl_ShowDP = 0
dbl_ShowMaxDP = 0
dbl_ShowMaxDPSupplyPressure = 0

If (dbl_MaxInletPressure > dbl_AllowANSIPressure) Then
    str_Message = ANSIMESSAGE
     dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowANSIPressure, str_PressureUM)
    GoTo EndOfRoutine
End If

If dbl_AllowPlugDP > 0 And (dbl_MaxInletPressure > dbl_AllowPlugDP) Then
    str_Message = TRIMMESSAGE
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowPlugDP, str_PressureUM)
    GoTo EndOfRoutine
End If

If (dbl_ActuatorForce > dbl_AllowableStemLoad) Then
    str_Message = STEMMESSAGE
    dbl_ShowDPSupplyPressure = 0
    dbl_ShowMaxDPSupplyPressure = 0
    GoTo EndOfRoutine
End If

If (dbl_MaxInletPressure <= dbl_MaxAllowDP) And str_Message = "" Then
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_CalibPressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, str_PressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_CalibPressureUM)
ElseIf (dbl_MaxInletPressure > dbl_MaxAllowDP And dbl_ShutOffDP <= dbl_MaxAllowDP And dbl_ShutOffDP > 0#) Then
    str_Message = LIMITEDDPMESSAGE
    dbl_ShowDP = dbl_ShowShutoffDP
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_CalibPressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, str_PressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_CalibPressureUM)
Else
    str_Message = ACTUATORMESSAGE
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, str_PressureUM)
End If

EndOfRoutine:
dbl_ShowDP = TrimDecimalPlaces(dbl_ShowDP)
dbl_ShowDPSupplyPressure = TrimDecimalPlaces(dbl_ShowDPSupplyPressure)
dbl_ShowMaxDP = TrimDecimalPlaces(dbl_ShowMaxDP)
dbl_ShowMaxDPSupplyPressure = TrimDecimalPlaces(dbl_ShowMaxDPSupplyPressure)

#If (DebugVersion = "yes") Then
MsgBox "long_TrimPackageID= " + CStr(long_TrimPackageID) + Chr(13) + Chr(10) + "long_ActuatorPartID= " + CStr(long_ActuatorPartID) + Chr(13) + Chr(10) + "dbl_InletPressure1= " + CStr(dbl_InletPressure1) + Chr(13) + Chr(10) + _
"dbl_InletPressure2= " + CStr(dbl_InletPressure2) + Chr(13) + Chr(10) + "dbl_InletPressure3= " + CStr(dbl_InletPressure3) + Chr(13) + Chr(10) + "dbl_InletPressure4= " + CStr(dbl_InletPressure4) + Chr(13) + Chr(10) + _
"dbl_OutletPressure1= " + CStr(dbl_OutletPressure1) + Chr(13) + Chr(10) + "dbl_OutletPressure2= " + CStr(dbl_OutletPressure2) + Chr(13) + Chr(10) + "dbl_OutletPressure3= " + CStr(dbl_OutletPressure3) + Chr(13) + Chr(10) + _
"dbl_OutletPressure4= " + CStr(dbl_OutletPressure4) + Chr(13) + Chr(10) + "dbl_InletTemperature1= " + CStr(dbl_InletTemperature1) + Chr(13) + Chr(10) + "dbl_InletTemperature2= " + CStr(dbl_InletTemperature2) + Chr(13) + Chr(10) + _
"dbl_InletTemperature3= " + CStr(dbl_InletTemperature3) + "dbl_InletTemperature4= " + CStr(dbl_InletTemperature4) + Chr(13) + Chr(10) + "dbl_ShutoffDP= " + CStr(dbl_ShutOffDP) + Chr(13) + Chr(10) + "str_PressureUM= " + str_PressureUM + _
Chr(13) + Chr(10) + "str_TemperatureUM= " + str_TemperatureUM + Chr(13) + Chr(10) + "str_CalibPressureUM= " + str_CalibPressureUM + Chr(13) + Chr(10) + "str_BodyMaterial= " + str_BodyMaterial + Chr(13) + Chr(10) + "str_Rating= " + str_Rating + _
Chr(13) + Chr(10) + "str_PackingMaterial= " + str_PackingMaterial + Chr(13) + Chr(10) + "str_LeakageClass= " + str_LeakageClass + Chr(13) + Chr(10) + "str_SpringRange= " + str_SpringRange + Chr(13) + Chr(10) + "dbl_Travel= " + _
CStr(dbl_Travel) + Chr(13) + Chr(10) + "str_LengthUM= " + str_LengthUM + Chr(13) + Chr(10) + "dbl_ShowDP= " + CStr(dbl_ShowDP) + Chr(13) + Chr(10) + "dbl_ShowDPSupplyPressure= " + CStr(dbl_ShowDPSupplyPressure) + Chr(13) + Chr(10) _
+ "dbl_ShowMaxDP= " + CStr(dbl_ShowMaxDP) + Chr(13) + Chr(10) + "dbl_ShowMaxDPSupplyPressure= " + CStr(dbl_ShowMaxDPSupplyPressure) + Chr(13) + Chr(10) + "str_Message= " + str_Message
#End If

Exit Sub

CalcDP_21000_1Stage_FTO_ATO_Error:
'Placeholder.  How do we want to handle error messages?

End Sub


Public Sub CalcDP_21000_1Stage_FTO_ATC(long_TrimPackageID As Long, _
    long_ActuatorPartID As Long, _
    dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, _
    dbl_InletPressure3 As Double, dbl_InletPressure4 As Double, _
    dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, _
    dbl_InletTemperature1 As Double, dbl_InletTemperature2 As Double, _
    dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, _
    dbl_ShutOffDP As Double, str_PressureUM As String, _
    str_TemperatureUM As String, str_CalibPressureUM As String, str_BodyMaterial As String, _
    str_Rating As String, str_PackingMaterial As String, _
    str_LeakageClass As String, str_SpringRange As String, dbl_Travel As Double, _
    str_LengthUM As String, dbl_MaxAllowSupplyPressure As Double, _
    dbl_ShowDP As Double, dbl_ShowDPSupplyPressure As Double, dbl_ShowMaxDP As Double, _
    dbl_ShowMaxDPSupplyPressure As Double, str_Message As String)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the selected valve and actuator shutoff capability and compare the result to defined requirements.
'This subroutine is specific to 21000 seies valves with single stage trim, flow to open, air to close.
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID
'long_ActuatorPartID
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff pressure drop
'str_PressureUM                 unit of measure associated with inlet and outlet pressures
'str_TemperatureUM              unit of measure associated with inlet temperature
'str_CalibPressureUM            unit of measure associated with spring range and supply pressure
'str_LengthUM                   unit of measure for length (Travel)
'dbl_MaxAllowSupplyPressure     maximum available supply pressure defined by user
'str_BodyMaterial
'str_Rating
'str_PackingMaterial            selected packing material
'str_LeakageClass               selected leakage class
'str_SpringRange                selected actuator spring range
'dbl_Travel                     Actuator travel
'
'Output Variables:
'dbl_ShowDP                     specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure       calculated required supply pressure for specified
'dbl_ShowMaxDP                  calculated maximum pressure drop for selected valve and actuator combination
'dbl_ShowMaxDPsupplyPressure    calculated required supply pressure for maximum pressure drop
'str_Message                    warning message if pressure drop requirement exceeds equipment capability
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/15 Larry Schoonover - not tested
'
'revised 7/8/98 John Haberlin - evaluation logic revised during testing
'                               test code remains as commented text
'revised 8/25/98 John Haberlin - duplicate stem calculation and evaluation removed
'revised 9/23/98 John Haberlin - added "dbl_MaxAllowSupplyPressure" as required input parameter
'revised 10/23/98 John Haberlin - revised evaluation of dbl_ShowMaxDP
'-----------------------------------------------------------------------------------------

On Error GoTo CalcDP_21000_1Stage_FTO_ATC_Error

#If (DebugVersion = "yes") Then
MsgBox "long_TrimPackageID= " + CStr(long_TrimPackageID) + Chr(13) + Chr(10) + "long_ActuatorPartID= " + CStr(long_ActuatorPartID) + Chr(13) + Chr(10) + "dbl_InletPressure1= " + CStr(dbl_InletPressure1) + Chr(13) + Chr(10) + _
"dbl_InletPressure2= " + CStr(dbl_InletPressure2) + Chr(13) + Chr(10) + "dbl_InletPressure3= " + CStr(dbl_InletPressure3) + Chr(13) + Chr(10) + "dbl_InletPressure4= " + CStr(dbl_InletPressure4) + Chr(13) + Chr(10) + _
"dbl_OutletPressure1= " + CStr(dbl_OutletPressure1) + Chr(13) + Chr(10) + "dbl_OutletPressure2= " + CStr(dbl_OutletPressure2) + Chr(13) + Chr(10) + "dbl_OutletPressure3= " + CStr(dbl_OutletPressure3) + Chr(13) + Chr(10) + _
"dbl_OutletPressure4= " + CStr(dbl_OutletPressure4) + Chr(13) + Chr(10) + "dbl_InletTemperature1= " + CStr(dbl_InletTemperature1) + Chr(13) + Chr(10) + "dbl_InletTemperature2= " + CStr(dbl_InletTemperature2) + Chr(13) + Chr(10) + _
"dbl_InletTemperature3= " + CStr(dbl_InletTemperature3) + "dbl_InletTemperature4= " + CStr(dbl_InletTemperature4) + Chr(13) + Chr(10) + "dbl_ShutoffDP= " + CStr(dbl_ShutOffDP) + Chr(13) + Chr(10) + "str_PressureUM= " + str_PressureUM + _
Chr(13) + Chr(10) + "str_TemperatureUM= " + str_TemperatureUM + Chr(13) + Chr(10) + "str_CalibPressureUM= " + str_CalibPressureUM + Chr(13) + Chr(10) + "str_BodyMaterial= " + str_BodyMaterial + Chr(13) + Chr(10) + "str_Rating= " + str_Rating + _
Chr(13) + Chr(10) + "str_PackingMaterial= " + str_PackingMaterial + Chr(13) + Chr(10) + "str_LeakageClass= " + str_LeakageClass + Chr(13) + Chr(10) + "str_SpringRange= " + str_SpringRange + Chr(13) + Chr(10) + "dbl_Travel= " + _
CStr(dbl_Travel) + Chr(13) + Chr(10) + "str_LengthUM= " + str_LengthUM + Chr(13) + Chr(10) + "dbl_ShowDP= " + CStr(dbl_ShowDP) + Chr(13) + Chr(10) + "dbl_ShowDPSupplyPressure= " + CStr(dbl_ShowDPSupplyPressure) + Chr(13) + Chr(10) _
+ "dbl_ShowMaxDP= " + CStr(dbl_ShowMaxDP) + Chr(13) + Chr(10) + "dbl_ShowMaxDPSupplyPressure= " + CStr(dbl_ShowMaxDPSupplyPressure) + Chr(13) + Chr(10) + "str_Message= " + str_Message + "dbl_MaxAllowSupplyPressure= " + CStr(dbl_MaxAllowSupplyPressure)
#End If

'***************************
'Declarations
'***************************
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_AllowANSIPressure As Double
Dim dbl_DynamicFactor As Double
Dim dbl_PlugDiameter As Double
Dim dbl_SeatDiameter As Double
Dim dbl_StemDiameter As Double
Dim dbl_PackingHeight As Double
Dim dbl_AllowPlugDP As Double
Dim dbl_ForcePacking As Double
Dim dbl_ForceSeating As Double
Dim dbl_ForceUnbalancedClosed As Double
Dim dbl_ReqForceClosed As Double
Dim dbl_AreaInitial As Double
Dim dbl_AreaFinal As Double
Dim dbl_MaxSupplyPressure As Double
Dim dbl_ActuatorFriction As Double
Dim dbl_PositionerLossCoef As Double
Dim dbl_PInitial As Double
Dim dbl_PFinal As Double
Dim dbl_SpringRate As Double
Dim dbl_RequiredSupplyPressure As Double
Dim dbl_ActuatorForce As Double
Dim dbl_AllowDPClosed As Double
Dim dbl_StemStressArea As Double
Dim dbl_Sy As Double
Dim dbl_E As Double
Dim dbl_Sd As Double
Dim dbl_L As Double
Dim dbl_AllowableCompressiveForce As Double
Dim dbl_AllowableTensileForce As Double
Dim dbl_AllowableStemLoad As Double
Dim dbl_AllowDPStability  As Double
Dim dbl_AllowDPOpen As Double
Dim dbl_MaxAllowDP As Double
Dim dbl_PSupplyDP As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim dbl_ReqDPForceClosed As Double
Dim dbl_ReqANSIForceClosed As Double
Dim dbl_ReqPlugLimitForceClosed As Double
Dim dbl_ReqShutoffForceClosed As Double
Dim dbl_PSupplyANSI As Double
Dim dbl_PSupplyStemAllow As Double
Dim dbl_PSupplyPlugAllow As Double
Dim dbl_PSupplyShutoffDP As Double
Dim dbl_MaxAllowSupply As Double
Dim bool_ANSIOK As Boolean, bool_StemOK As Boolean, bool_ActuatorOpenOK As Boolean
Dim bool_ActuatorClosedOK As Boolean, bool_ActuatorStabilityOK As Boolean
Dim bool_TrimOK As Boolean, bool_SupplyOK As Boolean
Dim str_StemMaterial As String
Dim dbl_PSupply As Double
Dim dbl_AllowDP As Double
Dim dbl_MaxSupplyDP As Double
Dim dbl_AllowableStemDP As Double

Dim dbl_InletPressure(4) As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double
Dim int_QtySeals As Integer
Dim dbl_StemSealWidth As Double
Dim dbl_PistonRingWidth As Double
Dim dbl_VolOver As Double
Dim dbl_ForceSpringClosed As Double
Dim dbl_ActuatorStemDiameter As Double
Dim dbl_LeverLength As Double

Dim str_NotUsed As String
Dim long_NotUsed As Long


'***************************
'Parameter validation
'***************************
'default for now
If (str_PackingMaterial = "") Then str_PackingMaterial = "PTFE"

If (str_PressureUM = "") Then GoTo CalcDP_21000_1Stage_FTO_ATC_Error
If (str_TemperatureUM = "") Then GoTo CalcDP_21000_1Stage_FTO_ATC_Error
If (str_CalibPressureUM = "") Then GoTo CalcDP_21000_1Stage_FTO_ATC_Error
If (str_PackingMaterial = "") Then GoTo CalcDP_21000_1Stage_FTO_ATC_Error
If (str_LeakageClass = "") Then GoTo CalcDP_21000_1Stage_FTO_ATC_Error
If (str_SpringRange = "") Then GoTo CalcDP_21000_1Stage_FTO_ATC_Error

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4

'determine requirements
Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature(), dbl_MaxInletTemperature)
dbl_ShowDP = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, str_PressureUM)
dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, str_TemperatureUM)
dbl_ShowShutoffDP = dbl_ShutOffDP
If (dbl_ShutOffDP > 0) Then dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, str_PressureUM)

'ANSI pressure
'-----------------------------------------------------------------------------------------
'If debug_on = False Then
    Call GetANSIPressure(str_BodyMaterial, dbl_MaxInletTemperature, str_Rating, _
    dbl_AllowANSIPressure)
'Else
'    dbl_AllowANSIPressure = 1500
'End If
'-----------------------------------------------------------------------------------------
'read valve properties
'-----------------------------------------------------------------------------------------
'If debug_on = False Then
    Call Get21000Data(long_TrimPackageID, dbl_DynamicFactor, _
        dbl_PlugDiameter, dbl_SeatDiameter, dbl_StemDiameter, dbl_PackingHeight, _
        dbl_AllowPlugDP, dbl_StemStressArea, dbl_Travel, str_StemMaterial)
    CnvrtLengthToMM dbl_Travel, "in", dbl_Travel
    CnvrtLengthFromMM dbl_Travel, str_LengthUM, dbl_Travel

    
'calculate required packing force
Call CalcPackingForce(dbl_MaxInletPressure, dbl_StemDiameter, str_PackingMaterial, _
    dbl_PackingHeight, dbl_ForcePacking)
    
'calculate required seating force
Call CalcSeatingForce(str_LeakageClass, dbl_SeatDiameter, dbl_ForceSeating)

'calculate required offbalance force
Call CalcSingleSeatUnbalancedForceFTO(dbl_MaxInletPressure, dbl_SeatDiameter, _
    dbl_ForceUnbalancedClosed)
dbl_ReqDPForceClosed = dbl_ForcePacking + dbl_ForceSeating + dbl_ForceUnbalancedClosed
Call CalcSingleSeatUnbalancedForceFTO(dbl_AllowANSIPressure, dbl_SeatDiameter, _
    dbl_ForceUnbalancedClosed)
dbl_ReqANSIForceClosed = dbl_ForcePacking + dbl_ForceSeating + dbl_ForceUnbalancedClosed
Call CalcSingleSeatUnbalancedForceFTO(dbl_AllowPlugDP, dbl_SeatDiameter, _
    dbl_ForceUnbalancedClosed)
dbl_ReqPlugLimitForceClosed = dbl_ForcePacking + dbl_ForceSeating + dbl_ForceUnbalancedClosed
Call CalcSingleSeatUnbalancedForceFTO(dbl_ShutOffDP, dbl_SeatDiameter, _
    dbl_ForceUnbalancedClosed)
dbl_ReqShutoffForceClosed = dbl_ForcePacking + dbl_ForceSeating + dbl_ForceUnbalancedClosed

'calculate stem limitation
'Call GetStemStressArea(dbl_StemDiameter, dbl_StemStressArea)now returned by get21000data

'-----------------------------------------------------------------------------------------
'If debug_on = False Then
    Call GetStemProperties(str_StemMaterial, dbl_MaxInletTemperature, dbl_Sy, _
        dbl_E, dbl_Sd)
'Else
'    Call screenStemData(dbl_E, dbl_Sd, dbl_Sy)
'End If
'-----------------------------------------------------------------------------------------


dbl_L = 6# 'default value

'----------revision jh 8/25/98--------------------------------------------------------

Call CalcAllowStemForce(dbl_Sy, dbl_E, dbl_StemDiameter, dbl_L, _
    dbl_Sd, dbl_StemStressArea, dbl_AllowableStemLoad)
'Call CalcAllowStemTensileForce(dbl_Sd, dbl_stemStressArea, dbl_AllowableTensileForce)
'
'If (dbl_AllowableTensileForce < dbl_AllowableCompressiveForce) Then
'    dbl_AllowableStemLoad = dbl_AllowableTensileForce
'Else
'    dbl_AllowableStemLoad = dbl_AllowableCompressiveForce
'End If
'-----------------------------------------------------------------------------------------
'calculate actuator capability
'-----------------------------------------------------------------------------------------
'If debug_on = False Then
    Call GetActuatorData(long_ActuatorPartID, dbl_AreaInitial, _
        dbl_AreaFinal, dbl_MaxSupplyPressure, dbl_ActuatorFriction, _
        dbl_PositionerLossCoef, str_NotUsed, long_NotUsed, _
        int_QtySeals, dbl_StemSealWidth, dbl_PistonRingWidth, _
        dbl_VolOver, dbl_ForceSpringClosed, dbl_ActuatorStemDiameter, dbl_LeverLength)
'Else
'    Call screenActuatorData(str_actuatorModel, str_actuatorSize, dbl_Travel, _
'        dbl_AreaInitial, dbl_AreaFinal, dbl_MaxAllowSupplyPressure, _
'        dbl_ActuatorFriction, dbl_PositionerLossCoef, str_SpringRange, str_LengthUM)
'End If

'-----------------------------------------------------------------------------------------
Call CalcSpringRate(str_SpringRange, dbl_AreaInitial, dbl_PInitial, dbl_AreaFinal, _
            dbl_PFinal, str_CalibPressureUM, dbl_Travel, _
            str_LengthUM, dbl_SpringRate)

Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, dbl_ReqANSIForceClosed, _
    dbl_PSupply)
dbl_PSupplyANSI = Int(dbl_PSupply)

Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, dbl_AllowableStemLoad, _
    dbl_PSupply)
dbl_PSupplyStemAllow = Int(dbl_PSupply)

Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, dbl_ReqPlugLimitForceClosed, _
    dbl_PSupply)
dbl_PSupplyPlugAllow = Int(dbl_PSupply)

Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, dbl_ReqShutoffForceClosed, _
    dbl_PSupply)
dbl_PSupplyShutoffDP = Int(dbl_PSupply)
    
Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, dbl_ReqDPForceClosed, _
    dbl_PSupply)
dbl_PSupplyDP = Int(dbl_PSupply)

'limit the supply pressure

dbl_MaxAllowSupply = dbl_MaxAllowSupplyPressure
If (dbl_PSupplyANSI < dbl_MaxAllowSupply) Then dbl_MaxAllowSupply = dbl_PSupplyANSI
If (dbl_PSupplyStemAllow < dbl_MaxAllowSupply) Then dbl_MaxAllowSupply = dbl_PSupplyStemAllow
If (dbl_PSupplyPlugAllow < dbl_MaxAllowSupply) Then dbl_MaxAllowSupply = dbl_PSupplyPlugAllow

'calculate max DP with max allow supply
Call CalcATCActuatorForce(dbl_AreaFinal, dbl_PFinal, dbl_MaxAllowSupply, dbl_ActuatorFriction, _
    dbl_ActuatorForce)
    
Call CalcAllowDPSingleStageUnbalFTO(dbl_ActuatorForce, dbl_ForceSeating, dbl_ActuatorFriction, _
    dbl_SeatDiameter, dbl_AllowDP)

'------------added 10/26/98-----------------------
dbl_ShowMaxDP = dbl_AllowPlugDP
If (dbl_AllowANSIPressure < dbl_ShowMaxDP) Then dbl_ShowMaxDP = dbl_AllowANSIPressure
If (dbl_AllowDP < dbl_ShowMaxDP) Then dbl_ShowMaxDP = dbl_AllowDP
dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_ShowMaxDP, str_PressureUM)
'-----------------------------------------------------------------------------------------

'calculate allowable pressure drop for stem
Call CalcAllowDPSingleStageUnbalFTO(dbl_AllowableStemLoad, dbl_ForceSeating, _
    dbl_ActuatorFriction, dbl_SeatDiameter, dbl_AllowDP)
dbl_AllowableStemDP = dbl_AllowDP

'evaluation
bool_ANSIOK = True
bool_StemOK = True
bool_ActuatorOpenOK = True
bool_ActuatorClosedOK = True
bool_ActuatorStabilityOK = True
bool_TrimOK = True
bool_SupplyOK = True

'-----------------------------------------------------------------------------------------
'If debug_on Then
'
'    form21000.txtAllowStemLoad.Text = dbl_AllowableStemLoad
'    form21000.txtAllowANSIpressure = dbl_AllowANSIPressure
'    form21000.txtPackingForce.Text = Format(dbl_ForcePacking, "#####")
'    form21000.txtSeatingForce = Format(dbl_ForceSeating, "#####")
'
'    form21000.txtUnbalForceClosed.Text = Format(dbl_ForceUnbalancedClosed, "#####")
'    form21000.txtUnbalForceOpen.Text = Format(dbl_ForceUnbalancedOpen, "#####")
'    form21000.txtReqForceClosed.Text = Format(dbl_ReqForceClosed, "#####")
'    form21000.txtReqForceOpen.Text = Format(dbl_ReqForceOpen, "#####")
'
'    form21000.txtActuatorForce.Text = dbl_ActuatorForce
'    form21000.txtPsupply.Text = dbl_PSupply
'
'End If
'-----------------------------------------------------------------------------------------

 str_Message = ""
If (dbl_MaxInletPressure > dbl_AllowANSIPressure) Then
    bool_ANSIOK = False
    str_Message = ANSIMESSAGE
    dbl_ShowDP = ConvertPressureFromPSIG(dbl_MaxInletPressure, str_PressureUM)
    '-----inactivated 10/26/98-----------------------
    'dbl_ShowMaxDP = dbl_AllowANSIPressure
    '------------------------------------------------
    dbl_ShowMaxDPSupplyPressure = dbl_MaxAllowSupply
    GoTo EndOfRoutine
End If

If (dbl_MaxInletPressure > dbl_AllowPlugDP) Then
    bool_TrimOK = False
    str_Message = TRIMMESSAGE
    dbl_ShowDP = ConvertPressureFromPSIG(dbl_MaxInletPressure, str_PressureUM)
    '-----inactivated 10/26/98-----------------------
    'dbl_ShowMaxDP = dbl_AllowPlugDP
    '------------------------------------------------
    dbl_ShowMaxDPSupplyPressure = dbl_MaxAllowSupply
    GoTo EndOfRoutine
End If

If (dbl_MaxInletPressure > dbl_AllowableStemDP) Then
    bool_StemOK = False
    str_Message = SUPPLYMESSAGE
    dbl_ShowDP = ConvertPressureFromPSIG(dbl_MaxInletPressure, str_PressureUM)
    '-----inactivated 10/26/98-----------------------
    'dbl_ShowMaxDP = dbl_AllowableStemDP
    '-----------------------------------------------    dbl_ShowMaxDPSupplyPressure = dbl_MaxAllowSupply
    GoTo EndOfRoutine
End If

If (dbl_PSupplyDP > dbl_MaxAllowSupply And dbl_PSupplyShutoffDP <= dbl_MaxAllowSupply) Then
    bool_SupplyOK = False
    str_Message = LIMITEDDPMESSAGE
ElseIf (dbl_PSupplyDP > dbl_MaxAllowSupply And dbl_PSupplyShutoffDP > dbl_MaxAllowSupply) Then
    bool_SupplyOK = False
    str_Message = SUPPLYMESSAGE
End If




If (bool_ANSIOK = True And bool_TrimOK = True And bool_StemOK = True And bool_SupplyOK = True) Then
    dbl_ShowDP = ConvertPressureFromPSIG(dbl_MaxInletPressure, str_PressureUM)
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_PressureUM)
End If

'-----inactivated 10/26/98-----------------------
'dbl_ShowMaxDP = dbl_AllowDP

'dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_ShowMaxDP, str_PressureUM)
'dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_MaxAllowSupply, _
    str_CalibPressureUM)
'-----------------------------------------------------------------------------------------

EndOfRoutine:
dbl_ShowDP = TrimDecimalPlaces(dbl_ShowDP)
dbl_ShowDPSupplyPressure = TrimDecimalPlaces(dbl_ShowDPSupplyPressure)
dbl_ShowMaxDP = TrimDecimalPlaces(dbl_ShowMaxDP)
dbl_ShowMaxDPSupplyPressure = TrimDecimalPlaces(dbl_ShowMaxDPSupplyPressure)

#If (DebugVersion = "yes") Then
MsgBox "long_TrimPackageID= " + CStr(long_TrimPackageID) + Chr(13) + Chr(10) + "long_ActuatorPartID= " + CStr(long_ActuatorPartID) + Chr(13) + Chr(10) + "dbl_InletPressure1= " + CStr(dbl_InletPressure1) + Chr(13) + Chr(10) + _
"dbl_InletPressure2= " + CStr(dbl_InletPressure2) + Chr(13) + Chr(10) + "dbl_InletPressure3= " + CStr(dbl_InletPressure3) + Chr(13) + Chr(10) + "dbl_InletPressure4= " + CStr(dbl_InletPressure4) + Chr(13) + Chr(10) + _
"dbl_OutletPressure1= " + CStr(dbl_OutletPressure1) + Chr(13) + Chr(10) + "dbl_OutletPressure2= " + CStr(dbl_OutletPressure2) + Chr(13) + Chr(10) + "dbl_OutletPressure3= " + CStr(dbl_OutletPressure3) + Chr(13) + Chr(10) + _
"dbl_OutletPressure4= " + CStr(dbl_OutletPressure4) + Chr(13) + Chr(10) + "dbl_InletTemperature1= " + CStr(dbl_InletTemperature1) + Chr(13) + Chr(10) + "dbl_InletTemperature2= " + CStr(dbl_InletTemperature2) + Chr(13) + Chr(10) + _
"dbl_InletTemperature3= " + CStr(dbl_InletTemperature3) + "dbl_InletTemperature4= " + CStr(dbl_InletTemperature4) + Chr(13) + Chr(10) + "dbl_ShutoffDP= " + CStr(dbl_ShutOffDP) + Chr(13) + Chr(10) + "str_PressureUM= " + str_PressureUM + _
Chr(13) + Chr(10) + "str_TemperatureUM= " + str_TemperatureUM + Chr(13) + Chr(10) + "str_CalibPressureUM= " + str_CalibPressureUM + Chr(13) + Chr(10) + "str_BodyMaterial= " + str_BodyMaterial + Chr(13) + Chr(10) + "str_Rating= " + str_Rating + _
Chr(13) + Chr(10) + "str_PackingMaterial= " + str_PackingMaterial + Chr(13) + Chr(10) + "str_LeakageClass= " + str_LeakageClass + Chr(13) + Chr(10) + "str_SpringRange= " + str_SpringRange + Chr(13) + Chr(10) + "dbl_Travel= " + _
CStr(dbl_Travel) + Chr(13) + Chr(10) + "str_LengthUM= " + str_LengthUM + Chr(13) + Chr(10) + "dbl_ShowDP= " + CStr(dbl_ShowDP) + Chr(13) + Chr(10) + "dbl_ShowDPSupplyPressure= " + CStr(dbl_ShowDPSupplyPressure) + Chr(13) + Chr(10) _
+ "dbl_ShowMaxDP= " + CStr(dbl_ShowMaxDP) + Chr(13) + Chr(10) + "dbl_ShowMaxDPSupplyPressure= " + CStr(dbl_ShowMaxDPSupplyPressure) + Chr(13) + Chr(10) + "str_Message= " + str_Message + "dbl_MaxAllowSupplyPressure= " + CStr(dbl_MaxAllowSupplyPressure)
#End If

Exit Sub

CalcDP_21000_1Stage_FTO_ATC_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcDP_21000_1Stage_FTC_ATO(long_TrimPackageID As Long, _
    long_ActuatorPartID As Long, _
    dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, _
    dbl_InletPressure3 As Double, dbl_InletPressure4 As Double, _
    dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, _
    dbl_InletTemperature1 As Double, dbl_InletTemperature2 As Double, _
    dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, _
    dbl_ShutOffDP As Double, str_PressureUM As String, _
    str_TemperatureUM As String, str_CalibPressureUM As String, str_BodyMaterial As String, _
    str_Rating As String, str_PackingMaterial As String, _
    str_LeakageClass As String, str_SpringRange As String, dbl_Travel As Double, str_LengthUM As String, _
    dbl_ShowDP As Double, dbl_ShowDPSupplyPressure As Double, dbl_ShowMaxDP As Double, _
    dbl_ShowMaxDPSupplyPressure As Double, str_Message As String)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the selected valve and actuator shutoff capability and compare the result to defined requirements.
'This subroutine is specific to 21000 seies valves with single stage trim, flow to close, air to open.
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID
'long_ActuatorPartID
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff pressure drop
'str_PressureUM                 unit of measure associated with inlet and outlet pressures
'str_TemperatureUM              unit of measure associated with inlet temperature
'str_CalibPressureUM            unit of measure associated with spring range and supply pressure
'str_LengthUM                   unit of measure for length (Travel)
'str_BodyMaterial
'str_Rating
'str_PackingMaterial            selected packing material
'str_LeakageClass               selected leakage class
'str_SpringRange                selected actuator spring range
'dbl_Travel                     Actuator travel
'
'Output Variables:
'dbl_ShowDP                     specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure       calculated required supply pressure for specified
'dbl_ShowMaxDP                  calculated maximum pressure drop for selected valve and actuator combination
'dbl_ShowMaxDPsupplyPressure    calculated required supply pressure for maximum pressure drop
'str_Message                    warning message if pressure drop requirement exceeds equipment capability
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/15 Larry Schoonover - not tested
'
'revised 7/8/98 John Haberlin - evaluation logic revised during testing
'                               test code remains as commented text
'revised 8/25/98 John Haberlin - duplicate stem calculation and evaluation removed
'-----------------------------------------------------------------------------------------
On Error GoTo CalcDP_21000_1Stage_FTC_ATO_Error

#If (DebugVersion = "yes") Then
MsgBox "long_TrimPackageID= " + CStr(long_TrimPackageID) + Chr(13) + Chr(10) + "long_ActuatorPartID= " + CStr(long_ActuatorPartID) + Chr(13) + Chr(10) + "dbl_InletPressure1= " + CStr(dbl_InletPressure1) + Chr(13) + Chr(10) + _
"dbl_InletPressure2= " + CStr(dbl_InletPressure2) + Chr(13) + Chr(10) + "dbl_InletPressure3= " + CStr(dbl_InletPressure3) + Chr(13) + Chr(10) + "dbl_InletPressure4= " + CStr(dbl_InletPressure4) + Chr(13) + Chr(10) + _
"dbl_OutletPressure1= " + CStr(dbl_OutletPressure1) + Chr(13) + Chr(10) + "dbl_OutletPressure2= " + CStr(dbl_OutletPressure2) + Chr(13) + Chr(10) + "dbl_OutletPressure3= " + CStr(dbl_OutletPressure3) + Chr(13) + Chr(10) + _
"dbl_OutletPressure4= " + CStr(dbl_OutletPressure4) + Chr(13) + Chr(10) + "dbl_InletTemperature1= " + CStr(dbl_InletTemperature1) + Chr(13) + Chr(10) + "dbl_InletTemperature2= " + CStr(dbl_InletTemperature2) + Chr(13) + Chr(10) + _
"dbl_InletTemperature3= " + CStr(dbl_InletTemperature3) + "dbl_InletTemperature4= " + CStr(dbl_InletTemperature4) + Chr(13) + Chr(10) + "dbl_ShutoffDP= " + CStr(dbl_ShutOffDP) + Chr(13) + Chr(10) + "str_PressureUM= " + str_PressureUM + _
Chr(13) + Chr(10) + "str_TemperatureUM= " + str_TemperatureUM + Chr(13) + Chr(10) + "str_CalibPressureUM= " + str_CalibPressureUM + Chr(13) + Chr(10) + "str_BodyMaterial= " + str_BodyMaterial + Chr(13) + Chr(10) + "str_Rating= " + str_Rating + _
Chr(13) + Chr(10) + "str_PackingMaterial= " + str_PackingMaterial + Chr(13) + Chr(10) + "str_LeakageClass= " + str_LeakageClass + Chr(13) + Chr(10) + "str_SpringRange= " + str_SpringRange + Chr(13) + Chr(10) + "dbl_Travel= " + _
CStr(dbl_Travel) + Chr(13) + Chr(10) + "str_LengthUM= " + str_LengthUM + Chr(13) + Chr(10) + "dbl_ShowDP= " + CStr(dbl_ShowDP) + Chr(13) + Chr(10) + "dbl_ShowDPSupplyPressure= " + CStr(dbl_ShowDPSupplyPressure) + Chr(13) + Chr(10) _
+ "dbl_ShowMaxDP= " + CStr(dbl_ShowMaxDP) + Chr(13) + Chr(10) + "dbl_ShowMaxDPSupplyPressure= " + CStr(dbl_ShowMaxDPSupplyPressure) + Chr(13) + Chr(10) + "str_Message= " + str_Message
#End If

'***************************
'Declarations
'***************************
Dim dbl_MaxPressureDrop As Double
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_AllowANSIPressure As Double
Dim dbl_DynamicFactor As Double
Dim dbl_PlugDiameter As Double
Dim dbl_SeatDiameter As Double
Dim dbl_StemDiameter As Double
Dim dbl_PackingHeight As Double
Dim dbl_AllowPlugDP As Double
Dim dbl_ForcePacking As Double
Dim dbl_ForceSeating As Double
Dim dbl_ForceUnbalancedClosed As Double
Dim dbl_ForceUnbalancedOpen As Double
Dim dbl_ReqForceClosed As Double
Dim dbl_ReqForceOpen As Double
Dim dbl_AreaInitial As Double
Dim dbl_AreaFinal As Double
Dim dbl_MaxSupplyPressure As Double
Dim dbl_ActuatorFriction As Double
Dim dbl_PositionerLossCoef As Double
Dim dbl_PInitial As Double
Dim dbl_PFinal As Double
Dim dbl_SpringRate As Double
Dim dbl_RequiredSupplyPressure As Double
Dim dbl_ActuatorForce As Double
Dim dbl_AllowDPClosed As Double
Dim dbl_StemStressArea As Double
Dim dbl_Sy As Double
Dim dbl_E As Double
Dim dbl_Sd As Double
Dim dbl_L As Double
Dim dbl_AllowableCompressiveForce As Double
Dim dbl_AllowableTensileForce As Double
Dim dbl_AllowableStemLoad As Double
Dim dbl_AllowDPStability  As Double
Dim dbl_AllowDPOpen As Double
Dim dbl_MaxAllowDP As Double
Dim dbl_PSupplyDP As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim dbl_ReqANSIForceClosed As Double
Dim dbl_ReqPlugLimitForceClosed As Double
Dim dbl_ReqShutoffForceClosed As Double
Dim dbl_PSupplyANSI As Double
Dim dbl_PSupplyStemAllow As Double
Dim dbl_PSupplyPlugAllow As Double
Dim dbl_PSupplyShutoffDP As Double
Dim dbl_MaxAllowSupply As Double
Dim dbl_EffectiveOrificeArea As Double
Dim str_StemMaterial As String
Dim dbl_PSupply As Double

Dim dbl_InletPressure(4) As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double
Dim int_QtySeals As Integer
Dim dbl_StemSealWidth As Double
Dim dbl_PistonRingWidth As Double
Dim dbl_VolOver As Double
Dim dbl_ForceSpringClosed As Double
Dim dbl_ActuatorStemDiameter As Double
Dim dbl_LeverLength As Double

Dim str_NotUsed As String
Dim long_NotUsed As Long


'***************************
'Parameter validation
'***************************
'default for now
If (str_PackingMaterial = "") Then str_PackingMaterial = "PTFE"

If (str_PressureUM = "") Then GoTo CalcDP_21000_1Stage_FTC_ATO_Error
If (str_TemperatureUM = "") Then GoTo CalcDP_21000_1Stage_FTC_ATO_Error
If (str_CalibPressureUM = "") Then GoTo CalcDP_21000_1Stage_FTC_ATO_Error
If (str_PackingMaterial = "") Then GoTo CalcDP_21000_1Stage_FTC_ATO_Error
If (str_LeakageClass = "") Then GoTo CalcDP_21000_1Stage_FTC_ATO_Error
If (str_SpringRange = "") Then GoTo CalcDP_21000_1Stage_FTC_ATO_Error

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4

'determine requirements
Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature, dbl_MaxInletTemperature)
Call CalcTestPressureDrop(dbl_InletPressure(), dbl_OutletPressure(), dbl_MaxPressureDrop)

dbl_ShowDP = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, str_PressureUM)
dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, str_TemperatureUM)
dbl_MaxPressureDrop = ConvertPressureToPSIG(dbl_MaxPressureDrop, str_PressureUM)
dbl_ShowShutoffDP = dbl_ShutOffDP
If (dbl_ShutOffDP > 0) Then dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, str_PressureUM)

'determine ANSI limitations
'-----------------------------------------------------------------------------------------
'If debug_on = False Then
    Call GetANSIPressure(str_BodyMaterial, dbl_MaxInletTemperature, str_Rating, _
    dbl_AllowANSIPressure)
'Else
'    dbl_AllowANSIPressure = 1500
'End If
'-----------------------------------------------------------------------------------------
'read valve properties
'-----------------------------------------------------------------------------------------
'If debug_on = False Then
    Call Get21000Data(long_TrimPackageID, dbl_DynamicFactor, _
        dbl_PlugDiameter, dbl_SeatDiameter, dbl_StemDiameter, dbl_PackingHeight, _
        dbl_AllowPlugDP, dbl_StemStressArea, dbl_Travel, str_StemMaterial)
    CnvrtLengthToMM dbl_Travel, "in", dbl_Travel
    CnvrtLengthFromMM dbl_Travel, str_LengthUM, dbl_Travel

    
'calculate required packing force
Call CalcPackingForce(dbl_MaxInletPressure, dbl_StemDiameter, str_PackingMaterial, _
    dbl_PackingHeight, dbl_ForcePacking)

'calculate required seating force
Call CalcSeatingForce(str_LeakageClass, dbl_SeatDiameter, dbl_ForceSeating)

'calculate required offbalance force
dbl_EffectiveOrificeArea = 0.2 'default
Call CalcSingleSeatUnbalancedForceFTC(dbl_MaxInletPressure, dbl_SeatDiameter, _
    dbl_StemDiameter, dbl_EffectiveOrificeArea, dbl_ForceUnbalancedClosed, _
    dbl_ForceUnbalancedOpen)

dbl_ReqForceClosed = dbl_ForcePacking + dbl_ForceSeating + dbl_ForceUnbalancedClosed
dbl_ReqForceOpen = dbl_ForcePacking + dbl_ForceUnbalancedOpen

    Call GetActuatorData(long_ActuatorPartID, dbl_AreaInitial, _
        dbl_AreaFinal, dbl_MaxSupplyPressure, dbl_ActuatorFriction, _
        dbl_PositionerLossCoef, str_NotUsed, long_NotUsed, _
        int_QtySeals, dbl_StemSealWidth, dbl_PistonRingWidth, _
        dbl_VolOver, dbl_ForceSpringClosed, dbl_ActuatorStemDiameter, dbl_LeverLength)

Call CalcSpringRate(str_SpringRange, dbl_AreaInitial, dbl_PInitial, dbl_AreaFinal, _
    dbl_PFinal, str_CalibPressureUM, dbl_Travel, str_LengthUM, dbl_SpringRate)

Call CalcATOActuatorForce(dbl_AreaInitial, dbl_PInitial, dbl_PFinal, _
    dbl_ActuatorFriction, dbl_RequiredSupplyPressure, dbl_ActuatorForce)

'calculate allowable pressure drop
Call CalcAllowDPSingleStageUnbalFTC_ATO(dbl_AreaInitial, dbl_PInitial, dbl_AreaFinal, _
    dbl_PFinal, dbl_SpringRate, dbl_RequiredSupplyPressure, dbl_ForcePacking, dbl_SeatDiameter, _
    dbl_StemDiameter, dbl_EffectiveOrificeArea, dbl_DynamicFactor, dbl_AllowDPStability, _
    dbl_AllowDPClosed, dbl_AllowDPOpen)

'calculate stem limitation
'Call GetStemStressArea(dbl_StemDiameter, dbl_StemStressArea)now returned by get21000data
'-----------------------------------------------------------------------------------------
'If debug_on = False Then
    Call GetStemProperties(str_StemMaterial, dbl_MaxInletTemperature, dbl_Sy, _
        dbl_E, dbl_Sd)
'Else
'    Call screenStemData(dbl_E, dbl_Sd, dbl_Sy)
'End If
'-----------------------------------------------------------------------------------------

dbl_L = 6# 'default value

'----------revision jh 8/25/98------------------------------------------------------------

Call CalcAllowStemForce(dbl_Sy, dbl_E, dbl_StemDiameter, dbl_L, _
    dbl_Sd, dbl_StemStressArea, dbl_AllowableStemLoad)

'Call CalcAllowStemTensileForce(dbl_Sd, dbl_stemStressArea, dbl_AllowableTensileForce)

'If (dbl_AllowableTensileForce < dbl_AllowableCompressiveForce) Then
'    dbl_AllowableStemLoad = dbl_AllowableTensileForce
'Else
 '   dbl_AllowableStemLoad = dbl_AllowableCompressiveForce
'End If
'-----------------------------------------------------------------------------------------
'calculate actuator pressure drop capability
Call CalcMaxDPAllowable(dbl_AllowDPStability, dbl_AllowDPClosed, dbl_AllowDPOpen, _
    dbl_AllowANSIPressure, dbl_MaxAllowDP)

'calculate required supply pressure
Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, dbl_ReqForceOpen, _
    dbl_PSupplyDP)
dbl_PSupply = dbl_RequiredSupplyPressure
'-----------------------------------------------------------------------------------------
'If debug_on Then
'
'    form21000.txtAllowStemLoad.Text = dbl_AllowableStemLoad
'    form21000.txtAllowANSIpressure = dbl_AllowANSIPressure
'    form21000.txtPackingForce.Text = Format(dbl_ForcePacking, "#####")
'    form21000.txtSeatingForce = Format(dbl_ForceSeating, "#####")
'
'    form21000.txtUnbalForceClosed.Text = Format(dbl_ForceUnbalancedClosed, "#####")
'    form21000.txtUnbalForceOpen.Text = Format(dbl_ForceUnbalancedOpen, "#####")
'    form21000.txtReqForceClosed.Text = Format(dbl_ReqForceClosed, "#####")
'    form21000.txtReqForceOpen.Text = Format(dbl_ReqForceOpen, "#####")
'
'    form21000.txtActuatorForce.Text = dbl_ActuatorForce
'    form21000.txtPsupply.Text = dbl_PSupply
'
'End If
'-----------------------------------------------------------------------------------------

'evaluation

str_Message = ""

dbl_ShowDP = 0
dbl_ShowDPSupplyPressure = 0
dbl_ShowMaxDP = 0
dbl_ShowMaxDPSupplyPressure = 0

If (dbl_MaxInletPressure > dbl_AllowANSIPressure) Then
    str_Message = ANSIMESSAGE
    dbl_ShowDP = ConvertPressureFromPSIG(dbl_MaxInletPressure, str_PressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowANSIPressure, str_PressureUM)
    GoTo EndOfRoutine
End If

If (dbl_MaxInletPressure > dbl_AllowPlugDP And dbl_AllowPlugDP > 0) Then
    str_Message = TRIMMESSAGE
    dbl_ShowDP = ConvertPressureFromPSIG(dbl_MaxInletPressure, str_PressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowPlugDP, str_PressureUM)
    GoTo EndOfRoutine
End If

If (dbl_ActuatorForce > dbl_AllowableStemLoad) Then
    str_Message = STEMMESSAGE
    GoTo EndOfRoutine
End If
    
'------------------------inactivated 10/26/98-----------------------------

'If (dbl_MaxInletPressure > dbl_AllowDPClosed Or dbl_MaxPressureDrop > dbl_AllowDPOpen _
    Or dbl_MaxPressureDrop > dbl_AllowDPStability) Then str_Message = ACTUATORMESSAGE

'If (dbl_ActuatorForce < (dbl_ReqForceClosed - dbl_ForceUnbalancedOpen)) Then _
    str_Message = ACTUATORMESSAGE
'-------------------------------------------------------------------------


dbl_ShowDP = dbl_MaxInletPressure
dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupply, str_CalibPressureUM)
dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, str_PressureUM)
dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupply, str_CalibPressureUM)

If (dbl_MaxInletPressure > dbl_AllowDPClosed) Then
    str_Message = ACTUATORMESSAGE + " - CLOSED"
    dbl_ShowDPSupplyPressure = 0
End If

If (dbl_MaxPressureDrop > dbl_AllowDPOpen) Then
    str_Message = ACTUATORMESSAGE + " - OPEN" + Chr(13) + "PRESSURE DROP EXCEEDS ALLOWABLE LIMIT"
    dbl_ShowDPSupplyPressure = 0
End If

If (dbl_MaxPressureDrop > dbl_AllowDPStability) Then
    str_Message = ACTUATORMESSAGE + " - STABILITY"
    dbl_ShowDPSupplyPressure = 0
End If

If (dbl_MaxInletPressure > dbl_ShutOffDP And dbl_ShutOffDP <= dbl_AllowDPClosed And dbl_ShutOffDP > 0# _
        And dbl_MaxPressureDrop < dbl_AllowDPOpen And dbl_MaxInletPressure > dbl_AllowDPClosed) Then
    str_Message = LIMITEDDPMESSAGE
    dbl_ShowDP = dbl_ShowShutoffDP
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupply, str_CalibPressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, str_PressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupply, str_CalibPressureUM)

End If

EndOfRoutine:
dbl_ShowDP = TrimDecimalPlaces(dbl_ShowDP)
dbl_ShowDPSupplyPressure = TrimDecimalPlaces(dbl_ShowDPSupplyPressure)
dbl_ShowMaxDP = TrimDecimalPlaces(dbl_ShowMaxDP)
dbl_ShowMaxDPSupplyPressure = TrimDecimalPlaces(dbl_ShowMaxDPSupplyPressure)

#If (DebugVersion = "yes") Then
MsgBox "long_TrimPackageID= " + CStr(long_TrimPackageID) + Chr(13) + Chr(10) + "long_ActuatorPartID= " + CStr(long_ActuatorPartID) + Chr(13) + Chr(10) + "dbl_InletPressure1= " + CStr(dbl_InletPressure1) + Chr(13) + Chr(10) + _
"dbl_InletPressure2= " + CStr(dbl_InletPressure2) + Chr(13) + Chr(10) + "dbl_InletPressure3= " + CStr(dbl_InletPressure3) + Chr(13) + Chr(10) + "dbl_InletPressure4= " + CStr(dbl_InletPressure4) + Chr(13) + Chr(10) + _
"dbl_OutletPressure1= " + CStr(dbl_OutletPressure1) + Chr(13) + Chr(10) + "dbl_OutletPressure2= " + CStr(dbl_OutletPressure2) + Chr(13) + Chr(10) + "dbl_OutletPressure3= " + CStr(dbl_OutletPressure3) + Chr(13) + Chr(10) + _
"dbl_OutletPressure4= " + CStr(dbl_OutletPressure4) + Chr(13) + Chr(10) + "dbl_InletTemperature1= " + CStr(dbl_InletTemperature1) + Chr(13) + Chr(10) + "dbl_InletTemperature2= " + CStr(dbl_InletTemperature2) + Chr(13) + Chr(10) + _
"dbl_InletTemperature3= " + CStr(dbl_InletTemperature3) + "dbl_InletTemperature4= " + CStr(dbl_InletTemperature4) + Chr(13) + Chr(10) + "dbl_ShutoffDP= " + CStr(dbl_ShutOffDP) + Chr(13) + Chr(10) + "str_PressureUM= " + str_PressureUM + _
Chr(13) + Chr(10) + "str_TemperatureUM= " + str_TemperatureUM + Chr(13) + Chr(10) + "str_CalibPressureUM= " + str_CalibPressureUM + Chr(13) + Chr(10) + "str_BodyMaterial= " + str_BodyMaterial + Chr(13) + Chr(10) + "str_Rating= " + str_Rating + _
Chr(13) + Chr(10) + "str_PackingMaterial= " + str_PackingMaterial + Chr(13) + Chr(10) + "str_LeakageClass= " + str_LeakageClass + Chr(13) + Chr(10) + "str_SpringRange= " + str_SpringRange + Chr(13) + Chr(10) + "dbl_Travel= " + _
CStr(dbl_Travel) + Chr(13) + Chr(10) + "str_LengthUM= " + str_LengthUM + Chr(13) + Chr(10) + "dbl_ShowDP= " + CStr(dbl_ShowDP) + Chr(13) + Chr(10) + "dbl_ShowDPSupplyPressure= " + CStr(dbl_ShowDPSupplyPressure) + Chr(13) + Chr(10) _
+ "dbl_ShowMaxDP= " + CStr(dbl_ShowMaxDP) + Chr(13) + Chr(10) + "dbl_ShowMaxDPSupplyPressure= " + CStr(dbl_ShowMaxDPSupplyPressure) + Chr(13) + Chr(10) + "str_Message= " + str_Message
#End If

Exit Sub

CalcDP_21000_1Stage_FTC_ATO_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcDP_21000_1Stage_FTC_ATC(long_TrimPackageID As Long, _
    long_ActuatorPartID As Long, _
    dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, _
    dbl_InletPressure3 As Double, dbl_InletPressure4 As Double, _
    dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, _
    dbl_InletTemperature1 As Double, dbl_InletTemperature2 As Double, _
    dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, _
    dbl_ShutOffDP As Double, str_PressureUM As String, _
    str_TemperatureUM As String, str_CalibPressureUM As String, str_BodyMaterial As String, _
    str_Rating As String, str_PackingMaterial As String, _
    str_LeakageClass As String, str_SpringRange As String, dbl_Travel As Double, str_LengthUM As String, _
    dbl_ShowDP As Double, dbl_ShowDPSupplyPressure As Double, dbl_ShowMaxDP As Double, _
    dbl_ShowMaxDPSupplyPressure As Double, str_Message As String)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the selected valve and actuator shutoff capability
'and compare the result to defined requirements.
'This subroutine is specific to 21000 seies valves with single stage trim,
'flow to close, air to close.
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID
'long_ActuatorPartID
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff pressure drop
'str_PressureUM                 unit of measure associated with inlet and outlet pressures
'str_TemperatureUM              unit of measure associated with inlet temperature
'str_CalibPressureUM            unit of measure associated with spring range and supply pressure
'str_LengthUM                   unit of measure for length (Travel)
'str_BodyMaterial
'str_Rating
'str_PackingMaterial            selected packing material
'str_LeakageClass               selected leakage class
'str_SpringRange                selected actuator spring range
'dbl_Travel                     Actuator travel
'
'Output Variables:
'dbl_ShowDP                     specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure       calculated required supply pressure for specified
'dbl_ShowMaxDP                  calculated maximum pressure drop for selected valve and actuator combination
'dbl_ShowMaxDPsupplyPressure    calculated required supply pressure for maximum pressure drop
'str_Message                    warning message if pressure drop requirement exceeds equipment capability
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/15 Larry Schoonover - not tested
'
'revised 7/8/98 John Haberlin - evaluation logic revised during testing
'                               test code remains as commented code
'revised 8/25/98 John Haberlin - duplicate stem calculation and evaluation removed
'revised 10/26/98 John Haberlin - initialize dbl_ShowDPSupplyPressure in ANSI test
'                               - added test for dbl_MaxInletPressure>dbl_allowDPOpen
'-----------------------------------------------------------------------------------------
On Error GoTo CalcDP_21000_1Stage_FTC_ATC_Error

#If (DebugVersion = "yes") Then
MsgBox "long_TrimPackageID= " + CStr(long_TrimPackageID) + Chr(13) + Chr(10) + "long_ActuatorPartID= " + CStr(long_ActuatorPartID) + Chr(13) + Chr(10) + "dbl_InletPressure1= " + CStr(dbl_InletPressure1) + Chr(13) + Chr(10) + _
"dbl_InletPressure2= " + CStr(dbl_InletPressure2) + Chr(13) + Chr(10) + "dbl_InletPressure3= " + CStr(dbl_InletPressure3) + Chr(13) + Chr(10) + "dbl_InletPressure4= " + CStr(dbl_InletPressure4) + Chr(13) + Chr(10) + _
"dbl_OutletPressure1= " + CStr(dbl_OutletPressure1) + Chr(13) + Chr(10) + "dbl_OutletPressure2= " + CStr(dbl_OutletPressure2) + Chr(13) + Chr(10) + "dbl_OutletPressure3= " + CStr(dbl_OutletPressure3) + Chr(13) + Chr(10) + _
"dbl_OutletPressure4= " + CStr(dbl_OutletPressure4) + Chr(13) + Chr(10) + "dbl_InletTemperature1= " + CStr(dbl_InletTemperature1) + Chr(13) + Chr(10) + "dbl_InletTemperature2= " + CStr(dbl_InletTemperature2) + Chr(13) + Chr(10) + _
"dbl_InletTemperature3= " + CStr(dbl_InletTemperature3) + Chr(13) + Chr(10) + "dbl_InletTemperature4= " + CStr(dbl_InletTemperature4) + Chr(13) + Chr(10) + "dbl_ShutoffDP= " + CStr(dbl_ShutOffDP) + Chr(13) + Chr(10) + "str_PressureUM= " + str_PressureUM + _
Chr(13) + Chr(10) + "str_TemperatureUM= " + str_TemperatureUM + Chr(13) + Chr(10) + "str_CalibPressureUM= " + str_CalibPressureUM + Chr(13) + Chr(10) + "str_BodyMaterial= " + str_BodyMaterial + Chr(13) + Chr(10) + "str_Rating= " + str_Rating + _
Chr(13) + Chr(10) + "str_PackingMaterial= " + str_PackingMaterial + Chr(13) + Chr(10) + "str_LeakageClass= " + str_LeakageClass + Chr(13) + Chr(10) + "str_SpringRange= " + str_SpringRange + Chr(13) + Chr(10) + "dbl_Travel= " + _
CStr(dbl_Travel) + Chr(13) + Chr(10) + "str_LengthUM= " + str_LengthUM + Chr(13) + Chr(10) + "dbl_ShowDP= " + CStr(dbl_ShowDP) + Chr(13) + Chr(10) + "dbl_ShowDPSupplyPressure= " + CStr(dbl_ShowDPSupplyPressure) + Chr(13) + Chr(10) _
+ "dbl_ShowMaxDP= " + CStr(dbl_ShowMaxDP) + Chr(13) + Chr(10) + "dbl_ShowMaxDPSupplyPressure= " + CStr(dbl_ShowMaxDPSupplyPressure) + Chr(13) + Chr(10) + "str_Message= " + str_Message
#End If


'***************************
'Declarations
'***************************
Dim dbl_MaxPressureDrop As Double
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_AllowANSIPressure As Double
Dim dbl_DynamicFactor As Double
Dim dbl_PlugDiameter As Double
Dim dbl_SeatDiameter As Double
Dim dbl_StemDiameter As Double
Dim dbl_PackingHeight As Double
Dim dbl_AllowPlugDP As Double
Dim dbl_ForcePacking As Double
Dim dbl_ForceSeating As Double
Dim dbl_ForceUnbalancedClosed As Double
Dim dbl_ForceUnbalancedOpen As Double
Dim dbl_ReqForceClosed As Double
Dim dbl_ReqForceOpen As Double
Dim dbl_AreaInitial As Double
Dim dbl_AreaFinal As Double
Dim dbl_MaxSupplyPressure As Double
Dim dbl_ActuatorFriction As Double
Dim dbl_PositionerLossCoef As Double
Dim dbl_PInitial As Double
Dim dbl_PFinal As Double
Dim dbl_SpringRate As Double
Dim dbl_RequiredSupplyPressure As Double
Dim dbl_ActuatorForce As Double
Dim dbl_AllowDPClosed As Double
Dim dbl_StemStressArea As Double
Dim dbl_Sy As Double
Dim dbl_E As Double
Dim dbl_Sd As Double
Dim dbl_L As Double
Dim dbl_AllowableCompressiveForce As Double
Dim dbl_AllowableTensileForce As Double
Dim dbl_AllowableStemLoad As Double
Dim dbl_AllowDPStability  As Double
Dim dbl_AllowDPOpen As Double
Dim dbl_MaxAllowDP As Double
Dim dbl_PSupplyDP As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim dbl_ReqANSIForceClosed As Double
Dim dbl_ReqPlugLimitForceClosed As Double
Dim dbl_ReqShutoffForceClosed As Double
Dim dbl_PSupplyANSI As Double
Dim dbl_PSupplyStemAllow As Double
Dim dbl_PSupplyPlugAllow As Double
Dim dbl_PSupplyShutoffDP As Double
Dim dbl_MaxAllowSupply As Double
Dim dbl_EffectiveOrificeArea As Double
Dim str_StemMaterial As String
Dim str_Model As String
Dim long_AirActionID As Long

Dim dbl_InletPressure(4) As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double
Dim int_QtySeals As Integer
Dim dbl_StemSealWidth As Double
Dim dbl_PistonRingWidth As Double
Dim dbl_VolOver As Double
Dim dbl_ForceSpringClosed As Double
Dim dbl_ActuatorStemDiameter As Double
Dim dbl_LeverLength As Double
'-----------------------------------------------------------------------------------------
'If debug_on = True Then
'    Dim dbl_PSupply As Double
'    Dim dbl_packingCoefficient As Double
'    Dim str_ActuatorModel As String
'    Dim str_actuatorSize As String
'
'    Dim dbl_Kv As Double
'    Dim dbl_Fcs As Double
'    Dim dbl_Au As Double
'End If
'-----------------------------------------------------------------------------------------

'***************************
'Parameter validation
'***************************
'default for now
If (str_PackingMaterial = "") Then str_PackingMaterial = "PTFE"

If (str_PressureUM = "") Then GoTo CalcDP_21000_1Stage_FTC_ATC_Error
If (str_TemperatureUM = "") Then GoTo CalcDP_21000_1Stage_FTC_ATC_Error
If (str_CalibPressureUM = "") Then GoTo CalcDP_21000_1Stage_FTC_ATC_Error
If (str_PackingMaterial = "") Then GoTo CalcDP_21000_1Stage_FTC_ATC_Error
If (str_LeakageClass = "") Then GoTo CalcDP_21000_1Stage_FTC_ATC_Error
If (str_SpringRange = "") Then GoTo CalcDP_21000_1Stage_FTC_ATC_Error

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4

'determine requirements
Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature(), dbl_MaxInletTemperature)
Call CalcTestPressureDrop(dbl_InletPressure(), dbl_OutletPressure(), dbl_MaxPressureDrop)

dbl_ShowDP = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, str_PressureUM)
dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, str_TemperatureUM)
dbl_MaxPressureDrop = ConvertPressureToPSIG(dbl_MaxPressureDrop, str_PressureUM)

dbl_ShowShutoffDP = dbl_ShutOffDP
If (dbl_ShutOffDP > 0) Then dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, str_PressureUM)

'determine ANSI limitations
'-----------------------------------------------------------------------------------------
'If debug_on = False Then
    Call GetANSIPressure(str_BodyMaterial, dbl_MaxInletTemperature, str_Rating, _
    dbl_AllowANSIPressure)
'Else
'    dbl_AllowANSIPressure = 1500
'End If
'-----------------------------------------------------------------------------------------

'read valve properties
'-----------------------------------------------------------------------------------------
'If debug_on = False Then
    Call Get21000Data(long_TrimPackageID, dbl_DynamicFactor, _
        dbl_PlugDiameter, dbl_SeatDiameter, dbl_StemDiameter, dbl_PackingHeight, _
        dbl_AllowPlugDP, dbl_StemStressArea, dbl_Travel, str_StemMaterial)
    CnvrtLengthToMM dbl_Travel, "in", dbl_Travel
    CnvrtLengthFromMM dbl_Travel, str_LengthUM, dbl_Travel

'calculate required packing force
Call CalcPackingForce(dbl_MaxInletPressure, dbl_StemDiameter, str_PackingMaterial, _
    dbl_PackingHeight, dbl_ForcePacking)

'calculate required seating force
Call CalcSeatingForce(str_LeakageClass, dbl_SeatDiameter, dbl_ForceSeating)

'calculate required offbalance force
dbl_EffectiveOrificeArea = 0.2 'default
Call CalcSingleSeatUnbalancedForceFTC(dbl_MaxInletPressure, dbl_SeatDiameter, _
    dbl_StemDiameter, dbl_EffectiveOrificeArea, dbl_ForceUnbalancedClosed, _
    dbl_ForceUnbalancedOpen)

dbl_ReqForceClosed = dbl_ForcePacking + dbl_ForceSeating + dbl_ForceUnbalancedClosed
dbl_ReqForceOpen = dbl_ForcePacking + dbl_ForceUnbalancedOpen

'calculate actuator capability

'-----------------------------------------------------------------------------------------
'If debug_on = False Then
    Call GetActuatorData(long_ActuatorPartID, dbl_AreaInitial, _
        dbl_AreaFinal, dbl_MaxSupplyPressure, dbl_ActuatorFriction, _
        dbl_PositionerLossCoef, str_Model, long_AirActionID, _
        int_QtySeals, dbl_StemSealWidth, dbl_PistonRingWidth, _
        dbl_VolOver, dbl_ForceSpringClosed, dbl_ActuatorStemDiameter, dbl_LeverLength)
'Else
'    Call screenActuatorData(str_ActuatorModel, str_actuatorSize, dbl_Travel, _
'        dbl_AreaInitial, dbl_AreaFinal, dbl_MaxSupplyPressure, _
'        dbl_ActuatorFriction, dbl_PositionerLossCoef, str_SpringRange, str_LengthUM)
'End If
'-----------------------------------------------------------------------------------------
    
Call CalcSpringRate(str_SpringRange, dbl_AreaInitial, dbl_PInitial, dbl_AreaFinal, _
    dbl_PFinal, str_CalibPressureUM, dbl_Travel, str_LengthUM, dbl_SpringRate)

Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, dbl_ReqForceClosed, _
    dbl_PSupplyDP)


Call CalcATCActuatorForce(dbl_AreaFinal, dbl_PFinal, dbl_PSupplyDP, _
    dbl_ActuatorFriction, dbl_ActuatorForce)

'calculate allowable pressure drop
Call CalcAllowDPSingleStageUnbalFTC_ATC(dbl_AreaInitial, dbl_PInitial, dbl_AreaFinal, _
    dbl_PFinal, dbl_SpringRate, dbl_ForcePacking, dbl_SeatDiameter, _
    dbl_StemDiameter, dbl_DynamicFactor, dbl_EffectiveOrificeArea, dbl_AllowDPStability, _
    dbl_AllowDPClosed, dbl_AllowDPOpen)

'calculate stem limitation
'Call GetStemStressArea(dbl_StemDiameter, dbl_StemStressArea) now returned by get21000data
'-----------------------------------------------------------------------------------------
'If debug_on = False Then
    Call GetStemProperties(str_StemMaterial, dbl_MaxInletTemperature, dbl_Sy, _
        dbl_E, dbl_Sd)
'Else
'    Call screenStemData(dbl_E, dbl_Sd, dbl_Sy)
'End If
'-----------------------------------------------------------------------------------------

dbl_L = 6# 'default value


'---------revision jh 8/25/98-------------------------------------------------------------

Call CalcAllowStemForce(dbl_Sy, dbl_E, dbl_StemDiameter, dbl_L, _
    dbl_Sd, dbl_StemStressArea, dbl_AllowableStemLoad)
    
'Call CalcAllowStemTensileForce(dbl_Sd, dbl_stemStressArea, dbl_AllowableTensileForce)
'
'If (dbl_AllowableTensileForce < dbl_AllowableCompressiveForce) Then
'    dbl_AllowableStemLoad = dbl_AllowableTensileForce
'Else
'    dbl_AllowableStemLoad = dbl_AllowableCompressiveForce
'End If
'-----------------------------------------------------------------------------------------
'calculate actuator pressure drop capability
Call CalcMaxDPAllowable(dbl_AllowDPStability, dbl_AllowDPClosed, dbl_AllowDPOpen, _
    dbl_AllowANSIPressure, dbl_MaxAllowDP)

'-----------------------------------------------------------------------------------------
'If debug_on Then
'
'    form21000.txtAllowStemLoad.Text = dbl_AllowableStemLoad
'    form21000.txtAllowANSIpressure = dbl_AllowANSIPressure
'    form21000.txtPackingForce.Text = Format(dbl_ForcePacking, "#####")
'    form21000.txtSeatingForce = Format(dbl_ForceSeating, "#####")
'
'    form21000.txtUnbalForceClosed.Text = Format(dbl_ForceUnbalancedClosed, "#####")
'    form21000.txtUnbalForceOpen.Text = Format(dbl_ForceUnbalancedOpen, "#####")
'    form21000.txtReqForceClosed.Text = Format(dbl_ReqForceClosed, "#####")
'    form21000.txtReqForceOpen.Text = Format(dbl_ReqForceOpen, "#####")
'
'    form21000.txtActuatorForce.Text = dbl_ActuatorForce
'    form21000.txtPsupply.Text = dbl_PSupply
'
'End If
'-----------------------------------------------------------------------------------------


'evaluation
str_Message = ""
dbl_ShowDP = ConvertPressureFromPSIG(dbl_MaxInletPressure, str_PressureUM)
dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, _
    str_CalibPressureUM)

dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, str_PressureUM)

dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, _
    str_CalibPressureUM)
    
If (dbl_MaxInletPressure > dbl_AllowANSIPressure) Then
    'bool_ANSIOK = False
    str_Message = ANSIMESSAGE
    dbl_ShowDP = ConvertPressureFromPSIG(dbl_MaxInletPressure, str_PressureUM)
    dbl_ShowMaxDP = dbl_AllowANSIPressure
    dbl_ShowMaxDPSupplyPressure = 0
    '---------added 10/26/98--------
    dbl_ShowDPSupplyPressure = 0
    '-------------------------------
    GoTo EndOfRoutine
End If

If (dbl_AllowPlugDP > 0 And dbl_MaxInletPressure > dbl_AllowPlugDP) Then
    'bool_TrimOK = False
    str_Message = TRIMMESSAGE
    dbl_ShowDP = ConvertPressureFromPSIG(dbl_MaxInletPressure, str_PressureUM)
    dbl_ShowMaxDP = dbl_AllowPlugDP
    dbl_ShowMaxDPSupplyPressure = dbl_MaxAllowSupply
    GoTo EndOfRoutine
End If

If (dbl_ActuatorForce > dbl_AllowableStemLoad) Then
    str_Message = STEMMESSAGE
    dbl_ShowDPSupplyPressure = 0
    dbl_ShowMaxDPSupplyPressure = 0
End If

If (dbl_MaxInletPressure > dbl_AllowDPClosed) Then
    str_Message = ACTUATORMESSAGE + " - CLOSED"
    dbl_ShowMaxDP = dbl_AllowDPClosed
    dbl_ShowDPSupplyPressure = 0
    dbl_ShowMaxDPSupplyPressure = 0
End If

If (dbl_MaxPressureDrop > dbl_AllowDPOpen) Then
    str_Message = ACTUATORMESSAGE + " - OPEN" + Chr(13) + "PRESSURE DROP EXCEEDS ALLOWABLE LIMIT"
    dbl_ShowMaxDP = dbl_AllowDPOpen
    dbl_ShowDPSupplyPressure = 0
    dbl_ShowMaxDPSupplyPressure = 0
End If

'---------added 10/26/98--------
If (dbl_MaxInletPressure > dbl_AllowDPOpen) Then
    str_Message = ACTUATORMESSAGE + " - OPEN" + Chr(13) + "INLET PRESSURE EXCEEDS ALLOWABLE PRESSURE DROP"
    dbl_ShowMaxDP = dbl_AllowDPOpen
    dbl_ShowDPSupplyPressure = 0
    dbl_ShowMaxDPSupplyPressure = 0
End If
'-----------------------------------------------------------------------------------------

If (dbl_MaxPressureDrop > dbl_AllowDPStability) Then
    str_Message = ACTUATORMESSAGE + " - STABILITY"
    dbl_ShowMaxDP = dbl_AllowDPStability
    dbl_ShowDPSupplyPressure = 0
    dbl_ShowMaxDPSupplyPressure = 0
End If


If (dbl_ActuatorForce < dbl_ReqForceOpen) Then
    str_Message = ACTUATORMESSAGE
End If

If (dbl_MaxInletPressure <= dbl_MaxAllowDP) Then
    str_Message = ""
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, _
        str_CalibPressureUM)
    dbl_ShowMaxDP = dbl_ShowDP
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, _
        str_CalibPressureUM)
End If

If (dbl_MaxInletPressure <= dbl_ShutOffDP And dbl_ShutOffDP <= dbl_MaxAllowDP And dbl_ShutOffDP > 0#) Then
    str_Message = LIMITEDDPMESSAGE
    dbl_ShowDP = dbl_ShowShutoffDP
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_CalibPressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, str_PressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, _
        str_CalibPressureUM)
End If
   
EndOfRoutine:
dbl_ShowDP = TrimDecimalPlaces(dbl_ShowDP)
dbl_ShowDPSupplyPressure = TrimDecimalPlaces(dbl_ShowDPSupplyPressure)
dbl_ShowMaxDP = TrimDecimalPlaces(dbl_ShowMaxDP)
dbl_ShowMaxDPSupplyPressure = TrimDecimalPlaces(dbl_ShowMaxDPSupplyPressure)

#If (DebugVersion = "yes") Then
MsgBox "long_TrimPackageID= " + CStr(long_TrimPackageID) + Chr(13) + Chr(10) + "long_ActuatorPartID= " + CStr(long_ActuatorPartID) + Chr(13) + Chr(10) + "dbl_InletPressure1= " + CStr(dbl_InletPressure1) + Chr(13) + Chr(10) + _
"dbl_InletPressure2= " + CStr(dbl_InletPressure2) + Chr(13) + Chr(10) + "dbl_InletPressure3= " + CStr(dbl_InletPressure3) + Chr(13) + Chr(10) + "dbl_InletPressure4= " + CStr(dbl_InletPressure4) + Chr(13) + Chr(10) + _
"dbl_OutletPressure1= " + CStr(dbl_OutletPressure1) + Chr(13) + Chr(10) + "dbl_OutletPressure2= " + CStr(dbl_OutletPressure2) + Chr(13) + Chr(10) + "dbl_OutletPressure3= " + CStr(dbl_OutletPressure3) + Chr(13) + Chr(10) + _
"dbl_OutletPressure4= " + CStr(dbl_OutletPressure4) + Chr(13) + Chr(10) + "dbl_InletTemperature1= " + CStr(dbl_InletTemperature1) + Chr(13) + Chr(10) + "dbl_InletTemperature2= " + CStr(dbl_InletTemperature2) + Chr(13) + Chr(10) + _
"dbl_InletTemperature3= " + CStr(dbl_InletTemperature3) + Chr(13) + Chr(10) + "dbl_InletTemperature4= " + CStr(dbl_InletTemperature4) + Chr(13) + Chr(10) + "dbl_ShutoffDP= " + CStr(dbl_ShutOffDP) + Chr(13) + Chr(10) + "str_PressureUM= " + str_PressureUM + _
Chr(13) + Chr(10) + "str_TemperatureUM= " + str_TemperatureUM + Chr(13) + Chr(10) + "str_CalibPressureUM= " + str_CalibPressureUM + Chr(13) + Chr(10) + "str_BodyMaterial= " + str_BodyMaterial + Chr(13) + Chr(10) + "str_Rating= " + str_Rating + _
Chr(13) + Chr(10) + "str_PackingMaterial= " + str_PackingMaterial + Chr(13) + Chr(10) + "str_LeakageClass= " + str_LeakageClass + Chr(13) + Chr(10) + "str_SpringRange= " + str_SpringRange + Chr(13) + Chr(10) + "dbl_Travel= " + _
CStr(dbl_Travel) + Chr(13) + Chr(10) + "str_LengthUM= " + str_LengthUM + Chr(13) + Chr(10) + "dbl_ShowDP= " + CStr(dbl_ShowDP) + Chr(13) + Chr(10) + "dbl_ShowDPSupplyPressure= " + CStr(dbl_ShowDPSupplyPressure) + Chr(13) + Chr(10) _
+ "dbl_ShowMaxDP= " + CStr(dbl_ShowMaxDP) + Chr(13) + Chr(10) + "dbl_ShowMaxDPSupplyPressure= " + CStr(dbl_ShowMaxDPSupplyPressure) + Chr(13) + Chr(10) + "str_Message= " + str_Message
#End If

Exit Sub

CalcDP_21000_1Stage_FTC_ATC_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcDP_21000_2Stage_FTO_ATO(long_TrimPackageID As Long, _
    long_ActuatorPartID As Long, _
    dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, _
    dbl_InletPressure3 As Double, dbl_InletPressure4 As Double, _
    dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, _
    dbl_InletTemperature1 As Double, dbl_InletTemperature2 As Double, _
    dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, _
    dbl_ShutOffDP As Double, str_PressureUM As String, _
    str_TemperatureUM As String, str_CalibPressureUM As String, str_BodyMaterial As String, _
    str_Rating As String, str_PackingMaterial As String, _
    str_LeakageClass As String, str_SpringRange As String, dbl_Travel As Double, str_LengthUM As String, _
    dbl_ShowDP As Double, dbl_ShowDPSupplyPressure As Double, dbl_ShowMaxDP As Double, _
    dbl_ShowMaxDPSupplyPressure As Double, str_Message As String)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the selected valve and actuator shutoff capability and compare the result to defined requirements.
'This subroutine is specific to 21000 seies valves with double stage trim, flow to open, air to open.
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID
'long_ActuatorPartID
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff pressure drop
'str_PressureUM                 unit of measure associated with inlet and outlet pressures
'str_TemperatureUM              unit of measure associated with inlet temperature
'str_CalibPressureUM            unit of measure associated with spring range and supply pressure
'str_LengthUM                   unit of measure for length (Travel)
'str_BodyMaterial
'str_Rating
'str_PackingMaterial            selected packing material
'str_LeakageClass               selected leakage class
'str_SpringRange                selected actuator spring range
'dbl_Travel                     Actuator travel
'
'Output Variables:
'dbl_ShowDP                     specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure       calculated required supply pressure for specified
'dbl_ShowMaxDP                  calculated maximum pressure drop for selected valve and actuator combination
'dbl_ShowMaxDPsupplyPressure    calculated required supply pressure for maximum pressure drop
'str_Message                    warning message if pressure drop requirement exceeds equipment capability
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/17 Larry Schoonover - not tested
'
'revised 7/8/98 John Haberlin - evaluation logic revised during testing
'                               test code remains as commented text
'revised 8/25/98 John Haberlin - duplicate stem calculation and evaluation removed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcDP_21000_2Stage_FTO_ATO_Error

#If (DebugVersion = "yes") Then
MsgBox "long_TrimPackageID= " + CStr(long_TrimPackageID) + Chr(13) + Chr(10) + "long_ActuatorPartID= " + CStr(long_ActuatorPartID) + Chr(13) + Chr(10) + "dbl_InletPressure1= " + CStr(dbl_InletPressure1) + Chr(13) + Chr(10) + _
"dbl_InletPressure2= " + CStr(dbl_InletPressure2) + Chr(13) + Chr(10) + "dbl_InletPressure3= " + CStr(dbl_InletPressure3) + Chr(13) + Chr(10) + "dbl_InletPressure4= " + CStr(dbl_InletPressure4) + Chr(13) + Chr(10) + _
"dbl_OutletPressure1= " + CStr(dbl_OutletPressure1) + Chr(13) + Chr(10) + "dbl_OutletPressure2= " + CStr(dbl_OutletPressure2) + Chr(13) + Chr(10) + "dbl_OutletPressure3= " + CStr(dbl_OutletPressure3) + Chr(13) + Chr(10) + _
"dbl_OutletPressure4= " + CStr(dbl_OutletPressure4) + Chr(13) + Chr(10) + "dbl_InletTemperature1= " + CStr(dbl_InletTemperature1) + Chr(13) + Chr(10) + "dbl_InletTemperature2= " + CStr(dbl_InletTemperature2) + Chr(13) + Chr(10) + _
"dbl_InletTemperature3= " + CStr(dbl_InletTemperature3) + Chr(13) + Chr(10) + "dbl_InletTemperature4= " + CStr(dbl_InletTemperature4) + Chr(13) + Chr(10) + "dbl_ShutoffDP= " + CStr(dbl_ShutOffDP) + Chr(13) + Chr(10) + "str_PressureUM= " + str_PressureUM + _
Chr(13) + Chr(10) + "str_TemperatureUM= " + str_TemperatureUM + Chr(13) + Chr(10) + "str_CalibPressureUM= " + str_CalibPressureUM + Chr(13) + Chr(10) + "str_BodyMaterial= " + str_BodyMaterial + Chr(13) + Chr(10) + "str_Rating= " + str_Rating + _
Chr(13) + Chr(10) + "str_PackingMaterial= " + str_PackingMaterial + Chr(13) + Chr(10) + "str_LeakageClass= " + str_LeakageClass + Chr(13) + Chr(10) + "str_SpringRange= " + str_SpringRange + Chr(13) + Chr(10) + "dbl_Travel= " + _
CStr(dbl_Travel) + Chr(13) + Chr(10) + "str_LengthUM= " + str_LengthUM + Chr(13) + Chr(10) + "dbl_ShowDP= " + CStr(dbl_ShowDP) + Chr(13) + Chr(10) + "dbl_ShowDPSupplyPressure= " + CStr(dbl_ShowDPSupplyPressure) + Chr(13) + Chr(10) _
+ "dbl_ShowMaxDP= " + CStr(dbl_ShowMaxDP) + Chr(13) + Chr(10) + "dbl_ShowMaxDPSupplyPressure= " + CStr(dbl_ShowMaxDPSupplyPressure) + Chr(13) + Chr(10) + "str_Message= " + str_Message
#End If


'***************************
'Declarations
'***************************
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_AllowANSIPressure As Double
Dim dbl_DynamicFactor As Double
Dim dbl_PlugDiameter As Double
Dim dbl_SeatDiameter As Double
Dim dbl_StemDiameter As Double
Dim dbl_PackingHeight As Double
Dim dbl_AllowPlugDP As Double
Dim dbl_ForcePacking As Double
Dim dbl_ForceSeating As Double
Dim dbl_ForceUnbalancedClosed As Double
Dim dbl_ForceUnbalancedOpen As Double
Dim dbl_ReqForceClosed As Double
Dim dbl_ReqForceOpen As Double
Dim dbl_AreaInitial As Double
Dim dbl_AreaFinal As Double
Dim dbl_MaxSupplyPressure As Double
Dim dbl_ActuatorFriction As Double
Dim dbl_PositionerLossCoef As Double
Dim dbl_PInitial As Double
Dim dbl_PFinal As Double
Dim dbl_SpringRate As Double
Dim dbl_PSupply As Double
Dim dbl_ActuatorForce As Double
Dim dbl_AllowDPClosed As Double
Dim dbl_StemStressArea As Double
Dim dbl_Sy As Double
Dim dbl_E As Double
Dim dbl_Sd As Double
Dim dbl_L As Double
Dim dbl_AllowableCompressiveForce As Double
Dim dbl_AllowableTensileForce As Double
Dim dbl_AllowableStemLoad As Double
Dim dbl_AllowDPStability  As Double
Dim dbl_AllowDPOpen As Double
Dim dbl_MaxAllowDP As Double
Dim dbl_PSupplyDP As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim str_StemMaterial As String
Dim dbl_MaxPressureDrop As Double
Dim dbl_ShowMaxPressureDrop As Double

Dim dbl_InletPressure(4) As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double
Dim int_QtySeals As Integer
Dim dbl_StemSealWidth As Double
Dim dbl_PistonRingWidth As Double
Dim dbl_VolOver As Double
Dim dbl_ForceSpringClosed As Double
Dim dbl_ActuatorStemDiameter As Double
Dim dbl_LeverLength As Double

Dim str_NotUsed As String
Dim long_NotUsed As Long
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'If debug_on = True Then
'    Dim dbl_packingCoefficient As Double
'    Dim dbl_Kv As Double
'    Dim dbl_Fcs As Double
'    Dim dbl_Au As Double
'
'    Dim str_ActuatorModel As String
'    Dim str_actuatorSize As String
'End If
'-----------------------------------------------------------------------------------------

'***************************
'Parameter validation
'***************************
'default for now
If (str_PackingMaterial = "") Then str_PackingMaterial = "PTFE"

If (str_PressureUM = "") Then GoTo CalcDP_21000_2Stage_FTO_ATO_Error
If (str_TemperatureUM = "") Then GoTo CalcDP_21000_2Stage_FTO_ATO_Error
If (str_CalibPressureUM = "") Then GoTo CalcDP_21000_2Stage_FTO_ATO_Error
If (str_PackingMaterial = "") Then GoTo CalcDP_21000_2Stage_FTO_ATO_Error
If (str_LeakageClass = "") Then GoTo CalcDP_21000_2Stage_FTO_ATO_Error
If (str_SpringRange = "") Then GoTo CalcDP_21000_2Stage_FTO_ATO_Error

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4

'determine requirements

Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature(), dbl_MaxInletTemperature)
Call CalcTestPressureDrop(dbl_InletPressure(), dbl_OutletPressure(), _
    dbl_MaxPressureDrop)
    
dbl_ShowDP = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, str_PressureUM)
dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, str_TemperatureUM)
dbl_ShowMaxPressureDrop = dbl_MaxPressureDrop
dbl_MaxPressureDrop = ConvertPressureToPSIG(dbl_MaxPressureDrop, str_PressureUM)
dbl_ShowShutoffDP = dbl_ShutOffDP
If (dbl_ShutOffDP > 0) Then dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, str_PressureUM)

'determine ANSI limitation

'-----------------------------------------------------------------------------------------
'If debug_on = False Then
    Call GetANSIPressure(str_BodyMaterial, dbl_MaxInletTemperature, str_Rating, _
    dbl_AllowANSIPressure)
'Else
'    dbl_AllowANSIPressure = 570
'End If
'-----------------------------------------------------------------------------------------

'read valve properties

    Call Get21000Data(long_TrimPackageID, dbl_DynamicFactor, _
        dbl_PlugDiameter, dbl_SeatDiameter, dbl_StemDiameter, dbl_PackingHeight, _
        dbl_AllowPlugDP, dbl_StemStressArea, dbl_Travel, str_StemMaterial)
    CnvrtLengthToMM dbl_Travel, "in", dbl_Travel
    CnvrtLengthFromMM dbl_Travel, str_LengthUM, dbl_Travel
    
'calculate required packing force
Call CalcPackingForce(dbl_MaxInletPressure, dbl_StemDiameter, str_PackingMaterial, _
    dbl_PackingHeight, dbl_ForcePacking)

'calculate required seating force
Call CalcSeatingForce(str_LeakageClass, dbl_SeatDiameter, dbl_ForceSeating)

'calculate required offbalance force

Call CalcDoubleStageForceFTO(dbl_MaxPressureDrop, dbl_SeatDiameter, dbl_PlugDiameter, _
    dbl_ForceUnbalancedClosed, dbl_ForceUnbalancedOpen)

    dbl_ReqForceClosed = dbl_ForcePacking + dbl_ForceSeating + dbl_ForceUnbalancedClosed
    dbl_ReqForceOpen = dbl_ForcePacking + dbl_ForceUnbalancedOpen

'calculate actuator capability

'-----------------------------------------------------------------------------------------
'If debug_on = False Then
    Call GetActuatorData(long_ActuatorPartID, dbl_AreaInitial, dbl_AreaFinal, _
        dbl_MaxSupplyPressure, dbl_ActuatorFriction, dbl_PositionerLossCoef, _
        str_NotUsed, long_NotUsed, _
        int_QtySeals, dbl_StemSealWidth, dbl_PistonRingWidth, _
        dbl_VolOver, dbl_ForceSpringClosed, dbl_ActuatorStemDiameter, dbl_LeverLength)
'Else
'    Call screenActuatorData(str_ActuatorModel, str_actuatorSize, dbl_Travel, _
'        dbl_AreaInitial, dbl_AreaFinal, dbl_MaxSupplyPressure, _
'        dbl_ActuatorFriction, dbl_PositionerLossCoef, str_SpringRange, str_LengthUM)
'End If
'-----------------------------------------------------------------------------------------

Call CalcSpringRate(str_SpringRange, dbl_AreaInitial, dbl_PInitial, _
    dbl_AreaFinal, dbl_PFinal, str_CalibPressureUM, dbl_Travel, str_LengthUM, _
    dbl_SpringRate)
    
Call CalcATOActuatorForce(dbl_AreaInitial, dbl_PInitial, dbl_PFinal, _
    dbl_ActuatorFriction, dbl_PSupply, dbl_ActuatorForce)
        
Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, dbl_ReqForceOpen, _
    dbl_PSupply)
    
'calculate allowable pressure drop

Call CalcAllowDPDoubleStageFTO(dbl_ActuatorForce, dbl_SpringRate, _
    dbl_ForcePacking, dbl_ForceSeating, dbl_SeatDiameter, dbl_PlugDiameter, _
    dbl_DynamicFactor, dbl_AllowDPStability, dbl_AllowDPClosed, dbl_AllowDPOpen)
    
Call CalcMaxDPAllowable(dbl_AllowDPStability, dbl_AllowDPClosed, dbl_AllowDPOpen, _
    dbl_AllowANSIPressure, dbl_MaxAllowDP)

'calculate stem limitation
'Call GetStemStressArea(dbl_StemDiameter, dbl_StemStressArea) now returned by get21000data

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'If debug_on = False Then
    Call GetStemProperties(str_StemMaterial, dbl_MaxInletTemperature, dbl_Sy, _
        dbl_E, dbl_Sd)
'Else
'    Call screenStemData(dbl_E, dbl_Sd, dbl_Sy)
'End If
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

dbl_L = 6# 'default value

'***********************revision jh 8/25/98********************************************

Call CalcAllowStemForce(dbl_Sy, dbl_E, dbl_StemDiameter, dbl_L, _
    dbl_Sd, dbl_StemStressArea, dbl_AllowableStemLoad)
'Call CalcAllowStemTensileForce(dbl_Sd, dbl_stemStressArea, dbl_AllowableTensileForce)
'
'If (dbl_AllowableTensileForce < dbl_AllowableCompressiveForce) Then
'    dbl_AllowableStemLoad = dbl_AllowableTensileForce
'Else
'    dbl_AllowableStemLoad = dbl_AllowableCompressiveForce
'End If

'-----------------------------------------------------------------------------------------
'If debug_on Then
'
'    form21000.txtAllowStemLoad.Text = dbl_AllowableStemLoad
'    form21000.txtAllowANSIpressure = dbl_AllowANSIPressure
'    form21000.txtPackingForce.Text = Format(dbl_ForcePacking, "#####")
'    form21000.txtSeatingForce = Format(dbl_ForceSeating, "#####")
'
'    form21000.txtUnbalForceClosed.Text = Format(dbl_ForceUnbalancedClosed, "#####")
'    form21000.txtUnbalForceOpen.Text = Format(dbl_ForceUnbalancedOpen, "#####")
'    form21000.txtReqForceClosed.Text = Format(dbl_ReqForceClosed, "#####")
'    form21000.txtReqForceOpen.Text = Format(dbl_ReqForceOpen, "#####")
'
'    form21000.txtActuatorForce.Text = dbl_ActuatorForce
'    form21000.txtPsupply.Text = dbl_PSupply
'
'End If
'-----------------------------------------------------------------------------------------

'evaluation

str_Message = ""

dbl_ShowDP = 0
dbl_ShowDPSupplyPressure = 0
dbl_ShowMaxDP = 0
dbl_ShowMaxDPSupplyPressure = 0

If (dbl_MaxInletPressure > dbl_AllowANSIPressure) Then
    str_Message = ANSIMESSAGE
    dbl_ShowDP = ConvertPressureFromPSIG(dbl_MaxInletPressure, str_PressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowANSIPressure, str_PressureUM)
    GoTo EndOfRoutine
End If

If (dbl_MaxInletPressure > dbl_AllowPlugDP And dbl_AllowPlugDP > 0) Then
    str_Message = TRIMMESSAGE
    dbl_ShowDP = ConvertPressureFromPSIG(dbl_MaxInletPressure, str_PressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowPlugDP, str_PressureUM)
    GoTo EndOfRoutine
End If

If (dbl_ActuatorForce > dbl_AllowableStemLoad) Then
    str_Message = STEMMESSAGE
    GoTo EndOfRoutine
End If
    
If (dbl_MaxInletPressure > dbl_AllowDPClosed Or dbl_MaxPressureDrop > dbl_AllowDPOpen _
    Or dbl_MaxPressureDrop > dbl_AllowDPStability) Then str_Message = ACTUATORMESSAGE

If (dbl_ActuatorForce < (dbl_ReqForceClosed - dbl_ForceUnbalancedOpen)) Then _
    str_Message = ACTUATORMESSAGE


dbl_ShowDP = dbl_MaxInletPressure
dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupply, str_CalibPressureUM)
dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, str_PressureUM)
dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupply, str_CalibPressureUM)

If (dbl_MaxInletPressure > dbl_AllowDPClosed) Then
    str_Message = ACTUATORMESSAGE + " - CLOSED"
    dbl_ShowDPSupplyPressure = 0
End If

If (dbl_MaxPressureDrop > dbl_AllowDPOpen) Then
    str_Message = ACTUATORMESSAGE + " - OPEN" + Chr(13) + "PRESSURE DROP EXCEEDS ALLOWABLE LIMIT"
    dbl_ShowDPSupplyPressure = 0
End If

If (dbl_MaxPressureDrop > dbl_AllowDPStability) Then
    str_Message = ACTUATORMESSAGE + " - STABILITY"
    dbl_ShowDPSupplyPressure = 0
End If

If (dbl_MaxInletPressure > dbl_ShutOffDP And dbl_ShutOffDP <= dbl_AllowDPClosed And dbl_ShutOffDP > 0# _
        And dbl_MaxPressureDrop < dbl_AllowDPOpen And dbl_MaxInletPressure > dbl_AllowDPClosed) Then
    str_Message = LIMITEDDPMESSAGE
    dbl_ShowDP = dbl_ShowShutoffDP
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupply, str_CalibPressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, str_PressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupply, str_CalibPressureUM)

End If

EndOfRoutine:
dbl_ShowDP = TrimDecimalPlaces(dbl_ShowDP)
dbl_ShowDPSupplyPressure = TrimDecimalPlaces(dbl_ShowDPSupplyPressure)
dbl_ShowMaxDP = TrimDecimalPlaces(dbl_ShowMaxDP)
dbl_ShowMaxDPSupplyPressure = TrimDecimalPlaces(dbl_ShowMaxDPSupplyPressure)

#If (DebugVersion = "yes") Then
MsgBox "long_TrimPackageID= " + CStr(long_TrimPackageID) + Chr(13) + Chr(10) + "long_ActuatorPartID= " + CStr(long_ActuatorPartID) + Chr(13) + Chr(10) + "dbl_InletPressure1= " + CStr(dbl_InletPressure1) + Chr(13) + Chr(10) + _
"dbl_InletPressure2= " + CStr(dbl_InletPressure2) + Chr(13) + Chr(10) + "dbl_InletPressure3= " + CStr(dbl_InletPressure3) + Chr(13) + Chr(10) + "dbl_InletPressure4= " + CStr(dbl_InletPressure4) + Chr(13) + Chr(10) + _
"dbl_OutletPressure1= " + CStr(dbl_OutletPressure1) + Chr(13) + Chr(10) + "dbl_OutletPressure2= " + CStr(dbl_OutletPressure2) + Chr(13) + Chr(10) + "dbl_OutletPressure3= " + CStr(dbl_OutletPressure3) + Chr(13) + Chr(10) + _
"dbl_OutletPressure4= " + CStr(dbl_OutletPressure4) + Chr(13) + Chr(10) + "dbl_InletTemperature1= " + CStr(dbl_InletTemperature1) + Chr(13) + Chr(10) + "dbl_InletTemperature2= " + CStr(dbl_InletTemperature2) + Chr(13) + Chr(10) + _
"dbl_InletTemperature3= " + CStr(dbl_InletTemperature3) + Chr(13) + Chr(10) + "dbl_InletTemperature4= " + CStr(dbl_InletTemperature4) + Chr(13) + Chr(10) + "dbl_ShutoffDP= " + CStr(dbl_ShutOffDP) + Chr(13) + Chr(10) + "str_PressureUM= " + str_PressureUM + _
Chr(13) + Chr(10) + "str_TemperatureUM= " + str_TemperatureUM + Chr(13) + Chr(10) + "str_CalibPressureUM= " + str_CalibPressureUM + Chr(13) + Chr(10) + "str_BodyMaterial= " + str_BodyMaterial + Chr(13) + Chr(10) + "str_Rating= " + str_Rating + _
Chr(13) + Chr(10) + "str_PackingMaterial= " + str_PackingMaterial + Chr(13) + Chr(10) + "str_LeakageClass= " + str_LeakageClass + Chr(13) + Chr(10) + "str_SpringRange= " + str_SpringRange + Chr(13) + Chr(10) + "dbl_Travel= " + _
CStr(dbl_Travel) + Chr(13) + Chr(10) + "str_LengthUM= " + str_LengthUM + Chr(13) + Chr(10) + "dbl_ShowDP= " + CStr(dbl_ShowDP) + Chr(13) + Chr(10) + "dbl_ShowDPSupplyPressure= " + CStr(dbl_ShowDPSupplyPressure) + Chr(13) + Chr(10) _
+ "dbl_ShowMaxDP= " + CStr(dbl_ShowMaxDP) + Chr(13) + Chr(10) + "dbl_ShowMaxDPSupplyPressure= " + CStr(dbl_ShowMaxDPSupplyPressure) + Chr(13) + Chr(10) + "str_Message= " + str_Message
#End If

Exit Sub

CalcDP_21000_2Stage_FTO_ATO_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcDP_21000_2Stage_FTO_ATC(long_TrimPackageID As Long, _
    long_ActuatorPartID As Long, _
    dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, _
    dbl_InletPressure3 As Double, dbl_InletPressure4 As Double, _
    dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, _
    dbl_InletTemperature1 As Double, dbl_InletTemperature2 As Double, _
    dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, _
    dbl_ShutOffDP As Double, str_PressureUM As String, _
    str_TemperatureUM As String, str_CalibPressureUM As String, str_BodyMaterial As String, _
    str_Rating As String, str_PackingMaterial As String, _
    str_LeakageClass As String, str_SpringRange As String, dbl_Travel As Double, _
    str_LengthUM As String, dbl_MaxAllowSupplyPressure As Double, _
    dbl_ShowDP As Double, dbl_ShowDPSupplyPressure As Double, dbl_ShowMaxDP As Double, _
    dbl_ShowMaxDPSupplyPressure As Double, str_Message As String)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the selected valve and actuator shutoff capability and compare the result to defined requirements.
'This subroutine is specific to 21000 seies valves with double stage trim, flow to open, air to open.
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID
'long_ActuatorPartID
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff pressure drop
'str_PressureUM                 unit of measure associated with inlet and outlet pressures
'str_TemperatureUM              unit of measure associated with inlet temperature
'str_CalibPressureUM            unit of measure associated with spring range and supply pressure
'str_LengthUM                   unit of measure for length (Travel)
'dbl_MaxAllowSupplyPressure     available supply pressure as defined by user
'str_BodyMaterial
'str_Rating
'str_PackingMaterial            selected packing material
'str_LeakageClass               selected leakage class
'str_SpringRange                selected actuator spring range
'dbl_Travel                     Actuator travel
'
'Output Variables:
'dbl_ShowDP                     specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure       calculated required supply pressure for specified
'dbl_ShowMaxDP                  calculated maximum pressure drop for selected valve and actuator combination
'dbl_ShowMaxDPsupplyPressure    calculated required supply pressure for maximum pressure drop
'str_Message                    warning message if pressure drop requirement exceeds equipment capability
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/17 Larry Schoonover - not tested
'
'revised 7/ 8/98 John Haberlin - evaluation logic revised during testing
'                               test code remains as commented text
'revised 8/25/98 John Haberlin - duplicate stem calculation and evaluation removed
'revised 9/22/98 John Haberlin - added "dbl_MaxAllowSupplyPressure" as a required input variable
'-----------------------------------------------------------------------------------------

On Error GoTo CalcDP_21000_2Stage_FTO_ATC_Error

#If (DebugVersion = "yes") Then
MsgBox "long_TrimPackageID= " + CStr(long_TrimPackageID) + Chr(13) + Chr(10) + "long_ActuatorPartID= " + CStr(long_ActuatorPartID) + Chr(13) + Chr(10) + "dbl_InletPressure1= " + CStr(dbl_InletPressure1) + Chr(13) + Chr(10) + _
"dbl_InletPressure2= " + CStr(dbl_InletPressure2) + Chr(13) + Chr(10) + "dbl_InletPressure3= " + CStr(dbl_InletPressure3) + Chr(13) + Chr(10) + "dbl_InletPressure4= " + CStr(dbl_InletPressure4) + Chr(13) + Chr(10) + _
"dbl_OutletPressure1= " + CStr(dbl_OutletPressure1) + Chr(13) + Chr(10) + "dbl_OutletPressure2= " + CStr(dbl_OutletPressure2) + Chr(13) + Chr(10) + "dbl_OutletPressure3= " + CStr(dbl_OutletPressure3) + Chr(13) + Chr(10) + _
"dbl_OutletPressure4= " + CStr(dbl_OutletPressure4) + Chr(13) + Chr(10) + "dbl_InletTemperature1= " + CStr(dbl_InletTemperature1) + Chr(13) + Chr(10) + "dbl_InletTemperature2= " + CStr(dbl_InletTemperature2) + Chr(13) + Chr(10) + _
"dbl_InletTemperature3= " + CStr(dbl_InletTemperature3) + Chr(13) + Chr(10) + "dbl_InletTemperature4= " + CStr(dbl_InletTemperature4) + Chr(13) + Chr(10) + "dbl_ShutoffDP= " + CStr(dbl_ShutOffDP) + Chr(13) + Chr(10) + "str_PressureUM= " + str_PressureUM + _
Chr(13) + Chr(10) + "str_TemperatureUM= " + str_TemperatureUM + Chr(13) + Chr(10) + "str_CalibPressureUM= " + str_CalibPressureUM + Chr(13) + Chr(10) + "str_BodyMaterial= " + str_BodyMaterial + Chr(13) + Chr(10) + "str_Rating= " + str_Rating + _
Chr(13) + Chr(10) + "str_PackingMaterial= " + str_PackingMaterial + Chr(13) + Chr(10) + "str_LeakageClass= " + str_LeakageClass + Chr(13) + Chr(10) + "str_SpringRange= " + str_SpringRange + Chr(13) + Chr(10) + "dbl_Travel= " + _
CStr(dbl_Travel) + Chr(13) + Chr(10) + "str_LengthUM= " + str_LengthUM + Chr(13) + Chr(10) + "dbl_ShowDP= " + CStr(dbl_ShowDP) + Chr(13) + Chr(10) + "dbl_ShowDPSupplyPressure= " + CStr(dbl_ShowDPSupplyPressure) + Chr(13) + Chr(10) _
+ "dbl_ShowMaxDP= " + CStr(dbl_ShowMaxDP) + Chr(13) + Chr(10) + "dbl_ShowMaxDPSupplyPressure= " + CStr(dbl_ShowMaxDPSupplyPressure) + Chr(13) + Chr(10) + "str_Message= " + str_Message
#End If

'***************************
'Declarations
'***************************
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_AllowANSIPressure As Double
Dim dbl_DynamicFactor As Double
Dim dbl_PlugDiameter As Double
Dim dbl_SeatDiameter As Double
Dim dbl_StemDiameter As Double
Dim dbl_PackingHeight As Double
Dim dbl_AllowPlugDP As Double
Dim dbl_ForcePacking As Double
Dim dbl_ForceSeating As Double
Dim dbl_ForceUnbalancedClosed As Double
Dim dbl_ForceUnbalancedOpen As Double
Dim dbl_ReqForceClosed As Double
Dim dbl_ReqForceOpen As Double
Dim dbl_AreaInitial As Double
Dim dbl_AreaFinal As Double
Dim dbl_MaxSupplyPressure As Double
Dim dbl_ActuatorFriction As Double
Dim dbl_PositionerLossCoef As Double
Dim dbl_PInitial As Double
Dim dbl_PFinal As Double
Dim dbl_SpringRate As Double
Dim dbl_PSupply As Double
Dim dbl_ActuatorForce As Double
Dim dbl_AllowDPClosed As Double
Dim dbl_StemStressArea As Double
Dim dbl_Sy As Double
Dim dbl_E As Double
Dim dbl_Sd As Double
Dim dbl_L As Double
Dim dbl_AllowableCompressiveForce As Double
Dim dbl_AllowableTensileForce As Double
Dim dbl_AllowableStemLoad As Double
Dim dbl_AllowDPStability  As Double
Dim dbl_AllowDPOpen As Double
Dim dbl_MaxAllowDP As Double
Dim dbl_PSupplyDP As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim str_StemMaterial As String
Dim dbl_ReqANSIForceClosed  As Double
Dim dbl_ReqANSIForceOpen As Double
Dim dbl_ReqPlugLimitForceClosed As Double
Dim dbl_ReqPlugLimitForceOpen As Double
Dim dbl_ReqShutoffForceClosed As Double
Dim dbl_ReqShutoffForceOpen As Double
Dim dbl_MaxAllowSupply As Double
Dim dbl_AllowableStemDP As Double
Dim dbl_PSupplyANSI As Double
Dim dbl_PSupplyStemAllow As Double
Dim dbl_PSupplyPlugAllow As Double
Dim dbl_PSupplyShutoffAllow As Double
Dim dbl_MaxPressureDrop As Double
Dim dbl_ShowMaxPressureDrop As Double

Dim dbl_InletPressure(4) As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double
Dim int_QtySeals As Integer
Dim dbl_StemSealWidth As Double
Dim dbl_PistonRingWidth As Double
Dim dbl_VolOver As Double
Dim dbl_ForceSpringClosed As Double
Dim dbl_ActuatorStemDiameter As Double
Dim dbl_LeverLength As Double

Dim str_NotUsed As String
Dim long_NotUsed As Long


'***************************
'Parameter validation
'***************************
'default for now
If (str_PackingMaterial = "") Then str_PackingMaterial = "PTFE"

If (str_PressureUM = "") Then GoTo CalcDP_21000_2Stage_FTO_ATC_Error
If (str_TemperatureUM = "") Then GoTo CalcDP_21000_2Stage_FTO_ATC_Error
If (str_CalibPressureUM = "") Then GoTo CalcDP_21000_2Stage_FTO_ATC_Error
If (str_PackingMaterial = "") Then GoTo CalcDP_21000_2Stage_FTO_ATC_Error
If (str_LeakageClass = "") Then GoTo CalcDP_21000_2Stage_FTO_ATC_Error
If (str_SpringRange = "") Then GoTo CalcDP_21000_2Stage_FTO_ATC_Error

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4

'determine requirements
Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature(), dbl_MaxInletTemperature)
Call CalcTestPressureDrop(dbl_InletPressure(), dbl_OutletPressure(), dbl_MaxPressureDrop)

    dbl_ShowDP = dbl_MaxInletPressure
    dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, str_PressureUM)
    dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
    dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, str_TemperatureUM)
    dbl_ShowMaxPressureDrop = dbl_MaxPressureDrop
    dbl_MaxPressureDrop = ConvertPressureToPSIG(dbl_MaxPressureDrop, str_PressureUM)
    dbl_ShowShutoffDP = dbl_ShutOffDP
    If (dbl_ShutOffDP > 0) Then dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, str_PressureUM)

'ANSI pressure
    Call GetANSIPressure(str_BodyMaterial, dbl_MaxInletTemperature, str_Rating, _
    dbl_AllowANSIPressure)

'End If


'read valve properties
    Call Get21000Data(long_TrimPackageID, dbl_DynamicFactor, _
        dbl_PlugDiameter, dbl_SeatDiameter, dbl_StemDiameter, dbl_PackingHeight, _
        dbl_AllowPlugDP, dbl_StemStressArea, dbl_Travel, str_StemMaterial)
    CnvrtLengthToMM dbl_Travel, "in", dbl_Travel
    CnvrtLengthFromMM dbl_Travel, str_LengthUM, dbl_Travel

    
'calculate required packing force
Call CalcPackingForce(dbl_MaxInletPressure, dbl_StemDiameter, str_PackingMaterial, _
    dbl_PackingHeight, dbl_ForcePacking)
    
'calculate required seating force
Call CalcSeatingForce(str_LeakageClass, dbl_SeatDiameter, dbl_ForceSeating)


'calculate required offbalance force
Call CalcDoubleStageForceFTO(dbl_MaxInletPressure, dbl_SeatDiameter, dbl_PlugDiameter, _
    dbl_ForceUnbalancedClosed, dbl_ForceUnbalancedOpen)
    dbl_ReqForceClosed = dbl_ForcePacking + dbl_ForceSeating + dbl_ForceUnbalancedClosed
    dbl_ReqForceOpen = dbl_ForcePacking + dbl_ForceUnbalancedOpen

Call CalcDoubleStageForceFTO(dbl_AllowANSIPressure, dbl_SeatDiameter, dbl_PlugDiameter, _
    dbl_ForceUnbalancedClosed, dbl_ForceUnbalancedOpen)
    dbl_ReqANSIForceClosed = dbl_ForceUnbalancedClosed
    dbl_ReqANSIForceOpen = dbl_ForceUnbalancedOpen

Call CalcDoubleStageForceFTO(dbl_AllowPlugDP, dbl_SeatDiameter, dbl_PlugDiameter, _
    dbl_ForceUnbalancedClosed, dbl_ForceUnbalancedOpen)
    dbl_ReqPlugLimitForceClosed = dbl_ForceUnbalancedClosed
    dbl_ReqPlugLimitForceOpen = dbl_ForceUnbalancedOpen

Call CalcDoubleStageForceFTO(dbl_ShutOffDP, dbl_SeatDiameter, dbl_PlugDiameter, _
    dbl_ForceUnbalancedClosed, dbl_ForceUnbalancedOpen)
    dbl_ReqShutoffForceClosed = dbl_ForcePacking + dbl_ForceSeating + dbl_ForceUnbalancedClosed
    dbl_ReqShutoffForceOpen = dbl_ForceUnbalancedOpen

'calculate stem limitation
'Call GetStemStressArea(dbl_StemDiameter, dbl_StemStressArea) now returned by get21000data
    
    Call GetStemProperties(str_StemMaterial, dbl_MaxInletTemperature, dbl_Sy, _
        dbl_E, dbl_Sd)

dbl_L = 6# 'default value


'*************************revision jh 8/25/98******************************************

Call CalcAllowStemForce(dbl_Sy, dbl_E, dbl_StemDiameter, dbl_L, _
    dbl_Sd, dbl_StemStressArea, dbl_AllowableStemLoad)
'Call CalcAllowStemTensileForce(dbl_Sd, dbl_stemStressArea, dbl_AllowableTensileForce)
'
'If (dbl_AllowableTensileForce < dbl_AllowableCompressiveForce) Then
'    dbl_AllowableStemLoad = dbl_AllowableTensileForce
'Else
'    dbl_AllowableStemLoad = dbl_AllowableCompressiveForce
'End If

'**************************************************************************************
'calculate required supply pressure
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    Call GetActuatorData(long_ActuatorPartID, dbl_AreaInitial, dbl_AreaFinal, _
        dbl_MaxSupplyPressure, dbl_ActuatorFriction, dbl_PositionerLossCoef, _
        str_NotUsed, long_NotUsed, _
        int_QtySeals, dbl_StemSealWidth, dbl_PistonRingWidth, _
        dbl_VolOver, dbl_ForceSpringClosed, dbl_ActuatorStemDiameter, dbl_LeverLength)

Call CalcSpringRate(str_SpringRange, dbl_AreaInitial, dbl_PInitial, _
    dbl_AreaFinal, dbl_PFinal, str_CalibPressureUM, dbl_Travel, str_LengthUM, _
    dbl_SpringRate)
        
dbl_MaxAllowSupply = dbl_MaxAllowSupplyPressure

Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, dbl_ReqForceClosed, _
    dbl_PSupply)
dbl_PSupplyDP = dbl_PSupply
If (dbl_PSupply < dbl_MaxAllowSupply) Then dbl_MaxAllowSupply = dbl_PSupply

Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, dbl_ReqANSIForceClosed, _
    dbl_PSupply)
dbl_PSupplyANSI = dbl_PSupply
If (dbl_PSupply < dbl_MaxAllowSupply) Then dbl_MaxAllowSupply = dbl_PSupply

Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, dbl_AllowableStemLoad, _
    dbl_PSupply)
dbl_PSupplyStemAllow = dbl_PSupply
If (dbl_PSupply < dbl_MaxAllowSupply) Then dbl_MaxAllowSupply = dbl_PSupply

Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, dbl_ReqPlugLimitForceClosed, _
    dbl_PSupply)
dbl_PSupplyPlugAllow = dbl_PSupply
If (dbl_PSupply < dbl_MaxAllowSupply) Then dbl_MaxAllowSupply = dbl_PSupply

Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, dbl_ReqShutoffForceClosed, _
    dbl_PSupply)
dbl_PSupplyShutoffAllow = dbl_PSupply
If (dbl_PSupply < dbl_MaxAllowSupply) Then dbl_MaxAllowSupply = dbl_PSupply

'calculate max DP with max allow supply

Call CalcATCActuatorForce(dbl_AreaFinal, dbl_PFinal, dbl_MaxAllowSupply, _
    dbl_ActuatorFriction, dbl_ActuatorForce)

Call CalcAllowDPDoubleStageFTO(dbl_ActuatorForce, dbl_SpringRate, dbl_ForcePacking, _
    dbl_ForceSeating, dbl_SeatDiameter, dbl_PlugDiameter, dbl_DynamicFactor, _
    dbl_AllowDPStability, dbl_AllowDPClosed, dbl_AllowDPOpen)

Call CalcMaxDPAllowable(dbl_AllowDPStability, dbl_AllowDPClosed, dbl_AllowDPOpen, _
    dbl_AllowANSIPressure, dbl_MaxAllowDP)
    
'calculate allowable pressure drop for stem

Call CalcAllowDPDoubleStageFTO(dbl_AllowableStemLoad, dbl_SpringRate, dbl_ForcePacking, _
    dbl_ForceSeating, dbl_SeatDiameter, dbl_PlugDiameter, dbl_DynamicFactor, _
    dbl_AllowDPStability, dbl_AllowDPClosed, dbl_AllowDPOpen)
    
Call CalcMaxDPAllowable(dbl_AllowDPStability, dbl_AllowDPClosed, dbl_AllowDPOpen, _
    dbl_AllowANSIPressure, dbl_AllowableStemDP)

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'If debug_on Then
'
'    form21000.txtAllowStemLoad.Text = dbl_AllowableStemLoad
'    form21000.txtAllowANSIpressure = dbl_AllowANSIPressure
'    form21000.txtPackingForce.Text = Format(dbl_ForcePacking, "#####")
'    form21000.txtSeatingForce = Format(dbl_ForceSeating, "#####")
'
'    form21000.txtUnbalForceClosed.Text = Format(dbl_ForceUnbalancedClosed, "#####")
'    form21000.txtUnbalForceOpen.Text = Format(dbl_ForceUnbalancedOpen, "#####")
'    form21000.txtReqForceClosed.Text = Format(dbl_ReqForceClosed, "#####")
'    form21000.txtReqForceOpen.Text = Format(dbl_ReqForceOpen, "#####")
'
'    form21000.txtActuatorForce.Text = dbl_ActuatorForce
'    form21000.txtPsupply.Text = dbl_PSupply
'
'End If
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

'evaluation
str_Message = ""

dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, str_PressureUM)
dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_MaxAllowSupply, str_CalibPressureUM)
dbl_ShowDP = ConvertPressureFromPSIG(dbl_MaxInletPressure, str_PressureUM)
dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_CalibPressureUM)


If (dbl_MaxInletPressure > dbl_AllowANSIPressure) Then
    str_Message = ANSIMESSAGE
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowANSIPressure, str_PressureUM)
    GoTo EndOfRoutine
End If
If (dbl_AllowPlugDP > 0 And dbl_MaxInletPressure > dbl_AllowPlugDP) Then
    str_Message = TRIMMESSAGE
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowPlugDP, str_PressureUM)
    GoTo EndOfRoutine
End If

If (dbl_ActuatorForce > dbl_AllowableStemLoad) Then
    str_Message = STEMMESSAGE
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowableStemDP, str_PressureUM)
    GoTo EndOfRoutine
End If

If (dbl_MaxInletPressure > dbl_AllowDPClosed) Then
    str_Message = ACTUATORMESSAGE + " - CLOSED"
    dbl_ShowDPSupplyPressure = 0
End If

If (dbl_MaxPressureDrop > dbl_AllowDPOpen) Then
    str_Message = ACTUATORMESSAGE + " - OPEN" + Chr(13) + "PRESSURE DROP EXCEEDS ALLOWABLE LIMIT"
    dbl_ShowDPSupplyPressure = 0
End If

If (dbl_MaxPressureDrop > dbl_AllowDPStability) Then
    str_Message = ACTUATORMESSAGE + " - STABILITY"
    dbl_ShowDPSupplyPressure = 0
End If

If (dbl_MaxInletPressure > dbl_ShutOffDP And dbl_ShutOffDP <= dbl_AllowDPClosed And dbl_ShutOffDP > 0# _
        And dbl_MaxPressureDrop < dbl_AllowDPOpen And dbl_MaxInletPressure > dbl_AllowDPClosed) Then
    str_Message = LIMITEDDPMESSAGE
    dbl_ShowDP = dbl_ShowShutoffDP
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupply, str_CalibPressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, str_PressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_ShowMaxDPSupplyPressure, str_CalibPressureUM)

End If

If (dbl_PSupplyDP > dbl_MaxAllowSupply And dbl_PSupplyShutoffAllow > dbl_MaxAllowSupply) Then
    str_Message = SUPPLYMESSAGE
End If

EndOfRoutine:
dbl_ShowDP = TrimDecimalPlaces(dbl_ShowDP)
dbl_ShowDPSupplyPressure = TrimDecimalPlaces(dbl_ShowDPSupplyPressure)
dbl_ShowMaxDP = TrimDecimalPlaces(dbl_ShowMaxDP)
dbl_ShowMaxDPSupplyPressure = TrimDecimalPlaces(dbl_ShowMaxDPSupplyPressure)

#If (DebugVersion = "yes") Then
MsgBox "long_TrimPackageID= " + CStr(long_TrimPackageID) + Chr(13) + Chr(10) + "long_ActuatorPartID= " + CStr(long_ActuatorPartID) + Chr(13) + Chr(10) + "dbl_InletPressure1= " + CStr(dbl_InletPressure1) + Chr(13) + Chr(10) + _
"dbl_InletPressure2= " + CStr(dbl_InletPressure2) + Chr(13) + Chr(10) + "dbl_InletPressure3= " + CStr(dbl_InletPressure3) + Chr(13) + Chr(10) + "dbl_InletPressure4= " + CStr(dbl_InletPressure4) + Chr(13) + Chr(10) + _
"dbl_OutletPressure1= " + CStr(dbl_OutletPressure1) + Chr(13) + Chr(10) + "dbl_OutletPressure2= " + CStr(dbl_OutletPressure2) + Chr(13) + Chr(10) + "dbl_OutletPressure3= " + CStr(dbl_OutletPressure3) + Chr(13) + Chr(10) + _
"dbl_OutletPressure4= " + CStr(dbl_OutletPressure4) + Chr(13) + Chr(10) + "dbl_InletTemperature1= " + CStr(dbl_InletTemperature1) + Chr(13) + Chr(10) + "dbl_InletTemperature2= " + CStr(dbl_InletTemperature2) + Chr(13) + Chr(10) + _
"dbl_InletTemperature3= " + CStr(dbl_InletTemperature3) + Chr(13) + Chr(10) + "dbl_InletTemperature4= " + CStr(dbl_InletTemperature4) + Chr(13) + Chr(10) + "dbl_ShutoffDP= " + CStr(dbl_ShutOffDP) + Chr(13) + Chr(10) + "str_PressureUM= " + str_PressureUM + _
Chr(13) + Chr(10) + "str_TemperatureUM= " + str_TemperatureUM + Chr(13) + Chr(10) + "str_CalibPressureUM= " + str_CalibPressureUM + Chr(13) + Chr(10) + "str_BodyMaterial= " + str_BodyMaterial + Chr(13) + Chr(10) + "str_Rating= " + str_Rating + _
Chr(13) + Chr(10) + "str_PackingMaterial= " + str_PackingMaterial + Chr(13) + Chr(10) + "str_LeakageClass= " + str_LeakageClass + Chr(13) + Chr(10) + "str_SpringRange= " + str_SpringRange + Chr(13) + Chr(10) + "dbl_Travel= " + _
CStr(dbl_Travel) + Chr(13) + Chr(10) + "str_LengthUM= " + str_LengthUM + Chr(13) + Chr(10) + "dbl_ShowDP= " + CStr(dbl_ShowDP) + Chr(13) + Chr(10) + "dbl_ShowDPSupplyPressure= " + CStr(dbl_ShowDPSupplyPressure) + Chr(13) + Chr(10) _
+ "dbl_ShowMaxDP= " + CStr(dbl_ShowMaxDP) + Chr(13) + Chr(10) + "dbl_ShowMaxDPSupplyPressure= " + CStr(dbl_ShowMaxDPSupplyPressure) + Chr(13) + Chr(10) + "str_Message= " + str_Message
#End If

Exit Sub

CalcDP_21000_2Stage_FTO_ATC_Error:
'Placeholder.  How do we want to handle error messages?

End Sub




Public Sub CalcDP_35000_1(long_TrimPackageID As Long, _
    long_ActuatorPartID As Long, dbl_ValveSize As Double, _
    dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, _
    dbl_InletPressure3 As Double, dbl_InletPressure4 As Double, _
    dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, _
    dbl_InletTemperature1 As Double, dbl_InletTemperature2 As Double, _
    dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, _
    dbl_ShutOffDP As Double, str_PressureUM As String, _
    str_TemperatureUM As String, str_CalibPressureUM As String, str_BodyMaterial As String, _
    str_Rating As String, str_PackingMaterial As String, _
    str_LeakageClass As String, str_SpringRange As String, str_FluidType As String, _
    dbl_Travel As Double, str_LengthUM As String, _
    dbl_ShowDP As Double, dbl_ShowDPSupplyPressure As Double, dbl_ShowMaxDP As Double, _
    dbl_ShowMaxDPSupplyPressure As Double, str_Message As String)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for rotary valves
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the selected valve and actuator shutoff capability and compare the result to the defined requirements.
'
'This subroutine is specific to  35000 series valves
'                     model 35 actuator     1"-4" valves
'                6"-16" valves ATO/FTC, ATC/FTO
'                     model 70 actuator     all
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID
'long_ActuatorPartID
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff pressure drop
'str_PressureUM                 unit of measure associated with inlet and outlet pressures
'str_TemperatureUM              unit of measure associated with inlet temperature
'str_CalibPressureUM            unit of measure associated with spring range and supply pressure
'str_LengthUM                   unit of measure for length (Travel)
'str_BodyMaterial
'str_Rating
'str_PackingMaterial            selected packing material
'str_LeakageClass               selected leakage class
'str_SpringRange                selected actuator spring range
'str_FluidType
'dbl_Travel                     Actuator travel
'
'Output Variables:
'dbl_ShowDP                     specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure       calculated required supply pressure for specified
'dbl_ShowMaxDP                  calculated maximum pressure drop for selected valve and actuator combination
'dbl_ShowMaxDPsupplyPressure    calculated required supply pressure for maximum pressure drop
'str_Message                    warning message if pressure drop requirement exceeds equipment capability
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/22 Larry Schoonover - not tested
'
'Calculation/evaluation logic corrections made 8/24/98 J.Haberlin
'   two subroutines added
'       calcReqSupplyPressureCamflex
'       calcAllowDPCamflexFromTorque
'   test code remains as comment text
'-----------------------------------------------------------------------------------------

On Error GoTo CalcDP_35000_1_Error

'***************************
'Declarations
'***************************
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_AllowANSIPressure As Double
Dim long_AirActionID As Long
Dim dbl_shaftDiameter As Double
Dim dbl_Eccentricity As Double
Dim dbl_LeverLength As Double
Dim dbl_MaxAllowValvePressure As Double
Dim dbl_MaxAllowTrimTorque As Double
Dim dbl_SeatRingDiameter As Double
Dim str_ShaftMaterial As String
Dim dbl_AreaInitial As Double
Dim dbl_AreaFinal As Double
Dim dbl_MaxSupplyPressure As Double
Dim dbl_ActuatorFriction As Double
Dim dbl_PositionerLossCoef As Double
Dim str_Model As String
Dim dbl_AllowShaftStress As Double
Dim str_ActuatorModel As String
Dim dbl_ServiceFactor As Double
Dim dbl_SeatingLoadCoefficient As Double
Dim dbl_PInitial As Double
Dim dbl_PFinal As Double
Dim dbl_SpringRate As Double
Dim dbl_OffBalTorque As Double
Dim dbl_Mu As Double
Dim dbl_FrictionTorque As Double
Dim dbl_SeatingTorque As Double
Dim dbl_RequiredTorque As Double
Dim dbl_PSupplyDP As Double
Dim dbl_AllowShaftTorque As Double
Dim dbl_MaxAllowShaftSupplyPressure As Double
Dim dbl_MaxAllowSupplyPressure As Double
Dim dbl_MaxAllowDP As Double
Dim dbl_RequiredSupplyPressure As Double
Dim dbl_packingCoefficient As Double
Dim dbl_AllowPlugDP As Double

Dim dbl_InletPressure(4) As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double
Dim int_QtySeals As Integer
Dim dbl_StemSealWidth As Double
Dim dbl_PistonRingWidth As Double
Dim dbl_VolOver As Double
Dim dbl_ForceSpringClosed As Double
Dim dbl_ActuatorStemDiameter As Double

#If (DebugVersion = "yes") Then
MsgBox "35000-1 Actuator" + Chr(10) + Chr(13) + _
     "long_TrimPackageI=D" + CStr(long_TrimPackageID) + Chr(10) + Chr(13) + _
     "long_ActuatorPartID=" + CStr(long_ActuatorPartID) + Chr(10) + Chr(13) + _
     "dbl_ValveSize=" + CStr(dbl_ValveSize) + Chr(10) + Chr(13) + _
     "dbl_InletPressure1=" + CStr(dbl_InletPressure1) + Chr(10) + Chr(13) + _
     "dbl_OutletPressure1=" + CStr(dbl_OutletPressure1) + Chr(10) + Chr(13) + _
     "dbl_InletTemperature1=" + CStr(dbl_InletTemperature1) + Chr(10) + Chr(13) + _
     "dbl_ShutoffDP=" + CStr(dbl_ShutOffDP) + Chr(10) + Chr(13) + _
     "str_PressureUM=" + str_PressureUM + Chr(10) + Chr(13) + _
     "str_TemperatureUM=" + str_TemperatureUM + Chr(10) + Chr(13) + _
     "str_CalibPressureUM=" + CStr(str_CalibPressureUM) + Chr(10) + Chr(13) + _
     "str_BodyMaterial=" + CStr(str_BodyMaterial) + Chr(10) + Chr(13) + _
     "str_Rating=" + str_Rating + Chr(10) + Chr(13) + _
     "str_PackingMaterial=" + str_PackingMaterial + Chr(10) + Chr(13) + _
     "str_LeakageClass=" + str_LeakageClass + Chr(10) + Chr(13) + _
     "str_SpringRange=" + str_SpringRange + Chr(10) + Chr(13) + _
     "str_FluidType=" + str_FluidType + Chr(10) + Chr(13) + _
     "dbl_Travel=" + CStr(dbl_Travel) + Chr(10) + Chr(13) + _
     "str_LengthUM=" + str_LengthUM + Chr(10) + Chr(13)
     
#End If

'***************************
'Parameter validation
'***************************
If (dbl_ShutOffDP < 0#) Then GoTo CalcDP_35000_1_Error
If (str_PressureUM = "") Then GoTo CalcDP_35000_1_Error
If (str_TemperatureUM = "") Then GoTo CalcDP_35000_1_Error
If (str_CalibPressureUM = "") Then GoTo CalcDP_35000_1_Error
If (str_PackingMaterial = "") Then GoTo CalcDP_35000_1_Error
If (str_LeakageClass = "") Then GoTo CalcDP_35000_1_Error
If (str_SpringRange = "") Then GoTo CalcDP_35000_1_Error

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4
'determine requirements
Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature, dbl_MaxInletTemperature)
dbl_ShowDP = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, str_PressureUM)
dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, str_TemperatureUM)
dbl_ShowShutoffDP = dbl_ShutOffDP
If (dbl_ShutOffDP > 0) Then dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, str_PressureUM)

'determine ANSI limitation
    Call GetANSIPressure(str_BodyMaterial, dbl_MaxInletTemperature, str_Rating, _
    dbl_AllowANSIPressure)

'get product data
    Call Get35000Data(long_TrimPackageID, dbl_shaftDiameter, dbl_Eccentricity, _
        dbl_AllowPlugDP, dbl_AllowShaftTorque, dbl_SeatRingDiameter, _
        str_ShaftMaterial)
    Call GetPackingCoef(str_PackingMaterial, dbl_packingCoefficient)

    Call GetActuatorData(long_ActuatorPartID, dbl_AreaInitial, _
        dbl_AreaFinal, dbl_MaxSupplyPressure, dbl_ActuatorFriction, _
        dbl_PositionerLossCoef, str_ActuatorModel, long_AirActionID, _
        int_QtySeals, dbl_StemSealWidth, dbl_PistonRingWidth, _
        dbl_VolOver, dbl_ForceSpringClosed, dbl_ActuatorStemDiameter, dbl_LeverLength)
    Call GetServiceFactor(dbl_ValveSize, long_AirActionID, str_FluidType, _
        str_ActuatorModel, dbl_ServiceFactor)
    
   ' restrict seating load for Camflex
    Call GetSeatingCoef(str_LeakageClass, dbl_SeatingLoadCoefficient)
'        historically the seating load for a camflex with class IV leakage is 0
'        other product lines use 30 psi/in of seat circumference.  For class VI
'        camflex uses the same seating load as other valve lines.  Therefor
'        the following compensation is made:
        If (str_LeakageClass = "IV") Then dbl_SeatingLoadCoefficient = 0

'-----------------------------------------------------------------------------------------
'calculate required torque

Call CalcSpringRate(str_SpringRange, dbl_AreaInitial, dbl_PInitial, dbl_AreaFinal, _
    dbl_PFinal, str_CalibPressureUM, 1, str_LengthUM, dbl_SpringRate)
Call CalcOffBalanceTorque(dbl_MaxInletPressure, dbl_SeatRingDiameter, _
    dbl_Eccentricity, dbl_OffBalTorque)
Call CalcFrictionTorque(dbl_MaxInletPressure, dbl_SeatRingDiameter, dbl_shaftDiameter, _
    dbl_packingCoefficient, dbl_FrictionTorque)
Call CalcSeatingTorque(dbl_AreaInitial, dbl_LeverLength, dbl_shaftDiameter, _
    dbl_SeatingLoadCoefficient, dbl_Eccentricity, dbl_SeatingTorque)
Call CalcRequiredTorque(dbl_OffBalTorque, dbl_SeatingTorque, dbl_FrictionTorque, _
    dbl_RequiredTorque)

Call CalcReqSupplyPressureCamflex(dbl_MaxInletPressure, dbl_Eccentricity, dbl_shaftDiameter, dbl_AreaFinal, dbl_SeatRingDiameter, _
    dbl_LeverLength, dbl_PFinal, dbl_RequiredSupplyPressure)
 
'calculate actuator / valve capability

'Call Get35AllowShaftTorque(dbl_ValveSize, str_ShaftMaterial, dbl_AllowShaftTorque)
'If dbl_MaxAllowTrimTorque < dbl_AllowShaftTorque Then dbl_AllowShaftTorque = dbl_MaxAllowTrimTorque
Call CalcMaxAllowSupplyPressureCamflex(dbl_AllowShaftTorque, dbl_ServiceFactor, _
    dbl_AreaFinal, dbl_LeverLength, dbl_PFinal, dbl_MaxAllowShaftSupplyPressure)
Call calcAllowDPCamflexFromTorque(dbl_AllowShaftTorque, dbl_PFinal, dbl_SeatRingDiameter, dbl_Eccentricity, _
    dbl_shaftDiameter, dbl_MaxAllowDP)
Call CalcRequiredSupplyPressureCamflex(dbl_RequiredTorque, dbl_AreaFinal, _
    dbl_LeverLength, dbl_PFinal, dbl_RequiredSupplyPressure)


If (dbl_MaxAllowShaftSupplyPressure < dbl_MaxSupplyPressure) Then
    dbl_MaxAllowSupplyPressure = dbl_MaxAllowShaftSupplyPressure
Else
    dbl_MaxAllowSupplyPressure = dbl_MaxSupplyPressure
End If

Call CalcAllowDPCamflex(dbl_AreaFinal, dbl_LeverLength, dbl_PFinal, dbl_SeatRingDiameter, _
    dbl_Eccentricity, dbl_MaxAllowSupplyPressure, dbl_shaftDiameter, dbl_MaxAllowDP)

If dbl_MaxAllowDP > dbl_AllowANSIPressure Then dbl_MaxAllowDP = dbl_AllowANSIPressure
If dbl_MaxAllowDP > dbl_AllowPlugDP Then dbl_MaxAllowDP = dbl_AllowPlugDP

'evaluation
If (dbl_MaxInletPressure > dbl_AllowANSIPressure) Then str_Message = ANSIMESSAGE
If (dbl_MaxInletPressure > dbl_AllowPlugDP) Then str_Message = TRIMMESSAGE


dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, str_PressureUM)
dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_MaxAllowSupplyPressure, _
        str_CalibPressureUM)


If (dbl_MaxAllowDP < dbl_MaxInletPressure) Then str_Message = ACTUATORMESSAGE

If (dbl_MaxInletPressure < dbl_ShutOffDP And dbl_ShutOffDP > 0#) Then
    If (dbl_ShutOffDP <= dbl_MaxAllowDP) Then
        str_Message = LIMITEDDPMESSAGE
    Else
        str_Message = ACTUATORMESSAGE
    End If
End If

If (dbl_MaxInletPressure <= dbl_MaxAllowDP) Then
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_RequiredSupplyPressure, str_CalibPressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, str_PressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_MaxAllowSupplyPressure, _
        str_CalibPressureUM)
End If

If (dbl_MaxInletPressure < dbl_ShutOffDP And dbl_ShutOffDP <= dbl_MaxAllowDP And dbl_ShutOffDP > 0#) Then
    dbl_ShowDP = dbl_ShowShutoffDP
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_RequiredSupplyPressure, str_CalibPressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, str_PressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_MaxAllowSupplyPressure, _
        str_CalibPressureUM)
End If

dbl_ShowDP = TrimDecimalPlaces(dbl_ShowDP)
dbl_ShowDPSupplyPressure = TrimDecimalPlaces(dbl_ShowDPSupplyPressure)
dbl_ShowMaxDP = TrimDecimalPlaces(dbl_ShowMaxDP)
dbl_ShowMaxDPSupplyPressure = TrimDecimalPlaces(dbl_ShowMaxDPSupplyPressure)

#If (DebugVersion = "yes") Then
    MsgBox "35000-1 Actuator" + Chr(10) + Chr(13) + _
    "dbl_ShowDP=" + CStr(dbl_ShowDP) + Chr(10) + Chr(13) + _
    "dbl_ShowDPSupplyPressure=" + CStr(dbl_ShowDPSupplyPressure) + Chr(10) + Chr(13) + _
    "dbl_ShowMaxDP=" + CStr(dbl_ShowMaxDP) + Chr(10) + Chr(13) + _
    "dbl_ShowMaxDPSupplyPressure=" + CStr(dbl_ShowMaxDPSupplyPressure) + Chr(10) + Chr(13) + _
    "str_Message=" + str_Message + Chr(10) + Chr(13)

#End If

Exit Sub

CalcDP_35000_1_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcDP_35000_2(long_TrimPackageID As Long, _
    long_ActuatorPartID As Long, dbl_ValveSize As Double, _
    dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, _
    dbl_InletPressure3 As Double, dbl_InletPressure4 As Double, _
    dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, _
    dbl_InletTemperature1 As Double, dbl_InletTemperature2 As Double, _
    dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, _
    dbl_ShutOffDP As Double, str_PressureUM As String, _
    str_TemperatureUM As String, str_CalibPressureUM As String, str_BodyMaterial As String, _
    str_Rating As String, str_PackingMaterial As String, _
    str_LeakageClass As String, str_SpringRange As String, str_FluidType As String, _
    dbl_Travel As Double, str_LengthUM As String, _
    dbl_ShowDP As Double, dbl_ShowDPSupplyPressure As Double, dbl_ShowMaxDP As Double, _
    dbl_ShowMaxDPSupplyPressure As Double, str_Message As String)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for rotary valves
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the selected valve and actuator shutoff capability and compare the result to the defined requirements.
'
'This subroutine is specific to  35000 series valves
'                     model 35 actuator    6"-16" valves ATC/FTC
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID
'long_ActuatorPartID
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff pressure drop
'str_PressureUM                 unit of measure associated with inlet and outlet pressures
'str_TemperatureUM              unit of measure associated with inlet temperature
'str_CalibPressureUM            unit of measure associated with spring range and supply pressure
'str_LengthUM                   unit of measure for length (Travel)
'str_BodyMaterial
'str_Rating
'str_PackingMaterial            selected packing material
'str_LeakageClass               selected leakage class
'str_SpringRange                selected actuator spring range
'str_FluidType
'dbl_Travel                     Actuator travel
'
'Output Variables:
'dbl_ShowDP                     specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure       calculated required supply pressure for specified
'dbl_ShowMaxDP                  calculated maximum pressure drop for selected valve and actuator combination
'dbl_ShowMaxDPsupplyPressure    calculated required supply pressure for maximum pressure drop
'str_Message                    warning message if pressure drop requirement exceeds equipment capability
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/22 Larry Schoonover - not tested
'
'Calculation/evaluation logic corrections made 8/24/98 J.Haberlin
'   test code remains as comment text
'-----------------------------------------------------------------------------------------

On Error GoTo CalcDP_35000_2_Error

'***************************
'Declarations
'***************************
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_AllowANSIPressure As Double
Dim long_AirActionID As Long
Dim dbl_shaftDiameter As Double
Dim dbl_Eccentricity As Double
Dim dbl_LeverLength As Double
Dim dbl_MaxAllowValvePressure As Double
Dim dbl_MaxAllowTrimTorque As Double
Dim dbl_SeatRingDiameter As Double
Dim str_ShaftMaterial As String
Dim dbl_AreaInitial As Double
Dim dbl_AreaFinal As Double
Dim dbl_MaxSupplyPressure As Double
Dim dbl_ActuatorFriction As Double
Dim dbl_PositionerLossCoef As Double
Dim str_Model As String
Dim dbl_AllowShaftStress As Double
Dim str_ActuatorModel As String
Dim dbl_ServiceFactor As Double
Dim dbl_SeatingLoadCoefficient As Double
Dim dbl_PInitial As Double
Dim dbl_PFinal As Double
Dim dbl_SpringRate As Double
Dim dbl_OffBalTorque As Double
Dim dbl_Mu As Double
Dim dbl_FrictionTorque As Double
Dim dbl_SeatingTorque As Double
Dim dbl_RequiredTorque As Double
Dim dbl_PSupplyDP As Double
Dim dbl_AllowShaftTorque As Double
Dim dbl_MaxAllowShaftSupplyPressure As Double
Dim dbl_MaxAllowSupplyPressure As Double
Dim dbl_MaxAllowDP As Double
Dim dbl_RequiredSupplyPressure As Double
Dim dbl_AllowDPStability As Double
Dim dbl_AllowDPStatic As Double

Dim dbl_InletPressure(4) As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double
Dim int_QtySeals As Integer
Dim dbl_StemSealWidth As Double
Dim dbl_PistonRingWidth As Double
Dim dbl_VolOver As Double
Dim dbl_ForceSpringClosed As Double
Dim dbl_ActuatorStemDiameter As Double

#If (DebugVersion = "yes") Then
MsgBox "35000-2 Actuator" + Chr(10) + Chr(13) + _
     "long_TrimPackageI=D" + CStr(long_TrimPackageID) + Chr(10) + Chr(13) + _
     "long_ActuatorPartID=" + CStr(long_ActuatorPartID) + Chr(10) + Chr(13) + _
     "dbl_ValveSize=" + CStr(dbl_ValveSize) + Chr(10) + Chr(13) + _
     "dbl_InletPressure1=" + CStr(dbl_InletPressure1) + Chr(10) + Chr(13) + _
     "dbl_OutletPressure1=" + CStr(dbl_OutletPressure1) + Chr(10) + Chr(13) + _
     "dbl_InletTemperature1=" + CStr(dbl_InletTemperature1) + Chr(10) + Chr(13) + _
     "dbl_ShutoffDP=" + CStr(dbl_ShutOffDP) + Chr(10) + Chr(13) + _
     "str_PressureUM=" + str_PressureUM + Chr(10) + Chr(13) + _
     "str_TemperatureUM=" + str_TemperatureUM + Chr(10) + Chr(13) + _
     "str_CalibPressureUM=" + CStr(str_CalibPressureUM) + Chr(10) + Chr(13) + _
     "str_BodyMaterial=" + CStr(str_BodyMaterial) + Chr(10) + Chr(13) + _
     "str_Rating=" + str_Rating + Chr(10) + Chr(13) + _
     "str_PackingMaterial=" + str_PackingMaterial + Chr(10) + Chr(13) + _
     "str_LeakageClass=" + str_LeakageClass + Chr(10) + Chr(13) + _
     "str_SpringRange=" + str_SpringRange + Chr(10) + Chr(13) + _
     "str_FluidType=" + str_FluidType + Chr(10) + Chr(13) + _
     "dbl_Travel=" + CStr(dbl_Travel) + Chr(10) + Chr(13) + _
     "str_LengthUM=" + str_LengthUM + Chr(10) + Chr(13)
     
#End If


'***************************
'Parameter validation
'***************************
If (dbl_ShutOffDP < 0#) Then GoTo CalcDP_35000_2_Error
If (str_PressureUM = "") Then GoTo CalcDP_35000_2_Error
If (str_TemperatureUM = "") Then GoTo CalcDP_35000_2_Error
If (str_CalibPressureUM = "") Then GoTo CalcDP_35000_2_Error
If (str_PackingMaterial = "") Then GoTo CalcDP_35000_2_Error
If (str_LeakageClass = "") Then GoTo CalcDP_35000_2_Error
If (str_SpringRange = "") Then GoTo CalcDP_35000_2_Error

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4

'determine requirements
Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature, dbl_MaxInletTemperature)
dbl_ShowDP = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, str_PressureUM)
dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, str_TemperatureUM)
dbl_ShowShutoffDP = dbl_ShutOffDP
If (dbl_ShutOffDP > 0) Then dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, str_PressureUM)

'determine ANSI limitation

    Call GetANSIPressure(str_BodyMaterial, dbl_MaxInletTemperature, str_Rating, _
    dbl_AllowANSIPressure)

'get product data
    Call Get35000Data(long_TrimPackageID, dbl_shaftDiameter, dbl_Eccentricity, _
        dbl_MaxAllowValvePressure, dbl_AllowShaftTorque, dbl_SeatRingDiameter, _
        str_ShaftMaterial)
    'Call GetPackingCoef(str_packingMaterial, dbl_packingCoefficient)
    
    Call GetActuatorData(long_ActuatorPartID, dbl_AreaInitial, _
        dbl_AreaFinal, dbl_MaxSupplyPressure, dbl_ActuatorFriction, _
        dbl_PositionerLossCoef, str_ActuatorModel, long_AirActionID, _
        int_QtySeals, dbl_StemSealWidth, dbl_PistonRingWidth, _
        dbl_VolOver, dbl_ForceSpringClosed, dbl_ActuatorStemDiameter, dbl_LeverLength)
    
    dbl_AreaFinal = dbl_AreaInitial 'both are equal in camflex - some database entries don't have area final
    dbl_ServiceFactor = 1
    
    
   ' restrict seating load for Camflex
 '  Call GetSeatingCoef(str_LeakageClass, dbl_SeatingLoadCoefficient)
'        historically the seating load for a camflex with class IV leakage is 0
'        other product lines use 30 psi/in of seat circumference.  For class VI
'        camflex uses the same seating load as other valve lines.  Therefor
'        the following compensation is made:
        If (str_LeakageClass = "IV") Then dbl_SeatingLoadCoefficient = 0

'calculate allowable pressure drop
'Call Get35AllowShaftTorque(dbl_ValveSize, str_ShaftMaterial, dbl_AllowShaftTorque)

Call CalcSpringRate(str_SpringRange, dbl_AreaInitial, dbl_PInitial, dbl_AreaFinal, _
    dbl_PFinal, str_CalibPressureUM, 1, str_LengthUM, dbl_SpringRate)
    
Call CalcDPAllowStabilityCamflex(dbl_AreaFinal, dbl_LeverLength, dbl_PInitial, _
    dbl_PFinal, dbl_SeatRingDiameter, dbl_Eccentricity, dbl_shaftDiameter, _
    dbl_AllowDPStability)
    
Call CalcDPAllowStaticCamflex(dbl_AreaInitial, dbl_LeverLength, dbl_PInitial, _
    dbl_SeatRingDiameter, dbl_Eccentricity, dbl_shaftDiameter, dbl_AllowDPStatic)
    
dbl_RequiredTorque = 0

Call CalcRequiredSupplyPressureCamflex(dbl_RequiredTorque, dbl_AreaFinal, _
    dbl_LeverLength, dbl_PFinal, dbl_RequiredSupplyPressure)


Call CalcMaxAllowSupplyPressureCamflex(dbl_AllowShaftTorque, dbl_ServiceFactor, _
    dbl_AreaFinal, dbl_LeverLength, dbl_PFinal, dbl_MaxAllowShaftSupplyPressure)
    
If (dbl_MaxAllowShaftSupplyPressure < dbl_MaxSupplyPressure) Then
    dbl_MaxAllowSupplyPressure = dbl_MaxAllowShaftSupplyPressure
Else
   dbl_MaxAllowSupplyPressure = dbl_MaxSupplyPressure
End If

If (dbl_AllowDPStability < dbl_AllowDPStatic) Then
    dbl_MaxAllowDP = dbl_AllowDPStability
Else
    dbl_MaxAllowDP = dbl_AllowDPStatic
End If


'evaluation
If (dbl_MaxInletPressure > dbl_AllowANSIPressure) Then str_Message = ANSIMESSAGE
'If (dbl_MaxInletPressure > dbl_AllowPlugDP) Then str_Message = TRIMMESSAGE
If (dbl_RequiredTorque > dbl_AllowShaftTorque) Then str_Message = STEMMESSAGE
If (dbl_MaxAllowDP < dbl_MaxInletPressure) Then str_Message = ACTUATORMESSAGE

If (dbl_MaxInletPressure < dbl_ShutOffDP And dbl_ShutOffDP > 0#) Then
    If (dbl_ShutOffDP <= dbl_MaxAllowDP) Then
        str_Message = LIMITEDDPMESSAGE
    Else
        str_Message = ACTUATORMESSAGE
    End If
End If

If (dbl_MaxInletPressure <= dbl_MaxAllowDP) Then
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_RequiredSupplyPressure, _
        str_CalibPressureUM)
    
End If

If (dbl_MaxInletPressure < dbl_ShutOffDP And dbl_ShutOffDP <= dbl_MaxAllowDP And dbl_ShutOffDP > 0#) Then
    dbl_ShowDP = dbl_ShowShutoffDP
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_CalibPressureUM)
    
End If

dbl_MaxAllowSupplyPressure = dbl_PFinal + 5
dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, str_PressureUM)
dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_MaxAllowSupplyPressure, _
        str_CalibPressureUM)

dbl_ShowDP = TrimDecimalPlaces(dbl_ShowDP)
dbl_ShowDPSupplyPressure = TrimDecimalPlaces(dbl_ShowDPSupplyPressure)
dbl_ShowMaxDP = TrimDecimalPlaces(dbl_ShowMaxDP)
dbl_ShowMaxDPSupplyPressure = TrimDecimalPlaces(dbl_ShowMaxDPSupplyPressure)


#If (DebugVersion = "yes") Then
    MsgBox "35000-2 Actuator" + Chr(10) + Chr(13) + _
    "dbl_ShowDP=" + CStr(dbl_ShowDP) + Chr(10) + Chr(13) + _
    "dbl_ShowDPSupplyPressure=" + CStr(dbl_ShowDPSupplyPressure) + Chr(10) + Chr(13) + _
    "dbl_ShowMaxDP=" + CStr(dbl_ShowMaxDP) + Chr(10) + Chr(13) + _
    "dbl_ShowMaxDPSupplyPressure=" + CStr(dbl_ShowMaxDPSupplyPressure) + Chr(10) + Chr(13) + _
    "str_Message=" + str_Message + Chr(10) + Chr(13)

#End If

Exit Sub

CalcDP_35000_2_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcDP_35000_3(long_TrimPackageID As Long, _
    long_ActuatorPartID As Long, dbl_ValveSize As Double, _
    dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, _
    dbl_InletPressure3 As Double, dbl_InletPressure4 As Double, _
    dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, _
    dbl_InletTemperature1 As Double, dbl_InletTemperature2 As Double, _
    dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, _
    dbl_ShutOffDP As Double, str_PressureUM As String, _
    str_TemperatureUM As String, str_CalibPressureUM As String, str_BodyMaterial As String, _
    str_Rating As String, str_PackingMaterial As String, _
    str_LeakageClass As String, str_SpringRange As String, str_FluidType As String, _
    dbl_Travel As Double, str_LengthUM As String, _
    dbl_ShowDP As Double, dbl_ShowDPSupplyPressure As Double, dbl_ShowMaxDP As Double, _
    dbl_ShowMaxDPSupplyPressure As Double, str_Message As String)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for rotary valves
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the selected valve and actuator shutoff capability and compare the result to the defined requirements.
'
'This subroutine is specific to  35000 series valves
'                     model 35 actuator           6"-16" valves ATO/FTO
'                     model 37 HPI actuator
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID
'long_ActuatorPartID
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff pressure drop
'str_PressureUM                 unit of measure associated with inlet and outlet pressures
'str_TemperatureUM              unit of measure associated with inlet temperature
'str_CalibPressureUM            unit of measure associated with spring range and supply pressure
'str_LengthUM                   unit of measure for length (Travel)
'str_BodyMaterial
'str_Rating
'str_PackingMaterial            selected packing material
'str_LeakageClass               selected leakage class
'str_SpringRange                selected actuator spring range
'str_FluidType
'dbl_Travel                     Actuator travel
'
'Output Variables:
'dbl_ShowDP                     specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure       calculated required supply pressure for specified
'dbl_ShowMaxDP                  calculated maximum pressure drop for selected valve and actuator combination
'dbl_ShowMaxDPsupplyPressure    calculated required supply pressure for maximum pressure drop
'str_Message                    warning message if pressure drop requirement exceeds equipment capability
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/22 Larry Schoonover - not tested
'
'Calculation/evaluation logic corrections made 8/24/98 J.Haberlin
'   test code remains as comment text
'-----------------------------------------------------------------------------------------

On Error GoTo CalcDP_35000_3_Error

'***************************
'Declarations
'***************************
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_AllowANSIPressure As Double
Dim long_AirActionID As Long
Dim dbl_shaftDiameter As Double
Dim dbl_Eccentricity As Double
Dim dbl_LeverLength As Double
Dim dbl_MaxAllowValvePressure As Double
Dim dbl_MaxAllowTrimTorque As Double
Dim dbl_SeatRingDiameter As Double
Dim str_ShaftMaterial As String
Dim dbl_AreaInitial As Double
Dim dbl_AreaFinal As Double
Dim dbl_MaxSupplyPressure As Double
Dim dbl_ActuatorFriction As Double
Dim dbl_PositionerLossCoef As Double
Dim str_Model As String
Dim dbl_AllowShaftStress As Double
Dim str_ActuatorModel As String
Dim dbl_ServiceFactor As Double
Dim dbl_SeatingLoadCoefficient As Double
Dim dbl_PInitial As Double
Dim dbl_PFinal As Double
Dim dbl_SpringRate As Double
Dim dbl_OffBalTorque As Double
Dim dbl_Mu As Double
Dim dbl_FrictionTorque As Double
Dim dbl_SeatingTorque As Double
Dim dbl_RequiredTorque As Double
Dim dbl_PSupplyDP As Double
Dim dbl_AllowShaftTorque As Double
Dim dbl_MaxAllowShaftSupplyPressure As Double
Dim dbl_MaxAllowSupplyPressure As Double
Dim dbl_MaxAllowDP As Double
Dim dbl_RequiredSupplyPressure As Double
Dim dbl_AllowDPStability As Double
Dim dbl_AllowDPStatic As Double
Dim dbl_PInitialEffective As Double
Dim dbl_packingCoefficient As Double

Dim dbl_InletPressure(4) As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double
Dim int_QtySeals As Integer
Dim dbl_StemSealWidth As Double
Dim dbl_PistonRingWidth As Double
Dim dbl_VolOver As Double
Dim dbl_ForceSpringClosed As Double
Dim dbl_ActuatorStemDiameter As Double


#If (DebugVersion = "yes") Then
MsgBox "35000-3 Actuator" + Chr(10) + Chr(13) + _
     "long_TrimPackageI=D" + CStr(long_TrimPackageID) + Chr(10) + Chr(13) + _
     "long_ActuatorPartID=" + CStr(long_ActuatorPartID) + Chr(10) + Chr(13) + _
     "dbl_ValveSize=" + CStr(dbl_ValveSize) + Chr(10) + Chr(13) + _
     "dbl_InletPressure1=" + CStr(dbl_InletPressure1) + Chr(10) + Chr(13) + _
     "dbl_OutletPressure1=" + CStr(dbl_OutletPressure1) + Chr(10) + Chr(13) + _
     "dbl_InletTemperature1=" + CStr(dbl_InletTemperature1) + Chr(10) + Chr(13) + _
     "dbl_ShutoffDP=" + CStr(dbl_ShutOffDP) + Chr(10) + Chr(13) + _
     "str_PressureUM=" + str_PressureUM + Chr(10) + Chr(13) + _
     "str_TemperatureUM=" + str_TemperatureUM + Chr(10) + Chr(13) + _
     "str_CalibPressureUM=" + CStr(str_CalibPressureUM) + Chr(10) + Chr(13) + _
     "str_BodyMaterial=" + CStr(str_BodyMaterial) + Chr(10) + Chr(13) + _
     "str_Rating=" + str_Rating + Chr(10) + Chr(13) + _
     "str_PackingMaterial=" + str_PackingMaterial + Chr(10) + Chr(13) + _
     "str_LeakageClass=" + str_LeakageClass + Chr(10) + Chr(13) + _
     "str_SpringRange=" + str_SpringRange + Chr(10) + Chr(13) + _
     "str_FluidType=" + str_FluidType + Chr(10) + Chr(13) + _
     "dbl_Travel=" + CStr(dbl_Travel) + Chr(10) + Chr(13) + _
     "str_LengthUM=" + str_LengthUM + Chr(10) + Chr(13)
     
#End If



'***************************
'Parameter validation
'***************************
If (dbl_ShutOffDP < 0#) Then GoTo CalcDP_35000_3_Error
If (str_PressureUM = "") Then GoTo CalcDP_35000_3_Error
If (str_TemperatureUM = "") Then GoTo CalcDP_35000_3_Error
If (str_CalibPressureUM = "") Then GoTo CalcDP_35000_3_Error
If (str_PackingMaterial = "") Then GoTo CalcDP_35000_3_Error
If (str_LeakageClass = "") Then GoTo CalcDP_35000_3_Error
If (str_SpringRange = "") Then GoTo CalcDP_35000_3_Error

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4

'determine requirements
Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature, dbl_MaxInletTemperature)
dbl_ShowDP = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, str_PressureUM)
dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, str_TemperatureUM)
dbl_ShowShutoffDP = dbl_ShutOffDP
If (dbl_ShutOffDP > 0) Then dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, str_PressureUM)

    Call GetANSIPressure(str_BodyMaterial, dbl_MaxInletTemperature, str_Rating, _
    dbl_AllowANSIPressure)

'get product data
    Call Get35000Data(long_TrimPackageID, dbl_shaftDiameter, dbl_Eccentricity, _
        dbl_MaxAllowValvePressure, dbl_MaxAllowTrimTorque, dbl_SeatRingDiameter, _
        str_ShaftMaterial)
    Call GetPackingCoef(str_PackingMaterial, dbl_packingCoefficient)
    
    Call GetActuatorData(long_ActuatorPartID, dbl_AreaInitial, _
        dbl_AreaFinal, dbl_MaxSupplyPressure, dbl_ActuatorFriction, _
        dbl_PositionerLossCoef, str_ActuatorModel, long_AirActionID, _
        int_QtySeals, dbl_StemSealWidth, dbl_PistonRingWidth, _
        dbl_VolOver, dbl_ForceSpringClosed, dbl_ActuatorStemDiameter, dbl_LeverLength)

    
    'restrict seating load for Camflex
    Call GetSeatingCoef(str_LeakageClass, dbl_SeatingLoadCoefficient)
       ' historically the seating load for a camflex with class IV leakage is 0
       ' other product lines use 30 psi/in of seat circumference.  For class VI
       ' camflex uses the same seating load as other valve lines.  Therefor
       ' the following compensation is made:
        If (str_LeakageClass = "IV") Then dbl_SeatingLoadCoefficient = 0



'calculate allowable pressure drop

Call CalcSpringRate(str_SpringRange, dbl_AreaInitial, dbl_PInitial, dbl_AreaFinal, _
    dbl_PFinal, str_CalibPressureUM, 1, str_LengthUM, dbl_SpringRate)

dbl_PInitialEffective = dbl_PInitial - 1.5
Call CalcDPAllowStaticCamflex(dbl_AreaInitial, dbl_LeverLength, dbl_PInitialEffective, _
    dbl_SeatRingDiameter, dbl_Eccentricity, dbl_shaftDiameter, dbl_MaxAllowDP)
dbl_RequiredTorque = 0

dbl_RequiredSupplyPressure = dbl_PFinal + 5

dbl_ServiceFactor = 1

'evaluation
 If (dbl_MaxInletPressure > dbl_AllowANSIPressure) Then str_Message = ANSIMESSAGE
'If (dbl_MaxInletPressure > dbl_AllowPlugDP) Then str_Message = TRIMMESSAGE
If (dbl_RequiredTorque > dbl_AllowShaftTorque) Then str_Message = STEMMESSAGE
If (dbl_MaxAllowDP < dbl_MaxInletPressure) Then str_Message = ACTUATORMESSAGE

If (dbl_MaxInletPressure < dbl_ShutOffDP) Then
    If (dbl_ShutOffDP <= dbl_MaxAllowDP) Then
        str_Message = LIMITEDDPMESSAGE
    Else
        str_Message = ACTUATORMESSAGE
    End If
End If

If (dbl_MaxInletPressure <= dbl_MaxAllowDP) Then
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_RequiredSupplyPressure, _
        str_CalibPressureUM)
    
End If

If (dbl_MaxInletPressure < dbl_ShutOffDP And dbl_ShutOffDP <= dbl_MaxAllowDP And dbl_ShutOffDP > 0#) Then
    dbl_ShowDP = dbl_ShowShutoffDP
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_CalibPressureUM)
   
End If

dbl_MaxAllowSupplyPressure = dbl_PFinal + 5
dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, str_PressureUM)
dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_MaxAllowSupplyPressure, _
        str_CalibPressureUM)

dbl_ShowDP = TrimDecimalPlaces(dbl_ShowDP)
dbl_ShowDPSupplyPressure = TrimDecimalPlaces(dbl_ShowDPSupplyPressure)
dbl_ShowMaxDP = TrimDecimalPlaces(dbl_ShowMaxDP)
dbl_ShowMaxDPSupplyPressure = TrimDecimalPlaces(dbl_ShowMaxDPSupplyPressure)

#If (DebugVersion = "yes") Then
    MsgBox "35000-3 Actuator" + Chr(10) + Chr(13) + _
    "dbl_ShowDP=" + CStr(dbl_ShowDP) + Chr(10) + Chr(13) + _
    "dbl_ShowDPSupplyPressure=" + CStr(dbl_ShowDPSupplyPressure) + Chr(10) + Chr(13) + _
    "dbl_ShowMaxDP=" + CStr(dbl_ShowMaxDP) + Chr(10) + Chr(13) + _
    "dbl_ShowMaxDPSupplyPressure=" + CStr(dbl_ShowMaxDPSupplyPressure) + Chr(10) + Chr(13) + _
    "str_Message=" + str_Message + Chr(10) + Chr(13)

#End If

Exit Sub

CalcDP_35000_3_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcDP_30000_FTO_ATC(long_TrimPackageID As Long, _
    long_ActuatorPartID As Long, _
    dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, _
    dbl_InletPressure3 As Double, dbl_InletPressure4 As Double, _
    dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, _
    dbl_InletTemperature1 As Double, dbl_InletTemperature2 As Double, _
    dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, _
    dbl_ShutOffDP As Double, str_PressureUM As String, _
    str_TemperatureUM As String, str_CalibPressureUM As String, str_BodyMaterial As String, _
    str_Rating As String, str_PackingMaterial As String, _
    str_LeakageClass As String, str_SpringRange As String, str_FluidType As String, _
    dbl_Travel As Double, str_LengthUM As String, _
    dbl_ShowDP As Double, dbl_ShowDPSupplyPressure As Double, dbl_ShowMaxDPOpen As Double, _
    dbl_ShowMaxDPClosed As Double, dbl_ShowMaxDP As Double, _
    dbl_ShowMaxDPSupplyPressure As Double, str_Message As String)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves, "Varimax allowable DP calculations" by C. Vlassoff Aug. 23, 1996
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the selected valve and actuator shutoff capability and compare the result to the defined requirements.
'
'This subroutine is specific to  30000 series valves FTO, ATC
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID
'long_ActuatorPartID
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff pressure drop
'str_PressureUM                 unit of measure associated with inlet and outlet pressures
'str_TemperatureUM              unit of measure associated with inlet temperature
'str_CalibPressureUM            unit of measure associated with spring range and supply pressure
'str_LengthUM                   unit of measure for length (Travel)
'str_BodyMaterial
'str_Rating
'str_PackingMaterial            selected packing material
'str_LeakageClass               selected leakage class
'str_SpringRange                selected actuator spring range
'str_FluidType
'dbl_Travel                     Actuator travel
'
'Output Variables:
'dbl_ShowDP                     specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure       calculated required supply pressure for specified
'dbl_ShowMaxDPOpen
'dbl_ShowMaxDPClosed
'dbl_ShowMaxDP                  calculated maximum pressure drop for selected valve and actuator combination
'dbl_ShowMaxDPsupplyPressure    calculated required supply pressure for maximum pressure drop
'str_Message                    warning message if pressure drop requirement exceeds equipment capability
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/23 Larry Schoonover - not tested
'
'-----------------------------------------------------------------------------------------

On Error GoTo CalcDP_30000_FTO_ATC_Error

'***************************
'Declarations
'***************************
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_AllowANSIPressure As Double
Dim dbl_AllowDPStatic As Double
Dim dbl_MaxPressureDrop As Double
Dim dbl_ShowMaxPressureDrop As Double
Dim dbl_AllowShaftDPOpen As Double
Dim dbl_AllowShaftDPClosed As Double
Dim dbl_AllowPlugDP As Double
Dim dbl_AllowSupplyMax As Double
Dim dbl_AllowSupplyStability As Double
Dim dbl_ActuatorOutputTorque As Double
Dim dbl_OffBalOpenTorque As Double
Dim dbl_OffBalClosedTorque As Double
Dim dbl_SeatingTorque As Double
Dim dbl_PInitial As Double
Dim dbl_PFinal As Double
Dim dbl_SpringRate As Double
Dim dbl_TrimLimitOpen  As Double
Dim dbl_TrimLimitClosed As Double
Dim dbl_AllowDPClosed As Double
Dim dbl_AllowDPOpen As Double
Dim dbl_ReqSupplyOpen As Double
Dim dbl_ReqSupplyClosed As Double
Dim dbl_PSupplyDP As Double
Dim dbl_MaxAllowSupplyPressure As Double
Dim dbl_LeverLength As Double

Dim dbl_InletPressure(4) As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double


'***************************
'Parameter validation
'***************************
If (dbl_ShutOffDP < 0#) Then GoTo CalcDP_30000_FTO_ATC_Error
If (str_PressureUM = "") Then GoTo CalcDP_30000_FTO_ATC_Error
If (str_TemperatureUM = "") Then GoTo CalcDP_30000_FTO_ATC_Error
If (str_CalibPressureUM = "") Then GoTo CalcDP_30000_FTO_ATC_Error
If (str_PackingMaterial = "") Then GoTo CalcDP_30000_FTO_ATC_Error
If (str_LeakageClass = "") Then GoTo CalcDP_30000_FTO_ATC_Error
If (str_SpringRange = "") Then GoTo CalcDP_30000_FTO_ATC_Error

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4

'determine requirements
Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature, dbl_MaxInletTemperature)
Call CalcTestPressureDrop(dbl_InletPressure(), dbl_OutletPressure(), dbl_MaxPressureDrop)

dbl_ShowDP = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, str_PressureUM)
dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, str_TemperatureUM)
dbl_ShowMaxPressureDrop = dbl_MaxPressureDrop
dbl_MaxPressureDrop = ConvertPressureToPSIG(dbl_MaxPressureDrop, str_PressureUM)
dbl_ShowShutoffDP = dbl_ShutOffDP
If (dbl_ShutOffDP > 0) Then dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, str_PressureUM)

'determine ANSI limitation
Call GetANSIPressure(str_BodyMaterial, dbl_MaxInletTemperature, str_Rating, _
    dbl_AllowANSIPressure)

'get product data
Call Get30000Data(long_TrimPackageID, dbl_AllowShaftDPOpen, dbl_AllowShaftDPClosed, _
    dbl_AllowPlugDP, dbl_AllowSupplyMax, dbl_AllowSupplyStability)
Call GetVarimaxActuatorData(long_ActuatorPartID, dbl_ActuatorOutputTorque, _
    dbl_OffBalOpenTorque, dbl_OffBalClosedTorque, dbl_SeatingTorque, dbl_LeverLength)
Call CalcSpringRate(str_SpringRange, 1#, dbl_PInitial, 1#, dbl_PFinal, str_CalibPressureUM, _
    1#, str_LengthUM, dbl_SpringRate)
    'only PInitial and PFinal are required from this routine

'calculate required supply pressure
dbl_TrimLimitOpen = dbl_AllowShaftDPOpen
dbl_TrimLimitClosed = dbl_AllowShaftDPClosed
If (dbl_AllowPlugDP < dbl_AllowShaftDPClosed) Then dbl_TrimLimitClosed = dbl_AllowPlugDP

Call CalcAllowDP_30000_FTO_ATC(dbl_ActuatorOutputTorque, dbl_OffBalOpenTorque, _
    dbl_OffBalClosedTorque, dbl_SeatingTorque, dbl_PInitial, dbl_PFinal, _
    dbl_AllowSupplyMax, dbl_AllowSupplyStability, dbl_AllowDPClosed, dbl_AllowDPOpen)
Call CalcRequiredSupplyPressure30000_FTO_ATC(dbl_ActuatorOutputTorque, _
    dbl_OffBalOpenTorque, dbl_OffBalClosedTorque, dbl_SeatingTorque, dbl_PInitial, _
    dbl_PFinal, dbl_MaxInletPressure, dbl_MaxPressureDrop, dbl_ReqSupplyOpen, _
    dbl_ReqSupplyClosed)
    
If (dbl_ReqSupplyOpen < dbl_ReqSupplyClosed) Then
    dbl_PSupplyDP = dbl_ReqSupplyClosed
Else
    dbl_PSupplyDP = dbl_ReqSupplyOpen
End If

If (dbl_PSupplyDP < 35) Then dbl_PSupplyDP = 35

'evlauation
If (dbl_MaxInletPressure > dbl_AllowANSIPressure) Then str_Message = ANSIMESSAGE
If (dbl_MaxInletPressure > dbl_AllowPlugDP) Then str_Message = TRIMMESSAGE
If (dbl_MaxInletPressure > dbl_AllowShaftDPClosed Or dbl_MaxPressureDrop > _
    dbl_AllowShaftDPOpen) Then str_Message = TRIMMESSAGE
If (dbl_ReqSupplyOpen > dbl_AllowSupplyMax Or dbl_ReqSupplyClosed > dbl_AllowSupplyMax) Then _
    str_Message = ACTUATORMESSAGE
If (dbl_ReqSupplyOpen > dbl_AllowSupplyStability) Then str_Message = ACTUATORMESSAGE

If (dbl_MaxInletPressure <= dbl_AllowDPClosed And dbl_MaxPressureDrop < dbl_AllowDPOpen) Then
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_CalibPressureUM)
    dbl_ShowMaxDPOpen = ConvertPressureFromPSIG(dbl_AllowDPOpen, str_PressureUM)
    dbl_ShowMaxDPClosed = ConvertPressureFromPSIG(dbl_AllowDPClosed, str_PressureUM)
    
    If (dbl_AllowSupplyMax < dbl_AllowSupplyStability) Then
        dbl_MaxAllowSupplyPressure = dbl_AllowSupplyMax
    Else
        dbl_MaxAllowSupplyPressure = dbl_AllowSupplyStability
    End If
    
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_MaxAllowSupplyPressure, _
        str_CalibPressureUM)
End If

Exit Sub

CalcDP_30000_FTO_ATC_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcDP_30000_FTC_ATC(long_TrimPackageID As Long, _
    long_ActuatorPartID As Long, _
    dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, _
    dbl_InletPressure3 As Double, dbl_InletPressure4 As Double, _
    dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, _
    dbl_InletTemperature1 As Double, dbl_InletTemperature2 As Double, _
    dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, _
    dbl_ShutOffDP As Double, str_PressureUM As String, _
    str_TemperatureUM As String, str_CalibPressureUM As String, str_BodyMaterial As String, _
    str_Rating As String, str_PackingMaterial As String, _
    str_LeakageClass As String, str_SpringRange As String, str_FluidType As String, _
    dbl_Travel As Double, str_LengthUM As String, _
    dbl_ShowDP As Double, dbl_ShowDPSupplyPressure As Double, dbl_ShowMaxDPOpen As Double, _
    dbl_ShowMaxDPClosed As Double, dbl_ShowMaxDP As Double, _
    dbl_ShowMaxDPSupplyPressure As Double, str_Message As String)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves, "Varimax allowable DP calculations" by C. Vlassoff Aug. 23, 1996
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the selected valve and actuator shutoff capability and compare the result to the defined requirements.
'
'This subroutine is specific to  30000 series valves FTC, ATC
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID
'long_ActuatorPartID
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff pressure drop
'str_PressureUM                 unit of measure associated with inlet and outlet pressures
'str_TemperatureUM              unit of measure associated with inlet temperature
'str_CalibPressureUM            unit of measure associated with spring range and supply pressure
'str_LengthUM                   unit of measure for length (Travel)
'str_BodyMaterial
'str_Rating
'str_PackingMaterial            selected packing material
'str_LeakageClass               selected leakage class
'str_SpringRange                selected actuator spring range
'str_FluidType
'dbl_Travel                     Actuator travel
'
'Output Variables:
'dbl_ShowDP                     specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure       calculated required supply pressure for specified
'dbl_ShowMaxDPOpen
'dbl_ShowMaxDPClosed
'dbl_ShowMaxDP                  calculated maximum pressure drop for selected valve and actuator combination
'dbl_ShowMaxDPsupplyPressure    calculated required supply pressure for maximum pressure drop
'str_Message                    warning message if pressure drop requirement exceeds equipment capability
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/23 Larry Schoonover - not tested
'
'-----------------------------------------------------------------------------------------

On Error GoTo CalcDP_30000_FTC_ATC_Error

'***************************
'Declarations
'***************************
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_AllowANSIPressure As Double
Dim dbl_AllowDPStatic As Double
Dim dbl_MaxPressureDrop As Double
Dim dbl_ShowMaxPressureDrop As Double
Dim dbl_AllowShaftDPOpen As Double
Dim dbl_AllowShaftDPClosed As Double
Dim dbl_AllowPlugDP As Double
Dim dbl_AllowSupplyMax As Double
Dim dbl_AllowSupplyStability As Double
Dim dbl_ActuatorOutputTorque As Double
Dim dbl_OffBalOpenTorque As Double
Dim dbl_OffBalClosedTorque As Double
Dim dbl_SeatingTorque As Double
Dim dbl_PInitial As Double
Dim dbl_PFinal As Double
Dim dbl_SpringRate As Double
Dim dbl_TrimLimitOpen  As Double
Dim dbl_TrimLimitClosed As Double
Dim dbl_AllowDPClosed As Double
Dim dbl_AllowDPOpen As Double
Dim dbl_ReqSupplyOpen As Double
Dim dbl_ReqSupplyClosed As Double
Dim dbl_PSupplyDP As Double
Dim dbl_MaxAllowSupplyPressure As Double
Dim dbl_LeverLength As Double

Dim dbl_InletPressure(4) As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double

'***************************
'Parameter validation
'***************************
If (dbl_ShutOffDP < 0#) Then GoTo CalcDP_30000_FTC_ATC_Error
If (str_PressureUM = "") Then GoTo CalcDP_30000_FTC_ATC_Error
If (str_TemperatureUM = "") Then GoTo CalcDP_30000_FTC_ATC_Error
If (str_CalibPressureUM = "") Then GoTo CalcDP_30000_FTC_ATC_Error
If (str_PackingMaterial = "") Then GoTo CalcDP_30000_FTC_ATC_Error
If (str_LeakageClass = "") Then GoTo CalcDP_30000_FTC_ATC_Error
If (str_SpringRange = "") Then GoTo CalcDP_30000_FTC_ATC_Error

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4

'determine requirements
Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature, dbl_MaxInletTemperature)
Call CalcTestPressureDrop(dbl_InletPressure(), dbl_OutletPressure(), dbl_MaxPressureDrop)

dbl_ShowDP = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, str_PressureUM)
dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, str_TemperatureUM)
dbl_ShowMaxPressureDrop = dbl_MaxPressureDrop
dbl_MaxPressureDrop = ConvertPressureToPSIG(dbl_MaxPressureDrop, str_PressureUM)
dbl_ShowShutoffDP = dbl_ShutOffDP
If (dbl_ShutOffDP > 0) Then dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, str_PressureUM)

'determine ANSI limitation
Call GetANSIPressure(str_BodyMaterial, dbl_MaxInletTemperature, str_Rating, _
    dbl_AllowANSIPressure)

'get product data
Call Get30000Data(long_TrimPackageID, dbl_AllowShaftDPOpen, dbl_AllowShaftDPClosed, _
    dbl_AllowPlugDP, dbl_AllowSupplyMax, dbl_AllowSupplyStability)
Call GetVarimaxActuatorData(long_ActuatorPartID, dbl_ActuatorOutputTorque, _
    dbl_OffBalOpenTorque, dbl_OffBalClosedTorque, dbl_SeatingTorque, dbl_LeverLength)
Call CalcSpringRate(str_SpringRange, 1#, dbl_PInitial, 1#, dbl_PFinal, str_CalibPressureUM, _
    1#, str_LengthUM, dbl_SpringRate)
    'only PInitial and PFinal are required from this routine

'calculate required supply pressure
dbl_TrimLimitOpen = dbl_AllowShaftDPOpen
dbl_TrimLimitClosed = dbl_AllowShaftDPClosed
If (dbl_AllowPlugDP < dbl_AllowShaftDPClosed) Then dbl_TrimLimitClosed = dbl_AllowPlugDP

Call CalcAllowDP_30000_FTC_ATC(dbl_ActuatorOutputTorque, dbl_OffBalOpenTorque, _
    dbl_OffBalClosedTorque, dbl_SeatingTorque, dbl_PInitial, dbl_PFinal, _
    dbl_AllowSupplyMax, dbl_AllowSupplyStability, dbl_AllowDPClosed, dbl_AllowDPOpen)
dbl_PSupplyDP = 35

'evaluation
If (dbl_MaxInletPressure > dbl_AllowANSIPressure) Then str_Message = ANSIMESSAGE
If (dbl_MaxInletPressure > dbl_AllowPlugDP) Then str_Message = TRIMMESSAGE
If (dbl_MaxInletPressure > dbl_AllowShaftDPClosed Or dbl_MaxPressureDrop > _
    dbl_AllowShaftDPOpen) Then str_Message = TRIMMESSAGE
If (dbl_MaxInletPressure > dbl_AllowDPClosed Or dbl_MaxPressureDrop > _
    dbl_AllowDPOpen) Then str_Message = ACTUATORMESSAGE
If (dbl_PSupplyDP > dbl_AllowSupplyStability Or dbl_PSupplyDP > dbl_AllowSupplyMax) _
    Then str_Message = ACTUATORMESSAGE

If (dbl_MaxInletPressure <= dbl_AllowDPClosed And dbl_MaxPressureDrop < dbl_AllowDPOpen) Then
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_CalibPressureUM)
    dbl_ShowMaxDPOpen = ConvertPressureFromPSIG(dbl_AllowDPOpen, str_PressureUM)
    dbl_ShowMaxDPClosed = ConvertPressureFromPSIG(dbl_AllowDPClosed, str_PressureUM)
    
    If (dbl_AllowSupplyMax < dbl_AllowSupplyStability) Then
        dbl_MaxAllowSupplyPressure = dbl_AllowSupplyMax
    Else
        dbl_MaxAllowSupplyPressure = dbl_AllowSupplyStability
    End If
    
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_MaxAllowSupplyPressure, _
        str_CalibPressureUM)
End If

Exit Sub

CalcDP_30000_FTC_ATC_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcDP_30000_FTO_ATO(long_TrimPackageID As Long, _
    long_ActuatorPartID As Long, _
    dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, _
    dbl_InletPressure3 As Double, dbl_InletPressure4 As Double, _
    dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, _
    dbl_InletTemperature1 As Double, dbl_InletTemperature2 As Double, _
    dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, _
    dbl_ShutOffDP As Double, str_PressureUM As String, _
    str_TemperatureUM As String, str_CalibPressureUM As String, str_BodyMaterial As String, _
    str_Rating As String, str_PackingMaterial As String, _
    str_LeakageClass As String, str_SpringRange As String, str_FluidType As String, _
    dbl_Travel As Double, str_LengthUM As String, _
    dbl_ShowDP As Double, dbl_ShowDPSupplyPressure As Double, dbl_ShowMaxDPOpen As Double, _
    dbl_ShowMaxDPClosed As Double, dbl_ShowMaxDP As Double, _
    dbl_ShowMaxDPSupplyPressure As Double, str_Message As String)
'--------------------------------------------------------------------------------------
'Reference: Actuator sizing methods for rotary valves, "Varimax allowable DP calculations" by C. Vlassoff Aug. 23, 1996
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the selected valve and actuator shutoff capability and compare the result to the defined requirements.
'
'This subroutine is specific to  30000 series valves FTO, ATO
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID
'long_ActuatorPartID
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff pressure drop
'str_PressureUM                 unit of measure associated with inlet and outlet pressures
'str_TemperatureUM              unit of measure associated with inlet temperature
'str_CalibPressureUM            unit of measure associated with spring range and supply pressure
'str_LengthUM                   unit of measure for length (Travel)
'str_BodyMaterial
'str_Rating
'str_PackingMaterial            selected packing material
'str_LeakageClass               selected leakage class
'str_SpringRange                selected actuator spring range
'str_FluidType
'dbl_Travel                     Actuator travel
'
'Output Variables:
'dbl_ShowDP                     specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure       calculated required supply pressure for specified
'dbl_ShowMaxDPOpen
'dbl_ShowMaxDPClosed
'dbl_ShowMaxDP                  calculated maximum pressure drop for selected valve and actuator combination
'dbl_ShowMaxDPsupplyPressure    calculated required supply pressure for maximum pressure drop
'str_Message                    warning message if pressure drop requirement exceeds equipment capability
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/23 Larry Schoonover - not tested
'
'-----------------------------------------------------------------------------------------

On Error GoTo CalcDP_30000_FTO_ATO_Error

'***************************
'Declarations
'***************************
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_AllowANSIPressure As Double
Dim dbl_AllowDPStatic As Double
Dim dbl_MaxPressureDrop As Double
Dim dbl_ShowMaxPressureDrop As Double
Dim dbl_AllowShaftDPOpen As Double
Dim dbl_AllowShaftDPClosed As Double
Dim dbl_AllowPlugDP As Double
Dim dbl_AllowSupplyMax As Double
Dim dbl_AllowSupplyStability As Double
Dim dbl_ActuatorOutputTorque As Double
Dim dbl_OffBalOpenTorque As Double
Dim dbl_OffBalClosedTorque As Double
Dim dbl_SeatingTorque As Double
Dim dbl_PInitial As Double
Dim dbl_PFinal As Double
Dim dbl_SpringRate As Double
Dim dbl_TrimLimitOpen  As Double
Dim dbl_TrimLimitClosed As Double
Dim dbl_AllowDPClosed As Double
Dim dbl_AllowDPOpen As Double
Dim dbl_ReqSupplyOpen As Double
Dim dbl_ReqSupplyClosed As Double
Dim dbl_PSupplyDP As Double
Dim dbl_MaxAllowSupplyPressure As Double
Dim dbl_LeverLength As Double

Dim dbl_InletPressure(4) As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double

'***************************
'Parameter validation
'***************************
If (dbl_ShutOffDP < 0#) Then GoTo CalcDP_30000_FTO_ATO_Error
If (str_PressureUM = "") Then GoTo CalcDP_30000_FTO_ATO_Error
If (str_TemperatureUM = "") Then GoTo CalcDP_30000_FTO_ATO_Error
If (str_CalibPressureUM = "") Then GoTo CalcDP_30000_FTO_ATO_Error
If (str_PackingMaterial = "") Then GoTo CalcDP_30000_FTO_ATO_Error
If (str_LeakageClass = "") Then GoTo CalcDP_30000_FTO_ATO_Error
If (str_SpringRange = "") Then GoTo CalcDP_30000_FTO_ATO_Error

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4

'determine requirements
Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature, dbl_MaxInletTemperature)
Call CalcTestPressureDrop(dbl_InletPressure(), dbl_OutletPressure(), dbl_MaxPressureDrop)

dbl_ShowDP = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, str_PressureUM)
dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, str_TemperatureUM)
dbl_ShowMaxPressureDrop = dbl_MaxPressureDrop
dbl_MaxPressureDrop = ConvertPressureToPSIG(dbl_MaxPressureDrop, str_PressureUM)
dbl_ShowShutoffDP = dbl_ShutOffDP
If (dbl_ShutOffDP > 0) Then dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, str_PressureUM)

'determine ANSI limitation
Call GetANSIPressure(str_BodyMaterial, dbl_MaxInletTemperature, str_Rating, _
    dbl_AllowANSIPressure)

'get product data
Call Get30000Data(long_TrimPackageID, dbl_AllowShaftDPOpen, dbl_AllowShaftDPClosed, _
    dbl_AllowPlugDP, dbl_AllowSupplyMax, dbl_AllowSupplyStability)
Call GetVarimaxActuatorData(long_ActuatorPartID, dbl_ActuatorOutputTorque, _
    dbl_OffBalOpenTorque, dbl_OffBalClosedTorque, dbl_SeatingTorque, dbl_LeverLength)
Call CalcSpringRate(str_SpringRange, 1#, dbl_PInitial, 1#, dbl_PFinal, str_CalibPressureUM, _
    1#, str_LengthUM, dbl_SpringRate)
    'only PInitial and PFinal are required from this routine

'calculate required supply pressure
dbl_TrimLimitOpen = dbl_AllowShaftDPOpen
dbl_TrimLimitClosed = dbl_AllowShaftDPClosed
If (dbl_AllowPlugDP < dbl_AllowShaftDPClosed) Then dbl_TrimLimitClosed = dbl_AllowPlugDP

Call CalcAllowDP_30000_FTO_ATO(dbl_ActuatorOutputTorque, dbl_OffBalOpenTorque, _
    dbl_OffBalClosedTorque, dbl_SeatingTorque, dbl_PInitial, dbl_PFinal, _
    dbl_AllowSupplyMax, dbl_AllowSupplyStability, dbl_AllowDPClosed, dbl_AllowDPOpen)
dbl_PSupplyDP = 35

'evaluation
If (dbl_MaxInletPressure > dbl_AllowANSIPressure) Then str_Message = ANSIMESSAGE
If (dbl_MaxInletPressure > dbl_AllowPlugDP) Then str_Message = TRIMMESSAGE
If (dbl_MaxInletPressure > dbl_AllowShaftDPClosed Or dbl_MaxPressureDrop > _
    dbl_AllowShaftDPOpen) Then str_Message = TRIMMESSAGE
If (dbl_MaxInletPressure > dbl_AllowDPClosed Or dbl_MaxPressureDrop > _
    dbl_AllowDPOpen) Then str_Message = ACTUATORMESSAGE
If (dbl_PSupplyDP > dbl_AllowSupplyStability Or dbl_PSupplyDP > dbl_AllowSupplyMax) _
    Then str_Message = ACTUATORMESSAGE

If (dbl_MaxInletPressure <= dbl_AllowDPClosed And dbl_MaxPressureDrop < dbl_AllowDPOpen) Then
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_CalibPressureUM)
    dbl_ShowMaxDPOpen = ConvertPressureFromPSIG(dbl_AllowDPOpen, str_PressureUM)
    dbl_ShowMaxDPClosed = ConvertPressureFromPSIG(dbl_AllowDPClosed, str_PressureUM)
    
    If (dbl_AllowSupplyMax < dbl_AllowSupplyStability) Then
        dbl_MaxAllowSupplyPressure = dbl_AllowSupplyMax
    Else
        dbl_MaxAllowSupplyPressure = dbl_AllowSupplyStability
    End If
    
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_MaxAllowSupplyPressure, _
        str_CalibPressureUM)
End If

Exit Sub

CalcDP_30000_FTO_ATO_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcDP_30000_FTC_ATO(long_TrimPackageID As Long, _
    long_ActuatorPartID As Long, _
    dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, _
    dbl_InletPressure3 As Double, dbl_InletPressure4 As Double, _
    dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, _
    dbl_InletTemperature1 As Double, dbl_InletTemperature2 As Double, _
    dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, _
    dbl_ShutOffDP As Double, str_PressureUM As String, _
    str_TemperatureUM As String, str_CalibPressureUM As String, str_BodyMaterial As String, _
    str_Rating As String, str_PackingMaterial As String, _
    str_LeakageClass As String, str_SpringRange As String, str_FluidType As String, _
    dbl_Travel As Double, str_LengthUM As String, _
    dbl_ShowDP As Double, dbl_ShowDPSupplyPressure As Double, dbl_ShowMaxDPOpen As Double, _
    dbl_ShowMaxDPClosed As Double, dbl_ShowMaxDP As Double, _
    dbl_ShowMaxDPSupplyPressure As Double, str_Message As String)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves, "Varimax allowable DP calculations" by C. Vlassoff Aug. 23, 1996
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the selected valve and actuator shutoff capability and compare the result to the defined requirements.
'
'This subroutine is specific to  30000 series valves FTO, ATC
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID
'long_ActuatorPartID
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff pressure drop
'str_PressureUM                 unit of measure associated with inlet and outlet pressures
'str_TemperatureUM              unit of measure associated with inlet temperature
'str_CalibPressureUM            unit of measure associated with spring range and supply pressure
'str_LengthUM                   unit of measure for length (Travel)
'str_BodyMaterial
'str_Rating
'str_PackingMaterial            selected packing material
'str_LeakageClass               selected leakage class
'str_SpringRange                selected actuator spring range
'str_FluidType
'dbl_Travel                     Actuator travel
'
'Output Variables:
'dbl_ShowDP                     specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure       calculated required supply pressure for specified
'dbl_ShowMaxDPOpen
'dbl_ShowMaxDPClosed
'dbl_ShowMaxDP                  calculated maximum pressure drop for selected valve and actuator combination
'dbl_ShowMaxDPsupplyPressure    calculated required supply pressure for maximum pressure drop
'str_Message                    warning message if pressure drop requirement exceeds equipment capability
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/23 Larry Schoonover - not tested
'
'-----------------------------------------------------------------------------------------

On Error GoTo CalcDP_30000_FTC_ATO_Error

'***************************
'Declarations
'***************************
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_AllowANSIPressure As Double
Dim dbl_AllowDPStatic As Double
Dim dbl_MaxPressureDrop As Double
Dim dbl_ShowMaxPressureDrop As Double
Dim dbl_AllowShaftDPOpen As Double
Dim dbl_AllowShaftDPClosed As Double
Dim dbl_AllowPlugDP As Double
Dim dbl_AllowSupplyMax As Double
Dim dbl_AllowSupplyStability As Double
Dim dbl_ActuatorOutputTorque As Double
Dim dbl_OffBalOpenTorque As Double
Dim dbl_OffBalClosedTorque As Double
Dim dbl_SeatingTorque As Double
Dim dbl_PInitial As Double
Dim dbl_PFinal As Double
Dim dbl_SpringRate As Double
Dim dbl_TrimLimitOpen  As Double
Dim dbl_TrimLimitClosed As Double
Dim dbl_AllowDPClosed As Double
Dim dbl_AllowDPOpen As Double
Dim dbl_ReqSupplyOpen As Double
Dim dbl_ReqSupplyClosed As Double
Dim dbl_PSupplyDP As Double
Dim dbl_MaxAllowSupplyPressure As Double
Dim dbl_LeverLength As Double

Dim dbl_InletPressure(4) As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double

'***************************
'Parameter validation
'***************************
If (dbl_ShutOffDP < 0#) Then GoTo CalcDP_30000_FTC_ATO_Error
If (str_PressureUM = "") Then GoTo CalcDP_30000_FTC_ATO_Error
If (str_TemperatureUM = "") Then GoTo CalcDP_30000_FTC_ATO_Error
If (str_CalibPressureUM = "") Then GoTo CalcDP_30000_FTC_ATO_Error
If (str_PackingMaterial = "") Then GoTo CalcDP_30000_FTC_ATO_Error
If (str_LeakageClass = "") Then GoTo CalcDP_30000_FTC_ATO_Error
If (str_SpringRange = "") Then GoTo CalcDP_30000_FTC_ATO_Error

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4

'determine requirements
Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature, dbl_MaxInletTemperature)
Call CalcTestPressureDrop(dbl_InletPressure(), dbl_OutletPressure(), dbl_MaxPressureDrop)

dbl_ShowDP = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, str_PressureUM)
dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, str_TemperatureUM)
dbl_ShowMaxPressureDrop = dbl_MaxPressureDrop
dbl_MaxPressureDrop = ConvertPressureToPSIG(dbl_MaxPressureDrop, str_PressureUM)
dbl_ShowShutoffDP = dbl_ShutOffDP
If (dbl_ShutOffDP > 0) Then dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, str_PressureUM)

'determine ANSI limitation
Call GetANSIPressure(str_BodyMaterial, dbl_MaxInletTemperature, str_Rating, _
    dbl_AllowANSIPressure)

'get product data
Call Get30000Data(long_TrimPackageID, dbl_AllowShaftDPOpen, dbl_AllowShaftDPClosed, _
    dbl_AllowPlugDP, dbl_AllowSupplyMax, dbl_AllowSupplyStability)
Call GetVarimaxActuatorData(long_ActuatorPartID, dbl_ActuatorOutputTorque, _
    dbl_OffBalOpenTorque, dbl_OffBalClosedTorque, dbl_SeatingTorque, dbl_LeverLength)
Call CalcSpringRate(str_SpringRange, 1#, dbl_PInitial, 1#, dbl_PFinal, str_CalibPressureUM, _
    1#, str_LengthUM, dbl_SpringRate)
    'only PInitial and PFinal are required from this routine

'calculate required supply pressure
dbl_TrimLimitOpen = dbl_AllowShaftDPOpen
dbl_TrimLimitClosed = dbl_AllowShaftDPClosed
If (dbl_AllowPlugDP < dbl_AllowShaftDPClosed) Then dbl_TrimLimitClosed = dbl_AllowPlugDP

Call CalcAllowDP_30000_FTC_ATO(dbl_ActuatorOutputTorque, dbl_OffBalOpenTorque, _
    dbl_OffBalClosedTorque, dbl_SeatingTorque, dbl_PInitial, dbl_PFinal, _
    dbl_AllowSupplyMax, dbl_AllowSupplyStability, dbl_AllowDPClosed, dbl_AllowDPOpen)
Call CalcRequiredSupplyPressure30000_FTC_ATO(dbl_ActuatorOutputTorque, _
    dbl_OffBalOpenTorque, dbl_OffBalClosedTorque, dbl_SeatingTorque, dbl_PInitial, _
    dbl_PFinal, dbl_MaxInletPressure, dbl_MaxPressureDrop, dbl_ReqSupplyOpen, _
    dbl_ReqSupplyClosed)
    
If (dbl_ReqSupplyOpen < dbl_ReqSupplyClosed) Then
    dbl_PSupplyDP = dbl_ReqSupplyClosed
Else
    dbl_PSupplyDP = dbl_ReqSupplyOpen
End If

If (dbl_PSupplyDP < 35) Then dbl_PSupplyDP = 35

'evlauation
If (dbl_MaxInletPressure > dbl_AllowANSIPressure) Then str_Message = ANSIMESSAGE
If (dbl_MaxInletPressure > dbl_AllowPlugDP) Then str_Message = TRIMMESSAGE
If (dbl_MaxInletPressure > dbl_AllowShaftDPClosed Or dbl_MaxPressureDrop > _
    dbl_AllowShaftDPOpen) Then str_Message = TRIMMESSAGE
If (dbl_ReqSupplyOpen > dbl_AllowSupplyMax Or dbl_ReqSupplyClosed > dbl_AllowSupplyMax) Then _
    str_Message = ACTUATORMESSAGE
If (dbl_ReqSupplyOpen > dbl_AllowSupplyStability) Then str_Message = ACTUATORMESSAGE

If (dbl_MaxInletPressure <= dbl_AllowDPClosed And dbl_MaxPressureDrop < dbl_AllowDPOpen) Then
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_CalibPressureUM)
    dbl_ShowMaxDPOpen = ConvertPressureFromPSIG(dbl_AllowDPOpen, str_PressureUM)
    dbl_ShowMaxDPClosed = ConvertPressureFromPSIG(dbl_AllowDPClosed, str_PressureUM)
    
    If (dbl_AllowSupplyMax < dbl_AllowSupplyStability) Then
        dbl_MaxAllowSupplyPressure = dbl_AllowSupplyMax
    Else
        dbl_MaxAllowSupplyPressure = dbl_AllowSupplyStability
    End If
    
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_MaxAllowSupplyPressure, _
        str_CalibPressureUM)
End If

Exit Sub

CalcDP_30000_FTC_ATO_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcDP_28000(long_TrimPackageID As Long, _
    dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, _
    dbl_InletPressure3 As Double, dbl_InletPressure4 As Double, _
    dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, _
    dbl_InletTemperature1 As Double, dbl_InletTemperature2 As Double, _
    dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, _
    dbl_ShutOffDP As Double, str_PressureUM As String, str_TemperatureUM As String, str_CalibPressureUM As String, _
    str_BodyMaterial As String, str_Rating As String, dbl_ShowDP As Double, _
    dbl_ShowDPSupplyPressure As Double, dbl_ShowMaxDP As Double, _
    dbl_ShowMaxDPSupplyPressure As Double, str_Message As String)
'--------------------------------------------------------------------------------------
'Due to the small orifice diameters used in this product line there is very little variation in allowable pressure drop limits between trim sizes.  As a result, the trim limit and the ANSI limit may be used in place of actuator calculations.
'This subroutine is specific to  28000 series valves
'---------------------------------------------------------------------------------------
'Input Variables:
'long_TrimPackageID
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff pressure                                 drop
'str_PressureUM                 unit of measure associated with inlet and outlet                                pressures
'str_TemperatureUM              unit of measure associated with inlet                                   temperature
'str_CalibPressureUM            unit of measure associated with spring range                                and supply pressure
'str_BodyMatl                   selected material of body
'str_Rating                     selected body rating
'
'Output Variables:
'dbl_ShowDP                  specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure    calculated required supply pressure for specified
'dbl_ShowMaxDPopen           calculated maximum pressure drop for selected
'                            valve and actuator combination
'dbl_ShowMaxDPclosed         calculated maximum pressure drop for selected
'                            valve and actuator combination
'
'dbl_ShowMaxDPsupplyPressure calculated required supply pressure for
'                            maximum pressure drop
'str_Message                 warning message if pressure drop requirement
'                            exceeds equipment capability
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/23 Larry Schoonover - not tested
'
'-----------------------------------------------------------------------------------------

On Error GoTo CalcDP_28000_Error

'***************************
'Declarations
'***************************
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim dbl_AllowANSIPressure As Double
Dim dbl_AllowTrimDP As Double
Dim dbl_PSupplyDP As Double

Dim dbl_InletPressure(4) As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double
Dim dbl_ShowShutoffDP As Double

'***************************
'Parameter validation
'***************************
If (dbl_ShutOffDP < 0#) Then GoTo CalcDP_28000_Error
If (str_PressureUM = "") Then GoTo CalcDP_28000_Error
If (str_TemperatureUM = "") Then GoTo CalcDP_28000_Error
If (str_CalibPressureUM = "") Then GoTo CalcDP_28000_Error

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4

'determine requirements
Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature, dbl_MaxInletTemperature)

dbl_ShowDP = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, str_PressureUM)
dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, str_TemperatureUM)
dbl_ShowShutoffDP = dbl_ShutOffDP
If (dbl_ShutOffDP > 0) Then dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, str_PressureUM)

'determine ANSI limitation
Call GetANSIPressure(str_BodyMaterial, dbl_MaxInletTemperature, str_Rating, _
    dbl_AllowANSIPressure)

'get product data
Call Get28000Data(long_TrimPackageID, dbl_AllowTrimDP, dbl_PSupplyDP)

'evaluation
If (dbl_MaxInletPressure > dbl_AllowANSIPressure) Then str_Message = ANSIMESSAGE
If (dbl_MaxInletPressure > dbl_AllowTrimDP) Then str_Message = TRIMMESSAGE

If (dbl_MaxInletPressure <= dbl_AllowTrimDP) Then
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_CalibPressureUM)
End If

dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowTrimDP, str_PressureUM)
dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_CalibPressureUM)

Exit Sub

CalcDP_28000_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' END OF Actuator UL
''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INCLUDE Actuator (low level)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub SetDatabasePath(str_Path As String)
'--------------------------------------------------------------------------------------
' this routine sets the global variable that contains the path to the database
'--------------------------------------------------------------------------------------
'Input Variables:
'str_Path
'
'Modified variables:
'g_str_DatabasePath  is set by this routine
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/11 Larry Schoonover
'
'-----------------------------------------------------------------------------------------
g_str_DatabasePath = str_Path

End Sub
Public Function InsertStr(str_SearchedString As String, str_FindString As String, _
    str_InsertString As String) As String
Dim Index As Integer
Index = InStr(1, str_SearchedString, str_FindString)
If (Index = 0) Then
    InsertStr = str_SearchedString
    Exit Function
End If
InsertStr = Left(str_SearchedString, Index) + str_InsertString + _
    Mid(str_SearchedString, Index + 1)

End Function

Public Function ReplaceStr(str_SearchedString As String, str_FindString As String, _
    ParamArray str_ReplaceString()) As String
    
Dim Index As Integer
Dim Count As Integer
Dim str_HoldString As String
Dim str_String As Variant

Count = 1
str_HoldString = str_SearchedString

For Each str_String In str_ReplaceString

    Index = InStr(Count, str_HoldString, str_FindString)
    If (Index = 0) Then
        ReplaceStr = str_HoldString
        Exit Function
    End If
    
    str_HoldString = Left(str_HoldString, Index - 1) + str_String + _
        Mid(str_HoldString, Index + Len(str_FindString))
        
    Count = Index + Len(str_String)
    
Next str_String

ReplaceStr = str_HoldString

End Function
' ****************** TEMP CODE UNTIL THIS ROUTINE IS SPEC'd
Public Sub GetPackingCoef(str_PackingMaterial As String, dbl_Mu As Double)

'--------------------------------------------------------------------------------------
' ****************** TEMP CODE UNTIL THIS ROUTINE IS SPEC'd
'--------------------------------------------------------------------------------------
'Input Variables:
'str_PackingMaterial    packing material
'                            "Teflon" or "Graphite"
'
'Output Variables:
'dbl_Mu                 friction  0.03 for teflon, 0.15 for graphite
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/28 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'8/4/98 - John Haberlin - additional packing choices
'-----------------------------------------------------------------------------------------

On Error GoTo GetPackingCoef_Error

'****************************
'Begin calculations
'****************************
Select Case str_PackingMaterial
    Case "Teflon", "TFE/aramid", "PTFE", "LE PACKING", "LE FIRE TESTED PACKING"
        dbl_Mu = 0.03
    Case "Graphite", "graphite", "Grafoil"
        dbl_Mu = 0.15
    Case Else
        GoTo GetPackingCoef_Error
End Select
    
Exit Sub

GetPackingCoef_Error:
'Placeholder.  How do we want to handle error messages?

End Sub
' ****************** TEMP CODE UNTIL THIS ROUTINE IS SPEC'd
Public Sub GetSeatingCoef(str_LeakageClass As String, dbl_SeatingLoadCoefficient As Double)

'--------------------------------------------------------------------------------------
' ****************** TEMP CODE UNTIL THIS ROUTINE IS SPEC'd
'--------------------------------------------------------------------------------------
'Input Variables:
'str_LeakageClass    Leakage class - "IV" "V" or "VI"
'
'Output Variables:
'dbl_SeatingLoadCoefficient coefficient
'                           "IV"=30
'                           "V" =150
'                           "VI"=50
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/28 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo GetSeatingCoef_Error

'****************************
'Begin calculations
'****************************
Select Case str_LeakageClass
    Case "IV"
        dbl_SeatingLoadCoefficient = 30#
    Case "V"
        dbl_SeatingLoadCoefficient = 150#
    Case "VI"
        dbl_SeatingLoadCoefficient = 50#
    Case Else
        GoTo GetSeatingCoef_Error
End Select
    
Exit Sub

GetSeatingCoef_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcAllowStemForce(dbl_Sy As Double, dbl_E As Double, _
    dbl_StemDiameter As Double, dbl_L As Double, dbl_Sd As Double, _
    dbl_As As Double, dbl_AllowableForce As Double)
'--------------------------------------------------------------------------------------
'Reference : Basic Practice Guide - Control Valve Body S/A Design, rev. C, 12/1/97
'
'This subroutine is used to calculate the allowable load limit for a valve stem
'   with consideration to buckling and compressive loading.
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_Sy                  allowable yield stress at temperature          psi
'dbl_E                   modulus of elasticity                          psi
'dbl_StemDiameter        stem diameter                                  in
'dbl_L                   supported stem length (default = 6)            in
'dbl_Sd                  allowable design stress at temperature         psi
'dbl_As                  stress area of stem                            in^2
'-----------------------------------------------------------------------------------------
'Output Variables:
'dbl_AllowableForce      calculated allowable stem load                 lbs
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/28 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'8/25/98 - John Haberlin - revised name from "CalcAllowStemCompressiveForce"
'                                       to   "CalcAllowStemForce"
'           reason: the subroutine had been previously revised to perform both
'                   compressive load and buckling calculations.
'-----------------------------------------------------------------------------------------

On Error GoTo CalcAllowStemForce_Error

'***************************
'Declarations
'***************************
Dim dbl_EndCondition As Double
Dim dbl_RadiusOfGyration As Double
Dim dbl_SlendernessRatioLimit As Double
Dim dbl_SlendernessRatio As Double
Dim dbl_CriticalLoad As Double
Dim dbl_AllowableBuckleLoad As Double
Dim dbl_AllowableCompressionLoad As Double

'***************************
'Parameter validation
'***************************
' if any of these are 0 we will get a divide by 0
If (dbl_Sy <= 0#) Then GoTo CalcAllowStemForce_Error
If (dbl_E = 0#) Then GoTo CalcAllowStemForce_Error
If (dbl_StemDiameter <= 0#) Then GoTo CalcAllowStemForce_Error
If (dbl_L <= 0#) Then GoTo CalcAllowStemForce_Error
If (dbl_As <= 0#) Then GoTo CalcAllowStemForce_Error

'****************************
'Begin calculations
'****************************
dbl_EndCondition = 2#
dbl_RadiusOfGyration = dbl_StemDiameter / 4#
dbl_SlendernessRatioLimit = Sqr((2# * PI2 * dbl_EndCondition * dbl_E) / dbl_Sy)
dbl_SlendernessRatio = dbl_L / dbl_RadiusOfGyration

If (dbl_SlendernessRatio <= dbl_SlendernessRatioLimit) Then
    dbl_CriticalLoad = dbl_As * (dbl_Sy - ((dbl_Sy * dbl_Sy) / _
        (4# * dbl_EndCondition * PI2 * dbl_E)) _
        * ((dbl_L * dbl_L) / (dbl_RadiusOfGyration * dbl_RadiusOfGyration)))
Else
    dbl_CriticalLoad = ((dbl_EndCondition * PI2 * dbl_E * dbl_As) / _
        ((dbl_L * dbl_L) / (dbl_RadiusOfGyration * dbl_RadiusOfGyration)))
End If

dbl_AllowableBuckleLoad = dbl_CriticalLoad / 1.3
dbl_AllowableCompressionLoad = dbl_Sd * dbl_As

If (dbl_AllowableBuckleLoad <= dbl_AllowableCompressionLoad) Then
    dbl_AllowableForce = dbl_AllowableBuckleLoad
Else
    dbl_AllowableForce = dbl_AllowableCompressionLoad
End If

Exit Sub

CalcAllowStemForce_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcAllowStemTensileForce(dbl_Sd As Double, dbl_As As Double, _
    dbl_AllowableForce As Double)
'-----------------------------------------------------------------------------------------
'Reference : Basic Practice Guide - Control Valve Body S/A Design, rev. C, 12/1/97, section 6.3.1
'--------------------------------------------------------------------------------------
'This subroutine is used to calculate the allowable tensile load limit for a valve stem.
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_Sd                  allowable design stress at temperature     psi
'dbl_As                  stress area of stem                        in^2
'-----------------------------------------------------------------------------------------
'Output Variables:
'dbl_AllowableForce      calculated allowable stem load             lbs
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/28 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcAllowStemTensileForce_Error

'***************************
'Parameter validation
'***************************
'neither input parameter can be <=0.0
If (dbl_Sd <= 0#) Then GoTo CalcAllowStemTensileForce_Error
If (dbl_As <= 0#) Then GoTo CalcAllowStemTensileForce_Error

'****************************
'Begin calculations
'****************************
dbl_AllowableForce = dbl_Sd * dbl_As

Exit Sub

CalcAllowStemTensileForce_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcPackingForce(dbl_MaxPressure As Double, dbl_StemDiameter As Double, _
    str_PackingMaterial As String, dbl_PackingHeight As Double, _
    dbl_ForcePacking As Double)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'--------------------------------------------------------------------------------------
'
'This subroutine is used to calculate the force needed to overcome packing friction in
'linear motion valves.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_MaxPressure        maximum inlet or shutoff pressure spec'd by cust.   psi
'dbl_StemDiameter       stem diameter                                       in.
'str_PackingMaterial    material of packing
'dbl_PackingHeight      height of packing set
'                       dbl_PackingHeight = dbl_StemDiameter + 0.6          in.
'
'Output Variables:
'dbl_ForcePacking       force req'd to overcome packing friction            lbs
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/28 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcPackingForce_Error

'***************************
'Declarations
'***************************
Dim dbl_Mu As Double
Dim dbl_PackingPressure As Double

'***************************
'Parameter validation
'***************************
If (dbl_StemDiameter <= 0#) Then GoTo CalcPackingForce_Error
If (str_PackingMaterial = "") Then GoTo CalcPackingForce_Error
If (dbl_PackingHeight <= 0#) Then GoTo CalcPackingForce_Error

'****************************
'Begin calculations
'****************************

'If debug_on = False Then
    Call GetPackingCoef(str_PackingMaterial, dbl_Mu)
'Else
'    dbl_Mu = 0.03
'End If
If (dbl_MaxPressure < 2000) Then
    dbl_PackingPressure = 1000
Else
    dbl_PackingPressure = dbl_MaxPressure / 2#
End If

dbl_ForcePacking = PI * dbl_StemDiameter * dbl_PackingHeight _
    * dbl_Mu * dbl_PackingPressure


'If debug_on Then form21000.txtPackingForce.Text = dbl_ForcePacking


Exit Sub

CalcPackingForce_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcSeatingForce(str_LeakageClass As String, dbl_SeatDiameter As Double, _
    dbl_ForceSeating As Double)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'--------------------------------------------------------------------------------------
'This subroutine is used to calculate the force needed to provide seating force for single seat  linear motion valves.
'---------------------------------------------------------------------------------------
'Input Variables:
'str_LeakageClass   specified leakage class
'                        "IV" or "V" or "VI"
'dbl_SeatDiameter   diameter of seat ring orifice       inches
'
'Output Variables:
'dbl_ForceSeating   force required to provide a seating
'                   load based on selected leakage class    lbs
'                    "IV"   30
'                    "V"    150
'                    "VI"   50
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/28 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcSeatingForce_Error

'***************************
'Declarations
'***************************
Dim dbl_SeatingLoadCoefficient As Double

'***************************
'Parameter validation
'***************************
If (str_LeakageClass = "") Then GoTo CalcSeatingForce_Error
If (dbl_SeatDiameter <= 0#) Then GoTo CalcSeatingForce_Error

'****************************
'Begin calculations
'****************************

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'If debug_on = False Then
    Call GetSeatingCoef(str_LeakageClass, dbl_SeatingLoadCoefficient)
'Else
'    Select Case Trim(str_LeakageClass)
'        Case "IV"
'            dbl_SeatingLoadCoefficient = 30
'        Case "V"
'            dbl_SeatingLoadCoefficient = 150
'        Case "VI"
'            dbl_SeatingLoadCoefficient = 50
'   End Select
'End If
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

dbl_ForceSeating = dbl_SeatDiameter * PI * dbl_SeatingLoadCoefficient

Exit Sub

CalcSeatingForce_Error:
'Placeholder.  How do we want to handle error messages?

End Sub


Public Sub CalcSingleSeatUnbalancedForceFTO(dbl_PressureDrop As Double, _
    dbl_SeatDiameter As Double, dbl_ForceUnbalanced As Double)

'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the force needed to provide force to overcome
'unbalanced force for single seat  linear motion valves.
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_pressureDrop       pressure drop acting on plug        lbs/in^2
'dbl_seatDiameter       diameter of seat ring orifice       inches
'
'Output Variables:
'dbl_forceUnbalanced    force required to overcome pressure
'                       drop acting on plug         lbs
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/28 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcSingleSeatUnblanacedForceFTO_Error

'***************************
'Parameter validation
'***************************
If (dbl_SeatDiameter <= 0#) Then GoTo CalcSingleSeatUnblanacedForceFTO_Error

'****************************
'Begin calculations
'****************************
dbl_ForceUnbalanced = PI * dbl_PressureDrop * dbl_SeatDiameter * dbl_SeatDiameter / 4

Exit Sub

CalcSingleSeatUnblanacedForceFTO_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcATOActuatorForce(dbl_AreaInitial As Double, dbl_PInitial As Double, dbl_PFinal As Double, _
        dbl_ActuatorFriction As Double, dbl_RequiredSupplyPressure As Double, dbl_ActuatorForce As Double)
        
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the force output of an air to open linear motion actuator.
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_AreaInitial        actuator area at initial position (valve closed)        in^2
'dbl_PIinitial          initial pressure provided by spring                     lb/in^2
'                           setting or air pressure (preload)
'dbl_PFinal             final pressure provided by spring
'dbl_ActuatorFriction   force necessary to overcome actuator friction           lb
'
'Output Variables:
'dbl_RequiredSupplyPressure
'dbl_ActuatorForce      output force of actuator            lb
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/28 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcATOActuatorForce_Error

'***************************
'Parameter validation
'***************************
If (dbl_AreaInitial <= 0) Then GoTo CalcATOActuatorForce_Error

'****************************
'Begin calculations
'****************************
dbl_ActuatorForce = (dbl_PInitial * dbl_AreaInitial) - dbl_ActuatorFriction
dbl_RequiredSupplyPressure = dbl_PFinal + 5#

Exit Sub

CalcATOActuatorForce_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcATCActuatorForce(dbl_AreaFinal As Double, dbl_PFinal As Double, _
    dbl_PSupply As Double, dbl_ActuatorFriction As Double, _
    dbl_ActuatorForce As Double)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'--------------------------------------------------------------------------------------
'This subroutine is used to calculate the force output of an air to close linear motion actuator.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_AreaFinal          actuator area at final position (valve closed)          in^2
'dbl_PFinal             final pressure provised by spring                       lb/in^2
'                       setting or air pressure (preload)
'dbl_PSupply            supply pressure to actuator                             lb/in^2
'dbl_ActuatorFriction   force necessary to overcome actuator friction           lb

'Output Variables:
'dbl_ActuatorForce      output force of actuator                                lb
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/29 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcATCActuatorForce_Error

'***************************
'Parameter validation
'***************************
If (dbl_AreaFinal < 0#) Then GoTo CalcATCActuatorForce_Error
If (dbl_PSupply <= dbl_PFinal) Then GoTo CalcATCActuatorForce_Error
If (dbl_PFinal < 0#) Then GoTo CalcATCActuatorForce_Error
If (dbl_ActuatorFriction < 0#) Then GoTo CalcATCActuatorForce_Error

'****************************
'Begin calculations
'****************************
dbl_ActuatorForce = (dbl_AreaFinal * (dbl_PSupply - dbl_PFinal) - dbl_ActuatorFriction)

Exit Sub

CalcATCActuatorForce_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcAllowDPSingleStageUnbalFTO(dbl_ActuatorForce As Double, dbl_ForceSeating As Double, _
    dbl_ForceFriction As Double, dbl_SeatDiameter As Double, dbl_AllowDP As Double)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the force output of an air to close linear motion actuator.
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_ActuatorForce      output force of Actuator                                lbs
'dbl_ForceSeating       required seating force                                  lbs
'dbl_ForceFriction      force needed to overcome packing friction               lbs
'dbl_SeatDiameter       diameter of seat ring orifice                           in
'
'Output Variables:
'dbl_AllowDP            allowable pressure drop for valve and actuator combo    lb/in^2
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/29 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcAllowDPSingleStageUnbalFTO_Error

'***************************
'Parameter validation
'***************************
If (dbl_ActuatorForce < 0#) Then GoTo CalcAllowDPSingleStageUnbalFTO_Error
If (dbl_ForceSeating < 0#) Then GoTo CalcAllowDPSingleStageUnbalFTO_Error
If (dbl_ForceFriction < 0#) Then GoTo CalcAllowDPSingleStageUnbalFTO_Error
If (dbl_SeatDiameter <= 0#) Then GoTo CalcAllowDPSingleStageUnbalFTO_Error

'****************************
'Begin calculations
'****************************
dbl_AllowDP = (dbl_ActuatorForce - dbl_ForceSeating - dbl_ForceFriction) * _
    4 / (PI * dbl_SeatDiameter * dbl_SeatDiameter)
    
Exit Sub

CalcAllowDPSingleStageUnbalFTO_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcValveSealRingFrictionForce(dbl_SealCoefficient As Double, dbl_SeatDiameter As Double, _
    dbl_SealWidth As Double, int_NumberOfSeals As Integer, dbl_SealForce As Double)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the force required to cvercome plug seal friction for linear motion actuator.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_SealCoefficient    balanced seal ring friction coefficient
'                           teflon 0.04
'                           graphite 0.21
'dbl_SealDiameter       diameter of seal ring                       inches
'dbl_SealWidth          width of seal ring                          inches
'int_NumberOfSeals      number of seals used
'
'Output Variables:
'dbl_SealForce          force needed to overcome seal ring friction lbs
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/29 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcValveSealRingFrictionForce_Error

'***************************
'Parameter validation
'***************************
If (dbl_SealCoefficient <= 0#) Then GoTo CalcValveSealRingFrictionForce_Error
If (dbl_SeatDiameter <= 0#) Then GoTo CalcValveSealRingFrictionForce_Error
If (dbl_SealWidth <= 0#) Then GoTo CalcValveSealRingFrictionForce_Error
If (int_NumberOfSeals < 0) Then GoTo CalcValveSealRingFrictionForce_Error

'****************************
'Begin calculations
'****************************
dbl_SealForce = (dbl_SealCoefficient * PI * dbl_SeatDiameter * dbl_SealWidth * _
    int_NumberOfSeals) / 2#

Exit Sub

CalcValveSealRingFrictionForce_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcPositionerLossForce(dbl_PSupply As Double, dbl_AreaActuator As Double, _
    dbl_PositionerLoss As Double)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the force necessary to compensate for positioner loss.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_PSupply            supply pressure                                 lb/in^2
'dbl_AreaActuator       effective area of actuator                      lb
'
'Output Variables:
'dbl_PositionerLoss     force necessary to overcome positioner loss
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/29 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcPositionerLoss_Error

'***************************
'Parameter validation
'***************************
If (dbl_PSupply <= 0#) Then GoTo CalcPositionerLoss_Error
If (dbl_AreaActuator <= 0#) Then GoTo CalcPositionerLoss_Error

'****************************
'Begin calculations
'****************************
dbl_PositionerLoss = 0.01 * dbl_PSupply * dbl_AreaActuator

Exit Sub

CalcPositionerLoss_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcSingleSeatUnbalancedForceFTC(dbl_PressureDrop As Double, dbl_SeatDiameter As Double, _
    dbl_StemDiameter As Double, dbl_EffectiveOrificeArea As Double, _
    dbl_ForceUnbalancedClosed As Double, dbl_ForceUnbalancedOpen As Double)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the force needed to provide force to overcome unbalanced force for single seat  linear motion valves.
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_PressureDrop               pressure drop acting on plug            lbs/in^2
'dbl_SeatDiameter               diameter of seat ring orifice           inches
'dbl_StemDiameter               diameter of plug stem                   inches
'dbl_EffectiveOrificeArea
'
'Output Variables:
'dbl_ForceUnbalancedClosed      force required to overcome pressure
'                                   drop acting on plug - closed        lbs
'dbl_ForceUnbalancedOpen        force required to overcome pressure
'                                   drop acting on plug - open      lbs
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/29 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcSingleSeatUnbalancedForceFTC_Error

'***************************
'Parameter validation
'***************************
If (dbl_PressureDrop < 0#) Then GoTo CalcSingleSeatUnbalancedForceFTC_Error
If (dbl_SeatDiameter <= 0#) Then GoTo CalcSingleSeatUnbalancedForceFTC_Error
If (dbl_StemDiameter <= 0#) Then GoTo CalcSingleSeatUnbalancedForceFTC_Error

'****************************
'Begin calculations
'****************************
dbl_ForceUnbalancedClosed = (dbl_SeatDiameter * dbl_SeatDiameter - dbl_StemDiameter * dbl_StemDiameter) _
    * dbl_PressureDrop * PI / 4

dbl_ForceUnbalancedOpen = (dbl_SeatDiameter * dbl_SeatDiameter - dbl_StemDiameter * dbl_StemDiameter) _
    * dbl_PressureDrop * dbl_EffectiveOrificeArea * PI / 4
    
Exit Sub

CalcSingleSeatUnbalancedForceFTC_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcAllowDPSingleStageUnbalFTC_ATO(dbl_AreaInitial As Double, dbl_PInitial As Double, _
    dbl_AreaFinal As Double, dbl_PFinal As Double, dbl_SpringRate As Double, dbl_PSupply As Double, _
    dbl_ForcePacking As Double, dbl_SeatDiameter As Double, dbl_StemDiameter As Double, _
    dbl_EffectiveOrificeArea As Double, dbl_FDynamic As Double, dbl_AllowDPStability As Double, _
    dbl_AllowDPClosed As Double, dbl_AllowDPOpen As Double)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the allowable pressure drop for unbalanced valves, FTC.
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_AreaInitial            actuator area at initial position               in^2
'dbl_PInitial               init. pressure prov'd by spring
'                             setting or air pressure (preload)             lb/in^2
'dbl_AreaFinal              actuator area at final position                 in^2
'dbl_PFinal                 final pressure provised by spring
'                             setting or air pressure (preload)             lb/in^2
'dbl_SpringRate             spring rate of actuator                         lb/in
'dbl_PSupply                supply pressure                                 lb/in
'dbl_ForcePacking           force needed to overcome packing friction       lb
'dbl_SeatDiameter           diameter of seat ring orifice                   in
'dbl_StemDiameter           diameter of plug stem                           in
'dbl_EffectiveOrificeArea
'dbl_FDynamic               stability factor
'--------------------------------------------------------------------------------------
'Output Variables:
'dbl_AllowDPstability       allowable pressure drop for valve               lb/in^2
'                             & actuator - stability
'dbl_AllowDPclosed          allowable pressure drop - valve closed          lb/in^2
'dbl_AllowDPopen            allowable pressure drop - valve open            lb/in^2
'--------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/29 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'--------------------------------------------------------------------------------------
On Error GoTo CalcAllowDPSingleStageUnbalFTC_ATO_Error

'***************************
'Parameter validation
'***************************
If (dbl_AreaInitial <= 0#) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATO_Error
If (dbl_PInitial < 0#) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATO_Error
If (dbl_AreaFinal <= 0#) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATO_Error
If (dbl_PFinal < 0#) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATO_Error
If (dbl_SpringRate < 0#) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATO_Error
If (dbl_SeatDiameter <= 0#) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATO_Error
If (dbl_StemDiameter <= 0#) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATO_Error
'the following would cause divide by 0 if not caught
If (dbl_FDynamic = 0#) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATO_Error
If (dbl_SeatDiameter <= dbl_StemDiameter) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATO_Error
If (dbl_EffectiveOrificeArea <= 0#) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATO_Error
'the following would cause the wrong sign
If (dbl_PSupply < dbl_PInitial) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATO_Error

'****************************
'Begin calculations
'****************************
dbl_AllowDPStability = dbl_SpringRate / dbl_FDynamic

dbl_AllowDPClosed = (((dbl_PSupply - dbl_PInitial) * dbl_AreaInitial) - dbl_ForcePacking) _
    / ((dbl_SeatDiameter * dbl_SeatDiameter - dbl_StemDiameter * dbl_StemDiameter) * PI / 4)
    
dbl_AllowDPOpen = (((dbl_PSupply - dbl_PFinal) * dbl_AreaFinal) - dbl_ForcePacking) _
    / ((dbl_SeatDiameter * dbl_SeatDiameter - dbl_StemDiameter * dbl_StemDiameter) * PI _
    * dbl_EffectiveOrificeArea / 4)

Exit Sub

CalcAllowDPSingleStageUnbalFTC_ATO_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcAllowDPSingleStageUnbalFTC_ATC(dbl_AreaInitial As Double, dbl_PInitial As Double, _
    dbl_AreaFinal As Double, dbl_PFinal As Double, dbl_SpringRate As Double, _
    dbl_ForcePacking As Double, dbl_SeatDiameter As Double, dbl_StemDiameter As Double, _
    dbl_FDynamic As Double, dbl_EffectiveOrificeArea As Double, dbl_AllowDPStability As Double, _
    dbl_AllowDPClosed As Double, dbl_AllowDPOpen As Double)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'--------------------------------------------------------------------------------------
'This subroutine is used to calculate the allowable pressure drop for unbalanced valves, FTC.
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_AreaInitial            actuator area at initial position               in^2
'dbl_PInitial               init. pressure prov'd by spring
'                             setting or air pressure (preload)             lb/in^2
'dbl_AreaFinal              actuator area at final position                 in^2
'dbl_PFinal                 final pressure provised by spring
'                             setting or air pressure (preload)             lb/in^2
'dbl_SpringRate             spring rate of actuator                         lb/in
'dbl_ForcePacking           force needed to overcome packing friction       lb
'dbl_SeatDiameter           diameter of seat ring orifice                   in
'dbl_StemDiameter           diameter of plug stem                           in
'dbl_FDynamic               stability factor
'dbl_EffectiveOrificeArea
'--------------------------------------------------------------------------------------
'Output Variables:
'dbl_AllowDPstability       allowable pressure drop for valve               lb/in^2
'                             & actuator - stability
'dbl_AllowDPclosed          allowable pressure drop - valve closed          lb/in^2                                valve closed
'dbl_AllowDPopen            allowable pressure drop - valve open            lb/in^2
'--------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/29 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'--------------------------------------------------------------------------------------
On Error GoTo CalcAllowDPSingleStageUnbalFTC_ATC_Error

'***************************
'Parameter validation
'***************************
If (dbl_AreaInitial <= 0#) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATC_Error
If (dbl_PInitial < 0#) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATC_Error
If (dbl_AreaFinal <= 0#) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATC_Error
If (dbl_PFinal < 0#) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATC_Error
If (dbl_SpringRate < 0#) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATC_Error
If (dbl_SeatDiameter <= 0#) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATC_Error
If (dbl_StemDiameter <= 0#) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATC_Error
'the following would cause divide by 0 if not caught
If (dbl_FDynamic = 0#) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATC_Error
If (dbl_SeatDiameter <= dbl_StemDiameter) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATC_Error
If (dbl_EffectiveOrificeArea <= 0#) Then GoTo CalcAllowDPSingleStageUnbalFTC_ATC_Error

'****************************
'Begin calculations
'****************************
dbl_AllowDPStability = dbl_SpringRate / dbl_FDynamic

dbl_AllowDPClosed = ((dbl_PFinal * dbl_AreaFinal) - dbl_ForcePacking) _
    / ((dbl_SeatDiameter * dbl_SeatDiameter - dbl_StemDiameter * dbl_StemDiameter) _
    * PI / 4)

dbl_AllowDPOpen = ((dbl_PInitial * dbl_AreaInitial) - dbl_ForcePacking) _
    / ((dbl_SeatDiameter * dbl_SeatDiameter - dbl_StemDiameter * dbl_StemDiameter) _
    * PI * dbl_EffectiveOrificeArea / 4)
    
Exit Sub

CalcAllowDPSingleStageUnbalFTC_ATC_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcDoubleStageForceFTO(dbl_PressureDrop As Double, dbl_SeatDiameter As Double, _
    dbl_PlugDiameter As Double, dbl_ForceUnbalancedClosed As Double, _
    dbl_ForceUnbalancedOpen As Double)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'--------------------------------------------------------------------------------------
'This subroutine is used to calculate the force needed to provide force to overcome
'unbalanced force for single seat, linear motion valves.
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_PressureDrop               pressure drop acting on plug            lbs/in^2
'dbl_SeatDiameter               diameter of seat ring orifice           in
'dbl_PlugDiameter               diameter of plug                        in
'--------------------------------------------------------------------------------------
'Output Variables:
'dbl_ForceUnbalancedClosed      force required to overcome pressure
'                                   drop acting on plug - closed        lbs
'dbl_ForceUnbalancedOpen        force required to overcome pressure
'                                   drop acting on plug - open          lbs
'--------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/29 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'--------------------------------------------------------------------------------------

On Error GoTo CalcDoubleStageForceFTO_Error

'***************************
'Parameter validation
'***************************
If (dbl_PressureDrop < 0#) Then GoTo CalcDoubleStageForceFTO_Error
If (dbl_SeatDiameter <= 0#) Then GoTo CalcDoubleStageForceFTO_Error
If (dbl_PlugDiameter <= 0#) Then GoTo CalcDoubleStageForceFTO_Error

'****************************
'Begin calculations
'****************************
dbl_ForceUnbalancedClosed = dbl_PressureDrop * dbl_SeatDiameter * dbl_SeatDiameter * PI / 4

dbl_ForceUnbalancedOpen = (dbl_SeatDiameter * dbl_SeatDiameter + dbl_PlugDiameter * dbl_PlugDiameter) _
    * dbl_PressureDrop * PI / 8
    
Exit Sub

CalcDoubleStageForceFTO_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcAllowDPDoubleStageFTO(dbl_ActuatorForce As Double, dbl_SpringRate As Double, _
    dbl_ForcePacking As Double, dbl_ForceSeating As Double, dbl_SeatDiameter As Double, _
    dbl_PlugDiameter As Double, dbl_FDynamic As Double, dbl_AllowDPStability As Double, _
    dbl_AllowDPClosed As Double, dbl_AllowDPOpen As Double)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'--------------------------------------------------------------------------------------
'This subroutine is used to calculate the allowable pressure drop for unbalanced valves, FTC.
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_ActuatorForce      output force of valve                                       lbs
'dbl_SpringRate         spring rate of actuator                                     lb/in.
'dbl_ForcePacking       force required to overcome packing friction                 lbs
'dbl_ForceSeating       " to provide seating load for selected leakage class        lbs
'dbl_SeatDiameter       diameter of seat ring orifice                               in.
'dbl_PlugDiameter       diameter of plug                                            in.
'dbl_FDynamic           stability factor
'--------------------------------------------------------------------------------------
'Output Variables:
'dbl_AllowDPstability   allowable pressure drop for valve                   lb/in^2
'                           and actuator combination - stability
'dbl_AllowDPclosed      allowable pressure drop for valve                   lb/in^2
'                           and actuator combination - closed
'dbl_AllowDPopen        allowable pressure drop for valve                   lb/in^2
'                           and actuator combination - open
'--------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/29 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'--------------------------------------------------------------------------------------
On Error GoTo CalcAllowDPDoubleStageFTO_Error

'***************************
'Parameter validation
'***************************
If (dbl_ActuatorForce < 0#) Then GoTo CalcAllowDPDoubleStageFTO_Error
If (dbl_SpringRate < 0#) Then GoTo CalcAllowDPDoubleStageFTO_Error
If (dbl_ForcePacking < 0#) Then GoTo CalcAllowDPDoubleStageFTO_Error
If (dbl_ForceSeating < 0#) Then GoTo CalcAllowDPDoubleStageFTO_Error
If (dbl_SeatDiameter <= 0#) Then GoTo CalcAllowDPDoubleStageFTO_Error
If (dbl_PlugDiameter <= 0#) Then GoTo CalcAllowDPDoubleStageFTO_Error

'****************************
'Begin calculations
'****************************
dbl_AllowDPStability = dbl_SpringRate / dbl_FDynamic

dbl_AllowDPClosed = (dbl_ActuatorForce - dbl_ForcePacking - dbl_ForceSeating) _
    / (dbl_SeatDiameter * dbl_SeatDiameter * PI / 4)
    
dbl_AllowDPOpen = (dbl_ActuatorForce - dbl_ForcePacking) / _
    ((dbl_SeatDiameter * dbl_SeatDiameter + dbl_PlugDiameter * dbl_PlugDiameter) * PI / 8)
    
Exit Sub

CalcAllowDPDoubleStageFTO_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Function ConvertPressureToPSIG(dbl_Pressure As Double, _
    str_UM As String) As Double
'--------------------------------------------------------------------------------------
'Reference :British Standards Institution: Conversion factors and tables.
'--------------------------------------------------------------------------------------
'This function converts pressure specified in given units of measure to psig.
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_Pressure           the pressure to be converted
'str_UM                 the unit of measure of the input pressure
'dbl_PressureAtmBarA    the atmospheric pressure in Bar A
'--------------------------------------------------------------------------------------
'Output Variables:
'function               pressure in psig
'
'Constants:
'Use theUM to find values for multiplier, adder.
'
'Input Units     Multiplier      Adder
'"psi g"        1.0             0.0
'"kgf/cm2 g"    14.22333        0.0
'"kPa g"        0.145038        0.0
'"MPa g"        145.038         0.0
'"bar g"        14.5038         0.0
'"psi a"        1.0             -14.6959
'"kgf/cm2 a"    14.22333        -14.6959
'"kPa a"        0.145038        -14.6959
'"MPa a"        145.038         -14.6959
'"bar a"        14.5038         -14.6959
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/29 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo ConvertPressureToPSIG_Error

'****************************
'Begin calculations
'****************************
Select Case str_UM
    Case "psi g", "psig"
        ConvertPressureToPSIG = dbl_Pressure
    Case "kgf/cm2 g"
        ConvertPressureToPSIG = dbl_Pressure * 14.22333
    Case "kPa g"
        ConvertPressureToPSIG = dbl_Pressure * 0.145038
     Case "MPa g"
        ConvertPressureToPSIG = dbl_Pressure * 145.038
     Case "bar g"
        ConvertPressureToPSIG = dbl_Pressure * 14.5038
     Case "psi a"
        ConvertPressureToPSIG = dbl_Pressure * 1# - 14.6959
     Case "kgf/cm2 a"
        ConvertPressureToPSIG = dbl_Pressure * 14.22333 - 14.6959
     Case "kPa a"
        ConvertPressureToPSIG = dbl_Pressure * 0.145038 - 14.6959
     Case "MPa a"
        ConvertPressureToPSIG = dbl_Pressure * 145.038 - 14.6959
     Case "bar a"
        ConvertPressureToPSIG = dbl_Pressure * 14.5038 - 14.6959
    Case Else
        GoTo ConvertPressureToPSIG_Error
End Select

Exit Function

ConvertPressureToPSIG_Error:
'Placeholder.  How do we want to handle error messages?

End Function

Public Function ConvertPressureFromPSIG(dbl_Pressure As Double, _
    str_UM As String) As Double
'--------------------------------------------------------------------------------------
'Reference :British Standards Institution: Conversion factors and tables.
'--------------------------------------------------------------------------------------
'This function converts pressure to specified iunits of measure from psig.
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_Pressure           the pressure to be converted
'str_UM                 the U/M of the output pressure
'
'Output Variables:
'function               pressure in psig
'
'constants:
'Use theUM to find values for multiplier, adder.
'
'Input Units     Multiplier      Adder
'"psi g"        1.0             0.0
'"kgf/cm2 g"    14.22333        0.0
'"kPa g"        0.145038        0.0
'"MPa g"        145.038         0.0
'"bar g"        14.5038         0.0
'"psi a"        1.0             -14.6959
'"kgf/cm2 a"    14.22333        -14.6959
'"kPa a"        0.145038        -14.6959
'"MPa a"        145.038         -14.6959
'"bar a"        14.5038         -14.6959
'--------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/29 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'--------------------------------------------------------------------------------------
On Error GoTo ConvertPressureFromPSIG_Error

'***************************
'Declarations
'***************************

'***************************
'Parameter validation
'***************************

'****************************
'Begin calculations
'****************************
Select Case str_UM
    Case "psi g", "psig"
        ConvertPressureFromPSIG = dbl_Pressure
    Case "kgf/cm2 g"
        ConvertPressureFromPSIG = dbl_Pressure / 14.22333
    Case "kPa g"
        ConvertPressureFromPSIG = dbl_Pressure / 0.145038
     Case "MPa g"
        ConvertPressureFromPSIG = dbl_Pressure / 145.038
     Case "bar g"
        ConvertPressureFromPSIG = dbl_Pressure / 14.5038
     Case "psi a"
        ConvertPressureFromPSIG = dbl_Pressure + 14.6959
     Case "kgf/cm2 a"
        ConvertPressureFromPSIG = (dbl_Pressure + 14.6959) / 14.22333
     Case "kPa a"
        ConvertPressureFromPSIG = (dbl_Pressure + 14.6959) / 0.145038
     Case "MPa a"
        ConvertPressureFromPSIG = (dbl_Pressure + 14.6959) / 145.038
     Case "bar a"
        ConvertPressureFromPSIG = (dbl_Pressure + 14.6959) / 14.5038
    Case Else
        GoTo ConvertPressureFromPSIG_Error
End Select

Exit Function

ConvertPressureFromPSIG_Error:
'Placeholder.  How do we want to handle error messages?

End Function

Public Function ConvertTemperatureToDegF(dbl_Temperature As Double, str_UM As String)
'--------------------------------------------------------------------------------------
'Reference :British Standards Institution: Conversion factors and tables.
'--------------------------------------------------------------------------------------
'This function converts temperature from specified units of measure to degrees F
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_Temperature     input temperature in UM units
'str_UM              units of measure
'--------------------------------------------------------------------------------------
'Output Variables:
'        function    temperature in degrees F
'--------------------------------------------------------------------------------------
'Equations:
'ConvertTemperatureToDegF = (9/5) degrees K - 459.67
'ConvertTemperatureToDegF = (9/5) degrees C + 32
'ConvertTemperatureToDegF = degrees R - 459.67
'--------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/29 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'--------------------------------------------------------------------------------------
On Error GoTo ConvertTemperatureToDegF_Error

'****************************
'Begin calculations
'****************************
Select Case str_UM
    Case "deg K"
        ConvertTemperatureToDegF = 1.8 * dbl_Temperature - 459.67
    Case "deg C"
        ConvertTemperatureToDegF = 1.8 * dbl_Temperature + 32#
    Case "deg R"
        ConvertTemperatureToDegF = dbl_Temperature - 459.67
    Case "deg F"
        ConvertTemperatureToDegF = dbl_Temperature
    Case Else
        GoTo ConvertTemperatureToDegF_Error
End Select

Exit Function

ConvertTemperatureToDegF_Error:
'Placeholder.  How do we want to handle error messages?

End Function

Public Function ConvertTemperatureFromDegF(dbl_Temperature As Double, str_UM As String)
'--------------------------------------------------------------------------------------
'Reference :British Standards Institution: Conversion factors and tables.
'--------------------------------------------------------------------------------------
'This function converts temperature to specified units of measure from degrees F
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_Temperature     input temperature in degrees F
'str_UM              units of measure

'Output Variables:
'        function        temperature in UM units
'--------------------------------------------------------------------------------------
'Equations:
'ConvertTemperatureFromDegF = (5/9) (degrees F + 459.67) (degrees K)
'ConvertTemperatureFromDegF = (5/9) (degrees F - 32)     (degrees C)
'ConvertTemperatureFromDegF = degrees F + 459.67     (degrees R)
'--------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/29 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'--------------------------------------------------------------------------------------
On Error GoTo ConvertTemperatureFromDegF_Error

'****************************
'Begin calculations
'****************************
Select Case str_UM
    Case "deg K"
        ConvertTemperatureFromDegF = 0.55555556 * (dbl_Temperature + 459.67)
    Case "deg C"
        ConvertTemperatureFromDegF = 0.55555556 * (dbl_Temperature - 32#)
    Case "deg R"
        ConvertTemperatureFromDegF = dbl_Temperature + 459.67
    Case "deg F"
        ConvertTemperatureFromDegF = dbl_Temperature
    Case Else
        GoTo ConvertTemperatureFromDegF_Error
End Select

Exit Function

ConvertTemperatureFromDegF_Error:
'Placeholder.  How do we want to handle error messages?

End Function

Public Function ConvertLengthToInches(dbl_Length As Double, str_UM As String) As Double
'--------------------------------------------------------------------------------------
'Reference :British Standards Institution: Conversion factors and tables.
'--------------------------------------------------------------------------------------
'This function converts length from specified units of measure to inches
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_Length                     the length in UM units
'str_UM                         the units of measure
'
'Output Variables:
'ConvertLengthToInches          the length in inches
'--------------------------------------------------------------------------------------
'Equations:
'"mm"    lengthOut = dbl_Length * .039370
'"m" lengthOut = dbl_Length * 39.3701
'"ft"    lengthOut = dbl_Length * 12
'"in"    lengthOut = dbl_Length
'"cm"    lengthOut = dbl_Length * .393701
'--------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/29 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'--------------------------------------------------------------------------------------'--------------------------------------------------------------------------------------
On Error GoTo ConvertLengthToInches_Error

'***************************
'Parameter validation
'***************************
If (dbl_Length < 0#) Then GoTo ConvertLengthToInches_Error

'****************************
'Begin calculations
'****************************
Select Case str_UM
    Case "mm"
        ConvertLengthToInches = dbl_Length * 0.03937
    Case "m"
        ConvertLengthToInches = dbl_Length * 39.3701
    Case "ft"
        ConvertLengthToInches = dbl_Length * 12
    Case "in"
        ConvertLengthToInches = dbl_Length
    Case "cm"
        ConvertLengthToInches = dbl_Length * 0.393701
    Case Else
        GoTo ConvertLengthToInches_Error
End Select
Exit Function

ConvertLengthToInches_Error:
'Placeholder.  How do we want to handle error messages?

End Function

Public Function ConvertLengthFromInches(dbl_Length As Double, str_UM As String) As Double
'--------------------------------------------------------------------------------------
'Reference :British Standards Institution: Conversion factors and tables.
'---------------------------------------------------------------------------------------
'This function converts length to specified units of measure from inches
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_Length                 the length in inches
'str_UM                     the units of measure
'
'Output Variables:
'ConvertLengthToInches      the length in UM units
'
'---------------------------------------------------------------------------------------
'Equations:
'"mm"    lengthOut = dbl_Length / 0.039370
'"m" lengthOut = dbl_Length / 39.3701
'"ft"    lengthOut = dbl_Length / 12
'"in"    lengthOut = dbl_Length
'"cm"    lengthOut = dbl_Length / 0.393701
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 5/29 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo ConvertLengthFromInches_Error

'***************************
'Parameter validation
'***************************
If (dbl_Length < 0#) Then GoTo ConvertLengthFromInches_Error

'****************************
'Begin calculations
'****************************
Select Case str_UM
    Case "mm"
        ConvertLengthFromInches = dbl_Length / 0.03937
    Case "m"
        ConvertLengthFromInches = dbl_Length / 39.3701
    Case "ft"
        ConvertLengthFromInches = dbl_Length / 12
    Case "in"
        ConvertLengthFromInches = dbl_Length
    Case "cm"
        ConvertLengthFromInches = dbl_Length / 0.393701
    Case Else
        GoTo ConvertLengthFromInches_Error
End Select
Exit Function

ConvertLengthFromInches_Error:
'Placeholder.  How do we want to handle error messages?

End Function

Public Sub CalcSpringRate(str_SpringRange As String, dbl_AreaInitial As Double, _
    dbl_PInitial As Double, dbl_AreaFinal As Double, dbl_PFinal As Double, _
    str_SpecifiedCalibPressureUM As String, dbl_Travel As Double, _
    str_SpecifiedLengthUM As String, dbl_SpringRate As Double)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to set variables Pinitial and Pfinal and then calculate the spring rate of an actuator.
'---------------------------------------------------------------------------------------
'Input Variables:
'
'str_springRange            spring range of actuator                user defined
'dbl_areaInitial            actuator area at initial position       in^2
'dbl_areaFinal              actuator area at final position         in^2
'str_specifiedCalibPressureUM   Pressure units
'dbl_travel                 actuator output travel                  user defined
'str_specifiedLengthUM      Length units

'Output Variables:
'
'dbl_Pinitial               initial pressure provised by spring
'                            setting or air pressure (preload)       lb/in^2
'dbl_Pfinal                 final pressure provised by spring
'                            setting or air pressure (preload)       lb/in^2
'dbl_springRate             springRate of actuator                   lb/in
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/1 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcSpringRate_Error

'***************************
'Declarations
'***************************
Dim int_ThisDash As Integer
Dim str_LocalSpringRange As String

'***************************
'Parameter validation
'***************************
If (str_SpringRange = "") Then GoTo CalcSpringRate_Error
If (dbl_AreaInitial <= 0#) Then GoTo CalcSpringRate_Error
If (dbl_AreaFinal <= 0#) Then GoTo CalcSpringRate_Error
If (dbl_Travel <= 0#) Then GoTo CalcSpringRate_Error

'****************************
'Begin calculations
'****************************
str_LocalSpringRange = Trim(str_SpringRange)
int_ThisDash = InStr(str_LocalSpringRange, "-")

'the spring range will always come in in psi
'dbl_PInitial = ConvertPressureToPSIG(CDbl(Left(str_LocalSpringRange, int_ThisDash - 1)), _
'    str_SpecifiedCalibPressureUM)
dbl_PInitial = CDbl(Left(str_LocalSpringRange, int_ThisDash - 1))

'dbl_PFinal = ConvertPressureToPSIG(CDbl(Mid(str_LocalSpringRange, int_ThisDash + 1)), _
'    str_SpecifiedCalibPressureUM)
dbl_PFinal = CDbl(Mid(str_LocalSpringRange, int_ThisDash + 1))

dbl_SpringRate = (dbl_PFinal * dbl_AreaFinal - dbl_PInitial * dbl_AreaInitial) / _
    ConvertLengthToInches(dbl_Travel, str_SpecifiedLengthUM)
   
   
 
   
Exit Sub

CalcSpringRate_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub GetANSIPressure(str_BodyMaterial As String, dbl_TestTemperature As Double, _
    str_SelectedRating As String, dbl_AllowANSIPressure As Double)
'--------------------------------------------------------------------------------------
'Reference : ANSI B16.34 - 1996 Valves - Flanged and Buttwelded Ends
'Solutions reference : 21000.frm, subroutine: getANSIpress
'---------------------------------------------------------------------------------------
'This subroutine is used to determine the maximum allowable working pressure with known body material and rating class.
'---------------------------------------------------------------------------------------
'Input Variables:
'
'str_BodyMatl               body material
'str_SelectedRating         selected ANSI rating class
'dbl_TestTemperature        the maximum temperature entered in
'                               service conditions and design conditions
'
'Output Variables:
'
'dbl_AllowANSIPressure      maximum pressure allowed for material       psi g
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/1 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo GetANSIPressure_Error

'***************************
'Declarations
'***************************
Dim dbl_LowerTemperature As Double
Dim dbl_UpperTemperature As Double
Dim dbl_LowerPressure As Double
Dim dbl_UpperPressure As Double
Dim rset_B16_34 As Recordset
Dim rset_TemperatureTable As Recordset

Dim str_DatabasePath As String      'path to the database file

Dim str_MaterialQuery As String     'query for material type
Dim str_Table As String             'the name of the proper table to get data
Dim str_PressureField As String      'name of pressure field
Dim str_TemperatureQuery As String  'query of temp vs pressure class data

'***************************
'Parameter validation
'***************************
If (str_BodyMaterial = "") Then GoTo GetANSIPressure_Error
If (str_SelectedRating = "") Then GoTo GetANSIPressure_Error

'****************************
'Begin calculations
'****************************

'get the proper table given the material type
str_MaterialQuery = "Select B16_34_table from BodyMaterials where material='?'"
str_MaterialQuery = ReplaceStr(str_MaterialQuery, "?", str_BodyMaterial)

Set rset_B16_34 = g_db_Gensz.OpenRecordset(str_MaterialQuery, dbReadOnly)

'there can only be one record
If (rset_B16_34.RecordCount <> 1) Then
    str_Table = "ceiling"
Else
    rset_B16_34.MoveFirst
    str_Table = Trim(rset_B16_34.Fields(0))
    If (str_Table = "") Then GoTo GetANSIPressure_Error
End If

'get the pressure column from the ansi rating
str_PressureField = Trim(str_SelectedRating)
str_PressureField = Mid(str_PressureField, InStr(1, str_PressureField, " ") + 1)
str_PressureField = "pressure" + str_PressureField
If (str_PressureField = "") Then GoTo GetANSIPressure_Error

'get the temp data from the database
str_TemperatureQuery = "Select tempF, [?] from [?]"
str_TemperatureQuery = ReplaceStr(str_TemperatureQuery, "?", str_PressureField, _
    str_Table)

Set rset_TemperatureTable = g_db_Gensz.OpenRecordset(str_TemperatureQuery)
If (rset_TemperatureTable.RecordCount = 0) Then GoTo GetANSIPressure_Error

'find the bounding temperatures from the table
dbl_LowerTemperature = -99999#
dbl_UpperTemperature = -99999#
rset_TemperatureTable.MoveFirst
Do Until rset_TemperatureTable.EOF
    dbl_LowerTemperature = dbl_UpperTemperature
    dbl_LowerPressure = dbl_UpperPressure
    dbl_UpperTemperature = CDbl(rset_TemperatureTable.Fields(0))
    dbl_UpperPressure = CDbl(rset_TemperatureTable.Fields(1))
    If (dbl_UpperTemperature > dbl_TestTemperature) Then Exit Do
    rset_TemperatureTable.MoveNext
Loop

rset_TemperatureTable.Close
If (dbl_LowerTemperature = -99999#) Then
    'first temp was above the test temp so we can't interpolate
    dbl_AllowANSIPressure = dbl_UpperPressure
Else
    'we have both upper and lower; interpolate
    If (dbl_TestTemperature > dbl_UpperTemperature) Then GoTo GetANSIPressure_Error
    dbl_AllowANSIPressure = dbl_LowerPressure + ((dbl_UpperPressure - dbl_LowerPressure) _
        * ((dbl_TestTemperature - dbl_LowerTemperature) / (dbl_UpperTemperature - dbl_LowerTemperature)))
End If

rset_B16_34.Close
rset_TemperatureTable.Close

Exit Sub

GetANSIPressure_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub GetStemStressArea(dbl_StemDiameter As Double, _
    dbl_StemStessArea As Double)
'--------------------------------------------------------------------------------------
'Reference : Machinist's handbook
'---------------------------------------------------------------------------------------
'This subroutine is used to determine the maximum allowable working pressure with known body material and rating class.
'---------------------------------------------------------------------------------------
'Input Variables:
'
'dbl_StemDiameter           diameter of plug stem      inches
'
'Output Variables:
'dbl_StemStressArea         stem stress area          inches
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/1 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo GetStemStressArea_Error

'***************************
'Declarations
'***************************
Dim rset_Stress As Recordset
Dim str_DatabasePath As String
Dim str_StressQuery As String

'***************************
'Parameter validation
'***************************
If (dbl_StemDiameter <= 0#) Then GoTo GetStemStressArea_Error

'****************************
'Begin calculations
'****************************
dbl_StemStessArea = -1 'initialize to bad value incase of error
str_StressQuery = "SELECT stemDia, stemStressArea from stemStressArea " _
    + "WHERE stemDia = " + Str$(dbl_StemDiameter)
Set rset_Stress = g_db_Gensz.OpenRecordset(str_StressQuery)
If (rset_Stress.RecordCount <> 1) Then GoTo GetStemStressArea_Error
rset_Stress.MoveFirst
dbl_StemStessArea = rset_Stress.Fields(1)

rset_Stress.Close

Exit Sub

GetStemStressArea_Error:
'Placeholder.  How do we want to handle error messages?

End Sub



Public Sub CalcTestPressure(dbl_InletPressure() As Double, dbl_ShutOffDP As Double, _
    dbl_MaxInletPressure As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to determine the largest value of inletPressure and shutoff pressure drop  specified by user.
'--------------------------------------------------------------------------------------
'Input Variables:
'
'dbl_InletPressure(1 to 4)      array of inletPressure values entered by user.  psi
'dbl_ShutoffDP                  specified maximum shutoff pressure drop psi
'
'Output Variables:
'
'dbl_MaxInletPressure           largest inlet pressure specified            psi
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/1 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcTestPressure_Error

'***************************
'Declarations
'***************************
Dim int_Index As Integer

'***************************
'Parameter validation
'***************************
If (dbl_ShutOffDP < 0#) Then GoTo CalcTestPressure_Error

'****************************
'Begin calculations
'****************************

'find the max inletpressure
dbl_MaxInletPressure = 0#
For int_Index = 1 To 4
If (dbl_InletPressure(int_Index) = Null) Then dbl_InletPressure(int_Index) = 0
If (dbl_InletPressure(int_Index) > dbl_MaxInletPressure) _
    Then dbl_MaxInletPressure = dbl_InletPressure(int_Index)
Next int_Index

'make sure its larger than the shutoff pressure
If (dbl_ShutOffDP > dbl_MaxInletPressure) Then dbl_MaxInletPressure = dbl_ShutOffDP

'************************
'If debug_on Then form21000.txtMaxInletPressure.Text = dbl_MaxInletPressure
'************************


Exit Sub

CalcTestPressure_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcTestTemperature(dbl_InletTemperature() As Double, dbl_MaxInletTemperature As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to determine the largest inlet temperature specified by the user
'--------------------------------------------------------------------------------------
'Input Variables:
'
'dbl_InletTemperature(1 to 4)   array of inlet temperature values entered by user.  degreeF
'
'Output Variables:
'
'dbl_MaxInletTemperature        largest inlet pressure specified      degreeF
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/1 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcTestTemperature_Error

'***************************
'Declarations
'***************************
Dim int_Index As Integer

'***************************
'Parameter validation
'***************************

'****************************
'Begin calculations
'****************************

'find the max inlet temperature
dbl_MaxInletTemperature = -99999#
For int_Index = 1 To 4
If (dbl_InletTemperature(int_Index) > dbl_MaxInletTemperature) _
    Then dbl_MaxInletTemperature = dbl_InletTemperature(int_Index)
Next int_Index

'***************************
'If debug_on Then form21000.txtMaxInletTemperature.Text = dbl_MaxInletTemperature
'***************************

Exit Sub

CalcTestTemperature_Error:
'Placeholder.  How do we want to handle error messages?

End Sub


Public Sub CalcMaxDPAllowable(dbl_AllowDPStability As Double, dbl_AllowDPClosed As Double, _
    dbl_AllowDPOpen As Double, dbl_AllowANSIPressure As Double, _
    dbl_MaxAllowDP)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to determine the maximum allowable pressure drop for valve while considering forces on open and closed position, stability, and ANSI limit.
'---------------------------------------------------------------------------------------
'Input Variables:
'
'dbl_AllowDPstability           stability limit         psi
'dbl_AllowDPclosed              limit when valve closed     psi
'dbl_AllowDPopen                limit when valve open       psi
'dbl_AllowANSIpressure          ANSI limit          psi
'
'Output Variables:
'
'dbl_MaxAllowDP                 overall valve limit     psi
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/1 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcMaxDPAllowable_Error

'***************************
'Parameter validation
'***************************
If (dbl_AllowDPStability < 0#) Then GoTo CalcMaxDPAllowable_Error
If (dbl_AllowDPClosed < 0#) Then GoTo CalcMaxDPAllowable_Error
If (dbl_AllowDPOpen < 0#) Then GoTo CalcMaxDPAllowable_Error
If (dbl_AllowANSIPressure < 0#) Then GoTo CalcMaxDPAllowable_Error

'****************************
'Begin calculations
'****************************
dbl_MaxAllowDP = dbl_AllowDPClosed
If (dbl_AllowDPStability > 0 And dbl_AllowDPStability < dbl_MaxAllowDP) Then
    dbl_MaxAllowDP = dbl_AllowDPStability
End If

If (dbl_AllowDPOpen > 0 And dbl_AllowDPOpen < dbl_MaxAllowDP) Then
    dbl_MaxAllowDP = dbl_AllowDPOpen
End If

If (dbl_AllowANSIPressure > 0 And dbl_AllowANSIPressure < dbl_MaxAllowDP) Then
    dbl_MaxAllowDP = dbl_AllowANSIPressure
End If

 Exit Sub

CalcMaxDPAllowable_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcRequiredSupplyPressure(dbl_AreaFinal As Double, dbl_PFinal As Double, _
    dbl_RequiredForce As Double, dbl_PSupply As Double)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate supply pressure required to provide the necessary output load for an air to close linear motion actuator.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_RequiredForce          required actuator thrust            lb
'dbl_AreaFinal              actuator area at final position     in^2
'                               (valve closed)
'dbl_PFinal                 final pressure provised by spring   lb/in^2
'                               setting or air pressure (preload)
'
'Output Variables:
'dbl_PSupply                supply pressure to actuator     lb/in^2
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcRequiredSupplyPressure_Error

'***************************
'Parameter validation
'***************************
If (dbl_RequiredForce < 0#) Then GoTo CalcRequiredSupplyPressure_Error
If (dbl_AreaFinal <= 0#) Then GoTo CalcRequiredSupplyPressure_Error
If (dbl_PFinal < 0#) Then GoTo CalcRequiredSupplyPressure_Error

'****************************
'Begin calculations
'****************************
dbl_PSupply = dbl_PFinal + dbl_RequiredForce / dbl_AreaFinal
If (dbl_PSupply < (dbl_PFinal + 5#)) Then
    dbl_PSupply = dbl_PFinal + 5#
End If

Exit Sub

CalcRequiredSupplyPressure_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcTestPressureDrop(dbl_InletPressure() As Double, dbl_OutletPressure() As Double, _
    dbl_MaxPressureDrop As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to determine the largest pressure drop specified by user.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_InletPressure(1 to 4)      array of inletPressure values entered by user.  psi
'dbl_OutletPressure(1 to 4)     array of outetPressure values entered by user.  psi
'
'Output Variables:
'dbl_MaxPressureDrop            largest pressure drop specified         psi
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcTestPressureDrop_Error

'***************************
'Declarations
'***************************
Dim int_Index As Integer        'index to loop
Dim dbl_PressureDrop As Double  'current pressure drop in loop

'***************************
'Parameter validation
'***************************

'****************************
'Begin calculations
'****************************
dbl_MaxPressureDrop = 0#

For int_Index = 1 To 4
    If (dbl_InletPressure(int_Index) = Null Or dbl_OutletPressure(int_Index) = Null) Then
        dbl_InletPressure(int_Index) = 0#
        dbl_OutletPressure(int_Index) = 0#
    End If
    
    dbl_PressureDrop = dbl_InletPressure(int_Index) - dbl_OutletPressure(int_Index)
    If (dbl_PressureDrop > dbl_MaxPressureDrop) Then dbl_MaxPressureDrop = dbl_PressureDrop
Next int_Index

'****************************
'If debug_on Then form21000.txtMaxPressureDrop.Text = dbl_MaxPressureDrop
'****************************

Exit Sub

CalcTestPressureDrop_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcOffBalanceTorque(dbl_PressureDrop As Double, dbl_SeatRingDiameter As Double, _
    dbl_Eccentricity As Double, dbl_OffBalanceTorque As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves.
'---------------------------------------------------------------------------------------
'This subroutineis used to calculate the unbalance torque resulting from pressure acting on a offset rotary style plug.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_PressureDrop           specified shutoff pressure  psi
'dbl_SeatRingDiameter       seat ring diameter      inches
'dbl_Eccentricity           eccentricity of plug        inches
'                               (distance from center of shaft
'                               to center of seat ring)
'
'Output Variables:
'dbl_OffBallanceTorque      torque required to overcome in-lb
'                               pressure acting on plug
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcOffBalanceTorque_Error

'***************************
'Parameter validation
'***************************
If (dbl_PressureDrop < 0#) Then GoTo CalcOffBalanceTorque_Error
If (dbl_SeatRingDiameter <= 0#) Then GoTo CalcOffBalanceTorque_Error
If (dbl_Eccentricity < 0#) Then GoTo CalcOffBalanceTorque_Error

'****************************
'Begin calculations
'****************************
dbl_OffBalanceTorque = dbl_PressureDrop * dbl_Eccentricity * PI _
    * dbl_SeatRingDiameter * dbl_SeatRingDiameter / 4
    
Exit Sub

CalcOffBalanceTorque_Error:
'Placeholder.  How do we want to handle error messages?

End Sub


Public Sub CalcFrictionTorque(dbl_PressureDrop As Double, dbl_SeatRingDiameter As Double, _
    dbl_shaftDiameter As Double, dbl_PackingCoef As Double, _
    dbl_FrictionTorque As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves.
'---------------------------------------------------------------------------------------
'This subroutineis used to calculate the torque necessary to overcome friction.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_PressureDrop           specified shutoff pressure  psi
'dbl_SeatRingDiameter       diameter of seat ring       inches
'dbl_ShaftDiameter          diameter of shaft       inches
'dbl_PackingCoef            coef. of friction for selected
'                               packing material
'
'Output Variables:
'dbl_FrictionTorque         torque required to overcome     in-lb
'                               friction
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcFrictionTorque_Error

'***************************
'Declarations
'***************************
Dim dbl_PackingTorque As Double
Dim dbl_PackingPressure As Double

'***************************
'Parameter validation
'***************************
If (dbl_PressureDrop < 0#) Then GoTo CalcFrictionTorque_Error
If (dbl_SeatRingDiameter <= 0#) Then GoTo CalcFrictionTorque_Error
If (dbl_shaftDiameter <= 0#) Then GoTo CalcFrictionTorque_Error
If (dbl_PackingCoef <= 0#) Then GoTo CalcFrictionTorque_Error

'****************************
'Begin calculations
'****************************

If (dbl_PressureDrop < 1500#) Then
    dbl_PackingPressure = 1500#
Else
    dbl_PackingPressure = dbl_PressureDrop
End If

dbl_PackingTorque = 0.314 * dbl_PackingPressure * dbl_PackingCoef _
    * dbl_shaftDiameter * dbl_shaftDiameter
    
dbl_FrictionTorque = dbl_PackingTorque + (0.25 * dbl_PressureDrop * dbl_shaftDiameter _
    * PI * dbl_SeatRingDiameter * dbl_SeatRingDiameter / 4)
    

Exit Sub

CalcFrictionTorque_Error:
'Placeholder.  How do we want to handle error messages?

End Sub


Public Sub CalcSeatingTorque(dbl_ActuatorArea As Double, dbl_LeverLength As Double, _
    dbl_shaftDiameter As Double, dbl_SeatingLoadCoef As Double, _
    dbl_Eccentricity As Double, dbl_SeatingTorque As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves.
'---------------------------------------------------------------------------------------
'This subroutineis used to calculate the torque necessary to provide a seating force adequate for the selected leakage class.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_ActuatorArea           effective diaphragm area of actuator    in^2
'dbl_ShaftDiameter          diameter of shaft           inches
'dbl_LeverLength            length of lever connecting actuator     inches
'                               to shaft
'dbl_SeatingLoadCoef        factor used to account for leakage class
'dbl_Eccentricity           eccentricity of plug           inches
'                               (distance from center of shaft
'                               to center of seat ring)
'
'Output Variables:
'dbl_SeatingTorque          torque necessary to provide seating     in-lb
'                               force
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcSeatingTorque_Error

'***************************
'Parameter validation
'***************************
If (dbl_ActuatorArea < 0#) Then GoTo CalcSeatingTorque_Error
If (dbl_shaftDiameter < 0#) Then GoTo CalcSeatingTorque_Error
If (dbl_LeverLength < 0#) Then GoTo CalcSeatingTorque_Error
If (dbl_SeatingLoadCoef < 0#) Then GoTo CalcSeatingTorque_Error
If (dbl_Eccentricity < 0#) Then GoTo CalcSeatingTorque_Error

'****************************
'Begin calculations
'****************************
dbl_SeatingTorque = (1.5 * dbl_ActuatorArea * dbl_LeverLength) + _
    (dbl_SeatingLoadCoef * PI * dbl_shaftDiameter * dbl_Eccentricity)
    
Exit Sub

CalcSeatingTorque_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcRequiredTorque(dbl_OffBalanceTorque As Double, dbl_SeatingTorque As Double, _
    dbl_FrictionTorque As Double, dbl_RequiredTorque As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves.
'---------------------------------------------------------------------------------------
'This subroutineis used to calculate the required torque for an offset rotary style plug.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_SeatingTorque          torque necessary to provide seating     in-lb
'                               force
'dbl_FrictionTorque         torque required to overcome         in-lb
'                               friction
'dbl_OffBalanceTorque       torque required to overcome     in-lb
'                               pressure acting on plug
'
'Output Variables:
'dbl_RequiredTorque         total torque required to operate        in-lb
'                               valve
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcRequiredTorque_Error

'***************************
'Parameter validation
'***************************
If (dbl_OffBalanceTorque < 0#) Then GoTo CalcRequiredTorque_Error
If (dbl_SeatingTorque < 0#) Then GoTo CalcRequiredTorque_Error
If (dbl_FrictionTorque < 0#) Then GoTo CalcRequiredTorque_Error

'****************************
'Begin calculations
'****************************
dbl_RequiredTorque = dbl_SeatingTorque + dbl_FrictionTorque + dbl_OffBalanceTorque

Exit Sub

CalcRequiredTorque_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcMaxAllowSupplyPressureCamflex(dbl_RequiredTorque As Double, _
    dbl_ServiceFactor As Double, dbl_AreaFinal As Double, dbl_LeverLength As Double, _
    dbl_PFinal As Double, dbl_MaxAllowSupplyPressure As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the maximum allowable supply pressure for Camflex.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_AreaFinal                  final area of actuator
'dbl_PFinal                     final actuator pressure         psi
'dbl_LeverLength                length of lever connecting actuator     inches
'                                   to shaft
'dbl_RequiredTorque             total torque required to operate        in-lb
'                                   valve
'dbl_ServiceFactor              factor used to compensate for flow
'                                   direction and fluid type
'
'Output Variables:
'dbl_MaxAllowSupplyPressure     max allowable supply pressure for   psi
'                                   Selected equipment
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcMaxAllowSupplyPressureCamflex_Error

'***************************
'Parameter validation
'***************************
If (dbl_RequiredTorque <= 0#) Then GoTo CalcMaxAllowSupplyPressureCamflex_Error
If (dbl_ServiceFactor < 0#) Then GoTo CalcMaxAllowSupplyPressureCamflex_Error
If (dbl_AreaFinal <= 0#) Then GoTo CalcMaxAllowSupplyPressureCamflex_Error
If (dbl_LeverLength <= 0#) Then GoTo CalcMaxAllowSupplyPressureCamflex_Error
If (dbl_PFinal < 0#) Then GoTo CalcMaxAllowSupplyPressureCamflex_Error

'****************************
'Begin calculations
'****************************
dbl_MaxAllowSupplyPressure = dbl_PFinal + (dbl_RequiredTorque * dbl_ServiceFactor) _
    / (dbl_AreaFinal * dbl_LeverLength)
    
Exit Sub

CalcMaxAllowSupplyPressureCamflex_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcAllowShaftTorque(dbl_AllowShaftStress As Double, dbl_shaftDiameter As Double, _
    dbl_AllowShaftTorque As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves.
'---------------------------------------------------------------------------------------
'This routine calculates the maximum allowable shaft torque from the shaft stress
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_AllowShaftStress           allowable shaft stress          psi
'dbl_ShaftDiameter              diameter of shaft           inches
'
'Output Variables:
'dbl_AllowShaftTorque           allowable shaft torque          in-lb
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcAllowShaftTorque_Error

'***************************
'Parameter validation
'***************************
If (dbl_AllowShaftStress <= 0#) Then GoTo CalcAllowShaftTorque_Error
If (dbl_shaftDiameter <= 0#) Then GoTo CalcAllowShaftTorque_Error

'****************************
'Begin calculations
'****************************
dbl_AllowShaftTorque = (dbl_AllowShaftStress * (0.85 * dbl_shaftDiameter) ^ 3) / 5.1

Exit Sub

CalcAllowShaftTorque_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcAllowDPCamflex(dbl_ActuatorArea As Double, dbl_LeverLength As Double, _
    dbl_PFinal As Double, dbl_SeatRingDiameter As Double, dbl_Eccentricity As Double, _
    dbl_SupplyPressure As Double, dbl_shaftDiameter As Double, _
    dbl_AllowDP As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves.

'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the max. allowable supply pressure for Camflex valves.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_ActuatorArea           effective diaphragm area        in^2
'dbl_LeverLength            length of lever connecting actuator     inches
'                               to shaft
'dbl_PFinal                 final spring setting of actuator        psig
'dbl_SeatRingDiameter       diameter of seat ring orifice       inches
'dbl_Eccentricity           eccentricity of plug            inches
'                               (distance from center of shaft
'                               to center of seat ring)
'dbl_SupplyPressure         supply pressure             psig
'dbl_ShaftDiameter          diameter of shaft           inches
'
'Output Variables:
'dbl_AllowDP                allowable pressure drop with specified
'                               valve and actuator
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcAllowDPCamflex_Error

'***************************
'Parameter validation
'***************************
If (dbl_ActuatorArea <= 0#) Then GoTo CalcAllowDPCamflex_Error
If (dbl_LeverLength <= 0#) Then GoTo CalcAllowDPCamflex_Error
If (dbl_PFinal < 0#) Then GoTo CalcAllowDPCamflex_Error
If (dbl_SeatRingDiameter <= 0#) Then GoTo CalcAllowDPCamflex_Error
If (dbl_Eccentricity <= 0#) Then GoTo CalcAllowDPCamflex_Error
If (dbl_SupplyPressure <= 0#) Then GoTo CalcAllowDPCamflex_Error
If (dbl_shaftDiameter <= 0#) Then GoTo CalcAllowDPCamflex_Error

'****************************
'Begin calculations
'****************************
dbl_AllowDP = (dbl_ActuatorArea * dbl_LeverLength * (dbl_SupplyPressure - dbl_PFinal)) _
    / ((PI * dbl_SeatRingDiameter * dbl_SeatRingDiameter / 4#) _
    * (dbl_Eccentricity + 0.25 * dbl_shaftDiameter))
    
Exit Sub

CalcAllowDPCamflex_Error:
'Placeholder.  How do we want to handle error messages?

End Sub
Sub calcAllowDPCamflexFromTorque(dbl_theTorque As Double, dbl_PFinal As Double, dbl_SeatRingDiameter As Double, dbl_Eccentricity As Double, _
     dbl_shaftDiameter As Double, dbl_AllowDP As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves.

'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the max. allowable supply pressure for Camflex valves when a
'given torque is known.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_theTorque              torque valve (required or maximum)  in-lb
'dbl_PFinal                 final spring setting of actuator    psig
'dbl_SeatRingDiameter       diameter of seat ring orifice       inches
'dbl_Eccentricity           eccentricity of plug                inches
'                               (distance from center of shaft
'                               to center of seat ring)
'dbl_SupplyPressure         supply pressure                     psig
'dbl_ShaftDiameter          diameter of shaft                   inches
'
'Output Variables:
'dbl_AllowDP                allowable pressure drop with specified
'                               valve and actuator
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 8/24/98 John Haberlin

'-----------------------------------------------------------------------------------------

On Error GoTo CalcAllowDPCamflex_Error

'***************************
'Parameter validation
'***************************

If (dbl_SeatRingDiameter <= 0#) Then GoTo CalcAllowDPCamflex_Error
If (dbl_Eccentricity <= 0#) Then GoTo CalcAllowDPCamflex_Error

If (dbl_shaftDiameter <= 0#) Then GoTo CalcAllowDPCamflex_Error

'****************************
'Begin calculations
'****************************
dbl_AllowDP = (dbl_theTorque) _
    / ((PI * dbl_SeatRingDiameter * dbl_SeatRingDiameter / 4#) _
    * (dbl_Eccentricity + 0.25 * dbl_shaftDiameter))
    
Exit Sub

CalcAllowDPCamflex_Error:
'Placeholder.  How do we want to handle error messages?
End Sub

Public Sub CalcRequiredSupplyPressureCamflex(dbl_Torque As Double, dbl_ActuatorArea As Double, _
    dbl_LeverLength As Double, dbl_PFinal As Double, dbl_RequiredSupplyPressure As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the required supply pressure for an rotary valves.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_Torque                  required or allowable torque           in-lb
'dbl_ActuatorArea            area of diaphragm                      in^2
'dbl_LeverLength             length of actuator lever               inches
'dbl_PFinal                  final actuator pressure                psi
'
'Output Variables:
'dbl_RequiredSupplyPressure  required supply pressure for spec'd DP psi
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcRequiredSupplyPressureCamflex_Error

'***************************
'Declarations
'***************************
Dim dbl_TempPRessure As Double

'***************************
'Parameter validation
'***************************
'If (dbl_Torque <= 0#) Then GoTo CalcRequiredSupplyPressureCamflex_Error
If (dbl_ActuatorArea <= 0#) Then GoTo CalcRequiredSupplyPressureCamflex_Error
If (dbl_LeverLength <= 0#) Then GoTo CalcRequiredSupplyPressureCamflex_Error
If (dbl_PFinal <= 0#) Then GoTo CalcRequiredSupplyPressureCamflex_Error

'****************************
'Begin calculations
'****************************

'dbl_RequiredSupplyPressure = (((PI * dbl_SeatRingDiameter * dbl_SeatRingDiameter / 4#) _
    * (dbl_Eccentricity + 0.25 * dbl_ShaftDiameter)) / (dbl_ActuatorArea * dbl_LeverLength)) + dbl_PFinal

dbl_TempPRessure = dbl_Torque / (dbl_ActuatorArea * dbl_LeverLength)
'If (dbl_TempPRessure < 5#) Then dbl_TempPRessure = 5#
dbl_RequiredSupplyPressure = dbl_TempPRessure + dbl_PFinal

Exit Sub

CalcRequiredSupplyPressureCamflex_Error:
'Placeholder.  How do we want to handle error messages?

End Sub
Sub CalcReqSupplyPressureCamflex(dbl_Pressure As Double, dbl_Eccentricity, dbl_shaftDiameter, dbl_ActuatorArea As Double, dbl_SeatRingDiameter, _
    dbl_LeverLength As Double, dbl_PFinal As Double, dbl_RequiredSupplyPressure As Double)
    
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the required supply pressure for an rotary valves.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_Pressure               specified service pressure                  psig
'dbl_Eccentricity           eccentricity of spcified valve plug         inches
'dbl_shaftDiameter          diameter of valve shaft                     inches
'dbl_seatRingDiameter       orifice diameter of seat ring               inches
'dbl_ActuatorArea           area of diaphragm                           in^2
'dbl_LeverLength            length of actuator lever                    inches
'dbl_PFinal                 final actuator pressure                     psi
'
'Output Variables:
'dbl_RequiredSupplyPressure required supply pressure for spec'd DP      psi
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 8/24/98 John Haberlin
'-----------------------------------------------------------------------------------------

On Error GoTo CalcRequiredSupplyPressureCamflex_Error

'***************************
'Declarations
'***************************
Dim dbl_TempPRessure As Double

'***************************
'Parameter validation
'***************************
'If (dbl_Torque <= 0#) Then GoTo CalcRequiredSupplyPressureCamflex_Error
If (dbl_ActuatorArea <= 0#) Then GoTo CalcRequiredSupplyPressureCamflex_Error
If (dbl_LeverLength <= 0#) Then GoTo CalcRequiredSupplyPressureCamflex_Error
If (dbl_PFinal <= 0#) Then GoTo CalcRequiredSupplyPressureCamflex_Error

'****************************
'Begin calculations
'****************************

dbl_RequiredSupplyPressure = ((dbl_Pressure * (PI * dbl_SeatRingDiameter * dbl_SeatRingDiameter / 4#) _
    * (dbl_Eccentricity + 0.25 * dbl_shaftDiameter)) / (dbl_ActuatorArea * dbl_LeverLength)) + dbl_PFinal

Exit Sub

CalcRequiredSupplyPressureCamflex_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcDPAllowStabilityCamflex(dbl_AreaFinal As Double, dbl_LeverLength As Double, _
    dbl_PInitial As Double, dbl_PFinal As Double, dbl_SeatRingDiameter As Double, _
    dbl_Eccentricity As Double, dbl_shaftDiameter As Double, dbl_AllowDPStability As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves.

'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the max. allowable supply pressure for Camflex valves.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_AreaFinal           effective diaphragm area                           in^2
'dbl_LeverLength         length of lever connecting actuator to shaft       in
'dbl_PInitial            initial spring setting of actuator                 psig
'dbl_PFinal              final spring setting of actuator                   psig
'dbl_SeatRingDiameter    diameter of seat ring orifice                      in
'dbl_Eccentricity        eccentricity of plug                               in
'                        (distance from ctr of shaft to ctr of seat ring)
'dbl_ShaftDiameter       diameter of shaft                                  in
'
'Output Variables:
'dbl_AllowDPstability    allowable pressure drop with specified
'                        valve and actuator
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcDPAllowStabilityCamflex_Error

'***************************
'Parameter validation
'***************************
If (dbl_AreaFinal <= 0#) Then GoTo CalcDPAllowStabilityCamflex_Error
If (dbl_LeverLength <= 0#) Then GoTo CalcDPAllowStabilityCamflex_Error
If (dbl_PInitial < 0#) Then GoTo CalcDPAllowStabilityCamflex_Error
If (dbl_PFinal <= 0#) Then GoTo CalcDPAllowStabilityCamflex_Error
If (dbl_SeatRingDiameter <= 0#) Then GoTo CalcDPAllowStabilityCamflex_Error
If (dbl_Eccentricity <= 0#) Then GoTo CalcDPAllowStabilityCamflex_Error
If (dbl_shaftDiameter <= 0#) Then GoTo CalcDPAllowStabilityCamflex_Error

'****************************
'Begin calculations
'****************************
dbl_AllowDPStability = (dbl_AreaFinal * dbl_LeverLength * (dbl_PFinal - dbl_PInitial) / 2#) _
    / ((PI * dbl_SeatRingDiameter * dbl_SeatRingDiameter / 4) _
    * (dbl_Eccentricity + 0.25 * dbl_shaftDiameter))

Exit Sub

CalcDPAllowStabilityCamflex_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcDPAllowStaticCamflex(dbl_AreaInitial As Double, dbl_LeverLength As Double, _
    dbl_PInitial As Double, dbl_SeatRingDiameter As Double, _
    dbl_Eccentricity As Double, dbl_shaftDiameter As Double, dbl_AllowDP As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the max. allowable supply pressure for Camflex valves.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_AreaInitial            effective diaphragm area                            in^2
'dbl_LeverLength            length of lever connecting actuator to shaft        in
'dbl_PInitial               initial spring setting of actuator                  psig
'dbl_SeatRingDiameter       diameter of seat ring orifice                       in
'dbl_Eccentricity           eccentricity of plug                                in
'                           (distance from ctr of shaft to ctr of seat ring)
'dbl_ShaftDiameter          diameter of shaft                                   in
'
'Output Variables:
'dbl_AllowDP                allowable pressure drop with specified
'                           valve and actuator
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcDPAllowStaticCamflex_Error


'****************************
'Begin calculations
'****************************

'Use equations for CalcAllowDPCamflex with SupplyPressure=PInitial and PFinal=0
Call CalcAllowDPCamflex(dbl_AreaInitial, dbl_LeverLength, 0#, dbl_SeatRingDiameter, _
    dbl_Eccentricity, dbl_PInitial, dbl_shaftDiameter, dbl_AllowDP)

Exit Sub

CalcDPAllowStaticCamflex_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcCylinderPreloadSupply(dbl_PSupplyMin As Double, dbl_PPreload As Double, _
    dbl_PSupply As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the preload and supply for cylinder type actuators..
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_PSupplyMin             calculated required supply pressure         psig

'Output Variables:
'dbl_PPreload               preload pressure                            psig
'dbl_PSupply                supply pressure                             psig

'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcCylinderPreloadSupply_Error

'***************************
'Parameter validation
'***************************
If (dbl_PSupplyMin < 0#) Then GoTo CalcCylinderPreloadSupply_Error

'****************************
'Begin calculations
'****************************
If dbl_PSupplyMin > 0# And dbl_PSupplyMin <= 15# Then
    dbl_PPreload = 15#
    dbl_PSupply = 30#
ElseIf dbl_PSupplyMin > 15# And dbl_PSupplyMin <= 25# Then
    dbl_PPreload = 15#
    dbl_PSupply = dbl_PSupplyMin + dbl_PPreload
ElseIf dbl_PSupplyMin > 25# And dbl_PSupplyMin <= 35# Then
    dbl_PPreload = 20#
    dbl_PSupply = dbl_PSupplyMin + dbl_PPreload
ElseIf (dbl_PSupplyMin > 35# And dbl_PSupplyMin <= 50#) Then
    dbl_PPreload = 25#
    dbl_PSupply = dbl_PSupplyMin + dbl_PPreload
Else
    'unacceptable application
    dbl_PPreload = 0#
    dbl_PSupply = 0#
End If

Exit Sub

CalcCylinderPreloadSupply_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcDPAllowStability41000(dbl_SpringRate As Double, dbl_Kv As Double, _
    str_ActuatorModel As String, dbl_AreaInitial As Double, dbl_Travel As Double, _
    dbl_MaxSupplyPressure As Double, dbl_AllowDPStability As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for reciprocating valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the allowable pressure drop 41000 series valves based on stability.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_SpringRate        spring rate of selected actuator                         lbs/in
'dbl_Kv                unbalanced area rated of change of selected plug/cage    in^2/in
'dbl_ActuatorModel     model of selected actuator
'dbl_AreaInitial       initial area of selected actuator                        in^2
'dbl_Travel            rated travel of plug                                     in
'dbl_MaxSupplyPressure maximum allowable supply for selected actuator           psig
'
'Output Variables:
'dbl_AllowDPstability  calculated allowable pressure drop                       psi
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcDPAllowStability41000_Error

'***************************
'Declarations
'***************************
Dim dbl_VolumeAbove As Double   'volume above piston seal
Dim dbl_VolumeBelow As Double   'volume below piston seal
Dim dbl_Pb As Double
Dim dbl_K As Double             'ratio of specific heats for air
Dim dbl_Kt As Double            'used to define additional actuator stiffness
                                '  from air above and below piston seal

'***************************
'Parameter validation
'***************************
If (dbl_SpringRate <= 0#) Then GoTo CalcDPAllowStability41000_Error
If (dbl_Kv <= 0#) Then GoTo CalcDPAllowStability41000_Error
If (dbl_AreaInitial <= 0#) Then GoTo CalcDPAllowStability41000_Error
If (dbl_Travel <= 0#) Then GoTo CalcDPAllowStability41000_Error
If (dbl_MaxSupplyPressure <= 0#) Then GoTo CalcDPAllowStability41000_Error

'****************************
'Begin calculations
'****************************
dbl_K = 1.4
dbl_Kt = 0#
dbl_Pb = dbl_MaxSupplyPressure / 2#

Select Case str_ActuatorModel
    Case "84", "85"
        dbl_VolumeAbove = (dbl_Travel + 0.4) * dbl_AreaInitial
        dbl_Kt = (dbl_K * dbl_AreaInitial * dbl_AreaInitial) _
            * ((14.7 + (dbl_SpringRate * dbl_Travel / dbl_AreaInitial)) _
            / dbl_VolumeAbove)
            
    Case "86", "89"
        dbl_VolumeAbove = (dbl_Travel + 0.4) * dbl_AreaInitial
        dbl_VolumeBelow = (16.4 * dbl_AreaInitial) - dbl_VolumeAbove
        
        dbl_Kt = (dbl_K * dbl_AreaInitial * dbl_AreaInitial) _
            * ((((dbl_Pb + 14.7) - (dbl_SpringRate * dbl_Travel / dbl_AreaInitial)) _
            / dbl_VolumeAbove) _
            + (((dbl_Pb + 14.7) + (dbl_SpringRate * dbl_Travel / dbl_AreaInitial)) _
            / dbl_VolumeBelow))
        
End Select

dbl_AllowDPStability = (dbl_SpringRate + dbl_Kt) / dbl_Kv

Exit Sub

CalcDPAllowStability41000_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcDPAllowOpen41000_ATC(dbl_AreaInitial As Double, dbl_PInitial As Double, _
    dbl_ForcePacking As Double, dbl_Au As Double, dbl_AllowDPOpen As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for reciprocating valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the allowable pressure drop in the open position for
'41000 series valves, ATC.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_AreaInitial     initial area of selected actuator                  in^2
'dbl_PInitial        initial spring setting of selected actuator        psig
'dbl_ForcePacking    initial spring setting of selected plug/cage       psig
'dbl_Au              effective unbalanced area                          in^2
'
'Output Variables:
'dbl_AllowDPopen     allowable pressure drop when open                  psi
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcDPAllowOpen41000_ATC_Error

'***************************
'Parameter validation
'***************************
If (dbl_AreaInitial <= 0#) Then GoTo CalcDPAllowOpen41000_ATC_Error
If (dbl_PInitial < 0#) Then GoTo CalcDPAllowOpen41000_ATC_Error
If (dbl_Au <= 0#) Then GoTo CalcDPAllowOpen41000_ATC_Error
If (dbl_ForcePacking < -0#) Then GoTo CalcDPAllowOpen41000_ATC_Error

'****************************
'Begin calculations
'****************************
dbl_AllowDPOpen = (dbl_AreaInitial * dbl_PInitial - dbl_ForcePacking) / dbl_Au

Exit Sub

CalcDPAllowOpen41000_ATC_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcRequiredSupplyPressure41000_FTC_ATC(dbl_PFinal As Double, _
    dbl_AreaFinal As Double, dbl_ForcePilotSpring As Double, _
    dbl_ForcePacking As Double, dbl_ForceSeating As Double, _
    dbl_RequiredSupplyPressure As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for reciprocating valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the required supply pressure for
'41000 series valves, ATC.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_AreaFinal              final area of selected actuator                 in^2
'dbl_PFinal                 final spring setting of selected actuator       psig
'dbl_ForcePilotSpring       force required to compress pilot spring         lbs
'dbl_ForcePacking           force required to overcome packing friction     lbs
'dbl_ForceSeating           force required to provide adequate seating      lbs
'
'Output Variables:
'dbl_RequiredSupplyPressure supply pressure necessary to operate valve      psig
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcRequiredSupplyPressure41000_FTC_ATC_Error

'***************************
'Parameter validation
'***************************
If (dbl_PFinal < 0#) Then GoTo CalcRequiredSupplyPressure41000_FTC_ATC_Error
If (dbl_AreaFinal <= 0#) Then GoTo CalcRequiredSupplyPressure41000_FTC_ATC_Error
If (dbl_ForcePilotSpring < 0#) Then GoTo CalcRequiredSupplyPressure41000_FTC_ATC_Error
If (dbl_ForcePacking < 0#) Then GoTo CalcRequiredSupplyPressure41000_FTC_ATC_Error
If (dbl_ForceSeating < 0#) Then GoTo CalcRequiredSupplyPressure41000_FTC_ATC_Error

'****************************
'Begin calculations
'****************************
dbl_RequiredSupplyPressure = ((dbl_PFinal * dbl_AreaFinal) + dbl_ForcePilotSpring _
    + dbl_ForcePacking + dbl_ForceSeating) / (dbl_AreaFinal)
    
Exit Sub

CalcRequiredSupplyPressure41000_FTC_ATC_Error:
'Placeholder.  How do we want to handle error messages?

End Sub


Public Sub CalcRequiredSupplyPressure41000_FTC_ATO(dbl_PFinal As Double, _
    dbl_AreaFinal As Double, dbl_ForcePacking As Double, dbl_Au As Double, _
    dbl_MaxInletPressure As Double, dbl_RequiredSupplyPressure As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for reciprocating valves.

'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the required supply pressure for
'41000 series valves, ATO.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_AreaFinal                  final area of selected actuator             in^2
'dbl_PFinal                     final spring setting of selected actuator   psig
'dbl_ForcePacking               force required to overcome packing friction lbs
'dbl_Au                         effective unbalanced area                   in^2
'dbl_MaxInletPressure           maximum inlet pressure specified            psig
'
'Output Variables:
'dbl_RequiredSupplyPressure     supply pressure necessary to operate valve  psig
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcRequiredSupplyPressure41000_FTC_ATO_Error

'***************************
'Parameter validation
'***************************
If (dbl_PFinal <= 0#) Then GoTo CalcRequiredSupplyPressure41000_FTC_ATO_Error
If (dbl_AreaFinal <= 0#) Then GoTo CalcRequiredSupplyPressure41000_FTC_ATO_Error
If (dbl_ForcePacking < 0#) Then GoTo CalcRequiredSupplyPressure41000_FTC_ATO_Error
If (dbl_Au < 0#) Then GoTo CalcRequiredSupplyPressure41000_FTC_ATO_Error
If (dbl_MaxInletPressure < 0#) Then GoTo CalcRequiredSupplyPressure41000_FTC_ATO_Error

'****************************
'Begin calculations
'****************************
dbl_RequiredSupplyPressure = ((dbl_PFinal * dbl_AreaFinal) + dbl_ForcePacking _
    + (dbl_Au * dbl_MaxInletPressure)) / (dbl_AreaFinal)
    
Exit Sub

CalcRequiredSupplyPressure41000_FTC_ATO_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcRequiredSupplyPressure41000_FTO_ATC(dbl_PFinal As Double, _
    dbl_AreaFinal As Double, dbl_ForcePacking As Double, dbl_ForceSeating As Double, _
    dbl_Au As Double, dbl_MaxInletPressure As Double, _
    dbl_RequiredSupplyPressure As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for reciprocating valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the required supply pressure for
'41000 series valves, FTO/ATC.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_AreaFinal              final area of selected actuator                     in^2
'dbl_PFinal                 final spring setting of selected actuator           psig
'dbl_ForcePacking           force required to overcome packing friction         lbs
'dbl_ForceSeating           force required to provide adequate seating          lbs
'dbl_Au                     effective offbalanced area                          in^2
'dbl_MaxInletPressure       maximum specified inlet pressure                    psig
'
'Output Variables:
'dbl_RequiredSupplyPressure supply pressure necessary to operate valve          psig
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcRequiredSupplyPressure41000_FTO_ATC_Error

'***************************
'Parameter validation
'***************************
If (dbl_PFinal <= 0#) Then GoTo CalcRequiredSupplyPressure41000_FTO_ATC_Error
If (dbl_AreaFinal <= 0#) Then GoTo CalcRequiredSupplyPressure41000_FTO_ATC_Error
If (dbl_ForcePacking < 0#) Then GoTo CalcRequiredSupplyPressure41000_FTO_ATC_Error
If (dbl_ForceSeating <= 0#) Then GoTo CalcRequiredSupplyPressure41000_FTO_ATC_Error
If (dbl_Au < 0#) Then GoTo CalcRequiredSupplyPressure41000_FTO_ATC_Error
If (dbl_MaxInletPressure <= 0#) Then GoTo CalcRequiredSupplyPressure41000_FTO_ATC_Error

'****************************
'Begin calculations
'****************************
dbl_RequiredSupplyPressure = (dbl_PFinal * dbl_AreaFinal + dbl_Au * dbl_MaxInletPressure _
    + dbl_ForcePacking + dbl_ForceSeating) / (dbl_AreaFinal)
    
Exit Sub

CalcRequiredSupplyPressure41000_FTO_ATC_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcRequiredSupplyPressure41000_FTO_ATO(dbl_PFinal As Double, _
    dbl_AreaFinal As Double, dbl_ForcePacking As Double, dbl_Au As Double, _
    dbl_MaxInletPressure As Double, dbl_RequiredSupplyPressure As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for reciprocating valves.
'--------------------------------------------------------------------------------------
'This subroutine is used to calculate the required supply pressure for
'41000 series valves, FTO/ATO.
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_AreaFinal              final area of selected actuator                     in^2
'dbl_PFinal                 final spring setting of selected actuator           psig
'dbl_ForcePacking           force required to overcome packing friction         lbs
'dbl_Au                     effective offbalanced area                          in^2
'dbl_MaxInletPressure       maximum specified inlet pressure                    psig
'
'Output Variables:
'dbl_RequiredSupplyPressure supply pressure necessary to operate valve          psig
'--------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcRequiredSupplyPressure41000_FTO_ATO_Error

'***************************
'Parameter validation
'***************************
If (dbl_PFinal <= 0#) Then GoTo CalcRequiredSupplyPressure41000_FTO_ATO_Error
If (dbl_AreaFinal <= 0#) Then GoTo CalcRequiredSupplyPressure41000_FTO_ATO_Error
If (dbl_ForcePacking < 0#) Then GoTo CalcRequiredSupplyPressure41000_FTO_ATO_Error
If (dbl_Au <= 0#) Then GoTo CalcRequiredSupplyPressure41000_FTO_ATO_Error
If (dbl_MaxInletPressure <= 0#) Then GoTo CalcRequiredSupplyPressure41000_FTO_ATO_Error

'****************************
'Begin calculations
'****************************
dbl_RequiredSupplyPressure = (dbl_PFinal * dbl_AreaFinal + dbl_Au * dbl_MaxInletPressure _
    + dbl_ForcePacking) / (dbl_AreaFinal)
    
Exit Sub

CalcRequiredSupplyPressure41000_FTO_ATO_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcAllowDP41000_FTO_ATC(dbl_AllowableStemLoad As Double, _
    dbl_ForceSeating As Double, dbl_ForcePacking As Double, dbl_Au As Double, _
    dbl_AllowDPStem As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for reciprocating valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the allowable pressure drop for 41000 series valves, FTO/ATC,
'based on the stem limit.
'
'NOTE: THIS SUBROUTINE MAY BE USED FOR ATC AND ATO.  THE SUBROUTINE WAS NAMED
'       ERRONEOUSLY.   J.HABERLIN 7/9/98
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_AllowableStemLoad    Allowable stem load for spec'd material and temp  lbs
'dbl_ForcePacking         Force required to overcome packing friction       lbs
'dbl_ForceSeating         Force required to provide specified leakage       lbs
'dbl_Au                   Effective offbalanced area                        in^2
'dbl_MaxInletPressure     maximum specified inlet pressure                  psig
'
'Output Variables:
'dbl_AllowDPStem          allowable pressure drop with specified stem       psig
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcAllowDP41000_FTO_ATC_Error

'***************************
'Parameter validation
'***************************
If (dbl_AllowableStemLoad <= 0#) Then GoTo CalcAllowDP41000_FTO_ATC_Error
If (dbl_ForceSeating < 0#) Then GoTo CalcAllowDP41000_FTO_ATC_Error
If (dbl_ForcePacking < 0#) Then GoTo CalcAllowDP41000_FTO_ATC_Error
If (dbl_Au <= 0#) Then GoTo CalcAllowDP41000_FTO_ATC_Error

'****************************
'Begin calculations
'****************************
dbl_AllowDPStem = (dbl_AllowableStemLoad - dbl_ForceSeating - dbl_ForcePacking) _
    / dbl_Au
    
Exit Sub

CalcAllowDP41000_FTO_ATC_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcAllowDP_30000_FTC_ATO(dbl_ActuatorOutputTorque As Double, _
    dbl_OffBalOpenTorque As Double, dbl_OffBalClosedTorque As Double, _
    dbl_SeatingTorque As Double, dbl_PInitial As Double, dbl_PFinal As Double, _
    dbl_AllowSupplyMax As Double, dbl_AllowSupplyStability As Double, _
    dbl_MaxDPClosed As Double, dbl_MaxDPOpen As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves, "Varimax allowable DP calculations" by C. Vlassoff Aug. 23, 1996.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate allowable pressure drop for 30000 series valves FTC, ATO.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_ActuatorOutputTorque    output torque of selected actuator                 daNm/psi
'dbl_OffBalOpenTorque        dynamic torque in open position                    daNm/psi
'dbl_OffBalClosedTorque      dynamic torque in closed position                  daNm/psi
'dbl_SeatingTorque           seating Torque                                     daNm/psi
'dbl_PInitial                initial spring setting of selected actuator        psig
'dbl_PFinal                  final spring setting of selected actuator          psig
'dbl_AllowSupplyMax          maximum allowable supply                           psig
'dbl_AllowSupplyStability    maximum allowable supply - stability               psig
'
'Output Variables:
'dbl_MaxDPclosed             maximum allowable pressure drop when valve closed  psig
'dbl_MaxDpopen               maximum allowable pressure drop when valve opened  psig
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/2 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcAllowDP_30000_FTC_ATO_Error

'***************************
'Declarations
'***************************
Dim dbl_Supply As Double

'***************************
'Parameter validation
'***************************
If (dbl_ActuatorOutputTorque <= 0#) Then GoTo CalcAllowDP_30000_FTC_ATO_Error
If (dbl_OffBalOpenTorque < 0#) Then GoTo CalcAllowDP_30000_FTC_ATO_Error
If (dbl_OffBalClosedTorque < 0#) Then GoTo CalcAllowDP_30000_FTC_ATO_Error
If (dbl_SeatingTorque <= 0#) Then GoTo CalcAllowDP_30000_FTC_ATO_Error
If (dbl_PInitial < 0#) Then GoTo CalcAllowDP_30000_FTC_ATO_Error
If (dbl_PFinal <= 0#) Then GoTo CalcAllowDP_30000_FTC_ATO_Error
If (dbl_AllowSupplyMax <= 0#) Then GoTo CalcAllowDP_30000_FTC_ATO_Error
If (dbl_AllowSupplyStability <= 0#) Then GoTo CalcAllowDP_30000_FTC_ATO_Error

'****************************
'Begin calculations
'****************************
If (dbl_AllowSupplyMax < dbl_AllowSupplyStability) Then
    dbl_Supply = dbl_AllowSupplyMax
Else
    dbl_Supply = dbl_AllowSupplyStability
End If

dbl_MaxDPClosed = 14.5 * (dbl_ActuatorOutputTorque * (dbl_Supply - dbl_PInitial) - dbl_SeatingTorque) _
    / dbl_OffBalClosedTorque

dbl_MaxDPOpen = 14.5 * (dbl_ActuatorOutputTorque * (dbl_AllowSupplyStability - dbl_PFinal)) _
    / dbl_OffBalOpenTorque
    
Exit Sub

CalcAllowDP_30000_FTC_ATO_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcAllowDP_30000_FTC_ATC(dbl_ActuatorOutputTorque As Double, _
    dbl_OffBalOpenTorque As Double, dbl_OffBalClosedTorque As Double, _
    dbl_SeatingTorque As Double, dbl_PInitial As Double, dbl_PFinal As Double, _
    dbl_AllowSupplyMax As Double, dbl_AllowSupplyStability As Double, _
    dbl_MaxDPClosed As Double, dbl_MaxDPOpen As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves, "Varimax allowable DP calculations" by C. Vlassoff Aug. 23, 1996.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate allowable pressure drop for 30000 series valves FTC, ATC.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_ActuatorOutputTorque       output torque of selected actuator      daNm/psi
'dbl_OffBalOpenTorque           dynamic torque in open position         daNm/psi
'dbl_OffBalClosedTorque         dynamic torque in closed position       daNm/psi
'dbl_SeatingTorque              seating Torque                  daNm/psi
'dbl_PInitial                   initial spring setting of selected
'                                   actuator psig
'dbl_PFinal                     final spring setting of selected
'                                   actuator psig
'dbl_AllowSupplyMax             maximum allowable supply            psig
'dbl_AllowSupplyStability       maximum allowable supply - stability        psig
'
'Output Variables:
'dbl_MaxDPclosed                maximum allowable pressure drop when    psig
'                                   valve closed
'dbl_MaxDPopen                  maximum allowable pressure drop when        psig
'                                   valve open
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/3 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcAllowDP_30000_FTC_ATC_Error

'***************************
'Declarations
'***************************
Dim dbl_Supply As Double

'***************************
'Parameter validation
'***************************
If (dbl_ActuatorOutputTorque <= 0#) Then GoTo CalcAllowDP_30000_FTC_ATC_Error
If (dbl_OffBalOpenTorque <= 0#) Then GoTo CalcAllowDP_30000_FTC_ATC_Error
If (dbl_OffBalClosedTorque <= 0#) Then GoTo CalcAllowDP_30000_FTC_ATC_Error
If (dbl_SeatingTorque <= 0#) Then GoTo CalcAllowDP_30000_FTC_ATC_Error
If (dbl_PInitial < 0#) Then GoTo CalcAllowDP_30000_FTC_ATC_Error
If (dbl_PFinal <= 0#) Then GoTo CalcAllowDP_30000_FTC_ATC_Error
If (dbl_AllowSupplyMax <= 0#) Then GoTo CalcAllowDP_30000_FTC_ATC_Error
If (dbl_AllowSupplyStability <= 0#) Then GoTo CalcAllowDP_30000_FTC_ATC_Error

'****************************
'Begin calculations
'****************************
If (dbl_AllowSupplyMax < dbl_AllowSupplyStability) Then
    dbl_Supply = dbl_AllowSupplyMax
Else
    dbl_Supply = dbl_AllowSupplyStability
End If

dbl_MaxDPClosed = 14.5 * (dbl_ActuatorOutputTorque * (dbl_PFinal - dbl_Supply) _
    - dbl_SeatingTorque) / dbl_OffBalClosedTorque

dbl_MaxDPOpen = 14.5 * dbl_ActuatorOutputTorque * dbl_PInitial / dbl_OffBalOpenTorque

Exit Sub

CalcAllowDP_30000_FTC_ATC_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcAllowDP_30000_FTO_ATO(dbl_ActuatorOutputTorque As Double, _
    dbl_OffBalOpenTorque As Double, dbl_OffBalClosedTorque As Double, _
    dbl_SeatingTorque As Double, dbl_PInitial As Double, dbl_PFinal As Double, _
    dbl_AllowSupplyMax As Double, dbl_AllowSupplyStability As Double, _
    dbl_MaxDPClosed As Double, dbl_MaxDPOpen As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves, "Varimax allowable DP calculations" by C. Vlassoff Aug. 23, 1996.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate allowable pressure drop for 30000 series valves FTO, ATO.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_ActuatorOutputTorque       output torque of selected actuator      daNm/psi
'dbl_OffBalOpenTorque           dynamic torque in open position         daNm/psi
'dbl_OffBalClosedTorque         dynamic torque in closed position       daNm/psi
'dbl_SeatingTorque              seating Torque                  daNm/psi
'dbl_PInitial                   initial spring setting of selected
'                                   actuator psig
'dbl_PFinal                     final spring setting of selected
'                                   actuator psig
'dbl_AllowSupplyMax             maximum allowable supply            psig
'dbl_AllowSupplyStability       maximum allowable supply - stability        psig
'
'Output Variables:
'dbl_MaxDPclosed                maximum allowable pressure drop when    psig
'                                   valve closed
'dbl_MaxDPopen                  maximum allowable pressure drop when        psig
'                                   valve open
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/3 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcAllowDP_30000_FTO_ATO_Error

'***************************
'Declarations
'***************************
Dim dbl_Supply As Double

'***************************
'Parameter validation
'***************************
If (dbl_ActuatorOutputTorque <= 0#) Then GoTo CalcAllowDP_30000_FTO_ATO_Error
If (dbl_OffBalOpenTorque <= 0#) Then GoTo CalcAllowDP_30000_FTO_ATO_Error
If (dbl_OffBalClosedTorque <= 0#) Then GoTo CalcAllowDP_30000_FTO_ATO_Error
If (dbl_SeatingTorque <= 0#) Then GoTo CalcAllowDP_30000_FTO_ATO_Error
If (dbl_PInitial < 0#) Then GoTo CalcAllowDP_30000_FTO_ATO_Error
If (dbl_PFinal <= 0#) Then GoTo CalcAllowDP_30000_FTO_ATO_Error
If (dbl_AllowSupplyMax <= 0#) Then GoTo CalcAllowDP_30000_FTO_ATO_Error
If (dbl_AllowSupplyStability <= 0#) Then GoTo CalcAllowDP_30000_FTO_ATO_Error

'****************************
'Begin calculations
'****************************
If (dbl_AllowSupplyMax < dbl_AllowSupplyStability) Then
    dbl_Supply = dbl_AllowSupplyMax
Else
    dbl_Supply = dbl_AllowSupplyStability
End If

dbl_MaxDPClosed = 14.5 * (dbl_ActuatorOutputTorque * dbl_PInitial - dbl_SeatingTorque) _
    / dbl_OffBalClosedTorque

dbl_MaxDPOpen = 14.5 * dbl_ActuatorOutputTorque * (dbl_PFinal - dbl_Supply) _
    / dbl_OffBalOpenTorque
    
Exit Sub

CalcAllowDP_30000_FTO_ATO_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcAllowDP_30000_FTO_ATC(dbl_ActuatorOutputTorque As Double, _
    dbl_OffBalOpenTorque As Double, dbl_OffBalClosedTorque As Double, _
    dbl_SeatingTorque As Double, dbl_PInitial As Double, dbl_PFinal As Double, _
    dbl_AllowSupplyMax As Double, dbl_AllowSupplyStability As Double, _
    dbl_MaxDPClosed As Double, dbl_MaxDPOpen As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves, "Varimax allowable DP calculations" by C. Vlassoff Aug. 23, 1996.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate allowable pressure drop for 30000 series valves FTO, ATC.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_ActuatorOutputTorque       output torque of selected actuator      daNm/psi
'dbl_OffBalOpenTorque           dynamic torque in open position         daNm/psi
'dbl_OffBalClosedTorque         dynamic torque in closed position       daNm/psi
'dbl_SeatingTorque              seating Torque                  daNm/psi
'dbl_PInitial                   initial spring setting of selected
'                                   actuator psig
'dbl_PFinal                     final spring setting of selected
'                                   actuator psig
'dbl_AllowSupplyMax             maximum allowable supply            psig
'dbl_AllowSupplyStability       maximum allowable supply - stability        psig
'
'Output Variables:
'dbl_MaxDPclosed                maximum allowable pressure drop when    psig
'                                   valve closed
'dbl_MaxDpopen                  maximum allowable pressure drop when        psig
'                                   valve open
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/3 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcAllowDP_30000_FTO_ATC_Error

'***************************
'Declarations
'***************************
Dim dbl_Supply As Double

'***************************
'Parameter validation
'***************************
If (dbl_ActuatorOutputTorque <= 0#) Then GoTo CalcAllowDP_30000_FTO_ATC_Error
If (dbl_OffBalOpenTorque <= 0#) Then GoTo CalcAllowDP_30000_FTO_ATC_Error
If (dbl_OffBalClosedTorque <= 0#) Then GoTo CalcAllowDP_30000_FTO_ATC_Error
If (dbl_SeatingTorque <= 0#) Then GoTo CalcAllowDP_30000_FTO_ATC_Error
If (dbl_PInitial < 0#) Then GoTo CalcAllowDP_30000_FTO_ATC_Error
If (dbl_PFinal <= 0#) Then GoTo CalcAllowDP_30000_FTO_ATC_Error
If (dbl_AllowSupplyMax <= 0#) Then GoTo CalcAllowDP_30000_FTO_ATC_Error
If (dbl_AllowSupplyStability <= 0#) Then GoTo CalcAllowDP_30000_FTO_ATC_Error

'****************************
'Begin calculations
'****************************
If (dbl_AllowSupplyMax < dbl_AllowSupplyStability) Then
    dbl_Supply = dbl_AllowSupplyMax
Else
    dbl_Supply = dbl_AllowSupplyStability
End If

dbl_MaxDPClosed = 14.5 * (dbl_ActuatorOutputTorque * (dbl_AllowSupplyMax - dbl_PFinal) - dbl_SeatingTorque) _
    / dbl_OffBalClosedTorque
    
dbl_MaxDPOpen = 14.5 * dbl_ActuatorOutputTorque * (dbl_AllowSupplyStability - dbl_PInitial) _
    / dbl_OffBalOpenTorque
    
Exit Sub

CalcAllowDP_30000_FTO_ATC_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcRequiredSupplyPressure30000_FTC_ATO(dbl_ActuatorOutputTorque As Double, _
    dbl_OffBalOpenTorque As Double, dbl_OffBalClosedTorque As Double, _
    dbl_SeatingTorque As Double, dbl_PInitial As Double, dbl_PFinal As Double, _
    dbl_MaxInletPressure As Double, dbl_MaxPressureDrop As Double, _
    dbl_RequiredSupplyOpen As Double, dbl_RequiredSupplyClosed As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves, "Varimax allowable DP calculations" by C. Vlassoff Aug. 23, 1996.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate required supply pressure  for 30000 series valves FTC, ATO.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_ActuatorOutputTorque       output torque of selected actuator      daNm/psi
'dbl_OffBalOpenTorque           dynamic torque in open position         daNm/psi
'dbl_OffBalClosedTorque         dynamic torque in closed position       daNm/psi
'dbl_SeatingTorque              seating Torque                  daNm/psi
'dbl_PInitial                   initial spring setting of selected
'                                   actuator psig
'dbl_PFinal                     final spring setting of selected
'                                   actuator psig
'dbl_MaxInletPressure           maximum specified inlet pressure        psi
'dbl_MaxPressureDrop            maximum specified pressure drop     psi
'
'Output Variables:
'dbl_RequiredSupplyClosed       supply pressure required            psig
'dbl_RequiredSupplyopen         supply pressure required            psig
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/3 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcRequiredSupplyPressure30000_FTC_ATO_Error

'***************************
'Parameter validation
'***************************
If (dbl_ActuatorOutputTorque <= 0#) Then GoTo CalcRequiredSupplyPressure30000_FTC_ATO_Error
If (dbl_OffBalOpenTorque <= 0#) Then GoTo CalcRequiredSupplyPressure30000_FTC_ATO_Error
If (dbl_OffBalClosedTorque <= 0#) Then GoTo CalcRequiredSupplyPressure30000_FTC_ATO_Error
If (dbl_SeatingTorque <= 0#) Then GoTo CalcRequiredSupplyPressure30000_FTC_ATO_Error
If (dbl_PInitial < 0#) Then GoTo CalcRequiredSupplyPressure30000_FTC_ATO_Error
If (dbl_PFinal <= 0#) Then GoTo CalcRequiredSupplyPressure30000_FTC_ATO_Error
If (dbl_MaxInletPressure <= 0#) Then GoTo CalcRequiredSupplyPressure30000_FTC_ATO_Error
If (dbl_MaxPressureDrop <= 0#) Then GoTo CalcRequiredSupplyPressure30000_FTC_ATO_Error

'****************************
'Begin calculations
'****************************
dbl_RequiredSupplyClosed = dbl_PInitial + (dbl_OffBalClosedTorque * dbl_MaxInletPressure / 14.5 _
    + dbl_SeatingTorque) / dbl_ActuatorOutputTorque
    
dbl_RequiredSupplyOpen = dbl_PFinal + (dbl_OffBalOpenTorque * dbl_MaxPressureDrop) _
    / (14.5 * dbl_ActuatorOutputTorque)
    
Exit Sub

CalcRequiredSupplyPressure30000_FTC_ATO_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub CalcRequiredSupplyPressure30000_FTO_ATC(dbl_ActuatorOutputTorque As Double, _
    dbl_OffBalOpenTorque As Double, dbl_OffBalClosedTorque As Double, _
    dbl_SeatingTorque As Double, dbl_PInitial As Double, dbl_PFinal As Double, _
    dbl_MaxInletPressure As Double, dbl_MaxPressureDrop As Double, _
    dbl_RequiredSupplyOpen As Double, dbl_RequiredSupplyClosed As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for rotary valves, "Varimax allowable DP calculations" by C. Vlassoff Aug. 23, 1996.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate required supply pressure  for 30000 series valves FTO, ATC.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_ActuatorOutputTorque       output torque of selected actuator      daNm/psi
'dbl_OffBalOpenTorque           dynamic torque in open position         daNm/psi
'dbl_OffBalClosedTorque         dynamic torque in closed position       daNm/psi
'dbl_SeatingTorque              seating Torque                  daNm/psi
'dbl_PInitial                   initial spring setting of selected
'                                   actuator psig
'dbl_PFinal                     final spring setting of selected
'                                   actuator psig
'dbl_MaxInletPressure           maximum specified inlet pressure        psi
'dbl_MaxPressureDrop            maximum specified pressure drop     psi
'
'Output Variables:
'dbl_RequiredSupplyClosed       supply pressure required            psig
'dbl_RequiredSupplyopen         supply pressure required            psig
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/3 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcRequiredSupplyPressure30000_FTO_ATC_Error

'***************************
'Parameter validation
'***************************
If (dbl_ActuatorOutputTorque <= 0#) Then GoTo CalcRequiredSupplyPressure30000_FTO_ATC_Error
If (dbl_OffBalOpenTorque <= 0#) Then GoTo CalcRequiredSupplyPressure30000_FTO_ATC_Error
If (dbl_OffBalClosedTorque <= 0#) Then GoTo CalcRequiredSupplyPressure30000_FTO_ATC_Error
If (dbl_SeatingTorque <= 0#) Then GoTo CalcRequiredSupplyPressure30000_FTO_ATC_Error
If (dbl_PInitial < 0#) Then GoTo CalcRequiredSupplyPressure30000_FTO_ATC_Error
If (dbl_PFinal <= 0#) Then GoTo CalcRequiredSupplyPressure30000_FTO_ATC_Error
If (dbl_MaxInletPressure <= 0#) Then GoTo CalcRequiredSupplyPressure30000_FTO_ATC_Error
If (dbl_MaxPressureDrop <= 0#) Then GoTo CalcRequiredSupplyPressure30000_FTO_ATC_Error

'****************************
'Begin calculations
'****************************
dbl_RequiredSupplyClosed = dbl_PFinal + (dbl_OffBalClosedTorque * dbl_MaxInletPressure / 14.5 _
    + dbl_SeatingTorque) / dbl_ActuatorOutputTorque
    
dbl_RequiredSupplyOpen = dbl_PInitial + (dbl_OffBalOpenTorque * dbl_MaxPressureDrop) _
    / (14.5 * dbl_ActuatorOutputTorque)

Exit Sub

CalcRequiredSupplyPressure30000_FTO_ATC_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub GetStemProperties(str_StemMaterial As String, _
    dbl_MaxInletTemperature As Double, dbl_Sy As Double, dbl_E As Double, _
    dbl_Sd As Double)
'--------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------
'This subroutine is used to determine the maximum allowable design stress, allowable yield stress, and the modulus of elasticity for the selected stem material at the maximum specified inlet temperature.
'---------------------------------------------------------------------------------------
'Input Variables:
'str_StemMaterial           stem material
'dbl_MaxInletTemperature    maximum specified inlet temperature         degF
'
'Output Variables:
'dbl_Sy                     allowable yield stress for material at temperature  psi
'dbl_Sd                     allowable design stress                 psi
'dbl_E                      modulus of elasticity                   psi
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/4 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo GetStemProperties_Error

'***************************
'Declarations
'***************************
Dim str_DatabasePath As String
Dim str_Query As String
Dim rset_Yield As Recordset
Dim str_Stress As String
Dim rset_Stress As Recordset
Dim str_Elasticity As String
Dim rset_Elasticity As Recordset

Dim dbl_LowValue As Double
Dim dbl_HighValue As Double
Dim dbl_LowTemperature As Double
Dim dbl_HighTemperature As Double

'***************************
'Parameter validation
'***************************
If (str_StemMaterial = "") Then GoTo GetStemProperties_Error

'****************************
'Begin calculations
'****************************

'get and interpolate yield stress
str_Query = "SELECT temperatureF, [?] FROM stemMinYield"
str_Query = ReplaceStr(str_Query, "?", str_StemMaterial)

Set rset_Yield = g_db_Gensz.OpenRecordset(str_Query)
If (rset_Yield.RecordCount <= 0) Then GoTo GetStemProperties_Error
rset_Yield.MoveFirst

dbl_HighTemperature = -99999#
dbl_HighValue = 0#

'find bracketing temperatures
Do Until rset_Yield.EOF
dbl_LowTemperature = dbl_HighTemperature
dbl_LowValue = dbl_HighValue
dbl_HighTemperature = rset_Yield.Fields(0)
dbl_HighValue = rset_Yield.Fields(1)
If (dbl_MaxInletTemperature <= dbl_HighTemperature) Then Exit Do
rset_Yield.MoveNext
Loop

'interpolate
If (dbl_HighTemperature <= dbl_LowTemperature) Then GoTo GetStemProperties_Error
dbl_Sy = dbl_LowValue + (dbl_MaxInletTemperature - dbl_LowTemperature) _
    * (dbl_HighValue - dbl_LowValue) / (dbl_HighTemperature - dbl_LowTemperature)

'get and interpolate design stress
str_Query = "SELECT temperatureF, [" + str_StemMaterial + "] FROM stemAllowDesignStress"
Set rset_Stress = g_db_Gensz.OpenRecordset(str_Query)
If (rset_Stress.RecordCount <= 0) Then GoTo GetStemProperties_Error
rset_Stress.MoveFirst

dbl_HighTemperature = -99999#
dbl_HighValue = 0#

'find bracketing temperatures
Do Until rset_Stress.EOF
dbl_LowTemperature = dbl_HighTemperature
dbl_LowValue = dbl_HighValue
dbl_HighTemperature = rset_Stress.Fields(0)
dbl_HighValue = rset_Stress.Fields(1)
If (dbl_MaxInletTemperature <= dbl_HighTemperature) Then Exit Do
rset_Stress.MoveNext
Loop

'interpolate
If (dbl_HighTemperature <= dbl_LowTemperature) Then GoTo GetStemProperties_Error
dbl_Sd = dbl_LowValue + (dbl_MaxInletTemperature - dbl_LowTemperature) _
    * (dbl_HighValue - dbl_LowValue) / (dbl_HighTemperature - dbl_LowTemperature)
    
'get and interpolate elasticity
str_Query = "SELECT temperatureF, [" + str_StemMaterial + "] FROM modulusOfElasticity"
Set rset_Elasticity = g_db_Gensz.OpenRecordset(str_Query)
If (rset_Elasticity.RecordCount <= 0) Then GoTo GetStemProperties_Error
rset_Elasticity.MoveFirst

dbl_HighTemperature = -99999#
dbl_HighValue = 0#

'find bracketing temperatures
Do Until rset_Elasticity.EOF
dbl_LowTemperature = dbl_HighTemperature
dbl_LowValue = dbl_HighValue
dbl_HighTemperature = rset_Elasticity.Fields(0)
dbl_HighValue = rset_Elasticity.Fields(1)
If (dbl_MaxInletTemperature <= dbl_HighTemperature) Then Exit Do
rset_Elasticity.MoveNext
Loop

'interpolate
If (dbl_HighTemperature <= dbl_LowTemperature) Then GoTo GetStemProperties_Error
dbl_E = dbl_LowValue + (dbl_MaxInletTemperature - dbl_LowTemperature) _
    * (dbl_HighValue - dbl_LowValue) / (dbl_HighTemperature - dbl_LowTemperature)
    
rset_Yield.Close
rset_Stress.Close
rset_Elasticity.Close

Exit Sub

GetStemProperties_Error:
'Placeholder.  How do we want to handle error messages?

End Sub


Public Sub GetServiceFactor(dbl_ValveSize As Double, long_AirActionID As Long, _
    str_FluidType As String, str_ActuatorModel As String, dbl_ServiceFactor)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing for rotary valves
'---------------------------------------------------------------------------------------
'This subroutine is used to determine the allowable service factor  with known valve size , actuator action, fluid type, and actuator model.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_ValveSize          selected valve size             mm
'str_ActuatorAction     selected actuator air action
'str_ActuatorModel      selected actuator model
'str_FluidType          type or phase of fliud
'
'Output Variables:
'dbl_ServiceFactor      factor used to compensate for different
'                           fluid types during actuator sizing
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/5 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo GetServiceFactor_Error

'***************************
'Declarations
'***************************

'***************************
'Parameter validation
'***************************
If (dbl_ValveSize <= 0#) Then GoTo GetServiceFactor_Error
If (str_ActuatorModel = "") Then GoTo GetServiceFactor_Error
If (str_FluidType = "") Then GoTo GetServiceFactor_Error

'****************************
'Begin calculations
'****************************
dbl_ServiceFactor = 1#

Select Case dbl_ValveSize
    Case Is < 150#
        dbl_ServiceFactor = 1#
    Case Is = 150#
        If (long_AirActionID = ATC) Then dbl_ServiceFactor = 1#
        If (long_AirActionID = ATO And str_FluidType = "Gas") Then dbl_ServiceFactor = 1#
        If (long_AirActionID = ATO And str_FluidType = "Liquid") Then dbl_ServiceFactor = 0.5
    Case Is > 150#
        If (long_AirActionID = ATC) Then dbl_ServiceFactor = 1#
        If (long_AirActionID = ATO And str_FluidType = "Gas") Then dbl_ServiceFactor = 0.5
        If (long_AirActionID = ATO And str_FluidType = "Liquid") Then dbl_ServiceFactor = 0.25
End Select

If (str_ActuatorModel = "70" Or str_ActuatorModel = "37HPI") Then dbl_ServiceFactor = 1#

Exit Sub

GetServiceFactor_Error:
'Placeholder.  How do we want to handle error messages?

End Sub



Public Sub CalcDPAllowOpen41000_ATO(dbl_AreaFinal As Double, dbl_PFinal As Double, _
    dbl_PSupply As Double, dbl_ForcePacking As Double, dbl_Au As Double, _
    dbl_AllowDPOpen As Double)
'--------------------------------------------------------------------------------------
'Reference:  Actuator sizing methods for reciprocating valves
'---------------------------------------------------------------------------------------
'This routine is used to calculate the allowable pressure drop in the open position
'   for the 41000 series ATO
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_AreaFinal          final area of the selected actuator   in^2
'dbl_PInitial           Initial pressure
'dbl_PSupply            supply pressure
'dbl_ForcePacking       packing force
'dbl_Au                 Effective unbalanced area
'
'Output Variables:
'dbl_AllowDPOpen        allowabel pressure drop when open
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/5 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo CalcDPAllowOpen41000_ATO_Error

'***************************
'Declarations
'***************************

'***************************
'Parameter validation
'***************************
If (dbl_AreaFinal <= 0#) Then GoTo CalcDPAllowOpen41000_ATO_Error
If (dbl_PFinal <= 0#) Then GoTo CalcDPAllowOpen41000_ATO_Error
If (dbl_PSupply <= 0#) Then GoTo CalcDPAllowOpen41000_ATO_Error
If (dbl_ForcePacking < 0#) Then GoTo CalcDPAllowOpen41000_ATO_Error
If (dbl_Au < 0#) Then GoTo CalcDPAllowOpen41000_ATO_Error

'****************************
'Begin calculations
'****************************
dbl_AllowDPOpen = ((dbl_PSupply - dbl_PFinal) * dbl_AreaFinal - dbl_ForcePacking) _
    / dbl_Au

Exit Sub

CalcDPAllowOpen41000_ATO_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

Public Sub Get35AllowShaftTorque(dbl_ValveSize As Double, str_ShaftMaterial As String, _
    dbl_AllowShaftTorque As Double)
'--------------------------------------------------------------------------------------
'Reference : Actuator sizing for rotary valves
'---------------------------------------------------------------------------------------
'This subroutine is used to determine the allowable shaft stress  with known valve size and shaft material.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_ValveSize               selected valve size             mm
'str_ShaftMaterial               selected material of shaft
'
'Output Variables:
'dbl_AllowShaftTorque    allowable shaft torque for Camflex      in-lb
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded 6/5 Larry Schoonover - not tested
'6/10 -Larry Schoonover -tested - passed
'-----------------------------------------------------------------------------------------

On Error GoTo Get35AllowShaftTorque_Error

'***************************
'Parameter validation
'***************************
If (dbl_ValveSize < 0#) Then GoTo Get35AllowShaftTorque_Error
If (str_ShaftMaterial = "") Then GoTo Get35AllowShaftTorque_Error

'****************************
'Begin calculations
'****************************

If (str_ShaftMaterial = "A564 type 17-4PH") Then
    Select Case dbl_ValveSize
        Case 25
            dbl_AllowShaftTorque = 226#
        Case 40
            dbl_AllowShaftTorque = 442#
        Case 50
            dbl_AllowShaftTorque = 442#
        Case 80
            dbl_AllowShaftTorque = 1810#
        Case 100
            dbl_AllowShaftTorque = 3520#
        Case 150
            dbl_AllowShaftTorque = 12200#
        Case 200
            dbl_AllowShaftTorque = 12200#
        Case 250
            dbl_AllowShaftTorque = 19400#
        Case 300
            dbl_AllowShaftTorque = 29000#
        Case 350
            dbl_AllowShaftTorque = 56400#
        Case 400
            dbl_AllowShaftTorque = 56400#
        Case Else
            GoTo Get35AllowShaftTorque_Error
    End Select
Else
    Select Case dbl_ValveSize
        Case 25
            dbl_AllowShaftTorque = 226#
        Case 40
            dbl_AllowShaftTorque = 442#
        Case 50
            dbl_AllowShaftTorque = 442#
        Case 80
            dbl_AllowShaftTorque = 1810#
        Case 100
            dbl_AllowShaftTorque = 3520#
        Case 150
            dbl_AllowShaftTorque = 6100#
        Case 200
            dbl_AllowShaftTorque = 6100#
        Case 250
            dbl_AllowShaftTorque = 9700#
        Case 300
            dbl_AllowShaftTorque = 14510#
        Case 350
            dbl_AllowShaftTorque = 28200#
        Case 400
            dbl_AllowShaftTorque = 28200#
        Case Else
            GoTo Get35AllowShaftTorque_Error
    End Select
End If
Exit Sub

Get35AllowShaftTorque_Error:
'Placeholder.  How do we want to handle error messages?

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' END OF Actuator lower lever
''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INCLUDE Conversion Routines
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'*********************************************************************
'This class module provides an interface allowing external applications
'to invoke the conversion routines in the MCommon standard module.
'Author: KPMG           Date:6/12/98
'*********************************************************************

Public Sub CnvrtDensityFromKg3H(densityKgM3 As Double, densityUM As String, densityOut As Double)
convertDensityFromKg3H densityKgM3, densityUM, densityOut
End Sub

Public Sub CnvrtDensityToKg3H(densityIN As Double, densityUM As String, denistyKgM3 As Double)
convertDensityToKg3H densityIN, densityUM, denistyKgM3
End Sub

Public Sub CnvrtEnthalpyFromMetric(enthalpyMetric As Double, enthalpyUM As String, enthalpyOut As Double)
convertEnthalpyFromMetric enthalpyMetric, enthalpyUM, enthalpyOut
End Sub

Public Sub CnvrtFlowRateFromKgh(flowKgh As Double, flowUM As String, inletDensityKgM3 As Double, flowOut As Double)
convertFlowRateFromKgh flowKgh, flowUM, inletDensityKgM3, flowOut
End Sub

Public Sub CnvrtFlowRateFromKgs(flowKgs As Double, flowUM As String, inletDensityKgM3 As Double, flowOut As Double)
convertFlowRateFromKgs flowKgs, flowUM, inletDensityKgM3, flowOut
End Sub

Public Sub CnvrtFlowRateToKgh(flowIn As Double, flowUM As String, inletDensityKgM3 As Double, flowKgh As Double)
convertFlowRateToKgh flowIn, flowUM, inletDensityKgM3, flowKgh
End Sub

Public Sub CnvrtFlowRateToKgs(flowIn As Double, flowUM As String, inletDensityKgM3 As Double, flowKgs As Double)
convertFlowRateToKgs flowIn, flowUM, inletDensityKgM3, flowKgs
End Sub

Public Sub CnvrtLengthFromM(lengthM As Double, lengthUM As String, lengthOut As Double)
convertLengthFromM lengthM, lengthUM, lengthOut
End Sub

Public Sub CnvrtLengthFromMM(lengthMM As Double, lengthUM As String, lengthOut As Double)
convertLengthFromMM lengthMM, lengthUM, lengthOut
End Sub

Public Sub CnvrtLengthToM(lengthIn As Double, lengthUM As String, lengthM As Double)
convertLengthToM lengthIn, lengthUM, lengthM
End Sub

Public Sub CnvrtLengthToMM(lengthIn As Double, lengthUM As String, lengthMM As Double)
convertLengthToMM lengthIn, lengthUM, lengthMM
End Sub

Public Sub CnvrtPressureFromBarA(pressureBarA As Double, pressureInputUM As String, pressureOut As Double)
convertPressureFromBarA pressureBarA, pressureInputUM, pressureOut
End Sub

Public Sub CnvrtPressureFromPaA(pressurePaA As Double, pressureInputUM As String, pressureOut As Double)
convertPressureFromPaA pressurePaA, pressureInputUM, pressureOut
End Sub

Public Sub CnvrtPressureToBarA(pressureInput As Double, pressureInputUM As String, pressureBarA As Double)
convertPressureToBarA pressureInput, pressureInputUM, pressureBarA
End Sub

Public Sub CnvrtPressureToPaA(pressureInput As Double, pressureInputUM As String, pressurePaA As Double)
convertPressureToPaA pressureInput, pressureInputUM, pressurePaA
End Sub

Public Sub CnvrtSpecVolFromMetric(specVolMetric As Double, specVolUM As String, specVolOut As Double)
convertSpecVolFromMetric specVolMetric, specVolUM, specVolOut
End Sub

Public Sub CnvrtTemperatureFromDegK(degK As Double, degInputUM As String, degOut As Double)
convertTemperatureFromDegK degK, degInputUM, degOut
End Sub

Public Sub CnvrtTemperatureToDegK(degInput As Double, degInputUM As String, degK As Double)
convertTemperatureToDegK degInput, degInputUM, degK
End Sub

Public Sub CnvrtVelocityFromMetric(velocityMPS As Double, velocityUM As String, velocityOUT As Double)
convertVelocityFromMetric velocityMPS, velocityUM, velocityOUT
End Sub

Public Sub CnvrtVelocityToMetric(velocityIN As Double, velocityUM As String, velocityMPS As Double)
convertVelocityToMetric velocityIN, velocityUM, velocityMPS
End Sub

Public Sub CnvrtDensityToKgM3AtInlet(dbl_Density As Double, str_DensityUM As String, dbl_T1 As Double, dbl_P1 As Double, dbl_InletDensity As Double)
convertDensityToKgM3AtInlet dbl_Density, str_DensityUM, dbl_T1, dbl_P1, dbl_InletDensity
End Sub

Function TrimDecimalPlaces(dbl_Value As Double) As Double
'routine to limit the number of decimal places printed
Dim lng_IntValue As Long

If (dbl_Value = 0#) Then Exit Function
If (dbl_Value > 1000) Then
    lng_IntValue = dbl_Value
    TrimDecimalPlaces = lng_IntValue
    Exit Function
End If
If (dbl_Value > 100) Then
    lng_IntValue = dbl_Value * 10
    TrimDecimalPlaces = lng_IntValue / 10#
    Exit Function
End If
If (dbl_Value > 10) Then
    lng_IntValue = dbl_Value * 100
    TrimDecimalPlaces = lng_IntValue / 100#
    Exit Function
End If
If (dbl_Value > 1) Then
    lng_IntValue = dbl_Value * 1000
    TrimDecimalPlaces = lng_IntValue / 1000#
    Exit Function
End If
If (dbl_Value > 0.1) Then
    lng_IntValue = dbl_Value * 10000
    TrimDecimalPlaces = lng_IntValue / 10000#
    Exit Function
Else
    lng_IntValue = dbl_Value * 100000
    TrimDecimalPlaces = lng_IntValue / 100000#
    Exit Function
End If

    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''
' END OF conversion routines
''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INCLUDE Gas LL Routines
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Public Sub CalcFp(dbl_ValveInletDia_mm As Double, dbl_Ksum As Double, _
dbl_Ki As Double, dbl_Cv As Double, dbl_FL As Double, _
dbl_Fp As Double, dbl_Flp As Double)
'-----------------------------------------------------------------------------------------------------
'Source: ISA 75.01 - 1985, section 4.3, piping geometry factor Fp
'
'
'This subroutine calculates FL and  FLp which are used by Cv calculations.
'-----------------------------------------------------------------------------------------------------
'Input           Output
'
'dbl_ValveInletDia          Valve inlet size                                mm
'dbl_Ksum                   Sum of effectibe velocity head coefficients
'dbl_Ki                     Sum of Bernoulli coefficients
'dbl_Cv                     Calculated Cv
'dbl_FL                     Pressure recovery factor
'
'               dbl_Fp      Piping geometry factor
'               dbl_FLp     Combined pressure recovery and piping geometry factor
'Coded By:                  C.Ugbesia
'Date:                      6/02/98
'-----------------------------------------------------------------------------------------------------
On Error GoTo CalcFp_Error

'**********************************
'Variable Declarations
'**********************************
'
'**********************************
'Check input validation
'**********************************
'
'All diameters must be greater than 0
If dbl_ValveInletDia_mm < 0 Then GoTo CalcFp_Error
'dbl_Cv,dbl_Ki,dbl_Ksum,dbl_FL, must be greater than 0
If dbl_Cv < 0 Then GoTo CalcFp_Error
If dbl_Ki < 0 Then GoTo CalcFp_Error
If dbl_FL < 0 Then GoTo CalcFp_Error
'
'**********************************
'Begin Calculation
'**********************************
'dbl_ValveInletDia_mm = dbl_ValveInletDia * 1000
dbl_Fp = Sqr((0.00214 * dbl_ValveInletDia_mm ^ 4) / (0.00214 * dbl_ValveInletDia_mm ^ 4 + (dbl_Ksum * dbl_Cv ^ 2)))
dbl_Flp = dbl_FL * Sqr((0.00214 * dbl_ValveInletDia_mm ^ 4) / (0.00214 * dbl_ValveInletDia_mm ^ 4 + (dbl_Ki * dbl_FL ^ 2 * dbl_Cv ^ 2)))

CalcFp_Error:
'Placeholder


End Sub
Public Sub PipeGeom(dbl_InletPipeDia As Double, dbl_OutletPipeDia As Double, _
dbl_ValveInletDia As Double, dbl_ValveOutletDia As Double, _
dbl_Ksum As Double, dbl_Ki As Double)
'------------------------------------------------------------------------------------------
'Source: ISA 75.01 - 1985, section 4.3, piping geometry factor Fp
'-------------------------------------------------------------------------------------------
'
' Input              Output
' dbl_InletPipeDia.              (All four must be in same units)
' dbl_OutletPipeDia              (All four must be in same units)
' dbl_ValveInletDia.             (All four must be in same units)
' dbl_ValveOutletDia.            (All four must be in same units)
'
'
'                    dbl_Ksum     (dimensionless)
'                    dbl_Ki       (dimensionless)
'
' Coded By:           C.Ugbesia
' Date:               6/02/98
'---------------------------------------------------------------------------------------------

On Error GoTo PipeGeom_Error
'**********************************
'Check input validation
'**********************************
'All diameters must be greater than 0
If dbl_InletPipeDia <= 0 Or dbl_OutletPipeDia <= 0 Then GoTo PipeGeom_Error

'**********************************
'Declarations
'**********************************
Dim dbl_K1 As Double
Dim dbl_K1B As Double
Dim dbl_K2 As Double
Dim dbl_K2B As Double

'**********************************
'Begin Calculation
'**********************************

If dbl_InletPipeDia > dbl_ValveInletDia Then
dbl_K1 = 0.5 * (1 - (dbl_ValveInletDia ^ 2 / (dbl_InletPipeDia ^ 2))) ^ 2
dbl_K1B = 1 - (dbl_ValveInletDia / dbl_InletPipeDia) ^ 4

Else
    dbl_K1 = 0
    dbl_K1B = 0
End If
     
If dbl_OutletPipeDia > dbl_ValveOutletDia Then
dbl_K2 = (1 - dbl_ValveOutletDia ^ 2 / (dbl_OutletPipeDia ^ 2)) ^ 2
dbl_K2B = 1 - (dbl_ValveOutletDia / dbl_OutletPipeDia) ^ 4

Else
    dbl_K2 = 0
    dbl_K2B = 0
End If

 dbl_Ksum = dbl_K1 + dbl_K2 + dbl_K1B - dbl_K2B
 dbl_Ki = dbl_K1 + dbl_K1B

Exit Sub
PipeGeom_Error:
'placeholder

End Sub

Public Function Log10(X)
If (X <= 0) Then
    Log10 = 0
    Exit Function
Else
    Log10 = Log(X) / Log(10)
End If
End Function
Public Sub SelectGasCVCalc(str_CalculationType As String, _
    dbl_InletPressure As Double, dbl_OutletPressure As Double, _
    dbl_InletTemperature As Double, dbl_FlowW As Double, dbl_K As Double, _
    dbl_rho As Double, dbl_MW As Double, dbl_Z1 As Double, _
    dbl_CriticalTemperature As Double, dbl_CriticalPressure As Double, _
    dbl_SL As Double, dbl_Gg As Double, dbl_FlowQ As Double, dbl_Mo As Double, _
    dbl_M2 As Double, _
    str_FlowUM As String, str_PressureUM As String, str_TemperatureUM As String, _
    dbl_ValveDia As Double, dbl_ValveInletDia As Double, _
    dbl_ValveOutletDia As Double, dbl_OutletPipeWallThickness As Double, _
    dbl_OutletPipeOD As Double, str_ValveLine As String, str_BodyMaterial As String, _
    str_ANSIRating As String, lng_TrimTypeID As Long, _
    lng_FlowDirectionID As Long, lng_TrimPackageID As Long, _
    dbl_InletDensity As Double, dbl_FL As Double, dbl_Pa As Double, _
    dbl_Cv As Double, dbl_RatedCv As Double, dbl_CalculatedCv As Double, _
    dbl_InletPipeDia As Double, dbl_OutletPipeDia As Double, _
    str_DensityUM As String, str_LengthUM As String, str_FluidName As String, _
    bln_ReducedArea As Boolean, dbl_Dsonic As Double, dbl_D3rdsonic As Double, _
    dbl_V2 As Double, dbl_Z As Double, lng_FlowStateFlag As Long, str_VelocityUM As String, _
    str_AbsPressureUM As String, dbl_PctTravel As Double)

On Error GoTo SelectGasCVCalc_Error
    Dim dbl_MassFlowRate As Double
    Dim dbl_P1 As Double
    Dim dbl_P2 As Double
    Dim dbl_rho1 As Double
    Dim dbl_rho2 As Double
    Dim dbl_T1 As Double
    Dim dbl_T2 As Double
    Dim dbl_W As Double
    Dim int_SelectionFlag As Integer
    Dim lng_IECTrimStyleID As Long
    Dim int_Count As Integer


#If (DebugVersion = "yes") Then
MsgBox "str_CalculationType= " + str_CalculationType + Chr(9) + Chr(9) + "bln_ReducedArea= " + CStr(bln_ReducedArea) + Chr(13) + "lng_FlowDirectionID= " + CStr(lng_FlowDirectionID) + Chr(9) + Chr(9) + _
"lng_TrimPackageID= " + CStr(lng_TrimPackageID) + Chr(13) + "lng_TrimTypeID= " + CStr(lng_TrimTypeID) + Chr(9) + Chr(9) + Chr(9) + "str_AbsPressureUM= " + str_AbsPressureUM + Chr(13) + "str_VelocityUM= " + str_VelocityUM + Chr(9) + Chr(9) + "str_DensityUM= " + str_DensityUM + Chr(13) + "str_FlowUM= " + str_FlowUM + Chr(13) + _
"str_LengthUM= " + str_LengthUM + Chr(9) + Chr(9) + Chr(9) + "str_PressureUM= " + str_PressureUM + Chr(13) + "str_TemperatureUM= " + str_TemperatureUM + Chr(9) + Chr(9) + "str_ANSIRating= " + str_ANSIRating + Chr(13) + _
"str_BodyMaterial= " + CStr(str_BodyMaterial) + Chr(9) + Chr(9) + Chr(9) + "str_FluidName= " + str_FluidName + Chr(13) + "str_ValveLine= " + str_ValveLine + Chr(9) + Chr(9) + "dbl_InletPressure= " + CStr(dbl_InletPressure) + Chr(13) + _
"dbl_OutletPressure= " + CStr(dbl_OutletPressure) + Chr(9) + Chr(9) + "dbl_InletTemperature= " + CStr(dbl_InletTemperature) + Chr(13) + "dbl_CriticalTemperature= " + CStr(dbl_CriticalTemperature) + Chr(9) + Chr(9) + "dbl_CriticalPressure= " + CStr(dbl_CriticalPressure) + Chr(13) + _
"dbl_FlowQ= " + CStr(dbl_FlowQ) + Chr(9) + Chr(9) + Chr(9) + "dbl_FlowW= " + CStr(dbl_FlowW) + Chr(13) + "dbl_k= " + CStr(dbl_K) + Chr(9) + Chr(9) + Chr(9) + "dbl_CalculatedCv= " + CStr(dbl_CalculatedCv) + Chr(13) + _
"dbl_Cv= " + CStr(dbl_Cv) + Chr(9) + Chr(9) + Chr(9) + "dbl_FL= " + CStr(dbl_FL) + Chr(13) + _
"dbl_Gg= " + CStr(dbl_Gg)
#End If

#If (DebugVersion = "yes") Then
MsgBox "dbl_InletPipeDia= " + CStr(dbl_InletPipeDia) + Chr(13) + _
"dbl_MW= " + CStr(dbl_MW) + Chr(13) + "dbl_OutletPipeDia= " + CStr(dbl_OutletPipeDia) + Chr(13) + _
"dbl_OutletPipeWallThickness= " + CStr(dbl_OutletPipeWallThickness) + Chr(13) + "dbl_OutletPipeOD= " + CStr(dbl_OutletPipeOD) + Chr(9) + Chr(9) + "dbl_Pa= " + CStr(dbl_Pa) + Chr(13) + _
"dbl_RatedCv= " + CStr(dbl_RatedCv) + Chr(9) + Chr(9) + Chr(9) + "dbl_rho= " + CStr(dbl_rho) + Chr(13) + "dbl_ValveDia= " + CStr(dbl_ValveDia) + Chr(9) + Chr(9) + Chr(9) + _
"dbl_ValveInletDia= " + CStr(dbl_ValveInletDia) + Chr(13) + "dbl_ValveOutletDia= " + CStr(dbl_ValveOutletDia) + Chr(9) + Chr(9) + "dbl_V2= " + CStr(dbl_V2) + Chr(13) + _
"dbl_Z= " + CStr(dbl_Z) + Chr(13) + "dbl_Z1= " + CStr(dbl_Z1) + Chr(9) + Chr(9) + Chr(9) + _
"dbl_Mo= " + CStr(dbl_Mo) + Chr(13) + "dbl_M2= " + CStr(dbl_M2) + Chr(9) + Chr(9) + Chr(9) + "dbl_SL= " + CStr(dbl_SL) + Chr(13) + "dbl_Dsonic= " + CStr(dbl_Dsonic) + Chr(9) + Chr(9) + Chr(9) + "dbl_D3rdsonic= " + CStr(dbl_D3rdsonic) + Chr(13) + _
"dbl_PctTravel= " + CStr(dbl_PctTravel) + Chr(9) + Chr(9) + Chr(9) + "lng_FlowStateFlag= " + CStr(lng_FlowStateFlag)
#End If
    
'look up the reduced area flag
If (str_ValveLine = "") Then
    bln_ReducedArea = True
Else
    Call GetReducedFlag(lng_TrimPackageID, bln_ReducedArea)
End If
    
If str_CalculationType = "Flow" Then
    Call SelectGasFlowCalc(str_CalculationType, _
        dbl_InletPressure, dbl_OutletPressure, dbl_InletTemperature, dbl_FlowW, _
        dbl_K, dbl_rho, dbl_MW, dbl_Z1, dbl_CriticalTemperature, _
        dbl_CriticalPressure, dbl_SL, dbl_Gg, dbl_FlowQ, dbl_Mo, dbl_M2, _
        str_FlowUM, str_PressureUM, str_TemperatureUM, _
        dbl_ValveDia, dbl_ValveInletDia, dbl_ValveOutletDia, _
        dbl_OutletPipeWallThickness, dbl_OutletPipeOD, str_ValveLine, _
        str_BodyMaterial, str_ANSIRating, lng_TrimTypeID, _
        lng_FlowDirectionID, lng_TrimPackageID, dbl_InletDensity, _
        dbl_FL, dbl_Pa, dbl_Cv, dbl_RatedCv, _
        dbl_CalculatedCv, dbl_InletPipeDia, dbl_OutletPipeDia, _
        str_DensityUM, str_LengthUM, str_FluidName, bln_ReducedArea, _
        dbl_Dsonic, dbl_D3rdsonic, dbl_V2, dbl_Z, lng_FlowStateFlag, _
        str_VelocityUM, str_AbsPressureUM, dbl_PctTravel)
    GoTo Leave
End If

int_SelectionFlag = 0
    If str_FluidName = "Steam" Then
        int_SelectionFlag = 5
    ElseIf (str_DensityUM = "kg/m3 1 bar, 0 deg C" Or str_DensityUM = "kg/m3 P1, T1" Or _
        str_DensityUM = "lb/ft3 14.73 psi a, 60 F" Or str_DensityUM = "lb/ft3 P1, T1") And _
        (str_FlowUM = "lb/h" Or str_FlowUM = "kg/h" Or str_FlowUM = "kg/s" Or str_FlowUM = "t/d" Or str_FlowUM = "t/h") Then
        int_SelectionFlag = 1
    ElseIf (str_DensityUM = "MW") And (str_FlowUM = "lb/h" _
        Or str_FlowUM = "kg/h" Or str_FlowUM = "kg/s" Or str_FlowUM = "t/d" Or str_FlowUM = "t/h") Then
        int_SelectionFlag = 2
    ElseIf (str_DensityUM = "MW") And (str_FlowUM = "scfh" Or _
        str_FlowUM = "scfm" Or str_FlowUM = "acfh" Or str_FlowUM = "acfm" Or _
        str_FlowUM = "MM scfd" Or str_FlowUM = "m3/h" Or str_FlowUM = "Nm3/h" Or _
        str_FlowUM = "sm3/h") Then
        int_SelectionFlag = 3
    ElseIf (str_DensityUM = "G") And (str_FlowUM = "scfh" Or _
        str_FlowUM = "scfm" Or str_FlowUM = "acfh" Or str_FlowUM = "acfm" Or _
        str_FlowUM = "MM scfd" Or str_FlowUM = "m3/h" Or str_FlowUM = "Nm3/h" Or _
        str_FlowUM = "sm3/h") Then
        int_SelectionFlag = 4
    Else
        GoTo SelectGasCVCalc_Error
    End If

Select Case int_SelectionFlag
    Case 1
        Call CVDataSet1(dbl_InletPressure, dbl_OutletPressure, _
            dbl_InletTemperature, dbl_FlowW, dbl_K, dbl_rho, dbl_rho1, _
            dbl_W, dbl_MW, dbl_P1, dbl_P2, dbl_T1, str_DensityUM, str_FlowUM, _
            str_PressureUM, str_TemperatureUM)
    Case 2
        Call CVDataSet2(dbl_InletPressure, dbl_OutletPressure, _
            dbl_InletTemperature, dbl_FlowW, dbl_K, dbl_Z, dbl_MW, _
            dbl_CriticalTemperature, dbl_CriticalPressure, dbl_rho1, _
            dbl_Gg, dbl_W, dbl_P1, dbl_P2, dbl_Pc, dbl_T1, dbl_Tc, _
            str_FlowUM, str_PressureUM, str_AbsPressureUM, str_TemperatureUM)
    Case 3
        Call CVDataSet3(dbl_InletPressure, dbl_OutletPressure, _
            dbl_InletTemperature, dbl_FlowQ, dbl_Z, dbl_MW, _
            dbl_CriticalTemperature, dbl_CriticalPressure, dbl_W, _
            dbl_rho1, dbl_Gg, dbl_P1, dbl_P2, dbl_Pc, dbl_T1, dbl_Tc, _
            str_FlowUM, str_PressureUM, str_TemperatureUM)
    Case 4
        Call CVDataSet4(dbl_InletPressure, dbl_OutletPressure, _
            dbl_InletTemperature, dbl_FlowQ, dbl_Z, _
            dbl_Gg, dbl_CriticalTemperature, dbl_CriticalPressure, _
            dbl_rho1, dbl_MW, dbl_W, dbl_P1, dbl_P2, dbl_Pc, dbl_T1, _
            dbl_Tc, str_FlowUM, str_PressureUM, str_TemperatureUM)
    Case 5
        Call CVDataSet5(dbl_InletPressure, dbl_OutletPressure, _
            dbl_InletTemperature, dbl_FlowW, dbl_K, dbl_rho1, _
            dbl_MW, dbl_W, dbl_P1, dbl_P2, dbl_T1, dbl_V2, str_FlowUM, str_PressureUM, _
            str_TemperatureUM)
    Case Else
        GoTo SelectGasCVCalc_Error
End Select

'Start preliminary calculation of 1 to 4 service conditions
dbl_T2 = dbl_T1
Call CalcCV_Data(dbl_ValveDia, dbl_ValveInletDia, _
    dbl_ValveOutletDia, dbl_RatedCv, str_ValveLine, _
    str_BodyMaterial, str_ANSIRating, dbl_rho2, dbl_SL, _
    dbl_P1, dbl_P2, dbl_K, dbl_W, dbl_rho1, dbl_Mo, dbl_M2, dbl_Dsonic, _
    dbl_D3rdsonic, lng_TrimTypeID, lng_FlowDirectionID, _
    lng_IECTrimStyleID, lng_TrimPackageID, str_FluidName, _
    dbl_OutletPipeWallThickness, dbl_OutletPipeOD, dbl_T1, dbl_T2, _
    dbl_CalculatedCv, dbl_MW, dbl_Cv, dbl_InletPipeDia, dbl_OutletPipeDia, _
    str_LengthUM, dbl_Pa, dbl_Z, dbl_Tc, dbl_Pc, dbl_V2, _
    bln_ReducedArea, dbl_FL, lng_FlowStateFlag, dbl_PctTravel)

dbl_Cv = dbl_CalculatedCv

    dbl_PctTravel = TrimDecimalPlaces(dbl_PctTravel)
    dbl_FL = TrimDecimalPlaces(dbl_FL)
    dbl_SL = TrimDecimalPlaces(dbl_SL)
    dbl_Cv = TrimDecimalPlaces(dbl_Cv)
    dbl_Mo = TrimDecimalPlaces(dbl_Mo)
    dbl_M2 = TrimDecimalPlaces(dbl_M2)
    dbl_Dsonic = TrimDecimalPlaces(dbl_Dsonic)
    dbl_D3rdsonic = TrimDecimalPlaces(dbl_D3rdsonic)

Leave:
#If (DebugVersion = "yes") Then
MsgBox "str_CalculationType= " + str_CalculationType + Chr(9) + Chr(9) + "bln_ReducedArea= " + CStr(bln_ReducedArea) + Chr(13) + "lng_FlowDirectionID= " + CStr(lng_FlowDirectionID) + Chr(9) + Chr(9) + _
"lng_TrimPackageID= " + CStr(lng_TrimPackageID) + Chr(13) + "lng_TrimTypeID= " + CStr(lng_TrimTypeID) + Chr(9) + Chr(9) + Chr(9) + "str_AbsPressureUM= " + str_AbsPressureUM + Chr(13) + "str_VelocityUM= " + str_VelocityUM + Chr(9) + Chr(9) + "str_DensityUM= " + str_DensityUM + Chr(13) + "str_FlowUM= " + str_FlowUM + Chr(13) + _
"str_LengthUM= " + str_LengthUM + Chr(9) + Chr(9) + Chr(9) + "str_PressureUM= " + str_PressureUM + Chr(13) + "str_TemperatureUM= " + str_TemperatureUM + Chr(9) + Chr(9) + "str_ANSIRating= " + str_ANSIRating + Chr(13) + _
"str_BodyMaterial= " + CStr(str_BodyMaterial) + Chr(9) + Chr(9) + Chr(9) + "str_FluidName= " + str_FluidName + Chr(13) + "str_ValveLine= " + str_ValveLine + Chr(9) + Chr(9) + "dbl_InletPressure= " + CStr(dbl_InletPressure) + Chr(13) + _
"dbl_OutletPressure= " + CStr(dbl_OutletPressure) + Chr(9) + Chr(9) + "dbl_InletTemperature= " + CStr(dbl_InletTemperature) + Chr(13) + "dbl_CriticalTemperature= " + CStr(dbl_CriticalTemperature) + Chr(9) + Chr(9) + "dbl_CriticalPressure= " + CStr(dbl_CriticalPressure) + Chr(13) + _
"dbl_FlowQ= " + CStr(dbl_FlowQ) + Chr(9) + Chr(9) + Chr(9) + "dbl_FlowW= " + CStr(dbl_FlowW) + Chr(13) + "dbl_k= " + CStr(dbl_K) + Chr(9) + Chr(9) + Chr(9) + "dbl_CalculatedCv= " + CStr(dbl_CalculatedCv) + Chr(13) + _
"dbl_Cv= " + CStr(dbl_Cv) + Chr(9) + Chr(9) + Chr(9) + "dbl_FL= " + CStr(dbl_FL) + Chr(13) + _
"dbl_Gg= " + CStr(dbl_Gg)
#End If

#If (DebugVersion = "yes") Then
MsgBox "dbl_InletPipeDia= " + CStr(dbl_InletPipeDia) + Chr(13) + _
"dbl_MW= " + CStr(dbl_MW) + Chr(13) + "dbl_OutletPipeDia= " + CStr(dbl_OutletPipeDia) + Chr(13) + _
"dbl_OutletPipeWallThickness= " + CStr(dbl_OutletPipeWallThickness) + Chr(13) + "dbl_OutletPipeOD= " + CStr(dbl_OutletPipeOD) + Chr(9) + Chr(9) + "dbl_Pa= " + CStr(dbl_Pa) + Chr(13) + _
"dbl_RatedCv= " + CStr(dbl_RatedCv) + Chr(9) + Chr(9) + Chr(9) + "dbl_rho= " + CStr(dbl_rho) + Chr(13) + "dbl_ValveDia= " + CStr(dbl_ValveDia) + Chr(9) + Chr(9) + Chr(9) + _
"dbl_ValveInletDia= " + CStr(dbl_ValveInletDia) + Chr(13) + "dbl_ValveOutletDia= " + CStr(dbl_ValveOutletDia) + Chr(9) + Chr(9) + "dbl_V2= " + CStr(dbl_V2) + Chr(13) + _
"dbl_Z= " + CStr(dbl_Z) + Chr(13) + "dbl_Z1= " + CStr(dbl_Z1) + Chr(9) + Chr(9) + Chr(9) + _
"dbl_Mo= " + CStr(dbl_Mo) + Chr(13) + "dbl_M2= " + CStr(dbl_M2) + Chr(9) + Chr(9) + Chr(9) + "dbl_SL= " + CStr(dbl_SL) + Chr(13) + "dbl_Dsonic= " + CStr(dbl_Dsonic) + Chr(9) + Chr(9) + Chr(9) + "dbl_D3rdsonic= " + CStr(dbl_D3rdsonic) + Chr(13) + _
"dbl_PctTravel= " + CStr(dbl_PctTravel) + Chr(9) + Chr(9) + Chr(9) + "lng_FlowStateFlag= " + CStr(lng_FlowStateFlag)
#End If

Exit Sub
SelectGasCVCalc_Error:
'Insert Error Message
End Sub



Public Sub SelectGasFlowCalc(str_CalculationType As String, _
    dbl_InletPressure As Double, dbl_OutletPressure As Double, _
    dbl_InletTemperature As Double, dbl_FlowW As Double, dbl_K As Double, _
    dbl_rho As Double, dbl_MW As Double, dbl_Z1 As Double, _
    dbl_CriticalTemperature As Double, dbl_CriticalPressure As Double, _
    dbl_SL As Double, dbl_Gg As Double, _
    dbl_FlowQ As Double, dbl_Mo As Double, dbl_M2 As Double, _
    str_FlowUM As String, str_PressureUM As String, str_TemperatureUM As String, _
    dbl_ValveDia As Double, dbl_ValveInletDia As Double, _
    dbl_ValveOutletDia As Double, dbl_OutletPipeWallThickness As Double, _
    dbl_OutletPipeOD As Double, str_ValveLine As String, _
    str_BodyMaterial As String, str_ANSIRating As String, lng_TrimTypeID As Long, _
    lng_FlowDirectionID As Long, lng_TrimPackageID As Long, _
    dbl_InletDensity As Double, dbl_FL As Double, dbl_Pa As Double, _
    dbl_Cv As Double, dbl_RatedCv As Double, dbl_CalculatedCv As Double, _
    dbl_InletPipeDia As Double, dbl_OutletPipeDia As Double, _
    str_DensityUM As String, str_LengthUM As String, str_FluidName As String, _
    bln_ReducedArea As Boolean, dbl_Dsonic As Double, dbl_D3rdsonic As Double, _
    dbl_V2 As Double, dbl_Z As Double, lng_FlowStateFlag As Long, _
    str_VelocityUM As String, str_AbsPressureUM As String, dbl_PctTravel As Double)


On Error GoTo SelectGasFlowCalc_Error:
Dim int_SelectionFlag As Integer
Dim dbl_MassFlowRate As Double
Dim dbl_P1 As Double
Dim dbl_P2 As Double
Dim dbl_T1 As Double
Dim dbl_T2 As Double
Dim dbl_W As Double
Dim dbl_rho1 As Double
Dim dbl_rho2 As Double
Dim lng_IECTrimStyleID As Long

#If (DebugVersion = "yes") Then
MsgBox "str_CalculationType= " + str_CalculationType + Chr(9) + Chr(9) + "bln_ReducedArea= " + CStr(bln_ReducedArea) + Chr(13) + "lng_FlowDirectionID= " + CStr(lng_FlowDirectionID) + Chr(9) + Chr(9) + _
"lng_TrimPackageID= " + CStr(lng_TrimPackageID) + Chr(13) + "lng_TrimTypeID= " + CStr(lng_TrimTypeID) + Chr(9) + Chr(9) + Chr(9) + "str_AbsPressureUM= " + str_AbsPressureUM + Chr(13) + "str_VelocityUM= " + str_VelocityUM + Chr(9) + Chr(9) + "str_DensityUM= " + str_DensityUM + Chr(13) + "str_FlowUM= " + str_FlowUM + Chr(13) + _
"str_LengthUM= " + str_LengthUM + Chr(9) + Chr(9) + Chr(9) + "str_PressureUM= " + str_PressureUM + Chr(13) + "str_TemperatureUM= " + str_TemperatureUM + Chr(9) + Chr(9) + "str_ANSIRating= " + str_ANSIRating + Chr(13) + _
"str_BodyMaterial= " + CStr(str_BodyMaterial) + Chr(9) + Chr(9) + Chr(9) + "str_FluidName= " + str_FluidName + Chr(13) + "str_ValveLine= " + str_ValveLine + Chr(9) + Chr(9) + "dbl_InletPressure= " + CStr(dbl_InletPressure) + Chr(13) + _
"dbl_OutletPressure= " + CStr(dbl_OutletPressure) + Chr(9) + Chr(9) + "dbl_InletTemperature= " + CStr(dbl_InletTemperature) + Chr(13) + "dbl_CriticalTemperature= " + CStr(dbl_CriticalTemperature) + Chr(9) + Chr(9) + "dbl_CriticalPressure= " + CStr(dbl_CriticalPressure) + Chr(13) + _
"dbl_FlowQ= " + CStr(dbl_FlowQ) + Chr(9) + Chr(9) + Chr(9) + "dbl_FlowW= " + CStr(dbl_FlowW) + Chr(13) + "dbl_k= " + CStr(dbl_K) + Chr(9) + Chr(9) + Chr(9) + "dbl_CalculatedCv= " + CStr(dbl_CalculatedCv) + Chr(13) + _
"dbl_Cv= " + CStr(dbl_Cv) + Chr(9) + Chr(9) + Chr(9) + "dbl_FL= " + CStr(dbl_FL) + Chr(13) + _
"dbl_Gg= " + CStr(dbl_Gg)
#End If

#If (DebugVersion = "yes") Then
MsgBox "dbl_InletDensity= " + CStr(dbl_InletDensity) + Chr(9) + Chr(9) + "dbl_InletPipeDia= " + CStr(dbl_InletPipeDia) + Chr(13) + _
"dbl_MW= " + CStr(dbl_MW) + Chr(13) + "dbl_OutletPipeDia= " + CStr(dbl_OutletPipeDia) + Chr(13) + _
"dbl_OutletPipeWallThickness= " + CStr(dbl_OutletPipeWallThickness) + Chr(13) + "dbl_OutletPipeOD= " + CStr(dbl_OutletPipeOD) + Chr(9) + Chr(9) + "dbl_Pa= " + CStr(dbl_Pa) + Chr(13) + _
"dbl_RatedCv= " + CStr(dbl_RatedCv) + Chr(9) + Chr(9) + Chr(9) + "dbl_rho= " + CStr(dbl_rho) + Chr(13) + "dbl_ValveDia= " + CStr(dbl_ValveDia) + Chr(9) + Chr(9) + Chr(9) + _
"dbl_ValveInletDia= " + CStr(dbl_ValveInletDia) + Chr(13) + "dbl_ValveOutletDia= " + CStr(dbl_ValveOutletDia) + Chr(9) + Chr(9) + "dbl_V2= " + CStr(dbl_V2) + Chr(13) + _
"dbl_Z= " + CStr(dbl_Z) + Chr(13) + "dbl_Z1= " + CStr(dbl_Z1) + Chr(9) + Chr(9) + Chr(9) + _
"dbl_Mo= " + CStr(dbl_Mo) + Chr(13) + "dbl_M2= " + CStr(dbl_M2) + Chr(9) + Chr(9) + Chr(9) + "dbl_SL= " + CStr(dbl_SL) + Chr(13) + "dbl_Dsonic= " + CStr(dbl_Dsonic) + Chr(9) + Chr(9) + Chr(9) + "dbl_D3rdsonic= " + CStr(dbl_D3rdsonic) + Chr(13) + _
"dbl_PctTravel= " + CStr(dbl_PctTravel) + Chr(9) + Chr(9) + Chr(9) + "lng_FlowStateFlag= " + CStr(lng_FlowStateFlag)
#End If




If str_CalculationType = "CV" Then
    Call SelectGasCVCalc(str_CalculationType, _
        dbl_InletPressure, dbl_OutletPressure, dbl_InletTemperature, dbl_FlowW, _
        dbl_K, dbl_rho, dbl_MW, dbl_Z1, dbl_CriticalTemperature, _
        dbl_CriticalPressure, dbl_SL, dbl_Gg, dbl_FlowQ, dbl_Mo, dbl_M2, _
        str_FlowUM, str_PressureUM, str_TemperatureUM, _
        dbl_ValveDia, dbl_ValveInletDia, dbl_ValveOutletDia, _
        dbl_OutletPipeWallThickness, dbl_OutletPipeOD, str_ValveLine, _
        str_BodyMaterial, str_ANSIRating, lng_TrimTypeID, _
        lng_FlowDirectionID, lng_TrimPackageID, _
        dbl_InletDensity, dbl_FL, dbl_Pa, dbl_Cv, dbl_RatedCv, _
        dbl_CalculatedCv, dbl_InletPipeDia, dbl_OutletPipeDia, _
        str_DensityUM, str_LengthUM, str_FluidName, bln_ReducedArea, _
        dbl_Dsonic, dbl_D3rdsonic, dbl_V2, dbl_Z, lng_FlowStateFlag, _
        str_VelocityUM, str_AbsPressureUM, dbl_PctTravel)
    
    Exit Sub
End If


int_SelectionFlag = 0
    If str_FluidName = "Steam" Then
        int_SelectionFlag = 5
    ElseIf (str_DensityUM = "kg/m3 1 bar, 0 deg C" Or str_DensityUM = "kg/m3 P1, T1" Or _
        str_DensityUM = "lb/ft3 14.73 psi a, 60 F" Or str_DensityUM = "lb/ft3 P1, T1") And _
        (str_FlowUM = "lb/h" Or str_FlowUM = "kg/h" Or str_FlowUM = "kg/s" Or str_FlowUM = "t/d" Or str_FlowUM = "t/h") Then
        int_SelectionFlag = 1
    ElseIf (str_DensityUM = "MW") And (str_FlowUM = "lb/h" _
        Or str_FlowUM = "kg/h" Or str_FlowUM = "kg/s" Or str_FlowUM = "t/d" Or str_FlowUM = "t/h") Then
        int_SelectionFlag = 2
    ElseIf (str_DensityUM = "MW") And (str_FlowUM = "scfh" Or _
        str_FlowUM = "scfm" Or str_FlowUM = "acfh" Or str_FlowUM = "acfm" Or _
        str_FlowUM = "MM scfd" Or str_FlowUM = "m3/h" Or str_FlowUM = "Nm3/h" Or _
        str_FlowUM = "sm3/h") Then
        int_SelectionFlag = 3
    ElseIf (str_DensityUM = "G") And (str_FlowUM = "scfh" Or _
        str_FlowUM = "scfm" Or str_FlowUM = "acfh" Or str_FlowUM = "acfm" Or _
        str_FlowUM = "MM scfd" Or str_FlowUM = "m3/h" Or str_FlowUM = "Nm3/h" Or _
        str_FlowUM = "sm3/h") Then
        int_SelectionFlag = 4
    ElseIf str_FluidName = "Steam" Then
        int_SelectionFlag = 5
    Else
        GoTo SelectGasFlowCalc_Error
    End If

Select Case int_SelectionFlag
    Case 1
        
        Call FlowDataSet1(dbl_InletPressure, dbl_OutletPressure, _
            dbl_InletTemperature, dbl_K, dbl_rho, dbl_rho1, dbl_FlowW, _
            dbl_MW, dbl_P1, dbl_P2, dbl_T1, str_DensityUM, str_FlowUM, _
            str_PressureUM, str_TemperatureUM)
    Case 2
        Call FlowDataSet2(dbl_InletPressure, dbl_OutletPressure, _
            dbl_InletTemperature, dbl_K, dbl_Z1, dbl_MW, _
            dbl_CriticalTemperature, dbl_CriticalPressure, dbl_FlowW, _
            dbl_rho1, dbl_P1, dbl_P2, dbl_Pc, dbl_T1, dbl_Tc, _
            str_DensityUM, str_FlowUM, _
            str_PressureUM, str_TemperatureUM)
    Case 3
        Call FlowDataSet3(dbl_InletPressure, dbl_OutletPressure, _
            dbl_InletTemperature, dbl_K, dbl_Z1, dbl_Gg, dbl_MW, _
            dbl_CriticalTemperature, dbl_CriticalPressure, dbl_FlowQ, _
            dbl_W, dbl_rho1, dbl_P1, dbl_P2, dbl_Pc, dbl_T1, _
            dbl_Tc, str_DensityUM, str_FlowUM, str_PressureUM, str_TemperatureUM)
    Case 4
        Call FlowDataSet4(dbl_InletPressure, dbl_OutletPressure, _
            dbl_InletTemperature, dbl_FlowQ, dbl_K, dbl_Z1, dbl_Gg, _
            dbl_CriticalTemperature, dbl_CriticalPressure, _
            dbl_W, dbl_MW, dbl_rho1, dbl_P1, dbl_P2, dbl_Pc, dbl_T1, _
            dbl_Tc, str_DensityUM, str_FlowUM, str_PressureUM, str_TemperatureUM)
    Case 5
        Call FlowDataSet5(dbl_InletPressure, dbl_OutletPressure, _
            dbl_InletTemperature, dbl_FlowW, dbl_K, dbl_rho1, _
            dbl_MW, dbl_P1, dbl_P2, dbl_T1, str_DensityUM, str_FlowUM, _
            str_PressureUM, str_TemperatureUM)
    Case Else
        GoTo SelectGasFlowCalc_Error
End Select

'Start preliminary calculation of 1 to 4 service conditions
'Continue at: Select Valve Model, Size, Rated Cv, Body Material,
'ANSI Rating.
dbl_T2 = dbl_T1
Call CalcFlow_Data(dbl_ValveDia, dbl_ValveInletDia, _
     dbl_ValveOutletDia, dbl_RatedCv, str_ValveLine, _
    str_BodyMaterial, str_ANSIRating, dbl_rho2, dbl_SL, _
    dbl_P1, dbl_P2, dbl_K, dbl_W, dbl_rho1, dbl_Mo, dbl_M2, _
    dbl_Dsonic, dbl_D3rdsonic, lng_TrimTypeID, lng_FlowDirectionID, _
    lng_IECTrimStyleID, lng_TrimPackageID, str_FluidName, _
    dbl_InletDensity, dbl_OutletPipeWallThickness, _
    dbl_OutletPipeOD, dbl_T1, dbl_T2, dbl_CalculatedCv, dbl_MW, dbl_Cv, _
    dbl_InletPipeDia, dbl_OutletPipeDia, str_LengthUM, dbl_Pa, dbl_Z, _
    dbl_Tc, dbl_Pc, dbl_V2, bln_ReducedArea, str_FlowUM, dbl_FL, _
    lng_FlowStateFlag, dbl_PctTravel)

    dbl_Dsonic = TrimDecimalPlaces(dbl_Dsonic)
    dbl_D3rdsonic = TrimDecimalPlaces(dbl_D3rdsonic)
    dbl_Mo = TrimDecimalPlaces(dbl_Mo)
    dbl_M2 = TrimDecimalPlaces(dbl_M2)
    dbl_FL = TrimDecimalPlaces(dbl_FL)
    dbl_SL = TrimDecimalPlaces(dbl_SL)
    dbl_W = TrimDecimalPlaces(dbl_W)


#If (DebugVersion = "yes") Then
MsgBox "str_CalculationType= " + str_CalculationType + Chr(9) + Chr(9) + "bln_ReducedArea= " + CStr(bln_ReducedArea) + Chr(13) + "lng_FlowDirectionID= " + CStr(lng_FlowDirectionID) + Chr(9) + Chr(9) + _
"lng_TrimPackageID= " + CStr(lng_TrimPackageID) + Chr(13) + "lng_TrimTypeID= " + CStr(lng_TrimTypeID) + Chr(9) + Chr(9) + Chr(9) + "str_AbsPressureUM= " + str_AbsPressureUM + Chr(13) + "str_VelocityUM= " + str_VelocityUM + Chr(9) + Chr(9) + "str_DensityUM= " + str_DensityUM + Chr(13) + "str_FlowUM= " + str_FlowUM + Chr(13) + _
"str_LengthUM= " + str_LengthUM + Chr(9) + Chr(9) + Chr(9) + "str_PressureUM= " + str_PressureUM + Chr(13) + "str_TemperatureUM= " + str_TemperatureUM + Chr(9) + Chr(9) + "str_ANSIRating= " + str_ANSIRating + Chr(13) + _
"str_BodyMaterial= " + CStr(str_BodyMaterial) + Chr(9) + Chr(9) + Chr(9) + "str_FluidName= " + str_FluidName + Chr(13) + "str_ValveLine= " + str_ValveLine + Chr(9) + Chr(9) + "dbl_InletPressure= " + CStr(dbl_InletPressure) + Chr(13) + _
"dbl_OutletPressure= " + CStr(dbl_OutletPressure) + Chr(9) + Chr(9) + "dbl_InletTemperature= " + CStr(dbl_InletTemperature) + Chr(13) + "dbl_CriticalTemperature= " + CStr(dbl_CriticalTemperature) + Chr(9) + Chr(9) + "dbl_CriticalPressure= " + CStr(dbl_CriticalPressure) + Chr(13) + _
"dbl_FlowQ= " + CStr(dbl_FlowQ) + Chr(9) + Chr(9) + Chr(9) + "dbl_FlowW= " + CStr(dbl_FlowW) + Chr(13) + "dbl_k= " + CStr(dbl_K) + Chr(9) + Chr(9) + Chr(9) + "dbl_CalculatedCv= " + CStr(dbl_CalculatedCv) + Chr(13) + _
"dbl_Cv= " + CStr(dbl_Cv) + Chr(9) + Chr(9) + Chr(9) + "dbl_FL= " + CStr(dbl_FL) + Chr(13) + _
"dbl_Gg= " + CStr(dbl_Gg)
#End If

#If (DebugVersion = "yes") Then
MsgBox "dbl_InletDensity= " + CStr(dbl_InletDensity) + Chr(9) + Chr(9) + "dbl_InletPipeDia= " + CStr(dbl_InletPipeDia) + Chr(13) + _
"dbl_MW= " + CStr(dbl_MW) + Chr(13) + "dbl_OutletPipeDia= " + CStr(dbl_OutletPipeDia) + Chr(13) + _
"dbl_OutletPipeWallThickness= " + CStr(dbl_OutletPipeWallThickness) + Chr(13) + "dbl_OutletPipeOD= " + CStr(dbl_OutletPipeOD) + Chr(9) + Chr(9) + "dbl_Pa= " + CStr(dbl_Pa) + Chr(13) + _
"dbl_RatedCv= " + CStr(dbl_RatedCv) + Chr(9) + Chr(9) + Chr(9) + "dbl_rho= " + CStr(dbl_rho) + Chr(13) + "dbl_ValveDia= " + CStr(dbl_ValveDia) + Chr(9) + Chr(9) + Chr(9) + _
"dbl_ValveInletDia= " + CStr(dbl_ValveInletDia) + Chr(13) + "dbl_ValveOutletDia= " + CStr(dbl_ValveOutletDia) + Chr(9) + Chr(9) + "dbl_V2= " + CStr(dbl_V2) + Chr(13) + _
"dbl_Z= " + CStr(dbl_Z) + Chr(13) + "dbl_Z1= " + CStr(dbl_Z1) + Chr(9) + Chr(9) + Chr(9) + _
"dbl_Mo= " + CStr(dbl_Mo) + Chr(13) + "dbl_M2= " + CStr(dbl_M2) + Chr(9) + Chr(9) + Chr(9) + "dbl_SL= " + CStr(dbl_SL) + Chr(13) + "dbl_Dsonic= " + CStr(dbl_Dsonic) + Chr(9) + Chr(9) + Chr(9) + "dbl_D3rdsonic= " + CStr(dbl_D3rdsonic) + Chr(13) + _
"dbl_PctTravel= " + CStr(dbl_PctTravel) + Chr(9) + Chr(9) + Chr(9) + "lng_FlowStateFlag= " + CStr(lng_FlowStateFlag)
#End If

Exit Sub
SelectGasFlowCalc_Error:
'Insert Error Message
End Sub

Sub GasParameters(str_FluidName As String, dbl_InletTemperature As Double, _
    str_TemperatureUM As String, dbl_Pressure As Double, _
    str_PressureUM As String, dbl_MW As Double, dbl_K As Double, dbl_Z As Double)
'*****************************************************************
' looks up standard gas data
'
'*****************************************************************
'Input Variables:
'str_FluidName      name of the gas
'dbl_Temperature    inlet temperature
'str_TemperatureUM  units of measure of the temperature
'dbl_Pressure       inlet pressure
'str_PressureUM     units of measure of the pressure
'
'Output Variables:
'dbl_MW             molecular weight
'dbl_K              specific heat
'dbl_Z              compressibility factor
'*****************************************************************
'   written by Schoonover 12/8/98
'*****************************************************************
On Error GoTo GasParameters_Error
'**********************************
'Check input validation
'**********************************

#If (DebugVersion = "yes") Then
MsgBox "GetGasParam" + Chr(13) + Chr(10) + _
    "FluidName=" + str_FluidName + Chr(13) + Chr(10) + _
    "Temperature=" + CStr(dbl_InletTemperature) + Chr(13) + Chr(10) + _
    "TemperatureUm=" + str_TemperatureUM + Chr(13) + Chr(10) + _
    "Pressure=" + CStr(dbl_Pressure) + Chr(13) + Chr(10) + _
    "PressureUM=" + str_PressureUM + Chr(13) + Chr(10) + _
    "dbl_MW=" + CStr(dbl_MW) + Chr(13) + Chr(10) + _
    "dbl_K=" + CStr(dbl_K) + Chr(13) + Chr(10) + _
    "dbl_Z=" + CStr(dbl_Z)
#End If


If (str_TemperatureUM = "") Then GoTo GasParameters_Error
If (str_PressureUM = "") Then GoTo GasParameters_Error
If (str_FluidName = "") Then GoTo GasParameters_Error
'**********************************
'Declaration
'**********************************
 Dim str_Query As String
 Dim str_Path As String
 Dim rst_Results As Recordset
 Dim dbl_Pc As Double
 Dim dbl_Tc As Double
 Dim dbl_Tr As Double
 Dim dbl_Pr As Double
 Dim dbl_Tk As Double
 Dim dbl_Pb As Double
 
'**********************************
'Begin Calculation
'**********************************
str_Query = "Select Pc, Tc, MW, K from tblGasStdParam where FLUID_NAME='" + str_FluidName + "'"
Set rst_Results = g_db_Gensz.OpenRecordset(str_Query)
If (rst_Results.RecordCount <> 1) Then GoTo GasParameters_Error
dbl_Pc = rst_Results.Fields!Pc
dbl_Tc = rst_Results.Fields!Tc
dbl_MW = rst_Results.Fields!MW
dbl_K = rst_Results.Fields!K

rst_Results.Close
    
    
'convert to proper units
Call CnvrtTemperatureToDegK(dbl_InletTemperature, str_TemperatureUM, dbl_Tk)
Call CnvrtPressureToBarA(dbl_Pressure, str_PressureUM, dbl_Pb)

dbl_Z = CompressibilityFactor(dbl_Tk, dbl_Tc, dbl_Pb, dbl_Pc)

#If (DebugVersion = "yes") Then
MsgBox "GetGasParam" + Chr(13) + Chr(10) + _
    "FluidName=" + str_FluidName + Chr(13) + Chr(10) + _
    "Temperature=" + CStr(dbl_InletTemperature) + Chr(13) + Chr(10) + _
    "TemperatureUm=" + str_TemperatureUM + Chr(13) + Chr(10) + _
    "Pressure=" + CStr(dbl_Pressure) + Chr(13) + Chr(10) + _
    "PressureUM=" + str_PressureUM + Chr(13) + Chr(10) + _
    "dbl_MW=" + CStr(dbl_MW) + Chr(13) + Chr(10) + _
    "dbl_K=" + CStr(dbl_K) + Chr(13) + Chr(10) + _
    "dbl_Z=" + CStr(dbl_Z)
#End If

Exit Sub
GasParameters_Error:
MsgBox "Error getting gas parameters " + str_FluidName + Chr(13) + Chr(10) + Error
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' END OF gas LL routines
''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INCLUDE  Liquid Routines
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GetLodbValveDataForNoise(str_ValveLine As String, long_TrimTypeID As Long, _
    dbl_HolesPerCv As Double, dbl_LDRatio As Double)

On Error GoTo GetLodbValveDataForNoise_Error
'--------------------------------------------------------------------
'This sub returns the Fd value for the product at the calculated Cv
'--------------------------------------------------------------------
Dim str_DbPath As String
Dim rst_Results As Recordset
Dim str_Query As String

'Select data
str_Query = "Select NUMBER_HOLES_PER_CV, LD from tblHolesPerCv where VALVE_LINE= '" _
    + str_ValveLine + "' and (TRIM_TYPE_ID=" + CStr(long_TrimTypeID) + " or TRIM_TYPE_ID" _
    & " is null)"

Set rst_Results = g_db_Gensz.OpenRecordset(str_Query)
If (rst_Results.RecordCount = 0) Then GoTo GetLodbValveDataForNoise_Error
dbl_HolesPerCv = IIf(IsNull(rst_Results.Fields!NUMBER_HOLES_PER_CV), 0, rst_Results.Fields!NUMBER_HOLES_PER_CV)
dbl_LDRatio = IIf(IsNull(rst_Results.Fields!LD), 0, rst_Results.Fields!LD)

rst_Results.Close


Exit Sub

GetLodbValveDataForNoise_Error:

End Sub

Sub GetStdValveDataForNoise(str_ValveLine As String, _
    lng_FlowDirectionID As Long, lng_TrimTypeID As Long, _
    dbl_CalculatedCv As Double, _
    dbl_RatedCv As Double, dbl_Fd As Double)

On Error GoTo GetStdValveDataForNoise_Error
'--------------------------------------------------------------------
'This sub returns the Fd value for the product at the calculated Cv
'
'written 11/20/98 Schoonover
'--------------------------------------------------------------------
Dim str_DbPath As String
Dim rst_Results As Recordset
Dim str_Query As String
Dim lng_FD_DATA_ID As Long
Dim dbl_PctCv As Double
Dim dbl_PctCvLowerBound As Double
Dim dbl_PctCvUpperBound As Double
Dim dbl_FdLowerBound As Double
Dim dbl_FdUpperBound As Double
Dim dbl_LocalRatedCV As Double

dbl_LocalRatedCV = dbl_RatedCv
If (dbl_LocalRatedCV = 0) Then dbl_LocalRatedCV = 2 * dbl_CalculatedCv

'Select FD_DATA_ID from tblFdKey
str_Query = "Select FD_DATA_ID from tblFdKey where VALVE_LINE= '" _
    + str_ValveLine + "' and (FLOW_DIRECTION_ID = " + Str(lng_FlowDirectionID) _
    + " or FLOW_DIRECTION_ID is null) and (TRIM_TYPE_ID= " _
    + Str(lng_TrimTypeID) + " or TRIM_TYPE_ID is null)"

Set rst_Results = g_db_Gensz.OpenRecordset(str_Query)

If (rst_Results.RecordCount = 0) Then GoTo GetStdValveDataForNoise_Error
If IsNull(rst_Results.Fields!FD_DATA_ID) Then GoTo GetStdValveDataForNoise_Error

lng_FD_DATA_ID = rst_Results.Fields!FD_DATA_ID
rst_Results.Close



'Calculated %RatedCv
dbl_PctCv = (dbl_CalculatedCv / dbl_LocalRatedCV) * 100
If dbl_PctCv >= 100 Then dbl_PctCv = 99.999
If (dbl_PctCv < 10) Then dbl_PctCv = 10

'Determine upper and lower bounds of the range dbl_PctCv falls in
If dbl_PctCv < 20 Then
    GetBounds dbl_PctCvLowerBound, dbl_PctCvUpperBound, dbl_PctCv, 10
Else
    GetBounds dbl_PctCvLowerBound, dbl_PctCvUpperBound, dbl_PctCv, 20
End If

'Retrieve Fd values from tblFd
str_Query = "Select FD" + CStr(dbl_PctCvLowerBound) + ", FD" _
+ CStr(dbl_PctCvUpperBound) + " from tblFd where FD_DATA_ID= " _
+ CStr(lng_FD_DATA_ID)

Set rst_Results = g_db_Gensz.OpenRecordset(str_Query)
With rst_Results
    dbl_FdLowerBound = IIf(IsNull(.Fields(0)), 0#, CDbl(.Fields(0)))
    dbl_FdUpperBound = IIf(IsNull(.Fields(1)), 0#, CDbl(.Fields(1)))
End With
rst_Results.Close

'Interpolate to find Fd
Interpolate dbl_PctCv, dbl_PctCvLowerBound, dbl_PctCvUpperBound, _
dbl_FdLowerBound, dbl_FdUpperBound, dbl_Fd

Exit Sub

GetStdValveDataForNoise_Error:
End Sub

Sub CvFlashingLiq(dbl_P1 As Double, dbl_P2 As Double, dbl_Pv As Double, _
dbl_Pc As Double, dbl_Gf As Double, dbl_Q As Double, dbl_W As Double, _
dbl_FL As Double, dbl_Fp As Double, dbl_Flp As Double, dbl_InletPipeDia As Double, _
dbl_ValveInletDia As Double, dbl_Cv As Double)

'------------------------------------------------------------------------------------------------
'This subroutine calculates the Cv for a flashing liquid with critical flow conditions.
'
'Reference: ISA S75.01 equations 12a and 12b
'           IEC 534-2-1 equations 3 and 4
'------------------------------------------------------------------------------------------------
'Input Variables                    output                    description                         units
'
'dbl_P1                                             inlet pressure                      bar a
'dbl_P2                                             outlet pressure                     bar a
'dbl_Pv                                             vapor pressure                      bar a
'dbl_Pc                                             critical pressure                   bar a
'r1/r0 » dbl_Gf                                     specific gravity or relative density
'dbl_W                                              flowrate (weight basis)             kg/h
'dbl_Q                                              flowrate (volume basis)             m3/h
'dbl_FL                                             liquid pressure recovery factor
'dbl_Fp                                             pipe geometry factor
'dbl_Flp                                            combined pipe geometry factor
'dbl_InletPipeDia                                   nominal pipe size                   inches/mm
'dbl_ValveInletDia                                  nominal valve size                  inches/mm
'                         dbl_Cv                    valve flow coefficient
'
'Author:  KPMG                  Date:  6/12/98
'------------------------------------------------------------------------------------------------

On Error GoTo CvFlashingLiquid_Error

'*****************
'Check input values
'*****************

'Pc must be greater than zero
If dbl_Pc < 0 Then GoTo CvFlashingLiquid_Error
'One and only one flowrate should be specified, and flowrates cannot be negative
If (dbl_W < 0 And dbl_Q < 0) Or (dbl_W > 0 And dbl_Q > 0) Then GoTo CvFlashingLiquid_Error
'inletValveDia and inletPipeDia must both be greater than zero
If (dbl_ValveInletDia < 0 Or dbl_InletPipeDia < 0) Then GoTo CvFlashingLiquid_Error

'****************
'Declarations
'****************

Dim dbl_Ff As Double

'*******************
'Begin Calculations
'*******************

dbl_Ff = 0.96 - 0.28 * Sqr(dbl_Pv / dbl_Pc)
    If dbl_InletPipeDia = dbl_ValveInletDia Then
        'Flowrate is in volume units
        If dbl_Q > 0 Then dbl_Cv = (dbl_Q / (0.865 * dbl_FL)) * Sqr(dbl_Gf / (dbl_P1 - dbl_Ff * dbl_Pv))
        'Flowrate is in weight units
        If dbl_W > 0 Then dbl_Cv = dbl_W / (27.3 * dbl_FL * Sqr((dbl_P1 - dbl_Ff * dbl_Pv) * dbl_Gf * 1000))
    Else
        'Flowrate is in volume units
        If dbl_Q > 0 Then dbl_Cv = (dbl_Q / (0.865 * dbl_Flp)) * Sqr(dbl_Gf / (dbl_P1 - dbl_Ff * dbl_Pv))
        'Flowrate is in weight units
        If dbl_W > 0 Then dbl_Cv = dbl_W / (27.3 * dbl_Flp * Sqr((dbl_P1 - dbl_Ff * dbl_Pv) * dbl_Gf * 1000))
    End If
        
Exit Sub

CvFlashingLiquid_Error:
'Placeholder

End Sub
Sub CvCriticalLiq(dbl_P1 As Double, dbl_P2 As Double, dbl_Pv As Double, dbl_Pc As Double, _
dbl_Gf As Double, dbl_Q As Double, dbl_W As Double, dbl_FL As Double, dbl_Fp As Double, _
dbl_Flp As Double, dbl_InletPipeDia As Double, dbl_ValveInletDia As Double, dbl_Cv As Double)

'------------------------------------------------------------------------------------------------
'This subroutine calculates the Cv for a flashing liquid with critical flow conditions.
'
'Reference: ISA S75.01
'                  IEC 534 - 2 - 1
'------------------------------------------------------------------------------------------------
'Input Variables                    output                                description                     units
'
'dbl_P1                                                         inlet pressure                  bar a
'dbl_P2                                                         outlet pressure                 bar a
'dbl_Pv                                                         vapor pressure                  bar a
'dbl_Pc                                                         critical pressure               bar a
'r1/r0 » dbl_Gf                                                 specific gravity or relative density
'dbl_W                                                          flowrate (weight basis)         kg/h
'dbl_Q                                                          flowrate (volume basis)         m3/h
'dbl_FL                                                         liquid pressure recovery factor
'dbl_Fp                                                         pipe geometry factor
'dbl_Flp                                                        combined pipe geometry factor
'dbl_InletPipeDia                                               nominal pipe size               inches/mm
'dbl_ValveInletDia                                              nominal valve size              inches/mm
'                         dbl_Cv                                valve flow coefficient
'
''Author:  KPMG                  Date:  6/12/98
'------------------------------------------------------------------------------------------------

On Error GoTo CvCriticalLiquid_Error

'*****************
'Check input values
'*****************

'dbl_Pc must be greater than zero
If dbl_Pc <= 0 Then GoTo CvCriticalLiquid_Error
'One and only one flowrate should be specified, and flowrates cannot be negative
If (dbl_W > 0 And dbl_Q <> 0) Or (dbl_Q > 0 And dbl_W <> 0) Then GoTo CvCriticalLiquid_Error
'inletValveDia and inletPipeDia must both be greater than zero
If (dbl_ValveInletDia <= 0 Or dbl_InletPipeDia <= 0) Then GoTo CvCriticalLiquid_Error

'******************
'Declarations
'******************
Dim dbl_Ff As Double

'******************
'Begin Calculations
'******************
dbl_Ff = 0.96 - 0.28 * Sqr(dbl_Pv / dbl_Pc)
    If dbl_InletPipeDia = dbl_ValveInletDia Then
        'If flowrate is in volume units
        If dbl_Q > 0 Then dbl_Cv = (dbl_Q / (0.865 * dbl_FL)) * Sqr(dbl_Gf / (dbl_P1 - dbl_Ff * dbl_Pv))
        'If flowrate is in weight units
        If dbl_W > 0 Then dbl_Cv = (dbl_W / (27.3 * dbl_FL)) / Sqr((dbl_P1 - dbl_Ff * dbl_Pv) * dbl_Gf * 1000)
    ElseIf dbl_InletPipeDia <> dbl_ValveInletDia Then
        'Flowrate is in volume units
        If dbl_Q > 0 Then dbl_Cv = dbl_Q * (1 / 0.865) * (1 / dbl_Flp) * Sqr(dbl_Gf / (dbl_P1 - dbl_Ff * dbl_Pv))
        'Flowrate is in weight units
        If dbl_W > 0 Then dbl_Cv = dbl_W * (1 / (27.3 * dbl_Flp)) * Sqr(1 / (dbl_P1 - dbl_Ff * dbl_Pv) * dbl_Gf * 1000)
    End If
        
Exit Sub
CvCriticalLiquid_Error:
'Placeholder

End Sub
Sub CvSubCriticalLiq(dbl_P1 As Double, dbl_P2 As Double, dbl_Pv As Double, _
dbl_Pc As Double, dbl_Gf As Double, dbl_Q As Double, dbl_W As Double, dbl_Fp As Double, _
dbl_InletPipeDia As Double, dbl_OutletPipeDia As Double, dbl_ValveInletDia As Double, _
dbl_ValveOutletDia As Double, dbl_Cv As Double)

'------------------------------------------------------------------------------------------------
'This subroutine calculates the Cv for a flashing liquid with critical flow conditions.
'
'Reference:  ISA S75.01
'        IEC 534 - 2 - 1
'------------------------------------------------------------------------------------------------
'Input Variables                    output                description                             units
'
'dbl_P1                                         inlet pressure                          bar a
'dbl_P2                                         outlet pressure                         bar a
'dbl_Pv                                         vapor pressure                          bar a
'dbl_Pc                                         critical pressure                       bar a
'r1/r0 »dbl_Gf                                  specific gravity or relative density
'dbl_W                                          flowrate (weight basis)                 kg/h
'dbl_Q                                          flowrate (volume basis)                 m3/h
'dbl_Fp                                         pipe geometry factor
'dbl_InletPipeDia                               nominal inlet pipe size                 inches/mm
'dbl_OutletPipeDia                              nominal outlet pipe size                inches/mm
'dbl_ValveInletDia                              nominal valve inlet size                inches/mm
'dbl_valveOutletDia                             nominal valve inlet size                inches/mm
'                         dbl_Cv                valve flow coefficient
'
'Author:  KPMG                  Date:  6/12/98
'------------------------------------------------------------------------------------------------

On Error GoTo CvSubCriticalLiq_Error

'*****************
'Check input values
'*****************

'All pipe and valve diameters must be greater than 0
If (dbl_InletPipeDia <= 0 Or dbl_OutletPipeDia <= 0 Or dbl_ValveInletDia <= 0 Or _
dbl_ValveOutletDia <= 0) Then GoTo CvSubCriticalLiq_Error
'dbl_P1 must be greater than dbl_P2
If dbl_P1 <= dbl_P2 Then GoTo CvSubCriticalLiq_Error
'dbl_Gf must be greater than 0
If dbl_Gf <= 0 Then GoTo CvSubCriticalLiq_Error
'One and only one flowrate must be specified
If (dbl_Q <= 0 And dbl_W <= 0) Or (dbl_Q > 0 And dbl_W > 0) Then GoTo CvSubCriticalLiq_Error

'************************
'Begin calculations
'************************
If (dbl_InletPipeDia = dbl_ValveInletDia And dbl_OutletPipeDia = dbl_ValveOutletDia) Then
    'If flowrate is in volume units
    If dbl_Q > 0 Then dbl_Cv = Sqr(dbl_Gf / (dbl_P1 - dbl_P2)) * dbl_Q / 0.865
    'If flowrate is in weight units
    If dbl_W > 0 Then dbl_Cv = dbl_W / (27.3 * Sqr(((dbl_P1 - dbl_P2) * dbl_Gf * 1000)))
Else
    'If flowrate is in volume units
    If dbl_Q > 0 Then dbl_Cv = Sqr(dbl_Gf / (dbl_P1 - dbl_P2)) * dbl_Q / (0.865 * dbl_Fp)
    'If flowrate is in weight units
    If dbl_W > 0 Then dbl_Cv = dbl_W / (27.3 * dbl_Fp * Sqr(((dbl_P1 - dbl_P2) * dbl_Gf * 1000)))
End If

Exit Sub

CvSubCriticalLiq_Error:
'Placeholder

End Sub
Sub LaminarCv(str_ValveLine As String, lng_FlowDirectionID As Long, _
    lng_TrimTypeID As Long, bln_ReducedArea As Boolean, dbl_Q As Double, _
    dbl_Viscosity As Double, dbl_Gf As Double, dbl_FL As Double, _
    dbl_ValveInletDia As Double, dbl_P1 As Double, dbl_P2 As Double, _
    dbl_RatedCv As Double, dbl_Cv As Double, int_FlowTypeFlag As Integer)

'--------------------------------------------------------------------------------------------------------------------
'
'Reference :  ISA S75.01, Appendix E
'
'This subroutine calculates the Cv for laminar or transitional flow
'
'--------------------------------------------------------------------------------------------------------------------
'Input Variables                 output               description                        units
'dbl_Q                                      Flowrate(volume)                   gpm
'dbl_Viscosity                              liquid viscosity(absolute)         centipoises
'dbl_Gf                                     specific gravity                   dimensionless
'dbl_FL                                     liquid pressure recovery factor    dimensionless
'dbl_ValveInletDia                          nominal valve size                 inches
'dbl_P1                                     inlet pressure                     psia
'dbl_P2                                     outlet pressure                    psia
'dbl_Fd                                     valve style modifier               dimensionless
'                      dbl_Cv               valve flow coefficient
'                      int_FlowTypeFlag     Indicates flow state and equations used to calculate Cv
'
'Author:  KPMG                  Date:  6/12/98
'--------------------------------------------------------------------------------------------------------------------

On Error GoTo LaminarCv_Error

'*********************
'Check input values
'*********************

'dbl_P1 must be greater than dbl_P2
If dbl_P1 <= dbl_P2 Then GoTo LaminarCv_Error
'Gf must be grater than 0
If dbl_Gf <= 0 Then GoTo LaminarCv_Error
'Viscosity must be greater than 0
If dbl_Viscosity <= 0 Then GoTo LaminarCv_Error
'FL must be greater than zero
If dbl_FL <= 0 Then GoTo LaminarCv_Error
'ValveInletDia must be greater than 0
If dbl_ValveInletDia <= 0 Then GoTo LaminarCv_Error
'Q must be greater than 0
If dbl_Q <= 0 Then GoTo LaminarCv_Error

'*********************
'Declarations
'*********************

Dim dbl_Cvt As Double
Dim dbl_ReynoldsNumber As Double
Dim dbl_Fr As Double
Dim dbl_Fd As Double
Dim dbl_LocalRatedCV As Double

'*******************
'Begin Calculations
'*******************


dbl_Cvt = dbl_Q / (0.865 * Sqr((dbl_P1 - dbl_P2) / dbl_Gf))
dbl_LocalRatedCV = dbl_RatedCv
If dbl_LocalRatedCV = 0 Then dbl_LocalRatedCV = 2 * dbl_Cvt

'temporary defaults
If (str_ValveLine = "") Then
    dbl_Fd = 1
    dbl_Cv = dbl_Cvt
Else
    GetStdValveDataForNoise str_ValveLine, lng_FlowDirectionID, _
        lng_TrimTypeID, dbl_Cvt, dbl_LocalRatedCV, dbl_Fd
End If

dbl_ReynoldsNumber = (76000 * dbl_Fd * dbl_Q * dbl_Gf / (dbl_Viscosity * Sqr(dbl_FL * dbl_Cvt))) _
* (((dbl_FL ^ 2) * (dbl_Cvt ^ 2) / (0.00214 * dbl_ValveInletDia ^ 4)) + 1) ^ 0.25

'*******************
'Check Calculations
'*******************

If (dbl_Cvt <= 0 Or dbl_ReynoldsNumber <= 0) Then GoTo LaminarCv_Error

'*******************
'Calculate Cv
'*******************

If dbl_ReynoldsNumber >= 40000 Then
    'Flow is turbulent
    int_FlowTypeFlag = 0
    Exit Sub
ElseIf dbl_ReynoldsNumber < 56 Then
    'Flow is laminar
    int_FlowTypeFlag = 5
    dbl_Fr = 0.019 * dbl_ReynoldsNumber ^ 0.67
    dbl_Cv = dbl_Cvt / dbl_Fr
    Exit Sub
ElseIf dbl_ReynoldsNumber >= 56 And dbl_ReynoldsNumber < 40000 Then
    'Flow is laminar
    Select Case dbl_ReynoldsNumber
        Case 56 To 66
            dbl_Fr = 0.284 + (dbl_ReynoldsNumber - 56) * (0.32 - 0.284) / (66 - 56)
        Case 66 To 79
            dbl_Fr = 0.32 + (dbl_ReynoldsNumber - 66) * (0.36 - 0.32) / (79 - 66)
        Case 79 To 94
            dbl_Fr = 0.36 + (dbl_ReynoldsNumber - 79) * (0.4 - 0.36) / (94 - 79)
        Case 94 To 110
            dbl_Fr = 0.4 + (dbl_ReynoldsNumber - 94) * (0.44 - 0.4) / (110 - 94)
        Case 110 To 130
            dbl_Fr = 0.44 + (dbl_ReynoldsNumber - 110) * (0.48 - 0.44) / (130 - 110)
        Case 130 To 154
            dbl_Fr = 0.48 + (dbl_ReynoldsNumber - 130) * (0.52 - 0.48) / (154 - 130)
        Case 154 To 188
            dbl_Fr = 0.52 + (dbl_ReynoldsNumber - 154) * (0.56 - 0.52) / (188 - 154)
        Case 188 To 230
            dbl_Fr = 0.56 + (dbl_ReynoldsNumber - 188) * (0.6 - 0.56) / (230 - 188)
        Case 230 To 278
            dbl_Fr = 0.6 + (dbl_ReynoldsNumber - 230) * (0.64 - 0.6) / (278 - 230)
        Case 278 To 340
            dbl_Fr = 0.64 + (dbl_ReynoldsNumber - 278) * (0.68 - 0.64) / (340 - 278)
        Case 340 To 471
            dbl_Fr = 0.68 + (dbl_ReynoldsNumber - 340) * (0.72 - 0.68) / (471 - 340)
        Case 471 To 620
            dbl_Fr = 0.72 + (dbl_ReynoldsNumber - 471) * (0.76 - 0.72) / (620 - 471)
        Case 620 To 980
            dbl_Fr = 0.76 + (dbl_ReynoldsNumber - 620) * (0.8 - 0.76) / (980 - 620)
        Case 980 To 1560
            dbl_Fr = 0.8 + (dbl_ReynoldsNumber - 980) * (0.84 - 0.8) / (1560 - 980)
        Case 1560 To 2470
            dbl_Fr = 0.84 + (dbl_ReynoldsNumber - 1560) * (0.88 - 0.84) / (2470 - 1560)
        Case 2470 To 4600
            dbl_Fr = 0.88 + (dbl_ReynoldsNumber - 2470) * (0.92 - 0.84) / (4600 - 2470)
        Case 4600 To 10200
            dbl_Fr = 0.92 + (dbl_ReynoldsNumber - 4600) * (0.96 - 0.92) / (10200 - 4600)
        Case 10200 To 40000
            dbl_Fr = 0.96 + (dbl_ReynoldsNumber - 10200) * (1 - 0.96) / (40000 - 10200)
    End Select
        int_FlowTypeFlag = 6
        dbl_Cv = dbl_Cvt / dbl_Fr
End If

Exit Sub

LaminarCv_Error:
'Placeholder

End Sub
Sub LiquidVelocity(dbl_Q As Double, dbl_W As Double, dbl_Gf As Double, _
dbl_InletPipeOD As Double, dbl_InletPipeWallThickness As Double, _
dbl_MaxVelocity As Double, dbl_InletPipeVelocity As Double, _
dbl_MaxVelocityDiameter As Double)

'------------------------------------------------------------------------------------------------
'This subroutine calculates the inlet velocity in the inlet pipe.
'It is an average velocity based on the pipe data.  TheMaxVelocityDiameter
'is calculated base on the flow rate and the value of MaxVelocity which is the
'limit given for pipe velocity.
'------------------------------------------------------------------------------------------------
'
'Input                  Output                  Description                                     Units
'dbl_InletPipeOD                                The outside diameter of the inlet pipe          mm
'dbl_InletPipeWallThickness                     The wall thickness of the inlet pipe            mm
'dbl_Q                                          The volume flow rate                            m3/h
'dbl_W                                          The weight flow rate                            kg/h
'r1/r0 » dbl_Gf                                 specific gravity or relative density            -
'dbl_MaxVelocity                                The maximum allowable velociy                   m/s
'                       dbl_MaxVelocityDiameter The minimum internal pipe diameter to meet      mm
'                                               the maximum velocity limit.
'                       dbl_InletPipeVelocity   The average velocity based on the inlet pipe    m/s
'                                               inside diameter
'
'Author:  KPMG                  Date:  6/12/98
'---------------------------------------------------------------------------------------------------

On Error GoTo LiquidVelocity_Error

'******************
'Check input values
'******************

'InletPipeOD must be greater than 2*InletPipeWallThickness
If dbl_InletPipeOD <= 2 * dbl_InletPipeWallThickness Then GoTo LiquidVelocity_Error
'MaxVelocity must be greater than 0
If dbl_MaxVelocity <= 0 Then GoTo LiquidVelocity_Error
'Gf must be greater than 0
If dbl_Gf <= 0 Then GoTo LiquidVelocity_Error

'*******************
'Begin calculations
'*******************

If dbl_W > 0 Then
    dbl_InletPipeVelocity = 353.7 * dbl_W / (1000 * dbl_Gf * (dbl_InletPipeOD - 2 * dbl_InletPipeWallThickness) ^ 2)
    dbl_MaxVelocityDiameter = 18.8 * Sqr(dbl_W / (1000 * dbl_Gf * dbl_MaxVelocity))
ElseIf dbl_Q > 0 Then
    dbl_InletPipeVelocity = 353.7 * dbl_Q / ((dbl_InletPipeOD - 2 * dbl_InletPipeWallThickness) ^ 2)
    dbl_MaxVelocityDiameter = 18.8 * Sqr(dbl_Q / dbl_MaxVelocity)
End If

Exit Sub

LiquidVelocity_Error:
'Placeholder

End Sub
Sub Lwae(dbl_Lwe1 As Double, dbl_Lwe2 As Double, dbl_Lwe3 As Double, dbl_Lwe4 As Double, dbl_Lwe5 As Double, dbl_Lwae As Double)

'**************************************************************************************
'  Purpose:    To calculate the force output of an air to open linear motion actuator
'
'  Inputs:     Lwe1
'              Lwe2
'              Lwe3
'              Lwe4
'              Lwe5
'
'  Outputs:    Lwae
'  Author:     Julie Fritz
'
'
'  Revisons:
'  Symbol  Date    SE              Purpose
'  ------- ----    --------------- ----------------------------------------------------
'          6/11/98  KPMG            Corrected formulae
'****************************************************************************************

On Error GoTo Lwae_Error

'****************************************************************************************
'  Declarations
'****************************************************************************************

Dim dbl_Lwa1 As Double          'L_wa(f) at each frequency of the octave band with A-weight factors applied
Dim dbl_Lwa2 As Double
Dim dbl_Lwa3 As Double
Dim dbl_Lwa4 As Double
Dim dbl_Lwa5 As Double

'***********************
'Begin calculations
'***********************

'L_wa(f) = L_we(f) + A_weight(f)
dbl_Lwa1 = dbl_Lwe1 - 3.2
dbl_Lwa2 = dbl_Lwe2
dbl_Lwa3 = dbl_Lwe3 + 1.2
dbl_Lwa4 = dbl_Lwe4 + 1
dbl_Lwa5 = dbl_Lwe5 - 1.1
dbl_Lwae = 10 * Log((10 ^ (0.1 * dbl_Lwa1)) + (10 ^ (0.1 * dbl_Lwa2)) + (10 ^ (0.1 * dbl_Lwa3)) + (10 ^ (0.1 * dbl_Lwa4)) + (10 ^ (0.1 * dbl_Lwa5))) / Log(10)

Lwae_Exit:
Exit Sub

Lwae_Error:
'Placeholder

End Sub
Sub Lpae(dbl_Lwae As Double, dbl_Ip As Double, dbl_Di As Double, _
dbl_Do As Double, dbl_Io As Double, dbl_LpAe As Double)

'***************************************************************************************
'*
'*
'*  Function/Sub:   Lpae
'*
'*
'*  Purpose:
'*
'*  Inputs:     Lwae
'*              Ip
'*              Di
'*              Dp
'*              Io
'*
'*  Outputs:    LPae
'*
'*
'*
'*  Author:     Julie Fritz
'*
'*
'*  Revisons:
'*  Symbol  Date    SE              Purpose
'*  ------- ----    --------------- ----------------------------------------------------
'*          6/11/98 KPMG            Corrected formulae
'*****************************************************************************************

'*****************************************************************************************
'*  Start Code
'*****************************************************************************************

On Error GoTo Lpae_Error:
dbl_LpAe = dbl_Lwae - 10 * Log((PI * dbl_Ip / dbl_Io) * (1 + dbl_Di / dbl_Do)) / Log(10)

Exit Sub

Lpae_Error:
    'Placeholder

End Sub


Sub Lwe_f(dbl_Lwi1 As Double, dbl_Lwi2 As Double, _
dbl_Lwi3 As Double, dbl_Lwi4 As Double, dbl_Lwi5 As Double, _
dbl_Tl1 As Double, dbl_Tl2 As Double, dbl_Tl3 As Double, dbl_Tl4 As Double, _
dbl_Tl5 As Double, dbl_Ip As Double, dbl_Do As Double, dbl_Lwe1 As Double, _
dbl_Lwe2 As Double, dbl_Lwe3 As Double, dbl_Lwe4 As Double, dbl_Lwe5 As Double)


'***************************************************************************************
'*
'*
'*  Function/Sub:   Lwe_f
'*
'*
'*  Purpose:    To calculate the external sound power level for each frequency, 500, 1000, 2000, 4000, adn 8000.
'*
'*  Inputs:     dbl_Lwi1-  Internal sound power level (f=500)       dB
'*              dbl_Lwi2-  Internal sound power level (f=1000)      dB
'*              dbl_Lwi3-  Internal sound power level (f=2000)      dB
'*              dbl_Lwi4-  Internal sound power level (f=5000)      dB
'*              dbl_Lwi5-  Internal sound power level (f=8000)      dB
'*              dbl_TL1-  Transmission Loss (f=500)                 dB
'*              dbl_TL2-  Transmission Loss (f=1000)                dB
'*              dbl_TL3-  Transmission Loss (f=2000)                dB
'*              dbl_TL4-  Transmission Loss (f=4000)                dB
'*              dbl_TL5-  Transmission Loss (f=8000)                dB
'*              dbl_Ip-  Refernece pipe length                      m
'*              dbl_Do-  Outside diameter of downstream pipe        m
'*
'*  Outputs:    dbl_Lwe1-  External sound power level (f=500)        dB
'*              dbl_Lwe2-  External sound power level (f=1000)       dB
'*              dbl_Lwe3-  External sound power level (f=2000)      dB
'*              dbl_Lwe4-  External sound power level (f=5000)      dB
'*              dbl_Lwe5-  External sound power level (f=8000)      dB
'*
'*
'*
'*
'*  Author:     Julie Fritz
'*
'*
'*  Revisons:
'*  Symbol  Date    SE              Purpose
'*  ------- ----    --------------- ----------------------------------------------------
'*          6/11/98 KPMG            Corrected formulae
'*****************************************************************************************

'*****************************************************************************************
'*  Start Code
'*****************************************************************************************

On Error GoTo Lwe_f_Error

'Lwe(f) = dbl_Lwi(f) - 17.37* (dbl_Ip/ (2*dbl_Do)) * (10^(-.1 * dbl_Tl(f))) - dbl_Tl(f) +
'(10 * log((4* dbl_Ip)/dbl_Do))

dbl_Lwe1 = dbl_Lwi1 - ((17.37 * dbl_Ip * 10 ^ (-0.1 * dbl_Tl1)) / (2 * dbl_Do)) - dbl_Tl1 + 10 * Log(4 * dbl_Ip / dbl_Do) / Log(10)
dbl_Lwe2 = dbl_Lwi2 - ((17.37 * dbl_Ip * 10 ^ (-0.1 * dbl_Tl2)) / (2 * dbl_Do)) - dbl_Tl2 + 10 * Log(4 * dbl_Ip / dbl_Do) / Log(10)
dbl_Lwe3 = dbl_Lwi3 - ((17.37 * dbl_Ip * 10 ^ (-0.1 * dbl_Tl3)) / (2 * dbl_Do)) - dbl_Tl3 + 10 * Log(4 * dbl_Ip / dbl_Do) / Log(10)
dbl_Lwe4 = dbl_Lwi4 - ((17.37 * dbl_Ip * 10 ^ (-0.1 * dbl_Tl4)) / (2 * dbl_Do)) - dbl_Tl4 + 10 * Log(4 * dbl_Ip / dbl_Do) / Log(10)
dbl_Lwe5 = dbl_Lwi5 - ((17.37 * dbl_Ip * 10 ^ (-0.1 * dbl_Tl5)) / (2 * dbl_Do)) - dbl_Tl5 + 10 * Log(4 * dbl_Ip / dbl_Do) / Log(10)

Exit Sub

Lwe_f_Error:
'Placeholder

End Sub
Sub Lwi_critical(dbl_Xfz As Double, dbl_Nf As Double, _
dbl_DeltaLfc As Double, dbl_Fb As Double, dbl_M As Double, _
dbl_Pfluid As Double, dbl_Xf As Double, dbl_P1Pa As Double, _
dbl_P2Pa As Double, dbl_Pcpa As Double, dbl_Pvpa As Double, _
dbl_FL As Double, dbl_Lwi As Double)

'**************************************************************************************
'*  Function/Sub:  Lwi_critical
'*
'*
'*  Purpose:    Reference IEC 534-8-4 equation 5 modified to include Fb,
'*              Fx adn Delta_Lfc.  To calculate the value of Lwi when
'*              flow is in the incipient cavitation or cavitation range.
'*
'*  Inputs:     dbl_Xfz- Characteristic pressure ratio
'*              dbl_Delta_LFc-  Valve specific correction factor coefficient
'*              dbl_Fb- Cavitating flow noise factor
'*              dbl_m- Mass flow rate
'*              dbl_Pfluid- Fluid density at inlet conditions
'*              dbl_Nf-  Acoustical efficiency factor
'*              dbl_Xf- Differential pressure ratio
'*              dbl_P1pa- Inlet pressure
'*              dbl_P2pa- Outlet pressure
'*              dbl_Pcpa- Critical pressure
'*              dbl_Pvpa- Vapor pressure
'*              dbl_Fl- Liquid pressure recovery factor
'*
'*  Outputs:    dbl_Lwi-  Internal sound power level
'*
'*
'*  Author:     Julie Fritz
'*
'*
'*  Revisons:
'*  Symbol  Date    SE              Purpose
'*  ------- ----    --------------- ----------------------------------------------------
'*          6/11/98 KPMG            Corrected formulae
'*****************************************************************************************

On Error GoTo Lwi_critical_Error

'********************
'Check input values
'********************

'********************
'Declarations
'********************
Dim dbl_Ff As Double
Dim dbl_Delta_p
Dim dbl_Delta_LF As Double
Dim dbl_Fx As Double

'Delta_LFc is a constant that is found in a database(changes with valve type)

'The term 10 log Delta P must be limited as follows

dbl_Ff = 0.96 - 0.28 * (Sqr(dbl_Pvpa / dbl_Pcpa))

dbl_Delta_p = dbl_P1Pa - dbl_P2Pa

If dbl_Delta_p <= (dbl_FL ^ 2) * (dbl_P1Pa - dbl_Ff * dbl_Pvpa) Then
    dbl_Delta_p = dbl_P1Pa - dbl_P2Pa
Else
    If dbl_Delta_p > (dbl_FL ^ 2) * (dbl_P1Pa - dbl_Ff * dbl_Pvpa) Then
        dbl_Delta_p = (dbl_FL ^ 2) * (dbl_P1Pa - dbl_Ff * dbl_Pvpa)
    End If
End If

dbl_Fx = (dbl_Xfz ^ 0.0625) * ((1 - dbl_Xf) ^ 0.8) * Log((1 - dbl_Xfz) / (1 - dbl_Xf)) / (Log(10) * (dbl_Xf ^ dbl_Xfz))
dbl_Delta_LF = dbl_DeltaLfc * dbl_Fx

dbl_Lwi = 120 + 10 * ((Log(dbl_Nf) + Log(dbl_M) + Log(dbl_Delta_p) - Log(dbl_Pfluid)) / Log(10)) + dbl_Fb * dbl_Fx * (dbl_DeltaLfc + 180)
Exit Sub

Lwi_critical_Error:
'Placeholder
 
End Sub

Sub Lwi_f(dbl_Lwi As Double, dbl_Lwi1 As Double, dbl_Lwi2 As Double, _
dbl_Lwi3 As Double, dbl_Lwi4 As Double, dbl_Lwi5 As Double)

'***************************************************************************************
'*
'*
'*  Function/Sub:   Lwi_f
'*
'*
'*  Purpose:    To calculate the external sound power level for each frequency, 500, 1000, 2000, 4000, adn 8000.
'*
'*  Inputs:     dbl_Lwi-  Internal sound power level                dB
'*
'*
'*  Outputs:    dbl_Lwi1-  Internal sound power level (f=500)       dB
'*              dbl_Lwi2-  Internal sound power level (f=1000)      dB
'*              dbl_Lwi3-  Internal sound power level (f=2000)      dB
'*              dbl_Lwi4-  Internal sound power level (f=5000)      dB
'*              dbl_Lwi5-  Internal sound power level (f=8000)      dB
'*
'*
'*
'*
'*  Author:     Julie Fritz
'*
'*
'*  Revisons:
'*  Symbol  Date    SE              Purpose
'*  ------- ----    --------------- ----------------------------------------------------
'*          6/11/98 KPMG            Corrected formulae
'*****************************************************************************************


'*****************************************************************************************
'*  Decalrations
'*****************************************************************************************



'*****************************************************************************************
'*  Start Code
'*****************************************************************************************
On Error GoTo Lwi_f_Error

dbl_Lwi1 = dbl_Lwi - 2.9
dbl_Lwi2 = dbl_Lwi - 10 * (Log(1000 / 500) / Log(10)) - 2.9
dbl_Lwi3 = dbl_Lwi - 10 * (Log(2000 / 500) / Log(10)) - 2.9
dbl_Lwi4 = dbl_Lwi - 10 * (Log(4000 / 500) / Log(10)) - 2.9
dbl_Lwi5 = dbl_Lwi - 10 * (Log(8000 / 500) / Log(10)) - 2.9

Lwi_f_Exit:
Exit Sub

Lwi_f_Error:
'Placeholder

End Sub

Sub Lwi_subcritical(dbl_P1Pa As Double, dbl_P2Pa As Double, dbl_Nf, dbl_Pfluid As Double, dbl_Wo As Double, dbl_M As Double, dbl_Lwi As Double)

'**************************************************************************************
'*
'*
'*  Function/Sub:  Lwi_subcritical
'*
'*
'*  Purpose:    Reference IEC 534-8-4.  To calculate the value of Lwi when flow is subcritical ( no incipient cavitation)
'*
'*  Inputs:     dbl_m- Mass flow rate
'*              dbl_Pfluid- Fluid density at inlet conditions
'*              dbl_Nf-  Acoustical efficiency factor
'*              dbl_P1pa- Inlet pressure
'*              dbl_P2pa- Outlet pressure
'*              dbl_Wo-  Ref sound power
'*
'*  Outputs:    dbl_Lwi-  Internal sound power level
'*
'*
'*  Author:     Julie Fritz
'*
'*
'*  Revisons:
'*  Symbol  Date    SE              Purpose
'*  ------- ----    --------------- ----------------------------------------------------
'*          6/11/98 KPMG            Corrected formulae
'*****************************************************************************************


'*****************************************************************************************
'*  Declarations
'*****************************************************************************************

Dim dbl_Delta_p

'*****************************************************************************************
'*  Start Code
'*****************************************************************************************
On Error GoTo Lwi_subcritical_Error

dbl_Delta_p = dbl_P1Pa - dbl_P2Pa
dbl_Lwi = 10 * (Log((dbl_Nf * dbl_M * dbl_Delta_p) / (dbl_Pfluid * dbl_Wo))) * 1 / Log(10)

'*****************************************************************************************
'*  Common Exit Point
'*****************************************************************************************
Lwi_subcritical_Exit:
Exit Sub

'*****************************************************************************************
'*  Standard Error Handling
'*****************************************************************************************
Lwi_subcritical_Error:
'Placeholder
End Sub
Sub TL_f(dbl_Cpipe As Double, dbl_Ppipe As Double, dbl_T As Double, dbl_Cfluid As Double, dbl_Pfluid As Double, dbl_Do As Double, dbl_Fr As Double, dbl_Tl1 As Double, dbl_Tl2 As Double, dbl_Tl3 As Double, dbl_Tl4 As Double, dbl_Tl5 As Double)

'***************************************************************************************
'*
'*
'*  Function/Sub:   TL_f
'*
'*
'*  Purpose:    TO calculate the transmission loss for each frequency. f= 500, 100, 2000, 4000 and 8000
'*
'*  Inputs:     dbl_Cpipe-  Speed of sound of longitudinal waves in pipe wall
'*              dbl_Ppipe-  Density of pipe material                kg/m^3
'*              dbl_t-  Pipe wall thickness                         m
'*              dbl_Cfluid-  Speed of sound in fluid                m
'*              dbl_Pfluid-  Density of fluid at inlet condition    kg/m^3
'*              dbl_Do-  Outside diameter of downstream pipe        m
'*              dbl_Fr- Ring frequency                              Hz
'*
'*
'*  Outputs:    dbl_TL1-  Transmission Loss (f=500)                 dB
'*              dbl_TL2-  Transmission Loss (f=1000)                dB
'*              dbl_TL3-  Transmission Loss (f=2000)                dB
'*              dbl_TL4-  Transmission Loss (f=4000)                dB
'*              dbl_TL5-  Transmission Loss (f=8000)                dB
'*
'*
'*
'*
'*  Author:     Julie Fritz
'*
'*
'*  Revisons:
'*  Symbol  Date    SE              Purpose
'*  ------- ----    --------------- ----------------------------------------------------
'*          6/11/98 KPMG            Corrected formulae
'*****************************************************************************************


'*****************************************************************************************
'*  Decalrations
'*****************************************************************************************



'*****************************************************************************************
'*  Start Code
'*****************************************************************************************
    On Error GoTo TL_f_Error

'TL(f) = 10 + 10*Log(((dbl_Cpipe * dbl_Ppipe * dbl_t)/(dbl_Cfluid * dbl_Pfluid * dbl_Do)))
'+ (10 * log((dbl_Fr/dbl_f + ((dbl_f/dbl_Fr)^1.5))^2
'Evaluate the equation for each frequency of the octave band

dbl_Tl1 = 10 + 10 * ((Log(dbl_Cpipe * dbl_Ppipe * dbl_T / (dbl_Cfluid * dbl_Pfluid * dbl_Do)) / Log(10)) + Log(((dbl_Fr / 500) + (500 / dbl_Fr) ^ 1.5) ^ 2) / Log(10))
dbl_Tl2 = 10 + 10 * ((Log(dbl_Cpipe * dbl_Ppipe * dbl_T / (dbl_Cfluid * dbl_Pfluid * dbl_Do)) / Log(10)) + Log(((dbl_Fr / 1000) + (1000 / dbl_Fr) ^ 1.5) ^ 2) / Log(10))
dbl_Tl3 = 10 + 10 * ((Log(dbl_Cpipe * dbl_Ppipe * dbl_T / (dbl_Cfluid * dbl_Pfluid * dbl_Do)) / Log(10)) + Log(((dbl_Fr / 2000) + (2000 / dbl_Fr) ^ 1.5) ^ 2) / Log(10))
dbl_Tl4 = 10 + 10 * ((Log(dbl_Cpipe * dbl_Ppipe * dbl_T / (dbl_Cfluid * dbl_Pfluid * dbl_Do)) / Log(10)) + Log(((dbl_Fr / 4000) + (4000 / dbl_Fr) ^ 1.5) ^ 2) / Log(10))
dbl_Tl5 = 10 + 10 * ((Log(dbl_Cpipe * dbl_Ppipe * dbl_T / (dbl_Cfluid * dbl_Pfluid * dbl_Do)) / Log(10)) + Log(((dbl_Fr / 8000) + (8000 / dbl_Fr) ^ 1.5) ^ 2) / Log(10))

'*****************************************************************************************
'*  Common Exit Point
'*****************************************************************************************
TL_f_Exit:
Exit Sub

'*****************************************************************************************
'*  Standard Error Handling
'*****************************************************************************************
TL_f_Error:
'Placeholder
End Sub

Sub Incipient(str_ValveLine As String, lng_FlowDirectionID As Long, _
lng_TrimTypeID As Long, bln_ReducedArea As Boolean, dbl_Cv As Double, _
dbl_RatedCv As Double, dbl_P1 As Double, dbl_P2 As Double, _
dbl_Pv As Double, dbl_Pc As Double, dbl_Xfz As Double, _
dbl_FL As Double, int_IncipientFlag As Integer, _
int_FlowStateFlag As Integer)
'------------------------------------------------------------------------------------------------
'Reference:  Masoneilan Handbook for Control Valve Sizing,  1987,  Seventh Edition
'        IEC 534-2-1 Sizing equations for fluid flow under installed conditions
'        IEC 534-8-4 Prediction of noise generated by hydrodynamic flow
'
'This routine checks for incipient cavitation.
'------------------------------------------------------------------------------------------------
'Input Variables   output         description                   units
'P1                     inlet pressure                bar a
'P2                     outlet pressure               bar a
'Pv                     vapor pressure                bar a
'Pc                     critical pressure             bar a
'FL                     liquid pressure recovery factor -
'XFZ
'        FlowStateFlag
'
'Author:  KPMG                  Date:  6/12/98
'------------------------------------------------------------------------------------------------
On Error GoTo Incipient_Error
'************************
'Check input data
'************************
'dbl_P1 must be greater than dbl_Pv
If (dbl_P1 <= dbl_Pv) Then GoTo Incipient_Error
'dbl_P1 must be greater than dbl_P2
If (dbl_P1 <= dbl_P2) Then GoTo Incipient_Error
'dbl_Pc must be greater than 0
If dbl_Pc <= 0 Then GoTo Incipient_Error

'************************
'Declarations
'************************
Dim dbl_Xf As Double
Dim dbl_DeltaPcritical As Double
Dim dbl_DLfc As Double

'***********************
'Begin calculations
'***********************

GetXFZ str_ValveLine, lng_FlowDirectionID, lng_TrimTypeID, _
    bln_ReducedArea, dbl_RatedCv, dbl_Cv, dbl_Xfz, dbl_DLfc

dbl_Xf = (dbl_P1 - dbl_P2) / (dbl_P1 - dbl_Pv)
dbl_DeltaPcritical = (dbl_P1 - dbl_Pv * (0.96 - 0.28 * Sqr(dbl_Pv / dbl_Pc))) * dbl_FL ^ 2

If (dbl_Xf >= dbl_Xfz And (dbl_P1 - dbl_P2) < dbl_DeltaPcritical) Then
    int_FlowStateFlag = 4
ElseIf (dbl_Xf < dbl_Xfz) Then
    int_FlowStateFlag = 1
ElseIf ((dbl_P1 - dbl_P2) >= dbl_DeltaPcritical) Then
    int_FlowStateFlag = 2
End If

Exit Sub

Incipient_Error:
'Placeholder
End Sub

Sub LiquidNoise(str_ValveLine As String, lng_FlowDirectionID As Long, _
    lng_TrimTypeID As Long, bln_ReducedArea As Boolean, _
    dbl_RatedCv As Double, dbl_Cv As Double, dbl_Fb As Double, _
    dbl_P1 As Double, dbl_P2 As Double, dbl_Pv As Double, _
    dbl_Pc As Double, dbl_Q As Double, dbl_W As Double, _
    dbl_Gf As Double, dbl_OutletPipeOD As Double, dbl_Tk As Double, _
    dbl_FL As Double, dbl_LpAe As Double)


'***************************************************************************************
'*  Reference:  IEC 534-8-4, First Edition 1994-05
'*              Prediction of noise generated by hydrodynamic flow.
'*
'*              Implementation of IEC 534-8-4, Hydrodynamic Noise Prediction
'*              (H. Boger Dec. 7, 1994)
'*
'*              Improvement to IEC method, Fb factor added
'*              (reference: memo from H. Boger
'*
'*  Function/Sub:   Liquid Noise
'*
'*
'*  Purpose:    To calculate athe external A-weighted sound pressure level, Lpae
'*
'*  Inputs:     P1
'*              P2
'*              Pv
'*              Pc
'*              Q
'*              W
'*              Gf
'*              OutletPipeOD
'*              tk
'*              Xfz
'*              Delta_Lfc
'*              Fl
'*
'*  Outputs:    Lpae
'*
'*
'*
'*  Author:     Julie Fritz
'*
'*
'*  Revisons:
'*  Symbol  Date    SE              Purpose
'*  ------- ----    --------------- ----------------------------------------------------
'*                                  1.  Added call to GetXFZ
'*                                  2.  Changed formulae as per liqnoise.doc
'*****************************************************************************************
'*****************************************************************************************
'*  Declarations
'*****************************************************************************************
 
 
'set constants
Const dbl_Wo = 10 ^ -12
Const dbl_Nf = 0.0000001
Const dbl_Cpipe = 5100
Const dbl_Cfluid = 1440
Const dbl_Ppipe = 7800
Const dbl_Lo = 1
Const dbl_Lp = 3
Dim dbl_DLfc As Double
Dim dbl_Pfluid As Double
Dim dbl_P1Pa As Double
Dim dbl_P2Pa As Double
Dim dbl_Pvpa As Double
Dim dbl_Pcpa As Double
Dim dbl_M As Double
Dim dbl_Do As Double
Dim dbl_Di As Double
Dim dbl_T As Double
Dim dbl_Fr As Double
Dim dbl_Lwi As Double
Dim dbl_Lwif As Double
Dim dbl_Lwef As Double
Dim dbl_Tlf As Double
Dim dbl_Lwae As Double
Dim dbl_Lwae2 As Double
Dim intCheck As Integer
Dim dbl_Xf As Double
Dim dbl_Lwi1 As Double
Dim dbl_Lwi2 As Double
Dim dbl_Lwi3 As Double
Dim dbl_Lwi4 As Double
Dim dbl_Lwi5 As Double
Dim dbl_Tl1 As Double
Dim dbl_Tl2 As Double
Dim dbl_Tl3 As Double
Dim dbl_Tl4 As Double
Dim dbl_Tl5 As Double
Dim dbl_Ip As Double
Dim dbl_Lwe1 As Double
Dim dbl_Lwe2 As Double
Dim dbl_Lwe3 As Double
Dim dbl_Lwe4 As Double
Dim dbl_Lwe5 As Double
Dim dbl_Xfz As Double
Dim dbl_DeltaLfc As Double

'******************************
'Check input data
'******************************
'One and only one flow rate can be specified, and flow rates cannot be negative.
If (dbl_W <= 0 And dbl_Q <= 0) Then GoTo LiquidNoise_Error
'Pipe diameters cannot be negative
If (dbl_OutletPipeOD <= 0 Or dbl_Tk <= 0) Then GoTo LiquidNoise_Error

'*****************************************************************************************
'*  Start Code
'*****************************************************************************************
On Error GoTo LiquidNoise_Error

'Convert pressure data to pressure in unit of Pascal:

dbl_P1Pa = dbl_P1 * 10000
dbl_P2Pa = dbl_P2 * 10000
dbl_Pvpa = dbl_Pv * 10000
dbl_Pcpa = dbl_Pc * 10000

'If flow rate is in weight units then convert the flow rate in kg/h to kg/sec
If dbl_W > 0 Then
    dbl_M = dbl_W / 3600
'If flow rate is in volume units, then convert from m3/h to kg/sec
Else
    dbl_M = dbl_Q * dbl_Gf * 1000 / 3600
End If

'Calculate the density of the fluid
dbl_Pfluid = 1000 * dbl_Gf

'Set pipe data
dbl_Do = dbl_OutletPipeOD / 1000
dbl_Di = (dbl_OutletPipeOD - 2 * dbl_Tk) / 1000
dbl_T = dbl_Tk / 1000

'calculate ring frequency
dbl_Fr = dbl_Cpipe / (PI * dbl_Do)

'calculate the differential pressure ratio Xf
dbl_Xf = (dbl_P1Pa - dbl_P2Pa) / (dbl_P1Pa - dbl_Pvpa)

'Get Xfz
GetXFZ str_ValveLine, lng_FlowDirectionID, lng_TrimTypeID, _
    bln_ReducedArea, dbl_RatedCv, dbl_Cv, dbl_Xfz, dbl_DLfc

If dbl_Xf < dbl_Xfz Then
    Call Lwi_subcritical(dbl_P1Pa, dbl_P2Pa, dbl_Nf, dbl_Pfluid, _
    dbl_Wo, dbl_M, dbl_Lwi)
ElseIf dbl_Xf >= dbl_Xfz Then
    Call Lwi_critical(dbl_Xfz, dbl_Nf, dbl_DeltaLfc, dbl_Fb, dbl_M, _
    dbl_Pfluid, dbl_Xf, dbl_P1Pa, dbl_P2Pa, dbl_Pcpa, dbl_Pvpa, _
    dbl_FL, dbl_Lwi)
End If
 
'Calculate the internal frequency spectrum of sound power level(ref IEC 534-8-4, 6.2.3)
Call Lwi_f(dbl_Lwi, dbl_Lwi1, dbl_Lwi2, dbl_Lwi3, dbl_Lwi4, dbl_Lwi5)

'Calculate the transmission loss
Call TL_f(dbl_Cpipe, dbl_Ppipe, dbl_T, dbl_Cfluid, dbl_Pfluid, _
dbl_Do, dbl_Fr, dbl_Tl1, dbl_Tl2, dbl_Tl3, dbl_Tl4, dbl_Tl5)

'Calculate the external sound power level
Call Lwe_f(dbl_Lwi1, dbl_Lwi2, dbl_Lwi3, dbl_Lwi4, dbl_Lwi5, _
dbl_Tl1, dbl_Tl2, dbl_Tl3, dbl_Tl4, dbl_Tl5, dbl_Lp, dbl_Do, _
dbl_Lwe1, dbl_Lwe2, dbl_Lwe3, dbl_Lwe4, dbl_Lwe5)

'calculate the A-weighted sound power level
Call Lwae(dbl_Lwe1, dbl_Lwe2, dbl_Lwe3, dbl_Lwe4, dbl_Lwe5, dbl_Lwae)
Call Lpae(dbl_Lwae, dbl_Lp, dbl_Di, dbl_Do, dbl_Lo, dbl_LpAe)


'*****************************************************************************************
'*  Common Exit Point
'*****************************************************************************************
LiquidNoise_Exit:
Exit Sub

'*****************************************************************************************
'*  Standard Error Handling
'*****************************************************************************************
LiquidNoise_Error:
'Placeholder
    
End Sub


Sub LiqParameters(str_FluidName As String, dbl_Temperature As Double, _
str_TemperatureUM As String, dbl_Pv As Double, dbl_Pc As Double, _
str_PressureUM As String, dbl_Viscosity As Double, dbl_Gf As Double, _
dbl_Fb As Double)

'------------------------------------------------------------------------------------------------
'This subroutine provides the vapor pressure and critical pressure for the
'fluid selected. The vapor pressure is calculated for fluids in table 1 and
'the critical pressure is stored by fluid name in table 1.   The specific
'gravity is calculated for water.
'------------------------------------------------------------------------------------------------
'Input                  Output         Description                      Units
'str_FluidName                             The fluid description            -
'dbl_TemperatureK    Temperature                      Kelvin

'Output Variables
'dbl_Gf                 Specific gravity                 -
'dbl_Pv                 Vapor pressure                   bar Absolute
'dbl_Pc                 Critical pressure                bar Absolute
'dbl_Viscosity          Absolute viscosity               centipoises
'dbl_Fb         Briggs factor
'Author:  KPMG                  Date:  6/12/98
'------------------------------------------------------------------------------------------------
On Error GoTo liqPvPcViscGf_Error

Dim dbl_TemperatureK As Double
Call convertTemperatureToDegK(dbl_Temperature, str_TemperatureUM, _
    dbl_TemperatureK)

'********************************
'Check input values
'********************************
'TemperatureK cannot be less than zero
If dbl_TemperatureK < 0 Then GoTo liqPvPcViscGf_Error
'Fluid name must be provided
If str_FluidName = "" Then GoTo liqPvPcViscGf_Error

'*******************************
'Declarations
'*******************************
Dim sng_tcelsius As Single
Dim sng_x As Single
Dim sng_A As Single
Dim sng_B As Single
Dim sng_C As Single
Dim sng_D As Single
Dim sng_E As Single
Dim dbl_Vpbar As Double
Dim sng_Tmax As Single
Dim sng_Tmin As Single
Dim dbl_PCritical As Double

'*******************************
'Begin calculations
'*******************************
'default in case we don't have it
dbl_Viscosity = 1#

Call convertTemperatureFromDegK(dbl_TemperatureK, "deg C", sng_tcelsius)
dbl_Fb = 1#
'Select values based on fluid name

Select Case str_FluidName
    Case "Ammonia"
        If (sng_tcelsius >= -65 And sng_tcelsius <= 132.4) Then
            sng_A = 38.44
            sng_B = -2066.2
            sng_C = -12.105
            sng_D = 7.7768
            sng_E = 0
            dbl_PCritical = 113.05
            dbl_Fb = 0.4
        End If
    Case "Andine"
        If (sng_tcelsius >= 77.4 And sng_tcelsius <= 426) Then
            sng_A = 18.893
            sng_B = -3104
            sng_C = -3.4714
            sng_D = 0.0274
            sng_E = 0
            dbl_PCritical = 52.77
        End If
    Case "Benzene"
        If (sng_tcelsius >= 7.6 And sng_tcelsius <= 289.4) Then
            sng_A = 51.204
            sng_B = -3245.7
            sng_C = -16.403
            sng_D = 7.54
            sng_E = 0
            dbl_PCritical = 48.94
        End If
    Case "n-Butanol"
        If (sng_tcelsius >= -1.2 And sng_tcelsius < 289.8) Then
            sng_A = -458.03
            sng_B = 7760.4
            sng_C = 199.14
            sng_D = -229.44
            sng_E = 95.289
            dbl_PCritical = 44
        End If
    Case "Butadiene"
        If (sng_tcelsius >= -108.9 And sng_tcelsius <= 152) Then
            sng_A = 41.401
            sng_B = -2181.2
            sng_C = -13.451
            sng_D = 8.4524
            sng_E = 0
            dbl_PCritical = 42.87
        End If
    Case "1-Butene"
        If (sng_tcelsius >= -130 And sng_tcelsius <= 146.3) Then
            sng_A = 41.61
            sng_B = -2158.8
            sng_C = -13.58
            sng_D = 8.6536
            sng_E = 0
            dbl_PCritical = 39.96
        End If
    Case "Butylene Oxide"
        If (sng_tcelsius >= -9 And sng_tcelsius <= 252.6) Then
            sng_A = 5.0952
            sng_B = -1361.9
            sng_C = 0.2377
            sng_D = 0.6957
            sng_E = 0
            dbl_PCritical = 4.38
        End If
    Case "Carbon Dioxide"
        If (sng_tcelsius >= -56 And sng_tcelsius <= 31.2) Then
            sng_A = 47.544
            sng_B = -1792.2
            sng_C = -16.559
            sng_D = 13.833
            sng_E = 0
            dbl_PCritical = 73.77
        End If
    Case "Carbon Monoxide"
        If (sng_tcelsius >= -200.7 And sng_tcelsius <= -140) Then
            sng_A = 32.863
            sng_B = -606.91
            sng_C = -12.969
            sng_D = 27.551
            sng_E = 0
            dbl_PCritical = 35.27
        End If
    Case "Carbon Tetrachloride"
        If (sng_tcelsius >= -69.7 And sng_tcelsius <= 283.2) Then
            sng_A = 50.612
            sng_B = -3135.7
            sng_C = -16.313
            sng_D = 7.8036
            sng_E = 0
            dbl_PCritical = 45.4
        End If
    Case "Chlorine"
        If (sng_tcelsius >= -100.4 And sng_tcelsius <= 144.1) Then
            sng_A = 42.262
            sng_B = -2009.9
            sng_C = -13.963
            sng_D = 9.3705
            sng_E = 0
            dbl_PCritical = 77.81
        End If
    Case "Chlorobenzene"
        If (sng_tcelsius >= -35.1 And sng_tcelsius <= 359.2) Then
            sng_A = -54.872
            sng_B = -975.28
            sng_C = 27.85
            sng_D = -37.491
            sng_E = 16.625
            dbl_PCritical = 45.42
        End If
    Case "Chloroform"
        If (sng_tcelsius >= -42 And sng_tcelsius <= 263) Then
            sng_A = 26.828
            sng_B = -2292.6
            sng_C = -7.186
            sng_D = 3.1365
            sng_E = 0
            dbl_PCritical = 55.65
        End If
    Case "Chloroprene"
        If (sng_tcelsius >= 6.4 And sng_tcelsius <= 261.7) Then
            sng_A = 11.416
            sng_B = -1785.1
            sng_C = -1.1467
            sng_D = -0.8341
            sng_E = 0
            dbl_PCritical = 42.5
        End If
    Case "Cumene"
        If (sng_tcelsius >= -20 And sng_tcelsius <= 360) Then
            sng_A = 136.71
            sng_B = -6044.1
            sng_C = -51.36
            sng_D = 42.022
            sng_E = -13.746
            dbl_PCritical = 31.54
        End If
    Case "Cyclopropane"
        If (sng_tcelsius >= -90 And sng_tcelsius <= 125.2) Then
            sng_A = 38.45
            sng_B = -1865.2
            sng_C = -12.578
            sng_D = 8.9375
            sng_E = 0
            dbl_PCritical = 55.77
        End If
    Case "Cyclobutane"
        If (sng_tcelsius >= -92 And sng_tcelsius <= 190.4) Then
            sng_A = 211.96
            sng_B = -5392.8
            sng_C = -89.626
            sng_D = 123.13
            sng_E = -64.64
            dbl_PCritical = 48.43
        End If
    Case "Cyclopentane"
        If (sng_tcelsius >= -60 And sng_tcelsius <= 238.5) Then
            sng_A = 45.484
            sng_B = -2726.5
            sng_C = -14.599
            sng_D = 7.67
            sng_E = 0
            dbl_PCritical = 45.27
        End If
    Case "Cyclohexane"
        If (sng_tcelsius >= -25.4 And sng_tcelsius <= 280.3) Then
            sng_A = 64.753
            sng_B = -3619.2
            sng_C = -21.753
            sng_D = 10.742
            sng_E = 0
            dbl_PCritical = 41.04
        End If
    Case "Ethane"
        If (sng_tcelsius >= -179.5 And sng_tcelsius <= 32.3) Then
            sng_A = 16.316
            sng_B = -1074.8
            sng_C = -3.1434
            sng_D = -4.5534
            sng_E = 10.373
            dbl_PCritical = 48.71
        End If
    Case "Ethanol"
        If (sng_tcelsius >= -45 And sng_tcelsius <= 243) Then
            sng_A = -10.967
            sng_B = -2212.6
            sng_C = -10.298
            sng_D = -21.061
            sng_E = 10.748
            dbl_PCritical = 63.06
        End If
    Case "Ethylbenzene"
        If (sng_tcelsius >= -30 And sng_tcelsius <= 346) Then
            sng_A = 60.335
            sng_B = -4012.8
            sng_C = -19.569
            sng_D = 8.462
            sng_E = 0
            dbl_PCritical = 38.46
        End If
    Case "Ethylene"
        If (sng_tcelsius >= -169 And sng_tcelsius <= 10.1) Then
            sng_A = 30.895
            sng_B = -1196.8
            sng_C = -10.153
            sng_D = 9.9351
            sng_E = 0
            dbl_PCritical = 51.47
        End If
    Case "Ethyl Oxide"
        If (sng_tcelsius >= 73.8 And sng_tcelsius <= 196) Then
            sng_A = 40.723
            sng_B = -2372.4
            sng_C = -12.865
            sng_D = 7.3243
            sng_E = 0
            dbl_PCritical = 72.53
        End If
    Case "Hydrazine"
        If (sng_tcelsius >= 15 And sng_tcelsius <= 380) Then
            sng_A = 60.878
            sng_B = -3880.3
            sng_C = -20.575
            sng_D = 15.585
            sng_E = -5.0525
            dbl_PCritical = 146.78
        End If
    Case "Hydrogen"
        If (sng_tcelsius >= -259.4 And sng_tcelsius <= -240.2) Then
            sng_A = 5.2366
            sng_B = -46.28
            sng_C = -0.4481
            sng_D = 25.29
            sng_E = 0
            dbl_PCritical = 12.89
        End If
    Case "Hydrogen Bromide"
        If (sng_tcelsius >= -83 And sng_tcelsius <= 90#) Then
            sng_A = -351.11
            sng_B = 5375.2
            sng_C = 161.89
            sng_D = -259.04
            sng_E = 157.16
            dbl_PCritical = 83.96
        End If
    Case "Hydrogen Chloride"
        If (sng_tcelsius >= -150.8 And sng_tcelsius <= 51.5) Then
            sng_A = 136.05
            sng_B = -3047.3
            sng_C = -58.416
            sng_D = 95.496
            sng_E = -58.507
            dbl_PCritical = 82.72
        End If
    Case "Hydrogen Fluoride"
        If (sng_tcelsius >= -74.7 And sng_tcelsius <= 188) Then
            sng_A = 66.244
            sng_B = -2588
            sng_C = -25.14
            sng_D = 28.493
            sng_E = -9.9602
            dbl_PCritical = 64.41
        End If
    Case "Hydrogen Iodide"
        If (sng_tcelsius >= -59.5 And sng_tcelsius <= 151) Then
            sng_A = 33.943
            sng_B = -1777.9
            sng_C = -10.62
            sng_D = 6.9457
            sng_E = 0
            dbl_PCritical = 82.73
        End If
    Case "Isobutylene"
        If (sng_tcelsius >= -140 And sng_tcelsius <= 144.9) Then
            sng_A = 37.475
            sng_B = -2038.7
            sng_C = -11.944
            sng_D = 7.6217
            sng_E = 0
            dbl_PCritical = 39.92
        End If
    Case "Isoprene"
        If (sng_tcelsius >= 78.1 And sng_tcelsius <= 210.2) Then
            sng_A = 37.675
            sng_B = -2427.2
            sng_C = -11.47
            sng_D = 5.3274
            sng_E = 0
            dbl_PCritical = 36.68
        End If
    Case "Methane"
        If (sng_tcelsius >= -182.2 And sng_tcelsius <= -81.9) Then
            sng_A = 22.573
            sng_B = -656.24
            sng_C = -7.3942
            sng_D = 11.896
            sng_E = 0
            dbl_PCritical = 46.86
        End If
    Case "Methanol"
        If (sng_tcelsius >= -67.4 And sng_tcelsius <= 240) Then
            sng_A = -42.629
            sng_B = -1186.2
            sng_C = 23.279
            sng_D = -35.082
            sng_E = 17.578
            dbl_PCritical = 79.82
        End If
    Case "Methyl Chloride"
        If (sng_tcelsius >= -99.5 And sng_tcelsius <= 143.1) Then
            sng_A = 42.827
            sng_B = -2100.9
            sng_C = -14.14
            sng_D = 9.5055
            sng_E = 0
            dbl_PCritical = 66.58
        End If
    Case "0-Methylene Chloride"
        If (sng_tcelsius >= -40 And sng_tcelsius <= 241) Then
            sng_A = 65.845
            sng_B = -2732.3
            sng_C = -25.385
            sng_D = 35.167
            sng_E = -19.359
            dbl_PCritical = 62.79
        End If
    Case "Napthalene"
        If (sng_tcelsius >= 82.9 And sng_tcelsius <= 475) Then
            sng_A = 192.16
            sng_B = -8336.3
            sng_C = -72.834
            sng_D = 56.768
            sng_E = -17.319
            dbl_PCritical = 39.35
        End If
    Case "Nitric Oxide"
        If (sng_tcelsius >= -156.8 And sng_tcelsius <= -93) Then
            sng_A = 258.32
            sng_B = -4361
            sng_C = -115.06
            sng_D = 167.15
            sng_E = 0
            dbl_PCritical = 65.45
        End If
    Case "Nitrogen"
        If (sng_tcelsius >= -210 And sng_tcelsius <= -146.8) Then
            sng_A = 21.623
            sng_B = -455.57
            sng_C = -7.5107
            sng_D = 17.214
            sng_E = 0
            dbl_PCritical = 34.1
        End If
    Case "Nitrogen Dioxide"
        If (sng_tcelsius >= -11.3 And sng_tcelsius <= 158.4) Then
            sng_A = 33.024
            sng_B = -2276.7
            sng_C = -10.143
            sng_D = 8.951
            sng_E = 0
            dbl_PCritical = 102.01
        End If
    Case "Nitrous Oxide"
        If (sng_tcelsius >= -88.5 And sng_tcelsius <= 36.5) Then
            sng_A = 54.061
            sng_B = -1894.7
            sng_C = 19.406
            sng_D = 16.572
            sng_E = 0
            dbl_PCritical = 72.51
        End If
    Case "Oxygen"
        If (sng_tcelsius >= -219.1 And sng_tcelsius <= -118.4) Then
            sng_A = 5.6486
            sng_B = -411.3
            sng_C = 1.8118
            sng_D = -25.042
            sng_E = 62.612
            dbl_PCritical = 50.91
        End If
    Case "Phenol"
        If (sng_tcelsius >= 67.8 And sng_tcelsius <= 420) Then
            sng_A = 672.71
            sng_B = -22197.3
            sng_C = -266.67
            sng_D = 226.43
            sng_E = -73.731
            dbl_PCritical = 59.25
        End If
    Case "Propane"
        If (sng_tcelsius >= -128.9 And sng_tcelsius <= 96.8) Then
            sng_A = 36.007
            sng_B = -1737.2
            sng_C = -11.666
            sng_D = 8.5187
            sng_E = 0
            dbl_PCritical = 42.4
        End If
    Case "n-Propanol"
        If (sng_tcelsius >= 0 And sng_tcelsius <= 263.7) Then
            sng_A = -338.31
            sng_B = 5127.5
            sng_C = 148.8
            sng_D = -175.79
            sng_E = 74.666
            dbl_PCritical = 51.79
        End If
    Case "Propylene"
        If (sng_tcelsius >= -150 And sng_tcelsius <= 92.1) Then
            sng_A = 36.877
            sng_B = -1725.5
            sng_C = -12.057
            sng_D = 8.9948
            sng_E = 0
            dbl_PCritical = 46.3
        End If
    Case "Propyl Oxide"
        If (sng_tcelsius >= -57.8 And sng_tcelsius <= 209.1) Then
            sng_A = 42.063
            sng_B = -2559.5
            sng_C = -13.309
            sng_D = 7.3076
            sng_E = 0
            dbl_PCritical = 49.27
        End If
    Case "Sulfur Dioxide"
        If (sng_tcelsius >= -67.6 And sng_tcelsius <= 157.6) Then
            sng_A = 46.554
            sng_B = -2456.3
            sng_C = -15.169
            sng_D = 9.0026
            sng_E = 0
            dbl_PCritical = 78.61
        End If
    Case "Sulfur Trioxide"
        If (sng_tcelsius >= -1 And sng_tcelsius <= 218.3) Then
            sng_A = 160.89
            sng_B = -8081.2
            sng_C = -54.24
            sng_D = 4.3154
            sng_E = 17.432
            dbl_PCritical = 82.45
        End If
    Case "Toluene"
        If (sng_tcelsius >= -60 And sng_tcelsius <= 320.6) Then
            sng_A = 115.21
            sng_B = -4918.1
            sng_C = -43.467
            sng_D = 38.548
            sng_E = -13.496
            dbl_PCritical = 41.8
        End If
    Case "M-xylene"
        If (sng_tcelsius >= -30 And sng_tcelsius <= 346) Then
            sng_A = 58.53
            sng_B = -3990.2
            sng_C = -18.835
            sng_D = 7.9678
            sng_E = 0
            dbl_PCritical = 36.28
        End If
    Case "O-xylene"
        If (sng_tcelsius >= -20 And sng_tcelsius <= 358.4) Then
            sng_A = 56.025
            sng_B = -3955.6
            sng_C = -17.831
            sng_D = 7.3259
            sng_E = 0
            dbl_PCritical = 37.92
        End If
    Case "P-xylene"
        If (sng_tcelsius >= 20 And sng_tcelsius <= 345) Then
            sng_A = 57.096
            sng_B = -3923.6
            sng_C = -18.309
            sng_D = 7.7401
            sng_E = 0
            dbl_PCritical = 35.94
        End If
    Case "Water"
        If (sng_tcelsius >= 0 And sng_tcelsius <= 374.2) Then
            sng_A = 16.373
            sng_B = -2818.6
            sng_C = -1.6908
            sng_D = -5.7546
            sng_E = 4.0073
            dbl_PCritical = 221.05
            dbl_Fb = 0.4
        End If
End Select

'Calculate sng_x and dbl_Vpbar
sng_x = sng_A + (sng_B / dbl_TemperatureK) + sng_C * Log(dbl_TemperatureK) / Log(10) + sng_D * (dbl_TemperatureK / 1000) + (sng_E * dbl_TemperatureK ^ 2) / 1000000
dbl_Vpbar = 0.00133 * (10 ^ sng_x)
dbl_Pv = dbl_Vpbar
dbl_Pc = dbl_PCritical
CnvrtPressureFromBarA dbl_Pv, str_PressureUM, dbl_Pv
CnvrtPressureFromBarA dbl_Pc, str_PressureUM, dbl_Pc

If str_FluidName = "Water" Then
    If (dbl_Vpbar < 0.05) Then
        dbl_Gf = 1
    ElseIf (dbl_Vpbar >= 0.05 And dbl_Vpbar < 80) Then
        Dim sng_a1 As Single
        Dim sng_b1 As Single
        Dim sng_c1 As Single
        Dim sng_d1 As Single
        Dim sng_e1 As Single
        Dim sng_f1 As Single
        Dim sng_g1 As Single
        
        sng_a1 = -0.0004775
        sng_b1 = 0.0001288
        sng_c1 = 0.03615
        sng_d1 = 0.004132
        sng_e1 = 0.00001103
        sng_f1 = 0.000000003544
        sng_g1 = 1.007
        
        dbl_Gf = 1 / (sng_a1 * dbl_Vpbar + (sng_b1 / dbl_Vpbar) + (sng_c1 * Sqr(dbl_Vpbar)) + (sng_d1 * Log(dbl_Vpbar)) + (sng_e1 * dbl_Vpbar ^ 2) + sng_f1 * dbl_Vpbar ^ 3 + sng_g1)
    ElseIf (dbl_Vpbar >= 80 And dbl_Vpbar <= 221) Then
        Dim sng_h1 As Single
        Dim sng_i1 As Single
        Dim sng_j1 As Single
        Dim sng_k1 As Single
        
        sng_h1 = -0.389
        sng_i1 = -0.00001519
        sng_j1 = 0.005294
        sng_k1 = 2.1
        
        dbl_Gf = sng_h1 * Log(dbl_Vpbar) + sng_i1 * dbl_Vpbar ^ 2 + sng_j1 * dbl_Vpbar + sng_k1
        
    End If
End If

If (str_FluidName = "Water" Or str_FluidName = "Ammonia" Or _
str_FluidName = "Sulfur Dioxide" Or str_FluidName = "Ethane" _
Or str_FluidName = "Propane" Or str_FluidName = "1-Methylchloride" _
Or str_FluidName = "Carbon Dioxide") Then
    
    Dim sng_a2 As Single
    Dim sng_b2 As Single
    Dim sng_c2 As Single
    Dim sng_d2 As Single
    Dim dbl_TemperatureF As Double
    Dim dbl_TemperatureR As Double
    
    Call convertTemperatureFromDegK(dbl_TemperatureK, "deg F", dbl_TemperatureF)
    Select Case str_FluidName
        Case "Water"
            If (dbl_TemperatureF >= 32 And dbl_TemperatureF <= 200) Then
                sng_a2 = 30.27791
                sng_b2 = -0.09514759
                sng_c2 = 0.00007535112
                sng_d2 = 0
            ElseIf (dbl_TemperatureF > 200 And dbl_TemperatureF <= 600) Then
                sng_a2 = 3.723846
                sng_b2 = -0.01038628
                sng_c2 = 0.00001006155
                sng_d2 = -0.000000003303412
            End If
        Case "Ammonia"
            If (dbl_TemperatureF >= -40 And dbl_TemperatureF <= 240) Then
                sng_a2 = 4.858626
                sng_b2 = -0.02203967
                sng_c2 = 0.00003409575
                sng_d2 = 0.00000001778858
            End If
        Case "Sulfur Dioxide"
            If (dbl_TemperatureF >= -30 And dbl_TemperatureF <= 30) Then
                sng_a2 = 5.544214
                sng_b2 = -0.01895345
                sng_c2 = 0.00004650391
                sng_d2 = 0
            End If
        Case "Ethane"
            If (dbl_TemperatureF >= 15 And dbl_TemperatureF <= 80) Then
                sng_a2 = 0.1070915
                sng_b2 = 0.0004650391
                sng_c2 = -0.000001287632
                sng_d2 = 0.000000003299098
            End If
        Case "Propane"
            If (dbl_TemperatureF >= 15 And dbl_TemperatureF <= 200) Then
                sng_a2 = 1.18963
                sng_b2 = -0.004417849
                sng_c2 = 0.000000006235607
                sng_d2 = -0.000000003299098
            End If
        Case "1-Methylchloride"
            If (dbl_TemperatureF >= -40 And dbl_TemperatureF <= 240) Then
                sng_a2 = 1.845704
                sng_b2 = -0.00655171
                sng_c2 = 0.0000088995
                sng_d2 = 0.000000004245605
            End If
        Case "Carbon Dioxide"
            If (dbl_TemperatureF >= -20 And dbl_TemperatureF <= 32) Then
                sng_a2 = 0.5073639
                sng_b2 = -0.0001229479
                sng_c2 = -0.0000017419
                sng_d2 = 4.697499E-11
            End If
    End Select
    'Convert temperature to Rankine
    Call convertTemperatureFromDegK(dbl_TemperatureK, "deg R", dbl_TemperatureR)
    dbl_Viscosity = sng_a2 + sng_b2 * dbl_TemperatureR + sng_c2 * dbl_TemperatureR ^ 2 + sng_d2 * dbl_TemperatureR ^ 3
End If

Exit Sub

liqPvPcViscGf_Error:
'Placeholder
End Sub

Public Function GetIECTrimType(long_TrimTypeID As Long) As Long
'---------------------------------------------------------------
'this routine gets the IEC trim type from the Trim Type table given
'   the regular trim type id
'
' written 11/20/98 Schoonover
'---------------------------------------------------------------
'************************
'Check input values
'************************
On Error GoTo GetIECTrimType_Error

'str_ValveLine must be provided
If (long_TrimTypeID = 0) Then GoTo GetIECTrimType_Error

'************************
'Declarations
'************************
Dim str_Query As String
Dim str_DbPath As String
Dim rst_Results As Recordset

'************************
'Process
'************************
str_Query = "select IEC_ID from tblTrimType where TRIM_TYPE_ID = " + CStr(long_TrimTypeID)

Set rst_Results = g_db_Prod.OpenRecordset(str_Query, dbReadOnly)

If (rst_Results.RecordCount = 0) Then GoTo GetIECTrimType_Error
If (IsNull(rst_Results.Fields!IEC_ID)) Then GoTo GetIECTrimType_Error

GetIECTrimType = rst_Results.Fields!IEC_ID
rst_Results.Close

Exit Function

GetIECTrimType_Error:
MsgBox "Error getting IEC Trim type =" + Error
End Function

Public Sub GetFm(str_ValveLine As String, long_TrimTypeID As Long, _
    lng_TrimPackageID As Long, dbl_Fm As Double)
'-------------------------------------------------------------------------
'This subroutine computes Fm from valve line and trim type - it currently
'   is hard coded rather than database lookup because there are very limited
'   choices available
'
'written 11/20/98 Schoonover
'-------------------------------------------------------------------------
'************************
'Check input values
'************************
On Error GoTo GetFm_Error

'str_ValveLine must be provided
If (long_TrimTypeID = 0) Then GoTo GetFm_Error

'************************
'Declarations
'************************
Dim long_IECTrimTypeID As Long

'************************
'Process
'************************
long_IECTrimTypeID = GetIECTrimType(long_TrimTypeID)
Select Case long_IECTrimTypeID
    Case 1 To 2
        dbl_Fm = 1#
    Case 3 To 4
        dbl_Fm = 0.74
End Select

Exit Sub
GetFm_Error:
End Sub

Sub GetFL(str_ValveLine As String, lng_FlowDirectionID As Long, _
    lng_TrimTypeID As Long, bln_ReducedArea As Boolean, _
    dbl_RatedCv As Double, dbl_Cv As Double, dbl_FL As Double)

'-------------------------------------------------------------------------
'This subroutine interpolates a set of data for FL to find the value
'of FL as the calculated value for Cv
'
'Input Variables:
'str_ModelNumber        Valve model number,
'lng_TrimTypeID         Trim Type ID
'lng_FlowDirectionID    Flow direction ID
'dbl_RatedCv            Rated Cv of the valve
'dbl_Cv                 Calculated Cv

'Output Variables:
'dbl_FL                 Liquid pressure recovery factor
'
'Author:  KPMG                  Date:  6/12/98
'modified: 11/20/98 Schoonover - use new key table in db
'-------------------------------------------------------------------------
On Error GoTo GetFL_Error

'************************
'Check input values
'************************

'str_ValveLine must be provided
If (str_ValveLine = "") Then GoTo GetFL_Error
'RatedCv must be greater than 0
If (dbl_RatedCv <= 0) Then GoTo GetFL_Error
'Cv cannot be less than 0
If (dbl_Cv < 0) Then GoTo GetFL_Error

'************************
'Declarations
'************************
Dim lng_FL_DATA_ID As Long
Dim dbl_PctTravel As Double
Dim dbl_FractionTravelLowerBound As Double
Dim dbl_FractionTravelUpperBound As Double
Dim dbl_FLLowerBound As Double
Dim dbl_FLUpperBound As Double
Dim str_Query As String
Dim str_DbPath As String
Dim rst_Results As Recordset

'************************
'Begin calculations
'************************

'Calculate %Travel
Call GetTravel(str_ValveLine, lng_TrimTypeID, dbl_RatedCv, dbl_Cv, dbl_PctTravel)

'Determine upper and lower bounds of FractionCv range
GetBounds dbl_FractionTravelLowerBound, dbl_FractionTravelUpperBound, dbl_PctTravel, 10
If (dbl_PctTravel >= 100) Then
    dbl_FractionTravelLowerBound = 90
    dbl_FractionTravelUpperBound = 100
End If
If (dbl_PctTravel < 10) Then
    dbl_FractionTravelLowerBound = 10
    dbl_FractionTravelUpperBound = 20
End If
    

'Build query to retrieve key into tblFL
str_Query = "Select FL_DATA_ID from tblFlKey where VALVE_LINE = '" _
    + str_ValveLine + "' and (FLOW_DIRECTION_ID = " + Str(lng_FlowDirectionID) _
    + " or FLOW_DIRECTION_ID is null) and (TRIM_TYPE_ID = " _
    + Str(lng_TrimTypeID) + " or TRIM_TYPE_ID is null) and REDUCED_AREA = " _
    + Str(bln_ReducedArea)

Set rst_Results = g_db_Gensz.OpenRecordset(str_Query, , dbReadOnly)

If (rst_Results.RecordCount = 0) Then GoTo GetFL_Error
If (IsNull(rst_Results.Fields!FL_DATA_ID)) Then GoTo GetFL_Error

lng_FL_DATA_ID = rst_Results.Fields!FL_DATA_ID
rst_Results.Close

'Build query to retrieve FL at FractionCv bounds
str_Query = "Select FL" + CStr(dbl_FractionTravelLowerBound) + ", FL" + _
CStr(dbl_FractionTravelUpperBound) + " from tblFL where FL_DATA_ID = " + Str(lng_FL_DATA_ID)

Set rst_Results = g_db_Gensz.OpenRecordset(str_Query, , dbReadOnly)
With rst_Results
    dbl_FLLowerBound = .Fields(0)
    dbl_FLUpperBound = .Fields(1)
    .Close
End With


Call Interpolate(dbl_PctTravel, dbl_FractionTravelLowerBound, _
    dbl_FractionTravelUpperBound, dbl_FLLowerBound, dbl_FLUpperBound, dbl_FL)

Exit Sub

GetFL_Error:
'Placeholder

End Sub


Sub LiquidCv(str_ValveLine As String, lng_FlowDirectionID As Long, _
    lng_TrimTypeID As Long, bln_ReducedArea As Boolean, int_FlowStateFlag As Integer, _
    dbl_P1 As Double, dbl_P2 As Double, dbl_Pv As Double, _
    dbl_Pc As Double, dbl_Gf As Double, dbl_FL As Double, _
    dbl_Fp As Double, dbl_Flp As Double, dbl_Q As Double, _
    dbl_W As Double, dbl_Viscosity As Double, _
    dbl_InletPipeDia As Double, dbl_OutletPipeDia As Double, _
    dbl_ValveInletDia As Double, dbl_ValveOutletDia As Double, _
    dbl_Ki As Double, dbl_Ksum As Double, dbl_Xfz As Double, dbl_RatedCv As Double, _
    dbl_PctTravel As Double, dbl_Cv As Double)
'-------------------------------------------------------------------------------------------reference:    Masoneilan Control Valve Sizing and Selection Handbook
'            ISA - S75.01-1985,  Flow Equations for Sizing Control Valves
'            IEC 534 - 2 - 1
'
'This subroutine is used to calculate the Cv for liquid process conditions.
'-------------------------------------------------------------------------------------------
'
'Input Variables         Description                                unit
'bln_FlowStateFlag       indicates flashing, cavitation, etc.       5 if flow is laminar, 6 if flow is transitional
'dbl_P1                  Inlet pressure                             bar a
'dbl_P2                  Outlet pressure                            bar a
'dbl_Pv                  Vapor pressure                             bar a
'dbl_Pc                  Critical pressure                          bar a
'r1/r0 » dbl_Gf          specific gravity or relative density       -
'dbl_FL                  liquid pressure recovery factor            -
'dbl_Fp                  pipe geometry factor                       -
'dbl_Flp                 combined liquid pressure recovery factor   -
'dbl_Q                   flowrate - volume basis                    m3/h
'dbl_W                   flowrate - weight basis                    kg/h
'dbl_InletPipeDia        Nominal inlet pipe dia                     inches/mm
'dbl_OutletPipeDia       Nominal outlet pipe dia                    inches/mm
'dbl_ValveInletDia       Nominal valve inlet connection size        inches/mm
'dbl_ValveOutletDia      Nominal valve outlet connection size       inches/mm

'dbl_Cv                  valve flow coefficient                     -
'
'Author:  KPMG                  Date:  6/12/98
'-------------------------------------------------------------------------------------------
Dim dbl_PcDrop As Double

On Error GoTo LiquidCv_Error

'***************************
'Check input data
'***************************

'The inlet pressure, outlet pressure, vapor pressure and critical pressure must be greater than 0
If (dbl_P1 <= 0 Or dbl_P2 <= 0 Or dbl_Pv <= 0 Or dbl_Pc <= 0) Then GoTo LiquidCv_Error
'The specific gravity must be greater than 0
If (dbl_Gf <= 0) Then GoTo LiquidCv_Error
'One, and only one flow rate should be defined.  Flow rates must be greater than 0.
If (dbl_Q < 0 Or dbl_W < 0) Or (dbl_Q = 0 And dbl_W = 0) Or _
(dbl_Q > 0 And dbl_W > 0) Then GoTo LiquidCv_Error
'The diameters must be greater than 0
If (dbl_InletPipeDia <= 0 Or dbl_OutletPipeDia <= 0 Or _
dbl_ValveInletDia = 0 Or dbl_ValveOutletDia = 0) Then GoTo LiquidCv_Error

'*****************************
'Declarations
'*****************************
Dim int_IncipientFlag As Integer
Dim dbl_Fd As Double
Dim dbl_DLfc As Double
Dim dbl_QLaminar As Double

'*****************************
'Begin calculations
'*****************************
If (dbl_Q = 0) Then
    Call convertFlowRateFromKgh(dbl_W, "m3/h", 1000 * dbl_Gf, dbl_QLaminar)
Else
    dbl_QLaminar = dbl_Q
End If

Call LaminarCv(str_ValveLine, lng_FlowDirectionID, lng_TrimTypeID, _
    bln_ReducedArea, dbl_QLaminar, dbl_Viscosity, dbl_Gf, dbl_FL, _
    dbl_ValveInletDia, dbl_P1, dbl_P2, dbl_RatedCv, dbl_Cv, int_FlowStateFlag)

'If flow is not laminar, then start turbulent flow calculation
If (int_FlowStateFlag <> 5 And int_FlowStateFlag <> 6) Then
    'calculate the critical pressure drop
    dbl_PcDrop = (dbl_P1 - (0.96 - 0.28 * Sqr(dbl_Pv / dbl_Pc)) * dbl_Pv) * dbl_FL ^ 2
        If dbl_P2 <= dbl_Pv Then
            'Fluid is flashing
            int_FlowStateFlag = 3
            Call CalcFp(dbl_ValveInletDia, dbl_Ksum, dbl_Ki, dbl_Cv, dbl_FL, dbl_Fp, dbl_Flp)
                If ((dbl_P1 - dbl_P2) > dbl_PcDrop) Then
                    Call CvFlashingLiq(dbl_P1, dbl_P2, dbl_Pv, dbl_Pc, _
                        dbl_Gf, dbl_Q, dbl_W, dbl_FL, dbl_Fp, dbl_Flp, _
                        dbl_InletPipeDia, dbl_ValveInletDia, dbl_Cv)
                Else
                    Call CvSubCriticalLiq(dbl_P1, dbl_P2, dbl_Pv, _
                        dbl_Pc, dbl_Gf, dbl_Q, dbl_W, dbl_Fp, dbl_InletPipeDia, _
                        dbl_OutletPipeDia, dbl_ValveInletDia, dbl_ValveOutletDia, dbl_Cv)
                End If
        ElseIf ((dbl_P1 - dbl_P2) > dbl_PcDrop) And dbl_P2 > dbl_Pv Then
            'Fluid is cavitating
            int_FlowStateFlag = 2
            Call CalcFp(dbl_ValveInletDia, dbl_Ksum, dbl_Ki, dbl_Cv, _
                dbl_FL, dbl_Fp, dbl_Flp)
            Call CvCriticalLiq(dbl_P1, dbl_P2, dbl_Pv, dbl_Pc, _
                dbl_Gf, dbl_Q, dbl_W, dbl_FL, dbl_Fp, dbl_Flp, _
                dbl_InletPipeDia, dbl_ValveInletDia, dbl_Cv)
        ElseIf ((dbl_P1 - dbl_P2) <= dbl_PcDrop) And dbl_P2 > dbl_Pv Then
            'Fluid is subcritical
            int_FlowStateFlag = 1
            Call CalcFp(dbl_ValveInletDia, dbl_Ksum, dbl_Ki, dbl_Cv, _
                dbl_FL, dbl_Fp, dbl_Flp)
            Call CvSubCriticalLiq(dbl_P1, dbl_P2, dbl_Pv, dbl_Pc, _
                dbl_Gf, dbl_Q, dbl_W, dbl_Fp, dbl_InletPipeDia, dbl_OutletPipeDia, _
                dbl_ValveInletDia, dbl_ValveOutletDia, dbl_Cv)
            Call GetXFZ(str_ValveLine, lng_FlowDirectionID, lng_TrimTypeID, _
                bln_ReducedArea, dbl_RatedCv, dbl_Cv, dbl_Xfz, dbl_DLfc)
            Call Incipient(str_ValveLine, lng_FlowDirectionID, _
                lng_TrimTypeID, bln_ReducedArea, dbl_Cv, dbl_RatedCv, _
                dbl_P1, dbl_P2, dbl_Pv, dbl_Pc, dbl_Xfz, dbl_FL, _
                int_IncipientFlag, int_FlowStateFlag)
        End If
End If
Exit Sub
LiquidCv_Error:
'Placeholder
End Sub


Sub SigmaCav(str_ModelNumber As String, _
    str_FlowDirection As String, str_TrimType As String, _
    dbl_P1 As Double, dbl_P2 As Double, dbl_Pv As Double, _
    dbl_Cv As Double, dbl_RatedCv As Double, _
    dbl_PctOpen As Double, dbl_ValveDia As Double, _
    dbl_Gf As Double, dbl_VS As Double, dbl_SS As Double)
'-------------------------------------------------------------------------------------------------------------
'Input Variables         Description
'str_ModelNumber         Valve model
'dbl_P1                  Inlet valve pressure                       psia
'dbl_P2                  Outlet valve presure                       psia
'dbl_Pv                  Vapor pressure                             psia
'dbl_Cv                  Calculated Cv for service condition
'dbl_RatedCv             Rated Cv for valve trim
'dbl_%Open               % of the opening of valve at calculated Cv
'dbl_ValveDia            Nominal valve size                         in.

'Output Variables:
'dbl_Gf                  Specific gravity
'dbl_SS                  Service Sigma
'dbl_VS                  Valve Sigma
'
'Author:  KPMG                  Date:  6/12/98
'-------------------------------------------------------------------------------------------------
On Error GoTo SigmaCav_Error
'***************************
'Check input values
'***************************
'dbl_P1 must be greater than dbl_P2
If (dbl_P1 <= dbl_P2) Then GoTo SigmaCav_Error
'The valve model must be supplied
If (str_ModelNumber = "") Then GoTo SigmaCav_Error
'% Travel cannot be less than 0
If (dbl_PctOpen < 0) Then GoTo SigmaCav_Error
'ValveDia must be greater than 0
If (dbl_ValveDia <= 0) Then GoTo SigmaCav_Error

'**************************
'Declarations
'**************************
Dim dbl_A As Double
Dim dbl_B As Double
Dim dbl_dr As Double
Dim dbl_PctOpenLowerBound As Double
Dim dbl_PctOpenUpperBound As Double
Dim dbl_Pref As Double
Dim dbl_SigmaA As Double
Dim dbl_SigmaALowerBound As Double
Dim dbl_SigmaAUpperBound As Double
Dim rst_Results As Recordset
Dim str_DbPath As String
Dim str_Query As String

'**************************
'Begin calculations
'**************************

'Determine upper and lower bounds of %Open
Call GetBounds(dbl_PctOpenLowerBound, dbl_PctOpenUpperBound, dbl_PctOpen, 10)

'Build query to retrieve values of  dbl_SigmaA, dbl_A, dbl_B and dbl_dr
'for a given value of the valve model, at dbl_PctOpenLowerBound, from the database
str_Query = "Select Open" + dbl_PctOpenLowerBound + ", Open" + dbl_PctOpenUpperBound + ", a, b, dr, Delta_Pref from SigmaCav" + _
"where valve_model = " + str_ModelNumber

Set rst_Results = g_db_Gensz.OpenRecordset(str_Query, , dbReadOnly)

With rst_Results
    dbl_SigmaALowerBound = .Fields(0)
    dbl_SigmaAUpperBound = .Fields(1)
    dbl_A = .Fields(2)
    dbl_B = .Fields(3)
    dbl_dr = .Fields(4)
    dbl_Pref = .Fields(5)
    .Close
End With

Call Interpolate(dbl_SigmaA, dbl_PctOpen, dbl_PctOpenLowerBound, _
    dbl_PctOpenUpperBound, dbl_SigmaALowerBound, dbl_SigmaAUpperBound)

dbl_SS = (dbl_P1 - dbl_Pv) / (dbl_P1 - dbl_P2)
dbl_VS = ((dbl_SigmaA * (dbl_ValveDia / dbl_dr) ^ dbl_B) - 1) * ((((dbl_P1 - dbl_Pv) / dbl_Pref) ^ dbl_A) - 1)

If dbl_VS < dbl_SS Then
    'Valve is acceptable
Else
    'Valve is unacceptable
End If

If dbl_P1 > 1500 And (str_ModelNumber = "41000" Or str_ModelNumber = "21000 containment cavitation") Then
    'Valve is not acceptable
End If

If (str_ModelNumber = "41000" Or str_ModelNumber = "21000 containment cavitation") Then
    Dim dbl_CageVelocity As Double
    dbl_CageVelocity = (560 / 62.4) * Sqr(dbl_Gf * (dbl_P1 - dbl_P2))
    If dbl_CageVelocity > 300 Then
        'Valve is not acceptable
    End If
End If

If dbl_SS < 1.11 And dbl_CageVelocity <= 200 Then
    'Valve is acceptable
Else
    'Valve is unacceptable
End If

If dbl_SS > 1.11 And dbl_SS < 1.23 And dbl_CageVelocity <= 250 Then
    'Valve is acceptable
Else
    'Valve is unacceptable
End If

If dbl_SS > 1.23 And dbl_CageVelocity <= 300 Then
    'Valve is acceptable
Else
    'Valve is unacceptable
End If

If (str_ModelNumber = "41000" Or str_ModelNumber = "21000 containment cavitation") Then
    Dim dbl_BodyVelocity As Double
    dbl_BodyVelocity = (25.5 * dbl_Cv / 62.4 * dbl_ValveDia ^ 2) * Sqr(dbl_Gf * (dbl_P1 - dbl_P2))
    If dbl_BodyVelocity > 50 Then
        'Valve is not acceptable
    Else
        'Valve is acceptable
    End If
End If

If dbl_SS < 1.23 And dbl_BodyVelocity <= 25 Then
    'Valve is acceptable
Else
    'Valve is not acceptable
End If

If dbl_SS >= 1.23 And dbl_BodyVelocity <= 50 Then
    'valve is unacceptable
Else
    'Valve is acceptable
End If

Exit Sub

SigmaCav_Error:
'Placeholder

End Sub

Sub GetXFZ(str_ValveLine As String, lng_FlowDirectionID As Long, _
    lng_TrimTypeID As Long, bln_ReducedArea As Boolean, _
    dbl_RatedCv As Double, dbl_Cv As Double, dbl_Xfz As Double, _
    dbl_DLfc As Double)
'*********************************************************************
'Input Variables:
'str_ModelNumber
'lng_FlowDirectionID
'lng_TrimTypeID
'lng_CharacteristicID
'dbl_Cv
'dbl_RatedCv

'Output Variables:
'dbl_Xfz
'dbl_DeltaLfc
'
'Author:  KPMG                  Date:  6/12/98
'modified: 11/20/98 Schoonover - use new key table in db
'*********************************************************************

On Error GoTo GetXfz_Error

'*******************************************
'Check input data
'*******************************************

'ModelNumber must be specified
If str_ValveLine = "" Then GoTo GetXfz_Error
'RatedCv must be greater than 0
If dbl_RatedCv <= 0 Then GoTo GetXfz_Error

'*******************************************
'Declarations
'*******************************************
Dim dbl_PctRatedCv As Double
Dim dbl_PctTravel As Double
Dim dbl_PctTravelLowerBound As Double
Dim dbl_PctTravelUpperBound As Double
Dim dbl_XfzLowerBound As Double
Dim dbl_XfzUpperBound As Double
Dim lng_NoiseDataID As Long
Dim rst_Results As Recordset
Dim str_DbPath As String
Dim str_Query As String

'*******************************************
'Begin calculations
'*******************************************
'default if no valve line
If (str_ValveLine = "") Then
    dbl_Xfz = 0.362
    dbl_DLfc = 0
    Exit Sub
End If

dbl_PctRatedCv = (dbl_Cv / dbl_RatedCv) * 100

If dbl_PctRatedCv >= 100 Then
    dbl_PctTravel = 99.999
ElseIf dbl_PctRatedCv < 10 Then
    dbl_PctTravel = 10
Else
    Call GetTravel(str_ValveLine, lng_TrimTypeID, _
        dbl_RatedCv, dbl_Cv, dbl_PctTravel)
End If
    

'Build query to determine NOISE_DATA_ID
str_Query = "Select NOISE_DATA_ID from tblXfzKey where VALVE_LINE = '" _
    + str_ValveLine + "' and (FLOW_DIRECTION_ID= " + CStr(lng_FlowDirectionID) _
    + " or FLOW_DIRECTION_ID is null) and (TRIM_TYPE_ID= " _
    + CStr(lng_TrimTypeID) + " or TRIM_TYPE_ID is null) and REDUCED_AREA= " + _
    CStr(bln_ReducedArea)

Set rst_Results = g_db_Gensz.OpenRecordset(str_Query, , dbForwardOnly)

If (rst_Results.RecordCount = 0) Then GoTo GetXfz_Error
If (IsNull(rst_Results.Fields!NOISE_DATA_ID)) Then GoTo GetXfz_Error

lng_NoiseDataID = rst_Results.Fields!NOISE_DATA_ID
rst_Results.Close


'Determine upper and lower bounds of % Travel
Call GetBounds(dbl_PctTravelLowerBound, dbl_PctTravelUpperBound, _
    dbl_PctTravel, 10)


'Build query to retrieve dbl_XfzLowerBound, dbl_XfzUpperBound and dbl_deltaLfc
If dbl_PctTravel < 50 Then
    str_Query = "Select XFZ" + CStr(dbl_PctTravelLowerBound) + ", XFZ" _
    + CStr(dbl_PctTravelUpperBound) + ", Dlf50 from tblXfz where NOISE_DATA_ID=" + CStr(lng_NoiseDataID)
Else
    str_Query = "Select XFZ" + CStr(dbl_PctTravelLowerBound) + ", XFZ" _
    + CStr(dbl_PctTravelUpperBound) + ", Dlf100 from tblXfz where NOISE_DATA_ID=" + CStr(lng_NoiseDataID)
End If

'Open database and retrieve results
Set rst_Results = g_db_Gensz.OpenRecordset(str_Query, , dbForwardOnly)
With rst_Results
    dbl_XfzLowerBound = .Fields(0)
    dbl_XfzUpperBound = .Fields(1)
    dbl_DLfc = .Fields(2)
End With

rst_Results.Close

'Interpolate to find dbl_Xfz at dbl_PctTravel
Call Interpolate(dbl_PctTravel, dbl_PctTravelLowerBound, dbl_PctTravelUpperBound, _
    dbl_XfzLowerBound, dbl_XfzUpperBound, dbl_Xfz)

Exit Sub

GetXfz_Error:
'Placeholder

End Sub

Sub GetTravel(str_ValveLine As String, _
    lng_TrimTypeID As Long, dbl_RatedCv As Double, _
    dbl_Cv As Double, dbl_PctTravel As Double)

'**********************************************************
'This subroutine returns the %Travel based on the
'characteristic, calculated Cv and rated Cv
'
'Input Variables:
'dbl_ModelNumber
'dbL_Characteristic
'dbl_Cv
'dbl_RatedCv

'Output Variables:
'dbl_PctTravel
'
'Author:  KPMG                  Date:  6/12/98
'modified: Schoonover 10/7/98 added db calls
'**********************************************************

On Error GoTo GetTravel_Error

'*******************************************
'Check input data
'*******************************************
'Valve Line must be specified
If str_ValveLine = "" Then GoTo GetTravel_Error
'RatedCv must be greater than zero
If dbl_RatedCv <= 0 Then GoTo GetTravel_Error

'*******************************************
'Declarations
'*******************************************
Dim dbl_PctRatedCv As Double
Dim dbl_PctRatedCvLowerBound As Double
Dim dbl_PctRatedCvUpperBound As Double
Dim dbl_PctTravelLowerBound As Double
Dim dbl_PctTravelUpperBound As Double
Dim long_CurveID As Long
Dim long_FieldLower As Long
Dim rset_Travel As Recordset
Dim str_DatabasePath As String
Dim str_FieldNameLower As String
Dim str_FieldNameUpper As String
Dim str_TravelQuery As String

'*******************************************
'Begin calculations
'*******************************************

dbl_PctRatedCv = (dbl_Cv / dbl_RatedCv) * 100

If dbl_PctRatedCv >= 100 Then
    dbl_PctTravel = 100
    Exit Sub
ElseIf dbl_PctRatedCv <= 0 Then
    dbl_PctTravel = 0
    Exit Sub
Else
    Call GetBounds(dbl_PctRatedCvLowerBound, dbl_PctRatedCvUpperBound, _
        dbl_PctRatedCv, 5)
End If

'get the upper and lower field names for the query
long_FieldLower = dbl_PctRatedCvLowerBound / 5
str_FieldNameLower = "CV" & CStr(long_FieldLower * 5)
str_FieldNameUpper = "CV" & CStr((long_FieldLower + 1) * 5)

'get the proper curve ID
If (lng_TrimTypeID = 2) Then
    If (str_ValveLine = "41000") Then
        long_CurveID = 3
    Else
        long_CurveID = 2
    End If
Else
    long_CurveID = 1
End If

str_TravelQuery = _
    "SELECT tblTravelFromCv.?, tblTravelFromCv.? " _
    & "From tblTravelFromCv " _
    & "WHERE tblTravelFromCv.CV_CURVE_ID=?"

str_TravelQuery = ReplaceStr(str_TravelQuery, "?", _
    str_FieldNameLower, str_FieldNameUpper, Str$(long_CurveID))

Set rset_Travel = g_db_Gensz.OpenRecordset(str_TravelQuery, dbReadOnly)
If (rset_Travel.RecordCount <> 1) Then GoTo GetTravel_Error
dbl_PctTravelLowerBound = rset_Travel.Fields(0)
dbl_PctTravelUpperBound = rset_Travel.Fields(1)
    
Call Interpolate(dbl_PctRatedCv, dbl_PctRatedCvLowerBound, dbl_PctRatedCvUpperBound, _
    dbl_PctTravelLowerBound, dbl_PctTravelUpperBound, dbl_PctTravel)

rset_Travel.Close
Exit Sub

GetTravel_Error:
'Placeholder

End Sub

Sub FlowRateSizing(str_ValveLine As String, lng_FlowDirectionID As Long, _
    lng_TrimTypeID As Long, bln_ReducedArea As Boolean, dbl_Fb As Double, _
    int_FlowStateFlag As Integer, dbl_P1 As Double, dbl_P2 As Double, _
    dbl_Pv As Double, dbl_Pc As Double, dbl_Gf As Double, _
    dbl_FL As Double, dbl_Fp As Double, dbl_Flp As Double, _
    dbl_Q As Double, dbl_W As Double, dbl_InletPipeDia As Double, _
    dbl_OutletPipeDia As Double, dbl_InletPipeOD As Double, dbl_OutletPipeOD As Double, _
    dbl_InletPipeWallThickness As Double, _
    dbl_OutletPipeThickness As Double, dbl_ValveInletDia As Double, _
    dbl_ValveOutletDia As Double, dbl_MaxVelocity As Double, _
    dbl_Cv As Double, dbl_RatedCv As Double, dbl_LpAe As Double, _
    dbl_InletPipeVelocity As Double, dbl_MaxVelocityDiameter As Double, _
    dbl_PctTravel As Double)

'-------------------------------------------------------------------------------------------reference:    Masoneilan Control Valve Sizing and Selection Handbook
'                    ISA - S75.01-1985,  Flow Equations for Sizing Control Valves
'            IEC 534 - 2 - 1
'
'This subroutine is used to calculate the flow rate for liquid process conditions.
'-------------------------------------------------------------------------------------------
'
'Input Variables           Description                             unit
'str_ValveLine
'lng_FlowDirectionID
'lng_TrimTypeID
'bln_ReducedArea
'dbl_Fb
'int_FlowStateFlag         Flow state flag                         None.  Indicates flashing, cavitation, etc.                 -
'dbl_P1                    Inlet pressure                          bar a
'dbl_P2                    Outlet pressure                         bar a
'dbl_Pv                    Vapor pressure                          bar a
'dbL_Pc                    Critical pressure                       bar a
'dbl_Gf » r1/r0            Specific gravity or relative density
'dbl_FL                    Liquid pressure recovery factor         -
'dbl_Fp                    Pipe geometry factor                    -
'dbl_Flp                   Combined liquid pressure recovery factor -
'dbl_InletPipeDia          Nominal inlet pipe dia                  inches/mm
'dbl_OutletPipeDia         Nominal outlet pipe dia                 inches/mm
'dbl_ValveInletDia         Nominal valve inlet connection size     inches/mm
'dbl_ValveOutletDia        Nominal valve outlet connection size    inches/mm
'dbl_Cv                    Valve flow coefficient                  -
'
'Output Variables:
'dbl_Q                     flowrate - volume basis                 m3/h
'dbl_W                     flowrate - weight basis                 kg/h

'note: laminar flow is not considered for flow rate sizing.
'
'Author:  KPMG                  Date:  6/12/98
'-------------------------------------------------------------------------------------------

On Error GoTo FlowRateSizing_Error
'
'*******************************************
'Check input data
'*******************************************
'All pressures must be greater than zero
If ((dbl_P1 <= 0) Or (dbl_P2 <= 0) Or (dbl_Pv <= 0) Or (dbl_Pc <= 0)) Then _
GoTo FlowRateSizing_Error

'All diameters must be greater than zero
If ((dbl_InletPipeDia <= 0) Or (dbl_OutletPipeDia <= 0) Or _
(dbl_ValveInletDia < 0) Or dbl_ValveOutletDia <= 0) Then GoTo FlowRateSizing_Error

'********************************************
'Declarations
'********************************************
Dim dbl_Ksum As Double
Dim dbl_Ki As Double
Dim dbl_Xfz As Double
Dim dbl_DLfc As Double
Dim dbl_PcDrop As Double
Dim int_Error As Integer
Dim int_IncipientFlag As Integer

'********************************************
'Begin calculations
'********************************************
'defaults
dbl_FL = 0.9
dbl_Fp = 1
dbl_Flp = dbl_FL

    
Call PipeGeom(dbl_InletPipeDia, dbl_OutletPipeDia, dbl_ValveInletDia, dbl_ValveOutletDia, dbl_Ksum, dbl_Ki)

Call GetFL(str_ValveLine, lng_FlowDirectionID, lng_TrimTypeID, bln_ReducedArea, dbl_RatedCv, dbl_Cv, dbl_FL)

'Calculate the critical pressure drop
dbl_PcDrop = (dbl_P1 - dbl_Pv * (0.96 - 0.28 * Sqr(dbl_Pv / dbl_Pc))) * dbl_FL ^ 2

'Get dbl_Xfz and dbl_DeltaLFC from database based on valvemodel and trim selected
Call GetXFZ(str_ValveLine, lng_FlowDirectionID, lng_TrimTypeID, _
    bln_ReducedArea, dbl_RatedCv, dbl_Cv, dbl_Xfz, dbl_DLfc)

If dbl_P2 <= dbl_Pv Then
    'Fluid is flashing
    int_FlowStateFlag = 3
    If dbl_P1 - dbl_P2 > dbl_Pc Then
        Call CalcFp(dbl_ValveInletDia, dbl_Ksum, dbl_Ki, dbl_Cv, dbl_FL, dbl_Fp, dbl_Flp)
        Call FlowFlashingLiq(dbl_P1, dbl_P2, dbl_Pv, dbl_Pc, dbl_Gf, dbl_FL, dbl_Fp, _
            dbl_Flp, dbl_InletPipeDia, dbl_ValveInletDia, dbl_Cv, dbl_Q, dbl_W)
    Else
        Call CalcFp(dbl_ValveInletDia, dbl_Ksum, dbl_Ki, dbl_Cv, dbl_FL, dbl_Fp, dbl_Flp)
        Call FlowSubcriticalLiq(dbl_P1, dbl_P2, dbl_Pv, dbl_Pc, dbl_Gf, dbl_Fp, _
            dbl_InletPipeDia, dbl_OutletPipeDia, dbl_ValveInletDia, dbl_ValveOutletDia, _
            dbl_Cv, dbl_Q, dbl_W, int_Error)
    End If
ElseIf dbl_P1 - dbl_P2 > dbl_Pc And dbl_P2 > dbl_Pv Then
    'Fluid is cavitating
    int_FlowStateFlag = 2
    Call CalcFp(dbl_ValveInletDia, dbl_Ksum, dbl_Ki, dbl_Cv, dbl_FL, dbl_Fp, dbl_Flp)
    Call FlowCriticalLiq(dbl_P1, dbl_P2, dbl_Pv, dbl_Pc, dbl_Gf, dbl_FL, dbl_Fp, _
        dbl_Flp, dbl_InletPipeDia, dbl_ValveInletDia, dbl_Cv, dbl_Q, dbl_W)
ElseIf dbl_P1 - dbl_P2 <= dbl_Pc And dbl_P2 > dbl_Pv Then
    'Fluid is subcritical
    int_FlowStateFlag = 1
    Call CalcFp(dbl_ValveInletDia, dbl_Ksum, dbl_Ki, dbl_Cv, dbl_FL, dbl_Fp, dbl_Flp)
    Call FlowSubcriticalLiq(dbl_P1, dbl_P2, dbl_Pv, dbl_Pc, dbl_Gf, dbl_Fp, _
        dbl_InletPipeDia, dbl_OutletPipeDia, dbl_ValveInletDia, dbl_ValveOutletDia, _
        dbl_Cv, dbl_Q, dbl_W, int_Error)
End If
Call Incipient(str_ValveLine, lng_FlowDirectionID, lng_TrimTypeID, _
    bln_ReducedArea, dbl_Cv, dbl_RatedCv, dbl_P1, dbl_P2, dbl_Pv, _
    dbl_Pc, dbl_Xfz, dbl_FL, int_IncipientFlag, int_FlowStateFlag)

Call LiquidNoise(str_ValveLine, lng_FlowDirectionID, lng_TrimTypeID, _
    bln_ReducedArea, dbl_RatedCv, dbl_Cv, dbl_Fb, dbl_P1, dbl_P2, _
    dbl_Pv, dbl_Pc, dbl_Q, dbl_W, dbl_Gf, dbl_OutletPipeOD, _
    dbl_OutletPipeThickness, dbl_FL, dbl_LpAe)

Call LiquidVelocity(dbl_Q, dbl_W, dbl_Gf, dbl_InletPipeOD, _
    dbl_InletPipeWallThickness, dbl_MaxVelocity, dbl_InletPipeVelocity, _
    dbl_MaxVelocityDiameter)

Call GetTravel(str_ValveLine, lng_TrimTypeID, _
    dbl_RatedCv, dbl_Cv, dbl_PctTravel)

Exit Sub
'
FlowRateSizing_Error:
'Placeholder
End Sub

Sub FlowSubcriticalLiq(dbl_P1 As Double, dbl_P2 As Double, _
    dbl_Pv As Double, dbl_Pc As Double, dbl_Gf As Double, _
    dbl_Fp As Double, dbl_InletPipeDia As Double, _
    dbl_OutletPipeDia As Double, dbl_ValveInletDia As Double, _
    dbl_ValveOutletDia As Double, dbl_Cv As Double, dbl_Q As Double, _
    dbl_W As Double, int_Error As Integer)
'********************************************************************
'This routine calculates the flowrate for subcritical flow conditions
'Input Variables:
'dbl_P1                 Inlet Pressure                                  bar a
'dbl_P2                 Outlet Pressure                                 bar a
'dbl_Pv                 Vapor Pressure                                  bar a
'dbl_Pc                 Critical Pressure                               bar a
'dbl_Gf                 Specific Gravity                                -
'dbl_Fp                 Pipe Geometry Factor                            -
'dbl_InletPipeDia       Nominal Inlet Pipe Diameter                     in or mm
'dbl_OutletPipeDia      Nominal Outlet Pipe Diameter                    in or mm
'dbl_ValveInletDia      Nominal Valve Inlet Connection Size             in or mm
'dbl_ValveOutletDia     Nominal Valve Outlet Connection Size            in or mm
'dbl_Cv                 Valve Flow Coefficient                          -

'Output Variables:
'dbl_Q                  Volume Flow Rate                                m3/h
'dbl_W                  Weight Flow Rate                                kg/h
'
'Author:  KPMG                  Date:  6/12/98
'********************************************************************
On Error GoTo FlowSubcriticalLiq_Error
'
'*******************************************
'Check input data
'*******************************************

'dbl_Gf must be greater than 0
If dbl_Gf <= 0 Then GoTo FlowSubcriticalLiq_Error
'dbl_P1 must be greater than dbl_P2
If (dbl_P1 - dbl_P2) <= 0 Then GoTo FlowSubcriticalLiq_Error
'All diameters must be greater than 0
If (dbl_InletPipeDia <= 0 Or dbl_OutletPipeDia <= 0 Or dbl_ValveInletDia <= 0 Or dbl_ValveOutletDia <= 0) Then GoTo FlowSubcriticalLiq_Error


'*******************************************
'Begin calculations
'*******************************************

'Calculate based on whether or not reducers exist

Select Case (dbl_InletPipeDia - dbl_ValveInletDia)
    Case 0 'no reducers
        dbl_Q = 0.865 * dbl_Cv * Sqr((dbl_P1 - dbl_P2) / dbl_Gf)
        dbl_W = 27.3 * dbl_Cv * Sqr((dbl_P1 - dbl_P2) * dbl_Gf * 1000)
    Case Is <> 0 'reducers
        dbl_Q = 0.865 * dbl_Fp * dbl_Cv * Sqr((dbl_P1 - dbl_P2) / dbl_Gf)
        dbl_W = 27.3 * dbl_Fp * dbl_Cv * Sqr((dbl_P1 - dbl_P2) * dbl_Gf * 1000)
End Select

Exit Sub

FlowSubcriticalLiq_Error:
'Placeholder

End Sub

Sub FlowCriticalLiq(dbl_P1 As Double, dbl_P2 As Double, _
    dbl_Pv As Double, dbl_Pc As Double, dbl_Gf As Double, _
    dbl_FL As Double, dbl_Fp As Double, dbl_Flp As Double, _
    dbl_InletPipeDia As Double, dbl_ValveInletDia As Double, _
    dbl_Cv As Double, dbl_Q As Double, dbl_W As Double)
'***********************************************************************************
'This routine calculates the fluid flowrates under critical flow conditions
'Input Variables:
'dbl_P1                 Inlet Pressure                                  bar a
'dbl_P2                 Outlet Pressure                                 bar a
'dbl_Pv                 Vapor Pressure                                  bar a
'dbl_Pc                 Critical Pressure                               bar a
'dbl_Gf                 Specific Gravity                                -
'dbl_Fp                 Pipe Geometry Factor                            -
'dbl_InletPipeDia       Nominal Inlet Pipe Diameter                     in or mm
'dbl_OutletPipeDia      Nominal Outlet Pipe Diameter                    in or mm
'dbl_ValveInletDia      Nominal Valve Inlet Connection Size             in or mm
'dbl_ValveOutletDia     Nominal Valve Outlet Connection Size            in or mm
'dbl_Cv                 Valve Flow Coefficient                          -

'Output Variables:
'dbl_Q                  Volume Flow Rate                                m3/h
'dbl_W                  Weight Flow Rate                                kg/h
'
'Author:  KPMG                  Date:  6/12/98
'********************************************************************
On Error GoTo FlowCriticalLiq_Error

'*******************************************
'Check input data
'*******************************************
'dbl_Gf must be greater than 0
If dbl_Gf <= 0 Then GoTo FlowCriticalLiq_Error
'All diameters must be greater than zero
If (dbl_InletPipeDia <= 0 Or dbl_ValveInletDia <= 0) Then GoTo FlowCriticalLiq_Error
'dbl_Pv must be greater than zero
If dbl_Pv <= 0 Then GoTo FlowCriticalLiq_Error

'*******************************************
'Declarations
'*******************************************

Dim dbl_Ff As Double

'*******************************************
'Begin calculations
'*******************************************

dbl_Ff = 0.96 - 0.28 * Sqr(dbl_Pv / dbl_Pc)

'Calculate based on whether or not reducers exist

Select Case (dbl_InletPipeDia - dbl_ValveInletDia)
    Case 0 'no reducers
            dbl_Q = 0.865 * dbl_FL * dbl_Cv * Sqr((dbl_P1 - dbl_Ff * dbl_Pv) / dbl_Gf)
            dbl_W = 27.3 * dbl_FL * dbl_Cv * Sqr((dbl_P1 - dbl_Ff * dbl_Pv) * dbl_Gf * 1000)
    Case Is <> 0 'reducers
            dbl_Q = 0.865 * dbl_Flp * dbl_Cv * Sqr((dbl_P1 - dbl_Ff * dbl_Pv) / dbl_Gf)
            dbl_W = 27.3 * dbl_Fp * dbl_Cv * Sqr((dbl_P1 - dbl_P2) * dbl_Gf * 1000)
End Select

Exit Sub

FlowCriticalLiq_Error:

End Sub

Sub FlowFlashingLiq(dbl_P1 As Double, dbl_P2 As Double, _
    dbl_Pv As Double, dbl_Pc As Double, dbl_Gf As Double, _
    dbl_FL As Double, dbl_Fp As Double, dbl_Flp As Double, _
    dbl_InletPipeDia As Double, dbl_ValveInletDia As Double, _
    dbl_Cv As Double, dbl_Q As Double, dbl_W As Double)
'***********************************************************************
'This routine calculates the flowrates under flashing fluid conditions
'
'Input Variables:
'dbl_P1             Input Pressure 1
'dbl_P2             Input Pressure 2
'dbl_Pv             Vapor Pressure
'dbl_Pc             Critical Pressure
'dbl_Gf
'dbl_FL
'dbl_Fp
'dbl_FLP
'dbl_InletPipeDia   Diameter of Inlet Pipe
'dbl_ValveInletDia  Diameter of Inlet Valve
'dbl_Cv
'
'Output Variables
'dbl_Q
'dbl_W

'Author:  KPMG                  Date:  6/12/98
'********************************************************************
On Error GoTo FlowFlashingLiquid_Error
'*******************************************
'Check input data
'*******************************************
'dbl_Gf must be greater than zero
If dbl_Gf <= 0 Then GoTo FlowFlashingLiquid_Error
'All diameters must be greater than zero
If (dbl_InletPipeDia <= 0 Or dbl_ValveInletDia <= 0) Then GoTo FlowFlashingLiquid_Error

'*******************************************
'Declarations
'*******************************************
Dim dbl_Ff As Double

'*******************************************
'Begin calculations
'*******************************************

dbl_Ff = 0.96 - 0.28 * Sqr(dbl_Pv / dbl_Pc)

Select Case (dbl_InletPipeDia - dbl_ValveInletDia)
    Case 0 'no reducers
            dbl_Q = 0.865 * dbl_FL * dbl_Cv * Sqr((dbl_P1 - dbl_Ff * dbl_Pv) / dbl_Gf)
            dbl_W = 27.3 * dbl_FL * dbl_Cv * Sqr((dbl_P1 - dbl_Ff * dbl_Pv) * dbl_Gf * 1000)
    Case Is <> 0 'reducers
            dbl_Q = 0.865 * dbl_Flp * dbl_Cv * Sqr((dbl_P1 - dbl_Ff * dbl_Pv) / dbl_Gf)
            dbl_W = 27.3 * dbl_Flp * dbl_Cv * Sqr((dbl_P1 - dbl_Ff * dbl_Pv) * dbl_Gf * 1000)
End Select


FlowFlashingLiquid_Error:
'Placeholder
End Sub

Sub LiqCalc(str_CalculationType As String, str_ValveLine As String, _
    lng_FlowDirectionID As Long, lng_TrimTypeID As Long, lng_TrimPackageID As Long, _
    dbl_Pv As Double, dbl_P1 As Double, dbl_P2 As Double, dbl_Pc As Double, _
    dbl_T1 As Double, dbl_W As Double, dbl_Q As Double, dbl_FL As Double, _
    dbl_Fp As Double, dbl_Flp As Double, dbl_Viscosity As Double, _
    dbl_Fb As Double, int_FlowStateFlag As Integer, dbl_InletPipeDia As Double, _
    dbl_OutletPipeDia As Double, dbl_ValveInletDia As Double, _
    dbl_ValveOutletDia As Double, dbl_RatedCv As Double, _
    dbl_Cv As Double, dbl_Gf As Double, dbl_LpAe As Double, _
    dbl_InletPipeOD As Double, dbl_InletPipeWallThickness As Double, _
    dbl_OutletPipeOD As Double, dbl_OutletPipeWallThickness As Double, _
    dbl_MaxVelocity As Double, dbl_InletPipeVelocity As Double, _
    dbl_MaxVelocityDiameter As Double, dbl_PctTravel As Double, _
    str_PUnit As String, str_PUnitAbs As String, str_TUnit As String, str_LUnit As String, _
    str_WUnit As String, str_QUnit As String, str_VelocityUM As String, _
    dbl_FLc As Double)

On Error GoTo LiqCalc_Error:

#If (DebugVersion = "yes") Then
MsgBox "str_CalculationType= " + str_CalculationType + Chr(13) + Chr(10) + "str_ValveLine= " + str_ValveLine + Chr(13) + Chr(10) + "lng_FlowDirectionID= " + CStr(lng_FlowDirectionID) + Chr(13) + Chr(10) + "lng_TrimTypeID= " + CStr(lng_TrimTypeID) + Chr(13) + Chr(10) + _
"lng_TrimPackageID= " + CStr(lng_TrimPackageID) + Chr(13) + Chr(10) + "dbl_Pv= " + CStr(dbl_Pv) + Chr(13) + Chr(10) + "dbl_P1= " + CStr(dbl_P1) + Chr(13) + Chr(10) + "dbl_P2= " + CStr(dbl_P2) + Chr(13) + Chr(10) + "dbl_Pc= " + CStr(dbl_Pc) + "dbl_T1= " + CStr(dbl_T1) + Chr(13) + Chr(10) + "dbl_W= " + CStr(dbl_W) + Chr(13) + Chr(10) + "dbl_Q= " + CStr(dbl_Q) + Chr(13) + Chr(10) + _
"dbl_FL= " + CStr(dbl_FL) + Chr(13) + Chr(10) + "dbl_Flp= " + CStr(dbl_Flp) + Chr(13) + Chr(10) + "dbl_Fp= " + CStr(dbl_Fp) + _
"dbl_Viscosity= " + CStr(dbl_Viscosity) + Chr(13) + Chr(10) + "dbl_Fb= " + CStr(dbl_Fb) + Chr(13) + Chr(10) + _
"int_FlowStateFlag= " + CStr(int_FlowStateFlag) + Chr(13) + Chr(10) + "dbl_InletPipeDia= " + CStr(dbl_InletPipeDia) + Chr(13) + Chr(10) + _
"dbl_OutletPipeDia= " + CStr(dbl_OutletPipeDia) + Chr(13) + Chr(10) + "dbl_ValveInletDia= " + CStr(dbl_ValveInletDia) + Chr(13) + Chr(10) + _
"dbl_ValveOutletDia= " + CStr(dbl_ValveOutletDia) + Chr(13) + Chr(10) + "dbl_RatedCv= " + CStr(dbl_RatedCv) + Chr(13) + Chr(10) + _
"dbl_Cv= " + CStr(dbl_Cv) + Chr(13) + Chr(10) + "dbl_Gf= " + CStr(dbl_Gf) + Chr(13) + Chr(10) + _
"dbl_Lpae= " + CStr(dbl_LpAe) + Chr(13) + Chr(10) + "dbl_InletPipeOD= " + CStr(dbl_InletPipeOD) + Chr(13) + Chr(10) + _
"dbl_InletPipeWallThickness= " + CStr(dbl_InletPipeWallThickness) + Chr(13) + Chr(10) + "dbl_OutletPipeOD= " + CStr(dbl_OutletPipeOD) + Chr(13) + Chr(10) + _
"dbl_OutletPipeWallThickness= " + CStr(dbl_OutletPipeWallThickness) + Chr(13) + Chr(10) + "dbl_MaxVelocity= " + CStr(dbl_MaxVelocity) + Chr(13) + Chr(10) + _
"dbl_InletPipeVelocity= " + CStr(dbl_InletPipeVelocity) + Chr(13) + Chr(10) + "dbl_MaxVelocityDiameter= " + CStr(dbl_MaxVelocityDiameter) + Chr(13) + Chr(10) + _
"dbl_PctTravel= " + CStr(dbl_PctTravel) + Chr(13) + Chr(10) + "str_PUnit= " + str_PUnit + Chr(13) + Chr(10) + _
"str_TUnit= " + str_TUnit + Chr(13) + Chr(10) + "str_LUnit= " + str_LUnit + Chr(13) + Chr(10) + "str_WUnit= " + str_WUnit + Chr(13) + Chr(10) + _
"str_QUnit= " + str_QUnit + Chr(13) + Chr(10) + "str_VelocityUM= " + str_VelocityUM
#End If


'**********************************************************
'Declarations
'**********************************************************
Static dbl_PreliminaryCv As Double
Dim dbl_Cv1 As Double
Dim dbl_Cv2 As Double
Dim int_NumIterations As Integer
Dim dbl_Ksum As Double
Dim dbl_Ki As Double
Dim dbl_Xfz As Double
Dim dbl_DeltaLfc As Double

'Declare variables to hold values converted to standard units
Dim dbl_PvSU As Double
Dim dbl_P1SU As Double
Dim dbl_P2SU As Double
Dim dbl_PcSU As Double
Dim dbl_T1SU As Double
Dim dbl_WSU As Double
Dim dbl_QSU As Double
Dim dbl_InletPipeDiaSU As Double
Dim dbl_OutletPipeDiaSU As Double
Dim dbl_OutletPipeWallThicknessSU As Double
Dim dbl_InletPipeWallThicknessSU As Double
Dim dbl_ValveInletDiaSU As Double
Dim dbl_ValveOutletDiaSU As Double
Dim dbl_MaxVelocitySU As Double
Dim dbl_InletPipeVelocitySU As Double
Dim dbl_MaxVelocityDiameterSU As Double
Dim dbl_InletPipeODSU As Double
Dim dbl_OutletPipeODSU As Double

Dim bln_ReducedArea As Boolean

'**********************************************************
'Begin calculations
'**********************************************************
'look up the reduced area flag
If (str_ValveLine = "") Then
    bln_ReducedArea = False
Else
    Call GetReducedFlag(lng_TrimPackageID, bln_ReducedArea)
End If
    

'Convert input values to standard units
Call convertPressureToBarA(dbl_Pv, str_PUnitAbs, dbl_PvSU)
Call convertPressureToBarA(dbl_P1, str_PUnit, dbl_P1SU)
Call convertPressureToBarA(dbl_P2, str_PUnit, dbl_P2SU)
Call convertPressureToBarA(dbl_Pc, str_PUnitAbs, dbl_PcSU)
Call convertTemperatureToDegK(dbl_T1, str_TUnit, dbl_T1SU)
Call convertFlowRateToKgh(dbl_W, str_WUnit, dbl_Gf * 1000, dbl_WSU)
Call convertFlowRatetoM3H(dbl_Q, str_QUnit, dbl_QSU)

dbl_InletPipeDiaSU = dbl_InletPipeDia
dbl_OutletPipeDiaSU = dbl_OutletPipeDia
dbl_ValveInletDiaSU = dbl_ValveInletDia
dbl_ValveOutletDiaSU = dbl_ValveOutletDia

If (dbl_ValveInletDiaSU = 0) Then dbl_ValveInletDiaSU = dbl_InletPipeDiaSU
If (dbl_ValveOutletDiaSU = 0) Then dbl_ValveOutletDiaSU = dbl_OutletPipeDiaSU

Call convertLengthToMM(dbl_InletPipeOD, str_LUnit, dbl_InletPipeODSU)
Call convertLengthToMM(dbl_OutletPipeOD, str_LUnit, dbl_OutletPipeODSU)
Call convertLengthToMM(dbl_InletPipeWallThickness, str_LUnit, dbl_InletPipeWallThicknessSU)
Call convertLengthToMM(dbl_OutletPipeWallThickness, str_LUnit, dbl_OutletPipeWallThicknessSU)
Call convertVelocityToMetric(dbl_MaxVelocity, str_VelocityUM, dbl_MaxVelocitySU)
'Determine calculation type. If no model number is specified,
'do preliminary calculations

If (str_CalculationType = "Flow") Then
    Call FlowRateSizing(str_ValveLine, lng_FlowDirectionID, _
        lng_TrimTypeID, bln_ReducedArea, dbl_Fb, int_FlowStateFlag, _
        dbl_P1SU, dbl_P2SU, dbl_PvSU, dbl_PcSU, dbl_Gf, _
        dbl_FL, dbl_Fp, dbl_Flp, dbl_QSU, dbl_WSU, _
        dbl_InletPipeDiaSU, dbl_OutletPipeDiaSU, dbl_InletPipeODSU, _
        dbl_OutletPipeODSU, dbl_InletPipeWallThicknessSU, _
        dbl_OutletPipeWallThicknessSU, dbl_ValveInletDiaSU, _
        dbl_ValveOutletDiaSU, dbl_MaxVelocitySU, dbl_Cv, dbl_RatedCv, _
        dbl_LpAe, dbl_InletPipeVelocitySU, dbl_MaxVelocityDiameterSU, _
        dbl_PctTravel)
'-----------------------------------------------------------------
'Added 12/2/98
    Call CriticalFL(dbl_P1SU, dbl_P2SU, dbl_PvSU, dbl_PcSU, dbl_FLc)
'---------------------------------------------------------------

    'convert back to input units
    Call convertFlowRateFromM3H(dbl_QSU, str_QUnit, dbl_Q)
    Call convertFlowRateFromKgh(dbl_WSU, str_WUnit, dbl_Gf * 1000, dbl_W)
    Call convertVelocityFromMetric(dbl_InletPipeVelocitySU, str_VelocityUM, dbl_InletPipeVelocity)
    Call convertLengthFromMM(dbl_MaxVelocityDiameterSU, str_LUnit, dbl_MaxVelocityDiameter)
    dbl_PctTravel = TrimDecimalPlaces(dbl_PctTravel)
    dbl_LpAe = TrimDecimalPlaces(dbl_LpAe)
    dbl_Q = TrimDecimalPlaces(dbl_Q)
    dbl_W = TrimDecimalPlaces(dbl_W)
    dbl_InletPipeVelocity = TrimDecimalPlaces(dbl_InletPipeVelocity)
    dbl_MaxVelocityDiameter = TrimDecimalPlaces(dbl_MaxVelocityDiameter)
    dbl_FLc = TrimDecimalPlaces(dbl_FLc)
    dbl_FL = TrimDecimalPlaces(dbl_FL)
    GoTo Leave
End If
    
    dbl_FL = 0.9
    dbl_Fp = 1
    dbl_Flp = dbl_FL
        
    Call PipeGeom(dbl_InletPipeDiaSU, dbl_OutletPipeDiaSU, dbl_ValveInletDiaSU, _
        dbl_ValveOutletDiaSU, dbl_Ksum, dbl_Ki)

    Call LiquidCv("", lng_FlowDirectionID, lng_TrimTypeID, _
        bln_ReducedArea, int_FlowStateFlag, dbl_P1SU, dbl_P2SU, _
        dbl_PvSU, dbl_PcSU, dbl_Gf, dbl_FL, dbl_Fp, dbl_Flp, dbl_QSU, _
        dbl_WSU, dbl_Viscosity, dbl_InletPipeDiaSU, _
        dbl_OutletPipeDiaSU, dbl_ValveInletDiaSU, dbl_ValveOutletDiaSU, _
        dbl_Ki, dbl_Ksum, dbl_Xfz, dbl_RatedCv, dbl_PctTravel, dbl_PreliminaryCv)

' for the preliminary calculation
If (str_ValveLine = "") Then
    dbl_Cv = dbl_PreliminaryCv
    'Calculate final velocity
    Call LiquidVelocity(dbl_QSU, dbl_WSU, dbl_Gf, _
        dbl_InletPipeODSU, dbl_InletPipeWallThicknessSU, _
        dbl_MaxVelocitySU, dbl_InletPipeVelocitySU, dbl_MaxVelocityDiameterSU)
'-----------------------------------------------------------------
'Added 12/2/98
    Call CriticalFL(dbl_P1, dbl_P2, dbl_Pv, dbl_Pc, dbl_FLc)
'---------------------------------------------------------------
    
    'Convert output to units expected by calling program
    Call convertVelocityFromMetric(dbl_InletPipeVelocitySU, str_VelocityUM, dbl_InletPipeVelocity)
    Call convertLengthFromMM(dbl_MaxVelocityDiameterSU, str_LUnit, dbl_MaxVelocityDiameter)
    'Exit Sub

'for real calculation
Else

    dbl_Cv1 = dbl_PreliminaryCv
    int_NumIterations = 0
    Do
        Call GetFL(str_ValveLine, lng_FlowDirectionID, lng_TrimTypeID, _
            bln_ReducedArea, dbl_RatedCv, dbl_Cv1, dbl_FL)
        
        Call CalcFp(dbl_ValveInletDiaSU, dbl_Ksum, dbl_Ki, dbl_Cv1, _
            dbl_FL, dbl_Fp, dbl_Flp)
        
        Call GetXFZ(str_ValveLine, lng_FlowDirectionID, _
            lng_TrimTypeID, bln_ReducedArea, dbl_RatedCv, dbl_Cv1, _
            dbl_Xfz, dbl_DeltaLfc)
        
        Call LiquidCv(str_ValveLine, lng_FlowDirectionID, _
            lng_TrimTypeID, bln_ReducedArea, int_FlowStateFlag, _
            dbl_P1SU, dbl_P2SU, dbl_PvSU, dbl_PcSU, dbl_Gf, dbl_FL, _
            dbl_Fp, dbl_Flp, dbl_QSU, dbl_WSU, dbl_Viscosity, _
            dbl_InletPipeDiaSU, dbl_OutletPipeDiaSU, dbl_ValveInletDiaSU, _
            dbl_ValveOutletDiaSU, dbl_Ki, dbl_Ksum, dbl_Xfz, _
            dbl_RatedCv, dbl_PctTravel, dbl_Cv2)
        
        If (Abs(dbl_Cv1 - dbl_Cv2) / dbl_Cv2) * 100 < 2 Then Exit Do
        dbl_Cv1 = dbl_Cv2
        int_NumIterations = int_NumIterations + 1
    Loop While (int_NumIterations < 10)
    dbl_Cv = dbl_Cv2
    'Calculate final noise prediction if int_FlowStateFlag is 1,2 or 4
    If int_FlowStateFlag = 1 Or int_FlowStateFlag = 2 Or int_FlowStateFlag = 4 Then
        Call LiquidNoise(str_ValveLine, lng_FlowDirectionID, _
            lng_TrimTypeID, bln_ReducedArea, dbl_RatedCv, dbl_Cv, _
            dbl_Fb, dbl_P1SU, dbl_P2SU, dbl_PvSU, dbl_PcSU, dbl_QSU, _
            dbl_WSU, dbl_Gf, dbl_OutletPipeODSU, _
            dbl_OutletPipeWallThicknessSU, dbl_FL, dbl_LpAe)
    End If
    'Calculate final velocity
    Call LiquidVelocity(dbl_QSU, dbl_WSU, dbl_Gf, dbl_InletPipeODSU, _
        dbl_InletPipeWallThicknessSU, dbl_MaxVelocitySU, _
        dbl_InletPipeVelocitySU, dbl_MaxVelocityDiameterSU)
'-----------------------------------------------------------------
'Added 12/2/98
    Call CriticalFL(dbl_P1SU, dbl_P2SU, dbl_PvSU, dbl_PcSU, dbl_FLc)
'---------------------------------------------------------------
    
    'Convert output to units expected by calling program
    Call convertVelocityFromMetric(dbl_InletPipeVelocitySU, str_VelocityUM, dbl_InletPipeVelocity)
    Call convertLengthFromMM(dbl_MaxVelocityDiameterSU, str_LUnit, dbl_MaxVelocityDiameter)
    
    'Calculate % travel
    Call GetTravel(str_ValveLine, lng_TrimTypeID, _
        dbl_RatedCv, dbl_Cv, dbl_PctTravel)
End If

Leave:
#If (DebugVersion = "yes") Then
MsgBox "str_CalculationType= " + str_CalculationType + Chr(13) + Chr(10) + "str_ValveLine= " + str_ValveLine + Chr(13) + Chr(10) + "lng_FlowDirectionID= " + CStr(lng_FlowDirectionID) + Chr(13) + Chr(10) + "lng_TrimTypeID= " + CStr(lng_TrimTypeID) + Chr(13) + Chr(10) + _
"bln_ReducedArea= " + CStr(bln_ReducedArea) + Chr(13) + Chr(10) + "dbl_Pv= " + CStr(dbl_Pv) + Chr(13) + Chr(10) + "dbl_P1= " + CStr(dbl_P1) + Chr(13) + Chr(10) + "dbl_P2= " + CStr(dbl_P2) + Chr(13) + Chr(10) + "dbl_Pc= " + CStr(dbl_Pc) + "dbl_T1= " + CStr(dbl_T1) + Chr(13) + Chr(10) + "dbl_W= " + CStr(dbl_W) + Chr(13) + Chr(10) + "dbl_Q= " + CStr(dbl_Q) + Chr(13) + Chr(10) + _
"dbl_FL= " + CStr(dbl_FL) + Chr(13) + Chr(10) + "dbl_Flp= " + CStr(dbl_Flp) + Chr(13) + Chr(10) + "dbl_Fp= " + CStr(dbl_Fp) + _
"dbl_Viscosity= " + CStr(dbl_Viscosity) + Chr(13) + Chr(10) + "dbl_Fb= " + CStr(dbl_Fb) + Chr(13) + Chr(10) + _
"int_FlowStateFlag= " + CStr(int_FlowStateFlag) + Chr(13) + Chr(10) + "dbl_InletPipeDia= " + CStr(dbl_InletPipeDia) + Chr(13) + Chr(10) + _
"dbl_OutletPipeDia= " + CStr(dbl_OutletPipeDia) + Chr(13) + Chr(10) + "dbl_ValveInletDia= " + CStr(dbl_ValveInletDia) + Chr(13) + Chr(10) + _
"dbl_ValveOutletDia= " + CStr(dbl_ValveOutletDia) + Chr(13) + Chr(10) + "dbl_RatedCv= " + CStr(dbl_RatedCv) + Chr(13) + Chr(10) + _
"dbl_Cv= " + CStr(dbl_Cv) + Chr(13) + Chr(10) + "dbl_Gf= " + CStr(dbl_Gf) + Chr(13) + Chr(10) + _
"dbl_Lpae= " + CStr(dbl_LpAe) + Chr(13) + Chr(10) + "dbl_InletPipeOD= " + CStr(dbl_InletPipeOD) + Chr(13) + Chr(10) + _
"dbl_InletPipeWallThickness= " + CStr(dbl_InletPipeWallThickness) + Chr(13) + Chr(10) + "dbl_OutletPipeOD= " + CStr(dbl_OutletPipeOD) + Chr(13) + Chr(10) + _
"dbl_OutletPipeWallThickness= " + CStr(dbl_OutletPipeWallThickness) + Chr(13) + Chr(10) + "dbl_MaxVelocity= " + CStr(dbl_MaxVelocity) + Chr(13) + Chr(10) + _
"dbl_InletPipeVelocity= " + CStr(dbl_InletPipeVelocity) + Chr(13) + Chr(10) + "dbl_MaxVelocityDiameter= " + CStr(dbl_MaxVelocityDiameter) + Chr(13) + Chr(10) + _
"dbl_PctTravel= " + CStr(dbl_PctTravel) + Chr(13) + Chr(10) + "str_PUnit= " + str_PUnit + Chr(13) + Chr(10) + _
"str_TUnit= " + str_TUnit + Chr(13) + Chr(10) + "str_LUnit= " + str_LUnit + Chr(13) + Chr(10) + "str_WUnit= " + str_WUnit + Chr(13) + Chr(10) + _
"str_QUnit= " + str_QUnit + Chr(13) + Chr(10) + "str_VelocityUM= " + str_VelocityUM
#End If

    dbl_PctTravel = TrimDecimalPlaces(dbl_PctTravel)
    dbl_LpAe = TrimDecimalPlaces(dbl_LpAe)
    dbl_Cv = TrimDecimalPlaces(dbl_Cv)
    dbl_InletPipeVelocity = TrimDecimalPlaces(dbl_InletPipeVelocity)
    dbl_MaxVelocityDiameter = TrimDecimalPlaces(dbl_MaxVelocityDiameter)
    dbl_FLc = TrimDecimalPlaces(dbl_FLc)
    dbl_FL = TrimDecimalPlaces(dbl_FL)

 Exit Sub
LiqCalc_Error:
'Placeholder
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' END OF liquid
''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INCLUDE Two Phase and Steam routines
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub CriticalOutletPressure(dbl_P1 As Double, dbl_P2 As Double, dbl_Pv As Double, _
    dbl_Pc As Double, dbl_FL As Double, dbl_P2c As Double)
'**********************************************************************
'This routine calculates the critical outlet pressure
'Input Variables:
'dbl_P1                 Inlet pressure                 bar a
'dbl_P2                 Outlet pressure                bar a
'dbl_Pv                 Vapor pressure                 bar a
'dbl_Pc                 Critical pressure              bar a
'dbl_FL                 Liquid pressure recovery factor

'Output Variables:
'dbl_P2c                Critical outlet pressure       bar a
'**********************************************************************
On Error GoTo CriticalOutletPressure_Error

'*****************************************
'Check input variables
'*****************************************
'All input variables must be positive
If (dbl_P1 < 0 Or dbl_P2 < 0 Or dbl_Pv < 0 Or dbl_Pc < 0 Or dbl_FL < 0) _
Then GoTo CriticalOutletPressure_Error

'*****************************************
'Begin calculation
'*****************************************
dbl_P2c = dbl_P1 - dbl_FL ^ 2 * (dbl_P1 - (0.96 - 0.28 * Sqr(dbl_Pv / dbl_Pc) * dbl_Pv))

Exit Sub

CriticalOutletPressure_Error:
'Placeholder

End Sub


Public Sub CriticalFL(dbl_P1 As Double, dbl_P2 As Double, _
    dbl_Pv As Double, dbl_Pc As Double, dbl_FLc As Double)

'-------------------------------------------------------------------
'This function calculates the critical FL.
'-------------------------------------------------------------------
'Reference: Technical Data Rev 1, Section 1:
'           Liquid Sizing, Part 3, Page 18, April 20, 1994
'-------------------------------------------------------------------
'Input Parameters:
'dbl_P1     Inlet Pressure            bar(absolute)
'dbl_P2     Outlet Pressure           bar(absolute)
'dbl_Pv     Vapor Pressure            bar(absolute)
'dbl_Pc     Critical Pressure         bar(absolute)

'Output Parameters:
'dbl_FLc    Critical (FL) liquid pressure recovery factor
'-------------------------------------------------------------------
'Coded by:  F.Krieger
'Date:      12/2/98
'-------------------------------------------------------------------
On Error GoTo CriticalFL_Error

'*********************************************
'Check input parameters
'*********************************************

'All the input pressures must be positive
If (dbl_P1 < 0 Or dbl_P2 < 0 Or dbl_Pv < 0 Or dbl_Pc < 0) Then GoTo CriticalFL_Error:

'*********************************************
'Begin calculations
'*********************************************
If (dbl_Pv >= dbl_P2) Then
    dbl_FLc = 0
Else
    dbl_FLc = Sqr((dbl_P1 - dbl_P2) / (dbl_P1 - (0.96 - 0.28 * Sqr(dbl_Pv / dbl_Pc)) * dbl_Pv))
End If
Exit Sub

CriticalFL_Error:
'Placeholder
 
End Sub

Sub CvWithFactor(dbl_C1, dbl_Cv, dbl_Cv1)
'**********************************************************************
'This routine mutiplies dbl_Cv and dbl_C1
'Input          Output      Description
'dbl_C1                     Multiplier
'dbl_Cv                     Required Cv
'               dbl_Cv1     dbl_Cv with multiplier
'**********************************************************************
On Error GoTo CvWithFactor_Error

'***************************************
'Check input variables
'***************************************
'dbl_C1 must be greater than 0
If dbl_C1 <= 0 Then GoTo CvWithFactor_Error

'***************************************
'Begin calculation
'***************************************
dbl_Cv1 = dbl_Cv * dbl_C1
Exit Sub

CvWithFactor_Error:
'Placeholder
End Sub

Function BL(dbl_Tr As Double)
'--------------------------------------------------------------------------------------
'Source: Properties of Water and Steam in SI Units, 2nd edition,
'Springer -Verlag, Berlin, 1979
'---------------------------------------------------------------------------------------
'This function is used to return intermediate values within various steam property subroutines.
'----------------------------------------------------------------------------------------
'Input       Output
'dbl_Tr
'           BL (dbl_Tr)
'
'Coded By:
'Date:
'-----------------------------------------------------------------------------------------

On Error GoTo BL_Error

'**********************************
'Check input validation
'**********************************
'
'**********************************
'Declarations
'**********************************

      Dim dbl_L(3) As Double
      Dim Tr As Double
      
      dbl_L(1) = 15.74373327
      dbl_L(2) = -34.17061978
      dbl_L(3) = 19.31380707
      
'**********************************
'Begin Calculation
'**********************************

     BL = dbl_L(1) + dbl_L(2) * dbl_Tr + dbl_L(3) * dbl_Tr * dbl_Tr
Exit Function
BL_Error:
'Placeholder

End Function


Function DBL(dbl_Tr As Double)
'--------------------------------------------------------------------------------------
'Source: Properties of Water and Steam in SI Units, 2nd edition,
'Springer -Verlag, Berlin, 1979
'
'---------------------------------------------------------------------------------------
'This function is used to return intermediate values within various steam property subroutines.
'
'----------------------------------------------------------------------------------------
'Input Variables:
'dbl_Tr

'Output Variables:
'DBL (dbl_Tr)
'
'Coded By:
'Date:
'-----------------------------------------------------------------------------------------

On Error GoTo DBL_Error

'**********************************
'Check input
'**********************************
'
'
'**********************************
'Declarations
'**********************************

     Dim dbl_L(2) As Double
      
     dbl_L(1) = -34.17061978
     dbl_L(2) = 19.31380707
      
'**********************************
'Begin Calculation
'**********************************

    DBL = dbl_L(1) + 2 * dbl_L(2) * dbl_Tr
    
Exit Function
DBL_Error:
'Placeholder

End Function


Function EnthalpyLiquid(dbl_T As Double, dbl_P As Double)
'--------------------------------------------------------------------------------------
'Source: Properties of Water and Steam in SI Units, 2nd edition,
'Springer -Verlag, Berlin, 1979
'
'--------------------------------------------------------------------------------------
'  This function returns the enthalpy of water in kJ/kg.
'
'   Temperature range   273.16 K < T < 623.15 K
'                       32 F < T < 752 F
'
'   Pressure range      P < 1000  bar a
'                       P < 14500 psi a
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_T              Temperature                 deg K
'dbl_P              Pressure                    bar a

'Output Variables:
'EnthalpyWater  Enthalpy of liquid          kJ/kg
'Coded By:
'Date:
'--------------------------------------------------------------------------------------
On Error GoTo EnthalpyLiquid_Error

'**********************************
'Check input validation
'**********************************
'
'
'**********************************
'Declarations
'**********************************
            
      Dim dbl_AK(12) As Double
      Dim dbl_AG(12) As Double
      Dim dbl_AGA(11) As Double
      Dim dbl_Tr As Double
      Dim dbl_Pr As Double
      Dim dbl_Y As Double
      Dim dbl_Y1 As Double
      Dim dbl_Z As Double
      Dim dbl_H1 As Double
      Dim dbl_H2 As Double
      Dim dbl_H3 As Double
      Dim dbl_H4 As Double
      Dim dbl_H5 As Double
      Dim dbl_H6 As Double
      Dim dbl_H7 As Double
      Dim dbl_H8 As Double
      Dim dbl_HR As Double
      Dim dbl_Tc As Double
      Dim dbl_Pc As Double
      Dim dbl_Hc As Double
      Dim I As Integer
      Dim I1 As Integer
            
'****************************************************
'Begin calculations
'****************************************************
      
      dbl_Tc = 647.3
      dbl_Pc = 221.2
      dbl_Hc = 70.1204
           
      dbl_AK(1) = 0.8438375405
      dbl_AK(2) = 0.0005362162162
      dbl_AK(3) = 1.72
      dbl_AK(4) = 0.07342278489
      dbl_AK(5) = 0.0497585887
      dbl_AK(6) = 0.65371543
      dbl_AK(7) = 0.00000115
      dbl_AK(8) = 0.000015108
      dbl_AK(9) = 0.14188
      dbl_AK(10) = 7.002753165
      dbl_AK(11) = 0.0002995284926
      dbl_AK(12) = 0.204
      
      dbl_AG(1) = 7.982692717
      dbl_AG(2) = -0.02616571843
      dbl_AG(3) = 0.00152241179
      dbl_AG(4) = 0.02284279054
      dbl_AG(5) = 242.1647003
      dbl_AG(6) = 1.269716088E-10
      dbl_AG(7) = 2.074838328E-07
      dbl_AG(8) = 2.17402035E-08
      dbl_AG(9) = 1.105710498E-09
      dbl_AG(10) = 12.93441934
      dbl_AG(11) = 0.00001308119072
      dbl_AG(12) = 6.047626338E-14
      
      dbl_AGA(1) = 6824.687741
      dbl_AGA(2) = -542.2063673
      dbl_AGA(3) = -20966.66205
      dbl_AGA(4) = 39412.86787
      dbl_AGA(5) = -67332.77739
      dbl_AGA(6) = 99023.81028
      dbl_AGA(7) = -109391.1774
      dbl_AGA(8) = 85908.41667
      dbl_AGA(9) = -45111.68742
      dbl_AGA(10) = 14181.38926
      dbl_AGA(11) = -2017.271113

      dbl_Tr = dbl_T / dbl_Tc
      dbl_Pr = dbl_P / dbl_Pc

      dbl_Y = 1 - dbl_AK(1) * dbl_Tr * dbl_Tr - dbl_AK(2) / dbl_Tr ^ 6
      dbl_Y1 = 6 * dbl_AK(2) / dbl_Tr ^ 7 - 2 * dbl_AK(1) * dbl_Tr
      dbl_Z = dbl_Y + Sqr(dbl_AK(3) * dbl_Y * dbl_Y - 2 * dbl_AK(4) * dbl_Tr + 2 * dbl_AK(5) * dbl_Pr)

      dbl_H1 = dbl_AGA(1) * dbl_Tr
      dbl_H2 = 0
      For I = 1 To 10
        I1 = I + 1
        dbl_H2 = dbl_H2 + (I - 2) * dbl_AGA(I1) * dbl_Tr ^ (I - 1)
      Next I
      dbl_H3 = dbl_AG(1) * (dbl_Z * (17 * (dbl_Z / 29 - dbl_Y / 12) + 5 * dbl_Tr * dbl_Y1 / 12) + dbl_AK(4) * dbl_Tr - (dbl_AK(3) - 1) * dbl_Tr * dbl_Y * dbl_Y1) / dbl_Z ^ (5 / 17)
      dbl_H4 = (dbl_AG(2) - dbl_AG(4) * dbl_Tr * dbl_Tr + dbl_AG(5) * (9 * dbl_Tr + dbl_AK(6)) * (dbl_AK(6) - dbl_Tr) ^ 9 + dbl_AG(6) * (20 * dbl_Tr ^ 19 + dbl_AK(7)) / (dbl_AK(7) + dbl_Tr ^ 19) ^ 2) * dbl_Pr
      dbl_H5 = (12 * dbl_Tr ^ 11 + dbl_AK(8)) / (dbl_AK(8) + dbl_Tr ^ 11) ^ 2 * (dbl_AG(7) * dbl_Pr + dbl_AG(8) * dbl_Pr * dbl_Pr + dbl_AG(9) * dbl_Pr ^ 3)
      dbl_H6 = dbl_AG(10) * dbl_Tr ^ 18 * (17# * dbl_AK(9) + 19# * dbl_Tr * dbl_Tr) * ((dbl_AK(10) + dbl_Pr) ^ (-3) + dbl_AK(11) * dbl_Pr)
      dbl_H7 = dbl_AG(11) * dbl_AK(12) * dbl_Pr ^ 3
      dbl_H8 = 21 * dbl_AG(12) * dbl_Pr ^ 4 / dbl_Tr ^ 20
      dbl_HR = dbl_H1 - dbl_H2 + dbl_H3 + dbl_H4 - dbl_H5 + dbl_H6 + dbl_H7 + dbl_H8
      EnthalpyLiquid = dbl_HR * dbl_Hc
      
EnthalpyLiquid_Error:
'Placeholder

End Function

Function EnthalpyVapor(dbl_T As Double, dbl_P As Double)
'--------------------------------------------------------------------------------------
'Source: Properties of Water and Steam in SI Units, 2nd edition,
'Springer -Verlag, Berlin, 1979
'--------------------------------------------------------------------------------------
' This function returns the enthalpy of steam in kJ/kg.
'
'  Temperature range  273.16 K < T < 1073.15 K
'                     32 F < T <  1472 F
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_T              Temperature                 deg K
'dbl_P              Pressure                    bar a
'--------------------------------------------------------------------------------------
'Output Variables:
'EnthalpySteam      Enthalpy of liquid          kJ/kg
'Coded By:
'Date:
'--------------------------------------------------------------------------------------
On Error GoTo EnthalpyVapor_Error

'**********************************
'Check input validation
'**********************************
'
'
'**********************************
'Declarations
'**********************************

      Dim dbl_NA(8) As Double
      Dim dbl_NZ(8, 3) As Double
      Dim dbl_NL(3) As Double
      Dim dbl_NX(3, 2) As Double
      Dim dbl_BGA(6) As Double
      Dim dbl_BG(8, 3) As Double
      Dim dbl_BK(3, 2) As Double
      Dim dbl_B9(7) As Double
      Dim I As Integer
      Dim IA As Integer
      Dim IB As Integer
      Dim IC As Integer
      Dim dbl_Tr As Double
      Dim dbl_Pr As Double
      Dim dbl_X As Double
      Dim dbl_BKA As Double
      Dim dbl_PRL  As Double
      Dim dbl_DPRL  As Double
      Dim dbl_BL As Double
      Dim dbl_DBL As Double
      Dim dbl_H1 As Double
      Dim dbl_H2 As Double
      Dim dbl_H3 As Double
      Dim dbl_H3A As Double
      Dim dbl_NAA As Double
      Dim dbl_H4 As Double
      Dim dbl_H4A As Double
      Dim dbl_H4B As Double
      Dim dbl_LAM As Double
      Dim dbl_NAB As Double
      Dim dbl_H4C As Double
      Dim dbl_HIL1 As Double
      Dim dbl_NAC As Double
      Dim dbl_H5 As Double
      Dim dbl_H5A As Double
      Dim dbl_HR As Double
            
      dbl_NA(1) = 2
      dbl_NA(2) = 3
      dbl_NA(3) = 2
      dbl_NA(4) = 2
      dbl_NA(5) = 3
      dbl_NA(6) = 2
      dbl_NA(7) = 2
      dbl_NA(8) = 2
      
      dbl_NZ(1, 1) = 13
      dbl_NZ(2, 1) = 18
      dbl_NZ(3, 1) = 18
      dbl_NZ(4, 1) = 25
      dbl_NZ(5, 1) = 32
      dbl_NZ(6, 1) = 12
      dbl_NZ(7, 1) = 24
      dbl_NZ(8, 1) = 24
      dbl_NZ(1, 2) = 3
      dbl_NZ(2, 2) = 2
      dbl_NZ(3, 2) = 10
      dbl_NZ(4, 2) = 14
      dbl_NZ(5, 2) = 28
      dbl_NZ(6, 2) = 11
      dbl_NZ(7, 2) = 18
      dbl_NZ(8, 2) = 14
      dbl_NZ(1, 3) = 0
      dbl_NZ(2, 3) = 1
      dbl_NZ(3, 3) = 0
      dbl_NZ(4, 3) = 0
      dbl_NZ(5, 3) = 24
      dbl_NZ(6, 3) = 0
      dbl_NZ(7, 3) = 0
      dbl_NZ(8, 3) = 0
      
      dbl_NL(1) = 1
      dbl_NL(2) = 1
      dbl_NL(3) = 2
      
      dbl_NX(1, 1) = 14
      dbl_NX(2, 1) = 19
      dbl_NX(3, 1) = 54
      dbl_NX(1, 2) = 0
      dbl_NX(2, 2) = 0
      dbl_NX(3, 2) = 27
      
      dbl_BGA(1) = 16.83599274
      dbl_BGA(2) = 28.56067796
      dbl_BGA(3) = -54.38923329
      dbl_BGA(4) = 0.4330662834
      dbl_BGA(5) = -0.6547711697
      dbl_BGA(6) = 0.08565182058
      
      dbl_BG(1, 1) = 0.06670375918
      dbl_BG(2, 1) = 0.08390104328
      dbl_BG(3, 1) = 0.4520918904
      dbl_BG(4, 1) = -0.5975336707
      dbl_BG(5, 1) = 0.5958051609
      dbl_BG(6, 1) = 0.1190610271
      dbl_BG(7, 1) = 0.1683998803
      dbl_BG(8, 1) = 0.006552390126
      dbl_BG(1, 2) = 1.388983801
      dbl_BG(2, 2) = 0.02614670893
      dbl_BG(3, 2) = 0.1069036614
      dbl_BG(4, 2) = -0.08847535804
      dbl_BG(5, 2) = -0.5159303373
      dbl_BG(6, 2) = -0.09867174132
      dbl_BG(7, 2) = -0.05809438001
      dbl_BG(8, 2) = 0.0005710218649
      dbl_BG(1, 3) = 0
      dbl_BG(2, 3) = -0.03373439453
      dbl_BG(3, 3) = 0
      dbl_BG(4, 3) = 0
      dbl_BG(5, 3) = 0.2075021122
      dbl_BG(6, 3) = 0
      dbl_BG(7, 3) = 0
      dbl_BG(8, 3) = 0
      
      dbl_BK(1, 1) = 0.4006073948
      dbl_BK(2, 1) = 0.08636081627
      dbl_BK(3, 1) = -0.8532322921
      dbl_BK(1, 2) = 0
      dbl_BK(2, 2) = 0
      dbl_BK(3, 2) = 0.3460208861
      
      dbl_B9(1) = 193.6587558
      dbl_B9(2) = -1388.522425
      dbl_B9(3) = 4126.607219
      dbl_B9(4) = -6508.211677
      dbl_B9(5) = 5745.984054
      dbl_B9(6) = -2693.088365
      dbl_B9(7) = 523.5718623
      
      dbl_BKA = 0.7633333333
      
'**********************************
'Begin Calculation
'**********************************

      dbl_Tr = dbl_T / dbl_Tc
      dbl_Pr = dbl_P / dbl_Pc

      dbl_X = Exp(dbl_BKA * (1 - dbl_Tr))
      dbl_PRL = BL(dbl_Tr)
      dbl_DPRL = DBL(dbl_Tr)

      dbl_H1 = dbl_BGA(1) * dbl_Tr
      dbl_H2 = 0
      For I = 2 To 6
        dbl_H2 = dbl_H2 + dbl_BGA(I) * (I - 3) * dbl_Tr ^ (I - 2)
      Next I
      dbl_H3 = 0
      For I = 1 To 5
        dbl_H3A = 0
        dbl_NAA = dbl_NA(I)
        For IA = 1 To dbl_NAA
            dbl_H3A = dbl_H3A + (dbl_BG(I, IA) * (1# + dbl_NZ(I, IA) * dbl_BKA * dbl_Tr)) * (dbl_X ^ dbl_NZ(I, IA))
        Next IA
        dbl_H3 = dbl_H3 + dbl_Pr ^ I * dbl_H3A
      Next I
      dbl_H4 = 0
      For I = 6 To 8
        dbl_H4A = 0
        dbl_H4B = 0
        dbl_LAM = I - 5
        dbl_NAB = dbl_NL(dbl_LAM)
        For IB = 1 To dbl_NAB
            dbl_H4B = dbl_H4B + dbl_NX(dbl_LAM, IB) * dbl_BK(dbl_LAM, IB) * dbl_X ^ dbl_NX(dbl_LAM, IB)
        Next IB
        dbl_H4C = 0
        For IC = 1 To dbl_NAB
            dbl_H4C = dbl_H4C + dbl_BK(dbl_LAM, IC) * dbl_X ^ dbl_NX(dbl_LAM, IC)
        Next IC
        dbl_HIL1 = dbl_Pr ^ (2 - I) + dbl_H4C
        dbl_NAC = dbl_NA(I)
        For IA = 1 To dbl_NAC
            dbl_H4A = dbl_H4A + dbl_BG(I, IA) * dbl_X ^ dbl_NZ(I, IA) * (1# + dbl_NZ(I, IA) * dbl_BKA * dbl_Tr - dbl_BKA * dbl_Tr * dbl_H4B / dbl_HIL1)
        Next IA
        dbl_H4 = dbl_H4 + dbl_H4A / dbl_HIL1
      Next I
      dbl_H5A = 0
      For I = 1 To 7
        dbl_H5A = dbl_H5A + (1 + dbl_Tr * (10 * dbl_DPRL / dbl_PRL + (I - 1) * dbl_BKA)) * dbl_B9(I) * dbl_X ^ (I - 1)
      Next I
      dbl_H5 = dbl_Pr * (dbl_Pr / dbl_PRL) ^ 10 * dbl_H5A
      dbl_HR = dbl_H1 - dbl_H2 - dbl_H3 - dbl_H4 + dbl_H5
      EnthalpyVapor = dbl_HR * dbl_Hc
      
EnthalpyVapor_Error:
'Placeholder

End Function

Sub OutletSpecVol(dbl_InletPressureBarA As Double, dbl_OutletPressureBarA As Double, _
    dbl_DegK As Double, str_EnthalpyUM As String, str_SpecVolUM As String, _
    dbl_OutletSpecVolume As Double, dbl_OutletEnthalpy As Double)
'--------------------------------------------------------------------------------------
'This subroutine is used to calculate the specific volume ot steam at OUTLET conditions.
'Entalpy is calculated using inlet pressure and temperature.  Then that enthalpy value
'is used to find the the temperature at outlet pressure that has approximately
'the same enthalpy.  Finally, that temperature and the outlet pressure are used
'to calculate specific volume.
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_InletPressureBarA                      bar a
'dbl_OutletPressureBarA                     bar s
'dbl_DegK                                   degrees K
'str_EnthalpyUM
'str_SpecVolUM
'--------------------------------------------------------------------------------------
'Output Variables:
'dbl_OutletSpecVolume                       kj/kg
'dbl_OutletEnthalpy                         kJ/kg

'Coded By:
'Date:
'--------------------------------------------------------------------------------------
On Error GoTo OutletSpecVol_Error

'**********************************
'Check input validation
'**********************************
'
'
'**********************************
'Declarations
'**********************************

    Dim dbl_EnthalpyInlet As Double
    Dim dbl_EnthalpyVapor As Double
    Dim dbl_TempTemperature As Double
    Dim dbl_SatTempK As Double
    Dim dbl_EnthalpyOutlet As Double
    Dim dbl_MyEnthalpyVapor As Double
    Dim dbl_LastEnthalpyOutlet As Double
    Dim dbl_MySpecVolVaporOut As Double
    Dim dbl_SpecVolVaporOut As Double
    
'**********************************
'Begin Calculation
'**********************************
  
    dbl_EnthalpyInlet = EnthalpyVapor(dbl_DegK, dbl_InletPressureBarA)
    
    dbl_TempTemperature = SatTempK(dbl_OutletPressureBarA)

    dbl_EnthalpyOutlet = EnthalpyVapor(dbl_TempTemperature, dbl_OutletPressureBarA)
    
    dbl_TempTemperature = SatTempK(dbl_OutletPressureBarA)
        
 
 ' Move toward equality in coarse steps
   
 Do Until dbl_EnthalpyOutlet > dbl_EnthalpyInlet
                       
    dbl_EnthalpyOutlet = EnthalpyVapor(dbl_TempTemperature, dbl_OutletPressureBarA)
                    
    dbl_TempTemperature = dbl_TempTemperature + 10
                               
    Call convertEnthalpyFromMetric(dbl_EnthalpyOutlet, str_EnthalpyUM, dbl_LastEnthalpyOutlet)
      
    dbl_MyEnthalpyVapor = Format(dbl_LastEnthalpyOutlet, "#####.0")
                  
   Loop
   
    'Move back
 
    dbl_TempTemperature = dbl_TempTemperature - 15
    dbl_EnthalpyOutlet = EnthalpyVapor(dbl_TempTemperature, dbl_OutletPressureBarA)
   
   
  'Start again with small increments
   
  Do Until dbl_EnthalpyOutlet > dbl_EnthalpyInlet
                       
    dbl_EnthalpyOutlet = EnthalpyVapor(dbl_TempTemperature, dbl_OutletPressureBarA)
                    
    dbl_TempTemperature = dbl_TempTemperature + 0.5
    Call convertEnthalpyFromMetric(dbl_EnthalpyOutlet, str_EnthalpyUM, dbl_LastEnthalpyOutlet)
             
     dbl_MyEnthalpyVapor = Format(dbl_LastEnthalpyOutlet, "#####.0")
                  
   Loop
      
      Call convertEnthalpyFromMetric(dbl_EnthalpyOutlet, str_EnthalpyUM, dbl_LastEnthalpyOutlet)
    
      dbl_MyEnthalpyVapor = Format(dbl_LastEnthalpyOutlet, "#####.0")
    
      dbl_OutletSpecVolume = SpecificVolumeVapor(dbl_TempTemperature, dbl_OutletPressureBarA)
    
      Call convertSpecVolFromMetric(dbl_OutletSpecVolume, str_SpecVolUM, dbl_SpecVolVaporOut)
    
      dbl_MySpecVolVaporOut = Format(dbl_SpecVolVaporOut, "#.0000")
      
      dbl_OutletEnthalpy = dbl_MyEnthalpyVapor
  
      dbl_OutletSpecVolume = dbl_MySpecVolVaporOut
     
OutletSpecVol_Error:
'Placeholder
    
End Sub

Sub RateOfFlashingWater(dbl_InletPressureBarA As Double, dbl_OutletPressureBarA As Double, _
    dbl_DegK As Double, dbl_X As Double, dbl_R As Double)
'--------------------------------------------------------------------------------------
'This subroutine calculate the rate of flashing for water, and the corresponding ratio
'of outlet volumetric flowrate and the inlet volumetric flowrate.
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_InletPressureBarA
'dbl_OutletPressureBarA
'dbl_DegK
'--------------------------------------------------------------------------------------
'Output Variables:
' dbl_X       Rate of flashing
' dbl_R       Ratio   Outlet volumetric flow rate / Inlet volumetric flowrate
'
' Coded By:
' Date:
'--------------------------------------------------------------------------------------
On Error GoTo RateOfFlashingWater_Error

'**********************************
'Check input validation
'**********************************
'
'
'**********************************
'Declarations
'**********************************

Dim dbl_InletEnthalpyLiquid As Double
Dim dbl_EnthalpyLiquid As Double
Dim dbl_SVLI As Double
Dim dbl_SpecificVolumeLiquid As Double
Dim dbl_SpecificVolumeVapor As Double
Dim dbl_InletDensityLiquid As Double
Dim dbl_OutletSaturationTemp As Double
Dim dbl_SVV2 As Double
Dim dbl_OutletDensityLiquid As Double
Dim dbl_SVL2 As Double
Dim dbl_SatTempK As Double
Dim dbl_OutletEnthalpyVapor As Double
Dim dbl_EnthalpyVapor As Double
Dim dbl_OutletEnthalpyLiquid As Double
Dim dbl_OutletDensityVapor As Double
Dim dbl_Z As Double


'**********************************
'Begin Calculation
'**********************************

dbl_InletEnthalpyLiquid = EnthalpyLiquid(dbl_DegK, dbl_InletPressureBarA)
    
dbl_SVLI = SpecificVolumeLiquid(dbl_DegK, dbl_InletPressureBarA)
dbl_InletDensityLiquid = 1 / dbl_SVLI
    
dbl_OutletSaturationTemp = SatTempK(dbl_OutletPressureBarA)
    
dbl_OutletEnthalpyVapor = EnthalpyVapor(dbl_OutletSaturationTemp, dbl_OutletPressureBarA)
dbl_OutletEnthalpyLiquid = EnthalpyLiquid(dbl_OutletSaturationTemp, dbl_OutletPressureBarA)
   
dbl_SVV2 = SpecificVolumeVapor(dbl_OutletSaturationTemp, dbl_OutletPressureBarA)
dbl_OutletDensityVapor = 1 / dbl_SVV2
    
    
dbl_SVL2 = SpecificVolumeLiquid(dbl_OutletSaturationTemp, dbl_OutletPressureBarA)
dbl_OutletDensityLiquid = 1 / dbl_SVL2
   
dbl_X = (dbl_InletEnthalpyLiquid - dbl_OutletEnthalpyLiquid) / (dbl_OutletEnthalpyVapor - dbl_OutletEnthalpyLiquid)
    
dbl_R = dbl_InletDensityLiquid * (dbl_X * dbl_OutletDensityLiquid + (1 - dbl_X) * dbl_OutletDensityVapor) / (dbl_OutletDensityLiquid * dbl_OutletDensityVapor)
   
dbl_Z = dbl_R
    
RateOfFlashingWater_Error:
'Placeholder
    
 
End Sub


Function SatTempK(dbl_P As Double)
'--------------------------------------------------------------------------------------
'Source: Properties of Water and Steam in SI Units, 2nd edition,
'Springer -Verlag, Berlin, 1979
'-------------------------------------------------------------------------------------
' This function returns the saturation temperature of water in deg K.
'
' Valid Range         0.01 bar < P < 220      bar
'                     0.145 psi a < P < 3190  psi a
'--------------------------------------------------------------------------------------
'Input Variables:
'dbl_P                  Pressure                        bar a
'--------------------------------------------------------------------------------------
'Output Variables:
'dbl_SatTempWater       Saturation Temperature          deg K
'
'Coded By:
'Date:
'-------------------------------------------------------------------------------------
On Error GoTo SatTempK_Error

'**********************************
'Check input validation
'**********************************
'
'
'**********************************
'Declarations
'**********************************

  Dim I As Integer
  Dim dbl_Tm As Double
  Dim dbl_Punt As Double
  Dim dbl_Pob As Double
  Dim dbl_Pm As Double
  Dim dbl_Dpunt As Double
  Dim dbl_Dpob As Double
  Dim dbl_Dpm As Double
  Dim dbl_VaporPressureBarA As Double
  Dim dbl_Tunt As Double
  Dim dbl_Tob As Double

    dbl_Tunt = 280.15
    dbl_Tob = 646.85
    
'**********************************
'Begin Calculation
'**********************************
 
    dbl_Tm = 0.5 * (dbl_Tunt + dbl_Tob)
    dbl_Punt = VaporPressureBarA(dbl_Tunt)
    dbl_Pob = VaporPressureBarA(dbl_Tob)
    dbl_Pm = VaporPressureBarA(dbl_Tm)
    dbl_Dpunt = dbl_Punt - dbl_P
    dbl_Dpob = dbl_Pob - dbl_P
    dbl_Dpm = dbl_Pm - dbl_P
    
    If (dbl_Dpunt * dbl_Dpob) > 0 Then
         SatTempK = -1
        Exit Function
        
    End If
    
    For I = 1 To 10000
        If (dbl_Dpm * dbl_Dpunt) < 0 Then
            dbl_Tob = dbl_Tm
            dbl_Pob = dbl_Pm
            dbl_Dpob = dbl_Dpm
        Else
            dbl_Tunt = dbl_Tm
            dbl_Punt = dbl_Pm
            dbl_Dpunt = dbl_Dpm
        End If
        
        dbl_Tm = 0.5 * (dbl_Tunt + dbl_Tob)
        SatTempK = dbl_Tm
        If (dbl_Tm - dbl_Tunt) < 0.0005 Then Exit Function
        dbl_Pm = VaporPressureBarA(dbl_Tm)
        dbl_Dpm = dbl_Pm - dbl_P
    Next I
    
SatTempK_Error:
'Placeholder

End Function

Function SpecificVolumeLiquid(dbl_T As Double, dbl_P As Double)
'--------------------------------------------------------------------------------------
'Source: Properties of Water and Steam in SI Units, 2nd edition,
'Springer -Verlag, Berlin, 1979
'---------------------------------------------------------------------------------------
'This function returns the specific volume of water in m3/kg.
'
'Temperature range              273.16 K < T < 623.15 K
'                               32 F < T < 753 F
'
'Pressure range                 P < 1000 BAR
'                               P < 14500 psi a
'----------------------------------------------------------------------------------------
'Input Variables:
'dbl_T                   Temperature                    K
'--------------------------------------------------------------------------------------
'Output Variables:
'dbl_P                   Pressure                       bar
'SpecificVolumeWater     Specific volume of liquid      m^3/kg
'
'Coded By:
'Date:
'---------------------------------------------------------------------------------------
On Error GoTo SpecificVolumeLiquid_Error

'**********************************
'Check input validation
'**********************************
'
'
'**********************************
'Declarations
'**********************************

      Dim dbl_AK(12) As Double
      Dim dbl_AG(12) As Double
      Dim dbl_Tr As Double
      Dim dbl_Pr As Double
      Dim dbl_Y As Double
      Dim dbl_Y1 As Double
      Dim dbl_Z As Double
      Dim dbl_S1 As Double
      Dim dbl_S2 As Double
      Dim dbl_S3 As Double
      Dim dbl_S4 As Double
      Dim dbl_S5 As Double
      Dim dbl_S6 As Double
      Dim dbl_VR As Double
      Dim dbl_Vc As Double
  

      dbl_AK(1) = 0.8438375405
      dbl_AK(2) = 0.0005362162162
      dbl_AK(3) = 1.72
      dbl_AK(4) = 0.07342278489
      dbl_AK(5) = 0.0497585887
      dbl_AK(6) = 0.65371543
      dbl_AK(7) = 0.00000115
      dbl_AK(8) = 0.000015108
      dbl_AK(9) = 0.14188
      dbl_AK(10) = 7.002753165
      dbl_AK(11) = 0.0002995284926
      dbl_AK(12) = 0.204
      
      dbl_AG(1) = 7.982692717
      dbl_AG(2) = -0.02616571843
      dbl_AG(3) = 0.00152241179
      dbl_AG(4) = 0.02284279054
      dbl_AG(5) = 242.1647003
      dbl_AG(6) = 1.269716088E-10
      dbl_AG(7) = 2.074838328E-07
      dbl_AG(8) = 2.17402035E-08
      dbl_AG(9) = 1.105710498E-09
      dbl_AG(10) = 12.93441934
      dbl_AG(11) = 0.00001308119072
      dbl_AG(12) = 6.047626338E-14
      
'**********************************
'Begin Calculation
'**********************************

     dbl_Tr = dbl_T / dbl_Tc
     dbl_Pr = dbl_P / dbl_Pc
      
     dbl_Y = 1 - dbl_AK(1) * dbl_Tr * dbl_Tr - dbl_AK(2) / dbl_Tr ^ 6
     dbl_Y1 = 6 * dbl_AK(2) / dbl_Tr ^ 7 - 2 * dbl_AK(1) * dbl_Tr
     dbl_Z = dbl_Y + Sqr(dbl_AK(3) * dbl_Y * dbl_Y - 2 * dbl_AK(4) * dbl_Tr + 2 * dbl_AK(5) * dbl_Pr)

     dbl_S1 = dbl_AG(1) * dbl_AK(5) / dbl_Z ^ (5 / 17)
     dbl_S2 = dbl_AG(2) + dbl_AG(3) * dbl_Tr + dbl_AG(4) * dbl_Tr * dbl_Tr + dbl_AG(5) * (dbl_AK(6) - dbl_Tr) ^ 10 + dbl_AG(6) / (dbl_AK(7) + dbl_Tr ^ 19)
     dbl_S3 = (dbl_AG(7) + 2 * dbl_AG(8) * dbl_Pr + 3# * dbl_AG(9) * dbl_Pr * dbl_Pr) / (dbl_AK(8) + dbl_Tr ^ 11)
     dbl_S4 = dbl_AG(10) * dbl_Tr ^ 18 * (dbl_AK(9) + dbl_Tr * dbl_Tr) * ((-3) / (dbl_AK(10) + dbl_Pr) ^ 4 + dbl_AK(11))
     dbl_S5 = 3 * dbl_AG(11) * (dbl_AK(12) - dbl_Tr) * dbl_Pr * dbl_Pr
     dbl_S6 = 4 * dbl_AG(12) * dbl_Pr ^ 3 / dbl_Tr ^ 20

     dbl_VR = dbl_S1 + dbl_S2 - dbl_S3 - dbl_S4 + dbl_S5 + dbl_S6
     SpecificVolumeLiquid = dbl_VR * dbl_Vc
      
SpecificVolumeLiquid_Error:
'Placeholder

End Function

Function SpecificVolumeVapor(dbl_T As Double, dbl_P As Double)
'--------------------------------------------------------------------------------------
'Source: Properties of Water and Steam in SI Units, 2nd edition,
'Springer -Verlag, Berlin, 1979
'---------------------------------------------------------------------------------------
'This function returns the specific volume of water in m3/kg.
'
'Temperature range           273.16 K < T < 1073.15 K
'                            32 F < T < 1472 F
'----------------------------------------------------------------------------------------
'Input Variables:
'dbl_T                      Temperature                 K

'Output Variables:
'dbl_P                      Pressure                    bar a
'SpecificVolumeSteam        Specific volume of vapor    m^3/kg
'
'Coded By:
'Date:
'----------------------------------------------------------------------------------

On Error GoTo SpecificVolumeVapor_Error

'**********************************
'Constants
'**********************************
Const dbl_Tc As Double = 647.3  'Critical Temperature (deg K)
Const dbl_Pc As Double = 221.2   'Critical Pressure (Bar a)
Const dbl_Vc As Double = 0.00317 'Specific Volume of Vapor (m3/kg)
'
'**********************************
'Declarations
'**********************************

      Dim dbl_NA(8) As Double
      Dim dbl_NZ(8, 3) As Double
      Dim dbl_NL(3) As Double
      Dim dbl_NX(3, 2) As Double
      Dim dbl_BG(8, 3) As Double
      Dim dbl_BK(3, 2) As Double
      Dim dbl_B9(7) As Double
      Dim I As Integer
      Dim IA As Integer
      Dim IB As Integer
      Dim dbl_Tr As Double
      Dim dbl_Pr As Double
      Dim dbl_X As Double
      Dim dbl_BKA As Double
      Dim dbl_PRL As Double
      Dim dbl_BL As Double
      Dim dbl_DPRL As Double
      Dim dbl_DBL As Double
      Dim dbl_RI1 As Double
      Dim dbl_S1 As Double
      Dim dbl_S2 As Double
      Dim dbl_S2A As Double
      Dim dbl_NAB  As Double
      Dim dbl_S3 As Double
      Dim dbl_LAM As Double
      Dim dbl_S3A As Double
      Dim dbl_S3B As Double
      Dim dbl_S3ZAE As Double
      Dim dbl_NAC As Double
      Dim dbl_S3HI As Double
      Dim dbl_S3NEN As Double
      Dim dbl_S4 As Double
      Dim dbl_VR As Double
        
      dbl_NA(1) = 2
      dbl_NA(2) = 3
      dbl_NA(3) = 2
      dbl_NA(4) = 2
      dbl_NA(5) = 3
      dbl_NA(6) = 2
      dbl_NA(7) = 2
      dbl_NA(8) = 2
      
      dbl_NZ(1, 1) = 13
      dbl_NZ(2, 1) = 18
      dbl_NZ(3, 1) = 18
      dbl_NZ(4, 1) = 25
      dbl_NZ(5, 1) = 32
      dbl_NZ(6, 1) = 12
      dbl_NZ(7, 1) = 24
      dbl_NZ(8, 1) = 24
      dbl_NZ(1, 2) = 3
      dbl_NZ(2, 2) = 2
      dbl_NZ(3, 2) = 10
      dbl_NZ(4, 2) = 14
      dbl_NZ(5, 2) = 28
      dbl_NZ(6, 2) = 11
      dbl_NZ(7, 2) = 18
      dbl_NZ(8, 2) = 14
      dbl_NZ(1, 3) = 0
      dbl_NZ(2, 3) = 1
      dbl_NZ(3, 3) = 0
      dbl_NZ(4, 3) = 0
      dbl_NZ(5, 3) = 24
      dbl_NZ(6, 3) = 0
      dbl_NZ(7, 3) = 0
      dbl_NZ(8, 3) = 0
      
      dbl_NL(1) = 1
      dbl_NL(2) = 1
      dbl_NL(3) = 2
      
      dbl_NX(1, 1) = 14
      dbl_NX(2, 1) = 19
      dbl_NX(3, 1) = 54
      dbl_NX(1, 2) = 0
      dbl_NX(2, 2) = 0
      dbl_NX(3, 2) = 27
      
      dbl_BG(1, 1) = 0.06670375918
      dbl_BG(2, 1) = 0.08390104328
      dbl_BG(3, 1) = 0.4520918904
      dbl_BG(4, 1) = -0.5975336707
      dbl_BG(5, 1) = 0.5958051609
      dbl_BG(6, 1) = 0.1190610271
      dbl_BG(7, 1) = 0.1683998803
      dbl_BG(8, 1) = 0.006552390126
      dbl_BG(1, 2) = 1.388983801
      dbl_BG(2, 2) = 0.02614670893
      dbl_BG(3, 2) = 0.1069036614
      dbl_BG(4, 2) = -0.08847535804
      dbl_BG(5, 2) = -0.5159303373
      dbl_BG(6, 2) = -0.09867174132
      dbl_BG(7, 2) = -0.05809438001
      dbl_BG(8, 2) = 0.0005710218649
      dbl_BG(1, 3) = 0
      dbl_BG(2, 3) = -0.03373439453
      dbl_BG(3, 3) = 0
      dbl_BG(4, 3) = 0
      dbl_BG(5, 3) = 0.2075021122
      dbl_BG(6, 3) = 0
      dbl_BG(7, 3) = 0
      dbl_BG(8, 3) = 0
      dbl_BK(1, 1) = 0.4006073948
      dbl_BK(2, 1) = 0.08636081627
      dbl_BK(3, 1) = -0.8532322921
      dbl_BK(1, 2) = 0
      dbl_BK(2, 2) = 0
      dbl_BK(3, 2) = 0.3460208861
      
      dbl_B9(1) = 193.6587558
      dbl_B9(2) = -1388.522425
      dbl_B9(3) = 4126.607219
      dbl_B9(4) = -6508.211677
      dbl_B9(5) = 5745.984054
      dbl_B9(6) = -2693.088365
      dbl_B9(7) = 523.5718623
      
      dbl_RI1 = 4.260321148
      dbl_BKA = 0.7633333333
      
'**********************************
'Begin Calculation
'**********************************

      dbl_Tr = dbl_T / dbl_Tc
      dbl_Pr = dbl_P / dbl_Pc
      
      dbl_X = Exp(dbl_BKA * (1 - dbl_Tr))
      dbl_PRL = BL(dbl_Tr)
      dbl_DPRL = DBL(dbl_Tr)

      dbl_S1 = dbl_RI1 * dbl_Tr / dbl_Pr

      dbl_S2 = 0
      For I = 1 To 5
       dbl_S2A = 0
       dbl_NAB = dbl_NA(I)
      For IA = 1 To dbl_NAB
       dbl_S2A = dbl_S2A + dbl_BG(I, IA) * dbl_X ^ dbl_NZ(I, IA)
      Next IA
       dbl_S2 = dbl_S2 + I * dbl_Pr ^ (I - 1) * dbl_S2A
      Next I

       dbl_S3 = 0
      For I = 6 To 8
         dbl_LAM = I - 5
         dbl_S3A = 0
         dbl_NAB = dbl_NA(I)
     For IA = 1 To dbl_NAB
         dbl_S3A = dbl_S3A + dbl_BG(I, IA) * dbl_X ^ dbl_NZ(I, IA)
     Next IA
         dbl_S3ZAE = (I - 2) * dbl_Pr ^ (1 - I) * dbl_S3A
         dbl_S3B = 0
         dbl_NAC = dbl_NL(dbl_LAM)
     For IB = 1 To dbl_NAC
         dbl_S3B = dbl_S3B + dbl_BK(dbl_LAM, IB) * dbl_X ^ dbl_NX(dbl_LAM, IB)
     Next IB
         dbl_S3HI = dbl_Pr ^ (2 - I) + dbl_S3B
         dbl_S3NEN = dbl_S3HI * dbl_S3HI
         dbl_S3 = dbl_S3 + dbl_S3ZAE / dbl_S3NEN
     Next I

        dbl_S4 = 0
    For I = 1 To 7
        dbl_S4 = dbl_S4 + dbl_B9(I) * dbl_X ^ (I - 1)
    Next I
       dbl_S4 = 11# * (dbl_Pr / dbl_PRL) ^ 10 * dbl_S4

       dbl_VR = dbl_S1 - dbl_S2 - dbl_S3 + dbl_S4
      
      SpecificVolumeVapor = dbl_VR * dbl_Vc 'Function undef. dbl_Vc=0
      
SpecificVolumeVapor_Error:
'Placeholder
      
End Function

Function VaporPressureBarA(dbl_T As Double)

'--------------------------------------------------------------------------------------
'Source: Properties of Water and Steam in SI Units, 2nd edition,
'Springer -Verlag, Berlin, 1979
'---------------------------------------------------------------------------------------
'This function returns the vapor pressure of water in bar A.
'----------------------------------------------------------------------------------------
'Input Variables:
'dbl_T                          Temperature          deg K
'--------------------------------------------------------------------------------------
'Output Variables:
'VaporPressureWaterBarA         Vapor pressure       bar a
'
'Coded By:
'Date:
'------------------------------------------------------------------------------------
On Error GoTo VaporPressureBarA_Error

'**********************************
'Check input validation
'**********************************
'
'
'**********************************
'Declarations
'**********************************

      Dim dbl_AK(12) As Double
      Dim dbl_BK(4) As Double
      Dim dbl_Tz(12) As Double
      Dim I As Integer
      Dim dbl_Tr As Double
      Dim dbl_U As Double
      Dim dbl_Sum As Double
      Dim dbl_Pr As Double
      
      dbl_AK(1) = -4.0596821
      dbl_AK(2) = 5.1322555
      dbl_AK(3) = -1.1842407
      dbl_AK(4) = 0.11779592
      dbl_AK(5) = -0.005157642
      dbl_AK(6) = -0.0014689537
      dbl_AK(7) = 0.00053622818
      dbl_AK(8) = 0.00012455399
      dbl_AK(9) = -0.000049154288
      dbl_AK(10) = 0.000046302565
      dbl_AK(11) = 0.000015301334
      dbl_AK(12) = -0.00002095453
      
      dbl_BK(1) = 2
      dbl_BK(2) = 0.95
      dbl_BK(3) = 1.45220717
      dbl_BK(4) = -0.84878953
      
'**********************************
'Begin Calculation
'**********************************

      dbl_Tr = dbl_T / dbl_Tc

      If dbl_Tr < 0.421 Or dbl_Tr > 1.658 Then
         VaporPressureBarA = -1
         
        Exit Function
        
      End If
  
      dbl_U = (dbl_BK(1) * (1 / dbl_Tr - dbl_BK(2)) ^ 0.4 - dbl_BK(3)) / dbl_BK(4)
      dbl_Tz(1) = 1
      dbl_Tz(2) = dbl_U
      For I = 3 To 12
        dbl_Tz(I) = 2 * dbl_U * dbl_Tz(I - 1) - dbl_Tz(I - 2)
      Next I
      dbl_Sum = 0
      For I = 1 To 12
        dbl_Sum = dbl_Sum + dbl_AK(I) * dbl_Tz(I)
      Next I
      dbl_Pr = Exp(dbl_Sum)
      VaporPressureBarA = dbl_Pr * dbl_Pc
      
VaporPressureBarA_Error:
'Placeholder
      
End Function

Sub CalcFf(dbl_Pv As Double, dbl_Pc As Double, dbl_Ff As Double)

'------------------------------------------------------------------------------------------
'Source: Masoneilan Control Valve Sizing and Selection Handbook
'    Bulletin OZ1000 6/94, p. 7
'
'This subroutine calculates the liquid critical pressure factor.
'----------------------------------------------------------------------------------------
'Input Variables:
'dbl_Pv          vapor pressure of liquid                    bar a
'dbl_Pc          critical pressure                           bar a
'
'Output Variables:
'dbl_Ff          liquid critical pressure factor             bar a
'
'Author:  KPMG
'Date:  6/12/98
'----------------------------------------------------------------------------------------

On Error GoTo CalcFf_Error

'**************************
'Check input data
'**************************

'Pc must be greater than 0
If dbl_Pc <= 0 Then GoTo CalcFf_Error

'**************************
'Begin calculation
'**************************

dbl_Ff = 0.96 - 0.28 * Sqr(dbl_Pv / dbl_Pc)

Exit Sub

CalcFf_Error:
'Placeholder

End Sub

Sub CalcXtp(dbl_P1 As Double, dbl_P2 As Double, dbl_Cv As Double, _
    dbl_Ki As Double, dbl_ValveInletDia As Double, dbl_FL As Double, _
    dbl_Fm As Double, dbl_Fp As Double, dbl_K As Double, _
    dbl_Xm As Double, dbl_Xtp As Double)

'--------------------------------------------------------------------------------------
'Source:    Masoneilan Control Valve Sizing and Selection Handbook
'           Bulletin OZ1000 6/94, p. 7, 12
'
'This subroutine calculates Xt and Xtp
'--------------------------------------------------------------------------------------
'
'Input Variables:
'
'dbl_P1             inlet pressure                                  bar a
'dbl_P2             outlet pressure                                 bar a
'dbl_Cv             Cv required based on spec'd service conditions
'dbl_Ki             sum of Bernoulli coefficients
'dbl_ValveInletDia  valve inlet diameter                            mm
'dbl_FL             liquid pressure recovery  factor
'dbl_Fm             multistage compressible flow factor
'dbl_Fp             piping geometry factor
'
'Output Variables:
'dbl_Xm             pressure drop ratio
'dbl_Xtp            pressure drop ratio factor including piping
'
'Author:  KPMG
'Date:  6/12/98
'--------------------------------------------------------------------------------------
On Error GoTo calcXtp_Error

'**************************
'Check input data
'**************************
'dbl_ValveInletDia, dbl_Fp and dbl_P1 must be greater than 0
If (dbl_ValveInletDia < 0 Or dbl_Fp < 0 Or dbl_P1 < 0) _
    Then GoTo calcXtp_Error

'*************************
'Declarations
'*************************
Dim dbl_Xt As Double

'*************************
'Begin calculation
'*************************

dbl_Xt = 0.84 * dbl_FL ^ 2
dbl_Xtp = (dbl_Xt / dbl_Fp ^ 2) * (0.00214 * dbl_ValveInletDia ^ 4 / (dbl_Xt * dbl_Ki * dbl_Cv ^ 2 + 0.00214 * dbl_ValveInletDia ^ 4))
dbl_Xm = dbl_Fm * (dbl_P1 - dbl_P2) / dbl_P1
If dbl_Xm > dbl_Xtp Then dbl_Xm = dbl_Xtp

Exit Sub

calcXtp_Error:
'Placeholder

End Sub
Sub CalcY(dbl_Xm As Double, dbl_Xtp As Double, dbl_K As Double, _
    dbl_Ym As Double)

'------------------------------------------------------------------------------------------
'Source: Masoneilan Control Valve Sizing and Selection Handbook
'    Bulletin OZ1000 6/94, p. 9
'
'This subroutine calculates the gas expansion factor, Y, which is used in gas Cv calculation.
'
'------------------------------------------------------------------------------------------
'Input Variables:
'dbl_Xm                 pressure drop ratio
'dbl_Xtp                pressure drop ratio factor
'dbl_k                  ratio of specific heats
'--------------------------------------------------------------------------------------
'Output Variables:
'dbl_Ym                 gas expansion factor

'Author:  KPMG
'Date:  6/12/98
'--------------------------------------------------------------------------------------

On Error GoTo calcY_Error

'*************************
'Check input values
'*************************

'dbl_k and dbl_Xtp must not be equal to 0
If (dbl_K = 0 Or dbl_Xtp = 0) Then GoTo calcY_Error

'*************************
'Declarations
'*************************
Dim dbl_Fk As Double

'*************************
'Begin calculation
'*************************

dbl_Fk = dbl_K / 1.4
dbl_Ym = 1 - (dbl_Xm / (3 * dbl_Fk * dbl_Xtp))
If dbl_Ym < 0.67 Then
    dbl_Ym = 0.67
ElseIf dbl_Ym > 1 Then
    dbl_Ym = 1
End If

Exit Sub

calcY_Error:
'Placeholder

End Sub

Sub CalcFLp(dbl_FL As Double, dbl_CalculatedCv As Double, _
    dbl_ValveInletDia As Double, dbl_Ki As Double, dbl_Flp As Double)

'----------------------------------------------------------------------------------------
'Source: Masoneilan Control Valve Sizing and Selection Handbook
'    Bulletin OZ1000 6/94, p. 6
'
'This subroutine is used to calculate the combined valve geometry factor and pressure recovery factor, FLp.
'
'----------------------------------------------------------------------------------------
'Input Variables:
'dbl_FL                 liquid pressure recovery factor
'dbl_CalculatedCv       required Cv for specified service
'dbl_ValveInletDia      size if valve inlet diameter            mm
'dbl_Ki                 sum of inlet loss and Bernouli coef.
'
'Output Variables:
'dbl_FLp                combined valve geometry and pressure
'                       recovery factors
'
'Author:  KPMG
'Date:  6/12/98
'------------------------------------------------------------------------------------------

On Error GoTo calcFlp_Error

'**************************
'Check input values
'**************************

'dbl_ValveInletDia must be greater than 0
If dbl_ValveInletDia <= 0 Then GoTo calcFlp_Error

'**************************
'Begin calculation
'**************************

dbl_Flp = Sqr(0.00214 * dbl_ValveInletDia ^ 4 / (0.00214 * dbl_ValveInletDia ^ 4 + dbl_Ki * dbl_FL ^ 2 * dbl_CalculatedCv ^ 2))

Exit Sub

calcFlp_Error:
'Placeholder

End Sub
Sub Calc2phaseMaxDP(dbl_P1 As Double, dbl_P2 As Double, _
    dbl_Pv As Double, dbl_Ff As Double, dbl_K As Double, _
    dbl_FL As Double, dbl_Xt As Double, dbl_DPLiquid As Double, _
    dbl_DPGas As Double)

'-----------------------------------------------------------------------------------------------------
'Source: Masoneilan Control Valve Sizing and Selection Handbook
'    Bulletin OZ1000 6/94, p. 11
'
'This subroutine calculates the liquid and gas pressure drops which are used to calculate two phase flow Cv.
'-----------------------------------------------------------------------------------------------------
'Input Variables:
'dbl_P1          inlet pressure                                         bar a
'dbl_P2          outlet pressure                                        bar a
'dbl_Pv          vapor pressure of liquid                               bar a
'dbl_Ff          liquid pressure recovery factor
'dbl_FL          pressure recovery factor
'dbl_Xt          pressure drop ratio factor
'dbl_k           ratio of specific heats
'
'Output Variables:
'dbl_DPliquid    pressure drop to be used for liquid phase calculation  bar
'dbl_DPgas       pressure drop to be used for gas phase calculation     bar
'
'Author:  KPMG
'Date:  6/12/98
'---------------------------------------------------------------------------------------------

On Error GoTo calc2phaseMaxDP_Error:

'************************
'Check input data
'************************

'dbl_P1, dbl_P2, dbl_Pv and dbl_k must be greater than 0
If (dbl_P1 <= 0 Or dbl_P2 <= 0 Or dbl_Pv <= 0 Or dbl_K <= 0) Then GoTo calc2phaseMaxDP_Error

'***********************
'Declarations
'***********************

Dim dbl_DP As Double
Dim dbl_MaxDPLiquid As Double
Dim dbl_MaxDPGas As Double
Dim dbl_Fk As Double

'**********************
'Begin calculations
'**********************

dbl_DP = dbl_P1 - dbl_P2
dbl_MaxDPLiquid = dbl_FL ^ 2 * (dbl_P1 - dbl_Ff * dbl_Pv)
dbl_Fk = dbl_K / 1.4
dbl_MaxDPGas = dbl_Fk * dbl_Xt * dbl_P1
If dbl_DP <= dbl_MaxDPLiquid Then
    dbl_DPLiquid = dbl_DP
Else
    dbl_DPLiquid = dbl_MaxDPLiquid
End If
If dbl_DP <= dbl_MaxDPGas Then
    dbl_DPGas = dbl_DP
Else
    dbl_DPGas = dbl_MaxDPGas
End If

Exit Sub

calc2phaseMaxDP_Error:
'Placeholder

End Sub


Sub CalcTwoPhaseCv(dbl_InletPipeDia As Double, _
    dbl_OutletPipeDia As Double, dbl_ValveInletDia As Double, _
    dbl_ValveOutletDia As Double, _
    lng_TrimTypeID As Long, lng_FlowDirectionID As Long, lng_TrimPackageID As Long, _
    str_ValveLine As String, dbl_RatedCv As Double, _
    dbl_W As Double, dbl_Inlet2phDensity As Double, _
    dbl_Pv As Double, dbl_Pc As Double, dbl_K As Double, dbl_Gf As Double, _
    dbl_wtFractionGas As Double, dbl_wtFractionLiquid As Double, _
    dbl_spWeightGas As Double, dbl_spWeightLiquid As Double, _
    bln_ReducedArea As Boolean, dbl_Cv As Double, _
    dbl_FL As Double, dbl_FlowWGas As Double, dbl_FlowWTotal As Double, _
    dbl_InletPressure As Double, dbl_OutletPressure As Double, _
    dbl_InletTemperature As Double, str_PressureUM As String, str_AbsPressureUM As String, _
    str_TemperatureUM As String, str_DensityUM As String, str_FlowUM As String, _
    dbl_InputDensity As Double, dbl_InputPc As Double, dbl_InputPv As Double, _
    dbl_PctTravel As Double)

'-----------------------------------------------------------------------------------------------------
'Source: Masoneilan Control Valve Sizing and Selection Handbook
'    Bulletin OZ1000 6/94, p. 11
'
'This subroutine is used to calculate the required Cv for two phase flow service.
'
'-----------------------------------------------------------------------------------------------------
'Input Variables:
'dbl_InletPipeDia        size of inlet pipe                         mm
'dbl_OutletPipeDia       size of outlet pipe                        mm
'dbl_ValveInletDia       size of valve inlet connection             mm
'dbl_ValveOutletDia      size of valve outlet connection            mm
'str_ValveLine           description of valve line/trim
'lng_TrimTypeID          ID of Valve Trim
'dbl_RatedCv             rated Cv of selected valve/trim
'str_AbsPressureUM       absolute pressure unit of measure
'str_PressureUM          pressure unit of measure
'str_TemperatureUM       temperature unit of measure
'str_FlowUM              flow rate unit of measure
'str_DensityUM           density unit of measure
'dbl_FlowWTotal          total flow rate of two phase mixture
'dbl_FlowWGas            flow rate of the gas phase entering the valve
'dbl_InletPressure       Inlet pressure
'dbl_OutletPressure      Outlet pressure
'dbl_InletTemperature    Inlet temperature

'dbl_Pv                  vapor pressure                                     bar a
'dbl_Pc                  critical pressure of liquid                        bar a
'dbl_k                   ratio of specific heats for gas
'---------------------------------------------------------------------------------------------
'Internal Variables:
'dbl_wtFractionGas       weight fraction of gas in two phase mixture     (range 0 - 1)
'dbl_wtFractionLiquid    weight fraction of liquid in two phase mixture      (range 0 - 1)
'dbl_spWeightGas         specific weight of the gas phase                kg/m3
'dbl_spWeightLiquid      specific weight of the liquid phase         kg/m3
'dbl_Ff                  calculated value returned from CalcFf
'dbl_DPLiquid            calculated value of max pressure drop for Liquid phase
'dbl_DPGas               calculated value of max pressure drop for Gas phase
'dbl_P1                  inlet pressure                                     bar a
'dbl_P2                  outlet pressure                                    bar a
'dbl_T1                  inlet temperature
'dbl_wgas                inlet flow rate of the gas phase
'dbl_W                   total inlet flow rate of the gas and liquid phases       kg/h
'---------------------------------------------------------------------------------------------
'Output Variables:
'dbl_Cv                  required Cv
'dbl_FL                  pressure recovery factor at calculated Cv
'
'Author:    KPMG
'Date:      6/12/98
'Revised:   F.Krieger
'Date:      12/07/98
'-----------------------------------------------------------------------------------------------------
On Error GoTo calcTwoPhaseCv_Error

#If (DebugVersion = "yes") Then
MsgBox "TwoPhase" + Chr(13) + Chr(10) + _
    Chr(13) + Chr(10) + "dbl_InletPipeDia=" + CStr(dbl_InletPipeDia) + _
    Chr(13) + Chr(10) + "dbl_OutletPipeDia=" + CStr(dbl_OutletPipeDia) + Chr(13) + Chr(10) + "dbl_ValveInletDia=" + CStr(dbl_ValveInletDia) + _
    Chr(13) + Chr(10) + "dbl_ValveOutletDia=" + CStr(dbl_ValveOutletDia) + Chr(13) + Chr(10) + _
    Chr(13) + Chr(10) + "lng_TrimTypeID=" + CStr(lng_TrimTypeID) + Chr(13) + Chr(10) + "lng_FlowDirectionID=" + CStr(lng_FlowDirectionID) + Chr(13) + Chr(10) + "lng_TrimPackageID=" + CStr(lng_TrimPackageID) + _
    Chr(13) + Chr(10) + "str_ValveLine=" + str_ValveLine + Chr(13) + Chr(10) + "dbl_RatedCv=" + CStr(dbl_RatedCv) + _
    Chr(13) + Chr(10) + "dbl_W=" + CStr(dbl_W) + Chr(13) + Chr(10) + "dbl_Inlet2phDensity=" + CStr(dbl_Inlet2phDensity) + _
    Chr(13) + Chr(10) + "dbl_Pv=" + CStr(dbl_Pv) + Chr(13) + Chr(10) + "dbl_Pc=" + CStr(dbl_Pc) + Chr(13) + Chr(10) + "dbl_K=" + CStr(dbl_K) + Chr(13) + Chr(10) + "dbl_Gf=" + CStr(dbl_Gf) + _
    Chr(13) + Chr(10) + "dbl_wtFractionGas=" + CStr(dbl_wtFractionGas) + Chr(13) + Chr(10) + "dbl_wtFractionLiquid=" + CStr(dbl_wtFractionLiquid) + _
    Chr(13) + Chr(10) + "dbl_spWeightGas=" + CStr(dbl_spWeightGas) + Chr(13) + Chr(10) + "dbl_spWeightLiquid=" + CStr(dbl_spWeightLiquid) + _
    Chr(13) + Chr(10) + "bln_ReducedArea=" + CStr(bln_ReducedArea) + Chr(13) + Chr(10) + "dbl_Cv=" + CStr(dbl_Cv) + _
    Chr(13) + Chr(10) + "dbl_FL=" + CStr(dbl_FL) + "dbl_FlowWGas=" + CStr(dbl_FlowWGas) + Chr(13) + Chr(10) + "dbl_FlowWTotal=" + CStr(dbl_FlowWTotal) + _
    Chr(13) + Chr(10) + "dbl_InletPressure=" + CStr(dbl_InletPressure) + Chr(13) + Chr(10) + "dbl_OutletPressure=" + CStr(dbl_OutletPressure) + _
    Chr(13) + Chr(10) + "dbl_InletTemperature=" + CStr(dbl_InletTemperature) + Chr(13) + Chr(10) + "str_PressureUM=" + str_PressureUM + Chr(13) + Chr(10) + "str_AbsPressureUM=" + str_AbsPressureUM + _
    Chr(13) + Chr(10) + "str_TemperatureUM=" + str_TemperatureUM + Chr(13) + Chr(10) + "str_DensityUM=" + str_DensityUM + Chr(13) + Chr(10) + "str_FlowUM=" + str_FlowUM + _
    Chr(13) + Chr(10) + "dbl_InputDensity=" + CStr(dbl_InputDensity) + Chr(13) + Chr(10) + "dbl_InputPc=" + CStr(dbl_InputPc) + Chr(13) + Chr(10) + "dbl_InputPv=" + CStr(dbl_InputPv) + _
    Chr(13) + Chr(10) + "dbl_PctTravel=" + CStr(dbl_PctTravel)
#End If


'********************
'Check input values
'********************


'******************
'Declarations
'******************

Dim dbl_CalculatedCv As Double
Dim dbl_DPGas As Double
Dim dbl_DPLiquid As Double
Dim dbl_Ff As Double
Dim dbl_Fk As Double
Dim dbl_Flp As Double
Dim dbl_Fm As Double
Dim dbl_Fp As Double
Dim dbl_Ki As Double
Dim dbl_Ksum As Double
Dim dbl_LastCv As Double
Dim dbl_P1 As Double
Dim dbl_P2 As Double
Dim dbl_T1 As Double
Dim dbl_wgas As Double
Dim dbl_Xm As Double
Dim dbl_Xt As Double
Dim dbl_Xtp As Double
Dim dbl_Y As Double
Dim int_Counter As Integer
Dim int_NumIterations As Integer

'************************
'Begin calculations
'************************
dbl_FL = 0.9 'default to start
dbl_wtFractionGas = dbl_FlowWGas / dbl_FlowWTotal

dbl_wtFractionLiquid = 1 - dbl_wtFractionGas

'Convert Input Data:

Call convertTemperatureToDegK(dbl_InletTemperature, str_TemperatureUM, dbl_T1)
Call convertPressureToBarA(dbl_InletPressure, str_PressureUM, dbl_P1)
Call convertPressureToBarA(dbl_OutletPressure, str_PressureUM, dbl_P2)
Call convertDensityToKgM3AtInlet(dbl_InputDensity, str_DensityUM, _
    dbl_T1, dbl_P1, dbl_spWeightGas)
Call convertFlowRateToKgh(dbl_FlowWGas, str_FlowUM, dbl_spWeightGas, dbl_wgas)
Call convertPressureToBarA(dbl_InputPc, str_AbsPressureUM, dbl_Pc)
Call convertPressureToBarA(dbl_InputPv, str_AbsPressureUM, dbl_Pv)

dbl_spWeightLiquid = dbl_Gf * 1000

dbl_Inlet2phDensity = 1 / ((dbl_wtFractionLiquid / dbl_spWeightLiquid) + _
    (dbl_wtFractionGas / dbl_spWeightGas))

Call convertFlowRateToKgh(dbl_FlowWTotal, str_FlowUM, dbl_Inlet2phDensity, dbl_W)


'Check str_ValveLine
If str_ValveLine = "" Then
    int_NumIterations = 0
Else
    int_NumIterations = 10
End If

Call CalcFf(dbl_Pv, dbl_Pc, dbl_Ff)
Call PipeGeom(dbl_InletPipeDia, dbl_OutletPipeDia, dbl_ValveInletDia, _
    dbl_ValveOutletDia, dbl_Ksum, dbl_Ki)
    
dbl_LastCv = 0
For int_Counter = 0 To int_NumIterations
    If int_Counter = 0 Then
        
        
        dbl_Flp = dbl_FL
        dbl_Fp = 1
        dbl_Fm = 1
        dbl_Xt = 0.84 * dbl_FL ^ 2
        dbl_Xtp = dbl_Xt
        dbl_Xm = ((dbl_P1 - dbl_P2) / dbl_P1) * dbl_Fm
        dbl_Fk = dbl_K / 1.4
        If dbl_Xm > dbl_Fk * dbl_Xtp Then
            dbl_Xm = dbl_Fk * dbl_Xtp
        End If
    Else
        Call GetTravel(str_ValveLine, lng_TrimTypeID, dbl_RatedCv, dbl_Cv, _
            dbl_PctTravel)
        Call GetFL(str_ValveLine, lng_FlowDirectionID, lng_TrimTypeID, _
            bln_ReducedArea, dbl_RatedCv, dbl_Cv, dbl_FL)
        Call GetFm(str_ValveLine, lng_TrimTypeID, lng_TrimPackageID, dbl_Fm)
        Call CalcFp(dbl_ValveInletDia, dbl_Ksum, dbl_Ki, dbl_Cv, dbl_FL, _
            dbl_Fp, dbl_Flp)
        Call CalcXtp(dbl_P1, dbl_P2, dbl_Cv, dbl_Ki, dbl_ValveInletDia, dbl_FL, _
            dbl_Fm, dbl_Fp, dbl_K, dbl_Xm, dbl_Xtp)
    End If

    Call CalcY(dbl_Xm, dbl_Xtp, dbl_K, dbl_Y)
    Call Calc2phaseMaxDP(dbl_P1, dbl_P2, dbl_Pv, dbl_Ff, dbl_K, dbl_FL, _
        dbl_Xtp, dbl_DPLiquid, dbl_DPGas)
        
    dbl_Cv = (dbl_W / (27.3 * dbl_Fp)) _
        * Sqr(dbl_wtFractionLiquid / (dbl_DPLiquid * dbl_spWeightLiquid) + _
        (dbl_wtFractionGas / (dbl_DPGas * dbl_spWeightGas * dbl_Y ^ 2)))

    If Abs((dbl_LastCv - dbl_Cv) / dbl_Cv) * 100 <= 0.01 Then Exit For
    dbl_LastCv = dbl_Cv
Next int_Counter

dbl_Cv = TrimDecimalPlaces(dbl_Cv)
dbl_FL = TrimDecimalPlaces(dbl_FL)

#If (DebugVersion = "yes") Then
MsgBox "TwoPhase" + Chr(13) + Chr(10) + _
    Chr(13) + Chr(10) + "dbl_InletPipeDia=" + CStr(dbl_InletPipeDia) + _
    Chr(13) + Chr(10) + "dbl_OutletPipeDia=" + CStr(dbl_OutletPipeDia) + Chr(13) + Chr(10) + "dbl_ValveInletDia=" + CStr(dbl_ValveInletDia) + _
    Chr(13) + Chr(10) + "dbl_ValveOutletDia=" + CStr(dbl_ValveOutletDia) + Chr(13) + Chr(10) + _
    Chr(13) + Chr(10) + "lng_TrimTypeID=" + CStr(lng_TrimTypeID) + Chr(13) + Chr(10) + "lng_FlowDirectionID=" + CStr(lng_FlowDirectionID) + Chr(13) + Chr(10) + "lng_TrimPackageID=" + CStr(lng_TrimPackageID) + _
    Chr(13) + Chr(10) + "str_ValveLine=" + str_ValveLine + Chr(13) + Chr(10) + "dbl_RatedCv=" + CStr(dbl_RatedCv) + _
    Chr(13) + Chr(10) + "dbl_W=" + CStr(dbl_W) + Chr(13) + Chr(10) + "dbl_Inlet2phDensity=" + CStr(dbl_Inlet2phDensity) + _
    Chr(13) + Chr(10) + "dbl_Pv=" + CStr(dbl_Pv) + Chr(13) + Chr(10) + "dbl_Pc=" + CStr(dbl_Pc) + Chr(13) + Chr(10) + "dbl_K=" + CStr(dbl_K) + Chr(13) + Chr(10) + "dbl_Gf=" + CStr(dbl_Gf) + _
    Chr(13) + Chr(10) + "dbl_wtFractionGas=" + CStr(dbl_wtFractionGas) + Chr(13) + Chr(10) + "dbl_wtFractionLiquid=" + CStr(dbl_wtFractionLiquid) + _
    Chr(13) + Chr(10) + "dbl_spWeightGas=" + CStr(dbl_spWeightGas) + Chr(13) + Chr(10) + "dbl_spWeightLiquid=" + CStr(dbl_spWeightLiquid) + _
    Chr(13) + Chr(10) + "bln_ReducedArea=" + CStr(bln_ReducedArea) + Chr(13) + Chr(10) + "dbl_Cv=" + CStr(dbl_Cv) + _
    Chr(13) + Chr(10) + "dbl_FL=" + CStr(dbl_FL) + "dbl_FlowWGas=" + CStr(dbl_FlowWGas) + Chr(13) + Chr(10) + "dbl_FlowWTotal=" + CStr(dbl_FlowWTotal) + _
    Chr(13) + Chr(10) + "dbl_InletPressure=" + CStr(dbl_InletPressure) + Chr(13) + Chr(10) + "dbl_OutletPressure=" + CStr(dbl_OutletPressure) + _
    Chr(13) + Chr(10) + "dbl_InletTemperature=" + CStr(dbl_InletTemperature) + Chr(13) + Chr(10) + "str_PressureUM=" + str_PressureUM + Chr(13) + Chr(10) + "str_AbsPressureUM=" + str_AbsPressureUM + _
    Chr(13) + Chr(10) + "str_TemperatureUM=" + str_TemperatureUM + Chr(13) + Chr(10) + "str_DensityUM=" + str_DensityUM + Chr(13) + Chr(10) + "str_FlowUM=" + str_FlowUM + _
    Chr(13) + Chr(10) + "dbl_InputDensity=" + CStr(dbl_InputDensity) + Chr(13) + Chr(10) + "dbl_InputPc=" + CStr(dbl_InputPc) + Chr(13) + Chr(10) + "dbl_InputPv=" + CStr(dbl_InputPv) + _
    Chr(13) + Chr(10) + "dbl_PctTravel=" + CStr(dbl_PctTravel)
#End If
Exit Sub

calcTwoPhaseCv_Error:
'Placeholder

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' END OF two phase and steam
''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INCLUDE DOMOTOR
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CalcReqDomotorPreloadSupply_FTO_ATC(dbl_ReqForceClosed As Double, _
    dbl_ForceSpringClosed As Double, dbl_AreaOver As Double, _
    dbl_RequiredPreload As Double, dbl_RequiredSupply As Double)
'---------------------------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for reciprocating valves,  Masoneilan Product Knowledge Program.
'---------------------------------------------------------------------------------------------------------
'This subroutine is used to calculate the required preload and supply pressures for
'Domotor actuators for flow to open air to closed operation.
'--------------------------------------------------------------------------------------------------------------------
'Input Variables:
'dbl_ReqForceClosed         actuator force required to closed valve lb
'dbl_ForceSpringClosed      compressed force for helper spring in Domotor lb
'dbl_AreaOver               area of top side of Domotor piston  sq.in.

'Output Variables:
'dbl_RequiredPreload        required preload pressure for service   psig
'dbl_RequiredSupply         required supply pressure to operate valve psig
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/22/98  Name: David Zhou
'Tested     Date: 10/23/98  Name: David Zhou        Passed(Y/N): Y
'-----------------------------------------------------------------------------------------
'**************************
'Declarations
'**************************
Dim I As Integer
On Error GoTo CalcReqDomotorPreloadSupply_FTO_ATC_Error

'***************************
'Parameter validation
'***************************
If dbl_ReqForceClosed < 0# Then GoTo CalcReqDomotorPreloadSupply_FTO_ATC_Error
If dbl_ForceSpringClosed < 0# Then GoTo CalcReqDomotorPreloadSupply_FTO_ATC_Error
If dbl_AreaOver < 0# Then GoTo CalcReqDomotorPreloadSupply_FTO_ATC_Error

'****************************
'Begin calculations
'****************************
dbl_RequiredPreload = -1
dbl_RequiredSupply = -1
dbl_RequiredPreload = 1.1 * (dbl_ReqForceClosed + dbl_ForceSpringClosed) / dbl_AreaOver

dbl_RequiredPreload = 5 * (Int(dbl_RequiredPreload / 5 - 0.000001) + 1)
If (dbl_RequiredPreload < 10) Then dbl_RequiredPreload = 10
If (dbl_RequiredPreload > 100) Then dbl_RequiredPreload = 100

If (dbl_RequiredPreload > 0) Then dbl_RequiredSupply = dbl_RequiredPreload + 10

Exit Sub

CalcReqDomotorPreloadSupply_FTO_ATC_Error:
'Placeholder

End Sub

Public Sub CalcDPAllowableDomotor_FTO_ATC(dbl_RequiredPreload As Double, dbl_AreaOver _
    As Double, dbl_ForceSpringClosed As Double, dbl_ForceSeating As Double, _
    dbl_ForcePacking As Double, dbl_SeatDiameter As Double, dbl_DPAllowable As Double)
'---------------------------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for reciprocating valves,  Masoneilan Product Knowledge Program.
'---------------------------------------------------------------------------------------------------------
'This subroutine is used to calculate the allowable pressure drop for Domotor actuators
'for flow to open air to closed operation.
'--------------------------------------------------------------------------------------------------------------------
'Input Variables:
'dbl_RequiredPreload        specified Preload pressure                        psig
'dbl_ForceSpringClosed      compressed force for helper spring in Domotor     lb
'dbl_ForceSeating           force necessary to provide selected leakage specification  lb
'dbl_ForcePacking           force necessary to overcome packing friction      lb
'dbl_SeatDiameter           diameter of seat ring orifice                     inches
'dbl_AreaOver               area of top side of Domotor piston                sq.in.
                        
'Output Variables:
'dbl_DPallowable            calculated allowable pressure drop with specified actuator  psi
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/23/98  Name: David Zhou
'Tested     Date: 10/23/98  Name: David Zhou        Passed(Y/N): Y
'-----------------------------------------------------------------------------------------
'**************************
'Declarations
'**************************
Dim dbl_ForceAvailable As Double
Const PI As Double = 3.1415927
On Error GoTo CalcDPAllowableDomotor_FTO_ATC_Error

'***************************
'Parameter validation
'***************************
If dbl_RequiredPreload < 0# Then GoTo CalcDPAllowableDomotor_FTO_ATC_Error
If dbl_AreaOver < 0# Then GoTo CalcDPAllowableDomotor_FTO_ATC_Error
If dbl_ForceSpringClosed < 0# Then GoTo CalcDPAllowableDomotor_FTO_ATC_Error
If dbl_ForceSeating < 0# Then GoTo CalcDPAllowableDomotor_FTO_ATC_Error
If dbl_ForcePacking < 0# Then GoTo CalcDPAllowableDomotor_FTO_ATC_Error
If dbl_SeatDiameter < 0# Then GoTo CalcDPAllowableDomotor_FTO_ATC_Error

'****************************
'Begin calculations
'****************************
dbl_ForceAvailable = dbl_RequiredPreload * dbl_AreaOver / 1.1 - dbl_ForceSpringClosed _
                - dbl_ForceSeating - dbl_ForcePacking
dbl_DPAllowable = dbl_ForceAvailable / (PI * dbl_SeatDiameter * dbl_SeatDiameter / 4)

Exit Sub

CalcDPAllowableDomotor_FTO_ATC_Error:
'Placeholder

End Sub

Public Sub CalcReqDomotorPreloadSupply_FTC_ATO(dbl_MaxInletPressure As Double, _
    dbl_AreaOver As Double, dbl_AreaUnder As Double, dbl_VolOver As Double, _
    dbl_Travel As Double, dbl_SeatDiameter As Double, dbl_DynamicFactor As Double, _
    dbl_ForcePacking As Double, dbl_ReqPreload As Double, dbl_ReqSupply As Double)
'---------------------------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for reciprocating valves,  Masoneilan Product
'Knowledge Program.
'---------------------------------------------------------------------------------------------------------
'This subroutine is used to calculate the required preload and supply pressures for
'Domotor actuators for flow to close air to open operation.
'--------------------------------------------------------------------------------------------------------------------
'Input Variables:
'dbl_MaxInletPressure   maximum pressure used for actuator sizing           psig
'dbl_AreaOver           area of top side of Domotor piston                  in^2
'dbl_AreaUnder          area of bottom side of Domotor piston               in^2
'dbl_VolOver            volume above the Domotor piston                     in^3
'dbl_Travel             full valve plug travel                              in.
'dbl_SeatDiameter       size of the valve orifice                           in.
'dbl_DynamicFactor      factor used to determine dynamic stability limit
'dbl_ForcePacking       frictional force caused by packing                  lbs

'Output Variables:
'dbl_ReqPreload         required preload pressure                           psig
'dbl_ReqSupply          required supply pressure                            psig
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/23/98  Name: David Zhou
'Tested     Date: 10/23/98  Name: David Zhou        Passed(Y/N): Y
'-----------------------------------------------------------------------------------------
'**************************
'Declarations
'**************************
Dim I As Integer
Dim dbl_AreaAve As Double
Dim dbl_VolDead As Double
Dim dbl_PistonDP As Double
Dim dbl_Temp As Double

On Error GoTo CalcReqDomotorPreloadSupply_FTC_ATO_Error

'***************************
'Parameter validation
'***************************
If dbl_MaxInletPressure < 0# Then GoTo CalcReqDomotorPreloadSupply_FTC_ATO_Error
If dbl_AreaOver < 0# Then GoTo CalcReqDomotorPreloadSupply_FTC_ATO_Error
If dbl_AreaUnder < 0# Then GoTo CalcReqDomotorPreloadSupply_FTC_ATO_Error
If dbl_VolOver < 0# Then GoTo CalcReqDomotorPreloadSupply_FTC_ATO_Error
If dbl_Travel < 0# Then GoTo CalcReqDomotorPreloadSupply_FTC_ATO_Error
If dbl_SeatDiameter < 0# Then GoTo CalcReqDomotorPreloadSupply_FTC_ATO_Error
If dbl_DynamicFactor < 0# Then GoTo CalcReqDomotorPreloadSupply_FTC_ATO_Error
If dbl_ForcePacking < 0# Then GoTo CalcReqDomotorPreloadSupply_FTC_ATO_Error

'****************************
'Begin calculations
'****************************
dbl_ReqPreload = -1
dbl_ReqSupply = -1

dbl_AreaAve = (dbl_AreaOver + dbl_AreaUnder) / 2
dbl_VolDead = 0.125 * dbl_AreaUnder
dbl_Temp = 1 / (dbl_VolOver - dbl_AreaAve * dbl_Travel) + 1 / (dbl_VolDead + dbl_AreaAve * dbl_Travel)
dbl_ReqPreload = dbl_MaxInletPressure * dbl_DynamicFactor / _
                (1.4 * dbl_AreaAve * dbl_AreaAve * dbl_Temp)

        dbl_ReqPreload = 5 * (Int(dbl_ReqPreload / 5 - 0.000001) + 1)
        If (dbl_ReqPreload < 30) Then dbl_ReqPreload = 30
        If (dbl_ReqPreload > 50) Then dbl_ReqPreload = 50
        dbl_Temp = PI * dbl_SeatDiameter * dbl_SeatDiameter / 4 + dbl_ForcePacking
        dbl_ReqSupply = (dbl_Temp + dbl_ReqPreload * dbl_AreaOver) / dbl_AreaUnder

    dbl_PistonDP = dbl_ReqSupply - dbl_ReqPreload
    If (dbl_ReqSupply <= 40) Then
        dbl_PistonDP = 5 * (Int(dbl_PistonDP / 5 - 0.000001) + 1)
        If (dbl_PistonDP > 40) Then dbl_PistonDP = 40
        If (dbl_PistonDP < 10) Then dbl_PistonDP = 10
        
        dbl_ReqSupply = dbl_ReqPreload + dbl_PistonDP
    Else
        dbl_ReqSupply = 0
    End If

Exit Sub

CalcReqDomotorPreloadSupply_FTC_ATO_Error:
'Placeholder

End Sub

Public Sub CalcAllowDPDomotor_FTC_ATO(dbl_AreaOver As Double, dbl_AreaUnder As Double, _
    dbl_VolOver As Double, dbl_Travel As Double, dbl_SeatDiameter As Double, _
    dbl_DynamicFactor As Double, dbl_ForcePacking As Double, dbl_ReqPreload As Double, _
    dbl_ReqSupply As Double, dbl_MaxAllowDP As Double)
'---------------------------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for reciprocating valves,  Masoneilan Product Knowledge Program.
'---------------------------------------------------------------------------------------------------------
'This subroutine is used to calculate the allowable pressure drop for Domotor actuators
'with given preload and supply for flow to close air to open operation.
'--------------------------------------------------------------------------------------------------------------------
'Input Variables:
'dbl_AreaOver          area of top side of Domotor piston               in^2
'dbl_AreaUnder         area of bottom side of Domotor piston            in^2
'dbl_VolOver           volume above the Domotor piston                  in^3
'dbl_Travel            full valve plug travel                           in.
'dbl_SeatDiameter      size of the valve orifice                        in.
'dbl_DynamicFactor     factor used to determine dynamic stability limit
'dbl_ForcePacking      frictional force caused by packing               lbs
'dbl_ReqPreload        required preload pressure                        psig
'dbl_ReqSupply         required supply pressure                         psig
'dbl_ReqSupply         unbalanced area spring rate/psi for              in^2/in.                     psig
'                      selected valve/trim
'Output Variables:
'dbl_MaxAllowDP        max allowable pressure drop with specified
'                      actuator, preload and supply                     psi
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/23/98  Name: David Zhou
'Tested     Date: 10/23/98  Name: David Zhou        Passed(Y/N): Y
'-----------------------------------------------------------------------------------------
'**************************
'Declarations
'**************************
Dim dbl_AreaAve As Double
Dim dbl_VolDead As Double
Dim dbl_DynamicDP As Double
Dim dbl_StaticDP As Double
Dim dbl_Temp As Double
Const PI As Double = 3.1415927

On Error GoTo CalcAllowDPDomotor_FTC_ATO_Error

'***************************
'Parameter validation
'***************************
If dbl_AreaOver < 0# Then GoTo CalcAllowDPDomotor_FTC_ATO_Error
If dbl_AreaUnder < 0# Then GoTo CalcAllowDPDomotor_FTC_ATO_Error
If dbl_VolOver < 0# Then GoTo CalcAllowDPDomotor_FTC_ATO_Error
If dbl_Travel < 0# Then GoTo CalcAllowDPDomotor_FTC_ATO_Error
If dbl_SeatDiameter < 0# Then GoTo CalcAllowDPDomotor_FTC_ATO_Error
If dbl_DynamicFactor < 0# Then GoTo CalcAllowDPDomotor_FTC_ATO_Error
If dbl_ForcePacking < 0# Then GoTo CalcAllowDPDomotor_FTC_ATO_Error
If dbl_ReqPreload < 0# Then GoTo CalcAllowDPDomotor_FTC_ATO_Error
If dbl_ReqSupply < 0# Then GoTo CalcAllowDPDomotor_FTC_ATO_Error

'****************************
'Begin calculations
'****************************
dbl_AreaAve = (dbl_AreaOver + dbl_AreaUnder) / 2
dbl_VolDead = 0.125 * dbl_AreaUnder
dbl_Temp = 1 / (dbl_VolOver - dbl_AreaAve * dbl_Travel) + 1 / (dbl_VolDead + dbl_AreaAve * dbl_Travel)
dbl_DynamicDP = 1.4 * dbl_AreaAve * dbl_AreaAve * dbl_Temp * dbl_ReqPreload / dbl_DynamicFactor
dbl_Temp = PI * dbl_SeatDiameter * dbl_SeatDiameter / 4
dbl_StaticDP = (dbl_ReqSupply * dbl_AreaUnder - dbl_ReqPreload * dbl_AreaOver - _
                dbl_ForcePacking) / dbl_Temp
dbl_MaxAllowDP = dbl_DynamicDP
If (dbl_StaticDP < dbl_DynamicDP) Then
    dbl_MaxAllowDP = dbl_StaticDP
End If

Exit Sub

CalcAllowDPDomotor_FTC_ATO_Error:
'Placeholder

End Sub

Public Sub CalcDomotorForce_ATO(dbl_AreaOver As Double, dbl_AreaUnder As Double, _
    dbl_ReqPreload As Double, dbl_ReqSupply As Double, dbl_DomotorForce_ATO As Double)
'---------------------------------------------------------------------------------------------------------
'Reference :Actuator sizing methods for linear travel valves.
'---------------------------------------------------------------------------------------------------------
'This subroutine is used to calculate the actuator output force  for Domotors, ATO.
'--------------------------------------------------------------------------------------------------------------------
'Input Variables:
'dbl_AreaOver          area of top side of Domotor piston               in^2
'dbl_AreaUnder         area of bottom side of Domotor piston            in^2
'dbl_ReqPreload        required preload pressure                        psig
'dbl_ReqSupply         required supply pressure                         psig
'
'Output Variables:
'dbl_DomotorForce_ATO  output force of actuator                         lbs
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/23/98  Name: David Zhou
'Tested     Date: 10/23/98  Name: David Zhou        Passed(Y/N): Y
'-----------------------------------------------------------------------------------------
'**************************
'Declarations
'**************************
On Error GoTo CalcDomotorForce_ATO_Error

'***************************
'Parameter validation
'***************************
If dbl_AreaOver < 0# Then GoTo CalcDomotorForce_ATO_Error
If dbl_AreaUnder < 0# Then GoTo CalcDomotorForce_ATO_Error
If dbl_ReqPreload < 0# Then GoTo CalcDomotorForce_ATO_Error
If dbl_ReqSupply < 0# Then GoTo CalcDomotorForce_ATO_Error

'****************************
'Begin calculations
'****************************
dbl_DomotorForce_ATO = dbl_ReqSupply * dbl_AreaUnder - dbl_ReqPreload * dbl_AreaOver

Exit Sub

CalcDomotorForce_ATO_Error:
'Placeholder

End Sub

'-New Actuator Code -- 12/30/98 --------------------------------------

Public Sub SelectActuator_2600(str_ActuatorModel As String, lng_FlowDirectionID As Long, _
    lng_AirActionID As Long, lng_ActuatorPartID As Long, lng_TrimPackageID As Long, _
    dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, dbl_InletPressure3 As Double, _
    dbl_InletPressure4 As Double, dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, dbl_InletTemperature1 As Double, _
    dbl_InletTemperature2 As Double, dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, _
    dbl_ShutOffDP As Double, str_SpecifiedPressureUM As String, _
    str_SpecifiedTemperatureUM As String, str_SpecifiedCalibPressureUM As String, str_LengthUM As String, _
    str_BodyMaterial As String, str_SelectedRating As String, str_ValveLine As String, _
    lng_TrimTypeID As Long, str_PackingMaterial As String, _
    str_LeakageClass As String, str_SpringRange As String, dbl_RatedCv As Double, dbl_ShowDP As Double, _
    dbl_ShowDPSupplyPressure As Double, dbl_ShowMaxDP As Double, dbl_ShowMaxDPSupplyPressure As Double, _
    dbl_ShowPreload As Double, dbl_ShowMaxPreload As Double, str_Message As String)
    
On Error GoTo SelectActuator_2600_Error:
   
#If (DebugVersion = "yes") Then
MsgBox "2600 Actuator" + Chr(10) + Chr(13) + _
    "lng_FlowDirectionID=" + CStr(lng_FlowDirectionID) + Chr(10) + Chr(13) + _
    "lng_AirActionID=" + CStr(lng_AirActionID) + Chr(10) + Chr(13) + _
    "str_ActuatorModel=" + str_ActuatorModel + Chr(10) + Chr(13) + _
    "str_ValveLine=" + str_ValveLine
#End If
   
'***************************
'Parameter validation
'***************************
If (lng_FlowDirectionID < 0) Then GoTo SelectActuator_2600_Error:
If (lng_AirActionID < 0) Then GoTo SelectActuator_2600_Error:
If (str_ValveLine = "") Then GoTo SelectActuator_2600_Error:
If (str_ActuatorModel = "") Then GoTo SelectActuator_2600_Error:

'****************************
'Begin calculations
'****************************

If str_ValveLine = "2600" Then
    If str_ActuatorModel = "71" Then
        If lng_FlowDirectionID = FTO And lng_AirActionID = ATC Then
            Call CalcDP_2600_Domotor_FTO_ATC(str_ActuatorModel, lng_FlowDirectionID, lng_AirActionID, _
                lng_ActuatorPartID, lng_TrimPackageID, dbl_InletPressure1, dbl_InletPressure2, _
                dbl_InletPressure3, dbl_InletPressure4, dbl_OutletPressure1, dbl_OutletPressure2, _
                dbl_OutletPressure3, dbl_OutletPressure4, dbl_InletTemperature1, _
                dbl_InletTemperature2, dbl_InletTemperature3, dbl_InletTemperature4, dbl_ShutOffDP, _
                str_SpecifiedPressureUM, str_SpecifiedTemperatureUM, str_SpecifiedCalibPressureUM, _
                str_LengthUM, str_BodyMaterial, str_SelectedRating, str_ValveLine, _
                lng_TrimTypeID, str_PackingMaterial, str_LeakageClass, str_SpringRange, dbl_RatedCv, _
                dbl_ShowDP, dbl_ShowDPSupplyPressure, dbl_ShowMaxDP, dbl_ShowMaxDPSupplyPressure, _
                dbl_ShowPreload, dbl_ShowMaxPreload, str_Message)
        ElseIf lng_FlowDirectionID = FTC And lng_AirActionID = ATO Then
            Call CalcDP_2600_Domotor_FTC_ATO(str_ActuatorModel, lng_FlowDirectionID, lng_AirActionID, _
                lng_ActuatorPartID, lng_TrimPackageID, dbl_InletPressure1, dbl_InletPressure2, _
                dbl_InletPressure3, dbl_InletPressure4, dbl_OutletPressure1, dbl_OutletPressure2, _
                dbl_OutletPressure3, dbl_OutletPressure4, dbl_InletTemperature1, _
                dbl_InletTemperature2, dbl_InletTemperature3, dbl_InletTemperature4, dbl_ShutOffDP, _
                str_SpecifiedPressureUM, str_SpecifiedTemperatureUM, str_SpecifiedCalibPressureUM, _
                str_LengthUM, str_BodyMaterial, str_SelectedRating, str_ValveLine, _
                lng_TrimTypeID, str_PackingMaterial, str_LeakageClass, str_SpringRange, dbl_RatedCv, _
                dbl_ShowDP, dbl_ShowDPSupplyPressure, dbl_ShowMaxDP, dbl_ShowMaxDPSupplyPressure, _
                dbl_ShowPreload, dbl_ShowMaxPreload, str_Message)
        Else
            str_Message = INVALIDACTUATOR
        End If
    ElseIf str_ActuatorModel = "84" Or str_ActuatorModel = "85" Then
        If lng_FlowDirectionID = FTO And lng_AirActionID = ATC Then
            Call CalcDP_2600_Cylinder_FTO_ATC(str_ActuatorModel, lng_FlowDirectionID, lng_AirActionID, _
                lng_ActuatorPartID, lng_TrimPackageID, dbl_InletPressure1, dbl_InletPressure2, _
                dbl_InletPressure3, dbl_InletPressure4, dbl_OutletPressure1, dbl_OutletPressure2, _
                dbl_OutletPressure3, dbl_OutletPressure4, dbl_InletTemperature1, _
                dbl_InletTemperature2, dbl_InletTemperature3, dbl_InletTemperature4, dbl_ShutOffDP, _
                str_SpecifiedPressureUM, str_SpecifiedTemperatureUM, str_SpecifiedCalibPressureUM, _
                str_LengthUM, str_BodyMaterial, str_SelectedRating, str_ValveLine, _
                lng_TrimTypeID, str_PackingMaterial, str_LeakageClass, str_SpringRange, dbl_RatedCv, _
                dbl_ShowDP, dbl_ShowDPSupplyPressure, dbl_ShowMaxDP, dbl_ShowMaxDPSupplyPressure, str_Message)
        ElseIf lng_FlowDirectionID = FTO And lng_AirActionID = ATO Then
            Call CalcDP_2600_Cylinder_FTO_ATO(str_ActuatorModel, lng_FlowDirectionID, lng_AirActionID, _
                lng_ActuatorPartID, lng_TrimPackageID, dbl_InletPressure1, dbl_InletPressure2, _
                dbl_InletPressure3, dbl_InletPressure4, dbl_OutletPressure1, dbl_OutletPressure2, _
                dbl_OutletPressure3, dbl_OutletPressure4, dbl_InletTemperature1, _
                dbl_InletTemperature2, dbl_InletTemperature3, dbl_InletTemperature4, dbl_ShutOffDP, _
                str_SpecifiedPressureUM, str_SpecifiedTemperatureUM, str_SpecifiedCalibPressureUM, _
                str_LengthUM, str_BodyMaterial, str_SelectedRating, str_ValveLine, _
                lng_TrimTypeID, str_PackingMaterial, str_LeakageClass, str_SpringRange, dbl_RatedCv, _
                dbl_ShowDP, dbl_ShowDPSupplyPressure, dbl_ShowMaxDP, dbl_ShowMaxDPSupplyPressure, _
                str_Message)
        Else
            str_Message = INVALIDACTUATOR
        End If
    Else
            str_Message = INVALIDACTUATOR
    End If
End If
Exit Sub
SelectActuator_2600_Error:

End Sub
 Sub CalcDP_2600_Domotor_FTO_ATC(str_ActuatorModel As String, lng_FlowDirectionID As Long, _
    lng_AirActionID As Long, lng_ActuatorPartID As Long, lng_TrimPackageID As Long, _
    dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, dbl_InletPressure3 As Double, _
    dbl_InletPressure4 As Double, dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, dbl_InletTemperature1 As Double, _
    dbl_InletTemperature2 As Double, dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, _
    dbl_ShutOffDP As Double, str_SpecifiedPressureUM As String, _
    str_SpecifiedTemperatureUM As String, str_SpecifiedCalibPressureUM As String, _
    str_LengthUM As String, str_BodyMaterial As String, _
    str_SelectedRating As String, str_ValveLine As String, lng_TrimTypeID As Long, _
    str_PackingMaterial As String, str_LeakageClass As String, _
    str_SpringRange As String, dbl_RatedCv As Double, _
    dbl_ShowDP As Double, dbl_ShowDPSupplyPressure As Double, _
    dbl_ShowMaxDP As Double, dbl_ShowMaxDPSupplyPressure As Double, dbl_ShowPreload As Double, _
    dbl_ShowMaxPreload As Double, str_Message As String)
  
'--------------------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.

'This subroutine is used to calculate the selected valve and actuator shutoff capability and compare
'the result to defined requirements.
'This subroutine is specific to 2600 series valves with single stage trim, flow to open, air to close.

'---------------------------------------------------------------------------------------------------
'Input Variables
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff pressure drop
'str_SpecifiedPressureUM        unit of measure associated with inlet and outlet pressures
'str_specifiedTemperatureUM     unit of measure associated with inlet temperature
'str_specifiedCalibPressureUM   unit of measure associated with spring range and supply pressure
'str_BodyMaterial               selected material of body
'str_SelectedRating             selected body rating
'str_ValveLine                  selected product line
'lng_TrimTypeID                 selected type of trim
'str_PackingMaterial            selected packing material
'str_LeakageClass               selected leakage class
'str_SpringRange                selected actuator spring range
'dbl_RatedCv                    selected rated Cv

'---------------------------------------------------------------------------------------------------
'Output Variables
'dbl_ShowDP                     specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure       calculated required supply pressure for specified
'dbl_ShowMaxDP                  calculated maximum pressure drop for selected valve and actuator combination
'dbl_ShowMaxDPsupplyPressure    calculated required supply pressure for maximum pressure drop
'str_Message                    warning message if pressure drop requirement exceeds equipment capability
'---------------------------------------------------------------------------------------------------
On Error GoTo CalcDP_2600_Domotor_FTO_ATC_Error:
 
Dim bln_ActuatorClosedOK As Boolean
Dim bln_ActuatorOpenOK As Boolean
Dim bln_ActuatorStabilityOK As Boolean
Dim bln_ANSIOK As Boolean
Dim bln_StemOK As Boolean
Dim bln_SupplyOK As Boolean
Dim bln_TrimOK As Boolean
Dim dbl_ActuatorFriction As Double
Dim dbl_ActuatorStemDiameter As Double
Dim dbl_AllowableCompressiveForce As Double
Dim dbl_AllowableStemDP As Double
Dim dbl_AllowableStemLoad As Double
Dim dbl_AllowableTensileForce As Double
Dim dbl_AllowANSIPressure As Double
Dim dbl_AllowDP As Double
Dim dbl_AllowPlugDP As Double
Dim dbl_AreaFinal As Double
Dim dbl_AreaInitial As Double
Dim dbl_AreaOver As Double
Dim dbl_AreaUnder As Double
Dim dbl_As As Double
Dim dbl_DynamicFactor As Double
Dim dbl_E As Double
Dim dbl_ForcePacking As Double
Dim dbl_ForceSeating As Double
Dim dbl_ForceSpringClosed As Double
Dim dbl_ForceUnbalancedClosed As Double
Dim dbl_InletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double
Dim dbl_L As Double
Dim dbl_LeverLength As Double
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_MaxPressureDrop As Double
Dim dbl_MaxSupplyDP As Double
Dim dbl_MaxSupplyPressure As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_PackingHeight As Double
Dim dbl_PistonRingWidth As Double
Dim dbl_PlugDiameter As Double
Dim dbl_PositionerLossCoef As Double
Dim dbl_PreloadANSI As Double
Dim dbl_PreloadDP As Double
Dim dbl_PreloadPlugAllow As Double
Dim dbl_PreloadShutoff As Double
Dim dbl_PreloadStemAllow As Double
Dim dbl_PSupply As Double
Dim dbl_PSupplyANSI As Double
Dim dbl_PSupplyDP As Double
Dim dbl_PSupplyPlugAllow As Double
Dim dbl_PSupplyShutoff As Double
Dim dbl_PSupplyStemAllow As Double
Dim dbl_ReqANSIForceClosed As Double
Dim dbl_ReqDPForceClosed As Double
Dim dbl_ReqPlugLimitForceClosed As Double
Dim dbl_ReqShutoffForceClosed As Double
Dim dbl_RequiredPreload As Double
Dim dbl_RequiredSupply As Double
Dim dbl_Sd As Double
Dim dbl_SeatDiameter As Double
Dim dbl_ShowMaxInletPressure As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim dbl_ShowMaxPressureDrop As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_StemDiameter As Double
Dim dbl_StemSealWidth As Double
Dim dbl_StemStressArea As Double
Dim dbl_StemTravel As Double
Dim dbl_Sy As Double
Dim dbl_TempShowMaxDP As Double
Dim dbl_theMaxPreload As Double
Dim dbl_theMaxSupply As Double
Dim dbl_VolOver As Double
Dim int_QtySeals As Integer
Dim long_AirActionID As Long
Dim str_StemMaterial As String

#If (DebugVersion = "yes") Then
MsgBox "2600-FTO-ATC Actuator" + Chr(10) + Chr(13) + _
     "str_ActuatorModel=" + str_ActuatorModel + Chr(10) + Chr(13) + _
     "lng_FlowDirectionID=" + CStr(lng_FlowDirectionID) + Chr(10) + Chr(13) + _
     "lng_AirActionID=" + CStr(lng_AirActionID) + Chr(10) + Chr(13) + _
     "lng_TrimPackageID=" + CStr(lng_TrimPackageID) + Chr(10) + Chr(13) + _
     "lng_ActuatorPartID=" + CStr(lng_ActuatorPartID) + Chr(10) + Chr(13) + _
     "dbl_InletPressure1=" + CStr(dbl_InletPressure1) + Chr(10) + Chr(13) + _
     "dbl_OutletPressure1=" + CStr(dbl_OutletPressure1) + Chr(10) + Chr(13) + _
     "dbl_InletTemperature1=" + CStr(dbl_InletTemperature1) + Chr(10) + Chr(13) + _
     "dbl_ShutOffDP=" + CStr(dbl_ShutOffDP) + Chr(10) + Chr(13) + _
     "str_SpecifiedPressureUM=" + str_SpecifiedPressureUM + Chr(10) + Chr(13) + _
     "str_SpecifiedTemperatureUM=" + str_SpecifiedTemperatureUM + Chr(10) + Chr(13) + _
     "str_SpecifiedCalibPressureUM=" + str_SpecifiedCalibPressureUM + Chr(10) + Chr(13) + _
     "str_LengthUM=" + str_LengthUM + Chr(10) + Chr(13) + _
     "str_BodyMaterial=" + str_BodyMaterial + Chr(10) + Chr(13) + _
     "str_SelectedRating=" + str_SelectedRating + Chr(10) + Chr(13) + _
     "str_ValveLine=" + str_ValveLine + Chr(10) + Chr(13) + _
     "lng_TrimTypeID=" + CStr(lng_TrimTypeID) + Chr(10) + Chr(13) + _
     "str_PackingMaterial=" + str_PackingMaterial + Chr(10) + Chr(13) + _
     "str_LeakageClass=" + str_LeakageClass + Chr(10) + Chr(13) + _
     "str_SpringRange=" + CStr(str_SpringRange) + Chr(10) + Chr(13) + _
     "dbl_RatedCv=" + CStr(dbl_RatedCv)
#End If

'***************************
'Parameter validation
'***************************
If (dbl_ShutOffDP < 0#) Then GoTo CalcDP_2600_Domotor_FTO_ATC_Error:
If (str_ValveLine = "") Then GoTo CalcDP_2600_Domotor_FTO_ATC_Error:
If (str_LengthUM = "") Then GoTo CalcDP_2600_Domotor_FTO_ATC_Error:
If (str_SpecifiedPressureUM = "") Then GoTo CalcDP_2600_Domotor_FTO_ATC_Error:
If (str_SpecifiedTemperatureUM = "") Then GoTo CalcDP_2600_Domotor_FTO_ATC_Error:
If (str_SpecifiedCalibPressureUM = "") Then GoTo CalcDP_2600_Domotor_FTO_ATC_Error:
If (str_PackingMaterial = "") Then GoTo CalcDP_2600_Domotor_FTO_ATC_Error:
If (str_LeakageClass = "") Then GoTo CalcDP_2600_Domotor_FTO_ATC_Error:
If (str_SpringRange = "") Then GoTo CalcDP_2600_Domotor_FTO_ATC_Error:

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4



'-Determine Requirements------------------------------
Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature(), dbl_MaxInletTemperature)
Call CalcTestPressureDrop(dbl_InletPressure(), dbl_OutletPressure(), dbl_MaxPressureDrop)

dbl_ShowMaxInletPressure = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, str_SpecifiedPressureUM)

dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, str_SpecifiedTemperatureUM)
dbl_ShowMaxPressureDrop = dbl_MaxPressureDrop
dbl_MaxPressureDrop = ConvertPressureToPSIG(dbl_MaxPressureDrop, str_SpecifiedPressureUM)
dbl_ShowShutoffDP = dbl_ShutOffDP
dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, str_SpecifiedPressureUM)

'-Determine ANSI limitation------------------------
Call GetANSIPressure(str_BodyMaterial, dbl_MaxInletTemperature, str_SelectedRating, _
    dbl_AllowANSIPressure)
'-Read Valve Properties------------------------------
Call Get2600Data(lng_TrimPackageID, dbl_DynamicFactor, dbl_PackingHeight, _
    dbl_PlugDiameter, dbl_SeatDiameter, dbl_StemDiameter, dbl_AllowPlugDP, _
    str_StemMaterial, dbl_StemTravel)
    
'-Calculate Required Packing Force---------------
Call CalcPackingForce(dbl_MaxInletPressure, dbl_StemDiameter, str_PackingMaterial, _
    dbl_PackingHeight, dbl_ForcePacking)
    
'-Calculate Required Seating Force-----------------
Call CalcSeatingForce(str_LeakageClass, dbl_SeatDiameter, dbl_ForceSeating)

'-Calculate Required Offbalance Force--------------
Call CalcSingleSeatUnbalancedForceFTO(dbl_MaxInletPressure, dbl_SeatDiameter, dbl_ForceUnbalancedClosed)
dbl_ReqDPForceClosed = dbl_ForcePacking + dbl_ForceSeating + dbl_ForceUnbalancedClosed

Call CalcSingleSeatUnbalancedForceFTO(dbl_AllowANSIPressure, dbl_SeatDiameter, dbl_ForceUnbalancedClosed)
dbl_ReqANSIForceClosed = dbl_ForcePacking + dbl_ForceSeating + dbl_ForceUnbalancedClosed

Call CalcSingleSeatUnbalancedForceFTO(dbl_AllowPlugDP, dbl_SeatDiameter, dbl_ForceUnbalancedClosed)
dbl_ReqPlugLimitForceClosed = dbl_ForcePacking + dbl_ForceSeating + dbl_ForceUnbalancedClosed

Call CalcSingleSeatUnbalancedForceFTO(dbl_ShutOffDP, dbl_SeatDiameter, dbl_ForceUnbalancedClosed)
dbl_ReqShutoffForceClosed = dbl_ForcePacking + dbl_ForceSeating + dbl_ForceUnbalancedClosed

'-calculate stem limitation----------------------
dbl_L = 6#
Call GetStemStressArea(dbl_StemDiameter, dbl_StemStressArea)
Call GetStemProperties(str_StemMaterial, dbl_MaxInletTemperature, dbl_Sy, dbl_E, dbl_Sd)
Call CalcAllowStemForce(dbl_Sy, dbl_E, dbl_StemDiameter, dbl_L, dbl_Sd, _
    dbl_StemStressArea, dbl_AllowableCompressiveForce)
Call CalcAllowStemTensileForce(dbl_Sd, dbl_StemStressArea, dbl_AllowableTensileForce)

If dbl_AllowableTensileForce < dbl_AllowableCompressiveForce Then
    dbl_AllowableStemLoad = dbl_AllowableTensileForce
Else
    dbl_AllowableStemLoad = dbl_AllowableCompressiveForce
End If

'-calculate required supply pressure-----------------
Call GetActuatorData(lng_ActuatorPartID, dbl_AreaOver, dbl_AreaUnder, dbl_MaxSupplyPressure, _
    dbl_ActuatorFriction, dbl_PositionerLossCoef, str_ActuatorModel, long_AirActionID, int_QtySeals, _
    dbl_StemSealWidth, dbl_PistonRingWidth, dbl_VolOver, dbl_ForceSpringClosed, _
    dbl_ActuatorStemDiameter, dbl_LeverLength)

Call CalcReqDomotorPreloadSupply_FTO_ATC(dbl_ReqANSIForceClosed, dbl_ForceSpringClosed, _
    dbl_AreaOver, dbl_PreloadANSI, dbl_PSupplyANSI)

Call CalcReqDomotorPreloadSupply_FTO_ATC(dbl_AllowableStemLoad, dbl_ForceSpringClosed, dbl_AreaOver, _
    dbl_PreloadStemAllow, dbl_PSupplyStemAllow)

Call CalcReqDomotorPreloadSupply_FTO_ATC(dbl_ReqDPForceClosed, dbl_ForceSpringClosed, dbl_AreaOver, _
    dbl_PreloadDP, dbl_PSupplyDP)

Call CalcReqDomotorPreloadSupply_FTO_ATC(dbl_ReqShutoffForceClosed, dbl_ForceSpringClosed, dbl_AreaOver, _
    dbl_PreloadShutoff, dbl_PSupplyShutoff)

'find max allow preload/supply
dbl_theMaxPreload = dbl_PreloadANSI
dbl_theMaxSupply = dbl_PSupplyANSI
If (dbl_PreloadStemAllow < dbl_theMaxPreload) Then
    dbl_theMaxPreload = dbl_PreloadStemAllow
    dbl_theMaxSupply = dbl_PSupplyStemAllow
End If

'-Calculate Max DP with Max Preload--------------------
If dbl_theMaxPreload > 0 Then
    Call CalcDPAllowableDomotor_FTO_ATC(dbl_theMaxPreload, dbl_AreaOver, dbl_ForceSpringClosed, _
        dbl_ForceSeating, dbl_ForcePacking, dbl_SeatDiameter, dbl_AllowDP)
Else
    dbl_AllowDP = 0
End If

'-calculate allowable pressure drop for stem-----------
Call CalcDPAllowableDomotor_FTO_ATC(dbl_PreloadStemAllow, dbl_AreaOver, dbl_ForceSpringClosed, _
    dbl_ForceSeating, dbl_ForcePacking, dbl_SeatDiameter, dbl_AllowableStemDP)
    
'-Evaluation------------------
bln_ANSIOK = True
bln_StemOK = True
bln_ActuatorOpenOK = True
bln_ActuatorClosedOK = True
bln_ActuatorStabilityOK = True
bln_TrimOK = True
bln_SupplyOK = True

If dbl_MaxInletPressure > dbl_AllowANSIPressure Then
    bln_ANSIOK = False
    str_Message = ANSIMESSAGE
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowANSIPressure, str_SpecifiedPressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyANSI, _
        str_SpecifiedCalibPressureUM)
    dbl_ShowMaxPreload = ConvertPressureFromPSIG(dbl_PreloadANSI, _
        str_SpecifiedCalibPressureUM)
End If

If dbl_MaxInletPressure > dbl_AllowableStemDP Then
    bln_StemOK = False
    str_Message = STEMMESSAGE
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowableStemDP, str_SpecifiedPressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyStemAllow, _
        str_SpecifiedCalibPressureUM)
    dbl_ShowMaxPreload = ConvertPressureFromPSIG(dbl_PreloadStemAllow, _
        str_SpecifiedCalibPressureUM)
End If

If dbl_PSupplyDP > dbl_MaxSupplyPressure And dbl_PSupplyShutoff <= dbl_MaxSupplyPressure Then
    bln_SupplyOK = False
    str_Message = LIMITEDDPMESSAGE
    dbl_ShowMaxDP = dbl_ShowMaxPressureDrop
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, _
        str_SpecifiedCalibPressureUM)
    dbl_ShowMaxPreload = ConvertPressureFromPSIG(dbl_PreloadDP, _
        str_SpecifiedCalibPressureUM)
        
ElseIf (dbl_PSupply > dbl_MaxSupplyPressure) And (dbl_PSupplyShutoff > dbl_MaxSupplyPressure) Then
    bln_SupplyOK = False
    str_Message = SUPPLYMESSAGE
    dbl_ShowDP = 0
    dbl_ShowMaxDPSupplyPressure = 0
    dbl_ShowMaxPreload = 0
    
End If

If (bln_ANSIOK = True) And (bln_StemOK = True) And (bln_SupplyOK = True) Then
    dbl_ShowDP = dbl_ShowMaxInletPressure
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_SpecifiedCalibPressureUM)
    dbl_ShowPreload = ConvertPressureFromPSIG(dbl_PreloadDP, str_SpecifiedCalibPressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowDP, str_SpecifiedPressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_theMaxSupply, str_SpecifiedCalibPressureUM)
    dbl_ShowMaxPreload = ConvertPressureFromPSIG(dbl_theMaxPreload, str_SpecifiedCalibPressureUM)
End If

dbl_ShowDP = TrimDecimalPlaces(dbl_ShowDP)
dbl_ShowDPSupplyPressure = TrimDecimalPlaces(dbl_ShowDPSupplyPressure)
dbl_ShowMaxDP = TrimDecimalPlaces(dbl_ShowMaxDP)
dbl_ShowMaxDPSupplyPressure = TrimDecimalPlaces(dbl_ShowMaxDPSupplyPressure)
dbl_ShowPreload = TrimDecimalPlaces(dbl_ShowPreload)
dbl_ShowMaxPreload = TrimDecimalPlaces(dbl_ShowMaxPreload)

#If (DebugVersion = "yes") Then
    MsgBox "2600-FTO-ATC Actuator" + Chr(10) + Chr(13) + _
    "dbl_ShowDP=" + CStr(dbl_ShowDP) + Chr(10) + Chr(13) + _
    "dbl_ShowDPSupplyPressure=" + CStr(dbl_ShowDPSupplyPressure) + Chr(10) + Chr(13) + _
    "dbl_ShowMaxDP=" + CStr(dbl_ShowMaxDP) + Chr(10) + Chr(13) + _
    "dbl_ShowMaxDPSupplyPressure=" + CStr(dbl_ShowMaxDPSupplyPressure) + Chr(10) + Chr(13) + _
    "dbl_ShowPreload=" + CStr(dbl_ShowPreload) + Chr(10) + Chr(13) + _
    "dbl_ShowMaxPreload=" + CStr(dbl_ShowMaxPreload) + Chr(10) + Chr(13) + _
    "str_Message=" + str_Message + Chr(10) + Chr(13)
#End If


Exit Sub
CalcDP_2600_Domotor_FTO_ATC_Error:

End Sub
 Sub CalcDP_2600_Domotor_FTC_ATO(str_ActuatorModel As String, lng_FlowDirectionID As Long, _
    lng_AirActionID As Long, lng_ActuatorPartID As Long, lng_TrimPackageID As Long, _
    dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, dbl_InletPressure3 As Double, _
    dbl_InletPressure4 As Double, dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, dbl_InletTemperature1 As Double, _
    dbl_InletTemperature2 As Double, dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, _
    dbl_ShutOffDP As Double, str_SpecifiedPressureUM As String, _
    str_SpecifiedTemperatureUM As String, str_SpecifiedCalibPressureUM As String, _
    str_LengthUM As String, str_BodyMaterial As String, _
    str_SelectedRating As String, str_ValveLine As String, lng_TrimTypeID As Long, _
    str_PackingMaterial As String, str_LeakageClass As String, _
    str_SpringRange As String, dbl_RatedCv As Double, _
    dbl_ShowDP As Double, dbl_ShowDPSupplyPressure As Double, _
    dbl_ShowMaxDP As Double, dbl_ShowMaxDPSupplyPressure As Double, dbl_ShowPreload As Double, _
    dbl_ShowMaxPreload As Double, str_Message As String)
'--------------------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.

'This subroutine is used to calculate the selected valve and actuator shutoff capability and compare
'the result to defined requirements.
'This subroutine is specific to 2600 series valves with single stage trim, flow to open, air to close.

'---------------------------------------------------------------------------------------------------
'Input Variables
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff pressure drop
'str_SpecifiedPressureUM        unit of measure associated with inlet and outlet pressures
'str_specifiedTemperatureUM     unit of measure associated with inlet temperature
'str_specifiedCalibPressureUM   unit of measure associated with spring range and supply pressure
'str_BodyMaterial               selected material of body
'str_SelectedRating             selected body rating
'str_ValveLine                  selected product line
'lng_TrimTypeID                 selected type of trim
'str_PackingMaterial            selected packing material
'str_LeakageClass               selected leakage class
'str_SpringRange                selected actuator spring range
'dbl_RatedCv                    selected rated Cv
'---------------------------------------------------------------------------------------------------
'Output Variables
'dbl_ShowDP                     specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure       calculated required supply pressure for specified
'dbl_ShowMaxDP                  calculated maximum pressure drop for selected valve and actuator combination
'dbl_ShowMaxDPsupplyPressure    calculated required supply pressure for maximum pressure drop
'str_Message                    warning message if pressure drop requirement exceeds equipment capability
'---------------------------------------------------------------------------------------------------
On Error GoTo CalcDP_2600_Domotor_FTC_ATO_Error:

Dim bln_ActuatorClosedOK As Boolean
Dim bln_ActuatorOpenOK As Boolean
Dim bln_ActuatorStabilityOK As Boolean
Dim bln_ANSIOK As Boolean
Dim bln_StemOK As Boolean
Dim bln_SupplyOK As Boolean
Dim bln_TrimOK As Boolean
Dim dbl_ActuatorFriction As Double
Dim dbl_ActuatorStemDiameter As Double
Dim dbl_AllowableCompressiveForce As Double
Dim dbl_AllowableStemDP As Double
Dim dbl_AllowableStemLoad As Double
Dim dbl_AllowableTensileForce As Double
Dim dbl_AllowANSIPressure As Double
Dim dbl_AllowDP As Double
Dim dbl_AllowPlugDP As Double
Dim dbl_AreaFinal As Double
Dim dbl_AreaInitial As Double
Dim dbl_AreaOver As Double
Dim dbl_AreaUnder As Double
Dim dbl_DomotorForce_ATO As Double
Dim dbl_DynamicFactor As Double
Dim dbl_E As Double
Dim dbl_ForcePacking As Double
Dim dbl_ForceSeating As Double
Dim dbl_ForceSpringClosed As Double
Dim dbl_ForceUnbalancedClosed As Double
Dim dbl_InletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double
Dim dbl_L As Double
Dim dbl_LeverLength As Double
Dim dbl_MaxAllowDP As Double
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_MaxPressureDrop As Double
Dim dbl_MaxSupplyDP As Double
Dim dbl_MaxSupplyPressure As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_PackingHeight As Double
Dim dbl_PistonRingWidth As Double
Dim dbl_PlugDiameter As Double
Dim dbl_PositionerLossCoef As Double
Dim dbl_PreloadANSI As Double
Dim dbl_PreloadDP As Double
Dim dbl_PreloadPlugAllow As Double
Dim dbl_PreloadShutoffDP As Double
Dim dbl_PreloadStemAllow As Double
Dim dbl_PSupply As Double
Dim dbl_PSupplyANSI As Double
Dim dbl_PSupplyDP As Double
Dim dbl_PSupplyPlugAllow As Double
Dim dbl_PSupplyShutoffDP As Double
Dim dbl_PSupplyStemAllow As Double
Dim dbl_ReqANSIForceClosed As Double
Dim dbl_ReqDPForceClosed As Double
Dim dbl_ReqPlugLimitForceClosed As Double
Dim dbl_ReqPreload As Double
Dim dbl_ReqShutoffForceClosed As Double
Dim dbl_ReqSupply As Double
Dim dbl_RequiredPreload As Double
Dim dbl_RequiredSupply As Double
Dim dbl_Sd As Double
Dim dbl_SeatDiameter As Double
Dim dbl_ShowMaxInletPressure As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim dbl_ShowMaxPressureDrop As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_StemDiameter As Double
Dim dbl_StemSealWidth As Double
Dim dbl_StemStressArea As Double
Dim dbl_Sy As Double
Dim dbl_TempShowMaxDP As Double
Dim dbl_theMaxPreload As Double
Dim dbl_Travel As Double
Dim dbl_VolOver As Double
Dim int_QtySeals As Integer
Dim long_AirActionID As Long
Dim str_StemMaterial As String

#If (DebugVersion = "yes") Then
MsgBox "2600-FTC-ATO Actuator" + Chr(10) + Chr(13) + _
     "str_ActuatorModel=" + str_ActuatorModel + Chr(10) + Chr(13) + _
     "lng_FlowDirectionID=" + CStr(lng_FlowDirectionID) + Chr(10) + Chr(13) + _
     "lng_AirActionID=" + CStr(lng_AirActionID) + Chr(10) + Chr(13) + _
     "lng_TrimPackageID=" + CStr(lng_TrimPackageID) + Chr(10) + Chr(13) + _
     "lng_ActuatorPartID=" + CStr(lng_ActuatorPartID) + Chr(10) + Chr(13) + _
     "dbl_InletPressure1=" + CStr(dbl_InletPressure1) + Chr(10) + Chr(13) + _
     "dbl_OutletPressure1=" + CStr(dbl_OutletPressure1) + Chr(10) + Chr(13) + _
     "dbl_InletTemperature1=" + CStr(dbl_InletTemperature1) + Chr(10) + Chr(13) + _
     "dbl_ShutOffDP=" + CStr(dbl_ShutOffDP) + Chr(10) + Chr(13) + _
     "str_SpecifiedPressureUM=" + str_SpecifiedPressureUM + Chr(10) + Chr(13) + _
     "str_SpecifiedTemperatureUM=" + str_SpecifiedTemperatureUM + Chr(10) + Chr(13) + _
     "str_SpecifiedCalibPressureUM=" + str_SpecifiedCalibPressureUM + Chr(10) + Chr(13) + _
     "str_LengthUM=" + str_LengthUM + Chr(10) + Chr(13) + _
     "str_BodyMaterial=" + str_BodyMaterial + Chr(10) + Chr(13) + _
     "str_SelectedRating=" + str_SelectedRating + Chr(10) + Chr(13) + _
     "str_ValveLine=" + str_ValveLine + Chr(10) + Chr(13) + _
     "lng_TrimTypeID=" + CStr(lng_TrimTypeID) + Chr(10) + Chr(13) + _
     "str_PackingMaterial=" + str_PackingMaterial + Chr(10) + Chr(13) + _
     "str_LeakageClass=" + str_LeakageClass + Chr(10) + Chr(13) + _
     "str_SpringRange=" + CStr(str_SpringRange) + Chr(10) + Chr(13) + _
     "dbl_RatedCv=" + CStr(dbl_RatedCv)
#End If

'***************************
'Parameter validation
'***************************
If (dbl_ShutOffDP < 0#) Then GoTo CalcDP_2600_Domotor_FTC_ATO_Error:
If (str_ValveLine = "") Then GoTo CalcDP_2600_Domotor_FTC_ATO_Error:
If (str_LengthUM = "") Then GoTo CalcDP_2600_Domotor_FTC_ATO_Error:
If (str_SpecifiedPressureUM = "") Then GoTo CalcDP_2600_Domotor_FTC_ATO_Error:
If (str_SpecifiedTemperatureUM = "") Then GoTo CalcDP_2600_Domotor_FTC_ATO_Error:
If (str_SpecifiedCalibPressureUM = "") Then GoTo CalcDP_2600_Domotor_FTC_ATO_Error:
If (str_PackingMaterial = "") Then GoTo CalcDP_2600_Domotor_FTC_ATO_Error:
If (str_LeakageClass = "") Then GoTo CalcDP_2600_Domotor_FTC_ATO_Error:
If (str_SpringRange = "") Then GoTo CalcDP_2600_Domotor_FTC_ATO_Error:

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4

'-determine requirements----------------------------------------------------

Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature(), dbl_MaxInletTemperature)
Call CalcTestPressureDrop(dbl_InletPressure(), dbl_OutletPressure(), dbl_MaxPressureDrop)

dbl_ShowMaxInletPressure = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, str_SpecifiedPressureUM)

dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, str_SpecifiedTemperatureUM)
dbl_ShowMaxPressureDrop = dbl_MaxPressureDrop
dbl_MaxPressureDrop = ConvertPressureToPSIG(dbl_MaxPressureDrop, str_SpecifiedPressureUM)
dbl_ShowShutoffDP = dbl_ShutOffDP
dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, str_SpecifiedPressureUM)

'-determine ANSI limitation----------------------------------------------------

Call GetANSIPressure(str_BodyMaterial, dbl_MaxInletTemperature, str_SelectedRating, _
    dbl_AllowANSIPressure)

'-read valve properties---------------------------------------------------------

Call Get2600Data(lng_TrimPackageID, dbl_DynamicFactor, dbl_PackingHeight, dbl_PlugDiameter, _
    dbl_SeatDiameter, dbl_StemDiameter, dbl_AllowPlugDP, str_StemMaterial, dbl_Travel)

Call CalcPackingForce(dbl_MaxInletPressure, dbl_StemDiameter, str_PackingMaterial, dbl_PackingHeight, _
    dbl_ForcePacking)

'-calculate stem limitation-------------------------------------------------------------

dbl_L = 6#
Call GetStemStressArea(dbl_StemDiameter, dbl_StemStressArea)
Call GetStemProperties(str_StemMaterial, dbl_MaxInletTemperature, dbl_Sy, dbl_E, dbl_Sd)
Call CalcAllowStemForce(dbl_Sy, dbl_E, dbl_StemDiameter, dbl_L, dbl_Sd, _
    dbl_StemStressArea, dbl_AllowableCompressiveForce)
Call CalcAllowStemTensileForce(dbl_Sd, dbl_StemStressArea, dbl_AllowableTensileForce)

If dbl_AllowableTensileForce < dbl_AllowableCompressiveForce Then
    dbl_AllowableStemLoad = dbl_AllowableTensileForce
Else
    dbl_AllowableStemLoad = dbl_AllowableCompressiveForce
End If

'-calculate required supply presure--------------------------------------------------
Call GetActuatorData(lng_ActuatorPartID, dbl_AreaOver, dbl_AreaUnder, dbl_MaxSupplyPressure, _
    dbl_ActuatorFriction, dbl_PositionerLossCoef, str_ActuatorModel, long_AirActionID, int_QtySeals, _
    dbl_StemSealWidth, dbl_PistonRingWidth, dbl_VolOver, dbl_ForceSpringClosed, _
    dbl_ActuatorStemDiameter, dbl_LeverLength)

Call CalcReqDomotorPreloadSupply_FTC_ATO(dbl_MaxInletPressure, dbl_AreaOver, _
    dbl_AreaUnder, dbl_VolOver, dbl_Travel, dbl_SeatDiameter, dbl_DynamicFactor, _
    dbl_ForcePacking, dbl_ReqPreload, dbl_ReqSupply)

Call CalcAllowDPDomotor_FTC_ATO(dbl_AreaOver, dbl_AreaUnder, dbl_VolOver, _
    dbl_Travel, dbl_SeatDiameter, dbl_DynamicFactor, dbl_ForcePacking, dbl_ReqPreload, _
    dbl_ReqSupply, dbl_MaxAllowDP)

Call CalcDomotorForce_ATO(dbl_AreaOver, dbl_AreaUnder, dbl_ReqPreload, dbl_ReqSupply, _
    dbl_DomotorForce_ATO)
 
'-evaluation--------------------------------------------------------------

bln_ANSIOK = True
bln_StemOK = True
bln_ActuatorOpenOK = True
bln_ActuatorClosedOK = True
bln_ActuatorStabilityOK = True
bln_TrimOK = True
bln_SupplyOK = True

dbl_ShowPreload = 0
dbl_ShowDPSupplyPressure = 0
dbl_ShowDP = 0
dbl_ShowMaxDPSupplyPressure = 0
dbl_ShowMaxDP = dbl_ShowMaxInletPressure
dbl_ShowMaxPreload = 0


If dbl_MaxInletPressure > dbl_AllowANSIPressure Then
    bln_ANSIOK = False
    str_Message = ANSIMESSAGE
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowANSIPressure, str_SpecifiedPressureUM)
End If


If dbl_DomotorForce_ATO > dbl_AllowableStemLoad Then
    bln_StemOK = False
    str_Message = STEMMESSAGE
    dbl_ShowMaxDP = 0
End If
 
If dbl_ReqSupply > dbl_MaxSupplyPressure Then
    bln_SupplyOK = False
    str_Message = SUPPLYMESSAGE
    dbl_ShowMaxDP = 0
End If

If bln_ANSIOK = True And bln_StemOK = True And bln_SupplyOK = True Then
    dbl_ShowDP = dbl_ShowMaxInletPressure
    dbl_PSupplyDP = dbl_ReqSupply
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_SpecifiedCalibPressureUM)
    dbl_ShowPreload = ConvertPressureFromPSIG(dbl_ReqPreload, str_SpecifiedCalibPressureUM)
    
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, str_SpecifiedPressureUM)
    dbl_ShowMaxDPSupplyPressure = dbl_ShowDPSupplyPressure
    dbl_ShowMaxPreload = dbl_ShowPreload

End If

dbl_ShowDP = TrimDecimalPlaces(dbl_ShowDP)
dbl_ShowDPSupplyPressure = TrimDecimalPlaces(dbl_ShowDPSupplyPressure)
dbl_ShowMaxDP = TrimDecimalPlaces(dbl_ShowMaxDP)
dbl_ShowMaxDPSupplyPressure = TrimDecimalPlaces(dbl_ShowMaxDPSupplyPressure)
dbl_ShowPreload = TrimDecimalPlaces(dbl_ShowPreload)
dbl_ShowMaxPreload = TrimDecimalPlaces(dbl_ShowMaxPreload)

#If (DebugVersion = "yes") Then
    MsgBox "2600-FTC-ATO Actuator" + Chr(10) + Chr(13) + _
    "dbl_ShowDP=" + CStr(dbl_ShowDP) + Chr(10) + Chr(13) + _
    "dbl_ShowDPSupplyPressure=" + CStr(dbl_ShowDPSupplyPressure) + Chr(10) + Chr(13) + _
    "dbl_ShowMaxDP=" + CStr(dbl_ShowMaxDP) + Chr(10) + Chr(13) + _
    "dbl_ShowMaxDPSupplyPressure=" + CStr(dbl_ShowMaxDPSupplyPressure) + Chr(10) + Chr(13) + _
    "dbl_ShowPreload=" + CStr(dbl_ShowPreload) + Chr(10) + Chr(13) + _
    "dbl_ShowMaxPreload=" + CStr(dbl_ShowMaxPreload) + Chr(10) + Chr(13) + _
    "str_Message=" + str_Message + Chr(10) + Chr(13)
#End If

Exit Sub
CalcDP_2600_Domotor_FTC_ATO_Error:

End Sub

 Sub CalcDP_2600_Cylinder_FTO_ATO(str_ActuatorModel As String, lng_FlowDirectionID As Long, _
    lng_AirActionID As Long, lng_ActuatorPartID As Long, lng_TrimPackageID As Long, _
    dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, dbl_InletPressure3 As Double, _
    dbl_InletPressure4 As Double, dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, dbl_InletTemperature1 As Double, _
    dbl_InletTemperature2 As Double, dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, _
    dbl_ShutOffDP As Double, str_SpecifiedPressureUM As String, _
    str_SpecifiedTemperatureUM As String, str_SpecifiedCalibPressureUM As String, _
    str_LengthUM As String, str_BodyMaterial As String, _
    str_SelectedRating As String, str_ValveLine As String, lng_TrimTypeID As Long, _
    str_PackingMaterial As String, str_LeakageClass As String, _
    str_SpringRange As String, dbl_RatedCv As Double, _
    dbl_ShowDP As Double, dbl_ShowDPSupplyPressure As Double, _
    dbl_ShowMaxDP As Double, dbl_ShowMaxDPSupplyPressure As Double, _
    str_Message As String)
'--------------------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.

'This subroutine is used to calculate the selected valve and actuator shutoff capability and compare
'the result to defined requirements.
'This subroutine is specific to 2600 series valves with single stage trim, flow to open, air to close.

'---------------------------------------------------------------------------------------------------
'Input Variables
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff pressure drop
'str_SpecifiedPressureUM        unit of measure associated with inlet and outlet pressures
'str_SpecifiedTemperatureUM     unit of measure associated with inlet temperature
'str_SpecifiedCalibPressureUM   unit of measure associated with spring range and supply pressure
'str_BodyMaterial               selected material of body
'str_SelectedRating             selected body rating
'str_ValveLine                  selected product line
'lng_TrimTypeID                 selected type of trim
'str_PackingMaterial            selected packing material
'str_LeakageClass               selected leakage class
'str_SpringRange                selected actuator spring range
'dbl_RatedCv                    selected rated Cv
'---------------------------------------------------------------------------------------------------
'Output Variables
'dbl_ShowDP                     specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure       calculated required supply pressure for specified
'dbl_ShowMaxDP                  calculated maximum pressure drop for selected valve and actuator combination
'dbl_ShowMaxDPsupplyPressure    calculated required supply pressure for maximum pressure drop
'str_Message                    warning message if pressure drop requirement exceeds equipment capability
'---------------------------------------------------------------------------------------------------
On Error GoTo CalcDP_2600_Cylinder_FTO_ATO_Error:

Dim bln_ActuatorClosedOK As Boolean
Dim bln_ActuatorOK As Boolean
Dim bln_ActuatorOpenOK As Boolean
Dim bln_ActuatorStabilityOK As Boolean
Dim bln_ANSIOK As Boolean
Dim bln_StemOK As Boolean
Dim bln_SupplyOK As Boolean
Dim bln_TrimOK As Boolean
Dim dbl_ActuatorForce As Double
Dim dbl_ActuatorFriction As Double
Dim dbl_ActuatorStemDiameter As Double
Dim dbl_AllowableCompressiveForce As Double
Dim dbl_AllowableStemDP As Double
Dim dbl_AllowableStemLoad As Double
Dim dbl_AllowableTensileForce As Double
Dim dbl_AllowANSIPressure As Double
Dim dbl_AllowDP As Double
Dim dbl_AllowDPClosed As Double
Dim dbl_AllowDPOpen As Double
Dim dbl_AllowDPStability As Double
Dim dbl_AllowPlugDP As Double
Dim dbl_AreaFinal As Double
Dim dbl_AreaInitial As Double
Dim dbl_AreaOver As Double
Dim dbl_AreaUnder As Double
Dim dbl_As As Double
Dim dbl_DynamicFactor As Double
Dim dbl_E As Double
Dim dbl_ForceFriction As Double
Dim dbl_ForcePacking As Double
Dim dbl_ForceSeating As Double
Dim dbl_ForceSpringClosed As Double
Dim dbl_ForceUnbalancedClosed As Double
Dim dbl_ForceUnbalancedOpen As Double
Dim dbl_InletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double
Dim dbl_L As Double
Dim dbl_LeverLength As Double
Dim dbl_MaxAllowDP As Double
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_MaxPressureDrop As Double
Dim dbl_MaxSupplyDP As Double
Dim dbl_MaxSupplyPressure As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_PackingHeight As Double
Dim dbl_PFinal As Double
Dim dbl_PInitial As Double
Dim dbl_PistonRingWidth As Double
Dim dbl_PlugDiameter As Double
Dim dbl_PositionerLossCoef As Double
Dim dbl_PreloadANSI As Double
Dim dbl_PreloadDP As Double
Dim dbl_PreloadPlugAllow As Double
Dim dbl_PreloadShutoffDP As Double
Dim dbl_PreloadStemAllow As Double
Dim dbl_PSupply As Double
Dim dbl_PSupplyANSI As Double
Dim dbl_PSupplyDP As Double
Dim dbl_PSupplyPlugAllow As Double
Dim dbl_PSupplyShutoffDP As Double
Dim dbl_PSupplyStemAllow As Double
Dim dbl_ReqANSIForceClosed As Double
Dim dbl_ReqDPForceClosed As Double
Dim dbl_ReqForceClosed As Double
Dim dbl_ReqForceOpen As Double
Dim dbl_ReqPlugLimitForceClosed As Double
Dim dbl_ReqShutoffForceClosed As Double
Dim dbl_RequiredPreload As Double
Dim dbl_RequiredSupply As Double
Dim dbl_RequiredSupplyPressure As Double
Dim dbl_Sd As Double
Dim dbl_SeatDiameter As Double
Dim dbl_ShowMaxInletPressure As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim dbl_ShowMaxPressureDrop As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_SpringRate As Double
Dim dbl_StemDiameter As Double
Dim dbl_StemSealWidth As Double
Dim dbl_StemStressArea As Double
Dim dbl_StemTravel As Double
Dim dbl_Sy As Double
Dim dbl_TempShowMaxDP As Double
Dim dbl_theMaxPreload As Double
Dim dbl_Travel As Double
Dim dbl_VolOver As Double
Dim int_QtySeals As Integer
Dim long_AirActionID As Long
Dim str_StemMaterial As String

#If (DebugVersion = "yes") Then
MsgBox "2600-Cylinder-FTO-ATO Actuator" + Chr(10) + Chr(13) + _
     "str_ActuatorModel=" + str_ActuatorModel + Chr(10) + Chr(13) + _
     "lng_FlowDirectionID=" + CStr(lng_FlowDirectionID) + Chr(10) + Chr(13) + _
     "lng_AirActionID=" + CStr(lng_AirActionID) + Chr(10) + Chr(13) + _
     "lng_TrimPackageID=" + CStr(lng_TrimPackageID) + Chr(10) + Chr(13) + _
     "lng_ActuatorPartID=" + CStr(lng_ActuatorPartID) + Chr(10) + Chr(13) + _
     "dbl_InletPressure1=" + CStr(dbl_InletPressure1) + Chr(10) + Chr(13) + _
     "dbl_OutletPressure1=" + CStr(dbl_OutletPressure1) + Chr(10) + Chr(13) + _
     "dbl_InletTemperature1=" + CStr(dbl_InletTemperature1) + Chr(10) + Chr(13) + _
     "dbl_ShutOffDP=" + CStr(dbl_ShutOffDP) + Chr(10) + Chr(13) + _
     "str_SpecifiedPressureUM=" + str_SpecifiedPressureUM + Chr(10) + Chr(13) + _
     "str_SpecifiedTemperatureUM=" + str_SpecifiedTemperatureUM + Chr(10) + Chr(13) + _
     "str_SpecifiedCalibPressureUM=" + str_SpecifiedCalibPressureUM + Chr(10) + Chr(13) + _
     "str_LengthUM=" + str_LengthUM + Chr(10) + Chr(13) + _
     "str_BodyMaterial=" + str_BodyMaterial + Chr(10) + Chr(13) + _
     "str_SelectedRating=" + str_SelectedRating + Chr(10) + Chr(13) + _
     "str_ValveLine=" + str_ValveLine + Chr(10) + Chr(13) + _
     "lng_TrimTypeID=" + CStr(lng_TrimTypeID) + Chr(10) + Chr(13) + _
     "str_PackingMaterial=" + str_PackingMaterial + Chr(10) + Chr(13) + _
     "str_LeakageClass=" + str_LeakageClass + Chr(10) + Chr(13) + _
     "str_SpringRange=" + CStr(str_SpringRange) + Chr(10) + Chr(13) + _
     "dbl_RatedCv=" + CStr(dbl_RatedCv)
#End If


'***************************
'Parameter validation
'***************************
If (dbl_ShutOffDP < 0#) Then GoTo CalcDP_2600_Cylinder_FTO_ATO_Error:
If (str_ValveLine = "") Then GoTo CalcDP_2600_Cylinder_FTO_ATO_Error:
If (str_LengthUM = "") Then GoTo CalcDP_2600_Cylinder_FTO_ATO_Error:
If (str_SpecifiedPressureUM = "") Then GoTo CalcDP_2600_Cylinder_FTO_ATO_Error:
If (str_SpecifiedTemperatureUM = "") Then GoTo CalcDP_2600_Cylinder_FTO_ATO_Error:
If (str_SpecifiedCalibPressureUM = "") Then GoTo CalcDP_2600_Cylinder_FTO_ATO_Error:
If (str_PackingMaterial = "") Then GoTo CalcDP_2600_Cylinder_FTO_ATO_Error:
If (str_LeakageClass = "") Then GoTo CalcDP_2600_Cylinder_FTO_ATO_Error:
If (str_SpringRange = "") Then GoTo CalcDP_2600_Cylinder_FTO_ATO_Error:

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4


'-determine requirements----------------------------------------------------

Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature(), dbl_MaxInletTemperature)

dbl_ShowMaxInletPressure = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, str_SpecifiedPressureUM)

dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, str_SpecifiedTemperatureUM)
dbl_ShowShutoffDP = dbl_ShutOffDP
dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, str_SpecifiedPressureUM)

'-determine ANSI limitation----------------------------------------------------

Call GetANSIPressure(str_BodyMaterial, dbl_MaxInletTemperature, str_SelectedRating, dbl_AllowANSIPressure)

'-read valve properties---------------------------------------------------------
Call Get2600Data(lng_TrimPackageID, dbl_DynamicFactor, dbl_PackingHeight, dbl_PlugDiameter, _
    dbl_SeatDiameter, dbl_StemDiameter, dbl_AllowPlugDP, str_StemMaterial, dbl_Travel)

'-calculate required packing force------------------------------------------

Call CalcPackingForce(dbl_MaxInletPressure, dbl_StemDiameter, str_PackingMaterial, _
    dbl_PackingHeight, dbl_ForcePacking)

'-calculate required seating force-------------------------------------------

Call CalcSeatingForce(str_LeakageClass, dbl_SeatDiameter, dbl_ForceSeating)

'-calculate required offbalance force---------------------------------------

Call CalcSingleSeatUnbalancedForceFTO(dbl_MaxInletPressure, dbl_SeatDiameter, dbl_ForceUnbalancedClosed)

dbl_ReqForceClosed = dbl_ForcePacking + dbl_ForceSeating + dbl_ForceUnbalancedClosed


'-calculate actuator capability--------------------------------------------------

Call GetActuatorData(lng_ActuatorPartID, dbl_AreaInitial, dbl_AreaFinal, dbl_MaxSupplyPressure, _
    dbl_ActuatorFriction, dbl_PositionerLossCoef, str_ActuatorModel, long_AirActionID, int_QtySeals, _
    dbl_StemSealWidth, dbl_PistonRingWidth, dbl_VolOver, dbl_ForceSpringClosed, _
    dbl_ActuatorStemDiameter, dbl_LeverLength)

Call CalcSpringRate(str_SpringRange, dbl_AreaInitial, dbl_PInitial, dbl_AreaFinal, dbl_PFinal, _
    str_SpecifiedCalibPressureUM, dbl_Travel, str_LengthUM, dbl_SpringRate)

Call CalcATOActuatorForce(dbl_AreaInitial, dbl_PInitial, dbl_PFinal, dbl_ActuatorFriction, _
    dbl_RequiredSupplyPressure, dbl_ActuatorForce)

'-calculate allowable pressure drop--------------------------------------------------

Call CalcAllowDPSingleStageUnbalFTO(dbl_ActuatorForce, dbl_ForceSeating, dbl_ForceFriction, _
    dbl_SeatDiameter, dbl_AllowDPClosed)

'-calculate stem limitation-------------------------------------------------------------
dbl_L = 6#
Call GetStemStressArea(dbl_StemDiameter, dbl_StemStressArea)
Call GetStemProperties(str_StemMaterial, dbl_MaxInletTemperature, dbl_Sy, dbl_E, dbl_Sd)
Call CalcAllowStemForce(dbl_Sy, dbl_E, dbl_StemDiameter, dbl_L, dbl_Sd, dbl_StemStressArea, _
    dbl_AllowableCompressiveForce)
Call CalcAllowStemTensileForce(dbl_Sd, dbl_StemStressArea, dbl_AllowableTensileForce)


If dbl_AllowableTensileForce < dbl_AllowableCompressiveForce Then
    dbl_AllowableStemLoad = dbl_AllowableTensileForce
Else
    dbl_AllowableStemLoad = dbl_AllowableCompressiveForce
End If
'-calculate actuator pressure drop capability---------------------------------------

dbl_AllowDPStability = 0
dbl_AllowDPOpen = 0
Call CalcMaxDPAllowable(dbl_AllowDPStability, dbl_AllowDPClosed, dbl_AllowDPOpen, _
    dbl_AllowANSIPressure, dbl_MaxAllowDP)

dbl_PSupplyDP = dbl_PFinal + 5

'-evaluation--------------------------------------------------------------
bln_ANSIOK = True
bln_StemOK = True
bln_ActuatorOK = True
bln_TrimOK = True

If dbl_MaxInletPressure > dbl_AllowANSIPressure Then
    bln_ANSIOK = False
    str_Message = ANSIMESSAGE
    dbl_ShowDP = dbl_ShowMaxInletPressure
    dbl_ShowDPSupplyPressure = 0
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowANSIPressure, str_SpecifiedPressureUM)
    dbl_ShowMaxDPSupplyPressure = 0
End If

If dbl_ActuatorForce > dbl_AllowableStemLoad Then
    bln_StemOK = False
    str_Message = STEMMESSAGE
    dbl_ShowDP = dbl_ShowMaxInletPressure
    dbl_ShowDPSupplyPressure = 0
    dbl_ShowMaxDP = 0
    dbl_ShowMaxDPSupplyPressure = 0
End If

If dbl_MaxInletPressure <= dbl_MaxAllowDP Then
    bln_ActuatorOK = True
    dbl_ShowDP = dbl_ShowMaxInletPressure
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_SpecifiedCalibPressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, str_SpecifiedPressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_SpecifiedCalibPressureUM)

ElseIf dbl_MaxInletPressure > dbl_MaxAllowDP And dbl_ShutOffDP <= dbl_MaxAllowDP Then
    bln_ActuatorOK = True
    str_Message = LIMITEDDPMESSAGE
    dbl_ShowDP = dbl_ShowShutoffDP
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_SpecifiedCalibPressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, str_SpecifiedPressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_SpecifiedCalibPressureUM)
Else
    bln_ActuatorOK = False
    str_Message = ACTUATORMESSAGE
    dbl_ShowDP = dbl_ShowMaxInletPressure
    dbl_ShowDPSupplyPressure = 0
    dbl_ShowMaxDP = 0
    dbl_ShowMaxDPSupplyPressure = 0
End If

dbl_ShowDP = TrimDecimalPlaces(dbl_ShowDP)
dbl_ShowDPSupplyPressure = TrimDecimalPlaces(dbl_ShowDPSupplyPressure)
dbl_ShowMaxDP = TrimDecimalPlaces(dbl_ShowMaxDP)
dbl_ShowMaxDPSupplyPressure = TrimDecimalPlaces(dbl_ShowMaxDPSupplyPressure)

#If (DebugVersion = "yes") Then
    MsgBox "2600-Cylinder-FTO-ATO Actuator" + Chr(10) + Chr(13) + _
    "dbl_ShowDP=" + CStr(dbl_ShowDP) + Chr(10) + Chr(13) + _
    "dbl_ShowDPSupplyPressure=" + CStr(dbl_ShowDPSupplyPressure) + Chr(10) + Chr(13) + _
    "dbl_ShowMaxDP=" + CStr(dbl_ShowMaxDP) + Chr(10) + Chr(13) + _
    "dbl_ShowMaxDPSupplyPressure=" + CStr(dbl_ShowMaxDPSupplyPressure) + Chr(10) + Chr(13) + _
    "str_Message=" + str_Message + Chr(10) + Chr(13)
#End If


Exit Sub
CalcDP_2600_Cylinder_FTO_ATO_Error:

End Sub
 
 Sub CalcDP_2600_Cylinder_FTO_ATC(str_ActuatorModel As String, lng_FlowDirectionID As Long, _
    lng_AirActionID As Long, lng_ActuatorPartID As Long, lng_TrimPackageID As Long, _
    dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, dbl_InletPressure3 As Double, _
    dbl_InletPressure4 As Double, dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, dbl_InletTemperature1 As Double, _
    dbl_InletTemperature2 As Double, dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, _
    dbl_ShutOffDP As Double, str_SpecifiedPressureUM As String, _
    str_SpecifiedTemperatureUM As String, str_SpecifiedCalibPressureUM As String, _
    str_LengthUM As String, str_BodyMaterial As String, _
    str_SelectedRating As String, str_ValveLine As String, _
    lng_TrimTypeID As Long, str_PackingMaterial As String, str_LeakageClass As String, _
    str_SpringRange As String, dbl_RatedCv As Double, _
    dbl_ShowDP As Double, dbl_ShowDPSupplyPressure As Double, _
    dbl_ShowMaxDP As Double, dbl_ShowMaxDPSupplyPressure As Double, _
    str_Message As String)
'--------------------------------------------------------------------------------------------------
'Reference : Actuator sizing methods for linear travel valves.

'This subroutine is used to calculate the selected valve and actuator shutoff capability and compare
'the result to defined requirements.
'This subroutine is specific to 2600 series valves with single stage trim,
'flow to open, air to close.
'---------------------------------------------------------------------------------------------------
'Input Variables
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff pressure drop
'str_SpecifiedPressureUM        unit of measure associated with inlet and outlet pressures
'str_specifiedTemperatureUM     unit of measure associated with inlet temperature
'str_specifiedCalibPressureUM   unit of measure associated with spring range
'                               and supply pressure
'str_BodyMaterial               selected material of body
'str_SelectedRating             selected body rating
'str_ValveLine                  selected product line
'lng_TrimTypeID                 selected type of trim
'str_PackingMaterial            selected packing material
'str_LeakageClass               selected leakage class
'str_SpringRange                selected actuator spring range
'dbl_RatedCv                    selected rated Cv
'---------------------------------------------------------------------------------------------------
'Output Variables
'dbl_ShowDP                     specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure       calculated required supply pressure for specified
'dbl_ShowMaxDP                  calculated maximum pressure drop for selected valve
'                               and actuator combination
'dbl_ShowMaxDPsupplyPressure    calculated required supply pressure for maximum pressure drop
'str_Message                    warning message if pressure drop requirement exceeds
'                               equipment capability
'---------------------------------------------------------------------------------------------------
On Error GoTo CalcDP_2600_Cylinder_FTO_ATC_Error:

Dim bln_ActuatorClosedOK As Boolean
Dim bln_ActuatorOK As Boolean
Dim bln_ActuatorOpenOK As Boolean
Dim bln_ActuatorStabilityOK As Boolean
Dim bln_ANSIOK As Boolean
Dim bln_StemOK As Boolean
Dim bln_SupplyOK As Boolean
Dim bln_TrimOK As Boolean
Dim dbl_ActuatorForce As Double
Dim dbl_ActuatorFriction As Double
Dim dbl_ActuatorStemDiameter As Double
Dim dbl_AllowableCompressiveForce As Double
Dim dbl_AllowableStemDP As Double
Dim dbl_AllowableStemLoad As Double
Dim dbl_AllowableTensileForce As Double
Dim dbl_AllowANSIPressure As Double
Dim dbl_AllowDP As Double
Dim dbl_AllowDPClosed As Double
Dim dbl_AllowDPOpen As Double
Dim dbl_AllowDPStability As Double
Dim dbl_AllowPlugDP As Double
Dim dbl_AreaFinal As Double
Dim dbl_AreaInitial As Double
Dim dbl_AreaOver As Double
Dim dbl_AreaUnder As Double
Dim dbl_As As Double
Dim dbl_DynamicFactor As Double
Dim dbl_E As Double
Dim dbl_ForceFriction As Double
Dim dbl_ForcePacking As Double
Dim dbl_ForceSeating As Double
Dim dbl_ForceSpringClosed As Double
Dim dbl_ForceUnbalancedClosed As Double
Dim dbl_ForceUnbalancedOpen As Double
Dim dbl_InletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double
Dim dbl_L As Double
Dim dbl_LeverLength As Double
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_MaxPressureDrop As Double
Dim dbl_MaxSupplyDP As Double
Dim dbl_MaxSupplyPressure As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_PackingHeight As Double
Dim dbl_PFinal As Double
Dim dbl_PInitial As Double
Dim dbl_PistonRingWidth As Double
Dim dbl_PlugDiameter As Double
Dim dbl_PositionerLossCoef As Double
Dim dbl_PreloadANSI As Double
Dim dbl_PreloadDP As Double
Dim dbl_PreloadPlugAllow As Double
Dim dbl_PreloadShutoffDP As Double
Dim dbl_PreloadStemAllow As Double
Dim dbl_PSupply As Double
Dim dbl_PSupplyANSI As Double
Dim dbl_PSupplyDP As Double
Dim dbl_PSupplyPlugAllow As Double
Dim dbl_PSupplyShutoffDP As Double
Dim dbl_PSupplyStemAllow As Double
Dim dbl_ReqANSIForceClosed As Double
Dim dbl_ReqDPForceClosed As Double
Dim dbl_ReqForceClosed As Double
Dim dbl_ReqForceOpen As Double
Dim dbl_ReqPlugLimitForceClosed As Double
Dim dbl_ReqShutoffForceClosed As Double
Dim dbl_RequiredPreload As Double
Dim dbl_RequiredSupply As Double
Dim dbl_RequiredSupplyPressure As Double
Dim dbl_Sd As Double
Dim dbl_SeatDiameter As Double
Dim dbl_ShowMaxInletPressure As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim dbl_ShowMaxPressureDrop As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_SpringRate As Double
Dim dbl_StemDiameter As Double
Dim dbl_StemSealWidth As Double
Dim dbl_StemStressArea As Double
Dim dbl_StemTravel As Double
Dim dbl_Sy As Double
Dim dbl_TempShowMaxDP As Double
Dim dbl_theMaxAllowSupply As Double
Dim dbl_theMaxPreload As Double
Dim dbl_Travel As Double
Dim dbl_VolOver As Double
Dim int_QtySeals As Integer
Dim long_AirActionID As Long
Dim str_StemMaterial As String

#If (DebugVersion = "yes") Then
MsgBox "2600-Cylinder-FTO-ATC Actuator" + Chr(10) + Chr(13) + _
     "str_ActuatorModel=" + str_ActuatorModel + Chr(10) + Chr(13) + _
     "lng_FlowDirectionID=" + CStr(lng_FlowDirectionID) + Chr(10) + Chr(13) + _
     "lng_AirActionID=" + CStr(lng_AirActionID) + Chr(10) + Chr(13) + _
     "lng_TrimPackageID=" + CStr(lng_TrimPackageID) + Chr(10) + Chr(13) + _
     "lng_ActuatorPartID=" + CStr(lng_ActuatorPartID) + Chr(10) + Chr(13) + _
     "dbl_InletPressure1=" + CStr(dbl_InletPressure1) + Chr(10) + Chr(13) + _
     "dbl_OutletPressure1=" + CStr(dbl_OutletPressure1) + Chr(10) + Chr(13) + _
     "dbl_InletTemperature1=" + CStr(dbl_InletTemperature1) + Chr(10) + Chr(13) + _
     "dbl_ShutOffDP=" + CStr(dbl_ShutOffDP) + Chr(10) + Chr(13) + _
     "str_SpecifiedPressureUM=" + str_SpecifiedPressureUM + Chr(10) + Chr(13) + _
     "str_SpecifiedTemperatureUM=" + str_SpecifiedTemperatureUM + Chr(10) + Chr(13) + _
     "str_SpecifiedCalibPressureUM=" + str_SpecifiedCalibPressureUM + Chr(10) + Chr(13) + _
     "str_LengthUM=" + str_LengthUM + Chr(10) + Chr(13) + _
     "str_BodyMaterial=" + str_BodyMaterial + Chr(10) + Chr(13) + _
     "str_SelectedRating=" + str_SelectedRating + Chr(10) + Chr(13) + _
     "str_ValveLine=" + str_ValveLine + Chr(10) + Chr(13) + _
     "lng_TrimTypeID=" + CStr(lng_TrimTypeID) + Chr(10) + Chr(13) + _
     "str_PackingMaterial=" + str_PackingMaterial + Chr(10) + Chr(13) + _
     "str_LeakageClass=" + str_LeakageClass + Chr(10) + Chr(13) + _
     "str_SpringRange=" + CStr(str_SpringRange) + Chr(10) + Chr(13) + _
     "dbl_RatedCv=" + CStr(dbl_RatedCv)
#End If

'***************************
'Parameter validation
'***************************
If (dbl_ShutOffDP < 0#) Then GoTo CalcDP_2600_Cylinder_FTO_ATC_Error:
If (str_ValveLine = "") Then GoTo CalcDP_2600_Cylinder_FTO_ATC_Error:
If (str_LengthUM = "") Then GoTo CalcDP_2600_Cylinder_FTO_ATC_Error:
If (str_SpecifiedPressureUM = "") Then GoTo CalcDP_2600_Cylinder_FTO_ATC_Error:
If (str_SpecifiedTemperatureUM = "") Then GoTo CalcDP_2600_Cylinder_FTO_ATC_Error:
If (str_SpecifiedCalibPressureUM = "") Then GoTo CalcDP_2600_Cylinder_FTO_ATC_Error:
If (str_PackingMaterial = "") Then GoTo CalcDP_2600_Cylinder_FTO_ATC_Error:
If (str_LeakageClass = "") Then GoTo CalcDP_2600_Cylinder_FTO_ATC_Error:
If (str_SpringRange = "") Then GoTo CalcDP_2600_Cylinder_FTO_ATC_Error:

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4

'-determine requirements----------------------------------------------------

Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature(), dbl_MaxInletTemperature)
Call CalcTestPressureDrop(dbl_InletPressure(), dbl_OutletPressure(), dbl_MaxPressureDrop)

dbl_ShowMaxInletPressure = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, str_SpecifiedPressureUM)

dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, str_SpecifiedTemperatureUM)
dbl_ShowMaxPressureDrop = dbl_MaxPressureDrop
dbl_MaxPressureDrop = ConvertPressureToPSIG(dbl_MaxPressureDrop, str_SpecifiedPressureUM)
dbl_ShowShutoffDP = dbl_ShutOffDP
dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, str_SpecifiedPressureUM)

'-determine ANSI limitation----------------------------------------------------

Call GetANSIPressure(str_BodyMaterial, dbl_MaxInletTemperature, str_SelectedRating, dbl_AllowANSIPressure)

'-read valve properties---------------------------------------------------------

Call Get2600Data(lng_TrimPackageID, dbl_DynamicFactor, dbl_PackingHeight, dbl_PlugDiameter, _
    dbl_SeatDiameter, dbl_StemDiameter, dbl_AllowPlugDP, str_StemMaterial, dbl_Travel)

'-calculate required packing force------------------------------------------

Call CalcPackingForce(dbl_MaxInletPressure, dbl_StemDiameter, str_PackingMaterial, dbl_PackingHeight, _
    dbl_ForcePacking)

'-calculate required seating force-------------------------------------------

Call CalcSeatingForce(str_LeakageClass, dbl_SeatDiameter, dbl_ForceSeating)

'-calculate required offbalance force---------------------------------------

Call CalcSingleSeatUnbalancedForceFTO(dbl_MaxInletPressure, dbl_SeatDiameter, dbl_ForceUnbalancedClosed)
dbl_ReqDPForceClosed = dbl_ForcePacking + dbl_ForceSeating + dbl_ForceUnbalancedClosed

Call CalcSingleSeatUnbalancedForceFTO(dbl_AllowANSIPressure, dbl_SeatDiameter, dbl_ForceUnbalancedClosed)
dbl_ReqANSIForceClosed = dbl_ForcePacking + dbl_ForceSeating + dbl_ForceUnbalancedClosed

Call CalcSingleSeatUnbalancedForceFTO(dbl_AllowPlugDP, dbl_SeatDiameter, dbl_ForceUnbalancedClosed)
dbl_ReqPlugLimitForceClosed = dbl_ForcePacking + dbl_ForceSeating + dbl_ForceUnbalancedClosed

Call CalcSingleSeatUnbalancedForceFTO(dbl_ShutOffDP, dbl_SeatDiameter, dbl_ForceUnbalancedClosed)
dbl_ReqShutoffForceClosed = dbl_ForcePacking + dbl_ForceSeating + dbl_ForceUnbalancedClosed

'-calculate stem limitation-------------------------------------------------------------

dbl_L = 6#
Call GetStemStressArea(dbl_StemDiameter, dbl_StemStressArea)
Call GetStemProperties(str_StemMaterial, dbl_MaxInletTemperature, dbl_Sy, dbl_E, dbl_Sd)
Call CalcAllowStemForce(dbl_Sy, dbl_E, dbl_StemDiameter, dbl_L, dbl_Sd, _
    dbl_StemStressArea, dbl_AllowableCompressiveForce)
Call CalcAllowStemTensileForce(dbl_Sd, dbl_StemStressArea, dbl_AllowableTensileForce)

If dbl_AllowableTensileForce < dbl_AllowableCompressiveForce Then
    dbl_AllowableStemLoad = dbl_AllowableTensileForce
Else
    dbl_AllowableStemLoad = dbl_AllowableCompressiveForce
End If

'-calculate required supply presure--------------------------------------------------

Call GetActuatorData(lng_ActuatorPartID, dbl_AreaInitial, dbl_AreaFinal, dbl_MaxSupplyPressure, _
    dbl_ActuatorFriction, dbl_PositionerLossCoef, str_ActuatorModel, long_AirActionID, int_QtySeals, _
    dbl_StemSealWidth, dbl_PistonRingWidth, dbl_VolOver, dbl_ForceSpringClosed, _
    dbl_ActuatorStemDiameter, dbl_LeverLength)

Call CalcSpringRate(str_SpringRange, dbl_AreaInitial, dbl_PInitial, dbl_AreaFinal, dbl_PFinal, _
    str_SpecifiedCalibPressureUM, dbl_Travel, str_LengthUM, dbl_SpringRate)

Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, dbl_ReqANSIForceClosed, _
    dbl_PSupplyANSI)

Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, dbl_AllowableStemLoad, _
    dbl_PSupplyStemAllow)

Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, dbl_ReqShutoffForceClosed, _
    dbl_PSupplyShutoffDP)

Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, dbl_ReqDPForceClosed, _
    dbl_PSupplyDP)

dbl_theMaxAllowSupply = dbl_MaxSupplyPressure
If dbl_PSupplyANSI < dbl_theMaxAllowSupply Then dbl_theMaxAllowSupply = dbl_PSupplyANSI
If dbl_PSupplyStemAllow < dbl_theMaxAllowSupply Then dbl_theMaxAllowSupply = dbl_PSupplyStemAllow

'-calculate max DP with max allow supply-------------------------------
Call CalcATCActuatorForce(dbl_AreaFinal, dbl_PFinal, dbl_theMaxAllowSupply, _
    dbl_ActuatorFriction, dbl_ActuatorForce)

Call CalcAllowDPSingleStageUnbalFTO(dbl_ActuatorForce, dbl_ForceSeating, dbl_ForceFriction, _
    dbl_SeatDiameter, dbl_MaxSupplyDP)

'-calculate allowable pressure drop for stem ----------------------

Call CalcAllowDPSingleStageUnbalFTO(dbl_ActuatorForce, dbl_ForceSeating, dbl_ForceFriction, _
    dbl_SeatDiameter, dbl_AllowableStemDP)

'calculate max allowable pressure drop
dbl_ShowMaxDP = dbl_AllowableStemDP
dbl_ShowMaxDPSupplyPressure = dbl_PSupplyStemAllow

If (dbl_AllowANSIPressure < dbl_ShowMaxDP) Then
    dbl_ShowMaxDP = dbl_AllowANSIPressure
    dbl_ShowMaxDPSupplyPressure = dbl_PSupplyANSI
End If

If (dbl_MaxSupplyDP < dbl_ShowMaxDP) Then
    dbl_ShowMaxDP = dbl_MaxSupplyDP
    dbl_ShowMaxDPSupplyPressure = dbl_MaxSupplyPressure
End If
dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_ShowMaxDP, str_SpecifiedPressureUM)
dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_ShowMaxDPSupplyPressure, str_SpecifiedCalibPressureUM)

'-evaluation--------------------------------------------------------------

bln_ANSIOK = True
bln_StemOK = True
bln_ActuatorOpenOK = True
bln_ActuatorClosedOK = True
bln_ActuatorStabilityOK = True
bln_TrimOK = True
bln_SupplyOK = True

If (dbl_MaxInletPressure > dbl_AllowANSIPressure) Then
    bln_ANSIOK = False
    str_Message = ANSIMESSAGE
    dbl_ShowDP = dbl_ShowMaxInletPressure
    dbl_ShowDPSupplyPressure = 0
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowANSIPressure, str_SpecifiedPressureUM)
    dbl_ShowMaxDPSupplyPressure = 0
End If

If dbl_MaxInletPressure > dbl_AllowableStemDP Then
    bln_StemOK = False
    str_Message = STEMMESSAGE
    dbl_ShowDP = dbl_ShowMaxInletPressure
    dbl_ShowDPSupplyPressure = 0
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowableStemDP, str_SpecifiedPressureUM)
    dbl_ShowMaxDPSupplyPressure = 0
End If
 
If dbl_PSupplyDP > dbl_MaxSupplyPressure And dbl_PSupplyShutoffDP <= dbl_MaxSupplyPressure Then
    bln_SupplyOK = False
    str_Message = LIMITEDDPMESSAGE
    dbl_ShowDP = ConvertPressureFromPSIG(dbl_ShutOffDP, str_SpecifiedPressureUM)
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_SpecifiedCalibPressureUM)
    dbl_ShowMaxDP = 0
    dbl_ShowMaxDPSupplyPressure = 0
    
ElseIf dbl_PSupplyDP > dbl_MaxSupplyPressure And dbl_PSupplyShutoffDP > dbl_MaxSupplyPressure Then
    bln_SupplyOK = False
    str_Message = SUPPLYMESSAGE
    dbl_ShowDP = dbl_ShowMaxInletPressure
    dbl_ShowDPSupplyPressure = 0
    dbl_ShowMaxDP = 0
    dbl_ShowMaxDPSupplyPressure = 0
End If


If bln_ANSIOK = True And bln_StemOK = True And bln_SupplyOK = True Then
    dbl_ShowDP = dbl_ShowMaxInletPressure
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG(dbl_PSupplyDP, str_SpecifiedCalibPressureUM)
    'dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxSupplyDP, str_SpecifiedPressureUM)
    'dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG(dbl_theMaxAllowSupply, str_SpecifiedCalibPressureUM)
End If


dbl_ShowDP = TrimDecimalPlaces(dbl_ShowDP)
dbl_ShowDPSupplyPressure = TrimDecimalPlaces(dbl_ShowDPSupplyPressure)
dbl_ShowMaxDP = TrimDecimalPlaces(dbl_ShowMaxDP)
dbl_ShowMaxDPSupplyPressure = TrimDecimalPlaces(dbl_ShowMaxDPSupplyPressure)

#If (DebugVersion = "yes") Then
    MsgBox "2600-Cylinder-FTO-ATC Actuator" + Chr(10) + Chr(13) + _
    "dbl_ShowDP=" + CStr(dbl_ShowDP) + Chr(10) + Chr(13) + _
    "dbl_ShowDPSupplyPressure=" + CStr(dbl_ShowDPSupplyPressure) + Chr(10) + Chr(13) + _
    "dbl_ShowMaxDP=" + CStr(dbl_ShowMaxDP) + Chr(10) + Chr(13) + _
    "dbl_ShowMaxDPSupplyPressure=" + CStr(dbl_ShowMaxDPSupplyPressure) + Chr(10) + Chr(13) + _
    "str_Message=" + str_Message + Chr(10) + Chr(13)
#End If


Exit Sub
CalcDP_2600_Cylinder_FTO_ATC_Error:

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' END OF DOMOTOR
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INCLUDE 41000 LL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'--------------------------------------------------
'  New added subroutines
'--------------------------------------------------
Public Sub CalcActuatorSealFrictionForce(dbl_AreaInitial As Double, dbl_ActuatorStemDiameter _
        As Double, dbl_PistonRingWidth As Double, dbl_StemSealWidth As Double, int_QtySeals _
        As Integer, dbl_PressureSupply As Double, dbl_ForceActSeal As Double)
'--------------------------------------------------------------------------------------
'Reference :CES 1019 Balanced and pilot balanced cage throttling globe valves
'           - Actuator Sizing Methods.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the force to overcome actuator seal friction force.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_AreaInitial          effective actuator area                           in^2
'dbl_PistonRingWidth      width of actuator piston seal ring                in.
'dbl_ActuatorStemDiameter diameter of stem for specified actuator           in.
'dbl_StemSealWidth        width of actuator stem seal                       in.
'int_qtySeals             number of stem seals used in actuator
'dbl_PressureSupply       supply pressure                                   psig
'
'Output Variables:
'dbl_ForceActSeal         force required to overcome actuator seal friction lbs.
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/14/98  Name: David Zhou
'Tested     Date: 10/19/98  Name: David Zhou     Passed(Y/N):Y
'-----------------------------------------------------------------------------------------

On Error GoTo CalcActuatorSealFrictionForce_Error

'***************************
'Parameter validation
'***************************
If dbl_AreaInitial < 0# Then GoTo CalcActuatorSealFrictionForce_Error
If dbl_ActuatorStemDiameter < 0# Then GoTo CalcActuatorSealFrictionForce_Error
If dbl_PistonRingWidth < 0# Then GoTo CalcActuatorSealFrictionForce_Error
If dbl_StemSealWidth < 0# Then GoTo CalcActuatorSealFrictionForce_Error
If int_QtySeals < 0 Then GoTo CalcActuatorSealFrictionForce_Error
If dbl_PressureSupply < 0# Then GoTo CalcActuatorSealFrictionForce_Error

'**************************
'Declarations
'**************************
Dim dbl_ActuatorDiameter As Double
Dim dbl_Mu As Double
Const PI As Double = 3.1415927

'****************************
'Begin calculations
'****************************
'greased rubber on steel
dbl_Mu = 0.05
dbl_ActuatorDiameter = Sqr(4 * dbl_AreaInitial / PI)
dbl_ForceActSeal = dbl_Mu * PI * (dbl_ActuatorDiameter * dbl_PistonRingWidth + _
    dbl_ActuatorStemDiameter * dbl_StemSealWidth * int_QtySeals) * dbl_PressureSupply / 2
Exit Sub

CalcActuatorSealFrictionForce_Error:
'Placeholder

End Sub

Public Sub CalcPlugSealRingFrictionForce(dbl_SealRingCoef As Double, dbl_TECSealCoef As Double, _
    dbl_TECCoefJ As Double, dbl_TECCoefS As Double, dbl_CageInsideDiameter As Double, dbl_SealRingWidth _
    As Double, dbl_ForceSealRingArea As Double, dbl_ForceSealRingSpring As Double)

'--------------------------------------------------------------------------------------
'Reference : CES 1019 Balanced and pilot balanced cage throttling globe valves - Actuator Sizing Methods.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the force necessary to overcome plug seal ring  friction force.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_SealRingCoef        friction coefficient for specified seal ring material
'dbl_TECsealCoef         friction coefficient for TEC seal
'dbl_TECcoefJ            friction coefficient for spring energized seal (area component)
'dbl_TECcoefS            friction coefficient for spring energized seal (spring component)
'dbl_CageInsideDiameter  inside diameter of cage where it is inches contacted by the seal ring
'dbl_SealRingWidth       width of plug seal ring         inches
'
'Output Variables:
'dbl_ForceSealRingArea   force required to overcome plug seal friction (area component) lbs
'dbl_ForceSealRingSpring force required to overcome plug seal friction (spring component) lbs
'
'NOTE: either dbl_SealRingCoef or dbl_TECSealCoef must be 0 (zero).
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/14/98  Name: David Zhou
'Tested     Date: 10/19/98  Name: David Zhou     Passed(Y/N):Y
'-----------------------------------------------------------------------------------------
'**************************
'Declarations
'**************************
Const PI As Double = 3.1415927
Const dbl_SmallNum = 1E-20

On Error GoTo CalcPlugSealRingFrictionForce_Error

'***************************
'Parameter validation
'***************************
If dbl_SealRingCoef < 0# Then GoTo CalcPlugSealRingFrictionForce_Error
If dbl_TECSealCoef < 0# Then GoTo CalcPlugSealRingFrictionForce_Error
If dbl_TECCoefJ < 0# Then GoTo CalcPlugSealRingFrictionForce_Error
If dbl_TECCoefS < 0# Then GoTo CalcPlugSealRingFrictionForce_Error
If dbl_CageInsideDiameter < 0# Then GoTo CalcPlugSealRingFrictionForce_Error
If dbl_SealRingWidth < 0# Then GoTo CalcPlugSealRingFrictionForce_Error
If Abs(dbl_SealRingCoef * dbl_TECSealCoef) > dbl_SmallNum Then GoTo CalcPlugSealRingFrictionForce_Error

'****************************
'Begin calculations
'****************************
If (dbl_TECSealCoef < dbl_SmallNum) Then
    dbl_ForceSealRingArea = dbl_SealRingCoef * PI * dbl_CageInsideDiameter * dbl_SealRingWidth / 2
    dbl_ForceSealRingSpring = 0#
Else
    dbl_ForceSealRingArea = dbl_TECSealCoef * PI * dbl_CageInsideDiameter * dbl_TECCoefJ
    dbl_ForceSealRingSpring = dbl_TECSealCoef * PI * dbl_CageInsideDiameter * dbl_TECCoefS
End If
Exit Sub

CalcPlugSealRingFrictionForce_Error:
'Placeholder

End Sub
Public Sub CalcReqActForce_41000_FTO(dbl_MaxInletPressure As Double, dbl_SeatDiameter As Double, _
       dbl_CageInsideDiameter As Double, dbl_StemDiameter As Double, dbl_ForceSeating As Double, _
       dbl_ForcePacking As Double, dbl_ForceActSeal As Double, dbl_ForcePositionerLoss _
       As Double, dbl_ForceSealRingArea As Double, dbl_ForceSealRingSpring _
       As Double, dbl_ForcePlugWt As Double, dbl_RequiredActuatorForce As Double)
'--------------------------------------------------------------------------------------
'Reference : CES 1019 Balanced and pilot balanced cage throttling globe valves - Actuator Sizing Methods.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the force necessary to operate a 41000 series balanced trim package when the flow direction is flow to open.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_MaxInletPressure     pressure that actuator must operate against           psia
'dbl_SeatDiameter         seat ring orifice diameter                            in.
'dbl_CageInsideDiameter   inside diameter of cage where it is                   in.
'                         contacted by the seal ring
'dbl_StemDiameter         diameter of the plug stem                             in.
'dbl_ForceSeating         force req'd to provide desired leakage class          lbs
'dbl_ForcePacking         force req'd to overcome packing friction              lbs
'dbl_ForceActSeal         force req'd to overcome actuator seal friction        lbs
'dbl_ForcePositionerLoss  force req'd to compensate for positioner loss         lbs
'dbl_ForceSealRingArea    force req'd to overcome plug seal ring friction (area)lbs
'dbl_ForceSealRingSpring  force req'd to overcome plug seal ring friction (spring)lbs
'dbl_ForcePlugWt          force req'd to overcome plug weight                   lbs
'
'Output Variables:
'dbl_RequiredActuatorForce force required to operate against specified          lbs
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/14/98  Name: David Zhou
'Tested     Date: 10/19/98  Name: David Zhou     Passed(Y/N):Y
'-----------------------------------------------------------------------------------------
'**************************
'Declarations
'**************************
Dim dbl_PressureComponent As Double
Dim dbl_Temp As Double
Const PI As Double = 3.1415927

On Error GoTo CalcReqActForce_41000_FTO_Error

'***************************
'Parameter validation
'***************************
If dbl_MaxInletPressure < 0# Then GoTo CalcReqActForce_41000_FTO_Error
If dbl_SeatDiameter < 0# Then GoTo CalcReqActForce_41000_FTO_Error
If dbl_CageInsideDiameter < 0# Then GoTo CalcReqActForce_41000_FTO_Error
If dbl_StemDiameter < 0# Then GoTo CalcReqActForce_41000_FTO_Error
If dbl_ForceSeating < 0# Then GoTo CalcReqActForce_41000_FTO_Error
If dbl_ForcePacking < 0# Then GoTo CalcReqActForce_41000_FTO_Error
If dbl_ForceActSeal < 0# Then GoTo CalcReqActForce_41000_FTO_Error
If dbl_ForcePositionerLoss < 0# Then GoTo CalcReqActForce_41000_FTO_Error
If dbl_ForceSealRingArea < 0# Then GoTo CalcReqActForce_41000_FTO_Error
If dbl_ForceSealRingSpring < 0# Then GoTo CalcReqActForce_41000_FTO_Error
If dbl_ForcePlugWt < 0# Then GoTo CalcReqActForce_41000_FTO_Error

'****************************
'Begin calculations
'****************************
dbl_PressureComponent = dbl_MaxInletPressure * ((dbl_SeatDiameter ^ 2 - dbl_CageInsideDiameter ^ 2 _
    + dbl_StemDiameter ^ 2) * PI / 4 + dbl_ForceSealRingArea)
dbl_Temp = dbl_ForceSeating + dbl_ForceActSeal + dbl_ForcePositionerLoss + _
            dbl_ForceSealRingSpring - dbl_ForcePlugWt

If (dbl_MaxInletPressure <= 2000#) Then
    dbl_RequiredActuatorForce = dbl_PressureComponent + dbl_Temp + dbl_ForcePacking
Else
    dbl_PressureComponent = dbl_PressureComponent + (dbl_ForcePacking / 1000) * dbl_MaxInletPressure
    dbl_RequiredActuatorForce = dbl_PressureComponent + dbl_Temp
End If
Exit Sub

CalcReqActForce_41000_FTO_Error:
'Placeholder

End Sub

Public Sub CalcAllowDP_41000_FTO(dbl_MaxInletPressure As Double, dbl_ForceActuator As Double, _
    dbl_SeatDiameter As Double, dbl_CageInsideDiameter As Double, dbl_StemDiameter As Double, _
    dbl_ForceSeating As Double, dbl_ForcePacking As Double, dbl_ForceActSeal As Double, _
    dbl_ForcePositionerLoss As Double, dbl_ForceSealRingArea As Double, _
    dbl_ForceSealRingSpring As Double, dbl_ForcePlugWt As Double, dbl_AllowDP As Double)
'--------------------------------------------------------------------------------------
'Reference : CES 1019 Balanced and pilot balanced cage throttling globe valves - Actuator Sizing Methods.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the allowable pressure drop for  41000 series balanced trim package when the flow direction is flow to open.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_MaxInletPressure         pressure that actuator must operate against           psia
'dbl_ForceActuator          available actuator output force                         psia
'dbl_SeatDiameter           seat ring orifice diameter                              in.
'dbl_CageInsideDiameter     inside diam of cage where it is contacted by seal ring  in.
'dbl_StemDiameter           diameter of the plug stem                               in.
'dbl_ForceSeating           force required to provide desired leakage class         lbs
'dbl_ForcePacking           force req'd to overcome packing friction                lbs
'dbl_ForceActSeal           force req'd to overcome actuator seal friction          lbs
'dbl_ForcePositionerLoss    force req'd to compensate for positioner loss           lbs
'dbl_ForceSealRingArea      force req'd to overcome plug seal ring friction (area)  lbs
'dbl_ForceSealRingSpring    force req'd to overcome plug seal ring friction (spring)lbs
'dbl_ForcePlugWt            force req'd to overcome plug weight                     lbs
'
'Output Variables:
'dbl_AllowDP                allowable pressure with spec'd actuator force available psi
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/14/98  Name: David Zhou
'Tested     Date: 10/19/98  Name: David Zhou     Passed(Y/N):Y
'-----------------------------------------------------------------------------------------
'**************************
'Declarations
'**************************
Dim dbl_PressureComponent As Double
Dim dbl_Temp As Double
Const PI As Double = 3.1415927

On Error GoTo CalcAllowDP_41000_FTO_Error

'***************************
'Parameter validation
'***************************
If dbl_MaxInletPressure < 0# Then GoTo CalcAllowDP_41000_FTO_Error
If dbl_SeatDiameter < 0# Then GoTo CalcAllowDP_41000_FTO_Error
If dbl_CageInsideDiameter < 0# Then GoTo CalcAllowDP_41000_FTO_Error
If dbl_StemDiameter < 0# Then GoTo CalcAllowDP_41000_FTO_Error
If dbl_ForceSeating < 0# Then GoTo CalcAllowDP_41000_FTO_Error
If dbl_ForcePacking < 0# Then GoTo CalcAllowDP_41000_FTO_Error
If dbl_ForceActSeal < 0# Then GoTo CalcAllowDP_41000_FTO_Error
If dbl_ForcePositionerLoss < 0# Then GoTo CalcAllowDP_41000_FTO_Error
If dbl_ForceSealRingArea < 0# Then GoTo CalcAllowDP_41000_FTO_Error
If dbl_ForceSealRingSpring < 0# Then GoTo CalcAllowDP_41000_FTO_Error
If dbl_ForcePlugWt < 0# Then GoTo CalcAllowDP_41000_FTO_Error

'****************************
'Begin calculations
'****************************
dbl_PressureComponent = ((dbl_SeatDiameter ^ 2 - dbl_CageInsideDiameter ^ 2 _
    + dbl_StemDiameter ^ 2) * PI / 4 + dbl_ForceSealRingArea)
dbl_Temp = dbl_ForceActuator - dbl_ForceSeating - dbl_ForcePacking - dbl_ForceActSeal - _
           dbl_ForcePositionerLoss - dbl_ForceSealRingSpring + dbl_ForcePlugWt

If (dbl_MaxInletPressure <= 2000#) Then
    dbl_AllowDP = dbl_Temp / dbl_PressureComponent
Else
    dbl_PressureComponent = dbl_PressureComponent + (dbl_ForcePacking / 1000)
    dbl_AllowDP = dbl_Temp / dbl_PressureComponent
End If
Exit Sub

CalcAllowDP_41000_FTO_Error:
'Placeholder

End Sub

Public Sub CalcDynamicFrictionRateChange(dbl_MaxInletPressure As Double, dbl_ForceSealRingArea _
    As Double, dbl_ForceSealRingSpring As Double, dbl_ForcePacking As Double, _
    dbl_Travel As Double, dbl_DynFrictionRateChange As Double)
'--------------------------------------------------------------------------------------
'Reference : CES 1019 Balanced and pilot balanced cage throttling globe valves - Actuator Sizing Methods.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the dynamic friction rate change.  This value is compared to spring rate in order to evaluate the stiffness of the selected actuator for the specified application.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_MaxInletPressure      pressure that actuator must operate against              psia
'dbl_ForceSealRingArea     force req'd to overcome plug seal ring friction (area)   lbs
'dbl_ForceSealRingSpring   force req'd to overcome plug seal friction(spring)       lbs
'dbl_ForcePacking          force req'd to overcome packing friction rated plug travel lbs
'dbl_Travel                                                                         in.
'
'Output Variables:
'dbl_DynFrictionRateChange calculated rate change of friction                       lb/in
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/14/98  Name: David Zhou
'Tested     Date: 10/19/98  Name: David Zhou     Passed(Y/N):Y
'-----------------------------------------------------------------------------------------
'**************************
'Declarations
'**************************
Dim dbl_Temp As Double
Const PI As Double = 3.1415927

On Error GoTo CalcDynamicFrictionRateChange_Error

'***************************
'Parameter validation
'***************************
If dbl_MaxInletPressure < 0# Then GoTo CalcDynamicFrictionRateChange_Error
If dbl_ForceSealRingArea < 0# Then GoTo CalcDynamicFrictionRateChange_Error
If dbl_ForcePacking < 0# Then GoTo CalcDynamicFrictionRateChange_Error

'****************************
'Begin calculations
'****************************
dbl_Temp = dbl_MaxInletPressure * 2# * dbl_ForceSealRingArea + dbl_ForceSealRingSpring

If (dbl_MaxInletPressure <= 2000#) Then
    dbl_DynFrictionRateChange = (dbl_Temp + dbl_ForcePacking) / 0.4 / dbl_Travel
Else
    dbl_DynFrictionRateChange = (dbl_Temp + dbl_ForcePacking / 1000#) / 0.4 / dbl_Travel
End If
Exit Sub

CalcDynamicFrictionRateChange_Error:
'Placeholder

End Sub

Public Sub CalcAllowDPClosed_41000_FTC(dbl_MaxInletPressure As Double, dbl_ActForceClosed _
    As Double, dbl_ForcePacking As Double, dbl_ForceActSeal As Double, _
    dbl_ForcePositionerLoss As Double, dbl_ForceSealRingSpring As Double, _
    dbl_ForcePilotSpring As Double, dbl_AreaUnbalancedClosed As Double, dbl_ForcePlugWt _
    As Double, dbl_AllowDPClosed As Double)
'--------------------------------------------------------------------------------------
'Reference : CES 1019 Balanced and pilot balanced cage throttling globe valves - Actuator Sizing Methods.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the allowable pressure drop in the closed position for  41000 series balanced trim package when the flow direction is flow to close.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_MaxInletPressure        pressure that actuator must operate against        psia
'dbl_ActuatorForceClosed     actuator force provided in closed position         lbs
'dbl_ForcePacking            force req'd to overcome packing friction           lbs
'dbl_ForceActSeal            force req'd to overcome actuator seal friction     lbs
'dbl_ForcePositionerLoss     force req'd to compensate for positioner loss      lbs
'dbl_ForceSealRingSpring     force req'd to overcome plug seal ring friction (spring)   lbs
'dbl_ForcePilotSpring        force req'd to compress pilot spring               lbs
'dbl_ForcePlugWt             force req'd to overcome plug weight                lbs
'dbl_AreaUnbalancedClosed    effective unbalanced area when valve is closed     in^2

'Output Variables:
'dbl_AllowDPclosed       allowable pressure with specified actuator force avail.psi
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/14/98  Name: David Zhou
'Tested     Date: 10/19/98  Name: David Zhou     Passed(Y/N):Y
'-----------------------------------------------------------------------------------------
'**************************
'Declarations
'**************************
Dim dbl_Temp As Double
Const PI As Double = 3.1415927

On Error GoTo CalcAllowDPClosed_41000_FTC_Error

'***************************
'Parameter validation
'***************************
If dbl_MaxInletPressure < 0# Then GoTo CalcAllowDPClosed_41000_FTC_Error
If dbl_ActForceClosed < 0# Then GoTo CalcAllowDPClosed_41000_FTC_Error
If dbl_ForcePacking < 0# Then GoTo CalcAllowDPClosed_41000_FTC_Error
If dbl_ForceActSeal < 0# Then GoTo CalcAllowDPClosed_41000_FTC_Error
If dbl_ForcePositionerLoss < 0# Then GoTo CalcAllowDPClosed_41000_FTC_Error
If dbl_ForcePilotSpring < 0# Then GoTo CalcAllowDPClosed_41000_FTC_Error
If dbl_ForceSealRingSpring < 0# Then GoTo CalcAllowDPClosed_41000_FTC_Error
If dbl_AreaUnbalancedClosed < 0# Then GoTo CalcAllowDPClosed_41000_FTC_Error
If dbl_ForcePlugWt < 0# Then GoTo CalcAllowDPClosed_41000_FTC_Error

'****************************
'Begin calculations
'****************************
dbl_Temp = dbl_ActForceClosed - dbl_ForceActSeal - dbl_ForcePositionerLoss - _
           dbl_ForceSealRingSpring - dbl_ForcePilotSpring - dbl_ForcePlugWt

If (dbl_MaxInletPressure <= 2000#) Then
    dbl_AllowDPClosed = (dbl_Temp - dbl_ForcePacking) / dbl_AreaUnbalancedClosed
Else
    dbl_AllowDPClosed = dbl_Temp / (dbl_AreaUnbalancedClosed + dbl_ForcePacking / 1000#)
End If
Exit Sub

CalcAllowDPClosed_41000_FTC_Error:
'Placeholder

End Sub

Public Sub CalcAllowDPOpen_41000_FTC(dbl_MaxInletPressure As Double, dbl_ActForceOpen As Double, _
    dbl_ForcePacking As Double, dbl_ForceActSeal As Double, dbl_ForcePositionerLoss As Double, _
    dbl_ForceSealRingSpring As Double, dbl_AreaUnbalancedOpen As Double, dbl_ForcePlugWt _
    As Double, dbl_AllowDPOpen As Double)
'--------------------------------------------------------------------------------------
'Reference : CES 1019 Balanced and pilot balanced cage throttling globe valves - Actuator Sizing Methods.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the allowable pressure drop in the closed position
'for 41000 series balanced trim package when the flow direction is flow to close.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_MaxInletPressure     pressure that actuator must operate against           psia
'dbl_ActuatorForceOpen    actuator force provided in open position            lbs
'dbl_ForcePacking         force required to overcome packing friction           lbs
'dbl_ForceActSeal         force required to overcome actuator seal friction     lbs
'dbl_ForcePositionerLoss  force required to compensate for positioner loss      lbs
'dbl_ForceSealRingSpring  force required to overcome plug seal ring friction (spring)   lbs
'dbl_ForcePilotSpring     force required to compress pilot spring               lbs
'dbl_ForcePlugWt          force required to overcome plug weight                lbs
'dbl_AreaUnbalancedOpen   effective unbalanced area when valve is open          sq.in.

'Output Variables:
'dbl_AllowDPOpen          allowable pressure with specified actuator force available   psi
'
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/14/98  Name: David Zhou
'Tested     Date: 10/19/98  Name: David Zhou     Passed(Y/N):Y
'-----------------------------------------------------------------------------------------
'**************************
'Declarations
'**************************
Dim dbl_Temp As Double
Const PI As Double = 3.1415927

On Error GoTo CalcAllowDPOpen_41000_FTC_Error

'***************************
'Parameter validation
'***************************
If dbl_MaxInletPressure < 0# Then GoTo CalcAllowDPOpen_41000_FTC_Error
If dbl_ActForceOpen < 0# Then GoTo CalcAllowDPOpen_41000_FTC_Error
If dbl_ForcePacking < 0# Then GoTo CalcAllowDPOpen_41000_FTC_Error
If dbl_ForceActSeal < 0# Then GoTo CalcAllowDPOpen_41000_FTC_Error
If dbl_ForcePositionerLoss < 0# Then GoTo CalcAllowDPOpen_41000_FTC_Error
If dbl_ForceSealRingSpring < 0# Then GoTo CalcAllowDPOpen_41000_FTC_Error
If dbl_ForcePlugWt < 0# Then GoTo CalcAllowDPOpen_41000_FTC_Error

'****************************
'Begin calculations
'****************************
dbl_Temp = dbl_ActForceOpen - dbl_ForceActSeal - dbl_ForcePositionerLoss - _
           dbl_ForceSealRingSpring - dbl_ForcePlugWt

If (dbl_MaxInletPressure <= 2000#) Then
    dbl_AllowDPOpen = (dbl_Temp - dbl_ForcePacking) / dbl_AreaUnbalancedOpen
Else
    dbl_AllowDPOpen = dbl_Temp / (dbl_AreaUnbalancedOpen + dbl_ForcePacking / 1000#)
End If
Exit Sub

CalcAllowDPOpen_41000_FTC_Error:
'Placeholder

End Sub

Public Sub CalcRequiredSeatingForce_41000(dbl_MaxInletPressure As Double, dbl_ForceSeating As Double, _
    dbl_ForcePacking As Double, dbl_ForceActSeal As Double, dbl_ForcePositionerLoss As Double, _
    dbl_ForceSealRingArea As Double, dbl_ForcePlugWt As Double, dbl_RequiredSeatingForce As Double)
'--------------------------------------------------------------------------------------
'Reference : CES 1019 Balanced and pilot balanced cage throttling globe valves - Actuator Sizing Methods.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the force necessary to seat a 41000 series balanced
'trim package when the flow direction is flow to close.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_MaxInletPressure     pressure that actuator must operate against       psia
'dbl_ForceSeating         force required to provide desired leakage class   lbs
'dbl_ForcePacking         " to overcome packing friction                    lbs
'dbl_ForceActSeal         " to overcome actuator seal friction              lbs
'dbl_ForcePositionerLoss  " to compensate for positioner loss               lbs
'dbl_ForceSealRingArea    " to overcome plug seal ring friction (area)      lbs
'dbl_ForcePlugWt          " to overcome plug weight                         lbs

'Output Variables:
'dbl_RequiredSeatingForce " to provide necessary seating force              lbs
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/15/98  Name: David Zhou
'Tested     Date: 10/19/98  Name: David Zhou     Passed(Y/N):Y
'-----------------------------------------------------------------------------------------
On Error GoTo CalcRequiredSeatingForce_41000_Error

'**************************
'Declarations
'**************************

'***************************
'Parameter validation
'***************************
If dbl_MaxInletPressure < 0# Then GoTo CalcRequiredSeatingForce_41000_Error
If dbl_ForceSeating < 0# Then GoTo CalcRequiredSeatingForce_41000_Error
If dbl_ForcePacking < 0# Then GoTo CalcRequiredSeatingForce_41000_Error
If dbl_ForceActSeal < 0# Then GoTo CalcRequiredSeatingForce_41000_Error
If dbl_ForcePositionerLoss < 0# Then GoTo CalcRequiredSeatingForce_41000_Error
If dbl_ForceSealRingArea < 0# Then GoTo CalcRequiredSeatingForce_41000_Error
If dbl_ForcePlugWt < 0# Then GoTo CalcRequiredSeatingForce_41000_Error

'****************************
'Begin calculations
'****************************
dbl_RequiredSeatingForce = dbl_MaxInletPressure * dbl_ForceSealRingArea + dbl_ForceSeating + _
    dbl_ForcePacking + dbl_ForceActSeal + dbl_ForcePositionerLoss - dbl_ForcePlugWt

Exit Sub

CalcRequiredSeatingForce_41000_Error:
'Placeholder

End Sub

Public Sub CalcMainPlugSeatStress_41000(dbl_MaxInletPressure As Double, dbl_SeatDiameter _
    As Double, dbl_StemDiameter As Double, dbl_AllowSeatYieldStress As Double, _
    dbl_OffBalanceForce As Double, dbl_AllowSeatForce As Double)
'--------------------------------------------------------------------------------------
'Reference : CES 1019 Balanced and pilot balanced cage throttling globe valves - Actuator Sizing Methods.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the force necessary to seat a 41000 series balanced
'trim package when the flow direction is flow to close.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_MaxInletPressure     pressure that actuator must operate against           psia
'dbl_SeatDiameter         seat ring orifice diameter                            inches
'dbl_StemDiameter         diameter of the plug stem                             inches
'dbl_AllowSeatYieldStress allowable yield stress of seat                        psi
'---------------------------------------------------------------------------------------
'Output Variables:
'dbl_OffbalanceForce      force resulting from pressure acting on pilot plug      lbs
'dbl_AllowSeatForce       allowable force on seat based on specified stress level lbs
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/15/98  Name: David Zhou
'Tested     Date: 10/19/98  Name: David Zhou     Passed(Y/N):Y
'-----------------------------------------------------------------------------------------
On Error GoTo CalcMainPlugSeatStress_41000_Error

'**************************
'Declarations
'**************************
Const PI As Double = 3.1415927


'***************************
'Parameter validation
'***************************
If dbl_MaxInletPressure < 0# Then GoTo CalcMainPlugSeatStress_41000_Error
If dbl_SeatDiameter < 0# Then GoTo CalcMainPlugSeatStress_41000_Error
If dbl_StemDiameter < 0# Then GoTo CalcMainPlugSeatStress_41000_Error
If dbl_AllowSeatYieldStress < 0# Then GoTo CalcMainPlugSeatStress_41000_Error

'****************************
'Begin calculations
'****************************
dbl_OffBalanceForce = dbl_MaxInletPressure * (dbl_SeatDiameter * dbl_SeatDiameter - _
    dbl_StemDiameter * dbl_StemDiameter) * PI / 4
dbl_AllowSeatForce = 0.034 * dbl_AllowSeatYieldStress * (dbl_SeatDiameter + 0.17) * PI
Exit Sub

CalcMainPlugSeatStress_41000_Error:
'Placeholder

End Sub

Public Sub CalcAllowDP_41000_FTC(dbl_AllowDPClosed As Double, dbl_AllowDPOpen As Double, _
    dbl_SpringRate As Double, dbl_Kv As Double, dbl_MaxAllowDP As Double)
'--------------------------------------------------------------------------------------
'Reference : CES 1019 Balanced and pilot balanced cage throttling globe valves - Actuator Sizing Methods.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the allowable pressure drop for a 41000 series valve
'when the flow direction is FTC.  The resust is the smallest valve calculated for both static and dynamic conditions.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_AllowDPclosed   calculated allowable DP when valve is closed               psi
'dbl_AllowDPopen     calculated allowable DP when valve is open                 psi
'dbl_SpringRate      spring rate of selected actuator                           lbs/in
'dbl_Kv              unbalanced area spring rate/psi for selected valve/trim    in^2/in

'Output Variables:
'dbl_MaxAllowDP      max allowable pressure drop for the selected valve/actuator psi
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/15/98  Name: David Zhou
'Tested     Date: 10/19/98  Name: David Zhou     Passed(Y/N):Y
'-----------------------------------------------------------------------------------------
On Error GoTo CalcAllowDP_41000_FTC_Error

'**************************
'Declarations
'**************************
Dim dbl_MaxAllowDPStability As Double

'***************************
'Parameter validation
'***************************
If dbl_AllowDPClosed < 0# Then GoTo CalcAllowDP_41000_FTC_Error
If dbl_AllowDPOpen < 0# Then GoTo CalcAllowDP_41000_FTC_Error
If dbl_SpringRate < 0# Then GoTo CalcAllowDP_41000_FTC_Error
If dbl_Kv < 0# Then GoTo CalcAllowDP_41000_FTC_Error

'****************************
'Begin calculations
'****************************
dbl_MaxAllowDPStability = dbl_SpringRate / dbl_Kv
If (dbl_MaxAllowDPStability > dbl_AllowDPOpen) Then
    dbl_MaxAllowDP = dbl_AllowDPOpen
Else
    dbl_MaxAllowDP = dbl_MaxAllowDPStability
End If

If (dbl_MaxAllowDP > dbl_AllowDPClosed) Then
    dbl_MaxAllowDP = dbl_AllowDPClosed
End If
Exit Sub

CalcAllowDP_41000_FTC_Error:
'Placeholder

End Sub

Public Sub CalcUnbalArea_41000_Pilot(dbl_SeatDiameter As Double, dbl_CageInsideDiameter _
    As Double, dbl_StemDiameter As Double, dbl_PlugSealDiameter As Double, _
    dbl_PlugThrottlingDiameter As Double, dbl_SealRingCoef As Double, dbl_SealRingWidth _
    As Double, dbl_B As Double, dbl_RatedCv As Double, str_FluidType As String, _
    dbl_AreaUnbalancedOpen As Double, dbl_AreaUnbalancedClosed As Double)

'--------------------------------------------------------------------------------------
'Reference : CES 1019 Balanced and pilot balanced cage throttling globe valves - Actuator Sizing Methods.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the effective unbalanced are for  41000 series valves
'with pilot style plugs.
'---------------------------------------------------------------------------------------
'dbl_SeatDiameter          seat ring orifice diameter                           in.
'dbl_CageInsideDiameter    inside diameter of cage where contacted by seal ring in.
'dbl_StemDiameter          diameter of the plug stem                            in.
'dbl_PlugSealDiameter      diameter of balanced plug next to seal ring groove   in.
'dbl_PlugThrottlingDiameter diameter at bottom of plug                          in.
'dbl_SealRingCoef          friction coefficient for specified seal ring material
'dbl_SealRingWidth         width of actuator seal ring                          in.
'dbl_RatedCv               rated Cv of the selected valve
'str_FluidType             type of fluid designated by user (liquid, gas, stem, two phase flow)
'dbl_B                     relative capacity
'                          use B = 0.1 with cyl./piston actuator
'                              B = 1 with other actuators
'
'Output Variables:
'dbl_AreaUnbalancedOpen    effective unbalanced area when valve is open         sq.in.
'dbl_AreaUnbalancedClosed  effective unbalanced area when valve is closed       sq.in.
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/15/98  Name: David Zhou
'Tested     Date: 10/19/98  Name: David Zhou     Passed(Y/N):Y
'-----------------------------------------------------------------------------------------

On Error GoTo CalcUnbalArea_41000_Pilot_Error

'***************************
'Parameter validation
'***************************
If dbl_SeatDiameter < 0# Then GoTo CalcUnbalArea_41000_Pilot_Error
If dbl_CageInsideDiameter < 0# Then GoTo CalcUnbalArea_41000_Pilot_Error
If dbl_StemDiameter < 0# Then GoTo CalcUnbalArea_41000_Pilot_Error
If dbl_PlugSealDiameter < 0# Then GoTo CalcUnbalArea_41000_Pilot_Error
If dbl_PlugThrottlingDiameter < 0# Then GoTo CalcUnbalArea_41000_Pilot_Error
If dbl_SealRingCoef < 0# Then GoTo CalcUnbalArea_41000_Pilot_Error
If dbl_SealRingWidth < 0# Then GoTo CalcUnbalArea_41000_Pilot_Error
If dbl_B < 0# Then GoTo CalcUnbalArea_41000_Pilot_Error
If dbl_RatedCv < 0# Then GoTo CalcUnbalArea_41000_Pilot_Error

'**************************
'Declarations
'**************************
Dim dbl_A As Double
Dim dbl_A2 As Double
Dim dbl_A3 As Double
Dim dbl_As As Double
Dim dbl_C5 As Double
Dim dbl_E As Double
Dim dbl_FdMax As Double
Dim dbl_FluidCoef As Double
Const PI As Double = 3.1415927
Const dbl_SmallNum = 1E-20

'****************************
'Begin calculations
'****************************
dbl_A = 1#
dbl_E = dbl_A
If (Abs(dbl_B) < dbl_SmallNum) Then dbl_B = 1#
If (str_FluidType = "Liquid" Or str_FluidType = "Two-Phase") Then
    dbl_FluidCoef = 0.5
Else
    dbl_FluidCoef = 0.23
End If
dbl_A2 = dbl_SeatDiameter * dbl_SeatDiameter * PI / 4
dbl_As = dbl_StemDiameter * dbl_StemDiameter * PI / 4
dbl_A3 = (dbl_PlugThrottlingDiameter ^ 2 - dbl_CageInsideDiameter ^ 2) * PI / 4
dbl_C5 = dbl_SealRingCoef * PI * dbl_CageInsideDiameter * dbl_SealRingWidth / 2
dbl_FdMax = dbl_FluidCoef * dbl_RatedCv / (38# * dbl_SeatDiameter * dbl_SeatDiameter * PI / 4)
dbl_AreaUnbalancedClosed = dbl_A3 + dbl_C5
dbl_AreaUnbalancedOpen = -dbl_A2 * dbl_FdMax * dbl_B ^ 3 - dbl_As * dbl_E * dbl_B * dbl_B + _
        1.2 * dbl_A2 * dbl_FdMax * dbl_B + dbl_A3 + dbl_C5

'change as per J Stares 12/15/98 to keep from going negative
If (dbl_AreaUnbalancedOpen < 0.1) Then dbl_AreaUnbalancedOpen = 0.1


Exit Sub

CalcUnbalArea_41000_Pilot_Error:
'Placeholder

End Sub

Public Sub CalcUnbalArea_41000_Nonpilot(dbl_SeatDiameter As Double, dbl_CageInsideDiameter _
    As Double, dbl_StemDiameter As Double, dbl_PlugSealDiameter As Double, _
    dbl_PlugThrottlingDiameter As Double, dbl_SealRingCoef As Double, dbl_SealRingWidth _
    As Double, dbl_B As Double, dbl_DynPlugDiameter, dbl_RatedCv As Double, str_FluidType _
    As String, dbl_AreaUnbalancedOpen As Double, dbl_AreaUnbalancedClosed As Double)

'--------------------------------------------------------------------------------------
'Reference : CES 1019 Balanced and pilot balanced cage throttling globe valves - Actuator Sizing Methods.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the effective unbalanced are for  41000 series valves
'with non-pilot style plugs.
'---------------------------------------------------------------------------------------
'dbl_SeatDiameter          seat ring orifice diameter                           in.
'dbl_CageInsideDiameter    inside diam of cage where contacted by seal ring     in.
'dbl_StemDiameter          diameter of the plug stem                            in.
'dbl_PlugSealDiameter      diameter of balanced plug next to seal ring groove   in.
'dbl_PlugThrottlingDiameter diameter at bottom of plug                          in.
'dbl_DynPlugDiameter       dynamically unaffected diameter                      in.
'dbl_SealRingCoef          friction coefficient for specified seal ring material
'dbl_SealRingWidth         width of actuator seal ring                          in.
'dbl_RatedCv               rated Cv of the selected valve
'str_FluidType             type of fluid designated by user (liquid, gas, stem, two phase flow)
'dbl_B                     relative capacity
'                          use B = 0.1 with cyl./piston actuator
'                              B = 1 with other actuators
'
'Output Variables:
'dbl_AreaUnbalancedOpen    effective unbalanced area when valve is open         sq.in.
'dbl_AreaUnbalancedClosed  effective unbalanced area when valve is closed       sq.in.
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/15/98  Name: David Zhou
'Tested     Date: 10/19/98  Name: David Zhou     Passed(Y/N):Y
'-----------------------------------------------------------------------------------------

On Error GoTo CalcUnbalArea_41000_Nonpilot_Error

'***************************
'Parameter validation
'***************************
If dbl_SeatDiameter < 0# Then GoTo CalcUnbalArea_41000_Nonpilot_Error
If dbl_CageInsideDiameter < 0# Then GoTo CalcUnbalArea_41000_Nonpilot_Error
If dbl_StemDiameter < 0# Then GoTo CalcUnbalArea_41000_Nonpilot_Error
If dbl_PlugSealDiameter < 0# Then GoTo CalcUnbalArea_41000_Nonpilot_Error
If dbl_PlugThrottlingDiameter < 0# Then GoTo CalcUnbalArea_41000_Nonpilot_Error
If dbl_DynPlugDiameter < 0# Then GoTo CalcUnbalArea_41000_Nonpilot_Error
If dbl_SealRingCoef < 0# Then GoTo CalcUnbalArea_41000_Nonpilot_Error
If dbl_SealRingWidth < 0# Then GoTo CalcUnbalArea_41000_Nonpilot_Error
If dbl_B < 0# Then GoTo CalcUnbalArea_41000_Nonpilot_Error
If dbl_RatedCv < 0# Then GoTo CalcUnbalArea_41000_Nonpilot_Error

'**************************
'Declarations
'**************************
Dim dbl_A As Double
Dim dbl_A2 As Double
Dim dbl_A3 As Double
Dim dbl_As As Double
Dim dbl_C5 As Double
Dim dbl_E As Double
Dim dbl_FdMax As Double
Dim dbl_FluidCoef As Double
Const PI As Double = 3.1415927
Const dbl_SmallNum = 1E-20

'****************************
'Begin calculations
'****************************
dbl_A = 1#
dbl_E = dbl_A
If (Abs(dbl_B) < dbl_SmallNum) Then dbl_B = 1#
If (str_FluidType = "Liquid" Or str_FluidType = "Two-Phase") Then
    dbl_FluidCoef = 0.5
Else
    dbl_FluidCoef = 0.23
End If
dbl_A2 = (dbl_SeatDiameter * dbl_SeatDiameter - dbl_DynPlugDiameter * dbl_DynPlugDiameter) * PI / 4
dbl_As = dbl_StemDiameter * dbl_StemDiameter * PI / 4
dbl_A3 = (dbl_PlugSealDiameter ^ 2 - dbl_CageInsideDiameter ^ 2) * PI / 4
dbl_C5 = dbl_SealRingCoef * PI * dbl_CageInsideDiameter * dbl_SealRingWidth / 2
dbl_FdMax = dbl_FluidCoef * dbl_RatedCv / (38# * dbl_SeatDiameter * dbl_SeatDiameter * PI / 4)
dbl_AreaUnbalancedClosed = dbl_A3 + dbl_C5
dbl_AreaUnbalancedOpen = -dbl_A2 * dbl_FdMax * dbl_B ^ 3 - dbl_As * dbl_E * dbl_B * dbl_B + _
        1.2 * dbl_A2 * dbl_FdMax * dbl_B + dbl_A3 + dbl_C5
'change as per J Stares 12/15/98 to keep from going negative
If (dbl_AreaUnbalancedOpen < 0.1) Then dbl_AreaUnbalancedOpen = 0.1

Exit Sub

CalcUnbalArea_41000_Nonpilot_Error:
'Placeholder

End Sub


Public Sub CalcSpringRateDACylinderATO(dbl_PInitial As Double, dbl_AreaInitial As Double, _
    dbl_SpringRate As Double, dbl_Travel As Double, str_SpecifiedLengthUM As String, _
    dbl_AreaUnbalancedOpen As Double, dbl_ShutOffDP As Double, dbl_MaxSupplyPressure _
    As Double, dbl_EffectiveSpringRate As Double)
'----------------------------------------------------------------------------------------
'Reference : CES 1019 Balanced and pilot balanced cage throttling globe valves
'            - Actuator Sizing Methods.
'----------------------------------------------------------------------------------------
'This subroutine is used to calculate the effective spring rate for ATO double acting
'cylinders.   The effective spring rate considers both the actuator spring and the
'effect of the volume of air in actuator.
'----------------------------------------------------------------------------------------
'Input Variables:
'dbl_Pinitial            initial spring setting of selected actuator          psig
'dbl_AreaInitial         area of actuator piston                              in^2
'dbl_SpringRate          spring rate of actuator spring                       lb/in
'dbl_Travel              rated plug travel                                    user def.
'str_SpecifiedLengthUM   unit of measure selected for travel                  user def.
'dbl_AreaUnbalancedOpen  effective unbalanced area of plug when valve is open in^2
'dbl_ShutoffDP           maximum pressure that the valve must shut against    psi
'dbl_MaxSupplyPressure   maximum supply pressure available                    psig

'Output Variables:
'dbl_EffectiveSpringRate effective spring rate of spring and effect of        lb/in
'                        volume of air in actuator
'----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/16/98  Name: David Zhou
'Tested     Date: 10/19/98  Name: David Zhou     Passed(Y/N):Y
'----------------------------------------------------------------------------------------

On Error GoTo CalcSpringRateDACylinderATO_Error

'***************************
'Parameter validation
'***************************
If dbl_PInitial < 0# Then GoTo CalcSpringRateDACylinderATO_Error
If dbl_AreaInitial < 0# Then GoTo CalcSpringRateDACylinderATO_Error
If dbl_SpringRate < 0# Then GoTo CalcSpringRateDACylinderATO_Error
If dbl_Travel < 0# Then GoTo CalcSpringRateDACylinderATO_Error
If dbl_ShutOffDP < 0# Then GoTo CalcSpringRateDACylinderATO_Error
If dbl_MaxSupplyPressure < 0# Then GoTo CalcSpringRateDACylinderATO_Error

'**************************
'Declarations
'**************************
Dim dbl_A As Double
Dim dbl_B As Double
Dim dbl_Fp As Double
Dim dbl_Pb As Double
Dim dbl_Va As Double
Dim dbl_Vb As Double
Dim dbl_Temp As Double
Dim dbl_Temp2 As Double
Dim dbl_Temp3 As Double
Dim dbl_Temp4 As Double
Dim dbl_SpecificHeatRatio As Double
Dim dbl_ThisTravel As Double
Dim dbl_SpringTravel As Double

'****************************
'Begin calculations
'****************************
dbl_SpecificHeatRatio = 1.4
dbl_A = 1#
dbl_B = 0.1

dbl_Fp = dbl_AreaUnbalancedOpen * (dbl_ShutOffDP / (1 + dbl_A * dbl_B * dbl_B))
dbl_ThisTravel = ConvertLengthToInches(dbl_Travel, str_SpecifiedLengthUM)
dbl_SpringTravel = dbl_PInitial * dbl_AreaInitial / dbl_SpringRate + dbl_ThisTravel
dbl_Pb = dbl_MaxSupplyPressure / 2
dbl_Va = (dbl_ThisTravel + 0.4) * dbl_AreaInitial
dbl_Vb = 16.4 * dbl_AreaInitial - dbl_Va

dbl_Temp = dbl_SpecificHeatRatio * dbl_AreaInitial * dbl_AreaInitial
dbl_Temp2 = (dbl_SpringRate * dbl_SpringTravel + dbl_Fp) / (2# * dbl_AreaInitial)
dbl_Temp3 = (dbl_Pb - dbl_Temp2) / (dbl_Va - dbl_AreaInitial * dbl_ThisTravel)
dbl_Temp4 = (dbl_Pb + dbl_Temp2) / (dbl_Vb + dbl_AreaInitial * dbl_ThisTravel)
dbl_EffectiveSpringRate = dbl_Temp * (dbl_Temp3 + dbl_Temp4) + dbl_SpringRate

Exit Sub

CalcSpringRateDACylinderATO_Error:
'Placeholder

End Sub

Public Sub CalcSpringRateDACylinderATC(dbl_PInitial As Double, dbl_AreaInitial As Double, _
    dbl_SpringRate As Double, dbl_Travel As Double, str_SpecifiedLengthUM As String, _
    dbl_AreaUnbalancedOpen As Double, dbl_ShutOffDP As Double, dbl_MaxSupplyPressure _
    As Double, dbl_EffectiveSpringRate As Double)
'----------------------------------------------------------------------------------------
'Reference : CES 1019 Balanced and pilot balanced cage throttling globe valves
'            - Actuator Sizing Methods.
'----------------------------------------------------------------------------------------
'This subroutine is used to calculate the effective spring rate for ATC double acting
'cylinders.   The effective spring rate considers both the actuator spring and
'the effect of the volume of air in actuator.
'----------------------------------------------------------------------------------------
'Input Variables:
'dbl_Pinitial             initial spring setting of selected actuator           psig
'dbl_AreaInitial          area of actuator piston                               in^2
'dbl_SpringRate           spring rate of actuator spring                        lb/in
'dbl_Travel               rated plug travel                                     user def.
'str_SpecifiedLengthUM    unit of measure selected for travel                   user def.
'dbl_AreaUnbalancedOpen   effective unbalanced area of plug when valve is open  in^2
'dbl_ShutoffDP            maximum pressure that the valve must shut against     psi
'dbl_MaxSupplyPressure    maximum supply pressure available                     psig

'Output Variables:
'dbl_EffectiveSpringRate  effective spring rate of spring and effect of         lb/in
'                         volume of air in actuator
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/16/98  Name: David Zhou
'Tested     Date: 10/19/98  Name: David Zhou     Passed(Y/N):Y
'-----------------------------------------------------------------------------------------

On Error GoTo CalcSpringRateDACylinderATC_Error

'***************************
'Parameter validation
'***************************
If dbl_PInitial < 0# Then GoTo CalcSpringRateDACylinderATC_Error
If dbl_AreaInitial < 0# Then GoTo CalcSpringRateDACylinderATC_Error
If dbl_SpringRate < 0# Then GoTo CalcSpringRateDACylinderATC_Error
If dbl_Travel < 0# Then GoTo CalcSpringRateDACylinderATC_Error
If dbl_ShutOffDP < 0# Then GoTo CalcSpringRateDACylinderATC_Error
If dbl_MaxSupplyPressure < 0# Then GoTo CalcSpringRateDACylinderATC_Error

'**************************
'Declarations
'**************************
Dim dbl_A As Double
Dim dbl_B As Double
Dim dbl_Fp As Double
Dim dbl_Pb As Double
Dim dbl_Va As Double
Dim dbl_Vb As Double
Dim dbl_Temp As Double
Dim dbl_Temp2 As Double
Dim dbl_Temp3 As Double
Dim dbl_Temp4 As Double
Dim dbl_SpecificHeatRatio As Double
Dim dbl_ThisTravel As Double
Dim dbl_SpringTravel As Double

'****************************
'Begin calculations
'****************************
dbl_SpecificHeatRatio = 1.4
dbl_A = 1#
dbl_B = 0.1

dbl_Fp = dbl_AreaUnbalancedOpen * (dbl_ShutOffDP / (1 + dbl_A * dbl_B * dbl_B))
dbl_ThisTravel = ConvertLengthToInches(dbl_Travel, str_SpecifiedLengthUM)
dbl_SpringTravel = dbl_PInitial * dbl_AreaInitial / dbl_SpringRate + dbl_ThisTravel
dbl_Pb = dbl_MaxSupplyPressure / 2
dbl_Va = (dbl_ThisTravel + 0.4) * dbl_AreaInitial
dbl_Vb = 16.4 * dbl_AreaInitial - dbl_Va

dbl_Temp = dbl_SpecificHeatRatio * dbl_AreaInitial * dbl_AreaInitial
dbl_Temp2 = (dbl_SpringRate * dbl_SpringTravel + dbl_Fp) / (2# * dbl_AreaInitial)
dbl_Temp3 = (dbl_Pb + dbl_Temp2) / (dbl_Va - dbl_AreaInitial * dbl_ThisTravel)
dbl_Temp4 = (dbl_Pb - dbl_Temp2) / (dbl_Vb + dbl_AreaInitial * dbl_ThisTravel)
dbl_EffectiveSpringRate = dbl_Temp * (dbl_Temp3 + dbl_Temp4) + dbl_SpringRate

Exit Sub

CalcSpringRateDACylinderATC_Error:
'Placeholder

End Sub

Public Sub CalcSpringRateSACylinderATO(dbl_PInitial As Double, dbl_AreaInitial As Double, _
    dbl_SpringRate As Double, dbl_Travel As Double, str_SpecifiedLengthUM As String, _
    dbl_AreaUnbalancedOpen As Double, dbl_ShutOffDP As Double, dbl_MaxSupplyPressure _
    As Double, dbl_EffectiveSpringRate)
'--------------------------------------------------------------------------------------
'Reference : CES 1019 Balanced and pilot balanced cage throttling globe valves - Actuator Sizing Methods.
'---------------------------------------------------------------------------------------
'This subroutine is used to calculate the effective spring rate for ATO single acting
'cylinders.   The effective spring rate considers both the actuator spring and the effect of the volume of air in actuator.
'---------------------------------------------------------------------------------------
'Input Variables:
'dbl_Pinitial            initial spring setting of selected actuator            psig
'dbl_AreaInitial         area of actuator piston                                in^2
'str_SpringRate          spring rate of actuator spring                         lb/in
'dbl_Travel              rated plug travel                                      user def.
'str_SpecifiedLengthUM   unit of measure selected for travel                    user def.
'dbl_AreaUnbalancedOpen  effective unbalanced area of plug when valve is open   in^2
'dbl_ShutoffDP           maximum pressure that the valve must shut against      psi
'dbl_MaxSupplyPressure   maximum supply pressure available                      psig

'Output Variables:
'dbl_EffectiveSpringRate effective spring rate of spring and effect of          lb/in
'                        the volume of air in actuator
'-----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/16/98  Name: David Zhou
'Tested     Date: 10/19/98  Name: David Zhou     Passed(Y/N):Y
'-----------------------------------------------------------------------------------------

On Error GoTo CalcSpringRateSACylinderATO_Error

'***************************
'Parameter validation
'***************************
If dbl_PInitial < 0# Then GoTo CalcSpringRateSACylinderATO_Error
If dbl_AreaInitial < 0# Then GoTo CalcSpringRateSACylinderATO_Error
If dbl_SpringRate < 0# Then GoTo CalcSpringRateSACylinderATO_Error
If dbl_Travel < 0# Then GoTo CalcSpringRateSACylinderATO_Error
If dbl_ShutOffDP < 0# Then GoTo CalcSpringRateSACylinderATO_Error
If dbl_MaxSupplyPressure < 0# Then GoTo CalcSpringRateSACylinderATO_Error

'**************************
'Declarations
'**************************
Dim dbl_A As Double
Dim dbl_B As Double
Dim dbl_Fp As Double
Dim dbl_Pb As Double
Dim dbl_Va As Double
Dim dbl_Vb As Double
Dim dbl_Temp As Double
Dim dbl_Temp2 As Double
Dim dbl_Temp3 As Double
Dim dbl_SpecificHeatRatio As Double
Dim dbl_ThisTravel As Double
Dim dbl_SpringTravel As Double

'****************************
'Begin calculations
'****************************
dbl_SpecificHeatRatio = 1.4
dbl_A = 1#
dbl_B = 0.1

dbl_Fp = dbl_AreaUnbalancedOpen * (dbl_ShutOffDP / (1 + dbl_A * dbl_B * dbl_B))
dbl_ThisTravel = ConvertLengthToInches(dbl_Travel, str_SpecifiedLengthUM)
dbl_SpringTravel = dbl_PInitial * dbl_AreaInitial / dbl_SpringRate + dbl_ThisTravel
dbl_Pb = dbl_MaxSupplyPressure / 2
dbl_Va = (dbl_ThisTravel + 0.4) * dbl_AreaInitial
dbl_Vb = 16.4 * dbl_AreaInitial - dbl_Va

dbl_Temp = dbl_SpecificHeatRatio * dbl_AreaInitial * dbl_AreaInitial
dbl_Temp2 = (dbl_SpringRate * dbl_SpringTravel + dbl_Fp) / dbl_AreaInitial
dbl_Temp3 = (14.7 + dbl_Temp2) / (dbl_Vb + dbl_AreaInitial * dbl_ThisTravel)
dbl_EffectiveSpringRate = dbl_Temp * dbl_Temp3 + dbl_SpringRate

Exit Sub

CalcSpringRateSACylinderATO_Error:
'Placeholder

End Sub


Public Sub CalcSpringRateSACylinderATC(dbl_PInitial As Double, dbl_AreaInitial As Double, _
    dbl_SpringRate As Double, dbl_Travel As Double, str_SpecifiedLengthUM As String, _
    dbl_AreaUnbalancedOpen As Double, dbl_ShutOffDP As Double, dbl_MaxSupplyPressure _
    As Double, dbl_EffectiveSpringRate)
'----------------------------------------------------------------------------------------
'Reference : CES 1019 Balanced and pilot balanced cage throttling globe valves
'            - Actuator Sizing Methods.
'----------------------------------------------------------------------------------------
'This subroutine is used to calculate the effective spring rate for ATC single acting
'cylinders.   The effective spring rate considers both the actuator spring and the effect
'of the volume of air in actuator.
'----------------------------------------------------------------------------------------
'Input Variables:
'dbl_Pinitial               initial spring setting of selected actuator         psig
'dbl_AreaInitial            area of actuator piston                             in^2
'dbl_SpringRate             spring rate of actuator spring                      lb/in
'dbl_Travel                 rated plug travel                                   user def.
'str_SpecifiedLengthUM      unit of measure selected for travel                 user def.
'dbl_AreaUnbalancedOpen     effective unbalanced area of plug
'                           when the  valve is open                             in^2
'dbl_ShutoffDP              maximum pressure that the valve must shut against   psi
'dbl_MaxSupplyPressure      maximum supply pressure available                   psig

'Output Variables:
'dbl_EffectiveSpringRate    effective spring rate of spring and effect of the   lb/in
'                           volume of air in actuator
'----------------------------------------------------------------------------------------
'Maintenance:
'Coded      Date: 10/16/98  Name: David Zhou
'Tested     Date: 10/19/98  Name: David Zhou     Passed(Y/N):Y
'----------------------------------------------------------------------------------------

On Error GoTo CalcSpringRateSACylinderATC_Error

'***************************
'Parameter validation
'***************************
If dbl_PInitial < 0# Then GoTo CalcSpringRateSACylinderATC_Error
If dbl_AreaInitial < 0# Then GoTo CalcSpringRateSACylinderATC_Error
If dbl_SpringRate < 0# Then GoTo CalcSpringRateSACylinderATC_Error
If dbl_Travel < 0# Then GoTo CalcSpringRateSACylinderATC_Error
If dbl_ShutOffDP < 0# Then GoTo CalcSpringRateSACylinderATC_Error
If dbl_MaxSupplyPressure < 0# Then GoTo CalcSpringRateSACylinderATC_Error

'**************************
'Declarations
'**************************
Dim dbl_A As Double
Dim dbl_B As Double
Dim dbl_Fp As Double
Dim dbl_Pb As Double
Dim dbl_Va As Double
Dim dbl_Vb As Double
Dim dbl_Temp As Double
Dim dbl_Temp2 As Double
Dim dbl_Temp3 As Double
Dim dbl_SpecificHeatRatio As Double
Dim dbl_ThisTravel As Double
Dim dbl_SpringTravel As Double

'****************************
'Begin calculations
'****************************
dbl_SpecificHeatRatio = 1.4
dbl_A = 1#
dbl_B = 0.1

dbl_Fp = dbl_AreaUnbalancedOpen * (dbl_ShutOffDP / (1 + dbl_A * dbl_B * dbl_B))
dbl_ThisTravel = ConvertLengthToInches(dbl_Travel, str_SpecifiedLengthUM)
dbl_SpringTravel = dbl_PInitial * dbl_AreaInitial / dbl_SpringRate + dbl_ThisTravel
dbl_Pb = dbl_MaxSupplyPressure / 2
dbl_Va = (dbl_ThisTravel + 0.4) * dbl_AreaInitial
'dbl_Vb = 16.4 * dbl_AreaInitial - dbl_Va

dbl_Temp = dbl_SpecificHeatRatio * dbl_AreaInitial * dbl_AreaInitial
dbl_Temp2 = (dbl_SpringRate * dbl_SpringTravel + dbl_Fp) / dbl_AreaInitial
dbl_Temp3 = (14.7 + dbl_Temp2) / (dbl_Va - dbl_AreaInitial * dbl_ThisTravel)
dbl_EffectiveSpringRate = dbl_Temp * dbl_Temp3 + dbl_SpringRate

Exit Sub

CalcSpringRateSACylinderATC_Error:
'Placeholder

End Sub

Public Sub CalcMaxDPCage_41000(dbl_ValveSize As Double, lng_TrimTypeID As Long, _
    dbl_MaxInletTemperature As Double, dbl_MaxAllowCageDP As Double)
    
'----------------------------------------------------------------------------------------
'Reference:  CES 1019 Balanced and Pilot Balanced Cage Throttling Globe
'            Values - Actuator Sizing Methods.
'----------------------------------------------------------------------------------------
'This subroutine is used to calculate the maximum allowable pressure drop
'for 41000 series valves based on the mechanical limitation of the cage at
'the maximum service temperature

'Using the valve size and trim type ID interpolate by temperature to find
'the allowable cage pressure drop.
'----------------------------------------------------------------------------------------
'Input Variables:
'dbl_ValveSize          Nominal size of the selected valve          mm
'lng_TrimTypeID         ID# of trim selected (linear, equal%, LODB, etc)
'dbl_MaxInletTemperatureMaximum specified service temperature       deg F
'----------------------------------------------------------------------------------------
'Output Variables:
'dbl_MaxAllowCageDP     Max allowable pressure drop for specified   psi
'                       cage at specified temperature
'----------------------------------------------------------------------------------------

On Error GoTo CalcMaxDPCage_41000_Error:

'*************************
'Declarations
'*************************
Dim str_DatabasePath As String
Dim str_Query As String
Dim rst_Result As Recordset

Dim dbl_Tn320F As Double
Dim dbl_T100F As Double
Dim dbl_T650F As Double
Dim dbl_T1050F As Double

'dbl_ValveSize                  'Independent Variable
Dim dbl_LowValue As Double      'Dependent Variable (MaxAllowCageDP)- Lower Bound
Dim dbl_HighValue As Double     'Dependent Variable (MaxAllowCageDP)- Upper Bound

If dbl_ValveSize <= 0 Then GoTo CalcMaxDPCage_41000_Error
If lng_TrimTypeID <= 0 Then GoTo CalcMaxDPCage_41000_Error

'*************************
'Begin Calculations
'*************************

str_Query = "SELECT Tn320F, T100F, T650F, T1050F From " + _
    "tblMaxDPCage41000 where ((TRIM_TYPE_ID=" + CStr(lng_TrimTypeID) + _
    ") AND (VALVE_SIZE=" + CStr(dbl_ValveSize) + "))"

Set rst_Result = g_db_Gensz.OpenRecordset(str_Query)
If (rst_Result.RecordCount <> 1) Then GoTo CalcMaxDPCage_41000_Error

'Get results
dbl_Tn320F = rst_Result.Fields!Tn320F
dbl_T100F = rst_Result.Fields!T100F
dbl_T650F = rst_Result.Fields!T650F
dbl_T1050F = rst_Result.Fields!T1050F

'Find Upper and Lower Bounds
If dbl_MaxInletTemperature > -320 And dbl_MaxInletTemperature <= 100 Then
    dbl_LowValue = -320
    dbl_HighValue = 100
    Call Interpolate(dbl_MaxInletTemperature, dbl_LowValue, dbl_HighValue, _
        dbl_Tn320F, dbl_T100F, dbl_MaxAllowCageDP)

ElseIf dbl_MaxInletTemperature > 100 And dbl_MaxInletTemperature <= 650 Then
    dbl_LowValue = 100
    dbl_HighValue = 650
    Call Interpolate(dbl_MaxInletTemperature, dbl_LowValue, dbl_HighValue, _
        dbl_T100F, dbl_T650F, dbl_MaxAllowCageDP)

ElseIf dbl_MaxInletTemperature > 650 And dbl_MaxInletTemperature <= 1050 Then
    dbl_LowValue = 650
    dbl_HighValue = 1050
    Call Interpolate(dbl_MaxInletTemperature, dbl_LowValue, dbl_HighValue, _
        dbl_T650F, dbl_T1050F, dbl_MaxAllowCageDP)

Else
    dbl_MaxAllowCageDP = dbl_T1050F

End If

rst_Result.Close

CalcMaxDPCage_41000_Error:
    
End Sub

Sub GetTrimTypeFromPackage(long_TrimPackageID As Long, long_TrimTypeID As Long)

On Error GoTo GetTrimTypeFromPackage_Error

'*************************
'Declarations
'*************************
Dim str_DatabasePath As String
Dim str_Query As String
Dim rst_Result As Recordset

'process

str_Query = "SELECT TRIM_TYPE_ID From " + _
    "tblPackageTrimType where ( PACKAGE_ID=" + CStr(long_TrimPackageID) + ")"

Set rst_Result = g_db_Prod.OpenRecordset(str_Query, dbReadOnly)
If (rst_Result.RecordCount <> 1) Then GoTo GetTrimTypeFromPackage_Error

'get results
long_TrimTypeID = rst_Result.Fields!TRIM_TYPE_ID

rst_Result.Close
Exit Sub

GetTrimTypeFromPackage_Error:
MsgBox "Error getting Trim type =" + CStr(long_TrimPackageID) + " - " + Error
End Sub
Sub GetPlugIDFromPackage(long_TrimPackageID As Long, long_PlugTypeID As Long)

On Error GoTo GetPlugIDFromPackage_Error

'*************************
'Declarations
'*************************
Dim str_DatabasePath As String
Dim str_Query As String
Dim rst_Result As Recordset

'process

str_Query = "SELECT tblTrimPlug.PLUG_TYPE_ID" _
    + " FROM tblTrimPackage INNER JOIN tblTrimPlug ON tblTrimPackage.PART_ID = tblTrimPlug.PART_ID " _
    + "WHERE (((tblTrimPackage.PACKAGE_ID)=" + CStr(long_TrimPackageID) + "))"

Set rst_Result = g_db_Prod.OpenRecordset(str_Query, dbReadOnly)
If (rst_Result.RecordCount <> 1) Then GoTo GetPlugIDFromPackage_Error


'get results
long_PlugTypeID = rst_Result.Fields!PLUG_TYPE_ID

rst_Result.Close

Exit Sub

GetPlugIDFromPackage_Error:
MsgBox "Error Plug ID =" + Error
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' END OF 41000 LL
''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INCLUDE 41000 UL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CalcDP_41000_FTO_ATC(long_TrimPackageID As Long, _
    long_ActuatorPartID As Long, dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, _
    dbl_InletPressure3 As Double, dbl_InletPressure4 As Double, dbl_OutletPressure1 As Double, _
    dbl_OutletPressure2 As Double, dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, _
    dbl_InletTemperature1 As Double, dbl_InletTemperature2 As Double, dbl_InletTemperature3 As Double, _
    dbl_InletTemperature4 As Double, dbl_ShutOffDP As Double, str_SpecifiedPressureUM As String, _
    str_SpecifiedTemperatureUM As String, str_SpecifiedCalibPressureUM As String, _
    str_SpecifiedLengthUM As String, str_BodyMatl As String, _
    str_SelectedRating As String, dbl_ValveSize As Double, _
    str_PackingMaterial As String, str_LeakageClass As String, _
    str_SpringRange As String, dbl_RatedCv As Double, dbl_ShowDP As Double, _
    dbl_ShowDPSupplyPressure As Double, dbl_ShowMaxDP As Double, _
    dbl_ShowMaxDPSupplyPressure As Double, str_Message As String)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'calcDP_41000_FTO_ATC(----see input output list below-)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Reference: CES 1019 Balanced and pilot balanced cage throttling
'   globe valves - Actuator Sizing Methods.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This subroutine is used to calculate the selected valve and
'  actuator shutoff capability and compare the result to defined
'  requirements.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This subroutine is specific to 41000 series valves, flow to open,
'  air to close.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Input Variables
'dbl_InletPressure(1 to 4)    specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)   specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4) specified inlet temperature (all cases)
'dbl_ShutoffDP                specified maximum required shutoff pressure Drop
'str_SpecifiedPressureUM      unit of measure associated with inlet
'                             and outlet pressures
'str_SpecifiedTemperatureUM   unit of measure associated with inlet
'                             temperature
'str_SpecifledCalibPressureUM unit of measure associated with spring
'                             range and supply pressure
'str_SpecifiedLengthUM        units of measure for length
'str_BodyMatl                 selected material of body
'str_SelectedRating           selected body rating
'long_TrimPackageID           ID of trim selected
'long_ActuatorID              ID of actuator selected
'dbl_ValveSize                valve size
'str_PackingMaterial          selected packing material
'str_leakageClass             selected leakage class
'str_SpringRange              selected actuator spring range
'dbl_RatedCv                  selected rated Cv

'Output Variables:
'dbl_ShowDP                   specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure     calculated required supply pressure
'                             for specified
'dbl_ShowMaxDP                calculated maximum pressure drop for
'                             selected valve and actuator combination
'dbl_ShowMaxDpsupplypressure  calculated required supply pressure
'                             for maximum pressure drop
'str_Message                  warning message if pressure drop
'                             requirement exceeds equipment capability
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'coded: 10/12/98    Larry Schoonover
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo CalcDP_41000_FTO_ATC_Error

'***************************
'Declarations
'***************************
Dim dbl_ActuatorFriction As Double
Dim dbl_ActuatorStemDiameter As Double
Dim long_AirActionID As Long
Dim dbl_AllowANSIPressure As Double
Dim dbl_AllowableStemForce As Double
Dim dbl_AllowDP As Double
Dim dbl_AllowDPStem As Double
Dim dbl_AreaInitial As Double
Dim dbl_AreaFinal As Double
Dim dbl_As As Double
Dim dbl_Au As Double
Dim dbl_CageInsideDiameter As Double
Dim dbl_DynFrictionRateChange As Double
Dim dbl_DynPlugDiameter As Double
Dim dbl_E As Double
Dim dbl_Fcs As Double
Dim dbl_ForceActSeal As Double
Dim dbl_ForcePacking As Double
Dim dbl_ForcePilotSpring As Double
Dim dbl_ForcePlugWt As Double
Dim dbl_ForcePositionerLoss As Double
Dim dbl_ForceSealRingArea As Double
Dim dbl_ForceSealRingSpring As Double
Dim dbl_ForceSeating As Double
Dim dbl_Kv As Double
Dim dbl_L As Double
Dim dbl_MaxAllowDP As Double
Dim dbl_MaxAllowDPStem As Double
Dim dbl_MaxAllowTrimDP As Double
Dim dbl_MaxAllowCageDP As Double
Dim dbl_MaxAllowFlatSpring As Double
Dim dbl_MaxAllowFlatSpringDP As Double
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_MaxSupplyPressure As Double
Dim dbl_PackingHeight As Double
Dim dbl_PFinal As Double
Dim dbl_PilotTravel As Double
Dim dbl_PInitial As Double
Dim dbl_PistonRingWidth As Double
Dim dbl_VolOver As Double
Dim dbl_ForceSpringClosed As Double

Dim dbl_PlugSealDiameter As Double
Dim dbl_PlugThrottlingDiameter As Double
Dim dbl_PositionerLossCoef As Double
Dim dbl_PSupplyDPStem As Double
Dim dbl_PSupplyDP As Double
Dim dbl_PSupplyDPANSI As Double
Dim dbl_PSupplyDPShutoff As Double
Dim int_QtySeals As Integer
Dim dbl_RequiredActuatorForce As Double
Dim dbl_RequiredANSIActuatorForce As Double
Dim dbl_RequiredShutoffDPActuatorForce As Double
Dim dbl_Sd As Double
Dim dbl_SealRingCoef As Double
Dim dbl_SealRingWidth As Double
Dim dbl_SeatDiameter As Double
Dim dbl_ShowMaxDPSupply As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_SpringEnergizedSealSpringCoef As Double
Dim dbl_SpringRate As Double
Dim dbl_StemDiameter As Double
Dim dbl_StemSealWidth As Double
Dim dbl_StemStressArea As Double
Dim dbl_Sy As Double
Dim dbl_TECCoefJ As Double
Dim dbl_TECCoefS As Double
Dim dbl_TECSealCoef As Double
Dim dbl_Travel As Double
Dim long_CageTypeID As Long
Dim long_PlugTypeID As Long
Dim long_TrimTypeID As Long
Dim str_Model As String
Dim str_StemMaterial As String
Dim dbl_SeatingDiameter As Double
Dim dbl_RequiredActuatorForceLimit As Double

Dim dbl_InletPressure(4) As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double
Dim dbl_LeverLength As Double

#If (DebugVersion = "yes") Then
MsgBox "CalcDP_41000_FTO_ATC" + Chr(13) + Chr(10) _
    + Chr(13) + Chr(10) + "long_TrimPackageID=" + CStr(long_TrimPackageID) _
    + Chr(13) + Chr(10) + "long_ActuatorPartID=" + CStr(long_ActuatorPartID) + Chr(13) + Chr(10) + "dbl_InletPressure1=" + CStr(dbl_InletPressure1) _
    + Chr(13) + Chr(10) + "dbl_InletPressure2=" + CStr(dbl_InletPressure2) + Chr(13) + Chr(10) + "dbl_InletPressure3=" + CStr(dbl_InletPressure3) _
    + Chr(13) + Chr(10) + "dbl_InletPressure4=" + CStr(dbl_InletPressure4) + Chr(13) + Chr(10) + "dbl_OutletPressure1=" + CStr(dbl_OutletPressure1) + Chr(13) + Chr(10) + "dbl_OutletPressure2=" + CStr(dbl_OutletPressure2) _
    + Chr(13) + Chr(10) + "dbl_OutletPressure3=" + CStr(dbl_OutletPressure3) + Chr(13) + Chr(10) + "dbl_OutletPressure4=" + CStr(dbl_OutletPressure4) + Chr(13) + Chr(10) + "dbl_InletTemperature1=" + CStr(dbl_InletTemperature1) _
    + Chr(13) + Chr(10) + "dbl_InletTemperature2=" + CStr(dbl_InletTemperature2) + Chr(13) + Chr(10) + "dbl_InletTemperature3=" + CStr(dbl_InletTemperature3) + Chr(13) + Chr(10) + "dbl_InletTemperature4=" + CStr(dbl_InletTemperature4) + Chr(13) + Chr(10) + "dbl_ShutoffDP=" + CStr(dbl_ShutOffDP) _
    + Chr(13) + Chr(10) + "str_SpecifiedPressureUM=" + str_SpecifiedPressureUM + Chr(13) + Chr(10) + "str_SpecifiedTemperatureUM=" + str_SpecifiedTemperatureUM _
    + Chr(13) + Chr(10) + "str_SpecifiedCalibPressureUM=" + str_SpecifiedCalibPressureUM _
    + Chr(13) + Chr(10) + "str_SpecifiedLengthUM=" + str_SpecifiedLengthUM + Chr(13) + Chr(10) + "str_BodyMatl=" + str_BodyMatl _
    + Chr(13) + Chr(10) + "str_SelectedRating=" + str_SelectedRating + Chr(13) + Chr(10) + "dbl_ValveSize=" + CStr(dbl_ValveSize) _
    + Chr(13) + Chr(10) + "str_PackingMaterial=" + str_PackingMaterial + Chr(13) + Chr(10) + "str_LeakageClass=" + str_LeakageClass _
    + Chr(13) + Chr(10) + "str_SpringRange=" + str_SpringRange + Chr(13) + Chr(10) + "dbl_RatedCv=" + CStr(dbl_RatedCv) _
    + Chr(13) + Chr(10) + "dbl_ShowDP=" + CStr(dbl_ShowDP) + Chr(13) + Chr(10) + "dbl_ShowDPSupplyPressure=" + CStr(dbl_ShowDPSupplyPressure) _
    + Chr(13) + Chr(10) + "dbl_ShowMaxDP=" + CStr(dbl_ShowMaxDP) + Chr(13) + Chr(10) + "dbl_ShowMaxDPSupplyPressure=" + CStr(dbl_ShowMaxDPSupplyPressure) _
    + Chr(13) + Chr(10) + "str_Message=" + str_Message
#End If

'***************************
'Parameter validation
'***************************
If dbl_ShutOffDP < 0 Then GoTo CalcDP_41000_FTO_ATC_Error
If long_TrimPackageID < 0 Then GoTo CalcDP_41000_FTO_ATC_Error
If long_ActuatorPartID < 0 Then GoTo CalcDP_41000_FTO_ATC_Error
If dbl_RatedCv < 0 Then GoTo CalcDP_41000_FTO_ATC_Error

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4

'Determine requirements
Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature(), dbl_MaxInletTemperature)
dbl_ShowDP = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, _
    str_SpecifiedPressureUM)
dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, _
    str_SpecifiedTemperatureUM)
dbl_ShowShutoffDP = dbl_ShutOffDP
dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, _
    str_SpecifiedPressureUM)

'Determine ANSI requirements
Call GetANSIPressure(str_BodyMatl, dbl_MaxInletTemperature, _
    str_SelectedRating, dbl_AllowANSIPressure)

'Get product data
Call Get41000Data(long_TrimPackageID, dbl_RatedCv, dbl_StemDiameter, _
    dbl_SeatDiameter, dbl_Kv, dbl_Au, dbl_PackingHeight, _
    dbl_StemStressArea, str_StemMaterial, dbl_ForcePilotSpring, dbl_Travel, _
    dbl_PlugSealDiameter, dbl_PlugThrottlingDiameter, dbl_DynPlugDiameter, _
    dbl_SealRingCoef, dbl_TECSealCoef, dbl_TECCoefJ, dbl_TECCoefS, _
    dbl_CageInsideDiameter, dbl_SealRingWidth, dbl_ForcePlugWt, dbl_MaxAllowFlatSpringDP, _
    dbl_SeatingDiameter, dbl_PilotTravel)

Call GetActuatorData(long_ActuatorPartID, dbl_AreaInitial, dbl_AreaFinal, _
    dbl_MaxSupplyPressure, dbl_ActuatorFriction, dbl_PositionerLossCoef, _
    str_Model, long_AirActionID, _
    int_QtySeals, dbl_StemSealWidth, dbl_PistonRingWidth, _
        dbl_VolOver, dbl_ForceSpringClosed, dbl_ActuatorStemDiameter, dbl_LeverLength)

Call GetTrimTypeFromPackage(long_TrimPackageID, long_TrimTypeID)

'Calculate required actuator evaluation data
Call CalcSeatingForce(str_LeakageClass, dbl_SeatDiameter, _
    dbl_ForceSeating)
Call CalcPackingForce(dbl_MaxInletPressure, dbl_StemDiameter, _
    str_PackingMaterial, dbl_PackingHeight, dbl_ForcePacking)
Call CalcActuatorSealFrictionForce(dbl_AreaInitial, _
    dbl_ActuatorStemDiameter, dbl_PistonRingWidth, _
    dbl_StemSealWidth, int_QtySeals, dbl_MaxSupplyPressure, _
    dbl_ForceActSeal)
Call CalcPositionerLossForce(dbl_MaxSupplyPressure, _
    dbl_AreaInitial, dbl_ForcePositionerLoss)
Call CalcPlugSealRingFrictionForce(dbl_SealRingCoef, _
    dbl_TECSealCoef, dbl_TECCoefJ, dbl_TECCoefS, _
    dbl_CageInsideDiameter, dbl_SealRingWidth, _
    dbl_ForceSealRingArea, dbl_ForceSealRingSpring)

'Calculate actuator stiffness
Call CalcDynamicFrictionRateChange(dbl_MaxInletPressure, _
    dbl_ForceSealRingArea, dbl_ForceSealRingSpring, _
    dbl_ForcePacking, dbl_Travel, dbl_DynFrictionRateChange)
Call CalcSpringRate(str_SpringRange, dbl_AreaInitial, _
    dbl_PInitial, dbl_AreaFinal, dbl_PFinal, str_SpecifiedCalibPressureUM, _
    dbl_Travel, str_SpecifiedLengthUM, dbl_SpringRate)

'Evaluate max inlet pressure
Call CalcReqActForce_41000_FTO(dbl_MaxInletPressure, _
    dbl_SeatDiameter, dbl_CageInsideDiameter, dbl_StemDiameter, _
    dbl_ForceSeating, dbl_ForcePacking, dbl_ForceActSeal, _
    dbl_ForcePositionerLoss, dbl_ForceSealRingArea, _
    dbl_ForceSealRingSpring, dbl_ForcePlugWt, dbl_RequiredActuatorForce)
Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, _
    dbl_RequiredActuatorForce, dbl_PSupplyDP)

'Evaluate shutoff DP specified
Call CalcReqActForce_41000_FTO(dbl_ShutOffDP, _
    dbl_SeatDiameter, dbl_CageInsideDiameter, dbl_StemDiameter, _
    dbl_ForceSeating, dbl_ForcePacking, dbl_ForceActSeal, _
    dbl_ForcePositionerLoss, dbl_ForceSealRingArea, _
    dbl_ForceSealRingSpring, dbl_ForcePlugWt, dbl_RequiredShutoffDPActuatorForce)
Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, _
    dbl_RequiredShutoffDPActuatorForce, dbl_PSupplyDPShutoff)


'Calculate stem limitations
dbl_L = 6# 'default value
Call GetStemStressArea(dbl_StemDiameter, dbl_StemStressArea)
Call GetStemProperties(str_StemMaterial, dbl_MaxInletTemperature, _
    dbl_Sy, dbl_E, dbl_Sd)
Call CalcAllowStemForce(dbl_Sy, dbl_E, dbl_StemDiameter, _
    dbl_L, dbl_Sd, dbl_StemStressArea, dbl_AllowableStemForce)

'Evaluate mechanical limit
Call CalcAllowDP_41000_FTO(dbl_MaxInletPressure, dbl_AllowableStemForce, dbl_PlugSealDiameter, _
    dbl_CageInsideDiameter, dbl_StemDiameter, dbl_ForceSeating, _
    dbl_ForcePacking, dbl_ForceActSeal, dbl_ForcePositionerLoss, _
    dbl_ForceSealRingArea, dbl_ForceSealRingSpring, _
    dbl_ForcePlugWt, dbl_MaxAllowDPStem)

'?Call CalcMaxDPcage_41000(dbl_ValveSize, long_TrimTypeID, _
    dbl_MaxInletTemperature, dbl_MaxAllowCageDP)

'Evaluation
str_Message = ""

'dbl_MaxAllowTrimDP = dbl_MaxAllowCageDP

'If (dbl_MaxAllowFlatSpringDP < dbl_MaxAllowTrimDP And _
'    dbl_MaxAllowFlatSpring <> 0) Then
'    dbl_MaxAllowTrimDP = dbl_MaxAllowCageDP
'End If

'If (dbl_MaxInletPressure > dbl_MaxAllowTrimDP) Then
'    str_Message = TRIMMESSAGE
'    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowTrimDP, _
'        str_SpecifiedPressureUM)
'End If

'determine limit
dbl_MaxAllowDP = dbl_AllowANSIPressure
'If (dbl_MaxAllowTrimDP < dbl_MaxAllowDP) Then dbl_MaxAllowDP = dbl_MaxAllowTrimDP
If (dbl_MaxAllowDPStem < dbl_MaxAllowDP) Then dbl_MaxAllowDP = dbl_MaxAllowDPStem
Call CalcReqActForce_41000_FTO(dbl_MaxAllowDP, _
    dbl_SeatDiameter, dbl_CageInsideDiameter, dbl_StemDiameter, _
    dbl_ForceSeating, dbl_ForcePacking, dbl_ForceActSeal, _
    dbl_ForcePositionerLoss, dbl_ForceSealRingArea, _
    dbl_ForceSealRingSpring, dbl_ForcePlugWt, dbl_RequiredActuatorForceLimit)
Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, _
    dbl_RequiredActuatorForceLimit, dbl_PSupplyDP)
dbl_ShowDPSupplyPressure = 0#
dbl_ShowMaxDPSupplyPressure = 0#

If (dbl_DynFrictionRateChange > dbl_SpringRate) Then
    str_Message = ACTUATORMESSAGE
    GoTo Leave
End If

If (dbl_MaxInletPressure > dbl_AllowANSIPressure) Then
    str_Message = ANSIMESSAGE
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowANSIPressure, str_SpecifiedPressureUM)
    GoTo Leave
End If

If (dbl_MaxInletPressure < dbl_MaxAllowTrimDP) Then
    str_Message = TRIMMESSAGE
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowTrimDP, str_SpecifiedPressureUM)
End If

If (dbl_MaxInletPressure > dbl_MaxAllowDPStem) Then
    str_Message = STEMMESSAGE
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowDPStem, _
        str_SpecifiedPressureUM)
    
End If

If (dbl_MaxInletPressure <= dbl_MaxAllowDP) Then
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG( _
        dbl_PSupplyDP, str_SpecifiedPressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG( _
        dbl_MaxAllowDP, str_SpecifiedPressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG( _
        dbl_PSupplyDP, str_SpecifiedCalibPressureUM)
ElseIf (dbl_MaxInletPressure > dbl_MaxAllowDP And _
    dbl_ShutOffDP <= dbl_MaxAllowDP And dbl_ShutOffDP > 0#) Then
    str_Message = LIMITEDDPMESSAGE
    dbl_ShowDP = dbl_ShowShutoffDP
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG( _
        dbl_PSupplyDP, str_SpecifiedCalibPressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, _
        str_SpecifiedPressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG( _
        dbl_PSupplyDP, str_SpecifiedCalibPressureUM)
Else
    str_Message = ACTUATORMESSAGE
End If

Leave:
dbl_ShowDP = TrimDecimalPlaces(dbl_ShowDP)
dbl_ShowDPSupplyPressure = TrimDecimalPlaces(dbl_ShowDPSupplyPressure)
dbl_ShowMaxDP = TrimDecimalPlaces(dbl_ShowMaxDP)
dbl_ShowMaxDPSupplyPressure = TrimDecimalPlaces(dbl_ShowMaxDPSupplyPressure)

#If (DebugVersion = "yes") Then
MsgBox "CalcDP_41000_FTO_ATC" + Chr(13) + Chr(10) _
    + Chr(13) + Chr(10) + "long_TrimPackageID=" + CStr(long_TrimPackageID) _
    + Chr(13) + Chr(10) + "long_ActuatorPartID=" + CStr(long_ActuatorPartID) + Chr(13) + Chr(10) + "dbl_InletPressure1=" + CStr(dbl_InletPressure1) _
    + Chr(13) + Chr(10) + "dbl_InletPressure2=" + CStr(dbl_InletPressure2) + Chr(13) + Chr(10) + "dbl_InletPressure3=" + CStr(dbl_InletPressure3) _
    + Chr(13) + Chr(10) + "dbl_InletPressure4=" + CStr(dbl_InletPressure4) + Chr(13) + Chr(10) + "dbl_OutletPressure1=" + CStr(dbl_OutletPressure1) + Chr(13) + Chr(10) + "dbl_OutletPressure2=" + CStr(dbl_OutletPressure2) _
    + Chr(13) + Chr(10) + "dbl_OutletPressure3=" + CStr(dbl_OutletPressure3) + Chr(13) + Chr(10) + "dbl_OutletPressure4=" + CStr(dbl_OutletPressure4) + Chr(13) + Chr(10) + "dbl_InletTemperature1=" + CStr(dbl_InletTemperature1) _
    + Chr(13) + Chr(10) + "dbl_InletTemperature2=" + CStr(dbl_InletTemperature2) + Chr(13) + Chr(10) + "dbl_InletTemperature3=" + CStr(dbl_InletTemperature3) + Chr(13) + Chr(10) + "dbl_InletTemperature4=" + CStr(dbl_InletTemperature4) + Chr(13) + Chr(10) + "dbl_ShutoffDP=" + CStr(dbl_ShutOffDP) _
    + Chr(13) + Chr(10) + "str_SpecifiedPressureUM=" + str_SpecifiedPressureUM + Chr(13) + Chr(10) + "str_SpecifiedTemperatureUM=" + str_SpecifiedTemperatureUM _
    + Chr(13) + Chr(10) + "str_SpecifiedCalibPressureUM=" + str_SpecifiedCalibPressureUM _
    + Chr(13) + Chr(10) + "str_SpecifiedLengthUM=" + str_SpecifiedLengthUM + Chr(13) + Chr(10) + "str_BodyMatl=" + str_BodyMatl _
    + Chr(13) + Chr(10) + "str_SelectedRating=" + str_SelectedRating + Chr(13) + Chr(10) + "dbl_ValveSize=" + CStr(dbl_ValveSize) _
    + Chr(13) + Chr(10) + "str_PackingMaterial=" + str_PackingMaterial + Chr(13) + Chr(10) + "str_LeakageClass=" + str_LeakageClass _
    + Chr(13) + Chr(10) + "str_SpringRange=" + str_SpringRange + Chr(13) + Chr(10) + "dbl_RatedCv=" + CStr(dbl_RatedCv) _
    + Chr(13) + Chr(10) + "dbl_ShowDP=" + CStr(dbl_ShowDP) + Chr(13) + Chr(10) + "dbl_ShowDPSupplyPressure=" + CStr(dbl_ShowDPSupplyPressure) _
    + Chr(13) + Chr(10) + "dbl_ShowMaxDP=" + CStr(dbl_ShowMaxDP) + Chr(13) + Chr(10) + "dbl_ShowMaxDPSupplyPressure=" + CStr(dbl_ShowMaxDPSupplyPressure) _
    + Chr(13) + Chr(10) + "str_Message=" + str_Message
#End If
    
Exit Sub
CalcDP_41000_FTO_ATC_Error:

End Sub


Public Sub CalcDP_41000_FTO_ATO(long_TrimPackageID As Long, _
    long_ActuatorPartID As Long, dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, dbl_InletPressure3 As Double, _
    dbl_InletPressure4 As Double, dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, dbl_InletTemperature1 As Double, _
    dbl_InletTemperature2 As Double, dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, dbl_ShutOffDP _
    As Double, str_SpecifiedPressureUM As String, str_SpecifiedTemperatureUM _
    As String, str_SpecifiedCalibPressureUM As String, _
    str_SpecifiedLengthUM As String, str_BodyMatl As String, _
    str_SelectedRating As String, dbl_ValveSize As Double, _
    str_PackingMaterial As String, str_LeakageClass As String, _
    str_SpringRange As String, dbl_RatedCv As Double, _
    dbl_ShowDP As Double, dbl_ShowDPSupplyPressure As Double, _
    dbl_ShowMaxDP As Double, dbl_ShowMaxDPSupplyPressure As Double, _
    str_Message As String)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'calcDP_41000_FTO_ATO(----see input output list below-)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Reference: CES 1019 Balanced and pilot balanced cage throttling
'   globe valves - Actuator Sizing Methods.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This subroutine is used to calculate the selected valve and
'  actuator shutoff capability and compare the result to defined
'  requirements.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This subroutine is specific to 41000 series valves, flow to open,
'  air to open.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Input Variables
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff
'                               pressure drop
'str_SpecifiedPressureUM        unit of measure associated with inlet
'                               and outlet pressures
'str_SpecifiedTemperatureUM     unit of measure associated with inlet
'                               temperature
'str_SpecifledCalibPressureUM   unit of measure assoeicated with spring
'                               range and supply pressure
'str_SpecifiedLengthUM          units of measure for length
'dbl_BodyMatl                   selected material of body
'str_SelectedRating             selected body rating
'long_TrimPackageID             ID of trim selected
'long_ActuatorID                ID of actuator selected
'dbl_ValveSize                  valve size
'str_PackingMaterial            selected packing material
'str_LeakageClass               selected leakage class
'str_SpringRange                selected actuator spring range
'dbl_RatedCv                    selected rated Cv
'
'Output Variables:
'dbl_ShowDP                     specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure       calculated required supply pressure
'                               for specified
'dbl_ShowMaxDP                  calculated maximum pressure drop for
'                               selected valve and actuator combination
'dbl_ShowMaxDpsupplypressure    calculated required supply pressure
'                               for maximum pressure drop
'str_Message                    warning message if pressure drop
'                               requirement exceeds equipment capability
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'coded: 10/15/98    Frank Krieger
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo CalcDP_41000_FTO_ATO_Error

'***************************
'Declarations
'***************************

Dim dbl_ActuatorForce As Double
Dim dbl_ActuatorFriction As Double
Dim dbl_ActuatorStemDiameter As Double
Dim dbl_AllowableStemForce As Double
Dim dbl_AllowANSIPressure As Double
Dim dbl_AllowDP As Double
Dim dbl_AllowDPStem As Double
Dim dbl_AllowStemLoad As Double
Dim dbl_AreaFinal As Double
Dim dbl_AreaInitial As Double
Dim dbl_As As Double
Dim dbl_Au As Double
Dim dbl_CageInsideDiameter As Double
Dim dbl_DynFrictionRateChange As Double
Dim dbl_DynPlugDiameter As Double
Dim dbl_E As Double
Dim dbl_Fcs As Double
Dim dbl_ForceActSeal As Double
Dim dbl_ForceActuator As Double
Dim dbl_ForcePacking As Double
Dim dbl_ForcePilotSpring As Double
Dim dbl_ForcePlugWt As Double
Dim dbl_ForcePositionerLoss As Double
Dim dbl_ForceSealRingArea As Double
Dim dbl_ForceSealRingSpring As Double
Dim dbl_ForceSeating As Double
Dim dbl_ForceSpringClosed As Double
Dim dbl_InletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double
Dim dbl_Kv As Double
Dim dbl_L As Double
Dim dbl_LeverLength As Double
Dim dbl_MaxAllowCageDP As Double
Dim dbl_MaxAllowDP As Double
Dim dbl_MaxAllowDPStem As Double
Dim dbl_MaxAllowFlatSpring As Double
Dim dbl_MaxAllowFlatSpringDP As Double
Dim dbl_MaxAllowTrimDP As Double
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_MaxSupplyPressure As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_PackingHeight As Double
Dim dbl_PFinal As Double
Dim dbl_PilotTravel As Double
Dim dbl_PInitial As Double
Dim dbl_PistonRingWidth As Double
Dim dbl_PlugSealDiameter As Double
Dim dbl_PlugThrottlingDiameter As Double
Dim dbl_PositionerLossCoef As Double
Dim dbl_PSupplyDP As Double
Dim dbl_PSupplyDPANSI As Double
Dim dbl_PSupplyDPShutoff As Double
Dim dbl_PSupplyDPStem As Double
Dim dbl_RequiredActuatorForce As Double
Dim dbl_RequiredANSIActuatorForce As Double
Dim dbl_RequiredShutoffDPActuatorForce As Double
Dim dbl_RequiredSupplyPressure As Double
Dim dbl_Sd As Double
Dim dbl_SealRingCoef As Double
Dim dbl_SealRingWidth As Double
Dim dbl_SeatDiameter As Double
Dim dbl_SeatingDiameter As Double
Dim dbl_ShowMaxDPSupply As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_SpringEnergizedSealSpringCoef As Double
Dim dbl_SpringRate As Double
Dim dbl_StemDiameter As Double
Dim dbl_StemSealWidth As Double
Dim dbl_StemStressArea As Double
Dim dbl_Sy As Double
Dim dbl_TECCoefJ As Double
Dim dbl_TECCoefS As Double
Dim dbl_TECSealCoef As Double
Dim dbl_Travel As Double
Dim dbl_VolOver As Double
Dim int_QtySeals As Integer
Dim long_AirActionID As Long
Dim long_CageTypeID As Long
Dim long_PlugTypeID As Long
Dim long_TrimTypeID As Long
Dim str_Model As String
Dim str_StemMaterial As String

#If (DebugVersion = "yes") Then
MsgBox "CalcDP_41000_FTO_ATO" + Chr(13) + Chr(10) _
    + Chr(13) + Chr(10) + "long_TrimPackageID=" + CStr(long_TrimPackageID) _
    + Chr(13) + Chr(10) + "long_ActuatorPartID=" + CStr(long_ActuatorPartID) + Chr(13) + Chr(10) + "dbl_InletPressure1=" + CStr(dbl_InletPressure1) _
    + Chr(13) + Chr(10) + "dbl_InletPressure2=" + CStr(dbl_InletPressure2) + Chr(13) + Chr(10) + "dbl_InletPressure3=" + CStr(dbl_InletPressure3) _
    + Chr(13) + Chr(10) + "dbl_InletPressure4=" + CStr(dbl_InletPressure4) + Chr(13) + Chr(10) + "dbl_OutletPressure1=" + CStr(dbl_OutletPressure1) + Chr(13) + Chr(10) + "dbl_OutletPressure2=" + CStr(dbl_OutletPressure2) _
    + Chr(13) + Chr(10) + "dbl_OutletPressure3=" + CStr(dbl_OutletPressure3) + Chr(13) + Chr(10) + "dbl_OutletPressure4=" + CStr(dbl_OutletPressure4) + Chr(13) + Chr(10) + "dbl_InletTemperature1=" + CStr(dbl_InletTemperature1) _
    + Chr(13) + Chr(10) + "dbl_InletTemperature2=" + CStr(dbl_InletTemperature2) + Chr(13) + Chr(10) + "dbl_InletTemperature3=" + CStr(dbl_InletTemperature3) + Chr(13) + Chr(10) + "dbl_InletTemperature4=" + CStr(dbl_InletTemperature4) + Chr(13) + Chr(10) + "dbl_ShutoffDP=" + CStr(dbl_ShutOffDP) _
    + Chr(13) + Chr(10) + "str_SpecifiedPressureUM=" + str_SpecifiedPressureUM + Chr(13) + Chr(10) + "str_SpecifiedTemperatureUM=" + str_SpecifiedTemperatureUM _
    + Chr(13) + Chr(10) + "str_SpecifiedCalibPressureUM=" + str_SpecifiedCalibPressureUM _
    + Chr(13) + Chr(10) + "str_SpecifiedLengthUM=" + str_SpecifiedLengthUM + Chr(13) + Chr(10) + "str_BodyMatl=" + str_BodyMatl _
    + Chr(13) + Chr(10) + "str_SelectedRating=" + str_SelectedRating + Chr(13) + Chr(10) + "dbl_ValveSize=" + CStr(dbl_ValveSize) _
    + Chr(13) + Chr(10) + "str_PackingMaterial=" + str_PackingMaterial + Chr(13) + Chr(10) + "str_LeakageClass=" + str_LeakageClass _
    + Chr(13) + Chr(10) + "str_SpringRange=" + str_SpringRange + Chr(13) + Chr(10) + "dbl_RatedCv=" + CStr(dbl_RatedCv) _
    + Chr(13) + Chr(10) + "dbl_ShowDP=" + CStr(dbl_ShowDP) + Chr(13) + Chr(10) + "dbl_ShowDPSupplyPressure=" + CStr(dbl_ShowDPSupplyPressure) _
    + Chr(13) + Chr(10) + "dbl_ShowMaxDP=" + CStr(dbl_ShowMaxDP) + Chr(13) + Chr(10) + "dbl_ShowMaxDPSupplyPressure=" + CStr(dbl_ShowMaxDPSupplyPressure) _
    + Chr(13) + Chr(10) + "str_Message=" + str_Message
#End If

'***************************
'Parameter validation
'***************************
If dbl_ShutOffDP < 0 Then GoTo CalcDP_41000_FTO_ATO_Error
If long_TrimPackageID < 0 Then GoTo CalcDP_41000_FTO_ATO_Error
If long_ActuatorPartID < 0 Then GoTo CalcDP_41000_FTO_ATO_Error
If dbl_RatedCv < 0 Then GoTo CalcDP_41000_FTO_ATO_Error

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4

'Determine requirements
Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature(), dbl_MaxInletTemperature)
dbl_ShowDP = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, _
    str_SpecifiedPressureUM)
dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, _
    str_SpecifiedTemperatureUM)
dbl_ShowShutoffDP = dbl_ShutOffDP
dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, _
    str_SpecifiedPressureUM)

'Determine ANSI requirements
Call GetANSIPressure(str_BodyMatl, dbl_MaxInletTemperature, _
    str_SelectedRating, dbl_AllowANSIPressure)

'Get product data
    Call Get41000Data(long_TrimPackageID, dbl_RatedCv, dbl_StemDiameter, _
    dbl_SeatDiameter, dbl_Kv, dbl_Au, dbl_PackingHeight, _
    dbl_StemStressArea, str_StemMaterial, dbl_ForcePilotSpring, dbl_Travel, _
    dbl_PlugSealDiameter, dbl_PlugThrottlingDiameter, dbl_DynPlugDiameter, _
    dbl_SealRingCoef, dbl_TECSealCoef, dbl_TECCoefJ, dbl_TECCoefS, _
    dbl_CageInsideDiameter, dbl_SealRingWidth, dbl_ForcePlugWt, dbl_MaxAllowFlatSpringDP, _
    dbl_SeatingDiameter, dbl_PilotTravel)

    Call GetActuatorData(long_ActuatorPartID, dbl_AreaInitial, dbl_AreaFinal, _
        dbl_MaxSupplyPressure, dbl_ActuatorFriction, dbl_PositionerLossCoef, _
        str_Model, long_AirActionID, _
        int_QtySeals, dbl_StemSealWidth, dbl_PistonRingWidth, _
        dbl_VolOver, dbl_ForceSpringClosed, dbl_ActuatorStemDiameter, dbl_LeverLength)
    
    Call GetTrimTypeFromPackage(long_TrimPackageID, long_TrimTypeID)

'Calculate required actuator evaluation data
Call CalcSeatingForce(str_LeakageClass, dbl_SeatDiameter, _
    dbl_ForceSeating)
Call CalcPackingForce(dbl_MaxInletPressure, dbl_StemDiameter, _
    str_PackingMaterial, dbl_PackingHeight, dbl_ForcePacking)
Call CalcActuatorSealFrictionForce(dbl_AreaInitial, _
    dbl_ActuatorStemDiameter, dbl_PistonRingWidth, _
    dbl_StemSealWidth, int_QtySeals, dbl_MaxSupplyPressure, _
    dbl_ForceActSeal)
Call CalcPositionerLossForce(dbl_MaxSupplyPressure, dbl_AreaInitial, _
    dbl_ForcePositionerLoss)
Call CalcPlugSealRingFrictionForce(dbl_SealRingCoef, _
    dbl_TECSealCoef, dbl_TECCoefJ, dbl_TECCoefS, _
    dbl_CageInsideDiameter, dbl_SealRingWidth, _
    dbl_ForceSealRingArea, dbl_ForceSealRingSpring)

'Calculate actuator stiffness
Call CalcDynamicFrictionRateChange(dbl_MaxInletPressure, _
    dbl_ForceSealRingArea, dbl_ForceSealRingSpring, _
    dbl_ForcePacking, dbl_Travel, dbl_DynFrictionRateChange)
Call CalcSpringRate(str_SpringRange, dbl_AreaInitial, _
    dbl_PInitial, dbl_AreaFinal, dbl_PFinal, str_SpecifiedCalibPressureUM, _
    dbl_Travel, str_SpecifiedLengthUM, dbl_SpringRate)

'Calculate available actuator force
Call CalcATOActuatorForce(dbl_AreaInitial, dbl_PInitial, _
    dbl_PFinal, dbl_ActuatorFriction, dbl_RequiredSupplyPressure, _
    dbl_ForceActuator)

'Calculate max allowable DP
    'replaced dbl_SeatDiameter with plugsealdiameter - schoonover 12/13/98
Call CalcAllowDP_41000_FTO(dbl_MaxInletPressure, _
    dbl_ForceActuator, dbl_PlugSealDiameter, _
    dbl_CageInsideDiameter, dbl_StemDiameter, dbl_ForceSeating, _
    dbl_ForcePacking, dbl_ForceActSeal, dbl_ForcePositionerLoss, _
    dbl_ForceSealRingArea, dbl_ForceSealRingSpring, _
    dbl_ForcePlugWt, dbl_MaxAllowDP)
    
'Evaluate mechanical limit
Call GetStemStressArea(dbl_StemDiameter, dbl_StemStressArea)
dbl_L = 6

Call GetStemProperties(str_StemMaterial, dbl_MaxInletTemperature, _
    dbl_Sy, dbl_E, dbl_Sd)
Call CalcAllowStemForce(dbl_Sy, dbl_E, dbl_StemDiameter, dbl_L, _
    dbl_Sd, dbl_StemStressArea, dbl_AllowableStemForce)
Call CalcAllowDP_41000_FTO(dbl_MaxInletPressure, _
    dbl_AllowableStemForce, dbl_PlugSealDiameter, _
    dbl_CageInsideDiameter, dbl_StemDiameter, dbl_ForceSeating, _
    dbl_ForcePacking, dbl_ForceActSeal, dbl_ForcePositionerLoss, _
    dbl_ForceSealRingArea, dbl_ForceSealRingSpring, _
    dbl_ForcePlugWt, dbl_MaxAllowDPStem)

If dbl_MaxAllowDPStem < dbl_MaxAllowDP Then dbl_MaxAllowDP = dbl_MaxAllowDPStem

Call CalcMaxDPCage_41000(dbl_ValveSize, long_TrimTypeID, _
    dbl_MaxInletTemperature, dbl_MaxAllowCageDP)

Call CalcRequiredSupplyPressure(dbl_AreaFinal, dbl_PFinal, _
    dbl_MaxInletPressure, dbl_PSupplyDP)

'Evaluation
str_Message = ""
'str_ANSIMessage = "SPECIFIED PRESSURE EXCEEDS ANSI LIMIT"
'str_StemMessage = "ACTUATOR OUTPUT EXCEEDS STEM LIMIT"
'str_ActuatorMessage = "ACTUATOR CAPABILITY EXCEEDED (CLOSED)"
'str_TrimMessage = "SPECIFIED PRESSURE EXCEEDS TRIM LIMIT"
'str_LimitedDPMessage = "WARNING - MAX INLET PRESSURE EXCEEDS ACTUATOR SHUTOFF LIMIT"

dbl_ShowDPSupplyPressure = 0#
dbl_ShowMaxDPSupplyPressure = 0#

If (dbl_DynFrictionRateChange > dbl_SpringRate) Then
    str_Message = ACTUATORMESSAGE
    GoTo Leave
End If

If (dbl_MaxInletPressure > dbl_AllowANSIPressure) Then
    str_Message = ANSIMESSAGE
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowANSIPressure, str_SpecifiedPressureUM)
End If

dbl_MaxAllowTrimDP = dbl_MaxAllowCageDP

If (dbl_MaxAllowFlatSpringDP < dbl_MaxAllowTrimDP And _
    dbl_MaxAllowFlatSpring <> 0) Then
    dbl_MaxAllowTrimDP = dbl_MaxAllowCageDP
End If

If (dbl_MaxInletPressure > dbl_MaxAllowTrimDP) Then
    str_Message = TRIMMESSAGE
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowTrimDP, _
        str_SpecifiedPressureUM)
End If

If (dbl_ForceActuator > dbl_AllowableStemForce) Then
    'stemOK = False
    str_Message = STEMMESSAGE
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDPStem, str_SpecifiedPressureUM)
End If


If (dbl_MaxInletPressure <= dbl_MaxAllowDP) Then
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG( _
        dbl_PSupplyDP, str_SpecifiedPressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG( _
        dbl_MaxAllowDP, str_SpecifiedPressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG( _
        dbl_PSupplyDP, str_SpecifiedCalibPressureUM)

ElseIf (dbl_MaxInletPressure > dbl_MaxAllowDP And _
    dbl_ShutOffDP <= dbl_MaxAllowDP And dbl_ShutOffDP > 0#) Then
    str_Message = LIMITEDDPMESSAGE
    dbl_ShowDP = dbl_ShowShutoffDP
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG( _
        dbl_PSupplyDPShutoff, str_SpecifiedCalibPressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, _
        str_SpecifiedPressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG( _
        dbl_PSupplyDP, str_SpecifiedCalibPressureUM)
Else
    str_Message = ACTUATORMESSAGE
End If

dbl_ShowDP = TrimDecimalPlaces(dbl_ShowDP)
dbl_ShowDPSupplyPressure = TrimDecimalPlaces(dbl_ShowDPSupplyPressure)
dbl_ShowMaxDP = TrimDecimalPlaces(dbl_ShowMaxDP)
dbl_ShowMaxDPSupplyPressure = TrimDecimalPlaces(dbl_ShowMaxDPSupplyPressure)

Leave:
#If (DebugVersion = "yes") Then
MsgBox "CalcDP_41000_FTO_ATO" + Chr(13) + Chr(10) _
    + Chr(13) + Chr(10) + "long_TrimPackageID=" + CStr(long_TrimPackageID) _
    + Chr(13) + Chr(10) + "long_ActuatorPartID=" + CStr(long_ActuatorPartID) + Chr(13) + Chr(10) + "dbl_InletPressure1=" + CStr(dbl_InletPressure1) _
    + Chr(13) + Chr(10) + "dbl_InletPressure2=" + CStr(dbl_InletPressure2) + Chr(13) + Chr(10) + "dbl_InletPressure3=" + CStr(dbl_InletPressure3) _
    + Chr(13) + Chr(10) + "dbl_InletPressure4=" + CStr(dbl_InletPressure4) + Chr(13) + Chr(10) + "dbl_OutletPressure1=" + CStr(dbl_OutletPressure1) + Chr(13) + Chr(10) + "dbl_OutletPressure2=" + CStr(dbl_OutletPressure2) _
    + Chr(13) + Chr(10) + "dbl_OutletPressure3=" + CStr(dbl_OutletPressure3) + Chr(13) + Chr(10) + "dbl_OutletPressure4=" + CStr(dbl_OutletPressure4) + Chr(13) + Chr(10) + "dbl_InletTemperature1=" + CStr(dbl_InletTemperature1) _
    + Chr(13) + Chr(10) + "dbl_InletTemperature2=" + CStr(dbl_InletTemperature2) + Chr(13) + Chr(10) + "dbl_InletTemperature3=" + CStr(dbl_InletTemperature3) + Chr(13) + Chr(10) + "dbl_InletTemperature4=" + CStr(dbl_InletTemperature4) + Chr(13) + Chr(10) + "dbl_ShutoffDP=" + CStr(dbl_ShutOffDP) _
    + Chr(13) + Chr(10) + "str_SpecifiedPressureUM=" + str_SpecifiedPressureUM + Chr(13) + Chr(10) + "str_SpecifiedTemperatureUM=" + str_SpecifiedTemperatureUM _
    + Chr(13) + Chr(10) + "str_SpecifiedCalibPressureUM=" + str_SpecifiedCalibPressureUM _
    + Chr(13) + Chr(10) + "str_SpecifiedLengthUM=" + str_SpecifiedLengthUM + Chr(13) + Chr(10) + "str_BodyMatl=" + str_BodyMatl _
    + Chr(13) + Chr(10) + "str_SelectedRating=" + str_SelectedRating + Chr(13) + Chr(10) + "dbl_ValveSize=" + CStr(dbl_ValveSize) _
    + Chr(13) + Chr(10) + "str_PackingMaterial=" + str_PackingMaterial + Chr(13) + Chr(10) + "str_LeakageClass=" + str_LeakageClass _
    + Chr(13) + Chr(10) + "str_SpringRange=" + str_SpringRange + Chr(13) + Chr(10) + "dbl_RatedCv=" + CStr(dbl_RatedCv) _
    + Chr(13) + Chr(10) + "dbl_ShowDP=" + CStr(dbl_ShowDP) + Chr(13) + Chr(10) + "dbl_ShowDPSupplyPressure=" + CStr(dbl_ShowDPSupplyPressure) _
    + Chr(13) + Chr(10) + "dbl_ShowMaxDP=" + CStr(dbl_ShowMaxDP) + Chr(13) + Chr(10) + "dbl_ShowMaxDPSupplyPressure=" + CStr(dbl_ShowMaxDPSupplyPressure) _
    + Chr(13) + Chr(10) + "str_Message=" + str_Message
#End If

Exit Sub
CalcDP_41000_FTO_ATO_Error:

End Sub


Public Sub CalcDP_41000_FTC_ATC(long_TrimPackageID As Long, _
    long_ActuatorPartID As Long, dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, dbl_InletPressure3 As Double, _
    dbl_InletPressure4 As Double, dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, dbl_InletTemperature1 As Double, _
    dbl_InletTemperature2 As Double, dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, dbl_ShutOffDP _
    As Double, str_SpecifiedPressureUM As String, str_SpecifiedTemperatureUM _
    As String, str_SpecifiedCalibPressureUM As String, _
    str_SpecifiedLengthUM As String, str_FluidType As String, str_BodyMatl As String, _
    str_SelectedRating As String, dbl_ValveSize As Double, _
    str_PackingMaterial As String, str_LeakageClass As String, _
    str_SpringRange As String, dbl_RatedCv As Double, _
    dbl_ShowDP As Double, dbl_ShowDPSupplyPressure As Double, _
    dbl_ShowMaxDP As Double, dbl_ShowMaxDPSupplyPressure As Double, _
    str_Message As String)
    
'----------------------------------------------------------------------------------------
'calcDP_41000_FTC_ATC(----see input output list below-)
'----------------------------------------------------------------------------------------
'Reference: CES 1019 Balanced and pilot balanced cage throttling
'   globe valves - Actuator Sizing Methods.
'----------------------------------------------------------------------------------------
'This subroutine is used to calculate the selected valve and
'  actuator shutoff capability and compare the result to defined
'  requirements.
'----------------------------------------------------------------------------------------
'This subroutine is specific to 41000 series valves with pilot,
'  flow to close, air to close.
'----------------------------------------------------------------------------------------
'Input Variables
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff
'                               pressure Drop
'str_SpecifiedPressureUM        unit of measure associated with inlet
'                               and outlet pressures
'str_SpecifiedTemperatureUM     unit of measure associated with inlet
'                               temperature
'str_SpecifledCalibPressureUM   unit of measure assoeicated with spring
'                               range and supply pressure
'str_SpecifiedLengthUM          units of measure for length
'str_BodyMatl                   selected material of body
'str_SelectedRating             selected body rating
'long_TrimPackageID             ID of trim selected
'long_ActuatorID                ID of actuator selected
'dbl_ValveSize                  valve size
'str_PackingMaterial            selected packing material
'str_LeakageClass               selected leakage class
'str_SpringRange                selected actuator spring range
'dbl_RatedCv                    selected rated Cv
'----------------------------------------------------------------------------------------
'Output Variables:
'dbl_ShowDP                     specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure       calculated required supply pressure
'                               for specified
'dbl_ShowMaxDP                  calculated maximum pressure drop for
'                               selected valve and actuator combination
'dbl_ShowMaxDpsupplypressure    calculated required supply pressure
'                               for maximum pressure drop
'str_Message                    warning message if pressure drop
'                               requirement exceeds equipment capability
'----------------------------------------------------------------------------------------
'coded: 10/15/98    Frank Krieger
'----------------------------------------------------------------------------------------
On Error GoTo CalcDP_41000_FTC_ATC_Error

'***************************
'Declarations
'***************************

Dim dbl_ActForceClosed As Double
Dim dbl_ActuatorForceOpen As Double
Dim dbl_ActuatorFriction As Double
Dim dbl_ActuatorStemDiameter As Double
Dim dbl_AllowableStemForce As Double
Dim dbl_AllowANSIPressure As Double
Dim dbl_AllowDP As Double
Dim dbl_AllowDPClosed As Double
Dim dbl_AllowDPOpen As Double
Dim dbl_AllowDPStability As Double
Dim dbl_AllowDPStem As Double
Dim dbl_AllowSeatForce As Double
Dim dbl_AllowSeatYieldStress As Double
Dim dbl_AreaFinal As Double
Dim dbl_AreaInitial As Double
Dim dbl_AreaUnbalancedClosed As Double
Dim dbl_AreaUnbalancedOpen As Double
Dim dbl_As As Double
Dim dbl_Au As Double
Dim dbl_B As Double
Dim dbl_CageInsideDiameter As Double
Dim dbl_DynFrictionRateChange As Double
Dim dbl_DynPlugDiameter As Double
Dim dbl_E As Double
Dim dbl_EffectiveSpringRate As Double
Dim dbl_Fcs As Double
Dim dbl_ForceActSeal As Double
Dim dbl_ForcePacking As Double
Dim dbl_ForcePilotSpring As Double
Dim dbl_ForcePlugWt As Double
Dim dbl_ForcePositionerLoss As Double
Dim dbl_ForceSealRingArea As Double
Dim dbl_ForceSealRingSpring As Double
Dim dbl_ForceSeating As Double
Dim dbl_ForceSpringClosed As Double
Dim dbl_InletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double
Dim dbl_Kv As Double
Dim dbl_L As Double
Dim dbl_LeverLength As Double
Dim dbl_MaxAllowCageDP As Double
Dim dbl_MaxAllowDP As Double
Dim dbl_MaxAllowFlatSpring As Double
Dim dbl_MaxAllowFlatSpringDP As Double
Dim dbl_MaxAllowTrimDP As Double
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_MaxSupplyPressure As Double
Dim dbl_OffBalanceForce As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_PackingHeight As Double
Dim dbl_PFinal As Double
Dim dbl_PilotTravel As Double
Dim dbl_PInitial As Double
Dim dbl_PistonRingWidth As Double
Dim dbl_PlugSealDiameter As Double
Dim dbl_PlugThrottlingArea As Double
Dim dbl_PlugThrottlingDiameter As Double
Dim dbl_PositionerLossCoef As Double
Dim dbl_PSupplyDP As Double
Dim dbl_PSupplyDPANSI As Double
Dim dbl_PSupplyDPShutoff As Double
Dim dbl_PSupplyDPStem As Double
Dim dbl_RequiredActuatorForce As Double
Dim dbl_RequiredANSIActuatorForce As Double
Dim dbl_RequiredSeatingForce As Double
Dim dbl_RequiredShutoffDPActuatorForce As Double
Dim dbl_Sd As Double
Dim dbl_SealRingCoef As Double
Dim dbl_SealRingWidth As Double
Dim dbl_SeatDiameter As Double
Dim dbl_SeatingDiameter As Double
Dim dbl_ShowMaxDPSupply As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_SpringEnergizedSealSpringCoef As Double
Dim dbl_SpringRate As Double
Dim dbl_StemDiameter As Double
Dim dbl_StemSealWidth As Double
Dim dbl_StemStressArea As Double
Dim dbl_Sy As Double
Dim dbl_TECCoefJ As Double
Dim dbl_TECCoefS As Double
Dim dbl_TECSealCoef As Double
Dim dbl_Travel As Double
Dim dbl_VolOver As Double
Dim int_QtySeals As Integer
Dim lng_TrimTypeID As Long
Dim long_AirActionID As Long
Dim long_CageTypeID As Long
Dim long_FluidTypeID As Long
Dim long_PlugTypeID As Long
Dim long_TrimTypeID As Long
Dim str_ActuatorModel As String
Dim str_Model As String
Dim str_StemMaterial As String

#If (DebugVersion = "yes") Then
MsgBox "CalcDP_41000_FTC_ATC" + Chr(13) + Chr(10) _
    + Chr(13) + Chr(10) + "long_TrimPackageID=" + CStr(long_TrimPackageID) _
    + Chr(13) + Chr(10) + "long_ActuatorPartID=" + CStr(long_ActuatorPartID) + Chr(13) + Chr(10) + "dbl_InletPressure1=" + CStr(dbl_InletPressure1) _
    + Chr(13) + Chr(10) + "dbl_InletPressure2=" + CStr(dbl_InletPressure2) + Chr(13) + Chr(10) + "dbl_InletPressure3=" + CStr(dbl_InletPressure3) _
    + Chr(13) + Chr(10) + "dbl_InletPressure4=" + CStr(dbl_InletPressure4) + Chr(13) + Chr(10) + "dbl_OutletPressure1=" + CStr(dbl_OutletPressure1) + Chr(13) + Chr(10) + "dbl_OutletPressure2=" + CStr(dbl_OutletPressure2) _
    + Chr(13) + Chr(10) + "dbl_OutletPressure3=" + CStr(dbl_OutletPressure3) + Chr(13) + Chr(10) + "dbl_OutletPressure4=" + CStr(dbl_OutletPressure4) + Chr(13) + Chr(10) + "dbl_InletTemperature1=" + CStr(dbl_InletTemperature1) _
    + Chr(13) + Chr(10) + "dbl_InletTemperature2=" + CStr(dbl_InletTemperature2) + Chr(13) + Chr(10) + "dbl_InletTemperature3=" + CStr(dbl_InletTemperature3) + Chr(13) + Chr(10) + "dbl_InletTemperature4=" + CStr(dbl_InletTemperature4) + Chr(13) + Chr(10) + "dbl_ShutoffDP=" + CStr(dbl_ShutOffDP) _
    + Chr(13) + Chr(10) + "str_SpecifiedPressureUM=" + str_SpecifiedPressureUM + Chr(13) + Chr(10) + "str_SpecifiedTemperatureUM=" + str_SpecifiedTemperatureUM _
    + Chr(13) + Chr(10) + "str_SpecifiedCalibPressureUM=" + str_SpecifiedCalibPressureUM _
    + Chr(13) + Chr(10) + "str_SpecifiedLengthUM=" + str_SpecifiedLengthUM + Chr(13) + Chr(10) + "str_BodyMatl=" + str_BodyMatl _
    + Chr(13) + Chr(10) + "str_SelectedRating=" + str_SelectedRating + Chr(13) + Chr(10) + "dbl_ValveSize=" + CStr(dbl_ValveSize) _
    + Chr(13) + Chr(10) + "str_PackingMaterial=" + str_PackingMaterial + Chr(13) + Chr(10) + "str_LeakageClass=" + str_LeakageClass _
    + Chr(13) + Chr(10) + "str_SpringRange=" + str_SpringRange + Chr(13) + Chr(10) + "dbl_RatedCv=" + CStr(dbl_RatedCv) _
    + Chr(13) + Chr(10) + "dbl_ShowDP=" + CStr(dbl_ShowDP) + Chr(13) + Chr(10) + "dbl_ShowDPSupplyPressure=" + CStr(dbl_ShowDPSupplyPressure) _
    + Chr(13) + Chr(10) + "dbl_ShowMaxDP=" + CStr(dbl_ShowMaxDP) + Chr(13) + Chr(10) + "dbl_ShowMaxDPSupplyPressure=" + CStr(dbl_ShowMaxDPSupplyPressure) _
    + Chr(13) + Chr(10) + "str_Message=" + str_Message
#End If


'***************************
'Parameter validation
'***************************
If dbl_ShutOffDP < 0 Then GoTo CalcDP_41000_FTC_ATC_Error
If long_TrimPackageID < 0 Then GoTo CalcDP_41000_FTC_ATC_Error
If long_ActuatorPartID < 0 Then GoTo CalcDP_41000_FTC_ATC_Error
If dbl_RatedCv < 0 Then GoTo CalcDP_41000_FTC_ATC_Error

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4

'Determine requirements
Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature(), dbl_MaxInletTemperature)
dbl_ShowDP = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, _
    str_SpecifiedPressureUM)
dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, _
    str_SpecifiedTemperatureUM)
dbl_ShowShutoffDP = dbl_ShutOffDP
dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, _
    str_SpecifiedPressureUM)

'Determine ANSI requirements
Call GetANSIPressure(str_BodyMatl, dbl_MaxInletTemperature, _
    str_SelectedRating, dbl_AllowANSIPressure)

'Get product data
    Call Get41000Data(long_TrimPackageID, dbl_RatedCv, dbl_StemDiameter, _
    dbl_SeatDiameter, dbl_Kv, dbl_Au, dbl_PackingHeight, _
    dbl_StemStressArea, str_StemMaterial, dbl_ForcePilotSpring, dbl_Travel, _
    dbl_PlugSealDiameter, dbl_PlugThrottlingDiameter, dbl_DynPlugDiameter, _
    dbl_SealRingCoef, dbl_TECSealCoef, dbl_TECCoefJ, dbl_TECCoefS, _
    dbl_CageInsideDiameter, dbl_SealRingWidth, dbl_ForcePlugWt, dbl_MaxAllowFlatSpringDP, _
    dbl_SeatingDiameter, dbl_PilotTravel)

    Call GetActuatorData(long_ActuatorPartID, dbl_AreaInitial, dbl_AreaFinal, _
        dbl_MaxSupplyPressure, dbl_ActuatorFriction, dbl_PositionerLossCoef, _
        str_Model, long_AirActionID, _
        int_QtySeals, dbl_StemSealWidth, dbl_PistonRingWidth, _
        dbl_VolOver, dbl_ForceSpringClosed, dbl_ActuatorStemDiameter, dbl_LeverLength)
    
    Call GetTrimTypeFromPackage(long_TrimPackageID, long_TrimTypeID)
    Call GetPlugIDFromPackage(long_TrimPackageID, long_PlugTypeID)

'Calculate required actuator evaluation data
Call CalcSeatingForce(str_LeakageClass, dbl_SeatDiameter, _
    dbl_ForceSeating)
Call CalcPackingForce(dbl_MaxInletPressure, dbl_StemDiameter, _
    str_PackingMaterial, dbl_PackingHeight, dbl_ForcePacking)
Call CalcActuatorSealFrictionForce(dbl_AreaInitial, _
    dbl_ActuatorStemDiameter, dbl_PistonRingWidth, _
    dbl_StemSealWidth, int_QtySeals, dbl_MaxSupplyPressure, _
    dbl_ForceActSeal)
Call CalcPositionerLossForce(dbl_MaxSupplyPressure, dbl_AreaInitial, _
    dbl_ForcePositionerLoss)
Call CalcPlugSealRingFrictionForce(dbl_SealRingCoef, _
    dbl_TECSealCoef, dbl_TECCoefJ, dbl_TECCoefS, _
    dbl_CageInsideDiameter, dbl_SealRingWidth, _
    dbl_ForceSealRingArea, dbl_ForceSealRingSpring)

'Calculate stem limitations
dbl_L = 6# 'default value
Call GetStemStressArea(dbl_StemDiameter, dbl_StemStressArea)
Call GetStemProperties(str_StemMaterial, dbl_MaxInletTemperature, _
    dbl_Sy, dbl_E, dbl_Sd)
Call CalcAllowStemForce(dbl_Sy, dbl_E, dbl_StemDiameter, _
    dbl_L, dbl_Sd, dbl_StemStressArea, dbl_AllowableStemForce)

'Evaluate mechanical limit
'Call CalcMainPlugSeatStress_41000(dbl_MaxInletPressure, _
'    dbl_SeatDiameter, dbl_StemDiameter, dbl_AllowSeatYieldStress, _
'    dbl_OffBalanceForce, dbl_AllowSeatForce)
Call CalcRequiredSeatingForce_41000(dbl_MaxInletPressure, _
    dbl_ForceSeating, dbl_ForcePacking, dbl_ForceActSeal, _
    dbl_ForcePositionerLoss, dbl_ForceSealRingArea, _
    dbl_ForcePlugWt, dbl_RequiredSeatingForce)
Call CalcMaxDPCage_41000(dbl_ValveSize, long_TrimTypeID, _
    dbl_MaxInletTemperature, dbl_MaxAllowCageDP)

'Evaluation

If long_PlugTypeID = 3 Then
    Call CalcUnbalArea_41000_Pilot(dbl_SeatDiameter, _
    dbl_CageInsideDiameter, dbl_StemDiameter, _
    dbl_PlugSealDiameter, dbl_PlugThrottlingDiameter, _
    dbl_SealRingCoef, dbl_SealRingWidth, _
    dbl_B, dbl_RatedCv, str_FluidType, _
    dbl_AreaUnbalancedOpen, dbl_AreaUnbalancedClosed)

Else
    Call CalcUnbalArea_41000_Nonpilot(dbl_SeatDiameter, _
    dbl_CageInsideDiameter, dbl_StemDiameter, _
    dbl_PlugSealDiameter, dbl_PlugThrottlingDiameter, _
    dbl_SealRingCoef, dbl_SealRingWidth, dbl_B, _
    dbl_DynPlugDiameter, dbl_RatedCv, str_FluidType, _
    dbl_AreaUnbalancedOpen, dbl_AreaUnbalancedClosed)

End If

' Calculate allowable DP based on dynamic stability

Call CalcDynamicFrictionRateChange(dbl_MaxInletPressure, _
    dbl_ForceSealRingArea, dbl_ForceSealRingSpring, _
    dbl_ForcePacking, dbl_Travel, dbl_DynFrictionRateChange)
    
Call CalcSpringRate(str_SpringRange, dbl_AreaInitial, _
    dbl_PInitial, dbl_AreaFinal, dbl_PFinal, _
    str_SpecifiedCalibPressureUM, dbl_Travel, _
    str_SpecifiedLengthUM, dbl_SpringRate)

'B = 0.1
'Do
'Do While B=0.1
    If str_Model = "84" Then
        Call CalcSpringRateSACylinderATC(dbl_PInitial, _
        dbl_AreaInitial, dbl_SpringRate, dbl_Travel, _
        str_SpecifiedLengthUM, dbl_AreaUnbalancedOpen, _
        dbl_ShutOffDP, dbl_MaxSupplyPressure, _
        dbl_EffectiveSpringRate)
    
    ElseIf str_Model = "86" Then
        Call CalcSpringRateDACylinderATC(dbl_PInitial, _
        dbl_AreaInitial, dbl_SpringRate, dbl_Travel, _
        str_SpecifiedLengthUM, dbl_AreaUnbalancedOpen, _
        dbl_ShutOffDP, dbl_MaxSupplyPressure, _
        dbl_EffectiveSpringRate)

        dbl_SpringRate = dbl_EffectiveSpringRate
    'Exit Do
    End If
'Loop
'Loop Until B=1

dbl_AllowDPStability = dbl_SpringRate / dbl_Kv

' Calculate allowable DP when valve is closed

dbl_ActForceClosed = dbl_PFinal * dbl_AreaFinal

Call CalcAllowDPClosed_41000_FTC(dbl_MaxInletPressure, _
    dbl_ActForceClosed, dbl_ForcePacking, dbl_ForceActSeal, _
    dbl_ForcePositionerLoss, dbl_ForceSealRingSpring, _
    dbl_ForcePilotSpring, dbl_AreaUnbalancedClosed, _
    dbl_ForcePlugWt, dbl_AllowDPClosed)
    
' Calculate allowable DP when valve is open

dbl_ActuatorForceOpen = dbl_PInitial * dbl_AreaInitial

Call CalcAllowDPOpen_41000_FTC(dbl_MaxInletPressure, dbl_ActuatorForceOpen, _
    dbl_ForcePacking, dbl_ForceActSeal, _
    dbl_ForcePositionerLoss, dbl_ForceSealRingSpring, _
    dbl_AreaUnbalancedOpen, _
    dbl_ForcePlugWt, dbl_AllowDPOpen)

' Calculate limiting allowable DP

Call CalcMaxDPAllowable(dbl_AllowDPStability, _
    dbl_AllowDPClosed, dbl_AllowDPOpen, _
    dbl_AllowANSIPressure, dbl_MaxAllowDP)

Call CalcRequiredSupplyPressure41000_FTC_ATC(dbl_PFinal, dbl_AreaFinal, _
    dbl_ForcePilotSpring, dbl_ForcePacking, dbl_ForceSeating, dbl_PSupplyDP)

'Evaluation
str_Message = ""
dbl_ShowDPSupplyPressure = 0#
dbl_ShowMaxDPSupplyPressure = 0#

If (dbl_MaxInletPressure > dbl_AllowANSIPressure) Then
    str_Message = ANSIMESSAGE
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowANSIPressure, str_SpecifiedPressureUM)
End If

dbl_MaxAllowTrimDP = dbl_MaxAllowCageDP

If (dbl_MaxAllowFlatSpringDP < dbl_MaxAllowTrimDP And _
    dbl_MaxAllowFlatSpring <> 0) Then
    dbl_MaxAllowTrimDP = dbl_MaxAllowCageDP
End If

If (dbl_MaxInletPressure > dbl_MaxAllowTrimDP) Then
    str_Message = TRIMMESSAGE
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowTrimDP, _
        str_SpecifiedPressureUM)
End If

If (dbl_MaxInletPressure <= dbl_MaxAllowDP) Then
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG( _
        dbl_PSupplyDP, str_SpecifiedCalibPressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG( _
        dbl_MaxAllowDP, str_SpecifiedPressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG( _
        dbl_PSupplyDP, str_SpecifiedCalibPressureUM)
ElseIf (dbl_MaxInletPressure > dbl_MaxAllowDP And _
    dbl_ShutOffDP <= dbl_MaxAllowDP And dbl_ShutOffDP > 0#) Then
    str_Message = LIMITEDDPMESSAGE
    dbl_ShowDP = dbl_ShowShutoffDP
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG( _
        dbl_PSupplyDP, str_SpecifiedCalibPressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, _
        str_SpecifiedPressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG( _
        dbl_PSupplyDP, str_SpecifiedCalibPressureUM)
Else
    str_Message = ACTUATORMESSAGE
End If

dbl_ShowDP = TrimDecimalPlaces(dbl_ShowDP)
dbl_ShowDPSupplyPressure = TrimDecimalPlaces(dbl_ShowDPSupplyPressure)
dbl_ShowMaxDP = TrimDecimalPlaces(dbl_ShowMaxDP)
dbl_ShowMaxDPSupplyPressure = TrimDecimalPlaces(dbl_ShowMaxDPSupplyPressure)

#If (DebugVersion = "yes") Then
MsgBox "CalcDP_41000_FTC_ATC" + Chr(13) + Chr(10) _
    + Chr(13) + Chr(10) + "long_TrimPackageID=" + CStr(long_TrimPackageID) _
    + Chr(13) + Chr(10) + "long_ActuatorPartID=" + CStr(long_ActuatorPartID) + Chr(13) + Chr(10) + "dbl_InletPressure1=" + CStr(dbl_InletPressure1) _
    + Chr(13) + Chr(10) + "dbl_InletPressure2=" + CStr(dbl_InletPressure2) + Chr(13) + Chr(10) + "dbl_InletPressure3=" + CStr(dbl_InletPressure3) _
    + Chr(13) + Chr(10) + "dbl_InletPressure4=" + CStr(dbl_InletPressure4) + Chr(13) + Chr(10) + "dbl_OutletPressure1=" + CStr(dbl_OutletPressure1) + Chr(13) + Chr(10) + "dbl_OutletPressure2=" + CStr(dbl_OutletPressure2) _
    + Chr(13) + Chr(10) + "dbl_OutletPressure3=" + CStr(dbl_OutletPressure3) + Chr(13) + Chr(10) + "dbl_OutletPressure4=" + CStr(dbl_OutletPressure4) + Chr(13) + Chr(10) + "dbl_InletTemperature1=" + CStr(dbl_InletTemperature1) _
    + Chr(13) + Chr(10) + "dbl_InletTemperature2=" + CStr(dbl_InletTemperature2) + Chr(13) + Chr(10) + "dbl_InletTemperature3=" + CStr(dbl_InletTemperature3) + Chr(13) + Chr(10) + "dbl_InletTemperature4=" + CStr(dbl_InletTemperature4) + Chr(13) + Chr(10) + "dbl_ShutoffDP=" + CStr(dbl_ShutOffDP) _
    + Chr(13) + Chr(10) + "str_SpecifiedPressureUM=" + str_SpecifiedPressureUM + Chr(13) + Chr(10) + "str_SpecifiedTemperatureUM=" + str_SpecifiedTemperatureUM _
    + Chr(13) + Chr(10) + "str_SpecifiedCalibPressureUM=" + str_SpecifiedCalibPressureUM _
    + Chr(13) + Chr(10) + "str_SpecifiedLengthUM=" + str_SpecifiedLengthUM + Chr(13) + Chr(10) + "str_BodyMatl=" + str_BodyMatl _
    + Chr(13) + Chr(10) + "str_SelectedRating=" + str_SelectedRating + Chr(13) + Chr(10) + "dbl_ValveSize=" + CStr(dbl_ValveSize) _
    + Chr(13) + Chr(10) + "str_PackingMaterial=" + str_PackingMaterial + Chr(13) + Chr(10) + "str_LeakageClass=" + str_LeakageClass _
    + Chr(13) + Chr(10) + "str_SpringRange=" + str_SpringRange + Chr(13) + Chr(10) + "dbl_RatedCv=" + CStr(dbl_RatedCv) _
    + Chr(13) + Chr(10) + "dbl_ShowDP=" + CStr(dbl_ShowDP) + Chr(13) + Chr(10) + "dbl_ShowDPSupplyPressure=" + CStr(dbl_ShowDPSupplyPressure) _
    + Chr(13) + Chr(10) + "dbl_ShowMaxDP=" + CStr(dbl_ShowMaxDP) + Chr(13) + Chr(10) + "dbl_ShowMaxDPSupplyPressure=" + CStr(dbl_ShowMaxDPSupplyPressure) _
    + Chr(13) + Chr(10) + "str_Message=" + str_Message
#End If


Exit Sub
CalcDP_41000_FTC_ATC_Error:

End Sub

Public Sub CalcDP_41000_FTC_ATO(long_TrimPackageID As Long, _
    long_ActuatorPartID As Long, dbl_InletPressure1 As Double, dbl_InletPressure2 As Double, dbl_InletPressure3 As Double, _
    dbl_InletPressure4 As Double, dbl_OutletPressure1 As Double, dbl_OutletPressure2 As Double, _
    dbl_OutletPressure3 As Double, dbl_OutletPressure4 As Double, dbl_InletTemperature1 As Double, _
    dbl_InletTemperature2 As Double, dbl_InletTemperature3 As Double, dbl_InletTemperature4 As Double, dbl_ShutOffDP _
    As Double, str_SpecifiedPressureUM As String, str_SpecifiedTemperatureUM _
    As String, str_SpecifiedCalibPressureUM As String, _
    str_SpecifiedLengthUM As String, str_FluidType As String, str_BodyMatl As String, _
    str_SelectedRating As String, dbl_ValveSize As Double, _
    str_PackingMaterial As String, str_LeakageClass As String, _
    str_SpringRange As String, dbl_RatedCv As Double, _
    dbl_ShowDP As Double, dbl_ShowDPSupplyPressure As Double, _
    dbl_ShowMaxDP As Double, dbl_ShowMaxDPSupplyPressure As Double, _
    str_Message As String)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'calcDP_41000_FTC_ATO(----see input output list below-)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Reference: CES 1019 Balanced and pilot balanced cage throttling
'   globe valves - Actuator Sizing Methods.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This subroutine is used to calculate the selected valve and
'  actuator shutoff capability and compare the result to defined
'  requirements.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This subroutine is specific to 41000 series valves, flow to close,
'  air to open.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Input Variables
'
'dbl_InletPressure(1 to 4)      specified inlet pressure (all cases)
'dbl_OutletPressure(1 to 4)     specified outlet pressure (all cases)
'dbl_InletTemperature(1 to 4)   specified inlet temperature (all cases)
'dbl_ShutoffDP                  specified maximum required shutoff
'                               pressure Drop
'str_SpecifiedPressureUM        unit of measure associated with inlet
'                               and outlet pressures
'str_SpecifiedTemperatureUM     unit of measure associated with inlet
'                               temperature
'str_SpecifiedCalibPressureUM   unit of measure assoeicated with spring
'                               range and supply pressure
'str_SpecifiedLengthUM          units of measure for length
'str_BodyMatl                   selected material of body
'str_SelectedRating             selected body rating
'long_TrimPackageID             ID of trim selected
'long_ActuatorID                ID of actuator selected
'dbl_ValveSize                  valve size
'str_PackingMaterial            selected packing material
'str_LeakageClass               selected leakage class
'str_SpringRange                selected actuator spring range
'dbl_RatedCv                    selected rated Cv
'
'Output Variables:
'dbl_ShowDP                     specified shutoff requirement (worst case)
'dbl_ShowDPsupplyPressure       calculated required supply pressure
'                               for specified
'dbl_ShowMaxDP                  calculated maximum pressure drop for
'                               selected valve and actuator combination
'dbl_ShowMaxDPSupplyPressure    calculated required supply pressure
'                               for maximum pressure drop
'str_Message                    warning message if pressure drop
'                               requirement exceeds equipment capability
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'coded: 10/16/98    Frank Krieger
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo CalcDP_41000_FTC_ATO_Error

'***************************
'Declarations
'***************************

Dim dbl_ActForceClosed As Double
Dim dbl_ActuatorForceOpen As Double
Dim dbl_ActuatorFriction As Double
Dim dbl_ActuatorStemDiameter As Double
Dim dbl_AllowableStemForce As Double
Dim dbl_AllowANSIPressure As Double
Dim dbl_AllowDP As Double
Dim dbl_AllowDPClosed As Double
Dim dbl_AllowDPOpen As Double
Dim dbl_AllowDPStability As Double
Dim dbl_AllowDPStem As Double
Dim dbl_AllowSeatForce As Double
Dim dbl_AllowSeatYieldStress As Double
Dim dbl_AreaFinal As Double
Dim dbl_AreaInitial As Double
Dim dbl_AreaUnbalancedClosed As Double
Dim dbl_AreaUnbalancedOpen As Double
Dim dbl_As As Double
Dim dbl_Au As Double
Dim dbl_B As Double
Dim dbl_CageInsideDiameter As Double
Dim dbl_DynFrictionRateChange As Double
Dim dbl_DynPlugDiameter As Double
Dim dbl_E As Double
Dim dbl_EffectiveSpringRate As Double
Dim dbl_Fcs As Double
Dim dbl_ForceActSeal As Double
Dim dbl_ForceOffBalance As Double
Dim dbl_ForcePacking As Double
Dim dbl_ForcePilotSpring As Double
Dim dbl_ForcePlugWt As Double
Dim dbl_ForcePositionerLoss As Double
Dim dbl_ForceSealRingArea As Double
Dim dbl_ForceSealRingSpring As Double
Dim dbl_ForceSeating As Double
Dim dbl_ForceSpringClosed As Double
Dim dbl_InletPressure(4) As Double
Dim dbl_InletTemperature(4) As Double
Dim dbl_Kv As Double
Dim dbl_L As Double
Dim dbl_LeverLength As Double
Dim dbl_MaxAllowCageDP As Double
Dim dbl_MaxAllowDP As Double
Dim dbl_MaxAllowFlatSpring As Double
Dim dbl_MaxAllowFlatSpringDP As Double
Dim dbl_MaxAllowTrimDP As Double
Dim dbl_MaxInletPressure As Double
Dim dbl_MaxInletTemperature As Double
Dim dbl_MaxSupplyPressure As Double
Dim dbl_OutletPressure(4) As Double
Dim dbl_PackingHeight As Double
Dim dbl_PFinal As Double
Dim dbl_PilotTravel As Double
Dim dbl_PInitial As Double
Dim dbl_PistonRingWidth As Double
Dim dbl_PlugSealDiameter As Double
Dim dbl_PlugThrottlingArea As Double
Dim dbl_PlugThrottlingDiameter As Double
Dim dbl_PositionerLossCoef As Double
Dim dbl_PSupplyDP As Double
Dim dbl_PSupplyDPANSI As Double
Dim dbl_PSupplyDPShutoff As Double
Dim dbl_PSupplyDPStem As Double
Dim dbl_RequiredActuatorForce As Double
Dim dbl_RequiredANSIActuatorForce As Double
Dim dbl_RequiredSeatingForce As Double
Dim dbl_RequiredShutoffDPActuatorForce As Double
Dim dbl_Sd As Double
Dim dbl_SealRingCoef As Double
Dim dbl_SealRingWidth As Double
Dim dbl_SeatDiameter As Double
Dim dbl_SeatingDiameter As Double
Dim dbl_ShowMaxDPSupply As Double
Dim dbl_ShowMaxInletTemperature As Double
Dim dbl_ShowShutoffDP As Double
Dim dbl_SpringEnergizedSealSpringCoef As Double
Dim dbl_SpringRate As Double
Dim dbl_StemDiameter As Double
Dim dbl_StemSealWidth As Double
Dim dbl_StemStressArea As Double
Dim dbl_Sy As Double
Dim dbl_TECCoefJ As Double
Dim dbl_TECCoefS As Double
Dim dbl_TECSealCoef As Double
Dim dbl_Travel As Double
Dim dbl_VolOver As Double
Dim int_QtySeals As Integer
Dim long_AirActionID As Long
Dim long_CageTypeID As Long
Dim long_FluidTypeID As Long
Dim long_PlugTypeID As Long
Dim long_TrimTypeID As Long
Dim str_ActuatorModel As String
Dim str_Model As String
Dim str_StemMaterial As String

#If (DebugVersion = "yes") Then
MsgBox "CalcDP_41000_FTC_ATO" + Chr(13) + Chr(10) _
    + Chr(13) + Chr(10) + "long_TrimPackageID=" + CStr(long_TrimPackageID) _
    + Chr(13) + Chr(10) + "long_ActuatorPartID=" + CStr(long_ActuatorPartID) + Chr(13) + Chr(10) + "dbl_InletPressure1=" + CStr(dbl_InletPressure1) _
    + Chr(13) + Chr(10) + "dbl_InletPressure2=" + CStr(dbl_InletPressure2) + Chr(13) + Chr(10) + "dbl_InletPressure3=" + CStr(dbl_InletPressure3) _
    + Chr(13) + Chr(10) + "dbl_InletPressure4=" + CStr(dbl_InletPressure4) + Chr(13) + Chr(10) + "dbl_OutletPressure1=" + CStr(dbl_OutletPressure1) + Chr(13) + Chr(10) + "dbl_OutletPressure2=" + CStr(dbl_OutletPressure2) _
    + Chr(13) + Chr(10) + "dbl_OutletPressure3=" + CStr(dbl_OutletPressure3) + Chr(13) + Chr(10) + "dbl_OutletPressure4=" + CStr(dbl_OutletPressure4) + Chr(13) + Chr(10) + "dbl_InletTemperature1=" + CStr(dbl_InletTemperature1) _
    + Chr(13) + Chr(10) + "dbl_InletTemperature2=" + CStr(dbl_InletTemperature2) + Chr(13) + Chr(10) + "dbl_InletTemperature3=" + CStr(dbl_InletTemperature3) + Chr(13) + Chr(10) + "dbl_InletTemperature4=" + CStr(dbl_InletTemperature4) + Chr(13) + Chr(10) + "dbl_ShutoffDP=" + CStr(dbl_ShutOffDP) _
    + Chr(13) + Chr(10) + "str_SpecifiedPressureUM=" + str_SpecifiedPressureUM + Chr(13) + Chr(10) + "str_SpecifiedTemperatureUM=" + str_SpecifiedTemperatureUM _
    + Chr(13) + Chr(10) + "str_SpecifiedCalibPressureUM=" + str_SpecifiedCalibPressureUM _
    + Chr(13) + Chr(10) + "str_SpecifiedLengthUM=" + str_SpecifiedLengthUM + Chr(13) + Chr(10) + "str_BodyMatl=" + str_BodyMatl _
    + Chr(13) + Chr(10) + "str_SelectedRating=" + str_SelectedRating + Chr(13) + Chr(10) + "dbl_ValveSize=" + CStr(dbl_ValveSize) _
    + Chr(13) + Chr(10) + "str_PackingMaterial=" + str_PackingMaterial + Chr(13) + Chr(10) + "str_LeakageClass=" + str_LeakageClass _
    + Chr(13) + Chr(10) + "str_SpringRange=" + str_SpringRange + Chr(13) + Chr(10) + "dbl_RatedCv=" + CStr(dbl_RatedCv) _
    + Chr(13) + Chr(10) + "dbl_ShowDP=" + CStr(dbl_ShowDP) + Chr(13) + Chr(10) + "dbl_ShowDPSupplyPressure=" + CStr(dbl_ShowDPSupplyPressure) _
    + Chr(13) + Chr(10) + "dbl_ShowMaxDP=" + CStr(dbl_ShowMaxDP) + Chr(13) + Chr(10) + "dbl_ShowMaxDPSupplyPressure=" + CStr(dbl_ShowMaxDPSupplyPressure) _
    + Chr(13) + Chr(10) + "str_Message=" + str_Message
#End If

'***************************
'Parameter validation
'***************************
If dbl_ShutOffDP < 0 Then GoTo CalcDP_41000_FTC_ATO_Error
If long_TrimPackageID < 0 Then GoTo CalcDP_41000_FTC_ATO_Error
If long_ActuatorPartID < 0 Then GoTo CalcDP_41000_FTC_ATO_Error
If dbl_RatedCv < 0 Then GoTo CalcDP_41000_FTC_ATO_Error

'****************************
'Begin calculations
'****************************
'set up arrays
dbl_InletPressure(1) = dbl_InletPressure1
dbl_InletPressure(2) = dbl_InletPressure2
dbl_InletPressure(3) = dbl_InletPressure3
dbl_InletPressure(4) = dbl_InletPressure4
dbl_OutletPressure(1) = dbl_OutletPressure1
dbl_OutletPressure(2) = dbl_OutletPressure2
dbl_OutletPressure(3) = dbl_OutletPressure3
dbl_OutletPressure(4) = dbl_OutletPressure4
dbl_InletTemperature(1) = dbl_InletTemperature1
dbl_InletTemperature(2) = dbl_InletTemperature2
dbl_InletTemperature(3) = dbl_InletTemperature3
dbl_InletTemperature(4) = dbl_InletTemperature4

'Determine requirements
Call CalcTestPressure(dbl_InletPressure(), dbl_ShutOffDP, dbl_MaxInletPressure)
Call CalcTestTemperature(dbl_InletTemperature(), dbl_MaxInletTemperature)
dbl_ShowDP = dbl_MaxInletPressure
dbl_MaxInletPressure = ConvertPressureToPSIG(dbl_MaxInletPressure, _
    str_SpecifiedPressureUM)
dbl_ShowMaxInletTemperature = dbl_MaxInletTemperature
dbl_MaxInletTemperature = ConvertTemperatureToDegF(dbl_MaxInletTemperature, _
    str_SpecifiedTemperatureUM)
dbl_ShowShutoffDP = dbl_ShutOffDP
dbl_ShutOffDP = ConvertPressureToPSIG(dbl_ShutOffDP, _
    str_SpecifiedPressureUM)

'Determine ANSI requirements
Call GetANSIPressure(str_BodyMatl, dbl_MaxInletTemperature, _
    str_SelectedRating, dbl_AllowANSIPressure)

'Get product data
    Call Get41000Data(long_TrimPackageID, dbl_RatedCv, dbl_StemDiameter, _
    dbl_SeatDiameter, dbl_Kv, dbl_Au, dbl_PackingHeight, _
    dbl_StemStressArea, str_StemMaterial, dbl_ForcePilotSpring, dbl_Travel, _
    dbl_PlugSealDiameter, dbl_PlugThrottlingDiameter, dbl_DynPlugDiameter, _
    dbl_SealRingCoef, dbl_TECSealCoef, dbl_TECCoefJ, dbl_TECCoefS, _
    dbl_CageInsideDiameter, dbl_SealRingWidth, dbl_ForcePlugWt, dbl_MaxAllowFlatSpringDP, _
    dbl_SeatingDiameter, dbl_PilotTravel)

    Call GetActuatorData(long_ActuatorPartID, dbl_AreaInitial, dbl_AreaFinal, _
        dbl_MaxSupplyPressure, dbl_ActuatorFriction, dbl_PositionerLossCoef, _
        str_Model, long_AirActionID, _
        int_QtySeals, dbl_StemSealWidth, dbl_PistonRingWidth, _
        dbl_VolOver, dbl_ForceSpringClosed, dbl_ActuatorStemDiameter, dbl_LeverLength)
        
    Call GetTrimTypeFromPackage(long_TrimPackageID, long_TrimTypeID)
    Call GetPlugIDFromPackage(long_TrimPackageID, long_PlugTypeID)

'Calculate required actuator evaluation data
Call CalcSeatingForce(str_LeakageClass, dbl_SeatDiameter, _
    dbl_ForceSeating)
Call CalcPackingForce(dbl_MaxInletPressure, dbl_StemDiameter, _
    str_PackingMaterial, dbl_PackingHeight, dbl_ForcePacking)
Call CalcActuatorSealFrictionForce(dbl_AreaInitial, _
    dbl_ActuatorStemDiameter, dbl_PistonRingWidth, _
    dbl_StemSealWidth, int_QtySeals, dbl_MaxSupplyPressure, _
    dbl_ForceActSeal)
Call CalcPositionerLossForce(dbl_MaxSupplyPressure, dbl_AreaInitial, _
    dbl_ForcePositionerLoss)
Call CalcPlugSealRingFrictionForce(dbl_SealRingCoef, _
    dbl_TECSealCoef, dbl_TECCoefJ, dbl_TECCoefS, _
    dbl_CageInsideDiameter, dbl_SealRingWidth, _
    dbl_ForceSealRingArea, dbl_ForceSealRingSpring)

'Calculate stem limitations
dbl_L = 6# 'default value
Call GetStemStressArea(dbl_StemDiameter, dbl_StemStressArea)
Call GetStemProperties(str_StemMaterial, dbl_MaxInletTemperature, _
    dbl_Sy, dbl_E, dbl_Sd)
Call CalcAllowStemForce(dbl_Sy, dbl_E, dbl_StemDiameter, _
    dbl_L, dbl_Sd, dbl_StemStressArea, dbl_AllowableStemForce)

'Evaluate mechanical limit
'Call CalcMainPlugSeatStress_41000(dbl_MaxInletPressure, _
'    dbl_SeatDiameter, dbl_StemDiameter, _
'    dbl_AllowSeatYieldStress, dbl_ForceOffBalance, _
'    dbl_AllowSeatForce)  removed per JH 12/11/98
Call CalcRequiredSeatingForce_41000(dbl_MaxInletPressure, _
    dbl_ForceSeating, dbl_ForcePacking, dbl_ForceActSeal, _
    dbl_ForcePositionerLoss, dbl_ForceSealRingArea, _
    dbl_ForcePlugWt, dbl_RequiredSeatingForce)
Call CalcMaxDPCage_41000(dbl_ValveSize, long_TrimTypeID, _
    dbl_MaxInletTemperature, dbl_MaxAllowCageDP)

'Calculate unbalanced areas

If long_PlugTypeID = 3 Then
    Call CalcUnbalArea_41000_Pilot(dbl_SeatDiameter, _
    dbl_CageInsideDiameter, dbl_StemDiameter, _
    dbl_PlugSealDiameter, dbl_PlugThrottlingDiameter, _
    dbl_SealRingCoef, dbl_SealRingWidth, dbl_B, _
    dbl_RatedCv, str_FluidType, _
    dbl_AreaUnbalancedOpen, dbl_AreaUnbalancedClosed)

Else
    Call CalcUnbalArea_41000_Nonpilot(dbl_SeatDiameter, _
    dbl_CageInsideDiameter, dbl_StemDiameter, _
    dbl_PlugSealDiameter, dbl_PlugThrottlingDiameter, _
    dbl_SealRingCoef, dbl_SealRingWidth, dbl_B, _
    dbl_DynPlugDiameter, dbl_RatedCv, str_FluidType, _
    dbl_AreaUnbalancedOpen, dbl_AreaUnbalancedClosed)

End If

' Calculate allowable DP based on dynamic stability

Call CalcDynamicFrictionRateChange(dbl_MaxInletPressure, _
    dbl_ForceSealRingArea, dbl_ForceSealRingSpring, _
    dbl_ForcePacking, dbl_Travel, dbl_DynFrictionRateChange)
    
Call CalcSpringRate(str_SpringRange, dbl_AreaInitial, _
    dbl_PInitial, dbl_AreaFinal, dbl_PFinal, _
    str_SpecifiedCalibPressureUM, dbl_Travel, _
    str_SpecifiedLengthUM, dbl_SpringRate)

If str_Model = "85" Then
    Call CalcSpringRateSACylinderATO(dbl_PInitial, _
    dbl_AreaInitial, dbl_SpringRate, dbl_Travel, _
    str_SpecifiedLengthUM, dbl_AreaUnbalancedOpen, _
    dbl_ShutOffDP, dbl_MaxSupplyPressure, _
    dbl_EffectiveSpringRate)
    
    dbl_SpringRate = dbl_EffectiveSpringRate
    
ElseIf str_Model = "86" Then
    Call CalcSpringRateDACylinderATO(dbl_PInitial, _
    dbl_AreaInitial, dbl_SpringRate, dbl_Travel, _
    str_SpecifiedLengthUM, dbl_AreaUnbalancedOpen, _
    dbl_ShutOffDP, dbl_MaxSupplyPressure, _
    dbl_EffectiveSpringRate)

    dbl_SpringRate = dbl_EffectiveSpringRate

End If

dbl_AllowDPStability = dbl_SpringRate / dbl_Kv

' Calculate allowable DP when valve is closed

dbl_ActForceClosed = dbl_PFinal * dbl_AreaFinal

Call CalcAllowDPClosed_41000_FTC(dbl_MaxInletPressure, _
    dbl_ActForceClosed, dbl_ForcePacking, dbl_ForceActSeal, _
    dbl_ForcePositionerLoss, dbl_ForceSealRingSpring, _
    dbl_ForcePilotSpring, dbl_AreaUnbalancedClosed, _
    dbl_ForcePlugWt, dbl_AllowDPClosed)
    
' Calculate allowable DP when valve is open

dbl_ActuatorForceOpen = dbl_PInitial * dbl_AreaInitial

Call CalcAllowDPOpen_41000_FTC(dbl_MaxInletPressure, dbl_ActuatorForceOpen, _
    dbl_ForcePacking, dbl_ForceActSeal, _
    dbl_ForcePositionerLoss, dbl_ForceSealRingSpring, _
    dbl_AreaUnbalancedOpen, _
    dbl_ForcePlugWt, dbl_AllowDPOpen)

' Calculate limiting allowable DP

Call CalcMaxDPAllowable(dbl_AllowDPStability, _
    dbl_AllowDPClosed, dbl_AllowDPOpen, _
    dbl_AllowANSIPressure, dbl_MaxAllowDP)
   
 Call CalcRequiredSupplyPressure41000_FTC_ATO(dbl_PFinal, dbl_AreaFinal, _
    dbl_ForcePacking, dbl_Au, dbl_MaxInletPressure, dbl_PSupplyDP)
  

'Evaluation
str_Message = ""
dbl_ShowDPSupplyPressure = 0#
dbl_ShowMaxDPSupplyPressure = 0#

If (dbl_MaxInletPressure > dbl_AllowANSIPressure) Then
    str_Message = ANSIMESSAGE
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_AllowANSIPressure, str_SpecifiedPressureUM)
End If

dbl_MaxAllowTrimDP = dbl_MaxAllowCageDP

If (dbl_MaxAllowFlatSpringDP < dbl_MaxAllowTrimDP And _
    dbl_MaxAllowFlatSpring <> 0) Then
    dbl_MaxAllowTrimDP = dbl_MaxAllowCageDP
End If

If (dbl_MaxInletPressure > dbl_MaxAllowTrimDP) Then
    str_Message = TRIMMESSAGE
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowTrimDP, _
        str_SpecifiedPressureUM)
End If

If (dbl_MaxInletPressure <= dbl_MaxAllowDP) Then
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG( _
        dbl_PSupplyDP, str_SpecifiedCalibPressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG( _
        dbl_MaxAllowDP, str_SpecifiedPressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG( _
        dbl_PSupplyDP, str_SpecifiedCalibPressureUM)
ElseIf (dbl_MaxInletPressure > dbl_MaxAllowDP And _
    dbl_ShutOffDP <= dbl_MaxAllowDP And dbl_ShutOffDP > 0#) Then
    str_Message = LIMITEDDPMESSAGE
    dbl_ShowDP = dbl_ShowShutoffDP
    dbl_ShowDPSupplyPressure = ConvertPressureFromPSIG( _
        dbl_PSupplyDP, str_SpecifiedCalibPressureUM)
    dbl_ShowMaxDP = ConvertPressureFromPSIG(dbl_MaxAllowDP, _
        str_SpecifiedPressureUM)
    dbl_ShowMaxDPSupplyPressure = ConvertPressureFromPSIG( _
        dbl_PSupplyDP, str_SpecifiedCalibPressureUM)
Else
    str_Message = ACTUATORMESSAGE
End If

dbl_ShowDP = TrimDecimalPlaces(dbl_ShowDP)
dbl_ShowDPSupplyPressure = TrimDecimalPlaces(dbl_ShowDPSupplyPressure)
dbl_ShowMaxDP = TrimDecimalPlaces(dbl_ShowMaxDP)
dbl_ShowMaxDPSupplyPressure = TrimDecimalPlaces(dbl_ShowMaxDPSupplyPressure)

#If (DebugVersion = "yes") Then
MsgBox "CalcDP_41000_FTC_ATO" + Chr(13) + Chr(10) _
    + Chr(13) + Chr(10) + "long_TrimPackageID=" + CStr(long_TrimPackageID) _
    + Chr(13) + Chr(10) + "long_ActuatorPartID=" + CStr(long_ActuatorPartID) + Chr(13) + Chr(10) + "dbl_InletPressure1=" + CStr(dbl_InletPressure1) _
    + Chr(13) + Chr(10) + "dbl_InletPressure2=" + CStr(dbl_InletPressure2) + Chr(13) + Chr(10) + "dbl_InletPressure3=" + CStr(dbl_InletPressure3) _
    + Chr(13) + Chr(10) + "dbl_InletPressure4=" + CStr(dbl_InletPressure4) + Chr(13) + Chr(10) + "dbl_OutletPressure1=" + CStr(dbl_OutletPressure1) + Chr(13) + Chr(10) + "dbl_OutletPressure2=" + CStr(dbl_OutletPressure2) _
    + Chr(13) + Chr(10) + "dbl_OutletPressure3=" + CStr(dbl_OutletPressure3) + Chr(13) + Chr(10) + "dbl_OutletPressure4=" + CStr(dbl_OutletPressure4) + Chr(13) + Chr(10) + "dbl_InletTemperature1=" + CStr(dbl_InletTemperature1) _
    + Chr(13) + Chr(10) + "dbl_InletTemperature2=" + CStr(dbl_InletTemperature2) + Chr(13) + Chr(10) + "dbl_InletTemperature3=" + CStr(dbl_InletTemperature3) + Chr(13) + Chr(10) + "dbl_InletTemperature4=" + CStr(dbl_InletTemperature4) + Chr(13) + Chr(10) + "dbl_ShutoffDP=" + CStr(dbl_ShutOffDP) _
    + Chr(13) + Chr(10) + "str_SpecifiedPressureUM=" + str_SpecifiedPressureUM + Chr(13) + Chr(10) + "str_SpecifiedTemperatureUM=" + str_SpecifiedTemperatureUM _
    + Chr(13) + Chr(10) + "str_SpecifiedCalibPressureUM=" + str_SpecifiedCalibPressureUM _
    + Chr(13) + Chr(10) + "str_SpecifiedLengthUM=" + str_SpecifiedLengthUM + Chr(13) + Chr(10) + "str_BodyMatl=" + str_BodyMatl _
    + Chr(13) + Chr(10) + "str_SelectedRating=" + str_SelectedRating + Chr(13) + Chr(10) + "dbl_ValveSize=" + CStr(dbl_ValveSize) _
    + Chr(13) + Chr(10) + "str_PackingMaterial=" + str_PackingMaterial + Chr(13) + Chr(10) + "str_LeakageClass=" + str_LeakageClass _
    + Chr(13) + Chr(10) + "str_SpringRange=" + str_SpringRange + Chr(13) + Chr(10) + "dbl_RatedCv=" + CStr(dbl_RatedCv) _
    + Chr(13) + Chr(10) + "dbl_ShowDP=" + CStr(dbl_ShowDP) + Chr(13) + Chr(10) + "dbl_ShowDPSupplyPressure=" + CStr(dbl_ShowDPSupplyPressure) _
    + Chr(13) + Chr(10) + "dbl_ShowMaxDP=" + CStr(dbl_ShowMaxDP) + Chr(13) + Chr(10) + "dbl_ShowMaxDPSupplyPressure=" + CStr(dbl_ShowMaxDPSupplyPressure) _
    + Chr(13) + Chr(10) + "str_Message=" + str_Message
#End If

Exit Sub
CalcDP_41000_FTC_ATO_Error:

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
' END OF 41000 UL
''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''
' INCLUDE Gas UL
''''''''''''''''''''''''''''''''''''''''''''''''''''

'---------------------------------------------------------------------
Public Sub CalcCn(str_ValveLine As String, dbl_Cv As Double, _
    dbl_RatedCv As Double, dbl_P1 As Double, dbl_P2 As Double, _
    dbl_Cn As Double)

'---------------------------------------------------------------------
'Source:  IEC 534-8
'
'This subroutine calculates the last stage Cv of a multistage valve.
'
'Coded by:              F.Krieger
'Date:                  10/19/98
'---------------------------------------------------------------------
'
'Input Variables:
'dbl_Cv                 calculated required Cv
'dbl_Rated Cv           Rated Cv of the valve
'dbl_P1                 Inlet Pressure (Absolute)
'dbl_P2                 Outlet Pressure (Absolute)
'str_ValveLine

'Output Variables:
'dbl_Cn                 Last Stage Cv

On Error GoTo CalcCn_Error

'***************************
'Check input values
'***************************

'dbl_RatedCv must be greater than 0
If (dbl_RatedCv = 0) Then GoTo CalcCn_Error

'**************************
'Declarations
'**************************

Dim dbl_X As Double

    If (str_ValveLine = 41000 Or str_ValveLine = 72000) Then

        dbl_X = (3.5 - (2 * (dbl_Cv / dbl_RatedCv))) * (1 - (dbl_P2 / dbl_P1) ^ (0.3 + (2 * (dbl_Cv) / dbl_RatedCv)))
        dbl_Cn = ((dbl_X ^ 2 + 1) ^ 0.5) * dbl_Cv

    ElseIf str_ValveLine = 21000 Then

        dbl_X = 1.5 - (1.5 * (dbl_P2 / dbl_P1) ^ 2)
        dbl_Cn = ((dbl_X ^ 2 + 1) ^ 0.5) * dbl_Cv

    End If

CalcCn_Error:
' Include Error Message
End Sub

Public Sub CalcPn(dbl_Cv As Double, dbl_Cn As Double, dbl_P1 As Double, _
    dbl_P2 As Double, dbl_Pn As Double)
    
'Source: IEC 534-8-3
'This subroutine calculates the stagnation pressure of a multistage
'trim.
'
'Coded by:             F.Krieger
'Date:                 10/19/98
'---------------------------------------------------------------------

'Input variables:
'dbl_Cv                Calculated required Cv
'dbl_P1                Inlet Pressure (Pa absolute)
'dbl_P2                Outlet Pressure (Pa absolute)
'dbl_Cn                Last stage Cv

'Output variables:
'dbl_Pn                Stagnation Pressure (absolute)
'---------------------------------------------------------------------

On Error GoTo CalcPn_Error

'************************
'Declarations
'************************
Dim dbl_PRatio As Double
Dim dbl_RRatio As Double

'************************
'Begin calculations
'************************

dbl_PRatio = dbl_P1 / dbl_P2

If dbl_PRatio >= 2 Then
    dbl_Pn = Sqr(((dbl_P1 * dbl_Cv) / (1.155 * dbl_Cn)) ^ 2 + (dbl_P2) ^ 2)
    
    If dbl_Pn >= (2 * dbl_P2) Then
        dbl_Pn = dbl_P1 * (dbl_Cv / dbl_Cn)
    End If
End If

dbl_RRatio = dbl_Pn / dbl_P2

If dbl_PRatio >= 2 And dbl_RRatio < 2 Then
    dbl_Pn = Sqr(((dbl_P1 * dbl_Cv) / (1.155 * dbl_Cn)) ^ 2 + (dbl_P2) ^ 2)
    
ElseIf dbl_PRatio >= 2 And dbl_RRatio >= 2 Then
        dbl_Pn = dbl_P1 * (dbl_Cv / dbl_Cn)
    
ElseIf dbl_PRatio < 2 Then
        dbl_Pn = Sqr(((dbl_Cv / dbl_Cn) ^ 2) * (dbl_P1 ^ 2 - dbl_P2 ^ 2) + (dbl_P2 ^ 2))
End If

CalcPn_Error:
'Insert Error Message
End Sub


Public Sub CalcMultiStageLodbJetDia(dbl_Fd, dbl_LDRatio, dbl_Cn, dbl_Dj)

'Source:  IEC 534-8-3, Eq.49
'This subroutine calculates the jet diameter for multistage lodb trim
'used by the gas noise calculation subroutines
'
'Coded by F.Krieger
'Date: 10/19/98
'---------------------------------------------------------------------
'
'Input Variables:
'dbl_Cn                Cv for last stage of trim
'dbl_Fd                Last stage Fd
'dbl_Ldratio

'Output Variables:
'dbl_Dj                Internal jet diameter, m

On Error GoTo CalcMultiStageLodbJetDia_Error

'************************
'Declarations
'************************

'Dim dbl_Fd As Double
'Dim dbl_Ldratio As Double
'Dim dbl_Cn As Double
'Dim dbl_Dj As Double

'************************
'Begin calculations
'************************

dbl_Dj = ((4.6 * 10 ^ -3) * dbl_Fd) * Sqr(dbl_Cn * (0.9 - (0.06 * dbl_LDRatio)))

CalcMultiStageLodbJetDia_Error:
'Insert Error Message

End Sub

Public Sub HighMachCorrection(str_ValveLine As String, lng_TrimTypeID As Long, _
    dbl_ValveInsideDia_m As Double, dbl_MassFlowRate As Double, dbl_OutletDensity As Double, _
    dbl_PaPa As Double, dbl_OutletPipeOD_m As Double, dbl_OutletPipeWallThickness_m As Double, _
    dbl_ValveOutletDia_m As Double, dbl_C2 As Double, dbl_Lg As Double, _
    dbl_LpAe1m As Double, dbl_Lpo As Double)
    
'Source: Section 7 of IEC 60534-8-3
'This subroutine calculates the combined A-weighted sound pressure
'level for the valve trim, valve and expander.
'
'Coded by F.Krieger
'Date: 10/19/98
'------------------------------------------------------------------------------------------
'
'Input Variables:
'str_Valveline                 valve model number
'lng_TrimTypeID                  valve trim style
'dbl_ValveInsideDia_m          inside diameter of the valve outlet, m
'dbl_MassFlowRate              mass flow rate, kg/s
'dbl_OutletDensity             fluid density at outlet conditions, kg/m3
'dbl_PaPa                      actual atmospheric pressure outside the pipe
'dbl_OutletPipeOD_m            actual outside diameter of pipe, m
'dbl_OutletPipeWallThickness_m pipe wall thickness, m
'dbl_ValveOutletDia_m          size of valve outlet, m
'dbl_C2                        sonic velocity
'dbl_Lg                        downstream pipe velocity correction
'dbl_LpAe1m                    A-weighted sound pressure level 1 meter from pipe wall, dBA

'Output Variables:
'dbl_Lpo                       combined A-weighted sound pressure level of the valve
'                              expander and valve trim 1 meter from pipe wall

On Error GoTo HighMachCorrection_Error

'************************
'Declarations
'************************
Dim dbl_PeakFrequency As Double
Dim dbl_Di As Double
Dim dbl_Up As Double
Dim dbl_Ur As Double
Dim dbl_Wmr As Double
Dim dbl_War As Double
Dim dbl_Fpr As Double
Dim dbl_Mr As Double
Dim dbl_Nr As Double
Dim dbl_Wr As Double
Dim dbl_Lpir As Double
Dim dbl_Lpe As Double
Dim dbl_TLr As Double
Dim dbl_PipeWallThickness As Double
Dim dbl_Pa As Double

Const dbl_Ps As Double = 101325#
Const dbl_B As Double = 0.93                 ' Contraction Coefficient

'***************************
'Parameter validation
'***************************
If dbl_ValveInsideDia_m < 0# Then GoTo HighMachCorrection_Error
If dbl_MassFlowRate < 0# Then GoTo HighMachCorrection_Error
If dbl_OutletDensity < 0# Then GoTo HighMachCorrection_Error
If dbl_PaPa < 0# Then GoTo HighMachCorrection_Error
If dbl_OutletPipeOD_m < 0# Then GoTo HighMachCorrection_Error
If dbl_OutletPipeWallThickness_m < 0# Then GoTo HighMachCorrection_Error
If dbl_ValveOutletDia_m < 0# Then GoTo HighMachCorrection_Error
If dbl_C2 < 0# Then GoTo HighMachCorrection_Error
If dbl_Lg < 0# Then GoTo HighMachCorrection_Error
If dbl_LpAe1m < 0# Then GoTo HighMachCorrection_Error
If dbl_Lpo < 0# Then GoTo HighMachCorrection_Error

'************************
'Begin calculations
'************************

' Calculate Pipe Inside Diameter (Di)

dbl_Di = dbl_OutletPipeOD_m - (2 * dbl_OutletPipeWallThickness_m)

dbl_Up = (4 * dbl_MassFlowRate) / (PI * dbl_OutletDensity * (dbl_Di ^ 2))
dbl_Ur = (dbl_Up * (dbl_Di ^ 2)) / (dbl_B * (dbl_ValveInsideDia_m ^ 2))
dbl_Wmr = ((dbl_MassFlowRate * dbl_Ur ^ 2) / 2) * ((1 - (dbl_ValveInsideDia_m ^ 2) / (dbl_Di ^ 2) + 0.2))
dbl_Fpr = (0.2 * dbl_Ur) / dbl_ValveInsideDia_m
dbl_Mr = dbl_Ur / dbl_C2
dbl_Nr = 0.001 * (dbl_Mr ^ 3)
dbl_War = dbl_Nr * dbl_Wmr
dbl_Lpir = 10 * Log10(((3200000000#) * dbl_War * dbl_OutletDensity * dbl_C2) / (dbl_Di ^ 2))

Call CalcTransmissionLoss(dbl_C2, dbl_OutletDensity, dbl_Di, _
    dbl_OutletPipeWallThickness_m, dbl_Pa, dbl_Fpr, dbl_TLr)
'dbl_TLr = -50.8

dbl_Lpe = 5 + dbl_Lpir + dbl_TLr + dbl_Lg - 10 * (Log10((dbl_Di + 2) / dbl_Di))
dbl_Lpo = 10 * (Log10((10 ^ (dbl_LpAe1m / 10)) + (10 ^ (dbl_Lpe / 10))))

HighMachCorrection_Error:
'Insert Error Message

End Sub

Public Sub CalcMultiStageMultiPathPeakFrq(dbl_P1 As Double, dbl_T1 As Double, _
    dbl_K As Double, dbl_Pvc As Double, dbl_InletDensity As Double, dbl_Dj As Double, _
    dbl_MW As Double, dbl_PeakFrequency As Double)
    
'Source: IEC 534-8-3, Eq. 9,11,12,51,52
'This subroutine calculates the peak frequency for multi-stage,
'multi-path trim.
'
'Coded by:  F.Krieger
'Date:      10/19/98
'------------------------------------------------------------------------------------------
'
'Input Variables
'dbl_P1             Inlet Pressure                              Pa a
'dbl_T1             Inlet Temperature                           deg K
'dbl_k              Ratio of specific heats
'dbl_Pvc            Absolute vena contracta pressure for
'                   subsonic flow                               Pa a
'dbl_InletDensity   Density of fluid at inlet pressure and temp kg/m3
'dbl_Dj             Calculated jet diam                         m
'dbl_MW             molecular weight

'Output Variables
'dbl_Peak Frequency Generated peak frequency                    Hz

On Error GoTo CalcMultiStageMultiPathPeakFrq_Error

'************************
'Declarations
'************************
Dim dbl_Uvc As Double
Dim dbl_Tvc As Double
Dim dbl_Cvc As Double
Dim dbl_Mjn As Double

'***************************
'Parameter validation
'***************************

If dbl_P1 < 0# Then GoTo CalcMultiStageMultiPathPeakFrq_Error
If dbl_T1 < 0# Then GoTo CalcMultiStageMultiPathPeakFrq_Error
If dbl_K < 0# Then GoTo CalcMultiStageMultiPathPeakFrq_Error
If dbl_Pvc < 0# Then GoTo CalcMultiStageMultiPathPeakFrq_Error
If dbl_InletDensity < 0# Then GoTo CalcMultiStageMultiPathPeakFrq_Error
If dbl_Dj < 0# Then GoTo CalcMultiStageMultiPathPeakFrq_Error
If dbl_MW < 0# Then GoTo CalcMultiStageMultiPathPeakFrq_Error
If dbl_PeakFrequency < 0# Then GoTo CalcMultiStageMultiPathPeakFrq_Error

dbl_Uvc = Sqr((2 * (dbl_K / (dbl_K - 1)) * (dbl_P1 / dbl_InletDensity) * ((1 - ((dbl_Pvc / dbl_P1) ^ ((dbl_K - 1) / dbl_K))))))
dbl_Tvc = dbl_T1 * ((dbl_Pvc / dbl_P1) ^ ((dbl_K - 1) / dbl_K))
dbl_Cvc = Sqr((dbl_K * 8314 * dbl_Tvc) / dbl_MW)
dbl_Mjn = dbl_Uvc / dbl_Cvc
dbl_PeakFrequency = ((2 * 10 ^ -1) * dbl_Mjn * dbl_Cvc) / dbl_Dj

CalcMultiStageMultiPathPeakFrq_Error:
' Insert Error Message

End Sub

Public Sub CalcSPLMultiStageMultiPath(dbl_Lpi As Double, dbl_TL As Double, _
    dbl_Lg As Double, dbl_Di As Double, dbl_Pn As Double, _
    dbl_InletPressure As Double, dbl_LpAe1m As Double)

'-------------------------------------------------------------------------
'Source: IEC 534-8-3, p.23
'This subroutine calculates A-weighted sound pressure level at a
'distance of 1 m from the pipe wall.
'
'Coded by:  F.Krieger
'Date:      10/19/98
'------------------------------------------------------------------------------------------
'
'Input Variables:
'dbl_Lpi                Internal sound pressure at pipe wall    dB
'dbl_Lg                 velocity correction                     dB
'dbl_TL                 Transmission loss corrected for
'                       peak frequency                          dB
'dbl_InletPressure      valve inlet pressure                    Pa abs.
'dbl_Pn                 stagnation pressure last stage          Pn
'dbl_Di                 internal pipe diameter                  m

'Output Variables:
'dbl_LpAe1m             A-weighted sound pressure level at
'                       1 m from pipe wall
'-------------------------------------------------------------------------

On Error GoTo CalcSPLMultiStageMultiPath_Error

'************************
'Declarations
'************************

Dim dbl_LpAe As Double
Dim dbl_L As Double

'************************
'Check input validation
'************************

If dbl_Lpi < 0# Then GoTo CalcSPLMultiStageMultiPath_Error
If dbl_Lg < 0# Then GoTo CalcSPLMultiStageMultiPath_Error
If dbl_Di < 0# Then GoTo CalcSPLMultiStageMultiPath_Error
If dbl_Pn < 0# Then GoTo CalcSPLMultiStageMultiPath_Error
If dbl_InletPressure < 0# Then GoTo CalcSPLMultiStageMultiPath_Error
If dbl_L < 0# Then GoTo CalcSPLMultiStageMultiPath_Error

'************************
'Begin Calculations
'************************

dbl_LpAe = 5 + dbl_Lpi + 10 * Log10(dbl_InletPressure / dbl_Pn) + dbl_TL + dbl_Lg

dbl_LpAe1m = dbl_LpAe - (10 * Log10((dbl_Di + 2) / dbl_Di))
    
CalcSPLMultiStageMultiPath_Error:
' Insert Error Message

End Sub

Public Sub CalcSPL(dbl_Lpi As Double, dbl_TL As Double, _
    dbl_Lg As Double, dbl_Di As Double, dbl_AddLoDBTrimCorrection As Double, _
    dbl_LpAe1m As Double)
'---------------------------------------------------------------------
' SOURCE: IEC 534-8-3, p.23
'
' This subroutine calculates A-weighted sound pressure level at
' a distance of 1 m from the pipe wall.
'---------------------------------------------------------------------
' Input Variables
' dbl_Lpi                   Internal sound pressure level at    dB
'                           pipe wall
' dbl_Lg                    Velocity correction                 dB
' dbl_TL                    Transmission loss corrected for     dB
'                           Peak frequency
' dbl_AddLoDBTrimCorrection                                     dB
' dbl_Di                    Internal pipe diameter              m

' Output Variables
' dbl_LpAe1m                A-weighted sound pressure level     dBA
'                           at 1 m from pipe wall
'
' Coded By:                 C.Ugbesia, F.Krieger
' Date:                     6/02/98
' Revised:                  10/19/98
'---------------------------------------------------------------------

 On Error GoTo CalcSPL_Error

'**********************************
'Check input validation
'**********************************
'dbl_Di must be greater than 0
 If dbl_Di < 0 Then GoTo CalcSPL_Error

'**********************************
'Declaration
'**********************************
 Dim dbl_LpAe As Double

'**********************************
'Begin Calculation
'**********************************

dbl_LpAe = 5 + dbl_Lpi + dbl_TL + dbl_Lg + dbl_AddLoDBTrimCorrection

dbl_LpAe1m = dbl_LpAe - (10 * Log10((dbl_Di + 2) / dbl_Di))

CalcSPL_Error:
'Placeholder

End Sub


Public Sub GetValveInsideDiameter(str_Rating As String, dbl_ValveDia As Double, _
    str_LengthUM As String, dbl_ValveInsideDia_m As Double)
'---------------------------------------------------------------------
' SOURCE: ANSI B16.34 Appendix A
'
' This subroutine gets the valve inside diameter in mm from a table.
'---------------------------------------------------------------------
' Input
' str_Rating            ANSI rating for the valve
' dbl_ValveDia           nominal valve size
' str_LengthUM        units of measure for valve size - mm or in

' Output
' dbl_ValveInsideDia_m  inside diameter of valve outlet, m
'---------------------------------------------------------------------
' completely rewritten 11/23/98 Schoonover
'---------------------------------------------------------------------
On Error GoTo GetValveInsideDiameter_Error

' Declarations
Dim str_Query As String
Dim str_QueryBase As String
Dim rs As Recordset
Dim str_DB As String
Dim dbl_ValveSizeLow As Double
Dim dbl_ValveSizeHigh As Double
Dim dbl_ValveInsideHigh As Double
Dim dbl_ValveInsideLow As Double
Dim bool_Extrapolate As Boolean

'input verification
If (str_Rating = "") Then GoTo GetValveInsideDiameter_Error
If (dbl_ValveDia <= 0) Then GoTo GetValveInsideDiameter_Error
If (str_LengthUM = "") Then GoTo GetValveInsideDiameter_Error

'process

bool_Extrapolate = False
dbl_ValveSizeHigh = dbl_ValveDia
If (str_LengthUM = "mm") Then
    If (dbl_ValveDia > 750) Then  'see if it exceeds data in table
        dbl_ValveSizeHigh = 750
        dbl_ValveSizeLow = 700
        bool_Extrapolate = True
    End If
str_QueryBase = "select [" + str_Rating + "] from tblValveInsideDia where " _
        & "VALVE_SIZE_MM = "
ElseIf (str_LengthUM = "in") Then
    If (dbl_ValveDia > 30) Then    'see if it exceeds data in table
        dbl_ValveSizeHigh = 30
        dbl_ValveSizeLow = 28
        bool_Extrapolate = True
    End If
    str_QueryBase = "select [" + str_Rating + "] from tblValveInsideDia where " _
        & "VALVE_SIZE_IN = "
End If

str_Query = str_QueryBase + CStr(dbl_ValveSizeHigh)
Set rs = g_db_Gensz.OpenRecordset(str_Query)
If (rs.RecordCount <> 1) Then GoTo GetValveInsideDiameter_Error
dbl_ValveInsideDia_m = rs.Fields(0)

If (bool_Extrapolate) Then
    'we have to extrapolate
    dbl_ValveInsideHigh = dbl_ValveInsideDia_m
    str_Query = str_QueryBase + CStr(dbl_ValveSizeLow)
    Set rs = g_db_Gensz.OpenRecordset(str_Query)
    If (rs.RecordCount <> 1) Then GoTo GetValveInsideDiameter_Error
    dbl_ValveInsideLow = rs.Fields(0)
    
    Interpolate dbl_ValveDia, dbl_ValveSizeLow, dbl_ValveSizeHigh, _
        dbl_ValveInsideLow, dbl_ValveInsideHigh, dbl_ValveInsideDia_m
End If
    rs.Close
    
Exit Sub
GetValveInsideDiameter_Error:
End Sub

Public Sub CalcAddLoDBTrimCorrection(dbl_Multiplier As Double, _
    dbl_Adder As Double, dbl_P1 As Double, dbl_P2 As Double, _
    dbl_AddLoDBTrimCorrection As Double)
'---------------------------------------------------------------------
' SOURCE: H.W.Boger document dated 3/30/94
'
' This subroutine calculates correction factor to compensate
' for multistage or special noise reduction trim sets.
'---------------------------------------------------------------------
' Input Variables
' dbl_Multiplier
' dbl_Adder
' dbl_P1
' dbl_P2
' dbl_Mo

' Output Variables
' dbl_AddLoDBtrimCorrection
'
' Coded By:             C.Ugbesia, F.Krieger
' Date:                 6/02/98
' Revised:              10/19/98
'---------------------------------------------------------------------

On Error GoTo CalcAddLoDBTrimCorrection_Error

'**********************************
'Declaration
'**********************************
  Dim dbl_Mo As Double
  Dim dbl_LodbMachCorrection As Double

'**********************************
'Check Input Validation
'**********************************
 If dbl_P1 < 0# Then GoTo CalcAddLoDBTrimCorrection_Error
 If dbl_P2 < 0# Then GoTo CalcAddLoDBTrimCorrection_Error

'**********************************
'Begin Calculation
'**********************************

 If dbl_Mo <= 0.3 Then
    dbl_LodbMachCorrection = 0
    
 ElseIf dbl_Mo > 0.3 And dbl_Mo <= 0.8 Then
    dbl_LodbMachCorrection = 30 * (dbl_Mo - 0.3)
    
 ElseIf dbl_Mo >= 0.8 Then
    dbl_LodbMachCorrection = 15
    
 End If

 dbl_AddLoDBTrimCorrection = dbl_Multiplier * ((dbl_P2 / dbl_P1) ^ 2) - dbl_Adder

CalcAddLoDBTrimCorrection_Error:
'Placeholder

End Sub


Public Sub AerodynamicNoise(str_ValveLine As String, lng_TrimTypeID As Long, lng_IECTrimStyleID As Long, _
    lng_FlowDirectionID As Long, str_ANSIRating As String, dbl_InletDensity As Double, _
    dbl_FL As Double, dbl_PaPa As Double, dbl_P1Pa As Double, dbl_P2Pa As Double, _
    dbl_MassFlowRate As Double, dbl_T1 As Double, dbl_K As Double, _
    dbl_MW As Double, dbl_OutletPipeOD_m As Double, dbl_OutletPipeWallThickness_m As Double, _
    dbl_CalculatedCv As Double, dbl_RatedCv As Double, _
    dbl_ValveDia As Double, dbl_ValveInletDia_m As Double, dbl_ValveOutletDia_m As Double, _
    dbl_InletPipeDia_m As Double, dbl_OutletPipeDia_m As Double, str_LengthUM As String, _
    dbl_M2 As Double, dbl_Mo As Double, dbl_SL As Double)
'---------------------------------------------------------------------
' SOURCE: IEC 543-8-3, edition 2 draft, 1995-08, part 8, Control valve aerodynamic noise.
'
' Noise prediction for valve with no restrictor downstream
' Revised:                     10/20/98
' Tested:                      11/17/98
'---------------------------------------------------------------------
'Input Variables                Description
'---------------------------------------------------------------------
'str_ValveLine                  Valve model number
'lng_IECTrimStyleID             Valve trim style
'dbl_InletDensity               Fluid density at inlet conditions, kg/m3
'dbl_PaPa                       Atmospheric pressure, Pa
'dbl_P1Pa                       Inlet pressure to the control valve, Pa
'dbl_P2Pa                       System outlet pressure, Pa
'dbl_MassFlowRate               Mass flow rate, kg/s
'dbl_T1                         Inlet temperature, Kelvin
'dbl_k                          Specific heat ratio
'dbl_MW                         Molecular weight
'dbl_OutletPipeOD_m             Outside diameter of pipe, m
'dbl_OutletPipeWallThickness_m  Pipe wall thickness, m
'dbl_Cv                         Calculated Cv of the control valve
'dbl_RatedCv                    Rated Cv of the control valve
'dbl_ValveDia                   Nominal valve diameter
'dbl_ValveInletDia_m            Size of inlet pipe, m
'dbl_ValveOutletDia_m           Size of outlet pipe, m
'dbl_InletPipeDia_m            Size of valve inlet, m
'dbl_OutletPipeDia_m           Size of vale outlet, m
'
'Output Variables               Description
'---------------------------------------------------------------------
'dbl_Mo                         Outlet Mach no at outlet of the valve
'dbl_M2                         Outlet mach number in downstream pipe
'dbl_SL                         Noise prediction for restrictor
'---------------------------------------------------------------------

On Error GoTo AerodynamicNoise_Error
 
'***************************
'Parameter validation
'***************************
If dbl_InletDensity < 0# Then GoTo AerodynamicNoise_Error
If dbl_PaPa < 0# Then GoTo AerodynamicNoise_Error
If dbl_P1Pa < 0# Then GoTo AerodynamicNoise_Error
If dbl_P2Pa < 0# Then GoTo AerodynamicNoise_Error
If dbl_MassFlowRate < 0# Then GoTo AerodynamicNoise_Error
If dbl_T1 < 0# Then GoTo AerodynamicNoise_Error
If dbl_K < 0# Then GoTo AerodynamicNoise_Error
If dbl_MW < 0# Then GoTo AerodynamicNoise_Error
If dbl_OutletPipeOD_m < 0# Then GoTo AerodynamicNoise_Error
If dbl_OutletPipeWallThickness_m < 0# Then GoTo AerodynamicNoise_Error
If dbl_CalculatedCv < 0# Then GoTo AerodynamicNoise_Error
If dbl_RatedCv < 0# Then GoTo AerodynamicNoise_Error
If dbl_ValveDia < 0# Then GoTo AerodynamicNoise_Error
If dbl_ValveInletDia_m < 0# Then GoTo AerodynamicNoise_Error
If dbl_ValveOutletDia_m < 0# Then GoTo AerodynamicNoise_Error
If dbl_InletPipeDia_m < 0# Then GoTo AerodynamicNoise_Error
If dbl_OutletPipeDia_m < 0# Then GoTo AerodynamicNoise_Error

'**********************************
'Declaration
'**********************************

Dim bln_ReducedArea As Boolean
Dim dbl_Adder As Double
Dim dbl_AddLoDBTrimCorrection As Double
Dim dbl_Alpha As Double
Dim dbl_C2 As Double
Dim dbl_Cn As Double
Dim dbl_Cv As Double
Dim dbl_Cvcc As Double
Dim dbl_Di As Double
Dim dbl_Dj As Double
Dim dbl_Fd As Double
Dim dbl_Flp As Double
Dim dbl_Fp As Double
Dim dbl_HolesPerCv As Double
Dim dbl_InletDensity1 As Double
Dim dbl_InletPressure As Double
Dim dbl_Ki As Double
Dim dbl_Ksum As Double
Dim dbl_LDRatio As Double
Dim dbl_Lg As Double
Dim dbl_LpAe1m As Double
Dim dbl_Lpi As Double
Dim dbl_Lpo As Double
Dim dbl_Mj As Double
Dim dbl_ModelNumber As Double
Dim dbl_Multiplier As Double
Dim dbl_OutletDensity As Double
Dim dbl_OutletDensityPn As Double
Dim dbl_P1 As Double
Dim dbl_P1Pa1 As Double
Dim dbl_P2 As Double
Dim dbl_P2b As Double
Dim dbl_P2c As Double
Dim dbl_P2ce As Double
Dim dbl_Pa As Double
Dim dbl_PeakFrequency As Double
Dim dbl_PipeWallThickness As Double
Dim dbl_Pn As Double
Dim dbl_Ps As Double
Dim dbl_Pvc As Double
Dim dbl_Pvcc As Double
Dim dbl_TL As Double
Dim dbl_Tvcc As Double
Dim dbl_ValveInletDia_mm As Double
Dim dbl_ValveInsideDia_m As Double
Dim dbl_ValveSize As Double
Dim dbl_Wa As Double
Dim dbl_Wms As Double
Dim str_DensityUM As String
Dim str_PressureUM As String
Dim str_TemperatureUM As String
Dim str_ValveModel As String

'**********************************
'Begin Calculation
'**********************************
If str_ValveLine = "" Then
    lng_IECTrimStyleID = 1
    'default stuff
    dbl_ValveInletDia_m = dbl_InletPipeDia_m
    dbl_ValveOutletDia_m = dbl_OutletPipeDia_m
Else
    lng_IECTrimStyleID = GetIECTrimType(lng_TrimTypeID)
End If
'--------------------------------------------------------------------------

'*************************************************
'Converting str_IECTrimStyle to lng_IECTrimStyleID
'Single Stage - Single Path = 1  (Use std trim method)
'Single Stage - Multi Path  = 2  (Use lodb Dj calc)
'Multi Stage - Single Path  = 3  (Use Fd of 21000 std trim, calc Lpa, 77000 a or b, use DSL)
'Multi Stage - Multi Path   = 4  (Use lodb Dj, Cn and Pn)
'*************************************************
   
'Calculate Pipe Inside Diameter
 dbl_Di = dbl_OutletPipeOD_m - 2 * dbl_OutletPipeWallThickness_m

 Call PipeGeom(dbl_InletPipeDia_m, dbl_OutletPipeDia_m, _
    dbl_ValveInletDia_m, dbl_ValveOutletDia_m, dbl_Ksum, dbl_Ki)

 dbl_ValveInletDia_mm = dbl_ValveInletDia_m * 1000
 dbl_InletDensity1 = dbl_InletDensity
 dbl_P1Pa1 = dbl_P1Pa
 
 Call CalcFp(dbl_ValveInletDia_mm, dbl_Ksum, dbl_Ki, dbl_CalculatedCv, _
    dbl_FL, dbl_Fp, dbl_Flp)
  
 If str_ValveLine = "" Then
    dbl_Fd = 0.28
 Else
    If lng_IECTrimStyleID = 1 Or lng_IECTrimStyleID = 3 Then

        Call GetStdValveDataForNoise(str_ValveLine, lng_FlowDirectionID, _
            lng_TrimTypeID, dbl_CalculatedCv, dbl_RatedCv, dbl_Fd)
        
    ElseIf lng_IECTrimStyleID = 2 Or lng_IECTrimStyleID = 4 Then
    
        Call GetLodbValveDataForNoise(str_ValveLine, lng_TrimTypeID, _
             dbl_HolesPerCv, dbl_LDRatio)
    
    End If
 
 End If
 If lng_IECTrimStyleID = 1 Or lng_IECTrimStyleID = 3 Then
    Call CalcStdJetDia(dbl_Fd, dbl_CalculatedCv, dbl_Flp, dbl_Fp, dbl_Dj)
 ElseIf lng_IECTrimStyleID = 2 Then
    Call CalcLodbJetDia(dbl_CalculatedCv, dbl_HolesPerCv, dbl_LDRatio, _
        dbl_Dj)
 ElseIf lng_IECTrimStyleID = 4 Then
    Call CalcCn(str_ValveLine, dbl_CalculatedCv, dbl_RatedCv, _
        dbl_P1Pa, dbl_P2Pa, dbl_Cn)
    Call CalcLodbJetDia(dbl_Cn, dbl_HolesPerCv, dbl_LDRatio, dbl_Dj)
    Call CalcPn(dbl_CalculatedCv, dbl_Cn, dbl_P1Pa, dbl_P2Pa, dbl_Pn)
    Call CalcOutletDensity(dbl_InletDensity, dbl_P1Pa, dbl_Pn, _
        dbl_OutletDensityPn)
        dbl_InletDensity1 = dbl_OutletDensityPn
        dbl_InletPressure = dbl_P1Pa
        dbl_P1Pa1 = dbl_Pn
 End If
 
 If lng_IECTrimStyleID = 4 Then

    Call CalcPressureRatios(dbl_P1Pa1, dbl_P2Pa, dbl_Flp, dbl_Fp, dbl_K, _
        dbl_Pvc, dbl_Pvcc, dbl_P2c, dbl_Alpha, dbl_P2b, dbl_P2ce)
        dbl_Pvc = dbl_P2Pa
 Else
    Call CalcPressureRatios(dbl_P1Pa1, dbl_P2Pa, dbl_Flp, dbl_Fp, dbl_K, _
        dbl_Pvc, dbl_Pvcc, dbl_P2c, dbl_Alpha, dbl_P2b, dbl_P2ce)
 End If

 
 If dbl_P2Pa >= dbl_P2c Then                                    '(Regime I)

    Call CalcRegime1(dbl_P1Pa1, dbl_T1, dbl_K, dbl_Pvc, dbl_InletDensity1, _
        dbl_MassFlowRate, dbl_Dj, dbl_MW, dbl_Flp, dbl_Fp, dbl_Wa, dbl_PeakFrequency)
    
 ElseIf (dbl_P2c > dbl_P2Pa) And (dbl_P2Pa >= dbl_Pvcc) Then                     '(Regime II)

    Call CalcCommonRegime2thru4(dbl_P1Pa1, dbl_P2Pa, dbl_T1, dbl_K, dbl_MassFlowRate, _
        dbl_MW, dbl_Alpha, dbl_Tvcc, dbl_Cvcc, dbl_Wms, dbl_Mj)
    Call CalcRegime2(dbl_P1Pa1, dbl_P2Pa, dbl_Pvcc, dbl_Dj, dbl_Cvcc, _
        dbl_Wms, dbl_Mj, dbl_Flp, dbl_Fp, dbl_Wa, dbl_PeakFrequency)

 ElseIf (dbl_Pvcc > dbl_P2Pa) And (dbl_P2Pa >= dbl_P2b) Then                     '(Regime III)

    Call CalcCommonRegime2thru4(dbl_P1Pa1, dbl_P2Pa, dbl_T1, dbl_K, dbl_MassFlowRate, _
        dbl_MW, dbl_Alpha, dbl_Tvcc, dbl_Cvcc, dbl_Wms, dbl_Mj)
    Call CalcRegime3(dbl_Mj, dbl_Flp, dbl_Fp, dbl_Wms, dbl_Dj, dbl_Cvcc, _
        dbl_Wa, dbl_PeakFrequency)
    
 ElseIf (dbl_P2b > dbl_P2Pa) And (dbl_P2Pa >= dbl_P2ce) Then                     '(Regime IV)

    Call CalcCommonRegime2thru4(dbl_P1Pa1, dbl_P2Pa, dbl_T1, dbl_K, dbl_MassFlowRate, _
        dbl_MW, dbl_Alpha, dbl_Tvcc, dbl_Cvcc, dbl_Wms, dbl_Mj)
    Call CalcRegime4(dbl_Mj, dbl_Flp, dbl_Fp, dbl_Wms, dbl_Dj, dbl_Cvcc, _
        dbl_Wa, dbl_PeakFrequency)
    
 ElseIf dbl_P2ce > dbl_P2Pa Then                                ' (Regime V)

    Call CalcRegime5(dbl_T1, dbl_Flp, dbl_Fp, dbl_Dj, dbl_K, dbl_MW, _
        dbl_MassFlowRate, dbl_Wa, dbl_PeakFrequency)

 End If
'Assume T2 approximately equals T1, use T1 in outlet sonic velocity calculation

 Call CalcOutletSonicVelocity(dbl_K, dbl_T1, dbl_MW, dbl_C2)
 Call CalcOutletDensity(dbl_InletDensity1, dbl_P1Pa1, dbl_P2Pa, dbl_OutletDensity)
     
 'If (str_ValveLine <> "") Then
    Call CalcMachNo(dbl_MassFlowRate, dbl_ValveOutletDia_m, _
       dbl_OutletDensity, dbl_C2, dbl_Mo)
    Call CalcMachNo(dbl_MassFlowRate, dbl_Di, _
       dbl_OutletDensity, dbl_C2, dbl_M2)
 'Else
 '   dbl_Mo = 0
 '   dbl_M2 = 0
 'End If

 Call CalcInternalSPL(dbl_Wa, dbl_OutletDensity, dbl_C2, dbl_Di, dbl_Lpi)
 
 If lng_IECTrimStyleID = 4 Then
    Call CalcMultiStageMultiPathPeakFrq(dbl_P1Pa1, dbl_T1, _
        dbl_K, dbl_Pvc, dbl_InletDensity, dbl_Dj, dbl_MW, dbl_PeakFrequency)
 End If
 Call CalcTransmissionLoss(dbl_C2, dbl_OutletDensity, dbl_Di, _
    dbl_OutletPipeWallThickness_m, dbl_PaPa, dbl_PeakFrequency, dbl_TL)
 Call CalcPipeVelocityCorrection(dbl_MassFlowRate, dbl_Di, _
    dbl_OutletDensity, dbl_C2, dbl_Lg)
 
 If str_ValveLine = "" Then
    dbl_Multiplier = 0
    dbl_Adder = 0
 Else
    Call GetAddLoDBTrimCorrectionData(str_ValveLine, lng_TrimTypeID, _
        dbl_ValveSize, dbl_Multiplier, dbl_Adder)
 End If
 
 Call CalcAddLoDBTrimCorrection(dbl_Multiplier, dbl_Adder, dbl_P1Pa1, _
    dbl_P2Pa, dbl_AddLoDBTrimCorrection)
 
 If lng_IECTrimStyleID = 4 Then
    Call CalcSPLMultiStageMultiPath(dbl_Lpi, dbl_TL, dbl_Lg, dbl_Di, _
       dbl_Pn, dbl_InletPressure, dbl_LpAe1m)
 Else
    Call CalcSPL(dbl_Lpi, dbl_TL, dbl_Lg, dbl_Di, _
        dbl_AddLoDBTrimCorrection, dbl_LpAe1m)
 End If
     
 If dbl_Mo > 0.3 And str_ValveLine <> "" Then
    If (str_ValveLine = "") Then
        Call GetValveInsideDiameter(str_ANSIRating, dbl_InletPipeDia_m * 1000, _
            "mm", dbl_ValveInsideDia_m)
    Else
        'valve diameter is always in mm
        Call GetValveInsideDiameter(str_ANSIRating, dbl_ValveDia, _
            "mm", dbl_ValveInsideDia_m)
    End If
    Call HighMachCorrection(str_ValveLine, lng_TrimTypeID, _
        dbl_ValveInsideDia_m, dbl_MassFlowRate, dbl_OutletDensity, _
        dbl_PaPa, dbl_OutletPipeOD_m, dbl_OutletPipeWallThickness_m, _
        dbl_ValveOutletDia_m, dbl_C2, dbl_Lg, dbl_LpAe1m, dbl_Lpo)
    dbl_SL = dbl_Lpo
 Else
    dbl_SL = dbl_LpAe1m
 End If

AerodynamicNoise_Error:
'Placeholder

End Sub


Public Sub CalcPressureRatios(dbl_P1 As Double, dbl_P2 As Double, _
dbl_Flp As Double, dbl_Fp As Double, dbl_K As Double, dbl_Pvc As Double, _
dbl_Pvcc As Double, dbl_P2c As Double, dbl_Alpha As Double, _
dbl_P2b As Double, dbl_P2ce As Double)
'---------------------------------------------------------------------
' SOURCE: IEC 534-8-3, Eqs.1,2,3,4,5,6
'
' This subroutine calculates the various pressures and pressure ratios
' used by the gas noise calculation subroutines.
'---------------------------------------------------------------------
'Input Variables
'dbl_P1                Valve inlet pressure                    Pa a
'dbl_P2                Valve outlet pressure                   Pa a
'dbl_Flp               Liquid pressure recovery factor
'                       with attached pipe fittings
'dbl_Fp                Piping geometry factor
'dbl_k                 Ratio of specific heats for selected gas

'Output Variables
'dbl_Pvc               Absolute vena contracta pressure
'                       for subsonic flow                       Pa a
'dbl_Pvcc              Absolute vena contracta pressure
'                       for critical flow                       Pa a
'dbl_P2c               Absolute outlet pressure
'                       at critical flow conditions             Pa a
'dbl_Alpha             Recovery correction factor
'dbl_P2b               Valve outlet absolute pressure
'                       at break point                          Pa a
'dbl_P2ce              Valve outlet absolute pressure where
'                       const. accoustical eff. begins          Pa a
'
'Coded By:             C.Ugbesia, F.Krieger
'Date:                 6/02/98
'Revised:              10/20/98
'---------------------------------------------------------------------
 On Error GoTo CalcPressureRatios_Error

'**********************************
'Check input validation
'**********************************
'dbl_Fp, dbl_K, dbl_P1,dbl_Alpha must be greater than 0
 If dbl_P1 < 0# Then GoTo CalcPressureRatios_Error
 If dbl_P2 < 0# Then GoTo CalcPressureRatios_Error
 If dbl_Flp < 0# Then GoTo CalcPressureRatios_Error
 If dbl_Fp < 0# Then GoTo CalcPressureRatios_Error
 If dbl_K < 0# Then GoTo CalcPressureRatios_Error
 If dbl_Alpha < 0# Then GoTo CalcPressureRatios_Error
 
'**********************************
'Declaration
'**********************************

'**********************************
'Begin Calculation
'**********************************
        
   dbl_Pvc = dbl_P1 - (dbl_P1 - dbl_P2) / ((dbl_Flp / dbl_Fp) ^ 2)
  
 If dbl_Pvc < 0 Then dbl_Pvc = dbl_P2
   
   dbl_Pvcc = dbl_P1 * (2 / (dbl_K + 1)) ^ (dbl_K / (dbl_K - 1))
 
   dbl_P2c = dbl_P1 - ((dbl_Flp / dbl_Fp) ^ 2) * (dbl_P1 - dbl_Pvcc)
    
   dbl_Alpha = dbl_Pvcc / dbl_P2c
    
   dbl_P2b = (dbl_P1 / dbl_Alpha) * (1 / dbl_K) ^ (dbl_K / (dbl_K - 1))
    
   dbl_P2ce = dbl_P1 / (22 * dbl_Alpha)
        
'End If
    
CalcPressureRatios_Error:
'Placeholder
               
End Sub

Public Sub CalcStdJetDia(dbl_Fd As Double, dbl_CalculatedCv As Double, _
    dbl_Flp As Double, dbl_Fp As Double, dbl_Dj As Double)
'---------------------------------------------------------------------
' SOURCE: IEC 534-8-3, p 21
'
' This subroutine calculates the jet diameter for standard trim used by
' the gas noise calculation subroutines.
'---------------------------------------------------------------------
' Input Variables
' dbl_Fd                Valve style modifier
' dbl_Calculated Cv     Calculated Cv for specified service conditions
' dbl_FLp               Combined pressure recovery factor and piping
'                       geometry factors
' dbl_Fp                Piping geometry factor

' Output Variables
' dbl_Dj                Internal jet diameter                   m
'
' Coded By:             C.Ugbesia, F.Krieger
' Date:                 6/02/98
' Revised:              10/20/98
'---------------------------------------------------------------------
 
  On Error GoTo CalcStdJetDia_Error
 
'**********************************
'Check input validation
'**********************************
'dbl_Fp must be greater than 0
 If dbl_Fd < 0# Then GoTo CalcStdJetDia_Error
 If dbl_CalculatedCv < 0# Then GoTo CalcStdJetDia_Error
 If dbl_Flp < 0# Then GoTo CalcStdJetDia_Error
 If dbl_Fp < 0# Then GoTo CalcStdJetDia_Error
 
'**********************************
'Declaration
'**********************************

'**********************************
'Begin Calculation
'**********************************
     
 dbl_Dj = (0.0046 * dbl_Fd) * Sqr(dbl_CalculatedCv * (dbl_Flp / dbl_Fp))
    
'End If
  
CalcStdJetDia_Error:
'Placeholder
    
End Sub

Public Sub CalcLodbJetDia(dbl_CalculatedCv As Double, _
    dbl_HolesPerCv As Double, dbl_LDRatio As Double, dbl_Dj As Double)
'---------------------------------------------------------------------
' SOURCE: IEC 534-8-3, p.21, eq.8
'
' This subroutine calculates the jet diameter for lodb trim used by the
' gas noise calculation subroutines.
'---------------------------------------------------------------------
' Input Variables
' dbl_Calculated Cv     Calculated Cv for specified service conditions
' dbl_HolesPerCv        Approx. number of holes per Cv
' dbl_LDratio           Length/dia ratio
'                       Piping geometry factor
' Output Variables
' dbl_Dj                Internal jet diameter                   m
'
' Coded By:             C.Ugbesia, F.Krieger
' Date:                 6/02/98
' Revised:              10/20/98
'---------------------------------------------------------------------

 On Error GoTo CalcLodbJetDia_Error

'**********************************
'Check input validation
'**********************************
'dbl_CalculatedCv and dbl_HolesPerCv must be greater than 0
 If dbl_CalculatedCv < 0# Then GoTo CalcLodbJetDia_Error
 If dbl_HolesPerCv < 0# Then GoTo CalcLodbJetDia_Error
 If dbl_LDRatio < 0# Then GoTo CalcLodbJetDia_Error

'**********************************
'Declaration
'**********************************

 Dim dbl_Fd As Double
 Dim dbl_Temp As Double
 
'**********************************
'Begin Calculation
'**********************************

   dbl_Fd = 1 / Sqr(dbl_CalculatedCv * dbl_HolesPerCv)
    
   dbl_Temp = (0.9 - (0.06 * dbl_LDRatio))
    
   If dbl_Temp < 0.66 Then
      dbl_Temp = 0.66
   End If
    
   dbl_Dj = 0.0046 * dbl_Fd * Sqr(dbl_CalculatedCv * dbl_Temp)
    
  'End If
  
CalcLodbJetDia_Error:
'Placeholder
    
End Sub


Public Sub CalcRegime1(dbl_P1 As Double, dbl_T1 As Double, dbl_K As Double, _
    dbl_Pvc As Double, dbl_InletDensity As Double, dbl_MassFlowRate As Double, _
    dbl_Dj As Double, dbl_MW As Double, dbl_Flp As Double, dbl_Fp As Double, _
    dbl_Wa As Double, dbl_PeakFrequency As Double)
'---------------------------------------------------------------------
' SOURCE: IEC 534-8-3, p 23-25, eq.9,10,11,12,13,14,15,17
'
'This subroutine calculates the intermediate variables used for
'Regime I (subsonic flow)
'---------------------------------------------------------------------
'Input Variables
'dbl_P1                Inlet pressure                          Pa a
'dbl_T1                Inlet temperature                       deg K
'dbl_K                 Ratio of specific heats
'dbl_Pvc               Absolute vena contracta pressure
'                       for subsonic flow                       Pa a
'dbl_InletDensity      Density of fluid at inlet pressure      kg/m3
'                       and temperature
'dbl_MassFlowRate      flow rate                               k/s
'dbl_Dj                Calculated internal pipe diameter       m
'dbl_MW                Molecular weight
'dbl_Flp               Liquid pressure recovery factor with
'                       attached pipe fittings
'dbl_Fp                piping geometry factor
'
'Output Variables
'dbl_Wa                Sound power                             W
'dbl_PeakFrequency     Generated peak frequency                Hz
'
'Coded By:             C.Ugbesia
'Date:                 6/02/98
'Revised:              10/20/98
'---------------------------------------------------------------------
 
 On Error GoTo CalcRegime1_Error

'**********************************
'Declaration
'**********************************
 Dim dbl_Cvc As Double
 Dim dbl_Eta1 As Double
 Dim dbl_Mvc As Double
 Dim dbl_Tvc As Double
 Dim dbl_Uvc As Double
 Dim dbl_Wm As Double
 Dim dbl_Wms As Double
 
'**********************************
'Check input validation
'**********************************
'dbl_P1, dbl_K, dbl_InletDensity must be greater Than 0
 If dbl_P1 < 0# Then GoTo CalcRegime1_Error
 If dbl_T1 < 0# Then GoTo CalcRegime1_Error
 If dbl_K < 0# Then GoTo CalcRegime1_Error
 If dbl_Pvc < 0# Then GoTo CalcRegime1_Error
 If dbl_InletDensity < 0# Then GoTo CalcRegime1_Error
 If dbl_MassFlowRate < 0# Then GoTo CalcRegime1_Error
 If dbl_Dj < 0# Then GoTo CalcRegime1_Error
 If dbl_MW < 0# Then GoTo CalcRegime1_Error
 If dbl_Flp < 0# Then GoTo CalcRegime1_Error
 If dbl_Fp < 0# Then GoTo CalcRegime1_Error

'**********************************
'Begin Calculation
'**********************************
 dbl_Uvc = Sqr((2 * (dbl_K / (dbl_K - 1)) * (dbl_P1 / dbl_InletDensity) * ((1 - ((dbl_Pvc / dbl_P1) ^ ((dbl_K - 1) / dbl_K))))))

 dbl_Wm = (dbl_MassFlowRate * (dbl_Uvc ^ 2)) / 2

 dbl_Tvc = dbl_T1 * (dbl_Pvc / dbl_P1) ^ ((dbl_K - 1) / dbl_K)

 dbl_Cvc = Sqr((dbl_K * 8314 * dbl_Tvc) / dbl_MW)

 dbl_Mvc = dbl_Uvc / dbl_Cvc

 dbl_Eta1 = 0.0001 * (dbl_Mvc ^ 3.6)

 dbl_Wa = dbl_Eta1 * dbl_Wm * ((dbl_Flp / dbl_Fp) ^ 2)

 dbl_PeakFrequency = (0.2 * dbl_Uvc) / dbl_Dj

CalcRegime1_Error:
'Placeholder

End Sub


Public Sub CalcRegime2(dbl_P1 As Double, dbl_P2 As Double, _
    dbl_Pvcc As Double, dbl_Dj As Double, dbl_Cvcc As Double, _
    dbl_Wms As Double, dbl_Mj As Double, dbl_Flp As Double, _
    dbl_Fp As Double, dbl_Wa As Double, dbl_PeakFrequency As Double)
'--------------------------------------------------------------------------------------------------------------------------
'SOURCE: IEC 534-8-3, p 27, eq.22,23,24
'
'This subroutine calculates the intermediate variables used for
'Regime II.
'---------------------------------------------------------------------
'Input Variables:
'dbl_P1                Inlet pressure                          Pa a
'dbl_P2                Outlet pressure                         Pa a
'dbl_Pvcc
'dbl_Dj                Calculated jet dia                      m
'dbl_Cvcc              Speed of sound at vena contracta
'                       at critical flow conditions             m/s
' dbl_Wms               Stream power of mass flow rate
'                       at sonic velocity   W
' dbl_Mj                Freely expanded jet Mach number
'                       in Regimes II-IV
' dbl_FLp
' dbl_Fp

' Output Variables
' dbl_Wa                Sound power                             W
' dbl_PeakFrequency     Generated peak frequency                Hz
'
' Coded By:             C.Ugbesia, F.Krieger
' Date:                 6/02/98
' Revised:              10/20/98
'---------------------------------------------------------------------
 On Error GoTo CalcRegime2_Error

'**********************************
'Check input validation
'**********************************
'dbl_P1 must be greater than dbl_P2 and dbl_Pvcc
 If dbl_P1 <= dbl_P2 Or dbl_P1 <= dbl_Pvcc Then GoTo CalcRegime2_Error
 'dbl_Dj and dbl_Fp can't be 0
 If dbl_P1 < 0# Then GoTo CalcRegime2_Error
 If dbl_P2 < 0# Then GoTo CalcRegime2_Error
 If dbl_Pvcc < 0# Then GoTo CalcRegime2_Error
 If dbl_Dj < 0# Then GoTo CalcRegime2_Error
 If dbl_Cvcc < 0# Then GoTo CalcRegime2_Error
 If dbl_Wms < 0# Then GoTo CalcRegime2_Error
 If dbl_Mj < 0# Then GoTo CalcRegime2_Error
 If dbl_Flp < 0# Then GoTo CalcRegime2_Error
 If dbl_Fp < 0# Then GoTo CalcRegime2_Error
 
'**********************************
'Declaration
'**********************************
 
 Dim dbl_Eta2 As Double

'**********************************
'Begin Calculation
'**********************************

 dbl_Eta2 = (0.0001) * (dbl_Mj ^ (6.6 * (dbl_Flp / dbl_Fp) ^ 2))

 dbl_Wa = dbl_Eta2 * dbl_Wms * ((dbl_P1 - dbl_P2) / (dbl_P1 - dbl_Pvcc))
 
 dbl_PeakFrequency = (0.2 * dbl_Mj * dbl_Cvcc) / dbl_Dj


CalcRegime2_Error:
'Placeholder

End Sub


Public Sub CalcCommonRegime2thru4(dbl_P1 As Double, dbl_P2 As Double, _
    dbl_T1 As Double, dbl_K As Double, dbl_MassFlowRate As Double, _
    dbl_MW As Double, dbl_Alpha As Double, dbl_Tvcc As Double, _
    dbl_Cvcc As Double, dbl_Wms As Double, dbl_Mj As Double)
'--------------------------------------------------------------------------------------------------------------------
' SOURCE: IEC 534-8-3, p 25-27, eq.18,19,20,21
'
' This subroutine calculates the intermediate variables used for
' Regime 2 thru 4
'
'---------------------------------------------------------------------
' Input Variables
' dbl_P1                Inlet Pressure                          Pa a
' dbl_P2                Outlet Pressure                         Pa a
' dbl_T1                Inlet temperature                       deg K
' dbl_k                 Ratio of specific heats
' dbl_MassFlowRate      Flow rate                               k/s
' dbl_MW                Molecular weight
' dbl_Alpha             Recovery correction factor

' Output Variables
' dbl_Tvcc              Absolute temperature at vena contracta
'                       at critical flow conditions             deg K
' dbl_Cvcc              Speed of sound at vena contracta
'                       at critical flow conditions             m/s
' dbl_Wms               Stream power of mass flow rate          W
'                       at sonic velocity
' dbl_Mj                Freely expanded jet Mach number
'                       in Regimes II-IV
'
' Coded By:             C.Ugbesia, F.Krieger
' Date:                 6/02/98
' Revised:              10/20/98
'---------------------------------------------------------------------

 On Error GoTo CalcCommonRegime2thru4_Error

'**********************************
