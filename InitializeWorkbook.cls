VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InitializeWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Operation")

Sub Main()
    ' Initialize the necessary worksheets
    InitializeDataJoint "Data_Joint"
    InitializeDataFrame "Data_Frame"
    InitializeDataJointConn "Data_JointConn"
    InitializeDataMemberList "Data_MemberList"
    InitializeDataFrameForce "Data_FrameForce"
    InitializeWSData "WS_Data"
    InitializeDefineConnectionType "Define_ConnectionType"
    InitializeDataSystem "Data_System"
End Sub

Sub InitializeDataJoint(sheetName As String)
    Dim ws As Worksheet
    Dim WB As Workbook
    
    ' Reference the current workbook
    Set WB = ThisWorkbook
    
    ' Check if the worksheet exists, if not, create it
    If Not WorksheetExists(sheetName) Then
        Set ws = WB.Worksheets.Add
        ws.name = sheetName
    Else
        Set ws = WB.Worksheets(sheetName)
    End If
    
    ' Initialize the worksheet headers assuming it's "Data_Joint"
    ' Adjust the following if different setups are needed for each sheet
    With ws
        If sheetName = "Data_Joint" Then
            .Cells(1, 1).value = "rRow"
            .Cells(2, 2).value = "Node ID"
            .Cells(2, 3).value = "X"
            .Cells(2, 4).value = "Y"
            .Cells(2, 5).value = "Z"
            .Cells(2, 6).value = "isRestraint"
        End If
    End With
End Sub

Sub InitializeDataJointConn(sheetName As String)
    Dim ws As Worksheet
    Dim WB As Workbook
    
    ' Reference the current workbook
    Set WB = ThisWorkbook
    
    ' Check if the worksheet exists, if not, create it
    If Not WorksheetExists(sheetName) Then
        Set ws = WB.Worksheets.Add
        ws.name = sheetName
    Else
        Set ws = WB.Worksheets(sheetName)
    End If
    
    ' Set up the headers in "Data_JointConn"
    With ws
        .Cells(1, 1).value = "rRow"
        .Cells(2, 1).value = "Node ID"
        .Cells(1, 2).value = "jtID"
        .Cells(2, 2).value = "Node ID"
        .Cells(1, 3).value = "memID"
        .Cells(2, 3).value = "Continuous Member"
        .Cells(1, 4).value = "eleID"
        .Cells(2, 4).value = "Corresponding Elements"
        .Cells(1, 5).value = "sections"
        .Cells(2, 5).value = "Corresponding Sections"
        .Cells(1, 6).value = "isRestraint"
        .Cells(2, 6).value = "Restraint"
        .Cells(1, 7).value = "isConn"
        .Cells(2, 7).value = "Is Possible Connection?"
        .Cells(1, 8).value = "matchedType"
        .Cells(2, 8).value = "Matched Joint Type"
    End With
End Sub

Sub InitializeDataFrame(sheetName As String)
    Dim ws As Worksheet
    Dim WB As Workbook
    
    ' Reference the current workbook
    Set WB = ThisWorkbook
    
    ' Check if the worksheet exists, if not, create it
    If Not WorksheetExists(sheetName) Then
        Set ws = WB.Worksheets.Add
        ws.name = sheetName
    Else
        Set ws = WB.Worksheets(sheetName)
    End If
    
    ' Set up the headers in "Data_Frame"
    With ws
        .Cells(1, 2).value = "eleID"
        .Cells(1, 3).value = "section"
        .Cells(3, 2).value = "Frame"
        .Cells(3, 3).value = "AnalSect"
        .Cells(3, 4).value = "JointI"
        .Cells(3, 5).value = "JointJ"
        .Cells(3, 6).value = "Length"
        .Cells(3, 7).value = "MemID"
        .Cells(3, 8).value = "Link? Ele?"
        .Cells(3, 9).value = "Orientation"
        .Cells(3, 10).value = "Vertical/ Horizontal/ Diagonal"
        .Cells(2, 2).value = "TABLE: Frame Section Assignments"
        .Cells(4, 1).value = "rRow"
    End With
End Sub

Sub InitializeDataFrameForce(sheetName As String)
    Dim ws As Worksheet
    Dim WB As Workbook
    
    ' Reference the current workbook
    Set WB = ThisWorkbook
    
    ' Check if the worksheet exists, if not, create it
    If Not WorksheetExists(sheetName) Then
        Set ws = WB.Worksheets.Add
        ws.name = sheetName
    Else
        Set ws = WB.Worksheets(sheetName)
    End If
    
    ' Set up the headers in "Data_FrameForce"
    With ws
        .Cells(1, 3).value = "eleID"
        .Cells(2, 3).value = "TABLE: Element"
        .Cells(3, 3).value = "Frame"
        .Cells(1, 4).value = "station"
        .Cells(3, 4).value = "Station"
        .Cells(1, 5).value = "loadComb"
        .Cells(3, 5).value = "OutputCase"
        .Cells(1, 6).value = "stepType"
        .Cells(3, 6).value = "StepType"
        .Cells(1, 7).value = "P"
        .Cells(3, 7).value = "P"
        .Cells(1, 8).value = "V2"
        .Cells(3, 8).value = "V2"
        .Cells(1, 9).value = "V3"
        .Cells(3, 9).value = "V3"
        .Cells(1, 10).value = "T"
        .Cells(3, 10).value = "T"
        .Cells(1, 11).value = "M2"
        .Cells(3, 11).value = "M2"
        .Cells(1, 12).value = "M3"
        .Cells(3, 12).value = "M3"
        .Cells(1, 13).value = "section"
        .Cells(3, 13).value = "AnalSect"
        .Cells(1, 14).value = "memID"
        .Cells(3, 14).value = "MemID"
        .Cells(1, 15).value = "subEleName"
        .Cells(3, 15).value = "Element"
        .Cells(4, 1).value = "rRow"

        ' Write formulas
        .Cells(4, 13).formula = "=XLOOKUP(C4,Data_Frame!$B:$B,Data_Frame!$C:$C)"
        .Cells(4, 14).formula = "=XLOOKUP(C4,Data_Frame!$B:$B,Data_Frame!$G:$G)"
    End With
End Sub

Sub InitializeDataMemberList(sheetName As String)
    Dim ws As Worksheet
        
    ' Check if the worksheet exists
    If Not WorksheetExists(sheetName) Then
        ' Add the worksheet if it does not exist
        Set ws = ThisWorkbook.Worksheets.Add
        ws.name = sheetName
    Else
        ' Set the worksheet if it exists
        Set ws = ThisWorkbook.Worksheets(sheetName)
    End If

    ' Initialize headers
    With ws
        .Cells(1, 1).value = "rRow"
        .Cells(1, 2).value = "MemberID"
        .Cells(1, 3).value = "fFrame"
        .Cells(1, 4).value = "lFrame"
        .Cells(1, 5).value = "totalLen"
        .Cells(1, 6).value = "eleList"
        .Cells(1, 7).value = "restrNodesAll"
        .Cells(1, 8).value = "eleListBetweenRestrNodesAll"
        .Cells(1, 9).value = "restrNodesMajor"
        .Cells(1, 10).value = "eleListBetweenRestrNodesMajor"
        .Cells(1, 11).value = "restrNodesMinor"
        .Cells(1, 12).value = "eleListBetweenRestrNodesMinor"
    End With
End Sub

Sub InitializeWSData(sheetName As String)
    Dim ws As Worksheet
    
    ' Check if the worksheet exists
    If Not WorksheetExists(sheetName) Then
        ' Add the worksheet if it does not exist
        Set ws = ThisWorkbook.Worksheets.Add
        ws.name = sheetName
    Else
        ' Set the worksheet if it exists
        Set ws = ThisWorkbook.Worksheets(sheetName)
    End If

    ' Initialize headers
    With ws
        .Cells(1, 1).value = "rRow_tarWSData"
        .Cells(1, 2).value = "Design Worksheet 1"
        .Cells(1, 5).value = "Design Worksheet 2"

        ' Set header for Design Worksheet 1
        .Cells(2, 2).value = "Design Input"
        .Cells(2, 3).value = "Location"
        .Cells(2, 4).value = "Design Output"
        .Cells(2, 5).value = "Location"

        ' Set header for Design Worksheet 2
        .Cells(2, 6).value = "Design Input"
        .Cells(2, 7).value = "Location"
        .Cells(2, 8).value = "Design Output"
        .Cells(2, 9).value = "Location"

        ' Merge and center headers across cells
        .Range("B1:I1").HorizontalAlignment = xlCenterAcrossSelection

        ' Color columns B:I, except for row 2
        .Range("B3:I50").Interior.Color = RGB(192, 192, 192)
        .Range("B1:I1").Interior.Color = RGB(192, 192, 192)
    End With
End Sub

Sub InitializeDefineConnectionType(sheetName As String)
    Dim ws As Worksheet

    ' Check if the worksheet exists
    If Not WorksheetExists(sheetName) Then
        ' Add the worksheet if it does not exist
        Set ws = ThisWorkbook.Worksheets.Add
        ws.name = sheetName
    Else
        ' Set the worksheet if it exists
        Set ws = ThisWorkbook.Worksheets(sheetName)
    End If

    ' Initialize headers and formatting
    With ws
        ' Header titles
        .Cells(1, 2).value = "connID"
        .Cells(1, 3).value = "mapJt"
        .Cells(1, 4).value = "mapMethod"
        .Cells(1, 5).value = "mapSections"
        .Cells(1, 6).value = "isRestraint"
        .Cells(1, 7).value = "extractSections"
        .Cells(1, 8).value = "isFullyPopulate"
        .Cells(1, 9).value = "endOffset"
        .Cells(1, 10).value = "matchedJoint"
        
        ' Merging and labeling
        .Range("B3:E3").Merge
        .Cells(3, 2).value = "Mapping Controls"
        .Range("F3:J3").Merge
        .Cells(3, 6).value = "Data Extraction Control"
        
        .Range("B4").value = "Connection Name"
        .Range("C4").value = "Joint ID"
        .Range("D4").value = "Method: Contains/ Contains Only"
        .Range("E4").value = "Section of Members (use ',' to separate sections)"
        .Range("F4").value = "is Restraint?"
        .Range("G4").value = "Output Section"
        .Range("H4").value = "Fully Populate?"
        .Range("I4").value = "End Offset Setting"
        .Range("J4").value = "MACRO USE: Matched Joint"
        
        ' Center and color formatting
        .Range("B1:J1").HorizontalAlignment = xlCenterAcrossSelection
        .Range("B3:J3").Interior.Color = RGB(192, 192, 192)
        .Range("B4:E4").Interior.Color = RGB(224, 224, 224)
        .Range("F4:J4").Interior.Color = RGB(255, 224, 192)
    End With
End Sub

Sub InitializeDataSystem(sheetName As String)
    Dim ws As Worksheet

    ' Check if the worksheet exists
    If Not WorksheetExists(sheetName) Then
        ' Add the worksheet if it does not exist
        Set ws = ThisWorkbook.Worksheets.Add
        ws.name = sheetName
    Else
        ' Set the worksheet if it exists
        Set ws = ThisWorkbook.Worksheets(sheetName)
    End If

    ' Initialize cells with values and formulas
    With ws
        ' Initialize <isWSImported>
        .Cells(1, 1).value = "<isWSImported>"
        .Cells(2, 1).value = "ws_joint"
        .Cells(2, 2).formula = "=NOT(ISBLANK(OFFSET(Data_Joint!A2,1,1)))"
        .Cells(3, 1).value = "ws_frame"
        .Cells(3, 2).formula = "=NOT(ISBLANK(OFFSET(Data_Frame!A4,1,4)))"
        .Cells(4, 1).value = "ws_frameSection"
        .Cells(4, 2).formula = "=NOT(ISBLANK(OFFSET(Data_Frame!A4,1,2)))"
        .Cells(5, 1).value = "ws_frameForce"
        .Cells(5, 2).formula = "=NOT(ISBLANK(OFFSET(Data_FrameForce!A3,2,2)))"
        .Cells(6, 1).value = "ws_restraint"
        .Cells(6, 2).formula = "=NOT(ISBLANK(OFFSET(Data_Joint!A2,1,4)))"
        .Cells(7, 1).value = "<End Property>"

        ' Initialize <isCreated>
        .Cells(9, 1).value = "<isCreated>"
        .Cells(10, 1).value = "member"
        .Cells(10, 2).value = False
        .Cells(11, 1).value = "frameOrientationType"
        .Cells(11, 2).value = False
        .Cells(12, 1).value = "isIdentifiedConn"
        .Cells(12, 2).value = False
        .Cells(13, 1).value = "isMappedConn"
        .Cells(13, 2).value = False
        .Cells(14, 1).value = "<End Property>"

        ' Initialize <MemberCreateMethod>
        .Cells(16, 1).value = "<MemberCreateMethod>"
        .Cells(17, 1).value = "frameConnectivity"
        .Cells(17, 2).formula = "=B10"
        .Cells(18, 1).value = "MatchSection"
        .Cells(18, 2).formula = "=NOT(ISBLANK(OFFSET(Data_MemberList!A7,1,1)))"
        .Cells(19, 1).value = "LimitKinkAngle"
        .Cells(19, 2).formula = "=NOT(ISBLANK(OFFSET(Data_MemberList!A8,1,1)))"
        .Cells(20, 1).value = "<End Property>"

        ' Initialize <WhatOperationCanProcess>
        .Cells(22, 1).value = "<WhatOperationCanProcess>"
        .Cells(23, 1).value = "CreateMemberBySection"
        .Cells(23, 2).formula = "=AND(B3,B4)"
        .Cells(24, 1).value = "CreateMemberByKinkAngle"
        .Cells(24, 2).formula = "=AND(B2,B3,B4)"
        .Cells(25, 1).value = "IdentifyRestraint"
        .Cells(25, 2).formula = "=AND(B10,B6,B2,B3,B4,B11)"
        .Cells(26, 1).value = "GetIdentifiedConnType"
        .Cells(26, 2).formula = "=B27"
        .Cells(27, 1).value = "MapConnection"
        .Cells(27, 2).formula = "=AND(B12,B2,B3,B4,B6)"
        .Cells(28, 1).value = "ExtractFrameForce"
        .Cells(28, 2).formula = "=AND(B13,B5)"
        .Cells(29, 1).value = "ExtractConnectionForce"
        .Cells(29, 2).formula = "=AND(B13,B5)"
        .Cells(30, 1).value = "<End Property>"

        ' Initialize <Setting_Conn_ByDir>
        .Cells(32, 1).value = "<Setting_Conn_ByDir>"
        .Cells(33, 1).value = "kinkAngle"
        .Cells(33, 2).value = "30"
        .Cells(34, 1).value = "<End Property>"

        ' Initialize <ElementTypeIdentification>
        .Cells(36, 1).value = "<ElementTypeIdentification>"
        .Cells(37, 1).value = "DegToHoriPlane_Vert"
        .Cells(37, 2).value = "75"
        .Cells(38, 1).value = "DegToHoriPlane_Hori"
        .Cells(38, 2).value = "15"
        .Cells(39, 1).value = "<End Property>"
        
        .Cells(42, 1).value = "<ProcessDataOperationSave>"
        .Cells(43, 1).value = "isCreateMember"
        .Cells(43, 2).value = True
        .Cells(44, 1).value = "isIdentifyRestraint"
        .Cells(44, 2).value = True
        .Cells(45, 1).value = "isMapBySection"
        .Cells(45, 2).value = True
        .Cells(46, 1).value = "isMapByKinkAngle"
        .Cells(46, 2).value = True
        .Cells(47, 1).value = "<End Property>"
        
        .Cells(49, 1).value = "<logFilePath>"
        .Cells(50, 1).value = "path"
        .Cells(50, 2).value = ""
        .Cells(51, 1).value = "<End Property>"
        
        .Cells(53, 1).value = "<OutputFolderPath>"
        .Cells(54, 1).value = "path"
        .Cells(54, 2).value = ""
        .Cells(55, 1).value = "<End Property>"
        
        .Cells(57, 1).value = "<ImportLog>"
        .Cells(58, 1).value = "folderPath"
        .Cells(58, 2).value = ""
        .Cells(59, 1).value = "isDetailMode"
        .Cells(59, 2).value = True
        .Cells(60, 1).value = "filePath"
        .Cells(60, 2).value = ""
        .Cells(61, 1).value = "<End Property>"
        
        .Cells(63, 1).value = "<CustomPropGroupList>"
        .Cells(64, 1).value = "Name"
        .Cells(64, 2).value = ""
        .Cells(65, 1).value = "<End Property>"

        ' Set the column width for better readability
        .Columns("A:D").AutoFit
    End With
End Sub

Function WorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    WorksheetExists = Not ws Is Nothing
End Function
