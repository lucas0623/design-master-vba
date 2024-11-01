Attribute VB_Name = "Module2"

'@Folder("Module")
'@Folder("VBAProject")
Sub ExportAllVBAComponents()
    Dim vbComp As Object
    Dim vbProj As Object
    Dim exportPath As String
    Dim fileName As String
    
    ' Set the export path
    exportPath = "C:\Users\lucasleung\OneDrive\12 Engineering\01 Resource Lib\04 Useful Excel Macro\Tools Addin\Useful Tools VBA Modules" ' Change this to your desired folder path
    
    ' Ensure the path ends with a backslash
    If Right(exportPath, 1) <> "\" Then exportPath = exportPath & "\"
    
    ' Loop through each component in the VBA project
    Debug.Print ThisWorkbook.name
    Set vbProj = ThisWorkbook.VBProject
    For Each vbComp In vbProj.VBComponents
        Select Case vbComp.Type
            Case 1 ' Standard Module
                fileName = exportPath & vbComp.name & ".bas"
                vbComp.Export fileName
            Case 2 ' Class Module
                fileName = exportPath & vbComp.name & ".cls"
                vbComp.Export fileName
            Case 3 ' UserForm
                fileName = exportPath & vbComp.name & ".frm"
                vbComp.Export fileName
            Case Else
                ' Handle other types if necessary
        End Select
    Next vbComp

    MsgBox "Export completed!"
End Sub

Sub TestDebugger()
    ' Create an instance of the Debugger class
    Dim dbg As New Debugger
    
    ' Test with a 1D array
    Dim arr1D(1 To 3) As Integer
    arr1D(1) = 1
    arr1D(2) = 20
    arr1D(3) = 30
    dbg.Log arr1D, "arr1D"
    
    ' Test with a 2D array
    Dim arr2D(1 To 2, 1 To 2) As String
    arr2D(1, 1) = "A"
    arr2D(1, 2) = "B"
    arr2D(2, 1) = "C"
    arr2D(2, 2) = "D"
    dbg.Log arr2D, "arr2D"
    
    ' Test with a Collection
    Dim coll As Collection
    Set coll = New Collection
    coll.Add "First", "Key1"
    coll.Add "Second", "Key2"
    coll.Add "Third", "Key3"
    dbg.Log coll, "coll"
    
    ' Test with a Dictionary
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "A", 1
    dict.Add "B", 2
    dict.Add "C", 3
    dbg.Log dict, "dict"
    
    ' Test with a nested Collection
    Dim nestedColl As Collection
    Set nestedColl = New Collection
    nestedColl.Add coll
    nestedColl.Add dict
    dbg.Log nestedColl, "nestedColl"
    
    ' Test with nested Dictionary
    Dim nestedDict As Object
    Set nestedDict = CreateObject("Scripting.Dictionary")
    nestedDict.Add "SubColl", coll
    nestedDict.Add "SubDict", dict
    dbg.Log nestedDict, "nestedDict"
    
    ' Test with a simple variable
    Dim simpleVar As String
    simpleVar = "Hello, World!"
    dbg.Log simpleVar, "simpleVar"
End Sub
