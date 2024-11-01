VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CreateNewSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Operation"
Option Explicit
Private newSheet As Worksheet
Public Sub CreateEmptySummary()
    CopyHiddenSheet "Summary_TemplateEmpty"
End Sub
Public Sub CreateForGeneralPurpose()

    CopyHiddenSheet "Summary_TemplateGeneral"
End Sub
Public Sub CreateForSteelMember()
    CopyHiddenSheet "Summary_TemplateStlMem"
End Sub
Public Sub CreateForSteelConnection()
    CopyHiddenSheet "Summary_TemplateStlConn"
End Sub
Public Sub CreateForRC()
    CopyHiddenSheet "Summary_TemplateRC"
End Sub

Private Sub CopyHiddenSheet(name As String)
    Dim oldSheet As Worksheet
    Set oldSheet = Worksheets(name)
    oldSheet.Visible = xlSheetVisible
    oldSheet.Copy Before:=ActiveSheet
    oldSheet.Visible = xlSheetHidden
End Sub
