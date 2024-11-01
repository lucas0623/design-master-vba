VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventComboBoxToListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Userform.GeneralForm")

Private pForm As UserForm
Private WithEvents pComboBox As msForms.ComboBox
Attribute pComboBox.VB_VarHelpID = -1
Private WithEvents pComboBox2 As msForms.ComboBox
Attribute pComboBox2.VB_VarHelpID = -1
Private pListBox1 As msForms.listBox
Private pDict As Object
Private iCount As Long
Private arr As Variant

Sub Init(form As UserForm, ListBox1 As msForms.listBox, _
                        ComboBox As msForms.ComboBox, dict As Object, Optional ComboBox2 As msForms.ComboBox)
                
    Set pForm = form
    Set pListBox1 = ListBox1
    Set pComboBox = ComboBox
    Set pComboBox2 = ComboBox2
    Set pDict = dict
    
End Sub

Private Sub pComboBox_Change()
    If Not pComboBox2 Is Nothing Then
        pListBox1.List = pDict(pComboBox.value & pComboBox2.value)
    Else
        pListBox1.List = pDict(pComboBox.value)
    End If
End Sub

Private Sub pComboBox2_Change()
    pListBox1.List = pDict(pComboBox.value & pComboBox2.value)
End Sub


