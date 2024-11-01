VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventSelectionBoxMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Userform.GeneralForm")

Private pForm As UserForm
Private pListBox1 As msForms.listBox
Private pListBox2 As msForms.listBox
Private WithEvents pAddBtm As msForms.CommandButton
Attribute pAddBtm.VB_VarHelpID = -1
Private WithEvents pRemoveBtm As msForms.CommandButton
Attribute pRemoveBtm.VB_VarHelpID = -1
Private iCount As Long



Sub Init(form As UserForm, ListBox1 As msForms.listBox, _
                        ListBox2 As msForms.listBox, addBtm As msForms.CommandButton, _
                        removeBtm As msForms.CommandButton)

    Set pForm = form
    Set pListBox1 = ListBox1
    Set pListBox2 = ListBox2
    Set pAddBtm = addBtm
    Set pRemoveBtm = removeBtm
End Sub

Private Sub pAddBtm_Click()

    Dim counter As Integer
    Dim i As Integer

    counter = 0
    For i = 0 To pListBox1.ListCount - 1
        If pListBox1.Selected(i - counter) Then
            pListBox2.AddItem pListBox1.List(i - counter)
            pListBox1.RemoveItem (i - counter)
            counter = counter + 1
        End If
    Next i
End Sub

Private Sub pRemoveBtm_Click()

    Dim counter As Integer
    Dim i As Integer

    For i = 0 To pListBox2.ListCount - 1
        If pListBox2.Selected(i - counter) Then
            pListBox1.AddItem pListBox2.List(i - counter)
            pListBox2.RemoveItem (i - counter)
            counter = counter + 1
        End If
    Next i
End Sub
