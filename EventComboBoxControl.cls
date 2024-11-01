VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventComboBoxControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Userform.GeneralForm")

Private pForm As UserForm
Private WithEvents pcb_master As msForms.ComboBox
Attribute pcb_master.VB_VarHelpID = -1
Private pcb_slave() As Object
Private pVal_visible As String

Sub Init(form As UserForm, cb_master As msForms.ComboBox, _
                        cb_slave() As Object, val_visible As String)
    Dim i As Long
    
    Set pForm = form
    Set pcb_master = cb_master
    ReDim pcb_slave(LBound(cb_slave) To UBound(cb_slave))
    For i = LBound(cb_slave) To UBound(cb_slave)
        Set pcb_slave(i) = cb_slave(i)
    Next i
    pVal_visible = val_visible
End Sub

Private Sub pcb_master_Change()
    Dim i As Long
    
    If pcb_master.value = pVal_visible Then
        For i = LBound(pcb_slave) To UBound(pcb_slave)
            pcb_slave(i).Visible = True
        Next i
    Else
        For i = LBound(pcb_slave) To UBound(pcb_slave)
            pcb_slave(i).Visible = False
        Next i
    End If
End Sub


