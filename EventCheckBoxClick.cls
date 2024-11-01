VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventCheckBoxClick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Userform.GeneralForm")

Private pForm As UserForm
Private WithEvents pcb_master As msForms.checkbox
Attribute pcb_master.VB_VarHelpID = -1
Private pcb_slave() As msForms.checkbox
Private pArrNum() As Long


Sub Init(form As UserForm, cb_master As msForms.checkbox, _
                        cb_slave() As msForms.checkbox, isAll As Boolean, arrNum() As Long)
    Dim i As Long
    
    Set pForm = form
    Set pcb_master = cb_master
    ReDim pcb_slave(LBound(cb_slave) To UBound(cb_slave))
    For i = LBound(cb_slave) To UBound(cb_slave)
        Set pcb_slave(i) = cb_slave(i)
    Next i
    
    If isAll Then
        ReDim pArrNum(LBound(cb_slave) To UBound(cb_slave))
        For i = LBound(cb_slave) To UBound(cb_slave)
            pArrNum(i) = i
        Next i
    Else
        ReDim pArrNum(LBound(arrNum) To UBound(arrNum))
        For i = LBound(arrNum) To UBound(arrNum)
            pArrNum(i) = arrNum(i)
        Next i
    End If

End Sub

Private Sub pcb_master_Click()
    Dim i As Long
    For i = LBound(pcb_slave) To UBound(pcb_slave)
        If IsInArr(i, pArrNum) Then
            pcb_slave(i).value = pcb_master.value
        Else
            pcb_slave(i).value = False
        End If
    Next i
End Sub


Private Function IsInArr(str As Variant, arr As Variant) As Boolean
    Dim i As Long
    If IsArray(arr) Then
        For i = LBound(arr) To UBound(arr)
            If str = arr(i) Then
                IsInArr = True
                Exit Function
            End If
        Next i
    Else
        If str = arr Then
            IsInArr = True
            Exit Function
        End If
    End If
    IsInArr = False
End Function


