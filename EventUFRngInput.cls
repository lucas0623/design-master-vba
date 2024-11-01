VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventUFRngInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Userform.GeneralForm")

Private WithEvents textBox1 As msForms.Textbox
Attribute textBox1.VB_VarHelpID = -1
Private uf As Object

Public Sub Initialize(ufObj As Object, textBoxObj As msForms.Textbox)
    Set uf = ufObj
    Set textBox1 = textBoxObj
End Sub

Private Sub textBox1_DropButtonClick()
    Dim var As Variant
    Dim rng As Range
    Dim sFullAddress As String
    Dim sAddress As String
    
    uf.Hide
    
    On Error Resume Next
    Set rng = Application.inputBox("Select the range containing your data", _
        "Select Range Data", textBox1.Text, Type:=8)
    On Error GoTo 0
  
'  If TypeName(var) = "String" Then
'    CheckAddress CStr(var)
'  End If
    If Not rng Is Nothing Then
        textBox1.Text = CStr(rng.address)
    End If
  uf.Show
End Sub

'Private Sub CheckAddress(sAddress As String)
'  ' changed following advice of Julien steelandt@yahoo.fr
'  Dim rng As Range
'  Dim sFullAddress As String
'
'  If left$(sAddress, 1) = "=" Then sAddress = mID$(sAddress, 2, 256)
'  If left$(sAddress, 1) = Chr(34) Then sAddress = mID$(sAddress, 2, 255)
'  If Right$(sAddress, 1) = Chr(34) Then sAddress = left$(sAddress, Len(sAddress) - 1)
'
'  On Error Resume Next
'  sAddress = Application.ConvertFormula(sAddress, xlR1C1, xlA1)
'
'  If IsRange(sAddress) Then
'    Set rng = Range(sAddress)
'  End If
'
'  If Not rng Is Nothing Then
'    sFullAddress = rng.Address(, , Application.ReferenceStyle, True)
'    If left$(sFullAddress, 1) = "'" Then
'      sAddress = "'"
'    Else
'      sAddress = ""
'    End If
'    sAddress = sAddress & mID$(sFullAddress, InStr(sFullAddress, "]") + 1)
'
'    rng.Parent.Activate
'
'    Me.txtRefChtData.text = sAddress
'  End If
'
'End Sub

'Public Property Let Address(sAddress As String)
'  CheckAddress sAddress
'End Property
'
'Public Property Get Address() As String
'  Dim sAddress As String
'
'  sAddress = Me.txtRefChtData.text
'  If IsRange(sAddress) Then
'    Address = sAddress
'  Else
'    sAddress = Application.ConvertFormula(sAddress, xlR1C1, xlA1)
'    If IsRange(sAddress) Then
'      Address = sAddress
'    End If
'  End If
'
'End Property

'Public Property Let Textbox(obj As MSForms.Textbox)
'  Set Textbox = obj
'End Property
