VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidateListBoxInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVVM.Infrastructure")
Option Explicit
Implements IValueValidator

Private genFuncs As New clsGeneralFunctions

 
Private Function IValueValidator_IsValid(ByVal value As Variant, ByVal Source As Object, ByVal Target As Object) As Boolean
    If genFuncs.isInitialised(value) Then
        If Not (UBound(value) = 0 And value(LBound(value)) = vbNullString) Then
            IValueValidator_IsValid = True
        End If
    End If
End Function
 
Private Property Get IValueValidator_Message() As String
    IValueValidator_Message = "Value cannot be empty."
End Property
