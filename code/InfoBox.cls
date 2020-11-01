VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InfoBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StowagePlan.feature.drawings.box")
Option Explicit

Private WithEvents InfoBoxFormObj As InfoForm
Attribute InfoBoxFormObj.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set InfoBoxFormObj = New InfoForm
End Sub

Public Sub InitiateProperties()
    Set InfoBoxFormObj = New InfoForm
End Sub

Public Sub Show()
    If Not InfoBoxFormObj.Visible Then
        InfoBoxFormObj.Show
    End If
End Sub

Private Sub InfoBoxFormObj_OnAddInfoBox()
    If Trim$(InfoBoxFormObj.InfoBoxTextArea.value) = vbNullString Then
        Unload InfoBoxFormObj
        Exit Sub
    End If
    InfoBoxShape
End Sub

Public Property Get TextValue() As String
    TextValue = InfoBoxFormObj.InfoBoxTextArea.value
End Property


