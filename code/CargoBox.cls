VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CargoBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StowagePlan.feature.drawings.box")
Option Explicit

Private WithEvents packagesForm As PkgsForm
Attribute packagesForm.VB_VarHelpID = -1
Private m_boxTag As String

Private Sub Class_Initialize()
    Set packagesForm = New PkgsForm
End Sub

Public Sub InitiateProperties(ByVal boxTag As String)
    m_boxTag = boxTag
    Set packagesForm = New PkgsForm
End Sub

Public Sub Show()
    If Not packagesForm.Visible Then
        packagesForm.Show
    End If
End Sub

Private Sub packagesForm_OnAddPkgs()
    AddStaticCargoShape m_boxTag
End Sub

Public Property Get TextBoxValue(dest As String, load As String, units As String) As String
    TextBoxValue = dest & " / " & load & vbCr & _
                   packagesForm.PkgsCountTextBox.value & " " & units & " - " & packagesForm.PkgsWeightTextBox.value & " mts"
End Property


