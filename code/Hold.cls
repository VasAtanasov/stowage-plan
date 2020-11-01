VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StowagePlan.core.model")
Private Type THold
    Count As Long
    Weight As Double
    holdNumber As Long
End Type

Private this As THold

Public Property Get Count() As Long
    Count = this.Count
End Property

Public Property Let Count(ByVal RHS As Long)
    this.Count = RHS
End Property

Public Property Get Weight() As Double
    Weight = this.Weight
End Property

Public Property Let Weight(ByVal RHS As Double)
    this.Weight = RHS
End Property

Public Property Get holdNumber() As Long
    holdNumber = this.holdNumber
End Property

Public Property Let holdNumber(ByVal RHS As Long)
    this.holdNumber = RHS
End Property

Private Sub Class_Initialize()
    this.Weight = 0
    this.Count = 0
End Sub
