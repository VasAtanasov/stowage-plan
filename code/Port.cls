VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Port"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("StowagePlan.core.model")
Private Type TPort
    PortName As String
    PortCode As String
    color As Long
    RowNumber As Long
    HoldData As Collection
    TotalUnits As Long
    TotalUnitsWeight As Double
    TotalPkgs As Long
    TotalPkgsWeight As Double
    LoadingPorts() As Port
End Type

Private this As TPort

Public Property Get PortName() As String
    PortName = this.PortName
End Property

Public Property Let PortName(ByVal RHS As String)
    this.PortName = RHS
End Property

Public Property Get PortCode() As String
    PortCode = this.PortCode
End Property

Public Property Let PortCode(ByVal RHS As String)
    this.PortCode = RHS
End Property

Public Property Get color() As Long
    color = this.color
End Property

Public Property Let color(ByVal RHS As Long)
    this.color = RHS
End Property

Public Property Get RowNumber() As Long
    RowNumber = this.RowNumber
End Property

Public Property Let RowNumber(ByVal RHS As Long)
    this.RowNumber = RHS
End Property

Public Property Get HoldData() As Collection
    Set HoldData = this.HoldData
End Property

Public Property Set HoldData(ByVal RHS As Collection)
    Set this.HoldData = RHS
End Property

Public Property Get TotalUnits() As Long
    TotalUnits = this.TotalUnits
End Property

Public Property Let TotalUnits(ByVal RHS As Long)
    this.TotalUnits = RHS
End Property

Public Property Get TotalUnitsWeight() As Double
    TotalUnitsWeight = this.TotalUnitsWeight
End Property

Public Property Let TotalUnitsWeight(ByVal RHS As Double)
    this.TotalUnitsWeight = RHS
End Property

Public Property Get TotalPkgs() As Long
    TotalPkgs = this.TotalPkgs
End Property

Public Property Let TotalPkgs(ByVal RHS As Long)
    this.TotalPkgs = RHS
End Property

Public Property Get TotalPkgsWeight() As Double
    TotalPkgsWeight = this.TotalPkgsWeight
End Property

Public Property Let TotalPkgsWeight(ByVal RHS As Double)
    this.TotalPkgsWeight = RHS
End Property

Public Property Get LoadingPorts() As Variant
    LoadingPorts = this.LoadingPorts
End Property

Private Sub Class_Initialize()
    Set this.HoldData = New Collection
    Dim holdNumber As Long
    Dim holdObj As Hold
    
    For holdNumber = 1 To HOLDS
        Set holdObj = New Hold
        holdObj.holdNumber = holdNumber
        this.HoldData.Add holdObj
    Next holdNumber
    
End Sub
