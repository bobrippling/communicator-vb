VERSION 5.00
Begin VB.UserControl Graph 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   1755
      Left            =   510
      ScaleHeight     =   1755
      ScaleWidth      =   3510
      TabIndex        =   0
      Top             =   630
      Visible         =   0   'False
      Width           =   3510
   End
End
Attribute VB_Name = "Graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Const SRCCOPY = &HCC0020
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private WithEvents mobjDatasets      As Datasets
Attribute mobjDatasets.VB_VarHelpID = -1

Private mudtControlProps    As gtypControlProps
Private mudtGraphProps      As gtypGraphProps

Private mblnDesignMode  As Boolean

Public Enum eBorderStyle
   egrNone = 0
   egrFixedSingle = 1
End Enum

Public Enum eAppearance
   egrFlat = 0
   egr3D = 1
End Enum

Private Type mtypPOINT
    X   As Long
    Y   As Long
End Type

Private Type mtypRECT
    Left    As Long
    Right   As Long
    Top     As Long
    Bottom  As Long
End Type

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub UserControl_Initialize()
    picDraw.FillStyle = vbFSSolid
    Set mobjDatasets = New Datasets
End Sub

Private Sub UserControl_InitProperties()
    InitProperties
    mblnDesignMode = Not UserControl.Ambient.UserMode
    DrawControl
    PropertyChanged PB_STATE
End Sub

Private Sub UserControl_Paint()
    DrawGraph
End Sub

Private Sub UserControl_Terminate()
    Set mobjDatasets = Nothing
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    State = PropBag.ReadProperty(PB_STATE, State)
    mblnDesignMode = Not UserControl.Ambient.UserMode
    DrawControl
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty PB_STATE, State
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    picDraw.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
On Error GoTo 0
    DrawGraph
End Sub

Private Sub mobjDatasets_Changed()
Static blnWorking As Boolean
    If Not blnWorking Then
        blnWorking = True
        RemovePoints
        If Not mblnDesignMode Then
            DrawGraph
        End If
        blnWorking = False
    End If
End Sub

Private Property Let GraphState(ByRef Value() As Byte)
Dim udtData     As gtypGraphData
    udtData.Data = Value
    LSet mudtGraphProps = udtData
End Property

Private Property Get GraphState() As Byte()
Dim udtData     As gtypGraphData
    LSet udtData = mudtGraphProps
    GraphState = udtData.Data
End Property

Friend Property Let ControlState(ByRef Value() As Byte)
Dim udtData     As gtypControlData
    udtData.Data = Value
    LSet mudtControlProps = udtData
End Property

Friend Property Get ControlState() As Byte()
Dim udtData     As gtypControlData
    LSet udtData = mudtControlProps
    ControlState = udtData.Data
End Property

Private Property Let State(ByRef Value() As Byte)
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .Contents = Value
        ControlState = .ReadProperty(PB_CONTROL)
        GraphState = .ReadProperty(PB_GRAPH)
    End With
    Set objPB = Nothing
End Property

Private Property Get State() As Byte()
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .WriteProperty PB_CONTROL, ControlState
        .WriteProperty PB_GRAPH, GraphState
        State = .Contents
    End With
    Set objPB = Nothing
End Property

Friend Property Let SuperState(ByRef Value() As Byte)
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .Contents = Value
        State = .ReadProperty(PB_STATE, State)
        mobjDatasets.SuperState = .ReadProperty(PB_DATASETS, mobjDatasets.SuperState)
    End With
    Set objPB = Nothing
End Property

Friend Property Get SuperState() As Byte()
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .WriteProperty PB_STATE, State
        .WriteProperty PB_DATASETS, mobjDatasets.SuperState
        SuperState = .Contents
    End With
    Set objPB = Nothing
End Property

Friend Property Let FileState(ByRef Value() As Byte)
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .Contents = Value
        GraphState = .ReadProperty(PB_GRAPH, GraphState)
        mobjDatasets.SuperState = .ReadProperty(PB_POINTS, mobjDatasets.SuperState)
    End With
    Set objPB = Nothing
End Property

Friend Property Get FileState() As Byte()
Dim objPB   As PropertyBag
    Set objPB = New PropertyBag
    With objPB
        .WriteProperty PB_GRAPH, GraphState
        .WriteProperty PB_POINTS, mobjDatasets.SuperState
        FileState = .Contents
    End With
    Set objPB = Nothing
End Property

Private Sub InitProperties()
    With mudtGraphProps
        .BackColor = RGB(255, 255, 255)
        .AxisColor = RGB(0, 0, 0)
        .GridColor = RGB(223, 223, 223)
        .FixedPoints = 20
        .XGridInc = 1
        .YGridInc = 10
        .MaxValue = 100
        .MinValue = 0
        .FadeIn = False
        .ShowGrid = True
        .ShowAxis = False
        .BarWidth = 0.8
    End With
    With mudtControlProps
        .Redraw = True
        .BorderStyle = eBorderStyle.egrFixedSingle
        .Appearance = eAppearance.egr3D
    End With
End Sub

Public Property Get Datasets() As Datasets
    Set Datasets = mobjDatasets
End Property

Public Property Let Redraw(ByVal Value As Boolean)
    mudtControlProps.Redraw = Value
    If Value Then
        Refresh
    End If
End Property

Public Property Get Redraw() As Boolean
    Redraw = mudtControlProps.Redraw
End Property

Public Property Let Appearance(ByVal Value As eAppearance)
    mudtControlProps.Appearance = Value
    UserControl.Appearance = Value
    DrawControl
    PropertyChanged PB_STATE
End Property

Public Property Get Appearance() As eAppearance
    Appearance = mudtControlProps.Appearance
End Property

Public Property Let BorderStyle(ByVal Value As eBorderStyle)
    mudtControlProps.BorderStyle = Value
    UserControl.BorderStyle = Value
    DrawControl
    PropertyChanged PB_STATE
End Property

Public Property Get BorderStyle() As eBorderStyle
    BorderStyle = mudtControlProps.BorderStyle
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
    mudtGraphProps.BackColor = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mudtGraphProps.BackColor
End Property

Public Property Let AxisColor(ByVal Value As OLE_COLOR)
    mudtGraphProps.AxisColor = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get AxisColor() As OLE_COLOR
    AxisColor = mudtGraphProps.AxisColor
End Property

Public Property Let GridColor(ByVal Value As OLE_COLOR)
    mudtGraphProps.GridColor = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get GridColor() As OLE_COLOR
    GridColor = mudtGraphProps.GridColor
End Property

Public Property Let FixedPoints(ByVal Value As Long)
    mudtGraphProps.FixedPoints = Value
    RemovePoints
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get FixedPoints() As Long
    FixedPoints = mudtGraphProps.FixedPoints
End Property

Public Property Let XGridInc(ByVal Value As Long)
    mudtGraphProps.XGridInc = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get XGridInc() As Long
    XGridInc = mudtGraphProps.XGridInc
End Property

Public Property Let YGridInc(ByVal Value As Double)
    mudtGraphProps.YGridInc = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get YGridInc() As Double
    YGridInc = mudtGraphProps.YGridInc
End Property

Public Property Let MaxValue(ByVal Value As Double)
    mudtGraphProps.MaxValue = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get MaxValue() As Double
    MaxValue = mudtGraphProps.MaxValue
End Property

Public Property Let MinValue(ByVal Value As Double)
    mudtGraphProps.MinValue = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get MinValue() As Double
    MinValue = mudtGraphProps.MinValue
End Property

Public Property Let ShowGrid(ByVal Value As Boolean)
    mudtGraphProps.ShowGrid = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get ShowGrid() As Boolean
    ShowGrid = mudtGraphProps.ShowGrid
End Property

Public Property Let ShowAxis(ByVal Value As Boolean)
    mudtGraphProps.ShowAxis = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get ShowAxis() As Boolean
    ShowAxis = mudtGraphProps.ShowAxis
End Property

Public Property Let FadeIn(ByVal Value As Boolean)
    mudtGraphProps.FadeIn = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get FadeIn() As Boolean
    FadeIn = mudtGraphProps.FadeIn
End Property

Public Property Let BarWidth(ByVal Value As Single)
    mudtGraphProps.BarWidth = Value
    DrawGraph
    PropertyChanged PB_STATE
End Property

Public Property Get BarWidth() As Single
    BarWidth = mudtGraphProps.BarWidth
End Property

Private Sub AddDefaultPoints()
    mobjDatasets.Clear
    mobjDatasets.Add
    AddDefaultPoint 80
    AddDefaultPoint 10
    AddDefaultPoint 70
    AddDefaultPoint 25
    AddDefaultPoint 50
    AddDefaultPoint 45
    AddDefaultPoint 15
    AddDefaultPoint 85
    AddDefaultPoint 5
    AddDefaultPoint 75
    AddDefaultPoint 65
End Sub

Private Sub AddDefaultPoint(ByVal plngPercent As Long)
    mobjDatasets.Item(1).Points.Add (plngPercent / 100) * (mudtGraphProps.MaxValue - mudtGraphProps.MinValue) + mudtGraphProps.MinValue
End Sub

Private Sub RemovePoints()
Dim objDataset  As Dataset
    If mudtGraphProps.FixedPoints > 0 Then
        For Each objDataset In mobjDatasets
            Do While objDataset.Points.Count > mudtGraphProps.FixedPoints
                objDataset.Points.Remove 1
            Loop
        Next objDataset
    End If
End Sub

Public Sub Refresh()
    DrawGraph
End Sub

Private Sub DrawControl()
    With UserControl
        .Appearance = mudtControlProps.Appearance
        .BorderStyle = mudtControlProps.BorderStyle
    End With
End Sub

Private Sub DrawGraph()
Dim lngWidth        As Long
Dim lngHeight       As Long
Dim lngIndex        As Long
Dim lngYAxis        As Long
Dim udtGrid         As mtypRECT
Dim objDataset      As Dataset
Dim udtRect()       As mtypRECT
Dim udtPoint        As mtypPOINT
Dim udtPrevPoint    As mtypPOINT
Dim udtCap          As mtypRECT
Dim udtBar          As mtypRECT
Dim lngCount        As Long
Dim lngOffset       As Long

    If UserControl.Height > 0 And UserControl.Width > 0 Then
        If mudtControlProps.Redraw Or mblnDesignMode Then
            If mblnDesignMode Then
                AddDefaultPoints
            End If
            With picDraw
                .Cls
                .BackColor = mudtGraphProps.BackColor
    
                udtGrid.Left = 0
                udtGrid.Top = 0
                udtGrid.Right = .ScaleWidth - 15
                udtGrid.Bottom = .ScaleHeight - 15
                lngWidth = udtGrid.Right - udtGrid.Left
                lngHeight = udtGrid.Bottom - udtGrid.Top
                lngYAxis = GetYAxis(udtGrid)
                DrawGrid udtGrid, mudtGraphProps.GridColor
                
                For Each objDataset In mobjDatasets
                    If objDataset.Points.Count > 0 And objDataset.Visible Then
                        If mudtGraphProps.ShowAxis Then
                            picDraw.Line (0, 0)-(0, lngHeight), mudtGraphProps.AxisColor
                            picDraw.Line (0, lngYAxis)-(lngWidth, lngYAxis), mudtGraphProps.AxisColor
                        End If
                        
                        udtRect = GetRectArrayForDataset(udtGrid, objDataset)
                        lngCount = UBound(udtRect)
                        If objDataset.ShowBars Or objDataset.ShowCaps Then
                            For lngIndex = 1 To lngCount
                                lngOffset = (udtRect(lngIndex).Right - udtRect(lngIndex).Left) * (1 - mudtGraphProps.BarWidth)
                                udtBar.Left = udtRect(lngIndex).Left + lngOffset
                                udtBar.Right = udtRect(lngIndex).Right - lngOffset
                                udtBar.Top = udtRect(lngIndex).Top
                                udtBar.Bottom = udtRect(lngIndex).Bottom
                                If objDataset.ShowBars Then
                                    DrawBar udtBar, objDataset.BarColor
                                End If
                                If objDataset.ShowCaps Then
                                    LSet udtCap = udtBar
                                    udtCap.Bottom = udtCap.Top - 15
                                    udtCap.Top = udtCap.Top + 15
                                    DrawBar udtCap, objDataset.CapColor
                                End If
                            Next lngIndex
                        End If
'                        If objDataset.ShowCaps Then
'                            For lngIndex = 1 To lngCount
'                                udtCap.Left = udtRect(lngIndex).Left + lngOffset
'                                udtCap.Right = udtRect(lngIndex).Right - lngOffset
'                                udtCap.Top = udtRect(lngIndex).Top + 15
'                                udtCap.Bottom = udtRect(lngIndex).Top - 15
'                                DrawBar udtCap, objDataset.CapColor
'                            Next lngIndex
'                        End If
                        If objDataset.ShowLines Then
                            For lngIndex = 2 To lngCount
                                udtPoint.X = udtRect(lngIndex).Left + ((udtRect(lngIndex).Right - udtRect(lngIndex).Left) / 2)
                                udtPoint.Y = udtRect(lngIndex).Top
                                udtPrevPoint.X = udtRect(lngIndex - 1).Left + ((udtRect(lngIndex - 1).Right - udtRect(lngIndex - 1).Left) / 2)
                                udtPrevPoint.Y = udtRect(lngIndex - 1).Top
                                DrawLine udtPoint, udtPrevPoint, objDataset.LineColor
                            Next lngIndex
                        End If
                        If objDataset.ShowPoints Then
                            For lngIndex = 1 To lngCount
                                If (lngIndex Mod mudtGraphProps.XGridInc = 0 Or lngIndex = 1) Then
                                    udtPoint.X = udtRect(lngIndex).Left + ((udtRect(lngIndex).Right - udtRect(lngIndex).Left) / 2)
                                    udtPoint.Y = udtRect(lngIndex).Top
                                    DrawPoint udtPoint, objDataset.PointColor
                                End If
                            Next lngIndex
                        End If
                    End If
                Next objDataset
                BitBlt UserControl.hDC, 0, 0, .ScaleWidth, .ScaleHeight, .hDC, 0, 0, SRCCOPY
            End With
        End If
    End If
End Sub

Private Sub DrawLine(ByRef pudtPt1 As mtypPOINT, ByRef pudtPt2 As mtypPOINT, ByVal plngColor As String)
    picDraw.Line (pudtPt1.X, pudtPt1.Y)-(pudtPt2.X, pudtPt2.Y), plngColor
End Sub

Private Sub DrawPoint(ByRef pudtPt As mtypPOINT, ByVal plngColor As Long)
    picDraw.FillColor = plngColor
    picDraw.Circle (pudtPt.X, pudtPt.Y), 40, 0
End Sub

Private Sub DrawBar(ByRef pudtRect As mtypRECT, ByVal plngColor As Long)
    picDraw.FillColor = plngColor
    With pudtRect
        picDraw.Line (.Left, .Top)-(.Right, .Bottom), 0, B
    End With
End Sub

Private Sub DrawGrid(ByRef pudtRect As mtypRECT, ByVal plngColor As Long)
Dim lngCount        As Long
Dim lngIndex        As Long
Dim lngX            As Long
Dim lngY            As Long
Dim lngFixedCt      As Long
Dim lngYAxis        As Long
Dim lngStepY        As Long
Dim lngHeight       As Long
Dim lngGap          As Long
Dim lngOffset       As Long
Dim lngRem          As Long
Dim lngWidth        As Long
    lngFixedCt = GetMaxPointCount
    If lngFixedCt > 0 And mudtGraphProps.ShowGrid Then
        lngWidth = pudtRect.Right - pudtRect.Left
        lngHeight = Abs(pudtRect.Bottom - pudtRect.Top)

        If mudtGraphProps.XGridInc > 0 Then
            GetGapAndOffset pudtRect, lngGap, lngOffset, lngRem
            If lngGap > 0 Then
                lngX = pudtRect.Left + (lngOffset + lngGap / 2)
                For lngIndex = 1 To lngFixedCt
                    If (lngIndex Mod mudtGraphProps.XGridInc = 0 Or lngIndex = 1) Then
                        picDraw.Line (lngX, pudtRect.Top)-(lngX, pudtRect.Bottom), mudtGraphProps.GridColor
                    End If
                    If lngIndex = lngFixedCt - lngRem Then
                        lngGap = lngGap + 1
                    End If
                    lngX = lngX + lngGap
                Next lngIndex
            End If
        End If

        If mudtGraphProps.YGridInc > 0 Then
            If (mudtGraphProps.MaxValue - mudtGraphProps.MinValue) > 0 Then
                lngYAxis = GetYAxis(pudtRect)
                lngStepY = (lngHeight / (mudtGraphProps.MaxValue - mudtGraphProps.MinValue)) * mudtGraphProps.YGridInc
                
                For lngY = lngYAxis To pudtRect.Top Step -lngStepY
                    picDraw.Line (pudtRect.Left, lngY)-(pudtRect.Right, lngY), mudtGraphProps.GridColor
                Next lngY
                
                For lngY = lngYAxis To pudtRect.Bottom Step lngStepY
                    picDraw.Line (pudtRect.Left, lngY)-(pudtRect.Right, lngY), mudtGraphProps.GridColor
                Next lngY
            End If
        End If
    End If
End Sub

Private Function GetYAxis(ByRef pudtRect As mtypRECT) As Long
Dim lngYAxis    As Long
    With mudtGraphProps
        If .MaxValue <= 0 And .MinValue < 0 Then
            lngYAxis = 0
        ElseIf .MaxValue > 0 And .MinValue >= 0 Then
            lngYAxis = pudtRect.Bottom
        Else
            lngYAxis = pudtRect.Top + (.MaxValue * ((pudtRect.Bottom - pudtRect.Top) / (.MaxValue - .MinValue)))
        End If
    End With
    GetYAxis = lngYAxis
End Function

Private Function GetMaxPointCount() As Long
Dim objDataset  As Dataset
Dim lngMaxCount    As Long
Dim lngPtCount  As Long
    If mudtGraphProps.FixedPoints = 0 Then
        For Each objDataset In mobjDatasets
            If objDataset.Visible Then
                lngPtCount = objDataset.Points.Count
                If lngPtCount > lngMaxCount Then
                    lngMaxCount = lngPtCount
                End If
            End If
        Next objDataset
    Else
        lngMaxCount = mudtGraphProps.FixedPoints
    End If
    GetMaxPointCount = lngMaxCount
End Function

Private Function AnyDatasetShowingBars() As Boolean
Dim blnShow     As Boolean
Dim objDataset  As Dataset
    For Each objDataset In mobjDatasets
        If (objDataset.ShowBars Or objDataset.ShowCaps) And objDataset.Visible Then
            blnShow = True
            Set objDataset = Nothing
            Exit For
        End If
    Next objDataset
    AnyDatasetShowingBars = blnShow
End Function

Private Function GetRectArrayForDataset(ByRef pudtGraph As mtypRECT, ByRef pobjDataset As Dataset) As mtypPOINT()
Dim lngWidth    As Long
Dim lngHeight   As Long
Dim lngFixedCt  As Long
Dim lngX        As Long
Dim lngY        As Long
Dim lngGap      As Long
Dim lngOffset   As Long
Dim dblRange    As Double
Dim lngYAxis    As Long
Dim udtRect()   As mtypRECT
Dim objPoint    As Point
Dim lngIndex    As Long
Dim lngRem      As Long
    If pobjDataset.Points.Count > 0 And pobjDataset.Visible Then
        With pudtGraph
            lngWidth = .Right - .Left
            lngHeight = .Bottom - .Top
        End With
        lngFixedCt = GetMaxPointCount
        GetGapAndOffset pudtGraph, lngGap, lngOffset, lngRem
        dblRange = MaxValue - MinValue
        lngYAxis = GetYAxis(pudtGraph)
        ReDim udtRect(pobjDataset.Points.Count) As mtypRECT
        lngX = pudtGraph.Left + lngOffset
        For Each objPoint In pobjDataset.Points
            lngIndex = lngIndex + 1
            If lngFixedCt - lngIndex - 1 = lngRem Then
                lngGap = lngGap + 1
            End If
            udtRect(lngIndex).Left = lngX
            udtRect(lngIndex).Right = lngX + lngGap
            udtRect(lngIndex).Bottom = lngYAxis
            udtRect(lngIndex).Top = lngYAxis - (objPoint.Value * (lngHeight / dblRange))
            lngX = lngX + lngGap
        Next objPoint
        GetRectArrayForDataset = udtRect
    End If
End Function

Private Sub GetGapAndOffset(ByRef pudtGraph As mtypRECT, ByRef plngGap As Long, ByRef plngOffset, ByRef plngReminder As Long)
Dim lngWidth    As Long
Dim lngGap      As Long
Dim lngOffset   As Long
Dim lngFixedCt  As Long
Dim lngRem      As Long
    lngFixedCt = GetMaxPointCount
    lngWidth = pudtGraph.Right - pudtGraph.Left
    If AnyDatasetShowingBars Then
        lngGap = lngWidth / lngFixedCt
        If lngGap > 0 Then
            If lngWidth - (lngGap * lngFixedCt) < 0 Then
                lngGap = lngGap - 1
            End If
            lngRem = lngWidth - (lngGap * lngFixedCt)
        End If
        lngOffset = 0
    Else
        If lngFixedCt > 1 Then
            lngGap = lngWidth / (lngFixedCt - 1)
            If lngGap > 0 Then
                If lngWidth - (lngGap * (lngFixedCt - 1)) < 0 Then
                    lngGap = lngGap - 1
                End If
                lngRem = lngWidth - (lngGap * (lngFixedCt - 1))
            End If
        Else
            lngGap = lngWidth
        End If
        lngOffset = -(lngGap / 2)
    End If
    plngGap = lngGap
    plngOffset = lngOffset
    plngReminder = lngRem
End Sub

Public Sub SaveSettings(ByVal Filename As String)
    If Len(Filename) > 0 Then
        If Dir(Filename) <> vbNullString Then
            Kill Filename
        End If
    End If
    SaveFile Filename, FileState
End Sub

Public Sub LoadSettings(ByVal Filename As String)
    FileState = GetFile(Filename)
    Refresh
End Sub

