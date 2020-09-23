VERSION 5.00
Begin VB.UserControl CircleProg 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1950
   PropertyPages   =   "CircleProg.ctx":0000
   ScaleHeight     =   116
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   130
   ToolboxBitmap   =   "CircleProg.ctx":0016
   Begin VB.PictureBox Board 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   840
      Top             =   120
   End
   Begin VB.PictureBox Buffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   480
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "CircleProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Const PI = 3.14159

Dim pValue As Long, pMax As Long, pShowPer As Boolean, pRad As Integer, pDepth As Integer
Dim cValue As Long, cNonValue As Long, cBack As Long, cPer As Long
Dim pMode As Boolean

Private Function DrawBar()
On Error Resume Next
Buffer.Cls

Dim I As Long, per, xs, ys, cx, cy

per = pValue / pMax * 100
per = per / 100
per = 360 * per

cx = Board.ScaleWidth \ 2
cy = Board.ScaleHeight \ 2

Buffer.DrawWidth = 3

For I = 0 To 360 Step 2
xs = Cos(I / 180 * PI) * pRad
ys = Sin(I / 180 * PI) * pRad
Buffer.Line (cx, cy)-(cx + xs, cy + ys), cNonValue

If pMode = True Then If I > 0 And I < 180 Then Buffer.Line (cx + xs, cy + ys)-(cx + xs, cy + ys + pDepth), GetDarkerColor(cNonValue)

DoEvents
Next I

For I = 0 To per Step 2
xs = Cos(I / 180 * PI) * pRad
ys = Sin(I / 180 * PI) * pRad
Buffer.Line (cx, cy)-(cx + xs, cy + ys), cValue

If pMode = True Then If I > 0 And I < 180 Then Buffer.Line (cx + xs, cy + ys)-(cx + xs, cy + ys + pDepth), GetDarkerColor(cValue)

DoEvents
Next I

Dim n As Boolean
n = pShowPer
per = pValue / pMax * 100

If n = True Then
Buffer.ForeColor = cPer
Buffer.CurrentX = ScaleWidth \ 2 - Buffer.TextWidth(per & "%") \ 2
Buffer.CurrentY = ScaleHeight \ 2 - Buffer.TextHeight("|") \ 2
Buffer.Print per & "%"
End If

Board.Cls
ModCircleProg.BitBlt Board.hDC, 0, 0, ScaleWidth, ScaleHeight, Buffer.hDC, 0, 0, vbSrcCopy
End Function

Private Sub UserControl_Initialize()
DrawBar
End Sub

'- Properties -
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
PropBag.ReadProperty "PRAD"
PropBag.ReadProperty "PSHOWPER"
PropBag.ReadProperty "PVALUE"
PropBag.ReadProperty "PMAX"
PropBag.ReadProperty "CVALUE"
PropBag.ReadProperty "CNONVALUE"
PropBag.ReadProperty "CPER"
PropBag.ReadProperty "CBACK"

PropBag.ReadProperty "PDEPTH"
PropBag.ReadProperty "PMODE"
DrawBar
End Sub

Private Sub UserControl_Resize()
Buffer.Left = 0
Buffer.Top = 0
Buffer.Width = ScaleWidth
Buffer.Height = ScaleHeight

Board.Left = 0
Board.Top = 0
Board.Width = ScaleWidth
Board.Height = ScaleHeight

DrawBar
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "PRAD", pRad
PropBag.WriteProperty "PSHOWPER", pShowPer
PropBag.WriteProperty "PVALUE", pValue
PropBag.WriteProperty "PMAX", pMax
PropBag.WriteProperty "CVALUE", cValue
PropBag.WriteProperty "CNONVALUE", cNonValue
PropBag.WriteProperty "CPER", cPer
PropBag.WriteProperty "CBACK", cBack

PropBag.WriteProperty "PDEPTH", pDepth
PropBag.WriteProperty "PMODE", pMode
End Sub

'- Setting and getting of properties -

'Colors
Property Get BackColor() As OLE_COLOR
BackColor = cBack
End Property

Property Let BackColor(V As OLE_COLOR)
Board.BackColor = V
cBack = V
DrawBar
End Property

Property Get ValueColor() As OLE_COLOR
ValueColor = cValue
End Property

Property Let ValueColor(V As OLE_COLOR)
cValue = V
DrawBar
End Property

Property Get NonValueColor() As OLE_COLOR
NonValueColor = cNonValue
End Property

Property Let NonValueColor(V As OLE_COLOR)
cNonValue = V
DrawBar
End Property

Property Get CaptionColor() As OLE_COLOR
CaptionColor = cPer
End Property

Property Let CaptionColor(V As OLE_COLOR)
cPer = V
DrawBar
End Property

'Other

Property Get ShowCaption() As Boolean
ShowCaption = pShowPer
End Property

Property Let ShowCaption(V As Boolean)
pShowPer = V
DrawBar
End Property

Property Get Max() As Long
Max = pMax
End Property

Property Let Max(V As Long)
pMax = V
If pMax < pValue Then pMax = pValue
DrawBar
End Property

Property Get Value() As Long
Value = pValue
End Property

Property Let Value(V As Long)
pValue = V
If pValue > pMax Then pValue = pMax
DrawBar
End Property

Property Get Radius() As Integer
Radius = pRad
End Property

Property Let Radius(V As Integer)
pRad = V
If pRad < 2 Then pRad = 2
If pRad > 50 Then pRad = 50
DrawBar
End Property

Property Get Depth() As Integer
Depth = pDepth
End Property

Property Let Depth(V As Integer)
pDepth = V
If pDepth < 2 Then pDepth = 2
If pDepth > 45 Then pDepth = 45
End Property

Property Get Is3D() As Boolean
Is3D = pMode
End Property

Property Let Is3D(V As Boolean)
pMode = V
DrawBar
End Property
