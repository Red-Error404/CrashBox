VERSION 5.00
Begin VB.Form TestWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CrashBox.Test"
   ClientHeight    =   7930
   ClientLeft      =   30
   ClientTop       =   370
   ClientWidth     =   7040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   793
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   704
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2650
      Left            =   240
      ScaleHeight     =   265
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   649
      TabIndex        =   5
      Top             =   4920
      Width           =   6490
      Begin VB.CommandButton Command2 
         Caption         =   "可视化"
         Height          =   370
         Left            =   5400
         TabIndex        =   6
         Top             =   2040
         Width           =   850
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "传统"
      Height          =   370
      Left            =   5640
      TabIndex        =   4
      Top             =   240
      Width           =   850
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   250
      Left            =   2400
      TabIndex        =   3
      Text            =   "0.5"
      Top             =   240
      Width           =   2890
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   3970
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   720
      Width           =   6490
   End
   Begin VB.CommandButton TestBtn 
      Caption         =   "高效"
      Height          =   370
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   850
   End
   Begin VB.Label Label1 
      Caption         =   "精确度："
      Height          =   250
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Width           =   1210
   End
End
Attribute VB_Name = "TestWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Rect
    x As Long
    y As Long
    w As Long
    h As Long
End Type
Dim rt() As Rect
Dim MapW As Long, MapH As Long

Private Sub Command1_Click()
    '王者铜出来挨打
    Dim Count As Long, t As Long, t2 As Long
    
    t = GetTickCount
    For i = 1 To UBound(rt)
        For s = 1 To UBound(rt)
            If i <> s Then
                If (rt(i).x >= rt(s).x And rt(i).x <= rt(s).x + rt(s).w And rt(i).y >= rt(s).y And rt(i).y <= rt(s).y + rt(s).h) Or _
                    (rt(i).x + rt(i).w >= rt(s).x And rt(i).x + rt(i).w <= rt(s).x + rt(s).w And rt(i).y >= rt(s).y And rt(i).y <= rt(s).y + rt(s).h) Or _
                    (rt(i).x >= rt(s).x And rt(i).x <= rt(s).x + rt(s).w And rt(i).y + rt(i).h >= rt(s).y And rt(i).y + rt(i).h <= rt(s).y + rt(s).h) Or _
                    (rt(i).x + rt(i).w >= rt(s).x And rt(i).x + rt(i).w <= rt(s).x + rt(s).w And rt(i).y + rt(i).h >= rt(s).y And rt(i).y + rt(i).h <= rt(s).y + rt(s).h) Then
                    Count = Count + 1
                    Exit For
                End If
            End If
        Next
    Next
    t2 = GetTickCount - t

    Dim text As String
    text = UBound(rt) & "个矩形碰撞检测(使用传统的两重循环)，耗时：" & t2 & "ms" & vbCrLf & "一共有" & Count & "个矩形与另一个矩形相撞(包括被撞的矩形)。"
    
    Text1.text = Text1.text & vbCrLf & vbCrLf & Now & vbCrLf & text
End Sub

Private Sub Command2_Click()
    Picture1.Cls
    
    Randomize
    MapW = Picture1.ScaleWidth: MapH = Picture1.ScaleHeight
    
    ReDim rt(10)
    For i = 1 To UBound(rt)
        With rt(i)
            .x = Int(Rnd * MapW)
            .w = 50
            .y = Int(Rnd * MapH)
            .h = 50
            Picture1.Line (.x, .y)-(.x + .w, .y + .h), RGB(Rnd * 255, 0, 0), BF
            Picture1.ForeColor = RGB(255, 255, 255)
            Picture1.CurrentX = .x: Picture1.CurrentY = .y
            Picture1.Print i
        End With
    Next
    
    Picture1.Refresh
End Sub

Private Sub Form_Load()
    MapW = 7000
    MapH = 4000
    
    Randomize
    ReDim rt(10000)
    For i = 1 To UBound(rt)
        With rt(i)
            .x = Int(Rnd * MapW)
            .w = 50
            .y = Int(Rnd * MapH)
            .h = 50
        End With
    Next
End Sub

Private Sub TestBtn_Click()
    Dim Crash As New GCrashBox
    Dim t As Long, i As Integer, t2 As Long, r As Integer
    
    Crash.Reset MapW, MapH, Val(Text2.text)
    
    Dim Count As Long
    t = GetTickCount
    For i = 1 To UBound(rt)
        r = Crash.CheckCrashRect(rt(i).x, rt(i).y, rt(i).w, rt(i).h, i, 0, False)
        If r <> 0 Then Count = Count + 1
    Next
    t2 = GetTickCount - t
    
    Dim text As String
    text = UBound(rt) & "个矩形碰撞检测，精度：" & Val(Text2.text) & "(可能造成" & Int(1 / Val(Text2.text) - 1) & "pixel的偏差)，耗时：" & t2 & "ms" & vbCrLf & "一共有" & Count & "个矩形与另一个矩形相撞（公共被碰撞的矩形被忽略）。"
    
    Text1.text = Text1.text & vbCrLf & vbCrLf & Now & vbCrLf & text
End Sub
