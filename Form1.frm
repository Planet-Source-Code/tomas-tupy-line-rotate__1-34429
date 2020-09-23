VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Mouse"
      Height          =   4185
      Left            =   3090
      TabIndex        =   2
      Top             =   285
      Width           =   2730
      Begin VB.Line Line3 
         BorderWidth     =   5
         X1              =   1365
         X2              =   2010
         Y1              =   1545
         Y2              =   900
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4200
      Left            =   240
      TabIndex        =   0
      Top             =   285
      Width           =   2685
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   150
         Top             =   690
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   285
         Left            =   630
         TabIndex        =   1
         Top             =   375
         Width           =   1365
      End
      Begin VB.Line Line1 
         X1              =   1020
         X2              =   1860
         Y1              =   1275
         Y2              =   1275
      End
      Begin VB.Line Line2 
         X1              =   345
         X2              =   2025
         Y1              =   3000
         Y2              =   3000
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PIE = 3.14159265 / 12
Private Th As Single
Private R1 As Single
Private R2 As Single
Private Lx1 As Single
Private Ly1 As Single
Private Lx2 As Single
Private Ly2 As Single
Private Lx3 As Single
Private Ly3 As Single
Private Sub Command1_Click()
Timer1.Enabled = True
End Sub


Private Sub Form_Load()
Dim dx As Single
Dim dy As Single

    dx = Line1.X2 - Line1.X1
    dy = Line1.Y2 - Line1.Y1
    R1 = Sqr(dx * dx + dy * dy)
    Lx1 = Line1.X1
    Ly1 = Line1.Y1
    
    dx = Line2.X2 - Line2.X1
    dy = Line2.Y2 - Line2.Y1
    R2 = Sqr(dx * dx + dy * dy) / 2
    Lx2 = (Line2.X1 + Line2.X2) / 2
    Ly2 = (Line2.Y1 + Line2.Y2) / 2
   
    dx = Line3.X2 - Line3.X1
    dy = Line3.Y2 - Line3.Y1
    R1 = Sqr(dx * dx + dy * dy)
    Lx3 = Line3.X1
    Ly3 = Line3.Y1

End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Th = Y / 300
Line3.X2 = Lx3 + Cos(Th) * R1
Line3.Y2 = Ly3 + Sin(Th) * R1
End Sub

Private Sub Timer1_Timer()
    Th = Th + PIE

    Line1.X2 = Lx1 + Cos(Th) * R1
    Line1.Y2 = Ly1 + Sin(Th) * R1

    Line2.X1 = Lx2 + Cos(Th) * R2
    Line2.Y1 = Ly2 + Sin(Th) * R2
    Line2.X2 = Lx2 - Cos(Th) * R2
    Line2.Y2 = Ly2 - Sin(Th) * R2
End Sub
