VERSION 5.00
Object = "{1C574F76-C2F9-4951-88C5-714FA56A9E03}#1.0#0"; "UltraContainer.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Ultracontainer Demo"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15945
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   9660
   ScaleWidth      =   15945
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin UltraContainer.UltraPictureBox upcEnterData 
      Height          =   6135
      Left            =   4050
      TabIndex        =   5
      Top             =   210
      Width           =   5385
      _extentx        =   9499
      _extenty        =   10821
      font            =   "Form1.frx":5D54E
      fontname        =   "MS Sans Serif"
      fontsize        =   8,25
      scalemode       =   3
      bordercolor     =   32768
      showtitle       =   -1  'True
      titlebackcolorfrom=   16384
      titlebackcolorto=   32768
      titlefontsize   =   12
      titlefontbold   =   -1  'True
      titlefontcolor  =   8454143
      showbackgroundgradient=   -1  'True
      backgroundgradientfrom=   16384
      backgroundgradientto=   12648447
      titleheight     =   32
      titlecaption    =   "Enter your data here"
      transparencydistance=   100
      showmirror      =   -1  'True
      mirrorpercent   =   50
      fadeenabled     =   -1  'True
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   465
         Left            =   2190
         TabIndex        =   13
         Top             =   2550
         Width           =   1305
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   465
         Left            =   3600
         TabIndex        =   12
         Top             =   2550
         Width           =   1305
      End
      Begin VB.TextBox txtComment 
         Appearance      =   0  'Flat
         Height          =   1065
         Left            =   1410
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "Form1.frx":5D57A
         Top             =   1410
         Width           =   3495
      End
      Begin VB.TextBox txtLastName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1410
         TabIndex        =   8
         Text            =   "Costinu"
         Top             =   990
         Width           =   3495
      End
      Begin VB.TextBox txtFirstName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1410
         TabIndex        =   6
         Text            =   "Costinu"
         Top             =   570
         Width           =   3495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Comment"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   150
         TabIndex        =   11
         Top             =   1410
         Width           =   1035
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Last name"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   150
         TabIndex        =   9
         Top             =   1020
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "First name"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   150
         TabIndex        =   7
         Top             =   600
         Width           =   1035
      End
   End
   Begin UltraContainer.UltraPictureBox upcMyData 
      Height          =   3165
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   3465
      _extentx        =   6112
      _extenty        =   5583
      backcolor       =   12632256
      font            =   "Form1.frx":5D5D0
      fontname        =   "Arial"
      fontsize        =   9,75
      scalemode       =   3
      showborder      =   -1  'True
      bordercolor     =   16697774
      showtitle       =   -1  'True
      titlebackcolorfrom=   8388608
      titlebackcolorto=   12282702
      titlefontbold   =   -1  'True
      titlefontcolor  =   16777215
      showbackgroundgradient=   -1  'True
      backgroundgradientfrom=   5451523
      backgroundgradientto=   12282702
      titlecaption    =   "This is my data"
      transparencydistance=   180
      fadeenabled     =   -1  'True
      Begin VB.Label lblMyData 
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1935
         Left            =   210
         TabIndex        =   1
         Top             =   510
         Width           =   3120
         WordWrap        =   -1  'True
      End
   End
   Begin UltraContainer.UltraPictureBox UltraPictureBox1 
      Height          =   5205
      Left            =   180
      TabIndex        =   2
      Top             =   3510
      Width           =   3465
      _extentx        =   6112
      _extenty        =   9181
      backcolor       =   12632256
      font            =   "Form1.frx":5D5F4
      fontname        =   "Arial"
      fontsize        =   9,75
      scalemode       =   3
      showborder      =   -1  'True
      bordercolor     =   16697774
      roundshape      =   -1  'True
      showtitle       =   -1  'True
      titlebackcolorfrom=   64
      titlebackcolorto=   192
      titlefontbold   =   -1  'True
      titlefontcolor  =   16777215
      showbackgroundgradient=   -1  'True
      backgroundgradientfrom=   128
      backgroundgradientto=   12632319
      titlecaption    =   "Round shape"
      transparencydistance=   180
      fadeenabled     =   -1  'True
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
         Height          =   465
         Left            =   1860
         TabIndex        =   4
         Top             =   4650
         Width           =   1455
      End
      Begin VB.ListBox List1 
         Height          =   3960
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   3195
      End
   End
   Begin UltraContainer.UltraPictureBox upcMyPics 
      Height          =   8535
      Left            =   9780
      TabIndex        =   14
      Top             =   210
      Width           =   5235
      _extentx        =   9234
      _extenty        =   15055
      backcolor       =   12632256
      font            =   "Form1.frx":5D618
      fontname        =   "Arial"
      fontsize        =   9,75
      showborder      =   -1  'True
      bordercolor     =   16576
      roundshape      =   -1  'True
      showtitle       =   -1  'True
      titlebackcolorfrom=   16512
      titlebackcolorto=   8438015
      titlefontbold   =   -1  'True
      titlefontcolor  =   16777215
      showbackgroundgradient=   -1  'True
      backgroundgradientfrom=   33023
      backgroundgradientto=   8438015
      backgroundgradientdirection=   1
      titlecaption    =   "My pictures"
      transparencydistance=   180
      Begin UltraContainer.UltraPictureBox UltraPicShow 
         Height          =   2835
         Left            =   2760
         TabIndex        =   20
         Top             =   540
         Width           =   2205
         _extentx        =   3889
         _extenty        =   5001
         font            =   "Form1.frx":5D63C
         fontname        =   "MS Sans Serif"
         fontsize        =   8,25
         picture         =   "Form1.frx":5D668
         scalemode       =   3
         bordercolor     =   32768
         transparencydistance=   210
         showmirror      =   -1  'True
         mirrorpercent   =   50
      End
      Begin UltraContainer.UltraPictureBox upcMyPic 
         Height          =   1485
         Index           =   4
         Left            =   300
         TabIndex        =   19
         Top             =   6900
         Width           =   2205
         _extentx        =   3889
         _extenty        =   2619
         font            =   "Form1.frx":62562
         fontname        =   "MS Sans Serif"
         fontsize        =   8,25
         picture         =   "Form1.frx":6258E
         scalemode       =   3
         showborder      =   -1  'True
         bordercolor     =   32768
         transparencydistance=   230
         mirrorpercent   =   30
         fadeenabled     =   -1  'True
      End
      Begin UltraContainer.UltraPictureBox upcMyPic 
         Height          =   1485
         Index           =   3
         Left            =   300
         TabIndex        =   18
         Top             =   5310
         Width           =   2205
         _extentx        =   3889
         _extenty        =   2619
         font            =   "Form1.frx":6ED64
         fontname        =   "MS Sans Serif"
         fontsize        =   8,25
         picture         =   "Form1.frx":6ED90
         scalemode       =   3
         showborder      =   -1  'True
         bordercolor     =   32768
         transparencydistance=   230
         mirrorpercent   =   30
         fadeenabled     =   -1  'True
      End
      Begin UltraContainer.UltraPictureBox upcMyPic 
         Height          =   1485
         Index           =   2
         Left            =   300
         TabIndex        =   17
         Top             =   3720
         Width           =   2205
         _extentx        =   3889
         _extenty        =   2619
         font            =   "Form1.frx":7423E
         fontname        =   "MS Sans Serif"
         fontsize        =   8,25
         picture         =   "Form1.frx":7426A
         scalemode       =   3
         showborder      =   -1  'True
         bordercolor     =   32768
         transparencydistance=   230
         mirrorpercent   =   30
         fadeenabled     =   -1  'True
      End
      Begin UltraContainer.UltraPictureBox upcMyPic 
         Height          =   1485
         Index           =   1
         Left            =   300
         TabIndex        =   16
         Top             =   2130
         Width           =   2205
         _extentx        =   3889
         _extenty        =   2619
         font            =   "Form1.frx":7BECC
         fontname        =   "MS Sans Serif"
         fontsize        =   8,25
         picture         =   "Form1.frx":7BEF8
         scalemode       =   3
         showborder      =   -1  'True
         bordercolor     =   32768
         transparencydistance=   230
         mirrorpercent   =   30
         fadeenabled     =   -1  'True
      End
      Begin UltraContainer.UltraPictureBox upcMyPic 
         Height          =   1485
         Index           =   0
         Left            =   300
         TabIndex        =   15
         Top             =   540
         Width           =   2205
         _extentx        =   3889
         _extenty        =   2619
         font            =   "Form1.frx":8129A
         fontname        =   "MS Sans Serif"
         fontsize        =   8,25
         picture         =   "Form1.frx":812C6
         scalemode       =   3
         showborder      =   -1  'True
         bordercolor     =   32768
         transparencydistance=   230
         mirrorpercent   =   30
         fadeenabled     =   -1  'True
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "<-- Go with mouse over the thumbnail pictures"
         Height          =   675
         Left            =   2820
         TabIndex        =   21
         Top             =   3690
         Width           =   1995
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim LastUsedIndex As Long

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    UltraPictureBox1.RedrawMirror
End Sub

Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    UltraPictureBox1.RedrawMirror
End Sub

Private Sub cmdClear_Click()
    txtComment.Text = ""
    txtFirstName.Text = ""
    txtLastName.Text = ""
End Sub

Private Sub cmdClear_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    upcEnterData.RedrawMirror
End Sub

Private Sub cmdClear_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    upcEnterData.RedrawMirror
End Sub

Private Sub cmdSave_Click()
    lblMyData.Caption = "First name: " & txtFirstName.Text & vbCrLf & _
                "Last name: " & txtLastName.Text & vbCrLf & _
                "Comment: " & txtComment.Text & vbCrLf
    
    upcMyData.Update
End Sub

Private Sub cmdSave_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    upcEnterData.RedrawMirror
End Sub

Private Sub cmdSave_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    upcEnterData.RedrawMirror
End Sub





Private Sub Form_Activate()
    upcEnterData.RedrawMirror
    upcMyPics.Update
    upcMyPics.SetTransparency 180
    UltraPicShow.Update
End Sub

Private Sub Text1_Change()
    UltraPictureBox1.RedrawMirror
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    upcMyData.Update
    upcEnterData.Update

    
    For i = 1 To 25
        List1.AddItem "List item #" & i
    Next i
    
    cmdSave_Click
    
End Sub

Private Sub txtComment_Change()
    upcEnterData.RedrawMirror
End Sub

Private Sub txtFirstName_Change()
    upcEnterData.RedrawMirror
End Sub

Private Sub txtLastName_Change()
    upcEnterData.RedrawMirror
End Sub


Private Sub upcMyPic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If LastUsedIndex <> Index Then
        LastUsedIndex = Index
        Set UltraPicShow.Picture = upcMyPic(Index).Image
        UltraPicShow.Update
    End If
End Sub
