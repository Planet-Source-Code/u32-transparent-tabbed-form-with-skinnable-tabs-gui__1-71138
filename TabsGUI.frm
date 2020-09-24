VERSION 5.00
Begin VB.Form frmTabsGUI 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTab 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Index           =   2
      Left            =   2040
      ScaleHeight     =   2715
      ScaleWidth      =   4155
      TabIndex        =   12
      Top             =   960
      Width           =   4215
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TAB 3 child control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   2400
         Width           =   1710
      End
   End
   Begin VB.PictureBox picTab 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Index           =   1
      Left            =   2040
      ScaleHeight     =   2715
      ScaleWidth      =   4155
      TabIndex        =   10
      Top             =   960
      Width           =   4215
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TAB 2 child control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   2400
         Width           =   1710
      End
   End
   Begin VB.PictureBox picTab 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Index           =   3
      Left            =   2040
      ScaleHeight     =   2715
      ScaleWidth      =   4155
      TabIndex        =   15
      Top             =   960
      Width           =   4215
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TAB 4 child control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   16
         Top             =   2400
         Width           =   1710
      End
   End
   Begin VB.PictureBox picTab 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Index           =   0
      Left            =   2040
      Picture         =   "TabsGUI.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   4155
      TabIndex        =   8
      Top             =   960
      Width           =   4215
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TAB 1 child control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   9
         Top             =   2400
         Width           =   1710
      End
   End
   Begin VB.Label lblSkin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sunken button"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   21
      ToolTipText     =   "Click..."
      Top             =   3960
      Width           =   1050
   End
   Begin VB.Label lblSkin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "edged creamy"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   20
      ToolTipText     =   "Click..."
      Top             =   3720
      Width           =   1005
   End
   Begin VB.Image Go 
      Height          =   300
      Left            =   6380
      Top             =   55
      Width           =   300
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5440
      TabIndex        =   19
      Top             =   4290
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3650
      TabIndex        =   18
      Top             =   4290
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2140
      TabIndex        =   17
      Top             =   4290
      Width           =   480
   End
   Begin VB.Image applyb 
      Height          =   360
      Left            =   4800
      Top             =   5040
      Width           =   1080
   End
   Begin VB.Image applya 
      Height          =   360
      Left            =   4800
      Top             =   4920
      Width           =   1080
   End
   Begin VB.Image apply 
      Height          =   390
      Left            =   1680
      Top             =   4200
      Width           =   1620
   End
   Begin VB.Label lblTab 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tab 4"
      Height          =   195
      Index           =   3
      Left            =   600
      TabIndex        =   14
      Top             =   2190
      Width           =   435
   End
   Begin VB.Image imgTab 
      Height          =   390
      Index           =   3
      Left            =   210
      MousePointer    =   99  'Custom
      Top             =   2080
      Width           =   1260
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select skin"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   780
   End
   Begin VB.Label lblSkin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Swirl mix"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Click..."
      Top             =   3480
      Width           =   600
   End
   Begin VB.Label lblSkin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Green mix"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Click..."
      Top             =   3240
      Width           =   705
   End
   Begin VB.Label lblSkin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Flat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Click..."
      Top             =   3000
      Width           =   330
   End
   Begin VB.Label lblTab 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tab 3"
      Height          =   195
      Index           =   2
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   435
   End
   Begin VB.Label lblTab 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tab 2"
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   435
   End
   Begin VB.Label lblTab 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tab 1"
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   1070
      Width           =   420
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "  Skinned Form and TabDialog"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6255
   End
   Begin VB.Image cancelb 
      Height          =   360
      Left            =   3600
      Top             =   5040
      Width           =   1080
   End
   Begin VB.Image cancela 
      Height          =   360
      Left            =   3600
      Top             =   4920
      Width           =   1080
   End
   Begin VB.Image okb 
      Height          =   360
      Left            =   2400
      Top             =   5040
      Width           =   1080
   End
   Begin VB.Image oka 
      Height          =   360
      Left            =   2400
      Top             =   4920
      Width           =   1080
   End
   Begin VB.Image cancel 
      Height          =   390
      Left            =   3000
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   1950
   End
   Begin VB.Image ok 
      Height          =   390
      Left            =   4680
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   1620
   End
   Begin VB.Image closeMea 
      Height          =   300
      Left            =   1680
      Top             =   5040
      Width           =   300
   End
   Begin VB.Image closeMeb 
      Height          =   300
      Left            =   1680
      Top             =   5160
      Width           =   300
   End
   Begin VB.Image tabLow 
      Height          =   315
      Left            =   120
      Top             =   5160
      Width           =   1410
   End
   Begin VB.Image tabHigh 
      Height          =   420
      Left            =   120
      Top             =   4920
      Width           =   1410
   End
   Begin VB.Image imgTab 
      Height          =   390
      Index           =   2
      Left            =   210
      MousePointer    =   99  'Custom
      Top             =   1710
      Width           =   1260
   End
   Begin VB.Image imgTab 
      Height          =   390
      Index           =   1
      Left            =   210
      MousePointer    =   99  'Custom
      Top             =   1330
      Width           =   1260
   End
   Begin VB.Image imgTab 
      Height          =   390
      Index           =   0
      Left            =   210
      MousePointer    =   99  'Custom
      Top             =   960
      Width           =   1260
   End
End
Attribute VB_Name = "frmTabsGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Example on GUI. In this case.. Tabbed Dialog
' Made by u32. June, 2008.
' Anyone that have a quest. E-mail me through PSC.
' Thanks!

Private CurrentTab As Integer

Private Sub apply_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    apply.Picture = applyb
    ok.Picture = oka
    cancel.Picture = cancela
    
End Sub

Private Sub cancel_Click()
    
    End
    
End Sub

Private Sub cancel_MouseMove(Button As Integer, Shift As Integer, Leave As Single, Y As Single)
   
    cancel.Picture = cancelb
    apply.Picture = applya
    ok.Picture = oka
   
End Sub

Private Sub Form_Load()

    Dim Idx As Long
    Me.Height = 4815
    Me.Width = 6750
    
    ' Set some defaults
    For Idx = 0 To picTab.UBound
        picTab(Idx).Picture = picTab(0).Picture
        picTab(Idx).CurrentX = 360
        picTab(Idx).CurrentY = 200
        picTab(Idx).Print "Tab " & Idx + 1 & " Container"
    Next
    
    ' Select default skin
    lblSkin_Click 1
    
    ' Default tab
    imgTab_Click 3
    
    ' Run Tranp code
    Transparency Me, 255, 0, 0
   
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Go.Picture = closeMea
End Sub

Private Sub Go_Click()
    End
End Sub

Private Sub Go_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Go = closeMeb
End Sub

Private Sub imgTab_Click(Index As Integer)
   
    Dim I As Long
    
    ' Index what's the current propertys
    For I = 0 To imgTab.UBound
       If I = Index Then
          CurrentTab = Index
          imgTab(I).Picture = tabHigh
          imgTab(I).Left = 270
          lblTab(I).FontBold = True
          lblTab(I).ForeColor = &H84FF&
          picTab(I).ZOrder
       Else
          imgTab(I).Picture = tabLow
          imgTab(I).Left = 210
          lblTab(I).FontBold = False
          lblTab(I).ForeColor = vbBlack
       End If
    Next
    
End Sub

Private Sub Label5_Click()

    cancel_Click
    
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    MoveMe Me, lblTitle
   
End Sub

Private Sub lblSkin_Click(Index As Integer)
   
    Dim NR As Long
    
    ' Choose index by Nr. and name of the skin
    SkinForm Me, Choose(Index + 1, "Flat", "Green", "Swirl", "Edged", "Sunken")
    
    For NR = 0 To lblSkin.UBound
        ' Bold font when tabs is selected
       lblSkin(NR).FontBold = NR = Index
    Next
    
    ' Choose any control with Click event
    imgTab_Click CurrentTab
    ' Refresh when new skin is selected
    Transparency Me, 255, 0, 0
    LoadIn
    
End Sub

Private Sub lblTab_Click(Index As Integer)

    imgTab_Click Index
   
End Sub

Private Sub ok_MouseMove(Button As Integer, Shift As Integer, Leave As Single, Y As Single)
   
    ok.Picture = okb
    apply.Picture = applya
    cancel.Picture = cancela
    
End Sub

Private Sub LoadIn()

    ok.Picture = oka
    cancel.Picture = cancela
    apply.Picture = applya
    Go.Picture = closeMea
   
End Sub
