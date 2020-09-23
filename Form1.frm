VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HU Card "
   ClientHeight    =   4890
   ClientLeft      =   2520
   ClientTop       =   1530
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   326
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   463
   Begin Threed.SSFrame LiteFrame 
      Height          =   480
      Left            =   4665
      TabIndex        =   68
      Top             =   4335
      Width           =   2220
      _Version        =   65536
      _ExtentX        =   3916
      _ExtentY        =   847
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowStyle     =   1
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "RX"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   2
         Left            =   1500
         TabIndex        =   71
         Top             =   225
         Width           =   195
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TX"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   1
         Left            =   855
         TabIndex        =   70
         Top             =   225
         Width           =   180
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Index           =   0
         Left            =   105
         TabIndex        =   69
         Top             =   225
         Width           =   285
      End
      Begin VB.Image RXLITE 
         Appearance      =   0  'Flat
         Height          =   210
         Left            =   1740
         Picture         =   "Form1.frx":0000
         Top             =   195
         Width           =   210
      End
      Begin VB.Image TXLITE 
         Appearance      =   0  'Flat
         Height          =   210
         Left            =   1080
         Picture         =   "Form1.frx":02AA
         Top             =   195
         Width           =   210
      End
      Begin VB.Image PORTLITE 
         Appearance      =   0  'Flat
         Height          =   210
         Left            =   435
         Picture         =   "Form1.frx":0554
         Top             =   195
         Width           =   210
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   675
      Left            =   7365
      TabIndex        =   66
      Top             =   45
      Width           =   1170
      _Version        =   65536
      _ExtentX        =   2064
      _ExtentY        =   1191
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Image RXOFF 
         Height          =   210
         Left            =   855
         Picture         =   "Form1.frx":07FE
         Top             =   390
         Width           =   210
      End
      Begin VB.Image TXOFF 
         Height          =   210
         Left            =   495
         Picture         =   "Form1.frx":0AA8
         Top             =   390
         Width           =   210
      End
      Begin VB.Image PortOFF 
         Height          =   210
         Left            =   120
         Picture         =   "Form1.frx":0D52
         Top             =   375
         Width           =   210
      End
      Begin VB.Image RXON 
         Height          =   210
         Left            =   840
         Picture         =   "Form1.frx":0FFC
         Top             =   120
         Width           =   210
      End
      Begin VB.Image TXON 
         Height          =   210
         Left            =   480
         Picture         =   "Form1.frx":12A6
         Top             =   120
         Width           =   210
      End
      Begin VB.Image PortON 
         Height          =   210
         Left            =   120
         Picture         =   "Form1.frx":1550
         Top             =   120
         Width           =   210
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Tech"
      Height          =   375
      Left            =   2220
      TabIndex        =   63
      Top             =   3180
      Width           =   735
   End
   Begin VB.TextBox R02Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Left            =   5355
      TabIndex        =   62
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox BYTESsentText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Left            =   6360
      TabIndex        =   60
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "DEMO"
      Height          =   375
      Left            =   1500
      TabIndex        =   59
      Top             =   3180
      Width           =   675
   End
   Begin VB.Frame Frame1 
      Caption         =   "PPV Purchases"
      Height          =   2685
      Left            =   210
      TabIndex        =   26
      Top             =   345
      Width           =   4635
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   24
         Left            =   3600
         TabIndex        =   51
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   23
         Left            =   2760
         TabIndex        =   50
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   22
         Left            =   1920
         TabIndex        =   49
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   21
         Left            =   1080
         TabIndex        =   48
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   20
         Left            =   240
         TabIndex        =   47
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   19
         Left            =   3600
         TabIndex        =   46
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   18
         Left            =   2760
         TabIndex        =   45
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   17
         Left            =   1920
         TabIndex        =   44
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   16
         Left            =   1080
         TabIndex        =   43
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   15
         Left            =   240
         TabIndex        =   42
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   14
         Left            =   3600
         TabIndex        =   41
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   13
         Left            =   2760
         TabIndex        =   40
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   12
         Left            =   1920
         TabIndex        =   39
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   1080
         TabIndex        =   38
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   240
         TabIndex        =   37
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   3600
         TabIndex        =   36
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   2760
         TabIndex        =   35
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   1920
         TabIndex        =   34
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   1080
         TabIndex        =   33
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   240
         TabIndex        =   32
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   3600
         TabIndex        =   31
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   2760
         TabIndex        =   30
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   1920
         TabIndex        =   29
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1080
         TabIndex        =   28
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "PPV #21  PPV #22   PPV #23  PPV #24  PPV #25"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Index           =   3
         Left            =   255
         TabIndex        =   56
         Top             =   2160
         Width           =   3960
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "PPV #16  PPV #17   PPV #18  PPV #19  PPV #20"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Index           =   2
         Left            =   270
         TabIndex        =   55
         Top             =   1680
         Width           =   3960
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "PPV #11  PPV #12   PPV #13  PPV #14  PPV #15"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Index           =   1
         Left            =   255
         TabIndex        =   54
         Top             =   1200
         Width           =   3960
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "PPV #06  PPV #07   PPV #08  PPV #09  PPV #10"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Index           =   0
         Left            =   255
         TabIndex        =   53
         Top             =   720
         Width           =   3960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "PPV #01  PPV #02   PPV #03  PPV #04  PPV #05"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   120
         Left            =   255
         TabIndex        =   52
         Top             =   240
         Width           =   3960
      End
   End
   Begin VB.TextBox SPENDINGLIMITtext 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5520
      TabIndex        =   24
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox RATINGtext 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5520
      TabIndex        =   22
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox GUIDEtext 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5520
      TabIndex        =   18
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox TIMEZONEtext 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox IRDText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5520
      TabIndex        =   16
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox USWtext 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5520
      TabIndex        =   13
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox FUSEtext 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5520
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox CardIDtext 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   480
      Width           =   1215
   End
   Begin VB.ComboBox COMMlist 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Form1.frx":17FA
      Left            =   3015
      List            =   "Form1.frx":17FC
      TabIndex        =   8
      Text            =   "Select COM Port"
      Top             =   3180
      Width           =   1815
   End
   Begin VB.TextBox BuffCntText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Left            =   4665
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   4080
      Width           =   180
   End
   Begin VB.TextBox TextInReadBuffer 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton CARDinfoBtn 
      Caption         =   "Get Card Info"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3180
      Width           =   1215
   End
   Begin VB.TextBox COMtext 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   5370
      Width           =   615
   End
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label HickLabel1 
      AutoSize        =   -1  'True
      Caption         =   "Copyright 2001"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   465
      TabIndex        =   65
      Top             =   4515
      Width           =   1695
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Now viewing realtime card responses   -    "
      Height          =   195
      Left            =   915
      TabIndex        =   64
      Top             =   4110
      Width           =   2985
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Sent:"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   5910
      TabIndex        =   61
      Top             =   4140
      Width           =   450
   End
   Begin VB.Label StatusLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   225
      TabIndex        =   58
      Top             =   3645
      Width           =   4620
   End
   Begin VB.Label HickLabel2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Hickware CREW"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2310
      TabIndex        =   57
      Top             =   4515
      Width           =   2055
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "LIMIT:"
      Height          =   195
      Left            =   5055
      TabIndex        =   25
      Top             =   3000
      Width           =   465
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "RATING:"
      Height          =   195
      Left            =   4860
      TabIndex        =   23
      Top             =   2640
      Width           =   660
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "TIME:"
      Height          =   195
      Left            =   5085
      TabIndex        =   21
      Top             =   2280
      Width           =   435
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "GUIDE:"
      Height          =   195
      Left            =   4965
      TabIndex        =   20
      Top             =   1920
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IRD:"
      Height          =   195
      Left            =   5190
      TabIndex        =   19
      Top             =   840
      Width           =   330
   End
   Begin VB.Label USWlabel 
      AutoSize        =   -1  'True
      Caption         =   "USW:"
      Height          =   195
      Left            =   5085
      TabIndex        =   15
      Top             =   1200
      Width           =   435
   End
   Begin VB.Label FuseLabel 
      AutoSize        =   -1  'True
      Caption         =   "FUSE:"
      Height          =   195
      Left            =   5055
      TabIndex        =   14
      Top             =   1560
      Width           =   465
   End
   Begin VB.Label CardIDlabel 
      AutoSize        =   -1  'True
      Caption         =   "CAMID:"
      Height          =   195
      Left            =   4965
      TabIndex        =   11
      Top             =   480
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "INS:"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   4980
      TabIndex        =   9
      Top             =   4155
      Width           =   360
   End
   Begin VB.Label BuufCntLabel 
      AutoSize        =   -1  'True
      Caption         =   "BytesIn:"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   3915
      TabIndex        =   7
      Top             =   4155
      Width           =   720
   End
   Begin VB.Label Label3 
      Caption         =   " Text in COMM buffer for DEBUG"
      Height          =   585
      Left            =   150
      TabIndex        =   5
      Top             =   5235
      Width           =   1320
   End
   Begin VB.Label ATRlabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   225
      TabIndex        =   1
      Top             =   30
      Width           =   6510
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "     Visual Basic Smartcard Interface       Release 2"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   5010
      TabIndex        =   67
      Top             =   3270
      Width           =   1920
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CardIDtext_GotFocus()

On Error Resume Next
 CARDinfoBtn.SetFocus
Exit Sub


End Sub

Private Sub Cardinfobtn_Click()

CardInserted = True

ShowStatus "clearing fields"

Call ToggleButtons

Call ClearVariables

 For xxx = 0 To 24
  Form1.Text1(xxx).Text = ""
 Next xxx

If port = "" Then
   MsgBox "No COM port selected!"
   Call CloseCOMM
   Exit Sub
End If

ShowStatus "resetting card"
Call ResetForWrite

DelaySecs 0.25

If AtrLen <> 59 Then Call CloseCOMM: Call ToggleButtons: Exit Sub

ShowStatus "sending data"
Call SendData(CardInfoStr)

ShowStatus "reading data"
Call ReadDATA
ShowStatus "parsing data"
Call ShowDATA: Call CardInfo2A(CardInfoBuffer)
Call ClearVariables
ShowStatus "sending data"
Call SendData(IRDinfoStr)
ShowStatus "reading data"
Call ReadDATA
ShowStatus "displaying data"
Call ShowDATA: Call CardInfo58(CardInfoBuffer)
ShowStatus "sending data"
Call SendData(PPVinfoStr)
ShowStatus "reading data"
Call ReadDATA
ShowStatus "displaying PPV data":
Call ShowDATA: Call CardInfoPPV(CardInfoBuffer)
ShowStatus "closing comport"
Call CloseCOMM

Call ToggleButtons

ShowStatus "done"

CardInserted = False

End Sub

Private Sub Command1_Click()
Dim k
Dim tempy

MsgBox "This will simulate writing to the card. It does not actually write because the packets are not HU compatible, they are simply here to show you HOW to do it.", 0, "NOTE"
Cardinfobtn_Click

Call ToggleButtons

Command2_Click

ShowStatus "clearing fields"

If port = "" Then
   MsgBox "No COM port selected!"
   Call CloseCOMM:
   Exit Sub
End If

ShowStatus "resetting card"

Call ResetForWrite

DelaySecs 0.25

If AtrLen <> 59 Then Call CloseCOMM: Call ToggleButtons: Exit Sub


ShowStatus "sending data"
CheckINS = True
'-----------------------------------------------------------
'The following shows how to send WRITE packets to the card
'This is only demonstrative packets, you have to do your own
'real packets to change the contents of the eeprom. This is
'here only to show you it CAN be done with VB. There IS some
'encryption that will have to be done in order to actually
'write to an HU card. We are not here to do everything for
'you, you have to do this yourself. The hard part has been
'done here for you already.
'-----------------------------------------------------------
'send packet header        check INS       do a slight delay
Call SendData(PWpacket01): Call GetReturn: DelaySecs (0.25)
'send packet data          check ACK       do a slight delay
Call SendData(PWpacket02): Call GetReturn: DelaySecs (0.25)
'etc
Call SendData(PWpacket03): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket04): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket05): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket06): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket07): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket08): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket09): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket10): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket11): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket12): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket13): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket14): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket15): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket16): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket17): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket18): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket19): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket20): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket21): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket22): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket23): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket24): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket25): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket26): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket27): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket28): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket29): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket30): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket31): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket32): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket33): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket34): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket35): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket36): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket37): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket38): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket39): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket40): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket41): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket42): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket43): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket44): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket45): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket46): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket47): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket48): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket49): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket50): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket51): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket52): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket53): Call GetReturn: DelaySecs (0.25)
Call SendData(PWpacket54): Call GetReturn: DelaySecs (0.25)

ShowStatus "reading data"
Call ReadDATA

Call ToggleButtons

Call CloseCOMM

ShowStatus "done"

End Sub

Private Sub Command2_Click()

If Form1.Height = Hite Then
  Form1.Height = 5300
  Command2.Caption = "Hide"
  Exit Sub
End If

If Form1.Height = 5300 Then
  Form1.Height = Hite
  Command2.Caption = "Tech"
  Exit Sub
End If
 
End Sub

Private Sub COMMlist_Click()

If port > "" Then Call CloseCOMM

port = Form1.COMMlist.Text

COMtext.Text = port

Call CheckCOM(port)
 DelaySecs 0.25

End Sub

Private Sub Form_Load()

'Use ascii to make it hard to hex edit our text`s

titleA$ = Chr$(72) + Chr$(85) + Chr$(32) + _
          Chr$(67) + Chr$(97) + Chr$(114) + Chr$(100)
titleB$ = Chr$(32)
titleC$ = Chr$(85) + Chr$(116) + Chr$(105) + Chr$(108) + _
          Chr$(105) + Chr$(116) + Chr$(121)
titleD$ = Chr$(32) + Chr$(32) + Chr$(32) + Chr$(32) + _
          Chr$(32) + Chr$(32) + Chr$(32) + Chr$(32) + _
          Chr$(32) + Chr$(32) + Chr$(32) + Chr$(32) + _
          Chr$(32) + Chr$(32) + Chr$(32) + Chr$(32) + _
          Chr$(32) + Chr$(32) + Chr$(32) + Chr$(32) + _
          Chr$(32) + Chr$(32) + Chr$(32) + Chr$(32) + _
          Chr$(32) + Chr$(32) + _
          Chr$(118) + Chr$(50) + Chr$(46) + Chr$(48)
titleE$ = titleA$ + titleB$ + titleC$ + titleD$

Form1.Caption = titleE$
Form1.HickLabel1.Caption = Chr$(67) + Chr$(111) + Chr$(112) + Chr$(121) + _
                           Chr$(114) + Chr$(105) + Chr$(103) + Chr$(104) + Chr$(116) + _
                           Chr$(32) + Chr$(50) + Chr$(48) + Chr$(48) + Chr$(49)
Form1.HickLabel2.Caption = Chr$(84) + Chr$(104) + Chr$(101) + Chr$(32) + _
                           Chr$(72) + Chr$(105) + Chr$(99) + Chr$(107) + _
                           Chr$(119) + Chr$(97) + Chr$(114) + Chr$(101) + Chr$(32) + _
                           Chr$(67) + Chr$(82) + Chr$(69) + Chr$(87)

Form1.Height = Hite

COMMlist.AddItem "COM1"
COMMlist.AddItem "COM2"
COMMlist.AddItem "COM3"
COMMlist.AddItem "COM4"
    
Me.Show

Call GetState

If port > "" Then
   Select Case port
    Case Is = "COM1"
     COMMlist.Text = COMMlist.List(0)
     Call ToggleButtons:
    Case Is = "COM2"
     COMMlist.Text = COMMlist.List(1)
     Call ToggleButtons:
    Case Is = "COM3"
     COMMlist.Text = COMMlist.List(2)
     Call ToggleButtons:
    Case Is = "COM4"
     COMMlist.Text = COMMlist.List(3)
     Call ToggleButtons:
    Case Else
    
  End Select
End If
  
 Call CheckCOM(port)
 
End Sub

Private Sub Form_Unload(Cancel As Integer)

Call SaveState

Call CloseCOMM
DelaySecs 0.0025

Unload Me
End

End Sub

Private Sub R02Label_Change()

If Len(R02Label.Text) > 2 Then
   Label5.Caption = "ACK:"
 Else
  Label5.Caption = "INS:"
End If

End Sub
