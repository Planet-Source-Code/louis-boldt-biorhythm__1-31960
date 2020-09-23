VERSION 5.00
Begin VB.Form BioRhythm 
   AutoRedraw      =   -1  'True
   Caption         =   "BioRhythm"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   684
   ClientWidth     =   7356
   FillStyle       =   0  'Solid
   Icon            =   "biorhythm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   7356
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   492
      Left            =   3360
      TabIndex        =   7
      Top             =   3840
      Width           =   852
   End
   Begin VB.Frame FraReportDate 
      Caption         =   "Report Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   4320
      TabIndex        =   48
      Top             =   3720
      Width           =   2892
      Begin VB.HScrollBar hsbRptYr 
         Height          =   225
         Left            =   2040
         Max             =   3000
         Min             =   1900
         TabIndex        =   5
         Top             =   492
         Value           =   1900
         Width           =   672
      End
      Begin VB.HScrollBar hsbRptMo 
         Height          =   225
         LargeChange     =   3
         Left            =   576
         Max             =   13
         TabIndex        =   3
         Top             =   492
         Value           =   1
         Width           =   435
      End
      Begin VB.HScrollBar hsbRptDy 
         Height          =   225
         LargeChange     =   3
         Left            =   1452
         Max             =   32
         TabIndex        =   4
         Top             =   492
         Value           =   1
         Width           =   435
      End
      Begin VB.Label lblRptYr 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1960"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2040
         TabIndex        =   51
         Top             =   240
         Width           =   672
      End
      Begin VB.Label lblRptMo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Septembre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   50
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label lblRptDy 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1452
         TabIndex        =   49
         Top             =   240
         Width           =   432
      End
   End
   Begin VB.Frame fraBirth 
      Caption         =   "Birthday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   240
      TabIndex        =   45
      Top             =   3720
      Width           =   2892
      Begin VB.HScrollBar hsbBirthDy 
         Height          =   225
         LargeChange     =   3
         Left            =   1440
         Max             =   32
         TabIndex        =   1
         Top             =   480
         Value           =   1
         Width           =   435
      End
      Begin VB.HScrollBar hsbBirthMo 
         Height          =   225
         LargeChange     =   3
         Left            =   480
         Max             =   13
         TabIndex        =   0
         Top             =   492
         Value           =   1
         Width           =   435
      End
      Begin VB.HScrollBar hsbBirthYr 
         Height          =   225
         LargeChange     =   3
         Left            =   2040
         Max             =   3000
         Min             =   1900
         TabIndex        =   2
         Top             =   480
         Value           =   1900
         Width           =   672
      End
      Begin VB.Label lblBirthDy 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1441
         TabIndex        =   47
         Top             =   240
         Width           =   432
      End
      Begin VB.Label lblBirthYr 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1960"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2040
         TabIndex        =   46
         Top             =   240
         Width           =   672
      End
      Begin VB.Label lblBirthMo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Septembre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   91
         TabIndex        =   10
         Top             =   240
         Width           =   1212
      End
   End
   Begin VB.CommandButton cmdCreateChart 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Create Chart"
      Height          =   492
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   852
   End
   Begin VB.PictureBox picChartBio 
      BackColor       =   &H00C0E0FF&
      FillColor       =   &H00C0C0C0&
      ForeColor       =   &H00E0E0E0&
      Height          =   2655
      Left            =   144
      ScaleHeight     =   200
      ScaleMode       =   0  'User
      ScaleWidth      =   29
      TabIndex        =   8
      Top             =   120
      Width           =   7095
      Begin VB.Line Line20 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.185
         Y1              =   190.015
         Y2              =   190.015
      End
      Begin VB.Line Line19 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.185
         Y1              =   180.031
         Y2              =   180.031
      End
      Begin VB.Line Line18 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29
         Y1              =   170.046
         Y2              =   170.046
      End
      Begin VB.Line Line16 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.185
         Y1              =   159.985
         Y2              =   159.985
      End
      Begin VB.Line Line15 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.185
         Y1              =   150
         Y2              =   150
      End
      Begin VB.Line Line14 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.185
         Y1              =   139.939
         Y2              =   139.939
      End
      Begin VB.Line Line13 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.185
         Y1              =   130.031
         Y2              =   130.031
      End
      Begin VB.Line Line12 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.185
         Y1              =   119.969
         Y2              =   119.969
      End
      Begin VB.Line Line11 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.185
         Y1              =   109.985
         Y2              =   109.985
      End
      Begin VB.Line Line10 
         BorderColor     =   &H000080FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.185
         Y1              =   90.015
         Y2              =   90.015
      End
      Begin VB.Line Line9 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.185
         Y1              =   80.031
         Y2              =   80.031
      End
      Begin VB.Line Line8 
         BorderColor     =   &H000080FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.185
         Y1              =   69.969
         Y2              =   69.969
      End
      Begin VB.Line Line7 
         BorderColor     =   &H0000FF00&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.185
         Y1              =   59.985
         Y2              =   59.985
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000080FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.185
         Y1              =   50
         Y2              =   50
      End
      Begin VB.Line Line5 
         BorderColor     =   &H0080FF80&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.185
         Y1              =   40.015
         Y2              =   40.015
      End
      Begin VB.Line Line4 
         BorderColor     =   &H008080FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29.185
         Y1              =   29.954
         Y2              =   29.954
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0080FF80&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29
         Y1              =   20.046
         Y2              =   20.046
      End
      Begin VB.Line Line1 
         BorderColor     =   &H008080FF&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   29
         Y1              =   9.985
         Y2              =   9.985
      End
      Begin VB.Line lnVert 
         X1              =   14.344
         X2              =   14.344
         Y1              =   0
         Y2              =   203.456
      End
      Begin VB.Line lnHor 
         X1              =   0
         X2              =   29.185
         Y1              =   101.767
         Y2              =   101.767
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   21
         X1              =   28.691
         X2              =   28.691
         Y1              =   0
         Y2              =   194.24
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   20
         X1              =   26.645
         X2              =   26.645
         Y1              =   0
         Y2              =   194.24
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   19
         X1              =   24.636
         X2              =   24.636
         Y1              =   0
         Y2              =   194.24
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   18
         X1              =   22.623
         X2              =   22.623
         Y1              =   0
         Y2              =   194.24
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   17
         X1              =   20.61
         X2              =   20.61
         Y1              =   0
         Y2              =   194.24
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   16
         X1              =   18.601
         X2              =   18.601
         Y1              =   0
         Y2              =   194.24
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   12
         X1              =   12.565
         X2              =   12.565
         Y1              =   0
         Y2              =   194.24
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   11
         X1              =   10.552
         X2              =   10.552
         Y1              =   0
         Y2              =   194.24
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   10
         X1              =   8.543
         X2              =   8.543
         Y1              =   0
         Y2              =   194.24
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   9
         X1              =   6.53
         X2              =   6.53
         Y1              =   0
         Y2              =   194.24
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   8
         X1              =   4.516
         X2              =   4.516
         Y1              =   0
         Y2              =   194.24
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   7
         X1              =   2.507
         X2              =   2.507
         Y1              =   0
         Y2              =   194.24
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   6
         X1              =   0.494
         X2              =   0.494
         Y1              =   0
         Y2              =   203.456
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   0
         X1              =   14.5
         X2              =   14.5
         Y1              =   0
         Y2              =   203.456
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   17
         Index           =   1
         X1              =   16.587
         X2              =   16.587
         Y1              =   0
         Y2              =   194.24
      End
   End
   Begin VB.Label lblAge 
      Alignment       =   2  'Center
      Caption         =   "Age"
      Height          =   252
      Left            =   1680
      TabIndex        =   52
      Top             =   4680
      Width           =   4092
   End
   Begin VB.Line linColorIntuition 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   6
      X1              =   5880
      X2              =   6960
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line linColorEmotion 
      BorderColor     =   &H00C00000&
      BorderWidth     =   6
      X1              =   4560
      X2              =   5520
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line linColorIntellect 
      BorderColor     =   &H00008000&
      BorderWidth     =   6
      X1              =   1680
      X2              =   2760
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line linColorPhysical 
      BorderColor     =   &H000000C0&
      BorderWidth     =   6
      X1              =   480
      X2              =   1440
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "Caution is indicated for activities on both Critical and Zero days."
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   2490
      TabIndex        =   44
      Top             =   7320
      Width           =   2370
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   $"biorhythm.frx":030A
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   495
      TabIndex        =   43
      Top             =   6840
      Width           =   6375
   End
   Begin VB.Line Line17 
      BorderStyle     =   3  'Dot
      X1              =   132
      X2              =   7212
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label5"
      Height          =   252
      Index           =   28
      Left            =   6972
      TabIndex        =   42
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Label5"
      Height          =   252
      Index           =   27
      Left            =   6744
      TabIndex        =   41
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label5"
      Height          =   252
      Index           =   26
      Left            =   6504
      TabIndex        =   40
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Label5"
      Height          =   252
      Index           =   25
      Left            =   6252
      TabIndex        =   39
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label5"
      Height          =   252
      Index           =   24
      Left            =   6012
      TabIndex        =   38
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Label5"
      Height          =   252
      Index           =   23
      Left            =   5772
      TabIndex        =   37
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label5"
      Height          =   252
      Index           =   22
      Left            =   5520
      TabIndex        =   36
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Label5"
      Height          =   252
      Index           =   21
      Left            =   5280
      TabIndex        =   35
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label5"
      Height          =   252
      Index           =   20
      Left            =   5040
      TabIndex        =   34
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Label5"
      Height          =   252
      Index           =   19
      Left            =   4788
      TabIndex        =   33
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label5"
      Height          =   252
      Index           =   18
      Left            =   4548
      TabIndex        =   32
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Label5"
      Height          =   252
      Index           =   17
      Left            =   4308
      TabIndex        =   31
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label5"
      Height          =   252
      Index           =   16
      Left            =   4044
      TabIndex        =   30
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Label5"
      Height          =   252
      Index           =   15
      Left            =   3804
      TabIndex        =   29
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label5"
      Height          =   252
      Index           =   14
      Left            =   3564
      TabIndex        =   28
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Label5"
      Height          =   252
      Index           =   13
      Left            =   3312
      TabIndex        =   27
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label5"
      Height          =   252
      Index           =   12
      Left            =   3072
      TabIndex        =   26
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Label5"
      Height          =   252
      Index           =   11
      Left            =   2832
      TabIndex        =   25
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label5"
      Height          =   252
      Index           =   10
      Left            =   2580
      TabIndex        =   24
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Label5"
      Height          =   252
      Index           =   9
      Left            =   2340
      TabIndex        =   23
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label5"
      Height          =   252
      Index           =   8
      Left            =   2100
      TabIndex        =   22
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Label5"
      Height          =   252
      Index           =   7
      Left            =   1848
      TabIndex        =   21
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label5"
      Height          =   252
      Index           =   6
      Left            =   1608
      TabIndex        =   20
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Label5"
      Height          =   252
      Index           =   5
      Left            =   1368
      TabIndex        =   19
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label5"
      Height          =   252
      Index           =   4
      Left            =   1104
      TabIndex        =   18
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Label5"
      Height          =   252
      Index           =   3
      Left            =   864
      TabIndex        =   17
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label5"
      Height          =   252
      Index           =   2
      Left            =   624
      TabIndex        =   16
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Label5"
      Height          =   252
      Index           =   1
      Left            =   372
      TabIndex        =   15
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label lblDy 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label5"
      Height          =   252
      Index           =   0
      Left            =   132
      TabIndex        =   14
      Top             =   2880
      Width           =   252
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "Intuitional"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   5880
      TabIndex        =   13
      Top             =   3240
      Width           =   972
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "Emotional"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   4560
      TabIndex        =   12
      Top             =   3240
      Width           =   972
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "Intellectual"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   1800
      TabIndex        =   11
      Top             =   3240
      Width           =   972
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "Physical"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   480
      TabIndex        =   9
      Top             =   3240
      Width           =   972
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "BioRhythm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Based on the core algorithm from a previous biorhythm
'program posted on Planet Source Code.

'Biorhythm study and use is considered a "pseudoscience"
'in the United State however it is widely accepted and
'utilized throughout Europe and much of the rest of the
'world as a valid tool for guaging personal capabilities
'and performance based on a group of cycles that begin
'on the day of birth and extend throughout life.

' This version was
' Plagiarize by Louis Boldt then changed somewhat.

' Code for the date selection was adapted
' from the Biorhythm posted by Auteur:Les Productions J.F.

' Code for Chart adapted from Michael Hebert at runes2@concentric.net
Option Explicit
Private masMoName(1 To 12) As String * 11

' --------------------------------------------------------------------
Private Sub Form_Load()
' --------------------------------------------------------------------
  'Set ScaleWidth and ScaleHeight of picture box
  picChartBio.ScaleWidth = 29 '14 days each side of report date
  picChartBio.ScaleHeight = 200 'Makes it easier to calc %
    
  'Makes Horizontal line in the middle of the picture
  lnHor.Y1 = picChartBio.ScaleHeight / 2
  lnHor.Y2 = lnHor.Y1
  lnHor.X1 = 0
  lnHor.X2 = picChartBio.ScaleWidth
  
  'Makes Vertical line in the middle of the picture
  lnVert.Y1 = 0
  lnVert.Y2 = picChartBio.ScaleHeight
  lnVert.X1 = picChartBio.ScaleWidth / 2
  lnVert.X2 = picChartBio.ScaleWidth / 2
  
  'Sets Labels to the current date
  
  Call Init_Date_Selection
  Call SetReportDateLabels(Date)
  Me.Show
  cmdCreateChart_Click
End Sub

' --------------------------------------------------------------------
Private Sub cmdExit_click()
' --------------------------------------------------------------------
 
  mnuFileExit_Click
 
End Sub
' --------------------------------------------------------------------
Private Sub cmdCreateChart_Click()
' --------------------------------------------------------------------
  Dim dblCurrentPoint As Double
  Dim intBioPeriod As Integer
  Dim cBioColor As ColorConstants
  Dim intBioType As Integer
  Dim datBirth As Date
  Dim datReport As Date
  Dim iYY As Integer
  Dim iMM As Integer
  Dim iDD As Integer
  'Clear the Screen
  picChartBio.Cls
  If Not IsDate(lblBirthMo.Caption & " " & lblBirthDy.Caption & ", " & lblBirthYr.Caption) Then
    MsgBox "Birth date is not valid." & vbNewLine & "" & vbNewLine & Trim$(lblBirthMo.Caption) & ", " & Trim$(lblBirthYr.Caption) & " does not have that" & vbNewLine & "many days.", 48, "Nut's n Boldt's Software"
    Exit Sub
  End If
  If Not IsDate(lblRptMo.Caption & " " & lblRptDy.Caption & ", " & lblRptYr.Caption) Then
    MsgBox "Report date is not valid." & vbNewLine & "" & vbNewLine & Trim$(lblRptMo.Caption) & ", " & Trim$(lblRptYr.Caption) & " does not have that" & vbNewLine & "many days.", 48, "Nut's n Boldt's Software"
    Exit Sub
  End If

  'Set Birth and report Dates
  datBirth = lblBirthMo.Caption & " " & lblBirthDy.Caption & ", " & lblBirthYr.Caption
  datReport = lblRptMo.Caption & " " & lblRptDy.Caption & ", " & lblRptYr.Caption
  SetReportDateLabels (datReport)
  
  lblAge.Caption = fnAge(datBirth, datReport)
  'Loop through biorhythm cycles and plot them
  'intBioPeriod is the number of days in a cycle
  'cBioColor is the color of the line to plot
  For intBioType = 1 To 4
    
    If intBioType = 1 Then 'Physical period
      intBioPeriod = 23
      cBioColor = vbRed
     End If
    
    If intBioType = 2 Then 'Intellectual period
       intBioPeriod = 33
       cBioColor = vbGreen
     End If
   
    If intBioType = 3 Then 'Emotional period
       intBioPeriod = 28
       cBioColor = vbBlue
     End If
    
    If intBioType = 4 Then 'Intuitional period
       intBioPeriod = 38
       cBioColor = vbCyan
     End If
     
    'Find the first peak to left of middle
    dblCurrentPoint = picChartBio.ScaleWidth / 2 - ((DateDiff("d", datBirth, datReport) Mod intBioPeriod) - intBioPeriod / 4)
    'Find the first peak that is off the chart
    Do While dblCurrentPoint > 0
      dblCurrentPoint = dblCurrentPoint - intBioPeriod
    Loop
        
    'Necessary because next loop add intBioPeriod/2 back to the variable
    dblCurrentPoint = dblCurrentPoint - intBioPeriod / 2
    
    'Find high and low points and plot parabolas
    Do While dblCurrentPoint < picChartBio.ScaleWidth
      dblCurrentPoint = dblCurrentPoint + intBioPeriod / 2
      If dblCurrentPoint + intBioPeriod / 4 >= 0 Then
        Parabola intBioType, dblCurrentPoint, 0, dblCurrentPoint + intBioPeriod / 4, cBioColor
      End If
        dblCurrentPoint = dblCurrentPoint + intBioPeriod / 2
      If dblCurrentPoint + intBioPeriod / 4 >= 0 Then
        Parabola intBioType, dblCurrentPoint, picChartBio.ScaleHeight, dblCurrentPoint + intBioPeriod / 4, cBioColor
      End If
    Loop
    
  Next intBioType

End Sub

' --------------------------------------------------------------------
Private Function Parabola(intBioType As Integer, dblXa As Double, dblYa As Double, dblLastXa As Double, RedGreenBlueCyan As ColorConstants)
' --------------------------------------------------------------------
    
  'This function creates a parabola when vertex and last point are given
  'Vertex = dblXa, dblYa
  'LastPoint = dblLastXa, intHorzCenter
  Dim intHorzCenter As Integer
  Dim dblSlope As Double
  Dim dblY As Double
  Dim dblX As Double
  picChartBio.DrawWidth = 4
  'Set intHorzCenter to be the horizontal center of the biorhythm picture
  intHorzCenter = picChartBio.ScaleHeight / 2
  'Find Slope of parabola
  dblSlope = (intHorzCenter - dblYa) / ((dblLastXa - dblXa) ^ 2)
  'Graph the parabola
  For dblX = (dblXa - (dblLastXa - dblXa)) To dblLastXa Step 0.01
    dblY = dblSlope * ((dblX - dblXa) ^ 2) + dblYa
    picChartBio.PSet (dblX, dblY), RedGreenBlueCyan
  Next dblX

End Function



' The following subs control the display date
' when the horizontal scroll bars are pressed
' --------------------------------------------------------------------
Private Sub hsbRptMo_Change()
' --------------------------------------------------------------------
  If hsbRptMo.Value = 13 Then
    hsbRptMo.Value = 1
  ElseIf hsbRptMo.Value = 0 Then
    hsbRptMo.Value = 12
  End If
  lblRptMo.Caption = masMoName(hsbRptMo.Value)
 End Sub
' --------------------------------------------------------------------
Private Sub hsbRptDy_Change()
' --------------------------------------------------------------------
  If hsbRptDy.Value = 32 Then
    hsbRptDy.Value = 1
  ElseIf hsbRptDy.Value = 0 Then
    hsbRptDy.Value = 31
  End If
  lblRptDy.Caption = Str$(hsbRptDy.Value)
End Sub

' --------------------------------------------------------------------
Private Sub hsbRptYr_Change()
' --------------------------------------------------------------------
  lblRptYr.Caption = Str$(hsbRptYr.Value)
End Sub
' --------------------------------------------------------------------
Private Sub hsbBirthMo_Change()
' --------------------------------------------------------------------
   If hsbBirthMo.Value = 13 Then
    hsbBirthMo.Value = 1
  ElseIf hsbBirthMo.Value = 0 Then
    hsbBirthMo.Value = 12
  End If
  lblBirthMo.Caption = masMoName(hsbBirthMo.Value)
End Sub

' --------------------------------------------------------------------
Private Sub hsbBirthDy_Change()
' --------------------------------------------------------------------
   If hsbBirthDy.Value = 32 Then
    hsbBirthDy.Value = 1
  ElseIf hsbBirthDy.Value = 0 Then
    hsbBirthDy.Value = 31
  End If
 
  lblBirthDy.Caption = Str$(hsbBirthDy.Value)
End Sub
' --------------------------------------------------------------------
Private Sub hsbBirthYr_Change()
' --------------------------------------------------------------------
  lblBirthYr.Caption = Str$(hsbBirthYr.Value)
End Sub
 ' --------------------------------------------------------------------
Private Sub Init_Date_Selection()
' --------------------------------------------------------------------
 
  masMoName(1) = "January"
  masMoName(2) = "February"
  masMoName(3) = "March"
  masMoName(4) = "April"
  masMoName(5) = "May"
  masMoName(6) = "June"
  masMoName(7) = "July"
  masMoName(8) = "August"
  masMoName(9) = "September"
  masMoName(10) = "October"
  masMoName(11) = "November"
  masMoName(12) = "December"
   
  ' Get the from birthday from the regestry
  hsbBirthDy.Value = GetSetting("BioRhythm", "Birth", "Day", 2)
  hsbBirthMo.Value = GetSetting("BioRhythm", "Birth", "Month", 3)
  hsbBirthYr.Value = GetSetting("BioRhythm", "Birth", "Year", 1950)
  
  lblBirthDy.Caption = Str$(hsbBirthDy.Value)
  lblBirthYr.Caption = Str$(hsbBirthYr.Value)
  lblBirthMo.Caption = masMoName(hsbBirthMo.Value)
  hsbRptDy.Value = Day(Now)
  hsbRptMo.Value = Month(Now)
  hsbRptYr.Value = Year(Now)
  lblRptDy.Caption = Str$(hsbRptDy.Value)
  lblRptYr.Caption = Str$(hsbRptYr.Value)
  lblRptMo.Caption = masMoName(hsbRptMo.Value)
End Sub
' --------------------------------------------------------------------
Private Function SetReportDateLabels(datReport As Date)
' --------------------------------------------------------------------
  Dim datWork As Date
  Dim iNdx As Integer

  'Sets datwork to the report date
  
  datWork = datReport
  'Fill Center dayLabel with Current Day
  lblDy(14) = DatePart("d", datWork)
  
  'Fill dayLabels - Left Side of Chart
  For iNdx = 13 To 0 Step -1
    datWork = datWork - 1
    lblDy(iNdx) = DatePart("d", datWork)
  Next iNdx
  
  'Reset datWork to Current System Date
  datWork = datReport
  
  'Fill dayLabels - Right Side of Chart
  For iNdx = 15 To 28 Step 1
    datWork = datWork + 1
    lblDy(iNdx) = DatePart("d", datWork)
  Next iNdx
     
End Function

' --------------------------------------------------------------------
Private Sub mnuAbout_Click()
' --------------------------------------------------------------------
  frmAbout.Show vbModal
End Sub
' --------------------------------------------------------------------
Private Sub mnuFileExit_Click()
' --------------------------------------------------------------------
'  save  the from birthday  to the regestry
  SaveSetting "BioRhythm", "Birth", "Day", CStr(hsbBirthDy.Value)
  SaveSetting "BioRhythm", "Birth", "Month", CStr(hsbBirthMo.Value)
  SaveSetting "BioRhythm", "Birth", "Year", CStr(hsbBirthYr.Value)

  Unload Me
  End
End Sub
