VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animasi teks menggunakan objek timer"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      Height          =   2115
      Left            =   0
      ScaleHeight     =   2055
      ScaleWidth      =   5595
      TabIndex        =   1
      Top             =   0
      Width           =   5655
      Begin VB.Timer tmrAnimasi 
         Enabled         =   0   'False
         Index           =   0
         Interval        =   35
         Left            =   3120
         Top             =   600
      End
      Begin VB.Timer tmrAnimasi 
         Enabled         =   0   'False
         Index           =   1
         Interval        =   35
         Left            =   3600
         Top             =   600
      End
      Begin VB.Timer tmrAnimasi 
         Enabled         =   0   'False
         Index           =   2
         Interval        =   35
         Left            =   4080
         Top             =   600
      End
      Begin VB.Timer tmrAnimasi 
         Enabled         =   0   'False
         Index           =   3
         Interval        =   35
         Left            =   4560
         Top             =   600
      End
      Begin VB.PictureBox picPanel 
         Height          =   495
         Left            =   240
         ScaleHeight     =   435
         ScaleWidth      =   4875
         TabIndex        =   8
         Top             =   120
         Width           =   4935
         Begin VB.Label lblAnimasi 
            Caption         =   "Animasi teks menggunakan objek timer"
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
            Left            =   960
            TabIndex        =   9
            Top             =   120
            Width           =   3315
         End
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdKeluar 
         Caption         =   "Keluar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   1560
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Caption         =   " [ Pilihan ] "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   4935
         Begin VB.OptionButton optAnimasi 
            Caption         =   "Animasi 2"
            Height          =   195
            Index           =   1
            Left            =   1440
            TabIndex        =   5
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optAnimasi 
            Caption         =   "Animasi 1"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   0
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optAnimasi 
            Caption         =   "Animasi 3"
            Height          =   195
            Index           =   2
            Left            =   2640
            TabIndex        =   3
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optAnimasi 
            Caption         =   "Animasi 4"
            Height          =   195
            Index           =   3
            Left            =   3840
            TabIndex        =   4
            Top             =   360
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************
' MMMM  MMMMM  OMMM   MMMO    OMMM    OMMM    OMMMMO     OMMMMO    OMMMMO  '
'  MM    MM   MM MM    MMMO  OMMM    MM MM    MM   MO   OM    MO  OM    MO '
'  MM  MM    MM  MM    MM  OO  MM   MM  MM    MM   MO   OM    MO       OMO '
'  MMMM     MMMMMMMM   MM  MM  MM  MMMMMMMM   MMMMMO     OMMMMO      OMO   '
'  MM  MM        MM    MM      MM       MM    MM   MO   OM    MO   OMO     '
'  MM    MM      MM    MM      MM       MM    MM    MO  OM    MO  OM   MM  '
' MMMM  MMMM    MMMM  MMMM    MMMM     MMMM  MMMM  MMMM  OMMMMO   MMMMMMM  '
'                                                                          '
' K4m4r82's Laboratory                                                     '
' http://coding4ever.wordpress.com                                         '
'***************************************************************************

Dim x           As Integer
Dim GerakX      As Integer
Dim GerakY      As Integer
Dim ZigZagX     As Integer
Dim ZigZagY     As Integer
Dim i           As Integer

Private Sub cmdStart_Click()
    If cmdStart.Caption = "Start" Then
        For i = optAnimasi.LBound To optAnimasi.UBound
            If optAnimasi(i).Value = True Then tmrAnimasi(i).Enabled = True
        Next
        
        Frame1.Enabled = False
        cmdStart.Caption = "Stop"
      
    Else
        For i = optAnimasi.LBound To optAnimasi.UBound
            tmrAnimasi(i).Enabled = False
        Next
        Frame1.Enabled = True
        cmdStart.Caption = "Start"
        lblAnimasi.Left = 1080
        lblAnimasi.Top = 120
    End If
End Sub

Private Sub cmdKeluar_Click()
    End
End Sub

Private Sub Form_Load()
    GerakY = 20
    x = 20
    ZigZagX = 20
    ZigZagY = 20
End Sub

Private Sub tmrAnimasi_Timer(Index As Integer)
    Select Case Index
        Case 0
            lblAnimasi.Move lblAnimasi.Left - 15
            If lblAnimasi.Left < -lblAnimasi.Width Then lblAnimasi.Left = picPanel.Width
         
        Case 1
'         lblAnimasi.Move lblAnimasi.Left + x
'         If lblAnimasi.Left < picPanel.ScaleLeft Then
'            x = 20
'         ElseIf lblAnimasi.Left + lblAnimasi.Width > picPanel.ScaleWidth + picPanel.ScaleLeft Then
'            x = -20
'         End If

            lblAnimasi.Move lblAnimasi.Left - x
            If lblAnimasi.Left < picPanel.ScaleLeft Then
                x = -20
            ElseIf lblAnimasi.Left + lblAnimasi.Width > picPanel.ScaleWidth + picPanel.ScaleLeft Then
                x = 20
            End If
         
        Case 2
            lblAnimasi.Move lblAnimasi.Left + GerakX, lblAnimasi.Top + GerakY
            If lblAnimasi.Top < picPanel.ScaleTop Then
                GerakY = 20
            ElseIf lblAnimasi.Top + lblAnimasi.Height > picPanel.ScaleHeight + picPanel.ScaleTop Then
                GerakY = -20
            End If
         
        Case 3
            lblAnimasi.Move lblAnimasi.Left + ZigZagX, lblAnimasi.Top + ZigZagY
            If lblAnimasi.Left < picPanel.ScaleLeft Then
                ZigZagX = 20
            ElseIf lblAnimasi.Left + lblAnimasi.Width > picPanel.ScaleWidth + picPanel.ScaleLeft Then
                ZigZagX = -20
            ElseIf lblAnimasi.Top < picPanel.ScaleTop Then
                ZigZagY = 20
            ElseIf lblAnimasi.Top + lblAnimasi.Height > picPanel.ScaleHeight + picPanel.ScaleTop Then
                ZigZagY = -20
            End If
    End Select
End Sub
