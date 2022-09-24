VERSION 5.00
Begin VB.Form FDestroyer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Application Title"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main Form.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   255
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   458
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   480
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   4560
      Top             =   3360
   End
   Begin VB.TextBox Judul 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "*** Application Title ***"
      Top             =   120
      Width           =   6615
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      ToolTipText     =   "Keluar dari program"
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      ToolTipText     =   "Tentang Brontok Destroyer"
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdHajar 
      Caption         =   "Hajar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Melumpuhkan Virus dari memory dan memperbaiki Registry"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Info 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2775
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   480
      Width           =   6615
   End
   Begin VB.ListBox ListProses 
      Height          =   2400
      ItemData        =   "Main Form.frx":08CA
      Left            =   240
      List            =   "Main Form.frx":08CC
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   6375
   End
End
Attribute VB_Name = "FDestroyer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ******************
' MyDoom Destroyer
' By RoNz
' Juni 2007
' ******************

Dim i As Integer

Private Sub Form_Load()

    ' *** Inisialisasi Variabel ***
    
    Me.Caption = Nama_Aplikasi & " By RoNz"
    Judul.Text = "*** " & Nama_Aplikasi & " " & Copyright & " ***"
    
    RefreshDaftarWindow Me, ListProses
    
End Sub

Private Sub cmdHajar_Click()
    
    Hajar_Virus
    
End Sub

Private Sub cmdKeluar_Click()

    End
    
End Sub

Private Sub cmdAbout_Click()

MsgBox Nama_Aplikasi & vbCrLf & _
      "" & vbCrLf & _
      Copyright & vbCrLf & _
      "Email: RoNz_327@Yahoo.Com", vbInformation

End Sub

Private Sub Judul_Click()

    RefreshDaftarWindow FDestroyer, ListProses
    
End Sub

Private Sub Judul_DblClick()

    Info.Visible = False
    ListProses.Visible = True
    
End Sub

Private Sub ListProses_DblClick()

    Info.Visible = True
    ListProses.Visible = False
    
End Sub

Private Sub Timer1_Timer()
    
    If i = 0 Then
        Judul.Text = "*** " & Nama_Aplikasi & " " & Copyright & " ***"
    ElseIf i = 1 Then
        Judul.Text = "*** #CyBeRz@Allnetwork.org ***"
    ElseIf i = 2 Then
        Judul.Text = "*** FeeLCoMz Community ***"
    ElseIf i = 3 Then
        Judul.Text = "*** Sistem Informasi, Filkom, UPI ""YPTK""  ***"
        i = -1
    End If
    
    i = i + 1

End Sub
