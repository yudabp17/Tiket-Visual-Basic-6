VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEMESANAN TIKET PESAWAT DOMESTIK"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   8415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   3480
         TabIndex        =   4
         Top             =   0
         Width           =   2175
         Begin VB.Label lbltanggal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.Timer Timer1 
         Left            =   0
         Top             =   7920
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1095
         Begin VB.Label lbljam 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.Image Image5 
         Height          =   975
         Left            =   1155
         Picture         =   "HALAMAN UTAMA.frx":0000
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   3375
      End
      Begin VB.Image Image4 
         Height          =   975
         Left            =   1155
         Picture         =   "HALAMAN UTAMA.frx":4EFD
         Stretch         =   -1  'True
         Top             =   4755
         Width           =   3375
      End
      Begin VB.Image Image3 
         Height          =   975
         Left            =   1155
         Picture         =   "HALAMAN UTAMA.frx":198F6
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   3375
      End
      Begin VB.Image Image2 
         Height          =   975
         Left            =   1155
         Picture         =   "HALAMAN UTAMA.frx":26BBD
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   3375
      End
      Begin VB.Image Image1 
         Height          =   975
         Left            =   120
         Picture         =   "HALAMAN UTAMA.frx":31E97
         Stretch         =   -1  'True
         Top             =   7320
         Width           =   5415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "=== PILIH MASKAPAI ==="
         BeginProperty Font 
            Name            =   "Adobe Caslon Pro"
            Size            =   18
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   1
         Top             =   1200
         Width           =   4335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sHari As String
Dim aHari


Private Sub Form_Load()
  aHari = Array("Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu")
  Timer1.Interval = 500
  Timer1.Enabled = True
End Sub

Private Sub Image2_Click()
Form2.Show
Unload Me
End Sub

Private Sub Image3_Click()
Form3.Show
Unload Me
End Sub

Private Sub Image4_Click()
Form4.Show
Unload Me
End Sub

Private Sub Image5_Click()
Form5.Show
Unload Me
End Sub

Private Sub Timer1_Timer()
  sHari = aHari(Abs(Weekday(Date) - 1))
  lbltanggal.Caption = Format(Date, "dd mmmm yyyy")
lbljam.Caption = Format(Time, "hh:mm:ss")
End Sub
