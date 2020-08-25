VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PEMESANAN TIKET LION AIR"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5910
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   8415
      Left            =   128
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.CommandButton Command2 
         Caption         =   "MENU UTAMA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   42
         Top             =   7680
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "HARGA"
         Height          =   375
         Left            =   3240
         TabIndex        =   34
         Top             =   3480
         Width           =   1095
      End
      Begin VB.ComboBox TAHUN 
         Height          =   315
         Left            =   4200
         TabIndex        =   33
         Top             =   4200
         Width           =   1215
      End
      Begin VB.ComboBox TGL 
         Height          =   315
         Left            =   1800
         TabIndex        =   32
         Top             =   4200
         Width           =   735
      End
      Begin VB.ComboBox BULAN 
         Height          =   315
         Left            =   2760
         TabIndex        =   31
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BAGASI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   30
         Top             =   6000
         Width           =   5175
         Begin VB.OptionButton Check3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "31 - 50 KG"
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
            Left            =   3480
            TabIndex        =   41
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton Check2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "16 - 30 KG"
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
            Left            =   1800
            TabIndex        =   40
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "0 - 15 KG"
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
            Left            =   240
            TabIndex        =   39
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "JUMLAH PENUMPANG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         TabIndex        =   23
         Top             =   4680
         Width           =   5175
         Begin VB.TextBox TxtDewasa 
            Height          =   285
            Left            =   1440
            TabIndex        =   25
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox TxtBayi 
            Height          =   285
            Left            =   1440
            TabIndex        =   24
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "USIA > 2 TAHUN"
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
            Left            =   3120
            TabIndex        =   29
            Top             =   360
            Width           =   1485
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "USIA 0 - 2 TAHUN"
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
            Left            =   3120
            TabIndex        =   28
            Top             =   720
            Width           =   1605
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DEWASA :"
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
            Left            =   360
            TabIndex        =   27
            Top             =   360
            Width           =   930
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BAYI :"
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
            Left            =   360
            TabIndex        =   26
            Top             =   720
            Width           =   555
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   5175
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DATA PEMESANAN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1200
            TabIndex        =   15
            Top             =   120
            Width           =   2865
         End
      End
      Begin VB.ComboBox CboTujuan 
         Height          =   315
         Left            =   435
         TabIndex        =   12
         Text            =   "Kota Tujuan"
         Top             =   3600
         Width           =   1695
      End
      Begin VB.ComboBox CboAsal 
         Height          =   315
         ItemData        =   "TIKET LION.frx":0000
         Left            =   435
         List            =   "TIKET LION.frx":0002
         TabIndex        =   10
         Text            =   "Kota Asal"
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   5175
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1800
            TabIndex        =   9
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1800
            TabIndex        =   7
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NO. TELEPON :"
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
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   1380
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NAMA PEMESAN :"
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
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1605
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   3
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
            TabIndex        =   4
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.Timer Timer1 
         Left            =   0
         Top             =   7920
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   3480
         TabIndex        =   1
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
            TabIndex        =   2
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "HARGA TIKET"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2400
         TabIndex        =   17
         Top             =   2640
         Width           =   3015
         Begin VB.Label TxtHarga 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Left            =   480
            TabIndex        =   19
            Top             =   360
            Width           =   120
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rp."
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
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   315
         End
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   5280
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Label txtTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   2040
         TabIndex        =   38
         Top             =   7080
         Width           =   120
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL BAYAR : Rp."
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
         Left            =   240
         TabIndex        =   37
         Top             =   7080
         Width           =   1770
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rp."
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
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   315
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   360
         TabIndex        =   35
         Top             =   0
         Width           =   120
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   22
         Top             =   4200
         Width           =   135
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   4200
         Width           =   135
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BERANGKAT :"
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
         Left            =   240
         TabIndex        =   20
         Top             =   4200
         Width           =   1515
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--- PILIH TUJUAN ---"
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
         Left            =   240
         TabIndex        =   16
         Top             =   2280
         Width           =   2115
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KOTA TUJUAN :"
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
         Left            =   555
         TabIndex        =   13
         Top             =   3360
         Width           =   1425
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KOTA ASAL :"
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
         Left            =   675
         TabIndex        =   11
         Top             =   2640
         Width           =   1155
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim harga, dewasa, bayi, hargabayi, subtotal, total As Variant
Dim sHari As String
Dim aHari

Private Sub Check1_Click()
dewasa = TxtDewasa.Text
bayi = TxtBayi.Text
harga = TxtHarga.Caption
hargabayi = 0.15 * harga
bagasi = 30000
subtotal = (dewasa * harga) + (hargabayi * bayi)
total = subtotal + bagasi
txtTotal.Caption = total
End Sub

Private Sub Check2_Click()
dewasa = TxtDewasa.Text
bayi = TxtBayi.Text
harga = TxtHarga.Caption
hargabayi = 0.15 * harga
bagasi = 60000
subtotal = (dewasa * harga) + (hargabayi * bayi)
total = subtotal + bagasi
txtTotal.Caption = total
End Sub

Private Sub Check3_Click()
dewasa = TxtDewasa.Text
bayi = TxtBayi.Text
harga = TxtHarga.Caption
hargabayi = 0.15 * harga
bagasi = 90000
subtotal = (dewasa * harga) + (hargabayi * bayi)
total = subtotal + bagasi
txtTotal.Caption = total
End Sub

Private Sub Command1_Click()
If CboAsal.ListIndex = 0 And CboTujuan.ListIndex = 1 Then
TxtHarga.Caption = 630000
ElseIf CboAsal.ListIndex = 0 And CboTujuan.ListIndex = 2 Then
TxtHarga.Caption = 535000
ElseIf CboAsal.ListIndex = 1 And CboTujuan.ListIndex = 0 Then
TxtHarga.Caption = 750000
ElseIf CboAsal.ListIndex = 1 And CboTujuan.ListIndex = 2 Then
TxtHarga.Caption = 450000
ElseIf CboAsal.ListIndex = 2 And CboTujuan.ListIndex = 0 Then
TxtHarga.Caption = 575000
ElseIf CboAsal.ListIndex = 2 And CboTujuan.ListIndex = 1 Then
TxtHarga.Caption = 500000
Else
MsgBox "Kota Asal dan Kota Tujuan Tidak Boleh Sama", vbInformation, "peringatan!"
End If
End Sub

Private Sub Command2_Click()
Form1.Show
Unload Me
End Sub

Private Sub Form_Load()
  aHari = Array("Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu")
  Timer1.Interval = 500
  Timer1.Enabled = True
  
  For t = 2014 To 2050
TAHUN.AddItem t
Next
For k = 1 To 31
TGL.AddItem k
Next
BULAN.List(0) = "Januari"
BULAN.List(1) = "Februari"
BULAN.List(2) = "Maret"
BULAN.List(3) = "April"
BULAN.List(4) = "Mei"
BULAN.List(5) = "Juni"
BULAN.List(6) = "Juli"
BULAN.List(7) = "Agustus"
BULAN.List(8) = "September"
BULAN.List(9) = "Oktober"
BULAN.List(10) = "November"
BULAN.List(11) = "Desember"

CboTujuan.List(0) = "Jakarta (Cgk)"
CboTujuan.List(1) = "Bali (DPS)"
CboTujuan.List(2) = "Surabaya (Sby)"

CboAsal.List(0) = "Jakarta (Cgk)"
CboAsal.List(1) = "Bali (DPS)"
CboAsal.List(2) = "Surabaya (Sby)"
End Sub


Private Sub Timer1_Timer()
  sHari = aHari(Abs(Weekday(Date) - 1))
  lbltanggal.Caption = Format(Date, "dd mmmm yyyy")
lbljam.Caption = Format(Time, "hh:mm:ss")
End Sub

