VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16290
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   16290
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Waktu"
      Height          =   1095
      Left            =   11160
      TabIndex        =   29
      Top             =   5160
      Width           =   2175
      Begin VB.Label Lbl_time 
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tanggal"
      Height          =   1215
      Left            =   7560
      TabIndex        =   27
      Top             =   5040
      Width           =   2415
      Begin VB.Label Lbl_date 
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Total Bayar"
      Height          =   495
      Left            =   7560
      TabIndex        =   23
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Beli"
      Height          =   2175
      Left            =   7560
      TabIndex        =   16
      Top             =   1080
      Width           =   5535
      Begin VB.TextBox Db6 
         Height          =   495
         Left            =   3840
         TabIndex        =   22
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Db5 
         Height          =   495
         Left            =   3840
         TabIndex        =   21
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Db4 
         Height          =   495
         Left            =   2040
         TabIndex        =   20
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Db3 
         Height          =   495
         Left            =   2040
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Db2 
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Db1 
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear"
      Height          =   615
      Left            =   11640
      TabIndex        =   15
      Top             =   240
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   13440
      Top             =   240
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add Bayar"
      Height          =   615
      Left            =   7560
      TabIndex        =   14
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   615
      Left            =   5400
      TabIndex        =   13
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   615
      Left            =   3480
      TabIndex        =   12
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Total Bayar"
      Height          =   615
      Left            =   3840
      TabIndex        =   10
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox Tjml_beli 
      Height          =   495
      Left            =   5040
      TabIndex        =   9
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Tjenis_brg 
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Thrg_brg 
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   960
      Width           =   1935
   End
   Begin VB.ListBox list_brg 
      Height          =   4935
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton Btn_addjenis 
      Caption         =   "Tambah Jenis Barang"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.ComboBox List_Jenis 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1680
      List            =   "Form1.frx":000A
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Lbl_thanks 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   26
      Top             =   4320
      Width           =   5775
   End
   Begin VB.Label Lbl_totalbeli 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10680
      TabIndex        =   25
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      TabIndex        =   24
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   7320
      X2              =   7320
      Y1              =   120
      Y2              =   5880
   End
   Begin VB.Label Lbl_totalbayar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   11
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Jumlah Beli"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Jenis Barang"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Harga Barang"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Jenis Barang"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()

End Sub

Private Sub Btn_addjenis_Click()
list_brg.Clear
If List_Jenis.Text = "Rokok" Then
list_brg.AddItem "Dji Sam Soe"
list_brg.AddItem "Djarum Super"
list_brg.AddItem "Starmild"
list_brg.AddItem "Neomild"
list_brg.AddItem "Signature"
ElseIf List_Jenis.Text = "Mie" Then
list_brg.AddItem "Indomie"
list_brg.AddItem "Supermie"
list_brg.AddItem "Popmie"
list_brg.AddItem "Sarimie"
list_brg.AddItem "Mie Sedap"
End If
End Sub

Private Sub Command1_Click()
Lbl_totalbayar.Caption = Val(Thrg_brg.Text) * Val(Tjml_beli.Text)
End Sub

Private Sub Command2_Click()
Thrg_brg.Text = ""
Tjenis_brg.Text = ""
Tjml_beli.Text = ""
Lbl_totalbayar.Caption = ""
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Dim total As String
total = Lbl_totalbayar.Caption
If Db1.Text = "" Then
Db1.Text = total
ElseIf Db2.Text = "" Then
Db2.Text = total
ElseIf Db3.Text = "" Then
Db3.Text = total
ElseIf Db4.Text = "" Then
Db4.Text = total
ElseIf Db5.Text = "" Then
Db5.Text = total
ElseIf Db6.Text = "" Then
Db6.Text = total
Else
MsgBox "Data bayar sudah penuh!"
End If
End Sub

Private Sub Command5_Click()
Db1.Text = ""
Db2.Text = ""
Db3.Text = ""
Db4.Text = ""
Db5.Text = ""
Db6.Text = ""
Lbl_totalbeli.Caption = ""
Lbl_thanks.Caption = ""

End Sub

Private Sub Command6_Click()
Dim total_beli As Double
total_beli = Val(Db1.Text) + Val(Db2.Text) + Val(Db3.Text) + Val(Db4.Text) + Val(Db5.Text) + Val(Db6.Text)
Lbl_totalbeli.Caption = total_beli
Lbl_thanks.Caption = "Terimakasih telah berbelanja di Nias market:) :*"
End Sub

Private Sub Lbl_totalbayar_Click()

End Sub

Private Sub list_brg_Click()
Dim harga As Double
Dim jenis As String
If List_Jenis.Text = "Rokok" Then
jenis = "Rokok"
Select Case list_brg.Text
Case "Dji Sam Soe"
harga = 12000
Case "Djarum Super"
harga = 10000
Case "Starmild"
harga = 11000
Case "Neomild"
harga = 10500
Case "Signature"
harga = 14000
End Select
ElseIf List_Jenis.Text = "Mie" Then
jenis = "Mie"
Select Case list_brg.Text
Case "Indomie"
harga = 1500
Case "Supermie"
harga = 1400
Case "Popmie"
harga = 6000
Case "Sarimie"
harga = 1300
Case "Mie Sedap"
harga = 1200
End Select
End If
Thrg_brg.Text = harga
Tjenis_brg.Text = jenis
Tjml_beli.Text = ""
Lbl_totalbayar.Caption = ""
End Sub

Private Sub List_Jenis_Change()

End Sub

Private Sub Timer1_Timer()
Lbl_date.Caption = Format(Now, "d mmmm yyyy")
Lbl_time.Caption = Format(Now, "hh : mm : ss")
End Sub
