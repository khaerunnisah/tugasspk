VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13815
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   13815
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Frame1"
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12255
      Begin VB.CommandButton Command1 
         Caption         =   "CARI"
         Height          =   375
         Left            =   10560
         TabIndex        =   24
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   7680
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   3000
         Width           =   2655
      End
      Begin VB.CommandButton cmdbatal 
         Caption         =   "BATAL"
         Height          =   375
         Left            =   7440
         TabIndex        =   22
         Top             =   6120
         Width           =   1095
      End
      Begin VB.CommandButton cmdhapus 
         Caption         =   "HAPUS"
         Height          =   375
         Left            =   6120
         TabIndex        =   21
         Top             =   6120
         Width           =   1095
      End
      Begin VB.CommandButton cmdubah 
         Caption         =   "UBAH"
         Height          =   375
         Left            =   4680
         TabIndex        =   20
         Top             =   6120
         Width           =   1095
      End
      Begin VB.CommandButton cmdsimpan 
         Caption         =   "SIMPAN"
         Height          =   375
         Left            =   3360
         TabIndex        =   19
         Top             =   6120
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2415
         Left            =   240
         TabIndex        =   18
         Top             =   3600
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   4260
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C000&
         Caption         =   "Kriteria"
         Height          =   2175
         Left            =   4560
         TabIndex        =   5
         Top             =   240
         Width           =   6255
         Begin VB.TextBox txtper 
            Height          =   285
            Left            =   4680
            TabIndex        =   17
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtpendapata 
            Height          =   285
            Left            =   4680
            TabIndex        =   15
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtlbbersih 
            Height          =   285
            Left            =   4680
            TabIndex        =   13
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtlbusaha 
            Height          =   285
            Left            =   1680
            TabIndex        =   11
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtlbkotor 
            Height          =   285
            Left            =   1680
            TabIndex        =   9
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtaset 
            Height          =   285
            Left            =   1680
            TabIndex        =   7
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "PER"
            Height          =   255
            Left            =   3240
            TabIndex        =   16
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label txtpendapatan 
            Caption         =   "Pendapatan"
            Height          =   255
            Left            =   3240
            TabIndex        =   14
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Laba Bersih"
            Height          =   255
            Left            =   3240
            TabIndex        =   12
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Laba Usaha"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Laba Kotor"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Aset"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox txtnmsaham 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtkdsaham 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "Nama Saham"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kode Saham"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bk As New ADODB.Recordset
Private Sub cmdsimpan_Click()
If txtkdsaham.Text = "" Then
MsgBox "Kode Saham Kosong", vbExclamation, "Pesan"
txtkdsaham.SetFocus
Exit Sub
End If

    If txtnmsaham.Text = "" Then
    MsgBox "Nama Saham Kosong", vbExclamation, "Pesan"
    txtjnmsaham.SetFocus
    Exit Sub
    End If
    
        If txtaset.Text = "" Then
        MsgBox "Nilai Aset Kosong", vbExclamation, "Pesan"
        txtaset.SetFocus
        Exit Sub
        End If
            If txtlbbersih.Text = "" Then
            MsgBox "Laba Bersih Kosong", vbExclamation, "Pesan"
            txtlbbersih.SetFocus
            Exit Sub
            End If
                If txtlbkotor.Text = "" Then
                MsgBox "Laba kotor Kosong", vbExclamation, "Pesan"
                txtlbkotor.SetFocus
                Exit Sub
                End If
                    If txtlbusaha.Text = "" Then
                    MsgBox "Laba Usaha Kosong", vbExclamation, "Pesan"
                    txtlbusaha.SetFocus
                    Exit Sub
                    End If
                        If txtpendapatan.Text = "" Then
                        MsgBox "Pendapatan Kosong", vbExclamation, "Pesan"
                        txtpendapatan.SetFocus
                        Exit Sub
                        End If
                            If txtper.Text = "" Then
                            MsgBox "Per Kosong", vbExclamation, "Pesan"
                            txtper.SetFocus
                            Exit Sub
                            End If
Set bk = New ADODB.Recordset
bk.Open "select *from tbl_bobot where kd_bobot= '" & txtkdbobot.Text & "'", koneksidb
If Not bobot.EOF Then
MsgBox "Kode Bobot Sudah Ada", vbCritical, "pesan"
txtkdbobot.Text = ""
txtkdbobot.SetFocus
Exit Sub
Else
    koneksidb.Execute "insert into tbl_bobot(kd_bobot,jns_bobot,nilai)value('" & txtkdbobot.Text & "','" & txtjnsbobot.Text & "','" & txtnilai.Text & "')"
    MsgBox "Data Tersimpan"
    Call bukadb
    Call tampil_data
    Set DataGrid1.DataSource = bobot
    With DataGrid1
    Call edit_grid
    Call kosong
    txtkdbobot.SetFocus
    End With
End If
End Sub
