VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form F_Saham 
   Caption         =   "Form1"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Frame1"
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7215
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add "
         Height          =   375
         Left            =   4680
         TabIndex        =   11
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtcari 
         Height          =   405
         Left            =   480
         TabIndex        =   10
         Top             =   3000
         Width           =   3975
      End
      Begin VB.CommandButton Cmdbatal 
         Caption         =   "BATAL"
         Height          =   375
         Left            =   5280
         TabIndex        =   9
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton Cmdhapus 
         Caption         =   "HAPUS"
         Height          =   375
         Left            =   5280
         TabIndex        =   8
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton Cmdubah 
         Caption         =   "UBAH"
         Height          =   375
         Left            =   5280
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton Cmdsimpan 
         Caption         =   "SIMPAN"
         Height          =   495
         Left            =   5280
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1455
         Left            =   480
         TabIndex        =   5
         Top             =   3720
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2566
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
      Begin VB.TextBox Txtnmsaham 
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox Txtkdsaham 
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C000&
         Caption         =   "Nama saham"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C000&
         Caption         =   "Kode Saham"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "F_Saham"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim saham As New ADODB.Recordset

Private Sub CmdAdd_Click()
If saham.State = adStateOpen Then saham.Close
saham.Open "select * from tbl_saham where kd_saham  like '%" & txtcari & "%'", koneksidb
If Not saham.EOF Then
    Txtkdsaham = saham!kd_saham
    Txtnmsaham = saham!nm_saham
    Call bukadb
    Call tampil_data
    Set DataGrid1.DataSource = saham
    With DataGrid1
    End With
    Call edit_grid
    End If
End Sub

Private Sub cmdbatal_Click()
Call kosong
Txtkdsaham.SetFocus
End Sub

Private Sub cmdhapus_Click()
Dim hapus As String
hapus = MsgBox("yakin ingin menghapus data ini", vbYesNo, "Pesan")
If hapus = vbYes Then
koneksidb.Execute "Delete from tbl_saham where kd_saham='" & Txtkdsaham.Text & "'"
Call refreshh
Txtkdsaham.SetFocus
Call kosong
End If
End Sub

Private Sub cmdsimpan_Click()
If Txtkdsaham.Text = "" Then
MsgBox "Kode Saham Kosong", vbExclamation, "Pesan"
Txtkdsaham.SetFocus
Exit Sub
End If

    If Txtnmsaham.Text = "" Then
    MsgBox "Nama Saham Kosong", vbExclamation, "Pesan"
    Txtnmsaham.SetFocus
    Exit Sub
    End If
Set saham = New ADODB.Recordset
saham.Open "select *from tbl_saham where kd_saham= '" & Txtkdsaham.Text & "'", koneksidb
If Not saham.EOF Then
MsgBox "Kode Saham Sudah Ada", vbCritical, "pesan"
Txtkdsaham.Text = ""
Txtkdsaham.SetFocus
Exit Sub
Else
    koneksidb.Execute "insert into tbl_saham(kd_saham,nm_saham)value('" & Txtkdsaham.Text & "','" & Txtnmsaham.Text & "')"
    MsgBox "Data Tersimpan"
    Call bukadb
    Call tampil_data
    Set DataGrid1.DataSource = saham
    With DataGrid1
    Call edit_grid
    Call kosong
    Txtkdsaham.SetFocus
    End With
End If
End Sub

Private Sub cmdubah_Click()
Dim ubah As String
ubah = MsgBox("yakin ingin mengubah data ini", vbYesNo, "Pesan")
If ubah = vbYes Then
koneksidb.Execute "update tbl_saham set nm_saham='" & Txtnmsaham.Text & "' where kd_saham='" & Txtkdsaham.Text & "'"
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = saham
With DataGrid1
End With
Call edit_grid
Call kosong
Txtkdsaham.SetFocus
End If
End Sub
Private Sub DataGrid1_Click()
Txtkdsaham.Text = saham!kd_saham
Txtnmsaham.Text = saham!nm_saham
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = saham
With DataGrid1
End With
Call edit_grid

End Sub

Private Sub Form_Load()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = saham
With saham
End With
End Sub
Sub tampil_data()
Set saham = New ADODB.Recordset
saham.ActiveConnection = koneksidb
saham.CursorLocation = adUseClient
saham.LockType = adLockOptimistic
saham.Source = "select*from tbl_saham"
saham.Open
End Sub
Sub edit_grid()
With DataGrid1
.Columns(0).Caption = "Koda Saham"
.Columns(1).Caption = "Deskripsi"
.Columns(0).Width = "1200"
.Columns(1).Width = "10000"
End With
End Sub
Sub kosong()
Txtkdsaham.Text = " "
Txtnmsaham.Text = " "
End Sub
Sub refreshh()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = saham
Call edit_grid
End Sub

Private Sub txtcari_Change()
Set saham = New ADODB.Recordset
saham.Open "select * from tbl_saham where kd_saham like '%" & txtcari & "%'", koneksidb
If Not saham.EOF Then
Set DataGrid1.DataSource = saham
Call edit_grid
End If
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If saham.State = adStateOpen Then saham.Close
saham.Open "select * from tbl_saham where kd_saham  like '%" & txtcari & "%'", koneksidb
If Not saham.EOF Then
    Txtkdsaham = saham!kd_saham
    Txtnmsaham = saham!nm_saham
    Call bukadb
    Call tampil_data
    Set DataGrid1.DataSource = saham
    With DataGrid1
    End With
    Call edit_grid
    End If
    End If
End Sub
