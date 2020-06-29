VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form F_Kriteria 
   Caption         =   "F_Kriteria"
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Frame1"
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1815
         Left            =   480
         TabIndex        =   11
         Top             =   3600
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   3201
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
      Begin VB.CommandButton cmdbatal 
         Caption         =   "BATAL"
         Height          =   375
         Left            =   5880
         TabIndex        =   10
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmdhapus 
         Caption         =   "HAPUS"
         Height          =   375
         Left            =   5880
         TabIndex        =   9
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmdubah 
         Caption         =   "UBAH"
         Height          =   375
         Left            =   5880
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdsimpan 
         Caption         =   "SIMPAN"
         Height          =   375
         Left            =   5880
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cmbatribut 
         Height          =   315
         Left            =   2400
         TabIndex        =   6
         Text            =   "Pilih"
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox txtnmkriteria 
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtkdkriteria 
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Atribut"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Nama Kriteria"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Kriteria"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
   End
End
Attribute VB_Name = "F_Kriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim kriteria As New ADODB.Recordset

Private Sub cmdbatal_Click()
Call kosong
txtkdkriteria.SetFocus
End Sub

Private Sub cmdhapus_Click()
Dim hapus As String
hapus = MsgBox("yakin ingin menghapus data ini", vbYesNo, "Pesan")
If hapus = vbYes Then
koneksidb.Execute "Delete from tbl_kriteria where kd_kriteria='" & txtkdkriteria.Text & "'"
Call refreshh
txtkdkriteria.SetFocus
Call kosong
End If
End Sub

Private Sub cmdsimpan_Click()
If txtkdkriteria.Text = "" Then
MsgBox "Kode Kriteria Kosong", vbExclamation, "Pesan"
txtkdkriteria.SetFocus
Exit Sub
End If

    If txtnmkriteria.Text = "" Then
    MsgBox "Nama Kriteria Kosong", vbExclamation, "Pesan"
    txtnmkriteria.SetFocus
    Exit Sub
    End If
    
            If cmbatribut.Text = "" Then
            MsgBox "Atribut Kosong", vbExclamation, "Pesan"
            cmbatribut.SetFocus
            Exit Sub
            End If

Set kriteria = New ADODB.Recordset
kriteria.Open "select *from tbl_kriteria where kd_kriteria= '" & txtkdkriteria.Text & "'", koneksidb
If Not kriteria.EOF Then
MsgBox "Kode Kriteria Sudah Ada", vbCritical, "pesan"
txtkdkriteria.Text = ""
txtkdkriteria.SetFocus
Exit Sub
Else
    koneksidb.Execute "insert into tbl_kriteria(kd_kriteria,nm_kriteria,atribut)value('" & txtkdkriteria.Text & "','" & txtnmkriteria.Text & "','" & cmbatribut.Text & "')"
    MsgBox "Data Tersimpan"
    Call bukadb
    Call tampil_data
    Set DataGrid1.DataSource = kriteria
    With DataGrid1
    Call edit_grid
    Call kosong
    txtkdkriteria.SetFocus
    End With
End If
End Sub

Private Sub cmdubah_Click()
Dim ubah As String
ubah = MsgBox("yakin ingin mengubah data ini", vbYesNo, "Pesan")
If ubah = vbYes Then
koneksidb.Execute "update tbl_kriteria set nm_kriteria='" & txtnmkriteria.Text & "',atribut='" & cmbatribut.Text & "' where kd_kriteria='" & txtkdkriteria.Text & "'"
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = kriteria
With DataGrid1
End With
Call edit_grid
Call kosong
txtkdkriteria.SetFocus
End If
End Sub



Private Sub DataGrid1_Click()
txtkdkriteria.Text = kriteria!kd_kriteria
txtnmkriteria.Text = kriteria!nm_kriteria
cmbatribut.Text = kriteria!atribut

Call bukadb
Call tampil_data
Set DataGrid1.DataSource = kriteria
With DataGrid1
End With
Call edit_grid
End Sub

Private Sub Form_Load()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = kriteria
With kriteria
End With
With DataGrid1

End With
With cmbatribut
.AddItem "Benefit"
.AddItem "Cost"
End With
Call edit_grid
End Sub



Sub tampil_data()
Set kriteria = New ADODB.Recordset
kriteria.ActiveConnection = koneksidb
kriteria.CursorLocation = adUseClient
kriteria.LockType = adLockOptimistic
kriteria.Source = "select*from tbl_kriteria"
kriteria.Open
End Sub

Sub edit_grid()
With DataGrid1
.Columns(0).Caption = "Kode Kriteria"
.Columns(1).Caption = "Nama Kriteria"
.Columns(2).Caption = "Atribut"

.Columns(0).Width = "1200"
.Columns(1).Width = "2000"
.Columns(2).Width = "2000"
End With
End Sub

Sub kosong()
txtkdkriteria.Text = " "
txtnmkriteria.Text = " "
cmbatribut.Text = " "
End Sub
Sub refreshh()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = kriteria
Call edit_grid
End Sub





