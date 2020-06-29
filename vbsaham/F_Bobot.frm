VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form F_Bobot 
   Caption         =   "F_Bobot"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Frame1"
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton cmdbatal 
         Caption         =   "BATAL"
         Height          =   375
         Left            =   5880
         TabIndex        =   11
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdhapus 
         Caption         =   "HAPUS"
         Height          =   375
         Left            =   5880
         TabIndex        =   10
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdubah 
         Caption         =   "UBAH"
         Height          =   375
         Left            =   5880
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdsimpan 
         Caption         =   "SIMPAN"
         Height          =   375
         Left            =   5880
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1815
         Left            =   240
         TabIndex        =   7
         Top             =   3240
         Width           =   7215
         _ExtentX        =   12726
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
      Begin VB.TextBox txtnilai 
         Height          =   375
         Left            =   2760
         TabIndex        =   6
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtjnsbobot 
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtkdbobot 
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Nilai"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Jenis Bobot"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Bobot"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "F_Bobot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bobot As New ADODB.Recordset

Private Sub cmdbatal_Click()
Call kosong
txtkdkriteria.SetFocus
End Sub

Private Sub cmdhapus_Click()
Dim hapus As String
hapus = MsgBox("yakin ingin menghapus data ini", vbYesNo, "Pesan")
If hapus = vbYes Then
koneksidb.Execute "Delete from tbl_bobot where kd_bobot='" & txtkdbobot.Text & "'"
Call refreshh
txtkdbobot.SetFocus
Call kosong
End If
End Sub

Private Sub cmdsimpan_Click()
If txtkdbobot.Text = "" Then
MsgBox "Kode Bobot Kosong", vbExclamation, "Pesan"
txtkdbobot.SetFocus
Exit Sub
End If

    If txtjnsbobot.Text = "" Then
    MsgBox "Jenis Bobot Kosong", vbExclamation, "Pesan"
    txtjnsbobot.SetFocus
    Exit Sub
    End If
    
        If txtnilai.Text = "" Then
        MsgBox "Nilai Kosong", vbExclamation, "Pesan"
        txtnilai.SetFocus
        Exit Sub
        End If
Set nilai = New ADODB.Recordset
bobot.Open "select *from tbl_bobot where kd_bobot= '" & txtkdbobot.Text & "'", koneksidb
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

Private Sub cmdubah_Click()
Dim ubah As String
ubah = MsgBox("yakin ingin mengubah data ini", vbYesNo, "Pesan")
If ubah = vbYes Then
koneksidb.Execute "update tbl_bobot set jns_bobot='" & txtjnsbobot.Text & "',nilai='" & txtnilai.Text & "' where kd_bobot='" & txtkdbobot.Text & "'"
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = bobot
With DataGrid1
End With
Call edit_grid
Call kosong
txtkdbobot.SetFocus
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
