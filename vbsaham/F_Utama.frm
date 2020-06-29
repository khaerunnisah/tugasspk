VERSION 5.00
Begin VB.Form F_Utama 
   Caption         =   "F_Utama"
   ClientHeight    =   8535
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   14685
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   8535
      Left            =   0
      Picture         =   "F_Utama.frx":0000
      ScaleHeight     =   8475
      ScaleWidth      =   14595
      TabIndex        =   0
      Top             =   0
      Width           =   14655
   End
   Begin VB.Menu DATA_SAHAM 
      Caption         =   "DATA_SAHAM"
   End
   Begin VB.Menu KRITERIA 
      Caption         =   "KRITERIA"
   End
   Begin VB.Menu BOBOT 
      Caption         =   "BOBOT"
   End
   Begin VB.Menu PENILAIAN 
      Caption         =   "PENILAIAN"
   End
   Begin VB.Menu HASIL 
      Caption         =   "HASIL"
   End
End
Attribute VB_Name = "F_Utama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOBOT_Click()
F_Bobot.Show
End Sub

Private Sub Data_Saham_Click()
F_Saham.Show
End Sub

Private Sub KRITERIA_Click()
F_Kriteria.Show
End Sub

Private Sub PENILAIAN_Click()
Form1.Show
End Sub
