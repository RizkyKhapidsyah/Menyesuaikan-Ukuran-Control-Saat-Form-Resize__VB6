VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menyesuaikan Ukuran Control saat Form Resize"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6900
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lngFormWidth As Long
Private lngFormHeight As Long

Private Sub Form_Load()
    Dim Ctl As Control
    'Tempatkan dimensi form dalam variabel
    lngFormWidth = ScaleWidth
    lngFormHeight = ScaleHeight
'Tempatkan inisialisasi dimensi control dalam 'property Tag - dengan penanganan error untuk 'controls yang tidak memiliki properties seperti 'Top (misalnya: control Line)
    On Error Resume Next
    For Each Ctl In Me
        Ctl.Tag = Ctl.Left & " " & Ctl.Top & " " & _
            Ctl.Width & " " & Ctl.Height & " "
            Ctl.Tag = Ctl.Tag & Ctl.FontSize & " "
    Next Ctl
    On Error GoTo 0
End Sub

Private Sub Form_Resize()
    Dim D(4) As Double
    Dim i As Long
    Dim TempPoz As Long
    Dim StartPoz As Long
    Dim Ctl As Control
    Dim TempVisible As Boolean
    Dim ScaleX As Double
    Dim ScaleY As Double
    'Hitung skala-nya
    ScaleX = ScaleWidth / lngFormWidth
    ScaleY = ScaleHeight / lngFormHeight
    On Error Resume Next
    'Untuk setiap control yang terdapat di form
    For Each Ctl In Me
        TempVisible = Ctl.Visible
        Ctl.Visible = False
        StartPoz = 1
        'Baca data dari property Tag
        For i = 0 To 4
            TempPoz = InStr(StartPoz, Ctl.Tag, " ", _
                vbTextCompare)
            If TempPoz > 0 Then
                D(i) = Mid(Ctl.Tag, StartPoz, _
                    TempPoz - StartPoz)
                StartPoz = TempPoz + 1
            Else
                D(i) = 0
            End If
            'Pindahkan control berdasarkan data
            'di property Tag dan di skala form
            Ctl.Move D(0) * ScaleX, D(1) * ScaleY, _
                D(2) * ScaleX, D(3) * ScaleY
            Ctl.Width = D(2) * ScaleX
            Ctl.Height = D(3) * ScaleY
            'Ganti ukuran huruf
            If ScaleX < ScaleY Then
                   Ctl.FontSize = D(4) * ScaleX
            Else
                   Ctl.FontSize = D(4) * ScaleY
            End If
        Next i
        Ctl.Visible = TempVisible
    Next Ctl
    On Error GoTo 0
End Sub


