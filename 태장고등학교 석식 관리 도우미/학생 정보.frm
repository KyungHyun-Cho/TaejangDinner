VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   1  '단일 고정
   Caption         =   "학생 정보"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3105
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   3105
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "검색된 리스트"
      ForeColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2895
      Begin VB.ListBox List1 
         Appearance      =   0  '평면
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3450
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "명"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "검색된 학생 수 :"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1320
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub
