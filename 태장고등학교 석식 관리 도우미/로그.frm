VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   BorderStyle     =   1  '���� ����
   Caption         =   "�α�â"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4695
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command2 
      Caption         =   "�˻�"
      Height          =   260
      Left            =   3240
      TabIndex        =   6
      Top             =   80
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '���
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   5
      Top             =   80
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4080
      Top             =   600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�α� ����"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   4455
   End
   Begin VB.ListBox List1 
      Appearance      =   0  '���
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�й� :"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1440
      TabIndex        =   7
      Top             =   75
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3600
      TabIndex        =   4
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�޽� ���� �л� �� :"
      BeginProperty Font 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1440
      TabIndex        =   3
      Top             =   360
      Width           =   2130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�α�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   27.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1140
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public UserData As String
Private Sub Command1_Click()
Dim Ar, B, C

 Set Ar = CreateObject("scripting.FileSystemObject")
 Set B = Ar.CreateTextFile("" & App.Path & "\Log\LogSave(" & Replace(Form1.Label4.Caption, "-", "") & "" & Replace(Replace(time, " ", ""), ":", "") & ").txt", True)
For i = 0 To List1.ListCount - 1
C = "" & C & "" & vbCrLf & "" & List1.List(i) & ""
DoEvents
Next i
C = "�޽� ���� �л� �� : " & Label3.Caption & "" & vbCrLf & "" & C & ""
 B.Write (C)
B.Close
MsgBox "������ �Ϸ�Ǿ����ϴ�.", vbInformation, "����"
End Sub

Private Sub Command2_Click()
Dim aA As String
For i = 0 To List1.ListCount - 1
If InStr(List1.List(i), Text1.Text) Then
List1.ListIndex = i
aA = "1"
Else
End If
DoEvents
Next i
If aA = "1" Then
MsgBox "�˻��� �Ϸ�Ǿ����ϴ�.", vbInformation, "�˻� �Ϸ�"
Else
MsgBox "��ġ�ϴ� ������ �����ϴ�.", vbCritical, "�˻� �Ϸ�."
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form7.DoingA = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = "13" Then Call Command2_Click
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim Ar, B
 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)
UserData = B.readall

Label3.Caption = UBound(Split(UserData, Date))
End Sub
