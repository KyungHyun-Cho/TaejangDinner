VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  '���� ����
   Caption         =   "����� ����"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2370
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   2370
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  '���
      Height          =   270
      IMEMode         =   3  '��� ����
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '���
      Height          =   270
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "UserPW"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "UserID"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   555
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text2.Text = "pass" Then
MsgBox "�������� ȯ���մϴ�.", vbInformation, "Success"
Form2.Show
Unload Me
Else
MsgBox "���̵� �Ǵ� ��й�ȣ�� Ȯ�����ּ���.", vbCritical, "Error"
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = "13" Then Call Command1_Click
End Sub
