VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '���� ����
   Caption         =   "���� ���� ���α׷� (Version : Alpha Test 2.4)"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   8505
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton Command6 
      Caption         =   "�Ϸ�"
      Height          =   615
      Left            =   4320
      TabIndex        =   18
      Top             =   3000
      Width           =   4095
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3600
      Top             =   600
   End
   Begin VB.CommandButton Command4 
      Caption         =   "���� �Ļ� �Ϸ� ����"
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   4095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7680
      Top             =   840
   End
   Begin VB.Frame Frame2 
      Caption         =   "�л� ã��"
      Height          =   615
      Left            =   2040
      TabIndex        =   6
      Top             =   1440
      Width           =   6375
      Begin VB.TextBox Text2 
         Appearance      =   0  '���
         Height          =   270
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "������ �޴�"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.CommandButton Command1 
         Caption         =   "�ҷ�����"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "(NULL)"
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "��ġ�ϴ� �л� ���� ����!"
      BeginProperty Font 
         Name            =   "����"
         Size            =   36
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   120
      TabIndex        =   19
      Top             =   3720
      Width           =   8160
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "��ġ�ϴ� �л� ���� ����!"
      BeginProperty Font 
         Name            =   "����"
         Size            =   36
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   -7920
      TabIndex        =   17
      Top             =   -4200
      Width           =   8370
   End
   Begin VB.Label Label15 
      Height          =   345
      Left            =   6015
      TabIndex        =   14
      Top             =   2535
      Width           =   2385
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���� ���� ����"
      Height          =   180
      Left            =   4560
      TabIndex        =   13
      Top             =   2640
      Width           =   1200
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "1990-00-00"
      Height          =   180
      Left            =   2640
      TabIndex        =   12
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "������ ���� ��"
      Height          =   180
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   1200
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "(�̸�)"
      Height          =   180
      Left            =   6960
      TabIndex        =   10
      Top             =   2280
      Width           =   510
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "00000"
      Height          =   180
      Left            =   2760
      TabIndex        =   9
      Top             =   2280
      Width           =   450
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�й�"
      Height          =   180
      Left            =   720
      TabIndex        =   8
      Top             =   2280
      Width           =   360
   End
   Begin VB.Line Line9 
      X1              =   6000
      X2              =   6000
      Y1              =   2160
      Y2              =   2880
   End
   Begin VB.Line Line8 
      X1              =   4320
      X2              =   4320
      Y1              =   2160
      Y2              =   2880
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   8400
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�л� �̸�"
      Height          =   180
      Left            =   4800
      TabIndex        =   7
      Top             =   2280
      Width           =   780
   End
   Begin VB.Line Line6 
      X1              =   1680
      X2              =   1680
      Y1              =   2160
      Y2              =   2880
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   8400
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   8400
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line3 
      X1              =   8400
      X2              =   8400
      Y1              =   2160
      Y2              =   2880
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   2160
      Y2              =   2880
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8400
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "1990-00-00"
      BeginProperty Font 
         Name            =   "����"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4200
      TabIndex        =   5
      Top             =   960
      Width           =   2580
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "���� ��¥ :"
      BeginProperty Font 
         Name            =   "����"
         Size            =   20.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      TabIndex        =   4
      Top             =   960
      Width           =   2025
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "���� ���� ���α׷�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   36
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   6390
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public UserData As String '�����ϴ� �л���
Public UserData2 As String '����л� ���ڵ� ����
Public AdminData As String
'UBound(Split(C, "" & vbCrLf & "" & vbCrLf & ""))
Public StatusA As String
Public temp As String
Public PersonalData As String
Public NPersonalData As String
Public winhttp As New winhttp.WinHttpRequest
Public DoingA As Boolean
Public Function Addlog(str As String) As String
Form4.List1.AddItem "" & time & " >> " & str & "", 0
End Function
Public Function Inj(Infor As String) As String
Form7.Label8.Caption = Infor
If InStr(UserData, "" & Infor & "/") Then
PersonalData = "" & Infor & "" & Split(Split(UserData, Infor)(1), vbCrLf)(0) & ""
Label10.Caption = Split(PersonalData, "/")(0)
Label11.Caption = Split(PersonalData, "/")(1)
Label13.Caption = Split(PersonalData, "/")(2)
Me.Height = "4140"
If Label13.Caption = Label4.Caption Then
StatusA = "2"
Form7.Label4.Caption = "�������� (Error Code : A01)"
Text2.Text = ""
Form7.Label4.Visible = True
Label15.BackColor = &HFF&
Else
Label15.BackColor = &H8000&
NPersonalData = Replace(PersonalData, Split(PersonalData, "/")(2), Label4.Caption)
UserData = Replace(UserData, PersonalData, NPersonalData)
Dim Ar, B
 Set Ar = CreateObject("scripting.FileSystemObject")
 Set B = Ar.CreateTextFile("" & App.Path & "\Data\DataBase.db", True)
 B.Write (UserData)
 B.Close
Addlog ("" & Label11.Caption & "(" & Label10.Caption & ") �л� ���� �Ϸ�")
StatusA = "1"
Form7.Label4.Caption = "��������"
Form7.Label4.Visible = True
Text2.Text = ""
Text2.SetFocus
End If
Else
Label10.Caption = Infor
Label11.Caption = "(�̸�)"
Label13.Caption = "1990-00-00"
StatusA = "2"
Form7.Label4.Caption = "�������� (Error Code : A02)"
Text2.Text = ""
Form7.Label4.Visible = True
Me.Height = "4875"
Text2.Text = ""
Text2.SetFocus
End If
End Function
Public Function DataLoad() As String
On Error Resume Next
Dim Ar, B
 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)
UserData = B.readall
 B.Close
  Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase2.db", 1, True)
UserData2 = B.readall
 B.Close
   Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase3.db", 1, True)
AdminData = B.readall
 B.Close
End Function

Private Sub Combo1_Change()

End Sub

Private Sub Command1_Click()
On Error GoTo ErrA:
winhttp.Open "GET", "http://www.taejang.hs.kr/main.php?menugrp=020801&master=meal2&act=list"
winhttp.Send
Label1.Caption = Replace(Replace(Split(Split(Split(winhttp.ResponseText, " �޴�")(1), "<td>")(1), "</td>")(0), " ", ""), ",", vbCrLf)
Exit Sub
ErrA:
MsgBox "���ͳ� ���� ���� �Ǵ� " & vbCrLf & "�Ĵ��� ��ϵǾ����� �ʰų�" & vbCrLf & "��Ÿ ������ �Ĵ��� �ҷ� �� �� �����ϴ�.", vbCritical, "Error"

End Sub





Private Sub Command4_Click()
On Error Resume Next
'If Label13.Caption <> Label4.Caption Then MsgBox "�ش� �л��� ���� ���� �� �ʿ䰡 �����ϴ�.", vbCritical, "Error": Exit Sub
NPersonalData = Replace(PersonalData, Split(PersonalData, "/")(2), "���� ����")
UserData = Replace(UserData, PersonalData, NPersonalData)
Dim Ar, B
 Set Ar = CreateObject("scripting.FileSystemObject")
 Set B = Ar.CreateTextFile("" & App.Path & "\Data\DataBase.db", True)
 B.Write (UserData)
 B.Close
Addlog ("" & Label11.Caption & "(" & Label10.Caption & ") �л� ���� ����")
MsgBox "���� �Ļ� ���� ó���Ǿ����ϴ�.", vbInformation, "Success"
PersonalData = "" & Label10.Caption & "" & Split(Split(UserData, Label10.Caption)(1), vbCrLf)(0) & ""
Label10.Caption = Split(PersonalData, "/")(0)
Label11.Caption = Split(PersonalData, "/")(1)
Label13.Caption = Split(PersonalData, "/")(2)
Me.Height = "4140"
If Label13.Caption = Label4.Caption Then
Label15.BackColor = &HFF&
Else
Label15.BackColor = &H8000&
End If
Form4.Label3.Caption = UBound(Split(UserData, Label4.Caption))
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command6_Click()
Form7.BorderStyle = 0
Form7.WindowState = "2"

End Sub

Private Sub Form_Load()
StatusA = "0"
Form7.Show
'Form3.Show
'Form4.Show
DataLoad
Label4.Caption = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form7.DoingA = False
End Sub

Private Sub Text2_Change()
If Form7.AdminMode = "0" Then
If StatusA = "0" Then
Form7.Label4.Visible = False
ElseIf StatusA = "1" Then
Form7.Label4.Visible = True
StatusA = "0"
ElseIf StatusA = "2" Then
StatusA = "0"
End If

If Len(Text2.Text) < "3" Then
Form3.List1.Clear
Form3.Label2.Caption = Form3.List1.ListCount
ElseIf Len(Text2.Text) = "3" Then

On Error Resume Next
Form3.List1.Clear

Dim Ar, B, C, D
 Set Ar = CreateObject("scripting.FileSystemObject")
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)
C = UBound(Split(B.readall, vbCrLf)) + 5
B.Close
Set B = Ar.OpenTextFile("" & App.Path & "\Data\DataBase.db", 1, True)



For i = 1 To C
D = B.ReadLine
If UBound(Split(D, "/")) = 2 Then
If Left(D, 3) = Text2.Text Then Form3.List1.AddItem D: D = ""
Form3.Label2.Caption = Form3.List1.ListCount
End If
DoEvents
Next i
End If
Form7.Label3.Caption = Text2.Text
Else
Form7.M1.ForeColor = RGB(0, 0, 0)
Form7.M2.ForeColor = RGB(0, 0, 0)
Form7.M3.ForeColor = RGB(0, 0, 0)
Form7.M4.ForeColor = RGB(0, 0, 0)
Form7.M5.ForeColor = RGB(0, 0, 0)
Form7.M6.ForeColor = RGB(0, 0, 0)
Form7.M7.ForeColor = RGB(0, 0, 0)
Form7.M8.ForeColor = RGB(0, 0, 0)
Form7.M9.ForeColor = RGB(0, 0, 0)
Form7.M0.ForeColor = RGB(0, 0, 0)
If Text2.Text = "1" Then
Form7.M1.ForeColor = RGB(255, 0, 0)
ElseIf Text2.Text = "2" Then
Form7.M2.ForeColor = RGB(255, 0, 0)
ElseIf Text2.Text = "3" Then
Form7.M3.ForeColor = RGB(255, 0, 0)
ElseIf Text2.Text = "4" Then
Form7.M4.ForeColor = RGB(255, 0, 0)
ElseIf Text2.Text = "5" Then
Form7.M5.ForeColor = RGB(255, 0, 0)
ElseIf Text2.Text = "6" Then
Form7.M6.ForeColor = RGB(255, 0, 0)
ElseIf Text2.Text = "7" Then
Form7.M7.ForeColor = RGB(255, 0, 0)
ElseIf Text2.Text = "8" Then
Form7.M8.ForeColor = RGB(255, 0, 0)
ElseIf Text2.Text = "9" Then
Form7.M9.ForeColor = RGB(255, 0, 0)
ElseIf Text2.Text = "0" Then
Form7.M0.ForeColor = RGB(255, 0, 0)
End If
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim TempLV As String
        If KeyAscii = 13 Then
        Text2.Text = Replace(Replace(Replace(Replace(Replace(Text2.Text, " ", ""), "+", ""), "-", ""), "*", ""), "/", "")
If Form7.AdminMode = "0" Then '�Ϲ� ���
If Text2.Text = "" Or Len(Text2.Text) < 4 Then Exit Sub
If InStr(AdminData, "" & Text2.Text & "/") Then
Form7.AdminMode = Split(AdminData, "/")(2)
If Split(AdminData, "/")(2) = "2" Then
TempLV = "Super Admin"
Else
TempLV = "Normal Admin"
End If
Form7.Label13.Caption = "" & Split(AdminData, "/")(1) & " (" & TempLV & ")"
Form7.AdminData1 = Text2.Text
Text2.Text = ""
ElseIf Len(Text2.Text) = "5" Then
Call Inj(Text2.Text)
ElseIf Len(Text2.Text) = "7" Then
If InStr(UserData2, Text2.Text) Then
Text2.Text = Split(Split(UserData2, Text2.Text)(1), "/")(1)
Else
Text2.Text = "00000"
End If
Call Inj(Text2.Text)
Else '�߸��Ȱ��
Exit Sub
End If

Else '������ ���
If Text2.Text = "0" Then
Form7.AdminMode = "0"
ElseIf Text2.Text = "1" Then
Form2.Show
Form7.AdminMode = "0"
ElseIf Text2.Text = "2" Then
If Form7.AdminMode = "2" Then
MsgBox "SuperAdmin �Դϴ�.", vbInformation, "Succes"
Else
MsgBox "�ش� �޴��� SuperAdmin �� ������ �����մϴ�.", vbCritical, "Error"
End If
Form7.AdminMode = "0"
ElseIf Text2.Text = "3" Then
Form7.AdminMode = "0"
ElseIf Text2.Text = "4" Then
Form7.AdminMode = "0"
ElseIf Text2.Text = "5" Then
Form7.AdminMode = "0"
ElseIf Text2.Text = "6" Then
Form4.Show
Form7.AdminMode = "0"
ElseIf Text2.Text = "7" Then
Form7.AdminMode = "0"
ElseIf Text2.Text = "8" Then
Form7.AdminMode = "0"
ElseIf Text2.Text = "9" Then
Form7.AdminMode = "0"
If InStr(UserData2, Form7.AdminData1) Then
Text2.Text = Split(Split(UserData2, Form7.AdminData1)(1), "/")(1)
Else
Text2.Text = "00000"
End If
Call Inj(Text2.Text)


End If
Text2.Text = ""
End If
End If
If Form7.AdminMode <> "0" Then Text2.Text = ""
Form4.Label3.Caption = UBound(Split(UserData, Label4.Caption))
End Sub

Private Sub Timer1_Timer()
Text2.SetFocus
End Sub

Private Sub Timer2_Timer()
DataLoad
End Sub
