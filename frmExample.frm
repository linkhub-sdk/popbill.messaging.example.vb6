VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "�˺� �޽�¡ SDK ����"
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   9825
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.Frame Frame6 
      Caption         =   " �˺� �޽�¡ ���� ���"
      Height          =   6975
      Left            =   75
      TabIndex        =   16
      Top             =   2730
      Width           =   10725
      Begin VB.CommandButton btnSearchPopup 
         Caption         =   "���۳�����ȸ �˾�"
         Height          =   390
         Left            =   8655
         TabIndex        =   37
         Top             =   195
         Width           =   1875
      End
      Begin VB.TextBox txtResult 
         Height          =   4680
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   3  '�����
         TabIndex        =   36
         Top             =   2205
         Width           =   10455
      End
      Begin VB.CommandButton btnCancelReserve 
         Caption         =   "���� ���� ���"
         Height          =   405
         Left            =   7185
         TabIndex        =   35
         Top             =   1755
         Width           =   2505
      End
      Begin VB.CommandButton btnGetMessages 
         Caption         =   "���ۻ���Ȯ��"
         Height          =   405
         Left            =   4485
         TabIndex        =   34
         Top             =   1755
         Width           =   2505
      End
      Begin VB.Frame Frame9 
         Caption         =   " �����ν� �ڵ� ���� ���� "
         Height          =   945
         Left            =   7095
         TabIndex        =   30
         Top             =   735
         Width           =   3465
         Begin VB.CommandButton btnSendXMS_One 
            Caption         =   "1�� ����"
            Height          =   465
            Left            =   120
            TabIndex        =   33
            Top             =   315
            Width           =   930
         End
         Begin VB.CommandButton btnSendXMS_Hundred 
            Caption         =   "100�� ����"
            Height          =   465
            Left            =   1125
            TabIndex        =   32
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendXMS_Same 
            Caption         =   "��������"
            Height          =   465
            Left            =   2310
            TabIndex        =   31
            Top             =   315
            Width           =   1020
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   " �幮 ���� ���� "
         Height          =   945
         Left            =   3570
         TabIndex        =   26
         Top             =   735
         Width           =   3465
         Begin VB.CommandButton btnSendLMS_One 
            Caption         =   "1�� ����"
            Height          =   465
            Left            =   120
            TabIndex        =   29
            Top             =   315
            Width           =   930
         End
         Begin VB.CommandButton btnSendLMS_Hundred 
            Caption         =   "100�� ����"
            Height          =   465
            Left            =   1125
            TabIndex        =   28
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendLMS_Same 
            Caption         =   "��������"
            Height          =   465
            Left            =   2310
            TabIndex        =   27
            Top             =   315
            Width           =   1020
         End
      End
      Begin VB.TextBox txtReceiptNum 
         Height          =   315
         Left            =   1125
         TabIndex        =   25
         Top             =   1785
         Width           =   3105
      End
      Begin VB.Frame Frame7 
         Caption         =   " �ܹ� ���� ���� "
         Height          =   945
         Left            =   45
         TabIndex        =   19
         Top             =   735
         Width           =   3465
         Begin VB.CommandButton btnSendSMS_Same 
            Caption         =   "��������"
            Height          =   465
            Left            =   2310
            TabIndex        =   22
            Top             =   315
            Width           =   1020
         End
         Begin VB.CommandButton btnSendSMS_hundredd 
            Caption         =   "100�� ����"
            Height          =   465
            Left            =   1125
            TabIndex        =   21
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendSMS_One 
            Caption         =   "1�� ����"
            Height          =   465
            Left            =   120
            TabIndex        =   20
            Top             =   315
            Width           =   930
         End
      End
      Begin VB.TextBox txtReserveDT 
         Height          =   315
         Left            =   3060
         TabIndex        =   17
         Top             =   255
         Width           =   3105
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "������ȣ : "
         Height          =   180
         Left            =   225
         TabIndex        =   24
         Top             =   1875
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "����ð�(yyyyMMddHHmmss) : "
         Height          =   180
         Left            =   225
         TabIndex        =   18
         Top             =   330
         Width           =   2790
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " �˺� �⺻ API "
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   10680
      Begin VB.Frame Frame2 
         Caption         =   " ȸ������"
         Height          =   1575
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "���� ���� Ȯ��"
            Height          =   495
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "ȸ�� ����"
            Height          =   495
            Left            =   240
            TabIndex        =   13
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " ����Ʈ ����"
         Height          =   1575
         Left            =   2160
         TabIndex        =   10
         Top             =   360
         Width           =   2160
         Begin VB.CommandButton btnUnitCost_LMS 
            Caption         =   "�幮 ���� �ܰ� Ȯ��"
            Height          =   495
            Left            =   165
            TabIndex        =   15
            Top             =   855
            Width           =   1815
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "�ܹ� ���� �ܰ� Ȯ��"
            Height          =   495
            Left            =   150
            TabIndex        =   11
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " ��Ʈ�� ����"
         Height          =   1575
         Left            =   4410
         TabIndex        =   8
         Top             =   405
         Width           =   2535
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ� ����Ʈ Ȯ��"
            Height          =   495
            Left            =   120
            TabIndex        =   23
            Top             =   270
            Width           =   1815
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "��Ʈ�� �ܿ� ����Ʈ Ȯ��"
            Height          =   495
            Left            =   120
            TabIndex        =   9
            Top             =   960
            Width           =   2295
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " ��Ÿ"
         Height          =   1575
         Left            =   7035
         TabIndex        =   5
         Top             =   390
         Width           =   2175
         Begin VB.CommandButton btnGetPopbillURL 
            Caption         =   " �˺� �⺻ URL Ȯ��"
            Height          =   495
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   1935
         End
         Begin VB.ComboBox cboPopbillTOGO 
            Height          =   300
            Left            =   120
            TabIndex        =   6
            Text            =   "LOGIN"
            Top             =   360
            Width           =   1935
         End
      End
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Top             =   165
      Width           =   1935
   End
   Begin VB.TextBox txtCorpNum 
      Height          =   315
      Left            =   1335
      TabIndex        =   1
      Text            =   "1231212312"
      Top             =   180
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�˺����̵� : "
      Height          =   180
      Left            =   3480
      TabIndex        =   2
      Top             =   240
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����ڹ�ȣ : "
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�������̵�
Private Const LinkID = "TESTER"
'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
Private Const SecretKey = "088b1258aoeMH5OtGjK4zaOlwZGVvSK40ceI8t4j7Hw="

Private MessageService As New PBMSGService



Private Sub btnCancelReserve_Click()
    Dim Response As PBResponse
    
    Set Response = MessageService.CancelReserve(txtCorpNum.Text, txtReceiptNum.Text, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
End Sub

Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = MessageService.CheckIsMember(txtCorpNum.Text, LinkID)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
End Sub


Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = MessageService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
    
    
End Sub


Private Sub btnGetMessages_Click()
    Dim sentMessages As Collection
    
    Set sentMessages = MessageService.GetMessages(txtCorpNum.Text, txtReceiptNum.Text, txtUserID.Text)
    
    If sentMessages Is Nothing Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    
    Dim sentMessage As PBSentMsg
    
    
    Dim tmp As String
    tmp = "state | subject | messageType | sendnum | receiveNum | receiveName | reserveDT | sendDT | sendResult" + vbCrLf
    
    For Each sentMessage In sentMessages
    
        tmp = tmp + CStr(sentMessage.state) + " | "
        tmp = tmp + sentMessage.subject + " | "
        tmp = tmp + sentMessage.messageType + " | "
        'tmp = tmp + sentMessage.content + " | " ' ���� ǥ�ô� ���̰���� �������� �����մϴ�.
        tmp = tmp + sentMessage.sendNum + " | "
        tmp = tmp + sentMessage.receiveNum + " | "
        tmp = tmp + sentMessage.receiveName + " | "
        tmp = tmp + sentMessage.reserveDT + " | "
        tmp = tmp + sentMessage.sendDT + " | "
        tmp = tmp + sentMessage.resultDT + " | "
        tmp = tmp + sentMessage.sendResult
        
        tmp = tmp + vbCrLf
    Next
    
    
    txtResult.Text = tmp
    
End Sub

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = MessageService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
    
End Sub

Private Sub btnGetPopbillURL_Click()
    Dim url As String
    
    url = MessageService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, cboPopbillTOGO.Text)
    
    If url = "" Then
         MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    joinData.LinkID = LinkID '���� ���̵�
    joinData.CorpNum = "1231212312" '����ڹ�ȣ "-" ����.
    joinData.CEOName = "��ǥ�ڼ���"
    joinData.CorpName = "ȸ����ȣ"
    joinData.Addr = "�ּ�"
    joinData.ZipCode = "500-100"
    joinData.BizType = "����"
    joinData.BizClass = "����"
    joinData.ID = "userid"      '6�� �̻� 20�� �̸�.
    joinData.PWD = "pwd_must_be_long_enough"    '6�� �̻� 20�� �̸�.
    joinData.ContactName = "����ڼ���"
    joinData.ContactTEL = "02-999-9999"
    joinData.ContactHP = "010-1234-5678"
    joinData.ContactFAX = "02-999-9998"
    joinData.ContactEmail = "test@test.com"
    
    Set Response = MessageService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (Response.message)
    
    
End Sub

Private Sub btnSearchPopup_Click()
    Dim url As String
    
    url = MessageService.GetURL(txtCorpNum.Text, txtUserID.Text, "BOX")
    
    If url = "" Then
         MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnSendLMS_Hundred_Click()
    Dim Messages As New Collection
    
    Dim message As PBMessage
    
    Dim i As Integer
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        message.sender = "07075106766"
        message.receiver = "11112222"
        message.receiverName = "�������̸�_" + CStr(i + 1)
        message.content = "�߽� ����. �幮�� 2000Byte�� ���̰� �����Ǿ� ���۵˴ϴ�. �˺��� �ְ��� ���ڼ��ݰ�꼭 ���񽺸� �����ϰ� �ֽ��ϴ�."
        message.subject = "�幮 �����Դϴ�."
        
        Messages.Add message
    Next
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendLMS(txtCorpNum.Text, "", "", "", Messages, txtReserveDT.Text, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
    
End Sub

Private Sub btnSendLMS_One_Click()
    
    Dim Messages As New Collection
    
    Dim message As New PBMessage
    
    message.sender = "07075106766"
    message.receiver = "11112222"
    message.receiverName = "�������̸�"
    message.content = "�߽� ����. �幮�� 2000Byte�� ���̰� �����Ǿ� ���۵˴ϴ�. �˺��� �ְ��� ���ڼ��ݰ�꼭 ���񽺸� �����ϰ� �ֽ��ϴ�."
    message.subject = "�幮 �����Դϴ�."
    
    Messages.Add message
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendLMS(txtCorpNum.Text, "", "", "", Messages, txtReserveDT.Text, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
    
    
End Sub

Private Sub btnSendLMS_Same_Click()
        
    Dim Messages As New Collection
    
    Dim message As PBMessage
    
    Dim i As Integer
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        message.receiver = "11112222"
        message.receiverName = "�������̸�_" + CStr(i + 1)
        
        Messages.Add message
    Next
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendLMS(txtCorpNum.Text, "07075106766", "�������� ����", "�߽� ����. �幮�� 2000Byte�� ���̰� �����Ǿ� ���۵˴ϴ�. �˺��� �ְ��� ���ڼ��ݰ�꼭 ���񽺸� �����ϰ� �ֽ��ϴ�.", _
                                    Messages, txtReserveDT.Text, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendSMS_hundredd_Click()
    Dim Messages As New Collection
    
    Dim message As PBMessage
    
    Dim i As Integer
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        message.sender = "07075106766"
        message.receiver = "11112222"
        message.receiverName = "�������̸�_" + CStr(i + 1)
        message.content = "�߽� ����. �ܹ��� 90Byte�� ���̰� �����Ǿ� ���۵˴ϴ�."
        
        Messages.Add message
    Next
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendSMS(txtCorpNum.Text, "", "", Messages, txtReserveDT.Text, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
    
    
End Sub

Private Sub btnSendSMS_One_Click()
    
    Dim Messages As New Collection
    
    Dim message As New PBMessage
    
    message.sender = "07075106766"
    message.receiver = "11112222"
    message.receiverName = "�������̸�"
    message.content = "�߽� ����. �ܹ��� 90Byte�� ���̰� �����Ǿ� ���۵˴ϴ�."
    
    Messages.Add message
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendSMS(txtCorpNum.Text, "", "", Messages, txtReserveDT.Text, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
    
    
End Sub

Private Sub btnSendSMS_Same_Click()
        
    Dim Messages As New Collection
    
    Dim message As PBMessage
    
    Dim i As Integer
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        message.receiver = "11112222"
        message.receiverName = "�������̸�_" + CStr(i + 1)
        
        Messages.Add message
    Next
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendSMS(txtCorpNum.Text, "07075106766", "�������� ���� 90byte�� ���̰� �����Ǹ�, Messages�� ������ ���� ���Űǿ� ����ó���˴ϴ�.", Messages, txtReserveDT.Text, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
    
End Sub

Private Sub btnSendXMS_Hundred_Click()
    Dim Messages As New Collection
    
    Dim message As PBMessage
    
    Dim i As Integer
    
    For i = 0 To 50
        
        Set message = New PBMessage
        
        message.sender = "07075106766"
        message.receiver = "11112222"
        message.receiverName = "�������̸�_" + CStr(i + 1)
        message.content = "�߽� ����. �� ������ �幮���� ���۵ɼ� �ֵ��� ���̸� �����Ͽ����ϴ�. �˺��� ���� �ְ��� ���ڼ��ݰ�꼭 ���� �Դϴ�."
        message.subject = "�幮 �����Դϴ�."
        
        Messages.Add message
    Next
    
    For i = 0 To 50
        
        Set message = New PBMessage
        
        message.sender = "07075106766"
        message.receiver = "11112222"
        message.receiverName = "�������̸�_" + CStr(i + 1)
        message.content = "�߽� ����. �� ������ �ܹ����� ���۵˴ϴ�."
        
        Messages.Add message
    Next
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendXMS(txtCorpNum.Text, "", "", "", Messages, txtReserveDT.Text, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendXMS_One_Click()
    
    Dim Messages As New Collection
    
    Dim message As New PBMessage
    
    message.sender = "07075106766"
    message.receiver = "01041680206"
    message.receiverName = "�������̸�"
    message.content = "�ڵ��ν� �߼��� ������ ���̸� 90Byte�������� ���ϴ� �ܹ�, �̻��� �幮���� �ڵ� �����մϴ�."
    message.subject = "�幮�� ��� �幮 �����Դϴ�."
    
    Messages.Add message
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendXMS(txtCorpNum.Text, "", "", "", Messages, txtReserveDT.Text, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendXMS_Same_Click()
        
    Dim Messages As New Collection
    
    Dim message As PBMessage
    
    Dim i As Integer
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        message.receiver = "11112222"
        message.receiverName = "�������̸�_" + CStr(i + 1)
        
        Messages.Add message
    Next
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendLMS(txtCorpNum.Text, "07075106766", "�������� ����, �幮�� �����", _
                                        "�ڵ��ν� �߼��� ������ ���̸� 90Byte�������� ���ϴ� �ܹ�, �̻��� �幮���� �ڵ� �����մϴ�.", _
                                        Messages, txtReserveDT.Text, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnUnitCost_Click()
    Dim unitCost As Single
    
    unitCost = MessageService.GetUnitCost(txtCorpNum.Text, SMS)
    
    If unitCost < 0 Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� �ܰ� : " + CStr(unitCost)
    
End Sub

Private Sub btnUnitCost_LMS_Click()
    Dim unitCost As Single
    
    unitCost = MessageService.GetUnitCost(txtCorpNum.Text, LMS)
    
    If unitCost < 0 Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� �ܰ� : " + CStr(unitCost)
End Sub


Private Sub Form_Load()
    MessageService.Initialize LinkID, SecretKey
    MessageService.IsTest = True
    
    
    cboPopbillTOGO.AddItem "LOGIN"
    cboPopbillTOGO.AddItem "CHRG"
   
 
    
End Sub
