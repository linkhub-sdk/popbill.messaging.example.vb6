VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "�˺� �޽�¡ SDK ����"
   ClientHeight    =   12000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17010
   LinkTopic       =   "Form1"
   ScaleHeight     =   12000
   ScaleWidth      =   17010
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton btnSendMMS_hundred 
      Caption         =   "100�� ����"
      Height          =   465
      Left            =   6240
      TabIndex        =   37
      Top             =   5475
      Width           =   1110
   End
   Begin VB.Frame Frame10 
      Caption         =   "���� ���۱��"
      Height          =   945
      Left            =   4920
      TabIndex        =   34
      Top             =   5160
      Width           =   3825
      Begin VB.CommandButton btnSendMMS 
         Caption         =   "1�� ����"
         Height          =   465
         Left            =   120
         TabIndex        =   36
         Top             =   315
         Width           =   1110
      End
      Begin VB.CommandButton btnSendMMS_Same 
         Caption         =   "�뷮����"
         Height          =   465
         Left            =   2520
         TabIndex        =   35
         Top             =   315
         Width           =   1110
      End
   End
   Begin VB.CommandButton btnUnitCost_MMS 
      Caption         =   "MMS ���۴ܰ� Ȯ��"
      Height          =   410
      Left            =   2320
      TabIndex        =   33
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Frame Frame6 
      Caption         =   " �˺� �޽�¡ ���� ���"
      Height          =   8655
      Left            =   240
      TabIndex        =   13
      Top             =   3240
      Width           =   13005
      Begin VB.Frame Frame17 
         Caption         =   "��û��ȣ �Ҵ� ���۰� ó��"
         Height          =   1335
         Left            =   4920
         TabIndex        =   62
         Top             =   3000
         Width           =   4215
         Begin VB.CommandButton btnCancelReserveRN 
            Caption         =   "���� ���� ���"
            Height          =   525
            Left            =   2160
            TabIndex        =   65
            Top             =   720
            Width           =   1935
         End
         Begin VB.CommandButton btnGetMessagesRN 
            Caption         =   "���ۻ���Ȯ��"
            Height          =   525
            Left            =   120
            TabIndex        =   64
            Top             =   720
            Width           =   1905
         End
         Begin VB.TextBox txtRequestNum 
            Height          =   435
            Left            =   1200
            TabIndex        =   63
            Top             =   240
            Width           =   2850
         End
         Begin VB.Label ��û��ȣ 
            Caption         =   "��û��ȣ : "
            Height          =   375
            Left            =   240
            TabIndex        =   66
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "�߽Ź�ȣ ����"
         Height          =   1455
         Left            =   10680
         TabIndex        =   50
         Top             =   240
         Width           =   2055
         Begin VB.CommandButton btnGetSenderNumberMgtURL 
            Caption         =   "�߽Ź�ȣ ���� �˾�"
            Height          =   495
            Left            =   120
            TabIndex        =   52
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton btnGetSenderNumberList 
            Caption         =   "�߽Ź�ȣ ��� ��ȸ"
            Height          =   495
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.CommandButton btnGetAutoDenyList 
         Caption         =   "080 ���Űźθ��"
         Height          =   495
         Left            =   8760
         TabIndex        =   47
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "���۳��� �����ȸ"
         Height          =   495
         Left            =   8760
         TabIndex        =   46
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton btnGetSentListURL 
         Caption         =   "���۳�����ȸ �˾�"
         Height          =   495
         Left            =   8760
         TabIndex        =   32
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtResult 
         Height          =   3840
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   3  '�����
         TabIndex        =   31
         Top             =   4560
         Width           =   12255
      End
      Begin VB.CommandButton btnCancelReserve 
         Caption         =   "���� ���� ���"
         Height          =   525
         Left            =   2640
         TabIndex        =   30
         Top             =   3720
         Width           =   1905
      End
      Begin VB.CommandButton btnGetMessages 
         Caption         =   "���ۻ���Ȯ��"
         Height          =   525
         Left            =   600
         TabIndex        =   29
         Top             =   3720
         Width           =   1905
      End
      Begin VB.Frame Frame9 
         Caption         =   " ��/�幮 �ڵ��ν� ���� ���� "
         Height          =   945
         Left            =   480
         TabIndex        =   25
         Top             =   1920
         Width           =   3945
         Begin VB.CommandButton btnSendXMS_One 
            Caption         =   "1�� ����"
            Height          =   465
            Left            =   240
            TabIndex        =   28
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendXMS_Hundred 
            Caption         =   "100�� ����"
            Height          =   465
            Left            =   1440
            TabIndex        =   27
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendXMS_Same 
            Caption         =   "�뷮����"
            Height          =   465
            Left            =   2640
            TabIndex        =   26
            Top             =   315
            Width           =   1110
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   " �幮 ���� ���� "
         Height          =   945
         Left            =   4680
         TabIndex        =   21
         Top             =   840
         Width           =   3825
         Begin VB.CommandButton btnSendLMS_One 
            Caption         =   "1�� ����"
            Height          =   465
            Left            =   120
            TabIndex        =   24
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendLMS_Hundred 
            Caption         =   "100�� ����"
            Height          =   465
            Left            =   1320
            TabIndex        =   23
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendLMS_Same 
            Caption         =   "�뷮����"
            Height          =   465
            Left            =   2520
            TabIndex        =   22
            Top             =   315
            Width           =   1110
         End
      End
      Begin VB.TextBox txtReceiptNum 
         Height          =   435
         Left            =   1680
         TabIndex        =   20
         Top             =   3240
         Width           =   2850
      End
      Begin VB.Frame Frame7 
         Caption         =   " �ܹ� ���� ���� "
         Height          =   945
         Left            =   480
         TabIndex        =   16
         Top             =   840
         Width           =   3945
         Begin VB.CommandButton btnSendSMS_Same 
            Caption         =   "�뷮����"
            Height          =   465
            Left            =   2640
            TabIndex        =   19
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendSMS_hundredd 
            Caption         =   "100�� ����"
            Height          =   465
            Left            =   1440
            TabIndex        =   18
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendSMS_One 
            Caption         =   "1�� ����"
            Height          =   465
            Left            =   240
            TabIndex        =   17
            Top             =   315
            Width           =   1110
         End
      End
      Begin VB.TextBox txtReserveDT 
         Height          =   315
         Left            =   3540
         TabIndex        =   14
         Top             =   375
         Width           =   3105
      End
      Begin VB.Frame Frame13 
         Caption         =   "�ΰ����"
         Height          =   2655
         Left            =   8640
         TabIndex        =   48
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton btnGetStates 
            Caption         =   "���۳��� �������"
            Height          =   495
            Left            =   120
            TabIndex        =   59
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "������ȣ ���� ��� (��û��ȣ ���Ҵ�)"
         Height          =   1335
         Left            =   480
         TabIndex        =   60
         Top             =   3000
         Width           =   4335
         Begin VB.Label ������ȣ 
            Caption         =   "������ȣ : "
            Height          =   255
            Left            =   240
            TabIndex        =   61
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "����ð�(yyyyMMddHHmmss) : "
         Height          =   180
         Left            =   705
         TabIndex        =   15
         Top             =   450
         Width           =   2790
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " �˺� �⺻ API "
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   16560
      Begin VB.Frame Frame16 
         Caption         =   "��Ʈ�ʰ��� ����Ʈ"
         Height          =   1935
         Left            =   8040
         TabIndex        =   54
         Top             =   240
         Width           =   2415
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "��Ʈ�� ����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   58
            Top             =   840
            Width           =   2175
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "��Ʈ�� �ܿ�����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "�������� ����Ʈ"
         Height          =   1935
         Left            =   6000
         TabIndex        =   53
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   "����Ʈ ���� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   56
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ� ����Ʈ Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   55
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "ȸ������ ����"
         Height          =   1935
         Left            =   14520
         TabIndex        =   43
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "ȸ������ ����"
            Height          =   410
            Left            =   120
            TabIndex        =   45
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "ȸ������ ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "����� ����"
         Height          =   1935
         Left            =   12480
         TabIndex        =   39
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "����� ���� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   42
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "����� ��� ��ȸ"
            Height          =   410
            Left            =   120
            TabIndex        =   41
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "����� �߰�"
            Height          =   410
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " ȸ������"
         Height          =   1935
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1695
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID �ߺ� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   38
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "���� ���� Ȯ��"
            Height          =   410
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "ȸ�� ����"
            Height          =   410
            Left            =   120
            TabIndex        =   10
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "���۴ܰ�"
         Height          =   1935
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   3840
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "�������� Ȯ��"
            Height          =   410
            Left            =   2160
            TabIndex        =   49
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnUnitCost_LMS 
            Caption         =   "LMS ���۴ܰ� Ȯ��"
            Height          =   410
            Left            =   160
            TabIndex        =   12
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "SMS ���۴ܰ� Ȯ��"
            Height          =   410
            Left            =   150
            TabIndex        =   8
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " �˺� �⺻ URL"
         ClipControls    =   0   'False
         Height          =   1935
         Left            =   10560
         TabIndex        =   5
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnGetAccessURL 
            Caption         =   " �˺� �α��� URL"
            Height          =   410
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1575
         End
      End
   End
   Begin VB.TextBox txtUserID 
      Height          =   315
      Left            =   6240
      TabIndex        =   3
      Text            =   "testkorea"
      Top             =   165
      Width           =   1935
   End
   Begin VB.TextBox txtCorpNum 
      Height          =   315
      Left            =   2295
      TabIndex        =   1
      Text            =   "1234567890"
      Top             =   180
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9840
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�˺�ȸ�� ���̵� : "
      Height          =   180
      Left            =   4680
      TabIndex        =   2
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�˺�ȸ�� ����ڹ�ȣ : "
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1860
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' �˺� ���� API VB 6.0 SDK Example
'
' - VB6 SDK ����ȯ�� ������� �ȳ� : http://blog.linkhub.co.kr/569/
' - ������Ʈ ���� : 2018-11-21
' - ���� ������� ����ó : 1600-9854 / 070-4304-2991
' - ���� ������� �̸��� : code@linkhub.co.kr
'
' <�׽�Ʈ �������� �غ����>
' 1) 25, 28�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
' 2) �˺� ���߿� ����Ʈ(test.popbill.com)�� ����ȸ������ �����մϴ�.
'=========================================================================

Option Explicit

'=========================================================================
' - ��������(��ũ���̵�, ���Ű)�� ��Ʈ���� ����ȸ���� �ĺ��ϴ�
'   ������ ���Ǵ� ������ ������� �ʵ��� �����Ͻñ� �ٶ��ϴ�.
' - ����� ��ȯ���Ŀ��� ��������(��ũ���̵�, ���Ű)�� ������� �ʽ��ϴ�.
'=========================================================================

'��ũ���̵�
Private Const LinkID = "TESTER"

'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

Private MessageService As New PBMSGService

'=========================================================================
' ���๮�������� ����մϴ�.
' - ������Ҵ� �������۽ð� 10���������� �����մϴ�.
'=========================================================================
Private Sub btnCancelReserve_Click()
    Dim Response As PBResponse
    
    Set Response = MessageService.CancelReserve(txtCorpNum.Text, txtReceiptNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �������ۿ�û�� �Ҵ��� ���ۿ�û��ȣ(requestNum)�� ���๮�� ������ ����մϴ�.
' - ��������Ҵ� �������۽ð� 10���������� �����մϴ�.
'=========================================================================
Private Sub btnCancelReserveRN_Click()
    Dim Response As PBResponse
    
    Set Response = MessageService.CancelReserveRN(txtCorpNum.Text, txtRequestNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �˺� ȸ�����̵� �ߺ����θ� Ȯ���մϴ�.
'=========================================================================

Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = MessageService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �ش� ������� ��Ʈ�� ����ȸ�� ���Կ��θ� Ȯ���մϴ�.
' - LinkID�� ���������� �����Ǿ� �ִ� ��ũ���̵� ���Դϴ�.
'=========================================================================

Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = MessageService.CheckIsMember(txtCorpNum.Text, LinkID)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub



'=========================================================================
' 080 ���� ���Űź� ����� Ȯ���մϴ�.
'=========================================================================

Private Sub btnGetAutoDenyList_Click()
    Dim AutoDenyList As Collection
    Dim tmp As String
    Dim AutoDenyInfo As PBAutoDenyInfo
    
    Set AutoDenyList = MessageService.GetAutoDenyList(txtCorpNum.Text)
    
    If AutoDenyList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "number(���Űźι�ȣ) | regDT(����Ͻ�)" + vbCrLf
    
    For Each AutoDenyInfo In AutoDenyList
        tmp = tmp + AutoDenyInfo.number + " | " + AutoDenyInfo.regDT + vbCrLf
    Next
    
    MsgBox tmp
    
End Sub

'=========================================================================
' ����ȸ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ��Ʈ�ʰ����� ��� ��Ʈ�� �ܿ�����Ʈ(GetPartnerBalance API)
'   �� ���� Ȯ���Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = MessageService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
    
End Sub

'=========================================================================
' ����ȸ���� ���� API ���� ���������� Ȯ���մϴ�.
'=========================================================================

Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim MType As MsgType
    Dim tmp As String
    
    MType = SMS     'SMS-�ܹ�, LMS-�幮 MMS-����
            
    Set ChargeInfo = MessageService.GetChargeInfo(txtCorpNum.Text, MType)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (���۴ܰ�) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (��������) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (��������) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub





'=========================================================================
' ����ȸ���� ȸ�������� Ȯ���մϴ�.
'=========================================================================

Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = MessageService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname(��ǥ�ڼ���) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName(��ȣ��) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr(�ּ�) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType(����) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass(����) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' �������ۿ�û�� ���� ���۰���� Ȯ���մϴ�.
'=========================================================================

Private Sub btnGetMessages_Click()
    Dim sentMessages As Collection
    Dim sentMessage As PBSentMsg
    Dim tmp As String
    
    Set sentMessages = MessageService.GetMessages(txtCorpNum.Text, txtReceiptNum.Text)
    
    If sentMessages Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "state | result | subject | messageType | sendnum | senderName | receiveNum | receiveName | receiptDT | reserveDT | sendDT | tranNet | receiptNum | requestNum" + vbCrLf
    
    For Each sentMessage In sentMessages
            
        ' ���ۻ��� �ڵ�
        tmp = tmp + CStr(sentMessage.state) + " | "
        
        ' ���۰�� �ڵ�
        tmp = tmp + CStr(sentMessage.result) + " | "
        
        ' �޽��� ����
        tmp = tmp + sentMessage.subject + " | "
        
        ' �޽��� ����
        tmp = tmp + sentMessage.messageType + " | "
        
        ' �޽��� ����
        'tmp = tmp + sentMessage.content + " | " ' ���� ǥ�ô� ���̰���� �������� �����մϴ�.
        
        ' �߽Ź�ȣ
        tmp = tmp + sentMessage.sendNum + " | "
        
        ' �߽��ڸ�
        tmp = tmp + sentMessage.senderName + " | "
        
        ' ���Ź�ȣ
        tmp = tmp + sentMessage.receiveNum + " | "
        
        ' �����ڸ�
        tmp = tmp + sentMessage.receiveName + " | "
        
        ' �����Ͻ�
        tmp = tmp + sentMessage.receiptDT + " | "
        
        ' �����Ͻ�
        tmp = tmp + sentMessage.reserveDT + " | "
        
        ' �����Ͻ�
        tmp = tmp + sentMessage.sendDT + " | "
        
        ' ���۰�� �����Ͻ�
        tmp = tmp + sentMessage.resultDT + " | "
        
        ' ����ó�� �̵���Ż��
        tmp = tmp + sentMessage.tranNet + " | "
        
        ' ������ȣ
        tmp = tmp + sentMessage.receiptNum + " | "
        
        ' ��û��ȣ
        tmp = tmp + sentMessage.requestNum
        
        tmp = tmp + vbCrLf
    Next
    
    txtResult.Text = tmp
    
End Sub

'=========================================================================
' �������ۿ�û�� �Ҵ��� ���ۿ�û��ȣ(requestNum)�� ���ۻ��¸� Ȯ���մϴ�
' - �����׸� ���� �ڼ��� ������ "[���� API �����Ŵ���] >
'  3.3.2. GetMessagesRN (���۳��� Ȯ�� - ��û��ȣ �Ҵ�)�� �����Ͻñ� �ٶ��ϴ�.
'=========================================================================
Private Sub btnGetMessagesRN_Click()
Dim sentMessages As Collection
    Dim sentMessage As PBSentMsg
    Dim tmp As String
    
    Set sentMessages = MessageService.GetMessagesRN(txtCorpNum.Text, txtRequestNum.Text)
    
    If sentMessages Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "state | result | subject | messageType | sendnum | senderName | receiveNum | receiveName | receiptDT | reserveDT | sendDT | tranNet | receiptNum | requestNum" + vbCrLf
    
    For Each sentMessage In sentMessages
            
        ' ���ۻ��� �ڵ�
        tmp = tmp + CStr(sentMessage.state) + " | "
        
        ' ���۰�� �ڵ�
        tmp = tmp + CStr(sentMessage.result) + " | "
        
        ' �޽��� ����
        tmp = tmp + sentMessage.subject + " | "
        
        ' �޽��� ����
        tmp = tmp + sentMessage.messageType + " | "
        
        ' �޽��� ����
        'tmp = tmp + sentMessage.content + " | " ' ���� ǥ�ô� ���̰���� �������� �����մϴ�.
        
        ' �߽Ź�ȣ
        tmp = tmp + sentMessage.sendNum + " | "
        
        ' �߽��ڸ�
        tmp = tmp + sentMessage.senderName + " | "
        
        ' ���Ź�ȣ
        tmp = tmp + sentMessage.receiveNum + " | "
        
        ' �����ڸ�
        tmp = tmp + sentMessage.receiveName + " | "
        
        ' �����Ͻ�
        tmp = tmp + sentMessage.receiptDT + " | "
        
        ' �����Ͻ�
        tmp = tmp + sentMessage.reserveDT + " | "
        
        ' �����Ͻ�
        tmp = tmp + sentMessage.sendDT + " | "
        
        ' ���۰�� �����Ͻ�
        tmp = tmp + sentMessage.resultDT + " | "
        
        ' ����ó�� �̵���Ż��
        tmp = tmp + sentMessage.tranNet + " | "
        
        ' ������ȣ
        tmp = tmp + sentMessage.receiptNum + " | "
                
        ' ��û��ȣ
        tmp = tmp + sentMessage.requestNum
        
        tmp = tmp + vbCrLf
    Next
    
    txtResult.Text = tmp
    
End Sub


'=========================================================================
' ��Ʈ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ���������� ��� ����ȸ�� �ܿ�����Ʈ(GetBalance API)��
'   �̿��Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = MessageService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
    
End Sub

'=========================================================================
' ��Ʈ�� ����Ʈ ���� URL�� ��ȯ�մϴ�.
' - URL ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
    
    url = MessageService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ���� URL�� ��ȯ�մϴ�.
' - URL ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim url As String
    
    url = MessageService.GetChargeURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �˺�(www.popbill.com)�� �α��ε� �˺� URL�� ��ȯ�մϴ�.
' - ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetAccessURL_Click()

    Dim url As String
    
    url = MessageService.GetAccessURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ���� �߽Ź�ȣ ����� ��ȸ�մϴ�.
'=========================================================================

Private Sub btnGetSenderNumberList_Click()
    Dim SenderNumberList As Collection
    Dim tmp As String
    Dim SenderNumber As PBMsgSenderNumber
    
    Set SenderNumberList = MessageService.GetSenderNumberList(txtCorpNum.Text)
    
    If SenderNumberList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
        
    For Each SenderNumber In SenderNumberList
        tmp = tmp + "�߽Ź�ȣ(number) : " + SenderNumber.number + vbCrLf
        tmp = tmp + "��ǥ��ȣ ��������(representYN) : " + CStr(SenderNumber.representYN) + vbCrLf
        tmp = tmp + "��ϻ���(state) : " + CStr(SenderNumber.state) + vbCrLf + vbCrLf
        
    Next
    
    MsgBox tmp
    
End Sub





'=========================================================================
' ���� ���۳��� ��������� Ȯ���մϴ�.
'=========================================================================

Private Sub btnGetStates_Click()
    Dim resultList As Collection
    Dim ReciptNumList As New Collection
    
    ', �ִ� 1000��
    ReciptNumList.Add "018061814000000039"
    ReciptNumList.Add "018061815000000002"

    
    Set resultList = MessageService.GetStates(txtCorpNum.Text, ReciptNumList)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "rNum | sn | stat | rlt | sDT | rDT | net | srt" + vbCrLf
    
    Dim info As PBMessageBriefInfo
    
    For Each info In resultList
        tmp = tmp + info.rNum + " | " + info.sn + " | " + info.stat + " | " + info.rlt + " | " + info.sDT + " | "
        tmp = tmp + info.rDT + " | " + info.net + " | " + info.srt + vbCrLf
    Next
    
    MsgBox tmp

End Sub

'=========================================================================
' �߽Ź�ȣ ���� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�� URL�� ������å�� ���� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetSenderNumberMgtURL_Click()

    Dim url As String
    
    url = MessageService.GetSenderNumberMgtURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
        
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' �˺� ����ȸ�� ������ ��û�մϴ�.
'=========================================================================

Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
 
    '��ũ ���̵�
    joinData.LinkID = LinkID
    
    '����ڹ�ȣ, '-'����, 10�ڸ�
    joinData.CorpNum = "1231212312"
    
    '��ǥ�ڼ���, �ִ� 30��
    joinData.ceoname = "��ǥ�ڼ���"
    
    '��ȣ��, �ִ� 70��
    joinData.corpName = "ȸ����ȣ"
    
    '�ּ�, �ִ� 300��
    joinData.addr = "�ּ�"
    
    '����, �ִ� 40��
    joinData.bizType = "����"
    
    '����, �ִ� 40��
    joinData.bizClass = "����"
    
    '���̵�, 6���̻� 20�� �̸�
    joinData.id = "userid"
    
    '��й�ȣ, 6���̻� 20�� �̸�
    joinData.pwd = "pwd_must_be_long_enough"
    
    '����ڸ�, �ִ� 30��
    joinData.ContactName = "����ڼ���"
    
    '����� ����ó, �ִ� 20��
    joinData.ContactTEL = "02-999-9999"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.ContactHP = "010-1234-5678"
    
    '����� �ѽ���ȣ, �ִ� 20��
    joinData.ContactFAX = "02-999-9998"
    
    '����� ����, �ִ� 70��
    joinData.ContactEmail = "test@test.com"
    
    Set Response = MessageService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
    
End Sub

'=========================================================================
' ����ȸ���� ����� ����� Ȯ���մϴ�.
'=========================================================================

Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = MessageService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id | email | hp | personName | searchAllAllowYN | tel | fax | mgrYN | regDT | state" + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.email + " | " + info.hp + " | " + info.personName + " | " + CStr(info.searchAllAllowYN) _
                + info.tel + " | " + info.fax + " | " + CStr(info.mgrYN) + " | " + info.regDT + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ���� ����ڸ� �űԷ� ����մϴ�.
'=========================================================================

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�, 6�� �̻� 20�� �̸�
    joinData.id = "testkorea_20161011"
    
    '��й�ȣ, 6�� �̻� 20�� �̸�
    joinData.pwd = "test@test.com"
    
    '����ڸ�, �ִ� 30��
    joinData.personName = "����ڸ�"
    
    '����� ����ó
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ
    joinData.hp = "010-1234-1234"
    
    '����� �����ּ�
    joinData.email = "test@test.com"
    
    '����� �ѽ���ȣ
    joinData.fax = "070-1234-1234"
    
    'ȸ����ȸ ���ѿ���, true-ȸ����ȸ / false-������ȸ
    joinData.searchAllAllowYN = True
    
    '������ ���ѿ���
    joinData.mgrYN = False
        
    Set Response = MessageService.RegistContact(txtCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �˻������� ����Ͽ� �������� ������ ��ȸ�մϴ�.
'  - �ִ� �˻��Ⱓ : 6���� �̳�
'=========================================================================

Private Sub btnSearch_Click()
    Dim msgSearchList As PBSearchList
    Dim SDate As String
    Dim EDate As String
    Dim state As New Collection
    Dim Item As New Collection
    Dim ReserveYN As Boolean
    Dim SenderYN As Boolean
    Dim Page As Integer
    Dim PerPage As Integer
    Dim Order As String
    Dim tmp As String
    Dim info As PBSentMsg
    Dim QString As String
    
    '[�ʼ�] ��������, ��������(yyyyMMdd)
    SDate = "20180801"
    
    '[�ʼ�] ��������, ��������(yyyyMMdd)
    EDate = "20180820"
    
    '���ۻ��°� �迭, 1-���, 2-����, 3-����, 4-���
    state.Add "1"
    state.Add "2"
    state.Add "3"
    state.Add "4"
    
    '�˻���� �迭, SMS(�ܹ�),LMS(�幮),MMS(����)
    Item.Add "SMS"
    Item.Add "LMS"
    Item.Add "MMS"
    
    '���๮�� �˻�����, True(���๮�ڸ� ��ȸ), False(��ü��ȸ)
    ReserveYN = False
    
    '������ȸ����, True(������ȸ), False(��ü��ȸ)
    SenderYN = False
    
    '������ ��ȣ
    Page = 1
    
    '������ ��ϰ���, �ִ� 1000��
    PerPage = 50
    
    '���Ĺ���, D-��������(�⺻��), A-��������
    Order = "D"
    
    '��ȸ �˻���, �߽��ڸ� �Ǵ� �����ڸ� ����
    QString = ""

    Set msgSearchList = MessageService.Search(txtCorpNum.Text, SDate, EDate, state, Item, ReserveYN, SenderYN, Page, PerPage, Order, QString)
     
    If msgSearchList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "code : " + CStr(msgSearchList.code) + vbCrLf
    tmp = tmp + "total : " + CStr(msgSearchList.total) + vbCrLf
    tmp = tmp + "perPage : " + CStr(msgSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum : " + CStr(msgSearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount : " + CStr(msgSearchList.pageCount) + vbCrLf
    tmp = tmp + "message : " + msgSearchList.message + vbCrLf + vbCrLf
    
    tmp = tmp + "state | result | subject | messageType | sendnum | senderName | receiveNum | receiveName | receiptDT | reserveDT | sendDT |  tranNet | receiptNum | requestNum" + vbCrLf
            
    For Each info In msgSearchList.list
    
        ' ���ۻ��� �ڵ�
        tmp = tmp + CStr(info.state) + " | "
        
        ' ���۰�� �ڵ�
        tmp = tmp + CStr(info.result) + " | "
        
        ' �޽��� ����
        tmp = tmp + info.subject + " | "
        
        ' �޽��� ����
        tmp = tmp + info.messageType + " | "
        
        ' �޽��� ����
        'tmp = tmp + sentMessage.content + " | " ' ���� ǥ�ô� ���̰���� �������� �����մϴ�.
        
        ' �߽Ź�ȣ
        tmp = tmp + info.sendNum + " | "
        
        ' �߽��ڸ�
        tmp = tmp + info.senderName + " | "
        
        ' ���Ź�ȣ
        tmp = tmp + info.receiveNum + " | "
        
        ' �����ڸ�
        tmp = tmp + info.receiveName + " | "
        
        ' �����Ͻ�
        tmp = tmp + info.receiptDT + " | "
        
        ' �����Ͻ�
        tmp = tmp + info.reserveDT + " | "
        
        ' ���۽ð�
        tmp = tmp + info.sendDT + " | "
        
        ' ���۰�� ���Žð�
        tmp = tmp + info.resultDT + " | "
        
        ' ����ó�� �̵���Ż��
        tmp = tmp + info.tranNet + " | "
        
        ' ������ȣ
        tmp = tmp + info.receiptNum + " | "
        
        ' ��û��ȣ
        tmp = tmp + info.requestNum
        tmp = tmp + vbCrLf
    Next
        
    txtResult.Text = tmp
End Sub

'=========================================================================
' ���ڸ޽��� ���۳��� �˾� URL�� ��ȯ�մϴ�.
' - ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================
Private Sub btnGetSentListURL_Click()
    Dim url As String
    
    url = MessageService.GetSentListURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnSendLMS_Hundred_Click()
    Dim Messages As New Collection
    Dim adsYN As Boolean
    Dim message As PBMessage
    Dim i As Integer
    Dim receiptNum As String
    Dim requestNum As String
    Dim UserID As String
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        '�߽Ź�ȣ
        message.sender = "07043042991"
        
        '�߽��ڸ�
        message.senderName = "�߽��ڸ�"
        
        '���Ź�ȣ
        message.receiver = "010111222"
        
        '�����ڸ�
        message.receiverName = "�������̸�_" + CStr(i + 1)
        
        '�޽��� ����, 2000byte �ʰ��� ������ �����Ǿ� ���۵�.
        message.content = "�߽� ����. �幮�� 2000Byte�� ���̰� �����Ǿ� ���۵˴ϴ�. �˺��� �ְ��� ���ڼ��ݰ�꼭 ���񽺸� �����ϰ� �ֽ��ϴ�."
        
        '�޽���
        message.subject = "�幮 �����Դϴ�."
        
        Messages.Add message
    Next
        
    '������ ���ۿ���
    adsYN = False
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    '�˺� ȸ�����̵�
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendLMS(txtCorpNum.Text, "", "", "", Messages, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

Private Sub btnSendLMS_One_Click()
    Dim Messages As New Collection
    Dim adsYN As Boolean
    Dim message As New PBMessage
    Dim receiptNum As String
    Dim requestNum As String
    Dim UserID As String
    
    '�߽Ź�ȣ
    message.sender = "07043042991"
    
    '�߽��ڸ�
    message.senderName = "�߽��ڸ�"
    
    '���Ź�ȣ
    message.receiver = "010111222"
    
    '�����ڸ�
    message.receiverName = "�������̸�"
    
    '�幮�޽��� ����
    message.subject = "�幮 �����Դϴ�."
    
    '�޽�������, 2000byte �ʰ��� ������ �����Ǿ� ���۵�
    message.content = "�߽� ����. �幮�� 2000Byte�� ���̰� �����Ǿ� ���۵˴ϴ�. �˺��� �ְ��� ���ڼ��ݰ�꼭 ���񽺸� �����ϰ� �ֽ��ϴ�."
    
    Messages.Add message
    
    
    '������ ���ۿ���
    adsYN = False
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    '�˺� ȸ�����̵�
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendLMS(txtCorpNum.Text, "", "", "", Messages, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

Private Sub btnSendLMS_Same_Click()
    Dim Messages As New Collection
    Dim sender As String
    Dim subject As String
    Dim Contents As String
    Dim adsYN As Boolean
    Dim message As PBMessage
    Dim i As Integer
    Dim receiptNum As String
    Dim requestNum As String
    Dim UserID As String
    
    '�߽Ź�ȣ
    sender = "07043042991"
    
    '�������� ����
    subject = "�������� ����"
    
    '�������� �޽���
    Contents = "�߽� ����. �幮�� 2000Byte�� ���̰� �����Ǿ� ���۵˴ϴ�. �˺��� �ְ��� ���ڼ��ݰ�꼭 ���񽺸� �����ϰ� �ֽ��ϴ�."
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        '���Ź�ȣ
        message.receiver = "010111222"
        
        '�����ڸ�
        message.receiverName = "�������̸�_" + CStr(i + 1)
        
        Messages.Add message
    Next
    
    '������ ���ۿ���
    adsYN = False
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    '�˺� ȸ�����̵�
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendLMS(txtCorpNum.Text, sender, subject, Contents, _
                                    Messages, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

Private Sub btnSendMMS_Click()
    Dim Messages As New Collection
    Dim FilePaths As New Collection
    Dim adsYN As Boolean
    Dim receiptNum As String
    Dim message As New PBMessage
    Dim requestNum As String
    Dim UserID As String
    
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    '���� �޽��� ���ϰ��
    FilePaths.Add CommonDialog1.FileName
    
    
    '�߽Ź�ȣ
    message.sender = "07043042991"
    
    '�߽��ڸ�
    message.senderName = "�߽��ڸ�"
    
    '���Ź�ȣ
    message.receiver = "010111222"
    
    '�����ڸ�
    message.receiverName = "�������̸�"
    
    '���� �޽��� ����
    message.subject = "�޽��� ����"
    
    '���� �޽��� ����
    message.content = "MMS �߽� �׽�Ʈ ����."
    
    Messages.Add message
    
    '������ ���ۿ���
    adsYN = False
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    '�˺� ȸ�����̵�
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendMMS(txtCorpNum.Text, "", "", "", Messages, FilePaths, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

Private Sub btnSendMMS_hundred_Click()
    Dim Messages As New Collection
    Dim FilePaths As New Collection
    Dim adsYN As Boolean
    Dim message As PBMessage
    Dim i As Integer
    Dim receiptNum As String
    Dim requestNum As String
    Dim UserID As String
    
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    FilePaths.Add CommonDialog1.FileName
  
    For i = 0 To 50
        
        Set message = New PBMessage
        
        '�߽Ź�ȣ
        message.sender = "07043042991"
        
        '�߽��ڸ�
        message.senderName = "�߽��ڸ�"
        
        '���Ź�ȣ
        message.receiver = "010111222"
        
        '�����ڸ�
        message.receiverName = "�������̸�_" + CStr(i + 1)
        
        '�޽��� ����
        message.subject = "����޽��� �����Դϴ�."
        
        '�޽��� ����
        message.content = "�߽� ����. �� ������ �幮���� ���۵ɼ� �ֵ��� ���̸� �����Ͽ����ϴ�. �˺��� ���� �ְ��� ���ڼ��ݰ�꼭 ���� �Դϴ�."
        
        Messages.Add message
    Next
    
    For i = 0 To 50
        
        Set message = New PBMessage
        
        '�߽Ź�ȣ
        message.sender = "07043042991"
        
        '�߽��ڸ�
        message.senderName = "�߽��ڸ�"
        
        '���Ź�ȣ
        message.receiver = "010111222"
        
        '�����ڸ�
        message.receiverName = "�������̸�_" + CStr(i + 1)
        
        '�޽��� ����
        message.subject = "���� �޽��� ����"
        
        '�޽��� ����
        message.content = "�߽� ����. �� ������ �ܹ����� ���۵˴ϴ�."
        
        Messages.Add message
    Next
    
    '������ ���ۿ���
    adsYN = False
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    '�˺� ȸ�����̵�
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendMMS(txtCorpNum.Text, "", "", "", Messages, FilePaths, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    txtReceiptNum.Text = receiptNum
End Sub

Private Sub btnSendMMS_Same_Click()
    Dim Messages As New Collection
    Dim FilePaths As New Collection
    Dim adsYN As Boolean
    Dim receiptNum As String
    Dim message As PBMessage
    Dim sender As String
    Dim subject As String
    Dim Contents As String
    Dim i As Integer
    Dim requestNum As String
    Dim UserID As String
    
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    FilePaths.Add CommonDialog1.FileName
    
    '�߽Ź�ȣ
    sender = "07043042991"
    
    '�����޽��� ����
    subject = "�����޽��� ����"
    
    '�����޽��� ����
    Contents = "�����޽��� ����"
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        '���Ź�ȣ
        message.receiver = "010111222"
        
        '�����ڸ�
        message.receiverName = "�������̸�_" + CStr(i + 1)
        
        Messages.Add message
    Next
   
    
    '������ ���ۿ���
    adsYN = False
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    '�˺� ȸ�����̵�
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendMMS(txtCorpNum.Text, sender, subject, Contents, Messages, FilePaths, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

Private Sub btnSendSMS_hundredd_Click()
    Dim Messages As New Collection
    Dim adsYN As Boolean
    Dim message As PBMessage
    Dim i As Integer
    Dim receiptNum As String
    Dim requestNum As String
    Dim UserID As String
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        '�߽Ź�ȣ
        message.sender = "07043042991"
        
        '�߽��ڸ�
        message.senderName = "�߽��ڸ�"
        
        '���Ź�ȣ
        message.receiver = "010111222"
        
        '�����ڸ�
        message.receiverName = "�������̸�_" + CStr(i + 1)
        
        '�޽��� ����, 90Byte �ʰ��� ������ �����Ǿ� ���۵�
        message.content = "�߽� ����. �ܹ��� 90Byte�� ���̰� �����Ǿ� ���۵˴ϴ�."
        
        Messages.Add message
    Next
    
    '������ ���ۿ���
    adsYN = False
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    '�˺� ȸ�����̵�
    UserID = txtUserID.Text
    
    
    receiptNum = MessageService.SendSMS(txtCorpNum.Text, "", "", Messages, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

Private Sub btnSendSMS_One_Click()
    Dim Messages As New Collection
    Dim adsYN As Boolean
    Dim message As New PBMessage
    Dim receiptNum As String
    Dim requestNum As String
    Dim UserID As String
    
    '�߽Ź�ȣ
    message.sender = "07043042991"
    
    '�߽��ڸ�
    message.senderName = "�߽��ڸ�"
    
    '���Ź�ȣ
    message.receiver = "010111222"
    
    '�����ڸ�
    message.receiverName = "�������̸�"
    
    '�޽��� ����
    message.content = "�߽� ����. �ܹ��� 90Byte�� ���̰� �����Ǿ� ���۵˴ϴ�."
    
    Messages.Add message
    
    '������ ���ۿ���
    adsYN = False
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    '�˺� ȸ�����̵�
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendSMS(txtCorpNum.Text, "", "", Messages, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    txtReceiptNum.Text = receiptNum
    
End Sub

Private Sub btnSendSMS_Same_Click()
    Dim Messages As New Collection
    Dim adsYN As Boolean
    Dim sender As String
    Dim Contents As String
    Dim message As PBMessage
    Dim i As Integer
    Dim receiptNum As String
    Dim requestNum As String
    Dim UserID As String
    
    '�߽Ź�ȣ
    sender = "07043042991"
    
    '�޽��� ����, 90byte �ʰ��� ������ �����Ǿ� ���۵�.
    Contents = "�������� ���� 90byte�� ���̰� �����Ǹ�, Messages�� ������ ���� ���Űǿ� ����ó���˴ϴ�."
    
    '������ ���ۿ���
    adsYN = False
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        '���Ź�ȣ
        message.receiver = "010111222"
        
        '�����ڸ�
        message.receiverName = "�������̸�_" + CStr(i + 1)
        
        Messages.Add message
    Next
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    '�˺� ȸ�����̵�
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendSMS(txtCorpNum.Text, sender, Contents, Messages, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

Private Sub btnSendXMS_Hundred_Click()
    Dim Messages As New Collection
    Dim adsYN As Boolean
    Dim receiptNum As String
    Dim message As PBMessage
    Dim i As Integer
    Dim requestNum As String
    Dim UserID As String
    
    For i = 0 To 50
        
        Set message = New PBMessage
        
        '�߽Ź�ȣ
        message.sender = "07043042991"
        
        '�߽��ڸ�
        message.senderName = "�߽��ڸ�"
        
        '���Ź�ȣ
        message.receiver = "010111222"
        
        '�����ڸ�
        message.receiverName = "�������̸�_" + CStr(i + 1)
        
        '�幮�޽��� ����
        message.subject = "�幮 �����Դϴ�."
        
        '�޽��� ����
        message.content = "�߽� ����. �� ������ �幮���� ���۵ɼ� �ֵ��� ���̸� �����Ͽ����ϴ�. �˺��� ���� �ְ��� ���ڼ��ݰ�꼭 ���� �Դϴ�."
        
        Messages.Add message
    Next
    
    For i = 0 To 50
        
        Set message = New PBMessage
        
        '�߽Ź�ȣ
        message.sender = "07043042992"
        
        '�߽��ڸ�
        message.senderName = "�߽��ڸ�"
        
        '���Ź�ȣ
        message.receiver = "010111222"
        
        '�����ڸ�
        message.receiverName = "�������̸�_" + CStr(i + 1)
        
        '�޽��� ����
        message.content = "�߽� ����. �� ������ �ܹ����� ���۵˴ϴ�."
        
        Messages.Add message
    Next
    
    
    '������ ���ۿ���
    adsYN = False
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    '�˺� ȸ�����̵�
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendXMS(txtCorpNum.Text, "", "", "", Messages, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

Private Sub btnSendXMS_One_Click()
    Dim Messages As New Collection
    Dim adsYN As Boolean
    Dim message As New PBMessage
    Dim receiptNum As String
    Dim requestNum As String
    Dim UserID As String
    
    '�߽Ź�ȣ
    message.sender = "07043042991"
    
    '�߽��ڸ�
    message.senderName = "�߽��ڸ�"
    
    '���Ź�ȣ
    message.receiver = "010111222"
    
    '�����ڸ�
    message.receiverName = "�������̸�"
    
    '�幮 �޽��� ����
    message.subject = "�幮�� ��� �幮 �����Դϴ�."
    
    '�޽��� ����, 90byte �������� ����Ÿ��(��/�幮)�� �ڵ����� �νĵǾ� ���۵�.
    message.content = "�ڵ��ν� �߼��� ������ ���̸� 90Byte�������� ���ϴ� �ܹ�, �̻��� �幮���� �ڵ� �����մϴ�."
    
    Messages.Add message
    
    '������ ���ۿ���
    adsYN = False
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    '�˺� ȸ�����̵�
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendXMS(txtCorpNum.Text, "", "", "", Messages, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    txtReceiptNum.Text = receiptNum
End Sub

Private Sub btnSendXMS_Same_Click()
    Dim Messages As New Collection
    Dim adsYN As Boolean
    Dim sender As String
    Dim subject As String
    Dim Contents As String
    Dim message As PBMessage
    Dim i As Integer
    Dim receiptNum As String
    Dim requestNum As String
    Dim UserID As String
    
    '�߽Ź�ȣ
    sender = "07043042991"
    
    '�����޽��� ����
    subject = "�������� ����, �幮�� �����"
    
    '�����޽��� ����
    Contents = "�ڵ��ν� �߼��� ������ ���̸� 90Byte�������� ���ϴ� �ܹ�, �̻��� �幮���� �ڵ� �����մϴ�."
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        '���Ź�ȣ
        message.receiver = "010111222"
        
        '�����ڸ�
        message.receiverName = "�������̸�_" + CStr(i + 1)
        
        Messages.Add message
    Next
    
    '������ ���ۿ���
    adsYN = False
    
    '���ۿ�û��ȣ, ��Ʈ�ʰ� ���ۿ�û�� ���� ������ȣ�� ���� �Ҵ��Ͽ� �����ϴ� ��� ����
    '�ִ� 36�ڸ�, ����, ����, �����('_'), ������('-')�� �����Ͽ� ����ں��� �ߺ����� �ʵ��� ����
    requestNum = ""
    
    '�˺� ȸ�����̵�
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendLMS(txtCorpNum.Text, sender, subject, Contents, _
                                        Messages, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "���� ��ȣ : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' �ܹ�(SMS) ���۴ܰ��� Ȯ���մϴ�.
'=========================================================================

Private Sub btnUnitCost_Click()
    Dim unitCost As Single
    
    unitCost = MessageService.GetUnitCost(txtCorpNum.Text, SMS)
    
    If unitCost < 0 Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "SMS ���� �ܰ� : " + CStr(unitCost)
    
End Sub

'=========================================================================
' �幮(LMS) ���۴ܰ��� Ȯ���մϴ�.
'=========================================================================

Private Sub btnUnitCost_LMS_Click()
    Dim unitCost As Single
    
    unitCost = MessageService.GetUnitCost(txtCorpNum.Text, LMS)
    
    If unitCost < 0 Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "LMS ���� �ܰ� : " + CStr(unitCost)
End Sub

'=========================================================================
' ����(MMS)�޽��� ���۴ܰ��� Ȯ���մϴ�.
'=========================================================================

Private Sub btnUnitCost_MMS_Click()
    Dim unitCost As Single
    
    unitCost = MessageService.GetUnitCost(txtCorpNum.Text, MMS)
    
    If unitCost < 0 Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "MMS ���� �ܰ� : " + CStr(unitCost)
End Sub

'=========================================================================
' ����ȸ���� ����� ������ �����մϴ�.
'=========================================================================

Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�
    joinData.id = txtUserID.Text
    
    '����ڸ�
    joinData.personName = "����ڸ�_����"
    
    '����ó
    joinData.tel = "070-4304-2991"
    
    '�޴�����ȣ
    joinData.hp = "010-1234-1234"
    
    '�̸��� �ּ�
    joinData.email = "test@test.com"
    
    '�ѽ���ȣ
    joinData.fax = "070-1234-1234"
    
    '��ü��ȸ����, Ture-ȸ����ȸ, False-������ȸ
    joinData.searchAllAllowYN = True
    
    '������ ���ѿ���
    joinData.mgrYN = False
    
    Set Response = MessageService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� �����մϴ�
'=========================================================================

Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '��ǥ�ڸ�
    CorpInfo.ceoname = "��ǥ��"
    
    '��ȣ
    CorpInfo.corpName = "��ȣ"
    
    '�ּ�
    CorpInfo.addr = "����Ư����"
    
    '����
    CorpInfo.bizType = "����"
    
    '����
    CorpInfo.bizClass = "����"
    
    Set Response = MessageService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(MessageService.LastErrCode) + vbCrLf + "����޽��� : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

Private Sub Form_Load()
    MessageService.Initialize LinkID, SecretKey
    
    '����ȯ�� ������ True-���߿�, False-�����
    MessageService.IsTest = True
    
End Sub

