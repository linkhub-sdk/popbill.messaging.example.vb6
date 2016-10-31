VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "팝빌 메시징 SDK 예제"
   ClientHeight    =   12540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13620
   LinkTopic       =   "Form1"
   ScaleHeight     =   12540
   ScaleWidth      =   13620
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton btnSendMMS_hundred 
      Caption         =   "100건 전송"
      Height          =   465
      Left            =   6360
      TabIndex        =   41
      Top             =   5830
      Width           =   1050
   End
   Begin VB.Frame Frame10 
      Caption         =   "포토 전송기능"
      Height          =   945
      Left            =   5040
      TabIndex        =   38
      Top             =   5520
      Width           =   3705
      Begin VB.CommandButton btnSendMMS 
         Caption         =   "1건 전송"
         Height          =   465
         Left            =   240
         TabIndex        =   40
         Top             =   315
         Width           =   930
      End
      Begin VB.CommandButton btnSendMMS_Same 
         Caption         =   "대량전송"
         Height          =   465
         Left            =   2520
         TabIndex        =   39
         Top             =   315
         Width           =   975
      End
   End
   Begin VB.CommandButton btnUnitCost_MMS 
      Caption         =   "MMS 전송단가 확인"
      Height          =   410
      Left            =   2320
      TabIndex        =   37
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Frame Frame6 
      Caption         =   " 팝빌 메시징 관련 기능"
      Height          =   8655
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Width           =   13005
      Begin VB.CommandButton btnGetAutoDenyList 
         Caption         =   "080 수신거부목록 확인"
         Height          =   495
         Left            =   9120
         TabIndex        =   52
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "전송내역 목록조회"
         Height          =   495
         Left            =   9120
         TabIndex        =   51
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton btnSearchPopup 
         Caption         =   "전송내역조회 팝업 URL"
         Height          =   495
         Left            =   9120
         TabIndex        =   36
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtResult 
         Height          =   4680
         Left            =   720
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   35
         Top             =   3720
         Width           =   11775
      End
      Begin VB.CommandButton btnCancelReserve 
         Caption         =   "예약 전송 취소"
         Height          =   525
         Left            =   6360
         TabIndex        =   34
         Top             =   3000
         Width           =   1665
      End
      Begin VB.CommandButton btnGetMessages 
         Caption         =   "전송상태확인"
         Height          =   525
         Left            =   4560
         TabIndex        =   33
         Top             =   3000
         Width           =   1665
      End
      Begin VB.Frame Frame9 
         Caption         =   " 단/장문 자동인식 문자 전송 "
         Height          =   945
         Left            =   720
         TabIndex        =   29
         Top             =   1920
         Width           =   3825
         Begin VB.CommandButton btnSendXMS_One 
            Caption         =   "1건 전송"
            Height          =   465
            Left            =   360
            TabIndex        =   32
            Top             =   315
            Width           =   930
         End
         Begin VB.CommandButton btnSendXMS_Hundred 
            Caption         =   "100건 전송"
            Height          =   465
            Left            =   1440
            TabIndex        =   31
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendXMS_Same 
            Caption         =   "대량전송"
            Height          =   465
            Left            =   2640
            TabIndex        =   30
            Top             =   315
            Width           =   1020
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   " 장문 문자 전송 "
         Height          =   945
         Left            =   4920
         TabIndex        =   25
         Top             =   840
         Width           =   3705
         Begin VB.CommandButton btnSendLMS_One 
            Caption         =   "1건 전송"
            Height          =   465
            Left            =   240
            TabIndex        =   28
            Top             =   315
            Width           =   930
         End
         Begin VB.CommandButton btnSendLMS_Hundred 
            Caption         =   "100건 전송"
            Height          =   465
            Left            =   1320
            TabIndex        =   27
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendLMS_Same 
            Caption         =   "대량전송"
            Height          =   465
            Left            =   2520
            TabIndex        =   26
            Top             =   315
            Width           =   1020
         End
      End
      Begin VB.TextBox txtReceiptNum 
         Height          =   315
         Left            =   1485
         TabIndex        =   24
         Top             =   3105
         Width           =   2850
      End
      Begin VB.Frame Frame7 
         Caption         =   " 단문 문자 전송 "
         Height          =   945
         Left            =   720
         TabIndex        =   18
         Top             =   840
         Width           =   3825
         Begin VB.CommandButton btnSendSMS_Same 
            Caption         =   "대량전송"
            Height          =   465
            Left            =   2640
            TabIndex        =   21
            Top             =   315
            Width           =   1020
         End
         Begin VB.CommandButton btnSendSMS_hundredd 
            Caption         =   "100건 전송"
            Height          =   465
            Left            =   1440
            TabIndex        =   20
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendSMS_One 
            Caption         =   "1건 전송"
            Height          =   465
            Left            =   360
            TabIndex        =   19
            Top             =   315
            Width           =   930
         End
      End
      Begin VB.TextBox txtReserveDT 
         Height          =   315
         Left            =   3540
         TabIndex        =   16
         Top             =   375
         Width           =   3105
      End
      Begin VB.Frame Frame13 
         Caption         =   "부가기능"
         Height          =   2295
         Left            =   8880
         TabIndex        =   53
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "접수번호 : "
         Height          =   180
         Left            =   585
         TabIndex        =   23
         Top             =   3195
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "예약시간(yyyyMMddHHmmss) : "
         Height          =   180
         Left            =   705
         TabIndex        =   17
         Top             =   450
         Width           =   2790
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 팝빌 기본 API "
      Height          =   2775
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   12960
      Begin VB.Frame Frame12 
         Caption         =   "회사정보 관련"
         Height          =   2415
         Left            =   10920
         TabIndex        =   48
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "회사정보 수정"
            Height          =   410
            Left            =   120
            TabIndex        =   50
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "회사정보 조회"
            Height          =   410
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "담당자 관련"
         Height          =   2415
         Left            =   8880
         TabIndex        =   44
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "담당자 정보 수정"
            Height          =   410
            Left            =   120
            TabIndex        =   47
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "담당자 목록 조회"
            Height          =   410
            Left            =   120
            TabIndex        =   46
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "담당자 추가"
            Height          =   410
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 회원정보"
         Height          =   2415
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1695
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID 중복 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   42
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "가입 여부 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "회원 가입"
            Height          =   410
            Left            =   120
            TabIndex        =   12
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " 포인트 관련"
         Height          =   2415
         Left            =   2040
         TabIndex        =   9
         Top             =   240
         Width           =   2160
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "과금정보 확인"
            Height          =   410
            Left            =   160
            TabIndex        =   54
            Top             =   1800
            Width           =   1815
         End
         Begin VB.CommandButton btnUnitCost_LMS 
            Caption         =   "LMS 전송단가 확인"
            Height          =   410
            Left            =   160
            TabIndex        =   14
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "SMS 전송단가 확인"
            Height          =   410
            Left            =   150
            TabIndex        =   10
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " 파트너 관련"
         Height          =   2415
         Left            =   4320
         TabIndex        =   7
         Top             =   240
         Width           =   2535
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여 포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "파트너 잔여포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   2295
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " 팝빌 기본 URL"
         ClipControls    =   0   'False
         Height          =   2415
         Left            =   6960
         TabIndex        =   5
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnGetPopbillURL_CHRG 
            Caption         =   "포인트 충전 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   43
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton btnGetPopbillURL_LOGIN 
            Caption         =   " 팝빌 로그인 URL"
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
      Caption         =   "팝빌회원 아이디 : "
      Height          =   180
      Left            =   4680
      TabIndex        =   2
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "팝빌회원 사업자번호 : "
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
' 팝빌 문자 API VB 6.0 SDK Example
'
' - VB6 SDK 연동환경 설정방법 안내 :
' - 업데이트 일자 : 2016-10-11
' - 연동 기술지원 연락처 : 1600-8536 / 070-4304-2991 (직통 / 정요한대리)
' - 연동 기술지원 이메일 : dev@linkhub.co.kr
'
' <테스트 연동개발 준비사항>
' 1) 25, 28번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
'    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
' 2) 팝빌 개발용 사이트(test.popbill.com)에 연동회원으로 가입합니다.
'=========================================================================

Option Explicit

'=========================================================================
' - 인증정보(링크아이디, 비밀키)는 파트너의 연동회원을 식별하는
'   인증에 사용되는 정보로 유출되지 않도록 주의하시기 바랍니다.
' - 상업용 전환이후에도 인증정보(링크아이디, 비밀키)는 변경되지 않습니다.
'=========================================================================

'링크아이디
Private Const LinkID = "TESTER"

'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

Private MessageService As New PBMSGService

'=========================================================================
' 예약문자전송을 취소합니다.
' - 예약취소는 예약전송시간 10분전까지만 가능합니다.
'=========================================================================
Private Sub btnCancelReserve_Click()
    Dim Response As PBResponse
    
    Set Response = MessageService.CancelReserve(txtCorpNum.Text, txtReceiptNum.Text, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 팝빌 회원아이디 중복여부를 확인합니다.
'=========================================================================

Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = MessageService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 해당 사업자의 파트너 연동회원 가입여부를 확인합니다.
' - LinkID는 인증정보로 설정되어 있는 링크아이디 값입니다.
'=========================================================================

Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = MessageService.CheckIsMember(txtCorpNum.Text, LinkID)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 080 서비스 수신거부 목록을 확인합니다.
'=========================================================================

Private Sub btnGetAutoDenyList_Click()
    Dim AutoDenyList As Collection
    Dim tmp As String
    Dim AutoDenyInfo As PBAutoDenyInfo
    
    Set AutoDenyList = MessageService.GetAutoDenyList(txtCorpNum.Text)
    
    If AutoDenyList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "number(수신거부번호) | regDT(등록일시)" + vbCrLf
    
    For Each AutoDenyInfo In AutoDenyList
        tmp = tmp + AutoDenyInfo.number + " | " + AutoDenyInfo.regDT + vbCrLf
    Next
    
    MsgBox tmp
    
End Sub

'=========================================================================
' 연동회원의 잔여포인트를 확인합니다.
' - 과금방식이 파트너과금인 경우 파트너 잔여포인트(GetPartnerBalance API)
'   를 통해 확인하시기 바랍니다.
'=========================================================================

Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = MessageService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
    
End Sub

'=========================================================================
' 연동회원의 문자 API 서비스 과금정보를 확인합니다.
'=========================================================================

Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim MType As MsgType
    Dim tmp As String
    
    MType = SMS     'SMS-단문, LMS-장문 MMS-포토
            
    Set ChargeInfo = MessageService.GetChargeInfo(txtCorpNum.Text, MType)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (전송단가) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (과금유형) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (과금제도) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원의 회사정보를 확인합니다.
'=========================================================================

Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = MessageService.GetCorpInfo(txtCorpNum.Text, txtUserID.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname(대표자성명) : " + CorpInfo.CEOName + vbCrLf
    tmp = tmp + "corpName(상호명) : " + CorpInfo.CorpName + vbCrLf
    tmp = tmp + "addr(주소) : " + CorpInfo.Addr + vbCrLf
    tmp = tmp + "bizType(업태) : " + CorpInfo.BizType + vbCrLf
    tmp = tmp + "bizClass(종목) : " + CorpInfo.BizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 문자전송요청에 대한 전송결과를 확인합니다.
'=========================================================================

Private Sub btnGetMessages_Click()
    Dim sentMessages As Collection
    Dim sentMessage As PBSentMsg
    Dim tmp As String
    
    Set sentMessages = MessageService.GetMessages(txtCorpNum.Text, txtReceiptNum.Text, txtUserID.Text)
    
    If sentMessages Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "state | subject | messageType | sendnum | senderName | receiveNum | receiveName | receiptDT | reserveDT | sendDT | sendResult | tranNet" + vbCrLf
    
    For Each sentMessage In sentMessages
    
        tmp = tmp + CStr(sentMessage.state) + " | "
        tmp = tmp + sentMessage.subject + " | "
        tmp = tmp + sentMessage.messageType + " | "
        'tmp = tmp + sentMessage.content + " | " ' 내용 표시는 길이관계상 예제에서 생략합니다.
        tmp = tmp + sentMessage.sendNum + " | "
        tmp = tmp + sentMessage.senderName + " | "
        tmp = tmp + sentMessage.receiveNum + " | "
        tmp = tmp + sentMessage.receiveName + " | "
        tmp = tmp + sentMessage.receiptDT + " | "
        tmp = tmp + sentMessage.reserveDT + " | "
        tmp = tmp + sentMessage.sendDT + " | "
        tmp = tmp + sentMessage.resultDT + " | "
        tmp = tmp + sentMessage.sendResult + " | "
        tmp = tmp + sentMessage.tranNet
        
        tmp = tmp + vbCrLf
    Next
    
    txtResult.Text = tmp
    
End Sub

'=========================================================================
' 파트너의 잔여포인트를 확인합니다.
' - 과금방식이 연동과금인 경우 연동회원 잔여포인트(GetBalance API)를
'   이용하시기 바랍니다.
'=========================================================================

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = MessageService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
    
End Sub

'=========================================================================
' 연동회원 포인트 충전 URL을 반환합니다.
' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================

Private Sub btnGetPopbillURL_CHRG_Click()
    Dim url As String
    
    url = MessageService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 팝빌(www.popbill.com)에 로그인된 팝빌 URL을 반환합니다.
' - 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================

Private Sub btnGetPopbillURL_LOGIN_Click()
    Dim url As String
    
    url = MessageService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "LOGIN")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 팝빌 연동회원 가입을 요청합니다.
'=========================================================================

Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
 
    '링크 아이디
    joinData.LinkID = LinkID
    
    '사업자번호, '-'제외, 10자리
    joinData.CorpNum = "1231212312"
    
    '대표자성명, 최대 30자
    joinData.CEOName = "대표자성명"
    
    '상호명, 최대 70자
    joinData.CorpName = "회원상호"
    
    '주소, 최대 300자
    joinData.Addr = "주소"
    
    '업태, 최대 40자
    joinData.BizType = "업태"
    
    '종목, 최대 40자
    joinData.BizClass = "종목"
    
    '아이디, 6자이상 20자 미만
    joinData.ID = "userid"
    
    '비밀번호, 6자이상 20자 미만
    joinData.PWD = "pwd_must_be_long_enough"
    
    '담당자명, 최대 30자
    joinData.ContactName = "담당자성명"
    
    '담당자 연락처, 최대 20자
    joinData.ContactTEL = "02-999-9999"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.ContactHP = "010-1234-5678"
    
    '담당자 팩스번호, 최대 20자
    joinData.ContactFAX = "02-999-9998"
    
    '담당자 메일, 최대 70자
    joinData.ContactEmail = "test@test.com"
    
    Set Response = MessageService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
    
End Sub

'=========================================================================
' 연동회원의 담당자 목록을 확인합니다.
'=========================================================================

Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = MessageService.ListContact(txtCorpNum.Text, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id | email | hp | personName | searchAllAllowYN | tel | fax | mgrYN | regDT " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.ID + " | " + info.email + " | " + info.hp + " | " + info.personName + " | " + CStr(info.searchAllAllowYN) _
                + info.tel + " | " + info.fax + " | " + CStr(info.mgrYN) + " | " + info.regDT + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원의 담당자를 신규로 등록합니다.
'=========================================================================

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디, 6자 이상 20자 미만
    joinData.ID = "testkorea_20161011"
    
    '비밀번호, 6자 이상 20자 미만
    joinData.PWD = "test@test.com"
    
    '담당자명, 최대 30자
    joinData.personName = "담당자명"
    
    '담당자 연락처
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호
    joinData.hp = "010-1234-1234"
    
    '담당자 메일주소
    joinData.email = "test@test.com"
    
    '담당자 팩스번호
    joinData.fax = "070-1234-1234"
    
    '회사조회 권한여부, true-회사조회 / false-개인조회
    joinData.searchAllAllowYN = True
    
    '관리자 권한여부
    joinData.mgrYN = False
        
    Set Response = MessageService.RegistContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 검색조건을 사용하여 문자전송 내역을 조회합니다.
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
    
    '[필수] 시작일자, 날자형식(yyyyMMdd)
    SDate = "20160901"
    
    '[필수] 종료일자, 날자형식(yyyyMMdd)
    EDate = "20161031"
    
    '전송상태값 배열, 1-대기, 2-성공, 3-실패, 4-취소
    state.Add "1"
    state.Add "2"
    state.Add "3"
    state.Add "4"
    
    '검색대상 배열, SMS(단문),LMS(장문),MMS(포토)
    Item.Add "SMS"
    Item.Add "LMS"
    Item.Add "MMS"
    
    '예약문자 검색여부, True(예약문자만 조회), False(전체조회)
    ReserveYN = False
    
    '개인조회여부, True(개인조회), False(전체조회)
    SenderYN = False
    
    '페이지 번호
    Page = 1
    
    '페이지 목록개수, 최대 1000건
    PerPage = 50
    
    '정렬방향, D-내림차순(기본값), A-오름차순
    Order = "D"

    Set msgSearchList = MessageService.Search(txtCorpNum.Text, SDate, EDate, state, Item, ReserveYN, SenderYN, Page, PerPage, Order)
     
    If msgSearchList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "code : " + CStr(msgSearchList.code) + vbCrLf
    tmp = tmp + "total : " + CStr(msgSearchList.total) + vbCrLf
    tmp = tmp + "perPage : " + CStr(msgSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum : " + CStr(msgSearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount : " + CStr(msgSearchList.pageCount) + vbCrLf
    tmp = tmp + "message : " + msgSearchList.message + vbCrLf + vbCrLf
    
    tmp = tmp + "state | subject | messageType | sendnum | senderName | receiveNum | receiveName | receiptDT | reserveDT | sendDT | sendResult | tranNet" + vbCrLf
            
    For Each info In msgSearchList.list
        tmp = tmp + CStr(info.state) + " | "
        tmp = tmp + info.subject + " | "
        tmp = tmp + info.messageType + " | "
        'tmp = tmp + sentMessage.content + " | " ' 내용 표시는 길이관계상 예제에서 생략합니다.
        tmp = tmp + info.sendNum + " | "
        tmp = tmp + info.senderName + " | "
        tmp = tmp + info.receiveNum + " | "
        tmp = tmp + info.receiveName + " | "
        tmp = tmp + info.receiptDT + " | "
        tmp = tmp + info.reserveDT + " | "
        tmp = tmp + info.sendDT + " | "
        tmp = tmp + info.resultDT + " | "
        tmp = tmp + info.sendResult + " | "
        tmp = tmp + info.tranNet
        tmp = tmp + vbCrLf
    Next
        
    txtResult.Text = tmp
End Sub

'=========================================================================
' 문자메시지 전송내역 팝업 URL을 반환합니다.
' - 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
'=========================================================================

Private Sub btnSearchPopup_Click()
    Dim url As String
    
    url = MessageService.GetURL(txtCorpNum.Text, txtUserID.Text, "BOX")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnSendLMS_Hundred_Click()
    Dim Messages As New Collection
    Dim adsYN As Boolean
    Dim message As PBMessage
    Dim i As Integer
    Dim ReceiptNum As String
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        '발신번호
        message.sender = "07043042991"
        
        '발신자명
        message.senderName = "발신자명"
        
        '수신번호
        message.receiver = "010111222"
        
        '수신자명
        message.receiverName = "수신자이름_" + CStr(i + 1)
        
        '메시지 내용, 2000byte 초과된 내용은 삭제되어 전송됨.
        message.content = "발신 내용. 장문은 2000Byte로 길이가 조정되어 전송됩니다. 팝빌은 최고의 전자세금계산서 서비스를 제공하고 있습니다."
        
        '메시지
        message.subject = "장문 제목입니다."
        
        Messages.Add message
    Next
        
    adsYN = False       '광고문자 전송여부
    
    ReceiptNum = MessageService.SendLMS(txtCorpNum.Text, "", "", "", Messages, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    
    txtReceiptNum.Text = ReceiptNum
    
End Sub

Private Sub btnSendLMS_One_Click()
    Dim Messages As New Collection
    Dim adsYN As Boolean
    Dim message As New PBMessage
    Dim ReceiptNum As String
    
    '발신번호
    message.sender = "07043042991"
    
    '발신자명
    message.senderName = "발신자명"
    
    '수신번호
    message.receiver = "010111222"
    
    '수신자명
    message.receiverName = "수신자이름"
    
    
    '장문메시지 제목
    message.subject = "장문 제목입니다."
    
    '메시지내용, 2000byte 초과한 내용은 삭제되어 전송됨
    message.content = "발신 내용. 장문은 2000Byte로 길이가 조정되어 전송됩니다. 팝빌은 최고의 전자세금계산서 서비스를 제공하고 있습니다."
    
    Messages.Add message
    
    
    '광고문자 전송여부
    adsYN = False
    
    ReceiptNum = MessageService.SendLMS(txtCorpNum.Text, "", "", "", Messages, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    
    txtReceiptNum.Text = ReceiptNum
    
End Sub

Private Sub btnSendLMS_Same_Click()
    Dim Messages As New Collection
    Dim sender As String
    Dim subject As String
    Dim Contents As String
    Dim adsYN As Boolean
    Dim message As PBMessage
    Dim i As Integer
    Dim ReceiptNum As String
    
    '발신번호
    sender = "07043042991"
    
    '동보전송 제목
    subject = "동보전송 제목"
    
    '동보전송 메시지
    Contents = "발신 내용. 장문은 2000Byte로 길이가 조정되어 전송됩니다. 팝빌은 최고의 전자세금계산서 서비스를 제공하고 있습니다."
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        '수신번호
        message.receiver = "010111222"
        
        '수신자명
        message.receiverName = "수신자이름_" + CStr(i + 1)
        
        Messages.Add message
    Next
    
    '광고문자 전송여부
    adsYN = False
    
    ReceiptNum = MessageService.SendLMS(txtCorpNum.Text, sender, subject, Contents, _
                                    Messages, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendMMS_Click()
    Dim Messages As New Collection
    Dim FilePaths As New Collection
    Dim adsYN As Boolean
    Dim ReceiptNum As String
    Dim message As New PBMessage
    
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    '포토 메시지 파일경로
    FilePaths.Add CommonDialog1.FileName
    
    
    '발신번호
    message.sender = "07043042991"
    
    '발신자명
    message.senderName = "발신자명"
    
    '수신번호
    message.receiver = "010111222"
    
    '수신자명
    message.receiverName = "수신자이름"
    
    '포토 메시지 제목
    message.subject = "메시지 제목"
    
    '포토 메시지 내용
    message.content = "MMS 발신 테스트 내용."
    
    Messages.Add message
    
    '광고문자 전송여부
    adsYN = False
    
    ReceiptNum = MessageService.SendMMS(txtCorpNum.Text, "", "", "", Messages, FilePaths, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    
    txtReceiptNum.Text = ReceiptNum
    
End Sub

Private Sub btnSendMMS_hundred_Click()
    Dim Messages As New Collection
    Dim FilePaths As New Collection
    Dim adsYN As Boolean
    Dim message As PBMessage
    Dim i As Integer
    Dim ReceiptNum As String
    
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    FilePaths.Add CommonDialog1.FileName
  
    For i = 0 To 50
        
        Set message = New PBMessage
        
        '발신번호
        message.sender = "07043042991"
        
        '발신자명
        message.senderName = "발신자명"
        
        '수신번호
        message.receiver = "010111222"
        
        '수신자명
        message.receiverName = "수신자이름_" + CStr(i + 1)
        
        '메시지 제목
        message.subject = "포토메시지 제목입니다."
        
        '메시지 내용
        message.content = "발신 내용. 이 내용은 장문으로 전송될수 있도록 길이를 설정하였습니다. 팝빌은 국내 최고의 전자세금계산서 서비스 입니다."
        
        Messages.Add message
    Next
    
    For i = 0 To 50
        
        Set message = New PBMessage
        
        '발신번호
        message.sender = "07043042991"
        
        '발신자명
        message.senderName = "발신자명"
        
        '수신번호
        message.receiver = "010111222"
        
        '수신자명
        message.receiverName = "수신자이름_" + CStr(i + 1)
        
        '메시지 제목
        message.subject = "포토 메시지 제목"
        
        '메시지 내용
        message.content = "발신 내용. 이 내용은 단문으로 전송됩니다."
        
        Messages.Add message
    Next
    
    
    '광고문자 전송여부
    adsYN = False
    
    ReceiptNum = MessageService.SendMMS(txtCorpNum.Text, "", "", "", Messages, FilePaths, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendMMS_Same_Click()
    Dim Messages As New Collection
    Dim FilePaths As New Collection
    Dim adsYN As Boolean
    Dim ReceiptNum As String
    Dim message As PBMessage
    Dim sender As String
    Dim subject As String
    Dim Contents As String
    Dim i As Integer
    
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    FilePaths.Add CommonDialog1.FileName
    
    '발신번호
    sender = "07043042991"
    
    '동보메시지 제목
    subject = "동보메시지 제목"
    
    '동보메시지 내용
    Contents = "동보메시지 내용"
    
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        '수신번호
        message.receiver = "010111222"
        
        '수신자명
        message.receiverName = "수신자이름_" + CStr(i + 1)
        
        Messages.Add message
    Next
   
    
    '광고문자 전송여부
    adsYN = False
    
    ReceiptNum = MessageService.SendMMS(txtCorpNum.Text, sender, subject, Contents, Messages, FilePaths, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    
    txtReceiptNum.Text = ReceiptNum
    
End Sub

Private Sub btnSendSMS_hundredd_Click()
    Dim Messages As New Collection
    Dim adsYN As Boolean
    Dim message As PBMessage
    Dim i As Integer
    Dim ReceiptNum As String
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        '발신번호
        message.sender = "07043042991"
        
        '발신자명
        message.senderName = "발신자명"
        
        '수신번호
        message.receiver = "010111222"
        
        '수신자명
        message.receiverName = "수신자이름_" + CStr(i + 1)
        
        '메시지 내용, 90Byte 초과된 내용은 삭제되어 전송됨
        message.content = "발신 내용. 단문은 90Byte로 길이가 조정되어 전송됩니다."
        
        Messages.Add message
    Next
    
    adsYN = False       '광고문자 전송여부
    
    ReceiptNum = MessageService.SendSMS(txtCorpNum.Text, "", "", Messages, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    
    txtReceiptNum.Text = ReceiptNum
    
End Sub

Private Sub btnSendSMS_One_Click()
    Dim Messages As New Collection
    Dim adsYN As Boolean
    Dim message As New PBMessage
    Dim ReceiptNum As String
    
    '발신번호
    message.sender = "07043042991"
    
    '발신자명
    message.senderName = "발신자명"
    
    '수신번호
    message.receiver = "010111222"
    
    '수신자명
    message.receiverName = "수신자이름"
    
    '메시지 내용
    message.content = "발신 내용. 단문은 90Byte로 길이가 조정되어 전송됩니다."
    
    Messages.Add message
    
    '광고문자 전송여부
    adsYN = False
    
    ReceiptNum = MessageService.SendSMS(txtCorpNum.Text, "", "", Messages, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
    
End Sub

Private Sub btnSendSMS_Same_Click()
    Dim Messages As New Collection
    Dim adsYN As Boolean
    Dim sender As String
    Dim Contents As String
    Dim message As PBMessage
    Dim i As Integer
    Dim ReceiptNum As String
    
    '발신번호
    sender = "07075103710"
        
    '메시지 내용, 90byte 초과된 내용은 삭제되어 전송됨.
    Contents = "동보전송 내용 90byte로 길이가 조정되며, Messages의 내용이 없는 수신건에 동보처리됩니다."
    
    '광고문자 전송여부
    adsYN = False
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        '수신번호
        message.receiver = "010111222"
        
        '수신자명
        message.receiverName = "수신자이름_" + CStr(i + 1)
        
        Messages.Add message
    Next
    
    ReceiptNum = MessageService.SendSMS(txtCorpNum.Text, sender, Contents, Messages, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    
    txtReceiptNum.Text = ReceiptNum
    
End Sub

Private Sub btnSendXMS_Hundred_Click()
    Dim Messages As New Collection
    Dim adsYN As Boolean
    Dim ReceiptNum As String
    Dim message As PBMessage
    Dim i As Integer
    
    For i = 0 To 50
        
        Set message = New PBMessage
        
        '발신번호
        message.sender = "07043042991"
        
        '발신자명
        message.senderName = "발신자명"
        
        '수신번호
        message.receiver = "010111222"
        
        '수신자명
        message.receiverName = "수신자이름_" + CStr(i + 1)
        
        '장문메시지 제목
        message.subject = "장문 제목입니다."
        
        '메시지 내용
        message.content = "발신 내용. 이 내용은 장문으로 전송될수 있도록 길이를 설정하였습니다. 팝빌은 국내 최고의 전자세금계산서 서비스 입니다."
        
        Messages.Add message
    Next
    
    For i = 0 To 50
        
        Set message = New PBMessage
        
        '발신번호
        message.sender = "07043042992"
        
        '발신자명
        message.senderName = "발신자명"
        
        '수신번호
        message.receiver = "010111222"
        
        '수신자명
        message.receiverName = "수신자이름_" + CStr(i + 1)
        
        '메시지 내용
        message.content = "발신 내용. 이 내용은 단문으로 전송됩니다."
        
        Messages.Add message
    Next
    
    
    '광고문자 전송여부
    adsYN = False
    
    ReceiptNum = MessageService.SendXMS(txtCorpNum.Text, "", "", "", Messages, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    
    txtReceiptNum.Text = ReceiptNum
    
End Sub

Private Sub btnSendXMS_One_Click()
    Dim Messages As New Collection
    Dim adsYN As Boolean
    Dim message As New PBMessage
    Dim ReceiptNum As String
    
    '발신번호
    message.sender = "07043042991"
    
    '발신자명
    message.senderName = "발신자명"
    
    '수신번호
    message.receiver = "010111222"
    
    '수신자명
    message.receiverName = "수신자이름"
    
    '장문 메시지 제목
    message.subject = "장문의 경우 장문 제목입니다."
    
    '메시지 내용, 90byte 기준으로 문자타입(단/장문)이 자동으로 인식되어 전송됨.
    message.content = "자동인식 발송은 내용의 길이를 90Byte기준으로 이하는 단문, 이상은 장문으로 자동 전송합니다."
    
    Messages.Add message
    
    '광고문자 전송여부
    adsYN = False
    
    ReceiptNum = MessageService.SendXMS(txtCorpNum.Text, "", "", "", Messages, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendXMS_Same_Click()
    Dim Messages As New Collection
    Dim adsYN As Boolean
    Dim sender As String
    Dim subject As String
    Dim Contents As String
    Dim message As PBMessage
    Dim i As Integer
    Dim ReceiptNum As String
    
    '발신번호
    sender = "07043042991"
    
    '동보메시지 제목
    subject = "동보전송 제목, 장문에 적용됨"
    
    '동보메시지 내용
    Contents = "자동인식 발송은 내용의 길이를 90Byte기준으로 이하는 단문, 이상은 장문으로 자동 전송합니다."
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        '수신번호
        message.receiver = "010111222"
        
        '수신자명
        message.receiverName = "수신자이름_" + CStr(i + 1)
        
        Messages.Add message
    Next
    
    '광고문자 전송여부
    adsYN = False
    
    ReceiptNum = MessageService.SendLMS(txtCorpNum.Text, sender, subject, Contents, _
                                        Messages, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    
    txtReceiptNum.Text = ReceiptNum
End Sub

'=========================================================================
' 단문(SMS) 전송단가를 확인합니다.
'=========================================================================

Private Sub btnUnitCost_Click()
    Dim unitCost As Single
    
    unitCost = MessageService.GetUnitCost(txtCorpNum.Text, SMS)
    
    If unitCost < 0 Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "SMS 전송 단가 : " + CStr(unitCost)
    
End Sub

'=========================================================================
' 장문(LMS) 전송단가를 확인합니다.
'=========================================================================

Private Sub btnUnitCost_LMS_Click()
    Dim unitCost As Single
    
    unitCost = MessageService.GetUnitCost(txtCorpNum.Text, LMS)
    
    If unitCost < 0 Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "LMS 전송 단가 : " + CStr(unitCost)
End Sub

'=========================================================================
' 포토(MMS)메시지 전송단가를 확인합니다.
'=========================================================================

Private Sub btnUnitCost_MMS_Click()
    Dim unitCost As Single
    
    unitCost = MessageService.GetUnitCost(txtCorpNum.Text, MMS)
    
    If unitCost < 0 Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "MMS 전송 단가 : " + CStr(unitCost)
End Sub

'=========================================================================
' 연동회원의 담당자 정보를 수정합니다.
'=========================================================================

Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse

    '담당자명
    joinData.personName = "담당자명_수정"
    
    '연락처
    joinData.tel = "070-4304-2991"
    
    '휴대폰번호
    joinData.hp = "010-1234-1234"
    
    '이메일 주소
    joinData.email = "test@test.com"
    
    '팩스번호
    joinData.fax = "070-1234-1234"
    
    '전체조회여부, Ture-회사조회, False-개인조회
    joinData.searchAllAllowYN = True
    
    '관리자 권한여부
    joinData.mgrYN = False
    
    Set Response = MessageService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 회사정보를 수정합니다
'=========================================================================

Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '대표자명
    CorpInfo.CEOName = "대표자"
    
    '상호
    CorpInfo.CorpName = "상호"
    
    '주소
    CorpInfo.Addr = "서울특별시"
    
    '업태
    CorpInfo.BizType = "업태"
    
    '종목
    CorpInfo.BizClass = "종목"
    
    Set Response = MessageService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

Private Sub Form_Load()
    MessageService.Initialize LinkID, SecretKey
    
    '연동환경 설정값 True-개발용, False-상업용
    MessageService.IsTest = True
    
End Sub

