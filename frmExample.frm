VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExample 
   Caption         =   "팝빌 메시징 SDK 예제"
   ClientHeight    =   12555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17010
   LinkTopic       =   "Form1"
   ScaleHeight     =   12555
   ScaleWidth      =   17010
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   12840
      TabIndex        =   71
      Top             =   165
      Width           =   4095
   End
   Begin VB.CommandButton btnUnitCost_MMS 
      Caption         =   "MMS 전송단가 확인"
      Height          =   410
      Left            =   2320
      TabIndex        =   32
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Frame Frame6 
      Caption         =   " 팝빌 메시징 관련 기능"
      Height          =   8655
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   13005
      Begin VB.Frame Frame10 
         Caption         =   "포토 전송기능"
         Height          =   945
         Left            =   4680
         TabIndex        =   62
         Top             =   1920
         Width           =   3825
         Begin VB.CommandButton btnSendMMS_Same 
            Caption         =   "동보전송"
            Height          =   465
            Left            =   2520
            TabIndex        =   65
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendMMS 
            Caption         =   "1건 전송"
            Height          =   465
            Left            =   120
            TabIndex        =   64
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendMMS_hundred 
            Caption         =   "100건 전송"
            Height          =   465
            Left            =   1320
            TabIndex        =   63
            Top             =   315
            Width           =   1110
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "요청번호 할당 전송건 처리"
         Height          =   1335
         Left            =   4920
         TabIndex        =   56
         Top             =   3000
         Width           =   4215
         Begin VB.CommandButton btnGetMessagesRN 
            Caption         =   "전송상태확인"
            Height          =   525
            Left            =   120
            TabIndex        =   60
            Top             =   720
            Width           =   1905
         End
         Begin VB.CommandButton btnCancelReserveRN 
            Caption         =   "예약 전송 취소"
            Height          =   525
            Left            =   2160
            TabIndex        =   58
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox txtRequestNum 
            Height          =   435
            Left            =   1200
            TabIndex        =   57
            Top             =   240
            Width           =   2850
         End
         Begin VB.Label 요청번호 
            Caption         =   "요청번호 : "
            Height          =   375
            Left            =   240
            TabIndex        =   59
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "발신번호 관리"
         Height          =   1455
         Left            =   10680
         TabIndex        =   45
         Top             =   240
         Width           =   2055
         Begin VB.CommandButton btnGetSenderNumberMgtURL 
            Caption         =   "발신번호 관리 팝업"
            Height          =   495
            Left            =   120
            TabIndex        =   47
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton btnGetSenderNuberList 
            Caption         =   "발신번호 목록 조회"
            Height          =   495
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.CommandButton btnGetAutoDenyList 
         Caption         =   "080 수신거부목록"
         Height          =   495
         Left            =   8760
         TabIndex        =   42
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "전송내역 목록조회"
         Height          =   495
         Left            =   8760
         TabIndex        =   41
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton btnGetSentListURL 
         Caption         =   "전송내역조회 팝업"
         Height          =   495
         Left            =   8760
         TabIndex        =   31
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtResult 
         Height          =   3840
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   30
         Top             =   4560
         Width           =   12255
      End
      Begin VB.CommandButton btnCancelReserve 
         Caption         =   "예약 전송 취소"
         Height          =   525
         Left            =   2640
         TabIndex        =   29
         Top             =   3720
         Width           =   1905
      End
      Begin VB.Frame Frame9 
         Caption         =   " 단/장문 자동인식 문자 전송 "
         Height          =   945
         Left            =   480
         TabIndex        =   25
         Top             =   1920
         Width           =   3945
         Begin VB.CommandButton btnSendXMS_One 
            Caption         =   "1건 전송"
            Height          =   465
            Left            =   240
            TabIndex        =   28
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendXMS_Hundred 
            Caption         =   "100건 전송"
            Height          =   465
            Left            =   1440
            TabIndex        =   27
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendXMS_Same 
            Caption         =   "동보전송"
            Height          =   465
            Left            =   2640
            TabIndex        =   26
            Top             =   315
            Width           =   1110
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   " 장문 문자 전송 "
         Height          =   945
         Left            =   4680
         TabIndex        =   21
         Top             =   840
         Width           =   3825
         Begin VB.CommandButton btnSendLMS_One 
            Caption         =   "1건 전송"
            Height          =   465
            Left            =   120
            TabIndex        =   24
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendLMS_Hundred 
            Caption         =   "100건 전송"
            Height          =   465
            Left            =   1320
            TabIndex        =   23
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendLMS_Same 
            Caption         =   "동보전송"
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
         Caption         =   " 단문 문자 전송 "
         Height          =   945
         Left            =   480
         TabIndex        =   16
         Top             =   840
         Width           =   3945
         Begin VB.CommandButton btnSendSMS_Same 
            Caption         =   "동보전송"
            Height          =   465
            Left            =   2640
            TabIndex        =   19
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendSMS_Hundred 
            Caption         =   "100건 전송"
            Height          =   465
            Left            =   1440
            TabIndex        =   18
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendSMS_One 
            Caption         =   "1건 전송"
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
         Caption         =   "부가기능"
         Height          =   2655
         Left            =   8640
         TabIndex        =   43
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton btnGetStates 
            Caption         =   "전송내역 요약정보"
            Height          =   495
            Left            =   120
            TabIndex        =   53
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "접수번호 관련 기능 (요청번호 미할당)"
         Height          =   1335
         Left            =   480
         TabIndex        =   54
         Top             =   3000
         Width           =   4335
         Begin VB.CommandButton btnGetMessages 
            Caption         =   "전송상태확인"
            Height          =   525
            Left            =   120
            TabIndex        =   61
            Top             =   720
            Width           =   1905
         End
         Begin VB.Label 접수번호 
            Caption         =   "접수번호 : "
            Height          =   255
            Left            =   240
            TabIndex        =   55
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "예약시간(yyyyMMddHHmmss) : "
         Height          =   180
         Left            =   705
         TabIndex        =   15
         Top             =   450
         Width           =   2790
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " 팝빌 기본 API "
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   16800
      Begin VB.Frame Frame16 
         Caption         =   "파트너과금 포인트"
         Height          =   2535
         Left            =   8280
         TabIndex        =   49
         Top             =   240
         Width           =   2415
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "파트너 포인트 충전 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   52
            Top             =   840
            Width           =   2175
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "파트너 잔여포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   51
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "연동과금 포인트"
         Height          =   2535
         Left            =   6000
         TabIndex        =   48
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   "포인트 충전 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   69
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton btnGetUseHistoryURL 
            Caption         =   "포인트 사용내역 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   67
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton btnGetPaymentURL 
            Caption         =   "포인트 결제내역 URL"
            Height          =   410
            Left            =   120
            TabIndex        =   66
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여 포인트 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "회사정보 관련"
         Height          =   2535
         Left            =   14880
         TabIndex        =   38
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "회사정보 수정"
            Height          =   410
            Left            =   120
            TabIndex        =   40
            Top             =   840
            Width           =   1575
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "회사정보 조회"
            Height          =   410
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "담당자 관련"
         Height          =   2535
         Left            =   12840
         TabIndex        =   34
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton btnGetContactInfo 
            Caption         =   "담당자 정보 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   68
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "담당자 정보 수정"
            Height          =   410
            Left            =   120
            TabIndex        =   37
            Top             =   1800
            Width           =   1695
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "담당자 목록 조회"
            Height          =   410
            Left            =   120
            TabIndex        =   36
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "담당자 추가"
            Height          =   410
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 회원정보"
         Height          =   2535
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1695
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID 중복 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   33
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "가입 여부 확인"
            Height          =   410
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "회원 가입"
            Height          =   410
            Left            =   120
            TabIndex        =   10
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "전송단가"
         Height          =   2535
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         Width           =   3840
         Begin VB.CommandButton btnGetChargeInfo 
            Caption         =   "과금정보 확인"
            Height          =   410
            Left            =   2160
            TabIndex        =   44
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnUnitCost_LMS 
            Caption         =   "LMS 전송단가 확인"
            Height          =   410
            Left            =   160
            TabIndex        =   12
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton btnGetUnitCost 
            Caption         =   "SMS 전송단가 확인"
            Height          =   410
            Left            =   150
            TabIndex        =   8
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " 팝빌 기본 URL"
         ClipControls    =   0   'False
         Height          =   2535
         Left            =   10920
         TabIndex        =   5
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnGetAccessURL 
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "URL : "
      Height          =   180
      Left            =   12120
      TabIndex        =   70
      Top             =   240
      Width           =   525
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
' - 업데이트 일자 : 2021-10-13
' - 연동 기술지원 연락처 : 1600-9854 / 070-4304-2991
' - 연동 기술지원 이메일 : code@linkhub.co.kr
'
' <테스트 연동개발 준비사항>
' 1) 29, 32번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
'    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
' 2) 팝빌 개발용 사이트(test.popbill.com)에 연동회원으로 가입합니다.
' 3) 문자를 전송하기 위해 발신번호 사전등록을 합니다. (등록방법은 사이트/API 두가지 방식이 있습니다.)
'     - 팝빌 사이트 로그인 > [문자/팩스] > [문자] > [발신번호 사전등록] 메뉴에서 등록
'     - getSenderNumberMgtURL API를 통해 반환된 URL을 이용하여 발신번호 등록

'=========================================================================

Option Explicit

'링크아이디
Private Const linkID = "TESTER"

'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'문자 서비스 클래스 선언
Private MessageService As New PBMSGService

'=========================================================================
' 사업자번호를 조회하여 연동회원 가입여부를 확인합니다.
' - LinkID는 인증정보로 설정되어 있는 링크아이디 값입니다.
' - https://docs.popbill.com/message/vb/api#CheckIsMember
'=========================================================================
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = MessageService.CheckIsMember(txtCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 사용하고자 하는 아이디의 중복여부를 확인합니다.
' - https://docs.popbill.com/message/vb/api#CheckID
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
' 사용자를 연동회원으로 가입처리합니다.
' - https://docs.popbill.com/message/vb/api#JoinMember
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '아이디, 6자이상 50자 미만
    joinData.id = "userid"
    
    '비밀번호, 8자 이상 20자 이하(영문, 숫자, 특수문자 조합)
    joinData.Password = "asdf$%^123"
    
    '파트너링크 아이디
    joinData.linkID = linkID
    
    '사업자번호, '-'제외, 10자리
    joinData.CorpNum = "1234567890"
    
    '대표자성명, 최대 100자
    joinData.ceoname = "대표자성명"
    
    '상호명, 최대 200자
    joinData.corpName = "회원상호"
    
    '사업장 주소, 최대 300자
    joinData.addr = "주소"
    
    '업태, 최대 100자
    joinData.bizType = "업태"
    
    '종목, 최대 100자
    joinData.bizClass = "종목"

    '담당자 성명, 최대 100자
    joinData.ContactName = "담당자성명"
    
    '담당자 이메일, 최대 100자
    joinData.ContactEmail = "test@test.com"
    
    '담당자 연락처, 최대 20자
    joinData.ContactTEL = "02-999-9999"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.ContactHP = "010-1234-5678"
    
    '담당자 팩스번호, 최대 20자
    joinData.ContactFAX = "02-999-9998"
    
    Set Response = MessageService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 팝빌 문자(SMS)API 서비스 과금정보를 확인합니다.
' - https://docs.popbill.com/message/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim MsgType As MsgType
    Dim tmp As String
    
    '문자전송 유형, SMS-단문, LMS-장문, MMS-포토
    MsgType = SMS
    
    Set ChargeInfo = MessageService.GetChargeInfo(txtCorpNum.Text, MsgType)
     
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
' 문자(SMS) 전송시 과금되는 포인트 단가를 확인합니다.
' - https://docs.popbill.com/message/vb/api#GetUnitCost
'=========================================================================
Private Sub btnGetUnitCost_Click()
    Dim unitCost As Single
    
    unitCost = MessageService.GetUnitCost(txtCorpNum.Text, SMS)
    
    If unitCost < 0 Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "SMS 전송 단가 : " + CStr(unitCost)
End Sub

'=========================================================================
' 문자(LMS) 전송시 과금되는 포인트 단가를 확인합니다.
' - https://docs.popbill.com/message/vb/api#GetUnitCost
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
' 문자(MMS) 전송시 과금되는 포인트 단가를 확인합니다.
' - https://docs.popbill.com/message/vb/api#GetUnitCost
'=========================================================================
Private Sub btnUnitCost_MMS_Click()
    Dim unitCost As Single
    
    unitCost = MessageService.GetUnitCost(txtCorpNum.Text, MMS)
    
    If unitCost < 0 Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "전송 단가 : " + CStr(unitCost)
End Sub

'=========================================================================
' 팝빌 사이트에 로그인 상태로 접근할 수 있는 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/message/vb/api#GetAccessURL
'==========================================================================
Private Sub btnGetAccessURL_Click()
    Dim url As String
    
    url = MessageService.GetAccessURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 연동회원 사업자번호에 담당자(팝빌 로그인 계정)를 추가합니다.
' - https://docs.popbill.com/message/vb/api#RegistContact
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디, 6자 이상 50자 미만
    joinData.id = "testkorea"
    
    '비밀번호, 8자 이상 20자 이하(영문, 숫자, 특수문자 조합)
    joinData.Password = "asdf#$%123"
    
    '담당자명, 최대 100자
    joinData.personName = "담당자명"
    
    '담당자 연락처, 최대 20자
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.hp = "010-1234-1234"
    
    '담당자 팩스번,최대 20자
    joinData.fax = "070-1234-1234"
    
    '담당자 메일주소, 최대 100자
    joinData.email = "test@test.com"
    
    '담당자 권한, 1-개인 / 2-읽기 / 3-회사
    joinData.searchRole = 3
        
    Set Response = MessageService.RegistContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 정보를 확인합니다.
' https://docs.popbill.com/message/vb/api#GetContactInfo
'=========================================================================
Private Sub btnGetContactInfo_Click()
    Dim tmp As String
    Dim info As PBContactInfo
    Dim ContactID As String
    
    ContactID = ""
    
    Set info = MessageService.GetContactInfo(txtCorpNum.Text, ContactID, txtUserID.Text)
    
    If info Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(아이디) | personName(성명) | email(이메일) | hp(휴대폰번호) |  fax(팩스번호) | tel(연락처) | " _
         + "regDT(등록일시) | searchRole(담당자 권한) | mgrYN(관리자 여부) | state(상태) " + vbCrLf
    
   
    tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
        
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 목록을 확인합니다.
' - https://docs.popbill.com/message/vb/api#ListContact
'=========================================================================
Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = MessageService.ListContact(txtCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(아이디) | personName(성명) | email(이메일) | hp(휴대폰번호) |  fax(팩스번호) | tel(연락처) | " _
         + "regDT(등록일시) | searchRole(담당자 권한) | mgrYN(관리자 여부) | state(상태) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 정보를 수정합니다.
' - https://docs.popbill.com/message/vb/api#UpdateContact
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디
    joinData.id = txtUserID.Text
    
    '담당자 성명, 최대 100자
    joinData.personName = "담당자명_수정"
    
    '담당자 연락처, 최대 20자
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.hp = "010-1234-1234"
        
    '담당자 팩스번호, 최대 20자
    joinData.fax = "070-1234-1234"
    
    '담당자 이메일, 최대 100자
    joinData.email = "test@test.com"

    '담당자 권한, 1-개인 / 2-읽기 / 3-회사
    joinData.searchRole = 3
                
    Set Response = MessageService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 회사정보를 확인합니다.
' - https://docs.popbill.com/message/vb/api#GetCorpInfo
'=========================================================================

Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = MessageService.GetCorpInfo(txtCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname(대표자성명) : " + CorpInfo.ceoname + vbCrLf
    tmp = tmp + "corpName(상호명) : " + CorpInfo.corpName + vbCrLf
    tmp = tmp + "addr(주소) : " + CorpInfo.addr + vbCrLf
    tmp = tmp + "bizType(업태) : " + CorpInfo.bizType + vbCrLf
    tmp = tmp + "bizClass(종목) : " + CorpInfo.bizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원의 회사정보를 수정합니다
' - https://docs.popbill.com/message/vb/api#UpdateCorpInfo
'=========================================================================
Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '대표자명, 최대 100자
    CorpInfo.ceoname = "대표자"
    
    '상호, 최대 200자
    CorpInfo.corpName = "상호"
    
    '주소, 최대 300자
    CorpInfo.addr = "서울특별시"
    
    '업태, 최대 100자
    CorpInfo.bizType = "업태"
    
    '종목, 최대 100자
    CorpInfo.bizClass = "종목"
    
    Set Response = MessageService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 잔여포인트를 확인합니다.
' - 과금방식이 파트너과금인 경우 파트너 잔여포인트(GetPartnerBalance API)를 통해 확인하시기 바랍니다.
' - https://docs.popbill.com/message/vb/api#GetBalance
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
' 연동회원 포인트 충전을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/message/vb/api#GetChargeURL
'=========================================================================
Private Sub btnGetChargeURL_Click()

    Dim url As String
    
    url = MessageService.GetChargeURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 연동회원 포인트 결제내역 확인을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/message/vb/api#GetPaymentURL
'=========================================================================
Private Sub btnGetPaymentURL_Click()
    Dim url As String
           
    url = MessageService.GetPaymentURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 연동회원 포인트 사용내역 확인을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/message/vb/api#GetUseHistoryURL
'=========================================================================
Private Sub btnGetUseHistoryURL_Click()
    Dim url As String
           
    url = MessageService.GetUseHistoryURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 파트너의 잔여포인트를 확인합니다.
' - 과금방식이 연동과금인 경우 연동회원 잔여포인트(GetBalance API)를 이용하시기 바랍니다.
' - https://docs.popbill.com/message/vb/api#GetPartnerBalance
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
' 파트너 포인트 충전 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/message/vb/api#GetPartnerURL
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
    
    url = MessageService.GetPartnerURL(txtCorpNum.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 최대 90byte의 단문(SMS) 메시지 1건 전송을 팝빌에 접수합니다.
' - https://docs.popbill.com/message/vb/api#SendSMS
'=========================================================================
Private Sub btnSendSMS_One_Click()
    Dim Messages As New Collection
    Dim message As New PBMessage
    Dim adsYN As Boolean
    Dim receiptNum As String
    Dim requestNum As String
    Dim UserID As String
    
    '발신번호
    message.sender = "07043042991"
    
    '발신자명
    message.senderName = "발신자명"
    
    '수신번호
    message.receiver = "010111222"
    
    '수신자명
    message.receiverName = "수신자이름"
    
    '메시지 내용, 최대 90Byte 길이를 초과한 내용은 삭제되어 전송됩니다.
    message.content = "발신 내용. 단문은 90Byte로 길이가 조정되어 전송됩니다."
    
    Messages.Add message
    
    '광고문자 전송여부
    adsYN = False
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    '팝빌 회원아이디
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendSMS(txtCorpNum.Text, "", "", Messages, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' 최대 90byte의 단문(SMS) 메시지 다수건 전송을 팝빌에 접수합니다. (최대 1,000건)
' - 수신자마다 개별 내용을 전송할 수 있습니다(대량전송).
' - https://docs.popbill.com/message/vb/api#SendSMS
'=========================================================================
Private Sub btnSendSMS_Hundred_Click()
    Dim Messages As New Collection
    Dim message As PBMessage
    Dim i As Integer
    Dim adsYN As Boolean
    Dim receiptNum As String
    Dim requestNum As String
    Dim UserID As String
    
    '전송정보 배열, 최대 1000건
    For i = 0 To 10
        
        Set message = New PBMessage
        
        '발신번호
        message.sender = "07043042991"
        
        '발신자명
        message.senderName = "발신자명"
        
        '수신번호
        message.receiver = "010111222"
        
        '수신자명
        message.receiverName = "수신자이름_" + CStr(i + 1)
        
        '메시지 내용, 최대 90Byte 길이를 초과한 내용은 삭제되어 전송됩니다.
        message.content = "발신 내용. 단문은 90Byte로 길이가 조정되어 전송됩니다."
        
        Messages.Add message
    Next
    
    '광고문자 전송여부
    adsYN = False
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    '팝빌 회원아이디
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendSMS(txtCorpNum.Text, "", "", Messages, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================
' 최대 90byte의 단문(SMS) 메시지 다수건 전송을 팝빌에 접수합니다. (최대 1,000건)
' - 모든 수신자에게 동일한 내용을 전송합니다(동보전송).
' - https://docs.popbill.com/message/vb/api#SendSMS
'=========================================================================
Private Sub btnSendSMS_Same_Click()
    Dim Messages As New Collection
    Dim message As PBMessage
    Dim sendNum As String
    Dim Contents As String
    Dim i As Integer
    Dim adsYN As Boolean
    Dim receiptNum As String
    Dim requestNum As String
    Dim UserID As String
    
    '동보전송 발신번호
    sendNum = "07043042991"
    
    '메시지 내용, 최대 90Byte 길이를 초과한 내용은 삭제되어 전송됩니다.
    Contents = "동보전송 내용 90byte로 길이가 조정되며, Messages의 내용이 없는 수신건에 동보처리됩니다."
    
    '전송정보 배열, 최대 1000건
    For i = 0 To 10
            
        Set message = New PBMessage
        
        '수신번호
        message.receiver = "010111222"
        
        '수신자명
        message.receiverName = "수신자이름_" + CStr(i + 1)
        
        Messages.Add message
    Next
        
    '광고문자 전송여부
    adsYN = False
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    '팝빌 회원아이디
    UserID = txtUserID.Text
        
    receiptNum = MessageService.SendSMS(txtCorpNum.Text, sendNum, Contents, Messages, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' 최대 2,000byte의 장문(LMS) 메시지 1건 전송을 팝빌에 접수합니다.
' - https://docs.popbill.com/message/vb/api#SendLMS
'=========================================================================
Private Sub btnSendLMS_One_Click()
    Dim Messages As New Collection
    Dim message As New PBMessage
    Dim adsYN As Boolean
    Dim receiptNum As String
    Dim requestNum As String
    Dim UserID As String
    
    '발신번호
    message.sender = "07043042991"
    
    '발신자명
    message.senderName = "발신자명"
    
    '수신번호
    message.receiver = "010111222"
    
    '수신자명
    message.receiverName = "수신자이름"
    
    '메시지 제목
    message.subject = "장문 제목입니다."
    
    '메시지 내용, 최대 2000Byte 길이를 초과한 내용은 삭제되어 전송됩니다.
    message.content = "발신 내용. 장문은 2000Byte로 길이가 조정되어 전송됩니다. 팝빌은 최고의 전자세금계산서 서비스를 제공하고 있습니다."
    
    Messages.Add message
    
    '광고문자 전송여부
    adsYN = False
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    '팝빌 회원아이디
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendLMS(txtCorpNum.Text, "", "", "", Messages, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' 최대 2,000byte의 장문(LMS) 메시지 다수건 전송을 팝빌에 접수합니다. (최대 1,000건)
' - 수신자마다 개별 내용을 전송할 수 있습니다(대량전송).
' - https://docs.popbill.com/message/vb/api#SendLMS
'=========================================================================
Private Sub btnSendLMS_Hundred_Click()
    Dim Messages As New Collection
    Dim message As PBMessage
    Dim i As Integer
    Dim adsYN As Boolean
    Dim receiptNum As String
    Dim requestNum As String
    Dim UserID As String
    
    '전송정보 배열, 최대 1000건
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
        
        '메시지 제목
        message.subject = "장문 제목입니다."
        
        '메시지 내용, 최대 2000Byte 길이를 초과한 내용은 삭제되어 전송됩니다.
        message.content = "발신 내용. 장문은 2000Byte로 길이가 조정되어 전송됩니다. 팝빌은 최고의 전자세금계산서 서비스를 제공하고 있습니다."
        
        Messages.Add message
    Next
    
    '광고문자 전송여부
    adsYN = False
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    '팝빌 회원아이디
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendLMS(txtCorpNum.Text, "", "", "", Messages, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' 최대 2,000byte의 장문(LMS) 메시지 다수건 전송을 팝빌에 접수합니다. (최대 1,000건)
' - 모든 수신자에게 동일한 내용을 전송합니다(동보전송).
' - https://docs.popbill.com/message/vb/api#SendLMS
'=========================================================================
Private Sub btnSendLMS_Same_Click()
    Dim Messages As New Collection
    Dim message As PBMessage
    Dim i As Integer
    Dim adsYN As Boolean
    Dim receiptNum As String
    Dim sender As String
    Dim senderName As String
    Dim subject As String
    Dim Contents As String
    Dim requestNum As String
    Dim UserID As String
    
    '전송정보 배열, 최대 1000건
    For i = 0 To 100

        Set message = New PBMessage
        
        '수신번호
        message.receiver = "11112222"
        
        '수신자명
        message.receiverName = "수신자이름_" + CStr(i + 1)
        Messages.Add message
    Next
    
    '발신번호
    sender = "07043042991"
    
    '발신자명
    senderName = "발신자명"
    
    '메시지 제목
    subject = "동보전송 메시지 제목"
    
    '메시지 내용, 최대 2000Byte 길이를 초과한 내용은 삭제되어 전송됩니다.
    Contents = "메시지 내용. 장문은 2000Byte로 길이가 조정되어 전송됩니다. 팝빌은 최고의 전자세금계산서 서비스를 제공하고 있습니다."
    
    '광고문자 전송여부
    adsYN = False
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    '팝빌 회원아이디
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendLMS(txtCorpNum.Text, sender, subject, Contents, Messages, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================
' 최대 2,000byte의 메시지와 이미지로 구성된 포토문자(MMS) 1건 전송을 팝빌에 접수합니다.
' - 이미지 파일 포맷/규격 : 최대 300Kbyte(JPEG, JPG), 가로/세로 1,000px 이하 권장
' - https://docs.popbill.com/message/vb/api#SendMMS
'=========================================================================
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
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    '팝빌 회원아이디
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendMMS(txtCorpNum.Text, "", "", "", Messages, FilePaths, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================
' 최대 2,000byte의 메시지와 이미지로 구성된 포토문자(MMS) 다수건 전송을 팝빌에 접수합니다. (최대 1,000건)
' - 수신자마다 개별 내용을 전송할 수 있습니다(대량전송).
' - 이미지 파일 포맷/규격 : 최대 300Kbyte(JPEG), 가로/세로 1,000px 이하 권장
' - https://docs.popbill.com/message/vb/api#SendMMS
'=========================================================================
Private Sub btnSendMMS_Hundred_Click()
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
  
   '전송정보 배열, 최대 1000건
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
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    '팝빌 회원아이디
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendMMS(txtCorpNum.Text, "", "", "", Messages, FilePaths, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    txtReceiptNum.Text = receiptNum

End Sub

'=========================================================================
' 최대 2,000byte의 메시지와 이미지로 구성된 포토문자(MMS) 다수건 전송을 팝빌에 접수합니다. (최대 1,000건)
' - 모든 수신자에게 동일한 내용을 전송합니다(동보전송).
' - 이미지 파일 포맷/규격 : 최대 300Kbyte(JPEG), 가로/세로 1,000px 이하 권장
' - https://docs.popbill.com/message/vb/api#SendMMS
'=========================================================================
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
    
    '발신번호
    sender = "07043042991"
    
    '동보메시지 제목
    subject = "동보메시지 제목"
    
    '동보메시지 내용
    Contents = "동보메시지 내용"
    
    '전송정보 배열, 최대 1000건
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
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    '팝빌 회원아이디
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendMMS(txtCorpNum.Text, sender, subject, Contents, Messages, FilePaths, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' 메시지 크기(90byte)에 따라 단문/장문(SMS/LMS)을 자동으로 인식하여 1건의 메시지를 전송을 팝빌에 접수합니다.
' - https://docs.popbill.com/message/vb/api#SendXMS
'=========================================================================
Private Sub btnSendXMS_One_Click()
    Dim Messages As New Collection
    Dim message As New PBMessage
    Dim adsYN As Boolean
    Dim receiptNum As String
    Dim requestNum As String
    Dim UserID As String
    
    '발신자 번호
    message.sender = "07043042991"
    
    '발신자명
    message.senderName = "발신자명"
    
    '수신자 번호
    message.receiver = "010111222"
    
    '수신자명
    message.receiverName = "수신자이름"
    
    '메시지 제목
    message.subject = "장문의 경우 장문 제목"
    
    '메시지 내용, 90byte를 기준으로 단/장문이 자동인식되어 전송됩니다.
    message.content = "자동인식 발송은 내용의 길이를 90Byte기준으로 이하는 단문, 이상은 장문으로 자동 전송합니다."
    
    Messages.Add message
    
    '광고문자 전송여부
    adsYN = False
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    '팝빌 회원아이디
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendXMS(txtCorpNum.Text, "", "", "", Messages, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================
' 메시지 크기(90byte)에 따라 단문/장문(SMS/LMS)을 자동으로 인식하여 다수건의 메시지 전송을 팝빌에 접수합니다. (최대 1,000건)
' - 수신자마다 개별 내용을 전송할 수 있습니다(대량전송).
' - https://docs.popbill.com/message/vb/api#SendXMS
'=========================================================================
Private Sub btnSendXMS_Hundred_Click()
    Dim Messages As New Collection
    Dim message As PBMessage
    Dim i As Integer
    Dim adsYN As Boolean
    Dim receiptNum As String
    Dim requestNum As String
    Dim UserID As String
    
    '전송정보 배열, 최대 1000건
    For i = 0 To 10
    
        Set message = New PBMessage
        
        '발신번호
        message.sender = "07043042991"
        
        '발신자명
        message.senderName = "발신자명"
        
        '수신번호
        message.receiver = "11112222"
        
        '수신자명
        message.receiverName = "수신자이름_" + CStr(i + 1)
        
        '메시지 제목
        message.subject = "장문 제목입니다."
        
        '메시지 내용, 90byte기준으로 단/장문이 자동인식되어 전송됩니다.
        message.content = "발신 내용. 이 내용은 장문으로 전송될수 있도록 길이를 설정하였습니다. 팝빌은 국내 최고의 전자세금계산서 서비스 입니다."
        
        Messages.Add message
    Next

    '광고문자 전송여부
    adsYN = False
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    '팝빌 회원아이디
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendXMS(txtCorpNum.Text, "", "", "", Messages, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    
    txtReceiptNum.Text = receiptNum
    
End Sub

'=========================================================================
' 메시지 크기(90byte)에 따라 단문/장문(SMS/LMS)을 자동으로 인식하여 다수건의 메시지 전송을 팝빌에 접수합니다. (최대 1,000건)
' - 모든 수신자에게 동일한 내용을 전송합니다(동보전송).
' - https://docs.popbill.com/message/vb/api#SendXMS
'=========================================================================
Private Sub btnSendXMS_Same_Click()
    Dim Messages As New Collection
    Dim message As PBMessage
    Dim i As Integer
    Dim subject As String
    Dim content As String
    Dim adsYN As Boolean
    Dim receiptNum As String
    Dim sender As String
    Dim senderName As String
    Dim requestNum As String
    Dim UserID As String
    
    '발신번호
    sender = "07043042991"
    
    '발신자명
    senderName = "발신자명"
    
    '메시지 제목
    subject = "동보전송 제목, 장문에 적용됨"
    
    '메시지 내용, 90byte를 기준으로 단/장문이 자동인식되어 전송됩니다.
    content = "자동인식 발송은 내용의 길이를 90Byte기준으로 이하는 단문, 이상은 장문으로 자동 전송합니다."
    
    '전송정보 배열, 최대 1000건
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
    
    '전송요청번호, 파트너가 전송요청에 대한 관리번호를 직접 할당하여 관리하는 경우 기재
    '최대 36자리, 영문, 숫자, 언더바('_'), 하이픈('-')을 조합하여 사업자별로 중복되지 않도록 구성
    requestNum = ""
    
    '팝빌 회원아이디
    UserID = txtUserID.Text
    
    receiptNum = MessageService.SendLMS(txtCorpNum.Text, sender, subject, content, Messages, txtReserveDT.Text, adsYN, UserID, requestNum)
    
    If receiptNum = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + receiptNum
    
    
    txtReceiptNum.Text = receiptNum
End Sub

'=========================================================================
' 팝빌에서 반환받은 접수번호를 통해 문자 전송상태 및 결과를 확인합니다.
' - https://docs.popbill.com/message/vb/api#GetMessages
'=========================================================================
Private Sub btnGetMessages_Click()
    Dim sentMessages As Collection
    Dim sentMessage As PBSentMsg
    Dim tmp As String
    
    Set sentMessages = MessageService.GetMessages(txtCorpNum.Text, txtReceiptNum.Text)
    
    If sentMessages Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "state(전송상태 코드) | result(전송결과 코드) | subject(메시지 제목) | messageType(메시지 유형) | content(메시지 내용) |  sendNum(발신번호) | senderName(발신자명) | "
    tmp = tmp + "receiveNum(수신번호) | receiveName(수신자명) | receiptDT(접수일시) | reserveDT(예약일시) | "
    tmp = tmp + "sendDT(전송일시) | resultDT(전송결과 수신일시) | tranNet(전송처리 이동통신사명) | receiptNum(접수번호) | requestNum(요청번호)" + vbCrLf
    
    For Each sentMessage In sentMessages
        
        '전송상태 코드
        tmp = tmp + CStr(sentMessage.state) + " | "
        
        '전송결과 코드
        tmp = tmp + CStr(sentMessage.result) + " | "
        
        '메시지 제목
        tmp = tmp + sentMessage.subject + " | "
        
        '메시지 유형
        tmp = tmp + sentMessage.messageType + " | "
        
        '메시지 내용
        tmp = tmp + sentMessage.content + " | "
        
        '발신번호
        tmp = tmp + sentMessage.sendNum + " | "
        
        '발신자명
        tmp = tmp + sentMessage.senderName + " | "
        
        '수신자명
        tmp = tmp + sentMessage.receiveName + " | "
        
        '수신번호
        tmp = tmp + sentMessage.receiveNum + " | "
        
        '접수일시
        tmp = tmp + sentMessage.receiptDT + " | "
        
        '예약일시
        tmp = tmp + sentMessage.reserveDT + " | "
        
        '전송일시
        tmp = tmp + sentMessage.sendDT + " | "
        
        '전송결과 수신일시
        tmp = tmp + sentMessage.resultDT + " | "
        
        '전송처리 이동통신사명
        tmp = tmp + sentMessage.tranNet + " | "
        
        '접수번호
        tmp = tmp + sentMessage.receiptNum + " | "
       
        '요청번호
        tmp = tmp + sentMessage.requestNum
        
        tmp = tmp + vbCrLf
    Next
    
    txtResult.Text = tmp
End Sub

'=========================================================================
' 팝빌에서 반환받은 접수번호를 통해 예약접수된 문자 메시지 전송을 취소합니다. (예약시간 10분 전까지 가능)
' - https://docs.popbill.com/message/vb/api#CancelReserve
'=========================================================================
Private Sub btnCancelReserve_Click()
    Dim Response As PBResponse

    Set Response = MessageService.CancelReserve(txtCorpNum.Text, txtReceiptNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 파트너가 할당한 전송요청 번호를 통해 문자 전송상태 및 결과를 확인합니다.
' - https://docs.popbill.com/message/vb/api#GetMessagesRN
'=========================================================================
Private Sub btnGetMessagesRN_Click()
Dim sentMessages As Collection
    Dim sentMessage As PBSentMsg
    Dim tmp As String
    
    Set sentMessages = MessageService.GetMessagesRN(txtCorpNum.Text, txtRequestNum.Text)
    
    If sentMessages Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "state(전송상태 코드) | result(전송결과 코드) | subject(메시지 제목) | messageType(메시지 유형) | content(메시지 내용) |  sendNum(발신번호) | senderName(발신자명) | "
    tmp = tmp + "receiveNum(수신번호) | receiveName(수신자명) | receiptDT(접수일시) | reserveDT(예약일시) | "
    tmp = tmp + "sendDT(전송일시) | resultDT(전송결과 수신일시) | tranNet(전송처리 이동통신사명) | receiptNum(접수번호) | requestNum(요청번호)" + vbCrLf
    
    For Each sentMessage In sentMessages
            
        ' 전송상태 코드
        tmp = tmp + CStr(sentMessage.state) + " | "
        
        ' 전송결과 코드
        tmp = tmp + CStr(sentMessage.result) + " | "
        
        ' 메시지 제목
        tmp = tmp + sentMessage.subject + " | "
        
        ' 메시지 유형
        tmp = tmp + sentMessage.messageType + " | "
        
        ' 메시지 내용
        tmp = tmp + sentMessage.content + " | "
        
        ' 발신번호
        tmp = tmp + sentMessage.sendNum + " | "
        
        ' 발신자명
        tmp = tmp + sentMessage.senderName + " | "
        
        ' 수신번호
        tmp = tmp + sentMessage.receiveNum + " | "
        
        ' 수신자명
        tmp = tmp + sentMessage.receiveName + " | "
        
        ' 접수일시
        tmp = tmp + sentMessage.receiptDT + " | "
        
        ' 예약일시
        tmp = tmp + sentMessage.reserveDT + " | "
        
        ' 전송일시
        tmp = tmp + sentMessage.sendDT + " | "
        
        ' 전송결과 수신일시
        tmp = tmp + sentMessage.resultDT + " | "
        
        ' 전송처리 이동통신사명
        tmp = tmp + sentMessage.tranNet + " | "
        
        ' 접수번호
        tmp = tmp + sentMessage.receiptNum + " | "
        
        ' 요청번호
        tmp = tmp + sentMessage.requestNum
        
        tmp = tmp + vbCrLf
    Next
    
    txtResult.Text = tmp
    
End Sub

'=========================================================================
' 팝빌에서 반환받은 접수번호를 통해 예약접수된 문자 메시지 전송을 취소합니다. (예약시간 10분 전까지 가능)
' - https://docs.popbill.com/message/vb/api#CancelReserveRN
'=========================================================================
Private Sub btnCancelReserveRN_Click()
    Dim Response As PBResponse
    
    Set Response = MessageService.CancelReserveRN(txtCorpNum.Text, txtRequestNum.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 검색조건에 해당하는 문자 전송내역을 조회합니다. (조회기간 단위 : 최대 2개월)
' - 문자 접수일시로부터 6개월 이내 접수건만 조회할 수 있습니다.
' - https://docs.popbill.com/message/vb/api#Search
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
    Dim QString As String
    
    '[필수] 시작일자, yyyyMMdd
    SDate = "20210901"
    
    '[필수] 종료일자, yyyyMMdd
    EDate = "20210910"
    
    '전송상태값 배열, 1-대기, 2-성공, 3-실패, 4-취소
    state.Add "1"
    state.Add "2"
    state.Add "3"
    
    '검색대상 배열, SMS(단문),LMS(장문),MMS(포토)
    Item.Add "SMS"
    Item.Add "LMS"
    Item.Add "MMS"
    
    '예약문자 검색여부, True(예약문자 조회), False(즉시전송 조회)
    ReserveYN = False
    
    '개인조회여부, True(개인조회), False(전체조회)
    SenderYN = False
    
    '페이지 번호, 기본값 '1'
    Page = 1
    
    '페이지 목록개수, 최대 1000건
    PerPage = 50
    
    '정렬방향, D-내림차순(기본값), A-오름차순
    Order = "D"
    
    '조회 검색어, 발신자명 또는 수신자명 기재
    QString = ""

    Set msgSearchList = MessageService.Search(txtCorpNum.Text, SDate, EDate, state, Item, ReserveYN, SenderYN, Page, PerPage, Order, QString)
     
    If msgSearchList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "code (응답코드) : " + CStr(msgSearchList.code) + vbCrLf
    tmp = tmp + "total (응답메시지) : " + CStr(msgSearchList.total) + vbCrLf
    tmp = tmp + "perPage (페이지당 검색개수) : " + CStr(msgSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum (페이지 번호) : " + CStr(msgSearchList.pageNum) + vbCrLf
    tmp = tmp + "pageCount (페이지 개수) : " + CStr(msgSearchList.pageCount) + vbCrLf
    tmp = tmp + "message (응답메시지) : " + msgSearchList.message + vbCrLf + vbCrLf
    
    tmp = "state(전송상태 코드) | result(전송결과 코드) | subject(팩스제목) | messageType(메시지 타입) | content(메시지 내용) |  sendnum(발신번호) | senderName(발신자명) | "
    tmp = tmp + "receiveNum(수신자명) | receiveName(수신번호) | receiptDT(접수일시) | reserveDT(예약일시) | "
    tmp = tmp + "sendDT(전송일시) | resultDT(전송결과 수신일시) | tranNet(전송처리 이동통신사명) | receiptNum(접수번호) | requestNum(요청번호)" + vbCrLf
            
    Dim info As PBSentMsg
    
    For Each info In msgSearchList.list
    
        '전송상태 코드
        tmp = tmp + CStr(info.state) + " | "
        
        '전송결과 코드
        tmp = tmp + CStr(info.result) + " | "
        
        '메시지 제목
        tmp = tmp + info.subject + " | "
        
        '메시지 유형
        tmp = tmp + info.messageType + " | "
        
        '메시지 내용
        'tmp = tmp + sentMessage.content + " | " ' 내용 표시는 길이관계상 예제에서 생략합니다.
        
        '발신번호
        tmp = tmp + info.sendNum + " | "
        
        '발신자명
        tmp = tmp + info.senderName + " | "
        
        '수신번호
        tmp = tmp + info.receiveNum + " | "
        
        '수신자명
        tmp = tmp + info.receiveName + " | "
        
        '접수일시
        tmp = tmp + info.receiptDT + " | "
        
        '예약일시
        tmp = tmp + info.reserveDT + " | "
        
        '전송일시
        tmp = tmp + info.sendDT + " | "
        
        '전송결과 수신일시
        tmp = tmp + info.resultDT + " | "
        
        '전송처리 이동통신사명
        tmp = tmp + info.tranNet + " | "
        
        '접수번호
        tmp = tmp + info.receiptNum + " | "
        
        '요청번호
        tmp = tmp + info.requestNum
        
        tmp = tmp + vbCrLf
    Next
        
    txtResult.Text = tmp
End Sub

'=========================================================================
' 문자 전송내역 요약정보를 확인합니다. (최대 1000건)
'=========================================================================
Private Sub btnGetStates_Click()
    Dim resultList As Collection
    Dim ReciptNumList As New Collection
    
    '문자 접수번호 배열, 최대 1000건
    ReciptNumList.Add "018061814000000039"
    ReciptNumList.Add "018061815000000002"

    
    Set resultList = MessageService.GetStates(txtCorpNum.Text, ReciptNumList)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "rNum(접수번호) | sn(일련번호) | stat(전송 상태코드) | rlt(전송 결과코드) | sDT(전송일시) | rDT(결과코드 수신일시) |" _
    + "net(전송 이동통신사명) | srt(구 전송결과 코드)" + vbCrLf
    
    Dim info As PBMessageBriefInfo
    
    For Each info In resultList
        tmp = tmp + info.rNum + " | " + info.sn + " | " + info.stat + " | " + info.rlt + " | " + info.sDT + " | "
        tmp = tmp + info.rDT + " | " + info.net + " | " + info.srt + vbCrLf
    Next
    
    MsgBox tmp

End Sub

'=========================================================================
' 팝빌 사이트와 동일한 문자 전송내역 확인 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/message/vb/api#GetSentListURL
'=========================================================================
Private Sub btnGetSentListURL_Click()

    Dim url As String
    
    url = MessageService.GetSentListURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

'=========================================================================
' 전용 080 번호에 등록된 수신거부 목록을 반환합니다.
' - https://docs.popbill.com/message/vb/api#GetAutoDenyList
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
' 팝빌에 등록한 연동회원의 문자 발신번호 목록을 확인합니다.
' - https://docs.popbill.com/message/vb/api#GetSenderNumberList
'=========================================================================
Private Sub btnGetSenderNuberList_Click()
    Dim SenderNumberList As Collection
    Dim tmp As String
    Dim SenderNumberInfo As PBMsgSenderNumber
    
    Set SenderNumberList = MessageService.GetSenderNumberList(txtCorpNum.Text)
    
    If SenderNumberList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
        
    For Each SenderNumberInfo In SenderNumberList
        tmp = tmp + "number(발신번호) : " + SenderNumberInfo.number + vbCrLf
        tmp = tmp + "representYN(대표번호 지정여부) : " + CStr(SenderNumberInfo.number) + vbCrLf
        tmp = tmp + "state(등록상태) : " + CStr(SenderNumberInfo.state) + vbCrLf + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 발신번호를 등록하고 내역을 확인하는 문자 발신번호 관리 페이지 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/message/vb/api#GetSenderNumberMgtURL
'=========================================================================
Private Sub btnGetSenderNumberMgtURL_Click()

    Dim url As String
    
    url = MessageService.GetSenderNumberMgtURL(txtCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(MessageService.LastErrCode) + vbCrLf + "응답메시지 : " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    txtURL.Text = url
End Sub

Private Sub Form_Load()

    '문자서비스 모듈 초기화
    MessageService.Initialize linkID, SecretKey
    
    '연동환경 설정값 True-개발용, False-상업용
    MessageService.IsTest = True
    
    '인증토큰 IP제한기능 사용여부, True-권장
    MessageService.IPRestrictOnOff = True
    
    ' 팝빌 API 서비스 고정 IP 사용여부, True-사용, False-미사용, 기본값(False)
    MessageService.UseStaticIP = False
    
    ' 로컬시스템 시간 사용여부 True-사용, Fasle-미사용, 기본값(False)
    MessageService.UseLocalTimeYN = False
    
End Sub

