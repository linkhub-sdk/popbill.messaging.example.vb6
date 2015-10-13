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
      Left            =   2040
      TabIndex        =   41
      Top             =   6075
      Width           =   1050
   End
   Begin VB.Frame Frame10 
      Caption         =   "포토 전송기능"
      Height          =   945
      Left            =   720
      TabIndex        =   38
      Top             =   5760
      Width           =   3825
      Begin VB.CommandButton btnSendMMS 
         Caption         =   "1건 전송"
         Height          =   465
         Left            =   240
         TabIndex        =   40
         Top             =   315
         Width           =   930
      End
      Begin VB.CommandButton btnSendMMS_Same 
         Caption         =   "동보 전송"
         Height          =   465
         Left            =   2520
         TabIndex        =   39
         Top             =   315
         Width           =   1110
      End
   End
   Begin VB.CommandButton btnUnitCost_MMS 
      Caption         =   "MMS 전송단가 확인"
      Height          =   495
      Left            =   2320
      TabIndex        =   37
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Frame Frame6 
      Caption         =   " 팝빌 메시징 관련 기능"
      Height          =   8415
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Width           =   13005
      Begin VB.CommandButton btnSearch 
         Caption         =   "전송내역 목록조회"
         Height          =   495
         Left            =   10560
         TabIndex        =   51
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton btnSearchPopup 
         Caption         =   "전송내역조회 팝업 URL"
         Height          =   495
         Left            =   7800
         TabIndex        =   36
         Top             =   240
         Width           =   2475
      End
      Begin VB.TextBox txtResult 
         Height          =   4680
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  '양방향
         TabIndex        =   35
         Top             =   3480
         Width           =   12615
      End
      Begin VB.CommandButton btnCancelReserve 
         Caption         =   "예약 전송 취소"
         Height          =   525
         Left            =   10560
         TabIndex        =   34
         Top             =   2160
         Width           =   1665
      End
      Begin VB.CommandButton btnGetMessages 
         Caption         =   "전송상태확인"
         Height          =   525
         Left            =   8760
         TabIndex        =   33
         Top             =   2160
         Width           =   1665
      End
      Begin VB.Frame Frame9 
         Caption         =   " 단/장문 자동인식 문자 전송 "
         Height          =   945
         Left            =   8760
         TabIndex        =   29
         Top             =   960
         Width           =   3825
         Begin VB.CommandButton btnSendXMS_One 
            Caption         =   "1건 전송"
            Height          =   465
            Left            =   240
            TabIndex        =   32
            Top             =   315
            Width           =   930
         End
         Begin VB.CommandButton btnSendXMS_Hundred 
            Caption         =   "100건 전송"
            Height          =   465
            Left            =   1320
            TabIndex        =   31
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendXMS_Same 
            Caption         =   "동보전송"
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
         Left            =   4680
         TabIndex        =   25
         Top             =   960
         Width           =   3825
         Begin VB.CommandButton btnSendLMS_One 
            Caption         =   "1건 전송"
            Height          =   465
            Left            =   360
            TabIndex        =   28
            Top             =   315
            Width           =   930
         End
         Begin VB.CommandButton btnSendLMS_Hundred 
            Caption         =   "100건 전송"
            Height          =   465
            Left            =   1440
            TabIndex        =   27
            Top             =   315
            Width           =   1110
         End
         Begin VB.CommandButton btnSendLMS_Same 
            Caption         =   "동보전송"
            Height          =   465
            Left            =   2640
            TabIndex        =   26
            Top             =   315
            Width           =   1020
         End
      End
      Begin VB.TextBox txtReceiptNum 
         Height          =   315
         Left            =   5685
         TabIndex        =   24
         Top             =   2265
         Width           =   2850
      End
      Begin VB.Frame Frame7 
         Caption         =   " 단문 문자 전송 "
         Height          =   945
         Left            =   600
         TabIndex        =   18
         Top             =   960
         Width           =   3825
         Begin VB.CommandButton btnSendSMS_Same 
            Caption         =   "동보전송"
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
         Left            =   3660
         TabIndex        =   16
         Top             =   375
         Width           =   3105
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "접수번호 : "
         Height          =   180
         Left            =   4785
         TabIndex        =   23
         Top             =   2355
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "예약시간(yyyyMMddHHmmss) : "
         Height          =   180
         Left            =   825
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
         Height          =   2295
         Left            =   10920
         TabIndex        =   48
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "회사정보 수정"
            Height          =   495
            Left            =   120
            TabIndex        =   50
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "회사정보 조회"
            Height          =   495
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "담당자 관련"
         Height          =   2295
         Left            =   8880
         TabIndex        =   44
         Top             =   240
         Width           =   1935
         Begin VB.CommandButton btnUpdateContact 
            Caption         =   "담당자 정보 수정"
            Height          =   495
            Left            =   120
            TabIndex        =   47
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton btnListContact 
            Caption         =   "담당자 목록 조회"
            Height          =   495
            Left            =   120
            TabIndex        =   46
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton btnRegistContact 
            Caption         =   "담당자 추가"
            Height          =   495
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 회원정보"
         Height          =   2295
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1695
         Begin VB.CommandButton btnCheckID 
            Caption         =   "ID 중복 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   42
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton btnCheckIsMember 
            Caption         =   "가입 여부 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnJoinMember 
            Caption         =   "회원 가입"
            Height          =   495
            Left            =   120
            TabIndex        =   12
            Top             =   1560
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " 포인트 관련"
         Height          =   2295
         Left            =   2040
         TabIndex        =   9
         Top             =   240
         Width           =   2160
         Begin VB.CommandButton btnUnitCost_LMS 
            Caption         =   "LMS 전송단가 확인"
            Height          =   495
            Left            =   165
            TabIndex        =   14
            Top             =   975
            Width           =   1815
         End
         Begin VB.CommandButton btnUnitCost 
            Caption         =   "SMS 전송단가 확인"
            Height          =   495
            Left            =   150
            TabIndex        =   10
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " 파트너 관련"
         Height          =   2295
         Left            =   4320
         TabIndex        =   7
         Top             =   240
         Width           =   2535
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여 포인트 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   2295
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "파트너 잔여포인트 확인"
            Height          =   495
            Left            =   120
            TabIndex        =   8
            Top             =   960
            Width           =   2295
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " 팝빌 기본 URL"
         ClipControls    =   0   'False
         Height          =   2295
         Left            =   6960
         TabIndex        =   5
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnGetPopbillURL_CHRG 
            Caption         =   "포인트 충전 URL"
            Height          =   495
            Left            =   120
            TabIndex        =   43
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton btnGetPopbillURL_LOGIN 
            Caption         =   " 팝빌 로그인 URL"
            Height          =   495
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
Option Explicit

'링크아이디
Private Const LinkID = "TESTER"
'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

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

Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = MessageService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = MessageService.CheckIsMember(txtCorpNum.Text, LinkID)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub


Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = MessageService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
    
    
End Sub

Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    
    Set CorpInfo = MessageService.GetCorpInfo(txtCorpNum.Text, txtUserID.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "ceoname : " + CorpInfo.CEOName + vbCrLf
    tmp = tmp + "corpName : " + CorpInfo.CorpName + vbCrLf
    tmp = tmp + "addr : " + CorpInfo.Addr + vbCrLf
    tmp = tmp + "bizType : " + CorpInfo.BizType + vbCrLf
    tmp = tmp + "bizClass : " + CorpInfo.BizClass + vbCrLf
    
    MsgBox tmp
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
    tmp = "state | subject | messageType | sendnum | receiveNum | receiveName | reserveDT | sendDT | sendResult | tranNet" + vbCrLf
    
    For Each sentMessage In sentMessages
    
        tmp = tmp + CStr(sentMessage.state) + " | "
        tmp = tmp + sentMessage.subject + " | "
        tmp = tmp + sentMessage.messageType + " | "
        'tmp = tmp + sentMessage.content + " | " ' 내용 표시는 길이관계상 예제에서 생략합니다.
        tmp = tmp + sentMessage.sendNum + " | "
        tmp = tmp + sentMessage.receiveNum + " | "
        tmp = tmp + sentMessage.receiveName + " | "
        tmp = tmp + sentMessage.reserveDT + " | "
        tmp = tmp + sentMessage.sendDT + " | "
        tmp = tmp + sentMessage.resultDT + " | "
        tmp = tmp + sentMessage.sendResult + " | "
        tmp = tmp + sentMessage.tranNet
        
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
    
    MsgBox "잔여포인트 : " + CStr(balance)
    
End Sub

Private Sub btnGetPopbillURL_CHRG_Click()
    Dim url As String
    
    url = MessageService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "CHRG")
    
    If url = "" Then
         MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnGetPopbillURL_LOGIN_Click()
    Dim url As String
    
    url = MessageService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "LOGIN")
    
    If url = "" Then
         MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    joinData.LinkID = LinkID '링크 아이디
    joinData.CorpNum = "1231212312" '사업자번호 "-" 제외.
    joinData.CEOName = "대표자성명"
    joinData.CorpName = "회원상호"
    joinData.Addr = "주소"
    joinData.ZipCode = "500-100"
    joinData.BizType = "업태"
    joinData.BizClass = "업종"
    joinData.ID = "userid"      '6자 이상 20자 미만.
    joinData.PWD = "pwd_must_be_long_enough"    '6자 이상 20자 미만.
    joinData.ContactName = "담당자성명"
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

Private Sub btnListContact_Click()
    Dim resultList As Collection
        
    Set resultList = MessageService.ListContact(txtCorpNum.Text, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "id | email | hp | personName | searchAllAllowYN | tel | fax | mgrYN | regDT " + vbCrLf
    
    Dim info As PBContactInfo
    
    For Each info In resultList
        tmp = tmp + info.ID + " | " + info.email + " | " + info.hp + " | " + info.personName + " | " + CStr(info.searchAllAllowYN) _
                + info.tel + " | " + info.fax + " | " + CStr(info.mgrYN) + " | " + info.regDT + vbCrLf
    Next
    
    MsgBox tmp
End Sub

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    joinData.ID = "testkorea_20151007"      '담당자 아이디
    joinData.PWD = "test@test.com"          '비밀번호
    joinData.personName = "담당자명"        '담당자명
    joinData.tel = "070-1234-1234"          '연락처
    joinData.hp = "010-1234-1234"           '휴대폰번호
    joinData.email = "test@test.com"        '이메일 주소
    joinData.fax = "070-1234-1234"          '팩스번호
    joinData.searchAllAllowYN = True        '전체조회여부, Ture-회사조회, False-개인조회
    joinData.mgrYN = False                  '관리자 권한여부
        
    Set Response = MessageService.RegistContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

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
    
    SDate = "20151001"    '[필수] 시작일자, yyyyMMdd
    EDate = "20151008"    '[필수] 종료일자, yyyyMMdd
    
    state.Add "1"         '전송상태값 배열, 1-대기, 2-성공, 3-실패, 4-취소
    state.Add "2"
    state.Add "3"
    
    Item.Add "SMS"        '검색대상 배열, SMS(단문),LMS(장문),MMS(포토)
    Item.Add "LMS"
    
    ReserveYN = False     '예약문자 검색여부, True(예약문자만 조회), False(전체조회)
    SenderYN = False      '개인조회여부, True(개인조회), False(전체조회)
    
    Page = 1              '페이지 번호
    PerPage = 10          '페이지 목록개수, 최대 1000건

    Set msgSearchList = MessageService.Search(txtCorpNum.Text, SDate, EDate, state, Item, ReserveYN, SenderYN, Page, PerPage)
     
    If msgSearchList Is Nothing Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    tmp = "code : " + CStr(msgSearchList.code) + vbCrLf
    tmp = tmp + "total : " + CStr(msgSearchList.total) + vbCrLf
    tmp = tmp + "perPage : " + CStr(msgSearchList.PerPage) + vbCrLf
    tmp = tmp + "pageNum : " + CStr(msgSearchList.pageNum) + vbCrLf
    tmp = tmp + "perCount : " + CStr(msgSearchList.pageCount) + vbCrLf
    tmp = tmp + "message : " + msgSearchList.message + vbCrLf + vbCrLf
    
    tmp = tmp + "state | subject | messageType | sendnum | receiveNum | receiveName | reserveDT | sendDT | sendResult | tranNet" + vbCrLf
            
    Dim info As PBSentMsg
    
    For Each info In msgSearchList.list
        tmp = tmp + CStr(info.state) + " | "
        tmp = tmp + info.subject + " | "
        tmp = tmp + info.messageType + " | "
        'tmp = tmp + sentMessage.content + " | " ' 내용 표시는 길이관계상 예제에서 생략합니다.
        tmp = tmp + info.sendNum + " | "
        tmp = tmp + info.receiveNum + " | "
        tmp = tmp + info.receiveName + " | "
        tmp = tmp + info.reserveDT + " | "
        tmp = tmp + info.sendDT + " | "
        tmp = tmp + info.resultDT + " | "
        tmp = tmp + info.sendResult + " | "
        tmp = tmp + info.tranNet
        tmp = tmp + vbCrLf
    Next
        
    txtResult.Text = tmp
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
    Dim adsYN As Boolean
    adsYN = False       '광고문자 전송여부
    
    Dim message As PBMessage
    
    Dim i As Integer
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        message.sender = "07075106766"
        message.receiver = "11112222"
        message.receiverName = "수신자이름_" + CStr(i + 1)
        message.content = "발신 내용. 장문은 2000Byte로 길이가 조정되어 전송됩니다. 팝빌은 최고의 전자세금계산서 서비스를 제공하고 있습니다."
        message.subject = "장문 제목입니다."
        
        Messages.Add message
    Next
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendLMS(txtCorpNum.Text, "", "", "", Messages, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
    
End Sub

Private Sub btnSendLMS_One_Click()
    
    Dim Messages As New Collection
    Dim adsYN As Boolean
    adsYN = False       '광고문자 전송여부
    
    Dim message As New PBMessage
    
    message.sender = "07075106766"
    message.receiver = "11112222"
    message.receiverName = "수신자이름"
    message.content = "발신 내용. 장문은 2000Byte로 길이가 조정되어 전송됩니다. 팝빌은 최고의 전자세금계산서 서비스를 제공하고 있습니다."
    message.subject = "장문 제목입니다."
    
    Messages.Add message
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendLMS(txtCorpNum.Text, "", "", "", Messages, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
    
    
End Sub

Private Sub btnSendLMS_Same_Click()
        
    Dim Messages As New Collection
    Dim adsYN As Boolean
    adsYN = False       '광고문자 전송여부
    
    Dim message As PBMessage
    
    Dim i As Integer
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        message.receiver = "11112222"
        message.receiverName = "수신자이름_" + CStr(i + 1)
        
        Messages.Add message
    Next
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendLMS(txtCorpNum.Text, "07075106766", "동보전송 제목", "발신 내용. 장문은 2000Byte로 길이가 조정되어 전송됩니다. 팝빌은 최고의 전자세금계산서 서비스를 제공하고 있습니다.", _
                                    Messages, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendMMS_Click()
    Dim Messages As New Collection
    Dim FilePaths As New Collection
    Dim adsYN As Boolean
    adsYN = True       '광고문자 전송여부
    
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    FilePaths.Add CommonDialog1.FileName
    
    Dim message As New PBMessage
    
    message.sender = "07075103710"
    message.receiver = "010111222"
    message.receiverName = "수신자이름"
    message.content = "MMS 발신 테스트 내용."
    message.subject = "메시지 제목"
    
    Messages.Add message
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendMMS(txtCorpNum.Text, "07075103710", "동보제목", "동보내용", Messages, FilePaths, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
    
End Sub

Private Sub btnSendMMS_hundred_Click()
    Dim Messages As New Collection
    Dim FilePaths As New Collection
    Dim adsYN As Boolean
    adsYN = False       '광고문자 전송여부
    
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    FilePaths.Add CommonDialog1.FileName
  
    Dim message As PBMessage
    
    Dim i As Integer
    
    For i = 0 To 50
        
        Set message = New PBMessage
        
        message.sender = "07075106766"
        message.receiver = "11112222"
        message.receiverName = "수신자이름_" + CStr(i + 1)
        message.content = "발신 내용. 이 내용은 장문으로 전송될수 있도록 길이를 설정하였습니다. 팝빌은 국내 최고의 전자세금계산서 서비스 입니다."
        message.subject = "장문 제목입니다."
        
        Messages.Add message
    Next
    
    For i = 0 To 50
        
        Set message = New PBMessage
        
        message.sender = "07075106766"
        message.receiver = "11112222"
        message.receiverName = "수신자이름_" + CStr(i + 1)
        message.content = "발신 내용. 이 내용은 단문으로 전송됩니다."
        
        Messages.Add message
    Next
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendMMS(txtCorpNum.Text, "07075103710", "동보제목", "동보내용", Messages, FilePaths, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendMMS_Same_Click()
    Dim Messages As New Collection
    Dim FilePaths As New Collection
    Dim adsYN As Boolean
    adsYN = False       '광고문자 전송여부
    
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then Exit Sub
    
    FilePaths.Add CommonDialog1.FileName
  
    Dim message As PBMessage
    Dim i As Integer
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        message.receiver = "010111222"
        message.receiverName = "수신자이름_" + CStr(i + 1)
        
        Messages.Add message
    Next
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendMMS(txtCorpNum.Text, "07075103710", "동보제목", "동보내용", Messages, FilePaths, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
    
    
End Sub

Private Sub btnSendSMS_hundredd_Click()
    Dim Messages As New Collection
    Dim adsYN As Boolean
    adsYN = False       '광고문자 전송여부
    
    Dim message As PBMessage
    
    Dim i As Integer
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        message.sender = "07075106766"
        message.receiver = "11112222"
        message.receiverName = "수신자이름_" + CStr(i + 1)
        message.content = "발신 내용. 단문은 90Byte로 길이가 조정되어 전송됩니다."
        
        Messages.Add message
    Next
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendSMS(txtCorpNum.Text, "07075103710", "", Messages, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
    
    
End Sub

Private Sub btnSendSMS_One_Click()
    
    Dim Messages As New Collection
    Dim adsYN As Boolean
    adsYN = False       '광고문자 전송여부
    
    Dim message As New PBMessage
    
    message.sender = "07075103710"
    message.receiver = "010111222"
    message.receiverName = "수신자이름"
    message.content = "발신 내용. 단문은 90Byte로 길이가 조정되어 전송됩니다."
    
    Messages.Add message
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendSMS(txtCorpNum.Text, "", "", Messages, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
    
    
End Sub

Private Sub btnSendSMS_Same_Click()
        
    Dim Messages As New Collection
    Dim adsYN As Boolean
    adsYN = False       '광고문자 전송여부
    
    Dim message As PBMessage
    
    Dim i As Integer
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        message.receiver = "11112222"
        message.receiverName = "수신자이름_" + CStr(i + 1)
        
        Messages.Add message
    Next
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendSMS(txtCorpNum.Text, "07075106766", "동보전송 내용 90byte로 길이가 조정되며, Messages의 내용이 없는 수신건에 동보처리됩니다.", Messages, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
    
End Sub

Private Sub btnSendXMS_Hundred_Click()
    Dim Messages As New Collection
    Dim adsYN As Boolean
    adsYN = False       '광고문자 전송여부
    
    Dim message As PBMessage
    
    Dim i As Integer
    
    For i = 0 To 50
        
        Set message = New PBMessage
        
        message.sender = "07075106766"
        message.receiver = "11112222"
        message.receiverName = "수신자이름_" + CStr(i + 1)
        message.content = "발신 내용. 이 내용은 장문으로 전송될수 있도록 길이를 설정하였습니다. 팝빌은 국내 최고의 전자세금계산서 서비스 입니다."
        message.subject = "장문 제목입니다."
        
        Messages.Add message
    Next
    
    For i = 0 To 50
        
        Set message = New PBMessage
        
        message.sender = "07075106766"
        message.receiver = "11112222"
        message.receiverName = "수신자이름_" + CStr(i + 1)
        message.content = "발신 내용. 이 내용은 단문으로 전송됩니다."
        
        Messages.Add message
    Next
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendXMS(txtCorpNum.Text, "", "", "", Messages, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendXMS_One_Click()
    
    Dim Messages As New Collection
    Dim adsYN As Boolean
    adsYN = False       '광고문자 전송여부
    
    Dim message As New PBMessage
    
    message.sender = "07075106766"
    message.receiver = "01041680206"
    message.receiverName = "수신자이름"
    message.content = "자동인식 발송은 내용의 길이를 90Byte기준으로 이하는 단문, 이상은 장문으로 자동 전송합니다."
    message.subject = "장문의 경우 장문 제목입니다."
    
    Messages.Add message
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendXMS(txtCorpNum.Text, "", "", "", Messages, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnSendXMS_Same_Click()
        
    Dim Messages As New Collection
    Dim adsYN As Boolean
    adsYN = False       '광고문자 전송여부
    
    Dim message As PBMessage
    
    Dim i As Integer
    
    For i = 0 To 100
        
        Set message = New PBMessage
        
        message.receiver = "11112222"
        message.receiverName = "수신자이름_" + CStr(i + 1)
        
        Messages.Add message
    Next
    
    Dim ReceiptNum As String
    
    ReceiptNum = MessageService.SendLMS(txtCorpNum.Text, "07075106766", "동보전송 제목, 장문에 적용됨", _
                                        "자동인식 발송은 내용의 길이를 90Byte기준으로 이하는 단문, 이상은 장문으로 자동 전송합니다.", _
                                        Messages, txtReserveDT.Text, adsYN, txtUserID.Text)
    
    If ReceiptNum = "" Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "접수 번호 : " + ReceiptNum
    txtReceiptNum.Text = ReceiptNum
End Sub

Private Sub btnUnitCost_Click()
    Dim unitCost As Single
    
    unitCost = MessageService.GetUnitCost(txtCorpNum.Text, SMS)
    
    If unitCost < 0 Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "SMS 전송 단가 : " + CStr(unitCost)
    
End Sub

Private Sub btnUnitCost_LMS_Click()
    Dim unitCost As Single
    
    unitCost = MessageService.GetUnitCost(txtCorpNum.Text, LMS)
    
    If unitCost < 0 Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "LMS 전송 단가 : " + CStr(unitCost)
End Sub


Private Sub btnUnitCost_MMS_Click()
    Dim unitCost As Single
    
    unitCost = MessageService.GetUnitCost(txtCorpNum.Text, MMS)
    
    If unitCost < 0 Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "MMS 전송 단가 : " + CStr(unitCost)
End Sub

Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    joinData.personName = "담당자명_수정"  '담당자명
    joinData.tel = "070-1234-1234"         '연락처
    joinData.hp = "010-1234-1234"          '휴대폰번호
    joinData.email = "test@test.com"       '이메일 주소
    joinData.fax = "070-1234-1234"         '팩스번호
    joinData.searchAllAllowYN = True       '전체조회여부, Ture-회사조회, False-개인조
    joinData.mgrYN = False                 '관리자 권한여부
                
    Set Response = MessageService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    CorpInfo.CEOName = "대표자"         '대표자명
    CorpInfo.CorpName = "상호_수정"          '상호명
    CorpInfo.Addr = "서울특별시"        '주소
    CorpInfo.BizType = "업태"           '업태
    CorpInfo.BizClass = "업종"          '업종
    
    Set Response = MessageService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(MessageService.LastErrCode) + "] " + MessageService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub Form_Load()
    MessageService.Initialize LinkID, SecretKey
    
    '연동환경 설정값 True-테스트용, False-상업용
    MessageService.IsTest = True
    
End Sub
