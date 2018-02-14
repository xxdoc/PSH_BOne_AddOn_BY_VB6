Attribute VB_Name = "SubMain"
Public Sbo_Application  As SAPbouiCOM.Application
Public Sbo_Company      As SAPbobsCOM.Company
Public FormCurrentCount As Long '//현재 폼의총갯수
Public FormTotalCount   As Long '//생성한 폼의총갯수
Public ClassList        As Collection '//컬렉션 개체
Public ShareFolderPath  As String '//공유폴더주소
Public ServerPath       As String '//서버주소
Public oZSBO            As ZZMDC

Public Sub Main()
'******************************************************************************
'Function ID    : Main
'해 당 모 듈    : SubMain
'기       능    : ZZMDC 클래스의 인스턴스를 호출, 시스템에서 최초로 실행
'인       수    : 없음
'반   환  값    : 없음
'특이사항       : 없음
'******************************************************************************
    If App.PrevInstance = True Then
        Instance_Flg = False
        MsgBox App.EXEName + "는 현재실행중입니다.", vbExclamation, "oTemp.exe"
        End
    End If
    
    Set oZSBO = New ZZMDC

    Do While GetMessage(Msg, 0&, 0&, 0&)    'Message Loop
        TranslateMessage Msg
        DispatchMessage Msg
        DoEvents
    Loop
    
End Sub

'*******************************************************************
'//폼객체추가
'*******************************************************************
Public Sub AddForms(ByVal cObject As Variant, ByVal oFormUid As String)
    ClassList.Add cObject, oFormUid
    FormTotalCount = FormTotalCount + 1
    FormCurrentCount = FormCurrentCount + 1
End Sub

'*******************************************************************
'//폼객체제거
'*******************************************************************
Public Sub RemoveForms(ByVal oFormUniqueID As String)
    Dim oTempClass As Variant
    Set oTempClass = ClassList.Item(oFormUniqueID)
    ClassList.Remove oFormUniqueID
    Set oTempClass = Nothing
    FormCurrentCount = FormCurrentCount - 1
End Sub

'*******************************************************************
'//폼현재객체수
'*******************************************************************
Public Function GetCurrentFormsCount() As Long
    GetCurrentFormsCount = FormCurrentCount
End Function

'*******************************************************************
'//폼총객체수
'*******************************************************************
Public Function GetTotalFormsCount() As Long
    GetTotalFormsCount = FormTotalCount
End Function
