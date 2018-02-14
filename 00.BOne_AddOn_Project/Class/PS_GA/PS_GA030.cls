VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_GA030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File               : PS_GA030.cls
'//  Module             : 총무관리 > 예산 및 비용관리
'//  Desc               : 예산 등록
'//  FormType           : PS_GA030
'//  Create Date(Start) : 2014.03.28
'//  Create Date(End)   :
'//  Creator            : Song Myoung gyu
'//  Modified Date      :
'//  Modifier           :
'//  Company            : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01         As SAPbouiCOM.Form
Public oMat01          As SAPbouiCOM.Matrix

Private oDS_PS_GA030L  As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01  As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01  As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private gDeleteRowCount As Integer '행삭제 카운트

'//사용자구조체
Private Type DeleteRowInfo
    DocEntry As Integer
End Type

Private oLast_Mode&

Private DeleteRow() As DeleteRowInfo

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_GA030.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_GA030_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_GA030_CreateItems
    Call PS_GA030_ComboBox_Setting
    Call PS_GA030_Initial_Setting
    Call PS_GA030_CF_ChooseFromList
    Call PS_GA030_EnableMenus
    Call PS_GA030_SetDocument(oFromDocEntry01)
    Call PS_GA030_FormResize

'    Call PS_GA030_Add_MatrixRow(0, True)
    Call PS_GA030_LoadCaption
    Call PS_GA030_FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1284"), False         '// 취소
    oForm01.EnableMenu ("1293"), False         '// 행삭제
    oForm01.EnableMenu ("1281"), False
    oForm01.EnableMenu ("1282"), True

'    Dim sQry As String
'    Dim RecordSet01 As SAPbobsCOM.Recordset
'    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    '관리번호
'    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_GA030H]"
'    Call RecordSet01.DoQuery(sQry)
'    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
'        Call oDS_PS_GA030H.setValue("DocEntry", 0, 1)
'    Else
'        Call oDS_PS_GA030H.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
'    End If
    
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing
 
    Exit Sub
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA030_LoadCaption()
'******************************************************************************
'Function ID : PS_GA030_LoadCaption()
'해당모듈    : PS_GA030
'기능        : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA030_LoadCaption_Error

    Call oForm01.Freeze(True)

    If oForm01.Mode = fm_ADD_MODE Then
    
        oForm01.Items("BtnAdd").Specific.Caption = "추가"

'    ElseIf oForm01.Mode = fm_OK_MODE Then
'        oForm01.Items("BtnAdd").Specific.Caption = "확인"

    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
    
        oForm01.Items("BtnAdd").Specific.Caption = "수정"
        
    End If
    
    Call oForm01.Freeze(False)

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA030_LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/메트릭스 Row추가
Sub PS_GA030_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_GA030_Add_MatrixRow_Error

    If RowIserted = False Then '//행추가여부
        oDS_PS_GA030L.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_GA030L.Offset = oRow
    oDS_PS_GA030L.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA030_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PS_GA030_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_GA030_MTX01(ByVal pSearchType As String)
'******************************************************************************
'Function ID : PS_GA030_MTX01()
'해당모듈    : PS_GA030
'기능        : 데이터 조회
'인수        : pSearchType : 조회 구분(1:실적조회, 2:계획조회)
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA030_MTX01_Error

    Dim loopCount As Integer
    Dim sQry      As String
    Dim ErrNum    As Integer
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID      As String '사업장
    Dim StdYear    As String '기준년도
    Dim AcctCode01 As String '계정코드
    Dim AcctCode02 As String '계정과목코드
    Dim AcctCode03 As String '세부계정과목코드
    
    BPLID = oForm01.Items("BPLID").Specific.Selected.VALUE
    StdYear = oForm01.Items("StdYear").Specific.VALUE
    AcctCode01 = oForm01.Items("AcctCode01").Specific.VALUE
    AcctCode02 = oForm01.Items("AcctCode02").Specific.VALUE
    AcctCode03 = oForm01.Items("AcctCode03").Specific.VALUE
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    If pSearchType = "1" Then '실적조회
    
        sQry = "      EXEC [PS_GA030_01] '"
        sQry = sQry & BPLID & "','"
        sQry = sQry & StdYear & "','"
        sQry = sQry & AcctCode01 & "','"
        sQry = sQry & AcctCode02 & "','"
        sQry = sQry & AcctCode03 & "'"
    
    Else
    
        sQry = "      EXEC [PS_GA030_02] '" '계획조회
        sQry = sQry & BPLID & "','"
        sQry = sQry & StdYear & "','"
        sQry = sQry & AcctCode01 & "','"
        sQry = sQry & AcctCode02 & "','"
        sQry = sQry & AcctCode03 & "'"
    
    End If
    
    Call RecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_GA030L.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (RecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
'        Call PS_GA030_Add_MatrixRow(0, True)
        Call PS_GA030_LoadCaption
        
        GoTo PS_GA030_MTX01_Error
        
        Exit Sub
    End If
    
    For loopCount = 0 To RecordSet01.RecordCount - 1
        If loopCount + 1 > oDS_PS_GA030L.Size Then
            Call oDS_PS_GA030L.InsertRecord(loopCount)
        End If
        
        Call oMat01.AddRow
        oDS_PS_GA030L.Offset = loopCount
        
        Call oDS_PS_GA030L.setValue("U_LineNum", loopCount, loopCount + 1)
        Call oDS_PS_GA030L.setValue("U_ColReg01", loopCount, RecordSet01.Fields("StdYear").VALUE) '년도
        Call oDS_PS_GA030L.setValue("U_ColReg02", loopCount, RecordSet01.Fields("BPLID").VALUE) '사업장코드
        Call oDS_PS_GA030L.setValue("U_ColReg03", loopCount, RecordSet01.Fields("AcctCode01").VALUE) '계정코드
        Call oDS_PS_GA030L.setValue("U_ColReg04", loopCount, RecordSet01.Fields("AcctName01").VALUE) '계정명
        Call oDS_PS_GA030L.setValue("U_ColReg05", loopCount, RecordSet01.Fields("AcctCode02").VALUE) '계정과목코드
        Call oDS_PS_GA030L.setValue("U_ColReg06", loopCount, RecordSet01.Fields("AcctName02").VALUE) '계정과목명
        Call oDS_PS_GA030L.setValue("U_ColReg07", loopCount, RecordSet01.Fields("AcctCode03").VALUE) '세부계정과목코드
        Call oDS_PS_GA030L.setValue("U_ColReg08", loopCount, RecordSet01.Fields("AcctName03").VALUE) '세부계정과목명
        Call oDS_PS_GA030L.setValue("U_ColSum01", loopCount, RecordSet01.Fields("Month01").VALUE) '1월
        Call oDS_PS_GA030L.setValue("U_ColSum02", loopCount, RecordSet01.Fields("Month02").VALUE) '2월
        Call oDS_PS_GA030L.setValue("U_ColSum03", loopCount, RecordSet01.Fields("Month03").VALUE) '3월
        Call oDS_PS_GA030L.setValue("U_ColSum04", loopCount, RecordSet01.Fields("Month04").VALUE) '4월
        Call oDS_PS_GA030L.setValue("U_ColSum05", loopCount, RecordSet01.Fields("Month05").VALUE) '5월
        Call oDS_PS_GA030L.setValue("U_ColSum06", loopCount, RecordSet01.Fields("Month06").VALUE) '6월
        Call oDS_PS_GA030L.setValue("U_ColSum07", loopCount, RecordSet01.Fields("Month07").VALUE) '7월
        Call oDS_PS_GA030L.setValue("U_ColSum08", loopCount, RecordSet01.Fields("Month08").VALUE) '8월
        Call oDS_PS_GA030L.setValue("U_ColSum09", loopCount, RecordSet01.Fields("Month09").VALUE) '9월
        Call oDS_PS_GA030L.setValue("U_ColSum10", loopCount, RecordSet01.Fields("Month10").VALUE) '10월
        Call oDS_PS_GA030L.setValue("U_ColSum11", loopCount, RecordSet01.Fields("Month11").VALUE) '11월
        Call oDS_PS_GA030L.setValue("U_ColSum12", loopCount, RecordSet01.Fields("Month12").VALUE) '12월
        Call oDS_PS_GA030L.setValue("U_ColSum13", loopCount, RecordSet01.Fields("Total").VALUE) '계
        Call oDS_PS_GA030L.setValue("U_ColPrc01", loopCount, RecordSet01.Fields("Rate").VALUE) '비율
    
        RecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & RecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop
    Call oForm01.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set RecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA030_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set RecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_GA030_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_GA030_Calculate_Header()
'******************************************************************************
'Function ID : PS_GA030_Calculate_Header()
'해당모듈    : PS_GA030
'기능        : 실적*비율=예산 (전체 계산)
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA030_Calculate_Header_Error

    Dim loopCount As Integer
    Dim sQry      As String
    Dim ErrNum    As Integer
    
'    Dim RecordSet01 As SAPbobsCOM.Recordset
'    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim rate As Currency
    rate = oForm01.Items("Rate").Specific.VALUE * 0.01
    
    Call oForm01.Freeze(True)
    Call oMat01.FlushToDataSource

    For loopCount = 0 To oMat01.RowCount - 1
        
        Call oDS_PS_GA030L.setValue("U_ColSum01", loopCount, oDS_PS_GA030L.GetValue("U_ColSum01", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum01", loopCount) * rate)) '1월
        Call oDS_PS_GA030L.setValue("U_ColSum02", loopCount, oDS_PS_GA030L.GetValue("U_ColSum02", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum02", loopCount) * rate)) '2월
        Call oDS_PS_GA030L.setValue("U_ColSum03", loopCount, oDS_PS_GA030L.GetValue("U_ColSum03", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum03", loopCount) * rate)) '3월
        Call oDS_PS_GA030L.setValue("U_ColSum04", loopCount, oDS_PS_GA030L.GetValue("U_ColSum04", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum04", loopCount) * rate)) '4월
        Call oDS_PS_GA030L.setValue("U_ColSum05", loopCount, oDS_PS_GA030L.GetValue("U_ColSum05", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum05", loopCount) * rate)) '5월
        Call oDS_PS_GA030L.setValue("U_ColSum06", loopCount, oDS_PS_GA030L.GetValue("U_ColSum06", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum06", loopCount) * rate)) '6월
        Call oDS_PS_GA030L.setValue("U_ColSum07", loopCount, oDS_PS_GA030L.GetValue("U_ColSum07", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum07", loopCount) * rate)) '7월
        Call oDS_PS_GA030L.setValue("U_ColSum08", loopCount, oDS_PS_GA030L.GetValue("U_ColSum08", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum08", loopCount) * rate)) '8월
        Call oDS_PS_GA030L.setValue("U_ColSum09", loopCount, oDS_PS_GA030L.GetValue("U_ColSum09", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum09", loopCount) * rate)) '9월
        Call oDS_PS_GA030L.setValue("U_ColSum10", loopCount, oDS_PS_GA030L.GetValue("U_ColSum10", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum10", loopCount) * rate)) '10월
        Call oDS_PS_GA030L.setValue("U_ColSum11", loopCount, oDS_PS_GA030L.GetValue("U_ColSum11", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum11", loopCount) * rate)) '11월
        Call oDS_PS_GA030L.setValue("U_ColSum12", loopCount, oDS_PS_GA030L.GetValue("U_ColSum12", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum12", loopCount) * rate)) '12월
        Call oDS_PS_GA030L.setValue("U_ColSum13", loopCount, oDS_PS_GA030L.GetValue("U_ColSum13", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum13", loopCount) * rate)) '계
        Call oDS_PS_GA030L.setValue("U_ColPrc01", loopCount, rate * 100) '비율(%)
        
    Next
    
    Call oMat01.LoadFromDataSource

    Call MDC_Com.MDC_GF_Message("계산 완료!", "S")
    Call oForm01.Freeze(False)
    
'    Set RecordSet01 = Nothing
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA030_Calculate_Header_Error:

    Call oForm01.Freeze(False)
    If ErrNum = 1 Then
    
    ElseIf ErrNum = 99 Then
        
    Else
    
        Call MDC_Com.MDC_GF_Message("PS_GA030_Calculate_Header_Error:" & Err.Number & " - " & Err.Description, "E")
        
    End If
    
'    Set RecordSet01 = Nothing

End Sub

Public Sub PS_GA030_Calculate_Line()
'******************************************************************************
'Function ID : PS_GA030_Calculate_Line()
'해당모듈    : PS_GA030
'기능        : 실적*비율=예산 (라인별 계산)
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA030_Calculate_Line_Error

    Dim loopCount As Integer
    Dim sQry      As String
    Dim ErrNum    As Integer
    
'    Dim RecordSet01 As SAPbobsCOM.Recordset
'    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim rate As Currency
    
    Call oForm01.Freeze(True)
    Call oMat01.FlushToDataSource

    For loopCount = 0 To oMat01.RowCount - 1
    
        rate = oDS_PS_GA030L.GetValue("U_ColPrc01", loopCount) * 0.01
        
        Call oDS_PS_GA030L.setValue("U_ColSum01", loopCount, oDS_PS_GA030L.GetValue("U_ColSum01", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum01", loopCount) * rate)) '1월
        Call oDS_PS_GA030L.setValue("U_ColSum02", loopCount, oDS_PS_GA030L.GetValue("U_ColSum02", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum02", loopCount) * rate)) '2월
        Call oDS_PS_GA030L.setValue("U_ColSum03", loopCount, oDS_PS_GA030L.GetValue("U_ColSum03", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum03", loopCount) * rate)) '3월
        Call oDS_PS_GA030L.setValue("U_ColSum04", loopCount, oDS_PS_GA030L.GetValue("U_ColSum04", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum04", loopCount) * rate)) '4월
        Call oDS_PS_GA030L.setValue("U_ColSum05", loopCount, oDS_PS_GA030L.GetValue("U_ColSum05", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum05", loopCount) * rate)) '5월
        Call oDS_PS_GA030L.setValue("U_ColSum06", loopCount, oDS_PS_GA030L.GetValue("U_ColSum06", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum06", loopCount) * rate)) '6월
        Call oDS_PS_GA030L.setValue("U_ColSum07", loopCount, oDS_PS_GA030L.GetValue("U_ColSum07", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum07", loopCount) * rate)) '7월
        Call oDS_PS_GA030L.setValue("U_ColSum08", loopCount, oDS_PS_GA030L.GetValue("U_ColSum08", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum08", loopCount) * rate)) '8월
        Call oDS_PS_GA030L.setValue("U_ColSum09", loopCount, oDS_PS_GA030L.GetValue("U_ColSum09", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum09", loopCount) * rate)) '9월
        Call oDS_PS_GA030L.setValue("U_ColSum10", loopCount, oDS_PS_GA030L.GetValue("U_ColSum10", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum10", loopCount) * rate)) '10월
        Call oDS_PS_GA030L.setValue("U_ColSum11", loopCount, oDS_PS_GA030L.GetValue("U_ColSum11", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum11", loopCount) * rate)) '11월
        Call oDS_PS_GA030L.setValue("U_ColSum12", loopCount, oDS_PS_GA030L.GetValue("U_ColSum12", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum12", loopCount) * rate)) '12월
        Call oDS_PS_GA030L.setValue("U_ColSum13", loopCount, oDS_PS_GA030L.GetValue("U_ColSum13", loopCount) + (oDS_PS_GA030L.GetValue("U_ColSum13", loopCount) * rate)) '계
        Call oDS_PS_GA030L.setValue("U_ColPrc01", loopCount, oDS_PS_GA030L.GetValue("U_ColPrc01", loopCount)) '비율
        
    Next
    
    Call oMat01.LoadFromDataSource

    Call MDC_Com.MDC_GF_Message("개별 계산 완료!", "S")
    Call oForm01.Freeze(False)
    
'    Set RecordSet01 = Nothing
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA030_Calculate_Line_Error:

    Call oForm01.Freeze(False)
    If ErrNum = 1 Then
    
    ElseIf ErrNum = 99 Then
        
    Else
    
        Call MDC_Com.MDC_GF_Message("PS_GA030_Calculate_Line_Error:" & Err.Number & " - " & Err.Description, "E")
        
    End If
    
'    Set RecordSet01 = Nothing

End Sub


Public Sub PS_GA030_AddData()
'******************************************************************************
'Function ID : PS_GA030_AddData()
'해당모듈    : PS_GA030
'기능        : 데이터 INSERT
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA030_AddData_Error

    Dim loopCount As Integer
    Dim sQry      As String
    Dim ErrNum    As Integer
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim StdYear    As String '년도
    Dim BPLID      As String '사업장
    Dim AcctCode01 As String '계정코드
    Dim AcctName01 As String '계정명
    Dim AcctCode02 As String '계정과목코드
    Dim AcctName02 As String '계정과목명
    Dim AcctCode03 As String '세부계정과목코드
    Dim AcctName03 As String '세부계정과목명
    Dim Month01    As Currency '01월
    Dim Month02    As Currency '02월
    Dim Month03    As Currency '03월
    Dim Month04    As Currency '04월
    Dim Month05    As Currency '05월
    Dim Month06    As Currency '06월
    Dim Month07    As Currency '07월
    Dim Month08    As Currency '08월
    Dim Month09    As Currency '09월
    Dim Month10    As Currency '10월
    Dim Month11    As Currency '11월
    Dim Month12    As Currency '12월
    Dim Total      As Currency '계
    Dim rate       As Currency '비율
    
'    Rate = oForm01.Items("Rate").Specific.VALUE
    
    Call oMat01.FlushToDataSource
    
    For loopCount = 0 To oMat01.RowCount - 1
    
        StdYear = Trim(oDS_PS_GA030L.GetValue("U_ColReg01", loopCount)) '년도
        BPLID = Trim(oDS_PS_GA030L.GetValue("U_ColReg02", loopCount)) '사업장
        AcctCode01 = Trim(oDS_PS_GA030L.GetValue("U_ColReg03", loopCount)) '계정코드
        AcctName01 = Trim(oDS_PS_GA030L.GetValue("U_ColReg04", loopCount)) '계정명
        AcctCode02 = Trim(oDS_PS_GA030L.GetValue("U_ColReg05", loopCount)) '계정과목코드
        AcctName02 = Trim(oDS_PS_GA030L.GetValue("U_ColReg06", loopCount)) '계정과목명
        AcctCode03 = Trim(oDS_PS_GA030L.GetValue("U_ColReg07", loopCount)) '세부계정과목코드
        AcctName03 = Trim(oDS_PS_GA030L.GetValue("U_ColReg08", loopCount)) '세부계정과목명
        Month01 = Trim(oDS_PS_GA030L.GetValue("U_ColSum01", loopCount)) '01월
        Month02 = Trim(oDS_PS_GA030L.GetValue("U_ColSum02", loopCount)) '02월
        Month03 = Trim(oDS_PS_GA030L.GetValue("U_ColSum03", loopCount)) '03월
        Month04 = Trim(oDS_PS_GA030L.GetValue("U_ColSum04", loopCount)) '04월
        Month05 = Trim(oDS_PS_GA030L.GetValue("U_ColSum05", loopCount)) '05월
        Month06 = Trim(oDS_PS_GA030L.GetValue("U_ColSum06", loopCount)) '06월
        Month07 = Trim(oDS_PS_GA030L.GetValue("U_ColSum07", loopCount)) '07월
        Month08 = Trim(oDS_PS_GA030L.GetValue("U_ColSum08", loopCount)) '08월
        Month09 = Trim(oDS_PS_GA030L.GetValue("U_ColSum09", loopCount)) '09월
        Month10 = Trim(oDS_PS_GA030L.GetValue("U_ColSum10", loopCount)) '10월
        Month11 = Trim(oDS_PS_GA030L.GetValue("U_ColSum11", loopCount)) '11월
        Month12 = Trim(oDS_PS_GA030L.GetValue("U_ColSum12", loopCount)) '12월
        Total = Trim(oDS_PS_GA030L.GetValue("U_ColSum13", loopCount)) '계
        rate = Trim(oDS_PS_GA030L.GetValue("U_ColPrc01", loopCount)) '비율
    
        sQry = "            EXEC [PS_GA030_03] "
        sQry = sQry & "'" & StdYear & "'," '년도
        sQry = sQry & "'" & BPLID & "',"  '사업장
        sQry = sQry & "'" & AcctCode01 & "',"  '계정코드
        sQry = sQry & "'" & AcctName01 & "',"  '계정명
        sQry = sQry & "'" & AcctCode02 & "',"  '계정과목코드
        sQry = sQry & "'" & AcctName02 & "',"  '계정과목명
        sQry = sQry & "'" & AcctCode03 & "',"  '세부계정과목코드
        sQry = sQry & "'" & AcctName03 & "',"  '세부계정과목명
        sQry = sQry & "'" & Month01 & "',"  '01월
        sQry = sQry & "'" & Month02 & "',"  '02월
        sQry = sQry & "'" & Month03 & "',"  '03월
        sQry = sQry & "'" & Month04 & "',"  '04월
        sQry = sQry & "'" & Month05 & "',"  '05월
        sQry = sQry & "'" & Month06 & "',"  '06월
        sQry = sQry & "'" & Month07 & "',"  '07월
        sQry = sQry & "'" & Month08 & "',"  '08월
        sQry = sQry & "'" & Month09 & "',"  '09월
        sQry = sQry & "'" & Month10 & "',"  '10월
        sQry = sQry & "'" & Month11 & "',"  '11월
        sQry = sQry & "'" & Month12 & "',"  '12월
        sQry = sQry & "'" & Total & "',"  '계
        sQry = sQry & "'" & rate & "'"  '비율
        
        Call RecordSet01.DoQuery(sQry)
    
    Next

    MDC_Com.MDC_GF_Message "등록 완료!", "S"
    
    Set RecordSet01 = Nothing
    
    Exit Sub

PS_GA030_AddData_Error:

    If ErrNum = 1 Then
    
    ElseIf ErrNum = 99 Then
        
    Else
    
        Call MDC_Com.MDC_GF_Message("PS_GA030_AddData_Error:" & Err.Number & " - " & Err.Description, "E")
        
    End If
    
    Set RecordSet01 = Nothing
    
End Sub

Private Function PS_GA030_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PS_GA030_HeaderSpaceLineDel()
'해당모듈    : PS_GA030
'기능        : 필수입력사항 체크
'인수        : 없음
'반환값      : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA030_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm01.Items("BPLID").Specific.VALUE) = "%" '사업장
            ErrNum = 1
            GoTo PS_GA030_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZDPTCD").Specific.VALUE) = "%" '관련자팀
            ErrNum = 2
            GoTo PS_GA030_HeaderSpaceLineDel_Error
'        Case Trim(oForm01.Items("ZSECCD").Specific.VALUE) = "%" '관련자담당
'            ErrNum = 3
'            GoTo PS_GA030_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZEMPIE").Specific.VALUE) = "" '관련자사번
            ErrNum = 4
            GoTo PS_GA030_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZGODAY").Specific.VALUE) = "" '기준일자
            ErrNum = 5
            GoTo PS_GA030_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZGOTIME").Specific.VALUE) = "" '출입시간(Fr)
            ErrNum = 6
            GoTo PS_GA030_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZOUTIME").Specific.VALUE) = "" '출입시간(To)
            ErrNum = 7
            GoTo PS_GA030_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZUSRNN").Specific.VALUE) = "" '방문자성명
            ErrNum = 8
            GoTo PS_GA030_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZBUSNM").Specific.VALUE) = "" '업체명
            ErrNum = 9
            GoTo PS_GA030_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZRESID").Specific.VALUE) = "" '주소
            ErrNum = 10
            GoTo PS_GA030_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZCARJO").Specific.VALUE) = "%" '차량종류
            ErrNum = 11
            GoTo PS_GA030_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZMOKJU").Specific.VALUE) = "%" '출입목적
            ErrNum = 12
            GoTo PS_GA030_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("ZPOINT").Specific.VALUE) = "%" '장소
            ErrNum = 13
            GoTo PS_GA030_HeaderSpaceLineDel_Error
    End Select

    PS_GA030_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA030_HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("BPLID").Click(ct_Regular)
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "관련자팀은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZDPTCD").Click(ct_Regular)
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "관련자담당은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZSECCD").Click(ct_Regular)
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "관련자사번은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZEMPIE").Click(ct_Regular)
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "기준일자는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZGODAY").Click(ct_Regular)
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "출입시간(시작)은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZGOTIME").Click(ct_Regular)
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "출입시간(종료)은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZOUTIME").Click(ct_Regular)
    ElseIf ErrNum = 8 Then
        MDC_Com.MDC_GF_Message "방문자성명은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZUSRNN").Click(ct_Regular)
    ElseIf ErrNum = 9 Then
        MDC_Com.MDC_GF_Message "업체명은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZBUSNM").Click(ct_Regular)
    ElseIf ErrNum = 10 Then
        MDC_Com.MDC_GF_Message "주소는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZRESID").Click(ct_Regular)
    ElseIf ErrNum = 11 Then
        MDC_Com.MDC_GF_Message "차량종류는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZCARJO").Click(ct_Regular)
    ElseIf ErrNum = 12 Then
        MDC_Com.MDC_GF_Message "출입목적은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZMOKJU").Click(ct_Regular)
    ElseIf ErrNum = 13 Then
        MDC_Com.MDC_GF_Message "장소는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("ZPOINT").Click(ct_Regular)
    End If
    PS_GA030_HeaderSpaceLineDel = False
End Function

Private Sub PS_GA030_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_GA030_FlushToItemValue_Error

    Dim i        As Integer
    Dim ErrNum   As Integer
    Dim sQry     As String
    Dim ItemCode As String
    Dim Qty      As Integer
    Dim Total    As String
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
    
        Case "Mat01"
            
            Call oMat01.FlushToDataSource
            
            If oCol = "Month01" Or oCol = "Month02" Or oCol = "Month03" Or oCol = "Month04" Or oCol = "Month05" Or oCol = "Month06" Or oCol = "Month07" Or oCol = "Month08" Or oCol = "Month09" Or oCol = "Month10" Or oCol = "Month11" Or oCol = "Month12" Then
            
                Total = CDbl(oDS_PS_GA030L.GetValue("U_ColSum01", oRow - 1)) + CDbl(oDS_PS_GA030L.GetValue("U_ColSum02", oRow - 1)) + CDbl(oDS_PS_GA030L.GetValue("U_ColSum03", oRow - 1)) + CDbl(oDS_PS_GA030L.GetValue("U_ColSum04", oRow - 1)) + CDbl(oDS_PS_GA030L.GetValue("U_ColSum05", oRow - 1)) + CDbl(oDS_PS_GA030L.GetValue("U_ColSum06", oRow - 1)) + CDbl(oDS_PS_GA030L.GetValue("U_ColSum07", oRow - 1)) + CDbl(oDS_PS_GA030L.GetValue("U_ColSum08", oRow - 1)) + CDbl(oDS_PS_GA030L.GetValue("U_ColSum09", oRow - 1)) + CDbl(oDS_PS_GA030L.GetValue("U_ColSum10", oRow - 1)) + CDbl(oDS_PS_GA030L.GetValue("U_ColSum11", oRow - 1)) + CDbl(oDS_PS_GA030L.GetValue("U_ColSum12", oRow - 1))
                Call oDS_PS_GA030L.setValue("U_ColSum13", oRow - 1, Total)
                Call oMat01.LoadFromDataSource
                Call oMat01.Columns(oCol).Cells(oRow).Click
                
            End If
            
            Call oMat01.AutoResizeColumns

        Case "AcctCode01"
            
            oForm01.Items("AcctName01").Specific.VALUE = MDC_GetData.Get_ReData("AcctName", "AcctCode", "[OACT]", "'" & oForm01.Items("AcctCode01").Specific.VALUE & "'") '계정

        Case "AcctCode02"
            
            oForm01.Items("AcctName02").Specific.VALUE = MDC_GetData.Get_ReData("AcctName", "AcctCode", "[OACT]", "'" & oForm01.Items("AcctCode02").Specific.VALUE & "'") '계정과목
            
        Case "AcctCode03"
            
            oForm01.Items("AcctName03").Specific.VALUE = MDC_GetData.Get_ReData("AcctName", "AcctCode", "[OACT]", "'" & oForm01.Items("AcctCode03").Specific.VALUE & "'") '세부계정과목
        
    End Select

    Set oRecordSet01 = Nothing

    Exit Sub
PS_GA030_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("에러 메시지", "E")
    Else
        Call MDC_Com.MDC_GF_Message("PS_GA030_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

'/폼의 아이템 사용지정
Sub PS_GA030_FormItemEnabled()
On Error GoTo PS_GA030_FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA030_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_GA030_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/아이템 변경 이벤트
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error

    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            Call Raise_EVENT_ITEM_PRESSED(FormUID, pval, BubbleEvent)
        Case et_KEY_DOWN: '//2
            Call Raise_EVENT_KEY_DOWN(FormUID, pval, BubbleEvent)
        Case et_COMBO_SELECT: '//5
            Call Raise_EVENT_COMBO_SELECT(FormUID, pval, BubbleEvent)
        Case et_CLICK: '//6
            Call Raise_EVENT_CLICK(FormUID, pval, BubbleEvent)
        Case et_DOUBLE_CLICK: '//7
            Call Raise_EVENT_DOUBLE_CLICK(FormUID, pval, BubbleEvent)
        Case et_MATRIX_LINK_PRESSED '//8
            Call Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, pval, BubbleEvent)
        Case et_VALIDATE: '//10
            Call Raise_EVENT_VALIDATE(FormUID, pval, BubbleEvent)
        Case et_MATRIX_LOAD: '//11
            Call Raise_EVENT_MATRIX_LOAD(FormUID, pval, BubbleEvent)
        Case et_FORM_ACTIVATE: '//18
            '//et_FORM_ACTIVATE
        Case et_FORM_DEACTIVATE: '//19
            '//et_FORM_DEACTIVATE
        Case et_FORM_RESIZE '//20
            Call Raise_EVENT_RESIZE(FormUID, pval, BubbleEvent)
        Case et_CHOOSE_FROM_LIST '//27
            Call Raise_EVENT_CHOOSE_FROM_LIST(FormUID, pval, BubbleEvent)
        Case et_GOT_FOCUS: '//3
            Call Raise_EVENT_GOT_FOCUS(FormUID, pval, BubbleEvent)
        Case et_LOST_FOCUS: '//4
            '//et_LOST_FOCUS
        Case et_FORM_UNLOAD: '//17
            Call Raise_EVENT_FORM_UNLOAD(FormUID, pval, BubbleEvent)
    End Select
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_ItemEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
            Case "1282": '추가
                '/추가버튼 클릭시 메트릭스 insertrow
                
'                Call PS_GA030_FormReset
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
                oForm01.Mode = fm_ADD_MODE
                BubbleEvent = False
                Call PS_GA030_LoadCaption
                
                'oForm01.Items("GCode").Click ct_Regular
                
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call PS_GA030_Add_MatrixRow(oMat01.VisualRowCount)
                
            
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                '//Call PS_GA030_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_GA030H.Clear

'                Call PS_GA030_LoadCaption
'                Call PS_GA030_FormItemEnabled
                '//Call PS_GA030_FormItemEnabled '//UDO방식
                '//Call PS_GA030_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_GA030_FormItemEnabled
                
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm01.Freeze(True)
                Call oDS_PS_GA030L.RemoveRecord(oDS_PS_GA030L.Size - 1)
                Call oMat01.LoadFromDataSource
                Call oForm01.Freeze(False)
                
        End Select
    End If
    Exit Sub
Raise_MenuEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormDataEvent_Error
    If (BusinessObjectInfo.BeforeAction = True) Then '//BeforeAction = True
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD: '//33
            Case et_FORM_DATA_ADD: '//34
            Case et_FORM_DATA_UPDATE: '//35
            Case et_FORM_DATA_DELETE: '//36
        End Select
    ElseIf (BusinessObjectInfo.BeforeAction = False) Then '//BeforeAction = False
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD: '//33
            Case et_FORM_DATA_ADD: '//34
            Case et_FORM_DATA_UPDATE: '//35
            Case et_FORM_DATA_DELETE: '//36
        End Select
    End If
    Exit Sub
Raise_FormDataEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error

    If pval.BeforeAction = True Then
    ElseIf pval.BeforeAction = False Then
    End If
    
    If pval.ItemUID = "Mat01" Then
        If pval.Row > 0 Then
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = pval.ColUID
            oLastColRow01 = pval.Row
        End If
    Else
        oLastItemUID01 = pval.ItemUID
        oLastColUID01 = ""
        oLastColRow01 = 0
    End If
    
    Exit Sub
    
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ITEM_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ITEM_PRESSED_Error

    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "PS_GA030" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnAdd" Then '추가(수정) 버튼클릭
        
            If oForm01.Mode = fm_ADD_MODE Then
            
'                If PS_GA030_HeaderSpaceLineDel = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
                
                Call PS_GA030_AddData

'                Call PS_GA030_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_GA030_LoadCaption
'                Call PS_GA030_MTX01
                
                oLast_Mode = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            
'                If PS_GA030_HeaderSpaceLineDel = False Then
'                    BubbleEvent = False
'                    Exit Sub
'                End If
            
                '쿼리에서 로직 처리
                '여러 행을 동시에 처리하므로 DocEntry별로 행이 존재하면 UPDATE, 존재하지 않으면 INSERT
                Call PS_GA030_AddData
                    
'                Call PS_GA030_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_GA030_LoadCaption
'                Call PS_GA030_MTX01
                
'                oForm01.Items("GCode").Click ct_Regular
            End If
            
        ElseIf pval.ItemUID = "BtnSearch1" Then '실적조회
        
            Call PS_GA030_FormReset
            oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
            Call PS_GA030_LoadCaption
            Call PS_GA030_MTX01("1")
            
        ElseIf pval.ItemUID = "BtnSearch2" Then '계획조회

'            Call PS_GA030_FormReset
            oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
            Call PS_GA030_LoadCaption
            Call PS_GA030_MTX01("2")
            
        ElseIf pval.ItemUID = "BtnCal1" Then '전체 계산
        
            Call PS_GA030_Calculate_Header
             
        ElseIf pval.ItemUID = "BtnCal2" Then '라인별 계산
        
            Call PS_GA030_Calculate_Line
        
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_GA030" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
    End If
    
    Exit Sub
    
Raise_EVENT_ITEM_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ITEM_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_KEY_DOWN(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_KEY_DOWN_Error
    If pval.BeforeAction = True Then
    
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "AcctCode01", "") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "AcctCode02", "") '//사용자값활성
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "AcctCode03", "") '//사용자값활성
        
'        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
    
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CLICK_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
        If pval.ItemUID = "Mat01" Then
        
            If pval.Row > 0 Then
            
                Call oMat01.SelectRow(pval.Row, True, False)
                
            End If
            
        End If
    
    End If
    
    Exit Sub
    
Raise_EVENT_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
        Call PS_GA030_FlushToItemValue(pval.ItemUID)
    
    End If

    Exit Sub
    
Raise_EVENT_COMBO_SELECT_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_COMBO_SELECT_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_DOUBLE_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_DOUBLE_CLICK_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
    
Raise_EVENT_DOUBLE_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_DOUBLE_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LINK_PRESSED(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LINK_PRESSED_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
    
Raise_EVENT_MATRIX_LINK_PRESSED_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LINK_PRESSED_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_VALIDATE(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_VALIDATE_Error

    Call oForm01.Freeze(True)
    
    If pval.BeforeAction = True Then
    
        If pval.ItemChanged = True Then
        
            If pval.ItemUID = "Mat01" Then
            
                If pval.ColUID = "Month01" Or pval.ColUID = "Month02" Or pval.ColUID = "Month03" Or pval.ColUID = "Month04" Or pval.ColUID = "Month05" Or pval.ColUID = "Month06" Or pval.ColUID = "Month07" Or pval.ColUID = "Month08" Or pval.ColUID = "Month09" Or pval.ColUID = "Month10" Or pval.ColUID = "Month11" Or pval.ColUID = "Month12" Then
            
                    Call PS_GA030_FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                
                End If
            
            Else
    
                Call PS_GA030_FlushToItemValue(pval.ItemUID)
            
            End If
                
        End If
        
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
'            oForm01.Update
        
    ElseIf pval.BeforeAction = False Then

    End If
    
    Call oForm01.Freeze(False)
    
    Exit Sub
    
Raise_EVENT_VALIDATE_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_VALIDATE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_MATRIX_LOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_MATRIX_LOAD_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        Call PS_GA030_FormItemEnabled
        '//Call PS_GA030_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    
    Exit Sub
    
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error

    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
    
        Call PS_GA030_FormResize
        
    End If
    
    Exit Sub
    
Raise_EVENT_RESIZE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_RESIZE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CHOOSE_FROM_LIST(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CHOOSE_FROM_LIST_Error

    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
'        If (pval.ItemUID = "ItemCode") Then
'            Dim oDataTable01 As SAPbouiCOM.DataTable
'            Set oDataTable01 = pval.SelectedObjects
'            oForm01.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
'            Set oDataTable01 = Nothing
'        End If
'        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_GA030H", "U_CardCode,U_CardName")
'        End If
    End If
    
    Exit Sub
    
Raise_EVENT_CHOOSE_FROM_LIST_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CHOOSE_FROM_LIST_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Private Sub Raise_EVENT_GOT_FOCUS(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_GOT_FOCUS_Error

    If pval.ItemUID = "Mat01" Then
        If pval.Row > 0 Then
            oLastItemUID01 = pval.ItemUID
            oLastColUID01 = pval.ColUID
            oLastColRow01 = pval.Row
        End If
    Else
        oLastItemUID01 = pval.ItemUID
        oLastColUID01 = ""
        oLastColRow01 = 0
    End If
    
    Exit Sub
    
Raise_EVENT_GOT_FOCUS_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_GOT_FOCUS_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_FORM_UNLOAD(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_FORM_UNLOAD_Error

    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
        RemoveForms oFormUniqueID01
        Set oForm01 = Nothing
        Set oMat01 = Nothing
    End If
    
    Exit Sub
    
Raise_EVENT_FORM_UNLOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_FORM_UNLOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ROW_DELETE_Error

    Dim i As Long
    
    If (oLastColRow01 > 0) Then
        If pval.BeforeAction = True Then
'            If (PS_GA030_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
            
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_GA030L.RemoveRecord(oDS_PS_GA030L.Size - 1)
            oMat01.LoadFromDataSource
            
'            If oMat01.RowCount = 0 Then
'                Call PS_GA030_Add_MatrixRow(0)
'            Else
'                If Trim(oDS_PS_GA030L.GetValue("U_ColReg01", oMat01.RowCount - 1)) <> "" Then
'                    Call PS_GA030_Add_MatrixRow(oMat01.RowCount)
'                End If
'            End If

        End If
    End If
    
    Exit Sub
    
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_GA030_CreateItems() As Boolean
On Error GoTo PS_GA030_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_GA030L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '사업장_S
    Call oForm01.DataSources.UserDataSources.Add("BPLID", dt_SHORT_TEXT, 10)
    Call oForm01.Items("BPLID").Specific.DataBind.SetBound(True, "", "BPLID")
    '사업장_E
    
    '기준년도
    Call oForm01.DataSources.UserDataSources.Add("StdYear", dt_SHORT_TEXT, 4)
    Call oForm01.Items("StdYear").Specific.DataBind.SetBound(True, "", "StdYear")
    
    '계정
    Call oForm01.DataSources.UserDataSources.Add("AcctCode01", dt_SHORT_TEXT, 20)
    Call oForm01.Items("AcctCode01").Specific.DataBind.SetBound(True, "", "AcctCode01")
    
    '계정명
    Call oForm01.DataSources.UserDataSources.Add("AcctName01", dt_SHORT_TEXT, 50)
    Call oForm01.Items("AcctName01").Specific.DataBind.SetBound(True, "", "AcctName01")
    
    '계정과목
    Call oForm01.DataSources.UserDataSources.Add("AcctCode02", dt_SHORT_TEXT, 20)
    Call oForm01.Items("AcctCode02").Specific.DataBind.SetBound(True, "", "AcctCode02")
    
    '계정과목명
    Call oForm01.DataSources.UserDataSources.Add("AcctName02", dt_SHORT_TEXT, 50)
    Call oForm01.Items("AcctName02").Specific.DataBind.SetBound(True, "", "AcctName02")
    
    '세부계정과목
    Call oForm01.DataSources.UserDataSources.Add("AcctCode03", dt_SHORT_TEXT, 20)
    Call oForm01.Items("AcctCode03").Specific.DataBind.SetBound(True, "", "AcctCode03")
    
    '세부계정과목명
    Call oForm01.DataSources.UserDataSources.Add("AcctName03", dt_SHORT_TEXT, 50)
    Call oForm01.Items("AcctName03").Specific.DataBind.SetBound(True, "", "AcctName03")
    
    '비율
    Call oForm01.DataSources.UserDataSources.Add("Rate", dt_PRICE)
    Call oForm01.Items("Rate").Specific.DataBind.SetBound(True, "", "Rate")
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Function
    
PS_GA030_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA030_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_GA030_ComboBox_Setting()
On Error GoTo PS_GA030_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)

    '//////////기본정보//////////S
    '사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_GA030_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA030_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA030_Initial_Setting()
On Error GoTo PS_GA030_Initial_Setting_Error
    '사업장 사용자의 소속 사업장 선택
    Call oForm01.Items("BPLID").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
 
    '날짜 설정
'    oForm01.Items("ToDt").Specific.VALUE = Format(Date, "YYYY") + "1231"
'    oForm01.Items("FrDt").Specific.VALUE = Format(Now, "YYYY") + "0101"
    oForm01.Items("StdYear").Specific.VALUE = Format(Now, "YYYY")

    '년도, 사업장 코드는 Hidden 처리
    oMat01.Columns("StdYear").Visible = False
    oMat01.Columns("BPLID").Visible = False

    Exit Sub
PS_GA030_Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA030_Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA030_CF_ChooseFromList()
On Error GoTo PS_GA030_CF_ChooseFromList_Error

    Exit Sub
    
PS_GA030_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA030_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA030_EnableMenus()
On Error GoTo PS_GA030_EnableMenus_Error
    
    Exit Sub
    
PS_GA030_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA030_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA030_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_GA030_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_GA030_FormItemEnabled
        '//Call PS_GA030_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_GA030_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_GA030_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA030_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA030_FormResize()
On Error GoTo PS_GA030_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
    
PS_GA030_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA030_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_GA030_FormReset()
'******************************************************************************
'Function ID : PS_GA030_FormReset()
'해당모듈    : PS_GA030
'기능        : 화면 초기화
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_GA030_FormReset_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
'    '관리번호
'    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_GA030H]"
'    Call RecordSet01.DoQuery(sQry)
'
'    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
'        Call oDS_PS_GA030H.setValue("DocEntry", 0, 1)
'    Else
'        Call oDS_PS_GA030H.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
'    End If
    
'    Call oDS_PS_GA030H.setValue("DocEntry", 0, "") '관리번호
'    Call oDS_PS_GA030H.setValue("U_BPLId", 0, MDC_PS_Common.User_BPLId) '사업장
'    Call oDS_PS_GA030H.setValue("U_ZGODAY", 0, Format(Date, "YYYYMMDD")) '기준일자
'    Call oDS_PS_GA030H.setValue("U_ZEMPIE", 0, "") '관련자사번
'    Call oDS_PS_GA030H.setValue("U_ZEMPNM", 0, "") '관련자성명
'    Call PS_GA030_FlushToItemValue("BPLID") '관련자사업장에 따른 팀 세팅
'    Call oDS_PS_GA030H.setValue("U_ZDPTCD", 0, "%") '관련자팀
'    Call PS_GA030_FlushToItemValue("ZDPTCD") '관련자팀에 따른 담당 세팅
'    Call oDS_PS_GA030H.setValue("U_ZSECCD", 0, "%") '관련자담당
'    Call oDS_PS_GA030H.setValue("U_ZPOINT", 0, "%") '장소
'    Call oDS_PS_GA030H.setValue("U_ZUSRNN", 0, "") '방문자성명
'    Call oDS_PS_GA030H.setValue("U_ZBIRDT", 0, "") '생년월일
'    Call oDS_PS_GA030H.setValue("U_ZBUSNM", 0, "") '업체명
'    Call oDS_PS_GA030H.setValue("U_ZRESID", 0, "") '주소
'    Call oDS_PS_GA030H.setValue("U_ZGOTIME", 0, "") '출입시간(Fr)
'    Call oDS_PS_GA030H.setValue("U_ZOUTIME", 0, "") '출입시간(To)
'    Call oDS_PS_GA030H.setValue("U_ZCARJO", 0, "%") '차량종류
'    Call oDS_PS_GA030H.setValue("U_ZCARNO", 0, "") '차량번호
'    Call oDS_PS_GA030H.setValue("U_ZMOKJU", 0, "%") '출입목적
'    Call oDS_PS_GA030H.setValue("U_ZSPCIU", 0, "") '방문객비고
    
    gDeleteRowCount = 0
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)

    Exit Sub

PS_GA030_FormReset_Error:
    Call oForm01.Freeze(False)
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_GA030_FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub




