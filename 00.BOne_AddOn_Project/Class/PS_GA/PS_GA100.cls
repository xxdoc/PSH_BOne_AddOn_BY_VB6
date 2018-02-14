VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_GA100"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_GA100.cls
'//  Module : 총무관리>안전환경관리>무재해목표관리>무재해목표등록
'//  Desc : 무재해 목표 정보 입력
'//  FormType : PS_GA100
'//  Create Date(Start) : 2012.09.12
'//  Create Date(End) : 2012.09.12
'//  Creator : Song Myoung gyu
'//  Modified Date :
'//  Modifier :
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_GA100H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_GA100L As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'//사용자구조체
Private Type ItemInformations
    ItemCode As String
    LotNo As String
    Quantity As Long
    OPORNo As Long
    POR1No As Long
    Check As Boolean
    OPDNNo As Long
    PDN1No As Long
End Type

Private oLast_Mode&

Private ItemInformation() As ItemInformations
Private ItemInformationCount As Long

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_GA100.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_GA100_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
     
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Call PS_GA100_CreateItems
    Call PS_GA100_ComboBox_Setting
    Call PS_GA100_CF_ChooseFromList
    Call PS_GA100_EnableMenus
    Call PS_GA100_SetDocument(oFromDocEntry01)
    Call PS_GA100_FormResize

'    Call PS_GA100_Add_MatrixRow(0, True)
    Call PS_GA100_LoadCaption
    Call PS_GA100_FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1285"), False        '// 복원
    oForm01.EnableMenu ("1284"), False         '// 취소
    oForm01.EnableMenu ("1293"), False         '// 행삭제
    oForm01.EnableMenu ("1281"), False
    oForm01.EnableMenu ("1282"), True

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
        
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_GA100H]"
    Call RecordSet01.DoQuery(sQry)
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        Call oDS_PS_GA100H.setValue("DocEntry", 0, 1)
    Else
        Call oDS_PS_GA100H.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
    End If
        
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing

     'oForm01.Items("GCode").Click ct_Regular
    Exit Sub
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA100_LoadCaption()
'******************************************************************************
'Function ID : PS_GA100_LoadCaption()
'해당모듈 : PS_GA100
'기능 : Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA100_LoadCaption_Error

    Call oForm01.Freeze(True)

    If oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("BtnAdd").Specific.Caption = "추가"
        oForm01.Items("BtnDelete").Enabled = False
'    ElseIf oForm01.Mode = fm_OK_MODE Then
'        oForm01.Items("BtnAdd").Specific.Caption = "확인"
    ElseIf oForm01.Mode = fm_UPDATE_MODE Then
        oForm01.Items("BtnAdd").Specific.Caption = "수정"
        oForm01.Items("BtnDelete").Enabled = True
    End If
    
    Call oForm01.Freeze(False)

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA100_LoadCaption_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'/메트릭스 Row추가
Sub PS_GA100_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_GA100_Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_GA100L.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_GA100L.Offset = oRow
    oDS_PS_GA100L.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA100_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PS_GA100_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_GA100_MTX01()
'******************************************************************************
'Function ID : PS_GA100_MTX01()
'해당모듈 : PS_GA100
'기능 : 데이터 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA100_MTX01_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim STrgtCls As String '목표구분
    Dim sBPLId As String '사업장
    Dim STeamCd1 As String '팀1
    Dim SRspCd1 As String '담당1
    Dim STeamCd2 As String '팀2
    Dim SRspCd2 As String '담당2
    Dim SCmYN As String '달성여부
    
    STrgtCls = IIf(Trim(oForm01.Items("STrgtCls").Specific.VALUE) = "%", "", Trim(oForm01.Items("STrgtCls").Specific.VALUE))
    sBPLId = IIf(Trim(oForm01.Items("SBPLId").Specific.VALUE) = "%", "", Trim(oForm01.Items("SBPLId").Specific.VALUE))
    STeamCd1 = Trim(oForm01.Items("STeamCd1").Specific.VALUE)
    SRspCd1 = Trim(oForm01.Items("SRspCd1").Specific.VALUE)
    STeamCd2 = Trim(oForm01.Items("STeamCd2").Specific.VALUE)
    SRspCd2 = Trim(oForm01.Items("SRspCd2").Specific.VALUE)
    SCmYN = IIf(Trim(oForm01.Items("SCmYN").Specific.VALUE) = "%", "", Trim(oForm01.Items("SCmYN").Specific.VALUE))
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "EXEC [PS_GA100_01] '" & STrgtCls & "','" & sBPLId & "','" & STeamCd1 & "','" & SRspCd1 & "','" & STeamCd2 & "','" & SRspCd2 & "','" & SCmYN & "'"
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_GA100L.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
'        Call PS_GA100_Add_MatrixRow(0, True)
        Call PS_GA100_LoadCaption
        
        GoTo PS_GA100_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_GA100L.Size Then
            oDS_PS_GA100L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_GA100L.Offset = i
        
        Call oDS_PS_GA100L.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_GA100L.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("DocEntry").VALUE))
        Call oDS_PS_GA100L.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("TrgtCls").VALUE))
        Call oDS_PS_GA100L.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("BPLId").VALUE))
        Call oDS_PS_GA100L.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("TeamCd1").VALUE))
        Call oDS_PS_GA100L.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("TeamNm1").VALUE))
        Call oDS_PS_GA100L.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("RspCd1").VALUE))
        Call oDS_PS_GA100L.setValue("U_ColReg07", i, Trim(oRecordSet01.Fields("RspNm1").VALUE))
        Call oDS_PS_GA100L.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("TeamCd2").VALUE))
        Call oDS_PS_GA100L.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("TeamNm2").VALUE))
        Call oDS_PS_GA100L.setValue("U_ColReg10", i, Trim(oRecordSet01.Fields("RspCd2").VALUE))
        Call oDS_PS_GA100L.setValue("U_ColReg11", i, Trim(oRecordSet01.Fields("RspNm2").VALUE))
        Call oDS_PS_GA100L.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("TrgtNod").VALUE))
        Call oDS_PS_GA100L.setValue("U_ColReg12", i, Trim(oRecordSet01.Fields("Grd").VALUE))
        Call oDS_PS_GA100L.setValue("U_ColSum02", i, Trim(oRecordSet01.Fields("TrgtMtp").VALUE))
        Call oDS_PS_GA100L.setValue("U_ColDt01", i, Format(Trim(oRecordSet01.Fields("StrDt").VALUE), "yyyyMMdd"))
        Call oDS_PS_GA100L.setValue("U_ColSum03", i, Trim(oRecordSet01.Fields("StrPrns").VALUE))
        Call oDS_PS_GA100L.setValue("U_ColDt02", i, Format(Trim(oRecordSet01.Fields("CmDt").VALUE), "yyyyMMdd"))
        Call oDS_PS_GA100L.setValue("U_ColReg13", i, Trim(oRecordSet01.Fields("CmYN").VALUE))
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
        
    Next
    
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop
    Call oForm01.Freeze(False)
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA100_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_GA100_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Sub PS_GA100_DeleteData()
'******************************************************************************
'Function ID : PS_GA100_DeleteData()
'해당모듈 : PS_GA100
'기능 : 기본정보 삭제
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA100_DeleteData_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As String
    
    If oForm01.Mode = fm_UPDATE_MODE Then
    
        DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
        
        sQry = "SELECT COUNT(*) FROM [@PS_GA100H] WHERE DocEntry = '" & DocEntry & "'"
        Call oRecordSet01.DoQuery(sQry)
        
        If (oRecordSet01.RecordCount = 0) Then
            ErrNum = 1
            oForm01.Mode = fm_ADD_MODE
            GoTo PS_GA100_DeleteData_Error
        Else
            sQry = "DELETE FROM [@PS_GA100H] WHERE DocEntry = '" & DocEntry & "'"
            Call oRecordSet01.DoQuery(sQry)
        End If
    End If
    
    Call MDC_Com.MDC_GF_Message("삭제 완료!", "S")
    
'    Call PS_GA100_FormReset
    
'    oForm01.Mode = fm_ADD_MODE
    
'    Call oForm01.Items("BtnSearch").Click(ct_Regular)
    
'    oMat01.Clear
'    oMat01.FlushToDataSource
'    oMat01.LoadFromDataSource
'    Call PS_GA100_Add_MatrixRow(0, True)
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA100_DeleteData_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("삭제대상이 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_GA100_DeleteData_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PS_GA100_UpdateData() As Boolean
'******************************************************************************
'Function ID : PS_GA100_UpdateData()
'해당모듈 : PS_GA100
'기능 : 기본정보를 수정
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA100_UpdateData_Error

    Dim i As Integer
    Dim j As Integer
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As Integer
    Dim TrgtCls As String '목표구분
    Dim BPLId As String '사업장
    Dim TeamCd1 As String '팀1
    Dim TeamNm1 As String '팀명1
    Dim RspCd1 As String '담당1
    Dim RspNm1 As String '담당명1
    Dim TeamCd2 As String '팀2
    Dim TeamNm2 As String '팀명2
    Dim RspCd2 As String '담당2
    Dim RspNm2 As String '담당명2
    Dim TrgtNod As String '목표일수
    Dim Grd As String '설정등급
    Dim TrgtMtp As String '목표배수
    Dim StrDt As String '재개시일자
    Dim StrPrns As String '재개시인원
    Dim CmDt As String '달성예정일
    Dim CmYN As String '달성여부

    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    TrgtCls = Trim(oForm01.Items("TrgtCls").Specific.VALUE) '목표구분
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE) '사업장
    TeamCd1 = Trim(oForm01.Items("TeamCd1").Specific.VALUE) '팀1
    TeamNm1 = Trim(oForm01.Items("TeamNm1").Specific.VALUE) '팀명1
    RspCd1 = Trim(oForm01.Items("RspCd1").Specific.VALUE) '담당1
    RspNm1 = Trim(oForm01.Items("RspNm1").Specific.VALUE) '담당명1
    TeamCd2 = Trim(oForm01.Items("TeamCd2").Specific.VALUE) '팀2
    TeamNm2 = Trim(oForm01.Items("TeamNm2").Specific.VALUE) '팀명2
    RspCd2 = Trim(oForm01.Items("RspCd2").Specific.VALUE) '담당2
    RspNm2 = Trim(oForm01.Items("RspNm2").Specific.VALUE) '담당명2
    TrgtNod = Trim(oForm01.Items("TrgtNod").Specific.VALUE) '목표일수
    Grd = Trim(oForm01.Items("Grd").Specific.VALUE) '설정등급
    TrgtMtp = Trim(oForm01.Items("TrgtMtp").Specific.VALUE) '목표배수
    StrDt = Trim(oForm01.Items("StrDt").Specific.VALUE) '재개시일자
    StrPrns = Trim(oForm01.Items("StrPrns").Specific.VALUE) '재개시인원
    CmDt = Trim(oForm01.Items("CmDt").Specific.VALUE) '달성예정일
    CmYN = Trim(oForm01.Items("CmYN").Specific.VALUE) '달성여부
    
    If Trim(DocEntry) = "" Then
        MDC_Com.MDC_GF_Message "수정할 항목이 없습니다. 수정하실려면 항목을 선택하세요!", "E"
        PS_GA100_UpdateData = False
        Exit Function
    End If
    
    sQry = "           UPDATE   [@PS_GA100H]"
    sQry = sQry & " SET         U_TrgtCls = '" & TrgtCls & "',"
    sQry = sQry & "              U_BPLId = '" & BPLId & "'," '사업장
    sQry = sQry & "              U_TeamCd1 = '" & TeamCd1 & "'," '팀1
    sQry = sQry & "              U_TeamNm1 = '" & TeamNm1 & "',"  '팀명1
    sQry = sQry & "              U_RspCd1 = '" & RspCd1 & "',"  '담당1
    sQry = sQry & "              U_RspNm1 = '" & RspNm1 & "',"  '담당명1
    sQry = sQry & "              U_TeamCd2 = '" & TeamCd2 & "',"  '팀2
    sQry = sQry & "              U_TeamNm2 = '" & TeamNm2 & "',"  '팀명2
    sQry = sQry & "              U_RspCd2 = '" & RspCd2 & "',"  '담당2
    sQry = sQry & "              U_RspNm2 = '" & RspNm2 & "',"  '담당명2
    sQry = sQry & "              U_TrgtNod = '" & TrgtNod & "',"  '목표일수
    sQry = sQry & "              U_Grd = '" & Grd & "',"  '설정등급
    sQry = sQry & "              U_TrgtMtp = '" & TrgtMtp & "',"  '목표배수
    sQry = sQry & "              U_StrDt = '" & StrDt & "',"  '재개시일자
    sQry = sQry & "              U_StrPrns = '" & StrPrns & "',"  '재개시인원
    sQry = sQry & "              U_CmDt = '" & CmDt & "',"  '달성예정일
    sQry = sQry & "              U_CmYN = '" & CmYN & "'"  '달성여부
    sQry = sQry & " WHERE    DocEntry = '" & DocEntry & "'"
    
    Call RecordSet01.DoQuery(sQry)
    
    MDC_Com.MDC_GF_Message "수정 완료!", "S"
    
    Set RecordSet01 = Nothing
    PS_GA100_UpdateData = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA100_UpdateData_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_GA100_UpdateData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Public Function PS_GA100_AddData() As Boolean
'******************************************************************************
'Function ID : PS_GA100_AddData()
'해당모듈 : PS_GA100
'기능 : 데이터 INSERT
'인수 : 없음
'반환값 : 성공여부
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA100_AddData_Error

    Dim i As Integer
    Dim sQry As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim RecordSet02 As SAPbobsCOM.Recordset
    Set RecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocEntry As Integer
    Dim TrgtCls As String '목표구분
    Dim BPLId As String '사업장
    Dim TeamCd1 As String '팀1
    Dim TeamNm1 As String '팀명1
    Dim RspCd1 As String '담당1
    Dim RspNm1 As String '담당명1
    Dim TeamCd2 As String '팀2
    Dim TeamNm2 As String '팀명2
    Dim RspCd2 As String '담당2
    Dim RspNm2 As String '담당명2
    Dim TrgtNod As String '목표일수
    Dim Grd As String '설정등급
    Dim TrgtMtp As String '목표배수
    Dim StrDt As String '재개시일자
    Dim StrPrns As String '재개시인원
    Dim CmDt As String '달성예정일
    Dim CmYN As String '달성여부

'    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    TrgtCls = Trim(oForm01.Items("TrgtCls").Specific.VALUE) '목표구분
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE) '사업장
    TeamCd1 = Trim(oForm01.Items("TeamCd1").Specific.VALUE) '팀1
    TeamNm1 = Trim(oForm01.Items("TeamNm1").Specific.VALUE) '팀명1
    RspCd1 = Trim(oForm01.Items("RspCd1").Specific.VALUE) '담당1
    RspNm1 = Trim(oForm01.Items("RspNm1").Specific.VALUE) '담당명1
    TeamCd2 = Trim(oForm01.Items("TeamCd2").Specific.VALUE) '팀2
    TeamNm2 = Trim(oForm01.Items("TeamNm2").Specific.VALUE) '팀명2
    RspCd2 = Trim(oForm01.Items("RspCd2").Specific.VALUE) '담당2
    RspNm2 = Trim(oForm01.Items("RspNm2").Specific.VALUE) '담당명2
    TrgtNod = Trim(oForm01.Items("TrgtNod").Specific.VALUE) '목표일수
    Grd = Trim(oForm01.Items("Grd").Specific.VALUE) '설정등급
    TrgtMtp = Trim(oForm01.Items("TrgtMtp").Specific.VALUE) '목표배수
    StrDt = Trim(oForm01.Items("StrDt").Specific.VALUE) '재개시일자
    StrPrns = Trim(oForm01.Items("StrPrns").Specific.VALUE) '재개시인원
    CmDt = Trim(oForm01.Items("CmDt").Specific.VALUE) '달성예정일
    CmYN = Trim(oForm01.Items("CmYN").Specific.VALUE) '달성여부
    
    'DocEntry는 화면상의 DocEntry가 아닌 입력 시점의 최종 DocEntry를 조회한 후 +1하여 INSERT를 해줘야 함
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM[@PS_GA100H]"
    Call RecordSet01.DoQuery(sQry)
    
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        DocEntry = 1
    Else
        DocEntry = Trim(RecordSet01.Fields(0).VALUE) + 1
    End If
    
    sQry = "           INSERT INTO [@PS_GA100H]"
    sQry = sQry & " ("
    sQry = sQry & "     DocEntry,"
    sQry = sQry & "     DocNum,"
    
    sQry = sQry & "     U_TrgtCls,"
    sQry = sQry & "     U_BPLId,"
    sQry = sQry & "     U_TeamCd1,"
    sQry = sQry & "     U_TeamNm1,"
    sQry = sQry & "     U_RspCd1,"
    sQry = sQry & "     U_RspNm1,"
    sQry = sQry & "     U_TeamCd2,"
    sQry = sQry & "     U_TeamNm2,"
    sQry = sQry & "     U_RspCd2,"
    sQry = sQry & "     U_RspNm2,"
    sQry = sQry & "     U_TrgtNod,"
    sQry = sQry & "     U_Grd,"
    sQry = sQry & "     U_TrgtMtp,"
    sQry = sQry & "     U_StrDt,"
    sQry = sQry & "     U_StrPrns,"
    sQry = sQry & "     U_CmDt,"
    sQry = sQry & "     U_CmYN"
    
    sQry = sQry & " )"
    sQry = sQry & " VALUES"
    sQry = sQry & " ("
    sQry = sQry & DocEntry & ","
    sQry = sQry & DocEntry & ","
    
    sQry = sQry & "'" & TrgtCls & "',"
    sQry = sQry & "'" & BPLId & "',"
    sQry = sQry & "'" & TeamCd1 & "',"
    sQry = sQry & "'" & TeamNm1 & "',"
    sQry = sQry & "'" & RspCd1 & "',"
    sQry = sQry & "'" & RspNm1 & "',"
    sQry = sQry & "'" & TeamCd2 & "',"
    sQry = sQry & "'" & TeamNm2 & "',"
    sQry = sQry & "'" & RspCd2 & "',"
    sQry = sQry & "'" & RspNm2 & "',"
    sQry = sQry & "'" & TrgtNod & "',"
    sQry = sQry & "'" & Grd & "',"
    sQry = sQry & "'" & TrgtMtp & "',"
    sQry = sQry & "'" & StrDt & "',"
    sQry = sQry & "'" & StrPrns & "',"
    sQry = sQry & "'" & CmDt & "',"
    sQry = sQry & "'" & CmYN & "'"
    
    sQry = sQry & ")"
    
    Call RecordSet02.DoQuery(sQry)
    
    MDC_Com.MDC_GF_Message "등록 완료!", "S"
    
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    PS_GA100_AddData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA100_AddData_Error:
    PS_GA100_AddData = False
    Set RecordSet01 = Nothing
    Set RecordSet02 = Nothing
    MDC_Com.MDC_GF_Message "PS_GA100_AddData_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

Private Function PS_GA100_HeaderSpaceLineDel() As Boolean
'******************************************************************************
'Function ID : PS_GA100_HeaderSpaceLineDel()
'해당모듈 : PS_GA100
'기능 : 필수입력사항 체크
'인수 : 없음
'반환값 : True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA100_HeaderSpaceLineDel_Error

    Dim ErrNum As Integer
    ErrNum = 0

    Select Case True
        Case Trim(oForm01.Items("TrgtCls").Specific.VALUE) = "%" '목표구분
            ErrNum = 1
            GoTo PS_GA100_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("TrgtNod").Specific.VALUE) = "" '목표일수
            ErrNum = 2
            GoTo PS_GA100_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("TrgtMtp").Specific.VALUE) = "" '목표배수
            ErrNum = 3
            GoTo PS_GA100_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("StrDt").Specific.VALUE) = "" '재개시일자
            ErrNum = 4
            GoTo PS_GA100_HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("CmDt").Specific.VALUE) = "" '달성예정일
            ErrNum = 5
            GoTo PS_GA100_HeaderSpaceLineDel_Error
    End Select

    PS_GA100_HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA100_HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "목표구분은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("TrgtCls").Click(ct_Regular)
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "목표일수는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("TrgtNod").Click(ct_Regular)
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "목표배수는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("TrgtMtp").Click(ct_Regular)
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "재개시일자는 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("StrDt").Click(ct_Regular)
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "달성예정일은 필수사항입니다. 확인하세요.", "E"
        Call oForm01.Items("CmDt").Click(ct_Regular)
    End If
    PS_GA100_HeaderSpaceLineDel = False
End Function

'/ 메트릭스 필수 사항 check
Private Function PS_GA100_MatrixSpaceLineDel() As Boolean
On Error GoTo PS_GA100_MatrixSpaceLineDel_Error

    Dim i As Long
    Dim ErrNum As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry As String

    Set oRecordSet01 = Nothing
    PS_GA100_MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA100_MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 사원코드가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 시간이 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 등록일자가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인의 비가동코드가 없습니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "PS_GA100_MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    PS_GA100_MatrixSpaceLineDel = False
End Function

Private Sub PS_GA100_FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo PS_GA100_FlushToItemValue_Error

    Dim i As Integer
    Dim ErrNum As Integer
    Dim sQry As String
    Dim ItemCode As String
    Dim Qty As Integer
    Dim Calculate_Weight As Double
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
'    Select Case oUID
'
'        Case "BOSHour"
'
'    End Select

   Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
PS_GA100_FlushToItemValue_Error:

    Set oRecordSet01 = Nothing

    Call MDC_Com.MDC_GF_Message("PS_GA100_FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E")

End Sub

'/폼의 아이템 사용지정
Sub PS_GA100_FormItemEnabled()
On Error GoTo PS_GA100_FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
'        oMat01.Columns("ItemCode").Cells(1).Click ct_Regular
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
'        oForm01.Items("ItemCode").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA100_FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "PS_GA100_FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
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
            Case "1281": '찾기
            Case "1282": '추가
                '/추가버튼 클릭시 메트릭스 insertrow
                
                Call PS_GA100_FormReset
                
'                oMat01.Clear
'                oMat01.FlushToDataSource
'                oMat01.LoadFromDataSource
                
                oForm01.Mode = fm_ADD_MODE
                BubbleEvent = False
                Call PS_GA100_LoadCaption
                
                'oForm01.Items("GCode").Click ct_Regular
                
                
                Exit Sub
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
                '//Call PS_GA100_FormItemEnabled '//UDO방식
            Case "1282": '추가
'                oMat01.Clear
'                oDS_PS_GA100H.Clear

'                Call PS_GA100_LoadCaption
'                Call PS_GA100_FormItemEnabled
                '//Call PS_GA100_FormItemEnabled '//UDO방식
                '//Call PS_GA100_AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                '//Call PS_GA100_FormItemEnabled
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
    
        If pval.ItemUID = "PS_GA100" Then
            If oForm01.Mode = fm_ADD_MODE Then
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            ElseIf oForm01.Mode = fm_OK_MODE Then
            End If
        End If
        
        If pval.ItemUID = "BtnAdd" Then '/추가/확인 버튼클릭
        
            If oForm01.Mode = fm_ADD_MODE Then
            
                If PS_GA100_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                If PS_GA100_AddData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If

                Call PS_GA100_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_GA100_LoadCaption
                Call PS_GA100_MTX01
                
                oLast_Mode = oForm01.Mode
                
            ElseIf oForm01.Mode = fm_UPDATE_MODE Then
            
                If PS_GA100_HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
                If PS_GA100_UpdateData() = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
                
                Call PS_GA100_FormReset
                oForm01.Mode = fm_ADD_MODE
                
                Call PS_GA100_LoadCaption
                Call PS_GA100_MTX01
                
'                oForm01.Items("GCode").Click ct_Regular
            End If
            
        ElseIf pval.ItemUID = "BtnSearch" Then '/조회
            
            Call PS_GA100_FormReset
            oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
            
            Call PS_GA100_LoadCaption
            Call PS_GA100_MTX01
            
        ElseIf pval.ItemUID = "BtnDelete" Then '/삭제
        
            If Sbo_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", "1", "예", "아니오") = "1" Then
        
                Call PS_GA100_DeleteData
                Call PS_GA100_FormReset
                oForm01.Mode = fm_ADD_MODE '/fm_VIEW_MODE
                
                Call PS_GA100_LoadCaption
                Call PS_GA100_MTX01
                
            Else
            
            End If
            
        End If
  
    ElseIf pval.BeforeAction = False Then
        If pval.ItemUID = "PS_GA100" Then
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
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "TeamCd1", "") '기본정보-팀1 포맷서치
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "RspCd1", "") '기본정보-담당1 포맷서치
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "TeamCd2", "") '기본정보-팀2 포맷서치
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "RspCd2", "") '기본정보-담당2 포맷서치
        
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "STeamCd1", "") '조회조건-팀1 포맷서치
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "SRspCd1", "") '조회조건-담당1 포맷서치
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "STeamCd2", "") '조회조건-팀2 포맷서치
        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "SRspCd2", "") '조회조건-담당2 포맷서치
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
    
Raise_EVENT_KEY_DOWN_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_KEY_DOWN_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CLICK_Error

    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "Mat01" Then
        
            If pval.Row > 0 Then
            
                Call oMat01.SelectRow(pval.Row, True, False)
                
                Call oForm01.Freeze(True)

                'DataSource를 이용하여 각 컨트롤에 값을 출력
                Call oDS_PS_GA100H.setValue("DocEntry", 0, oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE) '관리번호
                Call oDS_PS_GA100H.setValue("U_TrgtCls", 0, oMat01.Columns("TrgtCls").Cells(pval.Row).Specific.VALUE) '목표구분
                Call oDS_PS_GA100H.setValue("U_BPLId", 0, oMat01.Columns("BPLId").Cells(pval.Row).Specific.VALUE) '사업장
                Call oDS_PS_GA100H.setValue("U_TeamCd1", 0, oMat01.Columns("TeamCd1").Cells(pval.Row).Specific.VALUE) '팀1
                Call oDS_PS_GA100H.setValue("U_TeamNm1", 0, oMat01.Columns("TeamNm1").Cells(pval.Row).Specific.VALUE) '팀명1
                Call oDS_PS_GA100H.setValue("U_RspCd1", 0, oMat01.Columns("RspCd1").Cells(pval.Row).Specific.VALUE) '담당1
                Call oDS_PS_GA100H.setValue("U_RspNm1", 0, oMat01.Columns("RspNm1").Cells(pval.Row).Specific.VALUE) '담당명1
                Call oDS_PS_GA100H.setValue("U_TeamCd2", 0, oMat01.Columns("TeamCd2").Cells(pval.Row).Specific.VALUE) '팀2
                Call oDS_PS_GA100H.setValue("U_TeamNm2", 0, oMat01.Columns("TeamNm2").Cells(pval.Row).Specific.VALUE) '팀명2
                Call oDS_PS_GA100H.setValue("U_RspCd2", 0, oMat01.Columns("RspCd2").Cells(pval.Row).Specific.VALUE) '담당2
                Call oDS_PS_GA100H.setValue("U_RspNm2", 0, oMat01.Columns("RspNm2").Cells(pval.Row).Specific.VALUE) '담당명2
                Call oDS_PS_GA100H.setValue("U_TrgtNod", 0, oMat01.Columns("TrgtNod").Cells(pval.Row).Specific.VALUE) '목표일수
                Call oDS_PS_GA100H.setValue("U_Grd", 0, oMat01.Columns("Grd").Cells(pval.Row).Specific.VALUE) '설정등급
                Call oDS_PS_GA100H.setValue("U_TrgtMtp", 0, oMat01.Columns("TrgtMtp").Cells(pval.Row).Specific.VALUE) '목표배수
                Call oDS_PS_GA100H.setValue("U_StrDt", 0, oMat01.Columns("StrDt").Cells(pval.Row).Specific.VALUE) '재개시일자
                Call oDS_PS_GA100H.setValue("U_StrPrns", 0, oMat01.Columns("StrPrns").Cells(pval.Row).Specific.VALUE) '재개시인원
                Call oDS_PS_GA100H.setValue("U_CmDt", 0, oMat01.Columns("CmDt").Cells(pval.Row).Specific.VALUE)  '달성예정일
                Call oDS_PS_GA100H.setValue("U_CmYN", 0, oMat01.Columns("CmYN").Cells(pval.Row).Specific.VALUE) '달성여부
                
                oForm01.Mode = fm_UPDATE_MODE
                Call PS_GA100_LoadCaption
                
                Call oForm01.Freeze(False)
                
            End If
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    
    Exit Sub
    
Raise_EVENT_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_COMBO_SELECT(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_COMBO_SELECT_Error

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
        
            If (pval.ItemUID = "Mat01") Then
'                If (pval.ColUID = "ItemCode") Then
'                    '//기타작업
'                    Call oDS_PS_GA100L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                    If oMat01.RowCount = pval.Row And Trim(oDS_PS_GA100L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                        PS_GA100_AddMatrixRow (pval.Row)
'                    End If
'                Else
'                    Call oDS_PS_GA100L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
'                End If
            Else
                
                Call PS_GA100_FlushToItemValue(pval.ItemUID)
                
                If pval.ItemUID = "TeamCd1" Then
                
                    oForm01.Items("TeamNm1").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm01.Items("TeamCd1").Specific.VALUE & "'") '팀1
                    
                ElseIf pval.ItemUID = "RspCd1" Then
                
                    oForm01.Items("RspNm1").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm01.Items("RspCd1").Specific.VALUE & "'")  '담당1
                    
                ElseIf pval.ItemUID = "TeamCd2" Then
                
                    oForm01.Items("TeamNm2").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm01.Items("TeamCd2").Specific.VALUE & "'") '팀2
                    
                ElseIf pval.ItemUID = "RspCd2" Then
                
                    oForm01.Items("RspNm2").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm01.Items("RspCd2").Specific.VALUE & "'")  '담당2
                    
                ElseIf pval.ItemUID = "STeamCd1" Then
                
                    oForm01.Items("STeamNm1").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm01.Items("STeamCd1").Specific.VALUE & "'") '팀1
                    
                ElseIf pval.ItemUID = "SRspCd1" Then
                
                    oForm01.Items("SRspNm1").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm01.Items("SRspCd1").Specific.VALUE & "'")  '담당1
                
                ElseIf pval.ItemUID = "STeamCd2" Then
                
                    oForm01.Items("STeamNm2").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm01.Items("STeamCd2").Specific.VALUE & "'") '팀2
                    
                ElseIf pval.ItemUID = "SRspCd2" Then
                
                    oForm01.Items("SRspNm2").Specific.VALUE = MDC_GetData.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" & oForm01.Items("SRspCd2").Specific.VALUE & "'")  '담당2
                    
                Else

                End If
                
            End If
'            oMat01.LoadFromDataSource
'            oMat01.AutoResizeColumns
'            oForm01.Update
        End If
        
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
        Call PS_GA100_FormItemEnabled
        '//Call PS_GA100_AddMatrixRow(oMat01.VisualRowCount) '//UDO방식
    End If
    Exit Sub
Raise_EVENT_MATRIX_LOAD_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_MATRIX_LOAD_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_RESIZE(Optional ByRef FormUID, Optional ByRef pval As SAPbouiCOM.ItemEvent, Optional ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_RESIZE_Error
    If pval.BeforeAction = True Then
        
    ElseIf pval.BeforeAction = False Then
        Call PS_GA100_FormResize
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
'            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_GA100H", "U_CardCode,U_CardName")
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
'            If (PS_GA100_Validate("행삭제") = False) Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_GA100H.RemoveRecord(oDS_PS_GA100H.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call PS_GA100_Add_MatrixRow(0)
            Else
                If Trim(oDS_PS_GA100H.GetValue("U_CntcCode", oMat01.RowCount - 1)) <> "" Then
                    Call PS_GA100_Add_MatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function PS_GA100_CreateItems() As Boolean
On Error GoTo PS_GA100_CreateItems_Error

    Call oForm01.Freeze(True)
    
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_GA100H = oForm01.DataSources.DBDataSources("@PS_GA100H")
    Set oDS_PS_GA100L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '목표구분_S
    Call oForm01.DataSources.UserDataSources.Add("STrgtCls", dt_SHORT_TEXT, 10)
    Call oForm01.Items("STrgtCls").Specific.DataBind.SetBound(True, "", "STrgtCls")
    '목표구분_E
    
    '사업장_S
    Call oForm01.DataSources.UserDataSources.Add("SBPLId", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SBPLId").Specific.DataBind.SetBound(True, "", "SBPLId")
    '사업장_E

    '팀1
    Call oForm01.DataSources.UserDataSources.Add("STeamCd1", dt_SHORT_TEXT, 20)
    Call oForm01.Items("STeamCd1").Specific.DataBind.SetBound(True, "", "STeamCd1")
    '팀1
    
    '팀명1
    Call oForm01.DataSources.UserDataSources.Add("STeamNm1", dt_SHORT_TEXT, 50)
    Call oForm01.Items("STeamNm1").Specific.DataBind.SetBound(True, "", "STeamNm1")
    '팀명1
    
    '담당1
    Call oForm01.DataSources.UserDataSources.Add("SRspCd1", dt_SHORT_TEXT, 20)
    Call oForm01.Items("SRspCd1").Specific.DataBind.SetBound(True, "", "SRspCd1")
    '담당1
    
    '담당명1
    Call oForm01.DataSources.UserDataSources.Add("SRspNm1", dt_SHORT_TEXT, 50)
    Call oForm01.Items("SRspNm1").Specific.DataBind.SetBound(True, "", "SRspNm1")
    '담당명1
    
    '팀2
    Call oForm01.DataSources.UserDataSources.Add("STeamCd2", dt_SHORT_TEXT, 20)
    Call oForm01.Items("STeamCd2").Specific.DataBind.SetBound(True, "", "STeamCd2")
    '팀2
    
    '팀명2
    Call oForm01.DataSources.UserDataSources.Add("STeamNm2", dt_SHORT_TEXT, 50)
    Call oForm01.Items("STeamNm2").Specific.DataBind.SetBound(True, "", "STeamNm2")
    '팀명2
    
    '담당2
    Call oForm01.DataSources.UserDataSources.Add("SRspCd2", dt_SHORT_TEXT, 20)
    Call oForm01.Items("SRspCd2").Specific.DataBind.SetBound(True, "", "SRspCd2")
    '담당2
    
    '담당명2
    Call oForm01.DataSources.UserDataSources.Add("SRspNm2", dt_SHORT_TEXT, 50)
    Call oForm01.Items("SRspNm2").Specific.DataBind.SetBound(True, "", "SRspNm2")
    '담당명2
    
    '달성여부
    Call oForm01.DataSources.UserDataSources.Add("SCmYN", dt_SHORT_TEXT, 10)
    Call oForm01.Items("SCmYN").Specific.DataBind.SetBound(True, "", "SCmYN")
    '달성여부
    
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Function
PS_GA100_CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA100_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

'/콤보박스 set
Sub PS_GA100_ComboBox_Setting()
On Error GoTo PS_GA100_ComboBox_Setting_Error

    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)

    '기본정보-목표구분
    sQry = "            SELECT  U_Code,"
    sQry = sQry & "             U_CodeNm"
    sQry = sQry & "  FROM   [@PS_GA050L]"
    sQry = sQry & "  WHERE  Code = '10'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    
    Call oForm01.Items("TrgtCls").Specific.ValidValues.Add("%", "선택")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("TrgtCls").Specific, sQry, "%", False, False)

    '기본정보-사업장
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    Set oCombo = oForm01.Items("BPLId").Specific
    Call oCombo.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    '기본정보-달성여부
    sQry = "            SELECT  U_Code,"
    sQry = sQry & "             U_CodeNm"
    sQry = sQry & "  FROM   [@PS_GA050L]"
    sQry = sQry & "  WHERE  Code = '11'"
    sQry = sQry & "             AND U_UseYN = 'Y'"

    Call MDC_SetMod.Set_ComboList(oForm01.Items("CmYN").Specific, sQry, "01", False, False)
    
    
    '조회조건-목표구분
    sQry = "            SELECT  U_Code,"
    sQry = sQry & "             U_CodeNm"
    sQry = sQry & "  FROM   [@PS_GA050L]"
    sQry = sQry & "  WHERE  Code = '10'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    
    Call oForm01.Items("STrgtCls").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("STrgtCls").Specific, sQry, "%", False, False)
    
    '조회조건-사업장
    Call oForm01.Items("SBPLId").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
    Set oCombo = oForm01.Items("SBPLId").Specific
    Call oCombo.Select(MDC_PS_Common.User_BPLId, psk_ByValue)
    
    '조회조건-달성여부
    sQry = "            SELECT  U_Code,"
    sQry = sQry & "             U_CodeNm"
    sQry = sQry & "  FROM   [@PS_GA050L]"
    sQry = sQry & "  WHERE  Code = '11'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    
    Call oForm01.Items("SCmYN").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("SCmYN").Specific, sQry, "01", False, False)
    
    
    '매트릭스-목표구분
    sQry = "            SELECT  U_Code,"
    sQry = sQry & "             U_CodeNm"
    sQry = sQry & "  FROM   [@PS_GA050L]"
    sQry = sQry & "  WHERE  Code = '10'"
    sQry = sQry & "             AND U_UseYN = 'Y'"

    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("TrgtCls"), sQry)
    
    '매트릭스-사업장
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId")
    
'    '매트릭스-팀1
'    sQry = "           SELECT      b.U_Code, "
'    sQry = sQry & "                b.U_CodeNm"
'    sQry = sQry & " FROM       [@PS_HR200H] a"
'    sQry = sQry & "                INNER JOIN"
'    sQry = sQry & "                [@PS_HR200L] b"
'    sQry = sQry & "                    On a.Code = b.Code"
'    sQry = sQry & "                    And a.Name = '부서'"
'
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("TeamCd1"), sQry)
'
'    '매트릭스-담당1
'    sQry = "           SELECT      b.U_Code, "
'    sQry = sQry & "                b.U_CodeNm"
'    sQry = sQry & " FROM       [@PS_HR200H] a"
'    sQry = sQry & "                INNER JOIN"
'    sQry = sQry & "                [@PS_HR200L] b"
'    sQry = sQry & "                    On a.Code = b.Code"
'    sQry = sQry & "                    And a.Name = '담당'"
'
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("RspCd1"), sQry)
'
'    '매트릭스-팀2
'    sQry = "           SELECT      b.U_Code, "
'    sQry = sQry & "                b.U_CodeNm"
'    sQry = sQry & " FROM       [@PS_HR200H] a"
'    sQry = sQry & "                INNER JOIN"
'    sQry = sQry & "                [@PS_HR200L] b"
'    sQry = sQry & "                    On a.Code = b.Code"
'    sQry = sQry & "                    And a.Name = '부서'"
'
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("TeamCd2"), sQry)
'
'    '매트릭스-담당2
'    sQry = "           SELECT      b.U_Code, "
'    sQry = sQry & "                b.U_CodeNm"
'    sQry = sQry & " FROM       [@PS_HR200H] a"
'    sQry = sQry & "                INNER JOIN"
'    sQry = sQry & "                [@PS_HR200L] b"
'    sQry = sQry & "                    On a.Code = b.Code"
'    sQry = sQry & "                    And a.Name = '담당'"
'
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("RspCd2"), sQry)
    
    '매트릭스-달성여부
    sQry = "            SELECT  U_Code,"
    sQry = sQry & "             U_CodeNm"
    sQry = sQry & "  FROM   [@PS_GA050L]"
    sQry = sQry & "  WHERE  Code = '11'"
    sQry = sQry & "             AND U_UseYN = 'Y'"
    
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("CmYN"), sQry)
    
    Call oForm01.Freeze(False)
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    
    Exit Sub
PS_GA100_ComboBox_Setting_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_GA100_ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PS_GA100_CF_ChooseFromList()
On Error GoTo PS_GA100_CF_ChooseFromList_Error

    Exit Sub
PS_GA100_CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA100_CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA100_EnableMenus()
On Error GoTo PS_GA100_EnableMenus_Error
    
    Exit Sub
PS_GA100_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA100_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA100_SetDocument(ByVal oFromDocEntry01 As String)
On Error GoTo PS_GA100_SetDocument_Error
    If (oFromDocEntry01 = "") Then
        Call PS_GA100_FormItemEnabled
        '//Call PS_GA100_AddMatrixRow(0, True) '//UDO방식일때
    Else
'        oForm01.Mode = fm_FIND_MODE
'        Call PS_GA100_FormItemEnabled
'        oForm01.Items("DocEntry").Specific.Value = oFromDocEntry01
'        oForm01.Items("1").Click ct_Regular
    End If
    Exit Sub
PS_GA100_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA100_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PS_GA100_FormResize()
On Error GoTo PS_GA100_FormResize_Error
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
PS_GA100_FormResize_Error:
    Sbo_Application.SetStatusBarMessage "PS_GA100_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub PS_GA100_FormReset()
'******************************************************************************
'Function ID : PS_GA100_FormReset()
'해당모듈 : PS_GA100
'기능 : 화면 초기화
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_GA100_FormReset_Error

    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    '관리번호
    sQry = "SELECT ISNULL(MAX(DocEntry), 0) FROM [@PS_GA100H]"
    Call RecordSet01.DoQuery(sQry)
    
    If Trim(RecordSet01.Fields(0).VALUE) = 0 Then
        Call oDS_PS_GA100H.setValue("DocEntry", 0, 1)
    Else
        Call oDS_PS_GA100H.setValue("DocEntry", 0, Trim(RecordSet01.Fields(0).VALUE) + 1)
    End If
    
    Call oDS_PS_GA100H.setValue("U_TrgtCls", 0, "%") '목표구분
    Call oDS_PS_GA100H.setValue("U_BPLId", 0, MDC_PS_Common.User_BPLId) '사업장
    Call oDS_PS_GA100H.setValue("U_TeamCd1", 0, "") '팀1
    Call oDS_PS_GA100H.setValue("U_TeamNm1", 0, "") '팀명1
    Call oDS_PS_GA100H.setValue("U_RspCd1", 0, "") '담당1
    Call oDS_PS_GA100H.setValue("U_RspNm1", 0, "") '담당명1
    Call oDS_PS_GA100H.setValue("U_TeamCd2", 0, "") '팀2
    Call oDS_PS_GA100H.setValue("U_TeamNm2", 0, "") '팀명2
    Call oDS_PS_GA100H.setValue("U_RspCd2", 0, "") '담당2
    Call oDS_PS_GA100H.setValue("U_RspNm2", 0, "") '담당명2
    Call oDS_PS_GA100H.setValue("U_TrgtNod", 0, "") '목표일수
    Call oDS_PS_GA100H.setValue("U_Grd", 0, "") '설정등급
    Call oDS_PS_GA100H.setValue("U_TrgtMtp", 0, "") '목표배수
    Call oDS_PS_GA100H.setValue("U_StrDt", 0, "") '재개시일자
    Call oDS_PS_GA100H.setValue("U_StrPrns", 0, "") '재개시인원
    Call oDS_PS_GA100H.setValue("U_CmDt", 0, "") '달성예정일
    Call oDS_PS_GA100H.setValue("U_CmYN", 0, "01") '달성여부
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)

Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_GA100_FormReset_Error:
    Call oForm01.Freeze(False)
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PS_GA100_FormReset_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


