VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_FX030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_FX030.cls
'//  Module : FX
'//  Description : 고정자산 현황 및 상태 등록
'//  FormType: PS_FX030
'//  Create Date : 2016.07.21
'//  Modified Date  :
'//  Creator: Song Myounggyu
'//  Company: Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_FX030L As SAPbouiCOM.DBDataSource '등록라인

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_FX030.srf")
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
    
    oFormUniqueID01 = "PS_FX030_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_OK_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
'    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), False        '// 행삭제
        
    oForm01.Update
    oForm01.Freeze False
    oForm01.Visible = True
    
    Set oXmlDoc01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    If (oForm01 Is Nothing) = False Then
        Set oForm01 = Nothing
    End If
    MDC_Com.MDC_GF_Message "LoadForm_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Initialization()
On Error GoTo Initialization_Error
    Dim oCombo As SAPbouiCOM.ComboBox
    
    '//아이디별 사업장 세팅
    Set oCombo = oForm01.Items("BPLID").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    '//기본년월
    oForm01.Items("StdYear").Specific.VALUE = Left(Format(Now, "YYYY"), 4)
    
    '//아이디별 사번 세팅
'    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD
    
    '//아이디별 부서 세팅
'    Set oCombo = oForm01.Items("DeptCode").Specific
'    oCombo.Select MDC_PS_Common.User_DeptCode, psk_ByValue
    Set oCombo = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Initialization_Error:
    Set oCombo = Nothing
    MDC_Com.MDC_GF_Message "Initialization_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//디비데이터 소스 개체 할당
'    Set oDS_PS_FX030H = oForm01.DataSources.DBDataSources("@PS_FX030H")
'    Set oDS_PS_FX030L = oForm01.DataSources.DBDataSources("@PS_FX030L")
    
    '// 메트릭스 개체 할당
'    Set oMat01 = oForm01.Items("Mat01").Specific
    
'    Call oForm01.DataSources.UserDataSources.Add("DocDateFr", dt_DATE, 10)
'    oForm01.Items("DocDateFr").Specific.DataBind.SetBound True, "", "DocDateFr"
'    oForm01.DataSources.UserDataSources.Item("DocDateFr").VALUE = Format(Now, "YYYYMMDD")

'    Call oForm01.DataSources.UserDataSources.Add("DocDateTo", dt_DATE, 10)
'    oForm01.Items("DocDateTo").Specific.DataBind.SetBound True, "", "DocDateTo"
'    oForm01.DataSources.UserDataSources.Item("DocDateTo").VALUE = Format(Now, "YYYYMMDD")

    '//디비데이터 소스 개체 할당
'    Set oDS_PS_FX030H = oForm01.DataSources.DBDataSources("@PS_FX030H")
    Set oDS_PS_FX030L = oForm01.DataSources.DBDataSources("@PS_USERDS01")

    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '사업장_S
    Call oForm01.DataSources.UserDataSources.Add("BPLID", dt_SHORT_TEXT, 5)
    Call oForm01.Items("BPLID").Specific.DataBind.SetBound(True, "", "BPLID")
'    oForm01.DataSources.UserDataSources.Item("DocDate").VALUE = Format(Now, "YYYYMMDD")
    '사업장_E
    
    '기준년도_S
    Call oForm01.DataSources.UserDataSources.Add("StdYear", dt_SHORT_TEXT, 4)
    Call oForm01.Items("StdYear").Specific.DataBind.SetBound(True, "", "StdYear")
    '기준년도_E
    
    '기준년월_S
    Call oForm01.DataSources.UserDataSources.Add("StdYM", dt_SHORT_TEXT, 6) '가장최근 감가상각 년월
    Call oForm01.Items("StdYM").Specific.DataBind.SetBound(True, "", "StdYM")
    '기준년월_E
    
    '자산분류_S
    Call oForm01.DataSources.UserDataSources.Add("FixType", dt_SHORT_TEXT, 10)
    Call oForm01.Items("FixType").Specific.DataBind.SetBound(True, "", "FixType")
    '자산분류_E
    
    '폐기대상 조회_S
    Call oForm01.DataSources.UserDataSources.Add("ChkDisu", dt_SHORT_TEXT, 10)
    Call oForm01.Items("ChkDisu").Specific.DataBind.SetBound(True, "", "ChkDisu")
    '폐기대상 조회_E
        
    
'    '부서분류_S
'    Call oForm01.DataSources.UserDataSources.Add("TeamCD", dt_SHORT_TEXT, 10)
'    Call oForm01.Items("TeamCD").Specific.DataBind.SetBound(True, "", "TeamCD")
'    '자산분류_E
'
'    '부서분류_S
'    Call oForm01.DataSources.UserDataSources.Add("TeamNM", dt_SHORT_TEXT, 10)
'    Call oForm01.Items("TeamNM").Specific.DataBind.SetBound(True, "", "TeamNM")
'    '자산분류_E
'
'
'    '담당분류_S
'    Call oForm01.DataSources.UserDataSources.Add("RspCD", dt_SHORT_TEXT, 10)
'    Call oForm01.Items("RspCD").Specific.DataBind.SetBound(True, "", "RspCD")
'    '자산분류_E
'
'
'    '담당분류_S
'    Call oForm01.DataSources.UserDataSources.Add("RspNM", dt_SHORT_TEXT, 10)
'    Call oForm01.Items("RspNM").Specific.DataBind.SetBound(True, "", "RspNM")
'    '자산분류_E
  
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CreateItems_Error:
    MDC_Com.MDC_GF_Message "CreateItems_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    '//콤보에 기본값설정
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
       
    '// 사업장
    Set oCombo = oForm01.Items("BPLID").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL]"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
       
    '자산분류
    sQry = "            SELECT     U_Minor,"
    sQry = sQry & "                U_CdName"
    sQry = sQry & "  FROM      [@PS_SY001L]"
    sQry = sQry & "  WHERE    Code = 'FX001'"
    sQry = sQry & "                AND U_UseYN = 'Y'"
    
    Call oForm01.Items("FixType").Specific.ValidValues.Add("%", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("FixType").Specific, sQry, "", False, False)
    Call oForm01.Items("FixType").Specific.Select(0, psk_Index)
     
    
    '상태(Line)
    sQry = "           SELECT      U_Minor, "
    sQry = sQry & "                U_CdName"
    sQry = sQry & " FROM       [@PS_SY001L]"
    sQry = sQry & " WHERE      Code = 'FX004'"
    sQry = sQry & "                AND U_UseYN = 'Y'"
    
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Status"), sQry)
       
'    Call oForm01.Freeze(True)
    
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "ComboBox_Setting_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'****************************************************************************************************************
'// ItemEventHander
'****************************************************************************************************************
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
    Dim i&
    Dim ErrNum&
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
'                        If HeaderSpaceLineDel = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
'                        If MatrixSpaceLineDel = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
                    End If
                
                ElseIf pval.ItemUID = "BtnSrch" Then
                
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        Call PS_FX030_MTX01
                    End If
                
                ElseIf pval.ItemUID = "BtnSave" Then '저장
                
                    Call PS_FX030_AddData
                
                ElseIf pval.ItemUID = "BtnPrint" Then '출력버튼 클릭시
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        Call PS_FX030_Print_Report01
                    End If
                    
                ElseIf pval.ItemUID = "BtnPrt02" Then
                
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        Call PS_FX030_Print_Report02
                    End If
            
                ElseIf pval.ItemUID = "BtnPrt03" Then
                
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        Call PS_FX030_Print_Report03
                    End If
                    
                ElseIf pval.ItemUID = "BtnPrt04" Then
                
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        Call PS_FX030_Print_Report04
                    End If
                
                ElseIf pval.ItemUID = "BtnAll" Then '전체선택(해제)
        
                    Call PS_FX030_CheckAll
                    
                End If
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                End If
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            
                If pval.ItemUID = "Mat01" Then
        
                    If pval.Row > 0 Then
                    
                        Call oMat01.SelectRow(pval.Row, True, False)
                        
                    End If
                    
                End If
            
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            
'                If pval.ItemChanged = True Then
'
'                    If pval.ItemUID = "Mat01" Then
'
'                        If pval.ColUID = "Comment" Then
'
'                            Call oForm01.Freeze(True)
'
'                            Call oMat01.AutoResizeColumns
'
'                            Call oForm01.Freeze(False)
'
'                        End If
'
'                    End If
'
'                Else
'
'                End If
                
'                Call FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                Call oForm01.Update
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
        End Select
        
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
'                If pval.ItemUID = "1" Then
'                    If oForm01.Mode = fm_ADD_MODE Then
'                        oForm01.Mode = fm_OK_MODE
'                        Call Sbo_Application.ActivateMenuItem("1282")
'                    ElseIf oForm01.Mode = fm_OK_MODE Then
'                        FormItemEnabled
'                        Call Matrix_AddRow(1, oMat01.RowCount, False) 'oMat01
'                    End If
'                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            
                Call FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
            
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
'                Set oMat01 = Nothing
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    Set ProgressBar01 = Nothing
    If ErrNum = 101 Then
        ErrNum = 0
        MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
        BubbleEvent = False
    Else
        MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    Dim i&
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
            Case "1285": '복원
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                Call oForm01.Freeze(True)
                Call PS_FX030_Add_MatrixRow(oMat01.VisualRowCount)
                Call oForm01.Freeze(False)
            
        End Select
        
    '-----------------------------------------------------------------------------------------------------------
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1285": '복원
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            
            Case "7169": '엑셀 내보내기
            
                '엑셀 내보내기 이후 처리
                Call oForm01.Freeze(True)
                Call oDS_PS_FX030L.RemoveRecord(oDS_PS_FX030L.Size - 1)
                Call oMat01.LoadFromDataSource
                Call oForm01.Freeze(False)
                
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_MenuEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If (eventInfo.BeforeAction = True) Then
        
    ElseIf (eventInfo.BeforeAction = False) Then
        '//작업
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_FormDataEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_FormDataEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error

    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet    As SAPbobsCOM.Recordset
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
     Select Case oUID
        Case "TeamCD" '//팀
            sQry = "SELECT  T1.U_CodeNm FROM [@PS_HR200H] AS T0 Inner Join [@PS_HR200L] AS T1 ON T0.Code = T1.Code "
            sQry = sQry & " WHERE   T0.Name = '부서' AND T1.U_UseYN = 'Y' AND T1.U_Char2 = '" & oForm01.Items("BPLID").Specific.VALUE & "' "
            sQry = sQry & " And T1.U_Code = '" & oForm01.Items("TeamCD").Specific.VALUE & "'"
            
            oRecordSet.DoQuery sQry
            oForm01.Items("TeamNM").Specific.VALUE = Trim(oRecordSet.Fields(0).VALUE)
        Case "RspCD" '//담당
            sQry = "SELECT  T1.U_CodeNm FROM [@PS_HR200H] AS T0 Inner Join [@PS_HR200L] AS T1 ON T0.Code = T1.Code "
            sQry = sQry & " WHERE   T0.Name = '담당' AND T1.U_UseYN = 'Y' AND T1.U_Char2 = '" & oForm01.Items("BPLID").Specific.VALUE & "' "
            sQry = sQry & " And T1.U_Char1 = '" & oForm01.Items("TeamCD").Specific.VALUE & "'"
            sQry = sQry & " And T1.U_Code = '" & oForm01.Items("RspCD").Specific.VALUE & "'"
            
            oRecordSet.DoQuery sQry
            oForm01.Items("RspNM").Specific.VALUE = Trim(oRecordSet.Fields(0).VALUE)
            
    End Select
    
    
    If oUID = "Mat01" Then

        If oCol = "Comment" Then

        End If
    
    Else
    
        If oUID = "StdYear" Then

            sQry = "SELECT MAX(U_YM) AS [MaxYM] FROM [@PS_FX020H] WHERE U_BPLId = '" & Trim(oForm01.Items("BPLID").Specific.VALUE) & "'"

            Call oRecordSet.DoQuery(sQry)

            If oForm01.Items("StdYM").Specific.VALUE = "" Then

                oForm01.DataSources.UserDataSources("StdYM").VALUE = oRecordSet.Fields("MaxYM").VALUE
            
            End If

        End If
    
    End If

    Call oForm01.Freeze(False)

    Set oRecordSet = Nothing
    
    Exit Sub

FlushToItemValue_Error:
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    
    ErrNum = 0

    '// Check
    Select Case True
'        Case Trim(oDS_PS_FX030H.GetValue("U_BPLId", 0)) = ""
'            ErrNum = 1
'            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하여 주십시오.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Public Sub PS_FX030_MTX01()
'******************************************************************************
'Function ID : PS_FX030_MTX01()
'해당모듈 : PS_FX030
'기능 : 데이터 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_FX030_MTX01_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim BPLID As String '사업장
    Dim StdYear As String '기준년도
    Dim StdYM As String '기준년월
    Dim FixType As String '자산분류
    Dim ChkDisu As String '폐기대상 조회
    Dim TeamCD As String '팀코드
    Dim RspCD As String '담당코드
    
    
    BPLID = Trim(oForm01.Items("BPLID").Specific.VALUE) '사업장
    StdYear = Trim(oForm01.Items("StdYear").Specific.VALUE) '기준년도
    StdYM = Trim(oForm01.Items("StdYM").Specific.VALUE) '기준년월
    FixType = Trim(oForm01.Items("FixType").Specific.VALUE) '자산분류
    TeamCD = Trim(oForm01.Items("TeamCD").Specific.VALUE) '기준년월
    RspCD = Trim(oForm01.Items("RspCD").Specific.VALUE) '자산분류
    ChkDisu = IIf(oForm01.Items("ChkDisu").Specific.Checked = True, "Y", "N") '폐기대상 조회
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    If ChkDisu = "N" Then '폐기대상 조회 체크 안됨
    
        sQry = "         EXEC [PS_FX030_01] '"
        sQry = sQry & BPLID & "','"
        sQry = sQry & StdYear & "','"
        sQry = sQry & StdYM & "','"
        sQry = sQry & FixType & "','"
        sQry = sQry & TeamCD & "','"
        sQry = sQry & RspCD & "'"
    
    Else '폐기대상 조회 체크
    
        sQry = "         EXEC [PS_FX030_02] '"
        sQry = sQry & BPLID & "','"
        sQry = sQry & StdYear & "','"
        sQry = sQry & StdYM & "','"
        sQry = sQry & FixType & "','"
        sQry = sQry & TeamCD & "','"
        sQry = sQry & RspCD & "'"
    
    End If
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_FX030L.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
'        Call PS_FX030_Add_MatrixRow(0, True)
        GoTo PS_FX030_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_FX030L.Size Then
            oDS_PS_FX030L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_FX030L.Offset = i
        
        Call oDS_PS_FX030L.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_FX030L.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("Check").VALUE)) '선택
        Call oDS_PS_FX030L.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("FixCode").VALUE)) '자산코드
        Call oDS_PS_FX030L.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("FixName").VALUE)) '품명
        Call oDS_PS_FX030L.setValue("U_ColReg04", i, Trim(oRecordSet01.Fields("FixSpec").VALUE)) '규격
        Call oDS_PS_FX030L.setValue("U_ColQty01", i, Trim(oRecordSet01.Fields("PostQty").VALUE)) '수량
        Call oDS_PS_FX030L.setValue("U_ColDt01", i, Format(Trim(oRecordSet01.Fields("PostDate").VALUE), "YYYYMMDD")) '구입일자
        Call oDS_PS_FX030L.setValue("U_ColDt02", i, Format(Trim(oRecordSet01.Fields("SPostDate").VALUE), "YYYYMMDD")) '기초취득일
        Call oDS_PS_FX030L.setValue("U_ColSum01", i, Trim(oRecordSet01.Fields("LongYear").VALUE)) '내용년수
        Call oDS_PS_FX030L.setValue("U_ColSum02", i, Trim(oRecordSet01.Fields("PostAmt").VALUE)) '취득금액
        Call oDS_PS_FX030L.setValue("U_ColSum03", i, Trim(oRecordSet01.Fields("CurAmt").VALUE)) '현재잔액
        Call oDS_PS_FX030L.setValue("U_ColReg10", i, Trim(oRecordSet01.Fields("TeamName").VALUE)) '부서
        Call oDS_PS_FX030L.setValue("U_ColReg11", i, Trim(oRecordSet01.Fields("RspName").VALUE)) '담당
        Call oDS_PS_FX030L.setValue("U_ColReg12", i, Trim(oRecordSet01.Fields("ClsName").VALUE)) '반
        Call oDS_PS_FX030L.setValue("U_ColQty02", i, Trim(oRecordSet01.Fields("Qty").VALUE)) '현황
        Call oDS_PS_FX030L.setValue("U_ColReg14", i, Trim(oRecordSet01.Fields("Status").VALUE)) '상태
        Call oDS_PS_FX030L.setValue("U_ColReg15", i, Trim(oRecordSet01.Fields("Comment").VALUE)) '비고
        Call oDS_PS_FX030L.setValue("U_ColReg16", i, Trim(oRecordSet01.Fields("FixType").VALUE)) '자산분류
        Call oDS_PS_FX030L.setValue("U_ColReg17", i, Trim(oRecordSet01.Fields("CDate").VALUE)) '등록일
        Call oDS_PS_FX030L.setValue("U_ColReg18", i, Trim(oRecordSet01.Fields("CUName").VALUE)) '등록자이름
        Call oDS_PS_FX030L.setValue("U_ColReg19", i, Trim(oRecordSet01.Fields("UDate").VALUE)) '수정일
        Call oDS_PS_FX030L.setValue("U_ColReg20", i, Trim(oRecordSet01.Fields("UUName").VALUE)) '수정자이름
            
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
PS_FX030_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_FX030_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PS_FX030_AddData() As Boolean
'******************************************************************************
'Function ID : PS_FX030_AddData()
'해당모듈 : PS_FX030
'기능 : 데이터 INSERT(UPDATE)
'인수 : 없음
'반환값 : 성공여부
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_FX030_AddData_Error

    Dim loopCount As Integer
    Dim sQry As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID As String '사업장
    Dim StdYear As String '기준년도
    Dim FixCode As String '자산코드
    Dim Qty As String '현황(수량)
    Dim Status As String '상태
    Dim Comment As String '비고
    
    Dim UserSign As String 'UserSign

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장 중...", 100, False)
    
    BPLID = oForm01.Items("BPLID").Specific.VALUE '사업장
    StdYear = oForm01.Items("StdYear").Specific.VALUE '기준년월
    UserSign = Sbo_Company.UserSignature '등록자(수정자)
    
    Call oMat01.FlushToDataSource
    For loopCount = 0 To oMat01.VisualRowCount - 1
    
'        If Trim(oDS_PS_FX030L.GetValue("U_ColReg01", loopCount)) = "Y" Then
        
        If oMat01.Columns("Check").Cells(loopCount + 1).Specific.Checked = True Then '체크 확인 로직 변경(2015.10.06 송명규 FlushDataSource 후 에도 Check box 값을 못 가져옴, Matrix의 값을 직접 가져옴)
        
            FixCode = Trim(oDS_PS_FX030L.GetValue("U_ColReg02", loopCount)) '자산코드
            Qty = Trim(oDS_PS_FX030L.GetValue("U_ColQty02", loopCount)) '현황
            Status = Trim(oDS_PS_FX030L.GetValue("U_ColReg14", loopCount)) '상태
            Comment = Trim(oDS_PS_FX030L.GetValue("U_ColReg15", loopCount)) '비고
            
            sQry = "                EXEC [PS_FX030_03] "
            sQry = sQry & "'" & BPLID & "'," '사업장
            sQry = sQry & "'" & StdYear & "'," '기준년도
            sQry = sQry & "'" & FixCode & "'," '자산코드
            sQry = sQry & "'" & Qty & "'," '현황
            sQry = sQry & "'" & Status & "'," '상태
            sQry = sQry & "'" & Comment & "'," '비고
            sQry = sQry & "'" & UserSign & "'"  'UserSign
            
            Call RecordSet01.DoQuery(sQry)
    
        End If
        
    Next
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    
    Call MDC_Com.MDC_GF_Message("등록 완료!", "S")
    
    PS_FX030_AddData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_FX030_AddData_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    
    MDC_Com.MDC_GF_Message "PS_FX030_AddData_Error:" & Err.Number & " - " & Err.Description, "E"
    
    PS_FX030_AddData = False
    
End Function

Private Sub PS_FX030_Print_Report01()

    Dim DocNum As String
    Dim ErrNum As Integer
    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PS_FX030_Print_Report01_Error
    
    Dim BPLID As String '사업장
    Dim StdYear As String '기준년도
    Dim StdYM As String '기준년월
    Dim FixType As String '자산분류
    Dim ChkDisu As String '폐기대상 조회
    
    BPLID = Trim(oForm01.Items("BPLID").Specific.VALUE) '사업장
    StdYear = Trim(oForm01.Items("StdYear").Specific.VALUE) '기준년도
    StdYM = Trim(oForm01.Items("StdYM").Specific.VALUE) '기준년월
    FixType = Trim(oForm01.Items("FixType").Specific.VALUE) '자산분류
    ChkDisu = IIf(oForm01.Items("ChkDisu").Specific.Checked = True, "Y", "N") '폐기대상 조회
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
     '/ ODBC 연결 체크
    Call ConnectODBC
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PS_FX030] 고정자산 재물조사"
    
    If BPLID = "1" Then '창원
        ReportName = "PS_FX030_01.rpt"
    ElseIf BPLID = "2" Then '동래
        ReportName = "PS_FX030_02.rpt"
    ElseIf BPLID = "3" Then '사상
        ReportName = "PS_FX030_03.rpt"
    End If
    
    '// Formula 수식필드
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
       
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    '/ Procedure 실행"
'    If ChkDisu = "N" Then '폐기대상 조회 체크 안됨
    
        sQry = "         EXEC [PS_FX030_51] '"
        sQry = sQry & BPLID & "','"
        sQry = sQry & StdYear & "','"
        sQry = sQry & StdYM & "','"
        sQry = sQry & FixType & "','"
        sQry = sQry & ChkDisu & "'"
    
'    Else '폐기대상 조회 체크
'
'        sQry = "         EXEC [PS_FX030_52] '"
'        sQry = sQry & BPLID & "','"
'        sQry = sQry & StdYear & "','"
'        sQry = sQry & FixType & "'"
'
'    End If
    
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PS_FX030_Print_Report01_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
    
PS_FX030_Print_Report01_Error:
'    If ErrNum = 1 Then
'        Set oRecordSet = Nothing
'        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
'    Else
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PS_FX030_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'    End If
    
End Sub

Private Sub PS_FX030_Print_Report02()
'******************************************************************************
'Function ID : PS_FX030_Print_Report02()
'해당모듈 : PS_FX030
'기능 : 고정자산보유현황 리포트 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_FX030_Print_Report02_Error
    
    Dim DocNum As String
    Dim ErrNum As Integer
    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
        
    Dim BPLID As String '사업장
    Dim StdYear As String '기준년도
    Dim StdYM As String '기준년월
    
    BPLID = Trim(oForm01.Items("BPLID").Specific.VALUE) '사업장
    StdYear = Trim(oForm01.Items("StdYear").Specific.VALUE) '기준년도
    StdYM = Trim(oForm01.Items("StdYM").Specific.VALUE) '기준년월
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
     '/ ODBC 연결 체크
    Call ConnectODBC
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PS_FX030] 고정자산 보유현황"
    
    If BPLID = "1" Then '창원
        ReportName = "PS_FX030_04.rpt"
    ElseIf BPLID = "2" Then '동래
        ReportName = "PS_FX030_05.rpt"
    ElseIf BPLID = "3" Then '사상
        ReportName = "PS_FX030_06.rpt"
    End If
    
    '// Formula 수식필드
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
       
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    '/ Procedure 실행"
    sQry = "         EXEC [PS_FX030_52] '"
    sQry = sQry & BPLID & "','"
    sQry = sQry & StdYear & "','"
    sQry = sQry & StdYM & "'"
    
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PS_FX030_Print_Report02_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
    
PS_FX030_Print_Report02_Error:
'    If ErrNum = 1 Then
'        Set oRecordSet = Nothing
'        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
'    Else
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PS_FX030_Print_Report02_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'    End If
    
End Sub

Private Sub PS_FX030_Print_Report03()
'******************************************************************************
'Function ID : PS_FX030_Print_Report03()
'해당모듈 : PS_FX030
'기능 : 재물조사 결과 현황 리포트 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_FX030_Print_Report03_Error
    
    Dim DocNum As String
    Dim ErrNum As Integer
    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
        
    Dim BPLID As String '사업장
    Dim StdYear As String '기준년도
    Dim StdYM As String '기준년월
    
    BPLID = Trim(oForm01.Items("BPLID").Specific.VALUE) '사업장
    StdYear = Trim(oForm01.Items("StdYear").Specific.VALUE) '기준년도
    StdYM = Trim(oForm01.Items("StdYM").Specific.VALUE) '기준년월
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
     '/ ODBC 연결 체크
    Call ConnectODBC
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PS_FX030] 고정자산 보유현황"
    
    If BPLID = "1" Then '창원
        ReportName = "PS_FX030_07.rpt"
    ElseIf BPLID = "2" Then '동래
        ReportName = "PS_FX030_08.rpt"
    ElseIf BPLID = "3" Then '사상
        ReportName = "PS_FX030_09.rpt"
    End If
    
    '// Formula 수식필드
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
       
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    '/ Procedure 실행"
    sQry = "         EXEC [PS_FX030_53] '"
    sQry = sQry & BPLID & "','"
    sQry = sQry & StdYear & "','"
    sQry = sQry & StdYM & "'"
    
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PS_FX030_Print_Report03_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
    
PS_FX030_Print_Report03_Error:
'    If ErrNum = 1 Then
'        Set oRecordSet = Nothing
'        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
'    Else
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PS_FX030_Print_Report03_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'    End If
    
End Sub

Private Sub PS_FX030_Print_Report04()
'******************************************************************************
'Function ID : PS_FX030_Print_Report04()
'해당모듈 : PS_FX030
'기능 : 결과현황 일괄 출력
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_FX030_Print_Report04_Error
    
    Dim DocNum As String
    Dim ErrNum As Integer
    Dim WinTitle As String
    Dim ReportName As String
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
        
    Dim BPLID As String '사업장
    Dim StdYear As String '기준년도
    Dim StdYM As String '기준년월
    
    BPLID = Trim(oForm01.Items("BPLID").Specific.VALUE) '사업장
    StdYear = Trim(oForm01.Items("StdYear").Specific.VALUE) '기준년도
    StdYM = Trim(oForm01.Items("StdYM").Specific.VALUE) '기준년월
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)
    
     '/ ODBC 연결 체크
    Call ConnectODBC
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

    WinTitle = "[PS_FX030] 결과현황 일괄조회"
    
    If BPLID = "1" Then '창원
        ReportName = "PS_FX030_10.rpt"
    ElseIf BPLID = "2" Then '동래
        ReportName = "PS_FX030_11.rpt"
    ElseIf BPLID = "3" Then '사상
        ReportName = "PS_FX030_12.rpt"
    End If
    
    '// Formula 수식필드
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
       
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    sQry = "         EXEC [PS_FX030_53] '"
    sQry = sQry & BPLID & "','"
    sQry = sQry & StdYear & "','"
    sQry = sQry & StdYM & "'"
    
    gRpt_SRptSqry(1) = sQry
    gRpt_SRptName(1) = "PS_FX030_SUB1"
    
    '/ Procedure 실행"
    sQry = "         EXEC [PS_FX030_52] '"
    sQry = sQry & BPLID & "','"
    sQry = sQry & StdYear & "','"
    sQry = sQry & StdYM & "'"
    
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PS_FX030_Print_Report04_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "", "Y", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
    
PS_FX030_Print_Report04_Error:
'    If ErrNum = 1 Then
'        Set oRecordSet = Nothing
'        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
'    Else
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PS_FX030_Print_Report04_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'    End If
    
End Sub


Private Sub PS_FX030_CheckAll()
On Error GoTo PS_FX030_CheckAll_Error

    Dim CheckType As String
    Dim loopCount As Integer
    
    Call oForm01.Freeze(True)
    CheckType = "Y"
    
    Call oMat01.FlushToDataSource
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_FX030L.GetValue("U_ColReg01", loopCount)) = "N" Then
            CheckType = "N"
            Exit For
        End If
    Next
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        oDS_PS_FX030L.Offset = loopCount
        If CheckType = "N" Then
            Call oDS_PS_FX030L.setValue("U_ColReg01", loopCount, "Y")
        Else
            Call oDS_PS_FX030L.setValue("U_ColReg01", loopCount, "N")
        End If
    Next
    
    Call oMat01.LoadFromDataSource
    
    Call oForm01.Freeze(False)

    Exit Sub
    
PS_FX030_CheckAll_Error:
    Call oForm01.Freeze(False)
    Call Sbo_Application.SetStatusBarMessage("PS_FX030_CheckAll_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Sub

'/메트릭스 Row추가
Sub PS_FX030_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_FX030_Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_FX030L.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_FX030L.Offset = oRow
    oDS_PS_FX030L.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_FX030_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PS_FX030_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
