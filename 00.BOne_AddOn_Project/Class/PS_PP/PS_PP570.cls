VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP570"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File        : PS_PP570.cls
'//  Module      : 생산관리
'//  Desc        : 작번 비용 이관
'//  FormType    : PS_PP570
'//  Create Date : 2016.08.02
'//  Creator     : Song Myoungyu
'//  Copyright  (c) Poongsan Holdings
'//********************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_PP570L As SAPbouiCOM.DBDataSource '등록라인

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP570.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)

    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_PP570_" & GetTotalFormsCount
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
'    Set oCombo = oForm01.Items("BPLID").Specific
'    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    '//기본년월
'    oForm01.Items("StdYear").Specific.VALUE = Left(Format(Now, "YYYY"), 4)
    
    '//아이디별 사번 세팅
    oForm01.Items("TCntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD
    
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
'    Set oDS_PS_PP570H = oForm01.DataSources.DBDataSources("@PS_PP570H")
'    Set oDS_PS_PP570L = oForm01.DataSources.DBDataSources("@PS_PP570L")
    
    '// 메트릭스 개체 할당
'    Set oMat01 = oForm01.Items("Mat01").Specific
    
'    Call oForm01.DataSources.UserDataSources.Add("DocDateFr", dt_DATE, 10)
'    oForm01.Items("DocDateFr").Specific.DataBind.SetBound True, "", "DocDateFr"
'    oForm01.DataSources.UserDataSources.Item("DocDateFr").VALUE = Format(Now, "YYYYMMDD")

'    Call oForm01.DataSources.UserDataSources.Add("DocDateTo", dt_DATE, 10)
'    oForm01.Items("DocDateTo").Specific.DataBind.SetBound True, "", "DocDateTo"
'    oForm01.DataSources.UserDataSources.Item("DocDateTo").VALUE = Format(Now, "YYYYMMDD")

    '//디비데이터 소스 개체 할당
'    Set oDS_PS_PP570H = oForm01.DataSources.DBDataSources("@PS_PP570H")
    Set oDS_PS_PP570L = oForm01.DataSources.DBDataSources("@PS_USERDS01")

    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    
    '이관전작번
    Call oForm01.DataSources.UserDataSources.Add("BOrdNum", dt_SHORT_TEXT, 11)
    Call oForm01.Items("BOrdNum").Specific.DataBind.SetBound(True, "", "BOrdNum")
    
    '이관전서브1
    Call oForm01.DataSources.UserDataSources.Add("BOrdSub1", dt_SHORT_TEXT, 2)
    Call oForm01.Items("BOrdSub1").Specific.DataBind.SetBound(True, "", "BOrdSub1")
    
    '이관전서브2
    Call oForm01.DataSources.UserDataSources.Add("BOrdSub2", dt_SHORT_TEXT, 3)
    Call oForm01.Items("BOrdSub2").Specific.DataBind.SetBound(True, "", "BOrdSub2")
    
    '이관전품명
    Call oForm01.DataSources.UserDataSources.Add("BItemName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("BItemName").Specific.DataBind.SetBound(True, "", "BItemName")
    
    '이관전규격
    Call oForm01.DataSources.UserDataSources.Add("BItemSpec", dt_SHORT_TEXT, 100)
    Call oForm01.Items("BItemSpec").Specific.DataBind.SetBound(True, "", "BItemSpec")
    
    '이관후작번
    Call oForm01.DataSources.UserDataSources.Add("AOrdNum", dt_SHORT_TEXT, 11)
    Call oForm01.Items("AOrdNum").Specific.DataBind.SetBound(True, "", "AOrdNum")
    
    '이관후서브1
    Call oForm01.DataSources.UserDataSources.Add("AOrdSub1", dt_SHORT_TEXT, 2)
    Call oForm01.Items("AOrdSub1").Specific.DataBind.SetBound(True, "", "AOrdSub1")
    
    '이관후서브2
    Call oForm01.DataSources.UserDataSources.Add("AOrdSub2", dt_SHORT_TEXT, 3)
    Call oForm01.Items("AOrdSub2").Specific.DataBind.SetBound(True, "", "AOrdSub2")
    
    '이관후품명
    Call oForm01.DataSources.UserDataSources.Add("AItemName", dt_SHORT_TEXT, 100)
    Call oForm01.Items("AItemName").Specific.DataBind.SetBound(True, "", "AItemName")
    
    '이관후규격
    Call oForm01.DataSources.UserDataSources.Add("AItemSpec", dt_SHORT_TEXT, 100)
    Call oForm01.Items("AItemSpec").Specific.DataBind.SetBound(True, "", "AItemSpec")
    
    '등록자사번
    Call oForm01.DataSources.UserDataSources.Add("TCntcCode", dt_SHORT_TEXT, 20)
    Call oForm01.Items("TCntcCode").Specific.DataBind.SetBound(True, "", "TCntcCode")
    
    '등록자성명
    Call oForm01.DataSources.UserDataSources.Add("TCntcName", dt_SHORT_TEXT, 20)
    Call oForm01.Items("TCntcName").Specific.DataBind.SetBound(True, "", "TCntcName")
    
    '요청자사번
    Call oForm01.DataSources.UserDataSources.Add("RCntcCode", dt_SHORT_TEXT, 20)
    Call oForm01.Items("RCntcCode").Specific.DataBind.SetBound(True, "", "RCntcCode")
    
    '요청자성명
    Call oForm01.DataSources.UserDataSources.Add("RCntcName", dt_SHORT_TEXT, 20)
    Call oForm01.Items("RCntcName").Specific.DataBind.SetBound(True, "", "RCntcName")
  
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CreateItems_Error:
    MDC_Com.MDC_GF_Message "CreateItems_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    '//콤보에 기본값설정
    Dim oCombo       As SAPbouiCOM.ComboBox
    Dim sQry         As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
       
'    '// 사업장
'    Set oCombo = oForm01.Items("BPLID").Specific
'    sQry = "SELECT BPLId, BPLName From [OBPL] order by 1"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
'        oRecordSet01.MoveNext
'    Loop
'
'    '자산분류
'    sQry = "            SELECT     U_Minor,"
'    sQry = sQry & "                U_CdName"
'    sQry = sQry & "  FROM      [@PS_SY001L]"
'    sQry = sQry & "  WHERE    Code = 'FX001'"
'    sQry = sQry & "                AND U_UseYN = 'Y'"
'
'    Call oForm01.Items("FixType").Specific.ValidValues.Add("%", "전체")
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("FixType").Specific, sQry, "", False, False)
'    Call oForm01.Items("FixType").Specific.Select(0, psk_Index)
'
'    '상태(Line)
'    sQry = "           SELECT      U_Minor, "
'    sQry = sQry & "                U_CdName"
'    sQry = sQry & " FROM       [@PS_SY001L]"
'    sQry = sQry & " WHERE      Code = 'FX004'"
'    sQry = sQry & "                AND U_UseYN = 'Y'"
'
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Status"), sQry)
       
'    Call oForm01.Freeze(True)
    
    '이관구분
    Call oForm01.Items("TransCls").Specific.ValidValues.Add("01", "전체비용")
    Call oForm01.Items("TransCls").Specific.ValidValues.Add("02", "공수")
    Call oForm01.Items("TransCls").Specific.Select(0, psk_Index)
    
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
'                        If DataValidCheck = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
'                        If MatrixSpaceLineDel = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
                    End If
                
                ElseIf pval.ItemUID = "BtnSrch" Then
                    
                    Call PS_PP570_MTX01
                    
                ElseIf pval.ItemUID = "BtnSave" Then '저장
                
                    If DataValidCheck = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        If Sbo_Application.MessageBox("[" & oForm01.Items("BOrdNum").Specific.VALUE & "] 작번의 비용이 [" & oForm01.Items("AOrdNum").Specific.VALUE & "] 작번으로 이관됩니다. " & Chr(13) & "이관후 취소는 불가능합니다. 정말로 저장하시겠습니까?", "1", "예", "아니오") = "1" Then
                    
                            Call PS_PP570_AddData
                            Call PS_PP570_MTX01
                        
                        Else
        
                            BubbleEvent = False
                            Exit Sub
        
                        End If
                    End If
                    
                ElseIf pval.ItemUID = "BtnCfm" Then
                
                    Call PS_PP570_Update_AfterOrdNum
                    
                ElseIf pval.ItemUID = "BtnAll" Then '전체선택(해제)
        
                    Call PS_PP570_CheckAll
                    
                End If
            Case et_KEY_DOWN: '//2
            
                If pval.CharPressed = 9 Then
                
                    Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "BOrdNum", "")
                    Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "AOrdNum", "")
                
                    Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "TCntcCode", "")
                    Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "RCntcCode", "")
                
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
                
                If pval.ItemChanged = True Then
                    Call FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                End If
            
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
                Call PS_PP570_Add_MatrixRow(oMat01.VisualRowCount)
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
                Call oDS_PS_PP570L.RemoveRecord(oDS_PS_PP570L.Size - 1)
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

    Dim ErrNum     As Integer
    Dim sQry       As String
    Dim oRecordSet As SAPbobsCOM.Recordset
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    If oUID = "Mat01" Then

        If oCol = "Comment" Then

        End If
    
    Else
    
        If oUID = "TCntcCode" Then

            sQry = "        SELECT  U_FullName"
            sQry = sQry & " FROM    [@PH_PY001A]"
            sQry = sQry & " WHERE   Code = '" & Trim(oForm01.Items("TCntcCode").Specific.VALUE) & "'"
            Call oRecordSet.DoQuery(sQry)

            oForm01.Items("TCntcName").Specific.VALUE = Trim(oRecordSet.Fields(0).VALUE)

        ElseIf oUID = "RCntcCode" Then
        
            sQry = "        SELECT  U_FullName"
            sQry = sQry & " FROM    [@PH_PY001A]"
            sQry = sQry & " WHERE   Code = '" & Trim(oForm01.Items("RCntcCode").Specific.VALUE) & "'"
            Call oRecordSet.DoQuery(sQry)

            oForm01.Items("RCntcName").Specific.VALUE = Trim(oRecordSet.Fields(0).VALUE)
            
        ElseIf oUID = "BOrdNum" Then
        
            sQry = "        SELECT  U_JakMyung,"
            sQry = sQry & "         U_JakSize"
            sQry = sQry & " FROM    [@PS_PP030H]"
            sQry = sQry & " WHERE   U_OrdNum = '" & Trim(oForm01.Items("BOrdNum").Specific.VALUE) & "'"
            sQry = sQry & "         AND U_OrdSub1 = '00'"
            sQry = sQry & "         AND U_OrdSub2 = '000'"
            Call oRecordSet.DoQuery(sQry)

            oForm01.Items("BOrdSub1").Specific.VALUE = "00" 'Trim(oRecordSet.Fields(0).VALUE)
            oForm01.Items("BOrdSub2").Specific.VALUE = "000"
            oForm01.Items("BItemName").Specific.VALUE = Trim(oRecordSet.Fields(0).VALUE)
            oForm01.Items("BItemSpec").Specific.VALUE = Trim(oRecordSet.Fields(1).VALUE)

        ElseIf oUID = "BOrdSub1" Then
        
            sQry = "        SELECT  U_JakMyung,"
            sQry = sQry & "         U_JakSize"
            sQry = sQry & " FROM    [@PS_PP030H]"
            sQry = sQry & " WHERE   U_OrdNum = '" & Trim(oForm01.Items("BOrdNum").Specific.VALUE) & "'"
            sQry = sQry & "         AND U_OrdSub1 = '" & Trim(oForm01.Items("BOrdSub1").Specific.VALUE) & "'"
            sQry = sQry & "         AND U_OrdSub2 = '000'"
            Call oRecordSet.DoQuery(sQry)

            'oForm01.Items("BOrdSub1").Specific.VALUE = "00" 'Trim(oRecordSet.Fields(0).VALUE)
            oForm01.Items("BOrdSub2").Specific.VALUE = "000"
            oForm01.Items("BItemName").Specific.VALUE = Trim(oRecordSet.Fields(0).VALUE)
            oForm01.Items("BItemSpec").Specific.VALUE = Trim(oRecordSet.Fields(1).VALUE)
            
        ElseIf oUID = "BOrdSub2" Then
        
            sQry = "        SELECT  U_JakMyung,"
            sQry = sQry & "         U_JakSize"
            sQry = sQry & " FROM    [@PS_PP030H]"
            sQry = sQry & " WHERE   U_OrdNum = '" & Trim(oForm01.Items("BOrdNum").Specific.VALUE) & "'"
            sQry = sQry & "         AND U_OrdSub1 = '" & Trim(oForm01.Items("BOrdSub1").Specific.VALUE) & "'"
            sQry = sQry & "         AND U_OrdSub2 = '" & Trim(oForm01.Items("BOrdSub2").Specific.VALUE) & "'"
            Call oRecordSet.DoQuery(sQry)

            'oForm01.Items("BOrdSub1").Specific.VALUE = "00" 'Trim(oRecordSet.Fields(0).VALUE)
            'oForm01.Items("BOrdSub2").Specific.VALUE = "000"
            oForm01.Items("BItemName").Specific.VALUE = Trim(oRecordSet.Fields(0).VALUE)
            oForm01.Items("BItemSpec").Specific.VALUE = Trim(oRecordSet.Fields(1).VALUE)

        ElseIf oUID = "AOrdNum" Then
        
            sQry = "        SELECT  U_JakMyung,"
            sQry = sQry & "         U_JakSize"
            sQry = sQry & " FROM    [@PS_PP030H]"
            sQry = sQry & " WHERE   U_OrdNum = '" & Trim(oForm01.Items("AOrdNum").Specific.VALUE) & "'"
            sQry = sQry & "         AND U_OrdSub1 = '00'"
            sQry = sQry & "         AND U_OrdSub2 = '000'"
            Call oRecordSet.DoQuery(sQry)

            oForm01.Items("AOrdSub1").Specific.VALUE = "00" 'Trim(oRecordSet.Fields(0).VALUE)
            oForm01.Items("AOrdSub2").Specific.VALUE = "000"
            oForm01.Items("AItemName").Specific.VALUE = Trim(oRecordSet.Fields(0).VALUE)
            oForm01.Items("AItemSpec").Specific.VALUE = Trim(oRecordSet.Fields(1).VALUE)

        ElseIf oUID = "AOrdSub1" Then
        
            sQry = "        SELECT  U_JakMyung,"
            sQry = sQry & "         U_JakSize"
            sQry = sQry & " FROM    [@PS_PP030H]"
            sQry = sQry & " WHERE   U_OrdNum = '" & Trim(oForm01.Items("AOrdNum").Specific.VALUE) & "'"
            sQry = sQry & "         AND U_OrdSub1 = '" & Trim(oForm01.Items("AOrdSub1").Specific.VALUE) & "'"
            sQry = sQry & "         AND U_OrdSub2 = '000'"
            Call oRecordSet.DoQuery(sQry)

            'oForm01.Items("AOrdSub1").Specific.VALUE = "00" 'Trim(oRecordSet.Fields(0).VALUE)
            oForm01.Items("AOrdSub2").Specific.VALUE = "000"
            oForm01.Items("AItemName").Specific.VALUE = Trim(oRecordSet.Fields(0).VALUE)
            oForm01.Items("AItemSpec").Specific.VALUE = Trim(oRecordSet.Fields(1).VALUE)
            
        ElseIf oUID = "AOrdSub2" Then
        
            sQry = "        SELECT  U_JakMyung,"
            sQry = sQry & "         U_JakSize"
            sQry = sQry & " FROM    [@PS_PP030H]"
            sQry = sQry & " WHERE   U_OrdNum = '" & Trim(oForm01.Items("AOrdNum").Specific.VALUE) & "'"
            sQry = sQry & "         AND U_OrdSub1 = '" & Trim(oForm01.Items("AOrdSub1").Specific.VALUE) & "'"
            sQry = sQry & "         AND U_OrdSub2 = '" & Trim(oForm01.Items("AOrdSub2").Specific.VALUE) & "'"
            Call oRecordSet.DoQuery(sQry)

            'oForm01.Items("AOrdSub1").Specific.VALUE = "00" 'Trim(oRecordSet.Fields(0).VALUE)
            'oForm01.Items("AOrdSub2").Specific.VALUE = "000"
            oForm01.Items("AItemName").Specific.VALUE = Trim(oRecordSet.Fields(0).VALUE)
            oForm01.Items("AItemSpec").Specific.VALUE = Trim(oRecordSet.Fields(1).VALUE)

        End If
    
    End If

    Call oForm01.Freeze(False)

    Set oRecordSet = Nothing
    
    Exit Sub

FlushToItemValue_Error:
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function DataValidCheck() As Boolean
On Error GoTo DataValidCheck_Error

    Dim ErrNum    As Integer
    Dim loopCount As Integer
    
    ErrNum = 0

    'Header Check
    Select Case True
'        Case Trim(oDS_PS_PP570H.GetValue("U_BPLId", 0)) = ""
'            ErrNum = 1
'            GoTo DataValidCheck_Error
        Case Trim(oForm01.Items("TCntcCode").Specific.VALUE) = "" '등록자 누락
            ErrNum = 1
            GoTo DataValidCheck_Error
        Case Trim(oForm01.Items("RCntcCode").Specific.VALUE) = "" '요청자 누락
            ErrNum = 2
            GoTo DataValidCheck_Error
    End Select

    'Line Check
    Call oMat01.FlushToDataSource
    For loopCount = 0 To oMat01.VisualRowCount - 1
    
        If oMat01.Columns("Check").Cells(loopCount + 1).Specific.Checked = True Then '체크 확인 로직 변경(FlushDataSource 후 에도 Check box 값을 못 가져옴, Matrix의 값을 직접 가져옴)
    
            Select Case True
                Case Trim(oDS_PS_PP570L.GetValue("U_ColReg18", loopCount)) = ""  '이관후작번
                    ErrNum = 50
                    GoTo DataValidCheck_Error
                Case Trim(oDS_PS_PP570L.GetValue("U_ColReg19", loopCount)) = "" '이관후서브작번1
                    ErrNum = 51
                    GoTo DataValidCheck_Error
                Case Trim(oDS_PS_PP570L.GetValue("U_ColReg20", loopCount)) = "" '이관후서브작번2
                    ErrNum = 52
                    GoTo DataValidCheck_Error
            End Select
    
        End If
    
    Next
    

    DataValidCheck = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
DataValidCheck_Error:
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("등록자정보는 필수사항입니다. 확인하여 주십시오.", "E")
    ElseIf ErrNum = 2 Then
        Call MDC_Com.MDC_GF_Message("요청자정보는 필수사항입니다. 확인하여 주십시오.", "E")
    ElseIf ErrNum = 50 Then
        Call MDC_Com.MDC_GF_Message("이관후 작번은 필수사항입니다. 확인하여 주십시오.", "E")
    ElseIf ErrNum = 51 Then
        Call MDC_Com.MDC_GF_Message("이관후 서브작번1은 필수사항입니다. 확인하여 주십시오.", "E")
    ElseIf ErrNum = 52 Then
        Call MDC_Com.MDC_GF_Message("이관후 서브작번2는 필수사항입니다. 확인하여 주십시오.", "E")
    Else
        Call MDC_Com.MDC_GF_Message("DataValidCheck_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
    DataValidCheck = False
End Function

Public Sub PS_PP570_MTX01()
'******************************************************************************
'Function ID : PS_PP570_MTX01()
'해당모듈    : PS_PP570
'기능        : 데이터 조회
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP570_MTX01_Error

    Dim i      As Integer
    Dim sQry   As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

'    Dim BPLID As String '사업장
    Dim BOrdNum  As String '이관전작번
    Dim BOrdSub1 As String '이관전서브1
    Dim BOrdSub2 As String '이관전서브2
    
    BOrdNum = Trim(oForm01.Items("BOrdNum").Specific.VALUE) '이관전작번
    BOrdSub1 = Trim(oForm01.Items("BOrdSub1").Specific.VALUE) '이관전서브1
    BOrdSub2 = Trim(oForm01.Items("BOrdSub2").Specific.VALUE) '이관전서브2

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "      EXEC [PS_PP570_01] '"
    sQry = sQry & BOrdNum & "','"
    sQry = sQry & BOrdSub1 & "','"
    sQry = sQry & BOrdSub2 & "'"
    
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_PP570L.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
'        Call PS_PP570_Add_MatrixRow(0, True)
        GoTo PS_PP570_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_PP570L.Size Then
            oDS_PS_PP570L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_PP570L.offset = i
        
        Call oDS_PS_PP570L.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_PP570L.setValue("U_ColReg01", i, Trim(oRecordSet01.Fields("Check").VALUE)) '선택
        Call oDS_PS_PP570L.setValue("U_ColReg02", i, Trim(oRecordSet01.Fields("CpCode").VALUE)) '공정코드
        Call oDS_PS_PP570L.setValue("U_ColReg03", i, Trim(oRecordSet01.Fields("CpName").VALUE)) '공정명
        Call oDS_PS_PP570L.setValue("U_ColDt01", i, Format(Trim(oRecordSet01.Fields("WorkDate").VALUE), "YYYYMMDD")) '작업일자
        Call oDS_PS_PP570L.setValue("U_ColReg05", i, Trim(oRecordSet01.Fields("WorkCode").VALUE)) '작업자사번
        Call oDS_PS_PP570L.setValue("U_ColReg06", i, Trim(oRecordSet01.Fields("WorkName").VALUE)) '작업자성명
        Call oDS_PS_PP570L.setValue("U_ColQty01", i, Trim(oRecordSet01.Fields("WorkTime").VALUE)) '공수
        Call oDS_PS_PP570L.setValue("U_ColReg08", i, Trim(oRecordSet01.Fields("PP040HNo").VALUE)) '작업일보번호
        Call oDS_PS_PP570L.setValue("U_ColReg09", i, Trim(oRecordSet01.Fields("PP040LNo").VALUE)) '작업일보라인
        Call oDS_PS_PP570L.setValue("U_ColReg10", i, Trim(oRecordSet01.Fields("BOrdNum").VALUE)) '이관전작번
        Call oDS_PS_PP570L.setValue("U_ColReg11", i, Trim(oRecordSet01.Fields("BOrdSub1").VALUE)) '이관전서브1
        Call oDS_PS_PP570L.setValue("U_ColReg12", i, Trim(oRecordSet01.Fields("BOrdSub2").VALUE)) '이관전서브2
        Call oDS_PS_PP570L.setValue("U_ColReg13", i, Trim(oRecordSet01.Fields("BItemName").VALUE)) '이관전품명
        Call oDS_PS_PP570L.setValue("U_ColReg14", i, Trim(oRecordSet01.Fields("BItemSpec").VALUE)) '이관전규격
        Call oDS_PS_PP570L.setValue("U_ColReg15", i, Trim(oRecordSet01.Fields("BPP030HNo").VALUE)) '이관전작지번호
        Call oDS_PS_PP570L.setValue("U_ColReg16", i, Trim(oRecordSet01.Fields("BPP030MNo").VALUE)) '이관전작지라인
        Call oDS_PS_PP570L.setValue("U_ColReg17", i, Trim(oRecordSet01.Fields("BPP030Seq").VALUE)) '이관전작업순서
        Call oDS_PS_PP570L.setValue("U_ColReg18", i, Trim(oRecordSet01.Fields("AOrdNum").VALUE)) '이관후작번
        Call oDS_PS_PP570L.setValue("U_ColReg19", i, Trim(oRecordSet01.Fields("AOrdSub1").VALUE)) '이관후서브1
        Call oDS_PS_PP570L.setValue("U_ColReg20", i, Trim(oRecordSet01.Fields("AOrdSub2").VALUE)) '이관후서브2
        Call oDS_PS_PP570L.setValue("U_ColReg21", i, Trim(oRecordSet01.Fields("AItemName").VALUE)) '이관후품명
        Call oDS_PS_PP570L.setValue("U_ColReg22", i, Trim(oRecordSet01.Fields("AItemSpec").VALUE)) '이관후규격
    
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
PS_PP570_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP570_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

Public Function PS_PP570_AddData() As Boolean
'******************************************************************************
'Function ID : PS_PP570_AddData()
'해당모듈    : PS_PP570
'기능        : 데이터 INSERT(UPDATE)
'인수        : 없음
'반환값      : 성공여부
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP570_AddData_Error

    Dim loopCount    As Integer
    Dim sQry         As String
    Dim ErrorMessage As String
    
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim CpCode    As String '공정코드
    Dim BOrdNum   As String '이관전작번
    Dim BOrdSub1  As String '이관전서브작번1
    Dim BOrdSub2  As String '이관전서브작번2
    Dim BPP030HNo As String '이관전작지문서번호
    Dim BPP030MNo As String '이관전작지라인
    Dim BPP030Seq As String '이관전작업순서
    
    Dim AOrdNum   As String '이관후작번
    Dim AOrdSub1  As String '이관후서브작번1
    Dim AOrdSub2  As String '이관후서브작번2
    
    Dim PP040HNo  As String '작업일보문서번호
    Dim PP040MNo  As String '작업일보라인번호
    
    Dim TDocDate  As String '이관일자
    Dim TCntcCode As String '등록자사번
    Dim TCntcName As String '등록자성명
    Dim RCntcCode As String '요청자사번
    Dim RCntcName As String '요청사성명
    Dim TransCls  As String '이관구분
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("저장중...", RecordSet01.RecordCount, False)

'    Screen.MousePointer = vbHourglass

    TDocDate = Format(Now, "YYYYMMDD") '이관일자
    TCntcCode = oForm01.Items("TCntcCode").Specific.VALUE '등록자사번
    TCntcName = oForm01.Items("TCntcName").Specific.VALUE '등록자성명
    RCntcCode = oForm01.Items("RCntcCode").Specific.VALUE '요청자사번
    RCntcName = oForm01.Items("RCntcName").Specific.VALUE '요청자성명
    TransCls = oForm01.Items("TransCls").Specific.VALUE '이관구분
        
    Call oMat01.FlushToDataSource
    For loopCount = 0 To oMat01.VisualRowCount - 1
    
'        If Trim(oDS_PS_PP570L.GetValue("U_ColReg01", loopCount)) = "Y" Then
        
        If oMat01.Columns("Check").Cells(loopCount + 1).Specific.Checked = True Then '체크 확인 로직 변경(FlushDataSource 후 에도 Check box 값을 못 가져옴, Matrix의 값을 직접 가져옴)
        
            CpCode = Trim(oDS_PS_PP570L.GetValue("U_ColReg02", loopCount)) '공정코드
            BOrdNum = Trim(oDS_PS_PP570L.GetValue("U_ColReg10", loopCount)) '이관전작번
            BOrdSub1 = Trim(oDS_PS_PP570L.GetValue("U_ColReg11", loopCount)) '이관전서브작번1
            BOrdSub2 = Trim(oDS_PS_PP570L.GetValue("U_ColReg12", loopCount)) '이관전서브작번2
            BPP030HNo = Trim(oDS_PS_PP570L.GetValue("U_ColReg15", loopCount)) '이관전작지문서번호
            BPP030MNo = Trim(oDS_PS_PP570L.GetValue("U_ColReg16", loopCount)) '이관전작지라인
            BPP030Seq = Trim(oDS_PS_PP570L.GetValue("U_ColReg17", loopCount)) '이관전작업순서
            
            AOrdNum = Trim(oDS_PS_PP570L.GetValue("U_ColReg18", loopCount)) '이관후작번
            AOrdSub1 = Trim(oDS_PS_PP570L.GetValue("U_ColReg19", loopCount)) '이관후서브작번1
            AOrdSub2 = Trim(oDS_PS_PP570L.GetValue("U_ColReg20", loopCount)) '이관후서브작번2
            
            PP040HNo = Trim(oDS_PS_PP570L.GetValue("U_ColReg08", loopCount)) '작업일보문서번호
            PP040MNo = Trim(oDS_PS_PP570L.GetValue("U_ColReg09", loopCount))  '작업일보라인번호
            
            If TransCls = "01" Then '비용전체
            
                sQry = "      EXEC [PS_PP570_02] '"
                sQry = sQry & CpCode & "','" '공정코드
                sQry = sQry & BOrdNum & "','" '이관전작번
                sQry = sQry & BOrdSub1 & "','" '이관전서브작번1
                sQry = sQry & BOrdSub2 & "','" '이관전서브작번2
                sQry = sQry & BPP030HNo & "','" '이관전작지문서번호
                sQry = sQry & BPP030MNo & "','" '이관전작지라인
                sQry = sQry & BPP030Seq & "','" '이관전작업순서
                sQry = sQry & AOrdNum & "','" '이관후작번
                sQry = sQry & AOrdSub1 & "','" '이관후서브작번1
                sQry = sQry & AOrdSub2 & "','" '이관후서브작번2
                sQry = sQry & PP040HNo & "','" '작업일보문서번호
                sQry = sQry & PP040MNo & "','" '작업일보라인번호
                sQry = sQry & TDocDate & "','" '이관일자
                sQry = sQry & TCntcCode & "','" '등록자사번
                sQry = sQry & TCntcName & "','" '등록사성명
                sQry = sQry & RCntcCode & "','" '요청자사번
                sQry = sQry & RCntcName & "'" '요청자성명
            
            Else '공수
            
                sQry = "      EXEC [PS_PP570_03] '"
                sQry = sQry & CpCode & "','" '공정코드
                sQry = sQry & BOrdNum & "','" '이관전작번
                sQry = sQry & BOrdSub1 & "','" '이관전서브작번1
                sQry = sQry & BOrdSub2 & "','" '이관전서브작번2
                sQry = sQry & BPP030HNo & "','" '이관전작지문서번호
                sQry = sQry & BPP030MNo & "','" '이관전작지라인
                sQry = sQry & BPP030Seq & "','" '이관전작업순서
                sQry = sQry & AOrdNum & "','" '이관후작번
                sQry = sQry & AOrdSub1 & "','" '이관후서브작번1
                sQry = sQry & AOrdSub2 & "','" '이관후서브작번2
                sQry = sQry & PP040HNo & "','" '작업일보문서번호
                sQry = sQry & PP040MNo & "','" '작업일보라인번호
                sQry = sQry & TDocDate & "','" '이관일자
                sQry = sQry & TCntcCode & "','" '등록자사번
                sQry = sQry & TCntcName & "','" '등록사성명
                sQry = sQry & RCntcCode & "','" '요청자사번
                sQry = sQry & RCntcName & "'" '요청자성명
            
            End If
            
            Call RecordSet01.DoQuery(sQry)
    
            If RecordSet01.Fields("ErrorMessage").VALUE <> "True" Then '오류 발생 시 For문 종료
                ErrorMessage = RecordSet01.Fields("ErrorMessage").VALUE
                GoTo PS_PP570_AddData_Error
            End If
    
        End If
        
        ProgBar01.VALUE = ProgBar01.VALUE + 1
'        ProgBar01.Text = ProgBar01.VALUE & "/" & oMat01.VisualRowCount & "건 저장중...!"
        
    Next
    
'    Screen.MousePointer = vbDefault
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set RecordSet01 = Nothing
    
    Call MDC_Com.MDC_GF_Message("등록 완료!", "S")
    
    PS_PP570_AddData = True

    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP570_AddData_Error:

'     Screen.MousePointer = vbDefault

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set RecordSet01 = Nothing
    
    If ErrorMessage <> "" Then
        Call MDC_Com.MDC_GF_Message(loopCount + 1 & "행 " & ErrorMessage, "E")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP570_AddData_Error : " & Err.Number & " - " & Err.Description, "E")
    End If
    
    PS_PP570_AddData = False
    
End Function

Private Sub PS_PP570_Print_Report01()

    Dim DocNum     As String
    Dim ErrNum     As Integer
    Dim WinTitle   As String
    Dim ReportName As String
    Dim sQry       As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PS_PP570_Print_Report01_Error
    
    Dim BPLID    As String '사업장
    Dim StdYear  As String '기준년도
    Dim StdYM    As String '기준년월
    Dim FixType  As String '자산분류
    Dim ChkDisu  As String '폐기대상 조회
    
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

    WinTitle = "[PS_PP570] 고정자산 재물조사"
    
    If BPLID = "1" Then '창원
        ReportName = "PS_PP570_01.rpt"
    ElseIf BPLID = "2" Then '동래
        ReportName = "PS_PP570_02.rpt"
    ElseIf BPLID = "3" Then '사상
        ReportName = "PS_PP570_03.rpt"
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
    
        sQry = "      EXEC [PS_PP570_51] '"
        sQry = sQry & BPLID & "','"
        sQry = sQry & StdYear & "','"
        sQry = sQry & StdYM & "','"
        sQry = sQry & FixType & "','"
        sQry = sQry & ChkDisu & "'"
    
'    Else '폐기대상 조회 체크
'
'        sQry = "         EXEC [PS_PP570_52] '"
'        sQry = sQry & BPLID & "','"
'        sQry = sQry & StdYear & "','"
'        sQry = sQry & FixType & "'"
'
'    End If
    
'    oRecordSet.DoQuery sQry
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo PS_PP570_Print_Report01_Error
'    End If
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
    
PS_PP570_Print_Report01_Error:
'    If ErrNum = 1 Then
'        Set oRecordSet = Nothing
'        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
'    Else
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PS_PP570_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'    End If
    
End Sub

Private Sub PS_PP570_CheckAll()
On Error GoTo PS_PP570_CheckAll_Error

    Dim CheckType As String
    Dim loopCount As Integer
    
    Call oForm01.Freeze(True)
    CheckType = "Y"
    
    Call oMat01.FlushToDataSource
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_PP570L.GetValue("U_ColReg01", loopCount)) = "N" Then
            CheckType = "N"
            Exit For
        End If
    Next
    
    For loopCount = 0 To oMat01.VisualRowCount - 1
        oDS_PS_PP570L.offset = loopCount
        If CheckType = "N" Then
            Call oDS_PS_PP570L.setValue("U_ColReg01", loopCount, "Y")
        Else
            Call oDS_PS_PP570L.setValue("U_ColReg01", loopCount, "N")
        End If
    Next
    
    Call oMat01.LoadFromDataSource
    
    Call oForm01.Freeze(False)

    Exit Sub
    
PS_PP570_CheckAll_Error:
    Call oForm01.Freeze(False)
    Call Sbo_Application.SetStatusBarMessage("PS_PP570_CheckAll_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Sub

'/메트릭스 Row추가
Sub PS_PP570_Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo PS_PP570_Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_PP570L.InsertRecord (oRow)
    End If
    
    oMat01.AddRow
    oDS_PS_PP570L.offset = oRow
    oDS_PS_PP570L.setValue "U_LineNum", oRow, oRow + 1
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PS_PP570_Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "PS_PP570_Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub PS_PP570_Update_AfterOrdNum()
'******************************************************************************
'Function ID : PS_PP570_Update_AfterOrdNum()
'해당모듈    : PS_PP570
'기능        : 이관후 작번을 매트릭스에 적용
'인수        : 없음
'반환값      : 없음
'특이사항    : 없음
'******************************************************************************
On Error GoTo PS_PP570_Update_AfterOrdNum_Error

    Dim loopCount As Integer

    Dim AOrdNum   As String '이관후 작번
    Dim AOrdSub1  As String '이관후 서브1
    Dim AOrdSub2  As String '이관후 서브2
    Dim AItemName As String '이관후 작명
    Dim AItemSpec As String '이관후 규격
    
    AOrdNum = oForm01.Items("AOrdNum").Specific.VALUE
    AOrdSub1 = oForm01.Items("AOrdSub1").Specific.VALUE
    AOrdSub2 = oForm01.Items("AOrdSub2").Specific.VALUE
    AItemName = oForm01.Items("AItemName").Specific.VALUE
    AItemSpec = oForm01.Items("AItemSpec").Specific.VALUE
    
    Call oForm01.Freeze(True)
    
    Call oMat01.FlushToDataSource
    For loopCount = 0 To oMat01.VisualRowCount - 1
    
'        If Trim(oDS_PS_PP570L.GetValue("U_ColReg01", loopCount)) = "Y" Then
        
        If oMat01.Columns("Check").Cells(loopCount + 1).Specific.Checked = True Then '체크 확인 로직 변경(FlushDataSource 후 에도 Check box 값을 못 가져옴, Matrix의 값을 직접 가져옴)
        
            Call oDS_PS_PP570L.setValue("U_ColReg18", loopCount, AOrdNum)
            Call oDS_PS_PP570L.setValue("U_ColReg19", loopCount, AOrdSub1)
            Call oDS_PS_PP570L.setValue("U_ColReg20", loopCount, AOrdSub2)
            Call oDS_PS_PP570L.setValue("U_ColReg21", loopCount, AItemName)
            Call oDS_PS_PP570L.setValue("U_ColReg22", loopCount, AItemSpec)
    
        End If
        
    Next
    
    Call oMat01.LoadFromDataSource
    Call oMat01.AutoResizeColumns

    Call oForm01.Freeze(False)


'    Dim CheckType As String
'    Dim loopCount As Integer
'
'    Call oForm01.Freeze(True)
'    CheckType = "Y"
'
'    Call oMat01.FlushToDataSource
'
'    For loopCount = 0 To oMat01.VisualRowCount - 1
'        If Trim(oDS_PS_PP570L.GetValue("U_ColReg01", loopCount)) = "N" Then
'            CheckType = "N"
'            Exit For
'        End If
'    Next
'
'    For loopCount = 0 To oMat01.VisualRowCount - 1
'        oDS_PS_PP570L.Offset = loopCount
'        If CheckType = "N" Then
'            Call oDS_PS_PP570L.setValue("U_ColReg01", loopCount, "Y")
'        Else
'            Call oDS_PS_PP570L.setValue("U_ColReg01", loopCount, "N")
'        End If
'    Next
'
'    Call oMat01.LoadFromDataSource
'
'    Call oForm01.Freeze(False)

    Exit Sub
    
PS_PP570_Update_AfterOrdNum_Error:
    Call oForm01.Freeze(False)
    Call Sbo_Application.SetStatusBarMessage("PS_PP570_Update_AfterOrdNum_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Sub
