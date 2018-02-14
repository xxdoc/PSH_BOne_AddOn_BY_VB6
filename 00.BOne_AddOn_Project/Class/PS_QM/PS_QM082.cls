VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM082"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_QM082.cls
'//  Module         : QM
'//  Description    : 방산부품 원재료 입고 등록
'//  FormType       : PS_QM082
'//  Create Date    : 2013.01.26
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_QM082H       As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_QM082L       As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private Last_CardCode$, Last_CardName$, Last_InspPrsn$

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM082.srf")
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
    
    oFormUniqueID01 = "PS_QM082_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "DocEntry"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call FormClear
'    Call Matrix_AddRow(1, 0, True) 'oMat01
'    FormItemEnabled
    
    oForm01.EnableMenu ("1283"), True         '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), True         '// 복제
    oForm01.EnableMenu ("1284"), False        '// 취소
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

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//디비데이터 소스 개체 할당
    Set oDS_PS_QM082H = oForm01.DataSources.DBDataSources("@PS_QM082H")
'    Set oDS_PS_QM082L = oForm01.DataSources.DBDataSources("@PS_QM082L")
    
    '// 메트릭스 개체 할당
'    Set oMat01 = oForm01.Items("Mat01").Specific

    oDS_PS_QM082H.setValue "U_DocDate", 0, Format(Now, "YYYYMMDD")
   
    'oForm01.Items("DocEntry").Visible = True
    'oForm01.Items("Code").Visible = False
   
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CreateItems_Error:
    MDC_Com.MDC_GF_Message "CreateItems_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    '//콤보에 기본값설정
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet      As SAPbobsCOM.Recordset
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// 납품완료여부(Y/N)
    Set oCombo = oForm01.Items("FinishYN").Specific
    oCombo.ValidValues.Add "N", "납품미완료(N)"
    oCombo.ValidValues.Add "Y", "납품완료(Y)"

    oCombo.Select "N", psk_ByValue
    
    '// 검사완료여부(Y/N)
    Set oCombo = oForm01.Items("TestYN").Specific
    oCombo.ValidValues.Add "N", "검사미완료(N)"
    oCombo.ValidValues.Add "Y", "검사완료(Y)"

    oCombo.Select "N", psk_ByValue
    
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet = Nothing
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
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
'                        If MatrixSpaceLineDel = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
                        If oForm01.Mode = fm_ADD_MODE Then
                            If Check_Exist = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    '//헤더
                    If pval.ItemUID = "BaseCode" Then
                        If oForm01.Items("BaseCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    '//라인
'                    If pval.ItemUID = "Mat01" Then
'                        If pval.ColUID = "PP070No" Then
'                            If oMat01.Columns("PP070No").Cells(pval.Row).Specific.Value = "" Then
'                                Sbo_Application.ActivateMenuItem ("7425")
'                                BubbleEvent = False
'                            End If
'                        End If
'                    End If
                End If
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
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
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                    ElseIf oForm01.Mode = fm_OK_MODE Then
                        FormItemEnabled
                        FlushToItemValue "BaseCode"
'                        Call Matrix_AddRow(1, oMat01.RowCount, False) 'oMat01
                    End If
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    '//헤더
                    If pval.ItemUID = "BaseCode" Then
                        FlushToItemValue pval.ItemUID
                    End If
                    '//라인
'                    If pval.ItemUID = "Mat01" And (pval.ColUID = "PP070No") Then
'                        Call FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
'                    End If
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
                Set oMat01 = Nothing
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
        End Select
        
    '-----------------------------------------------------------------------------------------------------------
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1285": '복원
            Case "1287": '복제
                
            Case "1293": '행삭제
'
            Case "1281": '찾기
                FormItemEnabled
                oForm01.Items("LotNo").Click ct_Regular
                
            Case "1282": '추가
                Call FormItemEnabled
                Call FormClear
                oDS_PS_QM082H.setValue "U_DocDate", 0, Format(Now, "YYYYMMDD")
                oForm01.Items("LotNo").Click ct_Regular
                
                
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                FormItemEnabled
                FlushToItemValue "BaseCode"
'                If oMat01.VisualRowCount > 0 Then
'                    If oMat01.Columns("PP070No").Cells(oMat01.VisualRowCount).Specific.Value <> "" Then
'                        If oDS_PS_QM082H.GetValue("Status", 0) = "O" Then
'                            Matrix_AddRow 1, oMat01.RowCount, False
'                        End If
'                    End If
'                End If
  
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

Private Sub FormItemEnabled()
    If oForm01.Mode = fm_FIND_MODE Then
        oForm01.Items("BaseCode").Enabled = True
        oForm01.Items("LotNo").Enabled = True
        
    ElseIf oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("BaseCode").Enabled = True
        oForm01.Items("LotNo").Enabled = True
        
    ElseIf oForm01.Mode = fm_OK_MODE Then
        oForm01.Items("BaseCode").Enabled = False
        oForm01.Items("LotNo").Enabled = False
    End If
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM082'", "")
    If DocNum = 0 Then
'        oForm01.Items("Code").Specific.Value = 1
        oDS_PS_QM082H.setValue "Code", 0, 1
    Else
'        oForm01.Items("Code").Specific.Value = DocNum
        oDS_PS_QM082H.setValue "Code", 0, DocNum
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormClear_Error:
    MDC_Com.MDC_GF_Message "FormClear_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&, DocNum$, LineId$
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet    As SAPbobsCOM.Recordset
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    '--------------------------------------------------------------
    'Header--------------------------------------------------------
    Select Case oUID
        Case "BaseCode"
            '원재료 사양서
'            sQry = "select cardname from ocrd where cardtype='C' and cardcode = '" & Trim(oDS_PS_QM082H.GetValue("U_BaseCode", 0)) & "'"
'            oRecordSet.DoQuery sQry
'            oDS_PS_QM082H.setValue "U_CardName", 0, Trim(oRecordSet.Fields(0).VALUE)
'
            '검사사양서내역
            sQry = "select U_ItemCode, U_ItemName, U_CItemCod, U_CItemNam, U_StdNum,"
            sQry = sQry + " U_C_min, U_C_max, U_Mn_min, U_Mn_max, U_P_min, U_S_min, U_Si_min, U_Si_max, U_Ni_min, U_Ni_max,"
            sQry = sQry + " U_Cr_min, U_Cr_max, U_Mo_min, U_Mo_max, U_Fe_min, U_Cu_min, U_Cu_max, U_Mg_min, U_Mg_max, U_Zn_min,"
            sQry = sQry + " U_Ti_min, U_ZrTi_min, U_Other, U_Al, U_TS_min,U_TS_max, U_YS_min, U_YS_max, U_EL_min, U_EL_max,"
            sQry = sQry + " U_HARD_min, U_HARD_max, U_GS_min, U_GS_max, U_Clean, U_METCH,U_EACH, U_TOTAL"
            sQry = sQry & " From [@PS_QM081H] "
            sQry = sQry & " where Code = '" & Trim(oDS_PS_QM082H.GetValue("U_BaseCode", 0)) & "' "
            
            oRecordSet.DoQuery sQry

            oForm01.Items("ItemCode").Specific.VALUE = Trim(oRecordSet.Fields(0).VALUE)
            oForm01.Items("ItemName").Specific.VALUE = Trim(oRecordSet.Fields(1).VALUE)
            oForm01.Items("CItemCod").Specific.VALUE = Trim(oRecordSet.Fields(2).VALUE)
            oForm01.Items("CItemNam").Specific.VALUE = Trim(oRecordSet.Fields(3).VALUE)
            oForm01.Items("StdNum").Specific.VALUE = Trim(oRecordSet.Fields(4).VALUE)
            oForm01.Items("C_min").Specific.VALUE = Trim(oRecordSet.Fields(5).VALUE)
            oForm01.Items("C_max").Specific.VALUE = Trim(oRecordSet.Fields(6).VALUE)
            oForm01.Items("Mn_min").Specific.VALUE = Trim(oRecordSet.Fields(7).VALUE)
            oForm01.Items("Mn_max").Specific.VALUE = Trim(oRecordSet.Fields(8).VALUE)
            oForm01.Items("P_min").Specific.VALUE = Trim(oRecordSet.Fields(9).VALUE)
            oForm01.Items("S_min").Specific.VALUE = Trim(oRecordSet.Fields(10).VALUE)
            oForm01.Items("Si_min").Specific.VALUE = Trim(oRecordSet.Fields(11).VALUE)
            oForm01.Items("Si_max").Specific.VALUE = Trim(oRecordSet.Fields(12).VALUE)
            oForm01.Items("Ni_min").Specific.VALUE = Trim(oRecordSet.Fields(13).VALUE)
            oForm01.Items("Ni_max").Specific.VALUE = Trim(oRecordSet.Fields(14).VALUE)
            oForm01.Items("Cr_min").Specific.VALUE = Trim(oRecordSet.Fields(15).VALUE)
            oForm01.Items("Cr_max").Specific.VALUE = Trim(oRecordSet.Fields(16).VALUE)
            oForm01.Items("Mo_min").Specific.VALUE = Trim(oRecordSet.Fields(17).VALUE)
            oForm01.Items("Mo_max").Specific.VALUE = Trim(oRecordSet.Fields(18).VALUE)
            oForm01.Items("Fe_min").Specific.VALUE = Trim(oRecordSet.Fields(19).VALUE)
            oForm01.Items("Cu_min").Specific.VALUE = Trim(oRecordSet.Fields(20).VALUE)
            oForm01.Items("Cu_max").Specific.VALUE = Trim(oRecordSet.Fields(21).VALUE)
            oForm01.Items("Mg_min").Specific.VALUE = Trim(oRecordSet.Fields(22).VALUE)
            oForm01.Items("Mg_max").Specific.VALUE = Trim(oRecordSet.Fields(23).VALUE)
            oForm01.Items("Zn_min").Specific.VALUE = Trim(oRecordSet.Fields(24).VALUE)
            oForm01.Items("Ti_min").Specific.VALUE = Trim(oRecordSet.Fields(25).VALUE)
            oForm01.Items("ZrTi_min").Specific.VALUE = Trim(oRecordSet.Fields(26).VALUE)
            oForm01.Items("Other").Specific.VALUE = Trim(oRecordSet.Fields(27).VALUE)
            oForm01.Items("Al").Specific.VALUE = Trim(oRecordSet.Fields(28).VALUE)
            oForm01.Items("EACH").Specific.VALUE = Trim(oRecordSet.Fields(41).VALUE)
            oForm01.Items("TOTAL").Specific.VALUE = Trim(oRecordSet.Fields(42).VALUE)
          
            oForm01.Items("TS_min").Specific.VALUE = Trim(oRecordSet.Fields(29).VALUE)
            oForm01.Items("TS_max").Specific.VALUE = Trim(oRecordSet.Fields(30).VALUE)
            oForm01.Items("YS_min").Specific.VALUE = Trim(oRecordSet.Fields(31).VALUE)
            oForm01.Items("YS_max").Specific.VALUE = Trim(oRecordSet.Fields(32).VALUE)
            oForm01.Items("EL_min").Specific.VALUE = Trim(oRecordSet.Fields(33).VALUE)
            oForm01.Items("EL_max").Specific.VALUE = Trim(oRecordSet.Fields(34).VALUE)
            oForm01.Items("HARD_min").Specific.VALUE = Trim(oRecordSet.Fields(35).VALUE)
            oForm01.Items("HARD_max").Specific.VALUE = Trim(oRecordSet.Fields(36).VALUE)
            oForm01.Items("GS_min").Specific.VALUE = Trim(oRecordSet.Fields(37).VALUE)
            oForm01.Items("GS_max").Specific.VALUE = Trim(oRecordSet.Fields(38).VALUE)
            oForm01.Items("Clean").Specific.VALUE = Trim(oRecordSet.Fields(39).VALUE)
            oForm01.Items("METCH").Specific.VALUE = Trim(oRecordSet.Fields(40).VALUE)
            
       
            
    End Select
        
    
     
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    Dim SPEC_MIN        As Currency
    Dim SPEC_MAX        As Currency
    Dim VALUE           As Currency
    Dim VALUE_MIN       As Currency
    Dim VALUE_MAX       As Currency
    
    ErrNum = 0

    '// Check
    Select Case True
        Case Trim(oDS_PS_QM082H.GetValue("U_LotNo", 0)) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_QM082H.GetValue("U_BaseCode", 0)) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_QM082H.GetValue("U_DocDate", 0)) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_QM082H.GetValue("U_Qty", 0)) = ""
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_QM082H.GetValue("U_Weight", 0)) = ""
            ErrNum = 5
            GoTo HeaderSpaceLineDel_Error
    End Select

'    '--------------------------------------------------------------------------------------------
'    '//측청치의 유효값 검사
'    '//인강강도 Min
'    If Val(oDS_PS_QM082H.GetValue("U_M_TS_min", 0)) <> 0 Then
'        VALUE_MIN = Val(oDS_PS_QM082H.GetValue("U_M_TS_min", 0))
'        SPEC_MIN = Val(oForm01.Items("TS_min").Specific.VALUE)
'        SPEC_MAX = Val(oForm01.Items("TS_max").Specific.VALUE)
'        If VALUE_MIN < SPEC_MIN Or VALUE_MIN > SPEC_MAX Then
'            ErrNum = 10
'            oForm01.Items("M_TS_min").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'
'    '//인강강도 Max
'    If Val(oDS_PS_QM082H.GetValue("U_M_TS_max", 0)) <> 0 Then
'        VALUE_MAX = Val(oDS_PS_QM082H.GetValue("U_M_TS_max", 0))
'        SPEC_MIN = Val(oForm01.Items("TS_min").Specific.VALUE)
'        SPEC_MAX = Val(oForm01.Items("TS_max").Specific.VALUE)
'        If VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
'            ErrNum = 11
'            oForm01.Items("M_TS_max").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//연신율 Min
'    If Val(oDS_PS_QM082H.GetValue("U_M_El_min", 0)) <> 0 Then
'        VALUE_MIN = Val(oDS_PS_QM082H.GetValue("U_M_El_min", 0))
'        SPEC_MIN = Val(oForm01.Items("Elong").Specific.VALUE)
'        If VALUE_MIN < SPEC_MIN Then
'            ErrNum = 12
'            oForm01.Items("M_El_min").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//연신율 Max
'    If Val(oDS_PS_QM082H.GetValue("U_M_El_max", 0)) <> 0 Then
'        VALUE_MAX = Val(oDS_PS_QM082H.GetValue("U_M_El_max", 0))
'        SPEC_MIN = Val(oForm01.Items("Elong").Specific.VALUE)
'        If VALUE_MAX < SPEC_MIN Then
'            ErrNum = 13
'            oForm01.Items("M_El_max").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//경도 Min
'    If Val(oDS_PS_QM082H.GetValue("U_M_Hd_min", 0)) <> 0 Then
'        VALUE_MIN = Val(oDS_PS_QM082H.GetValue("U_M_Hd_min", 0))
'        SPEC_MIN = Val(oForm01.Items("Hard_min").Specific.VALUE)
'        SPEC_MAX = Val(oForm01.Items("Hard_max").Specific.VALUE)
'        If VALUE_MIN < SPEC_MIN Or VALUE_MIN > SPEC_MAX Then
'            ErrNum = 14
'            oForm01.Items("M_Hd_min").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//경도 Max
'    If Val(oDS_PS_QM082H.GetValue("U_M_Hd_max", 0)) <> 0 Then
'        VALUE_MAX = Val(oDS_PS_QM082H.GetValue("U_M_Hd_max", 0))
'        SPEC_MIN = Val(oForm01.Items("Hard_min").Specific.VALUE)
'        SPEC_MAX = Val(oForm01.Items("Hard_max").Specific.VALUE)
'        If VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
'            ErrNum = 15
'            oForm01.Items("M_Hd_max").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//Fe
'    If Val(oDS_PS_QM082H.GetValue("U_M_Ch_Fe", 0)) <> 0 Then
'        VALUE = Val(oDS_PS_QM082H.GetValue("U_M_Ch_Fe", 0))
'        SPEC_MIN = Val(oForm01.Items("C_Fe_min").Specific.VALUE)
'        SPEC_MAX = Val(oForm01.Items("C_Fe_max").Specific.VALUE)
'        If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
'            ErrNum = 16
'            oForm01.Items("M_Ch_Fe").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//P
'    If Val(oDS_PS_QM082H.GetValue("U_M_Ch_P", 0)) <> 0 Then
'        VALUE = Val(oDS_PS_QM082H.GetValue("U_M_Ch_P", 0))
'        SPEC_MIN = Val(oForm01.Items("C_P_min").Specific.VALUE)
'        SPEC_MAX = Val(oForm01.Items("C_P_max").Specific.VALUE)
'        If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
'            ErrNum = 17
'            oForm01.Items("M_Ch_P").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//측정항목A Min
'    If Val(oDS_PS_QM082H.GetValue("U_M_A_min", 0)) <> 0 Then
'        VALUE_MIN = Val(oDS_PS_QM082H.GetValue("U_M_A_min", 0))
'        SPEC_MIN = Val(oForm01.Items("A_Val").Specific.VALUE) - Val(oForm01.Items("A_Mi").Specific.VALUE)
'        SPEC_MAX = Val(oForm01.Items("A_Val").Specific.VALUE) + Val(oForm01.Items("A_Pl").Specific.VALUE)
'        If VALUE_MIN < SPEC_MIN Or VALUE_MIN > SPEC_MAX Then
'            ErrNum = 18
'            oForm01.Items("M_A_min").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//측정항목A Max
'    If Val(oDS_PS_QM082H.GetValue("U_M_A_max", 0)) <> 0 Then
'        VALUE_MAX = Val(oDS_PS_QM082H.GetValue("U_M_A_max", 0))
'        SPEC_MIN = Val(oForm01.Items("A_Val").Specific.VALUE) - Val(oForm01.Items("A_Mi").Specific.VALUE)
'        SPEC_MAX = Val(oForm01.Items("A_Val").Specific.VALUE) + Val(oForm01.Items("A_Pl").Specific.VALUE)
'        If VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
'            ErrNum = 19
'            oForm01.Items("M_A_max").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//측정항목B Min
'    If Val(oDS_PS_QM082H.GetValue("U_M_B_min", 0)) <> 0 Then
'        VALUE_MIN = Val(oDS_PS_QM082H.GetValue("U_M_B_min", 0))
'        SPEC_MIN = Val(oForm01.Items("B_Val").Specific.VALUE) - Val(oForm01.Items("B_Mi").Specific.VALUE)
'        SPEC_MAX = Val(oForm01.Items("B_Val").Specific.VALUE) + Val(oForm01.Items("B_Pl").Specific.VALUE)
'        If VALUE_MIN < SPEC_MIN Or VALUE_MIN > SPEC_MAX Then
'            ErrNum = 20
'            oForm01.Items("M_B_min").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//측정항목A Max
'    If Val(oDS_PS_QM082H.GetValue("U_M_B_max", 0)) <> 0 Then
'        VALUE_MAX = Val(oDS_PS_QM082H.GetValue("U_M_B_max", 0))
'        SPEC_MIN = Val(oForm01.Items("B_Val").Specific.VALUE) - Val(oForm01.Items("B_Mi").Specific.VALUE)
'        SPEC_MAX = Val(oForm01.Items("B_Val").Specific.VALUE) + Val(oForm01.Items("B_Pl").Specific.VALUE)
'        If VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
'            ErrNum = 21
'            oForm01.Items("M_B_max").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//측정항목B1 Min
'    If Val(oDS_PS_QM082H.GetValue("U_M_B1_min", 0)) <> 0 Then
'        VALUE_MIN = Val(oDS_PS_QM082H.GetValue("U_M_B1_min", 0))
'        SPEC_MIN = Val(oForm01.Items("B1_Val").Specific.VALUE) - Val(oForm01.Items("B1_Mi").Specific.VALUE)
'        SPEC_MAX = Val(oForm01.Items("B1_Val").Specific.VALUE) + Val(oForm01.Items("B1_Pl").Specific.VALUE)
'        If VALUE_MIN < SPEC_MIN Or VALUE_MIN > SPEC_MAX Then
'            ErrNum = 22
'            oForm01.Items("M_B1_min").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//측정항목B1 Max
'    If Val(oDS_PS_QM082H.GetValue("U_M_B1_max", 0)) <> 0 Then
'        VALUE_MAX = Val(oDS_PS_QM082H.GetValue("U_M_B1_max", 0))
'        SPEC_MIN = Val(oForm01.Items("B1_Val").Specific.VALUE) - Val(oForm01.Items("B1_Mi").Specific.VALUE)
'        SPEC_MAX = Val(oForm01.Items("B1_Val").Specific.VALUE) + Val(oForm01.Items("B1_Pl").Specific.VALUE)
'        If VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
'            ErrNum = 23
'            oForm01.Items("M_B1_max").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//측정항목C
'    If Val(oDS_PS_QM082H.GetValue("U_M_C", 0)) <> 0 Then
'        VALUE = Val(oDS_PS_QM082H.GetValue("U_M_C", 0))
'        SPEC_MIN = Val(oForm01.Items("C_Val").Specific.VALUE) - Val(oForm01.Items("C_Mi").Specific.VALUE)
'        SPEC_MAX = Val(oForm01.Items("C_Val").Specific.VALUE) + Val(oForm01.Items("C_Pl").Specific.VALUE)
'        If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
'            ErrNum = 24
'            oForm01.Items("M_C").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//측정항목D
'    If Val(oDS_PS_QM082H.GetValue("U_M_D", 0)) <> 0 Then
'        VALUE = Val(oDS_PS_QM082H.GetValue("U_M_D", 0))
'        SPEC_MIN = Val(oForm01.Items("D_Val").Specific.VALUE) - Val(oForm01.Items("D_Mi").Specific.VALUE)
'        SPEC_MAX = Val(oForm01.Items("D_Val").Specific.VALUE) + Val(oForm01.Items("D_Pl").Specific.VALUE)
'        If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
'            ErrNum = 25
'            oForm01.Items("M_D").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'
'    '//측정항목E-E1
'    If Val(oDS_PS_QM082H.GetValue("U_M_EE1", 0)) <> 0 Then
'        VALUE = Val(oDS_PS_QM082H.GetValue("U_M_EE1", 0))
'        SPEC_MAX = Val(oForm01.Items("EE1_Val").Specific.VALUE)
'        If VALUE > SPEC_MAX Then
'            ErrNum = 26
'            oForm01.Items("M_EE1").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//측정항목F
'    If Val(oDS_PS_QM082H.GetValue("U_M_F", 0)) <> 0 Then
'        VALUE = Val(oDS_PS_QM082H.GetValue("U_M_F", 0))
'        If Val(oForm01.Items("F_Mi").Specific.VALUE) = 0 And Val(oForm01.Items("F_Pl").Specific.VALUE) = 0 Then
'            '//MAX 일때는 0과 MAX값 사이에 입력가능
'            If VALUE < 0 Or VALUE > Val(oForm01.Items("F_Val").Specific.VALUE) Then
'                ErrNum = 27
'                oForm01.Items("M_F").Click ct_Regular
'                GoTo HeaderSpaceLineDel_Error
'            End If
'        Else
'            SPEC_MIN = Val(oForm01.Items("F_Val").Specific.VALUE) - Val(oForm01.Items("F_Mi").Specific.VALUE)
'            SPEC_MAX = Val(oForm01.Items("F_Val").Specific.VALUE) + Val(oForm01.Items("F_Pl").Specific.VALUE)
'
'            If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
'                ErrNum = 27
'                oForm01.Items("M_F").Click ct_Regular
'                GoTo HeaderSpaceLineDel_Error
'            End If
'        End If
'    End If
'    '//측정항목F
'    If Val(oDS_PS_QM082H.GetValue("U_M_F1", 0)) <> 0 Then
'        VALUE = Val(oDS_PS_QM082H.GetValue("U_M_F1", 0))
'        If Val(oForm01.Items("F_Mi").Specific.VALUE) = 0 And Val(oForm01.Items("F_Pl").Specific.VALUE) = 0 Then
'            If VALUE < 0 Or VALUE > Val(oForm01.Items("F1_Val").Specific.VALUE) Then
'                ErrNum = 28
'                oForm01.Items("M_F1").Click ct_Regular
'                GoTo HeaderSpaceLineDel_Error
'            End If
'        Else
'            SPEC_MIN = Val(oForm01.Items("F1_Val").Specific.VALUE) - Val(oForm01.Items("F1_Mi").Specific.VALUE)
'            SPEC_MAX = Val(oForm01.Items("F1_Val").Specific.VALUE) + Val(oForm01.Items("F1_Pl").Specific.VALUE)
'            If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
'                ErrNum = 28
'                oForm01.Items("M_F1").Click ct_Regular
'                GoTo HeaderSpaceLineDel_Error
'            End If
'        End If
'    End If
'    '//측정항목G
'    If Val(oDS_PS_QM082H.GetValue("U_M_G", 0)) <> 0 Then
'        VALUE = Val(oDS_PS_QM082H.GetValue("U_M_G", 0))
'        SPEC_MIN = Val(oForm01.Items("G_Val").Specific.VALUE) - Val(oForm01.Items("G_Mi").Specific.VALUE)
'        SPEC_MAX = Val(oForm01.Items("G_Val").Specific.VALUE) + Val(oForm01.Items("G_Pl").Specific.VALUE)
'        If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
'            ErrNum = 29
'            oForm01.Items("M_G").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//측정항목G1
'    If Val(oDS_PS_QM082H.GetValue("U_M_G1", 0)) <> 0 Then
'        VALUE = Val(oDS_PS_QM082H.GetValue("U_M_G1", 0))
'        SPEC_MIN = Val(oForm01.Items("G1_Val").Specific.VALUE) - Val(oForm01.Items("G1_Mi").Specific.VALUE)
'        SPEC_MAX = Val(oForm01.Items("G1_Val").Specific.VALUE) + Val(oForm01.Items("G1_Pl").Specific.VALUE)
'        If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
'            ErrNum = 30
'            oForm01.Items("M_G1").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//측정항목A-G1
'    If Val(oDS_PS_QM082H.GetValue("U_M_AG1", 0)) <> 0 Then
'        VALUE = Val(oDS_PS_QM082H.GetValue("U_M_AG1", 0))
'        SPEC_MIN = Val(oForm01.Items("AG1_Val").Specific.VALUE) - Val(oForm01.Items("AG1_Mi").Specific.VALUE)
'        SPEC_MAX = Val(oForm01.Items("AG1_Val").Specific.VALUE) + Val(oForm01.Items("AG1_Pl").Specific.VALUE)
'        If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
'            ErrNum = 31
'            oForm01.Items("M_AG1").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//측정항목H-H1
'    If Val(oDS_PS_QM082H.GetValue("U_M_HH1", 0)) <> 0 Then
'        VALUE = Val(oDS_PS_QM082H.GetValue("U_M_HH1", 0))
'        SPEC_MAX = Val(oForm01.Items("HH1_Val").Specific.VALUE)
'        If VALUE > SPEC_MAX Then
'            ErrNum = 32
'            oForm01.Items("M_HH1").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//측정항목I
'    If Val(oDS_PS_QM082H.GetValue("U_M_I", 0)) <> 0 Then
'        VALUE = Val(oDS_PS_QM082H.GetValue("U_M_I", 0))
'        SPEC_MIN = Val(oForm01.Items("I_Val").Specific.VALUE) - Val(oForm01.Items("I_Mi").Specific.VALUE)
'        SPEC_MAX = Val(oForm01.Items("I_Val").Specific.VALUE) + Val(oForm01.Items("I_Pl").Specific.VALUE)
'        If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
'            ErrNum = 33
'            oForm01.Items("M_I").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//측정항목I1
'    If Val(oDS_PS_QM082H.GetValue("U_M_I1", 0)) <> 0 Then
'        VALUE = Val(oDS_PS_QM082H.GetValue("U_M_I1", 0))
'        SPEC_MIN = Val(oForm01.Items("I1_Val").Specific.VALUE) - Val(oForm01.Items("I1_Mi").Specific.VALUE)
'        SPEC_MAX = Val(oForm01.Items("I1_Val").Specific.VALUE) + Val(oForm01.Items("I1_Pl").Specific.VALUE)
'        If VALUE < SPEC_MIN Or VALUE > SPEC_MAX Then
'            ErrNum = 34
'            oForm01.Items("M_I1").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//측정항목J-J1
'    If Val(oDS_PS_QM082H.GetValue("U_M_JJ1", 0)) <> 0 Then
'        VALUE = Val(oDS_PS_QM082H.GetValue("U_M_JJ1", 0))
'        SPEC_MAX = Val(oForm01.Items("JJ1_Val").Specific.VALUE)
'        If VALUE > SPEC_MAX Then
'            ErrNum = 35
'            oForm01.Items("M_JJ1").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//측정항목Edge Burr
'    If Val(oDS_PS_QM082H.GetValue("U_M_EdgeBu", 0)) <> 0 Then
'        VALUE = Val(oDS_PS_QM082H.GetValue("U_M_EdgeBu", 0))
'        SPEC_MAX = Val(oForm01.Items("EdgeBurr").Specific.VALUE)
'        If VALUE > SPEC_MAX Then
'            ErrNum = 36
'            oForm01.Items("M_EdgeBu").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//측정항목Camber
'    If Val(oDS_PS_QM082H.GetValue("U_M_Camber", 0)) <> 0 Then
'        VALUE = Val(oDS_PS_QM082H.GetValue("U_M_Camber", 0))
'        If VALUE > 0.5 Then
'            ErrNum = 37
'            oForm01.Items("M_Camber").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//측정항목Surface Roughness Min
'    If Val(oDS_PS_QM082H.GetValue("U_M_SR_min", 0)) <> 0 Then
'        VALUE_MIN = Val(oDS_PS_QM082H.GetValue("U_M_SR_min", 0))
'        SPEC_MAX = Val(oForm01.Items("Surface").Specific.VALUE)
'        If VALUE_MIN > SPEC_MAX Then
'            ErrNum = 38
'            oForm01.Items("M_SR_min").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//측정항목Surface Roughness Max
'    If Val(oDS_PS_QM082H.GetValue("U_M_SR_max", 0)) <> 0 Then
'        VALUE_MAX = Val(oDS_PS_QM082H.GetValue("U_M_SR_max", 0))
'        SPEC_MAX = Val(oForm01.Items("Surface").Specific.VALUE)
'        If VALUE_MAX > SPEC_MAX Then
'            ErrNum = 39
'            oForm01.Items("M_SR_max").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
'    '//측정항목Cross Bow
'    If Val(oDS_PS_QM082H.GetValue("U_M_CrossB", 0)) <> 0 Then
'        VALUE = Val(oDS_PS_QM082H.GetValue("U_M_CrossB", 0))
'        SPEC_MAX = Val(oForm01.Items("CrossBow").Specific.VALUE)
'        If VALUE > SPEC_MAX Then
'            ErrNum = 40
'            oForm01.Items("M_CrossB").Click ct_Regular
'            GoTo HeaderSpaceLineDel_Error
'        End If
'    End If
    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "LotNo는 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "원재료선택은 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "입고일자는 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "수량은 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "중량은 필수사항입니다. 확인하여 주십시오.", "E"
            
    ElseIf ErrNum = 10 Then
        MDC_Com.MDC_GF_Message "tensile strength 인장강도(kgf/㎡)의 Min 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 11 Then
        MDC_Com.MDC_GF_Message "tensile strength 인장강도(kgf/㎡)의 Max 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 12 Then
        MDC_Com.MDC_GF_Message "elongation 연신율(%)의 Min 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 13 Then
        MDC_Com.MDC_GF_Message "elongation 연신율(%)의 Max 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 14 Then
        MDC_Com.MDC_GF_Message "Hardness 경도(Hv)의 Min 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 15 Then
        MDC_Com.MDC_GF_Message "Hardness 경도(Hv)의 Max 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 16 Then
        MDC_Com.MDC_GF_Message "Fe 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 17 Then
        MDC_Com.MDC_GF_Message "P 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 18 Then
        MDC_Com.MDC_GF_Message "A 의 Min 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 19 Then
        MDC_Com.MDC_GF_Message "A 의 Max 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 20 Then
        MDC_Com.MDC_GF_Message "B 의 Min 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 21 Then
        MDC_Com.MDC_GF_Message "B 의 Max 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 22 Then
        MDC_Com.MDC_GF_Message "B1 의 Min 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 23 Then
        MDC_Com.MDC_GF_Message "B1 의 Max 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 24 Then
        MDC_Com.MDC_GF_Message "C 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 25 Then
        MDC_Com.MDC_GF_Message "D 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 26 Then
        MDC_Com.MDC_GF_Message "E-E1 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 27 Then
        MDC_Com.MDC_GF_Message "F 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 28 Then
        MDC_Com.MDC_GF_Message "F1 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 29 Then
        MDC_Com.MDC_GF_Message "G 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 30 Then
        MDC_Com.MDC_GF_Message "G1 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 31 Then
        MDC_Com.MDC_GF_Message "AG1 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 32 Then
        MDC_Com.MDC_GF_Message "H-H1 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 33 Then
        MDC_Com.MDC_GF_Message "I 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 34 Then
        MDC_Com.MDC_GF_Message "I1 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 35 Then
        MDC_Com.MDC_GF_Message "J-J1 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 36 Then
        MDC_Com.MDC_GF_Message "Edge Burr 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 37 Then
        MDC_Com.MDC_GF_Message "Camber 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 38 Then
        MDC_Com.MDC_GF_Message "surface roughness의 Min 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 39 Then
        MDC_Com.MDC_GF_Message "surface roughness의 Max 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 40 Then
        MDC_Com.MDC_GF_Message "Cross Bow 의 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 41 Then
        MDC_Com.MDC_GF_Message "corner radius 의 Min 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    ElseIf ErrNum = 42 Then
        MDC_Com.MDC_GF_Message "corner radius 의 Max 측정치가 SPEC 규격 범위를 벗어났습니다", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function


Private Function Check_Exist() As Boolean
On Error GoTo Check_Exist_Error
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrNum          As Integer
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "select * from [@PS_QM082H] where U_DocDate = '" & Trim(oDS_PS_QM082H.GetValue("U_DocDate", 0)) & "' and "
    sQry = sQry + " U_LotNo = '" & Trim(oDS_PS_QM082H.GetValue("U_LotNo", 0)) & "'"
    oRecordSet.DoQuery sQry

    If oRecordSet.RecordCount <> 0 Then
        ErrNum = 1
        GoTo Check_Exist_Error
    End If
    
    Set oRecordSet = Nothing
    Check_Exist = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Check_Exist_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "이미 동일한 입고일자에 LotNo가 존재합니다. 확인하여 주십시오.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    Check_Exist = False
End Function

