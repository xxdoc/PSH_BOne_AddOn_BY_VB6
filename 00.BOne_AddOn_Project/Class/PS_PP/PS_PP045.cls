VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP045"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_PP045.cls
'//  Module         : PP
'//  Description    : 생산일보조회
'//  FormType       : PS_PP045
'//  Create Date    : 2010.12.22
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01              As SAPbouiCOM.Matrix
Private oDS_PS_USERDS01     As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private oSeq&

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm(Optional ByRef oForm02 As SAPbouiCOM.Form, Optional ByVal oItemUID02 As String, Optional ByVal oColUID02 As String, Optional ByVal oColRow02 As Long, Optional ByVal oTradeType02 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP045.srf")
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
    
    oFormUniqueID01 = "PS_PP045_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
        
    oForm01.Freeze True
        
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    
    oForm01.EnableMenu ("1281"), False        '// 찾기
    oForm01.EnableMenu ("1282"), False        '// 추가
'    oForm01.EnableMenu ("1283"), False        '// 삭제
'    oForm01.EnableMenu ("1286"), False         '// 닫기
'    oForm01.EnableMenu ("1287"), False        '// 복제
'    oForm01.EnableMenu ("1284"), False         '// 취소
'    oForm01.EnableMenu ("1293"), False         '// 행삭제
        
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

'****************************************************************************************************************
'// ItemEventHander
'****************************************************************************************************************
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
    Dim i&
    Dim ErrNum&
    Dim TempForm01 As Variant
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
        
    Dim ItemCode$, ItemName$, Size$, Qty&, Weight As Currency, Unit$, RequestDate$, DueDate$, ItemType$, RequestNo$
    Dim RFC_Sender$
    Dim Calculate_Weight As Double
    Dim Seq&
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
            
                If pval.ItemUID = "Btn01" Then
                
                    Call Me.PS_PP045_MTX01
                
                End If
            
'et_KEY_DOWN ///////////////////'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "CntcCode" Then
                        If oForm01.Items("CntcCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
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
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
            Case et_KEY_DOWN: '//2
'et_COMBO_SELECT ///////////////////'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
                If pval.ItemChanged = True Then
                    'If pval.ItemUID = "BPLId" Or pval.ItemUID = "OrdGbn" Then
                        Call FlushToItemValue(pval.ItemUID)
                    'End If
                End If
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
'et_MATRIX_LINK_PRESSED /////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_MATRIX_LINK_PRESSED '//8
                If pval.ItemUID = "Mat01" And pval.ColUID = "DocEntry" Then
                    If Trim(oMat01.Columns("BPLId").Cells(pval.Row).Specific.VALUE) = "3" Then
                        Set TempForm01 = New PS_PP043
                    Else
                        If Trim(oMat01.Columns("OrdGbn").Cells(pval.Row).Specific.VALUE) = "104" Or Trim(oMat01.Columns("OrdGbn").Cells(pval.Row).Specific.VALUE) = "107" Then
                            Set TempForm01 = New PS_PP041
                        Else
                            Set TempForm01 = New PS_PP040
                        End If
                    End If
                    Call TempForm01.LoadForm(oMat01.Columns("DocEntry").Cells(pval.Row).Specific.VALUE)
                    
                    Set TempForm01 = Nothing
                End If
'et_VALIDATE ///////////////////'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    'If pval.ItemUID = "DocDateFr" Or pval.ItemUID = "DocDateTo" Or pval.ItemUID = "CntcCode" Then
                        Call FlushToItemValue(pval.ItemUID)
                    'End If
                End If
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
'et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
                Set oDS_PS_USERDS01 = Nothing
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
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_MenuEvent_Error:" & Err.Number & " - " & Err.Description, "E"
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

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If (eventInfo.BeforeAction = True) Then
        '//작업
    ElseIf (eventInfo.BeforeAction = False) Then
        '//작업
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//디비데이터 소스 개체 할당
    Set oDS_PS_USERDS01 = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    '사업장
    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 10)
    Call oForm01.Items("BPLId").Specific.DataBind.SetBound(True, "", "BPLId")
    
    '팀
    Call oForm01.DataSources.UserDataSources.Add("TeamCode", dt_SHORT_TEXT, 10)
    Call oForm01.Items("TeamCode").Specific.DataBind.SetBound(True, "", "TeamCode")
    
    '담당
    Call oForm01.DataSources.UserDataSources.Add("RspCode", dt_SHORT_TEXT, 10)
    Call oForm01.Items("RspCode").Specific.DataBind.SetBound(True, "", "RspCode")
    
    '반
    Call oForm01.DataSources.UserDataSources.Add("ClsCode", dt_SHORT_TEXT, 10)
    Call oForm01.Items("ClsCode").Specific.DataBind.SetBound(True, "", "ClsCode")
    
    '작업타입
    Call oForm01.DataSources.UserDataSources.Add("OrdGbn", dt_SHORT_TEXT, 10)
    Call oForm01.Items("OrdGbn").Specific.DataBind.SetBound(True, "", "OrdGbn")
    
    '일자(시작)
    Call oForm01.DataSources.UserDataSources.Add("DocDateFr", dt_DATE, 8)
    oForm01.Items("DocDateFr").Specific.DataBind.SetBound True, "", "DocDateFr"
    oForm01.DataSources.UserDataSources.Item("DocDateFr").VALUE = Format(Date, "YYYYMM01")
    
    '일자(종료)
    Call oForm01.DataSources.UserDataSources.Add("DocDateTo", dt_DATE, 8)
    oForm01.Items("DocDateTo").Specific.DataBind.SetBound True, "", "DocDateTo"
    oForm01.DataSources.UserDataSources.Item("DocDateTo").VALUE = Format(Date, "YYYYMMDD")
    
    '작업자(사번)
    Call oForm01.DataSources.UserDataSources.Add("CntcCode", dt_SHORT_TEXT, 20)
    Call oForm01.Items("CntcCode").Specific.DataBind.SetBound(True, "", "CntcCode")
    
    '작업자(성명)
    Call oForm01.DataSources.UserDataSources.Add("CntcName", dt_SHORT_TEXT, 50)
    Call oForm01.Items("CntcName").Specific.DataBind.SetBound(True, "", "CntcName")

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
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// 사업장
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL] Order by BPLId"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oMat01.Columns("BPLId").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '// 작업구분
    Set oCombo = oForm01.Items("OrdGbn").Specific
    sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Where U_PudYN = 'Y' Order by Code"
    oRecordSet01.DoQuery sQry
    oCombo.ValidValues.Add "", ""
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oMat01.Columns("OrdGbn").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '// 사용자
'    sQry = "Select U_MSTCOD, lastName + firstName From OHEM Order by empID"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oMat01.Columns("CntcCode").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
'        oRecordSet01.MoveNext
'    Loop


'    '매트릭스
'    '작업구분
'    sQry = "           SELECT      a.Code, "
'    sQry = sQry & "                a.Name"
'    sQry = sQry & " FROM       [@PSH_ITMBSORT] a"
'    sQry = sQry & " WHERE      a.U_PudYN = 'Y'"
'    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("OrdType"), sQry)
    
    '작업구분
    oMat01.Columns("OrdType").ValidValues.Add "10", "일반"
    oMat01.Columns("OrdType").ValidValues.Add "20", "PSMT지원"
    oMat01.Columns("OrdType").ValidValues.Add "30", "외주"
    oMat01.Columns("OrdType").ValidValues.Add "40", "실적"
    oMat01.Columns("OrdType").ValidValues.Add "50", "일반조정"
    oMat01.Columns("OrdType").ValidValues.Add "60", "외주조정"
    oMat01.Columns("OrdType").ValidValues.Add "70", "설계시간"
    
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "ComboBox_Setting_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Initialization()
On Error GoTo Initialization_Error
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    '//아이디별 사업장 세팅
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    '//아이디별 사번 세팅
'    oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD '기능 취소(2013.06.10 송명규)
    
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

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim BPLId$, OrdGbn$, DocDateFr$, DocDateTo$, CntcCode$
    
    Dim TeamCode As String
    Dim RspCode As String
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
    
        Case "BPLId"
        
            BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
            
            If oForm01.Items("TeamCode").Specific.ValidValues.Count > 0 Then
                For i = oForm01.Items("TeamCode").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm01.Items("TeamCode").Specific.ValidValues.Remove(i, psk_Index)
                Next i
            End If
            
            '부서콤보세팅
            Call oForm01.Items("TeamCode").Specific.ValidValues.Add("%", "전체")
            sQry = "            SELECT      U_Code AS [Code],"
            sQry = sQry & "                 U_CodeNm As [Name]"
            sQry = sQry & "  FROM       [@PS_HR200L]"
            sQry = sQry & "  WHERE      Code = '1'"
            sQry = sQry & "                 AND U_UseYN = 'Y'"
            sQry = sQry & "                 AND U_Char2 = '" & BPLId & "'"
            sQry = sQry & "  ORDER BY  U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm01.Items("TeamCode").Specific, sQry, "", False, False)
            Call oForm01.Items("TeamCode").Specific.Select(0, psk_Index)
    
        Case "TeamCode"
            
            TeamCode = Trim(oForm01.Items("TeamCode").Specific.VALUE)
            
            If oForm01.Items("RspCode").Specific.ValidValues.Count > 0 Then
                For i = oForm01.Items("RspCode").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm01.Items("RspCode").Specific.ValidValues.Remove(i, psk_Index)
                Next i
            End If
            
            '담당콤보세팅
            Call oForm01.Items("RspCode").Specific.ValidValues.Add("%", "전체")
            sQry = "            SELECT      U_Code AS [Code],"
            sQry = sQry & "                 U_CodeNm As [Name]"
            sQry = sQry & "  FROM       [@PS_HR200L]"
            sQry = sQry & "  WHERE      Code = '2'"
            sQry = sQry & "                 AND U_UseYN = 'Y'"
            sQry = sQry & "                 AND U_Char1 = '" & TeamCode & "'"
            sQry = sQry & "  ORDER BY  U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm01.Items("RspCode").Specific, sQry, "", False, False)
            Call oForm01.Items("RspCode").Specific.Select(0, psk_Index)
        
        Case "RspCode"
            
            TeamCode = Trim(oForm01.Items("TeamCode").Specific.VALUE)
            RspCode = Trim(oForm01.Items("RspCode").Specific.VALUE)
            
            If oForm01.Items("ClsCode").Specific.ValidValues.Count > 0 Then
                For i = oForm01.Items("ClsCode").Specific.ValidValues.Count - 1 To 0 Step -1
                    Call oForm01.Items("ClsCode").Specific.ValidValues.Remove(i, psk_Index)
                Next i
            End If
            
            '반콤보세팅
            Call oForm01.Items("ClsCode").Specific.ValidValues.Add("%", "전체")
            sQry = "            SELECT      U_Code AS [Code],"
            sQry = sQry & "                 U_CodeNm As [Name]"
            sQry = sQry & "  FROM       [@PS_HR200L]"
            sQry = sQry & "  WHERE      Code = '9'"
            sQry = sQry & "                 AND U_UseYN = 'Y'"
            sQry = sQry & "                 AND U_Char1 = '" & RspCode & "'"
            sQry = sQry & "                 AND U_Char2 = '" & TeamCode & "'"
            sQry = sQry & "  ORDER BY  U_Seq"
            Call MDC_SetMod.Set_ComboList(oForm01.Items("ClsCode").Specific, sQry, "", False, False)
            Call oForm01.Items("ClsCode").Specific.Select(0, psk_Index)
    
    
        Case "CntcCode"
            
            sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" & Trim(oForm01.Items("CntcCode").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry

            oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
            
    
'        Case "BPLId", "OrdGbn", "DocDateFr", "DocDateTo", "CntcCode"
'            oForm01.Freeze True
'            If oUID = "CntcCode" Then
'                sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" & Trim(oForm01.Items("CntcCode").Specific.VALUE) & "'"
'                oRecordSet01.DoQuery sQry
'
'                oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
'            End If
'
'            oMat01.Clear
'            oDS_PS_USERDS01.Clear
'
'            BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
'            OrdGbn = Trim(oForm01.Items("OrdGbn").Specific.VALUE)
'            DocDateFr = Trim(oForm01.Items("DocDateFr").Specific.VALUE)
'            DocDateTo = Trim(oForm01.Items("DocDateTo").Specific.VALUE)
'            CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
'
'            If BPLId = "" Then BPLId = "%"
'            If OrdGbn = "" Then OrdGbn = "%"
'            If DocDateFr = "" Then DocDateFr = "18990101"
'            If DocDateTo = "" Then DocDateTo = "20991231"
'            If CntcCode = "" Then CntcCode = "%"
'
'            sQry = "EXEC [PS_PP045_01] '" & BPLId & "', '" & OrdGbn & "', '" & DocDateFr & "', '" & DocDateTo & "', '" & CntcCode & "'"
'            oRecordSet01.DoQuery sQry
'
'            If oRecordSet01.RecordCount = 0 Then
'                MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
'                Set oRecordSet01 = Nothing
'                oForm01.Freeze False
'                Exit Sub
'            End If
'
'            Dim ProgBar01 As SAPbouiCOM.ProgressBar
'            Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
'
'            For i = 0 To oRecordSet01.RecordCount - 1
'                If i + 1 > oDS_PS_USERDS01.Size Then
'                    oDS_PS_USERDS01.InsertRecord (i)
'                End If
'
'                oMat01.AddRow
'                oDS_PS_USERDS01.Offset = i
'                oDS_PS_USERDS01.setValue "U_LineNum", i, i + 1
'                oDS_PS_USERDS01.setValue "U_ColDt01", i, Format(Trim(oRecordSet01.Fields("U_DocDate").VALUE), "YYYYMMDD")
'                oDS_PS_USERDS01.setValue "U_ColReg01", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)
'                oDS_PS_USERDS01.setValue "U_ColReg02", i, Trim(oRecordSet01.Fields("U_BPLId").VALUE)
'                oDS_PS_USERDS01.setValue "U_ColReg03", i, Trim(oRecordSet01.Fields("U_OrdType").VALUE)
'                oDS_PS_USERDS01.setValue "U_ColReg04", i, Trim(oRecordSet01.Fields("U_OrdGbn").VALUE)
'                oDS_PS_USERDS01.setValue "U_ColReg05", i, Trim(oRecordSet01.Fields("U_WorkCode").VALUE)
'                oDS_PS_USERDS01.setValue "U_ColReg06", i, Trim(oRecordSet01.Fields("U_WorkName").VALUE)
'                oDS_PS_USERDS01.setValue "U_ColReg07", i, Trim(oRecordSet01.Fields("U_YTime").VALUE)
''----------------------------------------------------------------------------------------------------------
'                oRecordSet01.MoveNext
'                ProgBar01.VALUE = ProgBar01.VALUE + 1
'                ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
'            Next
'            oMat01.LoadFromDataSource
''            oMat01.AutoResizeColumns
'            ProgBar01.Stop
'            oForm01.Freeze False
    End Select
    
'    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
FlushToItemValue_Error:
    oForm01.Freeze False
'    ProgBar01.Stop
'    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub PS_PP045_MTX01()
'******************************************************************************
'Function ID : PS_PP045_MTX01()
'해당모듈 : PS_PP045
'기능 : 데이터 조회
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_PP045_MTX01_Error

    Dim i As Integer
    Dim sQry As String
    Dim ErrNum As Integer
    
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim BPLId As String '사업장
    Dim TeamCode As String '팀
    Dim RspCode As String '담당
    Dim ClsCode As String '반
    Dim OrdGbn As String
    Dim DocDateFr As String
    Dim DocDateTo As String
    Dim CntcCode As String '사번
    
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    TeamCode = Trim(oForm01.Items("TeamCode").Specific.VALUE)
    RspCode = Trim(oForm01.Items("RspCode").Specific.VALUE)
    ClsCode = Trim(oForm01.Items("ClsCode").Specific.VALUE)
    OrdGbn = Trim(oForm01.Items("OrdGbn").Specific.VALUE)
    DocDateFr = Trim(oForm01.Items("DocDateFr").Specific.VALUE)
    DocDateTo = Trim(oForm01.Items("DocDateTo").Specific.VALUE)
    CntcCode = Trim(oForm01.Items("CntcCode").Specific.VALUE)
    
    If BPLId = "" Then BPLId = "%"
    If OrdGbn = "" Then OrdGbn = "%"
    If DocDateFr = "" Then DocDateFr = "18990101"
    If DocDateTo = "" Then DocDateTo = "20991231"
    If CntcCode = "" Then CntcCode = "%"
    
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    Call oForm01.Freeze(True)
    
    sQry = "         EXEC [PS_PP045_01] '"
    sQry = sQry & BPLId & "','"
    sQry = sQry & TeamCode & "','"
    sQry = sQry & RspCode & "','"
    sQry = sQry & ClsCode & "','"
    sQry = sQry & OrdGbn & "','"
    sQry = sQry & DocDateFr & "','"
    sQry = sQry & DocDateTo & "','"
    sQry = sQry & CntcCode & "'"
    Call oRecordSet01.DoQuery(sQry)
    
    oMat01.Clear
    oDS_PS_USERDS01.Clear
    oMat01.FlushToDataSource
    oMat01.LoadFromDataSource

    If (oRecordSet01.RecordCount = 0) Then
    
        ErrNum = 1
        
        oForm01.Mode = fm_ADD_MODE
        
'        Call PS_PP045_Add_MatrixRow(0, True)
        GoTo PS_PP045_MTX01_Error
        
        Exit Sub
    End If
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_USERDS01.Size Then
            oDS_PS_USERDS01.InsertRecord (i)
        End If

        oMat01.AddRow
        oDS_PS_USERDS01.Offset = i
        
        oDS_PS_USERDS01.setValue "U_LineNum", i, i + 1
        oDS_PS_USERDS01.setValue "U_ColDt01", i, Format(Trim(oRecordSet01.Fields("DocDate").VALUE), "YYYYMMDD")
        oDS_PS_USERDS01.setValue "U_ColReg01", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)
        oDS_PS_USERDS01.setValue "U_ColReg02", i, Trim(oRecordSet01.Fields("BPLId").VALUE)
        oDS_PS_USERDS01.setValue "U_ColReg03", i, Trim(oRecordSet01.Fields("OrdType").VALUE)
        oDS_PS_USERDS01.setValue "U_ColReg04", i, Trim(oRecordSet01.Fields("OrdGbn").VALUE)
        oDS_PS_USERDS01.setValue "U_ColReg05", i, Trim(oRecordSet01.Fields("CntcCode").VALUE)
        oDS_PS_USERDS01.setValue "U_ColReg06", i, Trim(oRecordSet01.Fields("FullName").VALUE)
        oDS_PS_USERDS01.setValue "U_ColQty01", i, Trim(oRecordSet01.Fields("YTime").VALUE)
        oDS_PS_USERDS01.setValue "U_ColQty02", i, Trim(oRecordSet01.Fields("WorkTime").VALUE)
        oDS_PS_USERDS01.setValue "U_ColQty03", i, Trim(oRecordSet01.Fields("Diff").VALUE)
        
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
PS_PP045_MTX01_Error:
'    ProgBar01.Stop
    Call oForm01.Freeze(False)
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        Call MDC_Com.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W")
    Else
        Call MDC_Com.MDC_GF_Message("PS_PP045_MTX01_Error:" & Err.Number & " - " & Err.Description, "E")
    End If
End Sub

