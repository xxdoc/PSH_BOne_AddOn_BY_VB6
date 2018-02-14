VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP913"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_PP913.cls
'//  Module         : PP
'//  Description    : 장비가동시간등록
'//  FormType       : PS_PP913
'//  Create Date    : 2011.06.13
'//  Modified Date  :
'//  Creator        : Youn Je Hyung
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Public oGrid01              As SAPbouiCOM.Grid
Private oDS_PS_PP913H       As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_PP913L       As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP913.srf")
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
    
    oFormUniqueID01 = "PS_PP913_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call FormClear
    'Call Matrix_AddRow(1, 0, True) 'oMat01
    FormItemEnabled
    
    oForm01.EnableMenu ("1283"), True         '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
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
    Set oDS_PS_PP913H = oForm01.DataSources.DBDataSources("@PS_PP913H")
'    Set oDS_PS_PP913L = oForm01.DataSources.DBDataSources("@PS_PP913L")
    
    Set oGrid01 = oForm01.Items("Grid01").Specific
    oForm01.DataSources.DataTables.Add ("ZTEMP")
    
    '// 메트릭스 개체 할당
'    Set oMat01 = oForm01.Items("Mat01").Specific
    
    oDS_PS_PP913H.setValue "U_DocDate", 0, Format(Now - 1, "YYYYMMDD")
    
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
        oRecordSet01.MoveNext
    Loop

    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
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
                If pval.ItemUID = "Search" Then '조회버튼
                    If oForm01.Items("DocDate").Specific.VALUE = "" Then
                        oGrid01.DataTable.Clear
                        MDC_Com.MDC_GF_Message "일자를 입력 후 조회버튼을 눌러주십시오.", "E"
                    Else
                        Call Search_Grid_Data
                    End If
                    
                ElseIf pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        
'                        If MatrixSpaceLineDel = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
                    End If
                End If
                
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    '//헤더
                    If pval.ItemUID = "CpCode" Then
                        If oForm01.Items("CpCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    If pval.ItemUID = "JsCode" Then
                        If oForm01.Items("JsCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    If pval.ItemUID = "NCode01" Then
                        If oForm01.Items("NCode01").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    If pval.ItemUID = "NCode02" Then
                        If oForm01.Items("NCode02").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    If pval.ItemUID = "NCode03" Then
                        If oForm01.Items("NCode03").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    If pval.ItemUID = "NCode04" Then
                        If oForm01.Items("NCode04").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    '//라인
'                    If pval.ItemUID = "Mat01" Then
'                        If pval.ColUID = "PP070No" Then
'                            If oMat01.Columns("PP070No").Cells(pval.Row).Specific.VALUE = "" Then
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
                        Call Search_Grid_Data
                    ElseIf oForm01.Mode = fm_OK_MODE Then
                        FormItemEnabled
'                        Call Matrix_AddRow(1, oMat01.RowCount, False) 'oMat01
                        Call Search_Grid_Data
                    End If
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
                If pval.ItemUID = "Grid01" Then
                    Call Display_Data
                End If
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    '//헤더
                    If pval.ItemUID = "CpCode" Or pval.ItemUID = "JsCode" Or pval.ItemUID = "NCode01" Or pval.ItemUID = "NCode02" Or pval.ItemUID = "NCode03" Or pval.ItemUID = "NCode04" Then
                        FlushToItemValue pval.ItemUID
                    End If
'                    '//라인
'                    If pval.ItemUID = "Mat01" And (pval.ColUID = "DocDate" Or pval.ColUID = "LoanAmt" Or pval.ColUID = "RepayAmt") Then
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
            Case "1293": '행삭제
            Case "1281": '찾기
                FormItemEnabled
                oForm01.Items("DocNum").Click ct_Regular
                
            Case "1282": '추가
                Call FormItemEnabled
                Call FormClear
                oDS_PS_PP913H.setValue "U_DocDate", 0, Format(Now, "YYYYMMDD") - 1
                oForm01.Items("DocDate").Click ct_Regular
                
                Dim oCombo          As SAPbouiCOM.ComboBox
                Set oCombo = oForm01.Items("BPLId").Specific
                oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
                Set oCombo = Nothing
                
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                If oGrid01.Rows.Count <> 0 Then oGrid01.DataTable.Clear
                FormItemEnabled
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
        oForm01.Items("DocNum").Enabled = True
        oForm01.Items("Search").Enabled = False
       
    ElseIf oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("DocNum").Enabled = False
        oForm01.Items("Search").Enabled = True
    ElseIf oForm01.Mode = fm_OK_MODE Then
        oForm01.Items("DocNum").Enabled = False
        oForm01.Items("Search").Enabled = True
    End If
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP913'", "")
    If DocNum = 0 Then
        oForm01.Items("DocNum").Specific.VALUE = 1
    Else
        oForm01.Items("DocNum").Specific.VALUE = DocNum
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormClear_Error:
    MDC_Com.MDC_GF_Message "FormClear_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
'*******************************************************************
'// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
'*******************************************************************
Private Sub Matrix_AddRow(ByVal oMat As Integer, ByVal oRow As Long, Optional Insert_YN As Boolean)
On Error GoTo Matrix_AddRow_Error
        
    Select Case oMat
        Case 1: 'oMat01
'            If Insert_YN = False Then
'                oRow = oMat01.RowCount
'                oDS_PS_PP913L.InsertRecord (oRow)
'            End If
'            '수입내역
'            oDS_PS_PP913L.Offset = oRow
'            oDS_PS_PP913L.setValue "LineId", oRow, oRow + 1
'            oDS_PS_PP913L.setValue "U_DocDate", oRow, ""
'            oDS_PS_PP913L.setValue "U_Comments", oRow, ""
'            oDS_PS_PP913L.setValue "U_LoanAmt", oRow, ""
'            oDS_PS_PP913L.setValue "U_Interest", oRow, ""
'            oDS_PS_PP913L.setValue "U_RepayAmt", oRow, ""
'            oDS_PS_PP913L.setValue "U_Balance", oRow, ""
'            oMat01.LoadFromDataSource
        
    End Select
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Matrix_AddRow_Error:
    MDC_Com.MDC_GF_Message "Matrix_AddRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&, DocNum$, LineId$
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim Balance         As Currency
    Dim RBalance        As Currency
    
    Dim oRecordSet    As SAPbobsCOM.Recordset
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    '--------------------------------------------------------------
    'Header--------------------------------------------------------
    Select Case oUID
        Case "CpCode"
            sQry = "SELECT U_CpName FROM [@PS_PP001L] where U_CpCode = '" & Trim(oDS_PS_PP913H.GetValue("U_CpCode", 0)) & "'"
            oRecordSet.DoQuery sQry
            oDS_PS_PP913H.setValue "U_CpName", 0, Trim(oRecordSet.Fields(0).VALUE)
            
        Case "JsCode"
            sQry = "SELECT U_JsName FROM [@PS_PP007L] where U_JsCode = '" & Trim(oDS_PS_PP913H.GetValue("U_JsCode", 0)) & "'"
            oRecordSet.DoQuery sQry
            oDS_PS_PP913H.setValue "U_JsName", 0, Trim(oRecordSet.Fields(0).VALUE)
            
        Case "NCode01"
            
            sQry = "SELECT U_CdName FROM [@PS_SY001L] WHERE Code = 'P008' And U_Minor = '" & Trim(oDS_PS_PP913H.GetValue("U_NCode01", 0)) & "'"
            oRecordSet.DoQuery sQry
            oDS_PS_PP913H.setValue "U_NCdName1", 0, Trim(oRecordSet.Fields(0).VALUE)
            
        Case "NCode02"
            sQry = "SELECT U_CdName FROM [@PS_SY001L] WHERE Code = 'P008' And U_Minor = '" & Trim(oDS_PS_PP913H.GetValue("U_NCode02", 0)) & "'"
            oRecordSet.DoQuery sQry
            oDS_PS_PP913H.setValue "U_NCdName2", 0, Trim(oRecordSet.Fields(0).VALUE)
            
        Case "NCode03"
            sQry = "SELECT U_CdName FROM [@PS_SY001L] WHERE Code = 'P008' And U_Minor = '" & Trim(oDS_PS_PP913H.GetValue("U_NCode03", 0)) & "'"
            oRecordSet.DoQuery sQry
            oDS_PS_PP913H.setValue "U_NCdName3", 0, Trim(oRecordSet.Fields(0).VALUE)
            
        Case "NCode04"
            sQry = "SELECT U_CdName FROM [@PS_SY001L] WHERE Code = 'P008' And U_Minor = '" & Trim(oDS_PS_PP913H.GetValue("U_NCode04", 0)) & "'"
            oRecordSet.DoQuery sQry
            oDS_PS_PP913H.setValue "U_NCdName4", 0, Trim(oRecordSet.Fields(0).VALUE)
    End Select
        
    '--------------------------------------------------------------
    'Line----------------------------------------------------------
    If oUID = "Mat01" Then
        Select Case oCol
'            Case "DocDate"
'                oMat01.FlushToDataSource
'                oDS_PS_PP913L.Offset = oRow - 1
'                'oMat01.SetLineData oRow
'
'                '--------------------------------------------------------------------------------------------
'                If oRow = oMat01.RowCount And Trim(oDS_PS_PP913L.GetValue("U_DocDate", oRow - 1)) <> "" Then
'                '// 다음 라인 추가
'                    Call Matrix_AddRow(1, 0, False)
'                    oMat01.Columns("DocDate").Cells(oRow).Click ct_Regular
'                End If
         End Select
    End If
     
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    
    ErrNum = 0

    '// Check
    Select Case True
        Case Trim(oDS_PS_PP913H.GetValue("U_DocDate", 0)) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_PP913H.GetValue("U_CpCode", 0)) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_PP913H.GetValue("U_JsCode", 0)) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
    End Select
    
    If oGrid01.Rows.Count <> 0 Then oGrid01.DataTable.Clear

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "일자는 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "공정은 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "기계는 필수사항입니다. 확인하여 주십시오.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error
    Dim i               As Long
    Dim ErrNum          As Integer
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0
    
'    oMat01.FlushToDataSource

    '// 라인
'    If oMat01.VisualRowCount = 1 Then '// MAT01에 값이 있는지 확인 (ErrorNumber : 1)
'        ErrNum = 1
'        GoTo MatrixSpaceLineDel_Error
'    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//마지막 행 하나를 빼고 i=0부터 시작하므로 하나를 빼므로
    '//oMat01.RowCount - 2가 된다..반드시 들어 가야 하는 필수값을 확인한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'    If oMat01.VisualRowCount > 0 Then
'        For i = 0 To oMat01.VisualRowCount - 2      '// Mat1에 입력값이 올바르게 들어갔는지 확인 (ErrorNumber : 2)
'            oDS_PS_PP913L.Offset = i
'            If Trim(oDS_PS_PP913L.GetValue("U_DocDate", i)) = "" Then
'                ErrNum = 2
'                oMat01.Columns("DocDate").Cells(i + 1).Click ct_Regular
'                GoTo MatrixSpaceLineDel_Error
'            End If
'        Next i
'    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
    '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'    If oMat01.VisualRowCount > 0 Then
'        oDS_PS_PP913L.RemoveRecord oDS_PS_PP913L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
'    End If
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '행을 삭제하였으니 DB데이터 소스를 다시 가져온다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'    oMat01.LoadFromDataSource

    Set oRecordSet = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "일자(라인)는 필수사항입니다. 확인하여 주십시오.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function
Public Sub Search_Grid_Data()
    On Error GoTo Search_Grid_Data_Error
    Dim oRecordSet          As SAPbobsCOM.Recordset
    Dim sQry                As String
    Dim DocDate             As String
    Dim BPLId               As String
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oForm01.Freeze True
    
'    FrDate = Trim(oForm01.Items("FrDate").Specific.VALUE)
'    ToDate = Trim(oForm01.Items("ToDate").Specific.VALUE)
'    If FrDate = "" Then FrDate = "1900-01-01"
'    If ToDate = "" Then ToDate = "2100-12-31"
'    Status = Trim(oForm01.Items("DocStatus").Specific.Selected.VALUE)

    BPLId = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
    DocDate = Trim(oForm01.Items("DocDate").Specific.VALUE)
    
    sQry = "EXEC PS_PP913_01 '" & BPLId & "', '" & DocDate & "'"
        
    '/ Procedure 실행(Grid 사용)
    oForm01.DataSources.DataTables.Item(0).ExecuteQuery (sQry)
    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("ZTEMP")
    
    Call GridSetting
    
    oForm01.Freeze False
    
    Set oRecordSet = Nothing
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Search_Grid_Data_Error:
        Set oRecordSet = Nothing
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'****************************************************************************************************************
'// Grid 꾸며주기
'****************************************************************************************************************
Private Sub GridSetting()
On Error GoTo GridSetting_Error
    Dim i           As Integer
    Dim sColsTitle  As String
    Dim sColsLine   As String
    
    oForm01.Freeze True
    
'    oGrid01.SelectionMode = ms_Single
    
'    oGrid01.Columns.Item(0).LinkedObjectType = lf_BusinessPartner
'    oGrid01.Columns.Item(4).LinkedObjectType = lf_Invoice
'    oGrid01.Columns.Item(2).LinkedObjectType = lf_Items
       
    For i = 0 To oGrid01.Columns.Count - 1
        sColsTitle = oGrid01.Columns(i).TitleObject.Caption
        
        oGrid01.Columns(i).Editable = False
        
        If oGrid01.DataTable.Columns(i).Type = ft_Float Then
            oGrid01.Columns(i).RightJustified = True
        End If
        
'        If InStr(1, sColsTitle, "T(mm)") > 0 Or _
'           InStr(1, sColsTitle, "W(mm)") > 0 Or _
'           InStr(1, sColsTitle, "L(mm)") > 0 Or _
'           InStr(1, sColsTitle, "Weight") > 0 Or _
'           InStr(1, sColsTitle, "Qty(Kg)") > 0 Then
'            oGrid01.Columns(i).RightJustified = True
'            oGrid01.Columns(i).BackColor = &HE0E0E0
'        End If
'
'        If InStr(1, sColsTitle, "수량") > 0 Or _
'           InStr(1, sColsTitle, "중량") > 0 Then
'            oGrid01.Columns(i).RightJustified = True
'            oGrid01.Columns(i).BackColor = &HFFC0C0
'        End If
    Next i
        
    oForm01.Freeze False
    
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
GridSetting_Error:
    MDC_Com.MDC_GF_Message "GridSetting_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub Display_Data()
    On Error GoTo Search_Matrix_Data_Error
    Dim oRecordSet          As SAPbobsCOM.Recordset
    Dim sQry                As String
    
    Dim i                   As Integer
    Dim j                   As Integer
    Dim cnt                 As Integer
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    oForm01.Freeze True
    
     For i = 0 To oGrid01.Rows.Count - 1
        If oGrid01.Rows.IsSelected(i) = True Then
            
            oForm01.Mode = fm_FIND_MODE
            oForm01.Items("DocNum").Enabled = True
            oForm01.Items("DocNum").Specific.VALUE = Trim(oGrid01.DataTable.GetValue(0, i))
            oForm01.Items("1").Click ct_Regular
            GoTo Exit_grid
        
        End If
     Next i
    
Exit_grid:
    
'    oForm01.Items("1").Click ct_Regular
'    Sbo_Application.SendKeys ("Enter")
    oForm01.Freeze False

    Set oRecordSet = Nothing
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Search_Matrix_Data_Error:
    Set oRecordSet = Nothing
    MDC_Com.MDC_GF_Message "Display_Data_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

