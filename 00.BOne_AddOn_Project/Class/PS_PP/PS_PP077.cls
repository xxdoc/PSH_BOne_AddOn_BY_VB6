VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP077"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_PP077.cls
'//  Module         : PP
'//  Description    : 휘팅서울포장등록
'//  FormType       : PS_PP077
'//  Create Date    : 2010.10.26
'//  Modified Date  :
'//  Creator        : Youn Je Hyung
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Public oGrid01              As SAPbouiCOM.Grid
Private oDS_PS_PP077H       As SAPbouiCOM.DBDataSource    '등록헤더

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private VIDocNum$, VODocNum$

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP077.srf")
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
    
    oFormUniqueID01 = "PS_PP077_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
'    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
'    Call FormClear
'    Call Matrix_AddRow(1, 0, True) 'oMat01
'    FormItemEnabled

    oForm01.Items("MovDocNo").Click ct_Regular
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
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
    Set oDS_PS_PP077H = oForm01.DataSources.DBDataSources("@PS_PP077H")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    Set oGrid01 = oForm01.Items("Grid01").Specific
    oForm01.DataSources.DataTables.Add ("ZTEMP")

    Call oForm01.DataSources.UserDataSources.Add("FrDate", dt_DATE, 10)
    oForm01.Items("FrDate").Specific.DataBind.SetBound True, "", "FrDate"
    oForm01.DataSources.UserDataSources.Item("FrDate").VALUE = Format(Now, "YYYYMM") + "01"

    Call oForm01.DataSources.UserDataSources.Add("ToDate", dt_DATE, 10)
    oForm01.Items("ToDate").Specific.DataBind.SetBound True, "", "ToDate"
    oForm01.DataSources.UserDataSources.Item("ToDate").VALUE = Format(Now, "YYYYMMDD")
    
    Call oForm01.DataSources.UserDataSources.Add("DocDate", dt_DATE, 10)
    oForm01.Items("DocDate").Specific.DataBind.SetBound True, "", "DocDate"
    oForm01.DataSources.UserDataSources.Item("DocDate").VALUE = Format(Now, "YYYYMMDD")
    
'    oDS_PS_PP077H.setValue "U_RegiDate", 0, Format(Now, "YYYYMMDD")
    
    '작업자 기본설정
    oForm01.Items("EmpId").Specific.VALUE = MDC_PS_Common.User_MSTCOD()

    oMat01.Columns("NPkQty").Editable = True
    oMat01.Columns("NPkWt").Editable = True
    oMat01.Columns("InDate").Editable = True
    
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
'    Set oCombo = oForm01.Items("BPLId").Specific
'    sQry = "SELECT BPLId, BPLName From [OBPL] order by 1"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
'        oRecordSet01.MoveNext
'    Loop
    
    Set oCombo = oForm01.Items("DocStatus").Specific
    oCombo.ValidValues.Add "1", "포장대기"
    oCombo.ValidValues.Add "2", "포장완료"
    oCombo.Select 0, psk_Index
    
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
    Dim vReturnValue$
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                
                If pval.ItemUID = "Search" Then '조회버튼
                    Call Search_Grid_Data
                
                ElseIf pval.ItemUID = "Distribute" Then  '전량배부버튼
                    If oForm01.Items("DistrQty").Specific.VALUE <> "" And oForm01.Items("DistrQty").Specific.VALUE <> 0 Then
                        Call Distiribute_Data
                    End If
                
                ElseIf pval.ItemUID = "Save" Then '선택저장버튼
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                
                    If MatrixSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                                        
                    If Save_Data("A") = False Then '//이동요청중량 보다 기포장중량+포장중량이 클수없음, Check!!!
                        BubbleEvent = False
                        Exit Sub
                    End If
                    
                    vReturnValue = Sbo_Application.MessageBox("이 데이터를 추가한 후에는 변경할 수 없습니다. 계속하겠습니까?", 1, "&확인", "&취소")
                    Select Case vReturnValue
                        Case 1:
                            If Insert_oInventoryGenEntry(2) = False Then
                                BubbleEvent = False
                                Exit Sub
                            Else
'                                Sbo_Application.MessageBox ("선택저장 작업이 완료되었습니다.")
                                MDC_Com.MDC_GF_Message "선택저장 작업이 완료되었습니다.", "S"
                                Call Search_Grid_Data '새로고침
                            End If
                        Case 2:
                            BubbleEvent = False
                            Exit Sub
                    End Select
                End If
                
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    '//헤더
                    If pval.ItemUID = "ItemCode" Then
                        If oForm01.Items("ItemCode").Specific.VALUE = "" Then
'                            Sbo_Application.ActivateMenuItem ("7425")
'                            BubbleEvent = False
                            Dim ChildForm01 As Variant
                            Set ChildForm01 = New PS_SM010
                            Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
                        End If
                    End If
                    If pval.ItemUID = "EmpId" Then
                        If oForm01.Items("EmpId").Specific.VALUE = "" Then
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
                If pval.ItemUID = "DocStatus" Then
                    oForm01.Freeze True
                    If oForm01.Items("DocStatus").Specific.Selected.VALUE = "1" Then
                        oForm01.Items("s2").Specific.Caption = "이동등록일"
                        oForm01.Items("Save").Enabled = True
                    ElseIf oForm01.Items("DocStatus").Specific.Selected.VALUE = "2" Then
                        oForm01.Items("s2").Specific.Caption = "포장등록일"
                        oForm01.Items("Save").Enabled = False
                    End If
                    
                    If oGrid01.Columns.Count <> 0 Then oGrid01.DataTable.Clear
                    oMat01.Clear
                    oForm01.Freeze False
                End If
            Case et_CLICK: '//6
                If pval.ItemUID = "Grid01" Then
                    Call Search_Matrix_Data
                End If
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    '//헤더
                    If pval.ItemUID = "ItemCode" Then
                        FlushToItemValue pval.ItemUID
                    End If
                    If pval.ItemUID = "EmpId" Then
                        FlushToItemValue pval.ItemUID
                    End If
                    If pval.ItemUID = "DocDate" Then
                        FlushToItemValue pval.ItemUID
                    End If
                    '//라인
                    If pval.ItemUID = "Mat01" And (pval.ColUID = "PP070No" Or pval.ColUID = "NPkQty" Or pval.ColUID = "NPkWt") Then
                        Call FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                    End If
                End If
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
                Call Form_Resize(FormUID, pval, BubbleEvent)
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
    Dim vReturnValue$
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
                If HeaderSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
                If MatrixSpaceLineDel = False Then
                    BubbleEvent = False
                    Exit Sub
                End If
            
                vReturnValue = Sbo_Application.MessageBox("이 데이터를 취소한 후에는 변경할 수 없습니다. 계속하겠습니까?", 1, "&확인", "&취소")
                Select Case vReturnValue
                Case 1:
                    If Insert_oInventoryGenExit(2) = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        Sbo_Application.MessageBox ("선택취소 작업이 완료되었습니다.")
                    End If
                Case 2:
                    BubbleEvent = False
                    Exit Sub
            End Select
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
'                If oMat01.RowCount <> oMat01.VisualRowCount Then
'                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'                    '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
'                    '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
'                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'                    For i = 0 To oMat01.VisualRowCount - 1
'                        oMat01.Columns("LineId").Cells(i + 1).Specific.Value = i + 1
'                    Next i
'
'                    oMat01.FlushToDataSource
'                    oDS_PS_PP077H.RemoveRecord oDS_PS_PP077H.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
'                    oMat01.Clear
'                    oMat01.LoadFromDataSource
'                End If
            
            Case "1281": '찾기
'                FormItemEnabled
'                oForm01.Items("DocNum").Click ct_Regular
                
            Case "1282": '추가
'                Call FormItemEnabled
'                Call FormClear
'                oDS_PS_PP077H.setValue "U_RegiDate", 0, Format(Now, "YYYYMMDD")
'                Call Matrix_AddRow(1, 0, True) 'oMat01
                
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
'                FormItemEnabled
'                If oMat01.VisualRowCount > 0 Then
'                    If oMat01.Columns("PP070No").Cells(oMat01.VisualRowCount).Specific.Value <> "" Then
'                        If oDS_PS_PP077H.GetValue("Status", 0) = "O" Then
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
        oForm01.Items("DocNum").Enabled = True
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("DeliArea").Enabled = True
        oForm01.Items("CarNo").Enabled = True
        oForm01.Items("TransCom").Enabled = True
        oForm01.Items("Fee").Enabled = True
        oForm01.Items("RegiDate").Enabled = True
        oMat01.Columns("PP070No").Editable = True
        
    ElseIf oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("DocNum").Enabled = False
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("CntcCode").Enabled = True
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("DeliArea").Enabled = True
        oForm01.Items("CarNo").Enabled = True
        oForm01.Items("TransCom").Enabled = True
        oForm01.Items("Fee").Enabled = True
        oForm01.Items("RegiDate").Enabled = True
        oMat01.Columns("PP070No").Editable = True
        
    ElseIf oForm01.Mode = fm_OK_MODE Then
        If oForm01.Items("Canceled").Specific.VALUE = "Y" Then
            oForm01.Items("Comments").Click ct_Regular
            oForm01.Items("DocNum").Enabled = False
            oForm01.Items("BPLId").Enabled = False
            oForm01.Items("CntcCode").Enabled = False
            oForm01.Items("CardCode").Enabled = False
            oForm01.Items("DeliArea").Enabled = False
            oForm01.Items("CarNo").Enabled = False
            oForm01.Items("TransCom").Enabled = False
            oForm01.Items("Fee").Enabled = False
            oForm01.Items("RegiDate").Enabled = False
            oMat01.Columns("PP070No").Editable = False
        Else
            oForm01.Items("DocNum").Enabled = True
            oForm01.Items("BPLId").Enabled = True
            oForm01.Items("CntcCode").Enabled = True
            oForm01.Items("CardCode").Enabled = True
            oForm01.Items("DeliArea").Enabled = True
            oForm01.Items("CarNo").Enabled = True
            oForm01.Items("TransCom").Enabled = True
            oForm01.Items("Fee").Enabled = True
            oForm01.Items("RegiDate").Enabled = True
            oMat01.Columns("PP070No").Editable = False
        End If
    End If
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'*******************************************************************
'// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
'*******************************************************************
Private Sub Matrix_AddRow(ByVal oMat As Integer, ByVal oRow As Long, Optional Insert_YN As Boolean)
On Error GoTo Matrix_AddRow_Error
        
    Select Case oMat
'        Case 1: 'oMat01
'            If Insert_YN = False Then
'                oRow = oMat01.RowCount
'                oDS_PS_PP077L.InsertRecord (oRow)
'            End If
'            '수입내역
'            oDS_PS_PP077L.Offset = oRow
'            oDS_PS_PP077L.setValue "LineId", oRow, oRow + 1
'            oDS_PS_PP077L.setValue "U_PP070No", oRow, ""
'            oDS_PS_PP077L.setValue "U_ItemCode", oRow, ""
'            oDS_PS_PP077L.setValue "U_ItemName", oRow, ""
'            oDS_PS_PP077L.setValue "U_Size", oRow, ""
'            oDS_PS_PP077L.setValue "U_Mark", oRow, ""
'            oDS_PS_PP077L.setValue "U_Qty", oRow, ""
'            oDS_PS_PP077L.setValue "U_Weight", oRow, ""
'            oDS_PS_PP077L.setValue "U_DocDate", oRow, ""
'            oMat01.LoadFromDataSource

'        Case 2: 'oMat02
'            If Insert_YN = False Then
'                oRow = oMat2.RowCount
'                oDS_ZPP140M.InsertRecord (oRow)
'            End If
'            '수출내역
'            oDS_ZPP140M.Offset = oRow
'            oDS_ZPP140M.setValue "LineId", oRow, oRow + 1
'            oDS_ZPP140M.setValue "U_ConfDate", oRow, ""
'            oDS_ZPP140M.setValue "U_ConfNo", oRow, ""
'            oDS_ZPP140M.setValue "U_Size", oRow, ""
'            oDS_ZPP140M.setValue "U_ExpQty", oRow, ""
'            oDS_ZPP140M.setValue "U_RfndQty", oRow, ""
'            oMat02.LoadFromDataSource
        
    End Select
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Matrix_AddRow_Error:
    MDC_Com.MDC_GF_Message "Matrix_AddRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&, DocNum$, LineId$
    Dim SumNPkQty           As Double '포장수량합계
    Dim SumNPkWt            As Double '포장중량합계
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordset    As SAPbobsCOM.Recordset
        
    Set oRecordset = Sbo_Company.GetBusinessObject(BoRecordset)
    '--------------------------------------------------------------
    'Header--------------------------------------------------------
    Select Case oUID
        Case "ItemCode"
            sQry = "Select ItemName From OITM Where ItemCode = '" & Trim(oForm01.Items("ItemCode").Specific.String) & "'"
            oRecordset.DoQuery sQry
            oForm01.Items("ItemName").Specific.String = Trim(oRecordset.Fields(0).VALUE)
        Case "EmpId"
            sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" & Trim(oForm01.Items("EmpId").Specific.String) & "'"
            oRecordset.DoQuery sQry
            oForm01.Items("EmpName").Specific.String = Trim(oRecordset.Fields(0).VALUE)
        Case "DocDate"
                oForm01.Freeze True
                oMat01.FlushToDataSource
                For i = 0 To oMat01.VisualRowCount - 1
                    oDS_PS_PP077H.Offset = i
                    oDS_PS_PP077H.setValue "U_InDate", i, Format(oForm01.Items("DocDate").Specific.VALUE, "00000000")
                Next i
                oMat01.LoadFromDataSource
                oForm01.Freeze False
    End Select
        
    '--------------------------------------------------------------
    'Line----------------------------------------------------------
    If oUID = "Mat01" Then
        Select Case oCol
            Case "NPkQty", "NPkWt"
                oMat01.FlushToDataSource
                
                For i = 0 To oMat01.VisualRowCount - 1
                    oDS_PS_PP077H.Offset = i
                    
                    SumNPkQty = SumNPkQty + Trim(oDS_PS_PP077H.GetValue("U_NPkQty", i))
                    SumNPkWt = SumNPkWt + Trim(oDS_PS_PP077H.GetValue("U_NPkWt", i))
                Next i

                oForm01.Items("SumNPkQty").Specific.VALUE = SumNPkQty
                oForm01.Items("SumNPkWt").Specific.VALUE = SumNPkWt
                
                '--------------------------------------------------------------------------------------------
'                If oRow = oMat01.RowCount And Trim(oDS_PS_PP077L.GetValue("U_PP070No", oRow - 1)) <> "" Then
'                '// 다음 라인 추가
'                    Call Matrix_AddRow(1, 0, False)
'                    oMat01.Columns("PP070No").Cells(oRow).Click ct_Regular
'                End If

         End Select
    End If
     
    Set oRecordset = Nothing
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
        Case Trim(oForm01.Items("DocDate").Specific.VALUE) = ""
            oForm01.Items("DocDate").Click ct_Regular
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("EmpId").Specific.VALUE) = ""
            oForm01.Items("EmpId").Click ct_Regular
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("EmpName").Specific.VALUE) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "입고전기일은 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "작업자코드는 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "작업자명이 없습니다. 작업자코드를 확인하여 주십시오.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error
    Dim i               As Long
    Dim ErrNum          As Integer
    Dim oRecordset      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim j               As Integer
    
    Set oRecordset = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0
    
    oMat01.FlushToDataSource

    '// 라인
    If oMat01.VisualRowCount = 0 Then '// MAT01에 값이 있는지 확인 (ErrorNumber : 1)
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//마지막 행 하나를 빼고 i=0부터 시작하므로 하나를 빼므로
    '//oMat01.RowCount - 2가 된다..반드시 들어 가야 하는 필수값을 확인한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    If oMat01.VisualRowCount > 0 Then
        For i = 0 To oMat01.VisualRowCount - 1      '// Mat1에 입력값이 올바르게 들어갔는지 확인 (ErrorNumber : 2)
            If Trim(oDS_PS_PP077H.GetValue("U_Check", i)) = "Y" Then
                oDS_PS_PP077H.Offset = i
                If Trim(oDS_PS_PP077H.GetValue("U_NPkQty", i)) = "" Or Trim(oDS_PS_PP077H.GetValue("U_NPkQty", i)) = 0 Then
                    ErrNum = 2
                    oMat01.Columns("NPkQty").Cells(i + 1).Click ct_Regular
                    GoTo MatrixSpaceLineDel_Error
                ElseIf Trim(oDS_PS_PP077H.GetValue("U_NPkWt", i)) = "" Or Trim(oDS_PS_PP077H.GetValue("U_NPkWt", i)) = 0 Then
                    ErrNum = 3
                    oMat01.Columns("NPkWt").Cells(i + 1).Click ct_Regular
                    GoTo MatrixSpaceLineDel_Error
                End If
                
                j = j + 1 '체크된 라인Count
            End If
        Next i
    End If
    
    '// 체크된 라인Count
    If j = 0 Then
        ErrNum = 4
        GoTo MatrixSpaceLineDel_Error
    End If
        
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
    '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'    If oMat01.VisualRowCount > 0 Then
'        oDS_PS_PP077L.RemoveRecord oDS_PS_PP077L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
'    End If
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '행을 삭제하였으니 DB데이터 소스를 다시 가져온다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oMat01.LoadFromDataSource

    Set oRecordset = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordset = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "포장수량이 0 또는 없습니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "포장중량이 0 또는 없습니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "선택되어진 라인이 없습니다. 확인하여 주십시오.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Sub Form_Resize(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Form_Resize_Error
    Dim i   As Integer
    
    
    oForm01.Items("Grid01").Top = 40
    oForm01.Items("Grid01").Height = (oForm01.Height / 2) - 75
    oForm01.Items("Grid01").Left = 6
    oForm01.Items("Grid01").Width = oForm01.Width - 21
    
    oForm01.Items("DistrQty").Top = (oForm01.Height / 2) - 32
    oForm01.Items("Distribute").Top = (oForm01.Height / 2) - 37
    
    oForm01.Items("s9").Top = (oForm01.Height / 2) - 32
    oForm01.Items("s9").Left = (oForm01.Width / 2) - 65
    
    oForm01.Items("EmpId").Top = (oForm01.Height / 2) - 32
    oForm01.Items("EmpId").Left = (oForm01.Width / 2) + 16
    
    oForm01.Items("EmpName").Top = (oForm01.Height / 2) - 32
    oForm01.Items("EmpName").Left = (oForm01.Width / 2) + 116
    
    oForm01.Items("s8").Top = (oForm01.Height / 2) - 32
    oForm01.Items("s8").Left = (oForm01.Width) - 190
    
    oForm01.Items("DocDate").Top = (oForm01.Height / 2) - 32
    oForm01.Items("DocDate").Left = (oForm01.Width) - 112
    
    oForm01.Items("Mat01").Top = (oForm01.Height / 2) - 15
    oForm01.Items("Mat01").Height = (oForm01.Height / 2) - 38
    oForm01.Items("Mat01").Left = 6
    oForm01.Items("Mat01").Width = oForm01.Width - 21
    
    If oGrid01.Rows.Count > 0 Then
        oGrid01.AutoResizeColumns
    End If
    oMat01.AutoResizeColumns
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Form_Resize_Error:
    Sbo_Application.SetStatusBarMessage "Form_Resize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Search_Grid_Data()
    On Error GoTo Search_Grid_Data_Error
    Dim oRecordset          As SAPbobsCOM.Recordset
    Dim sQry                As String
    Dim MovDocNo            As String
    Dim FrDate              As String
    Dim ToDate              As String
    Dim ItemCode            As String
    Dim Status              As String
    
    Set oRecordset = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oForm01.Freeze True
    
    oForm01.Items("DistrQty").Specific.VALUE = ""
    oForm01.Items("SumPkQty").Specific.VALUE = ""
    oForm01.Items("SumPkWt").Specific.VALUE = ""
    oForm01.Items("SumNPkQty").Specific.VALUE = ""
    oForm01.Items("SumNPkWt").Specific.VALUE = ""
    oForm01.Items("EmpId").Specific.VALUE = MDC_PS_Common.User_MSTCOD()
    Call FlushToItemValue("EmpId")
    oForm01.Items("DocDate").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    oMat01.Clear
    
    MovDocNo = Trim(oForm01.Items("MovDocNo").Specific.String)
    If MovDocNo = "" Then MovDocNo = "%"
    
    FrDate = Trim(oForm01.Items("FrDate").Specific.VALUE)
    ToDate = Trim(oForm01.Items("ToDate").Specific.VALUE)
    If FrDate = "" Then FrDate = "1900-01-01"
    If ToDate = "" Then ToDate = "2100-12-31"
                        
    ItemCode = Trim(oForm01.Items("ItemCode").Specific.String)
    If ItemCode = "" Then ItemCode = "%"
    
    Status = Trim(oForm01.Items("DocStatus").Specific.Selected.VALUE)
    
    sQry = "EXEC PS_PP077_01 '" & MovDocNo & "','" & FrDate & "','" & ToDate & "','" & ItemCode & "','" & Status & "'"
        
    '/ Procedure 실행(Grid 사용)
    oForm01.DataSources.DataTables.Item(0).ExecuteQuery (sQry)
    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("ZTEMP")
    
    Call GridSetting
    
    
    oForm01.Freeze False
    
    Set oRecordset = Nothing
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Search_Grid_Data_Error:
        Set oRecordset = Nothing
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
    
    oGrid01.SelectionMode = ms_Single
    
'    oGrid01.Columns.Item(0).LinkedObjectType = lf_BusinessPartner
'    oGrid01.Columns.Item(4).LinkedObjectType = lf_Invoice
     oGrid01.Columns.Item(2).LinkedObjectType = lf_Items
       
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

Private Sub Search_Matrix_Data()
    On Error GoTo Search_Matrix_Data_Error
    Dim oRecordset          As SAPbobsCOM.Recordset
    Dim sQry                As String
    
    Dim i                   As Integer
    Dim j                   As Integer
    Dim Cnt                 As Integer
    Dim SumPkQty            As Double '이동요청수량합계
    Dim SumPkWt             As Double '이동요청중량합계
    Dim SumNPkQty           As Double '포장수량합계
    Dim SumNPkWt            As Double '포장중량합계
    
    
    Set oRecordset = Sbo_Company.GetBusinessObject(BoRecordset)
    oForm01.Freeze True
    
     For i = 0 To oGrid01.Rows.Count - 1
        If oGrid01.Rows.IsSelected(i) = True Then
            
            '/ Procedure 실행
            sQry = "EXEC PS_PP077_02 '" & Trim(oGrid01.DataTable.GetValue(0, i)) & "', '" & Trim(oGrid01.DataTable.GetValue(1, i)) & "', '" & Trim(oGrid01.DataTable.GetValue(2, i)) & "'"
            oRecordset.DoQuery sQry

            
             'Matrix에 출력===========================================================================================
             '// Line 초기화
             Cnt = oDS_PS_PP077H.Size
             If Cnt > 0 Then
                 For j = 0 To Cnt - 1
                     oDS_PS_PP077H.RemoveRecord oDS_PS_PP077H.Size - 1
                 Next j
                 If Cnt = 1 Then
                     oDS_PS_PP077H.Clear
                 End If
             End If
             oMat01.LoadFromDataSource
             '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                 
            '//Matrix에 Data 뿌려준다
            j = 1
             Do Until oRecordset.EOF
                     
                 If oDS_PS_PP077H.Size < j Then
                     oDS_PS_PP077H.InsertRecord j - 1 '라인추가
                 End If
                 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                 oDS_PS_PP077H.setValue "DocNum", j - 1, j
                 oDS_PS_PP077H.setValue "U_Check", j - 1, "Y"
                 oDS_PS_PP077H.setValue "U_MovDocNo", j - 1, oRecordset.Fields(0).VALUE
                 oDS_PS_PP077H.setValue "U_PP070No", j - 1, oRecordset.Fields(1).VALUE
                 oDS_PS_PP077H.setValue "U_PP070NoL", j - 1, oRecordset.Fields(2).VALUE
                 oDS_PS_PP077H.setValue "U_PorNum", j - 1, oRecordset.Fields(3).VALUE
                 oDS_PS_PP077H.setValue "U_ItemCode", j - 1, oRecordset.Fields(4).VALUE
                 oDS_PS_PP077H.setValue "U_ItemName", j - 1, oRecordset.Fields(5).VALUE
                 oDS_PS_PP077H.setValue "U_PkQty", j - 1, oRecordset.Fields(6).VALUE
                 oDS_PS_PP077H.setValue "U_PkWt", j - 1, oRecordset.Fields(7).VALUE
                 oDS_PS_PP077H.setValue "U_NPkWt", j - 1, oRecordset.Fields(8).VALUE
                 oDS_PS_PP077H.setValue "U_OPkQty", j - 1, oRecordset.Fields(9).VALUE
                 oDS_PS_PP077H.setValue "U_OPkWt", j - 1, oRecordset.Fields(10).VALUE
                 If oForm01.Items("DocStatus").Specific.Selected.VALUE = "1" Then '포장대기
                    oDS_PS_PP077H.setValue "U_InDate", j - 1, Format(Now, "YYYYMMDD")
                 Else '포장완료
                    oDS_PS_PP077H.setValue "U_InDate", j - 1, Format(oRecordset.Fields(11).VALUE, "YYYYMMDD")
                 End If
                 
'                 SumPkQty = SumPkQty + oRecordset.Fields(6).Value
'                 SumPkWt = SumPkWt + oRecordset.Fields(7).Value
                 
'                 SumPkQty = oRecordset.Fields(6).Value '이동요청수량합계
'                 SumPkWt = oRecordset.Fields(7).Value  '이동요청중량합계
                 
                 SumPkQty = Trim(oGrid01.DataTable.GetValue(7, i))
                 SumPkWt = Trim(oGrid01.DataTable.GetValue(8, i))
                 
                 SumNPkWt = SumNPkWt + oRecordset.Fields(8).VALUE '포장중량합계
                 
                 j = j + 1
                 oRecordset.MoveNext
             Loop
             oMat01.LoadFromDataSource
             ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
     Next i
    
    oForm01.Items("SumPkQty").Specific.VALUE = SumPkQty
    oForm01.Items("SumPkWt").Specific.VALUE = SumPkWt
    oForm01.Items("SumNPkQty").Specific.VALUE = 0
    oForm01.Items("SumNPkWt").Specific.VALUE = SumNPkWt
    
    oForm01.Freeze False
    Set oRecordset = Nothing
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Search_Matrix_Data_Error:
    Set oRecordset = Nothing
    MDC_Com.MDC_GF_Message "Search_Matrix_Data_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub Distiribute_Data() '포장중량을 기준으로 포장수량 전량배부
    On Error GoTo Distiribute_Data_Error
    Dim oRecordset          As SAPbobsCOM.Recordset
    Dim sQry                As String
    Dim i                   As Integer
    Dim DocNum              As Integer
    Dim SumPkWt             As Double
    Dim SumNPkWt            As Double
    Dim SumNPkQty           As Double
    Dim DistrQty            As Double
    
    Set oRecordset = Sbo_Company.GetBusinessObject(BoRecordset)
    oForm01.Freeze True
    
    oMat01.FlushToDataSource
    If oMat01.VisualRowCount > 0 Then
    
'        SumPkWt = Trim(oForm01.Items("SumPkWt").Specific.Value)
        SumNPkWt = Trim(oForm01.Items("SumNPkWt").Specific.VALUE)
        DistrQty = Trim(oForm01.Items("DistrQty").Specific.VALUE)
    
        For i = 0 To oMat01.VisualRowCount - 1
            oDS_PS_PP077H.Offset = i
            oDS_PS_PP077H.setValue "U_NPkQty", i, Round(DistrQty * (Trim(oDS_PS_PP077H.GetValue("U_NPkWt", i)) / SumNPkWt), 0)
            
            SumNPkQty = SumNPkQty + Round(DistrQty * (Trim(oDS_PS_PP077H.GetValue("U_NPkWt", i)) / SumNPkWt), 0)
        Next i
        
        oForm01.Items("SumNPkQty").Specific.VALUE = SumNPkQty
    End If
    
    oMat01.LoadFromDataSource
    
    oForm01.Freeze False
    Set oRecordset = Nothing
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Distiribute_Data_Error:
    Set oRecordset = Nothing
    MDC_Com.MDC_GF_Message "Distiribute_Data_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function Insert_oInventoryGenEntry(ChkType As Integer) As Boolean
On Error GoTo error_Message
'============================================================================
'   입고 DI
'============================================================================
    Dim DI_oInventoryGenEntry As SAPbobsCOM.Documents    '//재고입고 문서 객체
    Dim RetVal&
'    Dim VIDocNum$, VODocNum$
    Dim oRecordset      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrCode         As Long
    Dim ErrMsg          As String
    Dim i&, K&, oRow&
    Dim ErrNum          As Integer
    
    Set oRecordset = Sbo_Company.GetBusinessObject(BoRecordset)
    '/////////////////////////////////////////////////////////////////////////////////
    '입고
    '/////////////////////////////////////////////////////////////////////////////////
    Sbo_Company.StartTransaction
    Set DI_oInventoryGenEntry = Nothing
    Set DI_oInventoryGenEntry = Sbo_Company.GetBusinessObject(oInventoryGenEntry)   '// 문서타입(입고)

    i = 1
    With DI_oInventoryGenEntry
        
        '/ Header
        .DocDate = Format(Trim(oForm01.Items("DocDate").Specific.VALUE), "0000-00-00")
        .TaxDate = Format(Trim(oForm01.Items("DocDate").Specific.VALUE), "0000-00-00")
        .Comments = "포장처리등록 입고"
'        .Comments = "품목번호 변경(" & ODocNum & ") 입고"
    
        '/ Line
        For oRow = 0 To oDS_PS_PP077H.Size - 1
            If Trim(oDS_PS_PP077H.GetValue("U_Check", oRow)) = "Y" Then
                If .Lines.Count < i Then
                   .Lines.Add
                End If
                
                .Lines.SetCurrentLine i - 1
                .Lines.ItemCode = Trim(oDS_PS_PP077H.GetValue("U_ItemCode", oRow))
                .Lines.ItemDescription = Trim(oDS_PS_PP077H.GetValue("U_ItemName", oRow))
                .Lines.Quantity = Val(oDS_PS_PP077H.GetValue("U_NPkQty", oRow))
                .Lines.WarehouseCode = "104" '제품-서울
                
'                .Lines.PriceAfterVAT = Val(oDS_PS_PP077H.GetValue("U_Price", oRow))
'                .Lines.UserFields("U_Number").Value = Val(oDS_PS_PP077H.GetValue("U_Number", oRow))
'                .Lines.AccountCode = Trim(ZPSH_SetFunc.Get_ReData("StkInTnAct", "WhsCode", "OWHS", "'" & Trim(oDS_PS_PP077H.GetValue("U_AWhsCode", oRow)) & "'"))
'                .Lines.Currency = "MYR"
                
                i = i + 1
            End If
        Next oRow
    End With

    '/-------------------------------------------------------------------------/
    '/ 입고전송
'    If Trim(oForm01.Items("IDocNum").Specific.Value) = "" Then
        RetVal = DI_oInventoryGenEntry.Add
        If (0 <> RetVal) Then
             Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
             GoTo error_Message  '//저장시 에러 발생
        End If
        If ChkType <> 2 Then
            Sbo_Company.EndTransaction wf_RollBack
        Else
            Sbo_Company.GetNewObjectCode VIDocNum
'            oForm01.Items("IDocNum").Specific.Value = VIDocNum

             If Save_Data("B") = False Then '//[@PS_PP077H]에 Insert!!!
                 ErrNum = 100
                 GoTo error_Message
             End If

'             If Add_PS_PP040() = False Then '//[@PS_PP040H](작업지시)에 Insert!!!
'                If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
'                Insert_oInventoryGenEntry = False
'                Exit Function
'             End If
             
             Sbo_Company.EndTransaction wf_Commit
        End If
'     End If
     
    Insert_oInventoryGenEntry = True
    Set DI_oInventoryGenEntry = Nothing
    Set oRecordset = Nothing
    
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
error_Message:
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Insert_oInventoryGenEntry = False
    
    If ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "Add_PS_PP040_Error:" & Err.Number & " - " & Err.Description, "E"
    ElseIf ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "Save_Data_Error:" & Err.Number & " - " & Err.Description, "E"
    ElseIf ErrNum = 100 Then
    Else
        MDC_Com.MDC_GF_Message "Insert_oInventoryGenEntry_Error:" & ErrCode & " - " & ErrMsg, "E"
    End If
        
    Set DI_oInventoryGenEntry = Nothing
    Set oRecordset = Nothing
End Function

Private Function Insert_oInventoryGenExit(ChkType As Integer) As Boolean
On Error GoTo error_Message
'============================================================================
'   출고 DI
'============================================================================
    Dim DI_oInventoryGenExit As SAPbobsCOM.Documents    '//재고입고 문서 객체
    Dim RetVal&
    Dim VIDocNum$, VODocNum$
    Dim oRecordset      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrCode         As Long
    Dim ErrMsg          As String
    Dim i&, K&, oRow&
    Dim ErrNum          As Integer
    
    Set oRecordset = Sbo_Company.GetBusinessObject(BoRecordset)
    '/////////////////////////////////////////////////////////////////////////////////
    '출고
    '/////////////////////////////////////////////////////////////////////////////////
    Sbo_Company.StartTransaction
    Set DI_oInventoryGenExit = Nothing
    Set DI_oInventoryGenExit = Sbo_Company.GetBusinessObject(oInventoryGenExit)   '// 문서타입(입고)

    i = 1
    With DI_oInventoryGenExit
        
        '/ Header
        .DocDate = Format(Trim(oForm01.Items("DocDate").Specific.VALUE), "0000-00-00")
        .TaxDate = Format(Trim(oForm01.Items("DocDate").Specific.VALUE), "0000-00-00")
        .Comments = "포장처리등록 입고"
'        .Comments = "품목번호 변경(" & ODocNum & ") 입고"
    
        '/ Line
        For oRow = 0 To oDS_PS_PP077H.Size - 1
            If Trim(oDS_PS_PP077H.GetValue("U_Check", oRow)) = "Y" Then
                If .Lines.Count < i Then
                   .Lines.Add
                End If
                
                .Lines.SetCurrentLine i - 1
                .Lines.ItemCode = Trim(oDS_PS_PP077H.GetValue("U_ItemCode", oRow))
                .Lines.ItemDescription = Trim(oDS_PS_PP077H.GetValue("U_ItemName", oRow))
                .Lines.Quantity = Val(oDS_PS_PP077H.GetValue("U_NPkQty", oRow))
                .Lines.WarehouseCode = "104" '제품-서울
                
'                .Lines.PriceAfterVAT = Val(oDS_PS_PP077H.GetValue("U_Price", oRow))
'                .Lines.UserFields("U_Number").Value = Val(oDS_PS_PP077H.GetValue("U_Number", oRow))
'                .Lines.AccountCode = Trim(ZPSH_SetFunc.Get_ReData("StkInTnAct", "WhsCode", "OWHS", "'" & Trim(oDS_PS_PP077H.GetValue("U_AWhsCode", oRow)) & "'"))
'                .Lines.Currency = "MYR"
                
                i = i + 1
            End If
        Next oRow
    End With

    '/-------------------------------------------------------------------------/
    '/ 출고전송
'    If Trim(oForm01.Items("IDocNum").Specific.Value) = "" Then
        RetVal = DI_oInventoryGenExit.Add
        If (0 <> RetVal) Then
             Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
             GoTo error_Message  '//저장시 에러 발생
        End If
        If ChkType <> 2 Then
            Sbo_Company.EndTransaction wf_RollBack
        Else
'             Sbo_Company.GetNewObjectCode VIDocNum
'            oForm01.Items("IDocNum").Specific.Value = VIDocNum

             If Save_Data("B", True) = False Then '//[@PS_PP077H]에 -수량으로 Insert!!!
                 ErrNum = 2
                 GoTo error_Message
             End If
             
             Sbo_Company.EndTransaction wf_Commit
        End If
'     End If
     
    Insert_oInventoryGenExit = True
    Set DI_oInventoryGenExit = Nothing
    Set oRecordset = Nothing
    
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
error_Message:
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Insert_oInventoryGenExit = False
    
    If ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "Save_Data_Error:" & Err.Number & " - " & Err.Description, "E"
    Else
'        MDC_Com.MDC_GF_Message "Insert_oInventoryGenEntry_Error:" & Err.Number & " - " & Err.Description, "E"
        MDC_Com.MDC_GF_Message "Insert_oInventoryGenEntry_Error:" & ErrCode & " - " & ErrMsg, "E"
    End If
        
    Set DI_oInventoryGenExit = Nothing
    Set oRecordset = Nothing
End Function

Private Function Save_Data(ByVal A_B As String, Optional Cancel As Boolean) As Boolean
    On Error GoTo Save_Data_Error
    Dim oRecordset          As SAPbobsCOM.Recordset
    Dim sQry                As String
    Dim i                   As Integer
    Dim j                   As Integer
    Dim ErrNum              As Integer
    Dim DocNum              As Integer
    Dim PkWt                As Double
    Dim NPkWt               As Double
    Dim OPkWt               As Double
    
    Set oRecordset = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat01.FlushToDataSource
        
    If oMat01.VisualRowCount > 0 Then
        For i = 0 To oMat01.VisualRowCount - 1
            oDS_PS_PP077H.Offset = i
            If Trim(oDS_PS_PP077H.GetValue("U_Check", i)) = "Y" Then
                
                Select Case A_B
                    Case "A": '//이동요청중량 보다 기포장중량+포장중량이 클수없음, Check!!!
                        PkWt = Trim(oDS_PS_PP077H.GetValue("U_PkWt", i))   '이동요청중량
                        NPkWt = Trim(oDS_PS_PP077H.GetValue("U_NPkWt", i)) '포장중량
                        'OPkWt = Trim(oDS_PS_PP077H.GetValue("U_OPkWt", i)) '기포장중량
                        For j = 0 To oMat01.VisualRowCount - 1
                            If Trim(oDS_PS_PP077H.GetValue("U_Check", j)) = "Y" Then
                                OPkWt = OPkWt + Trim(oDS_PS_PP077H.GetValue("U_OPkWt", j)) '기포장중량
                            End If
                        Next j
                        
                        If PkWt < NPkWt + OPkWt Then
                            ErrNum = 1
                            oMat01.Columns("NPkWt").Cells(i + 1).Click ct_Regular
                            GoTo Save_Data_Error
                        End If
                        
                    Case "B": '//[@PS_PP077H]에 Insert!!!
                
                        sQry = "select max(DocNum) MaxDocNum from [@PS_PP077H]"
                        oRecordset.DoQuery sQry
                        
                        If oRecordset.RecordCount = 0 Then
                            DocNum = 1
                        Else
                            DocNum = oRecordset.Fields("MaxDocNum").VALUE + 1
                        End If
                        
                        sQry = "insert into [@PS_PP077H] values ("
                        sQry = sQry & "'" & DocNum & "','" & DocNum & "'"
                        sQry = sQry & ",Null,0,Null,'N','N',Null,Null,Null,'N','O',Null,Null,Null,Null,Null,"
                        sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_Check", i)) & "',"
                        sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_MovDocNo", i)) & "',"
                        sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_PP070No", i)) & "',"
                        sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_PP070NoL", i)) & "',"
                        sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_PorNum", i)) & "',"
                        sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_ItemCode", i)) & "',"
'                        sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_ItemName", i)) & "',"
                        sQry = sQry & "'" & MDC_PS_Common.Make_ItemName(Trim(oDS_PS_PP077H.GetValue("U_ItemName", i))) & "',"
                        sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_PkQty", i)) & "',"
                        sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_PkWt", i)) & "',"
                        
                        If Cancel = True Then '취소 마이너스(-) insert
                            sQry = sQry & "'" & -1 * Trim(oDS_PS_PP077H.GetValue("U_NPkQty", i)) & "',"
                            sQry = sQry & "'" & -1 * Trim(oDS_PS_PP077H.GetValue("U_NPkWt", i)) & "',"
                        Else
                            sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_NPkQty", i)) & "',"
                            sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_NPkWt", i)) & "',"
                        End If
                        sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_OPkQty", i)) & "',"
                        sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_OPkWt", i)) & "',"
                        sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_InDate", i)) & "'"
                        sQry = sQry & ",Null,Null,"
                        sQry = sQry & "'" & VIDocNum & "',"
                        sQry = sQry & "Null)"
                                                
                        oRecordset.DoQuery sQry
                        '--------------------------------------------------------------------------------------
                       
                       '//[@PS_PP040H](작업지시)에 Insert!!!
                        If Add_PS_PP040(i, DocNum) = False Then
                            ErrNum = 100
                            GoTo Save_Data_Error
                        End If
                                              
                End Select
            End If
        Next i
    End If
    
    Set oRecordset = Nothing
    Save_Data = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Save_Data_Error:
    Set oRecordset = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "기포장중량 + 포장중량은 이동요청중량을 초과할수 없습니다.. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "해당작지번호에 대한 작업지시등록 공정중 바렐 또는 포장공정이 없거나 이미등록되어 있습니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 100 Then
    Else
        MDC_Com.MDC_GF_Message "Save_Data_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    Save_Data = False
End Function

Public Function Add_PS_PP040(ByVal i As Integer, PP077HDocNum As Integer) As Boolean
On Error GoTo Add_PS_PP040_Error
    Dim j&
    Dim ErrNum&, RowNum&
    Dim sQry         As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim oRecordSet02 As SAPbobsCOM.Recordset
    Dim oRecordSet03 As SAPbobsCOM.Recordset
    
    Dim lTypeCount   As Integer
    
    Dim BPLId$, JakNum$, OutQty As Long, OutWt As Currency
    Dim PP040H_DocEntry$, PP030M_LineId$, PP030M_U_Sequence$, PP030M_U_CpCode$, PP030M_U_CpName$, AutoKey&
    Dim iTemp01&, sTemp01$
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    Set oRecordSet02 = Sbo_Company.GetBusinessObject(BoRecordset)
    Set oRecordSet03 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat01.FlushToDataSource
    
'    For i = 0 To oMat01.RowCount - 1
        If Trim(oDS_PS_PP077H.GetValue("U_Check", i)) = "Y" Then
            
            sQry = "EXEC [PS_PP077_03] '" & Trim(oDS_PS_PP077H.GetValue("U_MovDocNo", i)) & "','" & Trim(oDS_PS_PP077H.GetValue("U_PP070No", i)) & "','" & Trim(oDS_PS_PP077H.GetValue("U_PP070NoL", i)) & "'"
            oRecordSet01.DoQuery sQry

            '// DocEntry
            sQry = "Select AutoKey From [ONNM] Where ObjectCode = 'PS_PP040'"
            oRecordSet02.DoQuery sQry
            PP040H_DocEntry = Trim(oRecordSet02.Fields("AutoKey").VALUE)
            AutoKey = PP040H_DocEntry + 1

            '// Insert PS_PP040H
            sQry = "INSERT INTO [@PS_PP040H]"
            sQry = sQry & " ("
            sQry = sQry & " DocEntry,"
            sQry = sQry & " DocNum,"
            sQry = sQry & " Period,"
            sQry = sQry & " Series,"
            sQry = sQry & " Object,"
            sQry = sQry & " UserSign,"
            sQry = sQry & " CreateDate,"
            sQry = sQry & " CreateTime,"
            sQry = sQry & " U_OrdType,"
            sQry = sQry & " U_OrdGbn,"
            sQry = sQry & " U_BPLId,"
            sQry = sQry & " U_ItemCode,"
            sQry = sQry & " U_ItemName,"
            sQry = sQry & " U_OrdNum,"
            sQry = sQry & " U_OrdSub1,"
            sQry = sQry & " U_OrdSub2,"
            sQry = sQry & " U_PP030HNo,"
            sQry = sQry & " U_DocType,"
            sQry = sQry & " Canceled,"
            sQry = sQry & " U_DocDate"
            sQry = sQry & " ) "

            sQry = sQry & "VALUES("
            sQry = sQry & "'" & PP040H_DocEntry & "',"
            sQry = sQry & "'" & PP040H_DocEntry & "',"
            sQry = sQry & "'21',"
            sQry = sQry & "'-1',"
            sQry = sQry & "'PS_PP040',"
            sQry = sQry & "'1',"
            sQry = sQry & "'" & Format(Date, "YYYYMMDD") & "',"
            sQry = sQry & "'1000',"
            sQry = sQry & "'40',"
            sQry = sQry & "'101',"
            sQry = sQry & "'" & Trim(oRecordSet01.Fields("U_BPLId").VALUE) & "',"
            sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_ItemCode", i)) & "',"
'            sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_ItemName", i)) & "',"
            sQry = sQry & "'" & MDC_PS_Common.Make_ItemName(Trim(oDS_PS_PP077H.GetValue("U_ItemName", i))) & "',"
            sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_PorNum", i)) & "',"
            sQry = sQry & "'" & Trim(oRecordSet01.Fields("U_OrdSub1").VALUE) & "',"
            sQry = sQry & "'" & Trim(oRecordSet01.Fields("U_OrdSub2").VALUE) & "',"
            sQry = sQry & "'" & Trim(oRecordSet01.Fields("U_PP030HNo").VALUE) & "',"
            sQry = sQry & "'10',"
            sQry = sQry & "'N',"
            sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_InDate", i)) & "'"
            sQry = sQry & ")"
            oRecordSet02.DoQuery sQry

            '// Insert PS_PP040M
            sQry = "INSERT INTO [@PS_PP040M]"
            sQry = sQry & " ("
            sQry = sQry & " DocEntry,"
            sQry = sQry & " LineId,"
            sQry = sQry & " VisOrder,"
            sQry = sQry & " Object,"
            sQry = sQry & " U_LineNum,"
            sQry = sQry & " U_LineId,"
            sQry = sQry & " U_WorkCode,"
            sQry = sQry & " U_WorkName"
            sQry = sQry & " ) "

            sQry = sQry & "VALUES("
            sQry = sQry & "'" & PP040H_DocEntry & "',"
            sQry = sQry & "'1',"
            sQry = sQry & "'0',"
            sQry = sQry & "'PS_PP040',"
            sQry = sQry & "'1',"
            sQry = sQry & "'1',"
            sQry = sQry & "'" & Trim(oForm01.Items("EmpId").Specific.VALUE) & "',"
            sQry = sQry & "'" & Trim(oForm01.Items("EmpName").Specific.VALUE) & "'"
            sQry = sQry & ")"
            oRecordSet02.DoQuery sQry

            
            '바렐 및 포장 등록
            sQry = "EXEC [PS_PP077_04] '" & Trim(oDS_PS_PP077H.GetValue("U_PorNUm", i)) & "'"
            oRecordSet03.DoQuery sQry
            
            If oRecordSet03.RecordCount = 0 Then
                ErrNum = 1
                GoTo Add_PS_PP040_Error
            End If
            
            j = 0
            Do Until oRecordSet03.EOF
                '// Insert PS_PP040L
                sQry = "INSERT INTO [@PS_PP040L]"
                sQry = sQry & " ("
                sQry = sQry & " DocEntry,"
                sQry = sQry & " LineId,"
                sQry = sQry & " VisOrder,"
                sQry = sQry & " Object,"
                sQry = sQry & " U_LineNum,"
                sQry = sQry & " U_LineId,"
                sQry = sQry & " U_OrdMgNum,"
                sQry = sQry & " U_Sequence,"
                sQry = sQry & " U_CpCode,"
                sQry = sQry & " U_CpName,"
                sQry = sQry & " U_OrdGbn,"
                sQry = sQry & " U_BPLId,"
                sQry = sQry & " U_ItemCode,"
                sQry = sQry & " U_ItemName,"
                sQry = sQry & " U_OrdNum,"
                sQry = sQry & " U_OrdSub1,"
                sQry = sQry & " U_OrdSub2,"
                sQry = sQry & " U_PP030HNo,"
                sQry = sQry & " U_PP030MNo,"
                sQry = sQry & " U_PQty,"
                sQry = sQry & " U_PWeight,"
                sQry = sQry & " U_YQty,"
                sQry = sQry & " U_YWeight,"
                sQry = sQry & " U_NQty,"
                sQry = sQry & " U_NWeight"
                sQry = sQry & " ) "
    
                sQry = sQry & "VALUES("
                sQry = sQry & "'" & PP040H_DocEntry & "',"
                sQry = sQry & "'" & j + 1 & "',"
                sQry = sQry & "'" & j & "',"
                sQry = sQry & "'PS_PP040',"
                sQry = sQry & "'" & j + 1 & "',"
                sQry = sQry & "'" & j + 1 & "',"
                sQry = sQry & "'" & Trim(oRecordSet01.Fields("U_PP030HNo").VALUE) & "-" & Trim(oRecordSet03.Fields("LineId").VALUE) & "',"
                sQry = sQry & "'" & Trim(oRecordSet03.Fields("U_Sequence").VALUE) & "',"
                sQry = sQry & "'" & Trim(oRecordSet03.Fields("U_CpCode").VALUE) & "',"
                sQry = sQry & "'" & Trim(oRecordSet03.Fields("U_CpName").VALUE) & "',"
                sQry = sQry & "'101',"
                sQry = sQry & "'" & Trim(oRecordSet01.Fields("U_BPLId").VALUE) & "',"
                sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_ItemCode", i)) & "',"
'                sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_ItemName", i)) & "',"
                sQry = sQry & "'" & MDC_PS_Common.Make_ItemName(Trim(oDS_PS_PP077H.GetValue("U_ItemName", i))) & "',"
                sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_PorNum", i)) & "',"
                sQry = sQry & "'" & Trim(oRecordSet01.Fields("U_OrdSub1").VALUE) & "',"
                sQry = sQry & "'" & Trim(oRecordSet01.Fields("U_OrdSub2").VALUE) & "',"
                sQry = sQry & "'" & Trim(oRecordSet01.Fields("U_PP030HNo").VALUE) & "',"
                sQry = sQry & "'" & Trim(oRecordSet03.Fields("LineId").VALUE) & "',"
                sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_NPkQty", i)) & "',"
                sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_NPkWt", i)) & "',"
                sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_NPkQty", i)) & "',"
                sQry = sQry & "'" & Trim(oDS_PS_PP077H.GetValue("U_NPkWt", i)) & "',"
                sQry = sQry & "'" & 0 & "',"
                sQry = sQry & "'" & 0 & "'"
                sQry = sQry & ")"
                oRecordSet02.DoQuery sQry
    
                '// Insert PS_PP040N
                sQry = "INSERT INTO [@PS_PP040N]"
                sQry = sQry & " ("
                sQry = sQry & " DocEntry,"
                sQry = sQry & " LineId,"
                sQry = sQry & " VisOrder,"
                sQry = sQry & " Object,"
                sQry = sQry & " U_LineNum,"
                sQry = sQry & " U_LineId,"
                sQry = sQry & " U_OrdMgNum,"
                sQry = sQry & " U_CpCode,"
                sQry = sQry & " U_CpName"
                sQry = sQry & " ) "
    
                sQry = sQry & "VALUES("
                sQry = sQry & "'" & PP040H_DocEntry & "',"
                sQry = sQry & "'" & j + 1 & "',"
                sQry = sQry & "'" & j & "',"
                sQry = sQry & "'PS_PP040',"
                sQry = sQry & "'" & j + 1 & "',"
                sQry = sQry & "'" & j + 1 & "',"
                sQry = sQry & "'" & Trim(oRecordSet01.Fields("U_PP030HNo").VALUE) & "-" & Trim(oRecordSet03.Fields("LineId").VALUE) & "',"
                sQry = sQry & "'" & Trim(oRecordSet03.Fields("U_CpCode").VALUE) & "',"
                sQry = sQry & "'" & Trim(oRecordSet03.Fields("U_CpName").VALUE) & "'"
                sQry = sQry & ")"
                oRecordSet02.DoQuery sQry
                
                j = j + 1
                oRecordSet03.MoveNext
            Loop
                        
            '작업일보-문서번호를 [@PS_PP077H]에 update
            sQry = "Update [@PS_PP077H] Set U_PP040No = '" & PP040H_DocEntry & "' "
            sQry = sQry & "Where DocNum = '" & PP077HDocNum & "'"
            oRecordSet02.DoQuery sQry
            
            
            'AutoKey Update
            sQry = "Update [ONNM] Set AutoKey = '" & AutoKey & "' Where ObjectCode = 'PS_PP040'"
            oRecordSet02.DoQuery sQry
        End If
'    Next
       
    Set oRecordSet01 = Nothing
    Set oRecordSet02 = Nothing
    Set oRecordSet03 = Nothing
    Add_PS_PP040 = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_PS_PP040_Error:
    Set oRecordSet01 = Nothing
    Set oRecordSet02 = Nothing
    Set oRecordSet03 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "해당작지번호에 대한 작업지시등록 공정중 바렐 또는 포장공정이 없거나 이미등록되어 있습니다. 확인하여 주십시오.", "E"
    Else
        MDC_Com.MDC_GF_Message "Add_PS_PP040_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    Add_PS_PP040 = False
End Function
