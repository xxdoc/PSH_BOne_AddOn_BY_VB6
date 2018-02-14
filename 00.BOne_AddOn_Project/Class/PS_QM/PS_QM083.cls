VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM083"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_QM083.cls
'//  Module         : QM
'//  Description    : 방산부품 출고검사 등록
'//  FormType       : PS_QM083
'//  Create Date    : 2013.01.26
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_QM083H       As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_QM083L       As SAPbouiCOM.DBDataSource    '등록라인

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM083.srf")
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
    
    oFormUniqueID01 = "PS_QM083_" & GetTotalFormsCount
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
    Set oDS_PS_QM083H = oForm01.DataSources.DBDataSources("@PS_QM083H")
'    Set oDS_PS_QM083L = oForm01.DataSources.DBDataSources("@PS_QM083L")
    
    '// 메트릭스 개체 할당
'    Set oMat01 = oForm01.Items("Mat01").Specific

    oDS_PS_QM083H.setValue "U_DocDate", 0, Format(Now, "YYYYMMDD")
   
  '  oForm01.Items("DocEntry").Visible = True
  '  oForm01.Items("Code").Visible = True
    
    
    
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
    
'    '// 납품완료여부(Y/N)
    Set oCombo = oForm01.Items("FinishYN").Specific
    oCombo.ValidValues.Add "N", "납품미완료(N)"
    oCombo.ValidValues.Add "Y", "납품완료(Y)"

    oCombo.Select "N", psk_ByValue
    
    oForm01.Items("ExSize7").Specific.VALUE = "V.C"
    oForm01.Items("Weight7").Specific.VALUE = "저울"
    oForm01.Items("Length7").Specific.VALUE = "V.C"
    oForm01.Items("Exterio7").Specific.VALUE = "육안"
    oForm01.Items("Parall7").Specific.VALUE = "직각자"
    
    oForm01.ActiveItem = "DocDate"
    
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
    
    Dim FinishYN As String
    
    Dim sQry            As String
    Dim oRecordSet    As SAPbobsCOM.Recordset
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
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
                                MDC_Com.MDC_GF_Message "해당Lot의 입고중량보다 출고중량이 많습니다. 확인바랍니다.", "E"
                                BubbleEvent = False
                                Exit Sub
'                                Dim vReturnValue As Integer
'                                vReturnValue = Sbo_Application.MessageBox("해당Lot의 입고수량보다 출고수량이 많습니다. 계속하시겠습니까? ", 2, "&확인", "&취소")
'                                If vReturnValue = 2 Then
'                                    BubbleEvent = False
'                                    Exit Sub
'                                End If
                            End If
                        End If
                        
                        '//완료 여부를 Update
                        
'                        FinishYN = oForm01.Items("FinishYN").Specific.VALUE
'
'                        sQry = "Update [@PS_QM082H] Set U_FinishYN = '" & FinishYN & "' where Code = '" & oForm01.Items("BaseCode").Specific.VALUE & "'"
'                        oRecordSet.DoQuery sQry
                        
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
                ElseIf pval.ItemUID = "Btn03" Then
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                    Call PS_QM083_Print_Report01
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
                oDS_PS_QM083H.setValue "U_DocDate", 0, Format(Now, "YYYYMMDD")
                oForm01.Items("Qty").Click ct_Regular
                
                
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                FormItemEnabled
                FlushToItemValue "BaseCode"
'                If oMat01.VisualRowCount > 0 Then
'                    If oMat01.Columns("PP070No").Cells(oMat01.VisualRowCount).Specific.Value <> "" Then
'                        If oDS_PS_QM083H.GetValue("Status", 0) = "O" Then
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
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("LotNo").Enabled = True
        oForm01.Items("BaseCode").Enabled = True
        
    ElseIf oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("DocDate").Enabled = True
        oForm01.Items("LotNo").Enabled = True
        oForm01.Items("BaseCode").Enabled = True
        
    ElseIf oForm01.Mode = fm_OK_MODE Then
        oForm01.Items("DocDate").Enabled = False
        oForm01.Items("LotNo").Enabled = False
        oForm01.Items("BaseCode").Enabled = False
    End If
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM083'", "")
    If DocNum = 0 Then
'        oForm01.Items("Code").Specific.Value = 1
        oDS_PS_QM083H.setValue "Code", 0, 1
    Else
'        oForm01.Items("Code").Specific.Value = DocNum
        oDS_PS_QM083H.setValue "Code", 0, DocNum
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
'            sQry = "select cardname from ocrd where cardtype='C' and cardcode = '" & Trim(oDS_PS_QM083H.GetValue("U_BaseCode", 0)) & "'"
'            oRecordSet.DoQuery sQry
'            oDS_PS_QM083H.setValue "U_CardName", 0, Trim(oRecordSet.Fields(0).VALUE)
'
            '검사사양서내역
            sQry = "select DocDate = Convert(Char(8),a.U_DocDate,112), a.U_ItemCode, a.U_ItemName, a.U_CItemCod, a.U_CItemNam, a.U_StdNum, a.U_LotNo,"
            sQry = sQry + " b.U_ExSize, b.U_Weight, b.U_Length, U_Exterior, U_Parallel"
            sQry = sQry & " From [@PS_QM082H] a Inner Join [@PS_QM081H] b On a.U_BaseCode = b.Code  "
            sQry = sQry & " where a.Code = '" & Trim(oDS_PS_QM083H.GetValue("U_BaseCode", 0)) & "' "
            
            oRecordSet.DoQuery sQry
            oForm01.Items("InDate").Specific.VALUE = Trim(oRecordSet.Fields(0).VALUE)
            oForm01.Items("ItemCode").Specific.VALUE = Trim(oRecordSet.Fields(1).VALUE)
            oForm01.Items("ItemName").Specific.VALUE = Trim(oRecordSet.Fields(2).VALUE)
            oForm01.Items("CItemCod").Specific.VALUE = Trim(oRecordSet.Fields(3).VALUE)
            oForm01.Items("CItemNam").Specific.VALUE = Trim(oRecordSet.Fields(4).VALUE)
            oForm01.Items("StdNum").Specific.VALUE = Trim(oRecordSet.Fields(5).VALUE)
            oForm01.Items("LotNo").Specific.VALUE = Trim(oRecordSet.Fields(6).VALUE)
            oForm01.Items("M_ExSize").Specific.VALUE = Trim(oRecordSet.Fields(7).VALUE)
            oForm01.Items("M_Weight").Specific.VALUE = Trim(oRecordSet.Fields(8).VALUE)
            oForm01.Items("M_Length").Specific.VALUE = Trim(oRecordSet.Fields(9).VALUE)
            oForm01.Items("M_Exterior").Specific.VALUE = Trim(oRecordSet.Fields(10).VALUE)
            oForm01.Items("M_Parallel").Specific.VALUE = Trim(oRecordSet.Fields(11).VALUE)
       
            
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
        Case Trim(oDS_PS_QM083H.GetValue("U_DocDate", 0)) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_QM083H.GetValue("U_Qty", 0)) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_QM083H.GetValue("U_Weight", 0)) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_QM083H.GetValue("U_BaseCode", 0)) = ""
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
    End Select
    
    

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "입고일자는 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "수량은 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "중량은 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "원재료입고No 필수사항입니다. 확인하여 주십시오.", "E"
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
    Dim InWeight        As Double
    Dim OutWeight          As Double
    
    '//해당Lot의 입고수량과 출고수량을 비교 출고수량이 많으면 메세지 표시
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "select Sum(U_Weight) from [@PS_QM082H] "
    sQry = sQry + "where U_DocDate < '" & Trim(oDS_PS_QM083H.GetValue("U_DocDate", 0)) & "' and "
    sQry = sQry + " U_LotNo = '" & Trim(oDS_PS_QM083H.GetValue("U_LotNo", 0)) & "'"
    sQry = sQry + " And Code = '" & Trim(oDS_PS_QM083H.GetValue("U_BaseCode", 0)) & "'"
    oRecordSet.DoQuery sQry
    
    InWeight = oRecordSet.Fields(0).VALUE
    
    sQry = "select Sum(U_Weight) from [@PS_QM083H] where U_DocDate <= '" & Trim(oDS_PS_QM083H.GetValue("U_DocDate", 0)) & "' and "
    sQry = sQry + " U_LotNo = '" & Trim(oDS_PS_QM083H.GetValue("U_LotNo", 0)) & "'"
    sQry = sQry + " And U_BaseCode = '" & Trim(oDS_PS_QM083H.GetValue("U_BaseCode", 0)) & "'"
    oRecordSet.DoQuery sQry
    
    OutWeight = oRecordSet.Fields(0).VALUE + oDS_PS_QM083H.GetValue("U_Weight", 0)
    
    If InWeight < OutWeight Then
        Check_Exist = False
    Else
        Check_Exist = True
    End If

    Set oRecordSet = Nothing
    
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


Private Sub PS_QM083_Print_Report01()
On Error GoTo PS_QM083_Print_Report01_Error
    Dim WinTitle                    As String
    Dim ReportName                  As String
    Dim sQry                        As String
    Dim sQry01                      As String
    Dim PrtType                     As String
    Dim i                           As Long
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim Code$, BaseCode$, PtrType$

    Code = Trim(oForm01.Items("Code").Specific.VALUE)
    BaseCode = Trim(oForm01.Items("BaseCode").Specific.VALUE)
    
    sQry = "select a.U_PrtType From [@PS_QM081H] a Inner Join [@PS_QM082H] b On a.Code = b.U_BaseCode "
    sQry = sQry & " Where b.Code = '" & BaseCode & "'"
    
    oRecordSet.DoQuery sQry
    
    PrtType = Trim(oRecordSet.Fields(0).VALUE)
    
    If Trim(PrtType) = "" Then PrtType = "A"
    
    Call ConnectODBC
    
    Select Case PrtType
        Case "A"
            WinTitle = "[PS_QM083_01] 시험성적서A"
            ReportName = "PS_QM083_01.rpt"
        Case "B"
            WinTitle = "[PS_QM083_02] 시험성적서B"
            ReportName = "PS_QM083_02.rpt"
        Case "C"
            WinTitle = "[PS_QM083_03] 시험성적서C"
            ReportName = "PS_QM083_03.rpt"
    End Select
    
    sQry = "EXEC [PS_QM083_01] '" & Code & "'"
    
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") = False Then
        Sbo_Application.SetStatusBarMessage "gCryReport_Action : 실패!", bmt_Short, True
    End If
    Exit Sub
PS_QM083_Print_Report01_Error:
    Sbo_Application.SetStatusBarMessage "PS_QM083_Print_Report01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
