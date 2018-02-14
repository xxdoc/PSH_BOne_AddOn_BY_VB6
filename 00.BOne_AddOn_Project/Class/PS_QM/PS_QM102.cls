VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM102"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_QM102  .cls
'//  Module         : QM
'//  Description    : 기계공구류수입검사일지 등록
'//  FormType       : PS_QM102
'//  Create Date    : 2012.01.27
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_QM102H       As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_QM102L       As SAPbouiCOM.DBDataSource    '등록라인

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM102.srf")
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
    
    oFormUniqueID01 = "PS_QM102_" & GetTotalFormsCount
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
    Call Matrix_AddRow(1, 0, True) 'oMat01
    FormItemEnabled
    
    oForm01.EnableMenu ("1283"), True        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1284"), False        '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제
    oForm01.Items("DocDate").Click ct_Regular
    'oform01.Items("DocDate")..Click ct_Regular
    
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
    Set oDS_PS_QM102H = oForm01.DataSources.DBDataSources("@PS_QM102H")
    Set oDS_PS_QM102L = oForm01.DataSources.DBDataSources("@PS_QM102L")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    oDS_PS_QM102H.setValue "U_DocDate", 0, Format(Now, "YYYYMMDD")

    Call oForm01.Items("Opt01").Specific.GroupWith("Opt02")
    Call oForm01.Items("Opt01").Click
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CreateItems_Error:
    MDC_Com.MDC_GF_Message "CreateItems_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    '//콤보에 기본값설정
    Dim oCombo              As SAPbouiCOM.ComboBox
    Dim oCombo_Matrix       As SAPbouiCOM.Column
    Dim sQry                As String
    Dim oRecordSet01        As SAPbobsCOM.Recordset
        
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
    
      
    '매트릭스의 검사완료구분
    oMat01.Columns("FinYN").ValidValues.Add "N", "검사중"
    oMat01.Columns("FinYN").ValidValues.Add "Y", "검사완료"
    
        
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
    
    Dim ChildForm01 As Variant
    Set ChildForm01 = New PS_SM010

    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If

                    End If
                ElseIf pval.ItemUID = "Opt01" Then
                
                    Call oForm01.Freeze(True)
                    oMat01.Columns("OrdNum").Editable = False
                    oMat01.Columns("JakMyung").Editable = False
                    oMat01.Columns("JakSize").Editable = False
                    oMat01.Columns("JakUnit").Editable = False
                    oMat01.Columns("Qty").Editable = False
                    Call oMat01.AutoResizeColumns
                    Call oForm01.Freeze(False)
                    
                ElseIf pval.ItemUID = "Opt02" Then
                
                    Call oForm01.Freeze(True)
                    oMat01.Columns("OrdNum").Editable = True
                    oMat01.Columns("JakMyung").Editable = True
                    oMat01.Columns("JakSize").Editable = True
                    oMat01.Columns("JakUnit").Editable = True
                    oMat01.Columns("Qty").Editable = True
                    Call oMat01.AutoResizeColumns
                    Call oForm01.Freeze(False)
                    
                End If
                
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "CntcCode" Then
                        If oForm01.Items("CntcCode").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                        End If
                    End If
                    '//라인
                    If pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "GaDocLin" Then
                            If oMat01.Columns("GaDocLin").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        End If
                        If pval.ColUID = "FCode1" Then
                            If oMat01.Columns("FCode1").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        End If
                        If pval.ColUID = "FCode2" Then
                            If oMat01.Columns("FCode2").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
                        End If
                        If pval.ColUID = "FCode3" Then
                            If oMat01.Columns("FCode3").Cells(pval.Row).Specific.VALUE = "" Then
                                Sbo_Application.ActivateMenuItem ("7425")
                                BubbleEvent = False
                            End If
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
'                        Call Matrix_AddRow(1, oMat01.RowCount, False) 'oMat01
                    End If
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "Mat01" Then  'And (pval.ColUID = "GaDocLin" Or pval.ColUID = "FCode1" Or pval.ColUID = "FCode2" Or pval.ColUID = "FCode3" Or pval.ColUID = "OrdNum") Then
                    Call FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                End If
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
'                    //헤더
                    If pval.ItemUID = "CntcCode" Then
                        Call FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                    End If
                    
                    '//라인
                    If pval.ItemUID = "Mat01" Then  'And (pval.ColUID = "GaDocLin" Or pval.ColUID = "FCode1" Or pval.ColUID = "FCode2" Or pval.ColUID = "FCode3" Or pval.ColUID = "OrdNum") Then
                        Call FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID)
                    End If
                End If
            Case et_MATRIX_LOAD: '//11
                
                Call Matrix_AddRow(1, oMat01.VisualRowCount) 'oMat01
                Call oMat01.AutoResizeColumns
                
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
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                    '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
                    '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("LineNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i
    
                    oMat01.FlushToDataSource
                    oDS_PS_QM102L.RemoveRecord oDS_PS_QM102L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                End If
            
            Case "1281": '찾기
                FormItemEnabled
                oForm01.Items("DocEntry").Click ct_Regular
                
            Case "1282": '추가
                Call FormItemEnabled
                Call FormClear
                oDS_PS_QM102H.setValue "U_DocDate", 0, Format(Now, "YYYYMMDD")
                Call Matrix_AddRow(1, 0, True) 'oMat01
                
                '//-- Combo Box 초기화
                Dim oCombo          As SAPbouiCOM.ComboBox
                '// 사업장
                Set oCombo = oForm01.Items("BPLId").Specific
                oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
                
                Set oCombo = Nothing

'                oForm01.Items("DocDate").Click ct_Regular

            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                FormItemEnabled
                If oMat01.VisualRowCount > 0 Then
                    If oMat01.Columns("OrdNum").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
                        Matrix_AddRow 1, oMat01.RowCount, False
                    End If
                End If
  
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
        oForm01.Items("DocEntry").Enabled = True
        
        oForm01.Items("Opt01").Enabled = False
        oForm01.Items("Opt02").Enabled = False
        
    ElseIf oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("DocEntry").Enabled = False
        
        oForm01.Items("Opt01").Enabled = True
        oForm01.Items("Opt02").Enabled = True
        
        Call oForm01.Items("Opt01").Click
        
    ElseIf oForm01.Mode = fm_OK_MODE Then
        oForm01.Items("DocEntry").Enabled = False
        
        oForm01.Items("Opt01").Enabled = False
        oForm01.Items("Opt02").Enabled = False

        If Trim(oDS_PS_QM102H.GetValue("U_Opt02", 0)) = "1" Then
            
            oMat01.Columns("OrdNum").Editable = True
            oMat01.Columns("JakMyung").Editable = True
            oMat01.Columns("JakSize").Editable = True
            oMat01.Columns("Qty").Editable = True
            oMat01.Columns("JakUnit").Editable = True
            
        ElseIf Trim(oDS_PS_QM102H.GetValue("U_Opt02", 0)) = "2" Then
            
            oMat01.Columns("OrdNum").Editable = False
            oMat01.Columns("JakMyung").Editable = False
            oMat01.Columns("JakSize").Editable = False
            oMat01.Columns("JakUnit").Editable = False
            oMat01.Columns("Qty").Editable = False
        
        Else
        
            oMat01.Columns("OrdNum").Editable = False
            oMat01.Columns("JakMyung").Editable = False
            oMat01.Columns("JakSize").Editable = False
            oMat01.Columns("JakUnit").Editable = False
            oMat01.Columns("Qty").Editable = False
            
        End If

    End If
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM102'", "")
    If DocEntry = 0 Then
        oForm01.Items("DocEntry").Specific.VALUE = 1
    Else
        oForm01.Items("DocEntry").Specific.VALUE = DocEntry
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
            If Insert_YN = False Then
                oRow = oMat01.RowCount
                oDS_PS_QM102L.InsertRecord (oRow)
            End If
            '수입내역
            oDS_PS_QM102L.Offset = oRow
            oDS_PS_QM102L.setValue "U_LineNum", oRow, oRow + 1
'            oDS_PS_QM102L.setValue "U_ItmBsort", oRow, ""
'            oDS_PS_QM102L.setValue "U_FAmt01", oRow, ""
'            oDS_PS_QM102L.setValue "U_FAmt02", oRow, ""
'            oDS_PS_QM102L.setValue "U_TotFAmt", oRow, ""
            oMat01.LoadFromDataSource

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

    Dim i&, j&, Cnt&, DocNum$, LineId$
    Dim ErrNum As Integer
    Dim sQry As String
    Dim GADocLin As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Call oForm01.Freeze(True)
    
    '--------------------------------------------------------------
    'Header--------------------------------------------------------
    Select Case oUID
        Case "CntcCode"
            '//거래처명 검색
            sQry = "select U_FULLNAME from OHEM where U_MSTCOD = '" & Trim(oDS_PS_QM102H.GetValue("U_CntcCode", 0)) & "'"
            oRecordSet.DoQuery sQry
            oDS_PS_QM102H.setValue "U_CntcName", 0, Trim(oRecordSet.Fields(0).VALUE)
    End Select
        
    '--------------------------------------------------------------
    'Line----------------------------------------------------------
    If oUID = "Mat01" Then
        Call oDS_PS_QM102L.setValue("U_" & oCol, oRow - 1, oMat01.Columns(oCol).Cells(oRow).Specific.VALUE)
        
        Select Case oCol
        
            Case "GaDocLin"
            
                If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("GaDocLin").Cells(oRow).Specific.VALUE) <> "" Then
                    Call Matrix_AddRow(1, oMat01.RowCount, False)
                    
                End If
                GADocLin = Trim(oMat01.Columns("GaDocLin").Cells(oRow).Specific.VALUE)
                sQry = "EXEC [PS_QM102_02] '" & Trim(GADocLin) & "'"
                Call oRecordSet.DoQuery(sQry)

                Call oDS_PS_QM102L.setValue("U_OrdNum", oRow - 1, Trim(oRecordSet.Fields("OrdNum").VALUE))
                Call oDS_PS_QM102L.setValue("U_JakMyung", oRow - 1, Trim(oRecordSet.Fields("JakMyung").VALUE))
                Call oDS_PS_QM102L.setValue("U_JakSize", oRow - 1, Trim(oRecordSet.Fields("JakSize").VALUE))
                Call oDS_PS_QM102L.setValue("U_JakUnit", oRow - 1, Trim(oRecordSet.Fields("JakUnit").VALUE))
                Call oDS_PS_QM102L.setValue("U_Qty", oRow - 1, Trim(oRecordSet.Fields("Qty").VALUE))
                Call oDS_PS_QM102L.setValue("U_PP030HNo", oRow - 1, Trim(oRecordSet.Fields("PP030HNo").VALUE))
                Call oDS_PS_QM102L.setValue("U_PP030MNo", oRow - 1, Trim(oRecordSet.Fields("PP030MNo").VALUE))
                Call oDS_PS_QM102L.setValue("U_CardCode", oRow - 1, Trim(oRecordSet.Fields("CardCode").VALUE))
                Call oDS_PS_QM102L.setValue("U_CardName", oRow - 1, Trim(oRecordSet.Fields("CardName").VALUE))
                Call oDS_PS_QM102L.setValue("U_CpName", oRow - 1, Trim(oRecordSet.Fields("CpName").VALUE))
                Call oDS_PS_QM102L.setValue("U_CheckQty", oRow - 1, Trim(oRecordSet.Fields("Qty").VALUE))
                Call oDS_PS_QM102L.setValue("U_JanQty", oRow - 1, Trim(oRecordSet.Fields("JanQty").VALUE))
                
                Call oDS_PS_QM102L.setValue("U_FinYN", oRow - 1, "N")
'                Call oMat01.Columns("GaDocLin").Cells(oRow).Click(ct_Regular)
              
            Case "OrdNum"
            
                If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("OrdNum").Cells(oRow).Specific.VALUE) <> "" Then
                    Call Matrix_AddRow(1, oMat01.RowCount, False)
                End If
                Call oDS_PS_QM102L.setValue("U_FinYN", oRow - 1, "N")
                
            Case "FCode1"
            
                sQry = "select U_SmalName from [@PS_PP003L] where U_SmalCode = '" & Trim(oMat01.Columns("FCode1").Cells(oRow).Specific.VALUE) & "'"
                Call oRecordSet.DoQuery(sQry)
                Call oDS_PS_QM102L.setValue("U_FName1", oRow - 1, Trim(oRecordSet.Fields(0).VALUE))
               
            Case "FCode2"
            
                sQry = "select U_SmalName from [@PS_PP003L] where U_SmalCode = '" & Trim(oMat01.Columns("FCode2").Cells(oRow).Specific.VALUE) & "'"
                Call oRecordSet.DoQuery(sQry)
                Call oDS_PS_QM102L.setValue("U_FName2", oRow - 1, Trim(oRecordSet.Fields(0).VALUE))
               
            Case "FCode3"
            
                sQry = "select U_SmalName from [@PS_PP003L] where U_SmalCode = '" & Trim(oMat01.Columns("FCode3").Cells(oRow).Specific.VALUE) & "'"
                Call oRecordSet.DoQuery(sQry)
                Call oDS_PS_QM102L.setValue("U_FName3", oRow - 1, Trim(oRecordSet.Fields(0).VALUE))
                
        End Select
        
'        If oCol <> "GaDocLin" Then
'            Call oMat01.FlushToDataSource
'        End If
        
        Call oMat01.LoadFromDataSource
        Call oMat01.AutoResizeColumns

        If oCol <> "FinYN" Then
            Call oMat01.Columns(oCol).Cells(oRow).Click(ct_Regular)
        Else
            Call oMat01.Columns("PassQty").Cells(oRow).Click(ct_Regular)
        End If
        
    End If
     
    Call oForm01.Update
    Call oForm01.Freeze(False)
    
    Set oRecordSet = Nothing
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    
    Call oForm01.Freeze(False)
    Set oRecordSet = Nothing
    
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    
    ErrNum = 0

    '// Check
    Select Case True
        Case Trim(oDS_PS_QM102H.GetValue("U_DocDate", 0)) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "등록일자는 필수사항입니다. 확인하여 주십시오.", "E"
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
    
    oMat01.FlushToDataSource

    '// 라인
    If oMat01.VisualRowCount = 1 Then '// MAT01에 값이 있는지 확인 (ErrorNumber : 1)
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//마지막 행 하나를 빼고 i=0부터 시작하므로 하나를 빼므로
    '//oMat01.RowCount - 2가 된다..반드시 들어 가야 하는 필수값을 확인한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    If oMat01.VisualRowCount > 0 Then
        For i = 0 To oMat01.VisualRowCount - 2      '// Mat1에 입력값이 올바르게 들어갔는지 확인 (ErrorNumber : 2)
            oDS_PS_QM102L.Offset = i
            If Trim(oDS_PS_QM102L.GetValue("U_OrdNum", i)) = "" Then
                ErrNum = 2
                oMat01.Columns("OrdNum").Cells(i + 1).Click ct_Regular
                GoTo MatrixSpaceLineDel_Error
            End If
        Next i
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
    '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    If oMat01.VisualRowCount > 0 Then
        oDS_PS_QM102L.RemoveRecord oDS_PS_QM102L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
    End If
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '행을 삭제하였으니 DB데이터 소스를 다시 가져온다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oMat01.LoadFromDataSource

    Set oRecordSet = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "구분(작지번호)는 필수사항입니다. 확인하여 주십시오.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function






