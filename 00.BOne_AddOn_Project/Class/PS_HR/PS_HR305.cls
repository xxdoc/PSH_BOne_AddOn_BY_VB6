VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_HR305"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_HR305.cls
'//  Module         : HR
'//  Description    : 사업부장KPI목표등록
'//  FormType       : PS_HR305
'//  Create Date    : 2012.2.13
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************


Option Explicit

Public oFormUniqueID01 As String
Public oForm01             As SAPbouiCOM.Form
Public oMat01              As SAPbouiCOM.Matrix
Public oMat02              As SAPbouiCOM.Matrix
Private oDS_PS_HR305H As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_HR305L As SAPbouiCOM.DBDataSource    '등록라인
Private oDS_PS_USERDS01 As SAPbouiCOM.DBDataSource '등록헤더

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private oSeq&
Private oCount&

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm(Optional ByVal prmPS_HR345 As PS_HR345)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_HR305.srf")
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
    
    oFormUniqueID01 = "PS_HR305_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    If (prmPS_HR345 Is Nothing) Then
        oForm01.Mode = fm_ADD_MODE
    Else
        oForm01.Mode = fm_FIND_MODE
    End If
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    Call FormClear
    Call Add_MatrixRow(0, True)
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제
        
    '문서번호 필드 Hidden
    oMat01.Columns("DocEntry").Visible = False

    oForm01.Update
    oForm01.Freeze False
    
    If (prmPS_HR345 Is Nothing) Then

    Else
        oForm01.Items("DocNum").Specific.VALUE = prmPS_HR345.oMat01.Columns("BasEntry").Cells(1).Specific.VALUE '문서번호
        oForm01.Items("StdYear").Specific.VALUE = Left(prmPS_HR345.oForm01.Items("EvalYM").Specific.VALUE, 4) '년도
        Call oForm01.Items("AttBPLId").Specific.Select(prmPS_HR345.oForm01.Items("AttBPLId").Specific.VALUE, psk_ByValue) '사업부
        
        oForm01.Items("1").Click ct_Regular
    End If
    
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
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Dim SumWeight As Currency, SumQty As Long
    Dim ItemCode$, ItemName$, Size$, Qty&, Weight As Currency, Unit$, RequestDate$, DueDate$, ItemType$, RequestNo$, BPLId$
    Dim RFC_Sender$
    Dim Calculate_Weight As Double
    Dim Seq&
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
                    
                        '같은 기준년도, 같은 사업부의 내용이 이미 입력되었는지 체크(fm_ADD_MODE일경우만)
                        If oForm01.Mode = fm_ADD_MODE Then
                            If CheckDocument = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        If oForm01.Mode = fm_ADD_MODE Then
                            oMat01.FlushToDataSource

                            oMat01.LoadFromDataSource
                        End If

                        Call Delete_EmptyRow
                        oLast_Mode = oForm01.Mode
                    ElseIf oForm01.Mode = fm_FIND_MODE Then
                        Call PS_HR305_MTX01 '//조회
                        oLast_Mode = oForm01.Mode
                    End If
                End If
                If pval.ItemUID = "Button01" Then
                    Call PS_HR305_MTX01 '//조회
                    
                End If
       
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "EmpNo" Then
                        If oForm01.Items("EmpNo").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                End If
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "AttBPLId" Then
                    oForm01.Freeze True
                    oMat01.Clear
                    oDS_PS_HR305L.Clear
                    If oForm01.Mode = fm_ADD_MODE Then
                        Call Add_MatrixRow(0, False)
                    End If
                                        
                    oForm01.Freeze False
                End If
            Case et_CLICK: '//6
            
                If pval.ItemUID = "Mat01" Then
                    If pval.Row > 0 Then
                        oLast_Item_UID = pval.ItemUID
                        oLast_Col_UID = pval.ColUID
                        oLast_Col_Row = pval.Row
                        
                        Call oMat01.SelectRow(pval.Row, True, False)
                    End If
                ElseIf pval.ItemUID = "Mat02" Then
                    If pval.Row > 0 Then
                        oLast_Item_UID = pval.ItemUID
                        oLast_Col_UID = pval.ColUID
                        oLast_Col_Row = pval.Row
                        
                        Call oMat02.SelectRow(pval.Row, True, False)
                    End If
                End If
            
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then

                    If (pval.ItemUID = "Mat01") Then
                        If pval.ColUID = "FETask" Then
                            oMat01.FlushToDataSource
                            
                            Call oDS_PS_HR305L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                            If oMat01.RowCount = pval.Row And Trim(oDS_PS_HR305L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                                Add_MatrixRow (pval.Row)
                            End If
                            oMat01.LoadFromDataSource
                            
                        ElseIf pval.ColUID = "Sig" Then '가중치 수정 시
                        
                            Call FlushToItemValue(pval.ItemUID, pval.Row, pval.ColUID) '가중치 계산
                            
                        End If
                        
                        oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                        
                    Else
                        
                    End If
                    
                End If
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
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_OK_MODE Then
                        If oLast_Mode = fm_UPDATE_MODE Then
                            Add_MatrixRow oMat01.RowCount, False
                            oLast_Mode = 100
                        ElseIf oLast_Mode = fm_FIND_MODE Then
                            Add_MatrixRow oMat01.RowCount, False
                            FormItemEnabled
                            oLast_Mode = 100
                        End If
                    ElseIf oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                    End If
                End If
            Case et_KEY_DOWN: '//2
'et_COMBO_SELECT ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
                
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    '//헤더
                    If pval.ItemUID = "EmpNo" Then
                        FlushToItemValue pval.ItemUID
'                    ElseIf pval.ItemUID = "Mat01" Then
                        
                    End If
                    
                End If
            Case et_MATRIX_LOAD: '//11
'                Add_MatrixRow oMat01.RowCount, False
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            
                Call Form_Resize
            
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
'et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
                Set oDS_PS_HR305H = Nothing
                Set oDS_PS_HR305L = Nothing
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    Set ProgressBar01 = Nothing
    oForm01.Freeze False
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
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
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
'[1284:취소] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1284": '취소
                FormItemEnabled
                oForm01.Items("DocNum").Click ct_Regular
            Case "1286": '닫기
'[1293:행삭제] //////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1293": '행삭제
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("LineNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i
                    
                    oMat01.FlushToDataSource
                    oDS_PS_HR305L.RemoveRecord oDS_PS_HR305L.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                End If
'[1281:찾기] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1281": '찾기
                oForm01.Freeze True
                Call FormItemEnabled
                oForm01.Items("DocNum").Click ct_Regular
                
                '//아이디별 사업장 세팅
                oForm01.Items("AttBPLId").Specific.Select MDC_PS_Common.User_BPLId, psk_ByValue
                
                oForm01.Freeze False
'[1282:추가] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1282": '추가
                Call Initialization
                Call FormItemEnabled
                Call FormClear
                Call oDS_PS_HR305H.setValue("U_StdYear", 0, Format(Now, "YYYY"))
                Call oDS_PS_HR305H.setValue("U_DocDate", 0, Format(Now, "YYYYMMDD"))
                
                Call Add_MatrixRow(0, True)
                
                '//마감여부콤보
                Call oDS_PS_HR305H.setValue("U_FinYN", 0, "N")
                
'[1288~1291:네비게이션] /////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call FormItemEnabled
                If oMat01.VisualRowCount > 0 Then
                    If oMat01.Columns("FETask").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
                        If oDS_PS_HR305H.GetValue("Status", 0) = "O" Then
                            Add_MatrixRow oMat01.RowCount, False
                        End If
                    End If
                    
                    Call PS_HR305_MTX01
                    
                End If
        End Select
    End If
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    Set oRecordSet01 = Nothing
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
    Set oDS_PS_HR305H = oForm01.DataSources.DBDataSources("@PS_HR305H")
    Set oDS_PS_HR305L = oForm01.DataSources.DBDataSources("@PS_HR305L")
    Set oDS_PS_USERDS01 = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    
    Set oMat02 = oForm01.Items("Mat02").Specific
    oMat02.SelectionMode = ms_NotSupported
    
    Call oDS_PS_HR305H.setValue("U_StdYear", 0, Format(Now, "YYYY"))
    Call oDS_PS_HR305H.setValue("U_DocDate", 0, Format(Now, "YYYYMMDD"))
    
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
    Set oCombo = oForm01.Items("AttBPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL] order by 1"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    '마감여부
    Set oCombo = oForm01.Items("FinYN").Specific
    oCombo.ValidValues.Add "N", "진행"
    oCombo.ValidValues.Add "Y", "마감"
    oCombo.Select "N", psk_ByValue
    
    'KPI단위_S
    sQry = "            SELECT      T1.U_Code,"
    sQry = sQry & "                 T1.U_CodeNm"
    sQry = sQry & "  FROM       [@PS_HR200H] AS T0"
    sQry = sQry & "                 INNER JOIN"
    sQry = sQry & "                 [@PS_HR200L] AS T1"
    sQry = sQry & "                     ON T0.Code = T1.Code"
    sQry = sQry & "  WHERE      Name = 'KPI단위'"

    Call oMat01.Columns("KPIUnit").ValidValues.Add("", "")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("KPIUnit"), sQry)
    'KPI단위_E
    
    '시작_S
    sQry = "            SELECT      T1.U_Code,"
    sQry = sQry & "                 T1.U_CodeNm"
    sQry = sQry & "  FROM       [@PS_HR200H] AS T0"
    sQry = sQry & "                 INNER JOIN"
    sQry = sQry & "                 [@PS_HR200L] AS T1"
    sQry = sQry & "                     ON T0.Code = T1.Code"
    sQry = sQry & "  WHERE      Name = '월'"
    
    Call oMat01.Columns("Stt").ValidValues.Add("", "")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Stt"), sQry)
    '시작_E
    
    '종료_S
    sQry = "            SELECT      T1.U_Code,"
    sQry = sQry & "                 T1.U_CodeNm"
    sQry = sQry & "  FROM       [@PS_HR200H] AS T0"
    sQry = sQry & "                 INNER JOIN"
    sQry = sQry & "                 [@PS_HR200L] AS T1"
    sQry = sQry & "                     ON T0.Code = T1.Code"
    sQry = sQry & "  WHERE      Name = '월'"

    Call oMat01.Columns("End").ValidValues.Add("", "")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("End"), sQry)
    '종료_E
    
    '주기_S
    sQry = "            SELECT      T1.U_Code,"
    sQry = sQry & "                 T1.U_CodeNm"
    sQry = sQry & "  FROM       [@PS_HR200H] AS T0"
    sQry = sQry & "                 INNER JOIN"
    sQry = sQry & "                 [@PS_HR200L] AS T1"
    sQry = sQry & "                     ON T0.Code = T1.Code"
    sQry = sQry & "  WHERE      Name = '평가주기'"

    Call oMat01.Columns("Pred").ValidValues.Add("", "")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Pred"), sQry)
    '주기_E
    
    '등급기준_S
    sQry = "            SELECT      T1.U_Code,"
    sQry = sQry & "                 T1.U_CodeNm"
    sQry = sQry & "  FROM       [@PS_HR200H] AS T0"
    sQry = sQry & "                 INNER JOIN"
    sQry = sQry & "                 [@PS_HR200L] AS T1"
    sQry = sQry & "                     ON T0.Code = T1.Code"
    sQry = sQry & "  WHERE      Name = '등급기준'"

    Call oMat01.Columns("GrdStd").ValidValues.Add("", "")
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("GrdStd"), sQry)
    '등급기준_E
    
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
    Set oCombo = oForm01.Items("AttBPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    Set oCombo = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Initialization_Error:
    Set oCombo = Nothing
    MDC_Com.MDC_GF_Message "Initialization_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub CF_ChooseFromList()
On Error GoTo CF_ChooseFromList_Error
    '//ChooseFromList 설정
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CF_ChooseFromList_Error:
    MDC_Com.MDC_GF_Message "CF_ChooseFromList_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error

    Dim FinYN As String
    
    If oMat01.VisualRowCount > 1 Then
    
        FinYN = PS_HR305_CheckDocumentStatus(oMat01.Columns("DocEntry").Cells(1).Specific.VALUE)
        
    End If

    If (oForm01.Mode = fm_ADD_MODE) Then
    
        If FinYN = "Y" Then
    
            oForm01.Items("AttBPLId").Enabled = False
            oForm01.Items("StdYear").Enabled = False
            oForm01.Items("EmpNo").Enabled = False
            oForm01.Items("DocDate").Enabled = False
            oForm01.Items("Mat01").Enabled = False
            oForm01.Items("Button01").Enabled = False
        
        Else
        
            oForm01.Items("AttBPLId").Enabled = True
            oForm01.Items("StdYear").Enabled = True
            oForm01.Items("EmpNo").Enabled = True
            oForm01.Items("DocDate").Enabled = True
            oForm01.Items("Mat01").Enabled = True
            oForm01.Items("Button01").Enabled = True
            
        End If
        
        oForm01.Items("DocNum").Enabled = False
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
    
        If FinYN = "Y" Then
            
            oForm01.Items("AttBPLId").Enabled = False
            oForm01.Items("StdYear").Enabled = False
            oForm01.Items("EmpNo").Enabled = False
            oForm01.Items("DocDate").Enabled = False
            oForm01.Items("Mat01").Enabled = False
            oForm01.Items("Button01").Enabled = False
        
        Else
        
            oForm01.Items("AttBPLId").Enabled = True
            oForm01.Items("StdYear").Enabled = True
            oForm01.Items("EmpNo").Enabled = True
            oForm01.Items("DocDate").Enabled = True
            oForm01.Items("Mat01").Enabled = True
            oForm01.Items("Button01").Enabled = True
        
        End If
        
        oForm01.Items("DocNum").Enabled = True
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
    
        If FinYN = "Y" Then
    
            oForm01.Items("AttBPLId").Enabled = False
            oForm01.Items("StdYear").Enabled = False
            oForm01.Items("EmpNo").Enabled = False
            oForm01.Items("DocDate").Enabled = False
            oForm01.Items("Mat01").Enabled = False
            oForm01.Items("Button01").Enabled = False
        
        Else
        
            oForm01.Items("AttBPLId").Enabled = True
            oForm01.Items("StdYear").Enabled = True
            oForm01.Items("EmpNo").Enabled = True
            oForm01.Items("DocDate").Enabled = True
            oForm01.Items("Mat01").Enabled = True
            oForm01.Items("Button01").Enabled = True
        
        End If
        
        oForm01.Items("DocNum").Enabled = False
        
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_HR305'", "")
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

Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_HR305L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_HR305L.Offset = oRow
    oDS_PS_HR305L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource

    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error
    Dim i&
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim SumWeight As Currency, SumQty As Long
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "EmpNo"
            sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" & Trim(oDS_PS_HR305H.GetValue("U_EmpNo", 0)) & "'"
            oRecordSet01.DoQuery sQry
            oDS_PS_HR305H.setValue "U_EmpNm", 0, Trim(oRecordSet01.Fields(0).VALUE)
        Case "Mat01"
            If oCol = "Sig" Then
                
                Dim SigSum As Double '가중치합
                Dim SigOrg As Double '가중치(수정전)
                SigSum = 0
                SigOrg = oDS_PS_HR305L.GetValue("U_Sig", oRow - 1) 'Flush하기 전에 KPI가중치를 저장(수정전의 값을 저장)

                Call oMat01.FlushToDataSource

                For i = 0 To oMat01.VisualRowCount - 1
                    SigSum = SigSum + oDS_PS_HR305L.GetValue("U_Sig", i)
                Next i

                If SigSum > 100 Then
                
                    If Sbo_Application.MessageBox("가중치 합이 100을 초과하였습니다. 계속 진행하시겠습니까?", "1", "예", "아니오") = "1" Then
                        oForm01.Items("SigSum").Specific.VALUE = SigSum
                    Else '아니오 선택 시
                        Call oDS_PS_HR305L.setValue("U_Sig", oRow - 1, SigOrg) '해당행의 KPI 가중치를 수정 전의 값으로 환원
                    End If
                    
                Else
                
                    oForm01.Items("SigSum").Specific.VALUE = SigSum
                
                End If
                
                Call oMat01.LoadFromDataSource
                
            End If
    End Select
'        Case "ItmBsort"
'            sQry = "Select Name From [@PSH_ITMBSORT] Where Code = '" & Trim(oDS_PS_HR305H.GetValue("U_ItmBsort", 0)) & "'"
'            oRecordSet01.DoQuery sQry
'
'            oDS_PS_HR305H.setValue "U_ItmBname", 0, Trim(oRecordSet01.Fields(0).VALUE)
'        Case "Mat01"
'            Select Case oCol
'                Case "MachCode"
'                    oForm01.Freeze True
'                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("MachCode").Cells(oRow).Specific.VALUE) <> "" Then
'                        oMat01.FlushToDataSource
'                        Add_MatrixRow oMat01.RowCount, False
'                        oMat01.Columns("MachCode").Cells(oRow).Click ct_Regular
'                    End If
'
'                    sQry = "Select a.U_MachCode, a.U_MachName "
'                    sQry = sQry & "From [@PS_PP130H] a "
'                    sQry = sQry & "Where a.U_MachCode = '" & Trim(oMat01.Columns("MachCode").Cells(oRow).Specific.VALUE) & "'"
'                    oRecordSet01.DoQuery sQry
'
'                    If oRecordSet01.RecordCount = 0 Then
'                        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
'                        Set oRecordSet01 = Nothing
'                        oForm01.Freeze False
'                        Exit Sub
'                    End If
'
'                    oMat01.Columns("MachName").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields("U_MachName").VALUE)
'                    oMat01.Columns("MachCode").Cells(oRow).Click ct_Regular
'
'                    oForm01.Freeze False
'            End Select
'    End Select

    Set oRecordSet01 = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
FlushToItemValue_Error:
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    
    ErrNum = 0

    '// Check
    Select Case True
        Case oDS_PS_HR305H.GetValue("U_AttBPLId", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_HR305H.GetValue("U_StdYear", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
    End Select
    

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "기준년도 필수사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error
    Dim i               As Long
    Dim ErrNum          As Integer
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Call oForm01.Freeze(True)

    ErrNum = 0
    
    oMat01.FlushToDataSource

    '// 라인
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    ElseIf oMat01.VisualRowCount = 1 And oDS_PS_HR305L.GetValue("U_FETask", 0) = "" Then
        ErrNum = 2
        GoTo MatrixSpaceLineDel_Error
    End If
    
    
    For i = 0 To oMat01.VisualRowCount - 2
        Select Case True
            Case oDS_PS_HR305L.GetValue("U_FETask", i) = ""
                ErrNum = 3
                GoTo MatrixSpaceLineDel_Error
        End Select
    Next
    oMat01.LoadFromDataSource

    Set oRecordSet01 = Nothing
    MatrixSpaceLineDel = True
    Call oForm01.Freeze(False)
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    If ErrNum = 1 Or ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "중점실행과제는 필수사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Sub Delete_EmptyRow()
On Error GoTo Delete_EmptyRow_Error
    Dim i&
    
    oMat01.FlushToDataSource
    
    For i = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_HR305L.GetValue("U_FETask", i)) = "" Then
            oDS_PS_HR305L.RemoveRecord i   '// Mat01에 마지막라인(빈라인) 삭제
        End If
    Next i
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Delete_EmptyRow_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub PS_HR305_MTX01()
On Error GoTo PS_HR305_MTX01_Error
    '//메트릭스에 데이터 로드
    Call oForm01.Freeze(True)
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim Param01 As String
    Dim Param02 As String
    
    Param01 = oForm01.Items("StdYear").Specific.VALUE
    Param02 = oForm01.Items("AttBPLId").Specific.Selected.VALUE
    
    Query01 = "Select U_MngPlc From [@PS_HR300H] a Inner Join [@PS_HR300L] b On a.DocEntry = b.DocEntry And a.Canceled = 'N' And a.U_StdYear = '" + Param01 + "' And a.U_BPLId = '" + Param02 + "'"
    
    Call RecordSet01.DoQuery(Query01)
    
    oMat02.Clear
    oMat02.FlushToDataSource
    oMat02.LoadFromDataSource
    
    If (RecordSet01.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PS_HR305_MTX01_Exit
    End If
    
    For i = 0 To RecordSet01.RecordCount - 1
        If i <> 0 Then
            oDS_PS_USERDS01.InsertRecord (i)
        End If
        oDS_PS_USERDS01.Offset = i
        Call oDS_PS_USERDS01.setValue("U_LineNum", i, i + 1)
        Call oDS_PS_USERDS01.setValue("U_ColRgL01", i, RecordSet01.Fields(0).VALUE)
        RecordSet01.MoveNext
    Next
    oMat02.LoadFromDataSource
'    oMat02.AutoResizeColumns
    oForm01.Update
   
    '조회 시 동적으로 상위 경영방침의 행 수를 "상위KPI행"에 바인딩
    '상위KPI행_S
    Dim sQry As String
    Dim sCount As Integer
    Dim sSeq As Integer
    
    
    sQry = "            SELECT      '' AS [Code],"
    sQry = sQry & "                 '' AS [Name]"
    sQry = sQry & "  UNION ALL"
    sQry = sQry & "  SELECT      CAST(T1.U_LineNum AS VARCHAR(10)) AS [Code],"
    sQry = sQry & "                 CAST(T1.U_LineNum AS VARCHAR(10)) AS [Name]"
    sQry = sQry & "  FROM       [@PS_HR300H] AS T0"
    sQry = sQry & "                 INNER JOIN"
    sQry = sQry & "                 [@PS_HR300L] AS T1"
    sQry = sQry & "                     ON T0.DocEntry = T1.DocEntry"
    sQry = sQry & "  Where       T0.U_BPLId = '" & oForm01.Items("AttBPLId").Specific.VALUE & "'"
    sQry = sQry & "                 AND T0.U_StdYear = '" & oForm01.Items("StdYear").Specific.VALUE & "'"
    sQry = sQry & "                 AND T0.[Status] = 'O'"
    
    sCount = oMat01.Columns("HKLine").ValidValues.Count
    sSeq = sCount
    For i = 1 To sCount
        Call oMat01.Columns("HKLine").ValidValues.Remove(sSeq - 1, psk_Index) '동적 바인딩을 위한 기존 값 제거
        sSeq = sSeq - 1
    Next i
    
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("HKLine"), sQry)
    '상위KPI행_E
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Sub
PS_HR305_MTX01_Exit:
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    
    Exit Sub
PS_HR305_MTX01_Error:
    
    
    Set RecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PS_HR305_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Function CheckDocument() As Boolean
'******************************************************************************
'Function ID : CheckDocument()
'해당모듈 : PS_HR305
'기능 : 해당 기준년도에 입력된 KPI목표가 있는지 검사
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo CheckDocument_Error
    
    Dim i As Long
    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Query01 = "                 SELECT       COUNT(*)"
    Query01 = Query01 & "  FROM        [@PS_HR305H] AS T0"
    Query01 = Query01 & "                  INNER JOIN"
    Query01 = Query01 & "                  [@PS_HR305H] AS T1"
    Query01 = Query01 & "                      ON T0.DocEntry = T1.DocEntry"
    Query01 = Query01 & "  WHERE       T0.[Status] = 'O'"
    Query01 = Query01 & "                  AND T0.U_AttBPLId = '" & oForm01.Items("AttBPLId").Specific.VALUE & "'"
    Query01 = Query01 & "                  AND T0.U_StdYear = '" & oForm01.Items("StdYear").Specific.VALUE & "'"
    
    Call RecordSet01.DoQuery(Query01)
    
    If RecordSet01.Fields(0).VALUE = 0 Then
        CheckDocument = True
    Else
        CheckDocument = False
        Call MDC_Com.MDC_GF_Message("해당 기준년도에 이미 [사업부KPI목표]가 등록되었습니다. 확인하십시오.", "E")
    End If
    
    Set RecordSet01 = Nothing
    Exit Function
CheckDocument_Error:
    Set RecordSet01 = Nothing
    Call Sbo_Application.SetStatusBarMessage("CheckDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Function

Public Sub Form_Resize()
'******************************************************************************
'Function ID : Form_Resize()
'해당모듈 : PS_HR305
'기능 : 폼 크기 조정시 각 아이템들의 위치 및 크기 조정
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo Form_Resize_Error

    '조회버튼
    oForm01.Items("Button01").Left = 470
    oForm01.Items("Button01").Top = 10
    '조회버튼
    
    '상위부서KPI목표 매트릭스
    oForm01.Items("Mat02").Height = oForm01.Height / 2 - 140 '화면 전체 높이의 1/2에서 140만큼 뺌
    oForm01.Items("Mat02").Width = oForm01.Width / 2
    '상위부서KPI목표 매트릭스

    '사업부KPI목표 Label
    oForm01.Items("22").Top = oForm01.Items("Mat02").Top + oForm01.Items("Mat02").Height + 5
    '사업부KPI목표 Label

    '사업부KPI목표 매트릭스
    oForm01.Items("Mat01").Top = oForm01.Items("Mat02").Top + oForm01.Items("Mat02").Height + oForm01.Items("22").Height + 5
    oForm01.Items("Mat01").Height = oForm01.Height / 2 - 58
    oForm01.Items("Mat01").Width = oForm01.Width - 35
    '사업부KPI목표 매트릭스

    oMat02.Columns("LineNum").Width = 20
    oMat02.Columns("MngPlc").Width = oForm01.Items("Mat02").Width - oMat02.Columns("LineNum").Width
    Call oMat01.AutoResizeColumns
'    Call oMat02.AutoResizeColumns
    
    Exit Sub
    
Form_Resize_Error:
    Call Sbo_Application.SetStatusBarMessage("Form_Resize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Sub

Public Function PS_HR305_CheckDocumentStatus(ByVal prmDocEntry As Integer) As String
'******************************************************************************
'Function ID : PS_HR305_CheckDocumentStatus()
'해당모듈 : PS_HR305
'기능 : 마감여부에 따라 화면 전체의 기능 제어
'인수 : prmDocEntry(문서번호)
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PS_HR305_CheckDocumentStatus_Error

    Dim Query01 As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Query01 = "SELECT U_FinYN AS [FinYN] FROM [@PS_HR305H] WHERE DocEntry = " & prmDocEntry

    Call RecordSet01.DoQuery(Query01)
    
    PS_HR305_CheckDocumentStatus = RecordSet01.Fields("FinYN").VALUE
 
Exit Function

PS_HR305_CheckDocumentStatus_Error:
    Call Sbo_Application.SetStatusBarMessage("PS_HR305_CheckDocumentStatus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True)
End Function

