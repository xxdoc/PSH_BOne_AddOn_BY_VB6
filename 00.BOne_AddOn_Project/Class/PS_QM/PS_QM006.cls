VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_QM006.cls
'//  Module         : QM
'//  Description    : 분말 검사성적서등록
'//  FormType       : PS_QM006
'//  Create Date    : 2014.10.02
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_QM006H       As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_QM006L       As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private Last_CardCode$, Last_CardName$, Last_InspPrsn$
Private oCode$

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm(Optional sCode As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM006.srf")
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
    
    oFormUniqueID01 = "PS_QM006_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    
    If sCode <> "" Then
        oForm01.Mode = fm_FIND_MODE
    Else
        oForm01.Mode = fm_ADD_MODE
    End If
    
    
    
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "Code"
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
    
    oCode = sCode
    If oCode <> "" Then
    
        Call FormItemEnabled
    
        oForm01.Items("Code").Specific.VALUE = sCode
        Call oForm01.Items("1").Click
        
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

Private Sub CreateItems()
On Error GoTo CreateItems_Error
    '//디비데이터 소스 개체 할당
    Set oDS_PS_QM006H = oForm01.DataSources.DBDataSources("@PS_QM006H")
'    Set oDS_PS_QM006L = oForm01.DataSources.DBDataSources("@PS_QM006L")
    
    '// 메트릭스 개체 할당
'    Set oMat01 = oForm01.Items("Mat01").Specific

    oDS_PS_QM006H.setValue "U_InspDate", 0, Format(Now, "YYYYMMDD")
   
    oForm01.Items("Code").Visible = True 'False
   
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
    
    '// 검사자
    Set oCombo = oForm01.Items("InspPrsn").Specific
    sQry = "SELECT DISTINCT Code, U_FullName From [@PH_PY001A] "

    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        oCombo.ValidValues.Add Trim(oRecordSet.Fields(0).VALUE), Trim(oRecordSet.Fields(1).VALUE)
        oRecordSet.MoveNext
    Loop
    
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
    
    Dim sQry            As String
    Dim oRecordSet      As SAPbobsCOM.Recordset
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim InspNo As String
    
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
                            InspNo = oForm01.Items("InspNo").Specific.VALUE
                            '검사입력되면 검사완료 Sign Update
                            sQry = "Update [Z_PACKING_PD] Set QM006YN = 'Y' Where InspNo = '" & InspNo & "'"
                            oRecordSet.DoQuery sQry
                            
                        End If
                    End If
                End If
                
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    '//헤더
                    If pval.ItemUID = "InspNo" Then
                        If oForm01.Items("InspNo").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                    If pval.ItemUID = "CardCode" Then
                        If oForm01.Items("CardCode").Specific.VALUE = "" Then
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
                        FlushToItemValue "ItemCode"
'                        Call Matrix_AddRow(1, oMat01.RowCount, False) 'oMat01
                    End If
                End If
                
                If pval.ItemUID = "Btn_Prt" Then '검사성적서출력 버튼 클릭시
                    Call Print_Query
                End If
                
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    '//헤더
                    If pval.ItemUID = "InspNo" Or pval.ItemUID = "ItemCode" Or pval.ItemUID = "CardCode" Then
                        FlushToItemValue pval.ItemUID
                        FlushToItemValue "ItemCode"
                    End If
                    
                    If pval.ItemUID = "M_P60" Or pval.ItemUID = "M_P80" Or pval.ItemUID = "M_P100" Or pval.ItemUID = "M_P140" Or pval.ItemUID = "M_P200" Or pval.ItemUID = "M_P325" Or pval.ItemUID = "M_M325" Then
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
    Dim sQry            As String
    
    
    Dim oRecordSet      As SAPbobsCOM.Recordset
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    Dim InspNo As String
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1283": '삭제
                If oForm01.Mode = fm_OK_MODE Then
                    InspNo = oForm01.Items("InspNo").Specific.VALUE
                    '검사입력되면 검사완료 Sign Update
                    sQry = "Update [Z_PACKING_PD] Set QM006YN = 'N' Where InspNo = '" & InspNo & "'"
                    oRecordSet.DoQuery sQry
                End If
                            
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
                oForm01.Freeze True
                oDS_PS_QM006H.setValue "U_InspNo", 0, ""
                oDS_PS_QM006H.setValue "U_ItemCode", 0, ""
                oDS_PS_QM006H.setValue "U_ItemName", 0, ""
                oDS_PS_QM006H.setValue "U_InspDate", 0, Format(Now, "YYYYMMDD")
                oDS_PS_QM006H.setValue "U_InspPrsn", 0, ""
                oDS_PS_QM006H.setValue "U_CardCode", 0, ""
                oDS_PS_QM006H.setValue "U_CardName", 0, ""
                FormItemEnabled
                Call FormClear
                oForm01.Items("InspNo").Click ct_Regular
                oForm01.Freeze False
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
'                    oDS_PS_QM006L.RemoveRecord oDS_PS_QM006L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
'                    oMat01.Clear
'                    oMat01.LoadFromDataSource
'                End If
            
            Case "1281": '찾기
                FormItemEnabled
                oForm01.Items("InspNo").Click ct_Regular
                oForm01.Items("C_Sum").Specific.VALUE = 0
                
            Case "1282": '추가
                Call FormItemEnabled
                Call FormClear
                oDS_PS_QM006H.setValue "U_InspDate", 0, Format(Now, "YYYYMMDD")
                oForm01.Items("InspNo").Click ct_Regular
                
                oDS_PS_QM006H.setValue "U_CardCode", 0, Last_CardCode
                oDS_PS_QM006H.setValue "U_CardName", 0, Last_CardName
                oDS_PS_QM006H.setValue "U_InspPrsn", 0, Last_InspPrsn
                
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                FormItemEnabled
                FlushToItemValue "ItemCode"
                
                
'                If oMat01.VisualRowCount > 0 Then
'                    If oMat01.Columns("PP070No").Cells(oMat01.VisualRowCount).Specific.Value <> "" Then
'                        If oDS_PS_QM006H.GetValue("Status", 0) = "O" Then
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
    Dim M_P60 As Currency
    Dim M_P80 As Currency
    Dim M_P100 As Currency
    Dim M_P140 As Currency
    Dim M_P200 As Currency
    Dim M_P325 As Currency
    Dim M_M325 As Currency
    
    If oForm01.Mode = fm_FIND_MODE Then
        oForm01.Items("Code").Enabled = True
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("InspNo").Enabled = True
        oForm01.Items("InspDate").Specific.VALUE = ""
    ElseIf oForm01.Mode = fm_ADD_MODE Then
        oForm01.Items("Code").Enabled = False
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("InspNo").Enabled = True
        
    ElseIf oForm01.Mode = fm_OK_MODE Then
        oForm01.Items("CardCode").Enabled = True
        oForm01.Items("InspNo").Enabled = False
    End If
    
    M_P60 = Val(oForm01.Items("M_P60").Specific.VALUE)
    M_P80 = Val(oForm01.Items("M_P80").Specific.VALUE)
    M_P100 = Val(oForm01.Items("M_P100").Specific.VALUE)
    M_P140 = Val(oForm01.Items("M_P140").Specific.VALUE)
    M_P200 = Val(oForm01.Items("M_P200").Specific.VALUE)
    M_P325 = Val(oForm01.Items("M_P325").Specific.VALUE)
    M_M325 = Val(oForm01.Items("M_M325").Specific.VALUE)

    oForm01.Items("C_Sum").Specific.VALUE = M_P60 + M_P80 + M_P100 + M_P140 + M_P200 + M_P325 + M_M325
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM006'", "")
    If DocNum = 0 Then
'        oForm01.Items("Code").Specific.Value = 1
        oDS_PS_QM006H.setValue "Code", 0, 1
    Else
'        oForm01.Items("Code").Specific.Value = DocNum
        oDS_PS_QM006H.setValue "Code", 0, DocNum
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
'                oDS_PS_QM006L.InsertRecord (oRow)
'            End If
'            '수입내역
'            oDS_PS_QM006L.Offset = oRow
'            oDS_PS_QM006L.setValue "LineId", oRow, oRow + 1
'            oDS_PS_QM006L.setValue "U_PP070No", oRow, ""
'            oDS_PS_QM006L.setValue "U_ItemCode", oRow, ""
'            oDS_PS_QM006L.setValue "U_ItemName", oRow, ""
'            oDS_PS_QM006L.setValue "U_Size", oRow, ""
'            oDS_PS_QM006L.setValue "U_Mark", oRow, ""
'            oDS_PS_QM006L.setValue "U_Qty", oRow, ""
'            oDS_PS_QM006L.setValue "U_Weight", oRow, ""
'            oDS_PS_QM006L.setValue "U_DocDate", oRow, ""
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
    Dim oRecordSet    As SAPbobsCOM.Recordset
        
    Dim M_P60 As Currency
    Dim M_P80 As Currency
    Dim M_P100 As Currency
    Dim M_P140 As Currency
    Dim M_P200 As Currency
    Dim M_P325 As Currency
    Dim M_M325 As Currency
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    '--------------------------------------------------------------
    'Header--------------------------------------------------------
    Select Case oUID
        Case "CardCode", "ItemCode"
            '납품처명
            sQry = "select cardname from ocrd where cardtype='C' and cardcode = '" & Trim(oDS_PS_QM006H.GetValue("U_CardCode", 0)) & "'"
            oRecordSet.DoQuery sQry
            oDS_PS_QM006H.setValue "U_CardName", 0, Trim(oRecordSet.Fields(0).VALUE)
            
            '검사사양서내역
            sQry = "select U_Fe_max,U_Sn_max,U_Pb_max, U_Cu_min,U_Ad_min,U_Ad_max,U_Fr_max,U_Oxy_max,U_P60_min,U_P60_max, U_P60Bal, "
            sQry = sQry & "U_P80_min,U_P80_max, U_P80Bal, U_P100_min,U_P100_max, U_P100Bal, U_P140_min,U_P140_max, U_P140Bal, "
            sQry = sQry & "U_P200_min,U_P200_max, U_P200Bal, U_P325_min,U_P325_max, U_P325Bal, U_M325_min,U_M325_max, U_M325Bal "
            sQry = sQry & "From [@PS_QM005H] "
            sQry = sQry & "where U_CardCode = '" & Trim(oDS_PS_QM006H.GetValue("U_CardCode", 0)) & "' "
            sQry = sQry & "and   U_ItemCode = '" & Trim(oDS_PS_QM006H.GetValue("U_ItemCode", 0)) & "'"
            oRecordSet.DoQuery sQry
            
            
            oForm01.Items("Fe_max").Specific.VALUE = Trim(oRecordSet.Fields(0).VALUE)
            oForm01.Items("Sn_max").Specific.VALUE = Trim(oRecordSet.Fields(1).VALUE)
            oForm01.Items("Pb_max").Specific.VALUE = Trim(oRecordSet.Fields(2).VALUE)
            oForm01.Items("Cu_min").Specific.VALUE = Trim(oRecordSet.Fields(3).VALUE)
            oForm01.Items("Ad_min").Specific.VALUE = Trim(oRecordSet.Fields(4).VALUE)
            oForm01.Items("Ad_max").Specific.VALUE = Trim(oRecordSet.Fields(5).VALUE)
            oForm01.Items("Fr_max").Specific.VALUE = Trim(oRecordSet.Fields(6).VALUE)
            oForm01.Items("Oxy_max").Specific.VALUE = Trim(oRecordSet.Fields(7).VALUE)
            oForm01.Items("P60_min").Specific.VALUE = Trim(oRecordSet.Fields(8).VALUE)
            oForm01.Items("P60_max").Specific.VALUE = Trim(oRecordSet.Fields(9).VALUE)
            oForm01.Items("P60Bal").Specific.VALUE = Trim(oRecordSet.Fields(10).VALUE)
            oForm01.Items("P80_min").Specific.VALUE = Trim(oRecordSet.Fields(11).VALUE)
            oForm01.Items("P80_max").Specific.VALUE = Trim(oRecordSet.Fields(12).VALUE)
            oForm01.Items("P80Bal").Specific.VALUE = Trim(oRecordSet.Fields(13).VALUE)
            oForm01.Items("P100_min").Specific.VALUE = Trim(oRecordSet.Fields(14).VALUE)
            oForm01.Items("P100_max").Specific.VALUE = Trim(oRecordSet.Fields(15).VALUE)
            oForm01.Items("P100Bal").Specific.VALUE = Trim(oRecordSet.Fields(16).VALUE)
            oForm01.Items("P140_min").Specific.VALUE = Trim(oRecordSet.Fields(17).VALUE)
            oForm01.Items("P140_max").Specific.VALUE = Trim(oRecordSet.Fields(18).VALUE)
            oForm01.Items("P140Bal").Specific.VALUE = Trim(oRecordSet.Fields(19).VALUE)
            oForm01.Items("P200_min").Specific.VALUE = Trim(oRecordSet.Fields(20).VALUE)
            oForm01.Items("P200_max").Specific.VALUE = Trim(oRecordSet.Fields(21).VALUE)
            oForm01.Items("P200Bal").Specific.VALUE = Trim(oRecordSet.Fields(22).VALUE)
            oForm01.Items("P325_min").Specific.VALUE = Trim(oRecordSet.Fields(23).VALUE)
            oForm01.Items("P325_max").Specific.VALUE = Trim(oRecordSet.Fields(24).VALUE)
            oForm01.Items("P325Bal").Specific.VALUE = Trim(oRecordSet.Fields(25).VALUE)
            oForm01.Items("M325_min").Specific.VALUE = Trim(oRecordSet.Fields(26).VALUE)
            oForm01.Items("M325_max").Specific.VALUE = Trim(oRecordSet.Fields(27).VALUE)
            oForm01.Items("M325Bal").Specific.VALUE = Trim(oRecordSet.Fields(28).VALUE)
            
        Case "InspNo"
            If oForm01.Items("InspNo").Specific.VALUE <> "" Then
                sQry = "select ItemCode, ItemName from [Z_PACKING_PD] "
                sQry = sQry & "where InspNo = '" & Trim(oDS_PS_QM006H.GetValue("U_InspNo", 0)) & "'"
                oRecordSet.DoQuery sQry
                oDS_PS_QM006H.setValue "U_ItemCode", 0, Trim(oRecordSet.Fields(0).VALUE)
                oDS_PS_QM006H.setValue "U_ItemName", 0, Trim(oRecordSet.Fields(1).VALUE)
            Else
                oDS_PS_QM006H.setValue "U_ItemCode", 0, ""
                oDS_PS_QM006H.setValue "U_ItemName", 0, ""
            End If
            
        Case "M_P60", "M_P80", "M_P100", "M_P140", "M_P200", "M_P325", "M_M325"
            M_P60 = Val(oForm01.Items("M_P60").Specific.VALUE)
            M_P80 = Val(oForm01.Items("M_P80").Specific.VALUE)
            M_P100 = Val(oForm01.Items("M_P100").Specific.VALUE)
            M_P140 = Val(oForm01.Items("M_P140").Specific.VALUE)
            M_P200 = Val(oForm01.Items("M_P200").Specific.VALUE)
            M_P325 = Val(oForm01.Items("M_P325").Specific.VALUE)
            M_M325 = Val(oForm01.Items("M_M325").Specific.VALUE)
            
            oForm01.Items("C_Sum").Specific.VALUE = M_P60 + M_P80 + M_P100 + M_P140 + M_P200 + M_P325 + M_M325
            
            
    End Select
        
    '--------------------------------------------------------------
    'Line----------------------------------------------------------
    If oUID = "Mat01" Then
        Select Case oCol
        
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
    Dim SPEC_MIN        As Currency
    Dim SPEC_MAX        As Currency
    Dim VALUE           As Currency
    Dim VALUE_MIN       As Currency
    Dim VALUE_MAX       As Currency
    
    Dim M_P60 As Currency
    Dim M_P80 As Currency
    Dim M_P100 As Currency
    Dim M_P140 As Currency
    Dim M_P200 As Currency
    Dim M_P325 As Currency
    Dim M_M325 As Currency
    
    
    ErrNum = 0

    '// Check
    Select Case True
        Case Trim(oDS_PS_QM006H.GetValue("U_InspNo", 0)) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_QM006H.GetValue("U_ItemCode", 0)) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
'        Case Trim(oDS_PS_QM006H.GetValue("U_LotNo", 0)) = ""
'            ErrNum = 3
'            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_QM006H.GetValue("U_InspDate", 0)) = ""
            ErrNum = 4
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_QM006H.GetValue("U_InspPrsn", 0)) = ""
            ErrNum = 5
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_QM006H.GetValue("U_CardCode", 0)) = ""
            ErrNum = 6
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_QM006H.GetValue("U_CardName", 0)) = ""
            ErrNum = 7
            GoTo HeaderSpaceLineDel_Error
    End Select

    '--------------------------------------------------------------------------------------------
    '//측청치의 유효값 검사
    '//화학성분 Fe
    If Val(oDS_PS_QM006H.GetValue("U_M_Fe", 0)) <> 0 Then
        VALUE_MAX = Val(oDS_PS_QM006H.GetValue("U_M_Fe", 0))
        SPEC_MAX = Val(oForm01.Items("Fe_max").Specific.VALUE)
        If VALUE_MAX > SPEC_MAX Then
            ErrNum = 10
            oForm01.Items("M_Fe").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    '//화학성분 Sn
    If Val(oDS_PS_QM006H.GetValue("U_M_Sn", 0)) <> 0 Then
        VALUE_MAX = Val(oDS_PS_QM006H.GetValue("U_M_Sn", 0))
        SPEC_MAX = Val(oForm01.Items("Sn_max").Specific.VALUE)
        If VALUE_MAX > SPEC_MAX Then
            ErrNum = 20
            oForm01.Items("M_Sn").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    '//화학성분 Pb
    If Val(oDS_PS_QM006H.GetValue("U_M_Pb", 0)) <> 0 Then
        VALUE_MAX = Val(oDS_PS_QM006H.GetValue("U_M_Pb", 0))
        SPEC_MAX = Val(oForm01.Items("Pb_max").Specific.VALUE)
        If VALUE_MAX > SPEC_MAX Then
            ErrNum = 30
            oForm01.Items("M_Pb").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    '//Cu순도
'    If Val(oDS_PS_QM006H.GetValue("U_M_Cu", 0)) <> 0 Then
    If Trim(oDS_PS_QM006H.GetValue("U_M_Cu", 0)) <> "" Then
        VALUE_MIN = Val(oDS_PS_QM006H.GetValue("U_M_Cu", 0))
        SPEC_MIN = Val(oForm01.Items("Cu_min").Specific.VALUE)
        If VALUE_MIN < SPEC_MIN Then
            ErrNum = 40
            oForm01.Items("M_Cu").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    '//겉보기밀도 Ad
    If Trim(oDS_PS_QM006H.GetValue("U_M_Ad", 0)) <> "" Then
        VALUE_MIN = Val(oDS_PS_QM006H.GetValue("U_M_Ad", 0))
        SPEC_MIN = Val(oForm01.Items("Ad_min").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("Ad_max").Specific.VALUE)
        If VALUE_MIN < SPEC_MIN Or VALUE_MIN > SPEC_MAX Then
            ErrNum = 50
            oForm01.Items("M_Ad").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    '//유동도 Fr
    If Trim(oDS_PS_QM006H.GetValue("U_M_Fr", 0)) <> "" Then
        VALUE_MAX = Val(oDS_PS_QM006H.GetValue("U_M_Fr", 0))
        SPEC_MAX = Val(oForm01.Items("Fr_max").Specific.VALUE)
        If VALUE_MAX > SPEC_MAX Then
            ErrNum = 60
            oForm01.Items("M_Fr").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    '//산소함량 M_Oxygen
    If Trim(oDS_PS_QM006H.GetValue("U_M_Oxygen", 0)) <> "" Then
        VALUE_MAX = Val(oDS_PS_QM006H.GetValue("U_M_Oxygen", 0))
        SPEC_MAX = Val(oForm01.Items("Oxy_max").Specific.VALUE)
        If VALUE_MAX > SPEC_MAX Then
            ErrNum = 70
            oForm01.Items("Oxy_max").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    M_P60 = Val(oDS_PS_QM006H.GetValue("U_M_P60", 0))
    M_P80 = Val(oDS_PS_QM006H.GetValue("U_M_P80", 0))
    M_P100 = Val(oDS_PS_QM006H.GetValue("U_M_P100", 0))
    M_P140 = Val(oDS_PS_QM006H.GetValue("U_M_P140", 0))
    M_P200 = Val(oDS_PS_QM006H.GetValue("U_M_P200", 0))
    M_P325 = Val(oDS_PS_QM006H.GetValue("U_M_P325", 0))
    M_M325 = Val(oDS_PS_QM006H.GetValue("U_M_M325", 0))
    
    
    '//입도분포 +60
    If Trim(M_P60) <> "" Then
        VALUE_MAX = M_P60
        SPEC_MIN = Val(oForm01.Items("P60_min").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("P60_max").Specific.VALUE)
        If SPEC_MIN = 0 And SPEC_MAX = 0 Then
            If oForm01.Items("P60Bal").Specific.VALUE = "N" Then
                If VALUE_MAX <> 0 Then
                    ErrNum = 80
                    oForm01.Items("M_P60").Click ct_Regular
                    GoTo HeaderSpaceLineDel_Error
                End If
            End If
        ElseIf VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
            ErrNum = 80
            oForm01.Items("M_P60").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    '//입도분포 +80
    If Trim(M_P80) <> "" Then
        VALUE_MAX = M_P80
        SPEC_MIN = Val(oForm01.Items("P80_min").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("P80_max").Specific.VALUE)
        If SPEC_MIN = 0 And SPEC_MAX = 0 Then
            If oForm01.Items("P80Bal").Specific.VALUE = "N" Then
                If VALUE_MAX <> 0 Then
                    ErrNum = 90
                    oForm01.Items("M_P80").Click ct_Regular
                    GoTo HeaderSpaceLineDel_Error
                End If
            End If
        ElseIf VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
            ErrNum = 90
            oForm01.Items("M_P80").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    '//입도분포 +100
    If Trim(M_P100) <> "" Then
        VALUE_MAX = M_P100
        SPEC_MIN = Val(oForm01.Items("P100_min").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("P100_max").Specific.VALUE)
        If SPEC_MIN = 0 And SPEC_MAX = 0 Then
            If oForm01.Items("P100Bal").Specific.VALUE = "N" Then
                If VALUE_MAX <> 0 Then
                    ErrNum = 100
                    oForm01.Items("M_P100").Click ct_Regular
                    GoTo HeaderSpaceLineDel_Error
                End If
            End If
        ElseIf VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
            ErrNum = 100
            oForm01.Items("M_P100").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    
    '//입도분포 +140
    If Trim(M_P140) <> "" Then
        VALUE_MAX = M_P140
        SPEC_MIN = Val(oForm01.Items("P140_min").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("P140_max").Specific.VALUE)
        If SPEC_MIN = 0 And SPEC_MAX = 0 Then
            If oForm01.Items("P140Bal").Specific.VALUE = "N" Then
                If VALUE_MAX <> 0 Then
                    ErrNum = 110
                    oForm01.Items("M_P140").Click ct_Regular
                    GoTo HeaderSpaceLineDel_Error
                End If
            End If
        ElseIf VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
            ErrNum = 110
            oForm01.Items("M_P140").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    
    '//입도분포 +200
    If M_P200 <> 0 Then
        VALUE_MAX = M_P200
        SPEC_MIN = Val(oForm01.Items("P200_min").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("P200_max").Specific.VALUE)
        If SPEC_MIN = 0 And SPEC_MAX = 0 Then
            If oForm01.Items("P200Bal").Specific.VALUE = "N" Then
                If VALUE_MAX <> 0 Then
                    ErrNum = 120
                    oForm01.Items("M_P200").Click ct_Regular
                    GoTo HeaderSpaceLineDel_Error
                End If
            End If
        ElseIf VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
            ErrNum = 120
            oForm01.Items("M_P200").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    '//입도분포 +325
    If M_P325 <> 0 Then
        VALUE_MAX = M_P325
        SPEC_MIN = Val(oForm01.Items("P325_min").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("P325_max").Specific.VALUE)
        If SPEC_MIN = 0 And SPEC_MAX = 0 Then
            If oForm01.Items("P325Bal").Specific.VALUE = "N" Then
                If VALUE_MAX <> 0 Then
                    ErrNum = 130
                    oForm01.Items("M_P325").Click ct_Regular
                    GoTo HeaderSpaceLineDel_Error
                End If
            End If
        ElseIf VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
            ErrNum = 130
            oForm01.Items("M_P325").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    '//입도분포 -325
    If M_M325 <> 0 Then
        VALUE_MAX = M_M325
        SPEC_MIN = Val(oForm01.Items("M325_min").Specific.VALUE)
        SPEC_MAX = Val(oForm01.Items("M325_max").Specific.VALUE)
        If SPEC_MIN = 0 And SPEC_MAX = 0 Then
            If oForm01.Items("M325Bal").Specific.VALUE = "N" Then
                If VALUE_MAX <> 0 Then
                    ErrNum = 140
                    oForm01.Items("M_M325").Click ct_Regular
                    GoTo HeaderSpaceLineDel_Error
                End If
            End If
        ElseIf VALUE_MAX < SPEC_MIN Or VALUE_MAX > SPEC_MAX Then
            ErrNum = 140
            oForm01.Items("M_M325").Click ct_Regular
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
    
    
    '//입도분포 전체합이 값이 100이어야 한다.
    If M_P60 + M_P80 + M_P100 + M_P140 + M_P200 + M_P325 + M_M325 <> 100 Then
        ErrNum = 150
        oForm01.Items("M_M325").Click ct_Regular
        GoTo HeaderSpaceLineDel_Error
    End If
    
    
    '--------------------------------------------------------------------------------------------
    '//추가완료 후 다시 자동으로 CardCode,CardName,InspPrsn을 보여주기 위해 미리 저장
    Last_CardCode = Trim(oDS_PS_QM006H.GetValue("U_CardCode", 0))
    Last_CardName = Trim(oDS_PS_QM006H.GetValue("U_CardName", 0))
    Last_InspPrsn = Trim(oDS_PS_QM006H.GetValue("U_InspPrsn", 0))

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "작지번호는 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "제품코드는 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "LOT-NO는 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "검사일은 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "검사자는 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "납품처은 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "납품처명이 없습니다. 납품처코드를 확인하여 주십시오.", "E"
        
    ElseIf ErrNum = 10 Then
        MDC_Com.MDC_GF_Message "화학성분 Fe Max 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 20 Then
        MDC_Com.MDC_GF_Message "화학성분 Sn Max 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 30 Then
        MDC_Com.MDC_GF_Message "화학성분 Pb Max 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 40 Then
        MDC_Com.MDC_GF_Message "Cu순도의 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 50 Then
        MDC_Com.MDC_GF_Message "겉보기밀도 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 60 Then
        MDC_Com.MDC_GF_Message "유동도의 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 70 Then
        MDC_Com.MDC_GF_Message "산소함량 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 80 Then
        MDC_Com.MDC_GF_Message "입도분포 +60 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 90 Then
        MDC_Com.MDC_GF_Message "입도분포 +80 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 100 Then
        MDC_Com.MDC_GF_Message "입도분포 +100 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 110 Then
        MDC_Com.MDC_GF_Message "입도분포 +140 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 120 Then
        MDC_Com.MDC_GF_Message "입도분포 +200 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 130 Then
        MDC_Com.MDC_GF_Message "입도분포 +325 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 140 Then
        MDC_Com.MDC_GF_Message "입도분포 -325 측정치가 SPEC 규격 범위를 벗어났습니다.", "E"
    ElseIf ErrNum = 150 Then
        MDC_Com.MDC_GF_Message "입도분포 측정치 합이 100이 아닙니다. ", "E"
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
'        ErrNum = 1
'        GoTo MatrixSpaceLineDel_Error
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//마지막 행 하나를 빼고 i=0부터 시작하므로 하나를 빼므로
    '//oMat01.RowCount - 2가 된다..반드시 들어 가야 하는 필수값을 확인한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    If oMat01.VisualRowCount > 0 Then
'        For i = 0 To oMat01.VisualRowCount - 2      '// Mat1에 입력값이 올바르게 들어갔는지 확인 (ErrorNumber : 2)
'            oDS_PS_QM006L.Offset = i
'            If Trim(oDS_PS_QM006L.GetValue("U_PP070No", i)) = "" Then
'                ErrNum = 2
'                oMat01.Columns("PP070No").Cells(i + 1).Click ct_Regular
'                GoTo MatrixSpaceLineDel_Error
'            End If
'        Next i
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
    '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    If oMat01.VisualRowCount > 0 Then
'        oDS_PS_QM006L.RemoveRecord oDS_PS_QM006L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
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
        MDC_Com.MDC_GF_Message "벌크포장문서 번호는 필수입니다. 확인하여 주십시오.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Function Check_Exist() As Boolean
On Error GoTo Check_Exist_Error
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim ErrNum          As Integer
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "select * from [@PS_QM006H] where U_InspNo = '" & Trim(oDS_PS_QM006H.GetValue("U_InspNo", 0)) & "'"
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
        MDC_Com.MDC_GF_Message "이미 동일한 작지번호가 존재합니다. 확인하여 주십시오.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    Check_Exist = False
End Function

Private Sub Print_Query()
On Error GoTo Print_Query_Error
    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    Dim Sub_sQry            As String
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    
    Dim BPLId           As String
    Dim PackingNo       As String
    Dim ItemCode        As String
    Dim CardCode        As String
    Dim Chk             As String
    Dim BakeTestChk As String 'BakeTest 측정치 추가 품목 체크를 위한 변수(2012.11.26 송명규 추가)
    Dim PackNo As String '패킹번호(2012.11.26 송명규 추가)
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Call ConnectODBC
    
    BPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
    
    
    '// 임시테이블에 check된항목저장
    sQry = "Delete [Z_PS_QM040] WHERE BPLId = '" & BPLId & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_QM006L.GetValue("U_ColReg01", i)) = "Y" Then
            sQry = "Insert [Z_PS_QM040] values ('" & BPLId & "', '" & Trim(oDS_PS_QM006L.GetValue("U_ColReg02", i)) & "')"
            oRecordSet01.DoQuery sQry
            
            CardCode = Trim(oDS_PS_QM006L.GetValue("U_ColReg05", i)) '// (주)TSD '12440' 찿기위해 MOVE
            ItemCode = Trim(oDS_PS_QM006L.GetValue("U_ColReg03", i))
            
        End If
    Next
    
    '// B/G타입  체크
    If MDC_PS_Common.GetValue("SELECT count(*) FROM [@PS_PP090H] a inner join [@PS_PP090L] b on a.DocEntry = b.DocEntry INNER JOIN [Z_PS_QM040] z on a.U_BPLId = z.BPLId and a.U_PackNo = z.PackNo left  join [OITM] c on b.U_ItemCode = c.ItemCode WHERE z.BPLId = '" & BPLId & "' and c.U_ItemType in ('16','17')") > 0 Then
       Chk = "Y"
    Else
       Chk = "N"
    End If
    

    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_QM041] 검사성적서출력(신양식)"
    
    If Chk = "Y" Then
       ReportName = "PS_QM041_02.RPT"  '// B/G 타입
    Else
    
        If CardCode <> "12440" Then
            ReportName = "PS_QM041_01.RPT"  '// 일반
        Else
        
            If ItemCode <> "104010098" Then '응급조치로 ItemCode로 체크(추후 품목이 추가될 경우를 생각해서 마스터데이터의 특정 값 기준으로 수정 할 필요 있음, 2012.11.26 송명규)
                ReportName = "PS_QM041_03.RPT"  '// (주)TSD
            Else
                ReportName = "PS_QM041_04.RPT"  '// (주)TSD
            End If
          
       End If
       
    End If
    
    
    '// Formula 수식필드
    ReDim gRpt_Formula(1)
    ReDim gRpt_Formula_Value(1)
    
'    gRpt_Formula(1) = "YYYY"
'    gRpt_Formula_Value(1) = Left(YearMon, 4)
'    gRpt_Formula(2) = "MM"
'    gRpt_Formula_Value(2) = Right(YearMon, 2)
    
    '// SubReport
    ReDim gRpt_SRptSqry(3)
    ReDim gRpt_SRptName(3)
    
    ReDim gRpt_SFormula(3, 1)
    ReDim gRpt_SFormula_Value(3, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    gRpt_SFormula(2, 1) = ""
    gRpt_SFormula_Value(2, 1) = ""
    
    gRpt_SFormula(3, 1) = ""
    gRpt_SFormula_Value(3, 1) = ""
    
    '// Sub_Procedure 실행
    'Sub_sQry = "EXEC [PS_QM041_03] '" & PackingNo & "', '" & ItemCode & "','" & CardCode & "'"
    
    If Chk = "Y" Then
    
       'Sub_sQry = "EXEC [PS_QM041_04] '" & BPLId & "'"
       'gRpt_SRptSqry(1) = Sub_sQry
       'gRpt_SRptName(1) = "PS_QM041_SUB_01"
       
       Sub_sQry = "EXEC [PS_QM041_04] '" & BPLId & "'"
       gRpt_SRptSqry(1) = Sub_sQry
       gRpt_SRptName(1) = "PS_QM040_SUB_01"
       
       Sub_sQry = "EXEC [PS_QM041_05] '" & BPLId & "'"
       gRpt_SRptSqry(2) = Sub_sQry
       gRpt_SRptName(2) = "PS_QM040_SUB_02"
       
       Sub_sQry = "EXEC [PS_QM041_03] '" & BPLId & "'"
       gRpt_SRptSqry(3) = Sub_sQry
       gRpt_SRptName(3) = "PS_QM040_SUB_03"
       
    Else
    
       Sub_sQry = "EXEC [PS_QM041_03] '" & BPLId & "'"
       gRpt_SRptSqry(1) = Sub_sQry
       gRpt_SRptName(1) = "PS_QM040_SUB_01"
       
       Sub_sQry = "EXEC [PS_QM041_05] '" & BPLId & "'"
       gRpt_SRptSqry(2) = Sub_sQry
       gRpt_SRptName(2) = "PS_QM040_SUB_02"
       
       Sub_sQry = "EXEC [PS_QM041_03] '" & BPLId & "'"
       gRpt_SRptSqry(3) = Sub_sQry
       gRpt_SRptName(3) = "PS_QM040_SUB_03"
       
    End If
    
    
    '/ Procedure 실행
    'sQry = "EXEC [PS_QM041_02] '" & PackingNo & "', '" & ItemCode & "','" & CardCode & "'"
    
    sQry = "EXEC [PS_QM041_02] '" & BPLId & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Print_Query_Error
    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "", "N", "V") = False Then
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "선택하신 라인에 패킹번호가 없습니다. 확인해 주세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "선택하신 라인에 품목코드가 없습니다.확인해 주세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "선택하신 라인에 납품처코드가 없습니다.확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Query_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

