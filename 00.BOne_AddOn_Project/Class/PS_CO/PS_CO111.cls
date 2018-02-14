VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_CO111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_CO111.cls
'//  Module         : CO
'//  Description    : 배부사이클 실행 > 배부사이클 실행 조회
'//  FormType       : PS_CO111
'//  Create Date    : 2010.10.21
'//  Modified Date  :
'//  Creator        : Ryu Yung Jo
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01         As SAPbouiCOM.Form
Public oMat01          As SAPbouiCOM.Matrix
Public oMat02          As SAPbouiCOM.Matrix
Private oDS_PS_CO110M  As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_CO110N  As SAPbouiCOM.DBDataSource '등록라인

Public oBaseForm01     As SAPbouiCOM.Form '//부모폼
Public oBaseItemUID01  As String
Public oBaseColUID01   As String
Public oBaseColRow01   As Long
Public oBaseSelectedLineNum01 As Long

Private oLast_Item_UID As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID  As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row  As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(ByRef oForm02 As SAPbouiCOM.Form, ByVal oItemUID02 As String, ByVal oColUID02 As String, ByVal oColRow02 As Long, ByVal SelectedLineNum As Long)
On Error GoTo LoadForm_Error

    Dim i           As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01   As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_CO111.srf")
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    
    '매트릭스의 타이틀높이와 셀높이를 고정
    For i = 1 To (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    
    oFormUniqueID01 = "PS_CO111_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '//oForm01.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때
        
    oForm01.Freeze True
    Set oBaseForm01 = oForm02
    oBaseItemUID01 = oItemUID02
    oBaseColUID01 = oColUID02
    oBaseColRow01 = oColRow02
    oBaseSelectedLineNum01 = SelectedLineNum
    
    Call CreateItems
    Call LoadData01
'    Call ComboBox_Setting
'    Call FormItemEnabled
    
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
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

'****************************************************************************************************************
'// ItemEventHander
'****************************************************************************************************************
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error

    Dim i&
    Dim ErrNum&
    Dim TempForm01 As Variant
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "Btn01" Then
                    oForm01.Close
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
                Call Raise_EVENT_CLICK(FormUID, pval, BubbleEvent)
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
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
'et_DOUBLE_CLICK ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_DOUBLE_CLICK: '//7
                If pval.ItemUID = "Mat01" And pval.Row <> 0 Then
                    Call LoadData02(Trim(oMat01.Columns("LineId").Cells(pval.Row).Specific.VALUE))
                End If
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            Case et_MATRIX_LOAD: '//11
            
                Call oMat01.AutoResizeColumns
                
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
                Set oDS_PS_CO110M = Nothing
                Set oDS_PS_CO110N = Nothing
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
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

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If pval.BeforeAction = True Then
    
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub CreateItems()
On Error GoTo CreateItems_Error

    Dim oQuery01     As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_CO110M = oForm01.DataSources.DBDataSources("@PS_CO110M")
    Set oDS_PS_CO110N = oForm01.DataSources.DBDataSources("@PS_CO110N")
    
    Set oMat01 = oForm01.Items("Mat01").Specific
    Set oMat02 = oForm01.Items("Mat02").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
    oMat02.SelectionMode = ms_NotSupported
    oMat02.AutoResizeColumns
     
    Set oRecordSet01 = Nothing
    
    Call oForm01.DataSources.UserDataSources.Add("Opt01", dt_SHORT_TEXT, 10)
    Call oForm01.DataSources.UserDataSources.Add("Opt02", dt_SHORT_TEXT, 10)
    Call oForm01.Items("RadioBtn01").Specific.DataBind.SetBound(True, "", "Opt01")
    Call oForm01.Items("RadioBtn02").Specific.DataBind.SetBound(True, "", "Opt02")
    Call oForm01.Items("RadioBtn01").Specific.GroupWith("RadioBtn02")
    
    
    Call oForm01.DataSources.UserDataSources.Add("SumCost1", dt_SUM)
    Call oForm01.Items("SumCost1").Specific.DataBind.SetBound(True, "", "SumCost1")
    
    Call oForm01.DataSources.UserDataSources.Add("SumCost2", dt_SUM)
    Call oForm01.Items("SumCost2").Specific.DataBind.SetBound(True, "", "SumCost2")
    
    Call oForm01.DataSources.UserDataSources.Add("SumCost3", dt_SUM)
    Call oForm01.Items("SumCost3").Specific.DataBind.SetBound(True, "", "SumCost3")
    
    Call oForm01.DataSources.UserDataSources.Add("SumCost4", dt_SUM)
    Call oForm01.Items("SumCost4").Specific.DataBind.SetBound(True, "", "SumCost4")
    
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
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// 사업장
'    Set oCombo = oForm01.Items("BPLId").Specific
'    sQry = "SELECT BPLId, BPLName From [OBPL]"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).Value), Trim(oRecordSet01.Fields(1).Value)
'        oRecordSet01.MoveNext
'    Loop
'    oCombo.Select "4", psk_ByValue
    
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "ComboBox_Setting_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
On Error GoTo FlushToItemValue_Error

    Dim i&
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "CardCode"
            sQry = "Select CardName From [OCRD] Where CardCode = '" & Trim(oForm01.DataSources.UserDataSources("CardCode").VALUE) & "'"
            oRecordSet01.DoQuery sQry
            
            oForm01.DataSources.UserDataSources("CardName").VALUE = Trim(oRecordSet01.Fields(0).VALUE)
    End Select

    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub LoadData01()
On Error GoTo LoadData01_Error

    Dim i    As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim Code$
        
    Code = Trim(oBaseForm01.Items("Code").Specific.VALUE)
    
    sQry = "Select * From [@PS_CO110M] Where Code = '" & Code & "' And U_L_LineId = '" & oBaseSelectedLineNum01 & "'"
    oRecordSet01.DoQuery sQry
            
    oMat01.Clear
    oDS_PS_CO110M.Clear
    
    If oRecordSet01.RecordCount = 0 Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_CO110M.Size Then
            oDS_PS_CO110M.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_CO110M.Offset = i
        oDS_PS_CO110M.setValue "U_LineNum", i, i + 1
        oDS_PS_CO110M.setValue "U_CCCode", i, Trim(oRecordSet01.Fields("U_CCCode").VALUE)
        oDS_PS_CO110M.setValue "U_CCName", i, Trim(oRecordSet01.Fields("U_CCName").VALUE)
        oDS_PS_CO110M.setValue "U_CECode", i, Trim(oRecordSet01.Fields("U_CECode").VALUE)
        oDS_PS_CO110M.setValue "U_CEName", i, Trim(oRecordSet01.Fields("U_CEName").VALUE)
        oDS_PS_CO110M.setValue "U_Cost1", i, Trim(oRecordSet01.Fields("U_Cost1").VALUE)
        oDS_PS_CO110M.setValue "U_Cost2", i, Trim(oRecordSet01.Fields("U_Cost2").VALUE)
        oDS_PS_CO110M.setValue "U_Cost3", i, Trim(oRecordSet01.Fields("U_Cost3").VALUE)
        oDS_PS_CO110M.setValue "U_Cost4", i, Trim(oRecordSet01.Fields("U_Cost4").VALUE)
        oDS_PS_CO110M.setValue "LineId", i, oRecordSet01.Fields("LineId").VALUE
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat01.LoadFromDataSource
    oMat01.AutoResizeColumns
    ProgBar01.Stop
    oForm01.Freeze False
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData01_Error:
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData01_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub LoadData02(sLineNum As Long)
On Error GoTo LoadData02_Error

    Dim i    As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim Code$
    Dim SumCost1 As Long, SumCost2 As Long, SumCost3 As Long, SumCost4 As Long
    
    Code = Trim(oBaseForm01.Items("Code").Specific.VALUE)
    
    sQry = "Select * From [@PS_CO110N] Where Code = '" & Code & "' And U_M_LineId = '" & sLineNum & "'"
    oRecordSet01.DoQuery sQry
            
    oMat02.Clear
    oDS_PS_CO110N.Clear
    
    If oRecordSet01.RecordCount = 0 Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_CO110N.Size Then
            oDS_PS_CO110N.InsertRecord (i)
        End If
        
        oMat02.AddRow
        oDS_PS_CO110N.Offset = i
        oDS_PS_CO110N.setValue "U_LineNum", i, i + 1
        oDS_PS_CO110N.setValue "U_CCCode", i, Trim(oRecordSet01.Fields("U_CCCode").VALUE)
        oDS_PS_CO110N.setValue "U_CCName", i, Trim(oRecordSet01.Fields("U_CCName").VALUE)
        oDS_PS_CO110N.setValue "U_CECode", i, Trim(oRecordSet01.Fields("U_CECode").VALUE)
        oDS_PS_CO110N.setValue "U_CEName", i, Trim(oRecordSet01.Fields("U_CEName").VALUE)
        oDS_PS_CO110N.setValue "U_Cost1", i, Trim(oRecordSet01.Fields("U_Cost1").VALUE)
        oDS_PS_CO110N.setValue "U_Cost2", i, Trim(oRecordSet01.Fields("U_Cost2").VALUE)
        oDS_PS_CO110N.setValue "U_Cost3", i, Trim(oRecordSet01.Fields("U_Cost3").VALUE)
        oDS_PS_CO110N.setValue "U_Cost4", i, Trim(oRecordSet01.Fields("U_Cost4").VALUE)
        
        SumCost1 = SumCost1 + Trim(oRecordSet01.Fields("U_Cost1").VALUE)
        SumCost2 = SumCost2 + Trim(oRecordSet01.Fields("U_Cost2").VALUE)
        SumCost3 = SumCost3 + Trim(oRecordSet01.Fields("U_Cost3").VALUE)
        SumCost4 = SumCost4 + Trim(oRecordSet01.Fields("U_Cost4").VALUE)
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
    Next
    
    oForm01.Items("SumCost1").Specific.VALUE = SumCost1
    oForm01.Items("SumCost2").Specific.VALUE = SumCost2
    oForm01.Items("SumCost3").Specific.VALUE = SumCost3
    oForm01.Items("SumCost4").Specific.VALUE = SumCost4
    
    oMat02.LoadFromDataSource
    oMat02.AutoResizeColumns
    ProgBar01.Stop
    oForm01.Freeze False
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData02_Error:
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData02_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
Private Sub Raise_EVENT_CLICK(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_CLICK_Error

    Dim True_False$
    
    If pval.BeforeAction = True Then
    
        If pval.ItemUID = "Mat01" Then
            If pval.Row > 0 Then

                Call oMat01.SelectRow(pval.Row, True, False)
                
            End If
        End If
    
        If pval.ItemUID = "Mat02" Then
            If pval.Row > 0 Then

                Call oMat02.SelectRow(pval.Row, True, False)
                
            End If
        End If
    
        If pval.ItemUID = "RadioBtn01" Then
            oForm01.Freeze True
            oForm01.Settings.MatrixUID = "Mat01"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oForm01.Freeze False
        End If
        If pval.ItemUID = "RadioBtn02" Then
            oForm01.Freeze True
            oForm01.Settings.MatrixUID = "Mat02"
            oForm01.Settings.EnableRowFormat = True
            oForm01.Settings.Enabled = True
            oForm01.Freeze False
        End If
    ElseIf pval.BeforeAction = False Then
    
    End If
    Exit Sub
Raise_EVENT_CLICK_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_CLICK_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

