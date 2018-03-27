VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_CO080"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_CO080cls
'//  Module         : CO
'//  Description    : 코스트센터비용집계
'//  FormType       : PS_CO080
'//  Create Date    : 2010.11.02
'//  Creator        : Lee Byong Gak
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Private oDS_PS_CO080H As SAPbouiCOM.DBDataSource '등록헤더
Private oDS_PS_CO080L As SAPbouiCOM.DBDataSource '등록라인

Private oLastItemUID01 As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLastColUID01 As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLastColRow01 As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
Private ItemInformationCount As Long
'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_CO080.srf")
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
    
    oFormUniqueID01 = "PS_CO080_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    oForm01.DataBrowser.BrowseBy = "Code"     '//UDO방식일때

    oForm01.EnableMenu ("1293"), True      '// 행삭제
    oForm01.EnableMenu ("1287"), False     '// 복제
   
    oForm01.Freeze True
    Call CreateItems
    Call AddMatrixRow(0, True)
    Call ComboBox_Setting
    
    oForm01.Update
    oForm01.Freeze False
    
    oForm01.Visible = True
    Set oXmlDoc01 = Nothing
    Exit Sub
LoadForm_Error:
    oForm01.Update
    oForm01.Freeze False
    Set oXmlDoc01 = Nothing
    Set oForm01 = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error

    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim Code$
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    Code = Trim(oDS_PS_CO080H.GetValue("U_YM", 0)) & Trim(oDS_PS_CO080H.GetValue("U_BPLId", 0))
                    Call oDS_PS_CO080H.setValue("Code", 0, Code)
                    Call oDS_PS_CO080H.setValue("Name", 0, Code)
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
     End If
            Case et_KEY_DOWN: '//2
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
                If pval.ItemUID = "Btn01" Then
                    Call LoadData
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            '// 전기일 시작 종료
               sQry = "select F_RefDate,T_RefDate from OFPR Where CONVERT(CHAR(6),F_RefDate,112) = '" & Trim(oForm01.Items("YM").Specific.VALUE) & "'"
               oRecordSet01.DoQuery sQry
        
               oForm01.Items("FRefDate").Specific.VALUE = Format(Trim(oRecordSet01.Fields(0).VALUE), "YYYYMMDD")
               oForm01.Items("TRefDate").Specific.VALUE = Format(Trim(oRecordSet01.Fields(1).VALUE), "YYYYMMDD")
               
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then                         '코스트센터가 바뀌면 한 행을 추가
                    If pval.ColUID = "ActCode" Then
                       FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                    End If
                End If
            Case et_MATRIX_LOAD: '//11
                Call AddMatrixRow(pval.Row, False)
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_ItemEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo MenuEvent_Error
    
    Dim i&
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                  Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
                 oForm01.DataBrowser.BrowseBy = "Code" '//UDO방식일때
            Case "1282": '추가
                 oForm01.DataBrowser.BrowseBy = "Code" '//UDO방식일때
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 1 To oMat01.VisualRowCount
                        oMat01.Columns("LineId").Cells(i).Specific.VALUE = i
                    Next i
                    oMat01.FlushToDataSource          ' DBDataSource에 레코드가 한줄 더 생긴다.
                    Call oDS_PS_CO080L.RemoveRecord(oDS_PS_CO080L.Size - 1)     ' 레코드 한 줄을 지운다.
                    oMat01.LoadFromDataSource         ' DBDataSource를 매트릭스에 올리고
                    If oMat01.RowCount = 0 Then
'                        Call AddMatrixRow(1, 0, True)
                        Call AddMatrixRow(1, 0)
                    Else
                        If Trim(oDS_PS_CO080L.GetValue("U_ActCode", oMat01.RowCount - 1)) <> "" Then
'                            Call AddMatrixRow(1, oMat01.RowCount, True)
                            Call AddMatrixRow(1, oMat01.RowCount)
                        End If
                    End If
                End If
            Case "1281": '찾기
                 Call AddMatrixRow(0, True) '//UDO방식
                 oForm01.DataBrowser.BrowseBy = "Code" '//UDO방식일때        '찾기버튼 클릭시 Matrix에 행 추가
            Case "1282": '추가
'                 Call FormItemEnabled '//UDO방식
                 Call AddMatrixRow(0, True) '//UDO방식
            Case "1288", "1289", "1290", "1291": '레코드이동버튼             '추가버튼 클릭시 Matrix에 행 추가
                '//Call FormItemEnabled
        End Select
    End If
    Exit Sub
MenuEvent_Error:
    Sbo_Application.SetStatusBarMessage "MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function CreateItems() As Boolean
    On Error GoTo CreateItems_Error
    Call oForm01.Freeze(True)
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_CO080H = oForm01.DataSources.DBDataSources("@PS_CO080H")
    Set oDS_PS_CO080L = oForm01.DataSources.DBDataSources("@PS_CO080L")

    Set oMat01 = oForm01.Items("Mat01").Specific
    
    Call oForm01.DataSources.UserDataSources.Add("Amount", dt_SUM, 15)
    Call oForm01.Items("Amount").Specific.DataBind.SetBound(True, "", "Amount")
'    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
   
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Exit Function
CreateItems_Error:
    Set oRecordSet01 = Nothing
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    Call oForm01.Freeze(True)
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        'Call FormClear '//UDO방식
        Call oForm01.EnableMenu("1281", True) '//찾기
        Call oForm01.EnableMenu("1282", False) '//추가
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        Call oForm01.EnableMenu("1281", False) '//찾기
        Call oForm01.EnableMenu("1282", True) '//추가
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
    End If
    Call oForm01.Freeze(False)
    Exit Sub
FormItemEnabled_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub AddMatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo AddMatrixRow_Error
    
    Call oForm01.Freeze(True)
    If RowIserted = False Then '//행추가여부
        oRow = oMat01.RowCount
        oDS_PS_CO080L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_CO080L.Offset = oRow
    oDS_PS_CO080L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Call oForm01.Freeze(False)
    Exit Sub
    
AddMatrixRow_Error:
    Call oForm01.Freeze(False)
    Sbo_Application.SetStatusBarMessage "AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_RightClickEvent_Error
    If (eventInfo.BeforeAction = True) Then
        '//작업
    ElseIf (eventInfo.BeforeAction = False) Then
        '//작업
    End If
    Exit Sub
Raise_RightClickEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ROW_DELETE_Error
    Dim i As Long
    If (oLastColRow01 > 0) Then
        If pval.BeforeAction = True Then
            '//행삭제전 행삭제가능여부검사
        ElseIf pval.BeforeAction = False Then
            For i = 1 To oMat01.VisualRowCount
                oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
            Next i
            oMat01.FlushToDataSource
            Call oDS_PS_CO080L.RemoveRecord(oDS_PS_CO080L.Size - 1)
            oMat01.LoadFromDataSource
            If oMat01.RowCount = 0 Then
                Call AddMatrixRow(0)
            Else
                If Trim(oDS_PS_CO080L.GetValue("U_CECode", oMat01.RowCount - 1)) <> "" Then
                    Call AddMatrixRow(oMat01.RowCount)
                End If
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error
    Dim i             As Long
    Dim ErrNum        As Integer
    Dim oRecordSet01  As SAPbobsCOM.Recordset
    Dim sQry          As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            
    ErrNum = 0
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '// 화면상의 메트릭스에 입력된 내용을 모두 디비데이터소스로 넘긴다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oMat01.FlushToDataSource
    
    '// 라인
    If oMat01.VisualRowCount < 1 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '// 맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
    '// 이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    If oMat01.VisualRowCount > 0 Then
'        For i = 0 To oMat01.VisualRowCount - 1
'            oDS_PS_CO080L.Offset = i
'            Select Case True
'                Case oDS_PS_CO080L.GetValue("U_CoCtCode", i) = ""
'                    ErrNum = 2
'                    GoTo MatrixSpaceLineDel_Error
'                Case oDS_PS_CO080L.GetValue("U_CoCtName", i) = ""
'                    ErrNum = 3
'                    GoTo MatrixSpaceLineDel_Error
'                Case oDS_PS_CO080L.GetValue("U_CoElCode", i) = ""
'                    ErrNum = 4
'                    GoTo MatrixSpaceLineDel_Error
'                Case oDS_PS_CO080L.GetValue("U_CoElName", i) = ""
'                    ErrNum = 5
'                    GoTo MatrixSpaceLineDel_Error
'                Case oDS_PS_CO080L.GetValue("U_Class", i) = ""
'                    ErrNum = 6
'                    GoTo MatrixSpaceLineDel_Error
'                Case oDS_PS_CO080L.GetValue("U_Amount", i) = ""
'                    ErrNum = 7
'                    GoTo MatrixSpaceLineDel_Error
'            End Select
'        Next i
        
        If oDS_PS_CO080L.GetValue("U_CoCtCode", oMat01.VisualRowCount - 1) = "" Then
            oDS_PS_CO080L.RemoveRecord oMat01.VisualRowCount - 1
        End If
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '행을 삭제하였으니 DB데이터 소스를 다시 가져온다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oMat01.LoadFromDataSource
    
    Set oRecordSet01 = Nothing
    MatrixSpaceLineDel = True
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인데이타가 없습니다.. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "코스트센터는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "코스트센터명은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "원가요소는 필수입력사항입니다. 확인하세요", "E"
    ElseIf ErrNum = 5 Then
        MDC_Com.MDC_GF_Message "원가요소명은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 6 Then
        MDC_Com.MDC_GF_Message "구분은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 7 Then
        MDC_Com.MDC_GF_Message "금액은 필수입력사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum      As Integer
    Dim DocNum      As String
    
    ErrNum = 0
    
    '// Check
    Select Case True
        Case oDS_PS_CO080H.GetValue("Code", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_CO080H.GetValue("U_FRefDate", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_CO080H.GetValue("U_TRefDate", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
    End Select
   
   HeaderSpaceLineDel = True
      
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Description, "E"
    Else
        MDC_Com.MDC_GF_Message "정상등록 되었습니다.", "W"
    End If
'        HeaderSpaceLineDel = False
End Function

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
    Dim i&
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    
    Dim sSeq$

    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat01.FlushToDataSource
        
   Select Case oCol
        Case "ActCode"
            oMat01.FlushToDataSource
            oDS_PS_CO080L.Offset = oRow - 1
                
            oForm01.Freeze True
            oForm01.Freeze False
          oMat01.LoadFromDataSource
          
            '--------------------------------------------------------------------------------------------
            If oRow = oMat01.RowCount And Trim(oDS_PS_CO080L.GetValue("U_ActCode", oRow - 1)) <> "" Then
            '// 다음 라인 추가
            Call AddMatrixRow(0, False)
            oMat01.Columns("ActCode").Cells(oRow).Click ct_Regular
            End If
                   
   End Select
   
    Set oRecordSet01 = Nothing
    Exit Sub
       
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FlushToItemValue_Error:
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "원가요소코드를 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
    Else
        MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    '//콤보에 기본값설정
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim sQry            As String
    Dim oRecordSet01    As SAPbobsCOM.Recordset
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    '// 마감년월
'    Set oCombo = oForm01.Items("Code").Specific
'    sQry = "select Code, Name  from OFPR Order by Code"
'    oRecordSet01.DoQuery sQry
'    Do Until oRecordSet01.EOF
'        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
'        oRecordSet01.MoveNext
'    Loop

    '// 사업장
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL] order by 1"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop


    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_CO080", "Mat01", "Class", "11", "재료비")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_CO080", "Mat01", "Class", "12", "노무비")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_CO080", "Mat01", "Class", "13", "경비")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_CO080", "Mat01", "Class", "21", "매출")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_CO080", "Mat01", "Class", "22", "매출원가")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_CO080", "Mat01", "Class", "24", "판관비")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_CO080", "Mat01", "Class", "25", "영업외수익")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_CO080", "Mat01", "Class", "26", "영업외비용")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_CO080", "Mat01", "Class", "27", "특별이익")
    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_CO080", "Mat01", "Class", "28", "특별손실")
    
    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Class"), "PS_CO080", "Mat01", "Class")
    
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
ComboBox_Setting_Error:
    Set oCombo = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "ComboBox_Setting_Error:" & Err.Number & " - " & Err.Description, "W"
End Sub

Public Sub LoadData()
On Error GoTo LoadData_Error
    Dim i As Integer
    Dim sQry As String
    Dim cnt As Integer
    Dim j   As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
'    Dim FRefDate$, TRefDate$
    Dim ProfitCode$, OcrName$, U_CECode$, U_CEName$, U_Class$, Debit$, Credit$, PrcAmount$, OcrTotal$
    Dim TotalAmount As Currency
    oForm01.Freeze True
    
    Dim iFrDate$, iToDate$, iBPLId$
    
    iFrDate = Trim(oForm01.Items("FRefDate").Specific.VALUE)
    iToDate = Trim(oForm01.Items("TRefDate").Specific.VALUE)
    iBPLId = Trim(oForm01.Items("BPLId").Specific.VALUE)
   
    sQry = "EXEC [PS_CO080_01] '" & iFrDate & "','" & iToDate & "','" & iBPLId & "'"

    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_CO080L.Clear
    
    For i = 0 To oRecordSet01.RecordCount - 1
            If i + 1 > oDS_PS_CO080L.Size Then
                oDS_PS_CO080L.InsertRecord (i)
            End If
            
        oMat01.AddRow
        oDS_PS_CO080L.Offset = i
        oDS_PS_CO080L.setValue "U_LineNum", i, i + 1
        oDS_PS_CO080L.setValue "U_CoCtCode", i, Trim(oRecordSet01.Fields(0).VALUE)
        oDS_PS_CO080L.setValue "U_CoCtName", i, Trim(oRecordSet01.Fields(1).VALUE)
        oDS_PS_CO080L.setValue "U_CoElCode", i, Trim(oRecordSet01.Fields(2).VALUE)
        oDS_PS_CO080L.setValue "U_CoElName", i, Trim(oRecordSet01.Fields(3).VALUE)
        oDS_PS_CO080L.setValue "U_Class", i, Trim(oRecordSet01.Fields(4).VALUE)
        oDS_PS_CO080L.setValue "U_Amount", i, Trim(oRecordSet01.Fields(5).VALUE)
        
        TotalAmount = TotalAmount + Trim(oRecordSet01.Fields(5).VALUE)
   
        oRecordSet01.MoveNext
    Next
        oForm01.DataSources.UserDataSources("Amount").VALUE = TotalAmount
        oMat01.LoadFromDataSource
        Call AddMatrixRow(i, False)
        oForm01.Freeze False
  
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData_Error:
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub
