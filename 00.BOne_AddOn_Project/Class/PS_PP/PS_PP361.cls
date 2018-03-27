VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_PP361"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File : PS_PP361.cls
'//  Module : PP
'//  Description : 작지별 생산진행현황
'//  FormType : PS_PP361
'//  Create Date : 2011.03.18
'//  Modified Date :
'//  Creator : Lee Byong Gak
'//  Company : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01 As SAPbouiCOM.Form
Public oMat01 As SAPbouiCOM.Matrix
Public oGrid01 As SAPbouiCOM.Grid
Public oGrid02 As SAPbouiCOM.Grid
Public oGrid03 As SAPbouiCOM.Grid

Private oDS_PS_PP361H As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_PP361L As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID As String '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID As String '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row As Long '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01 As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_PP361.srf")
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
    
    oFormUniqueID01 = "PS_PP361_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_OK_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
'    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    oForm01.Items("SYYYYMM").Click ct_Regular
    
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
'    Set oDS_PS_PP361H = oForm01.DataSources.DBDataSources("@PS_PP361H")
'    Set oDS_PS_PP361L = oForm01.DataSources.DBDataSources("@PS_PP361L")
    
    '// 메트릭스 개체 할당
'    Set oMat01 = oForm01.Items("Mat01").Specific
    Set oGrid01 = oForm01.Items("Grid01").Specific
    oGrid01.SelectionMode = ms_Auto
    Set oGrid02 = oForm01.Items("Grid02").Specific
    oGrid02.SelectionMode = ms_Auto
    Set oGrid03 = oForm01.Items("Grid03").Specific
    oGrid03.SelectionMode = ms_Auto
    
    oForm01.DataSources.DataTables.Add ("ZTEMP01")
    oForm01.DataSources.DataTables.Add ("ZTEMP02")
    oForm01.DataSources.DataTables.Add ("ZTEMP03")
       
    Call oForm01.DataSources.UserDataSources.Add("SYYYYMM", dt_SHORT_TEXT, 7)
    oForm01.Items("SYYYYMM").Specific.DataBind.SetBound True, "", "SYYYYMM"
    oForm01.DataSources.UserDataSources.Item("SYYYYMM").VALUE = Format(Now, "YYYY-MM")
    
    Call oForm01.DataSources.UserDataSources.Add("EYYYYMM", dt_SHORT_TEXT, 7)
    oForm01.Items("EYYYYMM").Specific.DataBind.SetBound True, "", "EYYYYMM"
    oForm01.DataSources.UserDataSources.Item("EYYYYMM").VALUE = Format(Now, "YYYY-MM")
    
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

    '//거래처구분(사업장이 아님, 2012.01.02 송명규 수정)
    Call oForm01.Items("BPLId").Specific.ValidValues.Add("ALL", "전체")
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'C100' ORDER BY Code", "", False, False)
    Call oForm01.Items("BPLId").Specific.Select(0, psk_Index)
    
'    Set oCombo = oForm01.Items("BPLId").Specific
'    Call oCombo.ValidValues.Add("ALL", "전체")
'    Call oCombo.ValidValues.Add("A", "안강")
'    Call oCombo.ValidValues.Add("C", "온산")
'    Call oCombo.ValidValues.Add("D", "동래")
'    Call oCombo.ValidValues.Add("E", "신동")
'    Call oCombo.ValidValues.Add("F", "멀티")
'    Call oCombo.ValidValues.Add("H", "사외")
'    Call oCombo.ValidValues.Add("I", "계열사")
'    Call oCombo.ValidValues.Add("P", "부평")
'    Call oCombo.ValidValues.Add("U", "수출")
'    Call oCombo.Select(0, psk_Index)

    '//품목구분
    Set oCombo = oForm01.Items("ItemGB").Specific
    sQry = "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'S002' ORDER BY Code"
    oRecordSet01.DoQuery sQry
    
    Call oCombo.ValidValues.Add("A", "전체")
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    Call oCombo.Select(0, psk_Index)

    
    '//구분
    Set oCombo = oForm01.Items("Section").Specific
    Call oCombo.ValidValues.Add("A", "전체")
    Call oCombo.ValidValues.Add("B", "미완료")
    Call oCombo.ValidValues.Add("C", "완료")
    Call oCombo.Select(0, psk_Index)
    
    
    '//선택
    Set oCombo = oForm01.Items("Selection").Specific
    Call oCombo.ValidValues.Add("000", "선택")
    Call oCombo.ValidValues.Add("100", "자재비내역")
    Call oCombo.ValidValues.Add("200", "자체가공비내역")
    Call oCombo.ValidValues.Add("300", "설계비내역")
'    Call oCombo.ValidValues.Add("400", "공정내역")
    Call oCombo.ValidValues.Add("400", "외주가공비내역")
    Call oCombo.ValidValues.Add("500", "외주제작비내역")
'    Call oCombo.ValidValues.Add("600", "공정내역")
    Call oCombo.Select(0, psk_Index)
    
    
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
                If pval.ItemUID = "Search" Then 'Search 버튼
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        If oGrid02.Rows.Count > 0 Then
                            oGrid02.DataTable.Clear
                        End If
                        If oGrid03.Rows.Count > 0 Then
                            oGrid03.DataTable.Clear
                        End If
                        Call Search_Grid01_Data 'Main작번Grid
                        Call oForm01.Items("Selection").Specific.Select(0, psk_Index)
                        oForm01.Items("SYYYYMM").Click ct_Regular
                    End If
                End If

                If pval.ItemUID = "Btn01" Then
                    Call Print_Query
                End If
                
                If pval.ItemUID = "Btn02" Then
                    Call Print_Query1
                End If
        
                If pval.ItemUID = "Btn03" Then
                    Call Print_Query2
                End If

            Case et_KEY_DOWN: '//2
            
                If pval.ItemUID = "CardCode" Then
        
                    Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "CardCode", "") '거래처 포맷서치 설정
                    
                End If
            
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            
                Dim Query01 As String
                Dim RecordSet01 As SAPbobsCOM.Recordset
                Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            
            
                If pval.ItemUID = "CardCode" Then '거래처코드
            
                    Query01 = "SELECT CardName FROM OCRD WHERE CardCode = '" & oForm01.Items(pval.ItemUID).Specific.VALUE & "'"
                    Call RecordSet01.DoQuery(Query01)
                    oForm01.Items("CardName").Specific.VALUE = Trim(RecordSet01.Fields(0).VALUE)
                    
                End If
            
                Set RecordSet01 = Nothing
            
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
'            Case et_ITEM_PRESSED: '//1
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
                If pval.ItemUID = "Selection" And oGrid02.Rows.Count <> 0 Then
                    If oForm01.Items("Selection").Specific.Selected.VALUE = "100" Then '자재비내역
                        Call Search_Grid03_Data
                    End If
                    If oForm01.Items("Selection").Specific.Selected.VALUE = "200" Then '자체가공비내역
                        Call Search_Grid03_Data
                    End If
                    If oForm01.Items("Selection").Specific.Selected.VALUE = "300" Then '설계비내역
                        Call Search_Grid03_Data
                    End If
                    If oForm01.Items("Selection").Specific.Selected.VALUE = "400" Then '외주가공비내역
                        Call Search_Grid03_Data
                    End If
                    If oForm01.Items("Selection").Specific.Selected.VALUE = "500" Then '외주제작비내역
                        Call Search_Grid03_Data
                    End If
'                    If oForm01.Items("Selection").Specific.Selected.VALUE = "600" Then '공정내역
'                        Call Search_Grid03_Data
'                    End If
                End If
            Case et_CLICK: '//6
                If pval.ItemUID = "Grid01" And pval.ColUID = "RowsHeader" Then
                    If oGrid03.Rows.Count > 0 Then
                        oGrid03.DataTable.Clear
                    End If
                    If oForm01.Items("Selection").Specific.Selected.VALUE <> "000" Then
                        Call oForm01.Items("Selection").Specific.Select(0, psk_Index)
                        oForm01.Items("SYYYYMM").Click ct_Regular
                    End If
                    
                    Call Search_Grid02_Data
                End If
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
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
'                Set oMat01 = Nothing
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
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
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
    
    ErrNum = 0

    '// Check
    Select Case True
        Case Trim(oForm01.Items("SYYYYMM").Specific.VALUE) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Len(Trim(oForm01.Items("SYYYYMM").Specific.VALUE)) <> 7
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("EYYYYMM").Specific.VALUE) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Len(Trim(oForm01.Items("EYYYYMM").Specific.VALUE)) <> 7
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
    End Select
    
    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "해당년월의 시작은 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "시작년월(YYYY-MM)의 형식을 확인하여 주십시오.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "해당년월의 종료는 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "종료년월(YYYY-MM)의 형식을 확인하여 주십시오.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Sub Form_Resize(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Form_Resize_Error
    Dim i   As Integer
        
    'Grid01
    oForm01.Items("Grid01").Top = 64
    oForm01.Items("Grid01").Height = (oForm01.Height / 2) - 130
    oForm01.Items("Grid01").Left = 10
    oForm01.Items("Grid01").Width = oForm01.Width - 35

    'Btn01
    oForm01.Items("Btn01").Top = oForm01.Items("Grid01").Top + oForm01.Items("Grid01").Height + 10
    oForm01.Items("Btn01").Left = oForm01.Width - 90
    
    'Sub 작번
    oForm01.Items("S200").Top = oForm01.Items("Grid01").Top + oForm01.Items("Grid01").Height + 35

    'Grid02
    oForm01.Items("Grid02").Top = oForm01.Items("Grid01").Top + oForm01.Items("Grid01").Height + 50
    oForm01.Items("Grid02").Height = (oForm01.Height / 2) - 300
    oForm01.Items("Grid02").Left = 10
    oForm01.Items("Grid02").Width = oForm01.Width - 35
    
    'Btn02
    oForm01.Items("Btn02").Top = oForm01.Items("Grid02").Top + oForm01.Items("Grid02").Height + 10
    oForm01.Items("Btn02").Left = oForm01.Width - 90
    
    '선택
    oForm01.Items("22").Top = oForm01.Items("Grid02").Top + oForm01.Items("Grid02").Height + 35
    'Selection
    oForm01.Items("Selection").Top = oForm01.Items("Grid02").Top + oForm01.Items("Grid02").Height + 35

    'Grid03
    oForm01.Items("Grid03").Top = oForm01.Items("Grid02").Top + oForm01.Items("Grid02").Height + 50
    oForm01.Items("Grid03").Height = (oForm01.Height / 2) - 280
    oForm01.Items("Grid03").Left = 10
    oForm01.Items("Grid03").Width = oForm01.Width - 35
    
    'Btn03
    oForm01.Items("Btn03").Top = oForm01.Items("Grid03").Top + oForm01.Items("Grid03").Height + 10
    oForm01.Items("Btn03").Left = oForm01.Width - 90
    
    If oGrid01.Rows.Count > 0 Then
        oGrid01.AutoResizeColumns
    End If
    If oGrid02.Rows.Count > 0 Then
        oGrid02.AutoResizeColumns
    End If
    If oGrid03.Rows.Count > 0 Then
        oGrid03.AutoResizeColumns
    End If
    'oMat01.AutoResizeColumns
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Form_Resize_Error:
    Sbo_Application.SetStatusBarMessage "Form_Resize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Search_Grid01_Data()
    On Error GoTo Search_Grid01_Data_Error
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim sQry As String
    Dim SYYYYMM As String   '작번등록년월 시작
    Dim EYYYYMM As String   '작번등록년월 종료
    Dim BPLID As String   '사업장
    Dim ItemGB As String   '품목구분
    Dim Section As String   '구분
    Dim ItemName As String   '품명
    Dim Size As String   '규격
    Dim OrdNum As String   '작번
    Dim CardCode As String '거래처코드(2012.01.02 송명규 추가)
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    SYYYYMM = Trim(oForm01.Items("SYYYYMM").Specific.VALUE)
    EYYYYMM = Trim(oForm01.Items("EYYYYMM").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.Selected.VALUE)
    ItemGB = Trim(oForm01.Items("ItemGB").Specific.Selected.VALUE)
    Section = Trim(oForm01.Items("Section").Specific.Selected.VALUE)
    ItemName = Trim(oForm01.Items("ItemName").Specific.VALUE)
    Size = Trim(oForm01.Items("Size").Specific.VALUE)
    OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE)
    CardCode = IIf(Trim(oForm01.Items("CardCode").Specific.VALUE) = "", "%", Trim(oForm01.Items("CardCode").Specific.VALUE))
    
    
    If ItemGB = "A" Then ItemGB = "%"
    If Section = "A" Then Section = "%"
    If BPLID = "ALL" Then BPLID = "%"
    If ItemName = "" Then ItemName = "%"
    If Size = "" Then Size = "%"
    If OrdNum = "" Then OrdNum = "%"

    oForm01.Freeze True
    
    sQry = "EXEC PS_PP361_01 '" & SYYYYMM & "','" & EYYYYMM & "','" & BPLID & "','" & ItemGB & "', '" & Section & "','" & ItemName & "', '" & Size & "', '" & OrdNum & "','" & CardCode & "'"
        
    'Procedure 실행(Grid 사용)
    Call oForm01.DataSources.DataTables.Item("ZTEMP01").ExecuteQuery(sQry)
    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("ZTEMP01")
    
    Call GridSetting("Grid01")
    
    oForm01.Freeze False
    
    Set oRecordSet = Nothing
    
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Search_Grid01_Data_Error:
    Set oRecordSet = Nothing
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "Search_Grid01_Data_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Search_Grid02_Data()
    On Error GoTo Search_Grid02_Data_Error
    Dim oRecordSet          As SAPbobsCOM.Recordset
    Dim sQry                As String
    Dim i&
    Dim SYYYYMM             As String   '작번등록년월 시작
    Dim EYYYYMM             As String   '작번등록년월 종료
    Dim BPLID               As String   '사업장
    Dim ItemCode            As String   '품목
    Dim Section             As String   '구분
    Dim ItemName            As String   '품명
    Dim Size                As String   '규격
    Dim OrdNum              As String   '작번
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    oForm01.Freeze True
    
'    SYYYYMM = Trim(oForm01.Items("SYYYYMM").Specific.VALUE)
'    EYYYYMM = Trim(oForm01.Items("EYYYYMM").Specific.VALUE)
    
     For i = 0 To oGrid01.Rows.Count - 1
        If oGrid01.Rows.IsSelected(i) = True Then
            
            '/ Procedure 실행
             sQry = "EXEC PS_PP361_02 '" & Trim(oGrid01.DataTable.GetValue(0, i)) & "'"
'            sQry = "EXEC PS_PP361_02 '" & Trim(oGrid01.DataTable.GetValue(0, i)) & "','" & SYYYYMM & "','" & EYYYYMM & "'"
            oRecordSet.DoQuery sQry

            '// Procedure 실행(Grid 사용)
            oForm01.DataSources.DataTables.Item("ZTEMP02").ExecuteQuery (sQry)
            oGrid02.DataTable = oForm01.DataSources.DataTables.Item("ZTEMP02")
    
'            SumPkQty = Trim(oGrid02.DataTable.GetValue(7, i))
'            SumPkWt = Trim(oGrid02.DataTable.GetValue(8, i))

        End If
     Next i
    
    Call GridSetting("Grid02")
    
'    oForm01.Items("SumPkQty").Specific.VALUE = SumPkQty
'    oForm01.Items("SumPkWt").Specific.VALUE = SumPkWt
'    oForm01.Items("SumNPkQty").Specific.VALUE = 0
'    oForm01.Items("SumNPkWt").Specific.VALUE = SumNPkWt
    
    oForm01.Freeze False
    Set oRecordSet = Nothing

Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Search_Grid02_Data_Error:
        Set oRecordSet = Nothing
        MDC_Com.MDC_GF_Message "Search_Grid02_Data_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Search_Grid03_Data()
    On Error GoTo Search_Grid03_Data_Error
    Dim oRecordSet          As SAPbobsCOM.Recordset
    Dim ErrNum              As Integer
    Dim sQry                As String
    Dim i&, j&
    Dim SYYYYMM             As String   '작번등록년월 시작
    Dim EYYYYMM             As String   '작번등록년월 종료
    Dim BPLID               As String   '사업장
    Dim ItemCode            As String   '품목
    Dim Section             As String   '구분
    Dim ItemName            As String   '품명
    Dim Size                As String   '규격
    Dim OrdNum              As String   '작번
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    oForm01.Freeze True
    
    SYYYYMM = Trim(oForm01.Items("SYYYYMM").Specific.VALUE)
    EYYYYMM = Trim(oForm01.Items("EYYYYMM").Specific.VALUE)
    
     j = 0
     For i = 0 To oGrid02.Rows.Count - 1
        If oGrid02.Rows.IsSelected(i) = True Then
            
            '/ Procedure 실행
            
If oForm01.Items("Selection").Specific.Selected.VALUE = "100" Then '자재비내역
    sQry = "EXEC PS_PP361_03 '" & Trim(oGrid02.DataTable.GetValue(0, i)) & "','" & Trim(oGrid02.DataTable.GetValue(1, i)) & "'"
ElseIf oForm01.Items("Selection").Specific.Selected.VALUE = "200" Then '자체가공비내역
    sQry = "EXEC PS_PP361_05 '" & Trim(oGrid02.DataTable.GetValue(0, i)) & "','" & Trim(oGrid02.DataTable.GetValue(1, i)) & "'"
ElseIf oForm01.Items("Selection").Specific.Selected.VALUE = "300" Then '설계비내역
    sQry = "EXEC PS_PP361_04 '" & Trim(oGrid02.DataTable.GetValue(0, i)) & "','" & Trim(oGrid02.DataTable.GetValue(1, i)) & "'"
ElseIf oForm01.Items("Selection").Specific.Selected.VALUE = "400" Then '외주가공비
    sQry = "EXEC PS_PP361_06 '" & Trim(oGrid02.DataTable.GetValue(0, i)) & "','" & Trim(oGrid02.DataTable.GetValue(1, i)) & "'"
ElseIf oForm01.Items("Selection").Specific.Selected.VALUE = "500" Then '외주제작비
    sQry = "EXEC PS_PP361_07 '" & Trim(oGrid02.DataTable.GetValue(0, i)) & "','" & Trim(oGrid02.DataTable.GetValue(1, i)) & "'"
'ElseIf oForm01.Items("Selection").Specific.Selected.VALUE = "600" Then
'    sQry = "EXEC PS_PP361_08 '" & Trim(oGrid02.DataTable.GetValue(0, i)) & "','" & Trim(oGrid02.DataTable.GetValue(1, i)) & "'"
End If
            oRecordSet.DoQuery sQry

            '// Procedure 실행(Grid 사용)
            oForm01.DataSources.DataTables.Item("ZTEMP03").ExecuteQuery (sQry)
            oGrid03.DataTable = oForm01.DataSources.DataTables.Item("ZTEMP03")
            
            j = j + 1
        End If
     Next i
    
    If j = 0 Then
        oForm01.Freeze False
        oGrid03.DataTable.Clear
        ErrNum = 1
        GoTo Search_Grid03_Data_Error
    End If
    
    Call GridSetting("Grid03")
    
    oForm01.Freeze False
    Set oRecordSet = Nothing

Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Search_Grid03_Data_Error:
 Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "[Sub 작번]의 행을 선택한 뒤 [선택] 중 하나의 내역을 선택하여 주십시오.", "E"
    Else
        MDC_Com.MDC_GF_Message "Search_Grid03_Data_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

'****************************************************************************************************************
'// Grid 꾸며주기
'****************************************************************************************************************
Private Sub GridSetting(ByVal GridNo As String)
On Error GoTo GridSetting_Error
    Dim i           As Integer
    Dim sColsTitle  As String
    Dim sColsLine   As String
    
    oForm01.Freeze True
      
'    oGrid01.Columns.Item(0).LinkedObjectType = lf_BusinessPartner
'    oGrid01.Columns.Item(4).LinkedObjectType = lf_Invoice

    Select Case GridNo
         Case "Grid01"
             oGrid01.Columns.Item(2).LinkedObjectType = lf_Items
            For i = 0 To oGrid01.Columns.Count - 1
                If oGrid01.DataTable.Columns(i).Type = ft_Float Then
                    oGrid01.Columns(i).RightJustified = True
                End If
            Next i
        Case "Grid02"
            For i = 0 To oGrid02.Columns.Count - 1
                If oGrid02.DataTable.Columns(i).Type = ft_Float Then
                    oGrid02.Columns(i).RightJustified = True
                End If
            Next i
'        Case "Grid03"
'            For i = 0 To oGrid02.Columns.Count - 1
'                If oGrid03.DataTable.Columns(i).Type = ft_Float Then
'                    oGrid03.Columns(i).RightJustified = True
'                End If
'            Next i
    End Select
        
    oForm01.Freeze False
    
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
GridSetting_Error:
    MDC_Com.MDC_GF_Message "GridSetting_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub Print_Query()
On Error GoTo Print_Query_Error
    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    Dim Sub_sQry        As String
    
    Dim SYYYYMM         As String
    Dim EYYYYMM         As String
    Dim BPLID           As String
    Dim ItemGB           As String
    Dim Section           As String
    Dim ItemName           As String
    Dim Size           As String
    Dim OrdNum           As String
    
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Call ConnectODBC
       
    '// 조회조건문
    
    SYYYYMM = Trim(oForm01.Items("SYYYYMM").Specific.VALUE)
    EYYYYMM = Trim(oForm01.Items("EYYYYMM").Specific.VALUE)
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    ItemGB = Trim(oForm01.Items("ItemGB").Specific.VALUE)
    Section = Trim(oForm01.Items("Section").Specific.VALUE)
    ItemName = Trim(oForm01.Items("ItemName").Specific.VALUE)
    Size = Trim(oForm01.Items("Size").Specific.VALUE)
    OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE)
    
    If ItemGB = "A" Then ItemGB = "%"
    If Section = "A" Then Section = "%"
    If BPLID = "ALL" Then BPLID = "%"
    If ItemName = "" Then ItemName = "%"
    If Size = "" Then Size = "%"
    If OrdNum = "" Then OrdNum = "%"
    
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_PP361] 작번 List"
    ReportName = "PS_PP361_01.RPT"
    
    '// Formula 수식필드
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
    
    gRpt_Formula(1) = "SYYYYMM"
    gRpt_Formula_Value(1) = SYYYYMM
    
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    '/ Procedure 실행"
    sQry = "EXEC [PS_PP361_001] '" & SYYYYMM & "', '" & EYYYYMM & "', '" & BPLID & "', '" & ItemGB & "', '" & Section & "',  '" & ItemName & "',  '" & Size & "',  '" & OrdNum & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 1
        GoTo Print_Query_Error
    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Query_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Sub Print_Query1()
On Error GoTo Print_Query1_Error
    Dim i               As Integer
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    Dim Sub_sQry        As String
    Dim OrdNum           As String
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    Call ConnectODBC
    
    '// 조회조건문
       
    OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE)
       
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
    WinTitle = "[PS_PP361] 작번 List"
    ReportName = "PS_PP361_02.RPT"
    
    '// Formula 수식필드
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
    
'    gRpt_Formula(1) = "SYYYYMM"
'    gRpt_Formula_Value(1) = SYYYYMM
    
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    '/ Procedure 실행"
    
     For i = 0 To oGrid01.Rows.Count - 1
        If oGrid01.Rows.IsSelected(i) = True Then
        
            '/ Procedure 실행
             sQry = "EXEC PS_PP361_002 '" & Trim(oGrid01.DataTable.GetValue(0, i)) & "'"
            oRecordSet.DoQuery sQry

        End If
     Next i

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query1_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Query1_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub

Private Sub Print_Query2()
On Error GoTo Print_Query2_Error
    Dim i               As Integer
    Dim j               As Integer
    Dim ErrNum          As Integer
    Dim WinTitle        As String
    Dim ReportName      As String
    Dim sQry            As String
    Dim Sub_sQry        As String
    
    Dim OrdNum           As String
    
    Dim oRecordSet      As SAPbobsCOM.Recordset
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
            
    Call ConnectODBC
    
    '// 조회조건문
    
    OrdNum = Trim(oForm01.Items("OrdNum").Specific.VALUE)
       
    '/ Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
If oForm01.Items("Selection").Specific.Selected.VALUE = "100" Then
    WinTitle = "[PS_PP361] 자재비 내역"
    ReportName = "PS_PP361_03.RPT"
ElseIf oForm01.Items("Selection").Specific.Selected.VALUE = "200" Then
    WinTitle = "[PS_PP361] 자체가공비 내역"
    ReportName = "PS_PP361_05.RPT"
ElseIf oForm01.Items("Selection").Specific.Selected.VALUE = "300" Then
    WinTitle = "[PS_PP361] 설계비 내역"
    ReportName = "PS_PP361_04.RPT"
ElseIf oForm01.Items("Selection").Specific.Selected.VALUE = "400" Then
    WinTitle = "[PS_PP361] 외주가공비 내역"
    ReportName = "PS_PP361_06.RPT"
ElseIf oForm01.Items("Selection").Specific.Selected.VALUE = "500" Then
    WinTitle = "[PS_PP361] 외주제작비 내역"
    ReportName = "PS_PP361_07.RPT"
'ElseIf oForm01.Items("Selection").Specific.Selected.VALUE = "600" Then
'    WinTitle = "[PS_PP361] 공정내역"
'    ReportName = "PS_PP361_08.RPT"
End If

    '// Formula 수식필드
    ReDim gRpt_Formula(2)
    ReDim gRpt_Formula_Value(2)
    
'    gRpt_Formula(1) = "SYYYYMM"
'    gRpt_Formula_Value(1) = SYYYYMM
    
    '// SubReport
    ReDim gRpt_SRptSqry(1)
    ReDim gRpt_SRptName(1)
    
    ReDim gRpt_SFormula(1, 1)
    ReDim gRpt_SFormula_Value(1, 1)
    
    gRpt_SFormula(1, 1) = ""
    gRpt_SFormula_Value(1, 1) = ""
    
    '/ Procedure 실행"
    
'     j = 0
'     For i = 0 To oGrid03.Rows.Count - 1
'        If oGrid02.Rows.IsSelected(i) = True Then
            
            '/ Procedure 실행
            
If oForm01.Items("Selection").Specific.Selected.VALUE = "100" Then
    sQry = "EXEC PS_PP361_003 '" & Trim(oGrid02.DataTable.GetValue(0, i)) & "','" & Trim(oGrid02.DataTable.GetValue(1, i)) & "'"
ElseIf oForm01.Items("Selection").Specific.Selected.VALUE = "200" Then
    sQry = "EXEC PS_PP361_005 '" & Trim(oGrid02.DataTable.GetValue(0, i)) & "','" & Trim(oGrid02.DataTable.GetValue(1, i)) & "'"
ElseIf oForm01.Items("Selection").Specific.Selected.VALUE = "300" Then
    sQry = "EXEC PS_PP361_004 '" & Trim(oGrid02.DataTable.GetValue(0, i)) & "','" & Trim(oGrid02.DataTable.GetValue(1, i)) & "'"
ElseIf oForm01.Items("Selection").Specific.Selected.VALUE = "400" Then
    sQry = "EXEC PS_PP361_006 '" & Trim(oGrid02.DataTable.GetValue(0, i)) & "','" & Trim(oGrid02.DataTable.GetValue(1, i)) & "'"
ElseIf oForm01.Items("Selection").Specific.Selected.VALUE = "500" Then
    sQry = "EXEC PS_PP361_007 '" & Trim(oGrid02.DataTable.GetValue(0, i)) & "','" & Trim(oGrid02.DataTable.GetValue(1, i)) & "'"
'ElseIf oForm01.Items("Selection").Specific.Selected.VALUE = "600" Then
'    sQry = "EXEC PS_PP361_008 '" & Trim(oGrid02.DataTable.GetValue(0, i)) & "','" & Trim(oGrid02.DataTable.GetValue(1, i)) & "'"
End If
            oRecordSet.DoQuery sQry

            '// Procedure 실행(Grid 사용)
'            oForm01.DataSources.DataTables.Item("ZTEMP03").ExecuteQuery (sQry)
'            oGrid03.DataTable = oForm01.DataSources.DataTables.Item("ZTEMP03")
'
'            j = j + 1
'        End If
'     Next i

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "", "N", "V") = False Then
    End If
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query2_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "출력할 데이터가 없습니다. 확인해 주세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "Print_Query2_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub
