VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM041"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_QM041.cls
'//  Module         : QM
'//  Description    : 검사성적서출력(신양식)
'//  FormType       : PS_QM041
'//  Create Date    : 2012.06.28
'//  Modified Date  :
'//  Creator        : NGY
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_QM041L       As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm(Optional ByVal oYM01 As String, Optional ByVal oDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM041.srf")
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
    
    oFormUniqueID01 = "PS_QM041_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
'    oForm01.Mode = fm_ADD_MODE
    oForm01.Mode = fm_FIND_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    'oForm01.DataBrowser.BrowseBy = "Code"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    
    oForm01.Items("DocEntry").Specific.VALUE = oDocEntry01
    oForm01.Items("YYYYMM").Specific.VALUE = oYM01
    
    Set oCombo = oForm01.Items("Gubun").Specific
    If Trim(oDocEntry01) <> "" Then
        oCombo.Select "2", psk_ByValue
    Else
        oCombo.Select "1", psk_ByValue
    End If
    
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
    Set oDS_PS_QM041L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_Single
'    oMat01.AutoResizeColumns

    Call oForm01.DataSources.UserDataSources.Add("BPLId", dt_SHORT_TEXT, 10)
    oForm01.Items("BPLId").Specific.DataBind.SetBound True, "", "BPLId"
    
    
    Call oForm01.DataSources.UserDataSources.Add("YYYYMM", dt_SHORT_TEXT, 7)
    oForm01.Items("YYYYMM").Specific.DataBind.SetBound True, "", "YYYYMM"
    oForm01.Items("YYYYMM").Specific.VALUE = Format(Now, "YYYY-MM")
'    oForm01.DataSources.UserDataSources.Item("YYYYMM").VALUE = Format(Now, "YYYY-MM")
    
    
    '한글, 영문 출력 구분
    'Call oForm01.DataSources.UserDataSources.Add("Print_Div", dt_SHORT_TEXT, 10)
    'oForm01.Items("Print_Div").Specific.DataBind.SetBound True, "", "Print_Div"
    
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
    sQry = "SELECT BPLId, BPLName From [OBPL] Where BPLId in ('1', '2') order by 1"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        'oMat01.Columns("BPLId").ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '//기본사업장SETTING
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    Call oForm01.Items("Gubun").Specific.ValidValues.Add("1", "Packing기준")
    Call oForm01.Items("Gubun").Specific.ValidValues.Add("2", "납품기준")
    oForm01.Items("Gubun").Specific.Select "0", psk_Index
    
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
                If pval.ItemUID = "1" Then
                    If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
'                        If HeaderSpaceLineDel = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
'                        If MatrixSpaceLineDel = False Then
'                            BubbleEvent = False
'                            Exit Sub
'                        End If
                    End If
                ElseIf pval.ItemUID = "Search" Then '조회버튼 클릭시
                    If HeaderSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    Else
                        Call Search_Matrix_Data
                    End If
                ElseIf pval.ItemUID = "Print" Then '출력버튼 클릭시
                    Call Print_Query
                End If
                
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    '//헤더
'                    If pval.ItemUID = "LotNo" Then
'                        If oForm01.Items("LotNo").Specific.Value = "" Then
'                            Sbo_Application.ActivateMenuItem ("7425")
'                            BubbleEvent = False
'                        End If
'                    End If
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
                If pval.ItemUID = "Mat01" Then
                    If pval.Row = 0 Then
                        If pval.ColUID = "Check" Then
                            For i = 0 To oMat01.VisualRowCount - 1
                                
                                If Trim(oDS_PS_QM041L.GetValue("U_ColReg01", i)) = "N" Then
                                    oDS_PS_QM041L.setValue "U_ColReg01", i, "Y"
                                Else
                                    oDS_PS_QM041L.setValue "U_ColReg01", i, "N"
                                End If
                            Next i
                            oMat01.LoadFromDataSource
                            'oMat01.FlushToDataSource
                        Else
                            oMat01.Columns(pval.ColUID).TitleObject.Sortable = True
                            oMat01.FlushToDataSource
                        End If
                    End If
                End If
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
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    '//헤더
'                    If pval.ItemUID = "LotNo" Or pval.ItemUID = "ChemC_Fe" Or pval.ItemUID = "ChemC_P" Then
'                        FlushToItemValue pval.ItemUID
'                    End If
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
    Dim oCombo          As SAPbouiCOM.ComboBox
    
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
                oForm01.Items("DocEntry").Specific.VALUE = ""
                Set oCombo = oForm01.Items("Gubun").Specific
                oCombo.Select "1", psk_ByValue
                
                Set oCombo = oForm01.Items("BPLId").Specific
                oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
                
                oForm01.Items("YYYYMM").Specific.VALUE = Format(Now, "YYYY-MM")
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
    Dim i&, j&, DocNum$, LineId$, Counts$, Cnt$
    Dim ErrNum          As Integer
    Dim sQry            As String
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    Dim oRecordSet    As SAPbobsCOM.Recordset
        
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    '--------------------------------------------------------------
    'Header--------------------------------------------------------
    Select Case oUID
'        Case "LotNo"
        
    End Select
        
    '--------------------------------------------------------------
    'Line----------------------------------------------------------
    If oUID = "Mat01" Then
        Select Case oCol
'            Case "PP070No"
        
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
      Case oForm01.Items("YYYYMM").Specific.VALUE = ""
           ErrNum = 1
           GoTo HeaderSpaceLineDel_Error
      Case Len(oForm01.Items("YYYYMM").Specific.VALUE) <> 7
           ErrNum = 2
           GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        Sbo_Application.SetStatusBarMessage "조회년월은 필수입니다. 입력하여 주십시오.", bmt_Short, True
    ElseIf ErrNum = 2 Then
        Sbo_Application.SetStatusBarMessage "조회년월의 자리수(YYYY-MM)를 확인하여 주십시오.", bmt_Short, True
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Sub Search_Matrix_Data()
    On Error GoTo Search_Matrix_Data_Error
    Dim oRecordSet          As SAPbobsCOM.Recordset
    Dim sQry                As String
    
    Dim i                   As Integer
    Dim j                   As Integer
    Dim Cnt                 As Integer
    Dim BPLID$, YYYYMM$
    Dim Gubun As String
    Dim DocEntry As String
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    oForm01.Freeze True
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    YYYYMM = Trim(oForm01.Items("YYYYMM").Specific.VALUE)
    Gubun = Trim(oForm01.Items("Gubun").Specific.VALUE)
    DocEntry = Trim(oForm01.Items("DocEntry").Specific.VALUE)
    
    '/ Procedure 실행
    sQry = "EXEC PS_QM041_01 '" & BPLID & "', '" & YYYYMM & "', '" & Gubun & "', '" & DocEntry & "'"
    oRecordSet.DoQuery sQry
    
     'Matrix에 출력===========================================================================================
     '// Line 초기화
     Cnt = oDS_PS_QM041L.Size
     If Cnt > 0 Then
         For j = 0 To Cnt - 1
             oDS_PS_QM041L.RemoveRecord oDS_PS_QM041L.Size - 1
         Next j
         If Cnt = 1 Then
             oDS_PS_QM041L.Clear
         End If
     End If
     oMat01.LoadFromDataSource
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '//Matrix에 Data 뿌려준다
    j = 1
     Do Until oRecordSet.EOF

         If oDS_PS_QM041L.Size < j Then
             oDS_PS_QM041L.InsertRecord j - 1 '라인추가
         End If
         '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
         oDS_PS_QM041L.setValue "U_LineNum", j - 1, j
         If Gubun = "2" Then
            oDS_PS_QM041L.setValue "U_ColReg01", j - 1, "Y"
         Else
            oDS_PS_QM041L.setValue "U_ColReg01", j - 1, "N"
         End If
         oDS_PS_QM041L.setValue "U_ColReg02", j - 1, oRecordSet.Fields("U_PackNo").VALUE
         oDS_PS_QM041L.setValue "U_ColReg03", j - 1, oRecordSet.Fields("U_ItemCode").VALUE
         oDS_PS_QM041L.setValue "U_ColReg04", j - 1, oRecordSet.Fields("U_ItemName").VALUE
         oDS_PS_QM041L.setValue "U_ColReg05", j - 1, oRecordSet.Fields("U_CardCode").VALUE
         oDS_PS_QM041L.setValue "U_ColReg06", j - 1, oRecordSet.Fields("U_CardName").VALUE
         oDS_PS_QM041L.setValue "U_ColReg07", j - 1, oRecordSet.Fields("Type").VALUE

         j = j + 1
         oRecordSet.MoveNext
     Loop
     oMat01.LoadFromDataSource
             ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    oForm01.Freeze False
    Set oRecordSet = Nothing
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Search_Matrix_Data_Error:
    Set oRecordSet = Nothing
    MDC_Com.MDC_GF_Message "Search_Matrix_Data_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

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
    
    Dim BPLID           As String
    Dim PackingNo       As String
    Dim ItemCode        As String
    Dim CardCode        As String
    Dim Chk             As String
    Dim BakeTestChk As String 'BakeTest 측정치 추가 품목 체크를 위한 변수(2012.11.26 송명규 추가)
    Dim PackNo As String '패킹번호(2012.11.26 송명규 추가)
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, False)

    Call ConnectODBC
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    
    '// 임시테이블에 check된항목저장
    sQry = "Delete [Z_PS_QM040] WHERE BPLId = '" & BPLID & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_QM041L.GetValue("U_ColReg01", i)) = "Y" Then
            sQry = "Insert [Z_PS_QM040] values ('" & BPLID & "', '" & Trim(oDS_PS_QM041L.GetValue("U_ColReg02", i)) & "')"
            oRecordSet01.DoQuery sQry
            
            CardCode = Trim(oDS_PS_QM041L.GetValue("U_ColReg05", i)) '// (주)TSD '12440' 찿기위해 MOVE
            ItemCode = Trim(oDS_PS_QM041L.GetValue("U_ColReg03", i))
            
        End If
    Next
    
    '// B/G타입  체크
    If MDC_PS_Common.GetValue("SELECT count(*) FROM [@PS_PP090H] a inner join [@PS_PP090L] b on a.DocEntry = b.DocEntry INNER JOIN [Z_PS_QM040] z on a.U_BPLId = z.BPLId and a.U_PackNo = z.PackNo left  join [OITM] c on b.U_ItemCode = c.ItemCode WHERE z.BPLId = '" & BPLID & "' and c.U_ItemType in ('16','17','19')") > 0 Then
       Chk = "Y"
    Else
       Chk = "N"
    End If
    
'    sQry = "            SELECT      ISNULL(T0.U_M_BT_Val, 0) AS [BT_Val]"
'    sQry = sQry & "  FROM       [@PS_PP090H] a"
'    sQry = sQry & "                 INNER JOIN"
'    sQry = sQry & "                 [@PS_PP090L] b"
'    sQry = sQry & "                     ON a.DocEntry = b.DocEntry"
'    sQry = sQry & "                 INNER JOIN"
'    sQry = sQry & "                 [Z_PS_QM040] z"
'    sQry = sQry & "                     ON a.U_BPLId = z.BPLId"
'    sQry = sQry & "                     AND a.U_PackNo = z.PackNo"
'    sQry = sQry & "                 LEFT JOIN"
'    sQry = sQry & "                 [@PS_QM020H] AS T0"
'    sQry = sQry & "                     ON T0.U_OrdNum = b.U_LotNo"
'    sQry = sQry & "  WHERE      z.BPLId = '" & BPLId & "'"
'    sQry = sQry & "                 AND b.U_PackNo = '" & PackNo & "'"
            
    
'    BakeTestChk = MDC_PS_Common.GetValue(sQry)
    
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
       
       Sub_sQry = "EXEC [PS_QM041_04] '" & BPLID & "'"
       gRpt_SRptSqry(1) = Sub_sQry
       gRpt_SRptName(1) = "PS_QM040_SUB_01"
       
       Sub_sQry = "EXEC [PS_QM041_05] '" & BPLID & "'"
       gRpt_SRptSqry(2) = Sub_sQry
       gRpt_SRptName(2) = "PS_QM040_SUB_02"
       
       Sub_sQry = "EXEC [PS_QM041_03] '" & BPLID & "'"
       gRpt_SRptSqry(3) = Sub_sQry
       gRpt_SRptName(3) = "PS_QM040_SUB_03"
       
    Else
    
       Sub_sQry = "EXEC [PS_QM041_03] '" & BPLID & "'"
       gRpt_SRptSqry(1) = Sub_sQry
       gRpt_SRptName(1) = "PS_QM040_SUB_01"
       
       Sub_sQry = "EXEC [PS_QM041_05] '" & BPLID & "'"
       gRpt_SRptSqry(2) = Sub_sQry
       gRpt_SRptName(2) = "PS_QM040_SUB_02"
       
       Sub_sQry = "EXEC [PS_QM041_03] '" & BPLID & "'"
       gRpt_SRptSqry(3) = Sub_sQry
       gRpt_SRptName(3) = "PS_QM040_SUB_03"
       
    End If
    
    '/ Procedure 실행
    'sQry = "EXEC [PS_QM041_02] '" & PackingNo & "', '" & ItemCode & "','" & CardCode & "'"
    
    sQry = "EXEC [PS_QM041_02] '" & BPLID & "'"
    Call oRecordSet.DoQuery(sQry)
'    If oRecordSet.RecordCount = 0 Then
'        ErrNum = 1
'        GoTo Print_Query_Error
'    End If

    '/ Action (sub_query가 있을때는 'Y'로...)/
    If MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "", "N", "V") = False Then
    End If
    
    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing
    
    Set oRecordSet = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Print_Query_Error:

    ProgBar01.VALUE = 100
    ProgBar01.Stop
    Set ProgBar01 = Nothing

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
