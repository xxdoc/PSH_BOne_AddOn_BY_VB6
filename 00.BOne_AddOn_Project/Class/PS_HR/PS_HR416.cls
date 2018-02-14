VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_HR416"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_HR416.cls
'//  Module         : HR
'//  Description    : 전문직 평가
'//  FormType       : PS_HR416
'//  Create Date    : 2012.10.05
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Public oGrid01              As SAPbouiCOM.Grid
Public oGrid02              As SAPbouiCOM.Grid
Public oGrid03              As SAPbouiCOM.Grid
Private oDS_PS_HR416L       As SAPbouiCOM.DBDataSource    '

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
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_HR416.srf")
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
    
    oFormUniqueID01 = "PS_HR416_" & GetTotalFormsCount
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

    oForm01.Items("Year").Click ct_Regular
    
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
    Set oDS_PS_HR416L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    Set oGrid01 = oForm01.Items("Grid01").Specific
    Set oGrid02 = oForm01.Items("Grid02").Specific
    Set oGrid03 = oForm01.Items("Grid03").Specific
    
    oForm01.DataSources.DataTables.Add ("ZTEMP1")
    oForm01.DataSources.DataTables.Add ("ZTEMP2")
    oForm01.DataSources.DataTables.Add ("ZTEMP3")
    
    oGrid01.DataTable = oForm01.DataSources.DataTables.Item("ZTEMP1")
    oGrid02.DataTable = oForm01.DataSources.DataTables.Item("ZTEMP2")
    oGrid03.DataTable = oForm01.DataSources.DataTables.Item("ZTEMP3")
    
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
    sQry = "SELECT BPLId, BPLName From [OBPL]"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    Set oCombo = oForm01.Items("Number").Specific
    oCombo.ValidValues.Add "1", "1차평가"
    oCombo.ValidValues.Add "2", "2차평가"
    oCombo.Select 0, psk_Index
    
    Set oCombo = oForm01.Items("Evaluate").Specific
    oCombo.ValidValues.Add "1", "1차평가자"
    oCombo.ValidValues.Add "2", "2차평가자"
    oCombo.ValidValues.Add "3", "종합평가자"
    oCombo.Select 0, psk_Index
    
    '//사용여부
    oMat01.Columns("Grade").ValidValues.Add "", "선택"
    oMat01.Columns("Grade").ValidValues.Add "S", "S"
    oMat01.Columns("Grade").ValidValues.Add "A", "A"
    oMat01.Columns("Grade").ValidValues.Add "B", "B"
    oMat01.Columns("Grade").ValidValues.Add "C", "C"
    oMat01.Columns("Grade").ValidValues.Add "D", "D"
    
    '//아이디별 사업장 세팅
    Set oCombo = oForm01.Items("BPLId").Specific
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
    Dim vReturnValue$
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Dim Grade As String
    Dim Year As String
    Dim RateCode As String
    
    Dim BPLID   As String
    Dim Number  As String
    Dim MSTCOD  As String
    Dim FULLNAME As String
    
    Dim Svalue           As Double '점수합계
    
    Dim Complete        As String
    
    Dim ChildForm01 As Variant
    Set ChildForm01 = New PS_HR403
    
    Dim oRecordSet01    As SAPbobsCOM.Recordset
    Dim sQry$
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
                
                If pval.ItemUID = "Btn01" Then '조회버튼
                    BPLID = oForm01.Items("BPLId").Specific.VALUE
                    Year = oForm01.Items("Year").Specific.VALUE
                    Number = oForm01.Items("Number").Specific.VALUE
                    MSTCOD = oForm01.Items("MSTCOD").Specific.VALUE
                    FULLNAME = oForm01.Items("FULLNAME").Specific.VALUE
                    
                    '//평가서약을 했는지...
                    sQry = "Select Count(*) From Z_PS_HR403 Where BPLId = '" & BPLID & "' and Year = '" & Year & "' And Number = '" & Number & "' and MSTCOD = '" & MSTCOD & "'"
                    oRecordSet01.DoQuery sQry
                    
                    If oRecordSet01.Fields(0).VALUE <= 0 Then
                        '//평가서약 화면
                        Call ChildForm01.LoadForm(BPLID, Year, Number, MSTCOD, FULLNAME)
                    Else
                        If PasswordChk(pval) = False Then
                            Sbo_Application.SetStatusBarMessage "패스워드가 틀렸습니다. 확인바랍니다.", bmt_Short, True
                            oForm01.Items("PassWd").Click ct_Regular
                        Else
                            Call Search_Grid_Data
                        End If
                    End If
            
            
            
                    
                
                ElseIf pval.ItemUID = "Btn02" Then '저장버튼
                    Complete = oForm01.Items("Complete").Specific.VALUE
    
                    If Complete = "평가완료처리" Then
                        Sbo_Application.SetStatusBarMessage "이미 평가완료처리를 하였습니다.. 확인바랍니다.", bmt_Short, True
                        Call Search_Grid_Data '새로고침
                    Else
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
                        Else
                            Call Search_Grid_Data '새로고침
                        End If
                    End If
    
                    
                ElseIf pval.ItemUID = "Btn04" Then '완료처리
                    Complete = oForm01.Items("Complete").Specific.VALUE
    
                    If Complete = "평가완료처리" Then
                        Sbo_Application.SetStatusBarMessage "이미 평가완료처리를 하였습니다.. 확인바랍니다.", bmt_Short, True
                    Else
                        Call Complete_Data
                    End If
                   
                ElseIf pval.ItemUID = "Btn05" Then '완료 취소처리
                    Complete = oForm01.Items("Complete").Specific.VALUE
    
                    If Complete = "평가완료미처리" Then
                        Sbo_Application.SetStatusBarMessage "평가완료 취소할 자료가 없습니다.. 확인바랍니다.", bmt_Short, True
                    Else
                        Call InComplete_Data
                    End If
                End If
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    '//헤더
                    If pval.ItemUID = "Year" Then
                        If oForm01.Items("Year").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "MSTCOD" Then
                        If oForm01.Items("MSTCOD").Specific.VALUE = "" Then
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
                If pval.ItemUID = "Mat01" Then
                    If pval.ColUID = "Grade" Then
                        Year = oForm01.Items("Year").Specific.VALUE
                        RateCode = oMat01.Columns("RateCode").Cells(pval.Row).Specific.VALUE
                        Grade = oMat01.Columns("Grade").Cells(pval.Row).Specific.VALUE
                        
                        
                        sQry = "SELECT s = b.U_SValue, a = b.U_AValue, b= b.U_BValue, c = U_CValue, d = U_DValue From [@PS_HR400H] a Inner Join [@PS_HR400L] b On a.Code = b.Code "
                        sQry = sQry & "Where a.U_Year = '" & Year & "' And b.U_RateCode = '" & RateCode & "'"
                        oRecordSet01.DoQuery sQry
                        
                        oForm01.Freeze True

                        oMat01.FlushToDataSource
                        
                        Select Case Grade
                            Case "S"
                                Call oDS_PS_HR416L.setValue("U_ColQty02", pval.Row - 1, Trim(oRecordSet01.Fields(0).VALUE))
                            Case "A"
                                Call oDS_PS_HR416L.setValue("U_ColQty02", pval.Row - 1, Trim(oRecordSet01.Fields(1).VALUE))
                            Case "B"
                                Call oDS_PS_HR416L.setValue("U_ColQty02", pval.Row - 1, Trim(oRecordSet01.Fields(2).VALUE))
                            Case "C"
                                Call oDS_PS_HR416L.setValue("U_ColQty02", pval.Row - 1, Trim(oRecordSet01.Fields(3).VALUE))
                            Case "D"
                                Call oDS_PS_HR416L.setValue("U_ColQty02", pval.Row - 1, Trim(oRecordSet01.Fields(4).VALUE))
                            Case Else
                                Call oDS_PS_HR416L.setValue("U_ColQty02", pval.Row - 1, 0)
                        End Select
                        
                        For i = 0 To oMat01.VisualRowCount - 1
                            oDS_PS_HR416L.Offset = i
                    
                            Svalue = Svalue + Trim(oDS_PS_HR416L.GetValue("U_ColQty02", i))
                    
                        Next i

                        oForm01.Items("Svalue").Specific.VALUE = Svalue
                        
             
             ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
     
                        
                        oMat01.LoadFromDataSource
                        oForm01.Freeze False
                    End If
                End If
            Case et_CLICK: '//6
                If pval.ItemUID = "Grid01" Then
                    Call Search_Matrix_Data
                ElseIf pval.ItemUID = "Mat01" Then
                    
                End If
            Case et_DOUBLE_CLICK: '//7
                If pval.ItemUID = "Mat01" Then
                    Call Search_Grid_Data1(pval.Row)
                End If
                
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    '//헤더
                    If (pval.ItemUID = "Year") Then
                        If Trim(oForm01.Items("Year").Specific.VALUE) <> "" Then
                            
                            sQry = "select U_Number from [@PS_HR410H] a"
                            sQry = sQry + " Where Isnull(a.U_OpenYN,'N') = 'Y' and isnull(a.U_CloseYN,'N') = 'N' "
                            sQry = sQry + " and a.U_BPLId = '" & Trim(oForm01.Items("BPLId").Specific.VALUE) & "' "
                            sQry = sQry + " and a.U_Year = '" & Trim(oForm01.Items("Year").Specific.VALUE) & "' "
                            oRecordSet01.DoQuery sQry
                    
                            If Trim(oRecordSet01.Fields(0).VALUE) <> "" Then
                                Call oForm01.Items("Number").Specific.Select(Trim(oRecordSet01.Fields(0).VALUE), psk_ByValue)
                            End If
                        End If
                    End If
                                
                    If pval.ItemUID = "MSTCOD" Then
                        FlushToItemValue pval.ItemUID
                    End If
                    '//라인
                    If pval.ItemUID = "Mat01" And (pval.ColUID = "Grade") Then
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
'                    If Insert_oInventoryGenExit(2) = False Then
'                        BubbleEvent = False
'                        Exit Sub
'                    Else
'                        Sbo_Application.MessageBox ("선택취소 작업이 완료되었습니다.")
'                    End If
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
'                    oDS_PS_HR416H.RemoveRecord oDS_PS_HR416H.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
'                    oMat01.Clear
'                    oMat01.LoadFromDataSource
'                End If
            
            Case "1281": '찾기
'                FormItemEnabled
'                oForm01.Items("DocNum").Click ct_Regular
                
            Case "1282": '추가
'                Call FormItemEnabled
'                Call FormClear
'                oDS_PS_HR416H.setValue "U_RegiDate", 0, Format(Now, "YYYYMMDD")
'                Call Matrix_AddRow(1, 0, True) 'oMat01
                
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
'                FormItemEnabled
'                If oMat01.VisualRowCount > 0 Then
'                    If oMat01.Columns("PP070No").Cells(oMat01.VisualRowCount).Specific.Value <> "" Then
'                        If oDS_PS_HR416H.GetValue("Status", 0) = "O" Then
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
'        oForm01.Items("BPLId").Enabled = True
'        oForm01.Items("CntcCode").Enabled = True
'        oForm01.Items("CardCode").Enabled = True
'        oForm01.Items("DeliArea").Enabled = True
'        oForm01.Items("CarNo").Enabled = True
'        oForm01.Items("TransCom").Enabled = True
'        oForm01.Items("Fee").Enabled = True
'        oForm01.Items("RegiDate").Enabled = True
'        oMat01.Columns("PP070No").Editable = True
'
    ElseIf oForm01.Mode = fm_ADD_MODE Then
'        oForm01.Items("DocNum").Enabled = False
'        oForm01.Items("BPLId").Enabled = True
'        oForm01.Items("CntcCode").Enabled = True
'        oForm01.Items("CardCode").Enabled = True
'        oForm01.Items("DeliArea").Enabled = True
'        oForm01.Items("CarNo").Enabled = True
'        oForm01.Items("TransCom").Enabled = True
'        oForm01.Items("Fee").Enabled = True
'        oForm01.Items("RegiDate").Enabled = True
'        oMat01.Columns("PP070No").Editable = True
        
    ElseIf oForm01.Mode = fm_OK_MODE Then
'        If oForm01.Items("Canceled").Specific.VALUE = "Y" Then
'            oForm01.Items("Comments").Click ct_Regular
'            oForm01.Items("DocNum").Enabled = False
'            oForm01.Items("BPLId").Enabled = False
'            oForm01.Items("CntcCode").Enabled = False
'            oForm01.Items("CardCode").Enabled = False
'            oForm01.Items("DeliArea").Enabled = False
'            oForm01.Items("CarNo").Enabled = False
'            oForm01.Items("TransCom").Enabled = False
'            oForm01.Items("Fee").Enabled = False
'            oForm01.Items("RegiDate").Enabled = False
'            oMat01.Columns("PP070No").Editable = False
'        Else
'            oForm01.Items("DocNum").Enabled = True
'            oForm01.Items("BPLId").Enabled = True
'            oForm01.Items("CntcCode").Enabled = True
'            oForm01.Items("CardCode").Enabled = True
'            oForm01.Items("DeliArea").Enabled = True
'            oForm01.Items("CarNo").Enabled = True
'            oForm01.Items("TransCom").Enabled = True
'            oForm01.Items("Fee").Enabled = True
'            oForm01.Items("RegiDate").Enabled = True
'            oMat01.Columns("PP070No").Editable = False
'        End If
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
'                oDS_PS_HR416L.InsertRecord (oRow)
'            End If
'            '수입내역
'            oDS_PS_HR416L.Offset = oRow
'            oDS_PS_HR416L.setValue "LineId", oRow, oRow + 1
'            oDS_PS_HR416L.setValue "U_PP070No", oRow, ""
'            oDS_PS_HR416L.setValue "U_ItemCode", oRow, ""
'            oDS_PS_HR416L.setValue "U_ItemName", oRow, ""
'            oDS_PS_HR416L.setValue "U_Size", oRow, ""
'            oDS_PS_HR416L.setValue "U_Mark", oRow, ""
'            oDS_PS_HR416L.setValue "U_Qty", oRow, ""
'            oDS_PS_HR416L.setValue "U_Weight", oRow, ""
'            oDS_PS_HR416L.setValue "U_DocDate", oRow, ""
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
    
    Dim BPLID               As String
    Dim Year                As String
    Dim Number              As String
    Dim MSTCOD              As String
    Dim FULLNAME            As String
    
    Dim sQry            As String
    Dim oRecordSet    As SAPbobsCOM.Recordset
    
    Dim ChildForm01 As Variant
    Set ChildForm01 = New PS_HR403
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    '--------------------------------------------------------------
    'Header--------------------------------------------------------
    Select Case oUID
        Case "MSTCOD"
            BPLID = oForm01.Items("BPLId").Specific.VALUE
            Year = oForm01.Items("Year").Specific.VALUE
            Number = oForm01.Items("Number").Specific.VALUE
            MSTCOD = oForm01.Items("MSTCOD").Specific.VALUE
            
            sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" & Trim(oForm01.Items("MSTCOD").Specific.String) & "'"
            oRecordSet.DoQuery sQry
            oForm01.Items("FULLNAME").Specific.String = Trim(oRecordSet.Fields(0).VALUE)
            
            FULLNAME = Trim(oRecordSet.Fields(0).VALUE)
            '//평가서약을 했는지...
            sQry = "Select Count(*) From Z_PS_HR403 Where BPLId = '" & BPLID & "' and Year = '" & Year & "' And Number = '" & Number & "' and MSTCOD = '" & MSTCOD & "'"
            oRecordSet.DoQuery sQry
            
            If oRecordSet.Fields(0).VALUE <= 0 Then
                '//평가서약 화면
                Call ChildForm01.LoadForm(BPLID, Year, Number, MSTCOD, FULLNAME)
            End If
            
    End Select
        
    '--------------------------------------------------------------
    'Line----------------------------------------------------------
    If oUID = "Mat01" Then
        Select Case oCol
            Case "GRADE"
                oMat01.FlushToDataSource
                
'                For i = 0 To oMat01.VisualRowCount - 1
'                    oDS_PS_HR416H.Offset = i
'
'                    SumNPkQty = SumNPkQty + Trim(oDS_PS_HR416H.GetValue("U_NPkQty", i))
'                    SumNPkWt = SumNPkWt + Trim(oDS_PS_HR416H.GetValue("U_NPkWt", i))
'                Next i
'
'                oForm01.Items("SumNPkQty").Specific.VALUE = SumNPkQty
'                oForm01.Items("SumNPkWt").Specific.VALUE = SumNPkWt
                
                '--------------------------------------------------------------------------------------------
'                If oRow = oMat01.RowCount And Trim(oDS_PS_HR416L.GetValue("U_PP070No", oRow - 1)) <> "" Then
'                '// 다음 라인 추가
'                    Call Matrix_AddRow(1, 0, False)
'                    oMat01.Columns("PP070No").Cells(oRow).Click ct_Regular
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
        Case Trim(oForm01.Items("Year").Specific.VALUE) = ""
            oForm01.Items("Year").Click ct_Regular
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("MSTCOD").Specific.VALUE) = ""
            oForm01.Items("MSTCOD").Click ct_Regular
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oForm01.Items("EmpNo").Specific.VALUE) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "평가년도는 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "평가자사번은 필수사항입니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "피평가자가 선택되지 않았습니다. 확인하여 주십시오.", "E"
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
    Dim j               As Integer
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0
    
    oMat01.FlushToDataSource

    '// 라인
    If oMat01.VisualRowCount = 0 Then '// MAT01에 값이 있는지 확인 (ErrorNumber : 1)
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//평가점수가 완료되지 않으면 저장할 수 없다.
    '//oMat01.RowCount - 2가 된다..반드시 들어 가야 하는 필수값을 확인한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    If oMat01.VisualRowCount > 0 Then
        For i = 0 To oMat01.VisualRowCount - 1      '// Mat1에 입력값이 올바르게 들어갔는지 확인 (ErrorNumber : 2)
            If oDS_PS_HR416L.GetValue("U_ColQty02", i) = 0 Then
                ErrNum = 2
                GoTo MatrixSpaceLineDel_Error
                
            End If

        Next i
    End If
'
'    '// 체크된 라인Count
'    If j = 0 Then
'        ErrNum = 4
'        GoTo MatrixSpaceLineDel_Error
'    End If
        
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
    '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'    If oMat01.VisualRowCount > 0 Then
'        oDS_PS_HR416L.RemoveRecord oDS_PS_HR416L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
'    End If
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
        MDC_Com.MDC_GF_Message "평가가 다되지 않았습니다.확인하여 주십시오.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Sub Form_Resize(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Form_Resize_Error
    Dim i   As Integer
    
    
    oForm01.Items("Grid01").Top = 79
    oForm01.Items("Grid01").Height = (oForm01.Height / 2) - 110
    oForm01.Items("Grid01").Left = 10
    oForm01.Items("Grid01").Width = oForm01.Width / 2
    
    oForm01.Items("Mat01").Top = 79
    oForm01.Items("Mat01").Height = (oForm01.Height / 2) - 110
    oForm01.Items("Mat01").Left = (oForm01.Width / 2) + 30
    oForm01.Items("Mat01").Width = (oForm01.Width / 2) - 30
    
    
    oForm01.Items("Btn02").Top = oForm01.Items("Grid01").Height + 79
    oForm01.Items("2").Top = oForm01.Items("Grid01").Height + 79
    oForm01.Items("Btn03").Top = oForm01.Items("Grid01").Height + 79
    oForm01.Items("Btn04").Top = oForm01.Items("Grid01").Height + 79
    oForm01.Items("Btn05").Top = oForm01.Items("Grid01").Height + oForm01.Items("Btn04").Height + 79
    
    '합계
    oForm01.Items("t3").Top = oForm01.Items("Btn02").Top
    oForm01.Items("t3").Left = oForm01.Items("Btn01").Left
    oForm01.Items("Svalue").Top = oForm01.Items("Btn02").Top
    oForm01.Items("Svalue").Left = oForm01.Items("Btn01").Left + 80
    
    oForm01.Items("t2").Top = 61
    oForm01.Items("t2").Left = (oForm01.Width / 2) + 30

    oForm01.Items("t4").Top = oForm01.Items("Btn02").Top + 20
    
    oForm01.Items("Grid02").Top = oForm01.Items("Grid01").Height + 120
    oForm01.Items("Grid02").Height = 126
    oForm01.Items("Grid02").Left = 10
    oForm01.Items("Grid02").Width = oForm01.Width - 21
    
    oForm01.Items("t5").Top = oForm01.Items("Grid02").Top + 126
    
    oForm01.Items("Grid03").Top = oForm01.Items("t5").Top + 20
    oForm01.Items("Grid03").Height = 126
    oForm01.Items("Grid03").Left = 10
    oForm01.Items("Grid03").Width = oForm01.Width - 21
    

'    If oGrid01.Rows.Count > 0 Then
'        oGrid01.AutoResizeColumns
'    End If
    oMat01.AutoResizeColumns
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Form_Resize_Error:
    Sbo_Application.SetStatusBarMessage "Form_Resize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Search_Grid_Data()
    On Error GoTo Search_Grid_Data_Error
    Dim oRecordSet          As SAPbobsCOM.Recordset
    Dim sQry                As String
    Dim BPLID               As String
    Dim Year                As String
    Dim Number              As String
    Dim Evaluate            As String
    Dim MSTCOD              As String
    Dim Cnt                 As Integer
    Dim ProcYN              As String
    
    ProcYN = "Y"
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oForm01.Freeze True
    
'    oForm01.Items("DistrQty").Specific.VALUE = ""
'    oForm01.Items("SumPkQty").Specific.VALUE = ""
'    oForm01.Items("SumPkWt").Specific.VALUE = ""
'    oForm01.Items("SumNPkQty").Specific.VALUE = ""
'    oForm01.Items("SumNPkWt").Specific.VALUE = ""
'    oForm01.Items("EmpId").Specific.VALUE = MDC_PS_Common.User_MSTCOD()
'    Call FlushToItemValue("EmpId")
'    oForm01.Items("DocDate").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    oMat01.Clear
    
    oForm01.Items("Svalue").Specific.VALUE = 0
     
    oForm01.Items("Code").Specific.VALUE = ""
    oForm01.Items("EmpNo").Specific.VALUE = ""
    oForm01.Items("EmpName").Specific.VALUE = ""
    oForm01.Items("CName").Specific.VALUE = ""
            
    BPLID = oForm01.Items("BPLId").Specific.VALUE
    Year = oForm01.Items("Year").Specific.VALUE
    Number = oForm01.Items("Number").Specific.VALUE
    Evaluate = oForm01.Items("Evaluate").Specific.VALUE
    MSTCOD = oForm01.Items("MSTCOD").Specific.VALUE
     
     
    If Evaluate = "1" Then
        sQry = " select COUNT(*) "
        sQry = sQry + " from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code "
        sQry = sQry + " Where a.U_BPLId = '" & BPLID & "'"
        sQry = sQry + " and a.U_Year ='" & Year & "'"
        sQry = sQry + " and a.U_Number = '" & Number & "'"
        sQry = sQry + " and Isnull(b.U_MSTCOD1,'') = '" & MSTCOD & "'"
        sQry = sQry + " and Isnull(b.U_Complet1,'N') = 'Y' "
    ElseIf Evaluate = "2" Then
        sQry = " select COUNT(*) "
        sQry = sQry + " from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code "
        sQry = sQry + " Where a.U_BPLId = '" & BPLID & "'"
        sQry = sQry + " and a.U_Year ='" & Year & "'"
        sQry = sQry + " and a.U_Number = '" & Number & "'"
        sQry = sQry + " and Isnull(b.U_MSTCOD2,'') = '" & MSTCOD & "'"
        sQry = sQry + " and Isnull(b.U_Complet2,'N') = 'Y' "
    ElseIf Evaluate = "3" Then
        sQry = " select COUNT(*) "
        sQry = sQry + " from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code "
        sQry = sQry + " Where a.U_BPLId = '" & BPLID & "'"
        sQry = sQry + " and a.U_Year ='" & Year & "'"
        sQry = sQry + " and a.U_Number = '" & Number & "'"
        sQry = sQry + " and Isnull(b.U_MSTCOD3,'') = '" & MSTCOD & "'"
        sQry = sQry + " and Isnull(b.U_Complet3,'N') = 'Y' "
    End If
    
    oRecordSet.DoQuery sQry
    
    Cnt = oRecordSet.Fields(0).VALUE
    
    If Cnt > 0 Then
        oForm01.Items("Complete").Specific.VALUE = "평가완료처리"
    Else
        oForm01.Items("Complete").Specific.VALUE = "평가완료미처리"
    End If
    
    If Evaluate = "1" Then
        sQry = " select COUNT(*) "
        sQry = sQry + " from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code "
        sQry = sQry + " Where a.U_BPLId = '" & BPLID & "'"
        sQry = sQry + " and a.U_Year ='" & Year & "'"
        sQry = sQry + " and a.U_Number = '" & Number & "'"
        sQry = sQry + " and Isnull(b.U_MSTCOD1,'') = '" & MSTCOD & "'"
    ElseIf Evaluate = "2" Then
        sQry = " select COUNT(*) "
        sQry = sQry + " from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code "
        sQry = sQry + " Where a.U_BPLId = '" & BPLID & "'"
        sQry = sQry + " and a.U_Year ='" & Year & "'"
        sQry = sQry + " and a.U_Number = '" & Number & "'"
        sQry = sQry + " and Isnull(b.U_MSTCOD2,'') = '" & MSTCOD & "'"
        sQry = sQry + " and Isnull(b.U_Complet1,'N') = 'N' "
    ElseIf Evaluate = "3" Then
        sQry = " select COUNT(*) "
        sQry = sQry + " from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code "
        sQry = sQry + " Where a.U_BPLId = '" & BPLID & "'"
        sQry = sQry + " and a.U_Year ='" & Year & "'"
        sQry = sQry + " and a.U_Number = '" & Number & "'"
        sQry = sQry + " and Isnull(b.U_MSTCOD3,'') = '" & MSTCOD & "'"
        sQry = sQry + " and Isnull(b.U_Complet2,'N') = 'N' "
    End If
    
    oRecordSet.DoQuery sQry
    Cnt = oRecordSet.Fields(0).VALUE
    
    If Evaluate = "1" Then
        If Cnt <= 0 Then
            Call MDC_Com.MDC_GF_Message("평가대상자가 없습니다.", "W")
            ProcYN = "N"
        End If
    ElseIf Evaluate = "2" Then
        If Cnt > 0 Then
            Call MDC_Com.MDC_GF_Message("1차평가가 완료되지 않았습니다.", "W")
            ProcYN = "N"
        End If
        '3차평가를 하기위해 임시로 막음.
'    ElseIf Evaluate = "3" Then
'        If Cnt > 0 Then
'            Call MDC_Com.MDC_GF_Message("2차평가가 완료되지 않았습니다.", "W")
'            ProcYN = "N"
'        End If
    End If
    
    If ProcYN = "Y" Then
        sQry = "EXEC PS_HR416_01 '" & BPLID & "','" & Year & "','" & Number & "','" & Evaluate & "','" & MSTCOD & "'"
        '    oGrid01.DataTable.Clear
            
        '    oForm01.DataSources.DataTables.Add "DataTable"
            
            '/ Procedure 실행(Grid 사용)
        '    Call oForm01.DataSources.DataTables.Item(0).ExecuteQuery(sQry)
    
    
    
        oGrid01.DataTable.ExecuteQuery (sQry)
        Call GridSetting
    End If
    oForm01.Freeze False
    
    Set oRecordSet = Nothing
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Search_Grid_Data_Error:
        Set oRecordSet = Nothing
        MDC_Com.MDC_GF_Message "Search_Grid_Data_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Search_Grid_Data1(ByVal oRow As Long)
    On Error GoTo Search_Grid_Data1_Error
    Dim oRecordSet          As SAPbobsCOM.Recordset
    Dim sQry                As String
    Dim RateCode            As String
    Dim BPLID               As String
    Dim MSTCOD              As String
    Dim i                   As Integer
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oForm01.Freeze True
    
'    oForm01.Items("DistrQty").Specific.VALUE = ""
'    oForm01.Items("SumPkQty").Specific.VALUE = ""
'    oForm01.Items("SumPkWt").Specific.VALUE = ""
'    oForm01.Items("SumNPkQty").Specific.VALUE = ""
'    oForm01.Items("SumNPkWt").Specific.VALUE = ""
'    oForm01.Items("EmpId").Specific.VALUE = MDC_PS_Common.User_MSTCOD()
'    Call FlushToItemValue("EmpId")
'    oForm01.Items("DocDate").Specific.VALUE = Format(Now, "YYYYMMDD")
    
    RateCode = oMat01.Columns("RateCode").Cells(oRow).Specific.VALUE
    BPLID = oForm01.Items("BPLId").Specific.VALUE '사업장도 함께 매개변수로 전달할 필요가 있음(2013.06.14 송명규)
    MSTCOD = oForm01.Items("EmpNo").Specific.VALUE '평가대상자
    
    'sQry = "EXEC PS_HR416_03 '" & RateCode & "'"
    sQry = "EXEC PS_HR416_03 '" & BPLID & "','" & RateCode & "'"
    
    oGrid02.DataTable.ExecuteQuery (sQry)
    Call GridSetting
    
          
    sQry = "EXEC PS_HR416_04 '" & BPLID & "', '" & MSTCOD & "', '" & RateCode & "'"

    oGrid03.DataTable.ExecuteQuery (sQry)
    Call GridSetting
    
    oForm01.Freeze False
    
    Set oRecordSet = Nothing
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Search_Grid_Data1_Error:
        Set oRecordSet = Nothing
        MDC_Com.MDC_GF_Message "Search_Grid_Data1_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

'****************************************************************************************************************
'// Grid 꾸며주기
'****************************************************************************************************************
Private Sub GridSetting() '(ByVal GridNum As Long)
On Error GoTo GridSetting_Error
    Dim i           As Integer
    Dim sColsTitle  As String
    Dim sColsLine   As String
    
    oForm01.Freeze True
    
    oGrid01.SelectionMode = ms_Single
            For i = 0 To oGrid01.Columns.Count - 1
                sColsTitle = oGrid01.Columns(i).TitleObject.Caption

                oGrid01.Columns(i).Editable = False
                
                If sColsTitle = "1차" Or sColsTitle = "2차" Or sColsTitle = "3차" Or sColsTitle = "평균" Then
                    oGrid01.Columns(i).RightJustified = True
                End If
                
                If oGrid01.DataTable.Columns(i).Type = ft_Float Then
                    oGrid01.Columns(i).RightJustified = True
                End If
            Next i
            
'    Select Case GridNum
'        Case 1
'            oGrid01.SelectionMode = ms_Single
'            For i = 0 To oGrid01.Columns.Count - 1
'                sColsTitle = oGrid01.Columns(i).TitleObject.Caption
'
'                oGrid01.Columns(i).Editable = False
'
'                If sColsTitle = "1차" Or sColsTitle = "평균" Then
'                    oGrid01.Columns(i).RightJustified = True
'                End If
'
'                If oGrid01.DataTable.Columns(i).Type = ft_Float Then
'                    oGrid01.Columns(i).RightJustified = True
'                End If
'            Next i
'        Case 2
'            oGrid02.SelectionMode = ms_Single
'            For i = 0 To oGrid02.Columns.Count - 1
'                sColsTitle = oGrid02.Columns(i).TitleObject.Caption
'
'                oGrid02.Columns(i).Editable = False
'
'                If sColsTitle = "1차" Or sColsTitle = "2차" Or sColsTitle = "평균" Then
'                    oGrid01.Columns(i).RightJustified = True
'                End If
'
'
'                If oGrid02.DataTable.Columns(i).Type = ft_Float Then
'                    oGrid02.Columns(i).RightJustified = True
'                End If
'            Next i
'        Case 3
'            oGrid03.SelectionMode = ms_Single
'            For i = 0 To oGrid03.Columns.Count - 1
'                sColsTitle = oGrid03.Columns(i).TitleObject.Caption
'
'                oGrid03.Columns(i).Editable = False
'
'                If sColsTitle = "1차" Or sColsTitle = "2차" Or sColsTitle = "3차" Or sColsTitle = "평균" Then
'                    oGrid01.Columns(i).RightJustified = True
'                End If
'
'                If oGrid03.DataTable.Columns(i).Type = ft_Float Then
'                    oGrid03.Columns(i).RightJustified = True
'                End If
'            Next i
'
'        End Select
         
    
        
    oForm01.Freeze False
    
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
GridSetting_Error:
    MDC_Com.MDC_GF_Message "GridSetting_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub Search_Matrix_Data()
    On Error GoTo Search_Matrix_Data_Error
    Dim oRecordSet          As SAPbobsCOM.Recordset
    Dim sQry                As String
    
    Dim i                   As Integer
    Dim j                   As Integer
    Dim Cnt                 As Integer
    
    Dim Code                As String
    Dim BPLID               As String
    Dim Year                As String
    Dim MSTCOD              As String
    Dim Number              As String
    Dim Evaluate            As String
    Dim Svalue              As Double
    
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    oForm01.Freeze True
    
    BPLID = oForm01.Items("BPLId").Specific.VALUE
    Year = oForm01.Items("Year").Specific.VALUE
    MSTCOD = oForm01.Items("MSTCOD").Specific.VALUE
    Number = oForm01.Items("Number").Specific.VALUE
    Evaluate = oForm01.Items("Evaluate").Specific.VALUE
    
    
    
    oGrid02.DataTable.Clear
    oGrid03.DataTable.Clear
    
     For i = 0 To oGrid01.Rows.Count - 1
        If oGrid01.Rows.IsSelected(i) = True Then
            Code = Trim(oGrid01.DataTable.GetValue(0, i))
            
            MSTCOD = Trim(oGrid01.DataTable.GetValue(5, i))
            
            oForm01.Items("Code").Specific.VALUE = Code
            oForm01.Items("EmpNo").Specific.VALUE = MSTCOD
            oForm01.Items("EmpName").Specific.VALUE = Trim(oGrid01.DataTable.GetValue(3, i))
            oForm01.Items("CName").Specific.VALUE = Trim(oGrid01.DataTable.GetValue(4, i))
            
            '/ Procedure 실행
            sQry = "EXEC PS_HR416_02 '" & BPLID & "', '" & Year & "', '" & MSTCOD & "', '" & Number & "', '" & Evaluate & "'"
            oRecordSet.DoQuery sQry

            
             'Matrix에 출력===========================================================================================
             '// Line 초기화
             Cnt = oDS_PS_HR416L.Size
             If Cnt > 0 Then
                 For j = 0 To Cnt - 1
                     oDS_PS_HR416L.RemoveRecord oDS_PS_HR416L.Size - 1
                 Next j
                 If Cnt = 1 Then
                     oDS_PS_HR416L.Clear
                 End If
             End If
             oMat01.LoadFromDataSource
             '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                 
            '//Matrix에 Data 뿌려준다
            j = 1
             Do Until oRecordSet.EOF
                     
                 If oDS_PS_HR416L.Size < j Then
                     oDS_PS_HR416L.InsertRecord j - 1 '라인추가
                 End If
                 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                 oDS_PS_HR416L.setValue "U_LineNum", j - 1, j
                 oDS_PS_HR416L.setValue "U_ColReg01", j - 1, oRecordSet.Fields(0).VALUE
                 oDS_PS_HR416L.setValue "U_ColReg02", j - 1, oRecordSet.Fields(1).VALUE
                 oDS_PS_HR416L.setValue "U_ColReg03", j - 1, oRecordSet.Fields(2).VALUE
                 oDS_PS_HR416L.setValue "U_ColQty01", j - 1, oRecordSet.Fields(3).VALUE
                 oDS_PS_HR416L.setValue "U_ColReg04", j - 1, oRecordSet.Fields(4).VALUE
                 oDS_PS_HR416L.setValue "U_ColQty02", j - 1, oRecordSet.Fields(5).VALUE
                 
                 
                 Svalue = Svalue + oRecordSet.Fields(5).VALUE
                 
                 j = j + 1
                 oRecordSet.MoveNext
             Loop
             oMat01.LoadFromDataSource
             ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
     Next i
     If Svalue <> 0 Then '
        oForm01.Items("Btn02").Specific.Caption = "수정"
     Else
        oForm01.Items("Btn02").Specific.Caption = "추가"
     End If
     oForm01.Items("Svalue").Specific.VALUE = Svalue
    
    oForm01.Freeze False
    Set oRecordSet = Nothing
Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Search_Matrix_Data_Error:
    Set oRecordSet = Nothing
    MDC_Com.MDC_GF_Message "Search_Matrix_Data_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub


Private Function Save_Data(ByVal A_B As String, Optional Cancel As Boolean) As Boolean
    On Error GoTo Save_Data_Error
    Dim oRecordSet          As SAPbobsCOM.Recordset
    Dim sQry                As String
    Dim i                   As Integer
'    Dim j                   As Integer
    Dim ErrNum              As Integer
    
    Dim BPLID                As String
    Dim Year                As String
    Dim Number                As String
    
    Dim Code                As String
    Dim EmpNo               As String
    Dim Evaluate            As Integer
    
    Dim Cnt                 As Integer
    Dim Svalue              As Double
    Dim Avg                 As Double
    
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    BPLID = oForm01.Items("BPLId").Specific.VALUE
    Year = oForm01.Items("Year").Specific.VALUE
    Number = oForm01.Items("Number").Specific.VALUE
    Code = oForm01.Items("Code").Specific.VALUE
    EmpNo = oForm01.Items("EmpNo").Specific.VALUE
    Evaluate = oForm01.Items("Evaluate").Specific.VALUE
            
    sQry = "Select Count(*) From Z_PS_HR410L Where Code = '" & Code & "' And MSTCOD = '" & EmpNo & "' and Evaluate = '" & Evaluate & "'"
    oRecordSet.DoQuery sQry
    
    Cnt = oRecordSet.Fields(0).VALUE
    
    oMat01.FlushToDataSource
    
    If oMat01.VisualRowCount > 0 Then
        For i = 0 To oMat01.VisualRowCount - 1
            oDS_PS_HR416L.Offset = i
            If Cnt <= 0 Then '신규 insert
                sQry = "insert into Z_PS_HR410L values ("
                sQry = sQry & "'" & Code & "','" & EmpNo & "','" & Evaluate & "',"
                sQry = sQry & "'" & Trim(oDS_PS_HR416L.GetValue("U_LineNum", i)) & "'," 'LineNum
                sQry = sQry & "'" & Trim(oDS_PS_HR416L.GetValue("U_ColReg01", i)) & "'," 'RateCode
                sQry = sQry & "'" & Trim(oDS_PS_HR416L.GetValue("U_ColReg02", i)) & "',"
                sQry = sQry & "'" & Trim(oDS_PS_HR416L.GetValue("U_ColReg03", i)) & "',"
                sQry = sQry & "'" & Trim(oDS_PS_HR416L.GetValue("U_ColQty01", i)) & "',"
                sQry = sQry & "'" & Trim(oDS_PS_HR416L.GetValue("U_ColReg04", i)) & "',"
                sQry = sQry & "'" & Trim(oDS_PS_HR416L.GetValue("U_ColQty02", i)) & "')"
'
                oRecordSet.DoQuery sQry
                
                If Trim(oDS_PS_HR416L.GetValue("U_ColReg01", i)) <> "A12" Then
                    Svalue = Svalue + Trim(oDS_PS_HR416L.GetValue("U_ColQty02", i))
                End If
                
                   
            Else
                sQry = "Update Z_PS_HR410L"
                sQry = sQry + " Set Grade = '" & Trim(oDS_PS_HR416L.GetValue("U_ColReg04", i)) & "' ,"
                sQry = sQry + " Value = '" & Trim(oDS_PS_HR416L.GetValue("U_ColQty02", i)) & "'"
                sQry = sQry + " Where Code = '" & Code & "' And MSTCOD = '" & EmpNo & "' And Evaluate = '" & Evaluate & "'"
                sQry = sQry + " And LineNum = '" & Trim(oDS_PS_HR416L.GetValue("U_LineNum", i)) & "'"
                oRecordSet.DoQuery sQry
                
                If Trim(oDS_PS_HR416L.GetValue("U_ColReg01", i)) <> "A12" Then
                    Svalue = Svalue + Trim(oDS_PS_HR416L.GetValue("U_ColQty02", i))
                End If
                
            End If

        Next i
        
        Select Case Evaluate
            Case "1"
                Avg = Round(Svalue * 3 / 10, 1)
            
                sQry = "Update [@PS_HR410L] "
                sQry = sQry + " Set U_Value1 = '" & Svalue & "', U_Avg1 = '" & Avg & "'"
                sQry = sQry + " From [@PS_HR410H] a "
                sQry = sQry + " Where a.Code = [@PS_HR410L].Code "
                sQry = sQry + " And a.U_BPLId = '" & BPLID & "' and a.U_Year = '" & Year & "'"
                sQry = sQry + " And a.U_Number = '" & Number & "' and [@PS_HR410L].U_MSTCOD = '" & EmpNo & "'"
                oRecordSet.DoQuery sQry
            Case "2" '//2차
                Avg = Round(Svalue * 4 / 10, 1)
            
                sQry = "Update [@PS_HR410L] "
                sQry = sQry + " Set U_Value2 = '" & Svalue & "', U_Avg2 = U_Avg1 + '" & Avg & "'"
                sQry = sQry + " From [@PS_HR410H] a "
                sQry = sQry + " Where a.Code = [@PS_HR410L].Code "
                sQry = sQry + " And a.U_BPLId = '" & BPLID & "' and a.U_Year = '" & Year & "'"
                sQry = sQry + " And a.U_Number = '" & Number & "' and [@PS_HR410L].U_MSTCOD = '" & EmpNo & "'"
                oRecordSet.DoQuery sQry
            Case "3" '//3차
                Avg = Round(Svalue * 3 / 10, 1)
            
                sQry = "Update [@PS_HR410L] "
                sQry = sQry + " Set U_Value3 = '" & Svalue & "', U_Avg3 = U_Avg2 + '" & Avg & "'"
                sQry = sQry + " From [@PS_HR410H] a "
                sQry = sQry + " Where a.Code = [@PS_HR410L].Code "
                sQry = sQry + " And a.U_BPLId = '" & BPLID & "' and a.U_Year = '" & Year & "'"
                sQry = sQry + " And a.U_Number = '" & Number & "' and [@PS_HR410L].U_MSTCOD = '" & EmpNo & "'"
                oRecordSet.DoQuery sQry
        End Select
    End If
    
    Set oRecordSet = Nothing
    Save_Data = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Save_Data_Error:
    Set oRecordSet = Nothing
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

Private Function Complete_Data()
    On Error GoTo Complete_Data_Error
    Dim oRecordSet          As SAPbobsCOM.Recordset
    Dim sQry                As String
    Dim i                   As Integer

    Dim ErrNum              As Integer
    
    Dim BPLID                As String
    Dim Year                As String
    Dim Number                As String
    
    Dim Code                As String
    Dim MSTCOD               As String
    Dim Evaluate            As Integer
    Dim Complete            As String
    
    Dim Cnt                 As Integer
    Dim Svalue              As Double
    Dim Avg                 As Double
    
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    BPLID = oForm01.Items("BPLId").Specific.VALUE
    Year = oForm01.Items("Year").Specific.VALUE
    Number = oForm01.Items("Number").Specific.VALUE
    Code = oForm01.Items("Code").Specific.VALUE
    MSTCOD = oForm01.Items("MSTCOD").Specific.VALUE
    Evaluate = oForm01.Items("Evaluate").Specific.VALUE
    
    Complete = oForm01.Items("Complete").Specific.VALUE
    
'    If Complete = "평가완료처리" Then
'        Sbo_Application.SetStatusBarMessage "이미 평가완료처리를 하였습니다.. 확인바랍니다.", bmt_Short, True
'        Exit Function
'    End If
    
    
    
    If Evaluate = "1" Then
        sQry = "select COUNT(*) from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code"
        sQry = sQry + " left Join Z_PS_HR410L c On b.Code = c.Code and b.U_MSTCOD = c.MSTCOD and c.Evaluate = '" & Evaluate & "'"
        sQry = sQry + " Where a.U_BPLId = '" & BPLID & "'"
        sQry = sQry + " and a.U_Year ='" & Year & "'"
        sQry = sQry + " and a.U_Number = '" & Number & "'"
        sQry = sQry + " and Isnull(b.U_MSTCOD1,'') = '" & MSTCOD & "'"
        sQry = sQry + " and isnull(c.Value,0) = 0 "
    ElseIf Evaluate = "2" Then
        sQry = "select COUNT(*) from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code"
        sQry = sQry + " left Join Z_PS_HR410L c On b.Code = c.Code and b.U_MSTCOD = c.MSTCOD and c.Evaluate = '" & Evaluate & "'"
        sQry = sQry + " Where a.U_BPLId = '" & BPLID & "'"
        sQry = sQry + " and a.U_Year ='" & Year & "'"
        sQry = sQry + " and a.U_Number = '" & Number & "'"
        sQry = sQry + " and Isnull(b.U_MSTCOD2,'') = '" & MSTCOD & "'"
        sQry = sQry + " and isnull(c.Value,0) = 0 "
    ElseIf Evaluate = "3" Then
        sQry = "select COUNT(*) from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code"
        sQry = sQry + " left Join Z_PS_HR410L c On b.Code = c.Code and b.U_MSTCOD = c.MSTCOD and c.Evaluate = '" & Evaluate & "'"
        sQry = sQry + " Where a.U_BPLId = '" & BPLID & "'"
        sQry = sQry + " and a.U_Year ='" & Year & "'"
        sQry = sQry + " and a.U_Number = '" & Number & "'"
        sQry = sQry + " and Isnull(b.U_MSTCOD3,'') = '" & MSTCOD & "'"
        sQry = sQry + " and isnull(c.Value,0) = 0 "
    End If
    
    oRecordSet.DoQuery sQry
    
    Cnt = oRecordSet.Fields(0).VALUE
    
    If Cnt > 0 Then
        '//평가가 다되지 않았다.
        Sbo_Application.SetStatusBarMessage "평가자의 평가가 완료되지 않았습니다. 확인바랍니다.", bmt_Short, True
    Else
        If Evaluate = "1" Then
            '//평가가 전부 다되었으면
            sQry = " Update [@PS_HR410L] set U_Complet1 = 'Y' "
            sQry = sQry + " From [@PS_HR410H] a"
            sQry = sQry + " Where a.Code = [@PS_HR410L].Code"
            sQry = sQry + " And a.U_BPLId = '" & BPLID & "'"
            sQry = sQry + " And a.U_Year = '" & Year & "'"
            sQry = sQry + " And a.U_Number = '" & Number & "'"
            sQry = sQry + " And [@PS_HR410L].U_MSTCOD1 = '" & MSTCOD & "'"
        ElseIf Evaluate = "2" Then
            sQry = " Update [@PS_HR410L] set U_Complet2 = 'Y' "
            sQry = sQry + " From [@PS_HR410H] a"
            sQry = sQry + " Where a.Code = [@PS_HR410L].Code"
            sQry = sQry + " And a.U_BPLId = '" & BPLID & "'"
            sQry = sQry + " And a.U_Year = '" & Year & "'"
            sQry = sQry + " And a.U_Number = '" & Number & "'"
            sQry = sQry + " And [@PS_HR410L].U_MSTCOD2 = '" & MSTCOD & "'"
        ElseIf Evaluate = "3" Then
            sQry = " Update [@PS_HR410L] set U_Complet3 = 'Y' "
            sQry = sQry + " From [@PS_HR410H] a"
            sQry = sQry + " Where a.Code = [@PS_HR410L].Code"
            sQry = sQry + " And a.U_BPLId = '" & BPLID & "'"
            sQry = sQry + " And a.U_Year = '" & Year & "'"
            sQry = sQry + " And a.U_Number = '" & Number & "'"
            sQry = sQry + " And [@PS_HR410L].U_MSTCOD3 = '" & MSTCOD & "'"
        End If
        oRecordSet.DoQuery sQry
        Sbo_Application.SetStatusBarMessage Evaluate & "차 평가완료처리했습니다..", bmt_Short, True
    End If
    oForm01.Items("Btn01").Click ct_Regular
    Set oRecordSet = Nothing
    
    
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Complete_Data_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "기포장중량 + 포장중량은 이동요청중량을 초과할수 없습니다.. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "해당작지번호에 대한 작업지시등록 공정중 바렐 또는 포장공정이 없거나 이미등록되어 있습니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 100 Then
    Else
        MDC_Com.MDC_GF_Message "Save_Data_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    
End Function

Private Function InComplete_Data()
    On Error GoTo InComplete_Data_Error
    Dim oRecordSet          As SAPbobsCOM.Recordset
    Dim sQry                As String
    Dim i                   As Integer

    Dim ErrNum              As Integer
    
    Dim BPLID                As String
    Dim Year                As String
    Dim Number                As String
    
    Dim Code                As String
    Dim MSTCOD               As String
    Dim Evaluate            As Integer
    Dim Complete            As String
    
    Dim Cnt                 As Integer
    Dim Svalue              As Double
    Dim Avg                 As Double
    
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    BPLID = oForm01.Items("BPLId").Specific.VALUE
    Year = oForm01.Items("Year").Specific.VALUE
    Number = oForm01.Items("Number").Specific.VALUE
    Code = oForm01.Items("Code").Specific.VALUE
    MSTCOD = oForm01.Items("MSTCOD").Specific.VALUE
    Evaluate = oForm01.Items("Evaluate").Specific.VALUE
    
    Complete = oForm01.Items("Complete").Specific.VALUE
    
    If Complete = "평가완료미처리" Then
        Sbo_Application.SetStatusBarMessage "평가완료처리시만 작업할 수 있습니다. 확인바랍니다.", bmt_Short, True
        Exit Function
    End If
    
    
    
    If Evaluate = "1" Then
        sQry = "select COUNT(*) from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code"
        sQry = sQry + " Where a.U_BPLId = '" & BPLID & "'"
        sQry = sQry + " and a.U_Year ='" & Year & "'"
        sQry = sQry + " and a.U_Number = '" & Number & "'"
        sQry = sQry + " and Isnull(b.U_MSTCOD1,'') = '" & MSTCOD & "'"
        sQry = sQry + " and isnull(b.U_AValue1,0) <> 0 "  '1차계수조정
    ElseIf Evaluate = "2" Then
        sQry = "select COUNT(*) from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code"
        sQry = sQry + " Where a.U_BPLId = '" & BPLID & "'"
        sQry = sQry + " and a.U_Year ='" & Year & "'"
        sQry = sQry + " and a.U_Number = '" & Number & "'"
        sQry = sQry + " and Isnull(b.U_MSTCOD2,'') = '" & MSTCOD & "'"
        sQry = sQry + " and isnull(b.U_AValue2,0) <> 0 " '2차계수조정
    ElseIf Evaluate = "3" Then
        sQry = "select COUNT(*) from [@PS_HR410H] a Inner Join [@PS_HR410L] b On a.Code = b.Code"
        sQry = sQry + " Where a.U_BPLId = '" & BPLID & "'"
        sQry = sQry + " and a.U_Year ='" & Year & "'"
        sQry = sQry + " and a.U_Number = '" & Number & "'"
        sQry = sQry + " and Isnull(b.U_MSTCOD3,'') = '" & MSTCOD & "'"
        sQry = sQry + " and isnull(b.U_AValue3,0) <> 0 " '3차계수조정
    End If
    
    oRecordSet.DoQuery sQry
    
    Cnt = oRecordSet.Fields(0).VALUE
    
    If Cnt > 0 Then
        '//계수조정이 되었다.
        Sbo_Application.SetStatusBarMessage "계수조정처리가 되었습니다. 계수조정취소작업을 먼저하시기바랍니다.", bmt_Short, True
    Else
        If Evaluate = "1" Then
            '//계수조정이 안되었으면
            sQry = " Update [@PS_HR410L] set U_Complet1 = 'N' "
            sQry = sQry + " From [@PS_HR410H] a"
            sQry = sQry + " Where a.Code = [@PS_HR410L].Code"
            sQry = sQry + " And a.U_BPLId = '" & BPLID & "'"
            sQry = sQry + " And a.U_Year = '" & Year & "'"
            sQry = sQry + " And a.U_Number = '" & Number & "'"
            sQry = sQry + " And [@PS_HR410L].U_MSTCOD1 = '" & MSTCOD & "'"
        ElseIf Evaluate = "2" Then
            sQry = " Update [@PS_HR410L] set U_Complet2 = 'N' "
            sQry = sQry + " From [@PS_HR410H] a"
            sQry = sQry + " Where a.Code = [@PS_HR410L].Code"
            sQry = sQry + " And a.U_BPLId = '" & BPLID & "'"
            sQry = sQry + " And a.U_Year = '" & Year & "'"
            sQry = sQry + " And a.U_Number = '" & Number & "'"
            sQry = sQry + " And [@PS_HR410L].U_MSTCOD2 = '" & MSTCOD & "'"
        ElseIf Evaluate = "3" Then
            sQry = " Update [@PS_HR410L] set U_Complet3 = 'N' "
            sQry = sQry + " From [@PS_HR410H] a"
            sQry = sQry + " Where a.Code = [@PS_HR410L].Code"
            sQry = sQry + " And a.U_BPLId = '" & BPLID & "'"
            sQry = sQry + " And a.U_Year = '" & Year & "'"
            sQry = sQry + " And a.U_Number = '" & Number & "'"
            sQry = sQry + " And [@PS_HR410L].U_MSTCOD3 = '" & MSTCOD & "'"
        End If
        oRecordSet.DoQuery sQry
        Sbo_Application.SetStatusBarMessage Evaluate & "차 평가완료처리 취소했습니다.", bmt_Short, True
    End If
    oForm01.Items("Btn01").Click ct_Regular
    Set oRecordSet = Nothing
    
    
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
InComplete_Data_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "기포장중량 + 포장중량은 이동요청중량을 초과할수 없습니다.. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "해당작지번호에 대한 작업지시등록 공정중 바렐 또는 포장공정이 없거나 이미등록되어 있습니다. 확인하여 주십시오.", "E"
    ElseIf ErrNum = 100 Then
    Else
        MDC_Com.MDC_GF_Message "InComplete_Data_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    
End Function

'// 패스워드 Check
Public Function PasswordChk(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
On Error GoTo PasswordChk_Error
    Dim sQry As String
    Dim RecordSet01 As SAPbobsCOM.Recordset
    Set RecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim MSTCOD$, PassWd$
        
    MSTCOD = Trim(oForm01.Items("MSTCOD").Specific.VALUE)
    PassWd = Trim(oForm01.Items("PassWd").Specific.VALUE)
    
    If Trim(MSTCOD) = "" Then
        MDC_Com.MDC_GF_Message "사번이 없습니다. 입력바랍니다!", "E"
        PasswordChk = False
        Exit Function
    End If
        
    sQry = "Select Count(*) From Z_PS_HRPASS Where MSTCOD = '" & Trim(oForm01.Items("MSTCOD").Specific.VALUE) & "'"
    sQry = sQry & " And  BPLId = '" & oForm01.Items("BPLId").Specific.VALUE & "' "
    sQry = sQry & " And  PassWd = '" & oForm01.Items("PassWd").Specific.VALUE & "' "
    RecordSet01.DoQuery sQry
    
    If Trim(RecordSet01.Fields(0).VALUE) <= 0 Then
        PasswordChk = False
    Else
        PasswordChk = True
    End If
        
    Set RecordSet01 = Nothing
    
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PasswordChk_Error:
    Set RecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "PasswordChk_Error:" & Err.Number & " - " & Err.Description, "E"
End Function

