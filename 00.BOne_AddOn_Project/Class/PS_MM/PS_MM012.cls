VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM012.cls
'//  Module         : MM
'//  Description    : 자재 순환품 관리
'//  FormType       : PS_MM012
'//  Create Date    : 2012.10.10
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_MM012H       As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_MM012L       As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oLast_Mode&
Private oCode$

Private oSAP_Connection01   As Object
Private oYear               As String


'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM012.srf")
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
    
    oFormUniqueID01 = "PS_MM012_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
  
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_FIND_MODE
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "Code"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    Call FormItemEnabled
    
    oForm01.EnableMenu ("1283"), True         '// 삭제
    oForm01.EnableMenu ("1287"), True         '// 복제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1284"), False        '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제
        
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

'****************************************************************************************************************
'// ItemEventHander
'****************************************************************************************************************
Public Sub Raise_ItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_ItemEvent_Error
    Dim i&
    
    Dim sQry            As String
    Dim oRecordSet01      As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                        
                        If Qty_Check = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        If oForm01.Mode = fm_ADD_MODE Then
                            Dim Code$
                            Code = Trim(oDS_PS_MM012H.GetValue("U_BPLId", 0)) & Trim(oDS_PS_MM012H.GetValue("U_Year", 0))
                            Call oDS_PS_MM012H.setValue("Code", 0, Code)
                            Call oDS_PS_MM012H.setValue("Name", 0, Code)
                        End If
                        Call Delete_EmptyRow
                        
                        oCode = Trim(oDS_PS_MM012H.GetValue("Code", 0))
                    End If
                    
                    oYear = Trim(oDS_PS_MM012H.GetValue("U_Year", 0))
                    
                ElseIf pval.ItemUID = "Btn01" Then
'                    Dim DD As String
'                    DD = Right(Left(Format(Now, "YYYYMMDD"), 8), 2)
'                    If DD >= "26" Then
                        Call R3_Interface
'                    Else
'                        Sbo_Application.MessageBox ("26일 부터 말일까지만 전송가능합니다.")
'                    End If
                End If
'et_KEY_DOWN ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_KEY_DOWN: '//2
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "Mat01" And pval.ColUID = "PQDocNum" Then
                        If oMat01.Columns("PQDocNum").Cells(pval.Row).Specific.VALUE = "" Then
                            Dim ChildForm01 As Variant
                            Set ChildForm01 = New PS_MM013
                            Call ChildForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row, 0)
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                End If
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            
                If pval.ItemUID = "Mat01" Then
                    If pval.Row > 0 Then
                        Call oMat01.SelectRow(pval.Row, True, False)
                        oLast_Col_Row = pval.Row
                    End If
                End If
                        
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
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    oForm01.Freeze True
                    If oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1281")
                        Call oDS_PS_MM012H.setValue("Code", 0, oCode)
                        oForm01.Items("1").Click ct_Regular
                    End If
                    
                    Call FormItemEnabled
                    oForm01.Freeze False
                ElseIf pval.ItemUID = "Btn01" Then
                
                End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If oForm01.Mode = fm_UPDATE_MODE Then
                        oForm01.Items("Btn01").Enabled = False
                    End If
                    If pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "PQDocNum" Then
                            oForm01.Freeze True
                                If (pval.Row = oMat01.RowCount Or oMat01.VisualRowCount = 0) And _
                                Trim(oMat01.Columns("PQDocNum").Cells(pval.Row).Specific.VALUE) <> "" Then
                                    oMat01.FlushToDataSource
                                    oDS_PS_MM012L.InsertRecord pval.Row
                                    oDS_PS_MM012L.setValue "U_LineNum", pval.Row, pval.Row + 1
                                    oMat01.LoadFromDataSource '
                                End If
                            oForm01.Freeze False
                        ElseIf pval.ColUID = "Mm01" Or pval.ColUID = "Mm02" Or pval.ColUID = "Mm03" Or pval.ColUID = "Mm04" _
                        Or pval.ColUID = "Mm05" Or pval.ColUID = "Mm06" Or pval.ColUID = "Mm07" Or pval.ColUID = "Mm08" _
                        Or pval.ColUID = "Mm09" Or pval.ColUID = "Mm10" Or pval.ColUID = "Mm11" Or pval.ColUID = "Mm12" Then
                            oForm01.Freeze True
                                Dim MmTot As Long
                                Dim Mm01 As Long, Mm02 As Long, Mm03 As Long, Mm04 As Long
                                Dim Mm05 As Long, Mm06 As Long, Mm07 As Long, Mm08 As Long
                                Dim Mm09 As Long, Mm10 As Long, Mm11 As Long, Mm12 As Long
                                
                                If Trim(oMat01.Columns("Mm01").Cells(pval.Row).Specific.VALUE) = "" Then
                                    Mm01 = 0
                                Else
                                    Mm01 = oMat01.Columns("Mm01").Cells(pval.Row).Specific.VALUE
                                End If
                                If Trim(oMat01.Columns("Mm02").Cells(pval.Row).Specific.VALUE) = "" Then
                                    Mm02 = 0
                                Else
                                    Mm02 = oMat01.Columns("Mm02").Cells(pval.Row).Specific.VALUE
                                End If
                                If Trim(oMat01.Columns("Mm03").Cells(pval.Row).Specific.VALUE) = "" Then
                                    Mm03 = 0
                                Else
                                    Mm03 = oMat01.Columns("Mm03").Cells(pval.Row).Specific.VALUE
                                End If
                                If Trim(oMat01.Columns("Mm04").Cells(pval.Row).Specific.VALUE) = "" Then
                                    Mm04 = 0
                                Else
                                    Mm04 = oMat01.Columns("Mm04").Cells(pval.Row).Specific.VALUE
                                End If
                                If Trim(oMat01.Columns("Mm05").Cells(pval.Row).Specific.VALUE) = "" Then
                                    Mm05 = 0
                                Else
                                    Mm05 = oMat01.Columns("Mm05").Cells(pval.Row).Specific.VALUE
                                End If
                                If Trim(oMat01.Columns("Mm06").Cells(pval.Row).Specific.VALUE) = "" Then
                                    Mm06 = 0
                                Else
                                    Mm06 = oMat01.Columns("Mm06").Cells(pval.Row).Specific.VALUE
                                End If
                                If Trim(oMat01.Columns("Mm07").Cells(pval.Row).Specific.VALUE) = "" Then
                                    Mm07 = 0
                                Else
                                    Mm07 = oMat01.Columns("Mm07").Cells(pval.Row).Specific.VALUE
                                End If
                                If Trim(oMat01.Columns("Mm08").Cells(pval.Row).Specific.VALUE) = "" Then
                                    Mm08 = 0
                                Else
                                    Mm08 = oMat01.Columns("Mm08").Cells(pval.Row).Specific.VALUE
                                End If
                                If Trim(oMat01.Columns("Mm09").Cells(pval.Row).Specific.VALUE) = "" Then
                                    Mm09 = 0
                                Else
                                    Mm09 = oMat01.Columns("Mm09").Cells(pval.Row).Specific.VALUE
                                End If
                                If Trim(oMat01.Columns("Mm10").Cells(pval.Row).Specific.VALUE) = "" Then
                                    Mm10 = 0
                                Else
                                    Mm10 = oMat01.Columns("Mm10").Cells(pval.Row).Specific.VALUE
                                End If
                                If Trim(oMat01.Columns("Mm11").Cells(pval.Row).Specific.VALUE) = "" Then
                                    Mm11 = 0
                                Else
                                    Mm11 = oMat01.Columns("Mm11").Cells(pval.Row).Specific.VALUE
                                End If
                                If Trim(oMat01.Columns("Mm12").Cells(pval.Row).Specific.VALUE) = "" Then
                                    Mm12 = 0
                                Else
                                    Mm12 = oMat01.Columns("Mm12").Cells(pval.Row).Specific.VALUE
                                End If
                                
                                MmTot = Mm01 + Mm02 + Mm03 + Mm04 + Mm05 + Mm06 + Mm07 + Mm08 + Mm09 + Mm10 + Mm11 + Mm12
                                
                                oMat01.Columns("MmTot").Cells(pval.Row).Specific.VALUE = MmTot
                            oForm01.Freeze False
                        End If
                    End If
                End If
'et_MATRIX_LOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_MATRIX_LOAD: '//11
                Add_MatrixRow oMat01.RowCount, False
                Call FormItemEnabled
                
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            
                oMat01.AutoResizeColumns
            
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
'et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
                Set oDS_PS_MM012H = Nothing
                Set oDS_PS_MM012L = Nothing
                Set oSAP_Connection01 = Nothing
        End Select
    End If
    
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
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
            Case "1283": '삭제
                If Sbo_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") = 2 Then
                    BubbleEvent = False
                    Exit Sub
                End If
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행삭제
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 0 To oMat01.VisualRowCount - 1
                        oMat01.Columns("LineNum").Cells(i + 1).Specific.VALUE = i + 1
                    Next i
                    
                    oMat01.FlushToDataSource
                    oDS_PS_MM012L.RemoveRecord oDS_PS_MM012L.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                    
                    If oMat01.Columns("PQDocNum").Cells(oMat01.RowCount).Specific.VALUE <> "" Then
                        Call Add_MatrixRow(oMat01.RowCount, False)
                    End If
                End If
            Case "1281": '찾기
                oForm01.Freeze True
                Call FormItemEnabled
                Call Initialization
'                oForm01.Items("CycleCod").Click ct_Regular
                oForm01.Freeze False
            Case "1282": '추가
                oForm01.Freeze True
                Call FormItemEnabled
                Call Initialization
                Call Add_MatrixRow(0, True)
                oForm01.Freeze False
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                oForm01.Freeze True
                Call FormItemEnabled
'                If oMat01.VisualRowCount > 0 Then
'                    If oMat01.Columns("PQDocNum").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
'                        Add_MatrixRow oMat01.RowCount, False
'                    End If
'                End If
                oForm01.Freeze False
            Case "1287": '// 복제
                oForm01.Freeze True
                oDS_PS_MM012H.setValue "Code", 0, ""
                oDS_PS_MM012H.setValue "Name", 0, ""
                oDS_PS_MM012H.setValue "U_Year", 0, ""
                
                For i = 0 To oMat01.VisualRowCount - 1
                    oMat01.FlushToDataSource
                    oDS_PS_MM012L.setValue "Code", i, ""
                    oMat01.LoadFromDataSource
                Next i
                oForm01.Freeze False
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_MenuEvent_Error:
    oForm01.Freeze False
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
    Set oDS_PS_MM012H = oForm01.DataSources.DBDataSources("@PS_MM012H")
    Set oDS_PS_MM012L = oForm01.DataSources.DBDataSources("@PS_MM012L")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_Auto

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
    Set oCombo = oForm01.Items("BPLId").Specific
    sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
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
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
    '//아이디별 사번 세팅
'    oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD
    
    '//아이디별 부서 세팅
'    Set oCombo = oForm01.Items("DeptCode").Specific
'    oCombo.Select MDC_PS_Common.User_DeptCode, psk_ByValue

    oDS_PS_MM012H.setValue "U_Year", 0, Left(Now, 4)
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
    If (oForm01.Mode = fm_ADD_MODE) Then
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("Year").Enabled = True
        oForm01.Items("Btn01").Enabled = False
        oForm01.Items("Mat01").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("Year").Enabled = True
        oForm01.Items("Btn01").Enabled = False
        oForm01.Items("Mat01").Enabled = False
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("Year").Enabled = True
        oForm01.Items("Btn01").Enabled = True
        oForm01.Items("Mat01").Enabled = True
    End If
    
    Dim Today As String
    Dim YY As String
    Dim MM As String
    Dim DD As String
    Dim Year As String
    
    YY = Left(Format(Now, "YYYYMMDD"), 4)
    MM = Mid(Left(Format(Now, "YYYYMMDD"), 8), 5, 2)
    DD = Right(Left(Format(Now, "YYYYMMDD"), 8), 2)
    
    Year = Trim(oDS_PS_MM012H.GetValue("U_Year", 0))
    
    If YY < Year Then
        oMat01.Columns("Mm01").Editable = True
    Else
        oMat01.Columns("Mm01").Editable = False
    End If


    If YY = Year And MM = "01" Then 'And DD >= "26" Then
        oMat01.Columns("Mm01").Editable = True
        oMat01.Columns("Mm02").Editable = True
    Else
        oMat01.Columns("Mm01").Editable = False
        oMat01.Columns("Mm02").Editable = False
    End If

'    If YY = Year And Val(MM) = Val("02") And Val(DD) >= Val("26") Then
''    If YY = oYear And Val(MM) = Val("02") And Val(DD) >= Val("26") Then
'        oMat01.Columns("Mm03").Editable = True
'    Else
'        oMat01.Columns("Mm03").Editable = False
'    End If

    If YY = Year And MM = "02" Then 'And DD >= "26" Then '
        oMat01.Columns("Mm02").Editable = True
        oMat01.Columns("Mm03").Editable = True
    Else
        oMat01.Columns("Mm03").Editable = False
    End If
    
    If YY = Year And MM = "03" Then 'And DD >= "26" Then
        oMat01.Columns("Mm03").Editable = True
        oMat01.Columns("Mm04").Editable = True
    Else
        oMat01.Columns("Mm04").Editable = False
    End If

    If YY = Year And MM = "04" Then 'And DD >= "26" Then
        oMat01.Columns("Mm04").Editable = True
        oMat01.Columns("Mm05").Editable = True
    Else
        oMat01.Columns("Mm05").Editable = False
    End If

    If YY = Year And MM = "05" Then 'And DD >= "26" Then
        oMat01.Columns("Mm05").Editable = True
        oMat01.Columns("Mm06").Editable = True
    Else
        oMat01.Columns("Mm06").Editable = False
    End If

    If YY = Year And MM = "06" Then 'And DD >= "26" Then
        oMat01.Columns("Mm06").Editable = True
        oMat01.Columns("Mm07").Editable = True
    Else
        oMat01.Columns("Mm07").Editable = False
    End If

    If YY = Year And MM = "07" Then 'And DD >= "26" Then
        oMat01.Columns("Mm07").Editable = True
        oMat01.Columns("Mm08").Editable = True
    Else
        oMat01.Columns("Mm08").Editable = False
    End If

    If YY = Year And MM = "08" Then 'And DD >= "26" Then
        oMat01.Columns("Mm08").Editable = True
        oMat01.Columns("Mm09").Editable = True
    Else
        oMat01.Columns("Mm09").Editable = False
    End If

    If YY = Year And MM = "09" Then 'And DD >= "26" Then
        oMat01.Columns("Mm09").Editable = True
        oMat01.Columns("Mm10").Editable = True
    Else
        oMat01.Columns("Mm10").Editable = False
    End If

    If YY = Year And MM = "10" Then 'And DD >= "26" Then
        oMat01.Columns("Mm10").Editable = True
        oMat01.Columns("Mm11").Editable = True
    Else
        oMat01.Columns("Mm11").Editable = False
    End If

    If YY = Year And MM = "11" Then 'And DD >= "26" Then
        oMat01.Columns("Mm11").Editable = True
        oMat01.Columns("Mm12").Editable = True
    Else
        oMat01.Columns("Mm12").Editable = False
    End If
    
    Call oMat01.AutoResizeColumns
    
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub Add_MatrixRow(ByVal oRow As Long, Optional RowIserted As Boolean)
On Error GoTo Add_MatrixRow_Error
    If RowIserted = False Then '//행추가여부
        oDS_PS_MM012L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_MM012L.Offset = oRow
    oDS_PS_MM012L.setValue "U_LineNum", oRow, oRow + 1
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_MatrixRow_Error:
    MDC_Com.MDC_GF_Message "Add_MatrixRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    Dim DocNum          As String
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    ErrNum = 0

    '// Check
    Select Case True
        Case Trim(oDS_PS_MM012H.GetValue("U_Year", 0)) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case Trim(oDS_PS_MM012H.GetValue("U_BPLId", 0)) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
    End Select
    
    If oForm01.Mode = fm_ADD_MODE Then
        sQry = "Select Count(*) From [@PS_MM012H] Where U_BPLId = '" & Trim(oDS_PS_MM012H.GetValue("U_BPLId", 0)) & "' And U_Year = '" & Trim(oDS_PS_MM012H.GetValue("U_Year", 0)) & "'"
        oRecordSet.DoQuery sQry
        If oRecordSet.Fields(0).VALUE > 0 Then
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
        End If
    End If
        
    Set oRecordSet = Nothing
    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "년도는 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "사업장은 필수입력사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "데이터가 있습니다. 해당 사업장과 해당 년도로 데이터를 검색하세요.", "E"
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
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    For i = 0 To oMat01.VisualRowCount - 2
        Select Case True
            Case oDS_PS_MM012L.GetValue("U_PQDocNum", i) = ""
                ErrNum = 2
                GoTo MatrixSpaceLineDel_Error
        End Select
    Next i
    
    oMat01.LoadFromDataSource

    Set oRecordSet = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "" & i + 1 & "번 라인에 견적번호가 없습니다..행을 삭제해주세요.", "E"
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
        If Trim(oDS_PS_MM012L.GetValue("U_PQDocNum", i)) = "" Then
            oDS_PS_MM012L.RemoveRecord i   '// Mat01에 마지막라인(빈라인) 삭제
        End If
    Next i
    
    oMat01.LoadFromDataSource
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Delete_EmptyRow_Error:
    MDC_Com.MDC_GF_Message "Delete_EmptyRow_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function R3_Interface() As Boolean
On Error GoTo R3_Interface_Error
    Dim ErrNum&
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String

    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)

    
    If Qty_Check = False Then
        ErrNum = 3
        GoTo R3_Interface_Error
    End If
                        
    If SAP_Connection = False Then
        ErrNum = 1
        GoTo R3_Interface_Error
    End If
    
    If Send_RFC = False Then
        ErrNum = 2
        GoTo R3_Interface_Error
    Else                                 '2017.12.11 전송버튼 누를 경우 누른 시간이 화면에 표시됨.
        sQry = "update [@PS_MM012H] set u_tradate = convert(varchar(20), getdate(),120) Where U_BPLId = '" & Trim(oDS_PS_MM012H.GetValue("U_BPLId", 0)) & "' And U_Year = '" & Trim(oDS_PS_MM012H.GetValue("U_Year", 0)) & "'"
        oRecordSet.DoQuery sQry
    End If

    If Not (oSAP_Connection01.Connection Is Nothing) Then
        oSAP_Connection01.Connection.Logoff
        Set oSAP_Connection01 = Nothing
    End If

    R3_Interface = True
    
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
R3_Interface_Error:
'    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "Sap풍산 ERP 시스템(R/3)에 접속할 수 없습니다.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "Send_RFC에서 에러 발생", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "구매요청 수량 초과 오류", "E"
    Else
        MDC_Com.MDC_GF_Message "R3_Interface_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    R3_Interface = False
End Function

Private Function SAP_Connection() As Boolean
On Error GoTo SAP_Connection_Error
    Set oSAP_Connection01 = CreateObject("SAP.Functions")
    oSAP_Connection01.Connection.User = "ifuser"
    oSAP_Connection01.Connection.Password = "pdauser"
'    oSAP_Connection01.Connection.Client = "710"     '개발용
    oSAP_Connection01.Connection.Client = "210"
'    oSAP_Connection01.Connection.ApplicationServer = "192.1.11.3"   //개발할때 (풍산과 디버깅테스트 할때 사용)
'    oSAP_Connection01.Connection.ApplicationServer = "192.1.11.7"   '개발용
     oSAP_Connection01.Connection.ApplicationServer = "192.1.1.217"    '//운영
    oSAP_Connection01.Connection.language = "KO"
    oSAP_Connection01.Connection.SystemNumber = "01"
'    oSapConnection01.Connection.SystemNumber = "00" '192.1.11.3 일때는 "00"
    If Not oSAP_Connection01.Connection.Logon(0, True) Then
        GoTo SAP_Connection_Error
    End If
    
    SAP_Connection = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
SAP_Connection_Error:
    SAP_Connection = False
End Function

Private Function Send_RFC() As Boolean
On Error GoTo Send_RFC_Error
    Dim oFunction01         As Object
    
    Dim oRecordSet01        As SAPbobsCOM.Recordset
    Dim sQry                As String
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "Select * From [@PS_MM012L] Where Isnull(U_PQDocNum,'') <> '' And Code = '" & Trim(oDS_PS_MM012H.GetValue("Code", 0)) & "'"
    oRecordSet01.DoQuery sQry
    
    Do Until oRecordSet01.EOF
        Set oFunction01 = Nothing
        Set oFunction01 = oSAP_Connection01.Add("ZMM_INTF_GROUP2")
        
        oFunction01.Exports("I_BANFN") = Trim(oRecordSet01.Fields("U_E_BANFN").VALUE)
        oFunction01.Exports("I_BNFPO") = Trim(oRecordSet01.Fields("U_E_BNFPO").VALUE)
        oFunction01.Exports("I_LFDAT") = Trim(oRecordSet01.Fields("U_DueDate").VALUE)
        oFunction01.Exports("I_MEINS") = Trim(oRecordSet01.Fields("U_Unit").VALUE)
        oFunction01.Exports("I_MENGE") = Trim(oRecordSet01.Fields("U_Qty").VALUE)
        oFunction01.Exports("I_ZMM01") = Trim(oRecordSet01.Fields("U_Mm01").VALUE)
        oFunction01.Exports("I_ZMM02") = Trim(oRecordSet01.Fields("U_Mm02").VALUE)
        oFunction01.Exports("I_ZMM03") = Trim(oRecordSet01.Fields("U_Mm03").VALUE)
        oFunction01.Exports("I_ZMM04") = Trim(oRecordSet01.Fields("U_Mm04").VALUE)
        oFunction01.Exports("I_ZMM05") = Trim(oRecordSet01.Fields("U_Mm05").VALUE)
        oFunction01.Exports("I_ZMM06") = Trim(oRecordSet01.Fields("U_Mm06").VALUE)
        oFunction01.Exports("I_ZMM07") = Trim(oRecordSet01.Fields("U_Mm07").VALUE)
        oFunction01.Exports("I_ZMM08") = Trim(oRecordSet01.Fields("U_Mm08").VALUE)
        oFunction01.Exports("I_ZMM09") = Trim(oRecordSet01.Fields("U_Mm09").VALUE)
        oFunction01.Exports("I_ZMM10") = Trim(oRecordSet01.Fields("U_Mm10").VALUE)
        oFunction01.Exports("I_ZMM11") = Trim(oRecordSet01.Fields("U_Mm11").VALUE)
        oFunction01.Exports("I_ZMM12") = Trim(oRecordSet01.Fields("U_Mm12").VALUE)
        oFunction01.Exports("I_ZSUM") = Trim(oRecordSet01.Fields("U_MmTot").VALUE)
        
        
        oRecordSet01.MoveNext
        
        If Not (oFunction01.Call) Then
            MDC_Com.MDC_GF_Message "R/3 RFC 함수호출 중 오류발생", "E"
            GoTo Send_RFC_Error
        Else
        End If
    Loop
    
    
    
    Send_RFC = True
    Set oFunction01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Send_RFC_Error:
    If Not (oSAP_Connection01.Connection Is Nothing) Then
        oSAP_Connection01.Connection.Logoff
        Set oSAP_Connection01 = Nothing
    End If
    
    Send_RFC = False
    Set oFunction01 = Nothing
    Set oRecordSet01 = Nothing
End Function

Private Function Qty_Check() As Boolean
On Error GoTo Qty_Check_Error
    Dim ErrNum&
    Dim i As Long
    Dim ItemCode As String
    Dim Qty As Double
    Dim MmTot As Double
    
    For i = 1 To oMat01.VisualRowCount - 1
        Qty = oMat01.Columns("Qty").Cells(i).Specific.VALUE
        MmTot = oMat01.Columns("MmTot").Cells(i).Specific.VALUE
        
        If Qty < MmTot Then
           ItemCode = oMat01.Columns("ItemCode").Cells(i).Specific.VALUE
           ErrNum = 1
           GoTo Qty_Check_Error
        End If
    Next i
    
    Qty_Check = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Qty_Check_Error:
'    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "( " & ItemCode & " ) 구매요청 수량보다 초과할 수 없습니다. 확인바랍니다.", "E"
    Else
        MDC_Com.MDC_GF_Message "Qty_Check_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    Qty_Check = False
End Function
