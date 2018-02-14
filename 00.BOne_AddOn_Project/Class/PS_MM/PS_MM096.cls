VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM096"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM096.cls
'//  Module         : MM
'//  Description    : 원자재 불출 등록(포장사업팀)
'//  FormType       : PS_MM096
'//  Create Date    : 2013.04.24
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01 As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Public oMat02               As SAPbouiCOM.Matrix
Private oDS_PS_MM096H       As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_MM096L       As SAPbouiCOM.DBDataSource    '등록라인
Private oDS_PS_USERDS01     As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID      As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID       As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row       As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oClick_ColRow       As Long
Private oLast_Mode&

Private oBPLId  As String
Private oOrdGbn  As String

Private oQty As Long
Private oWeight As Currency

'****************************************************************************************************************
' .srf 파일로부터 폼을 로드한다.
'****************************************************************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM096.srf")
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
    
    oFormUniqueID01 = "PS_MM096_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
   
    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    oForm01.Freeze True
    Call CreateItems
    Call ComboBox_Setting
    Call Initialization
    Call FormClear
        
    Call FormItemEnabled
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), False        '// 복제
    oForm01.EnableMenu ("1284"), True         '// 취소
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
    Dim ErrNum&
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    
    Dim ItemCode$, ItemName$, Size$, Qty&, Weight As Currency, Unit$, RequestDate$, DueDate$, ItemType$, RequestNo$, UnWeight As Currency
    Dim RFC_Sender$
    Dim Calculate_Weight As Double
    Dim Seq&
    Dim WhsCode As String
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "1" Then
                    oBPLId = oForm01.Items("BPLId").Specific.VALUE
                    oOrdGbn = oForm01.Items("OrdGbn").Specific.VALUE
                    
                    If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
                        If HeaderSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        If MatrixSpaceLineDel = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                        
                        
                        If oForm01.Mode = fm_ADD_MODE Then

                            If Add_oInventoryGenExit(1) = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If

                        End If
                        
                        If oForm01.Mode = fm_UPDATE_MODE Then
                            Call Update_PP040L("UPDATE")
                        End If
                    End If
                End If
            Case et_KEY_DOWN: '//2
                If pval.ItemUID = "Mat01" Then
                    If pval.ColUID = "PItemCod" Then
                        If oMat01.Columns("PItemCod").Cells(pval.Row).Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ColUID = "CItemCod" Then
                        If oMat01.Columns("CItemCod").Cells(pval.Row).Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    End If
                End If
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
                If pval.ItemUID = "Mat02" Then
                    Dim TempForm01 As Variant
                   
                    If pval.ColUID = "ItemCode" Then
                        If oMat02.Columns("ItemCode").Cells(pval.Row).Specific.String <> "" Then
                            Set TempForm01 = New PS_MM002
                            Call TempForm01.LoadForm(oForm01.Items("BPLId").Specific.VALUE, oMat02.Columns("ItemCode").Cells(pval.Row).Specific.String)
                            BubbleEvent = False
                        Else
                            
                        End If
                    ElseIf pval.ColUID = "DocEntry" Then
                        If oMat02.Columns("WorkGbn").Cells(pval.Row).Specific.VALUE = "10" Then
                            Set TempForm01 = New PS_PP043
                            Call TempForm01.LoadForm(oMat02.Columns("DocEntry").Cells(pval.Row).Specific.String)
                            BubbleEvent = False
                        ElseIf oMat02.Columns("WorkGbn").Cells(pval.Row).Specific.VALUE = "30" Then
                            Set TempForm01 = New PS_MM138
                            Call TempForm01.LoadForm(oMat02.Columns("DocEntry").Cells(pval.Row).Specific.String)
                            BubbleEvent = False
                        ElseIf oMat02.Columns("WorkGbn").Cells(pval.Row).Specific.VALUE = "40" Then
                            Set TempForm01 = New PS_PP083
                            Call TempForm01.LoadForm(oMat02.Columns("DocEntry").Cells(pval.Row).Specific.String)
                            BubbleEvent = False
                            
                        End If
                    End If
                End If
                
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
                        oMat02.Clear
                        oDS_PS_USERDS01.Clear
                        oForm01.Mode = fm_OK_MODE
                        Call Sbo_Application.ActivateMenuItem("1282")
                        
                        
                        
                    End If
'                ElseIf pval.ItemUID = "Btn01" Then
'                    If oMat01.RowCount = 0 Then
'                        Add_MatrixRow 0, True
'                        oMat01.Columns("PItemCod").Editable = True
'                        oMat01.Columns("CItemCod").Editable = True
'                    End If
                End If
            Case et_KEY_DOWN: '//2
'et_COMBO_SELECT ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_COMBO_SELECT: '//5
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "BPLId" Or pval.ItemUID = "OrdGbn" Then
                        Call LoadData_Mat02
                    End If
                End If
'et_CLICK ///////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_CLICK: '//6
                If pval.ItemUID = "Mat02" And pval.ColUID = "LineNum" And oForm01.Mode = fm_ADD_MODE Then
                    oClick_ColRow = pval.Row
                    Call LoadData(oClick_ColRow)
                End If
'et_DOUBLE_CLICK ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_DOUBLE_CLICK: '//7
'                If pval.ItemUID = "Mat02" Then
'                    If oMat02.Columns("ItemCode").Cells(pval.Row).Specific.VALUE <> "" And oMat02.Columns("ItemName").Cells(pval.Row).Specific.VALUE <> "" And oMat02.Columns("DocEntry").Cells(pval.Row).Specific.VALUE <> "" Then
'                         Call LoadData
'                    Else
'                        MDC_Com.MDC_GF_Message "데이터가 불확실 합니다. 관리자에게 문의 바랍니다.", "E"
'                    End If
'                End If
            Case et_MATRIX_LINK_PRESSED '//8
'et_VALIDATE ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_VALIDATE: '//10
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "Mat01" Then
                        If pval.ColUID = "PItemCod" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                        ElseIf pval.ColUID = "CItemCod" Then
                            FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                       ElseIf pval.ColUID = "WhsCode" Then
                            oForm01.Freeze True
                            oMat01.FlushToDataSource
                            WhsCode = Trim(oDS_PS_MM096L.GetValue("U_WhsCode", pval.Row - 1))
                            If WhsCode = "" Then
                                oDS_PS_MM096L.setValue "U_CQty", pval.Row - 1, 0
                                oDS_PS_MM096L.setValue "U_CWeight", pval.Row - 1, 0
                                oMat01.LoadFromDataSource
                                oMat01.Columns("WhsCode").Cells(pval.Row).Click ct_Regular
                                MDC_Com.MDC_GF_Message "창고는 필수 선택입니다.", "E"
                            Else

                                FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                            End If
                            oForm01.Freeze False
                        End If
'                        ElseIf pval.ColUID = "IssueWt" Then
'                            oForm01.Freeze True
'                            oMat01.FlushToDataSource
'                            ItemCode = Trim(oDS_PS_MM096L.GetValue("U_CItemCod", pval.Row - 1))
'                            If ItemCode = "" Then
'                                oDS_PS_MM096L.setValue "U_IssueQty", pval.Row - 1, 0
'                                oDS_PS_MM096L.setValue "U_IssueWt", pval.Row - 1, 0
'                                oMat01.LoadFromDataSource
'                                oMat01.Columns("IssueQty").Cells(pval.Row).Click ct_Regular
'                                MDC_Com.MDC_GF_Message "불출할 원자재를 아래 매트릭스에서 먼저 선택해 주세요.", "E"
'                            End If
'                            oForm01.Freeze False
'                        End If
                    End If
                End If
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
'et_FORM_RESIZE /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_RESIZE '//20
                oForm01.Items("Mat02").Top = 81 '
                oForm01.Items("Mat02").Left = 6
                oForm01.Items("Mat02").Width = oForm01.Width - 18
                oForm01.Items("Mat02").Height = oForm01.Height / 3 ' - 280
'                oMat02.Columns("ItemCode").Width = 100
'                oMat02.Columns("ItemName").Width = 200
'                oMat02.Columns("WhsCode").Width = 100
'                oMat02.Columns("WhsName").Width = 100
'                oMat02.Columns("Qty").Width = 100
'                oMat02.Columns("Weight").Width = 100
                
                oForm01.Items("Mat01").Top = oForm01.Items("Mat02").Top + oForm01.Height / 3
                oForm01.Items("Mat01").Left = 6
                oForm01.Items("Mat01").Width = oForm01.Width - 18
                oForm01.Items("Mat01").Height = oForm01.Height / 4 ' - 280
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
'et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
                Set oDS_PS_MM096H = Nothing
                Set oDS_PS_MM096L = Nothing
                Set oDS_PS_USERDS01 = Nothing
        End Select
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Raise_ItemEvent_Error:
    Set ProgressBar01 = Nothing
    If oForm01 Is Nothing Then
    Else
        oForm01.Freeze False
    End If
    MDC_Com.MDC_GF_Message "Raise_ItemEvent_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    Dim i&
    Dim sReturnValue As Integer
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
'[1284:취소] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1284": '취소
                If Trim(oDS_PS_MM096H.GetValue("Canceled", 0)) = "N" Then
                    sReturnValue = Sbo_Application.MessageBox("해당 문서가 취소됩니다 계속하시겠습니까??", 1, "&확인", "&취소")
                    Select Case sReturnValue
                        Case 1:
                            If Add_oInventoryGenEntry(1) = False Then
                                BubbleEvent = False
                                Exit Sub
                            Else
                                If Add_oInventoryGenEntry(2) = False Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                        Case 2:
                            Sbo_Application.MessageBox ("실행이 취소되었습니다.")
                            BubbleEvent = False
                    End Select
                Else
                    MDC_Com.MDC_GF_Message "취소된 생산실적입니다. 확인하세요.", "E"
                    BubbleEvent = False
                    Exit Sub
                End If
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
                    oDS_PS_MM096L.RemoveRecord oDS_PS_MM096L.Size - 1       '// Mat01에 마지막라인(빈라인) 삭제
                    oMat01.Clear
                    oMat01.LoadFromDataSource
                    
'                    If oMat01.Columns("GADocLin").Cells(oMat01.RowCount).Specific.Value <> "" Then
'                        Call Add_MatrixRow(oMat01.RowCount, False)
'                    End If
                End If
'[1281:찾기] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1281": '찾기
                Call FormItemEnabled
                oForm01.Items("DocNum").Click ct_Regular
'[1282:추가] ////////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1282": '추가
                oForm01.Freeze True
                Call Initialization
                Call FormItemEnabled
                Call FormClear
                
'                oDS_PS_MM096H.setValue "U_BPLId", 0, oBPLId
'                oDS_PS_MM096H.setValue "U_OrdGbn", 0, oOrdGbn
                If Trim(oBPLId) = "" Then
                    oForm01.Items("BPLId").Specific.Select "0", psk_Index
                Else
                    oForm01.Items("BPLId").Specific.Select oBPLId, psk_ByValue
                End If
                
                If Trim(oOrdGbn) = "" Then
                    oForm01.Items("OrdGbn").Specific.Select "0", psk_Index
                Else
                    oForm01.Items("OrdGbn").Specific.Select oOrdGbn, psk_ByValue
                End If
                
                oForm01.Items("ExitGbn").Specific.Select "0", psk_Index
                
'                oForm01.Items("BPLId").Specific.Select "0", psk_Index
'                oForm01.Items("OrdGbn").Specific.Select "0", psk_Index
                oDS_PS_MM096H.setValue "U_DocDate", 0, Format(Now, "YYYYMMDD")
                oForm01.Freeze False
'[1288~1291:네비게이션] /////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                Call FormItemEnabled
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
                If oForm01.Items("ExitGbn").Specific.VALUE <> "9" Then
                    If Add_oInventoryGenExit(2) = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If
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
    Set oDS_PS_MM096H = oForm01.DataSources.DBDataSources("@PS_MM095H")
    Set oDS_PS_MM096L = oForm01.DataSources.DBDataSources("@PS_MM095L")
    Set oDS_PS_USERDS01 = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    '// 메트릭스 개체 할당
    Set oMat01 = oForm01.Items("Mat01").Specific
    Set oMat02 = oForm01.Items("Mat02").Specific
    
    oDS_PS_MM096H.setValue "U_DocDate", 0, Format(Now, "yyyymmdd")
    
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
    
    Set oCombo = oForm01.Items("ExitGbn").Specific
    oCombo.ValidValues.Add "1", "정상출고"
    oCombo.ValidValues.Add "9", "잔재사용"
    oCombo.Select "0", psk_Index
    
    
    '//작업구분
    Set oCombo = oForm01.Items("OrdGbn").Specific
    sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Where U_PudYN = 'Y' and Code in ('108','109')  Order by Code"
    oRecordSet01.DoQuery sQry
    Do Until oRecordSet01.EOF
        oCombo.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    oCombo.Select "0", psk_Index
    
    '//작업구분
    oMat01.Columns("WorkGbn").ValidValues.Add "10", "자가"
    oMat01.Columns("WorkGbn").ValidValues.Add "30", "외주"
    oMat01.Columns("WorkGbn").ValidValues.Add "40", "실적"
    
    '//작업구분
    oMat02.Columns("WorkGbn").ValidValues.Add "10", "자가"
    oMat02.Columns("WorkGbn").ValidValues.Add "30", "외주"
    oMat02.Columns("WorkGbn").ValidValues.Add "40", "실적"
    
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
    
    If Trim(oBPLId) = "" Then
        oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    Else
        oCombo.Select oBPLId, psk_ByValue
    End If
    
    If oMat01.RowCount = 0 Then
        Add_MatrixRow 0, True
    End If
    '//아이디별 사번 세팅
'    oForm01.Items("CntcCode").Specific.Value = MDC_PS_Common.User_MSTCOD
    
    '//아이디별 부서 세팅
'    Set oCombo = oForm01.Items("DeptCode").Specific
'    oCombo.Select MDC_PS_Common.User_DeptCode, psk_ByValue
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
        oForm01.Items("DocNum").Enabled = False

        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("OrdGbn").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        
        oMat01.Columns("IssueYN").Editable = True
        oMat01.Columns("IssueQty").Editable = True
        oMat01.Columns("IssueWt").Editable = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("DocNum").Enabled = True

        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("OrdGbn").Enabled = True
        oForm01.Items("DocDate").Enabled = True
        
        oMat01.Columns("IssueYN").Editable = False
        oMat01.Columns("IssueQty").Editable = False
        oMat01.Columns("IssueWt").Editable = False
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        oForm01.Items("DocNum").Enabled = False
        
        oForm01.Items("BPLId").Enabled = False
        oForm01.Items("OrdGbn").Enabled = False
        oForm01.Items("DocDate").Enabled = False

        oMat01.Columns("IssueYN").Editable = False
        oMat01.Columns("IssueQty").Editable = False
        oMat01.Columns("IssueWt").Editable = False
    End If
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
FormItemEnabled_Error:
    MDC_Com.MDC_GF_Message "FormItemEnabled_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM095'", "")
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
        oDS_PS_MM096L.InsertRecord (oRow)
    End If
    oMat01.AddRow
    oDS_PS_MM096L.Offset = oRow
    oDS_PS_MM096L.setValue "U_LineNum", oRow, oRow + 1
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
    Dim WhsCode         As String
        
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "Mat01"
            If oCol = "PItemCod" Then
                oForm01.Freeze True
                If oForm01.Mode = fm_ADD_MODE Then
                    If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And Trim(oMat01.Columns("PItemCod").Cells(oRow).Specific.VALUE) <> "" Then
                        oMat01.FlushToDataSource
                        Add_MatrixRow oMat01.RowCount, False
                    End If
                End If
                
                sQry = "Select ItemName  From OITM Where ItemCode = '" & Trim(oMat01.Columns("PItemCod").Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
                oMat01.Columns("PItemNam").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                
                oMat01.Columns("DocDate").Cells(oRow).Specific.VALUE = oForm01.Items("DocDate").Specific.VALUE
                oMat01.Columns("WorkGbn").Cells(oRow).Specific.Select "10"
                
                oMat01.Columns("PItemCod").Cells(oRow).Click ct_Regular
                oForm01.Freeze False
            ElseIf oCol = "CItemCod" Then
                oForm01.Freeze True
                
                WhsCode = "10" + oForm01.Items("BPLId").Specific.VALUE
                
                sQry = "Select ItemName  From OITM Where ItemCode = '" & Trim(oMat01.Columns("CItemCod").Cells(oRow).Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
                oMat01.Columns("CItemNam").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                
                sQry = "Select Qty = Case When Onhand = 0 Then 0 Else U_Qty End, Weight = Onhand  From OITW"
                sQry = sQry + " Where ItemCode = '" & Trim(oMat01.Columns("CItemCod").Cells(oRow).Specific.VALUE) & "' and WhsCode = '" & WhsCode & "'"
                oRecordSet01.DoQuery sQry
                oMat01.Columns("CQty").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                oMat01.Columns("CWeight").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(1).VALUE)
                
                oMat01.Columns("WhsCode").Cells(oRow).Specific.VALUE = WhsCode
                
                
                oMat01.Columns("CItemCod").Cells(oRow).Click ct_Regular
                oForm01.Freeze False
            ElseIf oCol = "WhsCode" Then
                oForm01.Freeze True
                
                WhsCode = oMat01.Columns("WhsCode").Cells(oRow).Specific.VALUE
                
                
                sQry = "Select Qty = Case When Onhand = 0 Then 0 Else U_Qty End, Weight = Onhand  From OITW"
                sQry = sQry + " Where ItemCode = '" & Trim(oMat01.Columns("CItemCod").Cells(oRow).Specific.VALUE) & "' and WhsCode = '" & WhsCode & "'"
                oRecordSet01.DoQuery sQry
                oMat01.Columns("CQty").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                oMat01.Columns("CWeight").Cells(oRow).Specific.VALUE = Trim(oRecordSet01.Fields(1).VALUE)
                
                
                oMat01.Columns("WhsCode").Cells(oRow).Click ct_Regular
                oForm01.Freeze False
            End If
    End Select

'    Set oRecordset = Nothing
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Exit Sub
FlushToItemValue_Error:
    oForm01.Freeze False
    MDC_Com.MDC_GF_Message "FlushToItemValue_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum          As Integer
    
    ErrNum = 0

    '// Check
    Select Case True
        Case oDS_PS_MM096H.GetValue("U_BPLId", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM096H.GetValue("U_OrdGbn", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_MM096H.GetValue("U_DocDate", 0) = ""
            ErrNum = 3
            GoTo HeaderSpaceLineDel_Error
    End Select

    HeaderSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "사업장은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "작업구분은 필수사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "전기일은 필수사항입니다. 확인하세요.", "E"
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
    Dim ItemCode        As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)

    ErrNum = 0

    '// 라인
    If oMat01.VisualRowCount = 0 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    oForm01.Freeze True

    oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 1
        '원재료에 대한 제품코드가 누락되었을때
        If oDS_PS_MM096L.GetValue("U_PItemCod", i) = "" And oDS_PS_MM096L.GetValue("U_CItemCod", i) <> "" Then
            ErrNum = 2
            GoTo MatrixSpaceLineDel_Error
        End If
        
        '원재료의 불출대상 제품코드가 정상인지 check
        ItemCode = oDS_PS_MM096L.GetValue("U_PItemCod", i)
        
        sQry = "Select Cnt = Count(*) From OITM Where ItmsGrpCod = '102' And ItemCode = '" & ItemCode & "'"
        
        oRecordSet01.DoQuery sQry
        
        If oRecordSet01.Fields(0).VALUE = 0 Then
            ErrNum = 3
            GoTo MatrixSpaceLineDel_Error
        End If
        
    Next i

    
    oMat01.LoadFromDataSource

    oMat01.FlushToDataSource
    For i = 0 To oMat01.VisualRowCount - 1
        oDS_PS_MM096L.setValue "U_LineNum", i, i + 1
    Next i
    oMat01.LoadFromDataSource
    
    oForm01.Freeze False
    
    Set oRecordSet01 = Nothing
    MatrixSpaceLineDel = True
Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "제품코드가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "제품코드를 확인하세요. 원재료 불출대상 제품이 아닙니다.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Public Sub LoadData(sRow As Long)
On Error GoTo LoadData_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID$, WorkGbn$, DocEntry$, LineId$
    
    oForm01.Items("DocDate").Specific.VALUE = Format(oMat02.Columns("DocDate").Cells(sRow).Specific.VALUE, "YYYYMMDD")
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    WorkGbn = Trim(oMat02.Columns("WorkGbn").Cells(sRow).Specific.VALUE)
    DocEntry = Trim(oMat02.Columns("DocEntry").Cells(sRow).Specific.VALUE)
    LineId = Trim(oMat02.Columns("LineId").Cells(sRow).Specific.VALUE)
    
    sQry = "EXEC [PS_MM096_02] '" & BPLID & "','" & WorkGbn & "', '" & DocEntry & "', '" & LineId & "'"
    oRecordSet01.DoQuery sQry
    
    oMat01.Clear
    oDS_PS_MM096L.Clear
    
    If (oRecordSet01.RecordCount = 0) Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_MM096L.Size Then
            oDS_PS_MM096L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_MM096L.Offset = i
        oDS_PS_MM096L.setValue "U_LineNum", i, i + 1
        oDS_PS_MM096L.setValue "U_DocDate", i, Format(Trim(oRecordSet01.Fields("DocDate").VALUE), "YYYYMMDD")
        oDS_PS_MM096L.setValue "U_WorkGbn", i, Trim(oRecordSet01.Fields("WorkGbn").VALUE)
        oDS_PS_MM096L.setValue "U_OrdNum", i, Trim(oRecordSet01.Fields("OrdNum").VALUE)
        oDS_PS_MM096L.setValue "U_CpCode", i, Trim(oRecordSet01.Fields("CpCode").VALUE)
        oDS_PS_MM096L.setValue "U_CpName", i, Trim(oRecordSet01.Fields("CpName").VALUE)
        oDS_PS_MM096L.setValue "U_PItemCod", i, Trim(oRecordSet01.Fields("PItemCod").VALUE)
        oDS_PS_MM096L.setValue "U_PItemNam", i, Trim(oRecordSet01.Fields("PItemNam").VALUE)
        oDS_PS_MM096L.setValue "U_PQty", i, Trim(oRecordSet01.Fields("PQty").VALUE)
        oDS_PS_MM096L.setValue "U_PWeight", i, Trim(oRecordSet01.Fields("PWeight").VALUE)
        oDS_PS_MM096L.setValue "U_CItemCod", i, Trim(oRecordSet01.Fields("CItemCod").VALUE)
        oDS_PS_MM096L.setValue "U_CItemNam", i, Trim(oRecordSet01.Fields("CItemNam").VALUE)
        oDS_PS_MM096L.setValue "U_CutCount", i, Trim(oRecordSet01.Fields("CutCount").VALUE)
        oDS_PS_MM096L.setValue "U_IssueYN", i, Trim(oRecordSet01.Fields("IssueYN").VALUE)
        oDS_PS_MM096L.setValue "U_IssueQty", i, Trim(oRecordSet01.Fields("IssueQty").VALUE)
        oDS_PS_MM096L.setValue "U_IssueWt", i, Trim(oRecordSet01.Fields("IssueWt").VALUE)
        oDS_PS_MM096L.setValue "U_PreQty", i, Trim(oRecordSet01.Fields("PreQty").VALUE)
        oDS_PS_MM096L.setValue "U_PreWt", i, Trim(oRecordSet01.Fields("PreWt").VALUE)
        oDS_PS_MM096L.setValue "U_WhsCode", i, Trim(oRecordSet01.Fields("WhsCode").VALUE)
        oDS_PS_MM096L.setValue "U_CQty", i, Trim(oRecordSet01.Fields("CQty").VALUE)
        oDS_PS_MM096L.setValue "U_CWeight", i, Trim(oRecordSet01.Fields("CWeight").VALUE)
        oDS_PS_MM096L.setValue "U_PP040Doc", i, Trim(oRecordSet01.Fields("PP040Doc").VALUE)
        oDS_PS_MM096L.setValue "U_PP040Lin", i, Trim(oRecordSet01.Fields("PP040Lin").VALUE)

        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat01.LoadFromDataSource
'    oMat01.AutoResizeColumns
    ProgBar01.Stop
    oForm01.Freeze False
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData_Error:
    ProgBar01.Stop
    oForm01.Freeze False
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub LoadData_Mat02() '(sRow As Long)
On Error GoTo LoadData_Mat02_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID$, OrdGbn$
    
    BPLID = Trim(oForm01.Items("BPLId").Specific.VALUE)
    OrdGbn = Trim(oForm01.Items("OrdGbn").Specific.VALUE)
           
    sQry = "EXEC [PS_MM096_01] '" & BPLID & "','" & OrdGbn & "'"
    
    oRecordSet01.DoQuery sQry
    
    oMat02.Clear

    oDS_PS_USERDS01.Clear
    
    If (oRecordSet01.RecordCount = 0) Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_USERDS01.Size Then
            oDS_PS_USERDS01.InsertRecord (i)
        End If
        
        oMat02.AddRow
        oDS_PS_USERDS01.Offset = i
        oDS_PS_USERDS01.setValue "U_LineNum", i, i + 1
        oDS_PS_USERDS01.setValue "U_ColReg01", i, Trim(oRecordSet01.Fields("WorkGbn").VALUE)
        oDS_PS_USERDS01.setValue "U_ColReg02", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)
        oDS_PS_USERDS01.setValue "U_ColReg03", i, Trim(oRecordSet01.Fields("LineId").VALUE)
        oDS_PS_USERDS01.setValue "U_ColReg04", i, Trim(oRecordSet01.Fields("DocDate").VALUE)
        oDS_PS_USERDS01.setValue "U_ColReg05", i, Trim(oRecordSet01.Fields("CpName").VALUE)
        oDS_PS_USERDS01.setValue "U_ColReg06", i, Trim(oRecordSet01.Fields("ItemCode").VALUE)
        oDS_PS_USERDS01.setValue "U_ColReg07", i, Trim(oRecordSet01.Fields("ItemName").VALUE)
        oDS_PS_USERDS01.setValue "U_ColNum01", i, Trim(oRecordSet01.Fields("Qty").VALUE)
        
        oRecordSet01.MoveNext
        ProgBar01.VALUE = ProgBar01.VALUE + 1
        ProgBar01.Text = ProgBar01.VALUE & "/" & oRecordSet01.RecordCount & "건 조회중...!"
    Next
    oMat02.LoadFromDataSource
    oMat02.AutoResizeColumns
    ProgBar01.Stop
    oForm01.Freeze False
  
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
LoadData_Mat02_Error:
    ProgBar01.Stop
    oForm01.Freeze False
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData_Mat02_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Public Sub Update_PP040L(sStatus As String)
On Error GoTo Update_PP040L_Error
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim DocDate$, DocNum$, CItemCod$, IssueQty As Long, IssueWt As Currency, WhsCode$, IssueYN$
    Dim PP040Doc$, PP040Lin$, WorkGbn$
        
    oForm01.Freeze True
    If sStatus = "ADD" Then
        IssueYN = "Y"
    ElseIf sStatus = "CANCEL" Then
        IssueYN = "N"
    End If
    
    
    If oMat01.VisualRowCount > 0 Then
        '//
        CItemCod = oMat01.Columns("CItemCod").Cells(1).Specific.VALUE
        IssueQty = oMat01.Columns("IssueQty").Cells(1).Specific.VALUE
        IssueWt = oMat01.Columns("IssueWt").Cells(1).Specific.VALUE
        WhsCode = oMat01.Columns("WhsCode").Cells(1).Specific.VALUE
        WorkGbn = oMat01.Columns("WorkGbn").Cells(1).Specific.VALUE
        
        If CItemCod <> "" And WhsCode <> "" Then
            If Trim(oMat01.Columns("IssueYN").Cells(1).Specific.VALUE) = "Y" Then
                PP040Doc = Trim(oMat01.Columns("PP040Doc").Cells(1).Specific.VALUE)
                PP040Lin = Trim(oMat01.Columns("PP040Lin").Cells(1).Specific.VALUE)
                If WorkGbn = "10" Then
                    '실동입력
                    sQry = "Update [@PS_PP040L] Set U_IssueYN = '" & IssueYN & "' Where DocEntry = '" & PP040Doc & "' And LineId = '" & PP040Lin & "'"
                ElseIf WorkGbn = "30" Then
                    '외주입고
                    sQry = "Update [@PS_MM138H] Set U_IssueYN = '" & IssueYN & "' Where DocEntry = '" & PP040Doc & "'"
                ElseIf WorkGbn = "40" Then
                    '생산완료
                    sQry = "Update [@PS_PP083L] Set U_IssueYN = '" & IssueYN & "' Where DocEntry = '" & PP040Doc & "' And LineId = '" & PP040Lin & "'"
                End If
                    
                oRecordSet01.DoQuery sQry
            End If
        End If
    End If
    
'    For i = 0 To oMat01.VisualRowCount - 1
'        CItemCod = oMat01.Columns("CItemCod").Cells(i + 1).Specific.VALUE
'        IssueQty = oMat01.Columns("IssueQty").Cells(i + 1).Specific.VALUE
'        IssueWt = oMat01.Columns("IssueWt").Cells(i + 1).Specific.VALUE
'        WhsCode = oMat01.Columns("WhsCode").Cells(i + 1).Specific.VALUE
'        If CItemCod <> "" And IssueQty >= 0 And IssueWt <> 0 And WhsCode <> "" Then
'            If Trim(oMat01.Columns("IssueYN").Cells(i + 1).Specific.VALUE) = "Y" Then
'                PP040Doc = Trim(oMat01.Columns("PP040Doc").Cells(i + 1).Specific.VALUE)
'                PP040Lin = Trim(oMat01.Columns("PP040Lin").Cells(i + 1).Specific.VALUE)
'                sQry = "Update [@PS_PP040L] Set U_IssueYN = '" & IssueYN & "' Where DocEntry = '" & PP040Doc & "' And LineId = '" & PP040Lin & "'"
'                oRecordSet01.DoQuery sQry
'            End If
'        End If
'    Next i
    oForm01.Freeze False
  
    Set oRecordSet01 = Nothing
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Update_PP040L_Error:
    oForm01.Freeze False
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "Update_PP040L_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Function Add_oInventoryGenExit(ChkType As Integer) As Boolean
On Error GoTo Add_oInventoryGenExit_Error
    Dim DI_oInventoryGenExit As SAPbobsCOM.Documents         '//재고출고 문서 객체
    
    Dim i&, j&
    Dim ErrNum          As Integer
    Dim ErrCode         As Long
    Dim ErrMsg          As String
    Dim RetVal&
    
    Dim DocDate$, DocNum$, CItemCod$, IssueQty As Long, IssueWt As Currency, WhsCode$
    Dim sDocEntry$
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
     
    
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat01.FlushToDataSource
    
    DocDate = Format(oDS_PS_MM096H.GetValue("U_DocDate", 0), "0000-00-00")
    If ChkType = 2 Then Call FormClear
    DocNum = Trim(oDS_PS_MM096H.GetValue("DocNum", 0))
    
    If oForm01.Items("ExitGbn").Specific.VALUE <> "9" Then
        If Trim(oDS_PS_MM096H.GetValue("U_OIGEDoc", 0)) = "" Then
            Sbo_Company.StartTransaction
            Set DI_oInventoryGenExit = Nothing
            Set DI_oInventoryGenExit = Sbo_Company.GetBusinessObject(oInventoryGenExit)
            
            With DI_oInventoryGenExit
                .DocDate = DocDate
                .TaxDate = DocDate
                .Comments = "원재료 불출 등록(" & DocNum & ") 출고"
                
                j = 0
                For i = 0 To oMat01.VisualRowCount - 1
                    CItemCod = oMat01.Columns("CItemCod").Cells(i + 1).Specific.VALUE
                    IssueQty = oMat01.Columns("IssueQty").Cells(i + 1).Specific.VALUE
                    IssueWt = oMat01.Columns("IssueWt").Cells(i + 1).Specific.VALUE
                    WhsCode = oMat01.Columns("WhsCode").Cells(i + 1).Specific.VALUE
                
                    If CItemCod <> "" And IssueQty <> 0 And IssueWt <> 0 And WhsCode <> "" Then
                        If j > 0 Then .Lines.Add
                        .Lines.SetCurrentLine j
                        .Lines.ItemCode = CItemCod
                        .Lines.WarehouseCode = WhsCode
            '            .Lines.AccountCode = Trim(sAccount)
                        .Lines.Quantity = IssueWt
                        .Lines.UserFields("U_Qty").VALUE = IssueQty
                        
                        j = j + 1
                    End If
                Next i
            End With
                
            '// 완료
            RetVal = DI_oInventoryGenExit.Add
            If (0 <> RetVal) Then
                Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
                ErrNum = 1
                GoTo Add_oInventoryGenExit_Error
            Else
                If ChkType = 1 Then
                    Sbo_Company.EndTransaction wf_RollBack
                ElseIf ChkType = 2 Then
                    Sbo_Company.GetNewObjectCode sDocEntry
                    oDS_PS_MM096H.setValue "U_OIGEDoc", 0, sDocEntry
                    Call Update_PP040L("ADD")
                    Sbo_Company.EndTransaction wf_Commit
                End If
            End If
            
        End If
    Else
        '잔재사용
        Call Update_PP040L("ADD")
    End If
    Set oRecordSet = Nothing
    Set DI_oInventoryGenExit = Nothing
    Add_oInventoryGenExit = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_oInventoryGenExit_Error:
    Set oRecordSet = Nothing
    Set DI_oInventoryGenExit = Nothing
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Add_oInventoryGenExit = False
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "Add_oInventoryGenExit_Error:" & ErrCode & " - " & ErrMsg, "E"
    Else
        MDC_Com.MDC_GF_Message "Add_oInventoryGenExit_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Function

Private Function Add_oInventoryGenEntry(ChkType As Integer) As Boolean
On Error GoTo Add_oInventoryGenEntry_Error
    Dim DI_oInventoryGenEntry As SAPbobsCOM.Documents         '//재고출고 문서 객체
    
    Dim i&, j&
    Dim ErrNum          As Integer
    Dim ErrCode         As Long
    Dim ErrMsg          As String
    Dim RetVal&
    
    Dim DocDate$, DocNum$, CItemCod$, IssueQty As Long, IssueWt As Currency, WhsCode$
    Dim sDocEntry$
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim OIGEDoc As String
    
    
    
    Set oRecordSet = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat01.FlushToDataSource
    
    DocDate = Format(oDS_PS_MM096H.GetValue("U_DocDate", 0), "0000-00-00")
    DocNum = Trim(oDS_PS_MM096H.GetValue("DocNum", 0))
    OIGEDoc = Trim(oDS_PS_MM096H.GetValue("U_OIGEDoc", 0))
    
    If oForm01.Items("ExitGbn").Specific.VALUE <> "9" Then
    
        If Trim(OIGEDoc) <> "" And Trim(oDS_PS_MM096H.GetValue("U_OIGNDoc", 0)) = "" Then
            Sbo_Company.StartTransaction
            Set DI_oInventoryGenEntry = Nothing
            Set DI_oInventoryGenEntry = Sbo_Company.GetBusinessObject(oInventoryGenEntry)
            
            With DI_oInventoryGenEntry
                .DocDate = DocDate
                .TaxDate = DocDate
                .Comments = "원재료 불출 등록 출고 취소 (" & DocNum & ") 입고"
                
                .UserFields("U_CancDoc").VALUE = OIGEDoc '//입고취소 문서번호
                
                j = 0
                For i = 0 To oMat01.VisualRowCount - 1
                    CItemCod = oMat01.Columns("CItemCod").Cells(i + 1).Specific.VALUE
                    IssueQty = oMat01.Columns("IssueQty").Cells(i + 1).Specific.VALUE
                    IssueWt = oMat01.Columns("IssueWt").Cells(i + 1).Specific.VALUE
                    WhsCode = oMat01.Columns("WhsCode").Cells(i + 1).Specific.VALUE
                
                    If CItemCod <> "" And IssueQty <> 0 And IssueWt <> 0 And WhsCode <> "" Then
                        If j > 0 Then .Lines.Add
                        .Lines.SetCurrentLine j
                        .Lines.ItemCode = CItemCod
                        .Lines.WarehouseCode = WhsCode
            '            .Lines.AccountCode = Trim(sAccount)
                        .Lines.Quantity = IssueWt
                        .Lines.UserFields("U_Qty").VALUE = IssueQty
                        
                        j = j + 1
                    End If
                Next i
            End With
                
            '// 완료
            RetVal = DI_oInventoryGenEntry.Add
            If (0 <> RetVal) Then
                Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
                ErrNum = 1
                GoTo Add_oInventoryGenEntry_Error
            Else
                If ChkType = 1 Then
                    Sbo_Company.EndTransaction wf_RollBack
                ElseIf ChkType = 2 Then
                    Sbo_Company.GetNewObjectCode sDocEntry
                    oDS_PS_MM096H.setValue "U_OIGNDoc", 0, sDocEntry
                    Call Update_PP040L("CANCEL")
                    Sbo_Company.EndTransaction wf_Commit
                End If
            End If
        End If
    Else
        Call Update_PP040L("CANCEL")
    End If
    Set oRecordSet = Nothing
    Set DI_oInventoryGenEntry = Nothing
    Add_oInventoryGenEntry = True
    Exit Function
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Add_oInventoryGenEntry_Error:
    Set oRecordSet = Nothing
    Set DI_oInventoryGenEntry = Nothing
    If Sbo_Company.InTransaction Then Sbo_Company.EndTransaction wf_RollBack
    Add_oInventoryGenEntry = False
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "Add_oInventoryGenEntry_Error:" & ErrCode & " - " & ErrMsg, "E"
    Else
        MDC_Com.MDC_GF_Message "Add_oInventoryGenEntry_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Function


