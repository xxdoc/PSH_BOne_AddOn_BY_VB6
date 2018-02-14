VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM013"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************************************
'//  File           : PS_MM013.cls
'//  Module         : MM
'//  Description    : 자재 순환품 관리 Sub Form
'//  FormType       : PS_MM013
'//  Create Date    : 2012.10.11
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Company        : Poongsan Holdings
'****************************************************************************************************************
Option Explicit

Public oFormUniqueID01          As String
Public oForm01                  As SAPbouiCOM.Form
Public oMat01                   As SAPbouiCOM.Matrix
Private oDS_PS_MM013L           As SAPbouiCOM.DBDataSource '등록라인

Public oBaseForm01              As SAPbouiCOM.Form '//부모폼
Public oBaseItemUID01           As String
Public oBaseColUID01            As String
Public oBaseColRow01            As Long
Public oBaseSelectedLineNum01   As Long

Private oLast_Item_UID          As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID           As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row           As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(ByRef oForm02 As SAPbouiCOM.Form, ByVal oItemUID02 As String, ByVal oColUID02 As String, ByVal oColRow02 As Long, ByVal SelectedLineNum As Long)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM013.srf")
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
    
    oFormUniqueID01 = "PS_MM013_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_OK_MODE
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
    
    Dim Chk$
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
'et_ITEM_PRESSED ////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_ITEM_PRESSED: '//1
                If pval.ItemUID = "Btn01" Then
                    Call SetBaseForm
                End If
            Case et_KEY_DOWN: '//2
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
            Case et_ITEM_PRESSED: '//1
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
                If pval.ItemUID = "Mat01" And pval.Row = "0" And pval.ColUID = "Check" Then
                    
                    oForm01.Freeze True
                    oMat01.FlushToDataSource
                    If Trim(oDS_PS_MM013L.GetValue("U_ColReg20", 0)) = "" Or Trim(oDS_PS_MM013L.GetValue("U_ColReg20", 0)) = "N" Then
                        Chk = "Y"
                    ElseIf Trim(oDS_PS_MM013L.GetValue("U_ColReg20", 0)) = "Y" Then
                        Chk = "N"
                    End If
                    For i = 0 To oMat01.VisualRowCount - 1
                        oDS_PS_MM013L.setValue "U_ColReg20", i, Chk
                    Next i
                    oMat01.LoadFromDataSource
                    oForm01.Freeze False
                End If
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            Case et_MATRIX_LOAD: '//11
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            
                Call oMat01.AutoResizeColumns
                
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
            Case et_LOST_FOCUS: '//4
'et_FORM_UNLOAD /////////////''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
                Set oDS_PS_MM013L = Nothing
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
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_MM013L = oForm01.DataSources.DBDataSources("@PS_USERDS01")
    
    Set oMat01 = oForm01.Items("Mat01").Specific
    oMat01.SelectionMode = ms_NotSupported
    oMat01.AutoResizeColumns
     
    Set oRecordSet01 = Nothing
    
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
    Dim i As Integer
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Dim BPLID$, Year$
        
    BPLID = Trim(oBaseForm01.Items("BPLId").Specific.VALUE)
    Year = Trim(oBaseForm01.Items("Year").Specific.VALUE)
           
'    sQry = "Select b.DocEntry, b.LineId, b.U_E_BANFN, b.U_E_BNFPO, DueDate = Convert(char(8),c.U_DueDate,112), b.U_ItemCode, b.U_ItemName, b.U_OutSize, b.U_OutUnit, b.U_Weight "
'    sQry = sQry & "From [@PS_MM010H] a Join [@PS_MM010L] b On a.DocEntry = b.DocEntry And a.Canceled = 'N' And a.Status = 'O' "
'    sQry = sQry & " Inner join [@PS_MM005H] c  On b.U_CGNo = c.U_CgNum And c.Status = 'O' "
'    sQry = sQry & "Where a.U_BPLId = '" & BPLID & "' And Convert(char(4),a.U_DocDate,112) = '" & Year & "' And a.U_RotateYN = 'Y' And a.Canceled = 'N'"
'    sQry = sQry & " And Not Exists (Select * From [@PS_MM012H] t Inner Join [@PS_MM012L] t1 On a.DocEntry = b.DocEntry"
'    sQry = sQry & " And t1.U_PQDocNum = b.DocEntry and t1.U_PQLinNum = b.LineId)"
'    sQry = sQry & " Order by b.DocEntry, b.LineId "
    
    sQry = "EXEC PS_MM013_01 '" & BPLID & "','" & Year & "'"
    
    oRecordSet01.DoQuery sQry
            
    oMat01.Clear
    oDS_PS_MM013L.Clear
    
    If oRecordSet01.RecordCount = 0 Then
        MDC_Com.MDC_GF_Message "조회 결과가 없습니다. 확인하세요.:" & Err.Number & " - " & Err.Description, "W"
        Set oRecordSet01 = Nothing
        Exit Sub
    End If
    
    oForm01.Freeze True
    Dim ProgBar01 As SAPbouiCOM.ProgressBar
    Set ProgBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, False)
    
    For i = 0 To oRecordSet01.RecordCount - 1
        If i + 1 > oDS_PS_MM013L.Size Then
            oDS_PS_MM013L.InsertRecord (i)
        End If
        
        oMat01.AddRow
        oDS_PS_MM013L.Offset = i
        oDS_PS_MM013L.setValue "U_LineNum", i, i + 1
        oDS_PS_MM013L.setValue "U_ColReg01", i, Trim(oRecordSet01.Fields("DocEntry").VALUE)
        oDS_PS_MM013L.setValue "U_ColReg02", i, Trim(oRecordSet01.Fields("LineId").VALUE)
        oDS_PS_MM013L.setValue "U_ColReg03", i, Trim(oRecordSet01.Fields("U_E_BANFN").VALUE)
        oDS_PS_MM013L.setValue "U_ColReg04", i, Trim(oRecordSet01.Fields("U_E_BNFPO").VALUE)
        oDS_PS_MM013L.setValue "U_ColReg05", i, Trim(oRecordSet01.Fields("DueDate").VALUE)
        oDS_PS_MM013L.setValue "U_ColReg06", i, Trim(oRecordSet01.Fields("U_ItemCode").VALUE)
        oDS_PS_MM013L.setValue "U_ColReg07", i, Trim(oRecordSet01.Fields("U_ItemName").VALUE)
        oDS_PS_MM013L.setValue "U_ColReg08", i, Trim(oRecordSet01.Fields("U_OutSize").VALUE)
        oDS_PS_MM013L.setValue "U_ColReg09", i, Trim(oRecordSet01.Fields("U_OutUnit").VALUE)
        oDS_PS_MM013L.setValue "U_ColNum01", i, oRecordSet01.Fields("U_Weight").VALUE
        oDS_PS_MM013L.setValue "U_ColReg10", i, Trim(oRecordSet01.Fields("U_CardCode").VALUE)
        oDS_PS_MM013L.setValue "U_ColReg11", i, Trim(oRecordSet01.Fields("U_CardName").VALUE)
        
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
LoadData01_Error:    ProgBar01.Stop
    Set ProgBar01 = Nothing
    Set oRecordSet01 = Nothing
    MDC_Com.MDC_GF_Message "LoadData01_Error:" & Err.Number & " - " & Err.Description, "E"
End Sub

Private Sub SetBaseForm()
On Error GoTo SetBaseForm_Error
    Dim i As Long
    Dim j As Long
    Dim ErrNum As Long
    Dim sRow As Long
    
    Dim oBaseMat01 As SAPbouiCOM.Matrix
'    Dim oBaseDS_POR1 As SAPbouiCOM.DBDataSource
    Set oBaseMat01 = oBaseForm01.Items("Mat01").Specific
'    Set oBaseDS_POR1 = oBaseForm01.DataSources.DBDataSources("POR1")
    
    oBaseForm01.Freeze True
'    oBaseMat01.Clear
'    oBaseMat01.FlushToDataSource
'    oBaseMat01.LoadFromDataSource
    
    oMat01.FlushToDataSource
    j = oBaseColRow01
    For i = 0 To oMat01.VisualRowCount - 1
        If Trim(oDS_PS_MM013L.GetValue("U_ColReg20", i)) = "Y" Then
        
            oBaseMat01.Columns("PQDocNum").Cells(j).Specific.VALUE = Trim(oDS_PS_MM013L.GetValue("U_ColReg01", i))
            oBaseMat01.Columns("PQLinNum").Cells(j).Specific.VALUE = Trim(oDS_PS_MM013L.GetValue("U_ColReg02", i))
            oBaseMat01.Columns("E_BANFN").Cells(j).Specific.VALUE = Trim(oDS_PS_MM013L.GetValue("U_ColReg03", i))
            oBaseMat01.Columns("E_BNFPO").Cells(j).Specific.VALUE = Trim(oDS_PS_MM013L.GetValue("U_ColReg04", i))
            oBaseMat01.Columns("DueDate").Cells(j).Specific.VALUE = Trim(oDS_PS_MM013L.GetValue("U_ColReg05", i))
            oBaseMat01.Columns("ItemCode").Cells(j).Specific.VALUE = Trim(oDS_PS_MM013L.GetValue("U_ColReg06", i))
            oBaseMat01.Columns("ItemName").Cells(j).Specific.VALUE = Trim(oDS_PS_MM013L.GetValue("U_ColReg07", i))
            oBaseMat01.Columns("Size").Cells(j).Specific.VALUE = Trim(oDS_PS_MM013L.GetValue("U_ColReg08", i))
            oBaseMat01.Columns("Unit").Cells(j).Specific.VALUE = Trim(oDS_PS_MM013L.GetValue("U_ColReg09", i))
            oBaseMat01.Columns("Qty").Cells(j).Specific.VALUE = Trim(oDS_PS_MM013L.GetValue("U_ColNum01", i))
            oBaseMat01.Columns("CardCode").Cells(j).Specific.VALUE = Trim(oDS_PS_MM013L.GetValue("U_ColReg10", i))
            oBaseMat01.Columns("CardName").Cells(j).Specific.VALUE = Trim(oDS_PS_MM013L.GetValue("U_ColReg11", i))
            oBaseMat01.Columns("Mm01").Cells(j).Specific.VALUE = 0
            oBaseMat01.Columns("Mm02").Cells(j).Specific.VALUE = 0
            oBaseMat01.Columns("Mm03").Cells(j).Specific.VALUE = 0
            oBaseMat01.Columns("Mm04").Cells(j).Specific.VALUE = 0
            oBaseMat01.Columns("Mm05").Cells(j).Specific.VALUE = 0
            oBaseMat01.Columns("Mm06").Cells(j).Specific.VALUE = 0
            oBaseMat01.Columns("Mm07").Cells(j).Specific.VALUE = 0
            oBaseMat01.Columns("Mm08").Cells(j).Specific.VALUE = 0
            oBaseMat01.Columns("Mm09").Cells(j).Specific.VALUE = 0
            oBaseMat01.Columns("Mm10").Cells(j).Specific.VALUE = 0
            oBaseMat01.Columns("Mm11").Cells(j).Specific.VALUE = 0
            oBaseMat01.Columns("Mm12").Cells(j).Specific.VALUE = 0
            oBaseMat01.Columns("MmTot").Cells(j).Specific.VALUE = 0
            
            j = j + 1
        End If
    Next i
    
    oBaseForm01.Freeze False
    Call oForm01.Close
    Exit Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
SetBaseForm_Error:
    oBaseForm01.Freeze False
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message " ", "E"
    Else
        MDC_Com.MDC_GF_Message "SetBaseForm_Error:" & Err.Number & " - " & Err.Description, "E"
    End If
End Sub
