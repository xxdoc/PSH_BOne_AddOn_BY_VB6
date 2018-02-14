VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_QM065"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_QM065.cls
'//  Module         : QM
'//  Desc           : 포장사업팀 제품검사사양서 등록
'//  FormType       : PS_QM065
'//  Create Date    : 2016.07.27
'//  Modified Date  :
'//  Creator        : N.G.Y
'//  Copyright  (c) Poongsan Holdings
'//********************************************************************************
Option Explicit         '모듈 수준에서 해당하는 모듈의 모든 변수를 명시적으로 선언하도록 할 때 사용합니다.

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_QM065H       As SAPbouiCOM.DBDataSource    '등록헤더 'Represents a SAP Business One database table to be attached to a form.
Private oDS_PS_QM065L       As SAPbouiCOM.DBDataSource    '등록라인


Private oLast_Item_UID   As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID    As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row    As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
Private oSeq&
Private TmpCode As String

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional ByVal oBPLId01 As String, Optional ByVal oItemCode01 As String)
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_QM065.srf")
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
    
    oFormUniqueID01 = "PS_QM065_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)
    ' 변수나 속성 등에 개체 참조를 지정합니다

    oForm01.SupportedModes = -1
    
    If Trim(oItemCode01) = "" Then
        oForm01.Mode = fm_ADD_MODE
    Else
        oForm01.Mode = fm_FIND_MODE
    End If
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "DocEntry"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////D////////////////////////////////////////////////
    
    
    oForm01.Freeze True
    Call CreateItems
    
    oForm01.EnableMenu ("1283"), True      '// 제거
    oForm01.EnableMenu ("1293"), True      '// 행삭제
    oForm01.EnableMenu ("1287"), True      '// 복제
    oForm01.EnableMenu ("1284"), False     '// 취소
    
    
    
    Call ComboBox_Setting
    Call CF_ChooseFromList
    Call Initial_Setting
    Call FormItemEnabled
    Call FormClear '//UDO방식일때
    Call AddMatrixRow(0, oMat01.RowCount, True) '//UDO방식일때
    oForm01.Update      ' Reloads the form from the application.
                        ' Use this method to update an object from the database.

'    If oBaseItemCode01 <> "" Then
'        oForm01.Items("1").Click ct_Regular
'    End If
    
    oForm01.Freeze False
       
    oForm01.Visible = True  ' Indicates whether the form is visible.
    
    
    
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
    Dim sQry As String
    Dim ErrNum As Integer
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    
    Dim ChildForm01 As Variant
    Set ChildForm01 = New PS_SM010
    
    Dim ItemCode As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
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
                        
                        If oForm01.Mode = fm_ADD_MODE Then
                            Call FormClear
                            '// Input Code, Name
                            
                        End If
'                        If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
'
'                            Dim OutSize As String
'
'                            OutSize = oForm01.Items("OutSize").Specific.VALUE
'                            ItemCode = oForm01.Items("ItemCode").Specific.VALUE
'                            sQry = "Update [OITM] set U_OutSize = '" & OutSize & "' Where ItemCode = '" & ItemCode & "'"
'                            oRecordSet01.DoQuery sQry
'                        End If
                 
                    End If
                End If
                
            Case et_KEY_DOWN: '//2
                ' 제품코드
                If pval.CharPressed = 9 Then
                    If pval.ItemUID = "U_ItmBsort" Then
                        If oForm01.Items("U_ItmBsort").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "ItmMsort" Then
                        If oForm01.Items("ItmMsort").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "InspEmp" Then
                        If oForm01.Items("InspEmp").Specific.VALUE = "" Then
                            Sbo_Application.ActivateMenuItem ("7425")
                            BubbleEvent = False
                        End If
                    ElseIf pval.ItemUID = "ConfEmp" Then
                        If oForm01.Items("ConfEmp").Specific.VALUE = "" Then
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
                 oLast_Item_UID = pval.ItemUID

            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
            '//메트릭스에 데이터 로드
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
              If pval.ItemUID = "1" Then
            
                If oForm01.Mode = fm_ADD_MODE And pval.Action_Success = True Then
                    oForm01.Mode = fm_OK_MODE
                    Call Sbo_Application.ActivateMenuItem("1282")
                    
                ElseIf oForm01.Mode = fm_ADD_MODE And pval.Action_Success = False Then
                    FormItemEnabled
                    AddMatrixRow 1, oMat01.RowCount, True
                End If
                
                
                
              End If
            Case et_KEY_DOWN: '//2
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
                
                Call oForm01.Freeze(True)
                If pval.ItemChanged = True Then
                    ' 제품코드
                    If pval.ItemUID = "U_ItmBsort" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "ItmMsort" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "InspEmp" Then
                        FlushToItemValue pval.ItemUID
                    ElseIf pval.ItemUID = "ConfEmp" Then
                        FlushToItemValue pval.ItemUID
                    
                    ElseIf (pval.ItemUID = "Mat01") Then
                        Call oDS_PS_QM065L.setValue("U_" & pval.ColUID, pval.Row - 1, oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.VALUE)
                        If oMat01.RowCount = pval.Row And Trim(oDS_PS_QM065L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
                           AddMatrixRow 1, oMat01.VisualRowCount, True
                           Call oDS_PS_QM065L.setValue("U_Seqno", pval.Row - 1, oMat01.Columns("LineNum").Cells(pval.Row).Specific.VALUE)
                        End If
                        oMat01.LoadFromDataSource
                        oMat01.AutoResizeColumns
                        oForm01.Update
                        oMat01.Columns(pval.ColUID).Cells(pval.Row).Click
                    End If
                    
                    
                End If
                
'                If oMat01.RowCount = pval.Row And Trim(oDS_PS_QM065L.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
'                   AddMatrixRow 1, oMat01.VisualRowCount, True
'                End If
                Call oForm01.Freeze(False)
            Case et_MATRIX_LOAD: '//11
                AddMatrixRow 1, oMat01.VisualRowCount, True
            Case et_FORM_ACTIVATE: '//18
            Case et_FORM_DEACTIVATE: '//19
            Case et_FORM_RESIZE '//20
            Case et_CHOOSE_FROM_LIST '//27
            Case et_GOT_FOCUS: '//3
                oLast_Item_UID = pval.ItemUID
            Case et_LOST_FOCUS: '//4
            Case et_FORM_UNLOAD: '//17
                RemoveForms oFormUniqueID01
                Set oForm01 = Nothing
                Set oMat01 = Nothing
        End Select
    End If
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_ItemEvent_Error:
Call oForm01.Freeze(False)
Set oRecordSet01 = Nothing
    
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "신규(추가)모드에서는 적용할 수 없습니다. 제품조회 후 처리하세요.", "E"
    Else
        Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    End If
        
    
End Sub


Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
    Dim i&
    If (pval.BeforeAction = True) Then      '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
            Case "1293"     '행삭제
        End Select
        
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1281": '찾기
                FormItemEnabled
            Case "1282": '추가
                FormItemEnabled
                FormClear
                AddMatrixRow 0, oMat01.RowCount, True
                oForm01.Items("U_ItmBsort").Click ct_Collapsed
                
            Case "1287": '복제
                oForm01.Freeze True
                Call FormClear
'                oForm01.Items("ItemCode").Enabled = True
                For i = 0 To oMat01.VisualRowCount - 1
                    oMat01.FlushToDataSource
                    oDS_PS_QM065L.setValue "Code", i, ""
                    oMat01.LoadFromDataSource
                Next i
                
                Call FormItemEnabled
                
                oForm01.Freeze False
                
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                               
            Case "1293"     '행삭제
                If oMat01.RowCount <> oMat01.VisualRowCount Then
                    For i = 1 To oMat01.VisualRowCount
                        oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
                    Next i
                    oMat01.FlushToDataSource    ' DBDataSource에 레코드가 한줄 더 생긴다.
                    Call oDS_PS_QM065L.RemoveRecord(oDS_PS_QM065L.Size - 1)     ' 레코드 한 줄을 지운다.
                    oMat01.LoadFromDataSource   ' DBDataSource를 매트릭스에 올리고
                    If oMat01.RowCount = 0 Then
                        Call AddMatrixRow(1, 0, True)
                    Else
                        If Trim(oDS_PS_QM065L.GetValue("U_InspItem", oMat01.RowCount - 1)) <> "" Then
                            Call AddMatrixRow(1, oMat01.RowCount, True)

                        End If
                    End If
                End If
        End Select
    End If
    Exit Sub
Raise_MenuEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub


Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_FormDataEvent_Error
    If (BusinessObjectInfo.BeforeAction = True) Then '//BeforeAction = True
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD: '//33
                FormItemEnabled
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
Raise_FormDataEvent_Error:
    Sbo_Application.SetStatusBarMessage "Raise_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
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

Function CreateItems() As Boolean
On Error GoTo CreateItems_Error
    Dim oQuery01 As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_QM065H = oForm01.DataSources.DBDataSources("@PS_QM065H")
    Set oDS_PS_QM065L = oForm01.DataSources.DBDataSources("@PS_QM065L")
    
    Set oMat01 = oForm01.Items("Mat01").Specific
    
    
    
    Set oRecordSet01 = Nothing
    Exit Function
CreateItems_Error:
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    '//콤보에 기본값설정
    Dim oCombo          As SAPbouiCOM.ComboBox
    
    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", False, False)
    
'    Set oCombo = oForm01.Items("AQLYN").Specific
'    oCombo.ValidValues.Add "N", "N"
'    oCombo.ValidValues.Add "Y", "Y"
'    oCombo.Select 0, psk_Index
    
    Exit Sub
ComboBox_Setting_Error:
    Sbo_Application.SetStatusBarMessage "ComboBox_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub CF_ChooseFromList()
On Error GoTo CF_ChooseFromList_Error
    '//ChooseFromList 설정
    Exit Sub
CF_ChooseFromList_Error:
    Sbo_Application.SetStatusBarMessage "CF_ChooseFromList_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub Initial_Setting()
On Error GoTo Initial_Setting_Error
    '//아이디별 사업장 세팅
    Dim oCombo          As SAPbouiCOM.ComboBox
    Set oCombo = oForm01.Items("BPLId").Specific
    oCombo.Select MDC_PS_Common.User_BPLId, psk_ByValue
    
Exit Sub
Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "Initial_Setting_Error: " & Err.Number & " _ " & Err.Description, bmt_Short, True
End Sub


Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
        '//각모드에따른 아이템설정
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("U_ItmBsort").Enabled = True
        oForm01.Items("ItmMsort").Enabled = True
        oForm01.Items("SizeNo").Enabled = True
        oForm01.Items("OutSize").Enabled = True
        
        
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        '//각모드에따른 아이템설정
        '//각모드에따른 아이템설정
        oForm01.Items("BPLId").Enabled = True
        oForm01.Items("U_ItmBsort").Enabled = True
        oForm01.Items("ItmMsort").Enabled = True
        oForm01.Items("SizeNo").Enabled = True
        oForm01.Items("OutSize").Enabled = True
        
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        '//각모드에따른 아이템설정
        '//각모드에따른 아이템설정
        oForm01.Items("BPLId").Enabled = False
        oForm01.Items("U_ItmBsort").Enabled = False
        oForm01.Items("ItmMsort").Enabled = False
        oForm01.Items("SizeNo").Enabled = False
        oForm01.Items("OutSize").Enabled = False
    End If
    Exit Sub
FormItemEnabled_Error:
    Sbo_Application.SetStatusBarMessage "FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub AddMatrixRow(ByVal oSeq As Integer, ByVal oRow As Long, Optional RowIserted As Boolean)
'On Error GoTo AddMatrixRow_Error
   Select Case oSeq
       Case 0:
            oMat01.AddRow
            oDS_PS_QM065L.setValue "U_LineNum", oRow, oRow + 1
            oMat01.LoadFromDataSource
       Case 1:
            oDS_PS_QM065L.InsertRecord oRow
            oDS_PS_QM065L.setValue "U_LineNum", oRow, oRow + 1
            oMat01.LoadFromDataSource
   End Select
'AddMatrixRow_Error:
'    Sbo_Application.SetStatusBarMessage "AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM065'", "")
    If DocNum = 0 Then
'        oForm01.Items("DocEntry").Specific.String = 1
        oDS_PS_QM065H.setValue "DocEntry", 0, "1"
        oDS_PS_QM065H.setValue "Code", 0, "1"
        oDS_PS_QM065H.setValue "Name", 0, "1"
    Else
'        oForm01.Items("DocEntry").Specific.String = DocNum
        oDS_PS_QM065H.setValue "DocEntry", 0, DocNum    ' 화면에 적용이 안되기 때문
        oDS_PS_QM065H.setValue "Code", 0, DocNum
        oDS_PS_QM065H.setValue "Name", 0, DocNum
    End If
          
    
    Exit Sub
FormClear_Error:
    Sbo_Application.SetStatusBarMessage "FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function DataValidCheck() As Boolean
On Error GoTo DataValidCheck_Error
    '//유효성검사
    Exit Function
DataValidCheck_Error:
    Sbo_Application.SetStatusBarMessage "DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub MTX01()
On Error GoTo MTX01_Error
    '//메트릭스에 데이터 로드
    Exit Sub
MTX01_Error:
Sbo_Application.SetStatusBarMessage "MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Function HeaderSpaceLineDel() As Boolean
On Error GoTo HeaderSpaceLineDel_Error
    Dim ErrNum      As Integer
    Dim DocNum      As String
    oForm01.Freeze True
    ErrNum = 0
    
    '// Check
    Select Case True
        Case oDS_PS_QM065H.GetValue("U_ItmBsort", 0) = ""
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
        Case oDS_PS_QM065H.GetValue("U_ItmMsort", 0) = ""
            ErrNum = 2
            GoTo HeaderSpaceLineDel_Error
   '     Case oDS_PS_QM065H.GetValue("U_Inspect", 0) = ""
   '         ErrNum = 3
   '         GoTo HeaderSpaceLineDel_Error
    End Select
    oForm01.Freeze False
    HeaderSpaceLineDel = True
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "대분류는 필수입력 사항입니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "중분류는 필수입력 사항입니다. 확인하세요.", "E"
   ' ElseIf ErrNum = 3 Then
   '     MDC_Com.MDC_GF_Message "검사방법은 필수입력 사항입니다. 확인하세요.", "E"
    
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function

Private Function MatrixSpaceLineDel() As Boolean
On Error GoTo MatrixSpaceLineDel_Error
    Dim i             As Long
    Dim ErrNum        As Integer
    Dim oRecordSet01  As SAPbobsCOM.Recordset
    Dim sQry          As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    oForm01.Freeze True
    ErrNum = 0
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '// 화면상의 메트릭스에 입력된 내용을 모두 디비데이터소스로 넘긴다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oMat01.FlushToDataSource
    
    '// 라인
    If oMat01.VisualRowCount <= 1 Then
        ErrNum = 1
        GoTo MatrixSpaceLineDel_Error
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '// 맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
    '// 이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    If oMat01.VisualRowCount > 0 Then
        If oDS_PS_QM065L.GetValue("U_InspItem", oMat01.VisualRowCount - 1) = "" Then
            oDS_PS_QM065L.RemoveRecord oMat01.VisualRowCount - 1
        End If
    End If
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '행을 삭제하였으니 DB데이터 소스를 다시 가져온다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oMat01.LoadFromDataSource
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
    MatrixSpaceLineDel = True
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////
MatrixSpaceLineDel_Error:
    Set oRecordSet01 = Nothing
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
    Dim i$
    Dim oRecordSet01        As SAPbobsCOM.Recordset
    Dim sQry                As String
    Dim MItemCod$
    Dim Qty As Currency
    Dim Calculate_Weight As Currency
    Dim vReturnValue$
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
        Case "U_ItmBsort"
            sQry = "SELECT Name FROM [@PSH_ITMBSORT] WHERE Code = '" & Trim(oForm01.Items("U_ItmBsort").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oDS_PS_QM065H.setValue "U_ItmBName", 0, Trim(oRecordSet01.Fields(0).VALUE)
        Case "ItmMsort"
            sQry = "SELECT U_CodeName FROM [@PSH_ITMMSORT] WHERE U_Code = '" & Trim(oForm01.Items("ItmMsort").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oDS_PS_QM065H.setValue "U_ItmMName", 0, Trim(oRecordSet01.Fields(0).VALUE)
        Case "InspEmp"
            sQry = "Select U_FullName From [@PH_PY001A] Where Code = '" & Trim(oForm01.Items("InspEmp").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oDS_PS_QM065H.setValue "U_InspName", 0, Trim(oRecordSet01.Fields(0).VALUE)
        Case "ConfEmp"
            sQry = "Select U_FullName From [@PH_PY001A] Where Code = '" & Trim(oForm01.Items("ConfEmp").Specific.VALUE) & "'"
            oRecordSet01.DoQuery sQry
            oDS_PS_QM065H.setValue "U_ConfName", 0, Trim(oRecordSet01.Fields(0).VALUE)
    End Select
    oForm01.Freeze False
    Set oRecordSet01 = Nothing
End Sub
Private Function Item_Error_Message(ErrNum As Integer)
On Error GoTo Item_Error_Message

    If ErrNum = 1 Then
        GoTo Item_Error_Message
    End If

Item_Error_Message:

    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "제품코드를 먼저 입력하세요!", "E"
    End If
    
End Function


