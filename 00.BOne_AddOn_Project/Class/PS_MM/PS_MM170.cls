VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PS_MM170"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PS_MM170.cls
'//  Module         : MM
'//  Desc           : 지체상금등록
'//  FormType       : PS_MM170
'//  Create Date    : 2010.08.31
'//  Modified Date  :
'//  Creator        : DONG SUB KIM
'//  Copyright  (c) Poongsan Holdings
'//********************************************************************************
Option Explicit

Public oFormUniqueID01      As String
Public oForm01              As SAPbouiCOM.Form
Public oForm02              As SAPbouiCOM.Form
Public oMat01               As SAPbouiCOM.Matrix
Private oDS_PS_MM170H       As SAPbouiCOM.DBDataSource    '등록헤더
Private oDS_PS_MM170L       As SAPbouiCOM.DBDataSource    '등록라인

Private oLast_Item_UID   As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private oLast_Col_UID    As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private oLast_Col_Row    As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

Private oSeq&
'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm()
On Error GoTo LoadForm_Error
    Dim i As Long
    Dim oInnerXml01 As String
    Dim oXmlDoc01             As New MSXML2.DOMDocument
    
    oXmlDoc01.Load (SubMain.ShareFolderPath & "ScreenPS\PS_MM170.srf")
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
    
    oFormUniqueID01 = "PS_MM170_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID01 '//폼추가
    Sbo_Application.LoadBatchActions oXmlDoc01.xml
    
    '폼 할당
    Set oForm01 = Sbo_Application.Forms.Item(oFormUniqueID01)

    oForm01.SupportedModes = -1
    oForm01.Mode = fm_ADD_MODE
        
    oForm01.Freeze True
    
    Call oForm01.EnableMenu("1293", True)
    Call CreateItems
    
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '************************************************************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm01.DataBrowser.BrowseBy = "DocNum"
    '************************************************************************************************************
    '////////////////////////////////////////////////////////////D////////////////////////////////////////////////
    
    Call ComboBox_Setting
    Call CF_ChooseFromList
    Call Initial_Setting
    Call FormItemEnabled       '//폼의 아이템 활성화 컨트롤
    Call FormClear             '//UDO방식일때
    Call AddMatrixRow(0, oMat01.RowCount, True) '//UDO방식일때 - 매트릭스 Call Function
    
    oForm01.EnableMenu ("1283"), False        '// 삭제
    oForm01.EnableMenu ("1286"), False        '// 닫기
    oForm01.EnableMenu ("1287"), True         '// 복제
    oForm01.EnableMenu ("1284"), True         '// 취소
    oForm01.EnableMenu ("1293"), True         '// 행삭제
    
    '//Call MDC_GP_EnableMenus(oForm01, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False) '//메뉴설정
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
Dim i As Integer
Dim sQry                As String
Dim oRecordSet01        As SAPbobsCOM.Recordset
Dim TempForm01 As Variant
Dim ErrNum As Integer

Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)   '// 객체 정의 및 데이터 할당
    
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.EventType
        
            Case et_ITEM_PRESSED: '//1
            If pval.ItemUID = "1" Then
                If oForm01.Mode = fm_ADD_MODE Or oForm01.Mode = fm_UPDATE_MODE Then
                   If HeaderSpaceLineDel = False Then
                        BubbleEvent = False                   ' BubbleEvent = True 이면, 사용자에게 제어권을 넘겨준다. BeforeAction = True일 경우만 쓴다.
                        Exit Sub
                    End If
                    If MatrixSpaceLineDel = False Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If
            ElseIf pval.ItemUID = "Btn01" Then      '// 품의서 출력
            
            ElseIf pval.ItemUID = "Btn02" Then      '// DI API - 분개 생성
               If oForm01.Mode = fm_OK_MODE Then
                  If oForm01.Items("JdtDate").Specific.VALUE = "" Then
                     ErrNum = 1
                     Call Item_Error_Message(ErrNum)
                     BubbleEvent = False
                     Exit Sub
                  ElseIf oForm01.Items("Status").Specific.VALUE = "C" Then
                     ErrNum = 2
                     Call Item_Error_Message(ErrNum)
                     BubbleEvent = False
                     Exit Sub
                  Else
                    If Create_oJournalEntries(1) = False Then
                      BubbleEvent = False
                      Exit Sub
                    End If
                  End If
               
               Else
                   Call MDC_Com.MDC_GF_Message("먼저 저장한 후 분개 처리 바랍니다.", "W")
                   BubbleEvent = False
                   Exit Sub
               End If
                       
            ElseIf pval.ItemUID = "Btn03" Then      '// DI API - 분개 취소
               If oForm01.Mode = fm_OK_MODE Then
                  If oForm01.Items("JdtDate").Specific.VALUE = "" Then
                     ErrNum = 1
                     Call Item_Error_Message(ErrNum)
                     BubbleEvent = False
                     Exit Sub
                  ElseIf oForm01.Items("JdtCC").Specific.VALUE <> "Y" Then
                     ErrNum = 3
                     Call Item_Error_Message(ErrNum)
                     BubbleEvent = False
                     Exit Sub
                  ElseIf oForm01.Items("Status").Specific.VALUE = "C" Then
                     ErrNum = 2
                     Call Item_Error_Message(ErrNum)
                     BubbleEvent = False
                     Exit Sub
                  Else
                     If Cancel_oJournalEntries(1) = False Then
                        BubbleEvent = False
                        Exit Sub
                     End If
                  End If
               Else
                     Call MDC_Com.MDC_GF_Message("먼저 저장한 후 분개 처리 바랍니다.", "W")
                     BubbleEvent = False
                     Exit Sub
               End If
               
            Else
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "Mat01" And pval.ColUID = "ItemCode" Then
                        FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                    End If
                End If
            End If
   
               
            Case et_KEY_DOWN: '//2
                ' 거래처코드
                If oForm01.Items("CardCode").Specific.VALUE = "" Then
                    If pval.ItemUID = "CardCode" And pval.CharPressed = 9 Then
                        oForm01.Items("CardCode").Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                
                ' 입고번호 - 서브폼 호출
                If pval.ItemUID = "Mat01" And pval.ColUID = "GRDocNum" And pval.CharPressed = 9 Then
                    If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.String = "" Then
                        If oForm01.Items("CardCode").Specific.VALUE = "" Or _
                           oForm01.Items("BPLId").Specific.VALUE = "" Then
                           ErrNum = 4
                           Call Item_Error_Message(ErrNum)
                        Else
                           Set TempForm01 = New PS_MM171
                           Call TempForm01.LoadForm(oForm01, pval.ItemUID, pval.ColUID, pval.Row)
                        End If
                           BubbleEvent = False
                    End If
                End If
                               
                ' 담당자
                If oForm01.Items("CntcCode").Specific.VALUE = "" Then
                    If pval.ItemUID = "CntcCode" And pval.CharPressed = 9 Then
                        oForm01.Items("CntcCode").Click ct_Regular
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
                
                ' 품목코드
                If pval.ItemUID = "Mat01" And pval.ColUID = "ItemCode" And pval.CharPressed = 9 Then
                   If oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.String = "" Then
                       oMat01.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
                       Sbo_Application.ActivateMenuItem ("7425")
                       BubbleEvent = False
                   End If
                End If
                          
            Case et_COMBO_SELECT: '//5
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            If pval.ItemChanged = True Then
                If pval.ItemUID = "Mat01" And pval.ColUID = "GRDocNum" Then
                    FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                End If
            End If
        
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            
                ' 사원 이름 Query
                If pval.ItemUID = "CntcCode" Then
                    sQry = "Select U_FULLNAME, U_MSTCOD From [OHEM] Where U_MSTCOD = '" & Trim(oForm01.Items("CntcCode").Specific.VALUE) & "'"
                    Call oRecordSet01.DoQuery(sQry)
                    oForm01.Items("CntcName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                End If
                                                      
                
                                                      
                ' 아이템 코드
                If pval.ItemChanged = True Then
                    If pval.ItemUID = "Mat01" And pval.ColUID = "GRDocNum" Then
                        FlushToItemValue pval.ItemUID, pval.Row, pval.ColUID
                    End If
                End If
                
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
        
    ElseIf (pval.BeforeAction = False) Then         '//BeforeAction = False
        Select Case pval.EventType
            Case et_ITEM_PRESSED: '//1
            
             ' 저장 후 추가 가능처리
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
                If pval.Action_Success = True Then
                    oSeq = 1
                End If
            Case et_COMBO_SELECT: '//5
            
            ' 지체금액 매트릭스 SUM
            If pval.ItemUID = "Mat01" And pval.ColUID = "RepayYN" Then
                oForm01.Items("Sum").Specific.VALUE = ""
                For i = 1 To oMat01.VisualRowCount - 1
                    If oMat01.Columns("RepayYN").Cells(i).Specific.VALUE = "Y" Then
                        oForm01.Items("Sum").Specific.VALUE = Val(oForm01.Items("Sum").Specific.VALUE) + _
                        Val(oMat01.Columns("RepayP").Cells(i).Specific.VALUE)
                        '//문자열 내에 포함된 숫자를 적절한 형식의 숫자값으로 반환합니다.
                    End If
                Next i
            End If
            Case et_CLICK: '//6
            Case et_DOUBLE_CLICK: '//7
            Case et_MATRIX_LINK_PRESSED '//8
            Case et_VALIDATE: '//10
            ' 고객명 Query
            If pval.ItemUID = "CardCode" And pval.ItemChanged = True Then
                sQry = "Select CardName From [OCRD] Where CardCode = '" & Trim(oForm01.Items("CardCode").Specific.VALUE) & "'"
                oRecordSet01.DoQuery sQry
                oForm01.Items("CardName").Specific.VALUE = Trim(oRecordSet01.Fields(0).VALUE)
                oDS_PS_MM170H.setValue "U_CardName", 0, Trim(oRecordSet01.Fields(0).VALUE)
            End If
            
            '//전기일자를 분개일자와 동일하게...
            If pval.ItemUID = "DocDate" Then
                oDS_PS_MM170H.setValue "U_JdtDate", 0, oForm01.Items("DocDate").Specific.VALUE
            End If
            
            Case et_MATRIX_LOAD: '//11
                AddMatrixRow 1, oMat01.VisualRowCount, True
            Case et_FORM_ACTIVATE: '//18
            If oSeq = 1 Then
                oSeq = 0
            End If
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_ItemEvent_Error:
        Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub



Public Sub Raise_MenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_MenuEvent_Error
  Dim i&
    If (pval.BeforeAction = True) Then '//BeforeAction = True
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1293": '행닫기
                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281": '찾기
            Case "1282": '추가
            Case "1288", "1289", "1290", "1291": '레코드이동버튼

        End Select
    ElseIf (pval.BeforeAction = False) Then '//BeforeAction = False
        Select Case pval.MenuUID
            Case "1284": '취소
            Case "1286": '닫기
            Case "1281": '찾기
                FormItemEnabled
'                oForm01.Items("ItemCode").Click ct_Regular
            Case "1282": '추가
                FormItemEnabled
                FormClear
                AddMatrixRow 0, oMat01.RowCount, True
                oForm01.Items("CardCode").Click ct_Collapsed
                
            Case "1288", "1289", "1290", "1291": '레코드이동버튼
                FormItemEnabled
                If oMat01.VisualRowCount > 0 Then
                    If oMat01.Columns("GRDocNum").Cells(oMat01.VisualRowCount).Specific.VALUE <> "" Then
                        AddMatrixRow 1, oMat01.RowCount, True
                    End If
                End If
            Case "1293": '행닫기
               
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
            Case et_FORM_DATA_ADD:  '//34 - 추가
'            FormClear
'            If Create_oJournalEntries(2) = False Then
'                BubbleEvent = False
'                Exit Sub
'            End If
            Case et_FORM_DATA_UPDATE: '//35 - 업데이트
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
    Dim sQry As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
     
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Set oDS_PS_MM170H = oForm01.DataSources.DBDataSources("@PS_MM170H")
    Set oDS_PS_MM170L = oForm01.DataSources.DBDataSources("@PS_MM170L")
    Set oMat01 = oForm01.Items("Mat01").Specific                        '// 매트릭스 데이터 셋
    
     oDS_PS_MM170H.setValue "U_DocDate", 0, Format(Now, "yyyymmdd")
     oDS_PS_MM170H.setValue "U_JdtDate", 0, Format(Now, "yyyymmdd")
     
    '// 사업장 리스트
    sQry = "SELECT BPLId, BPLName FROM [OBPL]"
    
    oRecordSet01.DoQuery sQry
    
    Do Until oRecordSet01.EOF
        oForm01.Items("BPLId").Specific.ValidValues.Add Trim(oRecordSet01.Fields(0).VALUE), Trim(oRecordSet01.Fields(1).VALUE)
        oRecordSet01.MoveNext
    Loop
    
    '// 지체상금유무
    oMat01.Columns("RepayYN").ValidValues.Add "Y", "부여"
    oMat01.Columns("RepayYN").ValidValues.Add "N", "면제"
    
    Set oRecordSet01 = Nothing
    Exit Function
    
CreateItems_Error:
    Set oRecordSet01 = Nothing
    Sbo_Application.SetStatusBarMessage "CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Sub ComboBox_Setting()
On Error GoTo ComboBox_Setting_Error
    '//콤보에 기본값설정
        Call oForm01.Freeze(True)
    
    '//콤보에 기본값설정
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM130", "Mat01", "ItemCode", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM130", "Mat01", "ItemCode", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueColumn(oMat01.Columns("Column"), "PS_MM130", "Mat01", "ItemCode")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM130", "ItemCode", "", "01", "완제품")
'    Call MDC_PS_Common.Combo_ValidValues_Insert("PS_MM130", "ItemCode", "", "02", "반제품")
'    Call MDC_PS_Common.Combo_ValidValues_SetValueItem(oForm01.Items("Item").Specific, "PS_MM130", "ItemCode")
'    Call MDC_SetMod.Set_ComboList(oForm01.Items("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL WHERE BPLId = '1' ORDER BY BPLId", "창원", False, False)

' 품목대분류
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItmBsort"), "SELECT Code, Name FROM [@PSH_ITMBSORT] ORDER BY Code")
' 품목중분류
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItmMsort"), "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] ORDER BY U_Code")
' 형태타입
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ItemType"), "SELECT Code, Name FROM [@PSH_SHAPE] ORDER BY Code")
' 질별
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Quality"), "SELECT Code, Name FROM [@PSH_QUALITY] ORDER BY Code")
' 인증기호
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("Mark"), "SELECT Code, Name FROM [@PSH_MARK] ORDER BY Code")
' 매입기준단위
    Call MDC_Com.MDC_GP_MatrixSetMatComboList(oMat01.Columns("ObasUnit"), "SELECT Code, Name FROM [@PSH_UOMORG] ORDER BY Code")
    
    Call oForm01.Freeze(False)
    Exit Sub
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
 ' 사업장
    Call oForm01.Items("BPLId").Specific.Select(MDC_PS_Common.User_BPLId(), psk_ByValue)
 ' 인수자
    oForm01.Items("CntcCode").Specific.VALUE = MDC_PS_Common.User_MSTCOD()
    Exit Sub
Initial_Setting_Error:
    Sbo_Application.SetStatusBarMessage "Initial_Setting_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub FormItemEnabled()
On Error GoTo FormItemEnabled_Error
    If (oForm01.Mode = fm_ADD_MODE) Then
        oForm01.Items("DocNum").Enabled = False
        oForm01.Items("JdtDate").Enabled = True
    ElseIf (oForm01.Mode = fm_FIND_MODE) Then
        oForm01.Items("DocNum").Enabled = True
        oForm01.Items("JdtDate").Enabled = True
    ElseIf (oForm01.Mode = fm_OK_MODE) Then
        oForm01.Items("DocNum").Enabled = False
        oForm01.Items("JdtDate").Enabled = False
    End If
    
    Exit Sub

FormItemEnabled_Error:
    Sbo_Application.SetStatusBarMessage "FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
    
End Sub

Sub AddMatrixRow(ByVal oSeq As Integer, ByVal oRow As Long, Optional RowIserted As Boolean)
'On Error GoTo AddMatrixRow_Error

   Select Case oSeq
        Case 0:
            oMat01.AddRow           ' 매트릭스에 새로운 로를 추가한다.
            oDS_PS_MM170L.setValue "LineId", oRow, oRow + 1
            oMat01.LoadFromDataSource
        Case 1:
            oDS_PS_MM170L.InsertRecord oRow
            oDS_PS_MM170L.setValue "LIneId", oRow, oRow + 1
            oMat01.LoadFromDataSource
    End Select
'AddMatrixRow_Error:
'    Sbo_Application.SetStatusBarMessage "AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub FormClear()
On Error GoTo FormClear_Error
    Dim DocNum As String
    DocNum = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM170'", "")
    If DocNum = 0 Then
        oDS_PS_MM170H.setValue "DocNum", 0, "1"
    Else
        oDS_PS_MM170H.setValue "DocNum", 0, DocNum    ' 화면에 적용이 안되기 때문
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

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long, Optional oCol As String)
    '//데이터변화에 따른처리
    Dim i$
    Dim oRecordSet01        As SAPbobsCOM.Recordset
    Dim sQry                As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    Select Case oUID
    
        Case "Mat01"    'rowcount: 로 카운트를 반환, VisualRowCount: 삭제된 로를 제외하고 현재 보이는 로 카운트를 반환
            If (oRow = oMat01.RowCount Or oMat01.VisualRowCount = 0) And _
            Trim(oMat01.Columns("GRDocNum").Cells(oRow).Specific.VALUE) <> "" Then
                oMat01.FlushToDataSource      '데이터 소스를 지우고 매트릭스로부터 데이터 소스 레코드로 각 로를 복사한다.
                AddMatrixRow 1, oMat01.RowCount, True
                oMat01.Columns("GRDocNum").Cells(oRow).Click ct_Regular 'Column: 칼럼 오브젝트의 collection을 반환한다.
            End If
    End Select
    Set oRecordSet01 = Nothing
End Sub

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
    
    ErrNum = 0
    
    '// Check
    Select Case True
        Case oDS_PS_MM170H.GetValue("U_CardCode", 0) = "" Or _
             oDS_PS_MM170H.GetValue("U_BPLId", 0) = "" Or _
             oDS_PS_MM170H.GetValue("U_DocDate", 0) = ""
        
            ErrNum = 1
            GoTo HeaderSpaceLineDel_Error
    End Select
    
    HeaderSpaceLineDel = True
Exit Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////
HeaderSpaceLineDel_Error:
    If ErrNum = 1 Then
        MDC_Com.MDC_GF_Message "거래처코드, 사업장, 요청일자는 필수입력 사항입니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "HeaderSpaceLineDel_Error:" & Err.Description, "E"
    End If
    HeaderSpaceLineDel = False
End Function
        
Private Function MatrixSpaceLineDel() As Boolean
'------------------------------------------------------------------------------
' 저장할 데이터의 유효성을 점검한다
'------------------------------------------------------------------------------
On Error GoTo MatrixSpaceLineDel_Error
    Dim i           As Long
    Dim K           As Long
    Dim ErrNum      As Integer
    Dim Chk_Data    As String
    Dim oRow        As Integer
    Dim oRecordSet01  As SAPbobsCOM.Recordset
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
            
    ErrNum = 0
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '// 화면상의 메트릭스에 입력된 내용을 모두 디비데이터소스로 넘긴다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    ' Flushes current data from the user interface to the bounded data source, as follows:
    ' 1. Cleans the data source.
    ' 2. Copies each row from the matrix to the corresponding data source record.
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
        
    For i = 0 To oMat01.VisualRowCount - 2
        oDS_PS_MM170L.Offset = i
            Select Case True
                Case oDS_PS_MM170L.GetValue("U_GRDocNum", i) = ""
                    ErrNum = 2
                    GoTo MatrixSpaceLineDel_Error
    
                Case oDS_PS_MM170L.GetValue("U_RepayYN", i) = ""
                    ErrNum = 3
                    GoTo MatrixSpaceLineDel_Error
                    
                Case oDS_PS_MM170L.GetValue("U_RepayP", i) = ""
                    ErrNum = 4
                    GoTo MatrixSpaceLineDel_Error
            End Select
     Next i
     End If

    If oDS_PS_MM170L.GetValue("U_GRDocNum", oMat01.VisualRowCount - 1) = "" Then
        oDS_PS_MM170L.RemoveRecord oMat01.VisualRowCount - 1
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
        MDC_Com.MDC_GF_Message "라인 데이터가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 2 Then
        MDC_Com.MDC_GF_Message "입고문서 데이터가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 3 Then
        MDC_Com.MDC_GF_Message "지체상금유무 데이터가 없습니다. 확인하세요.", "E"
    ElseIf ErrNum = 4 Then
        MDC_Com.MDC_GF_Message "지체금액 데이터가 없습니다. 확인하세요.", "E"
    Else
        MDC_Com.MDC_GF_Message "MatrixSpaceLineDel_Error:" & Err.Description, "E"
    End If
    MatrixSpaceLineDel = False
End Function

Private Function Exist_YN(DocNum$) As String

    Dim oRecordSet01   As SAPbobsCOM.Recordset
    Dim sQry           As String
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    sQry = "SELECT Top 1 T1.DocNum FROM [@PS_MM170H] T1 "
    sQry = sQry & " WHERE T1.DocNum  = '" & DocNum & "'"
    oRecordSet01.DoQuery sQry
    
    Do Until oRecordSet01.EOF
        Exist_YN = oRecordSet01(0).VALUE
        oRecordSet01.MoveNext
    Loop
    
    If Trim(Exist_YN) = "" Then
        Exist_YN = ""
        Exit Function
    End If
    
    Set oRecordSet01 = Nothing
End Function

Private Sub Raise_EVENT_ROW_DELETE(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
On Error GoTo Raise_EVENT_ROW_DELETE_Error
    Dim i As Long
    If pval.BeforeAction = True Then
        If Sbo_Application.MessageBox("정말 삭제 하시겠습니까?", 1, "OK", "NO") <> 1 Then
            BubbleEvent = False
        End If
        '//행삭제전 행삭제가능여부검사
    ElseIf pval.BeforeAction = False Then
        For i = 1 To oMat01.VisualRowCount
            oMat01.Columns("LineNum").Cells(i).Specific.VALUE = i
        Next i
'        oMat01.Clear
        oMat01.FlushToDataSource
        Call oDS_PS_MM170L.RemoveRecord(oDS_PS_MM170L.Size - 1)
        oMat01.LoadFromDataSource
        If oMat01.RowCount = 0 Then
            Call AddMatrixRow(0, oMat01.RowCount, True)
        Else
            If Trim(oDS_PS_MM170L.GetValue("U_ItemCode", oMat01.RowCount - 1)) <> "" Then
                Call AddMatrixRow(1, oMat01.RowCount, True)
            End If
        End If
    End If
    Exit Sub
Raise_EVENT_ROW_DELETE_Error:
    Sbo_Application.SetStatusBarMessage "Raise_EVENT_ROW_DELETE_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
        

Private Function Item_Error_Message(ErrNum As Integer)
On Error GoTo Item_Error_Message

GoTo Item_Error_Message

Item_Error_Message:
 If ErrNum = 1 Then
    MDC_Com.MDC_GF_Message "분개처리일을 먼저 입력하세요.", "E"
 ElseIf ErrNum = 2 Then
    MDC_Com.MDC_GF_Message "문서가 Close 또는 Cancel 되었습니다.", "E"
 ElseIf ErrNum = 3 Then
    MDC_Com.MDC_GF_Message "분개생성:Y일 때 취소 할 수 있습니다.", "E"
  ElseIf ErrNum = 4 Then
    MDC_Com.MDC_GF_Message "거래처코드와, 사업장을 먼저 입력하세요.", "E"
 End If
End Function
        
        
Private Function Create_oJournalEntries(ChkType As Integer) As Boolean
On Error GoTo Create_oJournalEntries
    Dim f_oJournalEntries As SAPbobsCOM.JournalEntries
    
    Dim i             As Long
    Dim ErrNum        As Integer
    Dim ErrCode       As Long
    Dim ErrMsg        As String
    Dim RetVal&
    Dim sTransId$
        
    Dim SCardCode$, sDocDate$
    Dim sSum            As Currency
    Dim sCC             As String
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry            As String
    
    If (Sbo_Company.InTransaction = True) Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Sbo_Company.StartTransaction
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat01.FlushToDataSource
    
    SCardCode = Trim(oDS_PS_MM170H.GetValue("U_CardCode", 0))
    sDocDate = Format(oDS_PS_MM170H.GetValue("U_JdtDate", 0), "0000-00-00")
    sSum = Trim(oDS_PS_MM170H.GetValue("U_Sum", 0))
 
    Set f_oJournalEntries = Sbo_Company.GetBusinessObject(oJournalEntries)
        
    With f_oJournalEntries
         .ReferenceDate = sDocDate            '전기일
         .DueDate = sDocDate
         .TaxDate = sDocDate
         
         .Lines.Add
         .Lines.SetCurrentLine 0
         .Lines.AccountCode = "21101010"     '관리계정
         .Lines.ShortName = SCardCode        'G/L계정/BP 코드
         .Lines.Debit = sSum                 '차변
         .Lines.LineMemo = "지체상금"
         
         .Lines.SetCurrentLine 1
         .Lines.AccountCode = "43125020"     '잡이익-기타
         .Lines.ShortName = "43125020"
         .Lines.Credit = sSum                '대변
         .Lines.LineMemo = "지체상금"
         .UserFields("U_BPLId").VALUE = oForm01.Items("BPLId").Specific.VALUE '사업장
         
    End With
    '// 완료
    RetVal = f_oJournalEntries.Add
    If (0 <> RetVal) Then
        Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
        GoTo Create_oJournalEntries
    End If

    sCC = "Y"
    
    If ChkType = 1 Then
        Sbo_Company.GetNewObjectCode sTransId
        sQry = "Update [@PS_MM170H] Set U_JdtNo = '" & sTransId & "', U_JdtDate = '" & sDocDate & "', U_Sum = '" & sSum & "', U_JdtCC = '" & sCC & "' "
        sQry = sQry & "Where DocNum = '" & Trim(oDS_PS_MM170H.GetValue("DocNum", 0)) & "'"
        oRecordSet01.DoQuery sQry
        If (Sbo_Company.InTransaction = True) Then
            Sbo_Company.EndTransaction wf_Commit
        End If
    End If

  oDS_PS_MM170H.setValue "U_JdtNo", 0, sTransId
  oDS_PS_MM170H.setValue "U_JdtDate", 0, Format(Now, "yyyymmdd")
  oDS_PS_MM170H.setValue "U_Sum", 0, sSum
  oDS_PS_MM170H.setValue "U_JdtCC", 0, sCC

  Set oRecordSet01 = Nothing
  Set f_oJournalEntries = Nothing
  Create_oJournalEntries = True
  
  oForm01.Items("Btn02").Enabled = False
  oForm01.Items("Btn03").Enabled = True
  
  Exit Function
  
'///////////////////////////////////////////////////////////////////////////////////////
Create_oJournalEntries:
    Set oRecordSet01 = Nothing
    Set f_oJournalEntries = Nothing
    If Sbo_Company.InTransaction Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Create_oJournalEntries = False
    MDC_Com.MDC_GF_Message "Create_oJournalEntries:" & Err.Description & ErrMsg, "E"
End Function
        
        
Private Function Cancel_oJournalEntries(ChkType As Integer) As Boolean
On Error GoTo Cancel_oJournalEntries
    Dim f_oJournalEntries As SAPbobsCOM.JournalEntries
    
    Dim i             As Long
    Dim ErrNum        As Integer
    Dim ErrCode       As Long
    Dim ErrMsg        As String
    Dim RetVal&
    Dim sTransId$
        
    Dim SCardCode$, sDocDate$
    Dim sSum            As Currency
    Dim sCC             As String
        
    Dim oRecordSet01 As SAPbobsCOM.Recordset
    Dim sQry            As String
    
    If (Sbo_Company.InTransaction = True) Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    
    Sbo_Company.StartTransaction
    
    Set oRecordSet01 = Sbo_Company.GetBusinessObject(BoRecordset)
    
    oMat01.FlushToDataSource
    
    Set f_oJournalEntries = Nothing
    Set f_oJournalEntries = Sbo_Company.GetBusinessObject(oJournalEntries)
        
    If (f_oJournalEntries.GetByKey(Trim(oDS_PS_MM170H.GetValue("U_JdtNo", 0))) = False) Then
        Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
        GoTo Cancel_oJournalEntries
    End If
    
    '// 완료
    RetVal = f_oJournalEntries.Cancel
    If (0 <> RetVal) Then
        Call Sbo_Company.GetLastError(ErrCode, ErrMsg)
        GoTo Cancel_oJournalEntries
    End If
    
    sCC = "N"
    
    If ChkType = 1 Then
        Sbo_Company.GetNewObjectCode sTransId
        sQry = "Update [@PS_MM170H] Set U_JdtCan = '" & sTransId & "', U_Sum = '" & sSum & "', U_JdtCC = '" & sCC & "' "
        sQry = sQry & "Where DocNum = '" & Trim(oDS_PS_MM170H.GetValue("DocNum", 0)) & "'"
        oRecordSet01.DoQuery sQry
        
        If (Sbo_Company.InTransaction = True) Then
            Sbo_Company.EndTransaction wf_Commit
        End If
    End If
    
'    oForm01.Update
  oDS_PS_MM170H.setValue "U_JdtCan", 0, sTransId
  oDS_PS_MM170H.setValue "U_Sum", 0, sSum
  oDS_PS_MM170H.setValue "U_JdtCC", 0, sCC
  
  Set oRecordSet01 = Nothing
  Set f_oJournalEntries = Nothing
  Cancel_oJournalEntries = True
  
  oForm01.Items("Btn02").Enabled = False
  oForm01.Items("Btn03").Enabled = False
  
  Exit Function
  
'///////////////////////////////////////////////////////////////////////////////////////
Cancel_oJournalEntries:
    Set oRecordSet01 = Nothing
    Set f_oJournalEntries = Nothing
    If Sbo_Company.InTransaction Then
        Sbo_Company.EndTransaction wf_RollBack
    End If
    Cancel_oJournalEntries = False
    MDC_Com.MDC_GF_Message "Create_oJournalEntries:" & Err.Description & ErrMsg, "E"
End Function
