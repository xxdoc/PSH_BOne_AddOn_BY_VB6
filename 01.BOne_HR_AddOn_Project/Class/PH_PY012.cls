VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PH_PY012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//********************************************************************************
'//  File           : PH_PY012.cls
'//  Module         : 인사관리 > 인사
'//  Desc           : 출장등록
'//********************************************************************************
Option Explicit

Public oFormUniqueID As String
Public oForm As SAPbouiCOM.Form

Public oMat1 As SAPbouiCOM.Matrix

Private oDS_PH_PY012A As SAPbouiCOM.DBDataSource
Private oDS_PH_PY012B As SAPbouiCOM.DBDataSource

Private oLastItemUID    As String
Private oLastColUID     As String
Private oLastColRow     As Long

Dim DestDivValue As String '이전 출장구분 값 저장용(전역변수)

Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)

    Dim i           As Long
    Dim oXmlDoc     As New MSXML2.DOMDocument
    
    On Error GoTo LoadForm_Error
    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\PH_PY012.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount * 10)
    For i = 1 To (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length)
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").Item(i - 1).nodeValue = 20
        oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight").Item(i - 1).nodeValue = 16
    Next
    oFormUniqueID = "PH_PY012_" & GetTotalFormsCount
    AddForms Me, oFormUniqueID, "PH_PY012"
    Sbo_Application.LoadBatchActions oXmlDoc.xml
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)

    oForm.SupportedModes = -1
    oForm.Mode = fm_ADD_MODE
    oForm.DataBrowser.BrowseBy = "DocEntry"
        
    oForm.Freeze True
    Call PH_PY012_CreateItems
    Call PH_PY012_EnableMenus
    Call PH_PY012_SetDocument(oFromDocEntry01)
'    Call PH_PY012_FormResize
    
    oForm.Update
    oForm.Freeze False
    
    oForm.Visible = True
    Set oXmlDoc = Nothing
    Exit Sub
    
LoadForm_Error:
    oForm.Update
    oForm.Freeze False
    Set oXmlDoc = Nothing
    Set oForm = Nothing
    Sbo_Application.SetStatusBarMessage "Form_Load Error:" & Err.Description, bmt_Short, True
End Sub

Private Function PH_PY012_CreateItems() As Boolean

    Dim sQry        As String
    Dim i           As Long
    
    Dim oEdit       As SAPbouiCOM.EditText
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns    As SAPbouiCOM.Columns
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY012_CreateItems_Error
    
    Call oForm.Freeze(True)
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Set oDS_PH_PY012A = oForm.DataSources.DBDataSources("@PH_PY012A")
    Set oDS_PH_PY012B = oForm.DataSources.DBDataSources("@PH_PY012B")

    
    Set oMat1 = oForm.Items("Mat01").Specific

    oMat1.SelectionMode = ms_Auto
    oMat1.AutoResizeColumns


    '//----------------------------------------------------------------------------------------------
    '// 기본사항
    '//----------------------------------------------------------------------------------------------
    
    '사업장
    Set oCombo = oForm.Items("CLTCOD").Specific
'    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
'    Call SetReDataCombo(oForm, sQry, oCombo)
    oForm.Items("CLTCOD").DisplayDesc = True

    '출장구분
    Set oCombo = oForm.Items("DestDiv").Specific
'    oCombo.ValidValues.Add "", ""
    oCombo.ValidValues.Add "01", "공용"
    oCombo.ValidValues.Add "02", "국내출장"
    oCombo.ValidValues.Add "03", "해외출장"
'    oCombo.Select 0, psk_Index
    oForm.Items("DestDiv").DisplayDesc = True
    
    '출장지역
    Set oCombo = oForm.Items("DestCode").Specific
    sQry = "            SELECT      '' AS [Code],"
    sQry = sQry & "                 '' AS [Name],"
    sQry = sQry & "                 -1 AS [Seq]"
    sQry = sQry & "  UNION ALL"
    sQry = sQry & "  SELECT      T1.U_Code AS [Code],"
    sQry = sQry & "                 T1.U_CodeNm AS [Name],"
    sQry = sQry & "                 T1.U_Seq AS [Seq]"
    sQry = sQry & "  FROM       [@PS_HR200H] AS T0"
    sQry = sQry & "                 INNER JOIN"
    sQry = sQry & "                 [@PS_HR200L] AS T1"
    sQry = sQry & "                     ON T0.Code = T1.Code"
    sQry = sQry & "  WHERE      T0.Code = 'P217'"
    sQry = sQry & "                 AND T1.U_UseYN = 'Y'"
    sQry = sQry & "  ORDER BY  Seq"
    Call SetReDataCombo(oForm, sQry, oCombo)
    oCombo.Select 0, psk_Index
    oForm.Items("DestCode").DisplayDesc = True

    '매트릭스-차량구분
    Set oColumn = oMat1.Columns("Vehicle")
    oColumn.ValidValues.Add "", ""
    sQry = "           SELECT      U_Code AS [Code],"
    sQry = sQry & "                U_CodeNm As Name"
    sQry = sQry & " FROM       [@PS_HR200L]"
    sQry = sQry & " WHERE      Code = 'P218'"
    sQry = sQry & "                AND U_UseYN = 'Y'"
    sQry = sQry & " ORDER BY  U_Seq"
    Call MDC_SetMod.GP_MatrixSetMatComboList(oColumn, sQry, False, False)
    oColumn.DisplayDesc = True
    
    '매트릭스-유종
    Set oColumn = oMat1.Columns("FuelType")
    oColumn.ValidValues.Add "", ""
    oColumn.ValidValues.Add "1", "휘발유"
    oColumn.ValidValues.Add "2", "가스"
    oColumn.ValidValues.Add "3", "경유"
    oColumn.DisplayDesc = True
    
    '매트릭스-통화
    Set oColumn = oMat1.Columns("Currency")
    oColumn.ValidValues.Add "", ""
    sQry = "           SELECT      T0.CurrCode AS [Code],"
    sQry = sQry & "                T0.CurrName AS [Name]"
    sQry = sQry & " FROM       [OCRN] AS T0"
    sQry = sQry & " ORDER BY  CurrCode"
    Call MDC_SetMod.GP_MatrixSetMatComboList(oColumn, sQry, False, False)

    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Function
    
PH_PY012_CreateItems_Error:
    Set oEdit = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oColumns = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY012_CreateItems_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY012_EnableMenus()
    
    On Error GoTo PH_PY012_EnableMenus_Error
    
    Call oForm.EnableMenu("1283", False)         '// 삭제
    Call oForm.EnableMenu("1287", False)          '// 복제
    Call oForm.EnableMenu("1286", True)         '// 닫기
    Call oForm.EnableMenu("1284", True)         '// 취소
    Call oForm.EnableMenu("1293", True)         '// 행삭제
    
    Exit Sub
    
PH_PY012_EnableMenus_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY012_EnableMenus_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Private Sub PH_PY012_SetDocument(ByVal oFromDocEntry01 As String)
    On Error GoTo PH_PY012_SetDocument_Error
    
    If (oFromDocEntry01 = "") Then
        Call PH_PY012_FormItemEnabled
        Call PH_PY012_AddMatrixRow
    Else
        oForm.Mode = fm_FIND_MODE
        Call PH_PY012_FormItemEnabled
        oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
        oForm.Items("1").CLICK ct_Regular
    End If
    Exit Sub
    
PH_PY012_SetDocument_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY012_SetDocument_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY012_FormItemEnabled()
    On Error GoTo PH_PY012_FormItemEnabled_Error
    
    Dim oCombo As SAPbouiCOM.ComboBox
    
    Call oForm.Freeze(True)
    If (oForm.Mode = fm_ADD_MODE) Then

        '폼 DocEntry 세팅
        Call PH_PY012_FormClear
        
'        '사업장 세팅
'        Call oDS_PH_PY012A.setValue("U_CLTCOD", 0, MDC_SetMod.Get_ReData("Branch", "USER_CODE", "OUSR", "'" & oCompany.UserName & "'"))
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        '출장구분 세팅
        Set oCombo = oForm.Items("DestDiv").Specific
        oCombo.Select 0, psk_Index

        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", False)    '//문서추가
        
    ElseIf (oForm.Mode = fm_FIND_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD")
        
        Call oForm.EnableMenu("1281", False)    '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가
        
    ElseIf (oForm.Mode = fm_OK_MODE) Then
        '// 접속자에 따른 권한별 사업장 콤보박스세팅
        Call CLTCOD_Select(oForm, "CLTCOD", False)
        
        Call oForm.EnableMenu("1281", True)     '//문서찾기
        Call oForm.EnableMenu("1282", True)     '//문서추가

    End If
    
    Call oForm.Freeze(False)
    Exit Sub
    
PH_PY012_FormItemEnabled_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY012_FormItemEnabled_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
    On Error GoTo Raise_FormItemEvent_Error
    
    Dim sQry As String
    Dim i As Long
    Dim oCombo As SAPbouiCOM.ComboBox
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Dim BPLID As String '사업장 코드
    Dim StdYear As String '기준년도
    Dim StdMonth As String '기준월
    
    Dim FuelType As String '유종
    Dim FuelPrc As Double '단가
    Dim Distance As Double '거리
    Dim FoodNum As Double '식수
    Dim TransExp As Long '교통비
    Dim InsurExp As Double '보험료
    Dim AirpExp As Double '공항세
    Dim DayExp As Double '일비
    Dim LodgExp As Double '숙박비
    Dim FoodExp As Double '식비
    Dim TotFoodExp As Integer '총식비
    Dim ParkExp As Double '주차비
    Dim TollExp As Double '도로비
    Dim TotalExp As Double '합계
    
    Select Case pval.EventType
        Case et_ITEM_PRESSED: '//1
            
            If pval.BeforeAction = True Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If PH_PY012_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                        
                        '//해야할일 작업
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If PH_PY012_DataValidCheck = False Then
                            BubbleEvent = False
                        End If
                        '//해야할일 작업
                        
                    ElseIf oForm.Mode = fm_OK_MODE Then
                    End If
                End If
            ElseIf pval.BeforeAction = False Then
                If pval.ItemUID = "1" Then
                    If oForm.Mode = fm_ADD_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY012_FormItemEnabled
                            Call PH_PY012_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY012_FormItemEnabled
                            Call PH_PY012_AddMatrixRow
                        End If
                    ElseIf oForm.Mode = fm_OK_MODE Then
                        If pval.ActionSuccess = True Then
                            Call PH_PY012_FormItemEnabled
                        End If
                    End If
                End If
            End If
     
'----------------------------------------------------------
        Case et_KEY_DOWN: '//2

            If pval.BeforeAction = True Then
            
                If pval.ItemUID = "Mat01" Then
        
'                    If pval.ColUID = "Name" And pval.CharPressed = "9" Then
'
'                        If oMat1.Columns.Item("Name").Cells(pval.Row).Specific.Value = "" Then
'                            Call Sbo_Application.ActivateMenuItem("7425")
'                            BubbleEvent = False
'                        End If
'
'                    End If
                
                ElseIf pval.ItemUID = "MSTCOD" And pval.CharPressed = "9" Then
                
                    If oForm.Items("MSTCOD").Specific.Value = "" Then
                        Call Sbo_Application.ActivateMenuItem("7425")
                        BubbleEvent = False
                    End If
            
                End If

            ElseIf pval.Before_Action = False Then

            End If
            
'----------------------------------------------------------
        Case et_GOT_FOCUS: '//3
            Select Case pval.ItemUID
            Case "Mat01"
                If pval.Row > 0 Then
                    oLastItemUID = pval.ItemUID
                    oLastColUID = pval.ColUID
                    oLastColRow = pval.Row
                End If
            Case Else
                oLastItemUID = pval.ItemUID
                oLastColUID = ""
                oLastColRow = 0
            End Select
            
'----------------------------------------------------------
        Case et_LOST_FOCUS: '//4

'----------------------------------------------------------
        Case et_COMBO_SELECT: '//5
        
            Call oForm.Freeze(True)
            
            If pval.BeforeAction = True Then
                
                If pval.ItemUID = "DestDiv" Then

                    If oMat1.RowCount > 1 Then

                        DestDivValue = oForm.Items("DestDiv").Specific.Value '"계속 진행하겠습니까?" 라는 질문에 No를 선택했을 경우 DestDiv 콤보박스의 Value를 돌려놓기 위해 이전 값을 저장

                    End If

                End If
                
            ElseIf pval.BeforeAction = False Then
            
                If pval.ItemChanged = True Then
                
                    If pval.ItemUID = "Mat01" Then
                    
                        If pval.ColUID = "Vehicle" Then

                            Call PH_PY012_AddMatrixRow
                        
                        ElseIf pval.ColUID = "FuelType" Then '유종을 선택하면 유종에 따른 단가를 조회
                        
                            oMat1.FlushToDataSource
                                
                            BPLID = Trim(oDS_PH_PY012A.GetValue("U_CLTCOD", 0)) '사업장
                            StdYear = Left(Trim(oDS_PH_PY012A.GetValue("U_FrDate", 0)), 4)
                            StdMonth = Mid(Trim(oDS_PH_PY012A.GetValue("U_FrDate", 0)), 5, 2)
                            FuelType = Trim(oDS_PH_PY012B.GetValue("U_FuelType", pval.Row - 1))
                            
                            FuelPrc = PH_PY012_GetFuelPrc(BPLID, StdYear, StdMonth, FuelType)
                            
                            Call oDS_PH_PY012B.setValue("U_FuelPrc", pval.Row - 1, FuelPrc)

                            oMat1.LoadFromDataSource

                        End If
                        
                    ElseIf pval.ItemUID = "DestDiv" Then
                    
                        If oMat1.RowCount > 1 Then

                            If Sbo_Application.MessageBox("저장하지 않은 데이터는 모두 삭제됩니다. 계속 진행하겠습니까?", 1, "Yes", "No") = 1 Then

                                Call oMat1.Clear
                                Call oMat1.FlushToDataSource
                                Call oMat1.LoadFromDataSource

                                Call PH_PY012_AddMatrixRow

                            Else

                                '이전 상태의 콤보값으로 되돌림
                                Call oDS_PH_PY012A.setValue("U_DestDiv", 0, DestDivValue) '콤보박스를 바로 Select 하면 ComboSelect 이벤트가 발생하기 때문에 UDO를 이용하여 Binding
                                Call oForm.Freeze(False)
                                
                                Exit Sub

                            End If

                        End If
                    
                        '출장구분에 따라 매트릭스 컬럼 Visible 설정
                        Call PH_PY012_ColumnStatusChange
                    
                    End If
                    
                    Call oMat1.AutoResizeColumns
                    
                End If
            End If
            
            Call oForm.Freeze(False)
            
'----------------------------------------------------------
        Case et_CLICK: '//6
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                Case "Mat01"
                    If pval.Row > 0 Then
                        Call oMat1.SelectRow(pval.Row, True, False)
                    End If
                End Select
                
                Select Case pval.ItemUID
                Case "Mat01"
                    If pval.Row > 0 Then
                        oLastItemUID = pval.ItemUID
                        oLastColUID = pval.ColUID
                        oLastColRow = pval.Row
                    End If
                Case Else
                    oLastItemUID = pval.ItemUID
                    oLastColUID = ""
                    oLastColRow = 0
                End Select
            ElseIf pval.BeforeAction = False Then

            End If
            
'----------------------------------------------------------
        Case et_DOUBLE_CLICK: '//7

'----------------------------------------------------------
        Case et_MATRIX_LINK_PRESSED '//8

'----------------------------------------------------------
        Case et_MATRIX_COLLAPSE_PRESSED '//9
        
'----------------------------------------------------------
        Case et_VALIDATE: '//10
            Call oForm.Freeze(True)
            If pval.BeforeAction = True Then
            
                If pval.ItemChanged = True Then

                End If
       
            ElseIf pval.BeforeAction = False Then
            
                If pval.ItemChanged = True Then
                    
                    Select Case pval.ItemUID
                    
                        Case "MSTCOD"
                        
                            Call oDS_PH_PY012A.setValue("U_MSTNAM", 0, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items(pval.ItemUID).Specific.Value & "'"))
                        
                        Case "Mat01"
                        
                            If pval.ColUID = "Vehicle" And oForm("DestDiv").Specific.Value = "01" Then '차량구분 선택 시
                            
'                                oMat1.FlushToDataSource

'                                oMat1.LoadFromDataSource
                                
                                Call PH_PY012_AddMatrixRow
                            
                            ElseIf pval.ColUID = "FrDate" And (oForm("DestDiv").Specific.Value = "02" Or oForm("DestDiv").Specific.Value = "03") Then '기간시작
                            
'                                oMat1.FlushToDataSource
'
'                                oMat1.LoadFromDataSource
                                
                                Call PH_PY012_AddMatrixRow
                                
                            ElseIf pval.ColUID = "FuelPrc" Then '단가 입력 시
                            
                                oMat1.FlushToDataSource
                                
                                FuelPrc = oDS_PH_PY012B.GetValue("U_FuelPrc", pval.Row - 1)
                                Distance = oDS_PH_PY012B.GetValue("U_Distance", pval.Row - 1)
                                
                                TransExp = FuelPrc * Distance * 0.1
                                
                                Call oDS_PH_PY012B.setValue("U_TransExp", pval.Row - 1, TransExp)

                                '합계 계산(교통비+일비+식비+주차비+도로비)
                                DayExp = oDS_PH_PY012B.GetValue("U_DayExp", pval.Row - 1) '일비
                                FoodExp = oDS_PH_PY012B.GetValue("U_FoodExp", pval.Row - 1) '식비
                                ParkExp = oDS_PH_PY012B.GetValue("U_ParkExp", pval.Row - 1) '주차비
                                TollExp = oDS_PH_PY012B.GetValue("U_TollExp", pval.Row - 1) '도로비

                                TotalExp = TransExp + DayExp + (FoodNum * FoodExp) + ParkExp + TollExp
                                
                                Call oDS_PH_PY012B.setValue("U_TotalExp", pval.Row - 1, TotalExp) '합계

                                oMat1.LoadFromDataSource
                                
                            ElseIf pval.ColUID = "Distance" Then '거리 입력 시
                            
                                oMat1.FlushToDataSource
                                
                                FuelPrc = oDS_PH_PY012B.GetValue("U_FuelPrc", pval.Row - 1)
                                Distance = oDS_PH_PY012B.GetValue("U_Distance", pval.Row - 1)
                                
                                TransExp = FuelPrc * Distance * 0.1
                                
                                Call oDS_PH_PY012B.setValue("U_TransExp", pval.Row - 1, TransExp)

                                '합계 계산(교통비+일비+식비+주차비+도로비)
                                DayExp = oDS_PH_PY012B.GetValue("U_DayExp", pval.Row - 1) '일비
                                FoodExp = oDS_PH_PY012B.GetValue("U_FoodExp", pval.Row - 1) '식비
                                ParkExp = oDS_PH_PY012B.GetValue("U_ParkExp", pval.Row - 1) '주차비
                                TollExp = oDS_PH_PY012B.GetValue("U_TollExp", pval.Row - 1) '도로비

                                TotalExp = TransExp + DayExp + (FoodNum * FoodExp) + ParkExp + TollExp
                                
                                Call oDS_PH_PY012B.setValue("U_TotalExp", pval.Row - 1, TotalExp) '합계

                                oMat1.LoadFromDataSource
                                
                            ElseIf pval.ColUID = "FoodNum" Then '식수 입력 시
                            
                                oMat1.FlushToDataSource
                                
                                FoodNum = oDS_PH_PY012B.GetValue("U_FoodNum", pval.Row - 1)
                                FoodExp = oDS_PH_PY012B.GetValue("U_FoodExp", pval.Row - 1)
                                
                                TotFoodExp = FoodNum * FoodExp
                                 
                                Call oDS_PH_PY012B.setValue("U_FoodExp", pval.Row - 1, FoodExp)

                                '합계 계산(교통비+일비+식비+주차비+도로비)
                                TransExp = oDS_PH_PY012B.GetValue("U_TransExp", pval.Row - 1) '교통비
                                DayExp = oDS_PH_PY012B.GetValue("U_DayExp", pval.Row - 1) '일비
                                ParkExp = oDS_PH_PY012B.GetValue("U_ParkExp", pval.Row - 1) '주차비
                                TollExp = oDS_PH_PY012B.GetValue("U_TollExp", pval.Row - 1) '도로비

                                TotalExp = TransExp + DayExp + (FoodNum * FoodExp) + ParkExp + TollExp
                                
                                Call oDS_PH_PY012B.setValue("U_TotalExp", pval.Row - 1, TotalExp) '합계

                                oMat1.LoadFromDataSource
                                
                            ElseIf pval.ColUID = "FoodExp" Then '식비 입력 시
                            
                                oMat1.FlushToDataSource
                                
                                FoodNum = oDS_PH_PY012B.GetValue("U_FoodNum", pval.Row - 1)
                                FoodExp = oDS_PH_PY012B.GetValue("U_FoodExp", pval.Row - 1)
                                
                                TotFoodExp = FoodNum * FoodExp
                                 
                                Call oDS_PH_PY012B.setValue("U_FoodExp", pval.Row - 1, FoodExp)

                                '합계 계산(교통비+일비+식비+주차비+도로비)
                                TransExp = oDS_PH_PY012B.GetValue("U_TransExp", pval.Row - 1) '교통비
                                DayExp = oDS_PH_PY012B.GetValue("U_DayExp", pval.Row - 1) '일비
                                ParkExp = oDS_PH_PY012B.GetValue("U_ParkExp", pval.Row - 1) '주차비
                                TollExp = oDS_PH_PY012B.GetValue("U_TollExp", pval.Row - 1) '도로비

                                TotalExp = TransExp + DayExp + (FoodNum * FoodExp) + ParkExp + TollExp
                                
                                Call oDS_PH_PY012B.setValue("U_TotalExp", pval.Row - 1, TotalExp) '합계

                                oMat1.LoadFromDataSource
                            
                            ElseIf pval.ColUID = "ParkExp" Then '주차비 입력 시
                            
'                                oMat1.FlushToDataSource
'
'                                '합계(교통비+일비+식비+주차비+도로비+숙박비+보험료+공항세)
'                                TransExp = oDS_PH_PY012B.GetValue("U_TransExp", pval.Row - 1) '교통비
'                                DayExp = oDS_PH_PY012B.GetValue("U_DayExp", pval.Row - 1) '일비
'                                FoodExp = oDS_PH_PY012B.GetValue("U_FoodExp", pval.Row - 1) '식비
'                                ParkExp = oDS_PH_PY012B.GetValue("U_ParkExp", pval.Row - 1) '주차비
'                                TollExp = oDS_PH_PY012B.GetValue("U_TollExp", pval.Row - 1) '도로비
'
'                                TotalExp = TransExp + DayExp + FoodExp + ParkExp + TollExp
'
'                                Call oDS_PH_PY012B.setValue("U_TotalExp", pval.Row - 1, TotalExp) '합계
'
'                                oMat1.LoadFromDataSource

                                Call PH_PY012_CalculateTotalExp(pval.Row - 1) '합계 계산
                            
                            ElseIf pval.ColUID = "TollExp" Then '도로비 입력 시
                            
'                                oMat1.FlushToDataSource
'
'                                '합계(교통비+일비+식비+주차비+도로비+숙박비+보험료+공항세)
'                                TransExp = oDS_PH_PY012B.GetValue("U_TransExp", pval.Row - 1) '교통비
'                                DayExp = oDS_PH_PY012B.GetValue("U_DayExp", pval.Row - 1) '일비
'                                FoodExp = oDS_PH_PY012B.GetValue("U_FoodExp", pval.Row - 1) '식비
'                                ParkExp = oDS_PH_PY012B.GetValue("U_ParkExp", pval.Row - 1) '주차비
'                                TollExp = oDS_PH_PY012B.GetValue("U_TollExp", pval.Row - 1) '도로비
'
'                                TotalExp = TransExp + DayExp + FoodExp + ParkExp + TollExp
'
'                                Call oDS_PH_PY012B.setValue("U_TotalExp", pval.Row - 1, TotalExp) '합계
'
'                                oMat1.LoadFromDataSource

                                Call PH_PY012_CalculateTotalExp(pval.Row - 1) '합계 계산
                            
                            ElseIf pval.ColUID = "DayExp" Then '일비 입력 시
                            
                                Call PH_PY012_CalculateTotalExp(pval.Row - 1) '합계 계산
                                
                            ElseIf pval.ColUID = "LodgExp" Then '숙박비 입력 시
                            
                                Call PH_PY012_CalculateTotalExp(pval.Row - 1) '합계 계산
                            
                            ElseIf pval.ColUID = "TransExp" Then '교통비 입력 시
                            
                                Call PH_PY012_CalculateTotalExp(pval.Row - 1) '합계 계산
                            
                            ElseIf pval.ColUID = "InsurExp" Then '보험료 입력 시
                            
                                Call PH_PY012_CalculateTotalExp(pval.Row - 1) '합계 계산
                            
                            ElseIf pval.ColUID = "FoodExp" Then '식비 입력 시
                            
                                Call PH_PY012_CalculateTotalExp(pval.Row - 1) '합계 계산
                            
                            ElseIf pval.ColUID = "AirpExp" Then '공항세 입력 시
                            
                                Call PH_PY012_CalculateTotalExp(pval.Row - 1) '합계 계산
                            
                            End If
                            
                            Call oMat1.Columns(pval.ColUID).Cells(pval.Row).CLICK(ct_Regular)
                            Call oMat1.AutoResizeColumns
                        
                    End Select

                End If
                
            End If
            Call oForm.Freeze(False)
            
'----------------------------------------------------------
        Case et_MATRIX_LOAD: '//11
            If pval.BeforeAction = True Then
            
            ElseIf pval.BeforeAction = False Then
                oMat1.LoadFromDataSource

                Call PH_PY012_FormItemEnabled
                Call PH_PY012_AddMatrixRow
                Call oMat1.AutoResizeColumns
                
            End If
            
'----------------------------------------------------------
        Case et_DATASOURCE_LOAD '//12
           
'----------------------------------------------------------
        Case et_FORM_LOAD: '//16
            
'----------------------------------------------------------
        Case et_FORM_UNLOAD: '//17
            If pval.BeforeAction = True Then
            ElseIf pval.BeforeAction = False Then
                RemoveForms oFormUniqueID
                Set oForm = Nothing
                Set oDS_PH_PY012A = Nothing
                Set oDS_PH_PY012B = Nothing
                
                Set oMat1 = Nothing
                
            End If
            
'----------------------------------------------------------
        Case et_FORM_ACTIVATE: '//18
            
'----------------------------------------------------------
        Case et_FORM_DEACTIVATE: '//19
            
'----------------------------------------------------------
        Case et_FORM_CLOSE '//20
            
'----------------------------------------------------------
        Case et_FORM_RESIZE '//21
            If pval.BeforeAction = True Then
        
            ElseIf pval.BeforeAction = False Then

                Call oMat1.AutoResizeColumns

            End If
            
'----------------------------------------------------------
        Case et_FORM_KEY_DOWN '//22
            
'----------------------------------------------------------
        Case et_FORM_MENU_HILIGHT '//23
            
'----------------------------------------------------------
        Case et_CHOOSE_FROM_LIST '//27
            If pval.BeforeAction = True Then
            
            ElseIf pval.Before_Action = False Then
'                If pval.ItemUID = "Code" Then
'                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY012A", "Code")
'                End If
            End If

'----------------------------------------------------------
        Case et_PICKER_CLICKED '//37
            
'----------------------------------------------------------
        Case et_GRID_SORT '//38
            
'----------------------------------------------------------
        Case et_Drag '//39
            
    End Select
    
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_FormItemEvent_Error:
    oForm.Freeze (False)
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_ItemEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
    Dim i As Long
    On Error GoTo Raise_FormMenuEvent_Error
    
    Dim loopCount As Integer
    Dim FeeTot As Double
    Dim TuiTot As Double
    Dim Total As Double
    
    Call oForm.Freeze(True)
    
    If (pval.BeforeAction = True) Then
        Select Case pval.MenuUID
            Case "1283":
                If Sbo_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") = 2 Then
                    BubbleEvent = False
                    Exit Sub
                End If
            Case "1284":
            Case "1286":
            Case "1293":
            Case "1281":
            Case "1282":
            Case "1288", "1289", "1290", "1291":
            
        End Select
    ElseIf (pval.BeforeAction = False) Then
        Select Case pval.MenuUID
            Case "1283":
                oForm.Mode = fm_ADD_MODE
                Call PH_PY012_FormItemEnabled
                Call PH_PY012_AddMatrixRow
            Case "1284":
            Case "1286":
'            Case "1293":
'                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
            Case "1281":    '//문서찾기
                Call PH_PY012_FormItemEnabled
                Call PH_PY012_AddMatrixRow
                oForm.Items("DocEntry").CLICK ct_Regular
                Call PH_PY012_ColumnStatusChange
            Case "1282":    '//문서추가
                Call PH_PY012_FormItemEnabled
                Call PH_PY012_AddMatrixRow
            Case "1288", "1289", "1290", "1291":
                Call PH_PY012_FormItemEnabled
                Call PH_PY012_AddMatrixRow
                Call PH_PY012_ColumnStatusChange
            Case "1293" '// 행삭제

                 If oMat1.RowCount <> oMat1.VisualRowCount Then
                    oMat1.FlushToDataSource
    
                    While (i <= oDS_PH_PY012B.Size - 1)
                        If oDS_PH_PY012B.GetValue("U_LineNum", i) = "" Then
                            oDS_PH_PY012B.RemoveRecord (i)
                            i = 0
                        Else
                            i = i + 1
                        End If
                    Wend
                    
                    For i = 0 To oDS_PH_PY012B.Size
                        Call oDS_PH_PY012B.setValue("U_LineNum", i, i + 1)
                    Next i
                    
                    oMat1.LoadFromDataSource
                End If
                Call PH_PY012_AddMatrixRow
                
'            '복제
'            Case "1287"
'
'                Call oForm.Freeze(True)
'                Call oDS_PH_PY012A.setValue("DocEntry", 0, "")
'
'                For i = 0 To oMat1.VisualRowCount - 1
'                    Call oMat1.FlushToDataSource
'                    Call oDS_PH_PY012B.setValue("DocEntry", i, "")
'                    Call oDS_PH_PY012B.setValue("U_PayYN", i, "N")
'                    Call oMat1.LoadFromDataSource
'                Next i
'                Call oForm.Freeze(False)
            
        End Select
    End If
    Call oForm.Freeze(False)
    Exit Sub
Raise_FormMenuEvent_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "Raise_MenuEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
    
    On Error GoTo Raise_FormDataEvent_Error
    
    If (BusinessObjectInfo.BeforeAction = True) Then
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD:     '//33
            Case et_FORM_DATA_ADD:      '//34
            Case et_FORM_DATA_UPDATE:   '//35
            Case et_FORM_DATA_DELETE:   '//36
        End Select
    ElseIf (BusinessObjectInfo.BeforeAction = False) Then
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD:     '//33
            Case et_FORM_DATA_ADD:      '//34
            Case et_FORM_DATA_UPDATE:   '//35
            Case et_FORM_DATA_DELETE:   '//36
        End Select
    End If
    Exit Sub
    
Raise_FormDataEvent_Error:
    
        Sbo_Application.SetStatusBarMessage "Raise_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True

End Sub

Public Sub Raise_RightClickEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
    
    On Error GoTo Raise_RightClickEvent_Error
    
    If pval.BeforeAction = True Then
    ElseIf pval.BeforeAction = False Then
    End If
    Select Case pval.ItemUID
    Case "Mat01"
        If pval.Row > 0 Then
            oLastItemUID = pval.ItemUID
            oLastColUID = pval.ColUID
            oLastColRow = pval.Row
        End If
    Case Else
        oLastItemUID = pval.ItemUID
        oLastColUID = ""
        oLastColRow = 0
    End Select
    Exit Sub
Raise_RightClickEvent_Error:

    Sbo_Application.SetStatusBarMessage "Raise_RightClickEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY012_AddMatrixRow()
    Dim oRow    As Long
    
    On Error GoTo PH_PY012_AddMatrixRow_Error
    
    Call oForm.Freeze(True)
    
    '//[Mat1]
    oMat1.FlushToDataSource
    oRow = oMat1.VisualRowCount
    
    Dim FirstColumnName As String
    
    '출장구분에 따른 매트릭스 컬럼 설정이 변경되므로 첫 컬럼의 컬럼명을 저장
    If oForm.Items("DestDiv").Specific.Value = "01" Then
    
        FirstColumnName = "U_Vehicle"
    
    ElseIf oForm.Items("DestDiv").Specific.Value = "02" Or oForm.Items("DestDiv").Specific.Value = "03" Then
    
        FirstColumnName = "U_FrDate"
    
    End If
    
    
    If oMat1.VisualRowCount > 0 Then
        If Trim(oDS_PH_PY012B.GetValue(FirstColumnName, oRow - 1)) <> "" Then
            If oDS_PH_PY012B.Size <= oMat1.VisualRowCount Then
                oDS_PH_PY012B.InsertRecord (oRow)
            End If
            oDS_PH_PY012B.Offset = oRow
            oDS_PH_PY012B.setValue "U_LineNum", oRow, oRow + 1 '라인번호
            oDS_PH_PY012B.setValue "U_FrDate", oRow, "" '기간시작일자
            oDS_PH_PY012B.setValue "U_ToDate", oRow, "" '기간종료일자
            oDS_PH_PY012B.setValue "U_Destinat", oRow, "" '행선지
            oDS_PH_PY012B.setValue "U_PayPstg", oRow, 0 '지급율
            oDS_PH_PY012B.setValue "U_Currency", oRow, "" '통화
            oDS_PH_PY012B.setValue "U_Rate", oRow, 0 '환율
            oDS_PH_PY012B.setValue "U_Vehicle", oRow, "" '차량구분
            oDS_PH_PY012B.setValue "U_FuelType", oRow, "" '유종
            oDS_PH_PY012B.setValue "U_FuelPrc", oRow, 0 '단가
            oDS_PH_PY012B.setValue "U_Distance", oRow, 0 '거리
            oDS_PH_PY012B.setValue "U_FoodNum", oRow, 0 '식수
            oDS_PH_PY012B.setValue "U_TransExp", oRow, 0 '교통비
            oDS_PH_PY012B.setValue "U_InsurExp", oRow, 0 '보험료
            oDS_PH_PY012B.setValue "U_AirpExp", oRow, 0 '공항세
            oDS_PH_PY012B.setValue "U_DayExp", oRow, 0 '일비
            oDS_PH_PY012B.setValue "U_FDayExp", oRow, 0 '일비(외화)
            oDS_PH_PY012B.setValue "U_LodgExp", oRow, 0 '숙박비
            oDS_PH_PY012B.setValue "U_FoodExp", oRow, 0 '숙박비(외화)
            oDS_PH_PY012B.setValue "U_ParkExp", oRow, 0 '주차비
            oDS_PH_PY012B.setValue "U_TollExp", oRow, 0 '도로비
            oDS_PH_PY012B.setValue "U_TotalExp", oRow, 0 '합계
            oMat1.LoadFromDataSource
        Else
            oDS_PH_PY012B.Offset = oRow - 1
            oDS_PH_PY012B.setValue "U_LineNum", oRow - 1, oRow '라인번호
            oDS_PH_PY012B.setValue "U_FrDate", oRow - 1, "" '기간시작일자
            oDS_PH_PY012B.setValue "U_ToDate", oRow - 1, "" '기간종료일자
            oDS_PH_PY012B.setValue "U_Destinat", oRow - 1, "" '행선지
            oDS_PH_PY012B.setValue "U_PayPstg", oRow - 1, 0 '지급율
            oDS_PH_PY012B.setValue "U_Currency", oRow - 1, "" '통화
            oDS_PH_PY012B.setValue "U_Rate", oRow - 1, 0 '환율
            oDS_PH_PY012B.setValue "U_Vehicle", oRow - 1, "" '차량구분
            oDS_PH_PY012B.setValue "U_FuelType", oRow - 1, "" '유종
            oDS_PH_PY012B.setValue "U_FuelPrc", oRow - 1, 0 '단가
            oDS_PH_PY012B.setValue "U_Distance", oRow - 1, 0 '거리
            oDS_PH_PY012B.setValue "U_FoodNum", oRow - 1, 0 '식수
            oDS_PH_PY012B.setValue "U_TransExp", oRow - 1, 0 '교통비
            oDS_PH_PY012B.setValue "U_InsurExp", oRow - 1, 0 '보험료
            oDS_PH_PY012B.setValue "U_AirpExp", oRow - 1, 0 '공항세
            oDS_PH_PY012B.setValue "U_DayExp", oRow - 1, 0 '일비
            oDS_PH_PY012B.setValue "U_FDayExp", oRow - 1, 0 '일비(외화)
            oDS_PH_PY012B.setValue "U_LodgExp", oRow - 1, 0 '숙박비
            oDS_PH_PY012B.setValue "U_FoodExp", oRow - 1, 0 '숙박비(외화)
            oDS_PH_PY012B.setValue "U_ParkExp", oRow - 1, 0 '주차비
            oDS_PH_PY012B.setValue "U_TollExp", oRow - 1, 0 '도로비
            oDS_PH_PY012B.setValue "U_TotalExp", oRow - 1, 0 '합계
            oMat1.LoadFromDataSource
        End If
    ElseIf oMat1.VisualRowCount = 0 Then
        oDS_PH_PY012B.Offset = oRow
        oDS_PH_PY012B.setValue "U_LineNum", oRow, oRow + 1 '라인번호
        oDS_PH_PY012B.setValue "U_FrDate", oRow, "" '기간시작일자
        oDS_PH_PY012B.setValue "U_ToDate", oRow, "" '기간종료일자
        oDS_PH_PY012B.setValue "U_Destinat", oRow, "" '행선지
        oDS_PH_PY012B.setValue "U_PayPstg", oRow, 0 '지급율
        oDS_PH_PY012B.setValue "U_Currency", oRow, "" '통화
        oDS_PH_PY012B.setValue "U_Rate", oRow, 0 '환율
        oDS_PH_PY012B.setValue "U_Vehicle", oRow, "" '차량구분
        oDS_PH_PY012B.setValue "U_FuelType", oRow, "" '유종
        oDS_PH_PY012B.setValue "U_FuelPrc", oRow, 0 '단가
        oDS_PH_PY012B.setValue "U_Distance", oRow, 0 '거리
        oDS_PH_PY012B.setValue "U_FoodNum", oRow, 0 '식수
        oDS_PH_PY012B.setValue "U_TransExp", oRow, 0 '교통비
        oDS_PH_PY012B.setValue "U_InsurExp", oRow, 0 '보험료
        oDS_PH_PY012B.setValue "U_AirpExp", oRow, 0 '공항세
        oDS_PH_PY012B.setValue "U_DayExp", oRow, 0 '일비
        oDS_PH_PY012B.setValue "U_FDayExp", oRow, 0 '일비(외화)
        oDS_PH_PY012B.setValue "U_LodgExp", oRow, 0 '숙박비
        oDS_PH_PY012B.setValue "U_FoodExp", oRow, 0 '숙박비(외화)
        oDS_PH_PY012B.setValue "U_ParkExp", oRow, 0 '주차비
        oDS_PH_PY012B.setValue "U_TollExp", oRow, 0 '도로비
        oDS_PH_PY012B.setValue "U_TotalExp", oRow, 0 '합계
        oMat1.LoadFromDataSource
    End If
    
    Call oForm.Freeze(False)
    Exit Sub
PH_PY012_AddMatrixRow_Error:
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY012_AddMatrixRow_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Sub PH_PY012_FormClear()
On Error GoTo PH_PY012_FormClear_Error
    Dim DocEntry As String
    DocEntry = MDC_GetData.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY012'", "")
    If DocEntry = 0 Then
        oForm.Items("DocEntry").Specific.Value = 1
    Else
        oForm.Items("DocEntry").Specific.Value = DocEntry
    End If
    Exit Sub
PH_PY012_FormClear_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY012_FormClear_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

'Private Sub PH_PY012_FormResize()
''******************************************************************************
''Function ID : PH_PY012_FormResize()
''해당모듈 : PH_PY012
''기능 : Form의 크기 변경 시 아이템들의 위치 및 크기 동적 변경
''인수 : 없음
''반환값 : 없음
''특이사항 : 없음
''******************************************************************************
'On Error GoTo PH_PY012_FormResize_Error
'
'    oForm.Items("Mat01").Left = 10
'    oForm.Items("Mat01").Top = 110
'    oForm.Items("Mat01").Height = oForm.Height - 330
'    oForm.Items("Mat01").Width = oForm.Width - 20
'
'
'    Exit Sub
'
'PH_PY012_FormResize_Error:
'    Sbo_Application.SetStatusBarMessage "PH_PY012_FormResize_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
'End Sub

Private Sub PH_PY012_ColumnStatusChange()
'******************************************************************************
'Function ID : PH_PY012_FormResize()
'해당모듈 : PH_PY012
'기능 : 출장구분에 따른 매트릭스 컬럼 내용 변경
'인수 : 없음
'반환값 : 없음
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY012_ColumnStatusChange_Error

    If oForm.Items("DestDiv").Specific.Value = "01" Then '공용
    
        oMat1.Columns("FrDate").Visible = False '기간시작일자
        oMat1.Columns("ToDate").Visible = False '기간종료일자
        oMat1.Columns("Destinat").Visible = False '행선지
        oMat1.Columns("PayPstg").Visible = False '지급율
        oMat1.Columns("Currency").Visible = False '통화
        oMat1.Columns("Rate").Visible = False '환율
        oMat1.Columns("Vehicle").Visible = True '차량구분
        oMat1.Columns("FuelType").Visible = True '유종
        oMat1.Columns("FuelPrc").Visible = True '단가
        oMat1.Columns("Distance").Visible = True '거리
        oMat1.Columns("FoodNum").Visible = True '식수
        oMat1.Columns("TransExp").Visible = True '교통비
        oMat1.Columns("InsurExp").Visible = False '보험료
        oMat1.Columns("AirpExp").Visible = False '공항세
        oMat1.Columns("DayExp").Visible = True '일비
        oMat1.Columns("FDayExp").Visible = False '일비(외화)
        oMat1.Columns("LodgExp").Visible = False '숙박비
        oMat1.Columns("FLodgExp").Visible = False '숙박비(외화)
        oMat1.Columns("FoodExp").Visible = True '식비
        oMat1.Columns("ParkExp").Visible = True '주차비
        oMat1.Columns("TollExp").Visible = True '도로비
        oMat1.Columns("TotalExp").Visible = True '합계
        
'        oMat1.Columns("TransExp").Editable = False '교통비 Editable 설정
'        oMat1.Columns("FoodExp").Editable = False '식비 Editable 설정
        
        Call oMat1.AutoResizeColumns
            
    ElseIf oForm.Items("DestDiv").Specific.Value = "02" Then '국내출장
    
        oMat1.Columns("FrDate").Visible = True '기간시작일자
        oMat1.Columns("ToDate").Visible = True '기간종료일자
        oMat1.Columns("Destinat").Visible = True '행선지
        oMat1.Columns("PayPstg").Visible = True '지급율
        oMat1.Columns("Currency").Visible = False '통화
        oMat1.Columns("Rate").Visible = False '환율
        oMat1.Columns("Vehicle").Visible = False '차량구분
        oMat1.Columns("FuelType").Visible = False '유종
        oMat1.Columns("FuelPrc").Visible = False '단가
        oMat1.Columns("Distance").Visible = False '거리
        oMat1.Columns("FoodNum").Visible = False '식수
        oMat1.Columns("TransExp").Visible = True '교통비
        oMat1.Columns("InsurExp").Visible = True '보험료
        oMat1.Columns("AirpExp").Visible = True '공항세
        oMat1.Columns("DayExp").Visible = True '일비
        oMat1.Columns("FDayExp").Visible = False '일비(외화)
        oMat1.Columns("LodgExp").Visible = True '숙박비
        oMat1.Columns("FLodgExp").Visible = False '숙박비(외화)
        oMat1.Columns("FoodExp").Visible = True '식비
        oMat1.Columns("ParkExp").Visible = False '주차비
        oMat1.Columns("TollExp").Visible = False '도로비
        oMat1.Columns("TotalExp").Visible = True '합계
        
'        oMat1.Columns("TransExp").Editable = True '교통비 Editable 설정
'        oMat1.Columns("FoodExp").Editable = True '식비 Editable 설정
        
        Call oMat1.AutoResizeColumns
    
    ElseIf oForm.Items("DestDiv").Specific.Value = "03" Then '해외출장
    
        oMat1.Columns("FrDate").Visible = True '기간시작일자
        oMat1.Columns("ToDate").Visible = True '기간종료일자
        oMat1.Columns("Destinat").Visible = True '행선지
        oMat1.Columns("PayPstg").Visible = True '지급율
        oMat1.Columns("Currency").Visible = True '통화
        oMat1.Columns("Rate").Visible = True '환율
        oMat1.Columns("Vehicle").Visible = False '차량구분
        oMat1.Columns("FuelType").Visible = False '유종
        oMat1.Columns("FuelPrc").Visible = False '단가
        oMat1.Columns("Distance").Visible = False '거리
        oMat1.Columns("FoodNum").Visible = False '식수
        oMat1.Columns("TransExp").Visible = True '교통비
        oMat1.Columns("InsurExp").Visible = True '보험료
        oMat1.Columns("AirpExp").Visible = True '공항세
        oMat1.Columns("DayExp").Visible = True '일비
        oMat1.Columns("FDayExp").Visible = True '일비(외화)
        oMat1.Columns("LodgExp").Visible = True '숙박비
        oMat1.Columns("FLodgExp").Visible = True '숙박비(외화)
        oMat1.Columns("FoodExp").Visible = True '식비
        oMat1.Columns("ParkExp").Visible = False '주차비
        oMat1.Columns("TollExp").Visible = False '도로비
        oMat1.Columns("TotalExp").Visible = True '합계
        
'        oMat1.Columns("TransExp").Editable = True '교통비 Editable 설정
'        oMat1.Columns("FoodExp").Editable = True '식비 Editable 설정
        
        Call oMat1.AutoResizeColumns
    
    End If

    Exit Sub

PH_PY012_ColumnStatusChange_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY012_ColumnStatusChange_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY012_DataValidCheck() As Boolean
On Error GoTo PH_PY012_DataValidCheck_Error
    PH_PY012_DataValidCheck = False
    Dim i       As Long
    Dim sQry    As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    '사업장
    If Trim(oDS_PH_PY012A.GetValue("U_CLTCOD", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "사업장은 필수입니다.", bmt_Short, True
        oForm.Items("CLTCOD").CLICK ct_Regular
        PH_PY012_DataValidCheck = False
        Exit Function
    End If

    '사원
    If Trim(oDS_PH_PY012A.GetValue("U_MSTCOD", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "사원정보는 필수입니다.", bmt_Short, True
        oForm.Items("MSTCOD").CLICK ct_Regular
        PH_PY012_DataValidCheck = False
        Exit Function
    End If

    '기간(From)
    If Trim(oDS_PH_PY012A.GetValue("U_FrDate", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "기간 시작은 필수입니다.", bmt_Short, True
        oForm.Items("FrDate").CLICK ct_Regular
        PH_PY012_DataValidCheck = False
        Exit Function
    End If
    
    '기간(To)
    If Trim(oDS_PH_PY012A.GetValue("U_ToDate", 0)) = "" Then
        Sbo_Application.SetStatusBarMessage "기간 종료는 필수입니다.", bmt_Short, True
        oForm.Items("ToDate").CLICK ct_Regular
        PH_PY012_DataValidCheck = False
        Exit Function
    End If
        
    '라인
    If oMat1.VisualRowCount > 1 Then
        For i = 1 To oMat1.VisualRowCount - 1
        
'            '학교
'            If oMat1.Columns("SchCls").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "학교는 필수입니다.", bmt_Short, True
'                oMat1.Columns("SchCls").Cells(i).CLICK ct_Regular
'                PH_PY012_DataValidCheck = False
'                Exit Function
'            End If
'
'            '학교명
'            If oMat1.Columns("SchName").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "학교명은 필수입니다.", bmt_Short, True
'                oMat1.Columns("SchName").Cells(i).CLICK ct_Regular
'                PH_PY012_DataValidCheck = False
'                Exit Function
'            End If
'
'            '학년
'            If oMat1.Columns("Grade").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "학년은 필수입니다.", bmt_Short, True
'                oMat1.Columns("Grade").Cells(i).CLICK ct_Regular
'                PH_PY012_DataValidCheck = False
'                Exit Function
'            End If
'
'            '회차
'            If oMat1.Columns("Count").Cells(i).Specific.Value = "" Then
'                Sbo_Application.SetStatusBarMessage "회차는 필수입니다.", bmt_Short, True
'                oMat1.Columns("Count").Cells(i).CLICK ct_Regular
'                PH_PY012_DataValidCheck = False
'                Exit Function
'            End If
            
        Next
    Else
        Sbo_Application.SetStatusBarMessage "라인 데이터가 없습니다.", bmt_Short, True
        PH_PY012_DataValidCheck = False
        Exit Function
    End If
    
        oMat1.FlushToDataSource
    '// Matrix 마지막 행 삭제(DB 저장시)
    If oDS_PH_PY012B.Size > 1 Then oDS_PH_PY012B.RemoveRecord (oDS_PH_PY012B.Size - 1)
    
    oMat1.LoadFromDataSource
    
    PH_PY012_DataValidCheck = True
    
    Set oRecordSet = Nothing
    
    Exit Function
    
PH_PY012_DataValidCheck_Error:
    Set oRecordSet = Nothing
    PH_PY012_DataValidCheck = False
    Sbo_Application.SetStatusBarMessage "PH_PY012_DataValidCheck_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY012_MTX01()

    '//메트릭스에 데이터 로드
    
    Dim i       As Long
    Dim sQry    As String
    
    Dim Param01 As String
    Dim Param02 As String
    Dim Param03 As String
    Dim Param04 As String
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    
    On Error GoTo PH_PY012_MTX01_Error

    Call oForm.Freeze(True)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

    Param01 = oForm.Items("Param01").Specific.Value
    Param02 = oForm.Items("Param01").Specific.Value
    Param03 = oForm.Items("Param01").Specific.Value
    Param04 = oForm.Items("Param01").Specific.Value
    
    sQry = "SELECT 10"
    Call oRecordSet.DoQuery(sQry)
    
    oMat1.Clear
    oMat1.FlushToDataSource
    oMat1.LoadFromDataSource
    
    If (oRecordSet.RecordCount = 0) Then
        Call MDC_Com.MDC_GF_Message("결과가 존재하지 않습니다.", "E")
        GoTo PH_PY012_MTX01_Exit
    End If
    
    Dim ProgressBar01 As SAPbouiCOM.ProgressBar
    Set ProgressBar01 = Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, False)
    
    For i = 0 To oRecordSet.RecordCount - 1
        If i <> 0 Then
            oDS_PH_PY012B.InsertRecord (i)
        End If
        oDS_PH_PY012B.Offset = i
        Call oDS_PH_PY012B.setValue("U_COL01", i, oRecordSet.Fields(0).Value)
        Call oDS_PH_PY012B.setValue("U_COL02", i, oRecordSet.Fields(1).Value)
        oRecordSet.MoveNext
        ProgressBar01.Value = ProgressBar01.Value + 1
        ProgressBar01.Text = ProgressBar01.Value & "/" & oRecordSet.RecordCount & "건 조회중...!"
    Next
    oMat1.LoadFromDataSource
    oMat1.AutoResizeColumns
    oForm.Update
    
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Exit Sub
PH_PY012_MTX01_Exit:
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    If Not ProgressBar01 Is Nothing Then
        ProgressBar01.Stop
    End If
    Exit Sub
PH_PY012_MTX01_Error:
    ProgressBar01.Stop
    Set ProgressBar01 = Nothing
    Set oRecordSet = Nothing
    Call oForm.Freeze(False)
    Sbo_Application.SetStatusBarMessage "PH_PY012_MTX01_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub

Function PH_PY012_Validate(ByVal ValidateType As String) As Boolean
On Error GoTo PH_PY012_Validate_Error
    PH_PY012_Validate = True
    Dim i, j As Long
    Dim sQry As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY012A] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "Y" Then
        Sbo_Application.SetStatusBarMessage "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", bmt_Short, True
        PH_PY012_Validate = False
        GoTo PH_PY012_Validate_Exit
    End If
'
    If ValidateType = "수정" Then

    ElseIf ValidateType = "행삭제" Then

    ElseIf ValidateType = "취소" Then

    End If
    Set oRecordSet = Nothing
    Exit Function
PH_PY012_Validate_Exit:
    Set oRecordSet = Nothing
    Exit Function
PH_PY012_Validate_Error:
    PH_PY012_Validate = False
    Sbo_Application.SetStatusBarMessage "PH_PY012_Validate_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Function PH_PY012_GetFuelPrc(ByVal pBPLId As String, ByVal pStdYear As String, ByVal pStdMonth As String, ByVal pFuelType As String) As Double
'******************************************************************************
'Function ID : PH_PY012_GetFuelPrc()
'해당모듈 : PH_PY012
'기능 : 유류단가 가져오기
'인수 : pBPLId:사업장, pStdYear:기준년도, pStdMonth:기준월, pFuelType:유종
'반환값 : 유류단가
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY012_GetFuelPrc_Error

    Dim loopCount As Integer
    Dim sQry As String
    Dim CheckAmt
    
    Dim oRecordSet As SAPbobsCOM.Recordset
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    sQry = "           SELECT      T0.U_Year AS [StdYear],"
    sQry = sQry & "                T1.U_Month AS [StdMonth],"
    sQry = sQry & "                T1.U_Gasoline AS [Gasoline],"
    sQry = sQry & "                T1.U_Diesel AS [Diesel],"
    sQry = sQry & "                T1.U_LPG AS [LPG]"
    sQry = sQry & " FROM       [@PH_PY007A] AS T0"
    sQry = sQry & "                INNER JOIN"
    sQry = sQry & "                [@PH_PY007B] AS T1"
    sQry = sQry & "                    ON T0.Code = T1.Code"
    sQry = sQry & " WHERE      T0.U_CLTCOD = '" & pBPLId & "'"
    sQry = sQry & "                AND T0.U_Year = '" & pStdYear & "'"
    sQry = sQry & "                AND T1.U_Month = '" & pStdMonth & "'"
    
    Call oRecordSet.DoQuery(sQry)
    
    If pFuelType = "1" Then '휘발유
        PH_PY012_GetFuelPrc = oRecordSet.Fields("Gasoline").Value
    ElseIf pFuelType = "2" Then '가스
        PH_PY012_GetFuelPrc = oRecordSet.Fields("LPG").Value
    ElseIf pFuelType = "3" Then '경유
        PH_PY012_GetFuelPrc = oRecordSet.Fields("Diesel").Value
    End If

    Exit Function
PH_PY012_GetFuelPrc_Error:
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "PH_PY012_GetFuelPrc_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Function

Private Sub PH_PY012_CalculateTotalExp(ByVal pRow As Integer)
'******************************************************************************
'Function ID : PH_PY012_CalculateTotalExp()
'해당모듈 : PH_PY012
'기능 : 합계 계산
'인수 : pRow : pval.Row 값
'반환값 : 비용 합계
'특이사항 : 없음
'******************************************************************************
On Error GoTo PH_PY012_CalculateTotalExp_Error
    
    Dim FuelType As String '유종
    Dim FuelPrc As Double '단가
    Dim Distance As Double '거리
    Dim FoodNum As Double '식수
    Dim TransExp As Double '교통비
    Dim InsurExp As Double '보험료
    Dim AirpExp As Double '공항세
    Dim DayExp As Double '일비
    Dim LodgExp As Double '숙박비
    Dim FoodExp As Double '식비
    Dim ParkExp As Double '주차비
    Dim TollExp As Double '도로비
    Dim TotalExp As Double '합계

    oMat1.FlushToDataSource
                                
    '합계(교통비+일비+식비+주차비+도로비+숙박비+보험료+공항세)
    TransExp = oDS_PH_PY012B.GetValue("U_TransExp", pRow) '교통비
    DayExp = oDS_PH_PY012B.GetValue("U_DayExp", pRow) '일비
    FoodExp = oDS_PH_PY012B.GetValue("U_FoodExp", pRow) '식비
    ParkExp = oDS_PH_PY012B.GetValue("U_ParkExp", pRow) '주차비
    TollExp = oDS_PH_PY012B.GetValue("U_TollExp", pRow) '도로비
    LodgExp = oDS_PH_PY012B.GetValue("U_LodgExp", pRow) '숙박비
    InsurExp = oDS_PH_PY012B.GetValue("U_InsurExp", pRow) '보험료
    AirpExp = oDS_PH_PY012B.GetValue("U_AirpExp", pRow) '공항세
    
    TotalExp = TransExp + DayExp + FoodExp + ParkExp + TollExp + LodgExp + InsurExp + AirpExp
    
    Call oDS_PH_PY012B.setValue("U_TotalExp", pRow, TotalExp) '합계
    
    oMat1.LoadFromDataSource

    Exit Sub
PH_PY012_CalculateTotalExp_Error:
    Sbo_Application.SetStatusBarMessage "PH_PY012_CalculateTotalExp_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
End Sub
