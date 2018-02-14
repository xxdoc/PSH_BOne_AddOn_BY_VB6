VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZPY505"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//  SAP MANAGE UI API 2004 SDK Sample
'//****************************************************************************
'//  File           : ZPY505.cls
'//  Module         : 인사관리>정산관리
'//  Desc           : 기부금명세 등록
'//  FormType       : 2000060505
'//  Create Date    : 2006.01.15
'//  Modified Date  :
'//  Creator        : Ham Mi Kyoung
'//  Modifier       :
'//  Copyright  (c) Morning Data
'//****************************************************************************
Option Explicit

Public oFormUniqueID    As String
Public oForm            As SAPbouiCOM.Form
Private oDS_ZPY505H     As SAPbouiCOM.DBDataSource    '시스템코드 헤더
Private oDS_ZPY505L     As SAPbouiCOM.DBDataSource    '시스템코드 라인
Private oMat1           As SAPbouiCOM.Matrix
Private Last_Item       As String                     '클래스에서 선택한 마지막 아이템 Uid값
Private Col_Last_Uid    As String                     '마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
Private Col_Last_Row    As Long                       '마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
Private oOLDCHK         As String

Private Sub FormItemEnabled()
    If oForm.Mode = fm_FIND_MODE Then
        oForm.Items("JSNYER").Enabled = True
        oForm.Items("MSTCOD").Enabled = True
        oForm.Items("MSTNAM").Enabled = True
        oForm.Items("DocNum").Enabled = True
        oForm.Items("ENDCHK").Enabled = True
    ElseIf oForm.Mode = fm_ADD_MODE Then
        oForm.Items("JSNYER").Enabled = True
        oForm.Items("MSTCOD").Enabled = True
        oForm.Items("MSTNAM").Enabled = False
        oForm.Items("DocNum").Enabled = False
        oForm.Items("ENDCHK").Enabled = True
    ElseIf oForm.Mode = fm_OK_MODE Then
        oForm.Items("JSNYER").Enabled = False
        oForm.Items("MSTCOD").Enabled = False
        oForm.Items("MSTNAM").Enabled = False
        oForm.Items("DocNum").Enabled = False
        '// 년마감된것은 비활성화
        oOLDCHK = oDS_ZPY505H.GetValue("U_ENDCHK", 0)
        If MDC_SetMod.Get_ReData("U_ENDCHK", "U_JOBYER", "[@ZPY509L]", "'" & oDS_ZPY505H.GetValue("U_JSNYER", 0) & "'", " AND Code = '" & oDS_ZPY505H.GetValue("U_CLTCOD", 0) & "'") = "Y" Then
            oForm.Items("ENDCHK").Enabled = False
        Else
            oForm.Items("ENDCHK").Enabled = True
        End If
        
    End If
End Sub

Private Sub FormClear()
    Dim DocNum&
    
    DocNum = MDC_SetMod.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'ZPY505'", "")
    
    If DocNum = 0 Then
        oForm.Items("DocNum").Specific.String = 1
    Else
        oForm.Items("DocNum").Specific.String = DocNum
    End If
    Call FlushToItemValue("JSNYER")
End Sub

Private Sub FlushToItemValue(ByVal oUID As String, Optional oRow As Long)
    Dim i&
    Dim oMast           As ZPAY_g_EmpID
    Dim TOTCNT          As Double
    Dim TOTAMT          As Double
    
    Select Case oUID
      Case "JSNYER"
            If Trim$(oForm.Items(oUID).Specific.String) <> "" Then
                ZPAY_GBL_JSNYER = oForm.Items(oUID).Specific.String
            Else
                oDS_ZPY505H.setValue "U_JSNYER", 0, ZPAY_GBL_JSNYER
            End If
            oForm.Items(oUID).Update
      Case "MSTCOD"
            If oForm.Items(oUID).Specific.String = "" Then
                oDS_ZPY505H.setValue "U_MSTCOD", 0, ""
                oDS_ZPY505H.setValue "U_MSTNAM", 0, ""
                oDS_ZPY505H.setValue "U_EmpID", 0, ""
                oDS_ZPY505H.setValue "U_CLTCOD", 0, ""
            Else
                oDS_ZPY505H.setValue "U_MSTCOD", 0, UCase$(oForm.Items(oUID).Specific.String)
                oMast = MDC_SetMod.Get_EmpID_InFo(oForm.Items(oUID).Specific.String)
                oDS_ZPY505H.setValue "U_MSTNAM", 0, oMast.MSTNAM
                oDS_ZPY505H.setValue "U_EmpID", 0, oMast.EmpID
                oDS_ZPY505H.setValue "U_CLTCOD", 0, oMast.CLTCOD
            End If
            oForm.Items("MSTNAM").Update
            oForm.Items("EmpID").Update
            oForm.Items("CLTCOD").Update
            oForm.Items(oUID).Update
    End Select
    
    '//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ/
    Select Case oUID
        Case "Col5", "Col6"
            oMat1.FlushToDataSource
            For i = 1 To oMat1.VisualRowCount
                TOTCNT = TOTCNT + Val(oDS_ZPY505L.GetValue("U_GBUCNT", i - 1))
                TOTAMT = TOTAMT + Val(oDS_ZPY505L.GetValue("U_GBUAMT", i - 1))
            Next i
            oDS_ZPY505H.setValue "U_TOTCNT", 0, TOTCNT
            oDS_ZPY505H.setValue "U_TOTAMT", 0, TOTAMT
            oForm.Items("TOTCNT").Update
            oForm.Items("TOTAMT").Update
            
            oDS_ZPY505L.Offset = oRow - 1
'            oMat1.SetLineData oRow
        Case "Col8"
             oDS_ZPY505L.Offset = oRow - 1
             oDS_ZPY505L.setValue "U_FAMNAM", oRow - 1, oMat1.Columns(oUID).Cells(oRow).Specific.Value
             Call Display_GibuMan(oRow - 1)
             oMat1.SetLineData oRow
        Case "Col1"
            oMat1.FlushToDataSource
            oDS_ZPY505L.Offset = oRow - 1
            
            If oRow = oMat1.RowCount And Trim$(oDS_ZPY505L.GetValue("U_GBUYMM", oRow - 1)) <> "" Then
                Matrix_AddRow oRow
                oMat1.Columns("Col1").Cells(oRow).CLICK ct_Regular
            End If
    End Select
End Sub

Private Sub Display_GibuMan(sRow As Long)
On Error GoTo Error_Message
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim MSTCOD          As String
    Dim JSNYER          As String
    Dim FAMNAM          As String
    
    
    JSNYER = oDS_ZPY505H.GetValue("U_JSNYER", 0)
    MSTCOD = oDS_ZPY505H.GetValue("U_MSTCOD", 0)
    FAMNAM = oDS_ZPY505L.GetValue("U_FAMNAM", sRow)
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)

  '// 기부자명입력시 소득자료등록화면의 부양가족명세중 본인에 해당하는 나머지 정보 표시해줌.
   sQry = "SELECT  T0.U_FAMNAM AS FAMNAM, T0.U_CHKINT AS INTGBN, T0.U_FAMPER AS PERNBR, "
   sQry = sQry & " CASE T0.U_CHKCOD WHEN '0' THEN '1' WHEN '3' THEN '2' WHEN '4' THEN '3' ELSE '' END AS GWANGE "
   sQry = sQry & " FROM [@ZPY501L] T0 INNER JOIN [@ZPY501H] T1 ON T0.DocEntry = T1.DocEntry"
   sQry = sQry & " WHERE T1.U_JSNYER = '" & Trim$(JSNYER) & "'"
   sQry = sQry & " AND   T1.U_MSTCOD = '" & Trim$(MSTCOD) & "'"
   sQry = sQry & " AND   T0.U_FAMNAM = '" & Trim$(FAMNAM) & "'"
   oRecordSet.DoQuery sQry
   If oRecordSet.RecordCount > 0 Then
        oDS_ZPY505L.setValue "U_INTGBN", sRow, oRecordSet.Fields("INTGBN").Value
        oDS_ZPY505L.setValue "U_PERNBR", sRow, oRecordSet.Fields("PERNBR").Value
        oDS_ZPY505L.setValue "U_GWANGE", sRow, oRecordSet.Fields("GWANGE").Value
   End If

    Set oRecordSet = Nothing
   Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oRecordSet = Nothing
     Sbo_Application.StatusBar.SetText "Display_GibuMan Error:" & Err.Number & Space$(10) & Err.Description, bmt_Short, smt_Error
End Sub

Private Function MatrixSpaceLineDel() As Boolean
'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
'저장할 데이터의 유효성을 점검한다
'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
On Error GoTo Error_Message
    Dim iRow        As Long
    Dim kRow        As Long
    Dim ErrNum      As Integer
    Dim Chk_Data    As String
'    Dim CHAAMT      As Double
    Dim GovidChk    As String
    ErrNum = 0
    '/ 헤더부분 체크
    Select Case True
      Case Len(Trim$(oDS_ZPY505H.GetValue("U_JSNYER", 0))) <> 4
           ErrNum = 4
           GoTo Error_Message
      Case oDS_ZPY505H.GetValue("U_MSTCOD", 0) = ""
           ErrNum = 5
           GoTo Error_Message
      Case oDS_ZPY505H.GetValue("U_CLTCOD", 0) = ""
           ErrNum = 17
           GoTo Error_Message
           
    End Select
    '/ 주민번호체크유무
    GovidChk = MDC_SetMod.Get_ReData("ISNULL(T0.U_GovIDChk,'N')", "T1.U_MSTCOD", "[@PH_PY005A] T0 INNER JOIN [@PH_PY001A] T1 ON T0.CODE = T1.U_CLTCOD", "'" & Trim$(oDS_ZPY505H.GetValue("U_MSTCOD", 0)) & "'", "")

    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '화면상의 메트릭스에 입력된 내용을 모두 디비데이터소스로 넘긴다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oMat1.FlushToDataSource
        
    If oMat1.RowCount = 1 Then          '// Mat1에 값이 있는지 확인 (ErrorNumber : 1)
        ErrNum = 1
        GoTo Error_Message
    End If
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//마지막 행 하나를 빼고 i=0부터 시작하므로 하나를 빼므로
    '//oMat1.RowCount - 2가 된다..반드시 들어 가야 하는 필수값을 확인한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    For iRow = 0 To oMat1.VisualRowCount - 2      '// Mat1에 입력값이 올바르게 들어갔는지 확인 (ErrorNumber : 3)
        oDS_ZPY505L.Offset = iRow
'        CHAAMT = Val(oDS_ZPY505L.GetValue("U_BEFAMT", irow)) + Val(oDS_ZPY505L.GetValue("U_CURAMT", irow)) + Val(oDS_ZPY505L.GetValue("U_CHAAMT", irow))
        If Trim$(oDS_ZPY505L.GetValue("U_GBUYMM", iRow)) = "" Then
            ErrNum = 2
            oMat1.Columns("Col1").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf MDC_SetMod.ChkYearMonth(oDS_ZPY505L.GetValue("U_GBUYMM", iRow)) = False Then
            ErrNum = 7
            oMat1.Columns("Col1").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Trim$(oDS_ZPY505L.GetValue("U_GBUNAM", iRow)) = "" Then
            ErrNum = 6
            oMat1.Columns("Col3").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Val(oDS_ZPY505L.GetValue("U_GBUCNT", iRow)) = 0 Then
            ErrNum = 8
            oMat1.Columns("Col5").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Val(oDS_ZPY505L.GetValue("U_GBUAMT", iRow)) = 0 Then
            ErrNum = 9
            oMat1.Columns("Col6").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Trim$(oDS_ZPY505L.GetValue("U_FAMNAM", iRow)) = "" Then
            ErrNum = 10
            oMat1.Columns("Col8").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Trim$(oDS_ZPY505L.GetValue("U_GWANGE", iRow)) = "" Then
            ErrNum = 11
            oMat1.Columns("Col9").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Trim$(oDS_ZPY505L.GetValue("U_INTGBN", iRow)) = "" Then
            ErrNum = 12
            oMat1.Columns("Col10").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
        ElseIf Trim$(oDS_ZPY505L.GetValue("U_PERNBR", iRow)) = "" Then
            ErrNum = 13
            oMat1.Columns("Col11").Cells(iRow + 1).CLICK ct_Regular
            GoTo Error_Message
'        ElseIf Trim$(oDS_ZPY505L.GetValue("U_GBUCOD", irow)) = "31" And CHAAMT <> Val(oDS_ZPY505L.GetValue("U_GBUAMT", irow)) Then
'            ErrNum = 14
'            oMat1.Columns("Col14").Cells(irow + 1).CLICK ct_Regular
'            GoTo Error_Message
'        ElseIf Trim$(oDS_ZPY505L.GetValue("U_GBUCOD", irow)) <> "31" And CHAAMT <> 0 Then
'            ErrNum = 15
'            oMat1.Columns("Col14").Cells(irow + 1).CLICK ct_Regular
'            GoTo Error_Message
        Else
            '// 6.주민번호 오류 체크
            If Trim$(GovidChk) = "Y" And Len(oDS_ZPY505L.GetValue("U_PERNBR", iRow)) > 0 Then
                If GovIDCheck(oDS_ZPY505L.GetValue("U_PERNBR", iRow)) = False Then
                    ErrNum = 16
                    oMat1.Columns("Col11").Cells(iRow + 1).CLICK ct_Regular
                    GoTo Error_Message
                End If
            End If
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            '중복체크작업
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            Chk_Data = Trim$(oDS_ZPY505L.GetValue("U_GBUYMM", iRow)) & Trim$(oDS_ZPY505L.GetValue("U_GBUCOD", iRow)) & _
                       Trim$(oDS_ZPY505L.GetValue("U_GBUNBR", iRow)) & Trim$(oDS_ZPY505L.GetValue("U_PERNBR", iRow))
            For kRow = iRow + 1 To oMat1.VisualRowCount - 2
                oDS_ZPY505L.Offset = kRow
                If Trim$(Chk_Data) = Trim$(oDS_ZPY505L.GetValue("U_GBUYMM", kRow)) & Trim$(oDS_ZPY505L.GetValue("U_GBUCOD", kRow)) & _
                                     Trim$(oDS_ZPY505L.GetValue("U_GBUNBR", kRow)) & Trim$(oDS_ZPY505L.GetValue("U_PERNBR", kRow)) Then
                    ErrNum = 3
                    oMat1.Columns("Col1").Cells(iRow + 1).CLICK ct_Regular
                    GoTo Error_Message
                End If
            Next kRow
        End If
    Next iRow
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
    '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oDS_ZPY505L.RemoveRecord oDS_ZPY505L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '행을 삭제하였으니 DB데이터 소스를 다시 가져온다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oMat1.LoadFromDataSource
    
    MatrixSpaceLineDel = True
    Exit Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "입력할 데이터가 없습니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "기부연월은 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "기부연월의 기부코드, 기부처 사업자번호가 중복입력되었습니다. 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 4 Then
        Sbo_Application.StatusBar.SetText "귀속년도는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 5 Then
        Sbo_Application.StatusBar.SetText "사원번호는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 6 Then
        Sbo_Application.StatusBar.SetText "기부처 상호는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 7 Then
        Sbo_Application.StatusBar.SetText "기부연월을 확인하세요. Ex)2006년1월->200601", bmt_Short, smt_Error
    ElseIf ErrNum = 8 Then
        Sbo_Application.StatusBar.SetText "기부금건수가 0입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 9 Then
        Sbo_Application.StatusBar.SetText "기부금금액가 0입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 10 Then
        Sbo_Application.StatusBar.SetText "기부자 성명은 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 11 Then
        Sbo_Application.StatusBar.SetText "기부자 관계코드는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 12 Then
        Sbo_Application.StatusBar.SetText "기부자 내외국인구분은 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 13 Then
        Sbo_Application.StatusBar.SetText "기부자 주민등록번호는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
'    ElseIf ErrNum = 14 Then
'        Sbo_Application.StatusBar.SetText "31-공익법인기부금신탁일 경우 기부금액과 (이월액잔액+해당과세기간공제액+이월액)이 일치하지 않습니다.", bmt_Short, smt_Error
'    ElseIf ErrNum = 15 Then
'        Sbo_Application.StatusBar.SetText "31-공익법인기부금신탁가 아닐경우 (이월액잔액+해당과세기간공제액+이월액)는 입력하지 않습니다.", bmt_Short, smt_Error
    ElseIf ErrNum = 16 Then
        Sbo_Application.StatusBar.SetText "주민등록번호를 확인하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 17 Then
        Sbo_Application.StatusBar.SetText "자사코드는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "MatrixSpaceLineDel Error:" & Err.Number & Space$(10) & Err.Description, bmt_Short, smt_Error
    End If
    MatrixSpaceLineDel = False
End Function
'*******************************************************************
'// ItemEventHander
'*******************************************************************
Public Sub Raise_FormItemEvent(ByRef FormUID, ByRef pval As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)

    Dim sQry        As String
    Dim i           As Long
    Dim oCombo      As SAPbouiCOM.ComboBox
    Dim oColumn     As SAPbouiCOM.Column
    Dim oColumns     As SAPbouiCOM.Columns
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    On Error GoTo Raise_FormItemEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    Select Case pval.EventType
'et_ITEM_PRESSED''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_ITEM_PRESSED
            If pval.BeforeAction Then
                If pval.ItemUid = "1" Then
                        'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                        '//추가및 업데이시에
                        'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                        If oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_UPDATE_MODE Then
                            If oForm.Mode = fm_ADD_MODE Then
                                If MDC_SetMod.Value_ChkYn("[@ZPY505H]", "U_JSNYER", "'" & oForm.Items("JSNYER").Specific.String & "'", " AND U_MSTCOD = '" & oForm.Items("MSTCOD").Specific.String & "'") = False Then
                                    Sbo_Application.StatusBar.SetText "이미 저장되어져 있는 헤더의 내용과 일치합니다.", bmt_Short, smt_Error
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                        
                        If Trim$(oDS_ZPY505H.GetValue("U_ENDCHK", 0)) = "Y" And Trim$(oOLDCHK) = "Y" Then
                             Sbo_Application.StatusBar.SetText "잠금 자료입니다. 확인하여 주십시오.", bmt_Short, smt_Error
                             BubbleEvent = False
                             Exit Sub
                        ElseIf MatrixSpaceLineDel = False Then
                                BubbleEvent = False
                            End If
                        End If
                ElseIf pval.ItemUid = "CBtn1" And oForm.Items("MSTCOD").Enabled = True Then '/ ChooseBtn사원리스트
                    oForm.Items("MSTCOD").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                ElseIf pval.ItemUid = "Btn1" And (oForm.Mode = fm_ADD_MODE Or oForm.Mode = fm_OK_MODE Or oForm.Mode = fm_UPDATE_MODE) Then
                    Call BeforeBalance
                End If
            Else
                If pval.ItemUid = "1" And pval.ActionSuccess = True And oForm.Mode = fm_ADD_MODE Then
                    oForm.Mode = fm_OK_MODE
                    Call Sbo_Application.ActivateMenuItem("1282")
                End If
            End If
'et_CLICK''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_CLICK
            If pval.BeforeAction = True And pval.ItemUid <> "1000001" And pval.ItemUid <> "2" And oForm.Mode <> fm_FIND_MODE Then
                If Last_Item = "MSTCOD" Then
                    If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & oForm.Items(Last_Item).Specific.String & "'", "") = True _
                    And oForm.Items(Last_Item).Specific.String <> "" And Last_Item <> pval.ItemUid Then
                        Sbo_Application.ActivateMenuItem ("7425")
                        BubbleEvent = False
                    End If
                End If
            End If
            If pval.FormUID = oForm.uniqueID And pval.BeforeAction = True And Last_Item = "Mat1" And Col_Last_Uid = "Col1" And Col_Last_Row > 0 _
              And (Col_Last_Uid <> pval.ColUID Or Col_Last_Row <> pval.Row) And pval.ItemUid <> "1000001" And pval.ItemUid <> "2" Then
                If Col_Last_Row > oMat1.VisualRowCount Then
                    Exit Sub
                End If
            End If
'et_VALIDATE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_VALIDATE
            If pval.BeforeAction = False And pval.ItemChanged = True And _
            (pval.ItemUid = "MSTCOD" Or pval.ItemUid = "JSNYER") Then
                FlushToItemValue pval.ItemUid
            ElseIf pval.BeforeAction = False And pval.ItemChanged = True And pval.ItemUid = "Mat1" _
            And (pval.ColUID = "Col5" Or pval.ColUID = "Col6" Or pval.ColUID = "Col1" Or pval.ColUID = "Col8") Then
                FlushToItemValue pval.ColUID, pval.Row
            End If
        
'et_KEY_DOWN''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_KEY_DOWN
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            '//추가모드에서 코드이벤트가 코드에서 일어 났을때
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
             If pval.BeforeAction = True And pval.ItemUid = "MSTCOD" And pval.CharPressed = 9 And pval.FormMode <> fm_FIND_MODE Then
                If MDC_SetMod.Value_ChkYn("[@PH_PY001A]", "Code", "'" & oForm.Items(pval.ItemUid).Specific.String & "'", "") = True Then
                    oForm.Items("MSTCOD").CLICK ct_Regular
                    Sbo_Application.ActivateMenuItem ("7425")
                    BubbleEvent = False
                Else
                    If oMat1.RowCount > 0 Then
                        oMat1.Columns("Col1").Cells(oMat1.VisualRowCount).CLICK ct_Regular
                        BubbleEvent = False
                    End If
                End If
             ElseIf pval.BeforeAction = True And pval.ColUID = "Col1" And pval.CharPressed = 9 Then
                If Trim$(oMat1.Columns(pval.ColUID).Cells(pval.Row).Specific.String) = "" Then
                    Call Sbo_Application.StatusBar.SetText("기부연월은 필수입니다. 입력하세요.", bmt_Short, smt_Error)
                    BubbleEvent = False
                End If
             ElseIf pval.BeforeAction = True And pval.ColUID = "Col3" And pval.CharPressed = 9 Then
                If Trim$(oMat1.Columns(pval.ColUID).Cells(pval.Row).Specific.String) = "" Then
                    Call Sbo_Application.StatusBar.SetText("기부처상호는 필수입니다. 입력하세요.", bmt_Short, smt_Error)
                    BubbleEvent = False
                End If
             ElseIf pval.BeforeAction = True And pval.ColUID = "Col4" And pval.CharPressed = 9 Then
                If Trim$(oMat1.Columns(pval.ColUID).Cells(pval.Row).Specific.String) = "" Then
                    Call Sbo_Application.StatusBar.SetText("기부처의 사업자(주민)번호는 필수입니다. 입력하세요.", bmt_Short, smt_Error)
                    BubbleEvent = False
                Else
                    '/ 사업자번호 체크
                    If Len(oMat1.Columns(pval.ColUID).Cells(pval.Row).Specific.String) <= 12 Then
                        If MDC_SetMod.TaxNoCheck(oMat1.Columns(pval.ColUID).Cells(pval.Row).Specific.String) = False Then
                            Call Sbo_Application.StatusBar.SetText("사업자번호가 틀립니다. 확인하세요.", bmt_Short, smt_Error)
                            BubbleEvent = False
                        End If
                    End If
                End If
             End If
'et_GOT_FOCUS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_GOT_FOCUS
            If Last_Item = "Mat1" Then
                If pval.Row > 0 Then
                    Last_Item = pval.ItemUid
                    Col_Last_Row = pval.Row
                    Col_Last_Uid = pval.ColUID
                End If
            Else
                Last_Item = pval.ItemUid
                Col_Last_Row = 0
                Col_Last_Uid = ""
            End If
'et_FORM_UNLOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_FORM_UNLOAD
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            '컬렉션에서 삭제및 모든 메모리 제거
            'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
            If pval.BeforeAction = False Then
               RemoveForms oFormUniqueID
                Set oForm = Nothing
                Set oDS_ZPY505H = Nothing
                Set oDS_ZPY505L = Nothing
                Set oMat1 = Nothing
            End If
'et_MATRIX_LOAD''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case et_MATRIX_LOAD:
            If pval.BeforeAction = False Then
                FormItemEnabled
                Call Matrix_AddRow(oMat1.VisualRowCount)
            End If
            
    End Select
    
    Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Raise_FormItemEvent_Error:
    Sbo_Application.StatusBar.SetText "Raise_FormItemEvent_Error:", bmt_Short, smt_Error
End Sub
'*******************************************************************
'// oPaneLevel ==> 0:All / 1:oForm.PaneLevel=1 / 2:oForm.PaneLevel=2
'*******************************************************************
Private Sub Matrix_AddRow(ByVal oRow As Long, Optional Insert_YN As Boolean)
        If Insert_YN = False Then
            oDS_ZPY505L.InsertRecord (oRow)
        End If
        oDS_ZPY505L.Offset = oRow
        oDS_ZPY505L.setValue "U_LineNum", oRow, oRow + 1
        oDS_ZPY505L.setValue "U_GBUYMM", oRow, ""
        oDS_ZPY505L.setValue "U_GBUCOD", oRow, ""
        oDS_ZPY505L.setValue "U_GBUNAM", oRow, ""
        oDS_ZPY505L.setValue "U_GBUNBR", oRow, ""
        oDS_ZPY505L.setValue "U_GBUSEQ", oRow, ""
        oDS_ZPY505L.setValue "U_GBUCNT", oRow, ""
        oDS_ZPY505L.setValue "U_GBUAMT", oRow, ""
        oDS_ZPY505L.setValue "U_GWANGE", oRow, ""
        oDS_ZPY505L.setValue "U_FAMNAM", oRow, ""
        oDS_ZPY505L.setValue "U_INTGBN", oRow, ""
        oDS_ZPY505L.setValue "U_PERNBR", oRow, ""
        oDS_ZPY505L.setValue "U_BEFAMT", oRow, ""
        oDS_ZPY505L.setValue "U_CURAMT", oRow, ""
        oDS_ZPY505L.setValue "U_CHAAMT", oRow, ""
        oMat1.LoadFromDataSource
End Sub
'*******************************************************************
'// MenuEventHander
'*******************************************************************
Public Sub Raise_FormMenuEvent(ByRef FormUID As String, ByRef pval As SAPbouiCOM.IMenuEvent, ByRef BubbleEvent As Boolean)
    Dim i&
    
    If pval.BeforeAction = True Then
        Select Case pval.MenuUID
            Case "1283" '/ 제거
                If Trim$(oDS_ZPY505H.GetValue("U_ENDCHK", 0)) = "Y" Then
                    Sbo_Application.StatusBar.SetText "잠금 자료입니다. 확인하여 주십시오.", bmt_Short, smt_Error
                    BubbleEvent = False
                    Exit Sub
                Else
                    If Sbo_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") = 2 Then
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If
            Case "1288", "1289", "1290", "1291"
                Call AuthorityCheck(oForm, "CLTCOD", "@ZPY505H", "DocNum")      '//접속자 권한에 따른 사업장 보기
            
            Case Else
                Exit Sub
        End Select
    Else
        
        Select Case pval.MenuUID
            Case "1287" '/ 복제
               ' oForm.Items("Btn1").Visible = True
            Case "1283" '/ 제거
                 FormItemEnabled
            Case "1281", "1282":
                FormItemEnabled
                If pval.MenuUID = "1282" Then
                    FormClear
                    Call Matrix_AddRow(0, True)
                    oForm.Items("JSNYER").CLICK ct_Regular
                End If
            Case "1288" To "1291":
            Case "1293"
                If oMat1.RowCount <> oMat1.VisualRowCount Then
                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                    '//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
                    '//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
                    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                    For i = 0 To oMat1.VisualRowCount - 1
                        oMat1.Columns("Col0").Cells(i + 1).Specific.Value = i + 1
                    Next i
                    
                    oMat1.FlushToDataSource
                    oDS_ZPY505L.RemoveRecord oDS_ZPY505L.Size - 1   '// Mat1에 마지막라인(빈라인) 삭제
                    oMat1.Clear
                    oMat1.LoadFromDataSource
                End If
                Call FlushToItemValue("Col5", 1)
        End Select
    End If
End Sub

Public Sub Raise_FormDataEvent(ByRef FormUID As String, ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
    Dim i       As Long
    Dim sQry    As String
    Dim oCombo  As SAPbouiCOM.ComboBox
    
    Dim oRecordSet  As SAPbobsCOM.Recordset
    
    
    On Error GoTo Raise_FormDataEvent_Error
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    If (BusinessObjectInfo.BeforeAction = False) Then
        Select Case BusinessObjectInfo.EventType
            Case et_FORM_DATA_LOAD:     '//33
            Case et_FORM_DATA_ADD:      '//34
            Case et_FORM_DATA_UPDATE:   '//35
            Case et_FORM_DATA_DELETE:   '//36
        End Select

    End If
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Exit Sub
    
Raise_FormDataEvent_Error:
    Set oCombo = Nothing
    Set oRecordSet = Nothing
    Sbo_Application.SetStatusBarMessage "Raise_FormDataEvent_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True

End Sub

'*******************************************************************
' .srf 파일로부터 폼을 로드한다.
'*******************************************************************
Public Sub LoadForm(Optional JSNYER$, Optional MSTCOD$, Optional CLTCOD$)
'Public Sub LoadForm(Optional ByVal oFromDocEntry01 As String)
On Error GoTo LoadForm_Error
    Dim i           As Long
    Dim oXmlDoc             As New MSXML2.DOMDocument

    
    oXmlDoc.Load (MDC_Globals.SP_Path & "\" & SP_Screen & "\ZPY505.srf")
    oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = _
        oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue & "_" & (GetTotalFormsCount)
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '//여러개의 메트릭스가 틀경우에 층계모양처럼 로드 되도록 만든 모양
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetTotalFormsCount * 10)
    oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = _
            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetTotalFormsCount * 10)

    Sbo_Application.LoadBatchActions oXmlDoc.xml
    
    oFormUniqueID = "ZPY505_" & GetTotalFormsCount
    
    '폼 할당
    Set oForm = Sbo_Application.Forms.Item(oFormUniqueID)
    
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    '컬렉션에 폼을 담는다   **컬렉션이란 개체를 담아 놓는 배열로서 여기서는 활성화되어져 있는 폼을 담고 있다
    'ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
    AddForms Me, oFormUniqueID, "ZPY505"
    oForm.SupportedModes = -1
    oForm.Mode = fm_FIND_MODE
    
    '//////////////////////////////////////////////////////////////////////////////
    '***************************************************************
    '화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
    oForm.DataBrowser.BrowseBy = "DocNum"
    '***************************************************************
    '//////////////////////////////////////////////////////////////////////////////
    oForm.Freeze True
    CreateItems
    
    oForm.EnableMenu ("1293"), True  '/ 행삭제
    oForm.EnableMenu ("1283"), True  '/ 제거
    oForm.EnableMenu ("1284"), False  '/ 취소
    
    oForm.Freeze False
    oForm.Update
    'oForm.Visible = True
    
    If JSNYER <> "" Then
        Call ShowSource(JSNYER, MSTCOD, CLTCOD)
    End If
    
    Set oXmlDoc = Nothing
    Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
LoadForm_Error:
    Set oXmlDoc = Nothing
    Sbo_Application.StatusBar.SetText "Form_Load Error:" & Err.Description, bmt_Short, smt_Error
    If (oForm Is Nothing) = False Then
        oForm.Freeze False
        Set oForm = Nothing
    End If
End Sub

Private Sub BeforeBalance()
On Error GoTo Error_Message
    Dim JSNYER      As String
    Dim MSTCOD      As String
    Dim oRecordSet  As SAPbobsCOM.Recordset
    Dim sQry        As String
    Dim ErrNum      As Long
    Dim iRow        As Long
    Dim DupChk      As Boolean

    JSNYER = Trim$(oDS_ZPY505H.GetValue("U_JSNYER", 0))
    MSTCOD = Trim$(oDS_ZPY505H.GetValue("U_MSTCOD", 0))
    
    If JSNYER = "" Or MSTCOD = "" Then
        ErrNum = 1
        GoTo Error_Message
    End If
    
    oMat1.FlushToDataSource
    DupChk = False
    For iRow = 0 To oDS_ZPY505L.Size - 1
        If Val(oDS_ZPY505L.GetValue("U_BEFDOC", iRow)) <> 0 Then
            DupChk = True
            Exit For
        End If
    Next iRow
    If DupChk = True Then
        ErrNum = 3
        GoTo Error_Message
    End If
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    sQry = "EXEC ZPY505_1 '" & JSNYER & "', '" & MSTCOD & "'"
    oRecordSet.DoQuery sQry
    
    If oRecordSet.RecordCount = 0 Then
        ErrNum = 2
        GoTo Error_Message
    End If
    
    iRow = oDS_ZPY505L.Size - 1
    Do Until oRecordSet.EOF
        If iRow = oDS_ZPY505L.Size Then
            oDS_ZPY505L.InsertRecord (iRow)
        End If
        oDS_ZPY505L.Offset = iRow
        oDS_ZPY505L.setValue "U_LINENUM", iRow, iRow + 1
        oDS_ZPY505L.setValue "U_GBUYMM", iRow, oRecordSet.Fields("U_GBUYMM").Value
        oDS_ZPY505L.setValue "U_GBUCOD", iRow, oRecordSet.Fields("U_GBUCOD").Value
        oDS_ZPY505L.setValue "U_GBUNAM", iRow, oRecordSet.Fields("U_GBUNAM").Value
        oDS_ZPY505L.setValue "U_GBUCNT", iRow, oRecordSet.Fields("U_GBUCNT").Value
        oDS_ZPY505L.setValue "U_GBUAMT", iRow, oRecordSet.Fields("U_GBUAMT").Value
        oDS_ZPY505L.setValue "U_FAMNAM", iRow, oRecordSet.Fields("U_FAMNAM").Value
        oDS_ZPY505L.setValue "U_GWANGE", iRow, oRecordSet.Fields("U_GWANGE").Value
        oDS_ZPY505L.setValue "U_INTGBN", iRow, oRecordSet.Fields("U_INTGBN").Value
        oDS_ZPY505L.setValue "U_PERNBR", iRow, oRecordSet.Fields("U_PERNBR").Value
        oDS_ZPY505L.setValue "U_BEFAMT", iRow, oRecordSet.Fields("U_CHAAMT").Value
        
        oRecordSet.MoveNext
        iRow = iRow + 1
    Loop
    Call Matrix_AddRow(iRow)
    
    Sbo_Application.StatusBar.SetText "이전년도 기부금 이월금액 가져오기를 완료하였습니다.", bmt_Short, smt_Success
    
    Set oRecordSet = Nothing
    
    Exit Sub
Error_Message:

    Set oRecordSet = Nothing
    If ErrNum = 1 Then
        Sbo_Application.StatusBar.SetText "정산년도와 사원번호는 필수입니다. 입력하여 주십시오.", bmt_Short, smt_Error
    ElseIf ErrNum = 2 Then
        Sbo_Application.StatusBar.SetText "이전년도에서 이월된 기부금 내역이 없습니다.", bmt_Short, smt_Error
    ElseIf ErrNum = 3 Then
        Sbo_Application.StatusBar.SetText "이미 이월금액 가져오기를 실행한 상태입니다.", bmt_Short, smt_Error
    Else
        Sbo_Application.StatusBar.SetText "BeforeBalance Error : " & Err.Description, bmt_Short, smt_Error
    End If
    
End Sub

Private Sub ShowSource(JSNYER$, MSTCOD$, CLTCOD$)
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim DocEntry$
    Dim oMast           As ZPAY_g_EmpID
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    sQry = "SELECT DocNum FROM [@ZPY505H]"
    sQry = sQry & "   WHERE U_JSNYER = N'" & JSNYER & "'"
    sQry = sQry & "   AND   U_MSTCOD = N'" & MSTCOD & "'"
    sQry = sQry & "   AND   U_CLTCOD = N'" & CLTCOD & "'"
    oRecordSet.DoQuery sQry
    If oRecordSet.RecordCount > 0 Then
        Do Until oRecordSet.EOF
            DocEntry = oRecordSet.Fields(0).Value
            oRecordSet.MoveNext
        Loop
        oForm.Items("JSNYER").Specific.Value = JSNYER
        oForm.Items("MSTCOD").Specific.String = MSTCOD
        Call oForm.Items("CLTCOD").Specific.Select(CLTCOD, psk_ByValue)
        oForm.Items("DocNum").Specific.Value = DocEntry
        
        oForm.Items("DocNum").Update
        oMat1.LoadFromDataSource
        oForm.Update
        oForm.Items("1").CLICK ct_Regular
        
    Else
        oForm.Mode = fm_OK_MODE
        Call Sbo_Application.ActivateMenuItem("1282")
        
        Call oDS_ZPY505H.setValue("U_JSNYER", 0, JSNYER)
        Call oDS_ZPY505H.setValue("U_MSTCOD", 0, MSTCOD)
        Call oDS_ZPY505H.setValue("U_CLTCOD", 0, CLTCOD)
        oMast = MDC_SetMod.Get_EmpID_InFo(MSTCOD)
        Call oDS_ZPY505H.setValue("U_MSTNAM", 0, oMast.MSTNAM)
        Call oDS_ZPY505H.setValue("U_EmpID", 0, oMast.EmpID)
        
        oForm.Update
        
        Sbo_Application.SendKeys "{TAB}"
    End If
    
    Set oRecordSet = Nothing
    
End Sub
'*******************************************************************
'
'*******************************************************************
Private Sub CreateItems()
On Error GoTo Error_Message
    Dim oCombo          As SAPbouiCOM.ComboBox
    Dim oRecordSet      As SAPbobsCOM.Recordset
    Dim sQry            As String
    Dim oColumn         As SAPbouiCOM.Column
    Dim oCheck          As SAPbouiCOM.CheckBox
    
    Set oRecordSet = oCompany.GetBusinessObject(BoRecordset)
    
    '//디비데이터 소스 개체 할당
    Set oDS_ZPY505H = oForm.DataSources.DBDataSources("@ZPY505H")
    Set oDS_ZPY505L = oForm.DataSources.DBDataSources("@ZPY505L")
    
    Set oMat1 = oForm.Items("Mat1").Specific
    
    '// 사업장
    Set oCombo = oForm.Items("CLTCOD").Specific
    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
    oRecordSet.DoQuery sQry
    Do Until oRecordSet.EOF
        oCombo.ValidValues.Add Trim$(oRecordSet.Fields(0).Value), Trim$(oRecordSet.Fields(1).Value)
        oRecordSet.MoveNext
    Loop
    
    '// 기부코드
    Set oColumn = oMat1.Columns("Col2")
    oColumn.ValidValues.Add "10", "법정기부금"
    oColumn.ValidValues.Add "20", "정치자금"
    'oColumn.ValidValues.Add "21", "문화예술진흥기금" '2011년
    oColumn.ValidValues.Add "30", "특례기부금"
    oColumn.ValidValues.Add "31", "공익법인기부금신탁"
    oColumn.ValidValues.Add "40", "지정기부금(종교단체외)"
    oColumn.ValidValues.Add "41", "지정기부금(종교단체)"
    oColumn.ValidValues.Add "42", "우리사주조합 기부금"
    oColumn.ValidValues.Add "50", "공제제외기부금"
    
    '// 관계코드
    Set oColumn = oMat1.Columns("Col9")
    oColumn.ValidValues.Add "1", "거주자(본인)"
    oColumn.ValidValues.Add "2", "배우자"
    oColumn.ValidValues.Add "3", "직계비속"
    oColumn.ValidValues.Add "4", "직계존속"
    oColumn.ValidValues.Add "5", "형제자매"
    oColumn.ValidValues.Add "6", "그 외"
 
    '// 내외국인
    Set oColumn = oMat1.Columns("Col10")
    oColumn.ValidValues.Add "1", "내국인"
    oColumn.ValidValues.Add "9", "외국인"
    
    Set oColumn = oMat1.Columns("Col19")
    oColumn.ValOff = "N":   oColumn.ValOn = "Y"
 
    '// 영수증일련번호(2008년 제외)
    oMat1.Columns("Col7").Visible = False
    
  '/ Check 버튼
    Set oCheck = oForm.Items("ENDCHK").Specific
    oCheck.ValOff = "N": oCheck.ValOn = "Y"
    
    Set oCheck = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oRecordSet = Nothing
   Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Error_Message:
    Set oCheck = Nothing
    Set oCombo = Nothing
    Set oColumn = Nothing
    Set oRecordSet = Nothing
     Sbo_Application.StatusBar.SetText "CreateItems Error:" & Err.Description, bmt_Short, smt_Error
End Sub













