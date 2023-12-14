Attribute VB_Name = "Module1"
Sub List_to_Analysis()
    Dim ws As Worksheet
    Dim Row_max As Long
    Dim i As Long
    Dim fullString As String
    Dim PartString_Type As String
    Dim headers As Variant
    Dim ii As Integer
    
    Set ws = Worksheets("List_Table")
    Set ws_M42 = Worksheets("Weight_Analysis")
    Set ws_Weight_Analysis = Worksheets("Weight_Analysis")
    ' �M���Ҧ����e
    ws_M42.Cells.ClearContents
    
    ' �˴��O�_����T�Y�L�h�l�[

    ' �w�q�C���D�M�������C��
    headers = Array(Array("A", "�ޤ伵����"), Array("B", "����"), Array("C", "�~�W"), Array("D", "�ؤo/�p��"), Array("E", "����"), Array("F", "�e��"), Array("G", "����"), Array("H", "�ƶq"), Array("I", "�C�̭�"), Array("J", "�歫"), Array("K", "���q�p�p"), Array("L", "���"), Array("M", "�ռ�"), Array("N", "���פp�p"), Array("O", "�ƶq�p�p"), Array("P", "���q�X�p"), Array("Q", "�ݩ�"))

    ' �M���Ʋըó]�m�C���D
       With ws_Weight_Analysis
    For ii = LBound(headers) To UBound(headers)
        If .Cells(1, headers(ii)(0)).value <> headers(ii)(1) Then
            .Cells(1, headers(ii)(0)).value = headers(ii)(1)
        End If
    Next ii
        End With
    
    ' �ק�F��M�̫�@�C����k
    Row_max = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To Row_max
        fullString = ws.Cells(i, "A").value
        PartString_Type = GetFirstPartOfString(fullString)
        Last_row_main_Title = GetNextRowInColumnB()
        ws_M42.Cells(Last_row_main_Title, "A") = fullString
        ' �o�̥i�H�ھ� PartString_Type �i�椣�P�ާ@
Select Case PartString_Type
    Case "01"
        Type_01 fullString
    ' Case "02"
        ' Type_02 fullString
    ' Case "03"
        ' Type_03 fullString
    ' Case "04"
        ' Type_04 fullString
     Case "05"
         Type_05 fullString
    ' Case "06"
        ' Type_06 fullString
    ' Case "07"
        ' Type_07 fullString
    ' Case "08"
        ' Type_08 fullString
     Case "09"
         Type_09 fullString
    ' Case "10"
        ' Type_10 fullString
    ' Case "11"
        ' Type_11 fullString
    ' Case "12"
        ' Type_12 fullString
    ' Case "13"
        ' Type_13 fullString
     Case "14"
         Type_14 fullString
    ' Case "15"
        ' Type_15 fullString
    Case "16"
        Type_16 fullString
    ' Case "17"
        ' Type_17 fullString
    ' Case "18"
        ' Type_18 fullString
    ' Case "19"
        ' Type_19 fullString
     Case "20"
         Type_20 fullString
     Case "21"
         Type_21 fullString
     Case "22"
         Type_22 fullString
     Case "23"
         Type_23 fullString
     Case "24"
         Type_24 fullString
     Case "25"
         Type_25 fullString
    ' Case "26"
        ' Type_26 fullString
    ' Case "27"
        ' Type_27 fullString
    ' Case "28"
        ' Type_28 fullString
    ' Case "29"
        ' Type_29 fullString
    Case "108"
        Type_108 fullString
    Case Else
        Exit Sub
End Select
    Next i
End Sub


Function GetFirstPartOfString(fullString As String) As String
    Dim splitString As Variant
    Dim firstPart As String
    
    splitString = Split(fullString, "-") ' �ϥ�"-"�@�����j�ŨӤ��Φr�Ŧ�
    
    If UBound(splitString) >= 1 Then ' �T�O�����������j��
        firstPart = splitString(0) ' ������Ϋ�Ʋժ��Ĥ@�Ӥ����A�Y�Ĥ@��"-"���e����
    Else
        firstPart = "N/A" ' �p�G�S�����������j�šA�]�m�@�ӿ��~�������q�{��
    End If

    GetFirstPartOfString = firstPart
End Function


Function GetSecondPartOfString(fullString As String) As String
    Dim splitString As Variant
    Dim secondPart As String
    
    splitString = Split(fullString, "-") ' �ϥ�"-"�@�����j�ŨӤ��Φr�Ŧ�
    
    If UBound(splitString) >= 1 Then ' �T�O�����������j��
        secondPart = splitString(1) ' ������Ϋ�Ʋժ��ĤG�Ӥ����A�Y�Ĥ@�өM�ĤG��"-"��������
    Else
        secondPart = "N/A" ' �p�G�S�����������j�šA�]�m�@�ӿ��~�������q�{��
    End If

    GetSecondPartOfString = secondPart
End Function

Function GetThirdPartOfString(fullString As String) As String
    Dim splitString As Variant
    splitString = Split(fullString, "-") ' �ϥ� "-" �Ӥ��Φr�Ŧ�

    If UBound(splitString) >= 2 Then ' �T�O�����������j��
        GetThirdPartOfString = splitString(2) ' �ĤT����
    Else
        GetThirdPartOfString = "N/A" ' �p�G�S�����������j�šA�]�m�@�ӿ��~�������q�{��
    End If
End Function
Function GetFourthPartOfString(fullString As String) As String
    Dim splitString As Variant
    splitString = Split(fullString, "-") ' �ϥ� "-" �Ӥ��Φr�Ŧ�

    If UBound(splitString) >= 3 Then ' �T�O�����������j��
        GetFourthPartOfString = splitString(3) ' �ĥ|����
    Else
        GetFourthPartOfString = "N/A" ' �p�G�S�����������j�šA�]�m�@�ӿ��~�������q�{��
    End If
End Function

Function GetNextRowInColumnB() As Long
    Dim ws As Worksheet
    Dim lastRow As Long

    ' �]�w�� "Weight_Analysis" �u�@���ޥ�
    Set ws = Worksheets("Weight_Analysis")

    ' ���� B �C���̫�@��
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' ��^�U�@�檺�渹
    GetNextRowInColumnB = lastRow + 1
End Function


Function CalculatePipeWeight(Pipe_Dn_inch As Double, Pipe_Weight_thickness_mm As Double) As Double
    Dim pi As Double
    pi = 4 * Atn(1)
    ' �p�⤽��
    CalculatePipeWeight = Round(((Pipe_Dn_inch - Pipe_Weight_thickness_mm) * pi / 1000 * 1 * Pipe_Weight_thickness_mm * 7.85), 2)
End Function
Function GetLookupValue(value As Variant) As Variant
    ' �N���ഫ���r�Ŧ�
    Dim strValue As String
    strValue = CStr(value)

    ' �ˬd�O�_�t���p���I
    If InStr(1, strValue, ".") > 0 Then
        ' �p�G���p���I�A�O�����r�Ŧ�
        If InStr(1, strValue, "'") = 0 Then
        
        GetLookupValue = "'" & strValue
        
       Else
        
        GetLookupValue = strValue
        
        End If
    Else
        ' �S���p���I�A���եh���D�Ʀr�r�ū��ഫ�����
        Dim numericValue As String
        numericValue = ""
        Dim i As Integer
        For i = 1 To Len(strValue)
            If IsNumeric(Mid(strValue, i, 1)) Then
                numericValue = numericValue & Mid(strValue, i, 1)
            End If
        Next i

        If Len(numericValue) > 0 Then
            GetLookupValue = CInt(numericValue)
        Else
            GetLookupValue = 0 ' �γ]�m�@�ӦX�z���q�{��
        End If
    End If
End Function


Function CalculatePipeDetails(PipeSize As Variant, PipeThickness As Variant) As Collection
    Dim PipeDetails As New Collection
    Dim PipeDiameterInch As Double
    Dim PipeThicknessColumn As Long
    Dim PipeWeightPerMeter As Double
    Dim LookupValue As Variant
    
    ' �]�m�u�@��ޥ�
    Set ws_Pipe_Table = Worksheets("Pipe_Table")
    
    ' ����d���
    LookupValue = GetLookupValue(PipeSize)

    ' ����޹D���|�]�^�o�^
    PipeDiameterInch = ws_Pipe_Table.Application.WorksheetFunction.VLookup(LookupValue, ws_Pipe_Table.Range("B:R"), 2, False)

    ' ����޹D�p�שҦb�C
    PipeThicknessColumn = ws_Pipe_Table.Application.WorksheetFunction.Match(PipeThickness, ws_Pipe_Table.Range("B3:R3"), 0)

    ' ����C�̭��q
    PipeWeightPerMeter = ws_Pipe_Table.Application.WorksheetFunction.VLookup(LookupValue, ws_Pipe_Table.Range("B:R"), PipeThicknessColumn, False)

    ' �p�⭫�q
    Dim TotalWeight As Double
    TotalWeight = CalculatePipeWeight(CDbl(PipeDiameterInch), CDbl(PipeWeightPerMeter))

    ' �K�[�춰�X
    PipeDetails.Add PipeDiameterInch, "DiameterInch"
    PipeDetails.Add PipeWeightPerMeter, "WeightPerMeter"
    PipeDetails.Add TotalWeight, "TotalWeight"

    ' ��^���X
    Set CalculatePipeDetails = PipeDetails
End Function


Function ExtractParts(fourthString As String) As Variant
    
'����ƭt�d�װťX �@�Ӧr�� �t��"()"�� �ä��Φ�0�Ϊ�1
'�Ҧp : A(S) �h needvalue(0) = "A" needValue(1) = (S)
'needValue = ExtractParts("A(S)")
    Dim openParenPos As Integer
    openParenPos = InStr(fourthString, "(")
    
    If openParenPos > 0 Then
        Dim partBeforeParen As String
        Dim partWithParen As String

        partBeforeParen = Left(fourthString, openParenPos - 1)
        partWithParen = Mid(fourthString, openParenPos)

        ExtractParts = Array(partBeforeParen, partWithParen)
    Else
        ExtractParts = Array(fourthString, "")
    End If
End Function
Function CalculateAngleDetail(Angle_A As Variant, Angle_B As Variant, Thickness As Variant) As Double
    ' Const density As Double = 7.85 ' ���K���K�סA���: kg/dm3
    ' Dim singleWeight As Double
    ' Dim A As Double, B As Double, t As Double

    ' ' ���ձN�Ѽ��ഫ�� Double ����
    ' On Error Resume Next
    ' A = CDbl(Angle_A)
    ' B = CDbl(Angle_B)
    ' t = CDbl(Thickness)
    ' If Err.Number <> 0 Then
        ' CalculateAngleDetail = 0 ' �p�G�ഫ���ѡA��^ 0
        ' Exit Function
    ' End If
    ' On Error GoTo 0

    ' ' �T�O t ���ȾA�X�i��p��
    ' If t <= 0 Then
        ' CalculateAngleDetail = 0 ' �p�G t �p��ε��� 0�A��^ 0
        ' Exit Function
    ' End If

    ' ' �p��歫
    ' singleWeight = (((A * t) + (B * t) - (t * t)) * density) / 1000
    
    ' ' ��^�p�⵲�G
    ' CalculateAngleDetail = singleWeight
End Function

Sub AddPlateEntry(PlateType As String, PipeSize As Variant)
    Dim Plate_Size As Double
    Dim Plate_Thickness As Double
    Dim Weight_calculator As Double
    Dim Plate_Name As String
    Dim RequireDrilling As Boolean
    Dim ws As Worksheet
    Dim i As Long

    ' �]�w��S�w�u�@���ޥ�
    Set ws = Worksheets("Weight_Analysis")
    Set ws_M42 = Worksheets("M_42_Table")

    ' ���C B ���̫�@��A�ì��s�ƾڷǳƤU�@��
    i = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row + 1

    ' �ھڪO�l�����M�w�O�_�ݭn�p��
    Select Case PlateType
        Case "d", "b", "c"
            RequireDrilling = True
        Case Else
            RequireDrilling = False
    End Select
    ' �ھڪO�l�����M�w�ݩ�(BXB EXE GXG CXC)
        Select Case PlateType
            Case "a"
                col_type = 3
            Case "b"
                col_type = 3
            Case "c"
                col_type = 3
            Case "d"
                col_type = 6
            Case "e"
                col_type = 8
        End Select
            
            
            
    ' �T�wPlate���ؤo�M�p��
    If InStr(1, PipeSize, "*") > 0 Then
        ' �YPipeSize���S�w�榡���r�Ŧ�
        col_type = col_type - 1
        Plate_Size = Application.WorksheetFunction.VLookup(PipeSize, ws_M42.Range("B:L"), col_type, False)
        Plate_Thickness = Application.WorksheetFunction.VLookup(PipeSize, ws_M42.Range("B:L"), 10, False)
    Else
        ' �YPipeSize���Ʀr
        PipeSize = GetLookupValue(PipeSize)
        Plate_Size = Application.WorksheetFunction.VLookup(PipeSize, ws_M42.Range("A:L"), col_type, False)
        Plate_Thickness = Application.WorksheetFunction.VLookup(PipeSize, ws_M42.Range("A:L"), 11, False)
    End If
    Weight_calculator = Plate_Size / 1000 * Plate_Size / 1000 * Plate_Thickness * 7.85

    ' �T�wPlate�W��
    Plate_Name = "Plate_" & PlateType & IIf(RequireDrilling, "_���p��", "_�����p��")

    ' ��R�ƾ�
    With ws
        .Cells(i, "B").value = .Cells(i - 1, "B").value + 1
        .Cells(i, "C").value = Plate_Name
        .Cells(i, "D").value = Plate_Thickness
        .Cells(i, "E").value = Plate_Size
        .Cells(i, "F").value = Plate_Size
        .Cells(i, "G").value = "A36/SS400"
        .Cells(i, "H").value = 1
        .Cells(i, "J").value = Weight_calculator
        .Cells(i, "K").value = Weight_calculator
        .Cells(i, "L").value = "PC"
        .Cells(i, "M").value = 1
        .Cells(i, "O").value = 1
        .Cells(i, "P").value = Weight_calculator
        .Cells(i, "Q").value = "���O��"
    End With
End Sub

Sub AddBoltEntry(PipeSize As Variant, Quantity As Integer)
    Dim ws As Worksheet
    Dim i As Long
    Dim BoltSize As String
    
    ' �]�w�� "Weight_Analysis" �u�@���ޥ�
    Set ws = Worksheets("Weight_Analysis")
    Set ws_M42 = Worksheets("M_42_Table")
    ' ���C B ���U�@�Ӫťզ�
    i = GetNextRowInColumnB()

    If InStr(1, PipeSize, "*") > 0 Then
        ' �YPipeSize���S�w�榡���r�Ŧ�
        BoltSize = Application.WorksheetFunction.VLookup(PipeSize, ws_M42.Range("B:L"), 9, False)
    Else
        ' �YPipeSize���Ʀr
        PipeSize = GetLookupValue(PipeSize)
        BoltSize = Application.WorksheetFunction.VLookup(PipeSize, ws_M42.Range("A:L"), 10, False)
    End If


    ' ��R�ƾ�
    With ws
        .Cells(i, "B").value = .Cells(i - 1, "B").value + 1
        .Cells(i, "C").value = "EXP.BOLT"
        .Cells(i, "D").value = "'" & BoltSize & """"
        .Cells(i, "G").value = "SUS304"
        .Cells(i, "H").value = Quantity
        .Cells(i, "J").value = 1 ' ���]�C�����ꪺ��ӭ��q�O1�]�i�H�ھڹ�ڱ��p�վ�^
        .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value
        .Cells(i, "L").value = "SET"
        .Cells(i, "M").value = 1
        .Cells(i, "O").value = .Cells(i, "M").value * .Cells(i, "H").value
        .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
        .Cells(i, "Q").value = "������"
    End With
End Sub


Sub AddSteelSectionEntry(SectionType As String, Section_Dim As String, Total_Length As Double)
    Dim ws As Worksheet
    Dim i As Long
    Dim SectionWeight As Double


    ' �]�w��U���ؤu�@���ޥ�
    Set ws = Worksheets("Weight_Analysis")
    Set ws_HBeam = Worksheets("For_HBeam_Weight_Table")
    Set ws_Channel = Worksheets("For_Channel_Weight_Table")
    Set ws_Angle = Worksheets("For_Angle_Weight_Table")

    ' �ѷӭ��q
    Select Case SectionType
        Case "Angle"
            SectionWeight = Application.WorksheetFunction.VLookup(Section_Dim, ws_Angle.Range("C:G"), 5, False)
        Case "Channel"
            SectionWeight = Application.WorksheetFunction.VLookup(Section_Dim, ws_Channel.Range("D:H"), 5, False)
        Case "H beam"
            SectionWeight = Application.WorksheetFunction.VLookup(Section_Dim, ws_HBeam.Range("E:H"), 4, False)
        Case Else
            SectionWeight = 0 ' IF NO THEN 0
    End Select

    ' ���� B �C���U�@�Ӫťզ�
    i = GetNextRowInColumnB()
     With ws
    ' �p�G
    If .Cells(i, "A").value <> "" Then
    First_Value_Checking = 1
    Else
    First_Value_Checking = .Cells(i - 1, "B").value + 1
    End If
    ' ��R�ƾ�
   
        .Cells(i, "B").value = First_Value_Checking
        .Cells(i, "C").value = SectionType
        .Cells(i, "D").value = Section_Dim
        .Cells(i, "E").value = Total_Length
        .Cells(i, "G").value = "A36/SS400"
        .Cells(i, "H").value = 1
        .Cells(i, "I").value = SectionWeight
        .Cells(i, "J").value = .Cells(i, "E").value / 1000 * .Cells(i, "I").value
        .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value '���q�p�p
        .Cells(i, "L").value = "M"
        .Cells(i, "M").value = 1
        .Cells(i, "N").value = .Cells(i, "M").value * .Cells(i, "E").value / 1000 * .Cells(i, "H").value
        .Cells(i, "O").value = .Cells(i, "M").value * .Cells(i, "H").value
        .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
        .Cells(i, "Q").value = "������"
    End With
End Sub

Sub addPipeEntry(PipeSize As Variant, PipeThickness As Variant, Pipe_Length As Double)
    
    Set ws = Worksheets("Weight_Analysis")
    
    
End Sub
