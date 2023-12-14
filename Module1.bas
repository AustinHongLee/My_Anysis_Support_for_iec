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
    ' 清除所有內容
    ws_M42.Cells.ClearContents
    
    ' 檢測是否有資訊若無則追加

    ' 定義列標題和對應的列號
    headers = Array(Array("A", "管支撐型號"), Array("B", "項次"), Array("C", "品名"), Array("D", "尺寸/厚度"), Array("E", "長度"), Array("F", "寬度"), Array("G", "材質"), Array("H", "數量"), Array("I", "每米重"), Array("J", "單重"), Array("K", "重量小計"), Array("L", "單位"), Array("M", "組數"), Array("N", "長度小計"), Array("O", "數量小計"), Array("P", "重量合計"), Array("Q", "屬性"))

    ' 遍歷數組並設置列標題
       With ws_Weight_Analysis
    For ii = LBound(headers) To UBound(headers)
        If .Cells(1, headers(ii)(0)).value <> headers(ii)(1) Then
            .Cells(1, headers(ii)(0)).value = headers(ii)(1)
        End If
    Next ii
        End With
    
    ' 修改了找尋最後一列的方法
    Row_max = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To Row_max
        fullString = ws.Cells(i, "A").value
        PartString_Type = GetFirstPartOfString(fullString)
        Last_row_main_Title = GetNextRowInColumnB()
        ws_M42.Cells(Last_row_main_Title, "A") = fullString
        ' 這裡可以根據 PartString_Type 進行不同操作
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
    
    splitString = Split(fullString, "-") ' 使用"-"作為分隔符來分割字符串
    
    If UBound(splitString) >= 1 Then ' 確保有足夠的分隔符
        firstPart = splitString(0) ' 獲取分割後數組的第一個元素，即第一個"-"之前的值
    Else
        firstPart = "N/A" ' 如果沒有足夠的分隔符，設置一個錯誤消息或默認值
    End If

    GetFirstPartOfString = firstPart
End Function


Function GetSecondPartOfString(fullString As String) As String
    Dim splitString As Variant
    Dim secondPart As String
    
    splitString = Split(fullString, "-") ' 使用"-"作為分隔符來分割字符串
    
    If UBound(splitString) >= 1 Then ' 確保有足夠的分隔符
        secondPart = splitString(1) ' 獲取分割後數組的第二個元素，即第一個和第二個"-"之間的值
    Else
        secondPart = "N/A" ' 如果沒有足夠的分隔符，設置一個錯誤消息或默認值
    End If

    GetSecondPartOfString = secondPart
End Function

Function GetThirdPartOfString(fullString As String) As String
    Dim splitString As Variant
    splitString = Split(fullString, "-") ' 使用 "-" 來分割字符串

    If UBound(splitString) >= 2 Then ' 確保有足夠的分隔符
        GetThirdPartOfString = splitString(2) ' 第三部分
    Else
        GetThirdPartOfString = "N/A" ' 如果沒有足夠的分隔符，設置一個錯誤消息或默認值
    End If
End Function
Function GetFourthPartOfString(fullString As String) As String
    Dim splitString As Variant
    splitString = Split(fullString, "-") ' 使用 "-" 來分割字符串

    If UBound(splitString) >= 3 Then ' 確保有足夠的分隔符
        GetFourthPartOfString = splitString(3) ' 第四部分
    Else
        GetFourthPartOfString = "N/A" ' 如果沒有足夠的分隔符，設置一個錯誤消息或默認值
    End If
End Function

Function GetNextRowInColumnB() As Long
    Dim ws As Worksheet
    Dim lastRow As Long

    ' 設定對 "Weight_Analysis" 工作表的引用
    Set ws = Worksheets("Weight_Analysis")

    ' 找到第 B 列的最後一行
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' 返回下一行的行號
    GetNextRowInColumnB = lastRow + 1
End Function


Function CalculatePipeWeight(Pipe_Dn_inch As Double, Pipe_Weight_thickness_mm As Double) As Double
    Dim pi As Double
    pi = 4 * Atn(1)
    ' 計算公式
    CalculatePipeWeight = Round(((Pipe_Dn_inch - Pipe_Weight_thickness_mm) * pi / 1000 * 1 * Pipe_Weight_thickness_mm * 7.85), 2)
End Function
Function GetLookupValue(value As Variant) As Variant
    ' 將值轉換為字符串
    Dim strValue As String
    strValue = CStr(value)

    ' 檢查是否含有小數點
    If InStr(1, strValue, ".") > 0 Then
        ' 如果有小數點，保持為字符串
        If InStr(1, strValue, "'") = 0 Then
        
        GetLookupValue = "'" & strValue
        
       Else
        
        GetLookupValue = strValue
        
        End If
    Else
        ' 沒有小數點，嘗試去除非數字字符後轉換成整數
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
            GetLookupValue = 0 ' 或設置一個合理的默認值
        End If
    End If
End Function


Function CalculatePipeDetails(PipeSize As Variant, PipeThickness As Variant) As Collection
    Dim PipeDetails As New Collection
    Dim PipeDiameterInch As Double
    Dim PipeThicknessColumn As Long
    Dim PipeWeightPerMeter As Double
    Dim LookupValue As Variant
    
    ' 設置工作表引用
    Set ws_Pipe_Table = Worksheets("Pipe_Table")
    
    ' 獲取查找值
    LookupValue = GetLookupValue(PipeSize)

    ' 獲取管道直徑（英寸）
    PipeDiameterInch = ws_Pipe_Table.Application.WorksheetFunction.VLookup(LookupValue, ws_Pipe_Table.Range("B:R"), 2, False)

    ' 獲取管道厚度所在列
    PipeThicknessColumn = ws_Pipe_Table.Application.WorksheetFunction.Match(PipeThickness, ws_Pipe_Table.Range("B3:R3"), 0)

    ' 獲取每米重量
    PipeWeightPerMeter = ws_Pipe_Table.Application.WorksheetFunction.VLookup(LookupValue, ws_Pipe_Table.Range("B:R"), PipeThicknessColumn, False)

    ' 計算重量
    Dim TotalWeight As Double
    TotalWeight = CalculatePipeWeight(CDbl(PipeDiameterInch), CDbl(PipeWeightPerMeter))

    ' 添加到集合
    PipeDetails.Add PipeDiameterInch, "DiameterInch"
    PipeDetails.Add PipeWeightPerMeter, "WeightPerMeter"
    PipeDetails.Add TotalWeight, "TotalWeight"

    ' 返回集合
    Set CalculatePipeDetails = PipeDetails
End Function


Function ExtractParts(fourthString As String) As Variant
    
'此函數負責修剪出 一個字串 含有"()"的 並切割成0或者1
'例如 : A(S) 則 needvalue(0) = "A" needValue(1) = (S)
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
    ' Const density As Double = 7.85 ' 鋼鐵的密度，單位: kg/dm3
    ' Dim singleWeight As Double
    ' Dim A As Double, B As Double, t As Double

    ' ' 嘗試將參數轉換為 Double 類型
    ' On Error Resume Next
    ' A = CDbl(Angle_A)
    ' B = CDbl(Angle_B)
    ' t = CDbl(Thickness)
    ' If Err.Number <> 0 Then
        ' CalculateAngleDetail = 0 ' 如果轉換失敗，返回 0
        ' Exit Function
    ' End If
    ' On Error GoTo 0

    ' ' 確保 t 的值適合進行計算
    ' If t <= 0 Then
        ' CalculateAngleDetail = 0 ' 如果 t 小於或等於 0，返回 0
        ' Exit Function
    ' End If

    ' ' 計算單重
    ' singleWeight = (((A * t) + (B * t) - (t * t)) * density) / 1000
    
    ' ' 返回計算結果
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

    ' 設定對特定工作表的引用
    Set ws = Worksheets("Weight_Analysis")
    Set ws_M42 = Worksheets("M_42_Table")

    ' 找到列 B 的最後一行，並為新數據準備下一行
    i = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row + 1

    ' 根據板子類型決定是否需要鑽孔
    Select Case PlateType
        Case "d", "b", "c"
            RequireDrilling = True
        Case Else
            RequireDrilling = False
    End Select
    ' 根據板子類型決定屬性(BXB EXE GXG CXC)
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
            
            
            
    ' 確定Plate的尺寸和厚度
    If InStr(1, PipeSize, "*") > 0 Then
        ' 若PipeSize為特定格式的字符串
        col_type = col_type - 1
        Plate_Size = Application.WorksheetFunction.VLookup(PipeSize, ws_M42.Range("B:L"), col_type, False)
        Plate_Thickness = Application.WorksheetFunction.VLookup(PipeSize, ws_M42.Range("B:L"), 10, False)
    Else
        ' 若PipeSize為數字
        PipeSize = GetLookupValue(PipeSize)
        Plate_Size = Application.WorksheetFunction.VLookup(PipeSize, ws_M42.Range("A:L"), col_type, False)
        Plate_Thickness = Application.WorksheetFunction.VLookup(PipeSize, ws_M42.Range("A:L"), 11, False)
    End If
    Weight_calculator = Plate_Size / 1000 * Plate_Size / 1000 * Plate_Thickness * 7.85

    ' 確定Plate名稱
    Plate_Name = "Plate_" & PlateType & IIf(RequireDrilling, "_需鑽孔", "_不需鑽孔")

    ' 填充數據
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
        .Cells(i, "Q").value = "鋼板類"
    End With
End Sub

Sub AddBoltEntry(PipeSize As Variant, Quantity As Integer)
    Dim ws As Worksheet
    Dim i As Long
    Dim BoltSize As String
    
    ' 設定對 "Weight_Analysis" 工作表的引用
    Set ws = Worksheets("Weight_Analysis")
    Set ws_M42 = Worksheets("M_42_Table")
    ' 找到列 B 的下一個空白行
    i = GetNextRowInColumnB()

    If InStr(1, PipeSize, "*") > 0 Then
        ' 若PipeSize為特定格式的字符串
        BoltSize = Application.WorksheetFunction.VLookup(PipeSize, ws_M42.Range("B:L"), 9, False)
    Else
        ' 若PipeSize為數字
        PipeSize = GetLookupValue(PipeSize)
        BoltSize = Application.WorksheetFunction.VLookup(PipeSize, ws_M42.Range("A:L"), 10, False)
    End If


    ' 填充數據
    With ws
        .Cells(i, "B").value = .Cells(i - 1, "B").value + 1
        .Cells(i, "C").value = "EXP.BOLT"
        .Cells(i, "D").value = "'" & BoltSize & """"
        .Cells(i, "G").value = "SUS304"
        .Cells(i, "H").value = Quantity
        .Cells(i, "J").value = 1 ' 假設每個螺栓的單個重量是1（可以根據實際情況調整）
        .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value
        .Cells(i, "L").value = "SET"
        .Cells(i, "M").value = 1
        .Cells(i, "O").value = .Cells(i, "M").value * .Cells(i, "H").value
        .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
        .Cells(i, "Q").value = "螺絲類"
    End With
End Sub


Sub AddSteelSectionEntry(SectionType As String, Section_Dim As String, Total_Length As Double)
    Dim ws As Worksheet
    Dim i As Long
    Dim SectionWeight As Double


    ' 設定對各鋼種工作表的引用
    Set ws = Worksheets("Weight_Analysis")
    Set ws_HBeam = Worksheets("For_HBeam_Weight_Table")
    Set ws_Channel = Worksheets("For_Channel_Weight_Table")
    Set ws_Angle = Worksheets("For_Angle_Weight_Table")

    ' 參照重量
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

    ' 找到第 B 列的下一個空白行
    i = GetNextRowInColumnB()
     With ws
    ' 如果
    If .Cells(i, "A").value <> "" Then
    First_Value_Checking = 1
    Else
    First_Value_Checking = .Cells(i - 1, "B").value + 1
    End If
    ' 填充數據
   
        .Cells(i, "B").value = First_Value_Checking
        .Cells(i, "C").value = SectionType
        .Cells(i, "D").value = Section_Dim
        .Cells(i, "E").value = Total_Length
        .Cells(i, "G").value = "A36/SS400"
        .Cells(i, "H").value = 1
        .Cells(i, "I").value = SectionWeight
        .Cells(i, "J").value = .Cells(i, "E").value / 1000 * .Cells(i, "I").value
        .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value '重量小計
        .Cells(i, "L").value = "M"
        .Cells(i, "M").value = 1
        .Cells(i, "N").value = .Cells(i, "M").value * .Cells(i, "E").value / 1000 * .Cells(i, "H").value
        .Cells(i, "O").value = .Cells(i, "M").value * .Cells(i, "H").value
        .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
        .Cells(i, "Q").value = "素材類"
    End With
End Sub

Sub addPipeEntry(PipeSize As Variant, PipeThickness As Variant, Pipe_Length As Double)
    
    Set ws = Worksheets("Weight_Analysis")
    
    
End Sub
