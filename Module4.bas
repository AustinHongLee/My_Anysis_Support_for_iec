Attribute VB_Name = "Module4"
Sub Type_01(ByVal fullString As String)
    Dim PartString_Type As String
    Dim PipeSize As String
    Dim letter As String
    Dim pi As Double
    
    Set ws_M42 = Worksheets("Weight_Analysis")
    Set ws_Pipe_Table = Worksheets("Pipe_Table")
    Set ws = Worksheets("Weight_Analysis")
          

    
    '給M42資料
    PartString_Type = GetSecondPartOfString(fullString)
    PipeSize = Replace(PartString_Type, "B", "")
    

    letter = GetThirdPartOfString(fullString)
    letter = Right(letter, 1)
    

    
    'Main_Pipe
    Third_Length_export = Replace(GetThirdPartOfString(fullString), letter, "") * 100
        ' 處理主管與輔助管的編制:
            
            Select Case PipeSize
               Case 2
                Support_Pipe_Size = "'1.5"
                Pipe_ThickNess_mm = "SCH.80"
                L_Value = 70
                
               Case 3
                Support_Pipe_Size = "'2"
                Pipe_ThickNess_mm = "SCH.40"
                L_Value = 93
               
               Case 4
                Support_Pipe_Size = "'3"
                Pipe_ThickNess_mm = "SCH.40"
                L_Value = 139

               Case 6
                Support_Pipe_Size = "'4"
                Pipe_ThickNess_mm = "SCH.40"
                L_Value = 186

               Case 8
                Support_Pipe_Size = "'6"
                Pipe_ThickNess_mm = "SCH.40"
                L_Value = 271

               Case 10
                Support_Pipe_Size = "'8"
                Pipe_ThickNess_mm = "SCH.40"
                L_Value = 353

               Case 12
                Support_Pipe_Size = "'8"
                Pipe_ThickNess_mm = "SCH.40"
                L_Value = 370

               Case 14
                Support_Pipe_Size = "'10"
                Pipe_ThickNess_mm = "SCH.40"
                L_Value = 473
                
               Case 16
                Support_Pipe_Size = "'10"
                Pipe_ThickNess_mm = "SCH.40"
                L_Value = 491

               Case 18
                Support_Pipe_Size = "'12"
                Pipe_ThickNess_mm = "STD.WT"
                L_Value = 572
                
                Case 20
                Support_Pipe_Size = "'12"
                Pipe_ThickNess_mm = "STD.WT"
                L_Value = 594

                Case 24
                Support_Pipe_Size = "'14"
                Pipe_ThickNess_mm = "STD.WT"
                L_Value = 677
                
                Case 28
                Support_Pipe_Size = "'16"
                Pipe_ThickNess_mm = "STD.WT"
                L_Value = 782
                
                Case 36
                Support_Pipe_Size = "'24"
                Pipe_ThickNess_mm = "STD.WT"
                L_Value = 1099

                
    Case Else
        Exit Sub
End Select
    
    '主管長度與副管長度演算
    
    '主管長度 - 通常為SUS304
        Main_Pipe_Length = L_Value + 100
        Main_Pipe_Thickness = Pipe_ThickNess_mm
        Main_Pipe_Size = Support_Pipe_Size
        
            If Main_Pipe_Thickness <> "STD.WT" Then
            Main_Pipe_Thickness = Replace(Pipe_ThickNess_mm, "SCH.", "") & "S"
            Else
            Main_Pipe_Thickness = Pipe_ThickNess_mm
            End If
        
      'CalculatePipeDetails Fuction ( PipeSize , thickness(**S),Length,ws)
     Set MainPipeDetails = CalculatePipeDetails(Main_Pipe_Size, Main_Pipe_Thickness)
    
    '副管長度 - 通常為C12
        Support_Pipe_Length = Third_Length_export - 100
            
            If Pipe_ThickNess_mm <> "STD.WT" Then
            Support_Pipe_Thickness = Replace(Pipe_ThickNess_mm, "SCH.", "") & "S"
            Else
            Support_Pipe_Thickness = Pipe_ThickNess_mm
            End If
    'CalculatePipeDetails Fuction ( PipeSize , thickness(**S),Length,ws)
    Set SupportPipeDetails = CalculatePipeDetails(Support_Pipe_Size, Support_Pipe_Thickness)
    
    '固有資訊傳入 - 主管 - 副管 - 底板
          '主管
    ' PipeDetails.Add PipeDiameterInch, "DiameterInch" 管徑MM
    ' PipeDetails.Add PipeWeightPerMeter, "WeightPerMeter" 每米重
    ' PipeDetails.Add TotalWeight, "TotalWeight" '重量
    ' PipeDetails.Add Length, "Length"
    '   MainPipeDetails.Item("DiameterInch")
          
          i = GetNextRowInColumnB()
            
            If Main_Pipe_Thickness <> "STD.WT" Then
            Main_Pipe_Thickness = "SCH" & Replace(Main_Pipe_Thickness, "S", "")
            Else
            Main_Pipe_Thickness = Pipe_ThickNess_mm
            End If
            
            With ws
                .Cells(i, "B").value = 1 '項次
                .Cells(i, "C").value = "Pipe" '品名
                .Cells(i, "D").value = Main_Pipe_Size & """" & "*" & Main_Pipe_Thickness '尺寸厚度
                .Cells(i, "E").value = Main_Pipe_Length '長度
                .Cells(i, "G").value = "SUS304" '材值
                .Cells(i, "H").value = 1 '數量
                .Cells(i, "I").value = MainPipeDetails.Item("WeightPerMeter") '每米重
                .Cells(i, "J").value = .Cells(i, "E").value / 1000 * .Cells(i, "I").value '單重
                .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value '重量小計
                .Cells(i, "L").value = "M"
                .Cells(i, "M").value = 1 '組數
                .Cells(i, "N").value = .Cells(i, "H").value * .Cells(i, "M").value * .Cells(i, "E").value / 1000 '長度小計 組數*數量*長度/1000
                .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
                .Cells(i, "Q").value = "素材類"
            End With
          
          '評估上若副管長度小於等於0 則跳過
          If Support_Pipe_Length > 0 Then
            
          
          
          '副管
          i = GetNextRowInColumnB()
            If Support_Pipe_Thickness <> "STD.WT" Then
            Support_Pipe_Thickness = "SCH" & Replace(Support_Pipe_Thickness, "S", "")
            Else
            Support_Pipe_Thickness = Pipe_ThickNess_mm
            End If
            
            With ws
                .Cells(i, "B").value = 2 '項次
                .Cells(i, "C").value = "Pipe" '品名
                .Cells(i, "D").value = Support_Pipe_Size & """" & "*" & Support_Pipe_Thickness '尺寸厚度
                .Cells(i, "E").value = Support_Pipe_Length '長度
                .Cells(i, "G").value = "A53Gr.B" '材值
                .Cells(i, "H").value = 1 '數量
                .Cells(i, "I").value = SupportPipeDetails.Item("WeightPerMeter") '每米重
                .Cells(i, "J").value = .Cells(i, "E").value / 1000 * .Cells(i, "I").value '單重
                .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value '重量小計
                .Cells(i, "L").value = "M"
                .Cells(i, "M").value = 1 '組數
                .Cells(i, "N").value = .Cells(i, "H").value * .Cells(i, "M").value * .Cells(i, "E").value / 1000 '長度小計 組數*數量*長度/1000
                .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
                .Cells(i, "Q").value = "素材類"
            End With
           End If
    
    PipeSize = Replace(Support_Pipe_Size, "'", "")
    PerformActionByLetter letter, PipeSize
End Sub
Sub Type_05(ByVal fullString As String)
    '範例格式A : 20-L50-05L
    Dim PartString_Type As String
    Dim PipeSize As String
    Dim letter As String
    Dim SectionType As String
    Dim Section_Dim As String
    Dim Total_Length As Double

    
   
    Set ws = Worksheets("Weight_Analysis")
    
    '區分出角鐵尺寸
    PartString_Type = GetSecondPartOfString(fullString)
        Select Case PartString_Type
            
            Case "L50"
               The_Section_Size = "L50*50*6"
               SectionType = "Angle"
            Case "L65"
               The_Section_Size = "L65*65*6"
               SectionType = "Angle"
            Case "L75"
               The_Section_Size = "L75*75*9"
               SectionType = "Angle"
            End Select

    '區分出M42類型
        Support_05_Type_Choice_M42 = Right(GetThirdPartOfString(fullString), 1)
    '區分出長度"H"
         Section_Length_H = Replace(GetThirdPartOfString(fullString), Support_05_Type_Choice_M42, "") * 100
         Section_Length_L = 130
         
       '轉換為部分必要需求 :
        letter = Support_05_Type_Choice_M42
        PipeSize = The_Section_Size

      
      
      '導入Function addSteelSectionEntry
            SectionType = SectionType
            Section_Dim = Replace(The_Section_Size, Left(The_Section_Size, 1), "")
            Total_Length = Section_Length_H + Section_Length_L
            AddSteelSectionEntry SectionType, Section_Dim, Total_Length
            PerformActionByLetter letter, PipeSize
End Sub
Sub Type_09(ByVal fullString As String)
    '範例格式A : 09-2B-05B
    Dim PartString_Type As String
    Dim PipeSize As String
    Dim letter As String
    Dim pi As Double
    
    Set ws_M42 = Worksheets("Weight_Analysis")
    Set ws_Pipe_Table = Worksheets("Pipe_Table")
    Set ws = Worksheets("Weight_Analysis")
          

    
    '給M42資料
    PartString_Type = GetSecondPartOfString(fullString)
    PipeSize = Replace(PartString_Type, "B", "")
    

    letter = GetThirdPartOfString(fullString)
    letter = Right(letter, 1)
    

    
    'Main_Pipe
    Third_Length_export = Replace(GetThirdPartOfString(fullString), letter, "") * 100
        ' 處理主管與輔助管的編制:
            
            Select Case PipeSize
               Case 2
                Support_Pipe_Size = "'2"
                Pipe_ThickNess_mm = "SCH.80"
                L_Value = 106
                
               Case 3
                Support_Pipe_Size = "'2"
                Pipe_ThickNess_mm = "SCH.40"
                L_Value = 93
               
               Case 4
                Support_Pipe_Size = "'2"
                Pipe_ThickNess_mm = "SCH.40"
                L_Value = 106


                
    Case Else
        Exit Sub
End Select
    
    '主管長度與副管長度演算
    
    '主管長度 - 通常為SUS304
        Main_Pipe_Length = L_Value + 100
        Main_Pipe_Thickness = Pipe_ThickNess_mm
        Main_Pipe_Size = Support_Pipe_Size
        
            If Main_Pipe_Thickness <> "STD.WT" Then
            Main_Pipe_Thickness = Replace(Pipe_ThickNess_mm, "SCH.", "") & "S"
            Else
            Main_Pipe_Thickness = Pipe_ThickNess_mm
            End If
        
      'CalculatePipeDetails Fuction ( PipeSize , thickness(**S),Length,ws)
     Set MainPipeDetails = CalculatePipeDetails(Main_Pipe_Size, Main_Pipe_Thickness)
    
    '副管長度 - 通常為C12
        Support_Pipe_Length = Third_Length_export - 100
            
            If Pipe_ThickNess_mm <> "STD.WT" Then
            Support_Pipe_Thickness = Replace(Pipe_ThickNess_mm, "SCH.", "") & "S"
            Else
            Support_Pipe_Thickness = Pipe_ThickNess_mm
            End If
    'CalculatePipeDetails Fuction ( PipeSize , thickness(**S),Length,ws)
    Set SupportPipeDetails = CalculatePipeDetails(Support_Pipe_Size, Support_Pipe_Thickness)
    
    '固有資訊傳入 - 主管 - 副管 - 底板
          '主管
    ' PipeDetails.Add PipeDiameterInch, "DiameterInch" 管徑MM
    ' PipeDetails.Add PipeWeightPerMeter, "WeightPerMeter" 每米重
    ' PipeDetails.Add TotalWeight, "TotalWeight" '重量
    ' PipeDetails.Add Length, "Length"
    '   MainPipeDetails.Item("DiameterInch")
          
          i = GetNextRowInColumnB()
            
            If Main_Pipe_Thickness <> "STD.WT" Then
            Main_Pipe_Thickness = "SCH" & Replace(Main_Pipe_Thickness, "S", "")
            Else
            Main_Pipe_Thickness = Pipe_ThickNess_mm
            End If
            
            With ws
                .Cells(i, "B").value = 1 '項次
                .Cells(i, "C").value = "Pipe" '品名
                .Cells(i, "D").value = Main_Pipe_Size & """" & "*" & Main_Pipe_Thickness '尺寸厚度
                .Cells(i, "E").value = Main_Pipe_Length '長度
                .Cells(i, "G").value = "SUS304" '材值
                .Cells(i, "H").value = 1 '數量
                .Cells(i, "I").value = MainPipeDetails.Item("WeightPerMeter") '每米重
                .Cells(i, "J").value = .Cells(i, "E").value / 1000 * .Cells(i, "I").value '單重
                .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value '重量小計
                .Cells(i, "L").value = "M"
                .Cells(i, "M").value = 1 '組數
                .Cells(i, "N").value = .Cells(i, "H").value * .Cells(i, "M").value * .Cells(i, "E").value / 1000 '長度小計 組數*數量*長度/1000
                .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
                .Cells(i, "Q").value = "素材類"
            End With
          
          '評估上若副管長度小於等於0 則跳過
          If Support_Pipe_Length > 0 Then
            
          
          
          '副管
          i = GetNextRowInColumnB()
            If Support_Pipe_Thickness <> "STD.WT" Then
            Support_Pipe_Thickness = "SCH" & Replace(Support_Pipe_Thickness, "S", "")
            Else
            Support_Pipe_Thickness = Pipe_ThickNess_mm
            End If
            
            With ws
                .Cells(i, "B").value = 2 '項次
                .Cells(i, "C").value = "Pipe" '品名
                .Cells(i, "D").value = Support_Pipe_Size & """" & "*" & Support_Pipe_Thickness '尺寸厚度
                .Cells(i, "E").value = Support_Pipe_Length '長度
                .Cells(i, "G").value = "A53Gr.B" '材值
                .Cells(i, "H").value = 1 '數量
                .Cells(i, "I").value = SupportPipeDetails.Item("WeightPerMeter") '每米重
                .Cells(i, "J").value = .Cells(i, "E").value / 1000 * .Cells(i, "I").value '單重
                .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value '重量小計
                .Cells(i, "L").value = "M"
                .Cells(i, "M").value = 1 '組數
                .Cells(i, "N").value = .Cells(i, "H").value * .Cells(i, "M").value * .Cells(i, "E").value / 1000 '長度小計 組數*數量*長度/1000
                .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
                .Cells(i, "Q").value = "素材類"
            End With
           End If
    
    PipeSize = Replace(Support_Pipe_Size, "'", "")
    PerformActionByLetter letter, PipeSize
       
       '導入09-Tpye特有屬性 : Machine Bolt
       ' 填充數據
    i = GetNextRowInColumnB()
    With ws
        .Cells(i, "B").value = .Cells(i - 1, "B").value + 1
        .Cells(i, "C").value = "MACHINE BOLT"
        .Cells(i, "D").value = "1-5/8""""*150L"
        .Cells(i, "G").value = "A307Gr.B(熱浸鋅)"
        .Cells(i, "H").value = 1
        .Cells(i, "J").value = 20 ' 假設每個螺栓的單個重量是20（可以根據實際情況調整）
        .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value
        .Cells(i, "L").value = "SET"
        .Cells(i, "M").value = 1
        .Cells(i, "O").value = .Cells(i, "M").value * .Cells(i, "H").value
        .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
        .Cells(i, "Q").value = "螺絲類"
    End With


End Sub
Sub Type_14(ByVal fullString As String)
    '範例格式A : 14-2B-1005
    Dim PartString_Type As String
    Dim PipeSize As String
    Dim letter As String
    Dim pi As Double
    Dim SectionType As String
    Dim Section_Dim As String
    Dim Total_Length As Double
    Dim BoltSize As String
    
    Set ws_M42 = Worksheets("Weight_Analysis")
    Set ws_Pipe_Table = Worksheets("Pipe_Table")
    Set ws = Worksheets("Weight_Analysis")
    Set Type_14_Table = Worksheets("For_14_Type_data")

    
    '給定管尺寸
    PartString_Type = GetSecondPartOfString(fullString)
    PipeSize = Replace(PartString_Type, "B", "")
    '給定H&L 長度
    Section_Length_L = Left(GetThirdPartOfString(fullString), 2) * 100
    '注意 以下給個H值為暫定
    Pipe_Length_H_part = Right(GetThirdPartOfString(fullString), 2) * 100
    
    '主管長度 - 通常為SUS304
        
        ' 計算實際管子需求長度
        PipeSize = GetLookupValue(PipeSize)
        BpLength = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 5, False) 'F
        SL = Replace(Left(Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 12, False), 4), "C", "")  'N
        BTLength = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 5, False) 'F
        ' 計算管子厚度
            Select Case PipeSize
               Case 2
                Pipe_ThickNess_mm = "SCH.40"
                
               Case 3
                Pipe_ThickNess_mm = "SCH.40"
               
               Case 4
                Pipe_ThickNess_mm = "SCH.40"

               Case 6
                Pipe_ThickNess_mm = "SCH.40"

               Case 8
                Pipe_ThickNess_mm = "SCH.40"

               Case 10
                Pipe_ThickNess_mm = "SCH.40"

               Case 12
                Pipe_ThickNess_mm = "STD.WT"


    Case Else
        Exit Sub
End Select
        
        Main_Pipe_Length = Pipe_Length_H_part - BpLength - SL - BTLength
        Main_Pipe_Thickness = Pipe_ThickNess_mm
        Main_Pipe_Size = PipeSize
        
            If Main_Pipe_Thickness <> "STD.WT" Then
            Main_Pipe_Thickness = Replace(Pipe_ThickNess_mm, "SCH.", "") & "S"
            Else
            Main_Pipe_Thickness = Pipe_ThickNess_mm
            End If
        
      'CalculatePipeDetails Fuction ( PipeSize , thickness(**S),Length,ws)
     Set MainPipeDetails = CalculatePipeDetails(Main_Pipe_Size, Main_Pipe_Thickness)
    
    '固有資訊傳入 - 主管 - 鋼構 - Plate(wing) - Plate(STOPPER) - Plate(BASE PLATE) - Plate(TOP)
          '主管
    ' PipeDetails.Add PipeDiameterInch, "DiameterInch" 管徑MM
    ' PipeDetails.Add PipeWeightPerMeter, "WeightPerMeter" 每米重
    ' PipeDetails.Add TotalWeight, "TotalWeight" '重量
    ' PipeDetails.Add Length, "Length"
    '   MainPipeDetails.Item("DiameterInch")
          
          i = GetNextRowInColumnB()
            
            If Main_Pipe_Thickness <> "STD.WT" Then
            Main_Pipe_Thickness = "SCH" & Replace(Main_Pipe_Thickness, "S", "")
            Else
            Main_Pipe_Thickness = Pipe_ThickNess_mm
            End If
            
            With ws
                .Cells(i, "B").value = 1 '項次
                .Cells(i, "C").value = "Pipe" '品名
                .Cells(i, "D").value = Main_Pipe_Size & """" & "*" & Main_Pipe_Thickness '尺寸厚度
                .Cells(i, "E").value = Main_Pipe_Length '長度
                .Cells(i, "G").value = "SUS304" '材值
                .Cells(i, "H").value = 1 '數量
                .Cells(i, "I").value = MainPipeDetails.Item("WeightPerMeter") '每米重
                .Cells(i, "J").value = .Cells(i, "E").value / 1000 * .Cells(i, "I").value '單重
                .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value '重量小計
                .Cells(i, "L").value = "M"
                .Cells(i, "M").value = 1 '組數
                .Cells(i, "N").value = .Cells(i, "H").value * .Cells(i, "M").value * .Cells(i, "E").value / 1000 '長度小計 組數*數量*長度/1000
                .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
                .Cells(i, "Q").value = "素材類"
            End With
'導入鋼構


               The_Section_Size = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 12, False)
               SectionType = "Channel"
       '導入Function addSteelSectionEntry
            SectionType = SectionType
            Section_Dim = Replace(The_Section_Size, Left(The_Section_Size, 1), "")
            Total_Length = Section_Length_L
            AddSteelSectionEntry SectionType, Section_Dim, Total_Length

'導入14-Tpye特有屬性 : Plate(wing)_14Type
' 填充數據
            PipeSize = GetLookupValue(PipeSize)
            Plate_Size_a = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 9, False) 'Q
            Plate_Size_b = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 8, False) 'P
            Plate_Thickness = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 5, False) 'F
            Weight_calculator = Plate_Size_a / 1000 * Plate_Size_b / 1000 * Plate_Thickness * 7.85

            i = GetNextRowInColumnB()
                  With ws
                      .Cells(i, "B").value = .Cells(i - 1, "B").value + 1 '項次
                      .Cells(i, "C").value = "Plate(wing)_14Type"
                      .Cells(i, "D").value = Plate_Thickness
                      .Cells(i, "E").value = Plate_Size_a
                      .Cells(i, "F").value = Plate_Size_b
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

'導入14-Tpye特有屬性 : Plate(STOPPER)_14Type
' 填充數據
            PipeSize = GetLookupValue(PipeSize)
            Plate_Size_a = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 7, False) 'M
            Plate_Size_b = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 6, False) 'K
            Plate_Thickness = 6
            Weight_calculator = Plate_Size_a / 1000 * Plate_Size_b / 1000 * Plate_Thickness * 7.85

            i = GetNextRowInColumnB()
                  With ws
                      .Cells(i, "B").value = .Cells(i - 1, "B").value + 1 '項次
                      .Cells(i, "C").value = "Plate(STOPPER)_14Type"
                      .Cells(i, "D").value = Plate_Thickness
                      .Cells(i, "E").value = Plate_Size_a
                      .Cells(i, "F").value = Plate_Size_b
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

'導入14-Tpye特有屬性 : Plate(BASE PLATE)_14Type
' 填充數據
            PipeSize = GetLookupValue(PipeSize)
            Plate_Size_a = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 2, False) 'C
            Plate_Size_b = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 2, False) 'C
            Plate_Thickness = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 5, False) 'F
            Weight_calculator = Plate_Size_a / 1000 * Plate_Size_b / 1000 * Plate_Thickness * 7.85

            i = GetNextRowInColumnB()
                  With ws
                      .Cells(i, "B").value = .Cells(i - 1, "B").value + 1 '項次
                      .Cells(i, "C").value = "Plate(BASE PLATE)_14Type"
                      .Cells(i, "D").value = Plate_Thickness
                      .Cells(i, "E").value = Plate_Size_a
                      .Cells(i, "F").value = Plate_Size_b
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

'導入14-Tpye特有屬性 : Plate(TOP)_14Type
' 填充數據
            PipeSize = GetLookupValue(PipeSize)
            Plate_Size_a = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 11, False) 'C
            Plate_Size_b = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 11, False) 'C
            Plate_Thickness = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 5, False) 'F
            Weight_calculator = Plate_Size_a / 1000 * Plate_Size_b / 1000 * Plate_Thickness * 7.85

            i = GetNextRowInColumnB()
                  With ws
                      .Cells(i, "B").value = .Cells(i - 1, "B").value + 1 '項次
                      .Cells(i, "C").value = "Plate(TOP)_14Type"
                      .Cells(i, "D").value = Plate_Thickness
                      .Cells(i, "E").value = Plate_Size_a
                      .Cells(i, "F").value = Plate_Size_b
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
'導入14-Tpye特有屬性 : EXP.BOLT
' 填充數據
   BoltSize = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 10, False) 'J
    
    With ws
        .Cells(i, "B").value = .Cells(i - 1, "B").value + 1
        .Cells(i, "C").value = "EXP.BOLT"
        .Cells(i, "D").value = "'" & BoltSize & """"
        .Cells(i, "G").value = "SUS304"
        .Cells(i, "H").value = 4
        .Cells(i, "J").value = 1 ' 假設每個螺栓的單個重量是1（可以根據實際情況調整）
        .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value
        .Cells(i, "L").value = "SET"
        .Cells(i, "M").value = 1
        .Cells(i, "O").value = .Cells(i, "M").value * .Cells(i, "H").value
        .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
        .Cells(i, "Q").value = "螺絲類"
    End With


End Sub
Sub Type_16(ByVal fullString As String)
    ' 這裡是 Type_16 的程式碼
    ' 您可以使用 fullString 參數來執行所需的操作
    Set ws_Pipe_Table = Worksheets("Pipe_Table")
    Set ws = Worksheets("Weight_Analysis")



        PartString_Type = GetSecondPartOfString(fullString)
 '分解部分 範例 : 16-2B-04
        
        'Second For Main Pipe Size
            PartString_Type = GetSecondPartOfString(fullString)
            PipeSize = Replace(PartString_Type, "B", "")
        'Third For H Value
            Third_Length_export = GetThirdPartOfString(fullString) * 100
            
            
        ' 處理主管與輔助管的編制:
            
            Select Case PipeSize
               Case 2
                Support_Pipe_Size = "'1.5"
                Pipe_ThickNess_mm = "SCH.80"
                Plate_Size = 70
                
               Case 3
                Support_Pipe_Size = "'2"
                Pipe_ThickNess_mm = "SCH.40"
                Plate_Size = 80
               
               Case 4
                Support_Pipe_Size = "'3"
                Pipe_ThickNess_mm = "SCH.40"
                Plate_Size = 110

               Case 6
                Support_Pipe_Size = "'4"
                Pipe_ThickNess_mm = "SCH.40"
                Plate_Size = 140

               Case 8
                Support_Pipe_Size = "'6"
                Pipe_ThickNess_mm = "SCH.40"
                Plate_Size = 190

               Case 10
                Support_Pipe_Size = "'8"
                Pipe_ThickNess_mm = "SCH.40"
                Plate_Size = 240

               Case 12
                Support_Pipe_Size = "'10"
                Pipe_ThickNess_mm = "SCH.40"
                Plate_Size = 290

               Case 14
                Support_Pipe_Size = "'12"
                Pipe_ThickNess_mm = "STD.WT"
                Plate_Size = 340
                
               Case 16
                Support_Pipe_Size = "'12"
                Pipe_ThickNess_mm = "STD.WT"
                Plate_Size = 340

               Case 18
                Support_Pipe_Size = "'14"
                Pipe_ThickNess_mm = "STD.WT"
                Plate_Size = 380
                
                Case 20
                Support_Pipe_Size = "'14"
                Pipe_ThickNess_mm = "STD.WT"
                Plate_Size = 380

                Case 24
                Support_Pipe_Size = "'16"
                Pipe_ThickNess_mm = "STD.WT"
                Plate_Size = 430
    Case Else
        Exit Sub
End Select
    
    '主管長度與副管長度演算
    
    '主管長度 - 通常為SUS304
        Main_Pipe_Length = Round((PipeSize * 1.5 * 25.4) + (Main_Pipe_inch_mm / 2) + 100)
        Main_Pipe_Thickness = "40S"
        Main_Pipe_Size = PipeSize
        
      'CalculatePipeDetails Fuction ( PipeSize , thickness(**S),Length,ws)
     Set MainPipeDetails = CalculatePipeDetails(Main_Pipe_Size, "40S")
    
    '副管長度 - 通常為C12
        Support_Pipe_Length = Round(Third_Length_export - (Support_Pipe_inch_mm / 2) - 100 + 300)
            
            If Pipe_ThickNess_mm <> "STD.WT" Then
            Support_Pipe_Thickness = Replace(Pipe_ThickNess_mm, "SCH.", "") & "S"
            Else
            Support_Pipe_Thickness = Pipe_ThickNess_mm
            End If
    'CalculatePipeDetails Fuction ( PipeSize , thickness(**S),Length,ws)
    Set SupportPipeDetails = CalculatePipeDetails(Support_Pipe_Size, Support_Pipe_Thickness)
    
    '固有資訊傳入 - 主管 - 副管 - 底板
          '主管
    ' PipeDetails.Add PipeDiameterInch, "DiameterInch" 管徑MM
    ' PipeDetails.Add PipeWeightPerMeter, "WeightPerMeter" 每米重
    ' PipeDetails.Add TotalWeight, "TotalWeight" '重量
    ' PipeDetails.Add Length, "Length"
    '   MainPipeDetails.Item("DiameterInch")
          
          i = GetNextRowInColumnB()
            With ws
                .Cells(i, "B").value = 1 '項次
                .Cells(i, "C").value = "Pipe" '品名
                .Cells(i, "D").value = Main_Pipe_Size & """" & "*" & "SCH" & Replace(Main_Pipe_Thickness, "S", "") '尺寸厚度
                .Cells(i, "E").value = Main_Pipe_Length '長度
                .Cells(i, "G").value = "SUS304" '材值
                .Cells(i, "H").value = 1 '數量
                .Cells(i, "I").value = MainPipeDetails.Item("WeightPerMeter") '每米重
                .Cells(i, "J").value = .Cells(i, "E").value / 1000 * .Cells(i, "I").value '單重
                .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value '重量小計
                .Cells(i, "L").value = "M"
                .Cells(i, "M").value = 1 '組數
                .Cells(i, "N").value = .Cells(i, "H").value * .Cells(i, "M").value * .Cells(i, "E").value / 1000 '長度小計 組數*數量*長度/1000
                .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
                .Cells(i, "Q").value = "素材類"
            End With
          
          '副管
          i = GetNextRowInColumnB()
            With ws
                .Cells(i, "B").value = 2 '項次
                .Cells(i, "C").value = "Pipe" '品名
                .Cells(i, "D").value = Support_Pipe_Size & """" & "*" & "SCH" & Replace(Support_Pipe_Thickness, "S", "") '尺寸厚度
                .Cells(i, "E").value = Support_Pipe_Length '長度
                .Cells(i, "G").value = "A53Gr.B" '材值
                .Cells(i, "H").value = 1 '數量
                .Cells(i, "I").value = SupportPipeDetails.Item("WeightPerMeter") '每米重
                .Cells(i, "J").value = .Cells(i, "E").value / 1000 * .Cells(i, "I").value '單重
                .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value '重量小計
                .Cells(i, "L").value = "M"
                .Cells(i, "M").value = 1 '組數
                .Cells(i, "N").value = .Cells(i, "H").value * .Cells(i, "M").value * .Cells(i, "E").value / 1000 '長度小計 組數*數量*長度/1000
                .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
                .Cells(i, "Q").value = "素材類"
            End With
            
            '底板
            Plate_Thickness = 6
            Weight_calculator = Plate_Size / 1000 * Plate_Size / 1000 * Plate_Thickness * 7.85

            i = GetNextRowInColumnB()
                  With ws
                      .Cells(i, "B").value = 3 '項次
                      .Cells(i, "C").value = "Plate"
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
Sub Type_20(ByVal fullString As String)
    '範例格式A : 20-L50-05A
    Dim SectionType As String
    Dim Section_Dim As String
    Dim Total_Length As Double

    
   
    Set ws = Worksheets("Weight_Analysis")
    
    '區分出角鐵尺寸
    PartString_Type = GetSecondPartOfString(fullString)
        Select Case PartString_Type
            
            Case "L50"
               The_Section_Size = "L50*50*6"
               SectionType = "Angle"
            Case "L65"
               The_Section_Size = "L65*65*6"
               SectionType = "Angle"
            Case "L75"
               The_Section_Size = "L75*75*9"
               SectionType = "Angle"
            Case "C100"
               The_Section_Size = "C100*50*5"
               SectionType = "Channel"
            End Select

    '區分出Fig類型
        Support_23_Type_Choice = Right(GetThirdPartOfString(fullString), 1)
    '區分出長度"H"
         Section_Length_H = Replace(GetThirdPartOfString(fullString), Support_23_Type_Choice, "") * 100
    
      

      
      
      '導入Function addSteelSectionEntry
            SectionType = SectionType
            Section_Dim = Replace(The_Section_Size, Left(The_Section_Size, 1), "")
            Total_Length = Section_Length_H
            AddSteelSectionEntry SectionType, Section_Dim, Total_Length

End Sub
Sub Type_21(ByVal fullString As String)
    '範例格式A : 21-L50-05A
    Dim SectionType As String
    Dim Section_Dim As String
    Dim Total_Length As Double
    '範例格式B : 21-L50-05C-07
    
   
    Set ws = Worksheets("Weight_Analysis")
    
    '區分出角鐵尺寸
    PartString_Type = GetSecondPartOfString(fullString)
        Select Case PartString_Type
            
            Case "L50"
               The_Angle_Size = "L50*50*6"
            Case "L65"
               The_Angle_Size = "L65*65*6"
            Case "L75"
               The_Angle_Size = "L75*75*9"
            End Select
                  ' For Angle
    '區分出Fig類型
        Support_21_Type_Choice = Right(GetThirdPartOfString(fullString), 1)
    '區分出長度"H"
         Section_Length_H = Replace(GetThirdPartOfString(fullString), Support_21_Type_Choice, "") * 100
    
    '區分出長度"L"
    'if Support_21_Type_Choice = "A" then Section_Length_L = 300
    'if Support_21_Type_Choice = "B" then Section_Length_L = 500
    'if Support_21_Type_Choice = "C" then Section_Length_L = GetFourthPartOfString(fullString) *100
    
    Select Case Support_21_Type_Choice
        Case "A"
            Section_Length_L = 300
        Case "B"
            Section_Length_L = 500
        Case "C"
            Section_Length_L = GetFourthPartOfString(fullString) * 100
    End Select
      

      
      
      '導入Function addSteelSectionEntry
            SectionType = "Angle"
            Section_Dim = Replace(The_Angle_Size, "L", "")
            Total_Length = Section_Length_H + Section_Length_L
            AddSteelSectionEntry SectionType, Section_Dim, Total_Length

End Sub
Sub Type_22(ByVal fullString As String)
    '範例格式A : 22-L50-05A(L)
    Dim PartString_Type As String
    Dim PipeSize As String
    Dim letter As String
    Dim SectionType As String
    Dim Section_Dim As String
    Dim Total_Length As Double
    '範例格式B : 21-L50-05(L)C-07
    
   
    Set ws = Worksheets("Weight_Analysis")
    
    '區分出角鐵尺寸
    PartString_Type = GetSecondPartOfString(fullString)
        Select Case PartString_Type
            
            Case "L50"
               The_Angle_Size = "L50*50*6"
            Case "L65"
               The_Angle_Size = "L65*65*6"
            Case "L75"
               The_Angle_Size = "L75*75*9"
            End Select
                  ' For Angle
    '區分出Fig類型
        Support_22_Type_Choice = Mid(Right(GetThirdPartOfString(fullString), 3), 1, 1)
    '區分出M-42類型
        Support_22_Type_Choice_M42 = Right(GetThirdPartOfString(fullString), 1)
    
    '修剪出 Replace 邏輯 for 長度
        Type_22_Replace_A = "(" & Support_22_Type_Choice & ")"
        Type_22_Replace_B = Support_22_Type_Choice_M42
    '區分出長度"H"
        Section_Length_H = Replace(Replace(GetThirdPartOfString(fullString), Type_22_Replace_A, ""), Type_22_Replace_B, "") * 100
        
    
    '區分出長度"L"
    'if Support_21_Type_Choice = "A" then Section_Length_L = 300
    'if Support_21_Type_Choice = "B" then Section_Length_L = 500
    'if Support_21_Type_Choice = "C" then Section_Length_L = GetFourthPartOfString(fullString) *100
    
    Select Case Support_22_Type_Choice
        Case "A"
            Section_Length_L = 300
        Case "B"
            Section_Length_L = 500
        Case "C"
            Section_Length_L = GetFourthPartOfString(fullString) * 100
    End Select
      '轉換為部分必要需求 :
        letter = Support_22_Type_Choice_M42
        PipeSize = The_Angle_Size
      '導入Function addSteelSectionEntry
            
            SectionType = "Angle"
            Section_Dim = Replace(The_Angle_Size, "L", "")
            Total_Length = Section_Length_H + Section_Length_L
            AddSteelSectionEntry SectionType, Section_Dim, Total_Length
            PerformActionByLetter letter, PipeSize
End Sub
Sub Type_23(ByVal fullString As String)
    '範例格式A : 23-L50-05A
    Dim SectionType As String
    Dim Section_Dim As String
    Dim Total_Length As Double
    '範例格式B : 23-L50-05C-07
    
   
    Set ws = Worksheets("Weight_Analysis")
    
    '區分出角鐵尺寸
    PartString_Type = GetSecondPartOfString(fullString)
        Select Case PartString_Type
            
            Case "L50"
               The_Section_Size = "L50*50*6"
               SectionType = "Angle"
            Case "L65"
               The_Section_Size = "L65*65*6"
               SectionType = "Angle"
            Case "L75"
               The_Section_Size = "L75*75*9"
               SectionType = "Angle"
            Case "L100"
               The_Section_Size = "L100*100*10"
               SectionType = "Angle"
            Case "C100"
               The_Section_Size = "C100*50*5"
               SectionType = "Channel"
            Case "C150"
               The_Section_Size = "C150*75*9"
               SectionType = "Channel"
            Case "H100"
               The_Section_Size = "H100*100*6"
               SectionType = "H beam"
            Case "H150"
               The_Section_Size = "H150*150*7"
               SectionType = "H beam"
                        
            End Select

    '區分出Fig類型
        Support_23_Type_Choice = Right(GetThirdPartOfString(fullString), 1)
    '區分出長度"H"
         Section_Length_H = Replace(GetThirdPartOfString(fullString), Support_23_Type_Choice, "") * 100
    
    '區分出長度"L"
    'if Support_23_Type_Choice = "A" then Section_Length_L = 300
    'if Support_23_Type_Choice = "B" then Section_Length_L = 500
    'if Support_23_Type_Choice = "C" then Section_Length_L = GetFourthPartOfString(fullString) *100
    
    Select Case Support_23_Type_Choice
        Case "A"
            Section_Length_L = 300
        Case "B"
            Section_Length_L = 500
        Case "C"
            Section_Length_L = GetFourthPartOfString(fullString) * 100
    End Select
      

      
      
      '導入Function addSteelSectionEntry
            SectionType = SectionType
            Section_Dim = Replace(The_Section_Size, Left(The_Section_Size, 1), "")
            Total_Length = Section_Length_H + Section_Length_L
            AddSteelSectionEntry SectionType, Section_Dim, Total_Length

End Sub
Sub Type_24(ByVal fullString As String)
    '範例格式A : 24-L50-05
    Dim SectionType As String
    Dim Section_Dim As String
    Dim Total_Length As Double
   
    Set ws = Worksheets("Weight_Analysis")
    
    '區分出角鐵尺寸
    PartString_Type = GetSecondPartOfString(fullString)
        Select Case PartString_Type
            
            Case "L50"
               The_Angle_Size = "L50*50*6"
            Case "L75"
               The_Angle_Size = "L75*75*9"
            End Select
                  
                  ' For Angle
    '區分出長度"H"
         Section_Length_H = GetThirdPartOfString(fullString) * 100

           
      '導入Function addSteelSectionEntry
            SectionType = "Angle"
            Section_Dim = Replace(The_Angle_Size, "L", "")
            Total_Length = Section_Length_H
            AddSteelSectionEntry SectionType, Section_Dim, Total_Length

End Sub
Sub Type_25(ByVal fullString As String)
    '範例格式A : 25-L50-0505A
    Dim SectionType As String
    Dim Section_Dim As String
    Dim Total_Length As Double
    '範例格式B : 23-L50-0505C-0401
    
   
    Set ws = Worksheets("Weight_Analysis")
    
    
    
    '區分出角鐵尺寸
    PartString_Type = GetSecondPartOfString(fullString)
        Select Case PartString_Type
            
            Case "L50"
               The_Section_Size = "L50*50*6"
               SectionType = "Angle"
            Case "L65"
               The_Section_Size = "L65*65*6"
               SectionType = "Angle"
            Case "L75"
               The_Section_Size = "L75*75*9"
               SectionType = "Angle"

                        
            End Select
    '區分出Fig類型
    Support_23_Type_Choice = Right(GetThirdPartOfString(fullString), 1)
    '區分出"H"值
    Section_Length_H = Left(GetThirdPartOfString(fullString), 2) * 100
    '區分出"L"值
    Section_Length_L = Replace(Right(GetThirdPartOfString(fullString), 3), Support_23_Type_Choice, "") * 100

      

      
      
      '導入Function addSteelSectionEntry
            SectionType = SectionType
            Section_Dim = Replace(The_Section_Size, Left(The_Section_Size, 1), "")
            Total_Length = Section_Length_H + Section_Length_L
            AddSteelSectionEntry SectionType, Section_Dim, Total_Length

End Sub
Sub Type_108(ByVal fullString As String)
    Dim needValue As Variant
    
    Dim PartString_Type As String
    Dim PipeSize As String
    Dim letter As String
    Dim pi As Double
    
    '範例格式 : 108-1B-12E-A(S)
    'Need use GetFourthPartOfString
    '108=Type
    '1B =Denote Line Size "D"
    '12 =Denote Dimension "H" (IN 100mm)
    'E  =Denote the M42 Type
    'A  =分為Fig.A & Fig.B & Fig.C Lug Plate 的區別
    '(S)=材質區分
    
    Set ws_M42 = Worksheets("Weight_Analysis")
    Set ws_Pipe_Table = Worksheets("Pipe_Table")
    Set ws = Worksheets("Weight_Analysis")
        
        '區分出尺吋 符合Line Size : "D"
            PartString_Type = GetSecondPartOfString(fullString)
            PipeSize = Replace(PartString_Type, "B", "")
          
          '區分出M42板類型
        letter = GetThirdPartOfString(fullString)
        letter = Right(letter, 1)
        
        '區分出"H"值並乘上100
        Main_Pipe_Length = Replace(GetThirdPartOfString(fullString), letter, "") * 100
        

        
        '區分出Fig number
        needValue = ExtractParts(GetFourthPartOfString(fullString))
        Fig_number = needValue(0)
                
                Select Case Fig_number
                Case "A"
                    Fig = "Fig_A"
                Case "B"
                    Fig = "Fig_B"
                Case "C"
                    Fig = "Fig_C"
                
                Case Else
                   Exit Sub
                End Select
        
        
        
        '區分出 材質
        Mtl = needValue(1)
            If Mtl = "" Then
                Mtl = "A36"
            Else
                Mtl = needValue(1)
                Select Case Mtl
                
                Case "(A)"
                    Mtl = "A36/SS400"
                Case "(S)"
                    Mtl = "SUS304"
                Case Else
                   Exit Sub
                End Select
        
        
        
            End If
        
        
        ' 具有雙向制約檢測
        ' 如果denote "D" = 3/4" H >1000 Then 1.5"_Sch80 else 1"_Sch80
        ' 如果denote "D" = 1" H >1000 Then 1.5"_Sch80 else 1"_Sch80
        ' 如果denote "D" = 1.5" H >1000 Then 2"_Sch40 else 1..5"_Sch80
        ' 如果denote "D" = 2" = 2"_Sch40
        
        '實際換算出 所需主管 與 主管厚度
     Select Case PipeSize
        Case "1/2"
            If Third_Length_export > 1001 Then
                Main_Pipe_Size = "1.5"
                Main_Pipe_Thickness = "80S"
                Set MainPipeDetails = CalculatePipeDetails(Main_Pipe_Size, Main_Pipe_Thickness)
            Else
                Main_Pipe_Size = "1"
                Main_Pipe_Thickness = "80S"
                Set MainPipeDetails = CalculatePipeDetails(Main_Pipe_Size, Main_Pipe_Thickness)
                
            End If
        Case "3/4"
            If Third_Length_export > 1001 Then
                Main_Pipe_Size = "1.5"
                Main_Pipe_Thickness = "80S"
                Set MainPipeDetails = CalculatePipeDetails(Main_Pipe_Size, Main_Pipe_Thickness)
            Else
                Main_Pipe_Size = "1"
                Main_Pipe_Thickness = "80S"
                Set MainPipeDetails = CalculatePipeDetails(Main_Pipe_Size, Main_Pipe_Thickness)
                
            End If
        
        Case "1"
            If Third_Length_export > 1001 Then
                Main_Pipe_Size = "1.5"
                Main_Pipe_Thickness = "80S"
                Set MainPipeDetails = CalculatePipeDetails(Main_Pipe_Size, Main_Pipe_Thickness)
            Else
                Main_Pipe_Size = "1"
                Main_Pipe_Thickness = "80S"
                Set MainPipeDetails = CalculatePipeDetails(Main_Pipe_Size, Main_Pipe_Thickness)
                
            End If
        
        Case "1-1/2"
            If Third_Length_export > 1001 Then
                Main_Pipe_Size = "2"
                Main_Pipe_Thickness = "80S"
                Set MainPipeDetails = CalculatePipeDetails(Main_Pipe_Size, Main_Pipe_Thickness)
            Else
                Main_Pipe_Size = "1.5"
                Main_Pipe_Thickness = "80S"
                Set MainPipeDetails = CalculatePipeDetails(Main_Pipe_Size, Main_Pipe_Thickness)
                
            End If
        
         Case "2"

                Main_Pipe_Size = "2"
                Main_Pipe_Thickness = "40S"
                Set MainPipeDetails = CalculatePipeDetails(Main_Pipe_Size, Main_Pipe_Thickness)
        End Select
   
   '導入式: 管 -> M42 -> 獨特板子
                  
           '管子導入式
                  i = GetNextRowInColumnB()
            With ws
                .Cells(i, "B").value = 1 '項次
                .Cells(i, "C").value = "Pipe" '品名
                .Cells(i, "D").value = Main_Pipe_Size & """" & "*" & "SCH" & Replace(Main_Pipe_Thickness, "S", "") '尺寸厚度
                .Cells(i, "E").value = Main_Pipe_Length - 100 '長度
                .Cells(i, "G").value = Mtl '材值
                .Cells(i, "H").value = 1 '數量
                .Cells(i, "I").value = MainPipeDetails.Item("WeightPerMeter") '每米重
                .Cells(i, "J").value = .Cells(i, "E").value / 1000 * .Cells(i, "I").value '單重
                .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value '重量小計
                .Cells(i, "L").value = "M"
                .Cells(i, "M").value = 1 '組數
                .Cells(i, "N").value = .Cells(i, "H").value * .Cells(i, "M").value * .Cells(i, "E").value / 1000 '長度小計 組數*數量*長度/1000
                .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
                .Cells(i, "Q").value = "素材類"
            End With
               
            'M42 導入式 :
            '重新演算
            PipeSize = Main_Pipe_Size
            PerformActionByLetter letter, PipeSize
            
            'Spacer Plate 選用
            Plate_Size = 120
            Plate_Size_b = 80
            Plate_Thickness = 6
            Weight_calculator = Plate_Size / 1000 * Plate_Size_b / 1000 * Plate_Thickness * 7.85

            i = GetNextRowInColumnB()
                  With ws
                      .Cells(i, "B").value = .Cells(i - 1, "B").value + 1 '項次
                      .Cells(i, "C").value = "Plate"
                      .Cells(i, "D").value = Plate_Thickness
                      .Cells(i, "E").value = Plate_Size
                      .Cells(i, "F").value = Plate_Size_b
                      .Cells(i, "G").value = Mtl
                      .Cells(i, "H").value = 1
                      .Cells(i, "J").value = Weight_calculator
                      .Cells(i, "K").value = Weight_calculator
                      .Cells(i, "L").value = "PC"
                      .Cells(i, "M").value = 1
                      .Cells(i, "O").value = 1
                      .Cells(i, "P").value = Weight_calculator
                      .Cells(i, "Q").value = "鋼板類"
                  End With
            


            '特殊板選用
            ' 給定定義特殊板 Fig.A = 108_Fig_A_Plate
            Select Case Fig
            Case "Fig_A"
            Plate_Size = 120
            Plate_Size_b = 100
            Plate_Thickness = 9
            Weight_calculator = Plate_Size / 1000 * Plate_Size_b / 1000 * Plate_Thickness * 7.85

            i = GetNextRowInColumnB()
                  With ws
                      .Cells(i, "B").value = .Cells(i - 1, "B").value + 1 '項次
                      .Cells(i, "C").value = "108_Fig_A_Plate"
                      .Cells(i, "D").value = Plate_Thickness
                      .Cells(i, "E").value = Plate_Size
                      .Cells(i, "F").value = Plate_Size_b
                      .Cells(i, "G").value = Mtl
                      .Cells(i, "H").value = 1
                      .Cells(i, "J").value = Weight_calculator
                      .Cells(i, "K").value = Weight_calculator
                      .Cells(i, "L").value = "PC"
                      .Cells(i, "M").value = 1
                      .Cells(i, "O").value = 1
                      .Cells(i, "P").value = Weight_calculator
                      .Cells(i, "Q").value = "鋼板類"
                  End With
             
             Case "Fig_B"
            Plate_Size = 120
            Plate_Size_b = 100
            Plate_Thickness = 9
            Weight_calculator = Plate_Size / 1000 * Plate_Size_b / 1000 * Plate_Thickness * 7.85

            i = GetNextRowInColumnB()
                  With ws
                      .Cells(i, "B").value = .Cells(i - 1, "B").value + 1 '項次
                      .Cells(i, "C").value = "108_Fig_B_Plate"
                      .Cells(i, "D").value = Plate_Thickness
                      .Cells(i, "E").value = Plate_Size
                      .Cells(i, "F").value = Plate_Size_b
                      .Cells(i, "G").value = Mtl
                      .Cells(i, "H").value = 1
                      .Cells(i, "J").value = Weight_calculator
                      .Cells(i, "K").value = Weight_calculator
                      .Cells(i, "L").value = "PC"
                      .Cells(i, "M").value = 1
                      .Cells(i, "O").value = 1
                      .Cells(i, "P").value = Weight_calculator
                      .Cells(i, "Q").value = "鋼板類"
                  End With
             
             Case "Fig_C"
            Plate_Size = 65
            Plate_Size_b = 210
            Plate_Thickness = 9
            Weight_calculator = Plate_Size / 1000 * Plate_Size_b / 1000 * Plate_Thickness * 7.85

            i = GetNextRowInColumnB()
                  With ws
                      .Cells(i, "B").value = .Cells(i - 1, "B").value + 1 '項次
                      .Cells(i, "C").value = "108_Fig_C_Plate"
                      .Cells(i, "D").value = Plate_Thickness
                      .Cells(i, "E").value = Plate_Size
                      .Cells(i, "F").value = Plate_Size_b
                      .Cells(i, "G").value = Mtl
                      .Cells(i, "H").value = 1
                      .Cells(i, "J").value = Weight_calculator
                      .Cells(i, "K").value = Weight_calculator
                      .Cells(i, "L").value = "PC"
                      .Cells(i, "M").value = 1
                      .Cells(i, "O").value = 1
                      .Cells(i, "P").value = Weight_calculator
                      .Cells(i, "Q").value = "鋼板類"
                  End With


            End Select

End Sub
