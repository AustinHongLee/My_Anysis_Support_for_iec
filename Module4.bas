Attribute VB_Name = "Module4"
Sub Type_01(ByVal fullString As String)
    Dim PartString_Type As String
    Dim PipeSize As String
    Dim letter As String
    Dim pi As Double
    
    Set ws_M42 = Worksheets("Weight_Analysis")
    Set ws_Pipe_Table = Worksheets("Pipe_Table")
    Set ws = Worksheets("Weight_Analysis")
          

    
    '��M42���
    PartString_Type = GetSecondPartOfString(fullString)
    PipeSize = Replace(PartString_Type, "B", "")
    

    letter = GetThirdPartOfString(fullString)
    letter = Right(letter, 1)
    

    
    'Main_Pipe
    Third_Length_export = Replace(GetThirdPartOfString(fullString), letter, "") * 100
        ' �B�z�D�޻P���U�ު��s��:
            
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
    
    '�D�ު��׻P�ƺު��׺t��
    
    '�D�ު��� - �q�`��SUS304
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
    
    '�ƺު��� - �q�`��C12
        Support_Pipe_Length = Third_Length_export - 100
            
            If Pipe_ThickNess_mm <> "STD.WT" Then
            Support_Pipe_Thickness = Replace(Pipe_ThickNess_mm, "SCH.", "") & "S"
            Else
            Support_Pipe_Thickness = Pipe_ThickNess_mm
            End If
    'CalculatePipeDetails Fuction ( PipeSize , thickness(**S),Length,ws)
    Set SupportPipeDetails = CalculatePipeDetails(Support_Pipe_Size, Support_Pipe_Thickness)
    
    '�T����T�ǤJ - �D�� - �ƺ� - ���O
          '�D��
    ' PipeDetails.Add PipeDiameterInch, "DiameterInch" �ޮ|MM
    ' PipeDetails.Add PipeWeightPerMeter, "WeightPerMeter" �C�̭�
    ' PipeDetails.Add TotalWeight, "TotalWeight" '���q
    ' PipeDetails.Add Length, "Length"
    '   MainPipeDetails.Item("DiameterInch")
          
          i = GetNextRowInColumnB()
            
            If Main_Pipe_Thickness <> "STD.WT" Then
            Main_Pipe_Thickness = "SCH" & Replace(Main_Pipe_Thickness, "S", "")
            Else
            Main_Pipe_Thickness = Pipe_ThickNess_mm
            End If
            
            With ws
                .Cells(i, "B").value = 1 '����
                .Cells(i, "C").value = "Pipe" '�~�W
                .Cells(i, "D").value = Main_Pipe_Size & """" & "*" & Main_Pipe_Thickness '�ؤo�p��
                .Cells(i, "E").value = Main_Pipe_Length '����
                .Cells(i, "G").value = "SUS304" '����
                .Cells(i, "H").value = 1 '�ƶq
                .Cells(i, "I").value = MainPipeDetails.Item("WeightPerMeter") '�C�̭�
                .Cells(i, "J").value = .Cells(i, "E").value / 1000 * .Cells(i, "I").value '�歫
                .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value '���q�p�p
                .Cells(i, "L").value = "M"
                .Cells(i, "M").value = 1 '�ռ�
                .Cells(i, "N").value = .Cells(i, "H").value * .Cells(i, "M").value * .Cells(i, "E").value / 1000 '���פp�p �ռ�*�ƶq*����/1000
                .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
                .Cells(i, "Q").value = "������"
            End With
          
          '�����W�Y�ƺު��פp�󵥩�0 �h���L
          If Support_Pipe_Length > 0 Then
            
          
          
          '�ƺ�
          i = GetNextRowInColumnB()
            If Support_Pipe_Thickness <> "STD.WT" Then
            Support_Pipe_Thickness = "SCH" & Replace(Support_Pipe_Thickness, "S", "")
            Else
            Support_Pipe_Thickness = Pipe_ThickNess_mm
            End If
            
            With ws
                .Cells(i, "B").value = 2 '����
                .Cells(i, "C").value = "Pipe" '�~�W
                .Cells(i, "D").value = Support_Pipe_Size & """" & "*" & Support_Pipe_Thickness '�ؤo�p��
                .Cells(i, "E").value = Support_Pipe_Length '����
                .Cells(i, "G").value = "A53Gr.B" '����
                .Cells(i, "H").value = 1 '�ƶq
                .Cells(i, "I").value = SupportPipeDetails.Item("WeightPerMeter") '�C�̭�
                .Cells(i, "J").value = .Cells(i, "E").value / 1000 * .Cells(i, "I").value '�歫
                .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value '���q�p�p
                .Cells(i, "L").value = "M"
                .Cells(i, "M").value = 1 '�ռ�
                .Cells(i, "N").value = .Cells(i, "H").value * .Cells(i, "M").value * .Cells(i, "E").value / 1000 '���פp�p �ռ�*�ƶq*����/1000
                .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
                .Cells(i, "Q").value = "������"
            End With
           End If
    
    PipeSize = Replace(Support_Pipe_Size, "'", "")
    PerformActionByLetter letter, PipeSize
End Sub
Sub Type_05(ByVal fullString As String)
    '�d�Ү榡A : 20-L50-05L
    Dim PartString_Type As String
    Dim PipeSize As String
    Dim letter As String
    Dim SectionType As String
    Dim Section_Dim As String
    Dim Total_Length As Double

    
   
    Set ws = Worksheets("Weight_Analysis")
    
    '�Ϥ��X���K�ؤo
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

    '�Ϥ��XM42����
        Support_05_Type_Choice_M42 = Right(GetThirdPartOfString(fullString), 1)
    '�Ϥ��X����"H"
         Section_Length_H = Replace(GetThirdPartOfString(fullString), Support_05_Type_Choice_M42, "") * 100
         Section_Length_L = 130
         
       '�ഫ���������n�ݨD :
        letter = Support_05_Type_Choice_M42
        PipeSize = The_Section_Size

      
      
      '�ɤJFunction addSteelSectionEntry
            SectionType = SectionType
            Section_Dim = Replace(The_Section_Size, Left(The_Section_Size, 1), "")
            Total_Length = Section_Length_H + Section_Length_L
            AddSteelSectionEntry SectionType, Section_Dim, Total_Length
            PerformActionByLetter letter, PipeSize
End Sub
Sub Type_09(ByVal fullString As String)
    '�d�Ү榡A : 09-2B-05B
    Dim PartString_Type As String
    Dim PipeSize As String
    Dim letter As String
    Dim pi As Double
    
    Set ws_M42 = Worksheets("Weight_Analysis")
    Set ws_Pipe_Table = Worksheets("Pipe_Table")
    Set ws = Worksheets("Weight_Analysis")
          

    
    '��M42���
    PartString_Type = GetSecondPartOfString(fullString)
    PipeSize = Replace(PartString_Type, "B", "")
    

    letter = GetThirdPartOfString(fullString)
    letter = Right(letter, 1)
    

    
    'Main_Pipe
    Third_Length_export = Replace(GetThirdPartOfString(fullString), letter, "") * 100
        ' �B�z�D�޻P���U�ު��s��:
            
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
    
    '�D�ު��׻P�ƺު��׺t��
    
    '�D�ު��� - �q�`��SUS304
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
    
    '�ƺު��� - �q�`��C12
        Support_Pipe_Length = Third_Length_export - 100
            
            If Pipe_ThickNess_mm <> "STD.WT" Then
            Support_Pipe_Thickness = Replace(Pipe_ThickNess_mm, "SCH.", "") & "S"
            Else
            Support_Pipe_Thickness = Pipe_ThickNess_mm
            End If
    'CalculatePipeDetails Fuction ( PipeSize , thickness(**S),Length,ws)
    Set SupportPipeDetails = CalculatePipeDetails(Support_Pipe_Size, Support_Pipe_Thickness)
    
    '�T����T�ǤJ - �D�� - �ƺ� - ���O
          '�D��
    ' PipeDetails.Add PipeDiameterInch, "DiameterInch" �ޮ|MM
    ' PipeDetails.Add PipeWeightPerMeter, "WeightPerMeter" �C�̭�
    ' PipeDetails.Add TotalWeight, "TotalWeight" '���q
    ' PipeDetails.Add Length, "Length"
    '   MainPipeDetails.Item("DiameterInch")
          
          i = GetNextRowInColumnB()
            
            If Main_Pipe_Thickness <> "STD.WT" Then
            Main_Pipe_Thickness = "SCH" & Replace(Main_Pipe_Thickness, "S", "")
            Else
            Main_Pipe_Thickness = Pipe_ThickNess_mm
            End If
            
            With ws
                .Cells(i, "B").value = 1 '����
                .Cells(i, "C").value = "Pipe" '�~�W
                .Cells(i, "D").value = Main_Pipe_Size & """" & "*" & Main_Pipe_Thickness '�ؤo�p��
                .Cells(i, "E").value = Main_Pipe_Length '����
                .Cells(i, "G").value = "SUS304" '����
                .Cells(i, "H").value = 1 '�ƶq
                .Cells(i, "I").value = MainPipeDetails.Item("WeightPerMeter") '�C�̭�
                .Cells(i, "J").value = .Cells(i, "E").value / 1000 * .Cells(i, "I").value '�歫
                .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value '���q�p�p
                .Cells(i, "L").value = "M"
                .Cells(i, "M").value = 1 '�ռ�
                .Cells(i, "N").value = .Cells(i, "H").value * .Cells(i, "M").value * .Cells(i, "E").value / 1000 '���פp�p �ռ�*�ƶq*����/1000
                .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
                .Cells(i, "Q").value = "������"
            End With
          
          '�����W�Y�ƺު��פp�󵥩�0 �h���L
          If Support_Pipe_Length > 0 Then
            
          
          
          '�ƺ�
          i = GetNextRowInColumnB()
            If Support_Pipe_Thickness <> "STD.WT" Then
            Support_Pipe_Thickness = "SCH" & Replace(Support_Pipe_Thickness, "S", "")
            Else
            Support_Pipe_Thickness = Pipe_ThickNess_mm
            End If
            
            With ws
                .Cells(i, "B").value = 2 '����
                .Cells(i, "C").value = "Pipe" '�~�W
                .Cells(i, "D").value = Support_Pipe_Size & """" & "*" & Support_Pipe_Thickness '�ؤo�p��
                .Cells(i, "E").value = Support_Pipe_Length '����
                .Cells(i, "G").value = "A53Gr.B" '����
                .Cells(i, "H").value = 1 '�ƶq
                .Cells(i, "I").value = SupportPipeDetails.Item("WeightPerMeter") '�C�̭�
                .Cells(i, "J").value = .Cells(i, "E").value / 1000 * .Cells(i, "I").value '�歫
                .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value '���q�p�p
                .Cells(i, "L").value = "M"
                .Cells(i, "M").value = 1 '�ռ�
                .Cells(i, "N").value = .Cells(i, "H").value * .Cells(i, "M").value * .Cells(i, "E").value / 1000 '���פp�p �ռ�*�ƶq*����/1000
                .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
                .Cells(i, "Q").value = "������"
            End With
           End If
    
    PipeSize = Replace(Support_Pipe_Size, "'", "")
    PerformActionByLetter letter, PipeSize
       
       '�ɤJ09-Tpye�S���ݩ� : Machine Bolt
       ' ��R�ƾ�
    i = GetNextRowInColumnB()
    With ws
        .Cells(i, "B").value = .Cells(i - 1, "B").value + 1
        .Cells(i, "C").value = "MACHINE BOLT"
        .Cells(i, "D").value = "1-5/8""""*150L"
        .Cells(i, "G").value = "A307Gr.B(�����N)"
        .Cells(i, "H").value = 1
        .Cells(i, "J").value = 20 ' ���]�C�����ꪺ��ӭ��q�O20�]�i�H�ھڹ�ڱ��p�վ�^
        .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value
        .Cells(i, "L").value = "SET"
        .Cells(i, "M").value = 1
        .Cells(i, "O").value = .Cells(i, "M").value * .Cells(i, "H").value
        .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
        .Cells(i, "Q").value = "������"
    End With


End Sub
Sub Type_14(ByVal fullString As String)
    '�d�Ү榡A : 14-2B-1005
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

    
    '���w�ޤؤo
    PartString_Type = GetSecondPartOfString(fullString)
    PipeSize = Replace(PartString_Type, "B", "")
    '���wH&L ����
    Section_Length_L = Left(GetThirdPartOfString(fullString), 2) * 100
    '�`�N �H�U����H�Ȭ��ȩw
    Pipe_Length_H_part = Right(GetThirdPartOfString(fullString), 2) * 100
    
    '�D�ު��� - �q�`��SUS304
        
        ' �p���ںޤl�ݨD����
        PipeSize = GetLookupValue(PipeSize)
        BpLength = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 5, False) 'F
        SL = Replace(Left(Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 12, False), 4), "C", "")  'N
        BTLength = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 5, False) 'F
        ' �p��ޤl�p��
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
    
    '�T����T�ǤJ - �D�� - ���c - Plate(wing) - Plate(STOPPER) - Plate(BASE PLATE) - Plate(TOP)
          '�D��
    ' PipeDetails.Add PipeDiameterInch, "DiameterInch" �ޮ|MM
    ' PipeDetails.Add PipeWeightPerMeter, "WeightPerMeter" �C�̭�
    ' PipeDetails.Add TotalWeight, "TotalWeight" '���q
    ' PipeDetails.Add Length, "Length"
    '   MainPipeDetails.Item("DiameterInch")
          
          i = GetNextRowInColumnB()
            
            If Main_Pipe_Thickness <> "STD.WT" Then
            Main_Pipe_Thickness = "SCH" & Replace(Main_Pipe_Thickness, "S", "")
            Else
            Main_Pipe_Thickness = Pipe_ThickNess_mm
            End If
            
            With ws
                .Cells(i, "B").value = 1 '����
                .Cells(i, "C").value = "Pipe" '�~�W
                .Cells(i, "D").value = Main_Pipe_Size & """" & "*" & Main_Pipe_Thickness '�ؤo�p��
                .Cells(i, "E").value = Main_Pipe_Length '����
                .Cells(i, "G").value = "SUS304" '����
                .Cells(i, "H").value = 1 '�ƶq
                .Cells(i, "I").value = MainPipeDetails.Item("WeightPerMeter") '�C�̭�
                .Cells(i, "J").value = .Cells(i, "E").value / 1000 * .Cells(i, "I").value '�歫
                .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value '���q�p�p
                .Cells(i, "L").value = "M"
                .Cells(i, "M").value = 1 '�ռ�
                .Cells(i, "N").value = .Cells(i, "H").value * .Cells(i, "M").value * .Cells(i, "E").value / 1000 '���פp�p �ռ�*�ƶq*����/1000
                .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
                .Cells(i, "Q").value = "������"
            End With
'�ɤJ���c


               The_Section_Size = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 12, False)
               SectionType = "Channel"
       '�ɤJFunction addSteelSectionEntry
            SectionType = SectionType
            Section_Dim = Replace(The_Section_Size, Left(The_Section_Size, 1), "")
            Total_Length = Section_Length_L
            AddSteelSectionEntry SectionType, Section_Dim, Total_Length

'�ɤJ14-Tpye�S���ݩ� : Plate(wing)_14Type
' ��R�ƾ�
            PipeSize = GetLookupValue(PipeSize)
            Plate_Size_a = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 9, False) 'Q
            Plate_Size_b = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 8, False) 'P
            Plate_Thickness = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 5, False) 'F
            Weight_calculator = Plate_Size_a / 1000 * Plate_Size_b / 1000 * Plate_Thickness * 7.85

            i = GetNextRowInColumnB()
                  With ws
                      .Cells(i, "B").value = .Cells(i - 1, "B").value + 1 '����
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
                      .Cells(i, "Q").value = "���O��"
                  End With

'�ɤJ14-Tpye�S���ݩ� : Plate(STOPPER)_14Type
' ��R�ƾ�
            PipeSize = GetLookupValue(PipeSize)
            Plate_Size_a = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 7, False) 'M
            Plate_Size_b = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 6, False) 'K
            Plate_Thickness = 6
            Weight_calculator = Plate_Size_a / 1000 * Plate_Size_b / 1000 * Plate_Thickness * 7.85

            i = GetNextRowInColumnB()
                  With ws
                      .Cells(i, "B").value = .Cells(i - 1, "B").value + 1 '����
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
                      .Cells(i, "Q").value = "���O��"
                  End With

'�ɤJ14-Tpye�S���ݩ� : Plate(BASE PLATE)_14Type
' ��R�ƾ�
            PipeSize = GetLookupValue(PipeSize)
            Plate_Size_a = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 2, False) 'C
            Plate_Size_b = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 2, False) 'C
            Plate_Thickness = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 5, False) 'F
            Weight_calculator = Plate_Size_a / 1000 * Plate_Size_b / 1000 * Plate_Thickness * 7.85

            i = GetNextRowInColumnB()
                  With ws
                      .Cells(i, "B").value = .Cells(i - 1, "B").value + 1 '����
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
                      .Cells(i, "Q").value = "���O��"
                  End With

'�ɤJ14-Tpye�S���ݩ� : Plate(TOP)_14Type
' ��R�ƾ�
            PipeSize = GetLookupValue(PipeSize)
            Plate_Size_a = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 11, False) 'C
            Plate_Size_b = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 11, False) 'C
            Plate_Thickness = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 5, False) 'F
            Weight_calculator = Plate_Size_a / 1000 * Plate_Size_b / 1000 * Plate_Thickness * 7.85

            i = GetNextRowInColumnB()
                  With ws
                      .Cells(i, "B").value = .Cells(i - 1, "B").value + 1 '����
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
                      .Cells(i, "Q").value = "���O��"
                  End With
'�ɤJ14-Tpye�S���ݩ� : EXP.BOLT
' ��R�ƾ�
   BoltSize = Application.WorksheetFunction.VLookup(PipeSize, Type_14_Table.Range("A:L"), 10, False) 'J
    
    With ws
        .Cells(i, "B").value = .Cells(i - 1, "B").value + 1
        .Cells(i, "C").value = "EXP.BOLT"
        .Cells(i, "D").value = "'" & BoltSize & """"
        .Cells(i, "G").value = "SUS304"
        .Cells(i, "H").value = 4
        .Cells(i, "J").value = 1 ' ���]�C�����ꪺ��ӭ��q�O1�]�i�H�ھڹ�ڱ��p�վ�^
        .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value
        .Cells(i, "L").value = "SET"
        .Cells(i, "M").value = 1
        .Cells(i, "O").value = .Cells(i, "M").value * .Cells(i, "H").value
        .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
        .Cells(i, "Q").value = "������"
    End With


End Sub
Sub Type_16(ByVal fullString As String)
    ' �o�̬O Type_16 ���{���X
    ' �z�i�H�ϥ� fullString �ѼƨӰ���һݪ��ާ@
    Set ws_Pipe_Table = Worksheets("Pipe_Table")
    Set ws = Worksheets("Weight_Analysis")



        PartString_Type = GetSecondPartOfString(fullString)
 '���ѳ��� �d�� : 16-2B-04
        
        'Second For Main Pipe Size
            PartString_Type = GetSecondPartOfString(fullString)
            PipeSize = Replace(PartString_Type, "B", "")
        'Third For H Value
            Third_Length_export = GetThirdPartOfString(fullString) * 100
            
            
        ' �B�z�D�޻P���U�ު��s��:
            
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
    
    '�D�ު��׻P�ƺު��׺t��
    
    '�D�ު��� - �q�`��SUS304
        Main_Pipe_Length = Round((PipeSize * 1.5 * 25.4) + (Main_Pipe_inch_mm / 2) + 100)
        Main_Pipe_Thickness = "40S"
        Main_Pipe_Size = PipeSize
        
      'CalculatePipeDetails Fuction ( PipeSize , thickness(**S),Length,ws)
     Set MainPipeDetails = CalculatePipeDetails(Main_Pipe_Size, "40S")
    
    '�ƺު��� - �q�`��C12
        Support_Pipe_Length = Round(Third_Length_export - (Support_Pipe_inch_mm / 2) - 100 + 300)
            
            If Pipe_ThickNess_mm <> "STD.WT" Then
            Support_Pipe_Thickness = Replace(Pipe_ThickNess_mm, "SCH.", "") & "S"
            Else
            Support_Pipe_Thickness = Pipe_ThickNess_mm
            End If
    'CalculatePipeDetails Fuction ( PipeSize , thickness(**S),Length,ws)
    Set SupportPipeDetails = CalculatePipeDetails(Support_Pipe_Size, Support_Pipe_Thickness)
    
    '�T����T�ǤJ - �D�� - �ƺ� - ���O
          '�D��
    ' PipeDetails.Add PipeDiameterInch, "DiameterInch" �ޮ|MM
    ' PipeDetails.Add PipeWeightPerMeter, "WeightPerMeter" �C�̭�
    ' PipeDetails.Add TotalWeight, "TotalWeight" '���q
    ' PipeDetails.Add Length, "Length"
    '   MainPipeDetails.Item("DiameterInch")
          
          i = GetNextRowInColumnB()
            With ws
                .Cells(i, "B").value = 1 '����
                .Cells(i, "C").value = "Pipe" '�~�W
                .Cells(i, "D").value = Main_Pipe_Size & """" & "*" & "SCH" & Replace(Main_Pipe_Thickness, "S", "") '�ؤo�p��
                .Cells(i, "E").value = Main_Pipe_Length '����
                .Cells(i, "G").value = "SUS304" '����
                .Cells(i, "H").value = 1 '�ƶq
                .Cells(i, "I").value = MainPipeDetails.Item("WeightPerMeter") '�C�̭�
                .Cells(i, "J").value = .Cells(i, "E").value / 1000 * .Cells(i, "I").value '�歫
                .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value '���q�p�p
                .Cells(i, "L").value = "M"
                .Cells(i, "M").value = 1 '�ռ�
                .Cells(i, "N").value = .Cells(i, "H").value * .Cells(i, "M").value * .Cells(i, "E").value / 1000 '���פp�p �ռ�*�ƶq*����/1000
                .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
                .Cells(i, "Q").value = "������"
            End With
          
          '�ƺ�
          i = GetNextRowInColumnB()
            With ws
                .Cells(i, "B").value = 2 '����
                .Cells(i, "C").value = "Pipe" '�~�W
                .Cells(i, "D").value = Support_Pipe_Size & """" & "*" & "SCH" & Replace(Support_Pipe_Thickness, "S", "") '�ؤo�p��
                .Cells(i, "E").value = Support_Pipe_Length '����
                .Cells(i, "G").value = "A53Gr.B" '����
                .Cells(i, "H").value = 1 '�ƶq
                .Cells(i, "I").value = SupportPipeDetails.Item("WeightPerMeter") '�C�̭�
                .Cells(i, "J").value = .Cells(i, "E").value / 1000 * .Cells(i, "I").value '�歫
                .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value '���q�p�p
                .Cells(i, "L").value = "M"
                .Cells(i, "M").value = 1 '�ռ�
                .Cells(i, "N").value = .Cells(i, "H").value * .Cells(i, "M").value * .Cells(i, "E").value / 1000 '���פp�p �ռ�*�ƶq*����/1000
                .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
                .Cells(i, "Q").value = "������"
            End With
            
            '���O
            Plate_Thickness = 6
            Weight_calculator = Plate_Size / 1000 * Plate_Size / 1000 * Plate_Thickness * 7.85

            i = GetNextRowInColumnB()
                  With ws
                      .Cells(i, "B").value = 3 '����
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
                      .Cells(i, "Q").value = "���O��"
                  End With


End Sub
Sub Type_20(ByVal fullString As String)
    '�d�Ү榡A : 20-L50-05A
    Dim SectionType As String
    Dim Section_Dim As String
    Dim Total_Length As Double

    
   
    Set ws = Worksheets("Weight_Analysis")
    
    '�Ϥ��X���K�ؤo
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

    '�Ϥ��XFig����
        Support_23_Type_Choice = Right(GetThirdPartOfString(fullString), 1)
    '�Ϥ��X����"H"
         Section_Length_H = Replace(GetThirdPartOfString(fullString), Support_23_Type_Choice, "") * 100
    
      

      
      
      '�ɤJFunction addSteelSectionEntry
            SectionType = SectionType
            Section_Dim = Replace(The_Section_Size, Left(The_Section_Size, 1), "")
            Total_Length = Section_Length_H
            AddSteelSectionEntry SectionType, Section_Dim, Total_Length

End Sub
Sub Type_21(ByVal fullString As String)
    '�d�Ү榡A : 21-L50-05A
    Dim SectionType As String
    Dim Section_Dim As String
    Dim Total_Length As Double
    '�d�Ү榡B : 21-L50-05C-07
    
   
    Set ws = Worksheets("Weight_Analysis")
    
    '�Ϥ��X���K�ؤo
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
    '�Ϥ��XFig����
        Support_21_Type_Choice = Right(GetThirdPartOfString(fullString), 1)
    '�Ϥ��X����"H"
         Section_Length_H = Replace(GetThirdPartOfString(fullString), Support_21_Type_Choice, "") * 100
    
    '�Ϥ��X����"L"
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
      

      
      
      '�ɤJFunction addSteelSectionEntry
            SectionType = "Angle"
            Section_Dim = Replace(The_Angle_Size, "L", "")
            Total_Length = Section_Length_H + Section_Length_L
            AddSteelSectionEntry SectionType, Section_Dim, Total_Length

End Sub
Sub Type_22(ByVal fullString As String)
    '�d�Ү榡A : 22-L50-05A(L)
    Dim PartString_Type As String
    Dim PipeSize As String
    Dim letter As String
    Dim SectionType As String
    Dim Section_Dim As String
    Dim Total_Length As Double
    '�d�Ү榡B : 21-L50-05(L)C-07
    
   
    Set ws = Worksheets("Weight_Analysis")
    
    '�Ϥ��X���K�ؤo
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
    '�Ϥ��XFig����
        Support_22_Type_Choice = Mid(Right(GetThirdPartOfString(fullString), 3), 1, 1)
    '�Ϥ��XM-42����
        Support_22_Type_Choice_M42 = Right(GetThirdPartOfString(fullString), 1)
    
    '�װťX Replace �޿� for ����
        Type_22_Replace_A = "(" & Support_22_Type_Choice & ")"
        Type_22_Replace_B = Support_22_Type_Choice_M42
    '�Ϥ��X����"H"
        Section_Length_H = Replace(Replace(GetThirdPartOfString(fullString), Type_22_Replace_A, ""), Type_22_Replace_B, "") * 100
        
    
    '�Ϥ��X����"L"
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
      '�ഫ���������n�ݨD :
        letter = Support_22_Type_Choice_M42
        PipeSize = The_Angle_Size
      '�ɤJFunction addSteelSectionEntry
            
            SectionType = "Angle"
            Section_Dim = Replace(The_Angle_Size, "L", "")
            Total_Length = Section_Length_H + Section_Length_L
            AddSteelSectionEntry SectionType, Section_Dim, Total_Length
            PerformActionByLetter letter, PipeSize
End Sub
Sub Type_23(ByVal fullString As String)
    '�d�Ү榡A : 23-L50-05A
    Dim SectionType As String
    Dim Section_Dim As String
    Dim Total_Length As Double
    '�d�Ү榡B : 23-L50-05C-07
    
   
    Set ws = Worksheets("Weight_Analysis")
    
    '�Ϥ��X���K�ؤo
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

    '�Ϥ��XFig����
        Support_23_Type_Choice = Right(GetThirdPartOfString(fullString), 1)
    '�Ϥ��X����"H"
         Section_Length_H = Replace(GetThirdPartOfString(fullString), Support_23_Type_Choice, "") * 100
    
    '�Ϥ��X����"L"
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
      

      
      
      '�ɤJFunction addSteelSectionEntry
            SectionType = SectionType
            Section_Dim = Replace(The_Section_Size, Left(The_Section_Size, 1), "")
            Total_Length = Section_Length_H + Section_Length_L
            AddSteelSectionEntry SectionType, Section_Dim, Total_Length

End Sub
Sub Type_24(ByVal fullString As String)
    '�d�Ү榡A : 24-L50-05
    Dim SectionType As String
    Dim Section_Dim As String
    Dim Total_Length As Double
   
    Set ws = Worksheets("Weight_Analysis")
    
    '�Ϥ��X���K�ؤo
    PartString_Type = GetSecondPartOfString(fullString)
        Select Case PartString_Type
            
            Case "L50"
               The_Angle_Size = "L50*50*6"
            Case "L75"
               The_Angle_Size = "L75*75*9"
            End Select
                  
                  ' For Angle
    '�Ϥ��X����"H"
         Section_Length_H = GetThirdPartOfString(fullString) * 100

           
      '�ɤJFunction addSteelSectionEntry
            SectionType = "Angle"
            Section_Dim = Replace(The_Angle_Size, "L", "")
            Total_Length = Section_Length_H
            AddSteelSectionEntry SectionType, Section_Dim, Total_Length

End Sub
Sub Type_25(ByVal fullString As String)
    '�d�Ү榡A : 25-L50-0505A
    Dim SectionType As String
    Dim Section_Dim As String
    Dim Total_Length As Double
    '�d�Ү榡B : 23-L50-0505C-0401
    
   
    Set ws = Worksheets("Weight_Analysis")
    
    
    
    '�Ϥ��X���K�ؤo
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
    '�Ϥ��XFig����
    Support_23_Type_Choice = Right(GetThirdPartOfString(fullString), 1)
    '�Ϥ��X"H"��
    Section_Length_H = Left(GetThirdPartOfString(fullString), 2) * 100
    '�Ϥ��X"L"��
    Section_Length_L = Replace(Right(GetThirdPartOfString(fullString), 3), Support_23_Type_Choice, "") * 100

      

      
      
      '�ɤJFunction addSteelSectionEntry
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
    
    '�d�Ү榡 : 108-1B-12E-A(S)
    'Need use GetFourthPartOfString
    '108=Type
    '1B =Denote Line Size "D"
    '12 =Denote Dimension "H" (IN 100mm)
    'E  =Denote the M42 Type
    'A  =����Fig.A & Fig.B & Fig.C Lug Plate ���ϧO
    '(S)=����Ϥ�
    
    Set ws_M42 = Worksheets("Weight_Analysis")
    Set ws_Pipe_Table = Worksheets("Pipe_Table")
    Set ws = Worksheets("Weight_Analysis")
        
        '�Ϥ��X�ئT �ŦXLine Size : "D"
            PartString_Type = GetSecondPartOfString(fullString)
            PipeSize = Replace(PartString_Type, "B", "")
          
          '�Ϥ��XM42�O����
        letter = GetThirdPartOfString(fullString)
        letter = Right(letter, 1)
        
        '�Ϥ��X"H"�Ȩí��W100
        Main_Pipe_Length = Replace(GetThirdPartOfString(fullString), letter, "") * 100
        

        
        '�Ϥ��XFig number
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
        
        
        
        '�Ϥ��X ����
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
        
        
        ' �㦳���V����˴�
        ' �p�Gdenote "D" = 3/4" H >1000 Then 1.5"_Sch80 else 1"_Sch80
        ' �p�Gdenote "D" = 1" H >1000 Then 1.5"_Sch80 else 1"_Sch80
        ' �p�Gdenote "D" = 1.5" H >1000 Then 2"_Sch40 else 1..5"_Sch80
        ' �p�Gdenote "D" = 2" = 2"_Sch40
        
        '��ڴ���X �һݥD�� �P �D�ޫp��
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
   
   '�ɤJ��: �� -> M42 -> �W�S�O�l
                  
           '�ޤl�ɤJ��
                  i = GetNextRowInColumnB()
            With ws
                .Cells(i, "B").value = 1 '����
                .Cells(i, "C").value = "Pipe" '�~�W
                .Cells(i, "D").value = Main_Pipe_Size & """" & "*" & "SCH" & Replace(Main_Pipe_Thickness, "S", "") '�ؤo�p��
                .Cells(i, "E").value = Main_Pipe_Length - 100 '����
                .Cells(i, "G").value = Mtl '����
                .Cells(i, "H").value = 1 '�ƶq
                .Cells(i, "I").value = MainPipeDetails.Item("WeightPerMeter") '�C�̭�
                .Cells(i, "J").value = .Cells(i, "E").value / 1000 * .Cells(i, "I").value '�歫
                .Cells(i, "K").value = .Cells(i, "J").value * .Cells(i, "H").value '���q�p�p
                .Cells(i, "L").value = "M"
                .Cells(i, "M").value = 1 '�ռ�
                .Cells(i, "N").value = .Cells(i, "H").value * .Cells(i, "M").value * .Cells(i, "E").value / 1000 '���פp�p �ռ�*�ƶq*����/1000
                .Cells(i, "P").value = .Cells(i, "M").value * .Cells(i, "K").value
                .Cells(i, "Q").value = "������"
            End With
               
            'M42 �ɤJ�� :
            '���s�t��
            PipeSize = Main_Pipe_Size
            PerformActionByLetter letter, PipeSize
            
            'Spacer Plate ���
            Plate_Size = 120
            Plate_Size_b = 80
            Plate_Thickness = 6
            Weight_calculator = Plate_Size / 1000 * Plate_Size_b / 1000 * Plate_Thickness * 7.85

            i = GetNextRowInColumnB()
                  With ws
                      .Cells(i, "B").value = .Cells(i - 1, "B").value + 1 '����
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
                      .Cells(i, "Q").value = "���O��"
                  End With
            


            '�S��O���
            ' ���w�w�q�S��O Fig.A = 108_Fig_A_Plate
            Select Case Fig
            Case "Fig_A"
            Plate_Size = 120
            Plate_Size_b = 100
            Plate_Thickness = 9
            Weight_calculator = Plate_Size / 1000 * Plate_Size_b / 1000 * Plate_Thickness * 7.85

            i = GetNextRowInColumnB()
                  With ws
                      .Cells(i, "B").value = .Cells(i - 1, "B").value + 1 '����
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
                      .Cells(i, "Q").value = "���O��"
                  End With
             
             Case "Fig_B"
            Plate_Size = 120
            Plate_Size_b = 100
            Plate_Thickness = 9
            Weight_calculator = Plate_Size / 1000 * Plate_Size_b / 1000 * Plate_Thickness * 7.85

            i = GetNextRowInColumnB()
                  With ws
                      .Cells(i, "B").value = .Cells(i - 1, "B").value + 1 '����
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
                      .Cells(i, "Q").value = "���O��"
                  End With
             
             Case "Fig_C"
            Plate_Size = 65
            Plate_Size_b = 210
            Plate_Thickness = 9
            Weight_calculator = Plate_Size / 1000 * Plate_Size_b / 1000 * Plate_Thickness * 7.85

            i = GetNextRowInColumnB()
                  With ws
                      .Cells(i, "B").value = .Cells(i - 1, "B").value + 1 '����
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
                      .Cells(i, "Q").value = "���O��"
                  End With


            End Select

End Sub
