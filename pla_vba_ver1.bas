Attribute VB_Name = "Module1"
'A beginners version of pla training function set built by excel vba ver.1
'For binary classification only
'input data requirement:
'(1)Feature variable#:2 (i.e. x1,x2)
'(2)class labled with 1 or -1
'(3)Warning!!! if data is non-linear seperable, endless loop occurs
' By Bryan Ni(K.H.Ni) 2017/03/20

Function perceptronOutput(x1, x2, w As Variant)
    'w:array(from w0 to w2) x1,x2:single value data
    perceptronOutput = w(1) + w(2) * x1 + w(3) * x2
    
End Function
Private Function perceptronOutput_private(x1, x2, w)
     'w:array(from w0 to w2) x1,x2:single value data
    perceptronOutput_private = w(0) + w(1) * x1 + w(2) * x2
    
End Function

Private Sub updateWeight(x1, x2, ByRef w, y)
    
    w(0) = w(0) + y
    w(1) = w(1) + y * x1
    w(2) = w(2) + y * x2

End Sub
'core part of pla_training
Private Function pla_training_2d(x1 As Range, x2 As Range, class_lable As Range, initialWeight As Double)
    Dim w(2) As Double
    Dim dataSize As Integer
    w(0) = initialWeight
    w(1) = initialWeight
    w(2) = initialWeight

    dataSize = class_lable.Count
        
    Dim i As Integer
    Dim isTermination As Boolean
    isTermination = False
          
        While (isTermination = False)
           isTermination = True
           For i = 1 To dataSize
                If Math.Sgn(perceptronOutput_private(x1(i), x2(i), w)) <> Math.Sgn(class_lable(i)) Then
                     updateWeight x1(i), x2(i), w, Math.Sgn(class_lable(i))
                     isTermination = False
                End If
            Next i
        Wend
    pla_training_2d = w
End Function
Function pla_training_getWeightString(x1 As Range, x2 As Range, class_lable As Range)
    Dim weight_string As String
    Dim i As Integer
    Dim j As Integer
    Dim w
    weight_string = ""
    w = pla_training_2d(x1, x2, class_lable, 0)
      
    For i = 0 To UBound(w)
        weight_string = weight_string & w(i)
        If (i <> UBound(w)) Then
        weight_string = weight_string & ","
       
        End If
    Next i
    pla_training_getWeightString = weight_string
End Function
Function pla_training_getWeightVal(x1 As Range, x2 As Range, class_lable As Range, weightIndex As Integer)
    Dim weight_string As String
    Dim w
    w = pla_training_2d(x1, x2, class_lable, 0)
    pla_training_getWeightVal = w(weightIndex)
End Function
'initialization of weight value by the param rndNum
Function pla_training_getVal_rndPrm(x1 As Range, x2 As Range, class_lable As Range, weightIndex As Integer, rndNum As Double)
    Dim w
    Dim rndSeed As Double
    rndSeed = rndNum
    w = pla_training_2d(x1, x2, class_lable, rndNum)
    pla_training_getVal_rndPrm = w(weightIndex)
End Function


