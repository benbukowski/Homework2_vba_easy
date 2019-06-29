VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock()

  
  Dim stockn As String


  Dim stockv As Double
  stockv = 0
  Dim stv As Integer
  stv = 2
  For i = 2 To 760192

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      stockn = Cells(i, 1).Value

      stockv = stockv + Cells(i, 7).Value

      Range("H" & stv).Value = stockn

      Range("I" & stv).Value = stockv

      stv = stv + 1
      
      stockv = 0

    Else

        stockv = stockv + Cells(i, 7).Value

    End If

  Next i

End Sub

