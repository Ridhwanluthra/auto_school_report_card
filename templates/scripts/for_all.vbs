Sub edit_report_cards_new_subject_descriptions()
'
' edit_report_cards_new_subject_descriptions Macro
'

'
    Range("22:22,21:21").Select
    Selection.Delete Shift:=xlUp
    Range("D21").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-1]C,R[-2]C,R[-3]C,R[-4]C,R[-5]C)"
    Range("D21").Select
    Selection.AutoFill Destination:=Range("D21:E21"), Type:=xlFillDefault
    Rows("29:29").Select
    Selection.Delete Shift:=xlUp
    Range("D29").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-3]C,R[-4]C,R[-5]C)"
    Range("D29").Select
    Selection.AutoFill Destination:=Range("D29:E29"), Type:=xlFillDefault
    
    Rows("33:33").Select
    Selection.Delete Shift:=xlUp
    Range("D36").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-1]C,R[-2]C,R[-3]C,R[-4]C)"
    Range("D36").Select
    Selection.AutoFill Destination:=Range("D36:E36"), Type:=xlFillDefault

    Range("43:43,39:39").Select
    Selection.Delete Shift:=xlUp
    Range("A39").Select
    ActiveCell.FormulaR1C1 = "EVS"
    Range("D42").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-1]C,R[-2]C,R[-3]C)"
    Range("D42").Select
    Selection.AutoFill Destination:=Range("D42:E42"), Type:=xlFillDefault
    
    Range("B16").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[-13]C[2]"
    Range("B17").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[-14]C[5]"
    Range("B18").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[-15]C[8]"
    Range("B16:B18").Select
    Selection.AutoFill Destination:=Range("B16:B20"), Type:=xlFillDefault

    Range("B19").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[-16]C[11]"
    Range("B20").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[-17]C[14]"
    Range("B24").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[34]C[2]"
    Range("B25").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[33]C[5]"
    Range("B26").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[32]C[8]"
    Range("B27").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[31]C[11]"
    Range("B28").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[30]C[14]"
    Range("B32").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[81]C[2]"
    Range("B33").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[80]C[11]"
    Range("B34").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[79]C[14]"
    Range("B35").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[78]C[20]"
    Range("B39").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[129]C[5]"
    Range("B40").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[128]C[8]"
    Range("B41").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[127]C[11]"
    Range("C18").Select
    ActiveCell.FormulaR1C1 = "10"
    Range("C26").Select
    ActiveCell.FormulaR1C1 = "10"
    Range("C32").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("C34").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("C39").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("C40").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("C39").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("C40").Select
    ActiveCell.FormulaR1C1 = "5"
    Rows("48:48").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("48:48,49:49").Select
    
    Selection.Copy
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Rows("48:48").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A39:G46").Select
    Selection.Copy
    Range("A45").Select
    ActiveSheet.Paste
    
    Range("D51").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-1]C,R[-7]C,R[-14]C,R[-22]C,R[-28]C)"
    Range("D51").Select
    Selection.AutoFill Destination:=Range("D51:E51"), Type:=xlFillDefault
    Range("A45").Select
    ActiveCell.FormulaR1C1 = "Social"
    Range("A46").Select
    ActiveCell.FormulaR1C1 = "Responsibility"
    Range("B45").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[178]C[5]"
    Range("B46").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[177]C[8]"
    Range("B47").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[176]C[11]"
    
    Range("A45:G52").Select
    
    Selection.Copy
    
    Rows("56:56").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    
    Range("A58:H69").Select
    
    Selection.ClearContents
    Rows("59:59").Select
    
    Selection.Delete Shift:=xlUp
    Rows("58:68").Select
    Selection.Delete Shift:=xlUp
    
    Rows("54:54").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Rows("55:55").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Rows("56:56").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Range("A45:G52").Select
    Selection.Copy
    Range("A51").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range("A52").Select
    ActiveCell.FormulaR1C1 = ""
    Range("A51").Select
    ActiveCell.FormulaR1C1 = "Art and craft"
    Range("A51").Select
    ActiveCell.FormulaR1C1 = "Art and Craft"
    Rows("61:61").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Rows("60:60").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Rows("61:61").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Rows("62:62").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A51:G58").Select
    Selection.Copy
    Range("A57").Select
    ActiveSheet.Paste
    Range("A57").Select
    ActiveCell.FormulaR1C1 = "Computer"
    Range("D63").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-1]C,R[-7]C,R[-14]C,R[-22]C,R[-26]C)"
    Range("D63").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-1]C,R[-7]C,R[-14]C,R[-22]C,R[-26]C)"
    Range("D63").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(R[-7]C,R[-1]C,R[-13]C,R[-19]C,R[-25]C,R[-32]C,R[-40]C)"
    Range("D63").Select
    Selection.AutoFill Destination:=Range("D63:E63"), Type:=xlFillDefault
    Range("A51").Select
    ActiveCell.FormulaR1C1 = "Computer"
    Range("A57").Select
    ActiveCell.FormulaR1C1 = "Art & Craft"
    Range("B51").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[282]C[14]"
    Range("B52").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[281]C[17]"
    Range("B53").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[280]C[20]"
    Rows("58:58").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Rows("59:59").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("B57").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[331]C[8]"
    Range("B58").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[330]C[11]"
    Range("B59").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[329]C[14]"
    Range("B60").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[328]C[17]"
    Range("B61").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[327]C[20]"
    Range("C57").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("C58").Select
    ActiveCell.FormulaR1C1 = "10"
    Range("C59").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("C60").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("C61").Select
    ActiveCell.FormulaR1C1 = "5"
    Range("D57").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[333]C[8]"
    Range("E57").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[333]C[35]"
    Range("D58").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[332]C[11]"
    Range("E58").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[332]C[38]"
    Range("D59").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[331]C[14]"
    Range("E59").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[331]C[41]"
    Range("D60").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[330]C[17]"
    Range("E60").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[330]C[44]"
    Range("D61").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[329]C[20]"
    Range("E61").Select
    ActiveCell.FormulaR1C1 = "=+subject_eval!R[329]C[47]"
    Range("F57").Select
    Selection.AutoFill Destination:=Range("F57:F59"), Type:=xlFillDefault
    Range("G57").Select
    Selection.AutoFill Destination:=Range("G57:G59"), Type:=xlFillDefault
    Range("C65").Select
    ActiveCell.FormulaR1C1 = _
        "=SUM(R[-1]C,R[-9]C,R[-15]C,R[-21]C,R[-27]C,R[-34]C,R[-42]C)"
    Range("C51").Select
    ActiveCell.FormulaR1C1 = "10"
    Range("C52").Select
    ActiveCell.FormulaR1C1 = "10"
End Sub

