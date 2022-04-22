VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Comparador()


Dim i, j, k, l, m, n, nSheets1, nSheets2 As Integer
Dim banco1, banco2 As String
Dim sh As Worksheet
Dim b1, b2 As Workbook

'adiciona um sheet com nome em branco e um valor aleatorio
Dim mySheet As String
mySheet = "BlankSheet" & Int(2 + Rnd * (300 - 2 + 1))
Sheets.Add
ActiveSheet.Name = mySheet

'Exclui todos os sheets que não são o criado anteriormente
Application.DisplayAlerts = False
For Each sh In ThisWorkbook.Worksheets
    If sh.Name <> mySheet Then
      sh.Delete
    End If
Next sh
Application.DisplayAlerts = True

'abre as janelas para selecionar os arquivos
banco1 = Application.GetOpenFilename(FileFilter:="Excel files,*.xlsx", Title:="SELECIONE O BANCO ANTIGO")
banco2 = Application.GetOpenFilename(FileFilter:="Excel files,*.xlsx", Title:="SELECIONE O BANCO NOVO")
'abre os dois arquivos
Set b1 = Workbooks.Open(banco1)
Set b2 = Workbooks.Open(banco2)


'MsgBox banco1


 Dim LR1, LR2 As Integer ' LR = Last Row
 Dim LC1, LC2 As Integer ' LC = Last Column
 
Dim LinhasExcluidas, LinhasDiferentes, LinhasNovas As Integer
LinhasExcluidas = 0
LinhasDiferentes = 0
LinhasNovas = 0

'carrega o conteudo do primeiro sheet
'Worksheets(b1.Sheets(1)).Activate


 b1.Activate
 
LR1 = Cells(Rows.Count, 1).End(xlUp).Row
' MsgBox LR1
LC1 = Cells(1, Columns.Count).End(xlToLeft).Column
' MsgBox LC1
  
 b1.Sheets(1).Copy Before:=ThisWorkbook.Sheets(1)
 'renomeia o primeiro sheet para o nome do primeiro arquivo aberto
 ActiveSheet.Name = b1.Name
 With ActiveWindow
    .SplitColumn = 0
    .SplitRow = 1
End With
ActiveWindow.FreezePanes = True
'Workbooks(b1).Sheets(b1.Sheets(1).Name).Copy Before:=Workbooks(ThisWorkbook).Sheets(1)
'Sheets(1).Range(LC, LR).Copy
'Worksheets(1).Copy 'Before:=Workbooks(banco1).Sheets(1)
 ' ActiveSheet.Copy
 ' Sheets(thisworksheet).Activate
' Range("A1").Select
'thisworksheet.Paste
b2.Activate

LR2 = Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox LR2
LC2 = Cells(1, Columns.Count).End(xlToLeft).Column
' MsgBox LC2
  
 b2.Sheets(1).Copy After:=ThisWorkbook.Sheets(1)
 ActiveSheet.Name = b2.Name
 Dim r As Range
 Set r = Worksheets(2).Range(Cells(1, 1), Cells(LR2, 1))
r.RowHeight = 15

 With ActiveWindow
    .SplitColumn = 0
    .SplitRow = 1
End With

ActiveWindow.FreezePanes = True
'carrega o conteudo do segundo sheet

'Worksheets(b1.Name).Activate

'Exclui todos os sheets que não são os dois copiados
Application.DisplayAlerts = False
For Each sh In ThisWorkbook.Worksheets
    If sh.Name <> b1.Name And sh.Name <> b2.Name Then
      sh.Delete
    End If
Next sh

b1.Close True
b2.Close True


Application.DisplayAlerts = True
Dim PNTNAM1, SUBNAM1, PNTNAM2, SUBNAM2 As String
Sheets.Add After:=Sheets(Sheets.Count)
ActiveSheet.Name = "RELATÓRIO"


For i = 1 To LC2 Step 1
    Sheets(3).Cells(1, i) = Sheets(2).Cells(1, i)
    Sheets(3).Cells(1, i).Interior.ColorIndex = 15
    Sheets(3).Cells(1, i).Font.Bold = True
    Sheets(3).Cells(1, i).HorizontalAlignment = xlCenter
    Sheets(3).Cells(1, i).BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThin
Next i
With ActiveWindow
    .SplitColumn = 0
    .SplitRow = 1
End With
ActiveWindow.FreezePanes = True


Dim excluidas() As Integer
ReDim excluidas(LR1)

Dim discrepantes() As Integer
ReDim discrepantes(LR1)

Dim novas() As Integer
ReDim novas(LR2)


For i = 1 To LR1 Step 1
 excluidas(i) = 0
 discrepantes(i) = 0
Next i

For i = 1 To LR2 Step 1
 novas(i) = 1
Next i
m = 0
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'PARTE DA COMPARAÇÃO
For i = 2 To LR2 Step 1

SUBNAM2 = ThisWorkbook.Sheets(2).Cells(i, 3).Text
PNTNAM2 = ThisWorkbook.Sheets(2).Cells(i, 4).Text

Dim TmpCampo1, TmpCampo2 As String

    For j = 2 To LR1 Step 1
        SUBNAM1 = ThisWorkbook.Sheets(1).Cells(j, 3).Text
        PNTNAM1 = ThisWorkbook.Sheets(1).Cells(j, 4).Text

        If StrComp(SUBNAM1, SUBNAM2, vbTextCompare) = 0 And StrComp(PNTNAM1, PNTNAM2, vbTextCompare) = 0 Then
         Sheets(2).Cells(i, 4).Interior.Color = RGB(254, 255, 217)
         Sheets(2).Cells(i, 3).Interior.Color = RGB(254, 255, 217)
         Sheets(1).Cells(j, 4).Interior.Color = RGB(254, 255, 217)
         Sheets(1).Cells(j, 3).Interior.Color = RGB(254, 255, 217)
          novas(i) = 0
          m = m + 1
        excluidas(j) = 1 'marcado como usado
         For k = 1 To LC1 Step 1
            TmpCampo1 = ThisWorkbook.Sheets(2).Cells(i, k).Text
            TmpCampo2 = ThisWorkbook.Sheets(1).Cells(j, k).Text
            If StrComp(TmpCampo1, TmpCampo2, vbTextCompare) = 0 Then
                Sheets(2).Cells(i, k).Interior.Color = RGB(254, 255, 217)
                Sheets(1).Cells(j, k).Interior.Color = RGB(254, 255, 217)
            Else
                Sheets(2).Cells(i, k).Interior.Color = RGB(255, 125, 127)
                Sheets(1).Cells(j, k).Interior.Color = RGB(255, 125, 127)
            discrepantes(j) = i
            End If
         Next k
        Exit For
        End If
    Next j
Next i



'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'PARTE DO RELATÓRIO

'--------------------------------------
'FAZ A PARTE DE LINHAS EXCLUIDAS
'ThisWorkbook.Sheet(3).Activate

Sheets(3).Cells(3, 1).Value = "LINHAS EXCLUÍDAS (Presentes no banco 1 e não no 2):"
Sheets(3).Range("A3:EF3").Font.Size = 20
Sheets(3).Range("A3:EF3").Merge

k = 5
Sheets(3).Cells(k, 1).Value = Sheets(1).Name + ":"

For i = 2 To LR1 Step 1
    If excluidas(i) = 0 Then
        For j = 1 To LC1 Step 1
            Sheets(3).Cells(1 + k, j).Value = Sheets(1).Cells(i, j).Text
            Sheets(3).Cells(1 + k, 3).Interior.Color = RGB(255, 125, 127)
            Sheets(3).Cells(1 + k, 4).Interior.Color = RGB(255, 125, 127)
            Sheets(1).Cells(i, 3).Interior.Color = RGB(255, 125, 127)
            Sheets(1).Cells(i, 4).Interior.Color = RGB(255, 125, 127)
            Sheets(3).Cells(1 + k, j).BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThin
        Next j
        k = k + 1
    End If

Next i

  k = k + 3
'--------------------------------------
'FAZ A PARTE DAS LINHAS COM VALORES DIFERENTES
Sheets(3).Cells(k + 1, 1).Value = "LINHAS DISCREPANTES:"

Set r = Worksheets(3).Range(Cells(1 + k, 1), Cells(1 + k, 136))
r.Font.Size = 20
r.Merge
k = k + 1
For i = 2 To LR1 Step 1

    If discrepantes(i) <> 0 Then
    Sheets(3).Cells(k + 2, 1).Value = Sheets(1).Name + ":"
    Sheets(3).Cells(k + 4, 1).Value = Sheets(2).Name + ":"
    For j = 1 To LC1 Step 1
            Sheets(3).Cells(3 + k, j).Value = Sheets(1).Cells(i, j).Text
            Sheets(3).Cells(3 + k + 2, j).Value = Sheets(2).Cells(discrepantes(i), j).Text
            
            'pinta de vermelho caso sejam diferentes
            If StrComp(Sheets(3).Cells(3 + k, j).Text, Sheets(3).Cells(3 + k + 2, j).Text, vbTextCompare) <> 0 Then
                Sheets(3).Cells(3 + k, j).Interior.Color = RGB(255, 125, 127)
                Sheets(3).Cells(3 + k + 2, j).Interior.Color = RGB(255, 125, 127)
            End If
            
            'borda
            Sheets(3).Cells(3 + k, j).BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThin
           Sheets(3).Cells(3 + k + 2, j).BorderAround _
           LineStyle:=xlContinuous, _
           Weight:=xlThin
           
        Next j
        k = k + 5
    End If
    
Next i
k = k + 3
'--------------------------------------
'FAZ A PARTE DAS LINHAS NOVAS
Sheets(3).Cells(k + 1, 1).Value = "LINHAS NOVAS (Presentes somente no banco 2):"

Set r = Worksheets(3).Range(Cells(1 + k, 1), Cells(1 + k, 136))
r.Font.Size = 20
r.Merge
k = k + 3
Sheets(3).Cells(k, 1).Value = Sheets(1).Name + ":"
k = k + 1
For i = 2 To LR2 Step 1
    If novas(i) = 1 Then
        For j = 1 To LC1 Step 1
            Sheets(3).Cells(k, j).Value = Sheets(2).Cells(i, j).Text
            Sheets(3).Cells(k, 3).Interior.Color = RGB(201, 254, 201)
            Sheets(3).Cells(k, 4).Interior.Color = RGB(201, 254, 201)
            Sheets(2).Cells(i, 3).Interior.Color = RGB(201, 254, 201)
            Sheets(2).Cells(i, 4).Interior.Color = RGB(201, 254, 201)
            Sheets(3).Cells(k, j).BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThin
        Next j
        k = k + 1
    End If
Next i
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Worksheets(1).Columns("A:EF").AutoFit
Worksheets(2).Columns("A:EF").AutoFit
Worksheets(3).Columns("A:EF").AutoFit

ActiveWindow.DisplayGridlines = False

End Sub



