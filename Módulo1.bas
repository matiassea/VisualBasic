Attribute VB_Name = "M�dulo1"
'https://wellsr.com/vba/excel/vba-variable-scope/
'http://www.excelhowto.com/5-ways-to-get-unique-values-in-excel/
'colores
'http://dmcritchie.mvps.org/excel/colors.htm
'http://www.databison.com/excel-vba-for-do-while-and-do-until-loop/
Option Explicit
Dim Ultimate_Row As Integer
Dim Ultimate_Column As Integer
Sub mejorado()
'el problema es que no disgrega el 91046 o 38902 debido a que no tiene el numero 1
'http://www.databison.com/vba-if-function-using-if-else-elseif-if-then-in-vba-code/
'http://www.excel-easy.com/vba/examples/logical-operators.html
Dim descripcion As Integer
Dim Num As Integer
Dim i As Integer
Dim Mensaje As Integer
Dim c As Integer

Worksheets("Base").Activate


Mensaje = MsgBox("Debe tener actualizada y procesada la planilla PO" & vbCrLf & "Ya que es necesario conocer cuantas PO se realizaron ayer" & vbCrLf & "Debe tener actualizada y procesada la planilla ayer" & vbCrLf & "Ya que es necesario saber las asignaciones de ayer ", vbOKCancel)
    If Mensaje = vbCancel Then
        Exit Sub
    End If
Range("A6:B10000").Clear
Range("AX5:CI10000").Clear
Ultimate_Row = Worksheets("Base").Range("D" & Rows.Count).End(xlUp).Row 'conteo de columna
'Ultimate_Column = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
Ultimate_Column = 49

Dim Alerta2 As Integer
Alerta2 = MsgBox("Tama�o del conjunto de datos, siempre debe tener 49 columnas" & vbCrLf & "Tama�o del conjunto de datos es " & Ultimate_Column & " columnas x " & Ultimate_Row & " filas", vbOKOnly)


'revision de columna Codigo Subcategoria, columna (G),l esto debido a que en ocaciones no aparece
'el codigo SXXXXXX aparece Sin Info
Dim codigo_subcategoria As Integer
Dim Alerta As Integer
Dim Ans As Integer
'ejemplo https://www.mrexcel.com/forum/excel-questions/802933-difference-between-msgbox-cancel-red-cross-close-button.html
For codigo_subcategoria = 6 To Ultimate_Row
    If (InStr(Cells(codigo_subcategoria, 7).Text, "Sin") Or InStr(Cells(codigo_subcategoria, 7).Text, "SIN")) Or InStr(Cells(codigo_subcategoria, 7).Text, "sin") Then
    Alerta = MsgBox("Favor revisar codigos de subcategoria debido a que hay una linea sin codigo", vbOKOnly, "Alerta!!!")
        If Alerta = vbOK Then
            Exit Sub
        End If
    End If
Next codigo_subcategoria


'revision de columna Codigo Id Articulo (columna 23) y Itm Id Vndr (columna 28) esto debido a que en ocaciones no aparece
'Sin Informaci�n, en ocaciones aparece vacio o Sin Info
Dim codigo_catalogo_contrato As Integer
Dim Alerta12 As Integer
Dim Ans2 As Integer
'ejemplo https://www.mrexcel.com/forum/excel-questions/802933-difference-between-msgbox-cancel-red-cross-close-button.html
For codigo_catalogo_contrato = 6 To Ultimate_Row
    If Len(Cells(codigo_catalogo_contrato, 23)) < 2 Or Len(Cells(codigo_catalogo_contrato, 28)) < 2 Then
    Alerta12 = MsgBox("Favor revisar codigos de Itm Id Vndr o Codigo Id Articulo debido a que hay una linea erronea", vbOKOnly, "Alerta!!!")
        If Alerta12 = vbOK Then
            Exit Sub
        End If
    End If
Next codigo_catalogo_contrato



'Para colocar encabezado
Cells(5, Ultimate_Column + 1).Value = "Taxonomia"
            Cells(5, Ultimate_Column + 1).Font.Bold = True 'negrita
Cells(5, Ultimate_Column + 2).Value = "Req ID de linea 1"
            Cells(5, Ultimate_Column + 2).Font.Bold = True 'negrita
Cells(5, Ultimate_Column + 3).Value = "Tipo de compra"
            Cells(5, Ultimate_Column + 3).Font.Bold = True 'negrita
Cells(5, Ultimate_Column + 4).Value = "Area de compra"
            Cells(5, Ultimate_Column + 4).Font.Bold = True 'negrita
Cells(5, Ultimate_Column + 7).Value = "Clasificacion de compras"
            Cells(5, Ultimate_Column + 7).Font.Bold = True 'negrita


'Encabezado de columna 1 y 2
Cells(5, 2).Value = "Und. Negocio"
            Cells(5, 2).Font.Bold = True 'negrita
Cells(5, 1).Value = "Llave"
            Cells(5, 1).Font.Bold = True 'negrita


'Unidad de negocios
'N�46 + 2 = Unidadnegocios
For i = 6 To Ultimate_Row
        If Cells(i, 48).Value = "Instituto Profesional AIEP S.A." Then
        Cells(i, 2) = "CHL04"
        ElseIf Cells(i, 48).Value = "UNAB" Then
        Cells(i, 2) = "CHL01"
        ElseIf Cells(i, 48).Value = "Universidad Privada del Norte" Then
        Cells(i, 2) = "PER03"
        ElseIf Cells(i, 48).Value = "Univ. De Vi�a del Mar Chile OP" Then
        Cells(i, 2) = "CHL32"
        ElseIf Cells(i, 48).Value = "Universidad Per� Ciencias Aplicadas" Then
        Cells(i, 2) = "PER05"
        ElseIf Cells(i, 48).Value = "UDLA Chile" Then
        Cells(i, 2) = "CHL02"
        ElseIf Cells(i, 48).Value = "Cibertec" Then
        Cells(i, 2) = "PER06"
        ElseIf Cells(i, 48).Value = "IEDE Chile" Then
        Cells(i, 2) = "CHL05"
        ElseIf Cells(i, 48).Value = "Inmobiliaria Educ SPA (IESA)" Then
        Cells(i, 2) = "CHL18"
        ElseIf Cells(i, 48).Value = "Laureate Chile II SPA" Then
        Cells(i, 2) = "CHL25"
        ElseIf Cells(i, 48).Value = "Servicios Andinos" Then
        Cells(i, 2) = "CHL28"
        ElseIf Cells(i, 48).Value = "Immob Inversiones SanGenarosDos" Then
        Cells(i, 2) = "CHL31"
        ElseIf Cells(i, 48).Value = "Servicios Profesionales Andr�s Bello" Then
        Cells(i, 2) = "CHL08"
        End If
Next i

'Llave = N�1, que se crea hasta la ultima descripcion
'N�12 + 2  = Desc Articulo
'ReqId = N�40 + 2
'Und. Negocio = N�2

For i = 6 To Ultimate_Row
    If Len(Cells(i, 14)) > 0 Then
    Cells(i, 1).Value = Cells(i, 2) & Cells(i, 42)
    End If
Next i


Range("AA6:AA10000").Clear
Cells(5, 27).Value = "Politica"
            Cells(5, 27).Font.Bold = True 'negrita


'llama a sub programa llamado revision_primera
Call revision_primera

'Linea (30) = 1 , Key (1) se pega en Req ID de linea 1 (Cells(i, Ultimate_Column + 2))
'Columna Concatenado, si es la linea 1, se copia y pega el Req ID
'N�28 + 2 = Linea
'N�40 + 2 = ReqId
For i = 6 To Ultimate_Row
    If Cells(i, 30).Value = 1 Then
        Cells(i, Ultimate_Column + 2) = Cells(i, 1)
    End If
Next i


'https://excelmacromastery.com/vba-vlookup/
'Cells(i, Ultimate_Column + 2).Find(What:=Cells(i, Ultimate_Column + 2), LookIn:=xlValues )
'https://msdn.microsoft.com/en-us/VBA/Excel-VBA/articles/range-find-method-excel
'Hoja1.Range("BC6:BE5000").Cells(i, Ultimate_Column + 2).Find(What:=Cells(i, Ultimate_Column + 2), LookIn:=xlValues)
'After:=celda Le podemos indicar a partir de que celda queremos empezar a buscar
'lookIn:=donde buscar Con dos opciones: buscar en valores (xlValues) y buscar en formulas (xlFormula)
'LookAt:=como buscar con xlWhole para b�squedas de palabra exacta, o xlPart para b�squedas con parte de la palabra
'SearchOrder:= orden de busqueda para indicarle si queremos realizar la busqueda por filas(xlRows) o por columnas(xlColumns)
'SearchDirection:= direccion de la busqueda Pensado para continuar con la busqueda con dos posibles opciones: xlNext � Continuar con siguiente, xlPrevious-Continuar con anterior
'MatchCase:=true/false Para indicar si ha de detectar may�sculas y min�sculas=largo(
'Cells(i, Ultimate_Column + 2) = Cells.Find( Cells(i, Ultimate_Column + 2),

'llama a sub programa llamado revision_segunda, que considera las solicitudes sin primera linea
Call revision_segunda


'Ciclo para identificar el "Tipo Compra" = Sourcing + Contrato + Catalogo
For c = 6 To Ultimate_Row
'N�27 + 2 = Itm Id Vndr
'N�21 + 2 = Id Articulo
'N�3 + 2 = Categoriadearticulos
   Select Case True
    
    Case Cells(c, 23).Value = "Sin Informaci�n" And Cells(c, 28).Value = "Sin Informaci�n" And Cells(c, 56).Value = "Sourcing"  'Len(Cells(c, 56).Text) = 0
            Cells(c, Ultimate_Column + 3) = "Sourcing"
            Cells(c, Ultimate_Column + 4) = "Compra"
            Cells(c, Ultimate_Column + 3).Font.Bold = True 'negrita
            Cells(c, Ultimate_Column + 3).Font.ColorIndex = 4 'pintar en verde
            Cells(c, Ultimate_Column + 7) = "Sourcing"
            
    Case Cells(c, 23).Value = "Sin Informaci�n" And Cells(c, 28).Value = "Sin Informaci�n" And Cells(c, 56).Value = "Negociado por SASPA"   'Len(Cells(c, 56).Text) = 0
            Cells(c, Ultimate_Column + 3) = "Sourcing"
            Cells(c, Ultimate_Column + 4) = "Compra"
            Cells(c, Ultimate_Column + 3).Font.Bold = True 'negrita
            Cells(c, Ultimate_Column + 3).Font.ColorIndex = 4 'pintar en verde
            Cells(c, Ultimate_Column + 7) = "Negociado por SASPA"
                       
    Case InStr(Cells(c, 23).Value, "CNTR") > 0 Or InStr(Cells(c, 28).Value, "CNTR") > 0 'Len(Cells(c, 56).Text) = 0
            Cells(c, Ultimate_Column + 3) = "Contrato"
            Cells(c, Ultimate_Column + 4) = "Gestion PO"
            Cells(c, Ultimate_Column + 3).Font.Bold = True 'negrita
            Cells(c, Ultimate_Column + 3).Font.ColorIndex = 17 'pintar en azulado
            Cells(c, Ultimate_Column + 7) = "Contrato"
            
    Case InStr(Cells(c, 23).Value, "PER") > 0 Or InStr(Cells(c, 28).Value, "PER") > 0 'Len(Cells(c, 56).Text) = 0
            Cells(c, Ultimate_Column + 3) = "Catalogo"
            Cells(c, Ultimate_Column + 4) = "Gestion PO"
            Cells(c, Ultimate_Column + 3).Font.Bold = True 'negrita
            Cells(c, Ultimate_Column + 3).Font.ColorIndex = 33 'pintar en calipto
            Cells(c, Ultimate_Column + 7) = "Catalogo"
    
    Case Cells(c, 23).Value <> "Sin Informaci�n" And Cells(c, 28).Value = "Sin Informaci�n" 'Len(Cells(c, 56).Text) = 0
            Cells(c, Ultimate_Column + 3) = "Catalogo"
            Cells(c, Ultimate_Column + 4) = "Gestion PO"
            Cells(c, Ultimate_Column + 3).Font.Bold = True 'negrita
            Cells(c, Ultimate_Column + 3).Font.ColorIndex = 33 'pintar en calipto
            Cells(c, Ultimate_Column + 7) = "Catalogo"
    
    Case Cells(c, 23).Value = "Sin Informaci�n" And Cells(c, 28).Value <> "Sin Informaci�n"  'Len(Cells(c, 56).Text) = 0
            Cells(c, Ultimate_Column + 3) = "Catalogo"
            Cells(c, Ultimate_Column + 4) = "Gestion PO"
            Cells(c, Ultimate_Column + 3).Font.Bold = True 'negrita
            Cells(c, Ultimate_Column + 3).Font.ColorIndex = 33 'pintar en calipto
            Cells(c, Ultimate_Column + 7) = "Catalogo"
            
    Case Cells(c, 23).Value <> "Sin Informaci�n" And Cells(c, 28).Value <> "Sin Informaci�n"  'Len(Cells(c, 56).Text) = 0
            Cells(c, Ultimate_Column + 3) = "Catalogo"
            Cells(c, Ultimate_Column + 4) = "Gestion PO"
            Cells(c, Ultimate_Column + 3).Font.Bold = True 'negrita
            Cells(c, Ultimate_Column + 3).Font.ColorIndex = 33 'pintar en calipto
            Cells(c, Ultimate_Column + 7) = "Catalogo"
    End Select

Next c

'--------------------------------------------------------------------------------------------------------------------
'Asignando "Politica", en "Area de Compra", "Tipo de compra", "Clasificacion de compras", Key Politica (27)
'Si en columna politica (27) no hay nada, busca la key de politica y rescata cada area de compra
For i = 6 To Ultimate_Row
    If IsError(Application.VLookup(Cells(i, 1), Range("AA6:BA10000"), 27, False)) = False Then
        Cells(i, Ultimate_Column + 4) = Application.VLookup(Cells(i, 1), Range("AA6:BA10000"), 27, False)
        Cells(i, Ultimate_Column + 3) = Application.VLookup(Cells(i, 1), Range("AA6:AZ10000"), 26, False)
        Cells(i, Ultimate_Column + 7) = Application.VLookup(Cells(i, 1), Range("AA6:BD10000"), 30, False)
    End If
Next i
'--------------------------------------------------------------------------------------------------------------------
'Asignando "Area de Compra" Compra o Gestion PO , cuando tienen "Area de Compra" con error
'https://www.exceltrick.com/formulas_macros/vlookup-in-vba/
'For i = 6 To Ultimate_Row
'N�28 + 2 = Linea
'N�40 + 2 = ReqId
'AN = Req Id
'Ax = Ultimate_Column + 3
    'If Len(Cells(i, Ultimate_Column + 2)) = 0 Then
        'Cells(i, Ultimate_Column + 4) = Application.VLookup(Cells(i, 1), Range("z6:BA10000"), 3, False)
    'End If
'Next i
'--------------------------------------------------------------------------------------------------------------------
'Completar "Tipo de compra", para las lineas sin "Tipo de compra", por lo comun son compras por Politica
'Adendum, Negociado por SASPA, Negociado por COE, Regularizacion, Comprometida, Boleta Honorario,Gobierno, Emergencia
Dim k As Integer
For k = 6 To Ultimate_Row
    If Len(Cells(k, 14)) > 0 And Len(Cells(k, 52)) = 0 Then
    Select Case True
    Case Cells(k, Ultimate_Column + 7).Value <> "Sourcing" Or Cells(k, Ultimate_Column + 7).Value <> "Catalogo" Or Cells(k, Ultimate_Column + 7).Value <> "Contrato"
            Cells(k, Ultimate_Column + 3) = "Politica"
            Cells(k, Ultimate_Column + 3).Font.Bold = True 'negrita
            Cells(k, Ultimate_Column + 3).Font.ColorIndex = 4 'pintar en verde
    End Select
    End If
Next k
'--------------------------------------------------------------------------------------------------------------------
'Asignando la politica de compra, en caso de no detectar nada, quiere decir que no hay politica, por lo que entra como Sourcing
'For i = 6 To Ultimate_Row
'si Cells(Ultimate_Column + 4) no hay caracteres y la celda "linea" > 1 entonces
'And IsError(Application.VLookup(Cells(i, 42), Range("AY6:BD10000"), 6, False)) = False
 'If IsError(Application.VLookup(Cells(i, 1), Range("Z6:BD10000"), 31, False)) = False Then
    'If Len(Cells(i, 52)) = 0 Then
        'Cells(i, Ultimate_Column + 3) = Application.VLookup(Cells(i, 1), Range("AY6:BD10000"), 6, False)
    'End If
'Next i

'Para compras catalogadas por Id articulo y Itm Id Vndr como catologo o contrato
'DAndo como resultado Cells(i, Ultimate_Column + 4) = "Gestion PO"
'Call tercera_revision
Call countPO2

'--------------------------------------------------------------------------------------------------------------------

'Definicion de los compradores de la Taxonomia
'si dice COE en campo "Coe Compra" ==> "TAXONOMIA" = "COMPRADOR_CARGO1"
'N�32 + 2 = Comprador
'N�6 + 2 = Coe Compra

Dim r As Integer

For r = 6 To Ultimate_Row
    If Len(Cells(r, 14)) > 0 Then 'que la descripcion no este vacia, si esta vacia se detiene.
        Select Case True

'si en "Coe Compra" (N�8) = "COE", "TOTAL_DOLAR" > 35.0000, "Clasificacion de compras"="Sourcing" => Taxonomia(N�41)="COMPRADOR_CARGO1" (N�9)
        Case Cells(r, 8).Value = "COE" And Cells(r, 46).Value >= 35000 And Cells(r, Ultimate_Column + 7).Value = "Sourcing" And Cells(r, Ultimate_Column + 6).Value = 1
            Cells(r, Ultimate_Column + 1) = Cells(r, 9).Text 'COMPRADOR_CARGO1
            Cells(r, Ultimate_Column + 1).Font.Bold = True 'negrita
            Cells(r, Ultimate_Column + 1).Font.ColorIndex = 26 'rosado
            Cells(r, Ultimate_Column + 4) = "COE"

'si en "Coe Compra" (N�8) = "COE", "TOTAL_DOLAR" > 35.0000, "Clasificacion de compras"="Sourcing" => Taxonomia(N�41)="COMPRADOR_CARGO1" (N�9)
        Case Cells(r, 8).Value = "COE" And Cells(r, 46).Value >= 35000 And Cells(r, Ultimate_Column + 7).Value = "Negociado por SASPA" And Cells(r, Ultimate_Column + 6).Value = 1
            Cells(r, Ultimate_Column + 1) = Application.VLookup(Cells(r, 7), Worksheets("data").Range("A2:B382"), 2, False) 'COMPRADOR_CARGO1
            Cells(r, Ultimate_Column + 1).Font.Bold = True 'negrita
            Cells(r, Ultimate_Column + 1).Font.ColorIndex = 5 'azul
            
'si en "Coe Compra" (N�8) = "Compra", "TOTAL_DOLAR" > 35.0000, "Clasificacion de compras"="Sourcing" => Taxonomia(N�41)="COMPRADOR_CARGO1" (N�9)
        Case Cells(r, 8).Value = "Compra" And Cells(r, 46).Value >= 35000 And Cells(r, Ultimate_Column + 7).Value = "Sourcing" And Cells(r, Ultimate_Column + 6).Value = 1
            Cells(r, Ultimate_Column + 1) = Cells(r, 9).Text 'COMPRADOR_CARGO1
            Cells(r, Ultimate_Column + 1).Font.Bold = True 'negrita
            Cells(r, Ultimate_Column + 1).Font.ColorIndex = 26 'rosado
            Cells(r, Ultimate_Column + 4) = "COE"
                       
'si en "Coe Compra" (N�8) = "COE", "TOTAL_DOLAR" < 35.0000, "COMPRADOR_CARGO1" (N�9) <> "Pablo Villarroel", "Clasificacion de compras"="Sourcing" => Taxonomia(N�41)="COMPRADOR_CARGO1" (N�9)
         Case Cells(r, 8).Value = "COE" And Cells(r, 46).Value < 35000 And Cells(r, Ultimate_Column + 6).Value = 1 And Cells(r, 9).Value <> "Pablo Villarroel" And Cells(r, Ultimate_Column + 7).Value = "Sourcing"
            Cells(r, Ultimate_Column + 1) = Cells(r, 9).Text 'COMPRADOR_CARGO1
            Cells(r, Ultimate_Column + 1).Font.Bold = True 'negrita
            Cells(r, Ultimate_Column + 1).Font.ColorIndex = 26 'rosado
            Cells(r, Ultimate_Column + 4) = "COE"
                  
'si en "Coe Compra" (N�8) = "COE", "TOTAL_DOLAR" < 35.0000, "COMPRADOR_CARGO1" (N�9) <> "Pablo Villarroel", "Clasificacion de compras"="Sourcing" => Taxonomia(N�41)="COMPRADOR_CARGO1" (N�9)
         Case Cells(r, 8).Value = "COE" And Cells(r, 46).Value < 35000 And Cells(r, Ultimate_Column + 6).Value = 1 And Cells(r, 9).Value <> "Pablo Villarroel" And Cells(r, Ultimate_Column + 7).Value = "Negociado por SASPA"
            Cells(r, Ultimate_Column + 1) = Cells(r, 9).Text 'COMPRADOR_CARGO1
            Cells(r, Ultimate_Column + 1).Font.Bold = True 'negrita
            Cells(r, Ultimate_Column + 1).Font.ColorIndex = 26 'rosado
            Cells(r, Ultimate_Column + 4) = "COE"
              
'si en "Coe Compra" (N�8) = "Compra", "TOTAL_DOLAR" < 35.0000, "COMPRADOR_CARGO1" (N�9) = "Pablo Villarroel", "Clasificacion de compras"="Negociado por SASPA" => Taxonomia(N�41)= Segun subcategoria
        Case Cells(r, 8).Value = "Compra" And Cells(r, 46).Value < 35000 And Cells(r, Ultimate_Column + 7).Value = "Negociado por SASPA" And Cells(r, Ultimate_Column + 6).Value = 1
            Cells(r, Ultimate_Column + 1) = Application.VLookup(Cells(r, 7), Worksheets("data").Range("A2:B382"), 2, False)
            Cells(r, Ultimate_Column + 1).Font.Bold = True 'negrita
            Cells(r, Ultimate_Column + 1).Font.ColorIndex = 5 'azul

'si en "Coe Compra" (N�8) = "COE", "TOTAL_DOLAR" < 35.0000, "COMPRADOR_CARGO1" (N�9) = "Pablo Villarroel", "Clasificacion de compras"="Sourcing" => Taxonomia(N�41)= Segun subcategoria
        Case Cells(r, 8).Value = "COE" And Cells(r, 46).Value < 35000 And Cells(r, Ultimate_Column + 6).Value = 1 And Cells(r, 9).Value = "Pablo Villarroel" And Cells(r, Ultimate_Column + 7).Value = "Sourcing"
            Cells(r, Ultimate_Column + 1) = Application.VLookup(Cells(r, 7), Worksheets("data").Range("A2:B382"), 2, False)
            Cells(r, Ultimate_Column + 1).Font.Bold = True 'negrita
            Cells(r, Ultimate_Column + 1).Font.ColorIndex = 5 'azul
            
'si en "Coe Compra" (N�8) = "COE", "TOTAL_DOLAR" < 35.0000, "COMPRADOR_CARGO1" (N�9) = "Pablo Villarroel", "Clasificacion de compras"="Negociado por SASPA" => Taxonomia(N�41)= Segun subcategoria
        Case Cells(r, 8).Value = "COE" And Cells(r, 46).Value < 35000 And Cells(r, Ultimate_Column + 6).Value = 1 And Cells(r, 9).Value = "Pablo Villarroel" And Cells(r, Ultimate_Column + 7).Value = "Negociado por SASPA"
            Cells(r, Ultimate_Column + 1) = Application.VLookup(Cells(r, 7), Worksheets("data").Range("A2:B382"), 2, False)
            Cells(r, Ultimate_Column + 1).Font.Bold = True 'negrita
            Cells(r, Ultimate_Column + 1).Font.ColorIndex = 5 'azul

'si en "Coe Compra" (N�8) = "Compra", "TOTAL_DOLAR" < 35.0000, "COMPRADOR_CARGO1" (N�9) = "Pablo Villarroel", "Clasificacion de compras"="Negociado por SASPA" => Taxonomia(N�41)= Segun subcategoria
        Case Cells(r, 8).Value = "Compra" And Cells(r, 46).Value < 35000 And Cells(r, Ultimate_Column + 7).Value = "Sourcing" And Cells(r, Ultimate_Column + 6).Value = 1
            Cells(r, Ultimate_Column + 1) = Application.VLookup(Cells(r, 7), Worksheets("data").Range("A2:B382"), 2, False)
            Cells(r, Ultimate_Column + 1).Font.Bold = True 'negrita
            Cells(r, Ultimate_Column + 1).Font.ColorIndex = 5 'azul
                     

'si en "Coe Compra" (N�8) = "Compra", "Area de compra" (N�53) = "Compra Normal", "Clasificacion de compras"="Sourcing" => Taxonomia(N�41)= Segun subcategoria
        Case Cells(r, 8).Value = "Compra" And Cells(r, Ultimate_Column + 4).Value = "Compra" And Cells(r, Ultimate_Column + 6).Value = 1 And Cells(r, Ultimate_Column + 7).Value = "Sourcing"
            Cells(r, Ultimate_Column + 1) = Application.VLookup(Cells(r, 7), Worksheets("data").Range("A2:B382"), 2, False)
            Cells(r, Ultimate_Column + 1).Font.Bold = True 'negrita
            Cells(r, Ultimate_Column + 1).Font.ColorIndex = 5 'azul
           
'si en "Coe Compra" (N�8) = "Compra", "Area de compra" (N�53) = "Compra Normal", "Clasificacion de compras"="Negociado por SASPA" => Taxonomia(N�41)= Segun subcategoria
        Case Cells(r, 8).Value = "Compra" And Cells(r, Ultimate_Column + 4).Value = "Compra" And Cells(r, Ultimate_Column + 6).Value = 1 And Cells(r, Ultimate_Column + 7).Value = "Negociado por SASPA"
            Cells(r, Ultimate_Column + 1) = Application.VLookup(Cells(r, 7), Worksheets("data").Range("A2:B382"), 2, False)
            Cells(r, Ultimate_Column + 1).Font.Bold = True 'negrita
            Cells(r, Ultimate_Column + 1).Font.ColorIndex = 5 'azul
   
'si en "Coe Compra" (N�8) = "Compra", "Area de compra" (N�53) = "Gestion PO" => Taxonomia(N�41)= "Gestion PO"
'en Taxonomia (N�41) se indica "Gestion Po", no asigna comprador
        Case Cells(r, 8).Value = "Compra" And Cells(r, Ultimate_Column + 4).Value = "Gestion PO" And Cells(r, Ultimate_Column + 6).Value = 1
            Cells(r, Ultimate_Column + 1) = "Gestion PO"
            Cells(r, Ultimate_Column + 1).Font.Bold = True 'negrita
            Cells(r, Ultimate_Column + 1).Font.ColorIndex = 30 'Cafe
            
'si en "Coe Compra" (N�8) = "COE", "Area de compra" (N�53) = "Gestion PO" => Taxonomia(N�41)= "Gestion PO"
'en Taxonomia (N�41) se indica "Gestion Po", no asigna comprador
        Case Cells(r, 8).Value = "COE" And Cells(r, Ultimate_Column + 4).Value = "Gestion PO" And Cells(r, Ultimate_Column + 6).Value = 1
            Cells(r, Ultimate_Column + 1) = "Gestion PO"
            Cells(r, Ultimate_Column + 1).Font.Bold = True 'negrita
            Cells(r, Ultimate_Column + 1).Font.ColorIndex = 30 'Cafe

        End Select
    End If
Next r

'--------------------------------------------------------------------------------------------------------------------
'Sobreescribir la columna 18, titulo de la columna "ReqId + CountPO2"
'si la cantidad de linea = 1 entonces colocar Llave en columna 18
'Debe ser Cells(i, 55) = 1 ya que hay solicitudes que no comienzan en 1

Range("R6:R10000").Clear
Cells(5, 18).Value = "ReqId + CountPO2"
            Cells(5, 18).Font.Bold = True 'negrita

For i = 6 To Ultimate_Row
    If Len(Cells(i, 14)) > 0 And Cells(i, 55) = 1 Then
            Cells(i, 18) = Cells(i, 1)
    End If
Next i

'Busca el comprador para las lineas Req Id + Linea <> 1, lo busca en Req Id + Linea = 1, que esta dado por Taxonomia o criterio anterior
'con ese comprador queda la solicitud completa asignada a solamente un comprador y con un criterio COE o NO COE
'El criterio de busqueda es por medio de BuscarV
Dim e As Integer
For e = 6 To Ultimate_Row
'si Cells(Ultimate_Column + 4) no hay caracteres y la celda "linea" > 1 entonces
'N�28 + 2 = Linea
'N�40 + 2 = ReqId
'AN = Req Id
'AX = Ultimate_Column + 3
    If Len(Cells(e, 14)) > 0 Then
    Select Case True
    Case Cells(e, Ultimate_Column + 6).Value > 1
            Cells(e, Ultimate_Column + 1) = Application.VLookup(Cells(e, 1).Text, Range("R6:AX10000"), 33, False)
            Cells(e, Ultimate_Column + 1).Font.Bold = True 'negrita
            Cells(e, Ultimate_Column + 1).Font.ColorIndex = 5 'azul
            Cells(e, Ultimate_Column + 4) = Application.VLookup(Cells(e, 1).Text, Range("R6:BA10000"), 36, False)
    End Select
    End If
Next e
'--------------------------------------------------------------------------------------------------------------------
'Reclasificacion de las Areas de Compras, para las lineas sin Area de compra, por lo comun son compras por Politica
'Busca el Area de Compras para las lineas Req Id + Linea <> 1, lo busca en Req Id + Linea = 1, que esta dado por Politica
'Dim f As Integer
'For f = 6 To Ultimate_Row
    'If Len(Cells(f, 14)) > 0 Then
    'Select Case True
    'Case Cells(f, Ultimate_Column + 6).Value > "1" And Len(Cells(f, Ultimate_Column + 6)) = 0
            'Cells(f, Ultimate_Column + 4) = Application.VLookup(Cells(f, 1).Text, Range("Z6:BC10000"), 28, False)
    'End Select
    'End If
'Next f
'--------------------------------------------------------------------------------------------------------------------

'En esta etapa hace la revision segun tipo de compra ==> gestion PO si es contrato o catalogo
'buscar si esta realziada la compra en la planilla PO que se obtiene de Query de PS
Call compra_realizada
'busca si esta asignado ayer,para mantener al comprador
Call asignado_ayer
Call avance_PO

'Cells(5, Ultimate_Column + 1)= Taxonomia
'Cells(5, Ultimate_Column + 2).Value = "Concatenado"  SE ELIMINA
'Cells(5, Ultimate_Column + 3).Value = "Tipo de compra"
'Cells(5, Ultimate_Column + 4).Value = "Area de compra"

'borrado de columnas para que se vea mejor
Columns(Ultimate_Column + 2).EntireColumn.Delete

'--------------------------------------------------------------------------------------------------------------------
'Contadores para definir Gestion PO y Compra Normal, en Gestion PO
'Conteo a nivel de lineas segun "Area de compra" = Gestion PO / Compra
Dim pega5 As Integer
Dim count_GestionPO As Integer
Dim count_Compra_Normal As Integer
count_GestionPO = 0
count_Compra_Normal = 0

For pega5 = 6 To Ultimate_Row
    If (InStr(Cells(pega5, 52).Text, "Gestion PO")) Then
    count_GestionPO = count_GestionPO + 1
    ElseIf (InStr(Cells(pega5, 52).Text, "Compra")) Then
    count_Compra_Normal = count_Compra_Normal + 1
    End If
Next pega5

'Conteo a nivel de solicitud de las Taxonomias que tienen emisores de "Gestion PO",
'debido a Reasignacion (COMPRADOR_CARGO1), columna 9.
'Esto es antes de la asignacion equitativa
Dim pega8 As Integer
Dim count_Matias_Carranza As Integer 'Matias Carranza
Dim count_Rosirys_Matos As Integer 'Rosirys Matos
Dim count_Daniel_Alvarez As Integer 'Daniel Alvarez
Dim count_Angelica_Luza As Integer 'Angeliza Luza
Dim count_Maite_Yanez As Integer 'Maite Yanez

count_Rosirys_Matos = 0
count_Daniel_Alvarez = 0
count_Matias_Carranza = 0
count_Angelica_Luza = 0
count_Maite_Yanez = 0

For pega8 = 6 To Ultimate_Row
If (InStr(Cells(pega8, 50).Text, "Matias Carranza")) And Cells(pega8, 54).Value = 1 Then
    count_Matias_Carranza = count_Matias_Carranza + 1
ElseIf (InStr(Cells(pega8, 50).Text, "Angelica Luza")) And Cells(pega8, 54).Value = 1 Then
    count_Angelica_Luza = count_Angelica_Luza + 1
ElseIf (InStr(Cells(pega8, 50).Text, "Maite Yanez")) And Cells(pega8, 54).Value = 1 Then
    count_Maite_Yanez = count_Maite_Yanez + 1
ElseIf (InStr(Cells(pega8, 50).Text, "Rosirys Matos")) And Cells(pega8, 54).Value = 1 Then
    count_Rosirys_Matos = count_Rosirys_Matos + 1
ElseIf (InStr(Cells(pega8, 50).Text, "Daniel Alvarez")) And Cells(pega8, 54).Value = 1 Then
    count_Daniel_Alvarez = count_Daniel_Alvarez + 1
End If
Next pega8


'Conteo a nivel de solicitudes de Taxonomias que tienen "Gestion PO"
'Son las solicutdes no asignadas.
Dim contador_gestion_PO_file50 As Integer
Dim count123 As Integer
contador_gestion_PO_file50 = 0

For count123 = 6 To Ultimate_Row
    If (InStr(Cells(count123, 50).Text, "Gestion PO")) And Cells(count123, 54).Value = 1 Then
    contador_gestion_PO_file50 = contador_gestion_PO_file50 + 1
    End If
Next count123

'Conteo a nivel de solictudes de Taxonomias que tienen "Gestion PO"
'Es decir son solicitudes no asignadas, en los "0 dias" y "1 dias"
Dim recientes As Integer
Dim count12345 As Integer
recientes = 0

For count12345 = 6 To Ultimate_Row
    Select Case True
    Case Cells(count12345, 3).Text = "0 D�as" And InStr(Cells(count12345, 50).Text, "Gestion PO") And Cells(count12345, 54).Value = 1
        recientes = recientes + 1
    Case Cells(count12345, 3).Text = "1 D�as" And InStr(Cells(count12345, 50).Text, "Gestion PO") And Cells(count12345, 54).Value = 1
        recientes = recientes + 1
    End Select
Next count12345

'Cantidad de solicitudes a repartir en Gestion PO
'https://www.exceltrick.com/formulas_macros/vba-msgbox/
'https://powerspreadsheets.com/excel-vba-inputbox/
Dim emisores As Integer
emisores = InputBox("Cantidad de emisores en Gestion PO", "Gestion PO", "5")
MsgBox "La cantidad de emisores es de: " & emisores
Dim division As Integer
division = ((recientes) / emisores) 'esto debido a que hay OC de Gestion PO que son Sourcing por lo tanto compra normal

'La cantidad de solicitudes asignadas, se suma mas lineas hasta que sea igual o menor
'a la division de solicitudes sin asignar del dia 0 y dia 1
Dim countCARRANZA As Integer
Dim countLUZA As Integer
Dim countYANEZ As Integer
Dim countMATOS As Integer
Dim countALVAREZ As Integer

countCARRANZA = 0
countLUZA = 0
countYANEZ = 0
countMATOS = 0
countALVAREZ = 0

Dim Alerta4 As Integer
Alerta4 = MsgBox("Cantidad de solicitudes recientes (0 Dias y 1 Dias) = " & recientes & vbCrLf & "Cantidad de solicitudes a repartir = " & division & vbCrLf & "Total Lineas Gestion PO = " & count_GestionPO & vbCrLf & "Total Lineas Compra normal = " & count_Compra_Normal & vbCrLf & "Solicitudes a Carranza = " & count_Matias_Carranza & vbCrLf & "Solicitudes a Luza = " & count_Angelica_Luza & vbCrLf & "Solicitudes a Yanez = " & count_Maite_Yanez & vbCrLf & "Solicitudes a Matus = " & count_Rosirys_Matos & vbCrLf & "Solicitudes a Alvarez = " & count_Daniel_Alvarez & vbCrLf, vbOKOnly)


Dim pa As Integer
For pa = 6 To Ultimate_Row
    Select Case True
    Case (countCARRANZA <= division) And (Cells(pa, 50).Text = "Gestion PO") And Cells(pa, 54).Value = 1
        Cells(pa, 50) = "Matias Carranza"
        Cells(pa, 50).Font.ColorIndex = 30 'Cafe
        countCARRANZA = countCARRANZA + 1
    
    Case (countLUZA <= division) And (Cells(pa, 50).Text = "Gestion PO") And Cells(pa, 54).Value = 1
        Cells(pa, 50) = "Angelica Luza"
        Cells(pa, 50).Font.ColorIndex = 30 'Cafe
        countLUZA = countLUZA + 1
    
    Case (countYANEZ <= division) And (Cells(pa, 50).Text = "Gestion PO") And Cells(pa, 54).Value = 1
        Cells(pa, 50) = "Maite Yanez"
        Cells(pa, 50).Font.ColorIndex = 30 'Cafe
        countYANEZ = countYANEZ + 1
        
    Case (countMATOS <= division) And (Cells(pa, 50).Text = "Gestion PO") And Cells(pa, 54).Value = 1
        Cells(pa, 50) = "Rosirys Matos"
        Cells(pa, 50).Font.ColorIndex = 30 'Cafe
        countMATOS = countMATOS + 1
        
    Case (countALVAREZ <= division) And (Cells(pa, 50).Text = "Gestion PO") And Cells(pa, 54).Value = 1
        Cells(pa, 50) = "Daniel Alvarez"
        Cells(pa, 50).Font.ColorIndex = 30 'Cafe
        countALVAREZ = countALVAREZ + 1
    End Select
Next pa

'Repartir las solicitudes que tengan un numero de linea <> 1, en Gestion PO en los
'dias 0 y dias 1
Dim repartir_gestion_po As Integer
For repartir_gestion_po = 6 To Ultimate_Row
    If Cells(repartir_gestion_po, 54).Value > 1 Then
    Select Case True
    Case Cells(repartir_gestion_po, 3).Text = "0 D�as"
        Cells(repartir_gestion_po, 50) = Application.VLookup(Cells(repartir_gestion_po, 1), Range("R6:AX10000"), 33, False)
        Cells(repartir_gestion_po, 50).Font.ColorIndex = 30 'Cafe
    Case Cells(repartir_gestion_po, 3).Text = "1 D�as"
        Cells(repartir_gestion_po, 50) = Application.VLookup(Cells(repartir_gestion_po, 1), Range("R6:AX10000"), 33, False)
        Cells(repartir_gestion_po, 50).Font.ColorIndex = 30 'Cafe
    End Select
    End If
Next repartir_gestion_po


'conteo de solicitudes por personas en gestion PO
Dim pega1 As Integer
Dim count1 As Integer
Dim count2 As Integer
Dim count3 As Integer
Dim count4 As Integer
Dim count5 As Integer

count1 = 0
count2 = 0
count3 = 0
count4 = 0
count5 = 0


For pega1 = 6 To Ultimate_Row
If (InStr(Cells(pega1, Ultimate_Column + 1).Text, "Matias Carranza")) And Cells(pega1, 54).Value = 1 Then
    count1 = count1 + 1
ElseIf (InStr(Cells(pega1, Ultimate_Column + 1).Text, "Angelica Luza")) And Cells(pega1, 54).Value = 1 Then
    count2 = count2 + 1
ElseIf (InStr(Cells(pega1, Ultimate_Column + 1).Text, "Maite Yanez")) And Cells(pega1, 54).Value = 1 Then
    count3 = count3 + 1
ElseIf (InStr(Cells(pega1, Ultimate_Column + 1).Text, "Rosirys Matos")) And Cells(pega1, 54).Value = 1 Then
    count4 = count4 + 1
ElseIf (InStr(Cells(pega1, Ultimate_Column + 1).Text, "Daniel Alvarez")) And Cells(pega1, 54).Value = 1 Then
    count5 = count5 + 1
End If
Next pega1

'Conteo de solicitudes "Compra Normal" y "Gestion PO"
Dim pega12 As Integer
Dim count_Sourcing As Integer 'Sourcing
Dim count_Contrato As Integer 'Contrato
Dim count_Catalogo As Integer 'Catalogo
count_Sourcing = 0
count_Contrato = 0
count_Catalogo = 0

For pega12 = 6 To Ultimate_Row
If (InStr(Cells(pega12, 51).Text, "Sourcing")) And Cells(pega12, 54).Value = 1 Then
    count_Sourcing = count_Sourcing + 1
ElseIf (InStr(Cells(pega12, 51).Text, "Contrato")) And Cells(pega12, 54).Value = 1 Then
    count_Contrato = count_Contrato + 1
ElseIf (InStr(Cells(pega12, 51).Text, "Catalogo")) And Cells(pega12, 54).Value = 1 Then
    count_Catalogo = count_Catalogo + 1
End If
Next pega12

Call Clasificacion_Compras

'mensaje final
MsgBox ("Total de solicitudes Matias Carranza " & count1 & vbCrLf & "Total de solicitudes Angelica Luza " & count2 & vbCrLf & "Total de solicitudes Maite Yanez " & count3 & vbCrLf & "Total de solicitudes Rosirys Matos " & count4 & vbCrLf & "Total de solicitudes Daniel Alvarez " & count5 & vbCrLf & "Total lineas Gestion PO " & count_GestionPO & vbCrLf & "Total lineas Compra normal " & count_Compra_Normal & vbCrLf & "Total de solicitudes Sourcing " & count_Sourcing & vbCrLf & "Total de solicitudes Contrato " & count_Contrato & vbCrLf & "Total de solicitudes Catalogo " & count_Catalogo)
 
'count1 Matias Carranza & vbCrLf & "Total de lineas para Maite Yanez " & count1 & vbCrLf &
'count2 Angeliza Luza & vbCrLf & "Total de lineas para Maite Yanez " & count2 & vbCrLf &
'count3 Maite Yanez & vbCrLf & "Total de lineas para Maite Yanez " & count3 & vbCrLf &
'count4 Rosirys Matos & vbCrLf & "Total de lineas para Rosirys Matos " & count4 & vbCrLf &
'count5 Daniel Alvarez & vbCrLf & "Total de lineas para Daniel Alvarez " & count5 & vbCrLf &

End Sub

Sub revision_primera()
Dim i As Integer

'escenario combinado
'compra
'Compra
'COMPRA
'comp,Comp, COMP, Comprometida, COMPROMETIDA, comprometida,COMPROMETIDO, Comprometido, comprometido

'en el scripts comp,Comp, COMP
'en el scripts omprometi  = Comprometido, comprometido, Comprometida, comprometida
'en el scripts COMPROMETI = COMPROMETIDO, COMPROMETIDA
'en el scripts egulariza , EGULARIZA = Regularizacion, regularizacion , REGULARIZACION
'en el scripts bh , bh



For i = 6 To Ultimate_Row


'N�40 = Req Id + 2
'N�28 = Linea + 2
    
If Cells(i, 42).Value > 0 And Cells(i, 30) = 1 Then
    Select Case True

'compra, Compra, COMPRA
'N�14 = Desc Articulo
    Case InStr(Cells(i, 14).Value, "ompra") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "ompra", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        ElseIf InStr(Cells(i, Ultimate_Column + 1).Value, "EGULARIZA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "egulariza") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
        ElseIf InStr(Cells(i, Ultimate_Column + 1).Value, "BH") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "bh") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"
        ElseIf InStr(Cells(i, Ultimate_Column + 1).Value, "gencia") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "GENCIA") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Emergencia"
        Else
            Call Compra_PO(i)
        End If
        
    Case InStr(Cells(i, 14).Value, "COMPRA") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPRA", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        ElseIf InStr(Cells(i, Ultimate_Column + 1).Value, "EGULARIZA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "egulariza") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
        ElseIf InStr(Cells(i, Ultimate_Column + 1).Value, "BH") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "bh") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"
        ElseIf InStr(Cells(i, Ultimate_Column + 1).Value, "gencia") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "GENCIA") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Emergencia"
        Else
            Call Compra_PO(i)
        End If
    
'COMPLETO OR ompleto
    Case InStr(Cells(i, 14).Value, "ompleto") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "ompleto", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
        
    Case InStr(Cells(i, 14).Value, "COMPLETO") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPLETO", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If

'COMPUTACIONA OR omputaciona
    Case InStr(Cells(i, 14).Value, "COMPUTACIONA") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPUTACIONA", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
        
    Case InStr(Cells(i, 14).Value, "omputaciona") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "omputaciona", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If

'Complete
    Case InStr(Cells(i, 14).Value, "omplete") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "omplete", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
            
            
    Case InStr(Cells(i, 14).Value, "COMPLETE") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPLETE", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
        
'COMPATIBL OR ompatibl
    Case InStr(Cells(i, 14).Value, "COMPATIBL") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPATIBL", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
             Call Compra_PO(i)
        End If
            
    Case InStr(Cells(i, 14).Value, "ompatibl") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "ompatibl", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
            
'Complejidad OR omplejidad
    Case InStr(Cells(i, 14).Value, "omplejidad") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "omplejidad", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
            
    Case InStr(Cells(i, 14).Value, "COMPLEJIDAD") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPLEJIDAD", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
    
'COMPETITIVA OR ompetitiva
    Case InStr(Cells(i, 14).Value, "ompetitiva") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "ompetitiva", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
            
    Case InStr(Cells(i, 14).Value, "COMPETITIVA") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPETITIVA", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
        
'COMPRENDE OR omprende
    Case InStr(Cells(i, 14).Value, "COMPRENDE") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPRENDE", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
            
    Case InStr(Cells(i, 14).Value, "omprende") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "omprende", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
            
'COMPARTIR OR ompartir
    Case InStr(Cells(i, 14).Value, "COMPARTIR") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPARTIR", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
            
            
    Case InStr(Cells(i, 14).Value, "ompartir") > 0
         Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "ompartir", "", 1)
         If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If


'Comp = Comp, comp, COMP
'omprometid = Comprometido + Comprometida + comprometido + comprometida
'COMPROMETID = COMPROMETIDO + COMPROMETIDA
'rometid = comprometido + comprometida

    Case InStr(Cells(i, 14).Value, "Comp") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "Comp", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometid") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETID") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "rometid") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "BH") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "bh") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If

    Case InStr(Cells(i, 14).Value, "COMP") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMP", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometid") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETID") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ROMETID") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "BH") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "bh") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
      
    Case InStr(Cells(i, 14).Value, "comp") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "comp", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometid") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETID") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "rometid") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "BH") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "bh") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If

'COMPUTADORA OR omputadora
    Case InStr(Cells(i, 14).Value, "COMPUTADORA") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPUTADORA", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
            
    Case InStr(Cells(i, 14).Value, "computadora") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "computadora", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If

        
'Detector de las siguientes palabras COMPROMETIDA - BH
'Comprometido - Infraestructura - Negociado por COE - Adjudicado por el COE - COE - Addendum - Proyecto de gobierno.


'COMPROMETIDA
    Case InStr(Cells(i, 14).Value, "Comprometida") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
    
    Case InStr(Cells(i, 14).Value, "COMPROMETIDA") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"

    Case InStr(Cells(i, 14).Value, "comprometida") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
            
'Comprometido
    Case InStr(Cells(i, 14).Value, "Compremetido") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
            
    Case InStr(Cells(i, 14).Value, "COMPROMETIDO") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
       
    Case InStr(Cells(i, 14).Value, "compremetido") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"

'flujo especial considerando que el compro es parte de comprometida, por lo cual se agrega metida
            
'compro
    Case InStr(Cells(i, 14).Value, "compro") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "compro", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "Comprometida") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETIDA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comprometida") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "Compremetido") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETIDO") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comprometido") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "metido") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "metida") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
        
    Case InStr(Cells(i, 14).Value, "Compro") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "Compro", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "Comprometida") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETIDA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comprometida") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "Compremetido") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETIDO") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comprometido") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "metido") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "metida") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
        
    Case InStr(Cells(i, 14).Value, "COMPRO") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPRO", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "Comprometida") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETIDA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comprometida") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "Compremetido") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETIDO") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comprometido") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "METIDO") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "METIDA") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If

'BH
    Case InStr(Cells(i, 14).Value, "BH") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"

    Case InStr(Cells(i, 14).Value, "bh") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"

    Case InStr(Cells(i, 14).Value, "honorarios") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"
            
    Case InStr(Cells(i, 14).Value, "Honorarios") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"

    Case InStr(Cells(i, 14).Value, "HONORARIOS") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"

    Case InStr(Cells(i, 14).Value, "honorario") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"

    Case InStr(Cells(i, 14).Value, "Honorario") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"

    Case InStr(Cells(i, 14).Value, "HONORARIO") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"
            
    Case InStr(Cells(i, 14).Value, "Bh") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"
            
    Case InStr(Cells(i, 14).Value, "bH") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"
            
'Infraestructura
    Case InStr(Cells(i, 14).Value, "Infraestructura") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
           
    Case InStr(Cells(i, 14).Value, "infraestructura") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
  
    Case InStr(Cells(i, 14).Value, "INFRAESTRUCTURA") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
    
    Case InStr(Cells(i, 14).Value, "fraestructu") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
    
    Case InStr(Cells(i, 14).Value, "FRAESTRUCTU") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"

            
'Negociado por COE
    Case InStr(Cells(i, 14).Value, "Negociado por COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "negociado por COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "NEGOCIADO por COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "NEGOCIADO POR COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "negociado por coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "negociada por el coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "negociado por el coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "Negociado por el COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
                 
    Case InStr(Cells(i, 14).Value, "Negociad por el COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
    
    Case InStr(Cells(i, 14).Value, "negociad por el COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
                      
    Case InStr(Cells(i, 14).Value, "negociad por el coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "NEGOCIADO POR EL COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
 
    Case InStr(Cells(i, 14).Value, "NEGOC X COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
 
     Case InStr(Cells(i, 14).Value, "negoc X COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
 
     Case InStr(Cells(i, 14).Value, "negoc X coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"

     Case InStr(Cells(i, 14).Value, "Negociado por CoE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "negociado por CoE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "NEGOCIADO POR CoE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "Negociado por Coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "NEGOCIADO_POR_EL_COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "negociado_por_el_coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
                       
    Case InStr(Cells(i, 14).Value, "Negociado_por_el_Coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"

'Negociado por SASPA = EGOCIAD ,egociad
'Criterios complementarios: SSC, SC, SSA, SSAA, SA, S.A, ERVICIO NDINO, ervicio ndino, SSS, SS.AA., CSC, SASPA, SCC
            
     Case InStr(Cells(i, 14).Value, "NEG") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "NEG", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "COE") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "coe") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "CoE") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "Coe") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
        ElseIf InStr(Cells(i, Ultimate_Column + 1).Value, "SSC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SSA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SSAA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "S.A") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ERVICIO NDINO") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ervicio ndino") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SSS") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SS.AA.") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "CSC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SASPA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SCC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "NDINOS") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ndino") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ervicio") > 0 Then
            Cells(i, Ultimate_Column + 7) = "Negociado por SASPA"
            Cells(i, Ultimate_Column + 4) = "Compra"
            Cells(i, 14).Font.Bold = True 'negrita
            Cells(i, 14).Font.ColorIndex = 3 'pintar en rojo
            Cells(i, 27) = Cells(i, 1)
        Else
            Call Compra_PO(i)
        End If
            
    Case InStr(Cells(i, 14).Value, "neg") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "neg", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "COE") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "coe") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "CoE") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "Coe") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
        ElseIf InStr(Cells(i, Ultimate_Column + 1).Value, "SSC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SSA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SSAA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "S.A") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ERVICIO NDINO") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ervicio ndino") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SSS") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SS.AA.") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "CSC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SASPA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SCC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "NDINOS") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ndino") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ervicio") > 0 Then
            Cells(i, Ultimate_Column + 7) = "Negociado por SASPA"
            Cells(i, Ultimate_Column + 4) = "Compra"
            Cells(i, 14).Font.Bold = True 'negrita
            Cells(i, 14).Font.ColorIndex = 3 'pintar en rojo
            Cells(i, 27) = Cells(i, 1)
        Else
            Call Compra_PO(i)
        End If
            
    Case InStr(Cells(i, 14).Value, "Neg") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "Neg", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "COE") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "coe") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "CoE") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "Coe") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
        ElseIf InStr(Cells(i, Ultimate_Column + 1).Value, "SSC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SSA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SSAA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "S.A") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ERVICIO NDINO") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ervicio ndino") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SSS") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SS.AA.") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "CSC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SASPA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SCC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "NDINOS") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ndino") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ervicio") > 0 Then
            Cells(i, Ultimate_Column + 7) = "Negociado por SASPA"
            Cells(i, Ultimate_Column + 4) = "Compra"
            Cells(i, 14).Font.Bold = True 'negrita
            Cells(i, 14).Font.ColorIndex = 3 'pintar en rojo
            Cells(i, 27) = Cells(i, 1)
        Else
            Call Compra_PO(i)
        End If

            
'Adjudicado por el COE
    Case InStr(Cells(i, 14).Value, "Adjudicado por el COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"

    Case InStr(Cells(i, 14).Value, "adjudicado por el COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "adjudicado por el coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"

    Case InStr(Cells(i, 14).Value, "ADJUDICADO por el COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
                        
     Case InStr(Cells(i, 14).Value, "ADJUDICADO POR COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
     Case InStr(Cells(i, 14).Value, "adjudicado por COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
      Case InStr(Cells(i, 14).Value, "adjudicado por coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
      Case InStr(Cells(i, 14).Value, "Adjudicado por coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
    
    Case InStr(Cells(i, 14).Value, "Adjudicado por COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"

    Case InStr(Cells(i, 14).Value, "Adjudicado por Coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
                               
'Addendum
    Case InStr(Cells(i, 14).Value, "Addendum") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
                        
    Case InStr(Cells(i, 14).Value, "addendum") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
              
    Case InStr(Cells(i, 14).Value, "ADDENDUM") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
               
    Case InStr(Cells(i, 14).Value, "Adendum") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
               
    Case InStr(Cells(i, 14).Value, "adendum") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"

    Case InStr(Cells(i, 14).Value, "ADENDUM") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"

'Proyecto de gobierno.
    Case InStr(Cells(i, 14).Value, "Proyecto de gobierno") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Proyecto"
      
'--------------------------------------------------------------------------------------
'criterios agregados el 11-01
'Gastronomia
    Case InStr(Cells(i, 14).Value, "Gastronomia") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
           
    Case InStr(Cells(i, 14).Value, "GASTRONOMIA") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
          
    Case InStr(Cells(i, 14).Value, "gastronomia") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"

    Case InStr(Cells(i, 14).Value, "gastronom�a") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "Gastronom�a") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
                
'Medios
    Case InStr(Cells(i, 14).Value, "Medios") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
           
    Case InStr(Cells(i, 14).Value, "medios") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
          
    Case InStr(Cells(i, 14).Value, "MEDIOS") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
'Gobierno
    Case InStr(Cells(i, 14).Value, "Gobierno") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Gobierno"
           
    Case InStr(Cells(i, 14).Value, "gobierno") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Gobierno"
          
    Case InStr(Cells(i, 14).Value, "GOBIERNO") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Gobierno"
            
'Regularizacion agregados el 23-01

    Case InStr(Cells(i, 14).Value, "egulariza") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
            
    Case InStr(Cells(i, 14).Value, "EGULARIZA") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
    
    Case InStr(Cells(i, 14).Value, "Regularizacion") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
            
    Case InStr(Cells(i, 14).Value, "REGULARIZACION") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
    
    Case InStr(Cells(i, 14).Value, "regularizacion") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
            
    Case InStr(Cells(i, 14).Value, "REGULARIZACI�N") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
            
    Case InStr(Cells(i, 14).Value, "Regularizaci�n") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
            
    Case InStr(Cells(i, 14).Value, "Regularizaci�n") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
            
    Case InStr(Cells(i, 14).Value, "regularizaci�n") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
            
    Case InStr(Cells(i, 14).Value, "REGULARIZACI�N") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
                 
    Case InStr(Cells(i, 14).Value, "REGULARIZACI�") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
            
    Case InStr(Cells(i, 14).Value, "REGULARIZACON") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
          
    Case InStr(Cells(i, 14).Value, "Regularizai�n") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
            
    Case InStr(Cells(i, 14).Value, "REGULARZACION") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
            
'--------------------------------------------------------------------------------------
'Suscripcion agregados el 05-02
    
    Case InStr(Cells(i, 14).Value, "suscripcion") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
    
    Case InStr(Cells(i, 14).Value, "Suscripcion") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
        
    Case InStr(Cells(i, 14).Value, "SUSCRIPCION") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
    
    Case InStr(Cells(i, 14).Value, "suscripci�n") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
    
    Case InStr(Cells(i, 14).Value, "Suscripci�n") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
    
    Case InStr(Cells(i, 14).Value, "SUSCRIPCI�N") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "SUSCRIPCI�NES") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"

    Case InStr(Cells(i, 14).Value, "suscripci�nes") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"

    Case InStr(Cells(i, 14).Value, "SUSCRIPCIONES") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "suscripciones") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "suscr") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"

    Case InStr(Cells(i, 14).Value, "SUSCRIPCI�N") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"

'--------------------------------------------------------------------------------------
'Membresia agregados el 10-08
    
    Case InStr(Cells(i, 14).Value, "Membresia") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
       
    Case InStr(Cells(i, 14).Value, "MEMBRESIA") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "membresia") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "Membrecia") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "membrecia") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "MEMBRECIA") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
'Investigacion agregados el 13-08
    
    Case InStr(Cells(i, 14).Value, "INVESTIGACION") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "INVESTIGACI�N") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "investigacion") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "investigaci�n") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "Investigacion") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "Investigaci�n") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
'Emergencia agregados el 07-09
'Criterios complementarios , uces, UCES, �mpara, �MPARA, ampara, ampara, OTON, OTON, uz

    Case InStr(Cells(i, 14).Value, "Emergencia") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "Emergencia", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "UCES") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "uces") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "�mpara") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "�MPARA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "AMPARA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ampara") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "OTON") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "oton") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "uz") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "UZ") > 0 Then
            Call Compra_PO(i)
        Else
            Cells(i, Ultimate_Column + 7) = "Emergencia"
            Call Gestion_PO(i)
        End If

    Case InStr(Cells(i, 14).Value, "EMERGENCIA") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "EMERGENCIA", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "UCES") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "uces") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "�mpara") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "�MPARA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "AMPARA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ampara") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "OTON") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "oton") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "uz") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "UZ") > 0 Then
            Call Compra_PO(i)
        Else
            Cells(i, Ultimate_Column + 7) = "Emergencia"
            Call Gestion_PO(i)
        End If
                                           
    Case InStr(Cells(i, 14).Value, "emergencia") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "emergencia", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "UCES") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "uces") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "�mpara") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "�MPARA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "AMPARA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ampara") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "OTON") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "oton") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "uz") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "UZ") > 0 Then
            Call Compra_PO(i)
        Else
            Cells(i, Ultimate_Column + 7) = "Emergencia"
            Call Gestion_PO(i)
        End If
           
    Case Else
            Call Compra_PO(i)
    End Select
            
End If
Next i
End Sub
Sub revision_segunda()
Dim i As Integer
'comp,Comp, COMP, Comprometida, COMPROMETIDA, comprometida,COMPROMETIDO, Comprometido, comprometido,compro, Compro, COMPRO,
'213

For i = 6 To Ultimate_Row
'escenario combinado
    
'If Len(Cells(i, Ultimate_Column + 4)) = 1 Or Len(Cells(i, Ultimate_Column + 4)) = 0 Then
If Len(Cells(i, Ultimate_Column + 7)) = 0 Then

    Select Case True


'compra, Compra, COMPRA
'N�14 = Desc Articulo
    Case InStr(Cells(i, 14).Value, "ompra") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "ompra", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        ElseIf InStr(Cells(i, Ultimate_Column + 1).Value, "EGULARIZA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "egulariza") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
        ElseIf InStr(Cells(i, Ultimate_Column + 1).Value, "BH") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "bh") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"
        ElseIf InStr(Cells(i, Ultimate_Column + 1).Value, "gencia") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "GENCIA") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Emergencia"
        Else
            Call Compra_PO(i)
        End If
        
    Case InStr(Cells(i, 14).Value, "COMPRA") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPRA", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        ElseIf InStr(Cells(i, Ultimate_Column + 1).Value, "EGULARIZA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "egulariza") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
        ElseIf InStr(Cells(i, Ultimate_Column + 1).Value, "BH") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "bh") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"
        ElseIf InStr(Cells(i, Ultimate_Column + 1).Value, "gencia") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "GENCIA") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Emergencia"
        Else
            Call Compra_PO(i)
        End If
    
'COMPLETO OR ompleto
    Case InStr(Cells(i, 14).Value, "ompleto") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "ompleto", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
        
    Case InStr(Cells(i, 14).Value, "COMPLETO") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPLETO", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If

'COMPUTACIONA OR omputaciona
    Case InStr(Cells(i, 14).Value, "COMPUTACIONA") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPUTACIONA", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
        
    Case InStr(Cells(i, 14).Value, "omputaciona") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "omputaciona", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If

'Complete
    Case InStr(Cells(i, 14).Value, "omplete") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "omplete", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
            
            
    Case InStr(Cells(i, 14).Value, "COMPLETE") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPLETE", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
             Call Compra_PO(i)
        End If
        
'COMPATIBL OR ompatibl
    Case InStr(Cells(i, 14).Value, "COMPATIBL") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPATIBL", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
            
    Case InStr(Cells(i, 14).Value, "ompatibl") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "ompatibl", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
            
'Complejidad OR omplejidad
    Case InStr(Cells(i, 14).Value, "omplejidad") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "omplejidad", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
            
    Case InStr(Cells(i, 14).Value, "COMPLEJIDAD") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPLEJIDAD", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
    
'COMPETITIVA OR ompetitiva
    Case InStr(Cells(i, 14).Value, "ompetitiva") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "ompetitiva", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
            
    Case InStr(Cells(i, 14).Value, "COMPETITIVA") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPETITIVA", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
        
'COMPRENDE OR omprende
    Case InStr(Cells(i, 14).Value, "COMPRENDE") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPRENDE", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
            
    Case InStr(Cells(i, 14).Value, "omprende") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "omprende", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
            
'COMPARTIR OR ompartir
    Case InStr(Cells(i, 14).Value, "COMPARTIR") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPARTIR", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
            
            
    Case InStr(Cells(i, 14).Value, "ompartir") > 0
         Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "ompartir", "", 1)
         If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If


'Comp = Comp, comp, COMP
'omprometid = Comprometido + Comprometida + comprometido + comprometida
'COMPROMETID = COMPROMETIDO + COMPROMETIDA
'rometid = comprometido + comprometida

    Case InStr(Cells(i, 14).Value, "Comp") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "Comp", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometid") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETID") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "rometid") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "BH") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "bh") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If

    Case InStr(Cells(i, 14).Value, "COMP") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMP", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometid") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETID") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ROMETID") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "BH") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "bh") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
      
    Case InStr(Cells(i, 14).Value, "comp") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "comp", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometid") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETID") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "rometid") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "BH") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "bh") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If

'COMPUTADORA OR omputadora
    Case InStr(Cells(i, 14).Value, "COMPUTADORA") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPUTADORA", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
            
    Case InStr(Cells(i, 14).Value, "computadora") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "computadora", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETI") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "omprometi") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If

'Detector de las siguientes palabras COMPROMETIDA - BH
'Comprometido - Infraestructura - Negociado por COE - Adjudicado por el COE - COE - Addendum - Proyecto de gobierno.


'COMPROMETIDA
    Case InStr(Cells(i, 14).Value, "Comprometida") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
    
    Case InStr(Cells(i, 14).Value, "COMPROMETIDA") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
    
    Case InStr(Cells(i, 14).Value, "comprometida") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        
'Comprometido
    Case InStr(Cells(i, 14).Value, "Compremetido") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
            
      Case InStr(Cells(i, 14).Value, "COMPROMETIDO") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        
     Case InStr(Cells(i, 14).Value, "compremetido") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
       

'flujo especial considerando que el compro es parte de comprometida, por lo cual se agrega metida
            
'compro
    Case InStr(Cells(i, 14).Value, "compro") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "compro", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "Comprometida") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETIDA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comprometida") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "Compremetido") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETIDO") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comprometido") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "metido") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "metida") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
        
    Case InStr(Cells(i, 14).Value, "Compro") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "Compro", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "Comprometida") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETIDA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comprometida") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "Compremetido") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETIDO") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comprometido") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "metido") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "metida") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If
        
    Case InStr(Cells(i, 14).Value, "COMPRO") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "COMPRO", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "Comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMP") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comp") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "Comprometida") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETIDA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comprometida") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "Compremetido") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "COMPROMETIDO") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "comprometido") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "METIDO") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "METIDA") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Comprometida"
        Else
            Call Compra_PO(i)
        End If

'BH
    Case InStr(Cells(i, 14).Value, "BH") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"

    Case InStr(Cells(i, 14).Value, "bh") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"
            
    Case InStr(Cells(i, 14).Value, "honorarios") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"
            
    Case InStr(Cells(i, 14).Value, "Honorarios") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"
    
    Case InStr(Cells(i, 14).Value, "HONORARIOS") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"
                    
    Case InStr(Cells(i, 14).Value, "honorario") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"
            
    Case InStr(Cells(i, 14).Value, "Honorario") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"
            
    Case InStr(Cells(i, 14).Value, "HONORARIO") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"
            
    Case InStr(Cells(i, 14).Value, "Bh") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"
    
    Case InStr(Cells(i, 14).Value, "bH") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Boleta Honorario"
            
'Infraestructura
    Case InStr(Cells(i, 14).Value, "Infraestructura") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
           
    Case InStr(Cells(i, 14).Value, "infraestructura") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
          
    Case InStr(Cells(i, 14).Value, "INFRAESTRUCTURA") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"

    Case InStr(Cells(i, 14).Value, "fraestructu") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
    
    Case InStr(Cells(i, 14).Value, "FRAESTRUCTU") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
    
'Negociado por COE
    Case InStr(Cells(i, 14).Value, "Negociado por COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"

    Case InStr(Cells(i, 14).Value, "negociado por COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "NEGOCIADO por COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "NEGOCIADO POR COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
    
    Case InStr(Cells(i, 14).Value, "negociado por coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "negociada por el coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "negociado por el coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "Negociado por el COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
                            
    Case InStr(Cells(i, 14).Value, "Negociad por el COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
    
    Case InStr(Cells(i, 14).Value, "negociad por el COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
                      
    Case InStr(Cells(i, 14).Value, "negociad por el coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
    
    Case InStr(Cells(i, 14).Value, "NEGOCIADO POR EL COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
 
    Case InStr(Cells(i, 14).Value, "NEGOC X COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
 
     Case InStr(Cells(i, 14).Value, "negoc X COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
 
     Case InStr(Cells(i, 14).Value, "negoc X coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"

    Case InStr(Cells(i, 14).Value, "Negociado por CoE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "negociado por CoE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "NEGOCIADO POR CoE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "Negociado por Coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
  
    Case InStr(Cells(i, 14).Value, "NEGOCIADO_POR_EL_COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "negociado_por_el_coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
                       
    Case InStr(Cells(i, 14).Value, "Negociado_por_el_Coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"

'Negociado por SASPA = EGOCIAD ,egociad
'Criterios complementarios: SSC, SC, SSA, SSAA, SA, S.A, ERVICIO NDINO, ervicio ndino, SSS, SS.AA., CSC, SASPA, SCC
            
     Case InStr(Cells(i, 14).Value, "NEG") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "NEG", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "COE") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "coe") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "CoE") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "Coe") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
        ElseIf InStr(Cells(i, Ultimate_Column + 1).Value, "SSC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SSA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SSAA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "S.A") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ERVICIO NDINO") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ervicio ndino") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SSS") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SS.AA.") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "CSC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SASPA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SCC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "NDINOS") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ndino") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ervicio") > 0 Then
            Cells(i, Ultimate_Column + 7) = "Negociado por SASPA"
            Cells(i, Ultimate_Column + 4) = "Compra"
            Cells(i, 14).Font.Bold = True 'negrita
            Cells(i, 14).Font.ColorIndex = 3 'pintar en rojo
            Cells(i, 27) = Cells(i, 1)
        Else
            Call Compra_PO(i)
        End If
            
    Case InStr(Cells(i, 14).Value, "neg") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "neg", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "COE") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "coe") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "CoE") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "Coe") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
        ElseIf InStr(Cells(i, Ultimate_Column + 1).Value, "SSC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SSA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SSAA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "S.A") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ERVICIO NDINO") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ervicio ndino") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SSS") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SS.AA.") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "CSC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SASPA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SCC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "NDINOS") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ndino") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ervicio") > 0 Then
            Cells(i, Ultimate_Column + 7) = "Negociado por SASPA"
            Cells(i, Ultimate_Column + 4) = "Compra"
            Cells(i, 14).Font.Bold = True 'negrita
            Cells(i, 14).Font.ColorIndex = 3 'pintar en rojo
            Cells(i, 27) = Cells(i, 1)
        Else
            Call Compra_PO(i)
        End If
            
    Case InStr(Cells(i, 14).Value, "Neg") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "Neg", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "COE") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "coe") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "CoE") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "Coe") > 0 Then
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
        ElseIf InStr(Cells(i, Ultimate_Column + 1).Value, "SSC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SSA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SSAA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "S.A") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ERVICIO NDINO") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ervicio ndino") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SSS") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SS.AA.") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "CSC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SASPA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "SCC") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "NDINOS") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ndino") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ervicio") > 0 Then
            Cells(i, Ultimate_Column + 7) = "Negociado por SASPA"
            Cells(i, Ultimate_Column + 4) = "Compra"
            Cells(i, 14).Font.Bold = True 'negrita
            Cells(i, 14).Font.ColorIndex = 3 'pintar en rojo
            Cells(i, 27) = Cells(i, 1)
        Else
            Call Compra_PO(i)
        End If

            
'Adjudicado por el COE
    Case InStr(Cells(i, 14).Value, "Adjudicado por el COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"

    Case InStr(Cells(i, 14).Value, "adjudicado por el COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
    Case InStr(Cells(i, 14).Value, "adjudicado por el coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"

    Case InStr(Cells(i, 14).Value, "ADJUDICADO por el COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
                        
     Case InStr(Cells(i, 14).Value, "ADJUDICADO POR COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
     Case InStr(Cells(i, 14).Value, "adjudicado por COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
      Case InStr(Cells(i, 14).Value, "adjudicado por coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
            
      Case InStr(Cells(i, 14).Value, "Adjudicado por coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"
    
    Case InStr(Cells(i, 14).Value, "Adjudicado por COE") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"

    Case InStr(Cells(i, 14).Value, "Adjudicado por Coe") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Negociado por COE"


'Addendum
    Case InStr(Cells(i, 14).Value, "Addendum") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
                        
    Case InStr(Cells(i, 14).Value, "addendum") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
              
    Case InStr(Cells(i, 14).Value, "ADDENDUM") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
               
    Case InStr(Cells(i, 14).Value, "Adendum") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
               
    Case InStr(Cells(i, 14).Value, "adendum") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"

    Case InStr(Cells(i, 14).Value, "ADENDUM") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
               
'Proyecto de gobierno.
    Case InStr(Cells(i, 14).Value, "Proyecto de gobierno") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Proyecto"
       
'--------------------------------------------------------------------------------------
'criterios agregados el 11-01
'Gastronomia
    Case InStr(Cells(i, 14).Value, "Gastronomia") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
           
    Case InStr(Cells(i, 14).Value, "GASTRONOMIA") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
    
    Case InStr(Cells(i, 14).Value, "GASTRONOM�A") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
          
    Case InStr(Cells(i, 14).Value, "gastronomia") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "gastronom�a") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "Gastronom�a") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"

'Medios
    Case InStr(Cells(i, 14).Value, "Medios") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
           
    Case InStr(Cells(i, 14).Value, "medios") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
          
    Case InStr(Cells(i, 14).Value, "MEDIOS") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
'Gobierno
    Case InStr(Cells(i, 14).Value, "Gobierno") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Gobierno"
           
    Case InStr(Cells(i, 14).Value, "gobierno") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Gobierno"
          
    Case InStr(Cells(i, 14).Value, "GOBIERNO") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Gobierno"

'Regularizacion agregados el 23-01

    Case InStr(Cells(i, 14).Value, "egulariza") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
            
    Case InStr(Cells(i, 14).Value, "EGULARIZA") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
    
    Case InStr(Cells(i, 14).Value, "Regularizacion") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
    
    Case InStr(Cells(i, 14).Value, "REGULARIZACION") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
    
    Case InStr(Cells(i, 14).Value, "regularizacion") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
    
    Case InStr(Cells(i, 14).Value, "REGULARIZACI�N") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
            
    Case InStr(Cells(i, 14).Value, "Regularizaci�n") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"

    Case InStr(Cells(i, 14).Value, "Regularizaci�n") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
            
    Case InStr(Cells(i, 14).Value, "regularizaci�n") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"

    Case InStr(Cells(i, 14).Value, "REGULARIZACI�N") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
                 
    Case InStr(Cells(i, 14).Value, "REGULARIZACI�") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
            
    Case InStr(Cells(i, 14).Value, "REGULARIZACON") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
            
    Case InStr(Cells(i, 14).Value, "Regularizai�n") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
            
        Case InStr(Cells(i, 14).Value, "REGULARZACION") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Regularizacion"
'--------------------------------------------------------------------------------------
'Suscripcion agregados el 05-02
    
    Case InStr(Cells(i, 14).Value, "suscripcion") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "Suscripcion") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "SUSCRIPCION") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
    
    Case InStr(Cells(i, 14).Value, "suscripci�n") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "Suscripci�n") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "SUSCRIPCI�N") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "SUSCRIPCI�NES") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "suscripci�nes") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "SUSCRIPCIONES") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "suscripciones") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "suscr") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"

    Case InStr(Cells(i, 14).Value, "SUSCRIPCI�N") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
'--------------------------------------------------------------------------------------
'Membresia agregados el 10-08
    
    Case InStr(Cells(i, 14).Value, "Membresia") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "MEMBRESIA") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "membresia") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "Membrecia") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "membrecia") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
     Case InStr(Cells(i, 14).Value, "MEMBRECIA") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"

'Investigacion agregados el 13-08
    
    Case InStr(Cells(i, 14).Value, "INVESTIGACION") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "INVESTIGACI�N") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "investigacion") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
            
    Case InStr(Cells(i, 14).Value, "investigaci�n") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
    Case InStr(Cells(i, 14).Value, "Investigacion") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
    
    Case InStr(Cells(i, 14).Value, "Investigaci�n") > 0
            Call Gestion_PO(i)
            Cells(i, Ultimate_Column + 7) = "Adendum"
   
'Emergencia agregados el 07-09
'Criterios complementarios , uces, UCES, �mpara, �MPARA, ampara, ampara, OTON, OTON, uz

            
    Case InStr(Cells(i, 14).Value, "Emergencia") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "Emergencia", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "UCES") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "uces") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "�mpara") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "�MPARA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "AMPARA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ampara") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "OTON") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "oton") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "uz") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "UZ") > 0 Then
            Call Compra_PO(i)
        Else
            Cells(i, Ultimate_Column + 7) = "Emergencia"
            Call Gestion_PO(i)
        End If

    Case InStr(Cells(i, 14).Value, "EMERGENCIA") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "EMERGENCIA", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "UCES") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "uces") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "�mpara") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "�MPARA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "AMPARA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ampara") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "OTON") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "oton") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "uz") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "UZ") > 0 Then
            Call Compra_PO(i)
        Else
            Cells(i, Ultimate_Column + 7) = "Emergencia"
            Call Gestion_PO(i)
        End If
                                           
    Case InStr(Cells(i, 14).Value, "emergencia") > 0
        Cells(i, Ultimate_Column + 1).Value = Replace(Cells(i, 14).Value, "emergencia", "", 1)
        If InStr(Cells(i, Ultimate_Column + 1).Value, "UCES") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "uces") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "�mpara") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "�MPARA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "AMPARA") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "ampara") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "OTON") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "oton") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "uz") > 0 Or InStr(Cells(i, Ultimate_Column + 1).Value, "UZ") > 0 Then
            Call Compra_PO(i)
        Else
            Cells(i, Ultimate_Column + 7) = "Emergencia"
            Call Gestion_PO(i)
        End If
    
    Case Else
            Call Compra_PO(i)
    
    End Select
        
    End If
Next i
End Sub
Sub tercera_revision()
Dim i As Integer
'Para compras catalogadas por Id articulo y Itm Id Vndr como catologo o contrato
'sean gestionadas por Gestion PO

For i = 6 To Ultimate_Row
  
If Cells(i, Ultimate_Column + 3).Value = "Catalogo" Or Cells(i, Ultimate_Column + 3).Value = "Contrato" Then
            Cells(i, Ultimate_Column + 4) = "Gestion PO"
            'Cells(i, 14).Font.Bold = True 'negrita
            'Cells(i, 14).Font.ColorIndex = 3 'pintar en rojo
        End If
Next i
End Sub
Sub limpiar2()
Dim Mensaje As Integer
'Mensaje As String
Mensaje = MsgBox("Limpiara todas las filas desde la fila 6 hasta la fila 8000, dejando el nombre d elas columnas donde esta la base", vbOKCancel)
If Mensaje = vbCancel Then
Exit Sub
End If
Worksheets("Base").Activate
Rows("6:9000").Delete
'Equipo = N�47 + 2
Cells(5, 50).Clear
Cells(5, 51).Clear
Cells(5, 52).Clear
Cells(5, 53).Clear
Cells(5, 54).Clear
Cells(5, 55).Clear
Cells(5, 1).Clear
Cells(5, 2).Clear
End Sub
Sub compra_realizada()
Dim i As Integer
'buscar si esta realziada la compra en la planilla PO que se obtiene de Query de PS
'https://www.exceltrick.com/formulas_macros/vlookup-in-vba/
Worksheets("Base").Activate
For i = 6 To Ultimate_Row
    If IsError(Application.VLookup(Cells(i, 1), Worksheets("PO").Range("A6:G4000"), 7, False)) = False Then
            Cells(i, Ultimate_Column + 1) = "OC Realizada"
            Cells(i, Ultimate_Column + 1).Font.Bold = True 'negrita
            Cells(i, Ultimate_Column + 1).Font.ColorIndex = 5 'color azul
    End If
Next i
End Sub
'este es de comprobacion solamente
Sub avance_PO()
Dim g As Integer
'https://www.exceltrick.com/formulas_macros/vlookup-in-vba/
Worksheets("Base").Activate

Cells(5, Ultimate_Column + 5) = "Compra realizada"
            Cells(5, Ultimate_Column + 5).Font.Bold = True 'negrita

For g = 6 To Ultimate_Row
    If IsError(Application.VLookup(Cells(g, 1), Worksheets("PO").Range("A6:G4000"), 7, False)) = False Then
            Cells(g, Ultimate_Column + 5) = "OC Realizada"
            Cells(g, Ultimate_Column + 5).Font.Bold = True 'negrita
            Cells(g, Ultimate_Column + 5).Font.ColorIndex = 17 'color azulado
    Else
            Cells(g, Ultimate_Column + 5) = "OC No Realizada"
            Cells(g, Ultimate_Column + 5).Font.Bold = True 'negrita
            Cells(g, Ultimate_Column + 5).Font.ColorIndex = 3 'color rojo
    
    End If
Next g
End Sub

Sub asignado_ayer()
Dim j As Integer
'busca si esta asignado en "Ayer"
'Si el IsError = False ==> encontro un valor.

'https://www.exceltrick.com/formulas_macros/vlookup-in-vba/
Worksheets("Base").Activate
For j = 6 To Ultimate_Row

    If IsError(Application.VLookup(Cells(j, 1), Worksheets("Ayer").Range("A2:F3000"), 6, False)) = False Then
            Cells(j, Ultimate_Column + 1) = Application.WorksheetFunction.VLookup(Cells(j, 1), Worksheets("Ayer").Range("A2:F3000"), 6, False)
            Cells(j, Ultimate_Column + 1).Font.Bold = True 'negrita
            Cells(j, Ultimate_Column + 1).Font.ColorIndex = 3 'color rojo
    End If
Next j
End Sub
'https://excelmacromastery.com/vba-dictionary/
'https://www.experts-exchange.com/articles/3391/Using-the-Dictionary-Class-in-VBA.html
'para contar las lineas de cada PO, esto es para ver cuantas PO hay en el listado
'https://stackoverflow.com/questions/915317/does-vba-have-dictionary-structure
'https://github.com/VBA-tools/VBA-Dictionary/blob/master/Dictionary.cls
Sub countPO2()

Dim ws As Worksheet
Dim lastRow As Long, x As Long
Dim items As Object

Application.ScreenUpdating = False
  
Set ws = Worksheets("Base")
Cells(5, Ultimate_Column + 6).Value = "Cantidad de lineas" 'Donde va el titulo
            Cells(5, Ultimate_Column + 6).Font.Bold = True 'Donde va el titulo en negrita
    
lastRow = ws.Range("D" & Rows.Count).End(xlUp).Row 'conteo de columna
    
    Set items = CreateObject("Scripting.Dictionary")
    For x = 6 To lastRow
        If Not items.Exists(ws.Cells(x, 1).Value) Then 'columna de conteo columna 1
            items.Add ws.Cells(x, 1).Value, 1 'columna de conteo columna 1
            ws.Cells(x, Ultimate_Column + 6).Value = items(ws.Cells(x, 1).Value) 'columna donde deja = columna de conteo columna 1
        Else
            items(ws.Cells(x, 1).Value) = items(ws.Cells(x, 1).Value) + 1 'columna de conteo columna 1 = columna de conteo columna 1 + 1
            ws.Cells(x, Ultimate_Column + 6).Value = items(ws.Cells(x, 1).Value) 'columna donde deja = columna de conteo columna 1
        End If
    Next x
End Sub
Sub Gestion_PO(ByVal i As Integer)
            Cells(i, Ultimate_Column + 4) = "Gestion PO"
            Cells(i, 14).Font.Bold = True 'negrita
            Cells(i, 14).Font.ColorIndex = 3 'pintar en rojo
            Cells(i, 27) = Cells(i, 1)
End Sub
Sub Compra_PO(ByVal i As Integer)
            Cells(i, Ultimate_Column + 4) = "Compra"
            Cells(i, Ultimate_Column + 7) = "Sourcing"

End Sub
Sub Clasificacion_Compras()
Dim k As Integer

Worksheets("Base").Activate
For k = 6 To Ultimate_Row 'Ultimate_Row

Cells(5, 56).Value = "Clasificacion categoria" 'Donde va el titulo
            Cells(5, 56).Font.Bold = True 'Donde va el titulo en negrita
            
    If IsError(Application.VLookup(Cells(k, 50), Worksheets("data").Range("H2:I30"), 2, False)) = False Then
        Cells(k, 56) = Application.VLookup(Cells(k, 50), Worksheets("data").Range("H2:I30"), 2, False)
    Else
        Cells(k, 56) = "Sin clasificacion"
    End If
Next k
End Sub



   

