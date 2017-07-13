'Author: Humberto Della Torre
'Contenido de ValidacionCampos
Option Explicit
Private Const ERRCOLOR = 22 'Rojo
Private Const EMPCOLOR = 37 'Azul

Sub ValidacionText(columna As Range)
    Dim celda As Range 'Rango de celda
    Dim maxLen As Integer: maxLen = Int(Replace((columna.Rows(1).Comment.Text), "Text", "")) 'Maxima cantidad de caracteres
    Dim errCount As Integer: errCount = 0 'Cuenta de errores
        
    For Each celda In columna.Rows.Offset(1).Resize(columna.Rows.Count - 1) 'Por cada celda en la columna, excluyendo la primera
    
        If celda.Value = "" Then 'Si esta vacia
            celda.Interior.ColorIndex = EMPCOLOR 'Cambiamos el color de celda para indicar que esta vacia
            celda.ClearComments 'Borramos comentarios
            
        ElseIf Len(celda.Value) > maxLen Then 'Si el valor en la celda execede la maxima cantidad de caracteres
            errCount = errCount + 1 'Incrementamos la cuenta de errores
            celda.Interior.ColorIndex = ERRCOLOR 'Cambiamos el color de celda para indicar que hay un problema
            If celda.Comment Is Nothing Then celda.AddComment 'Si no tenia comentarios la celda se la agregamos
            celda.Comment.Text "Este valor usa " & Len(celda.Value) & " caracteres y el valor maximo es " & maxLen & "." 'Ingresamos un comentario informativo
        
        ElseIf Not celda.Interior.ColorIndex = xlNone Then 'Si ya no tiene problemas
            celda.Interior.ColorIndex = xlNone 'Revertimos el color
            celda.ClearComments 'Borramos comentarios
        
        End If
    Next celda
            
    'Mensaje informativo para indicar la cantidad de errores en la columna
    If errCount > 1 Then
        MsgBox ("El campo " & columna.Rows(1).Value & " tiene " & errCount & " problemas.")
    ElseIf errCount = 1 Then
        MsgBox ("El campo " & columna.Rows(1).Value & " tiene " & errCount & " problemas.")
    End If
            
End Sub

Sub ValidacionOption(columna As Range)
    Dim celda As Range 'Rango de celda
    Dim errCount As Integer: errCount = 0 'Cuenta de errores
    Dim optionsArr() As String 'Lista de opciones
    Dim i As Integer 'Iterador

    If Left(columna.Rows(1).Comment.Text, 6) = "Option" Then 'Si el comentario es opciones ocupamos conseguir los detalles de las opciones
        optionsArr = Split(columna.Rows(1).Comment.Text, vbLf) 'Dividimos el comentario en un array {"Option","0: Opcion0",....,"N: OpcionN"}
        For i = 1 To UBound(optionsArr) 'Para cada valor en el array
            optionsArr(i - 1) = LCase(optionsArr(i)) 'Movemos todos los valores uno atras {"0: Opcion0",...,"N: OpcionN", "N: OpcionN"}
        Next i
        ReDim Preserve optionsArr(UBound(optionsArr) - 1) 'Quitamos el ultimo valor ya que es duplicado {"0: Opcion0",...,"N: OpcionN"}
        For i = 0 To UBound(optionsArr) 'Para cada valor que queda en el array
            optionsArr(i) = (Right(optionsArr(i), Len(optionsArr(i)) - 3)) 'Eliminamos los primeros tres caracteres "N: " {"Opcion0",...,"OpcionN"}
        Next i
    ElseIf columna.Rows(1).Comment.Text = "Boolean" Then 'Si es boolean
        optionsArr = Split("false no 0 true sí 1") 'Las opciones son pueden ser de las siguientes
    End If
            
    For Each celda In columna.Rows.Rows.Offset(1).Resize(columna.Rows.Count - 1) 'Por cada celda en la columna, excluyendo la primera
    
        If celda.Value = "" Then 'Si esta vacia
            celda.Interior.ColorIndex = EMPCOLOR 'Cambiamos el color de celda para indicar que esta vacia
            celda.ClearComments 'Borramos comentarios
            
        ElseIf InStr("|" & Join(optionsArr, "|"), "|" & LCase(celda.Value)) = 0 And columna.Rows(1).Comment.Text = "Boolean" Then 'Si no es boolean y no es una de las opciones
            errCount = errCount + 1 'Incrementamos la cuenta de errores
            celda.Interior.ColorIndex = ERRCOLOR 'Cambiamos el color de celda para indicar que tiene un error
            If celda.Comment Is Nothing Then celda.AddComment 'Si no tiene un comentario agregamos uno
            celda.Comment.Text "El valor no es un boolean válido. Los siguientes son ejemplos válidos de booleans: " & Join(optionsArr, ", ") & "." 'Mensaje informativo en el comentario
        
        ElseIf InStr("|" & Join(optionsArr, "|") & "|", "|" & LCase(celda.Value) & "|") = 0 And Left(columna.Rows(1).Comment.Text, 6) = "Option" Then 'Si es option y no es una de las opciones
            errCount = errCount + 1 'Incrementamos la cuenta de errores
            celda.Interior.ColorIndex = ERRCOLOR 'Cambiamos el color de celda para indicar que tiene un error
            If celda.Comment Is Nothing Then celda.AddComment 'Si no tiene un comentario agregamos uno
            celda.Comment.Text "El valor no es una opcion válida. Las siguientes son las opciones válidas: " & Join(optionsArr, ", ") & "." 'Mensaje informativo en el comentario
        
        ElseIf Not celda.Interior.ColorIndex = xlNone Then 'Si ya no tiene problemas
            celda.Interior.ColorIndex = xlNone 'Revertimos el color
            celda.ClearComments 'Borramos comentarios
            
        End If
    Next celda
            
    'Mensaje informativo para indicar la cantidad de errores y en que campos
    If errCount > 1 Then
        MsgBox ("El campo " & columna.Rows(1).Value & " tiene " & errCount & " problemas.")
    ElseIf errCount = 1 Then
        MsgBox ("El campo " & columna.Rows(1).Value & " tiene " & errCount & " problema.")
    End If
    
End Sub

Sub ValidacionDateFormula(columna As Range)
    Dim celda As Range 'Rango de celda
    Dim errCount As Integer: errCount = 0 'Cuenta de errores
    
    Dim sign As String: sign = "[+-]?" '<Sign> = +|-
    Dim number As String: number = "[0]*[0-9]{1,4}" '<Number> = Positive integer
    Dim enUnit As String: enUnit = "(wd|d|w|m|q|y)" '<Unit> ingles
    Dim esUnit As String: esUnit = "(ds|d|s|m|t|a)" '<Unit> espanol
    Dim enPrefix As String: enPrefix = "[c]" '<Prefix> ingles
    Dim esPrefix As String: esPrefix = "[p]" '<Prefix> espanol
    Dim enTerm As String: enTerm = "(" & enUnit & number & "|" & number & enUnit & "|" & enPrefix & enUnit & ")" '<Term> = <Number><Unit>|<Unit><Number>|<Prefix><Unit> ingles
    Dim esTerm As String: esTerm = "(" & esUnit & number & "|" & number & esUnit & "|" & esPrefix & esUnit & ")" '<Term> espanol
    Dim enSubExp As String: enSubExp = "(" & sign & enTerm & ")" '<SubExpression> [<Sign>]<Term> ingles
    Dim esSubExp As String: esSubExp = "(" & sign & esTerm & ")" '<SubExpression> espanol
    
    Dim dfRegEx As String: dfRegEx = enSubExp & "+|" & esSubExp & "+" 'Date expression = [<SubExpression>][<SubExpression>]...
               
    For Each celda In columna.Rows.Offset(1).Resize(columna.Rows.Count - 1) 'Por cada celda en la columna, excluyendo la primera
        
        If celda.Value = "" Then 'Si esta vacia
            celda.Interior.ColorIndex = EMPCOLOR 'Cambiamos el color de celda para indicar que esta vacia
            celda.ClearComments 'Borramos comentarios
        
        ElseIf Not CStr(regEx(LCase(celda.Value), dfRegEx)) = LCase(celda.Text) Or dateFormulaBoundCheck(LCase(celda.Value)) Or Len(celda.Value) > 32 Then 'Si no cumple con el formato
            errCount = errCount + 1 'Incrementamos la cuenta de errores
            celda.Interior.ColorIndex = ERRCOLOR 'Cambiamos el color de celda para indicar error
            If celda.Comment Is Nothing Then celda.AddComment 'Si no tenia comentarios agregamos
            celda.Comment.Text "Para referencia del formato DateFormula consulte el siguiente link https://msdn.microsoft.com/en-us/library/dd301368.aspx" 'Comentario informativo
        
        ElseIf Not celda.Interior.ColorIndex = xlNone Then 'Si ya no tiene problemas
            celda.Interior.ColorIndex = xlNone 'Revertir color
            celda.ClearComments 'Borramos comentarios
            
        End If
    Next celda
            
    'Mensaje informativo para indicar la cantidad de errores y en que campos
    If errCount > 1 Then
        MsgBox ("El campo " & columna.Rows(1).Value & " tiene " & errCount & " problemas.")
    ElseIf errCount = 1 Then
        MsgBox ("El campo " & columna.Rows(1).Value & " tiene " & errCount & " problema.")
    End If

End Sub

Sub ValidacionInteger(columna As Range)
    Dim celda As Range 'Celda
    Dim errCount As Integer: errCount = 0 'Cuenta de errores
            
    For Each celda In columna.Rows.Offset(1).Resize(columna.Rows.Count - 1) 'Por cada celda en la columna, excluyendo la primera
    
        If celda.Value = "" Then 'Si esta vacia
            celda.Interior.ColorIndex = EMPCOLOR 'Cambiamos el color de celda para indicar que esta vacia
            celda.ClearComments 'Borramos comentarios
        
        ElseIf isInteger(celda.Value) <> "" Then 'Si el valor en la celda no es un entero, la funcion regresa una cadena con el problem o una cadena vacia para indicar que no hay problema
            errCount = errCount + 1 'Incrementamos la cuenta de errores
            celda.Interior.ColorIndex = ERRCOLOR 'Cambiamos el color de celda para indicar que hay un problema
            If celda.Comment Is Nothing Then celda.AddComment 'Si no tenia comentarios la celda se la agregamos
            celda.Comment.Text isInteger(celda.Value) 'Ingresamos un comentario informativo
        
        ElseIf Not celda.Interior.ColorIndex = xlNone Then 'Si ya no tiene problemas
            celda.Interior.ColorIndex = xlNone 'Revertir color
            celda.ClearComments 'Borramos comentarios
            
        End If
    Next celda
            
    'Mensaje informativo para indicar la cantidad de errores y en que campos
    If errCount > 1 Then
        MsgBox ("El campo " & columna.Rows(1).Value & " tiene " & errCount & " problemas.")
    ElseIf errCount = 1 Then
        MsgBox ("El campo " & columna.Rows(1).Value & " tiene " & errCount & " problemas.")
    End If
            
End Sub

Sub ValidacionDecimal(columna As Range)
    Dim celda As Range 'Celda
    Dim errCount As Integer: errCount = 0 'Cuenta de errores
            
    For Each celda In columna.Rows.Offset(1).Resize(columna.Rows.Count - 1) 'Por cada celda en la columna, excluyendo la primera
                
        If celda.Value = "" Then 'Si esta vacia
            celda.Interior.ColorIndex = EMPCOLOR 'Cambiamos el color de celda para indicar que esta vacia
            celda.ClearComments 'Borramos comentarios
            
        ElseIf isDecimal(celda.Value) <> "" Then 'Si el valor en la celda no es un decimal, funcion regresa cadena con problema o una cadena vacia si no hay problemas
            errCount = errCount + 1 'Incrementamos la cuenta de errores
            celda.Interior.ColorIndex = ERRCOLOR 'Cambiamos el color de celda para indicar que hay un problema
            If celda.Comment Is Nothing Then celda.AddComment 'Si no tenia comentarios la celda se la agregamos
            celda.Comment.Text isDecimal(celda.Value) 'Ingresamos un comentario informativo
        
        ElseIf Not celda.Interior.ColorIndex = xlNone Then 'Si ya no tiene problemas
            celda.Interior.ColorIndex = xlNone 'Revertir color
            celda.ClearComments 'Borramos comentarios
        
        End If
    Next celda
            
    'Mensaje informativo para indicar la cantidad de errores y en que campos
    If errCount > 1 Then
        MsgBox ("El campo " & columna.Rows(1).Value & " tiene " & errCount & " problemas.")
    ElseIf errCount = 1 Then
        MsgBox ("El campo " & columna.Rows(1).Value & " tiene " & errCount & " problemas.")
    End If
            
End Sub

Sub ValidacionDate(columna As Range)
    Dim celda As Range 'Celda
    Dim errCount As Integer: errCount = 0 'Cuenta de errores
    Dim dateStr As String
        
    For Each celda In columna.Rows.Offset(1).Resize(columna.Rows.Count - 1) 'Por cada celda en la columna, excluyendo la primera
                
        If celda.Value = "" Then 'Si esta vacia
            celda.Interior.ColorIndex = EMPCOLOR 'Cambiamos el color de celda para indicar que esta vacia
            celda.ClearComments 'Borramos comentarios
                    
        ElseIf Not isDate(celda.Text) Then 'Si el valor en la celda no es una fecha
            errCount = errCount + 1 'Incrementamos la cuenta de errores
            celda.Interior.ColorIndex = ERRCOLOR 'Cambiamos el color de celda para indicar que hay un problema
            If celda.Comment Is Nothing Then celda.AddComment 'Si no tenia comentarios la celda se la agregamos
                celda.Comment.Text "El valor en la celda no es una fecha valida o tiene un formato incorrecto, pruebe con el siguiente formato: aaaa-mm-dd." 'Ingresamos un comentario informativo
        
        ElseIf isDate(celda.Text) Then 'Si es fecha
            celda.ClearComments 'Removemos comentarios
            celda.Interior.ColorIndex = xlNone 'Remover color para indicar que esta correcto
            celda.Value = Format(celda.Value, "yyyy-mm-dd") 'Cambiar el formato de fecha a aaaa-mm-dd
            celda.NumberFormat = "yyyy-mm-dd" 'Cambiar el formato de fecha
            dateStr = celda.Text 'Guardamos el valor del campo
            celda.NumberFormat = "@" 'Cambiamos el formato a Texto
            celda.Value = dateStr 'Ingresamos el valor
        
        End If
    Next celda
            
    'Mensaje informativo para indicar la cantidad de errores y en que campos
    If errCount > 1 Then
        MsgBox ("El campo " & columna.Rows(1).Value & " tiene " & errCount & " problemas.")
    ElseIf errCount = 1 Then
        MsgBox ("El campo " & columna.Rows(1).Value & " tiene " & errCount & " problema.")
    End If
            
End Sub

Sub ValidacionCode(columna As Range)
    Dim celda As Range 'Rango de celda
    Dim errCount As Integer: errCount = 0 'Cuenta de errores
    Dim maxLen As Integer: maxLen = Int(Replace((columna.Rows(1).Comment.Text), "Code", "")) 'Maxima cantidad de caracteres
    Dim relacion As String: relacion = tRelation(columna.Rows(1))
    
    Dim src As Workbook
    Dim clm As Range
    Dim cld As Range
    Dim col As New Collection
    
    If relacion <> "" Then 'Si hay una relacion
        Set src = Workbooks.Open(Application.ActiveWorkbook.Path & "\" & Split(relacion, "|")(0), True, True) 'Abrimos archivo de Excel
        
        Set clm = (src.Worksheets(1).Range(Split(relacion, "|")(1))) 'Conseguimos el rango de la columna de la tabla especificada
        For Each cld In clm.Rows 'Por cada valor
            col.Add cld.Text 'Agregamos los valores en la columna a una coleccion
        Next cld
        
        src.Close 'Cerramos el archivo Excel
    End If
    
    For Each celda In columna.Rows.Offset(1).Resize(columna.Rows.Count - 1) 'Por cada celda en la columna, excluyendo la primera
        
        If celda.Value = "" Then 'Si esta vacia
            celda.Interior.ColorIndex = EMPCOLOR 'Cambiamos el color de la celda para indicarlo
            celda.ClearComments 'Borramos comentarios
            
        ElseIf Len(celda.Value) > maxLen Then 'Si el valor en la celda execede la maxima cantidad de caracteres
            errCount = errCount + 1 'Incrementamos la cuenta de errores
            celda.Interior.ColorIndex = ERRCOLOR 'Cambiamos el color de celda para indicar que hay un problema
            If celda.Comment Is Nothing Then celda.AddComment 'Si no tenia comentarios la celda se los agregamos
            celda.Comment.Text "Este valor usa " & Len(celda.Value) & " caracteres y el valor maximo es " & maxLen & "." 'Ingresamos un comentario informativo

        ElseIf relacion <> "" And Not inCol(col, celda.Value) Then 'Si existe una relacion y el valor no pertenece a la coleccion
            errCount = errCount + 1 'Incrementamos la cuenta de errores
            celda.Interior.ColorIndex = ERRCOLOR 'Cambiamos el color de celda para indicar que hay un problema
            If celda.Comment Is Nothing Then celda.AddComment 'Si no tenia comentarios la celda se los agregamos
            celda.Comment.Text "Tiene que ser uno de los valores bajo el campo " & Split(Split(Split(relacion, "|")(1), "[")(1), "]")(0) & " en el archivo " & Split(relacion, "|")(0) & "." 'Ingresamos un comentario informativo
        
        ElseIf celda.Interior.ColorIndex <> xlNone Then
            celda.Interior.ColorIndex = xlNone 'Revertir color
            celda.ClearComments 'Borramos comentarios
        
        End If
siguiente:
    Next celda
    
    'Mensaje informativo para indicar la cantidad de errores y en que campos
    If errCount > 1 Then
        MsgBox ("El campo " & columna.Rows(1).Value & " tiene " & errCount & " problemas.")
    ElseIf errCount = 1 Then
        MsgBox ("El campo " & columna.Rows(1).Value & " tiene " & errCount & " problema.")
    End If
    
End Sub

Sub ValidacionCampos()
    'Declaracion de variables
    Dim tabla As ListObject 'Tabla
    Dim columna As Range 'Rango de columna
    Set tabla = ActiveSheet.ListObjects(1) 'Obtenemos referencia a la tabla, *** Pasar nombre de tabla? Siempre sera ListObjects(1)
    'Debug.Print TypeName(ActiveSheet.ListObjects(1))
    For Each columna In tabla.Range.Columns 'Por cada columna en la tabla
    
        If Left(columna.Rows(1).Comment.Text, 4) = "Text" Then 'Si el comentario indican Text#
            Call ValidacionText(columna) 'Validacion Text#
        ElseIf Left(columna.Rows(1).Comment.Text, 4) = "Code" Then 'Si el comentario indica Code#
            Call ValidacionCode(columna) 'Validacion Code#
        ElseIf Left(columna.Rows(1).Comment.Text, 6) = "Option" Or columna.Rows(1).Comment.Text = "Boolean" Then 'Si el comentario indica Option
            Call ValidacionOption(columna) 'Validacion Option
        ElseIf columna.Rows(1).Comment.Text = "DateFormula" Then 'Si el comentario indica DateFormula
            Call ValidacionDateFormula(columna) 'Validacion DateFormula
        ElseIf columna.Rows(1).Comment.Text = "Date" Then 'Si el comentario indica validacion date
            Call ValidacionDate(columna) 'Validacion Date
        ElseIf columna.Rows(1).Comment.Text = "Integer" Then 'Si el comentario indica Integer
            Call ValidacionInteger(columna) 'Validacion Integer
        ElseIf columna.Rows(1).Comment.Text = "Decimal" Then 'Si el comentario indica Decimal
            Call ValidacionDecimal(columna) 'Validacion Decimal
        Else 'Si el campo no incluye uno de los casos definidos
            Debug.Print ("El campo " & columna.Rows(1) & " no puede ser validado.") 'Para revisar si en la plantilla hay algun problema, uso interno
        End If
    Next columna
End Sub
