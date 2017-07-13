'Author: Humberto Della Torre
'Funciones de ayuda
Option Explicit
Private Const MAXLONG = 2147483648# 'Maximo en NAV
Private Const MINLONG = -2147483648# 'Minimo en NAV
Private Const NUMREF = 1 'Numero de referencias

'https://stackoverflow.com/questions/22542834/how-to-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-and-loops
Function regEx(strInput As String, matchPattern As String, Optional ByVal outputPattern As String = "$0") As Variant
    Dim inputRegexObj As New VBScript_RegExp_55.RegExp, outputRegexObj As New VBScript_RegExp_55.RegExp, outReplaceRegexObj As New VBScript_RegExp_55.RegExp
    Dim inputMatches As Object, replaceMatches As Object, replaceMatch As Object
    Dim replaceNumber As Integer

    With inputRegexObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = matchPattern
    End With
    With outputRegexObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = "\$(\d+)"
    End With
    With outReplaceRegexObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
    End With

    Set inputMatches = inputRegexObj.Execute(strInput)
    If inputMatches.Count = 0 Then
        regEx = False
    Else
        Set replaceMatches = outputRegexObj.Execute(outputPattern)
        For Each replaceMatch In replaceMatches
            replaceNumber = replaceMatch.SubMatches(0)
            outReplaceRegexObj.Pattern = "\$" & replaceNumber

            If replaceNumber = 0 Then
                outputPattern = outReplaceRegexObj.Replace(outputPattern, inputMatches(0).Value)
            Else
                If replaceNumber > inputMatches(0).SubMatches.Count Then
                    'regex = "A to high $ tag found. Largest allowed is $" & inputMatches(0).SubMatches.Count & "."
                    regEx = CVErr(xlErrValue)
                    Exit Function
                Else
                    outputPattern = outReplaceRegexObj.Replace(outputPattern, inputMatches(0).SubMatches(replaceNumber - 1))
                End If
            End If
        Next
        regEx = outputPattern
    End If
End Function

Function dateFormulaBoundCheck(exp As String) As String

    'Variables
    Dim regexStr As String: regexStr = LCase(exp) 'Cambiamos la cadena a minusculas
    Dim regexArr() As String: regexArr = Split(Replace(Replace(regexStr, "+", "|"), "-", "|"), "|") 'dividimos la cadena en subexpressiones
    Dim subExpStr As Variant 'Para iterar por el array
    Dim expUnit As String '<unit> de la expression
    Dim expNum As Integer 'numero de la expression
    Dim err As Boolean: err = False 'valor de retorno, solo puede cambiar a True si en algun caso encuentra un error
    Dim number As String: number = "[0]*[0-9]{1,4}" 'expresion regular de numero
    Dim unit As String: unit = "(wd|ds|w|m|q|y|d|s|t|a)" 'expression regular de unidades
    
    For Each subExpStr In regexArr 'por cada subexpresion en el array
        If CStr(regEx((subExpStr), unit & number)) = subExpStr Then 'Para las subexpresiones <unit><number>
            expUnit = CStr(regEx((subExpStr), unit)) 'extraemos el <unit>
            expNum = CInt(CStr(regEx((subExpStr), number))) 'extraemos el <number>
            Select Case expUnit 'Dependiendo de que unit el caso
                Case "wd", "ds"
                    If expNum > 7 Or expNum = 0 Then 'Hay 7 dias de semanas
                        err = True
                    End If
                Case "d"
                    If expNum > 31 Or expNum = 0 Then 'Hay maximo 31 dias en un mes
                        err = True
                    End If
                Case "w", "s"
                    If expNum > 53 Or expNum = 0 Then 'Hay 53 semanas en un año
                        err = True
                    End If
                Case "m"
                    If expNum > 12 Or expNum = 0 Then 'Hay 12 meses en un año
                        err = True
                    End If
                Case "q", "t"
                    If expNum > 4 Or expNum = 0 Then 'Hay 4 trimestres en un año
                        err = True
                    End If
                Case "y", "a"
                    If expNum > 99 Then 'El valor de año llega hasta 99
                        err = True
                    End If
            End Select
        End If
    Next subExpStr
    
    dateFormulaBoundCheck = err 'Este valor empieza false y cambia a True solo si en una o mas sub expresiones hubo un problema
End Function

Function isInteger(entero As String) As String
    
    isInteger = "El valor en la celda no es numerico. El formato correcto usa puntos como separador de millares y coma como separador decimal." 'Mensaje si no es numerico
    If IsNumeric(entero) Then 'Revisar si es numerico
        isInteger = "El valor en la celda no es un entero. Despues de una coma solo pueden haber 0s." 'Mensaje si no es entero
        If CDbl(entero) = 0 Then GoTo Salto 'Si es 0 saltamos lo siguiente
        If Int(entero) / CDbl(entero) = 1 Then 'Revisamos si es entero
Salto:
            isInteger = "El valor en la celda excede el rango [-2147483647,2147483647]." 'Mensaje si esta fuera del rango de entero
            If CDbl(entero) > MINLONG And CDbl(entero) < MAXLONG Then 'Revisa si el valor esta en el rango correcto
                isInteger = "" 'Cadena vacia si no hay problemas
            End If
        End If
    End If
    
End Function

Function isDecimal(entero As String) As String

    isDecimal = "El valor en la celda no es numerico. El formato correcto usa puntos como separador de millares y coma como separador decimal." 'Mensaje si no es numerico
    If IsNumeric(entero) Then 'Revisar si es numerico
        isDecimal = "" 'Cadena vacia si no hay problemas
    End If
    
End Function

Function inCol(col As Collection, str As String) As Boolean
    Dim i As Long 'Iterador
    
    inCol = False 'Asumption
    For i = 1 To col.Count 'Por cada valor
        If col.Item(i) = str Then 'Si encontramos un match
            inCol = True 'Cambiamos asumption
            Exit Function 'Salimos de la funcion
        End If
    Next i
    
End Function

Function tRelation(campo As String) As String
    Dim relaciones(0 To NUMREF, 0 To 1) As String '2D Array ***Maybe change to jagged for better performance
    Dim iOuter As Integer 'Iterador (iOuter,X)
    
    relaciones(0, 0) = "Nº": relaciones(0, 1) = "testEmptyTable.xlsm|CatalogoDeCuenta[Nº]" 'Incrementar public const NUMREF y seguir el mismo formato para agregar dependencias
    'relaciones(1, 0) = "campo": relaciones(1, 1) = "archivoExcel.xlsm|nombretabla[campo]"
    '...
    relaciones(NUMREF, 0) = "": relaciones(NUMREF, 1) = ""
    
    For iOuter = 0 To NUMREF 'Por cada valor
        If relaciones(iOuter, 0) = campo Then GoTo found 'Si el campo tiene una referencia salimos a found
    Next iOuter
    
    tRelation = "" 'No se encontro ninguna referencia, regresamos cadena vacia
    Exit Function

found:
    tRelation = relaciones(iOuter, 1) 'Se encontro una referencia, regresamos la cadena
    
End Function
