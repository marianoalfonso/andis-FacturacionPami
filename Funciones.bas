Attribute VB_Name = "Funciones"
'FUNCIONES:
'Setear_Configuracion
'ConfigGet
'EnviarLog
'Eliminar_Archivo
'BorrarArchivo
'MoveFile
'Copiar_Archivo
'Validar_Expresion
'Validar_Periodo
'Completar_Cadena
'ConvertirDecimal
'Control_Ultimo_Archivo
'Obtener_ProximoArchivo_Exportacion
'Incrementar_Archivo_Exportacion
'Obtener_ProximoArchivo_Importacion
'Incrementar_Archivo_Importacion
'ArchivoExiste
'Verificar_Existencia_Archivo
'Actualizar_Config_INI
'Validar_Cuit
'Validar_TextBoxNulo


Public Sub Setear_Configuracion()
'PROCEDIMIENTO: Setear_Configuracion
'FECHA DE CREACION: 27 de febrero de 2007
'AUTOR: Mariano Alfonso (NetworkIt)
'DESCRIPCION: Setea las variables del sistema en base a la lectura de
'             un archivo .INI parametrizable
'PARAMETROS: ()
'DEVOLUCION: ()
    Dim tmpArchivo As String
'    C:\Facturacion Pami\Facturaciones Cerradas
    tmpArchivo = Trim(App.Path & "\Facturaciones Cerradas\default.txt")
    If ArchivoExiste(tmpArchivo) Then
        Kill tmpArchivo
    End If

    Gstr_Cuit = ConfigGet("DATOS_PRESTADOR", "Cuit")
    Gstr_RazonSocial = ConfigGet("DATOS_PRESTADOR", "RazonSocial")
    Gstr_TipoPrestador = ConfigGet("DATOS_PRESTADOR", "TipoPrestador")
    Gstr_Provincia = ConfigGet("DATOS_PRESTADOR", "NombreProvincia")
    
    Gstr_DestinationPath = ConfigGet("DATOS_CONFIGURACION", "DestinationPath")
End Sub


Public Function ConfigGet(strSection As String, strItem As String) As String
'FUNCION: ConfigGet
'FECHA DE CREACION: 27 de febrero de 2007
'AUTOR: Mariano Alfonso (NetworkIt)
'DESCRIPCION: Dado un encabezado de grupo y un valor, devuelve el
'             resultado obtenido de un archivo externo de configuracion
'             parametrizable.
'PARAMETROS: strSelection --> Cabecera de grupo (String)
'            strItem      --> Item a buscar (String)
'DEVOLUCION: (String)

On Error GoTo TrapError
Dim intFile As Integer
Dim Aux As String
Dim bytFound As Byte
  
ConfigGet = ""
intFile = FreeFile  'Toma el próximo archivo disponible

Open App.Path & "\Config.ini" For Input As #intFile
'Open App.Path & "\Config.txt" For Input As #intFile

Do While Not EOF(intFile)       'Busca strSection pasada en el archivo
  Line Input #intFile, Aux
  bytFound = InStr(Aux, strSection)
  If bytFound > 0 Then Exit Do
Loop

If bytFound > 0 Then    'Si encontro la cadena
  bytFound = 0
  Do While Not EOF(intFile)
    Line Input #intFile, Aux
    bytFound = InStr(Aux, strItem)  'InStr Busca strItem en el archivo
    If bytFound > 0 Then
      ConfigGet = Mid(Aux, InStr(Aux, "=") + 1) 'Extrae el resultado
      Exit Do
    End If
  Loop
End If

Close #intFile
  
Exit Function
TrapError:
Err.Clear
Exit Function
End Function


'SE PASA COMO PARAMETRO EL NOMBRE DEL ARCHIVO
Public Function ConfigGetEXTENDIDO(strArchivo As String, strSection As String, strItem As String, Optional Flag As Integer) As String
'FUNCION: ConfigGet
'FECHA DE CREACION: 27 de febrero de 2007
'AUTOR: Mariano Alfonso (NetworkIt)
'DESCRIPCION: Dado un encabezado de grupo y un valor, devuelve el
'             resultado obtenido de un archivo externo de configuracion
'             parametrizable.
'PARAMETROS: strSelection --> Cabecera de grupo (String)
'            strItem      --> Item a buscar (String)
'DEVOLUCION: (String)

On Error GoTo TrapError
Dim intFile As Integer
Dim Aux As String
Dim bytFound As Byte

strArchivo = Trim(strArchivo)

ConfigGetEXTENDIDO = ""
intFile = FreeFile  'Toma el próximo archivo disponible

If Flag = 1 Then
    Open App.Path & "\" & strArchivo For Input As #intFile
  Else
    Open App.Path & "\Facturaciones Cerradas\" & strArchivo For Input As #intFile
End If

Do While Not EOF(intFile)       'Busca strSection pasada en el archivo
  Line Input #intFile, Aux
  bytFound = InStr(Aux, strSection)
  If bytFound > 0 Then Exit Do
Loop

If bytFound > 0 Then    'Si encontro la cadena
  bytFound = 0
  Do While Not EOF(intFile)
    Line Input #intFile, Aux
    bytFound = InStr(Aux, strItem)  'InStr Busca strItem en el archivo
    If bytFound > 0 Then
      ConfigGetEXTENDIDO = Mid(Aux, InStr(Aux, "=") + 1) 'Extrae el resultado
      Exit Do
    End If
  Loop
End If

Close #intFile
  
Exit Function
TrapError:
Err.Clear
Close #intFile
Exit Function
End Function

'Sub ConfigSet(strSection As String, strItem As String, strValue As String)
''FUNCION: ConfigSet
''FECHA DE CREACION: 04 de julio de 2012
''AUTOR: Mariano Alfonso (Genesys)
''DESCRIPCION: Dado un encabezado de grupo y dos valores, graba el valor pasado en un archivo txt de configuracion
''PARAMETROS: strSection --> Cabecera de grupo (String)
'''            strItem    --> Item a buscar (String)
''            strValor   --> Valor a grabar
''DEVOLUCION: (String)
'MsgBox "ERRORRRRRRRRRRRRRRRRRRRRRRRRR"
'On Error GoTo TrapError
'Dim intFile As Integer
'Dim Aux As String
'Dim bytFound As Byte
'Dim Resultado As String
'
'''ConfigGet = ""
'intFile = FreeFile  'Toma el próximo archivo disponible''''

'Open App.Path & "\Config.ini" For Input As #intFile'

'Do While Not EOF(intFile)       'Busca strSection pasada en el archivo
'  Line Input #intFile, Aux
'  bytFound = InStr(Aux, strSection)
'  If bytFound > 0 Then Exit Do
'Loop

'If bytFound > 0 Then    'Si encontro la cadena
'  bytFound = 0
'  Do While Not EOF(intFile)
'    Line Input #intFile, Aux
'    bytFound = InStr(Aux, strItem)  'InStr Busca strItem en el archivo
'    If bytFound > 0 Then
'      ConfigGet = Mid(Aux, InStr(Aux, "=") + 1) 'Extrae el resultado
'      Exit Do
'    End If
'  Loop
'End If

'Close #intFile
  
'Exit Sub
'TrapError:
'Err.Clear
'Exit Sub
'End Sub


'Sub EnviarLog(ByVal filename As String, ByVal strMsj As String)
''FUNCION: EnviarLog
''FECHA DE CREACION: 27 de febrero de 2007
''AUTOR: Mariano Alfonso (NetworkIt)
''DESCRIPCION: Graba en un archivo de texto el error generado, verificando
''             si el mismo se encuentra en la lista de errores predefinidos o no
''PARAMETROS:
''
''DEVOLUCION: ()
'Dim intFile As Integer
'Dim strPaste As String
'Dim LstrHora As String
'Dim LstrUsuario As String

'intFile = FreeFile
'LstrHora = Now()

'Open filename For Append Access Write As #intFile
'Print #intFile, LstrHora & " - " & strMsj
'Close intFile

'End Sub

Function Eliminar_Archivo(KF)
'FUNCION: Eliminar_Archivo
'FECHA DE CREACION: 15 de marzo de 2007
'AUTOR: Mariano Alfonso (NetworkIt)
'DESCRIPCION: Elimina el archivo pasado como parametro
'PARAMETROS: KF --> Archivo a borrar (string)
'DEVOLUCION
On Error GoTo ErrorEliminarArchivo
    Kill KF
Exit Function
ErrorEliminarArchivo:
    MsgBox "no pudo eliminarse el archivo", vbCritical
End Function

Public Function BorrarArchivo(ByVal FilePath As String) As Boolean
'FUNCION ALTERNATIVA DE ELIMINACION DE ARCHIVO (VERIFICAR VELOCIDAD)
    On Error GoTo error
    Kill FilePath$
    BorrarArchivo = True
    Exit Function
error:
    BorrarArchivo = False
    MsgBox Err.Description, vbExclamation, "Error al eliminar la facturacion Abierta", vbCritical
    Resume
End Function

'MOVEMOS UN ARCHIVO DE FACTURACION DEL APP.PATH A LA CARPETA FACTURACIONES CERRADAS
Public Sub MoveFile(StartPath As String, EndPath As String)
    On Error GoTo error
    FileCopy StartPath$, EndPath$
    Kill StartPath$
Exit Sub
error:      MsgBox Err.Description, vbExclamation, "Error moviendo la facturacion cerrada.", vbCritical
End Sub

Public Function Copiar_Archivo(ByVal ArchivoOrigen As String, ByVal ArchivoDestino As String)
'FUNCION: Copiar_Archivo
'FECHA DE CREACION: 15 de marzo de 2007
'AUTOR: Mariano Alfonso (NetworkIt)
'DESCRIPCION: Copia un archivo destino en un nombre y ubicacion dadas
'PARAMETROS: ArchivoOrigen  --> Archivo a copiar (string)
'            ArchivoDestino --> Archivo destino (string)
'DEVOLUCION
Dim a%, Buffer%, Temp$, fRead&, fSize&, b%
On Error GoTo ErrHan:
    a = FreeFile
    Buffer = 4048
    Open ArchivoOrigen For Binary Access Read As a
    b = FreeFile
    Open ArchivoDestino For Binary Access Write As b
    fSize = FileLen(ArchivoOrigen)
    While fRead < fSize
    DoEvents
    If Buffer > (fSize - fRead) Then Buffer = (fSize - fRead)
    Temp = Space(Buffer)
    Get a, , Temp
    Put b, , Temp
    fRead = fRead + Buffer
    Wend
    Close b
    Close a
    Copiar_Archivo = 1
Exit Function
ErrHan:
    Copiar_Archivo = 0
    MsgBox Err.Description
    Resume
End Function



'CONVIERTE EL VALOR DECIMAL DE COMA A PUNTO
Function ConvertirDecimal(ByVal Valor As String) As String
Dim Lint_Pos1, LintPos2 As Integer
Dim Lstr_Aux1, Lstr_Aux2 As String
    On Error GoTo ErrorDecimal
    If Valor <> "" Then
        Lint_Pos1 = InStr(1, Valor, ",")
        If Lint_Pos1 > 0 Then
            Lstr_Aux1 = Mid(Valor, 1, Lint_Pos1 - 1) & "." & Mid(Valor, Lint_Pos1 + 1, (Len(Valor) - Lint_Pos1))
          Else
            Lstr_Aux1 = Valor
        End If
        ConvertirDecimal = Lstr_Aux1
      Else
        ConvertirDecimal = "0.00"
    End If
Exit Function
ErrorDecimal:
    ConvertirDecimal = "0.00"
End Function


'CONTROLA SI ES EL ULTIMO ARCHIVO O SE PRODUJO UN ERROR
Function Control_Ultimo_Archivo(Ultimo As String, Importado As String) As Boolean
On Error GoTo ErrControl
    Dim Lint_File1, Lint_File2 As Integer
    Dim Lint_Posicion As Integer
    Lint_File1 = Val(Mid(Ultimo, 2, 7))
    Lint_File2 = Val(Mid(Importado, 2, 7))
    If Lint_File2 > Lint_File1 Then
        Control_Ultimo_Archivo = True
      Else
        Control_Ultimo_Archivo = False
    End If
Exit Function
ErrControl:
    Control_Ultimo_Archivo = False
End Function


'VERIFICA SI EL ARCHIVO WEB EXISTE O NO
Function ArchivoExiste(LstrArchivo As String) As Boolean
Dim f%
On Error GoTo ErrorX
    LstrArchivo = Trim(LstrArchivo)
   ' Trap any errors that may occur
   On Error Resume Next
   ' Get a free file handle to avoid using a file handle already in use
   f% = FreeFile
   ' Open the file for reading
   Open LstrArchivo For Input As #f%
   ' Close it
   Close #f%
   ' If there was an error, Err will be <> 0. In that case, we return False
   ArchivoExiste = Not (Err <> 0)
Exit Function
ErrorX:
    ArchivoExiste = False
End Function


'valida un numero de cuit pasado como parametro
Public Function Validar_Cuit(mk_p_nro As String) As Boolean

    Dim mk_suma As Integer
    Dim mk_valido As String
    'mk_p_nro = Replace("-", "")
    mk_p_nro = Replace$(mk_p_nro, "-", "")

    If IsNumeric(mk_p_nro) Then
        If Len(mk_p_nro) <> 11 Then
            mk_valido = False
          Else
            mk_suma = 0
            mk_suma = mk_suma + CInt(Mid$(mk_p_nro, 1, 1)) * 5
            mk_suma = mk_suma + CInt(Mid(mk_p_nro, 2, 1)) * 4
            mk_suma = mk_suma + CInt(Mid(mk_p_nro, 3, 1)) * 3
            mk_suma = mk_suma + CInt(Mid(mk_p_nro, 4, 1)) * 2
            mk_suma = mk_suma + CInt(Mid(mk_p_nro, 5, 1)) * 7
            mk_suma = mk_suma + CInt(Mid(mk_p_nro, 6, 1)) * 6
            mk_suma = mk_suma + CInt(Mid(mk_p_nro, 7, 1)) * 5
            mk_suma = mk_suma + CInt(Mid(mk_p_nro, 8, 1)) * 4
            mk_suma = mk_suma + CInt(Mid(mk_p_nro, 9, 1)) * 3
            mk_suma = mk_suma + CInt(Mid(mk_p_nro, 10, 1)) * 2
            mk_suma = mk_suma + CInt(Mid(mk_p_nro, 11, 1)) * 1
            
            If Math.Round(mk_suma / 11, 0) = (mk_suma / 11) Then
                mk_valido = True
              Else
                mk_valido = False
            End If
        End If
      Else
        mk_valido = False
    End If
    Validar_Cuit = mk_valido
End Function

'actualiza el archivo de configuracion
Public Sub Actualizar_Config_INI(LineaModificacion As Integer, Valor As String)
On Error GoTo ErrorX:
    Dim Lectura As String
    Linea = 0
    Open App.Path & "\Config.ini" For Input As #1
'    Open App.Path & "\Config.txt" For Input As #1
    Open App.Path & "\Config.tmp" For Output As #2
    Do While Not EOF(1)
    Linea = Linea + 1
    Line Input #1, Lectura
    If Linea = LineaModificacion Then
        Print #2, Valor
      Else
        Print #2, Lectura
    End If
    Loop
    Close #1
    Close #2
    Kill App.Path & "\Config.ini"
'    Kill App.Path & "\Config.txt"
    Name App.Path & "\Config.tmp" As App.Path & "\Config.ini"
'    Name App.Path & "\Config.tmp" As App.Path & "\Config.txt"
    Exit Sub
ErrorX:
    Close #1
    Close #2
End Sub

'valida cualquier cadena si es nula o contiene datos (no importa el formato
Public Function Validar_TextBoxNulo(cadena As String) As Boolean
    If Len(cadena) = 0 Then
        Validar_TextBoxNulo = False
      Else
'        Validar_Beneficio cadena
        Validar_TextBoxNulo = True
    End If
End Function

'VALIDA EL NUMERO DE BENEFICIO PERO SOLO COMO ADVERTENCIA
Public Function Validar_Beneficio(cadena As String) As Boolean
    Dim vMensaje As String
    If Len(cadena) <> 14 Then
        vMensaje = "Por favor chequee el numero de beneficio. El mismo debería contar con 14 caracteres." & Chr(13) & _
                   "El parentesco se indica con dos digitos al final" & Chr(13) & Chr(13) & _
                   "Ejemplos de formato" & Chr(13) & _
                   "                   : 40500008070100 (el parentesco en este caso es 00)" & Chr(13) & _
                   "                   : 40506220860716 (el parentesco en este caso es 16)" & Chr(13) & _
                   "                   : 43660124240601 (el parentesco en este caso es 01)" & Chr(13) & Chr(13) & _
                   "El sistema no valida el numero de beneficio online. Debe revisar cuidadosamente sus datos ingresados." & Chr(13) & _
                   "Se le permitirá continuar la carga, pero es importante que REVISE LA NUMERACION PARA EVITAR POSIBLES DEBITOS"
        MsgBox vMensaje, vbInformation
        Validar_Beneficio = False
      Else
        Validar_Beneficio = True
    End If
End Function

'ANALIZAMOS EL STRING DEL IMPORTE Y LO FORMATEAMOS
Public Function Validar_Importe(Importe As String) As String
    Dim iImporteNumerico As Double
    iImporteNumerico = Val(Importe)
    iImporteNumerico = Format(Val(iImporteNumerico), "currency")
    Validar_Importe = Str(iImporteNumerico)
End Function

'actualiza el archivo de configuracion
Public Function Generar_Exportacion(ByVal Archivo As String, Valor As String, Flag As Boolean) As Boolean
On Error GoTo ErrorXX
    Dim Lectura As String
    Linea = 0
    Archivo = "\" & Archivo & ".txt"
    If Flag Then
        Open App.Path & Archivo For Output As #1
      Else
        Open App.Path & Archivo For Append As #1
    End If
    Print #1, Valor
    Close #1
    Generar_Exportacion = True
    Exit Function
ErrorXX:
    Close #1
End Function

'COMPLETA UNA CADENA DE TEXTO PASADO CON EL CARACTER/ES PASADO COMO PARAMETRO ALINEADO A LA IZQUIERDA/DERECHA
Function Completar_Cadena(ByVal Texto As String, Longitud As Integer, Caracter As String, Alineacion As String) As String
Dim Lint_Temp1 As Integer
Dim Lstr_Cadena As String
    Lstr_Cadena = Texto
    For Lint_Temp1 = 1 To (Longitud - Len(Trim(Texto)))
        If Alineacion = "Izquierda" Then
            Lstr_Cadena = Caracter & Lstr_Cadena
          ElseIf Alineacion = "Derecha" Then
            Lstr_Cadena = Lstr_Cadena & Caracter
        End If
    Next
    Completar_Cadena = Lstr_Cadena
End Function


'VALIDA EL PERIODO A FACTURAR
Public Function Validar_Periodo(Periodo As String) As Boolean
    Validar_Periodo = True
    Periodo = Trim(Periodo)
    If Len(Periodo) <> 7 Then
        Validar_Periodo = False
        Exit Function
    End If
    If Left(Periodo, 2) < 1 Or Left(Periodo, 2) > 12 Then
        Validar_Periodo = False
        Exit Function
    End If
    If Mid(Periodo, 3, 1) <> "-" Then
        Validar_Periodo = False
        Exit Function
    End If
    If Right(Periodo, 4) < 2012 Or Right(Periodo, 4) > 2080 Then
        Validar_Periodo = False
        Exit Function
    End If
End Function

'GENERA EL NOMBRE DEL ARCHIVO QUE SE GRABARA CON LA EXPORTACION
Public Function Generar_Archivo_Exportacion(Prestador As String, Periodo As String, Cuit As String, Factura As String, Optional Estado As String) As String
Dim sNombreArchivo As String
Dim iPosicionSeparador As Integer
    Prestador = UCase(Left(Prestador, 1))
    Periodo = Trim(Left(Periodo, 2)) & Trim(Right(Periodo, 4))
    Do While InStr(1, Cuit, "-") > 0
        iPosicionSeparador = InStr(1, Cuit, "-")
        Cuit = Trim(Mid(Cuit, 1, iPosicionSeparador - 1)) & Trim(Mid(Cuit, iPosicionSeparador + 1, Len(Cuit) - iPosicionSeparador))
    Loop
'    Do While InStr(1, Factura, "-") > 0
'        iPosicionSeparador = InStr(1, Factura, "-")
'        Factura = Trim(Mid(Factura, 1, iPosicionSeparador - 1)) & Trim(Mid(Factura, iPosicionSeparador + 1, Len(Factura) - iPosicionSeparador))
'    Loop
'    Factura = Trim(Right(Factura, 3))
    Factura = Trim(Factura)
    Generar_Archivo_Exportacion = Trim(Estado) & Trim(Prestador) & Trim(Periodo) & "-" & Trim(Cuit) & "-" & Trim(Factura)
End Function

'CHEQUEA SI EXISTE UNA FACTURACION ABIERTA O NO
Public Function Estado_Facturacion() As Boolean
Dim sEstadoFacturacion As String
    sEstadoFacturacion = ConfigGet("DATOS_FACTURACION", "File")
    If Trim(sEstadoFacturacion = "null") Then
        Estado_Facturacion = False
      Else
        Estado_Facturacion = True
    End If
End Function

'VALIDA QUE EL NUMERO DE FACTURA NO SEA NULO Y TENGA AL MENOS 3 CARACTERES NUMERICOS
Public Function Validar_Factura(Numero As String) As Boolean
Dim PosGuion As Integer
Dim ParteA As String
Dim ParteB As String
Dim auxCadena1, auxCadena2, auxCadena3 As String
    'CHEQUEAMOS QUE TENGA UN SOLO GUION
    PosGuion = InStr(1, Numero, "-")
    If PosGuion = 0 Then
        Validar_Factura = False
        Exit Function
    End If
    auxCadena1 = Left(Numero, PosGuion - 1)
    auxCadena2 = Right(Numero, Len(Numero) - PosGuion)
    auxCadena3 = Trim(auxCadena1) & Trim(auxCadena2)
    If InStr(1, auxCadena3, "-") > 0 Then
        Validar_Factura = False
        Exit Function
      Else
        If Not IsNumeric(auxCadena3) Then
            Validar_Factura = False
            Exit Function
        End If
    End If
    If PosGuion = 0 Then
        Validar_Factura = False
        Exit Function
      Else
        ParteA = Left(Numero, PosGuion - 1)
        If (Len(ParteA) < 1) Or (Len(ParteA) > 4) Then
            Validar_Factura = False
            Exit Function
          Else
            ParteA = Completar_Cadena(ParteA, 4, "0", "Izquierda")
        End If
        ParteB = Right(Numero, Len(Numero) - PosGuion)
        If (Len(ParteB) < 1) Or (Len(ParteB) > 8) Then
            Validar_Factura = False
            Exit Function
          Else
            ParteB = Completar_Cadena(ParteB, 8, "0", "Izquierda")
        End If
    End If
    Gstr_NumeroFactura = ParteA & "-" & ParteB
    Validar_Factura = True
End Function

'FUNCION PARA IMPRIMIR UN LISTVIEW
'A esta función se le envía el control LV a imprimir
Public Sub Imprimir_ListView(ListView As ListView, Archivo As String, Prestador As String, Cuit As String, PeriodoFacturado As String, NumeroFactura As String, TotalGeneral As String, Total1 As String, Optional Total2 As String, Optional Total3 As String, Optional Total4 As String)
  
Dim i As Integer, AnchoCol As Single, Espacio As Integer, x As Integer
    
    AnchoCol = 0
    'Recorremos desde la primer columna hasta la última para almacenar el ancho total
    For i = 1 To ListView.ColumnHeaders.Count
       AnchoCol = AnchoCol + ListView.ColumnHeaders(i).Width
    Next
      
    Espacio = 0
      
    'Encabezado de ejemplo
    Printer.FontBold = True
    Printer.FontSize = 12
    Printer.Print "PRESTADOR: " & Trim(Prestador)
    Printer.Print "CUIT: " & Trim(Cuit)
    Printer.Print "PERIODO FACTURADO: " & Trim(PeriodoFacturado)
    Printer.Print "NUMERO DE FACTURA: " & Trim(NumeroFactura)
    Printer.Print "FECHA DE IMPRESION: " & Now()
    Printer.Print "NOMBRE DEL ARCHIVO: " & Trim(Archivo)
      
    Printer.Print
      
    Printer.FontBold = False
    Printer.FontSize = 8
    'Imprime una línea
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
      
    With ListView
      
    'Acá se imprimen los encabezados del ListView
    For i = 1 To .ColumnHeaders.Count
        Espacio = Espacio + CInt(.ColumnHeaders(i).Width * Printer.ScaleWidth / AnchoCol)
        Printer.Print ListView.ColumnHeaders(i).Text;
        Printer.CurrentX = Espacio
    Next
    
    Printer.Print
      
    'Imprime una línea
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
      
    'Imprime Línea en blanco
    Printer.Print
      
    'Este bucle recorre los items y subitems del ListView  y los imprime
    For i = 1 To .ListItems.Count
         Espacio = 0
           
         Set Item = .ListItems(i)
         Printer.Print Item.Text;
         'Recorremos las columnas
         For x = 1 To .ColumnHeaders.Count - 1
            Espacio = Espacio + CInt(.ColumnHeaders(x).Width * Printer.ScaleWidth / AnchoCol)
            Printer.CurrentX = Espacio
            If (x = 4) Then
                Printer.Print "$ " & Item.SubItems(x);
              Else
                Printer.Print Item.SubItems(x);
            End If
         Next
           
         'Otro espacio en blanco
         Printer.Print
    Next
      
    End With
      
    Printer.Print
    'Imprime la línea de final de impresión
    Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.ScaleWidth, Printer.CurrentY)
    
    Printer.Print
    Printer.Print
    
    'IMPRESION DE TOTALES
    Printer.FontBold = True
    Printer.FontSize = 10
    Select Case Gstr_TipoPrestador
        Case "Discapacidad"
            Printer.Print "Total Prestacion: $ " & Trim(Total1)
            Printer.Print "Total Transporte: $ " & Trim(Total2)
            Printer.Print "Total Apoyo     : $ " & Trim(Total3)
            Printer.Print "Total Matricula : $ " & Trim(Total4)
            Printer.Print
            Printer.Print "TOTAL GENERAL   : $ " & Trim(TotalGeneral)
        Case "Hemodialisis"
            Printer.Print "Total Prestacion: $ " & Trim(Total1)
            Printer.Print "Total Transporte: $ " & Trim(Total2)
            Printer.Print
            Printer.Print "TOTAL GENERAL   : $ " & Trim(TotalGeneral)
        Case "Geriatria"
            Printer.Print "Total Prestacion: $ " & Trim(Total1)
            Printer.Print
            Printer.Print "TOTAL GENERAL   : $ " & Trim(TotalGeneral)
    End Select
    
    'Comenzamos la impresión
    Printer.EndDoc
End Sub



