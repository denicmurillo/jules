Attribute VB_Name = "modChecklist"
Option Explicit

Public Sub ProcesarCambioCasilla(ByRef chkControl As ContentControl)
    Dim tagVal As String
    Dim secID As Integer
    Dim isChecked As Boolean
    Dim secName As String
    Dim respuesta As VbMsgBoxResult

    tagVal = chkControl.Tag
    secID = CInt(Mid(tagVal, 5))
    isChecked = chkControl.Checked
    secName = chkControl.Title

    If isChecked Then
        ' Agregar sección en su orden correspondiente
        InsertarSeccion secID
    Else
        ' Advertencia previa antes de eliminar
        respuesta = MsgBox("ADVERTENCIA: ¿Está seguro de que desea eliminar la sección '" & secName & "'?" & vbCrLf & _
                           "Esta acción eliminará la tabla y observaciones asociadas y no se puede deshacer.", _
                           vbWarning + vbYesNo + vbDefaultButton2, "Confirmar eliminación")
        If respuesta = vbYes Then
            EliminarSeccion secID
        Else
            ' Re-marcar la casilla si el usuario cancela
            chkControl.Checked = True
        End If
    End If
End Sub

Public Sub InsertarSeccion(ByVal secID As Integer)
    Dim bmName As String
    Dim rng As Range

    bmName = "SEC_" & secID

    ' Si ya existe el marcador de la sección, no volvemos a insertarla
    If ThisDocument.Bookmarks.Exists(bmName) Then Exit Sub

    ' Buscar la posición de inserción adecuada (orden numérico de secID)
    Set rng = ObtenerRangoInsercion(secID)

    If rng Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    ' Generar el contenido de la sección dinámicamente en el rango
    CrearContenidoSeccion secID, rng

    Application.ScreenUpdating = True
End Sub

Public Sub EliminarSeccion(ByVal secID As Integer)
    Dim bmName As String
    Dim bmTarget As Bookmark
    Dim rng As Range

    bmName = "SEC_" & secID

    If ThisDocument.Bookmarks.Exists(bmName) Then
        Set bmTarget = ThisDocument.Bookmarks(bmName)
        Set rng = bmTarget.Range
        rng.Delete
        ' Eliminar explícitamente el objeto Marcador de la colección para permitir re-inserción futura
        If ThisDocument.Bookmarks.Exists(bmName) Then
            ThisDocument.Bookmarks(bmName).Delete
        End If
    End If
End Sub

Private Function ObtenerRangoInsercion(ByVal secID As Integer) As Range
    Dim i As Integer
    Dim bmName As String

    ' Buscar si hay una sección posterior (secID + 1 ... 23) ya presente
    For i = secID + 1 To 23
        bmName = "SEC_" & i
        If ThisDocument.Bookmarks.Exists(bmName) Then
            Set ObtenerRangoInsercion = ThisDocument.Bookmarks(bmName).Range.Duplicate
            ObtenerRangoInsercion.Collapse wdCollapseStart
            Exit Function
        End If
    Next i

    ' Si no hay secciones posteriores, insertar en el marcador final 'MARCADOR_INSERCION'
    If ThisDocument.Bookmarks.Exists("MARCADOR_INSERCION") Then
        Set ObtenerRangoInsercion = ThisDocument.Bookmarks("MARCADOR_INSERCION").Range.Duplicate
        ObtenerRangoInsercion.Collapse wdCollapseStart
        Exit Function
    End If

    Set ObtenerRangoInsercion = ThisDocument.Content
    ObtenerRangoInsercion.Collapse wdCollapseEnd
End Function

Private Sub CrearContenidoSeccion(ByVal secID As Integer, ByRef targetRng As Range)
    Dim titulo As String
    Dim numItems As Integer
    Dim tbl As Table
    Dim startPos As Long, endPos As Long
    Dim bmName As String

    bmName = "SEC_" & secID

    ' Obtener información de la sección según secID
    titulo = ObtenerTituloSeccion(secID)

    ' Colapsar rango al inicio de inserción
    targetRng.Collapse wdCollapseStart
    startPos = targetRng.Start

    ' Insertar Título de la sección
    targetRng.Text = titulo & vbCrLf
    targetRng.Style = ThisDocument.Styles(wdStyleHeading2)
    targetRng.Collapse wdCollapseEnd

    ' Crear Tabla para los Items
    numItems = ObtenerCantidadItems(secID)
    Set tbl = ThisDocument.Tables.Add(Range:=targetRng, NumRows:=numItems + 1, NumColumns:=4)

    ' Formatear la tabla
    With tbl
        .Borders.Enable = True
        .Range.Font.Name = "Calibri"
        .Range.Font.Size = 10

        ' Encabezados
        .Cell(1, 1).Range.Text = "Item"
        .Cell(1, 2).Range.Text = "Descripción"
        .Cell(1, 3).Range.Text = "Estado"
        .Cell(1, 4).Range.Text = "Prioridad"

        .Rows(1).Range.Font.Bold = True
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15

        ' Llenar datos
        PoblarDatosTabla secID, tbl

        ' Ajustar anchos de columnas
        .Columns(1).Width = CentimetersToPoints(1.5)
        .Columns(2).Width = CentimetersToPoints(9.5)
        .Columns(3).Width = CentimetersToPoints(2.5)
        .Columns(4).Width = CentimetersToPoints(2.5)
    End With

    ' Mover el rango objetivo al final de la tabla
    Set targetRng = tbl.Range
    targetRng.Collapse wdCollapseEnd
    targetRng.InsertParagraphAfter
    targetRng.Collapse wdCollapseEnd

    ' Insertar Campo de Observaciones
    targetRng.Text = "Observaciones: " & vbCrLf & vbCrLf
    targetRng.Font.Bold = True
    targetRng.Collapse wdCollapseEnd

    endPos = targetRng.End

    ' Crear/Actualizar Marcador de la Sección que abarca todo el contenido insertado
    ThisDocument.Bookmarks.Add Name:=bmName, Range:=ThisDocument.Range(startPos, endPos)
End Sub

Private Function ObtenerTituloSeccion(ByVal secID As Integer) As String
    Select Case secID
        Case 1: ObtenerTituloSeccion = "Documentación"
        Case 2: ObtenerTituloSeccion = "Transformador principal"
        Case 3: ObtenerTituloSeccion = "Sistema de puesta a tierra"
        Case 4: ObtenerTituloSeccion = "Sistema contra descargas atmosféricas"
        Case 5: ObtenerTituloSeccion = "Espacios físicos cuartos eléctricos"
        Case 6: ObtenerTituloSeccion = "Transformadores de baja tensión"
        Case 7: ObtenerTituloSeccion = "Tableros y protecciones"
        Case 8: ObtenerTituloSeccion = "Conductores"
        Case 9: ObtenerTituloSeccion = "Canalizaciones, bandejas, cajas de registro y conexiones"
        Case 10: ObtenerTituloSeccion = "Iluminación"
        Case 11: ObtenerTituloSeccion = "Tomacorrientes y apagadores"
        Case 12: ObtenerTituloSeccion = "Equipos de aire acondicionado, ventilación y refrigeración"
        Case 13: ObtenerTituloSeccion = "Motores"
        Case 14: ObtenerTituloSeccion = "UPS"
        Case 15: ObtenerTituloSeccion = "Sistemas de emergencia"
        Case 16: ObtenerTituloSeccion = "16.1. Lugares Clase I"
        Case 17: ObtenerTituloSeccion = "16.2. Lugares Clase II"
        Case 18: ObtenerTituloSeccion = "16.3. Lugares Clase III"
        Case 19: ObtenerTituloSeccion = "Sitios de Reunión"
        Case 20: ObtenerTituloSeccion = "Teatros y lugares similares"
        Case 21: ObtenerTituloSeccion = "Carnavales, circos, ferias y lugares similares"
        Case 22: ObtenerTituloSeccion = "Piscinas"
        Case 23: ObtenerTituloSeccion = "Instalaciones de asistencia médica"
    End Select
End Function

Private Function ObtenerCantidadItems(ByVal secID As Integer) As Integer
    Select Case secID
        Case 1: ObtenerCantidadItems = 2
        Case 2: ObtenerCantidadItems = 9
        Case 3: ObtenerCantidadItems = 14
        Case 4: ObtenerCantidadItems = 7
        Case 5: ObtenerCantidadItems = 3
        Case 6: ObtenerCantidadItems = 8
        Case 7: ObtenerCantidadItems = 14
        Case 8: ObtenerCantidadItems = 7
        Case 9: ObtenerCantidadItems = 16
        Case 10: ObtenerCantidadItems = 5
        Case 11: ObtenerCantidadItems = 7
        Case 12: ObtenerCantidadItems = 12
        Case 13: ObtenerCantidadItems = 13
        Case 14: ObtenerCantidadItems = 4
        Case 15: ObtenerCantidadItems = 10
        Case 16: ObtenerCantidadItems = 7
        Case 17: ObtenerCantidadItems = 7
        Case 18: ObtenerCantidadItems = 7
        Case 19: ObtenerCantidadItems = 3
        Case 20: ObtenerCantidadItems = 8
        Case 21: ObtenerCantidadItems = 7
        Case 22: ObtenerCantidadItems = 7
        Case 23: ObtenerCantidadItems = 27
    End Select
End Function

Private Sub PoblarDatosTabla(ByVal secID As Integer, ByRef tbl As Table)
    Select Case secID
        Case 1
            SetFila tbl, 2, "1", "¿Se cuenta con planos actualizados de la instalación eléctrica y distribución arquitectónica?", "No aplica", "Leve"
            SetFila tbl, 3, "2", "Revisar que se tenga documentación adicional como memoria de cálculo, especificaciones, manuales, capacidades de interrupción, bitácoras de mantenimiento, etc.", "No aplica", "Leve"
        Case 2
            SetFila tbl, 2, "1", "¿La placa de datos se encuentra en su lugar y tiene los datos completos y visibles?", "No aplica", "Leve"
            SetFila tbl, 3, "2", "¿El estado físico del transformador no evidencia fugas, corrosión o daño físico?", "No aplica", "Alta"
            SetFila tbl, 4, "3", "¿Las condiciones de instalación del transformador cumplen los requisitos?", "No aplica", "Moderada"
            SetFila tbl, 5, "4", "¿Existe algún registro de mantenimiento del transformador? Como pruebas de aceite, aislamiento o relación de transformación, entre otros.", "No aplica", "Leve"
            SetFila tbl, 6, "5", "¿El transformador cuenta con protección de sobrecorriente para el primario de 600V o más, y está dimensionado correctamente?", "No aplica", "Alta"
            SetFila tbl, 7, "6", "¿El transformador cuenta con protección de sobrecorriente para el secundario de 600V o menos, y está dimensionado correctamente?", "No aplica", "Alta"
            SetFila tbl, 8, "7", "¿Existe protección contra sobrecorriente para los conductores del primario del transformador?", "No aplica", "Alta"
            SetFila tbl, 9, "8", "¿Existe protección contra sobrecorriente para los conductores del secundario del transformador?", "No aplica", "Alta"
            SetFila tbl, 10, "9", "¿La bóveda del transformador está correctamente construida, posee acceso, ventilación y drenaje? ¿Presenta grietas o filtraciones?", "No aplica", "Alta"
        Case 3
            SetFila tbl, 2, "1", "¿Existe sistema de puesta a tierra en el sitio?¿Cumple con los requisitos mínimos?", "No aplica", "Alta"
            SetFila tbl, 3, "2", "¿El(los) conductor(es) del electrodo de puesta a tierra está(n) dimensionado(s) apropiadamente?", "No aplica", "Alta"
            SetFila tbl, 4, "3", "¿Se cuenta con acceso a los electrodos que conforman la malla de puesta a tierra?", "No aplica", "Leve"
            SetFila tbl, 5, "4", "¿Se cuenta con correcta soldadura o medio de sujeción entre los electrodos y el cable?", "No aplica", "Moderada"
            SetFila tbl, 6, "5", "¿Hay evidencia de daño físico en el estado del sistema de puestas a tierra?", "No aplica", "Alta"
            SetFila tbl, 7, "6", "¿Los registros cumplen con las dimensiones y requerimientos constructivos y se encuentran en un buen estado de conservación?", "No aplica", "Moderada"
            SetFila tbl, 8, "7", "¿La medición de la puesta a tierra está dentro del rango aceptable para este tipo de instalación?", "No aplica", "Moderada"
            SetFila tbl, 9, "8", "¿Hay corrientes en el sistema de puestas a tierra de la instalación?", "No aplica", "Alta"
            SetFila tbl, 10, "9", "¿Hay continuidad entre los conductores de tierras y conexiones equipotenciales?", "No aplica", "Alta"
            SetFila tbl, 11, "10", "¿Existe aterrizamiento equipotencial de estructura metálica expuesta?", "No aplica", "Alta"
            SetFila tbl, 12, "11", "¿El tamaño del puente de conexión equipotencial principal de acometida cumple conforme a NEC 2020 artículo 250.28 y 250.102(C)?", "No aplica", "Alta"
            SetFila tbl, 13, "12", "¿Los equipos tienen puesta a tierra adecuada, tanto para equipos fijos en su sitio ó conectados mediante métodos de alambrado permanente?", "No aplica", "Alta"
            SetFila tbl, 14, "13", "¿Las cajas de salida de tomas y apagadores se encuentran puestas a tierra?", "No aplica", "Alta"
            SetFila tbl, 15, "14", "¿El equipo especial cuenta con puesta a tierra? (Lavandería, Cocina, Aire Acondicionado, Equipo de laboratorio, equipo de producción industrial, equipo médico, ascensores, calderas, autoclaves, entre otros)", "No aplica", "Alta"
        Case 4
            SetFila tbl, 2, "1", "¿Existe un sistema de protección contra descargas atmosféricas? ¿Revisar si la instalación lo requiere?", "No aplica", "Moderada"
            SetFila tbl, 3, "2", "¿Qué tipo de pararrayos está instalado?", "No aplica", "Leve"
            SetFila tbl, 4, "3", "¿Las condiciones y el estado de la instalación cumplen los requisitos de la normativa aplicable?", "No aplica", "Moderada"
            SetFila tbl, 5, "4", "¿Se cuenta con correcta protección mecánica de los bajantes?", "No aplica", "Moderada"
            SetFila tbl, 6, "5", "¿Cumple la malla de tierras con los valores requeridos para el correcto funcionamiento del pararrayos instalado?", "No aplica", "Moderada"
            SetFila tbl, 7, "6", "¿Existe un registro de mantenimiento del pararrayos?", "No aplica", "Leve"
            SetFila tbl, 8, "7", "¿Existe un registro de conteo de eventos en el pararrayos (Contador de eventos)?", "No aplica", "Leve"
        Case 5
            SetFila tbl, 2, "1", "¿Se respetan las distancias de seguridad mínimas en el equipamiento de los cuartos eléctricos?", "No aplica", "Alta"
            SetFila tbl, 3, "2", "¿Se respetan los accesos, iluminación y ventilación requeridos para los cuartos eléctricos?", "No aplica", "Alta"
            SetFila tbl, 4, "3", "¿Las condiciones generales como pintura, techos, material de construcción, almacenamiento de otros materiales inflamables, entre otros, son las adecuadas?", "No aplica", "Alta"
        Case 6
            SetFila tbl, 2, "1", "¿La placa de datos se encuentra en su lugar y tiene los datos completos y visibles?", "No aplica", "Leve"
            SetFila tbl, 3, "2", "¿Evidencia fugas, corrosión o daño físico?", "No aplica", "Alta"
            SetFila tbl, 4, "3", "¿Es el Montaje y ventilación segura?", "No aplica", "Moderada"
            SetFila tbl, 5, "4", "¿El transformador cuenta con las Protecciones y cables alimentadores adecuados?", "No aplica", "Alta"
            SetFila tbl, 6, "5", "¿El transformador cuenta con protección de sobrecorriente para el secundario de 600V o menos, y está dimensionado correctamente? (NEC 2020 Tabla 450.3(A))", "No aplica", "Alta"
            SetFila tbl, 7, "6", "¿Existe protección contra sobrecorriente para los conductores del primario del transformador?", "No aplica", "Alta"
            SetFila tbl, 8, "7", "¿Existe protección contra sobrecorriente para los conductores del secundario del transformador?", "No aplica", "Alta"
            SetFila tbl, 9, "8", "Sistema derivado separado: ¿El sistema tiene los electrodos de puesta a tierra, conductores de electrodos de puesta a tierra y puentes de conexión equipotencial adecuados? Acorde a NEC 2020 artículo 250.30.", "No aplica", "Alta"
        Case 7
            SetFila tbl, 2, "1", "¿Los tableros se encuentran debidamente rotulados con la información necesaria y suficiente para su operación segura? ¿Son accesibles?", "No aplica", "Moderada"
            SetFila tbl, 3, "2", "¿El estado físico de los tableros en su exterior evidencia daños que pongan en peligro la seguridad de usuarios y personal a su alrededor?", "No aplica", "Alta"
            SetFila tbl, 4, "3", "¿La altura de los tableros es correcta? ¿Es el montaje y ventilación segura? NEC 2020 Artículo 110.26 (E)", "No aplica", "Moderada"
            SetFila tbl, 5, "4", "¿El espacio de frente y ancho de espacio es correcto? NEC 2020 Artículo 110.26 (A)", "No aplica", "Alta"
            SetFila tbl, 6, "5", "¿El estado físico de los tableros en su interior evidencia daños ó evidencia de manipulación que pongan en peligro la seguridad de usuarios y personal a su alrededor?", "No aplica", "Alta"
            SetFila tbl, 7, "6", "¿Existen espacios libres de breakers sin tapas dentro del tablero? ¿Existen aberturas sin utilizar de forma que las partes energizadas del tablero queden expuestas a contacto involuntario?", "No aplica", "Alta"
            SetFila tbl, 8, "7", "¿Faltan tornillos en las tapas, barras, cables o breakers?", "No aplica", "Moderada"
            SetFila tbl, 9, "8", "¿Uso incorrecto de tornillos en tableros (Tornillos que no son de uso eléctrico)?", "No aplica", "Alta"
            SetFila tbl, 10, "9", "¿Están los circuitos y tuberías correctamente identificados y rotulados?", "No aplica", "Moderada"
            SetFila tbl, 11, "10", "¿Se usan interruptores con protección de falla a tierra (GFCIs) en zonas húmedas?", "No aplica", "Alta"
            SetFila tbl, 12, "11", "¿Es requerida la protección contra falla de arco (AFCI) y está siendo utilizada?", "No aplica", "Moderada"
            SetFila tbl, 13, "12", "¿Está la capacidad nominal de interruptores de acuerdo con los cables?", "No aplica", "Alta"
            SetFila tbl, 14, "13", "¿Está asegurada la coordinación entre interruptores y cables? ¿Existe la memoria de esta coordinación?", "No aplica", "Leve"
            SetFila tbl, 15, "14", "¿Existe continuidad eléctrica entre la barra de neutro y la barra de tierra?", "No aplica", "Alta"
        Case 8
            SetFila tbl, 2, "1", "¿Se encuentran los conductores en buen estado de conservación?", "No aplica", "Alta"
            SetFila tbl, 3, "2", "¿Existe rotulación para la identificación de conductores de fases, neutro y de tierras, clara y evidente?", "No aplica", "Moderada"
            SetFila tbl, 4, "3", "¿Existe documentación de mediciones de resistencia de aislamiento de bajo voltaje en los conductores de alimentadores y subalimentadores?", "No aplica", "Leve"
            SetFila tbl, 5, "4", "¿Los cables flexibles, extensiones y regletas están siendo utilizados de manera correcta?", "No aplica", "Moderada"
            SetFila tbl, 6, "5", "¿Cuentan con un correcto código de colores?", "No aplica", "Moderada"
            SetFila tbl, 7, "6", "¿Los conductores de cada circuito se encuentran agrupados?", "No aplica", "Moderada"
            SetFila tbl, 8, "7", "¿Los empalmes están realizados por métodos aprobados?", "No aplica", "Moderada"
        Case 9
            SetFila tbl, 2, "1", "¿Cumplen los requisitos de accesibilidad y conservación las canalizaciones y registros de la instalación?", "No aplica", "Moderada"
            SetFila tbl, 3, "2", "¿Están los soportes instalados y conservados de manera correcta?", "No aplica", "Moderada"
            SetFila tbl, 4, "3", "¿Están las canalizaciones, conexiones y cajas de registro protegidas de daños mecánicos en los pasos de muros y suelos?", "No aplica", "Moderada"
            SetFila tbl, 5, "4", "¿Las canalizaciones instaladas son de material acorde al sitio en que se encuentran? (expuestas, embebidas en concreto, al aire libre, en zonas clasificadas)", "No aplica", "Moderada"
            SetFila tbl, 6, "5", "¿Las cajas de registro se encuentran con cantidad de volumen de cables y empalmes adecuados? NEC 2020 Artículo 314.16(A)", "No aplica", "Moderada"
            SetFila tbl, 7, "6", "¿Cuentan con terminaciones adecuadas en las cajas de distribución?", "No aplica", "Moderada"
            SetFila tbl, 8, "7", "¿Las canalizaciones cuentan con correcta capacidad de llenado?", "No aplica", "Moderada"
            SetFila tbl, 9, "8", "¿Las canalizaciones, bandejas o portacables eléctricas se usan exclusivamente para cables eléctricos?", "No aplica", "Moderada"
            SetFila tbl, 10, "9", "¿Las canalizaciones, bandejas o portacables eléctricas se encuentran correctamente soportados?", "No aplica", "Moderada"
            SetFila tbl, 11, "10", "¿Las canalizaciones, bandejas o portacables metálicos se encuentran puestas a tierra?", "No aplica", "Moderada"
            SetFila tbl, 12, "11", "¿Los conductores dentro de las cajas de registro tienen una longitud adecuada?", "No aplica", "Moderada"
            SetFila tbl, 13, "12", "¿Las cajas de registro están instaladas en puntos de unión, empalme, salidas de tomas, luminarias, salidas especiales, interruptores?", "No aplica", "Moderada"
            SetFila tbl, 14, "13", "¿Las cajas cuentan con cubiertas y tapas en buen estado?", "No aplica", "Moderada"
            SetFila tbl, 15, "14", "¿Se utilizan conectores contra la abrasión en tuberías con conductores mayores al calibre 4 AWG?", "No aplica", "Moderada"
            SetFila tbl, 16, "15", "¿Existe el espacio adecuado para los conductores, empalmes o derivaciones en gabinetes y ductos cuadrados?", "No aplica", "Moderada"
            SetFila tbl, 17, "16", "¿Existe cableado adecuado en ductos, cámaras de distribución de aire, plenum y otros espacios de circulación de aire?", "No aplica", "Moderada"
        Case 10
            SetFila tbl, 2, "1", "¿La especificación de las luminarias son acorde al área en que se encuentra instalado? Ejemplo: oficinas, pasillos, exteriores, áreas húmedas, etc.", "No aplica", "Moderada"
            SetFila tbl, 3, "2", "¿Las luminarias instaladas son certificadas y cumplen con la normativa del Código Eléctrico Nacional 2020?", "No aplica", "Moderada"
            SetFila tbl, 4, "3", "¿Las luminarias se encuentran en buen estado?", "No aplica", "Moderada"
            SetFila tbl, 5, "4", "¿La instalación requiere de un sistema de iluminación de emergencia? En caso positivo, ¿Cuenta con este sistema?", "No aplica", "Alta"
            SetFila tbl, 6, "5", "¿La instalación requiere de un sistema de rótulos luminosos de salida? En caso positivo, ¿Cuenta con este sistema?", "No aplica", "Alta"
        Case 11
            SetFila tbl, 2, "1", "¿La especificación de los tomacorrientes y apagadores son acorde al área en que se encuentran instalados? Ejemplo: oficinas, pasillos, exteriores, etc.", "No aplica", "Moderada"
            SetFila tbl, 3, "2", "¿El montaje, las condiciones y el estado de la instalación cumplen los requisitos?¿Permite una operación segura?", "No aplica", "Moderada"
            SetFila tbl, 4, "3", "¿Los apagadores y tomacorrientes son certificados?", "No aplica", "Moderada"
            SetFila tbl, 5, "4", "¿Las capacidades nominales de los tomacorrientes son compatibles con las capacidades nominales de la protección?", "No aplica", "Alta"
            SetFila tbl, 6, "5", "Verificar que la polaridad de los tomacorrientes es la adecuada", "No aplica", "Moderada"
            SetFila tbl, 7, "6", "¿En lugares húmedos y mojados, se tiene las cubiertas y cerramientos apropiados?", "No aplica", "Moderada"
            SetFila tbl, 8, "7", "¿En exteriores, baños, cerca de pilas, se cuenta con protección contra falla a tierra (GFCI)?", "No aplica", "Alta"
        Case 12
            SetFila tbl, 2, "1", "¿Está la placa de datos legible y disponible en el equipo?", "No aplica", "Leve"
            SetFila tbl, 3, "2", "¿El estado físico del equipo es tal que no representa un riesgo para el personal que lo opera o lo rodea?", "No aplica", "Alta"
            SetFila tbl, 4, "3", "¿Es el Montaje y ventilación segura?", "No aplica", "Moderada"
            SetFila tbl, 5, "4", "¿Los calibres de los conductores del circuito ramal son adecuados y en base a la información de la placa característica?", "No aplica", "Alta"
            SetFila tbl, 6, "5", "¿Los calibres de los conductores del circuito ramal múltiples unidades es adecuado y en base a la información de la placa característica?", "No aplica", "Alta"
            SetFila tbl, 7, "6", "¿Cuenta con las protecciones de sobrecorriente adecuadas?", "No aplica", "Alta"
            SetFila tbl, 8, "7", "¿Cuenta con las protecciones contra cortocircuito y falla a tierra del ramal y alimentador dimensionados adecuadamente?", "No aplica", "Alta"
            SetFila tbl, 9, "8", "Equipos con controladores: ¿El controlador es del tipo apropiado y posee capacidades nominales adecuadas (Cuando no son parte de un equipo multimotor certificado o un equipo de carga combinada certificado)?", "No aplica", "Moderada"
            SetFila tbl, 10, "9", "¿El medio de desconexión tiene capacidades nominales adecuadas para el equipo?", "No aplica", "Alta"
            SetFila tbl, 11, "10", "¿El medio de desconexión se encuentra al alcance de la vista del equipo?", "No aplica", "Alta"
            SetFila tbl, 12, "11", "¿Cuenta con espacios de acceso adecuado para su mantenimiento?", "No aplica", "Alta"
            SetFila tbl, 13, "12", "¿Cuenta con tomacorrientes e iluminación adecuados para el mantenimiento y reparación de los equipos?", "No aplica", "Moderada"
        Case 13
            SetFila tbl, 2, "1", "¿Está la placa de datos legible y disponible en el equipo?", "No aplica", "Leve"
            SetFila tbl, 3, "2", "¿El estado físico del equipo es tal que no representa un riesgo para el personal que lo opera o lo rodea?", "No aplica", "Alta"
            SetFila tbl, 4, "3", "¿Es el Montaje y ventilación segura?", "No aplica", "Moderada"
            SetFila tbl, 5, "4", "¿Los calibres de los conductores del circuito ramal son adecuados y en base a la información de la placa característica y al 125% del FLC?", "No aplica", "Alta"
            SetFila tbl, 6, "5", "¿Los calibres de los conductores del circuito ramal múltiples motores es adecuado?", "No aplica", "Alta"
            SetFila tbl, 7, "7", "¿Cuenta con las protecciones de sobrecarga adecuadas?", "No aplica", "Alta"
            SetFila tbl, 8, "8", "¿Cuenta con las protecciones del circuito ramal contra cortocircuito y fallas a tierra?", "No aplica", "Alta"
            SetFila tbl, 9, "9", "¿Cuenta con las protecciones del alimentador contra cortocircuito y fallas a tierra?", "No aplica", "Alta"
            SetFila tbl, 10, "10", "¿Cuenta con la protección contra sobrecorriente apropiada de los circuitos de control de los motores?", "No aplica", "Moderada"
            SetFila tbl, 11, "11", "Motores con controladores: ¿El controlador es del tipo apropiado y posee capacidades nominales adecuadas?", "No aplica", "Moderada"
            SetFila tbl, 12, "12", "¿El medio de desconexión es accesible y posee espacio de trabajo adecuado?", "No aplica", "Alta"
            SetFila tbl, 13, "13", "¿El medio de desconexión del motor posee capacidad nominal apropiada?", "No aplica", "Alta"
            SetFila tbl, 14, "14", "¿El desconectador del motor se encuentra a la vista desde el motor, tiene espacio de trabajo adecuado y bloqueo por llave?", "No aplica", "Alta"
        Case 14
            SetFila tbl, 2, "1", "¿Está la placa de datos disponible y legible en el equipo?", "No aplica", "Leve"
            SetFila tbl, 3, "2", "¿El estado físico del equipo es tal que no representa un riesgo para el personal que lo opera o lo rodea?", "No aplica", "Alta"
            SetFila tbl, 4, "3", "¿Es el Montaje y ventilación segura?", "No aplica", "Moderada"
            SetFila tbl, 5, "4", "¿Cuenta con las protecciones, controles y cables acordes con sus características?", "No aplica", "Alta"
        Case 15
            SetFila tbl, 2, "1", "¿El sistema de emergencia es legalmente exigido", "No aplica", "Leve"
            SetFila tbl, 3, "2", "¿Está la placa de datos legible y disponible en el equipo?", "No aplica", "Leve"
            SetFila tbl, 4, "3", "¿El estado físico del equipo es tal que no representa un riesgo para el personal que lo opera o lo rodea?", "No aplica", "Alta"
            SetFila tbl, 5, "4", "¿Es el Montaje y ventilación segura?", "No aplica", "Moderada"
            SetFila tbl, 6, "5", "¿Cuenta con las protecciones, controles y cables acordes con sus características?", "No aplica", "Alta"
            SetFila tbl, 7, "6", "¿Se cuenta con una transferencia automática, con medios de derivación, se opere eléctricamente y esté retenido mecánicamente?", "No aplica", "Alta"
            SetFila tbl, 8, "7", "¿Se cuenta con interruptor principal para desconexión de la energía?", "No aplica", "Alta"
            SetFila tbl, 9, "8", "¿Existen registros de mantenimiento preventivo de la planta eléctrica de emergencia?", "No aplica", "Leve"
            SetFila tbl, 10, "9", "¿El cableado de los sistemas de emergencia son los adecuados?", "No aplica", "Alta"
            SetFila tbl, 11, "10", "¿La instalación requiere de sistema de protección contra incendios y este cumple de conformidad con las normas aplicables?", "No aplica", "Alta"
        Case 16
            SetFila tbl, 2, "1", "¿Los materiales utilizados en las conexiones flexibles, como herrajes flexibles y cordones flexibles a prueba de explosión son correctos para la zona?", "No aplica", "Alta"
            SetFila tbl, 3, "2", "¿Las canalizaciones y cableado cumplen con las disposiciones de sellado y drenaje?", "No aplica", "Alta"
            SetFila tbl, 4, "3", "¿Las temperaturas marcadas en los equipos no son mayores que la temperatura de ignición de los gases o vapores involucrados?", "No aplica", "Alta"
            SetFila tbl, 5, "4", "¿Las capacidades nominales de las protecciones son adecuadas para la alimentación de equipo?", "No aplica", "Alta"
            SetFila tbl, 6, "5", "¿Se cuenta con protección contra tensiones transitorias?", "No aplica", "Moderada"
            SetFila tbl, 7, "6", "¿Los equipos eléctricos transformadores, condensadores, motores, luminarias, tomacorrientes instalados en el lugar, cumplen a cabalidad con los requerimientos del NEC 2020 artículo 501?", "No aplica", "Alta"
            SetFila tbl, 8, "7", "¿Las trayectorias de puesta a tierra y conexión equipotencial al medio de desconexión de la edificación son correctas?", "No aplica", "Alta"
        Case 17
            SetFila tbl, 2, "1", "¿Los materiales utilizados en las conexiones flexibles, como herrajes flexibles y cordones flexibles a prueba de explosión son correctos para la zona?", "No aplica", "Alta"
            SetFila tbl, 3, "2", "¿Las canalizaciones y cableado cumplen con las disposiciones de sellado y drenaje que solicita el NEC 2020 en su artículo 502.15(A) hasta (F)?", "No aplica", "Alta"
            SetFila tbl, 4, "3", "¿Las temperaturas marcadas en los equipos no son mayores que la temperatura de ignición de los gases o vapores involucrados?", "No aplica", "Alta"
            SetFila tbl, 5, "4", "¿Las capacidades nominales de las protecciones son adecuadas para la alimentación de equipo y cumplen con lo solicitado en NEC 2020 Artículo 502.115?", "No aplica", "Alta"
            SetFila tbl, 6, "5", "¿Los equipos eléctricos transformadores, condensadores, motores, luminarias, tomacorrientes instalados en el lugar, cumplen a cabalidad con los requerimientos del NEC 2020 artículo 502?", "No aplica", "Alta"
            SetFila tbl, 7, "6", "¿Las trayectorias de puesta a tierra y conexión equipotencial al medio de desconexión de la edificación son correctas? NEC 2020 Artículo 502.30", "No aplica", "Alta"
            SetFila tbl, 8, "7", "¿Se cuenta con tuberías de ventilación? De ser así: Cumple con NEC 2020 Artículo 502.128", "No aplica", "Alta"
        Case 18
            SetFila tbl, 2, "1", "¿Los materiales utilizados en las conexiones flexibles, como herrajes flexibles y cordones flexibles a prueba de explosión son correctos para la zona? NEC 2020 Artículo 501.10(A) y (B) y 503.140", "No aplica", "Alta"
            SetFila tbl, 3, "2", "¿Las canalizaciones, barras conductoras, terminales, componentes y cableado cumplen con las disposiciones que solicita el NEC 2020 en su artículo 503.25?", "No aplica", "Alta"
            SetFila tbl, 4, "3", "¿Las temperaturas marcadas en los equipos no son mayores que la temperatura de ignición de los gases o vapores involucrados? NEC 2020 Artículo 500.8(C)", "No aplica", "Alta"
            SetFila tbl, 5, "4", "¿Las capacidades nominales de las protecciones son adecuadas para la alimentación de equipo y cumplen con lo solicitado en NEC 2020 Artículo 503.115?", "No aplica", "Alta"
            SetFila tbl, 6, "5", "¿Los equipos eléctricos transformadores, condensadores, motores, luminarias, tomacorrientes instalados en el lugar, cumplen a cabalidad con los requerimientos del NEC 2020 artículo 503?", "No aplica", "Alta"
            SetFila tbl, 7, "6", "¿Las trayectorias de puesta a tierra y conexión equipotencial al medio de desconexión de la edificación son correctas? NEC 2020 Artículo 503.30", "No aplica", "Alta"
            SetFila tbl, 8, "7", "¿Se cuenta con tuberías de ventilación? De ser así: Cumple con NEC 2020 Artículo 503.128", "No aplica", "Alta"
        Case 19
            SetFila tbl, 2, "1", "¿El edificio clasifica como sitio de reunión (al menos 100 personas)?", "No aplica", "Leve"
            SetFila tbl, 3, "2", "¿Los métodos de alambrado son los adecuados para el tipo de inmueble y la clasificación de resistencia al fuego de las áreas?", "No aplica", "Alta"
            SetFila tbl, 4, "3", "¿El alambrado y los equipos de distribución portátiles en instalaciones provisionales para eventos son las adecuadas?", "No aplica", "Alta"
        Case 20
            SetFila tbl, 2, "1", "¿Los métodos de alambrado son los adecuados para el tipo de inmueble?", "No aplica", "Alta"
            SetFila tbl, 3, "2", "¿Los reguladores de intensidad (Dimmers) cumplen con los requisitos de desconexión, protección contra sobrecorriente y alimentadores?", "No aplica", "Moderada"
            SetFila tbl, 4, "3", "¿Los equipos fijos del escenario cumplen con los requisitos de construcción e instalación?", "No aplica", "Moderada"
            SetFila tbl, 5, "4", "¿Los tableros y equipos portátiles cumplen con la construcción, alimentadores, protecciones y certificados?", "No aplica", "Alta"
            SetFila tbl, 6, "5", "¿Los camerinos utilizan porta lámparas colgantes o bombillas incandescentes expuestas sin resguardo?", "No aplica", "Alta"
            SetFila tbl, 7, "6", "¿Las luminarias y tomacorrientes adyacentes a los espejos y sobre mesa tocadores en camerinos y áreas de maquillaje son controlados por interruptor con luz piloto?", "No aplica", "Moderada"
            SetFila tbl, 8, "7", "¿Las canalizaciones metálicas, cables con cubiertas metálicas, bastidores y envolventes metálicos de todos los equipos están conectados a un conductor de puesta a tierra de equipos?", "No aplica", "Alta"
            SetFila tbl, 9, "8", "En cines, ¿Las salas de proyección cumplen con los requisitos de ubicación, alimentación, protección y listado?", "No aplica", "Moderada"
        Case 21
            SetFila tbl, 2, "1", "¿Se cumplen con las distancias de seguridad de los conductores aéreos?", "No aplica", "Alta"
            SetFila tbl, 3, "2", "¿Los equipos eléctricos y métodos de alambrado en estructuras portátiles están equipados con protección mecánica?", "No aplica", "Alta"
            SetFila tbl, 4, "3", "¿Los equipos de acometida o los generadores portátiles cumplen con los requisitos de montaje, ubicación, cableado y protección?", "No aplica", "Alta"
            SetFila tbl, 5, "4", "¿Los métodos de alambrado cumplen con los requisitos para este tipo de lugares? tipo, protección mecánica, soporte, cajas y accesorios", "No aplica", "Moderada"
            SetFila tbl, 6, "5", "¿Los tomacorrientes y equipos de fácil acceso al público general tienen protección GFCI?", "No aplica", "Alta"
            SetFila tbl, 7, "6", "¿Las canalizaciones, envolventes metálicos y estructuras están unidas equipotencialmente?", "No aplica", "Alta"
            SetFila tbl, 8, "7", "¿Todos los equipos eléctricos poseen conductor de puesta a tierra de equipos?", "No aplica", "Alta"
        Case 22
            SetFila tbl, 2, "1", "¿Se cumplen con las distancias de seguridad de los conductores aéreos?", "No aplica", "Alta"
            SetFila tbl, 3, "2", "¿Los métodos de alambrado en sitios con ambientes corrosivos son de los materiales adecuados?", "No aplica", "Moderada"
            SetFila tbl, 4, "3", "¿Las partes metálicas de las piscinas y otros equipos metálicos están unidos equipotencialmente y se utilizan accesorios aptos para ambientes húmedos y corrosivos?", "No aplica", "Alta"
            SetFila tbl, 5, "4", "¿Los motores cumplen con los requisitos? ¿Cuando son de 60 A o menos utilizan protección GFCI?", "No aplica", "Alta"
            SetFila tbl, 6, "5", "¿La ubicación, métodos de cableado y protecciones de los tomacorrientes y luminarias cumplen con los requisitos para piscinas?", "No aplica", "Alta"
            SetFila tbl, 7, "6", "¿Las luminarias subacuáticas cumplen con los requisitos de instalación y protección?", "No aplica", "Alta"
            SetFila tbl, 8, "7", "¿Todos los equipos eléctricos para piscinas poseen protección contra falla a tierra?", "No aplica", "Alta"
        Case 23
            SetFila tbl, 2, "1", "¿Hay selectividad entre las protecciones de falla a tierra de la acometida y el alimentador?", "No aplica", "Moderada"
            SetFila tbl, 3, "2", "¿Los dispositivos de protección contra sobrecorriente que alimentan el sistema esencial están coordinados en fallas mayores a 0.1s?", "No aplica", "Moderada"
            SetFila tbl, 4, "3", "¿Se utilizan conductores de cobre aislados en canalizaciones metálicas, para brindar la puesta a tierra de los equipos en las áreas de cuidado de los pacientes?", "No aplica", "Alta"
            SetFila tbl, 5, "4", "¿Se cuenta con conexión equipotencial entre los paneles de distribución de circuitos ramales esenciales que alimentan áreas de pacientes?", "No aplica", "Moderada"
            SetFila tbl, 6, "5", "¿Se cuenta con al menos dos circuitos ramales, en el área del cuidado del paciente del área de servicios generales, y que estos circuitos sean uno del sistema Normal y el otro del sistema crítico?", "No aplica", "Moderada"
            SetFila tbl, 7, "6", "¿Cada cama de un paciente del área de cuidados generales tiene al menos 8 receptáculos y que al menos 4 de los tomacorrientes sean alimentados del circuito normal o crítico y estén alimentados de otra transferencia que la de los demás receptáculos?", "No aplica", "Moderada"
            SetFila tbl, 8, "7", "¿Cada cama de un paciente del área de cuidados críticos tiene al menos 14 receptáculos y que al menos 7 de los tomacorrientes sean alimentados del circuito normal o crítico y estén alimentados de otra transferencia que la de los demás receptáculos?", "No aplica", "Moderada"
            SetFila tbl, 9, "8", "¿En quirófanos hay al menos 36 receptáculos y que al menos 12 de los tomacorrientes sean alimentados del circuito normal o crítico y estén alimentados de otra transferencia que la de los demás receptáculos?", "No aplica", "Moderada"
            SetFila tbl, 10, "9", "¿Se cuenta con algún punto de conexión de puesta a tierra de equipos en la vecindad de las camas de pacientes?", "No aplica", "Alta"
            SetFila tbl, 11, "10", "¿Los tomacorrientes de las áreas de las camas de los pacientes, críticos y quirófanos son de grado hospitalario?", "No aplica", "Moderada"
            SetFila tbl, 12, "11", "¿Existen tomacorrientes con protección contra falla a tierra tipo A en los lugares requeridos por características húmedas, ó en donde no se puedan tolerar interrupciones mediante sistemas de potencia separados?", "No aplica", "Alta"
            SetFila tbl, 13, "12", "¿Los sistemas eléctricos esenciales constan con los sistemas de equipos y de emergencia?", "No aplica", "Moderada"
            SetFila tbl, 14, "13", "¿El sistema de emergencia posee ramales de seguridad para la vida y ramales críticos?", "No aplica", "Moderada"
            SetFila tbl, 15, "14", "¿La carga del sistema esencial posee la capacidad adecuada de las fuentes de alimentación y de los alimentadores?", "No aplica", "Alta"
            SetFila tbl, 16, "15", "¿El alambrado del sistema de emergencia (ramales de seguridad para la vida y ramales críticos) es independiente y se encuentra separado de otros alambrados y equipos?", "No aplica", "Alta"
            SetFila tbl, 17, "16", "¿Se suministra protección mecánica para el alambrado del sistema esencial, con canalizaciones metálicas no flexibles o cable tipo MI?", "No aplica", "Moderada"
            SetFila tbl, 18, "17", "¿Las cargas previstas para conexión al ramal de seguridad de la vida se encuentran alimentadas solamente por dicho ramal?", "No aplica", "Alta"
            SetFila tbl, 19, "18", "¿Las cargas previstas para conexión al crítico se encuentran alimentadas solamente por dicho ramal?", "No alta", "Alta"
            SetFila tbl, 20, "19", "¿Las cargas previstas para conexión al ramal de equipos se encuentran alimentadas solamente por dicho ramal?", "No aplica", "Moderada"
            SetFila tbl, 21, "20", "¿Se suministran dos fuentes independientes para el sistema eléctrico esencial y la fuente alternativa es adecuada?", "No aplica", "Alta"
            SetFila tbl, 22, "21", "¿El equipo de rayos x posee un medio de desconexión de la capacidad adecuada y fácilmente accesible?", "No aplica", "Alta"
            SetFila tbl, 23, "22", "¿Los alimentadores y los dispositivos de protección contra sobrecorriente de equipos de rayos x están dimensionados correctamente?", "No aplica", "Moderada"
            SetFila tbl, 24, "23", "¿Los equipos de rayos X están puestos a tierra correctamente?", "No aplica", "Alta"
            SetFila tbl, 25, "24", "¿El monitor de aislamiento de línea (LIM) está operativo y señaliza correctamente (verde = Normal, Rojo = Falla)?", "No aplica", "Alta"
            SetFila tbl, 26, "25", "¿El LIM se ubica dentro de las salas de operaciones?", "No aplica", "Moderada"
            SetFila tbl, 27, "26", "¿El LIM dispara una alarma visual/auditiva cuando la corriente de fuga total alcanza los 5mA?", "No aplica", "Alta"
            SetFila tbl, 28, "27", "¿Los conductores aislados son identificados correctamente (naranja, café)?", "No aplica", "Moderada"
    End Select
End Sub

Private Sub SetFila(ByRef tbl As Table, ByVal rowIdx As Integer, ByVal itemNum As String, ByVal desc As String, ByVal estado As String, ByVal prioridad As String)
    tbl.Cell(rowIdx, 1).Range.Text = itemNum
    tbl.Cell(rowIdx, 2).Range.Text = desc
    tbl.Cell(rowIdx, 3).Range.Text = estado
    tbl.Cell(rowIdx, 4).Range.Text = prioridad
End Sub
