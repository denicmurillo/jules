# Lista de Chequeo para Instalaciones Eléctricas en Microsoft Word (.docm)

Este proyecto proporciona una plantilla dinámica habilitada para macros de Microsoft Word (`.docm`) diseñada para la evaluación de instalaciones eléctricas según la normativa NFPA 70 / Código Eléctrico Nacional (NEC 2020).

## Características Principales

1. **Documento Dinámico e Interactivo:**
   - Contiene un **Panel de Control** en el cuerpo del documento con 23 casillas de verificación (Controles de Contenido de Casilla de Verificación / Checkboxes).
   - Al marcar una casilla, el bloque correspondiente (título, tabla de chequeo con items, prioridades y campo de observaciones) se inserta dinámicamente en el documento en el orden predeterminado.
   - Al desmarcar una casilla, el sistema muestra una advertencia al usuario: *"ADVERTENCIA: ¿Está seguro de que desea eliminar la sección '[Nombre]'? Esta acción no se puede deshacer."*
   - Si el usuario selecciona *"No"*, la casilla se re-marca automáticamente y el contenido permanece intacto. Si selecciona *"Sí"*, la sección y su tabla asociada se eliminan sin dejar rastro.

2. **Detección Automática de Eventos:**
   - La plantilla utiliza el evento nativo `Document_ContentControlOnExit` en `ThisDocument`. No requiere presionar botones de acción adicionales; al interactuar con las casillas de verificación, la macro responde de manera inmediata.

3. **Estructura Fija y Orden Riguroso:**
   - **Encabezado Fijo:** Se mantiene siempre visible en la parte superior del documento con la sección *"CLASIFICACIÓN DE NO CONFORMIDADES"* y la tabla explicativa con código de colores para las prioridades (**Alta**, **Moderada**, **Leve**).
   - **Orden Predeterminado:** Independientemente del orden en que el inspector seleccione o desmarque las casillas, las secciones se posicionan y organizan dinámicamente manteniendo la secuencia numérica del 1 al 23.

---

## Secciones Incluidas (23 Categorías)

1. **Documentación**
2. **Transformador principal**
3. **Sistema de puesta a tierra**
4. **Sistema contra descargas atmosféricas**
5. **Espacios físicos cuartos eléctricos**
6. **Transformadores de baja tensión**
7. **Tableros y protecciones**
8. **Conductores**
9. **Canalizaciones, bandejas, cajas de registro y conexiones**
10. **Iluminación**
11. **Tomacorrientes y apagadores**
12. **Equipos de aire acondicionado, ventilación y refrigeración**
13. **Motores**
14. **UPS**
15. **Sistemas de emergencia**
16. **16.1. Lugares Clase I**
17. **16.2. Lugares Clase II**
18. **16.3. Lugares Clase III**
19. **Sitios de Reunión**
20. **Teatros y lugares similares**
21. **Carnavales, circos, ferias y lugares similares**
22. **Piscinas**
23. **Instalaciones de asistencia médica**

---

## Archivos del Proyecto

- `Lista_de_Chequeo_Electrica.docm`: Plantilla principal habilitada para macros lista para usar en Microsoft Word.
- `src/ThisDocument.cls`: Código VBA del módulo de clase del documento (manejador de eventos).
- `src/modChecklist.bas`: Código VBA del módulo estándar con la lógica de inserción, eliminación y formateo de tablas.
- `build_template.py`, `ole_writer.py`, `generate_vba_bin.py`, `build_docm.py`: Scripts automatizados para la generación y ensamblaje de la plantilla OpenXML.
- `tests/test_docm.py`: Suite de pruebas automatizadas de verificación de estructura y macros.

---

## Instrucciones de Uso

1. Abra el archivo `Lista_de_Chequeo_Electrica.docm` en Microsoft Word (Windows).
2. Asegúrese de **Habilitar Contenido / Habilitar Macros** en la barra amarilla de seguridad de Word.
3. En la sección **ASPECTOS A EVALUAR**, marque la casilla de verificación del área que desea evaluar (por ejemplo, `[X] 22. Piscinas`).
4. La sección con su tabla de items y observaciones aparecerá automáticamente debajo en el documento.
5. Si desmarca la casilla, confirme la advertencia si desea eliminar la sección.
