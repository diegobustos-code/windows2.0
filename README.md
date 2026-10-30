# Asistencia Zoom (Windows)

Programa de escritorio para Windows que cruza los reportes de participantes de
Zoom (CSV) con el listado oficial de socios (Excel) y entrega la asistencia de
cada socio como **P** (presente), **A** (ausente) o **J** (ausencia justificada).

Nació como una app Android y se portó a Windows. Este documento resume cómo se
usa, cómo decide el cotejo y todo lo que se ha ido agregando, de la primera
versión a la última.

---

## Qué puedes hacer

Al abrir el programa aparece una ventana para elegir una de tres opciones:

| Opción | Para qué sirve | Resultado |
|---|---|---|
| **Asistencia de UN mes** | Cotejar un solo CSV de Zoom | Excel del mes, con pestañas de revisión |
| **Reporte Multi-Mes** | Cotejar varios meses a la vez (o todo el año) | Un Excel con una fila por socio y una columna por mes |
| **Rellenar Listado Existente** | Completar el listado que ya llevas a mano | Una copia de tu listado con el mes rellenado |

Los archivos se guardan en `Descargas / Asistencia Zoom / Cotejados` (y los
pendientes en `Por revisar`). Cada pop-up de guardado trae un botón para abrir
el archivo y otro para abrir la carpeta.

---

## Formato de los reportes

Es el mismo en los tres modos:

- Una fila por socio, **en el mismo orden del listado oficial**.
- Cada fila lleva el **color de su sede** (los mismos colores del listado).
- Celdas de asistencia, todas en negrita:
  - **P**: letra verde, sin relleno propio (se ve el color de la sede).
  - **A**: letra roja sobre relleno rojo suave.
  - **J**: letra naranja sobre relleno amarillo.
- Lo que hay que revisar a mano va en **otras pestañas del mismo libro**:
  - **Ambiguos**: participantes que coinciden con más de un socio; incluye los
    posibles socios para decidir.
  - **No Encontrados**: participantes que no calzan con nadie.
  - **Cotejos aproximados**: personas que el programa sí asignó, pero sin nombre
    y apellido exactos (por ejemplo, solo el nombre de pila, o una sede que
    desempató). Ya cuentan en la asistencia; se listan para que las confirmes.
  - **Revisar diferencias** (solo en Rellenar Listado): ver más abajo.

El botón **Pendientes** del modo mensual genera un Excel con esas mismas
pestañas (antes era un CSV que Excel abría desordenado).

---

## Cómo decide el cotejo

El nombre que la gente escribe en Zoom es poco confiable, así que se usa el
listado oficial como fuente de verdad. Todo ignora tildes y mayúsculas. El
orden de decisión es:

1. **Por apellido.** El apellido (paterno o materno) debe aparecer en el nombre
   de Zoom. Entre quienes comparten apellido se prefiere a quien además coincide
   en el nombre de pila.
2. **Solo apellido.** Si nadie coincide en el nombre, se asigna únicamente cuando
   ese apellido lo lleva un solo socio, la sede escrita (si hay) no lo contradice
   y el nombre de pila escrito no es el de otra persona del listado. Si no,
   queda pendiente.
3. **Solo nombre de pila.** Si nadie coincide por apellido, se busca por nombre.
   Si es único en el listado, se asigna.
4. **Ortografía parecida e iniciales.** Un error de tipeo en el apellido
   ("gonzales" por "Gonzalez", "Irribarren" por "Iribarren") o iniciales
   ("Ana Carolina RM" por Rojas Millán), siempre junto a un nombre que coincida.
   También acepta pequeñas diferencias en el nombre de pila ("Gisselle" por
   "Giselle").
5. **Desempate por sede.** Si hay empate, se usa la sede que la persona escribió
   en Zoom ("Maria Fernanda Cpech Viña" → socia de Viña del Mar).
6. **Nombre escrito todo junto.** "Mariateresacortes" se separa en María Teresa
   Cortés cuando todas las partes existen en el listado.
7. **Si nada funciona**, el participante queda en *No Encontrados* o *Ambiguos*.
   Nunca se adivina entre varias personas.

Otras reglas:
- Si un socio entra varias veces a la reunión, sus minutos se **suman**.
- Es **Presente (P)** quien acumula al menos el umbral de minutos (30 por
  defecto, editable en pantalla).
- La sede se detecta desde el nombre de Zoom con una lista de sedes y variantes
  ("sn bdo", "pto montt", "pte alto", "viña", "valpo", "losangeles", entre otras).
  Si aparece una sede nueva, se agrega en `KNOWN_SEDES` de `csv_processor.py`.

---

## Reporte Multi-Mes

1. Carga uno o varios CSV de Zoom (por ejemplo, un año completo de una vez).
2. El mes se detecta desde el nombre del archivo (`Julio 2026`, `sep2026`,
   `participants_..._2026_07_23.csv`, etc.). **Siempre se te pide confirmarlo**,
   y puedes corregirlo con doble clic, porque hay archivos con nombres irregulares
   (por ejemplo, reuniones extraordinarias).
3. Carga el listado de socios y genera el Excel.

---

## Rellenar Listado Existente

Para el listado que ya llevas a mano, al que le agregas una columna por mes y
escribes "J" en las ausencias justificadas. El programa **rellena** la columna
del mes en vez de generar un listado nuevo.

- **Respeta las J** por defecto, con una excepción: la casilla **"Cambiar J a P si el Zoom dice que estuvo presente"** (activada de fábrica) corrige las J puestas por error cuando el Zoom muestra que la persona sí cumplió el mínimo de minutos. Cada cambio queda anotado en la pestaña "Revisar diferencias". Si la desmarcas, todas las J se mantienen. Una J con Zoom ausente siempre se conserva.
- Reconoce el mes aunque el encabezado tenga texto extra (`Octubre
  (extraordinaria)`, `Asistencia Agosto`). Si el mes no tiene columna, la crea al
  final.
- **Ignora la columna "Sede"/"Sedes"** y las columnas que no son de mes (como
  "Licencias").
- Mantiene el formato de tu archivo y guarda siempre una **copia**; tu original
  no se modifica.
- Si el archivo trae varias hojas, usa la primera cuyo formato de columnas
  reconoce.
- Acepta uno o varios meses a la vez.
- Reformatea al nuevo estilo (colores de P, A y J) también las celdas que ya
  estaban llenas de meses anteriores.
- Repara automáticamente un problema típico de archivos exportados desde Google
  Sheets (un filtro de tabla desalineado) que hacía que Excel mostrara "problema
  con el contenido".

Pestañas que agrega: *No Encontrados*, *Ambiguos*, *Cotejos aproximados* y
**Revisar diferencias**. Esta última lista a quienes el Zoom marca Presentes pero
el listado tenía como **J** (se respeta o se cambia a P según la casilla anterior) o **A**, y a quienes pasaron de
**P a A**, con sus minutos y qué hizo el programa, para corregir a mano si
corresponde. Ojo: solo las J están protegidas; una A o una P ya escrita se
actualiza al resultado de Zoom, pero queda anotada en esa pestaña.

---

## Formatos de listado de socios aceptados

El orden de las columnas no importa; se detectan por el título del encabezado.

1. **Formato completo:** `Sede`, `Apellido Paterno`, `Apellido Materno`, `Nombres`.
2. **Formato simple:** `Sede` y `Apellido` (el apellido completo en una columna),
   más una columna con el nombre, con cualquier título.

`Sede`/`Sedes` y `Apellido`/`Apellidos` se consideran equivalentes.

---

## Historial de cambios

Del más reciente al más antiguo.

### Cotejo más intensivo y exportación unificada
- Cotejo por aproximación: nombre escrito junto, apellido con otra ortografía,
  iniciales, desempate por sede y por nombre parecido.
- Corrección de un error antiguo: un apellido repetido escrito sin nombre le
  asignaba los minutos a otro socio con ese apellido (inflaba duraciones y
  podía marcar presentes por error). Ahora esos casos se asignan bien o quedan
  pendientes.
- Apellido completo en una sola columna: ya no se penaliza a quien tiene dos
  apellidos cuando en Zoom se escribe solo uno.
- Nuevas variantes de sede (viña, valpo, losangeles) y "Cpech" ignorado como
  palabra del nombre.
- Los tres modos exportan en el mismo orden del listado, con color por sede.
- Un solo libro con pestañas *Ambiguos*, *No Encontrados* y *Cotejos
  aproximados*; "Pendientes" ahora es Excel.
- *Ambiguos* muestra los posibles socios.
- Rellenar Listado: nueva pestaña *Revisar diferencias*.
- Reporte Multi-Mes: las filas se identifican por posición, así que dos socios
  con los mismos datos ya no se fusionan.
- Archivos nuevos: `sede_colors.py`. Modificados: `roster_matcher.py`,
  `csv_processor.py`, `excel_exporter.py`, `multi_month_processor.py`,
  `multi_month_window.py`, `roster_sheet_filler.py`, `roster_fill_window.py`,
  `desktop_app.py`.

### Nuevo formato de celdas
- P en verde sin relleno propio, A en rojo con relleno rojo suave, J en naranja
  con relleno amarillo; todo en negrita.

### Reparación del archivo "dañado" al rellenar
- Corrige el aviso de Excel por una tabla con filtro desalineado en listados
  exportados desde Google Sheets.

### Columnas "Apellidos" y "Sedes"
- El detector de formato acepta singular y plural.

### Rellenar Listado Existente
- Función nueva (ver arriba). Archivos: `roster_sheet_filler.py`,
  `roster_fill_window.py`; en `multi_month_processor.py` se agregó
  `extract_month_name_only`.

### Segundo formato de listado
- Se acepta el formato simple (`Sede`, `Apellido` y una columna de nombre) además
  del formato original.

### Cotejo por nombre de pila y apellido flexible
- Sin tildes (ya funcionaba) y, además, coincidencia por nombre de pila cuando es
  único en el listado; si lo comparten varios socios, va a *Ambiguos*.
- Hojas separadas *No Encontrados* y *Ambiguos* dentro del Excel Multi-Mes.

### Ventana de inicio
- Al abrir, se elige entre "Asistencia de UN mes" y "Reporte Multi-Mes" en vez de
  pedir siempre el mes.

### Reporte Multi-Mes
- Varios CSV de Zoom (un año completo, si quieres) en un solo Excel con una
  columna P/A por mes. Detección del mes con confirmación obligatoria.
- Archivos nuevos: `multi_month_processor.py`, `multi_month_window.py` y
  `ui_common.py` (colores y ventanas compartidos, para evitar importaciones
  circulares).

### Compilación a .exe
- GitHub Actions (`build-exe.yml`) compila el .exe con PyInstaller en modo
  carpeta (`--onedir`), que genera menos falsos positivos de antivirus que
  `--onefile`.
- Corrección del error de arranque `expected integer but got "9.5"` (tamaños de
  letra fraccionarios en Tkinter).

### Versión inicial (port desde Android)
- Interfaz Tkinter para Windows con la misma lógica de la app Android: mes
  obligatorio, abrir CSV de Zoom y listado de socios, cotejar, buscar, filtrar
  por duración mínima, umbral de "Presente" editable y orden; exportar a CSV,
  Excel y Pendientes.

---

## Archivos del proyecto

```
├── desktop_app.py              # Ventana principal y exportación del modo mensual
├── multi_month_window.py       # Ventana del Reporte Multi-Mes
├── multi_month_processor.py    # Combina varios meses; detecta el mes por nombre de archivo
├── roster_fill_window.py       # Ventana de Rellenar Listado Existente
├── roster_sheet_filler.py      # Lógica de rellenar y revisar un listado existente
├── roster_matcher.py           # Carga del listado de socios y algoritmo de cotejo
├── csv_processor.py            # Lectura del CSV de Zoom y detección de la sede
├── excel_exporter.py           # Genera los Excel (formato, pestañas de revisión)
├── sede_colors.py              # Color de fila por sede
├── ui_common.py                # Colores, pop-ups y botones compartidos
├── logo.png / logo.ico         # Logo (el .ico es el ícono del .exe)
├── requirements-desktop.txt    # Dependencias para compilar
└── .github/workflows/build-exe.yml   # Compila el .exe en GitHub
```

---

## Cómo obtener el .exe

1. Sube todos los archivos al repositorio (con `build-exe.yml` dentro de
   `.github/workflows/`). Si falta alguno, el .exe falla al abrir con un error
   `No module named ...`.
2. En GitHub, pestaña **Actions** → **Compilar EXE de Windows** → **Run
   workflow**. Tarda unos 3 a 5 minutos.
3. Descarga el artefacto `AsistenciaZoom-EXE` (viene en .zip) y descomprímelo.
   Dentro hay una carpeta `AsistenciaZoom`: copia **la carpeta completa** a
   cualquier PC con Windows y abre `AsistenciaZoom.exe`. No separes el .exe del
   resto de los archivos. No hace falta instalar Python.

Para probarlo sin compilar, con Python instalado:

```
pip install -r requirements-desktop.txt
python desktop_app.py
```

### Windows Defender / SmartScreen

El .exe no está firmado digitalmente, así que la primera vez Windows puede
mostrar "Windows protegió su PC". Se soluciona con **Más información** →
**Ejecutar de todas formas**. Si el antivirus lo bloquea, se puede enviar el
archivo a Microsoft para que lo revisen como falso positivo
(*Submit a file for malware analysis*).
