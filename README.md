# Consolidador de tomas

Página web estática para juntar dos archivos del formato de captura antropométrica de Pardo Rubio. Funciona en el navegador con JSZip incluido en `vendor/`: los Excel y los datos personales nunca se envían a un servidor.

Si después de subir una actualización aparece un mensaje de una versión anterior, verifica que `index.html`, `app.js` y `engine.js` se hayan enviado juntos. El identificador de versión en las rutas de los scripts hace que el navegador descargue el código actualizado.

## Uso

1. Abrir la página publicada y seleccionar el archivo anterior y el actual. Para probarla en el computador antes de publicarla, ejecutar `python3 -m http.server 8000` dentro de esta carpeta y abrir `http://localhost:8000`.
2. Pulsar **Comparar archivos**. Las coincidencias inequívocas se asignan automáticamente. Resolver los casos que aparezcan en pantalla; se indican la hoja, la fila y las celdas B (NUIP), C (nombres), D (apellidos) y F (nacimiento). Una persona anterior no se puede asignar dos veces.
3. Pulsar **Ver consolidado**. Con **Editar** se pueden corregir los campos de identificación y ambas tomas; cada campo muestra su celda de origen. Las correcciones se aplican únicamente al archivo exportado y vuelven a ordenar el consolidado por primer nombre.
4. Pulsar **Descargar Excel consolidado**. Incluye continuidades, nuevos ingresos y egresos, todos en orden alfabético y con hasta 20 registros en cada hoja institucional.

## Publicación en GitHub Pages

Crear un repositorio nuevo (público en GitHub Free). Subir **los archivos de esta carpeta**, incluidos `vendor/`, al directorio raíz del repositorio (o a `docs/`). En **Settings → Pages**, elegir **Deploy from a branch**, `main`, y `/ (root)` (o `/docs`). Abrir la URL que indique GitHub. **No subir los Excel de los niños al repositorio:** los datos se cargan solo al usar la página. GitHub Pages publica sitios accesibles públicamente, incluso si el repositorio es privado en un plan que lo permita.

## Reglas

- Se leen todas las hojas excepto `Instrucciones` (también se reconoce la grafía `Instruccciones`). En cada hoja se ubica la tabla por su encabezado `No. de orden` y sus filas 1–20; las notas y leyendas al pie se omiten. Por eso funcionan las plantillas cuyos niños comienzan en la fila 15 o 16. Los campos de niño son A:AE.
- Coincidencia automática: NUIP único y nombres, apellidos y fecha de nacimiento presentes e iguales tras normalizar mayúsculas y espacios. Cualquier duda exige elección manual. La página propone candidatos por documento o identidad y muestra también todas las personas anteriores sin asignar.
- Continuidad: conserva Toma 1 anterior y **descarta todo lo que tenga la Toma 2 del archivo anterior**; usa Toma 2 actual (T:AE) si viene diligenciada, o bien la medición actual (H:S) como Toma 2. Si todos los datos diligenciados de la medición actual coinciden con Toma 1 anterior, Toma 2 queda vacía y la persona continúa, sin EGRESO. Los campos vacíos del archivo actual no borran datos anteriores. Datos personales: toma los del archivo actual.
- Nuevo: datos y Toma 1 del archivo actual (H:S). Egreso: datos y Toma 1 del anterior; `EGRESO` en la fecha de Toma 2, columna T.
- El resultado se ordena por el **primer nombre** (primera palabra de la columna NOMBRES); si coincide, se desempata con el nombre completo, los apellidos y el documento.
- El formato de salida se basa en el archivo **actual**, conservando encabezados, estilos, instrucciones y el resto de la plantilla. Hojas de captura adicionales se crean automáticamente cuando son necesarias; las hojas vacías sobrantes se retiran. Los campos auxiliares fuera de A:AE no pasan al formato.
- Algunas plantillas contienen tablas internas de Excel con rangos y encabezados fijos. El archivo exportado retira únicamente esas definiciones de tabla para evitar que Excel solicite repararlo después de reorganizar los niños. Se conservan las celdas, los estilos y las hojas de instrucciones.
- Si dos personas comparten documento en cualquiera de los archivos, la asignación deja de ser automática. No se permite que una persona anterior se use más de una vez.

**Nota:** La marca `EGRESO` refleja la ausencia en el archivo actual; conviene comprobar que este archivo incluya todas las hojas y todos los niños del periodo.
