
# Change Log
Documentación de cambios del proyecto.

## [ 1.9.0 ] - 2023/08/24

### Added

- Nueva herramienta: Registros Morph. Es un sistema de registros y anotaciones en documentos de Google Suite para hacer apuntes a nivel de administración y gestores de documentos (el sistema de comentarios de Google está más destinado a la comunicación directa entre usuarios y resolución de problemas).

### Changed

- El optimizador de documentos ha pasado a la sección "Ayuda".
- La sección Cuadros se mueve a la primera posición en la lista desplegable, para mover a primer plano el actualizador / congelador.
- Se ha mejorado el diseño de las tablas en el "Gestor de hojas" y "Registros Morph".

### Fixed

- Revisión de bugs en Morph Document Studio.

## [ 1.8.7 ] - 2023/08/02

### Added

- Función en desarrollo: Morph Internal Logger.

### Changed

- Optimizaciones de largo alcance en el código de Morph Document Studio (x3 reducción de tiempo de ejecución).

### Fixed

- Arreglos estéticos en la interfaz de Document Studio y Gestor de hojas.

## [ 1.8.6 ] - 2023/07/13

### Added

- Función añadida al gestor de hojas: refrescar hoja de plantilla, para sustituir una hoja por otra copiada de la plantilla sin estropear las referencias a la hoja original dentro del cuadro.
- Ahora se pueden añadir archivos .csv a la lista de archivos importados.
- Nuevo rol de usuario (formulaMod) para administradores de tablas.

### Fixed

- Arreglado el bug al cambiar el color de los Tabs.

## [ 1.8.5 ] - 2023/05/24

### Fixed

- Bugs arreglados y código optimizado en la herramienta para actualizar rangos nombrados.
- Notas adaptadas al horario de verano.

## [ 1.8.4 ] - 2023/05/19

### Added

- Nueva herramienta para eliminar todos los rangos nombrados / actualizar los rangos nombrados según la plantilla Morph.
- Botones "Formula W+/U-" para hacer "wrap" y "unwrap" de la fórmula de una celda (WIP)
- Listado de hojas conectadas en hoja LINK.

### Changed

- HOJA LINK MODULAR: las tres secciones de la hoja Link pueden actualizarse independientemente. El código del actualizador se ha modificado para adaptarse a este nuevo formato.
- Se ha eliminado el botón del histórico de superficies y se ha integrado en las opciones del actualizador.
- Eliminado botón para crear lista de miembros Morph.
- Eliminada la herramienta para borrar filas innecesarias.
- Cambios estéticos en la herramienta del gestor de hojas.

### Fixed

- Arreglado el error en el botón para limpiar la caché del documento.

## [ 1.8.3 ] - 2023/03/06

### Added

- Al actualizar/congelar, creará la carpeta automáticamente si no encuentra la carpeta de exportación/congelados.

### Fixed

- Arreglado bug en el supercongelador.

## [ 1.8.2 ] - 2023/03/01

### Added

- Botón para reformatear la hoja LINK.
- Botón para limpiar la caché del documento en el apartado "Optimizar documento".

### Changed

- Cambio de la imagen del logo principal, letras en texto en lugar de en imagen.
- Cambio general de organización y diseño en la sección "Cuadros".

### Fixed

- Bug con los datos "backup" de la hoja LINK a la hora de actualizar a la última versión.

## [ 1.8.1 ] - 2023/01/22

### Fixed

- Bugs generales en el manejo de cuadros debido a los cambios en la última actualización.

## [ 1.8.0 ] - 2023/01/18

### Added

- Añadido botón para crear marcas automáticas en el historial de superficies construidas.
- Añadida opción en el actualizador de cuadros para mantener la visibilidad de las hojas.
- Nueva función en el gestor de hojas para mover hojas individualmente.
- Añadidos iconos de interrogación para ofrecer ayuda contextual junto a las checkbox de opciones.
- Añadido panel de información sobre cuadros Morph.
- Nueva herramienta: MODIFICAR FÓRMULAS. Genera automáticamente un registro de las fórmulas modificadas en cuadros.
- Debug Report: opción para desarrolladores que facilita la detección de errores en cuadros.
- Nueva herramienta para borrar filas en masa (desde la pestaña Cuadros).

### Changed

- Simplificación de botones en la sección Google Sheets.
- Cambio de nombre de la sección "Superficies" por "Cuadros".
- Cambio de diseño en la sección "Cuadros". La casilla "Actualizar hoja LINK" activa los prefijos y ejecuta la función completa (búsqueda de carpetas, panel de control, formato de hoja, etc.) mientras que tenerla desactivada ejecuta un código simple leyendo las IDs directamente de la hoja LINK. De este modo se pueden optimizar los tiempos de ejecución.
- Se elimina de la vista general la herramienta "Transición de cuadros". Solo podrán acceder a ella los desarrolladores.
- Las celdas "Hojas conectadas" de la hoja LINK ahora se muestran en las columnas E-F.
- Cambios de diseño en el gestor de hojas, añadiendo un apartado específico para reordenar hojas.
- La función "Lista de miembros Morph" ya no aplica formato, solo añade los datos en crudo y a partir de la celda seleccionada.
- Eliminada la herramienta para modificar la altura de celdas, por redundancia con las posibilidades de Google Sheets.
- Gran optimización de código para el congelador Morph, reduciendo tiempos de espera.

### Fixed

- Arreglado el link a la hoja conectada en la hoja LINK con la herramienta "Conectar G-Sheets".
- Optimización del código para actualizar cuadros para solventar el fallo de importación en cuadros grandes.
- Arreglado el fallo al buscar el panel de control en el cuadro de mediciones.
- Arreglado bug al usar nota en celda A1 para el listado de archivos.
- Arreglado bug al usar la función para eliminar filas/columnas vacías.
- Cambios menores de diseño.

## [ 1.7.1 ] - 2022/12/27

### Fixed

- Eliminado botón innecesario en la sección Google Sheets.
- Bug en el display de imágenes en Document Studio.

## [ 1.7.0 ] - 2022/12/18

### Added

- Nuevas funciones en el gestor de hojas:
  - Opción para ocultar / mostrar hojas.
  - En la tabla se añade un botón para ir rápidamente a cada hoja.
  - Ahora la lista de hojas puede ordenarse manualmente mediante Drag & Drop, pudiendo reordenarse las hojas de un documento a través del botón "Reordenar".
  - Nueva herramienta: congelador parcial, para crear un documento congelado solo con las hojas seleccionadas.
- Nueva herramienta: Crear PDF Secuencial. Imprime múltiples PDF basados para cada opción de una lista desplegable en Google Sheets.

### Changed

- En el gestor de hojas las hojas ocultas ahora aparecen en color gris claro.
- Cambios estéticos y optimización del código de la sección de desarrollo.

## [ 1.6.1 ] - 2022/11/04

### Fixed

- Optimización del problema de importación TXT con archivos grandes. El nuevo código consigue mejores tiempos en la actualización de cuadros.

## [ 1.6.0 ] - 2022/11/04

### Added

- Herramienta para ajustar la altura de las filas seccionadas.
- Añadida opción en el "listado de archivos" para insertar imágenes de una carpeta.
- Actualizador para el cuadro de mediciones.
- Nueva sección Interoperabilidad, con las siguientes herramientas:
  - Herramienta para conectar hojas de Google Sheets.
  - Herramienta para importar rápidamente múltiples archivos CSV a una hoja de Google Sheets, especificando la celda de inserción.
- Herramienta para crear marcas temporales en el histórico de superficies.

### Changed

- La herramienta de listado de archivos empieza el listado en la celda seleccionada.
- Retirada la herramienta para crear estructuras de carpetas (temporalmente).
- Eliminada la paginación en el gestor de hojas: ahora aparece toda la lista de una.

### Fixed

- Arreglado error al importar TXT vacíos en el actualizador de cuadros.

## [ 1.5.3 ] - 2022/10/18

### Fixed

- Arreglada búsqueda "Case Sensitive" del panel de control y carpeta de superficies.
- Selector de hojas para la herramienta Importar CSV.

## [ 1.5.2 ] - 2022/10/07

### Added

- Se ha añadido un botón temporal para adaptar los cuadros de superficies anteriores al nuevo sistema de actualización.

## [ 1.5.1 ] - 2022/10/03

### Added

- Opción para añadir todos los archivos en el sector de prefijo del actualizador de cuadro de superficies.

### Changed

- Arreglos estéticos en la tabla de Sheet Manager.
- Arreglos en el input de claves de la sección de desarrollo.

### Fixed

- Al actualizar el cuadro de superficies se eliminan los archivos .txt duplicados en la carpeta de exportaciones.
- Mantener selección en la tabla de Sheet Manager al cambiar de página.

## [ 1.5.0 ] - 2022/09/28

### Added

- Nueva función: generar identificadores únicos (en todas las celdas seleccionadas).
- Añadida sección para desarrolladores en la página de Gsuite Morph Tools. Se puede acceder con permiso directo desde la base de datos o clave maestra.
- Nueva herramienta: gestor de hojas para Google Sheets. Permite eliminar, ocultar, limpiar o duplicar hojas en masa, reduciendo considerablemente el tiempo necesario para realizar estas tareas.
- Morph Document Studio:
  - Añadidas opciones para guardar / borrar la configuración de Document Studio en el documento.
  - Añadido botón para refrescar los campos de selección de columnas.
- Nueva función: Font Size +, Font Size -, para incremental proporcionalmente el tamaño de letra.
- Nueva función: Fit Filas + Fit Columnas para mejorar visualmente la estructura de la hoja de cálculo.

### Changed

- Los botones de funciones se ordenan alfabéticamente de forma automática.
- Ahora las funciones directas no muestran mensaje al ejecutarse correctamente.
- Arreglos menores de diseño (tipografía y colores).
- Modificado el editor visual de Document Studio para mostrar todas las opciones (sin menú adicional).
- Añadido un selector de prefijo para los documentos del actualizador del cuadro de superficies.
- Nuevo logo para la aplicación con tamaño reducido.

### Fixed

- Arreglado el error al congelar documentos vinculados a un formulario de Google.
- Arreglados los problemas al ocultar elementos de Document Studio en función de las opciones seleccionadas.

## [ 1.4.0 ] - 2022/08/10

### Added

- Primera versión implementada públicamente en el servidor de Morph.
- Añadido un menú desplegable interno como encabezado del complemento. El logo aparece dentro de este encabezado.
- Nueva herramienta: exportar TSV con fórmulas.
- Nota informando de la última vez actualizado / congelado en el cuadro de superficies.

### Changed

- Completada la plantilla para el generador de estructura de carpetas.
- Actualizada la página Gsuite Morph Tools de Morph Pills.
- Las funciones similares (p.ej: congelador / supercongelador) se han optimizado en una sola función que reconoce desde qué botón se ejecuta.
- Cambiado el diseño de la plantilla del cuadro de superficies.

### Fixed

- Arreglado: la pantalla de carga no rellena la barra lateral.
- Arreglado problema de variables en onOpen(e) impide mostrar el menú en la aplicación publicada.
- Arreglado error al acceder al documento al actualizar cuadro de superficies.

## [ 1.3.0 ] - 2022/07/27

### Added

- Primera versión funcional de Document Studio by Morph:
  - Opción para añadir marcadores automáticamente a Google Sheets.
  - Añadida programación para que la aplicación funcione con plantillas en Google Slides.
  - Las imágenes en Document Studio Docs por defecto se adaptan al ancho del documento, pero se ha añadido un snippet para elegir ancho de imagen **LINK{w=400}**, siendo 400px el ancho de imagen requerido.
- Nueva función: listar archivos de una carpeta.
- Añadido método manual para actualizar / congelar en casos especiales de proyecto.
- Generación de XLSX automático para archivos congelados.
- Congelador rápido (desarrollo en fase Beta).
- Archivos añadidos al repositorio en Github.
- Creada la página de Morph Pills para la sección de ayuda.

### Changed

- Implementada pantalla de carga al abrir las secciones de Gsuite Morph Tools.
- El changelog se ha trasladado de la barra lateral a un documento del repositorio.

### Fixed
 
- Se pueden añadir links a imágenes externas dentro de Document Studio.
- Eliminados los pasos intermedios en el generador de estructura de carpetas, ahora funciona con un solo clic.

## [ 1.2.0 ] - 2022/06/30

### Added

- Nueva función: 'supercongelador', un congelador genérico para cualquier tipo de archivo.
- Nueva herramienta: contador de celdas disponible en el menú Ayuda.
- Nueva función: lista de usuarios extraída del servidor de Morph.
- Generador de plantillas Document Studio by Morph: creación de documentos en masa a partir de plantillas en Google Docs (Fase inicial de desarrollo).
- Nueva función: eliminar todas las filas y columnas vacías.
- Nueva función: añadir paleta de colores Morph.
- Nueva función: generador de estructuras de carpetas.

### Changed

- Actualizar cuadro de superficies:
  - Abstracción del código: los botones para actualizar / congelar cuadro ya no necesitan introducir ningún dato para funcionar.
  - Busca el archivo Panel de Control en cualquier lugar dentro de la estructura directa de carpetas del proyecto.
  - Añade a la lista de archivos exportados todos los archivos con prefijo TXT de la carpeta Exportaciones (Sin límite de archivos).

### Fixed
 
- Arreglados los fallos principales del código de congelación, ya no genera error crítico por límite de celdas.
- El cuadro de superficies otorga automáticamente los permisos de ImportRange.

## [ 1.1.0 ] - 2022/06/15

### Added

- Configuración del proyecto en Google Cloud Platform.
- Gsuite Morph Tools publicado en Google Marketplace.
- Implementación de la interfaz básica Morph Tools.
- Añadidos botones del cuadro de superficies (Actualizar / congelar).