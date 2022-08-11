
# Change Log
Todos los cambios en el proyecto serán documentados en este archivo.

## [ 1.5 ] - 2022/08/10

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

## [ 1.3 ] - 2022/07/27

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

## [ 1.2 ] - 2022/06/30

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

## [ 1.0 ] - 2022/06/15

### Added

- Configuración del proyecto en Google Cloud Platform.
- Gsuite Morph Tools publicado en Google Marketplace.
- Implementación de la interfaz básica Morph Tools.
- Añadidos botones del cuadro de superficies (Actualizar / congelar).