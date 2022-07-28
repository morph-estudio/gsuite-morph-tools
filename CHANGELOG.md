
# Change Log
Todos los cambios notables en el proyecto serán documentados en este archivo.
 
## [1.0] - 2022-06-15
 
### Added
- Configuración del proyecto en Google Cloud Platform.
- Gsuite Morph Tools configurado y publicado en Google Marketplace.
- Implementación de la interfaz Morph Tools.
- Añadidos botones del cuadro de superficies (Actualizar/Congelar).
- Nueva función: eliminar todas las filas y columnas vacías.
- Nueva función: añadir paleta de colores Morph.
- Nueva función: generador de estructuras de carpetas.
 
## [1.1] - 2022-06-30
 
### Added

- Nueva función: 'Supercongelador', un congelador genérico para cualquier tipo de archivo.
- Generación de XLSX automático para archivos congelados.
- Nueva herramienta: contador de celdas disponible en el menú Ayuda.
- Nueva función: lista de usuarios Morph extraída del servidor de Morph.
- Generador de plantillas Document Studio by Morph: creación de documentos en masa a partir de plantillas en Google Docs (fase beta).
 
### Changed

- Actualizar cuadro de superficies:
  - Los botones para Actualizar / Congelar cuadro ya no necesitan introducir ningún dato de forma manual para funcionar.
  - Busca el archivo Panel de Control en cualquier lugar dentro de la estructura de carpetas del proyecto.
  - Añade a la lista de archivos exportados todos los archivos con prefijo TXT de la carpeta Exportaciones (sin límite).

### Fixed
 
- Arreglados bug mayor en el código de congelación, ya no genera fallo por celdas máximas.
- El cuadro de superficies otorga automáticamente los permisos de ImportRange.

## [1.3] - 2022-07-27
 
### Added

- Primera versión funcional de Document Studio by Morph:
  - Opción para añadir marcadores automáticamente a Google Sheets.
  - Añadido el código para que la aplicación funcione con plantillas en Google Slides.
  - Las imágenes en Document Studio Docs por defecto se adaptan al ancho del documento, pero se ha programado un código para elegir ancho de imagen **LINK{w=400}**, siendo 400px el ancho de imagen requerido.
- Nueva función: listar archivos de una carpeta determinada.
- Añadido método manual para actualizar / congelar para casos especiales de proyecto.
- Congelador rápido en fase beta (72" vs 240").
- Archivos añadidos al repositorio en Github.
 
### Changed

- Implementada pantalla de carga al abrir los menús de Gsuite Morph Tools.
- El changelog se ha trasladado de la barra lateral de GMT a un documento del repositorio.

### Fixed
 
- Se pueden añadir links a imágenes externas dentro de Document Studio.
- El generador de carpetas ahora funciona automáticamente sin introducir datos intermedios.