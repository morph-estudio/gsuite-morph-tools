// JSON TEMPLATE DATA

/**
 * Devuelve el ID de diferentes archivos clave de Morph.
 *
 * @param {string} file - El nombre del archivo para el cual se desea obtener el ID.
 * @return {string} El ID del archivo clave correspondiente.
 */
function naveNodrizaIDS(file) {
  switch (file) {
    case 'Cuadro Superficies':
      return '1_Qq8y_cC5V9lSThCypq0Qpbj6JXrtv-QBBlt7ghy0pI';
    case 'Panel de control':
      return '1UlBPMINIHB1rCY9YNJqW0ia5sCSx93XUgPnvrW36_Hk';
    case 'Exportación Superficies':
      return '1t7r-aTKL9nrjgMJosloOMfoGqt3R1SJC2MthLOL9IFo';
    case 'Cuadro Mediciones':
      return '1Sy9ch1kQMhhfRR0_pyQXKQASS0RBjVIem85e3skjAi0';
    default:
  }
}

/**
 * Devuelve el nombre clave de diferentes archivos de plantillas Morph.
 *
 * @param {string} file - El nombre del archivo para el cual se desea obtener el ID.
 * @return {string} El nombre del archivo clave correspondiente.
 */

function getTipoArchivo(file, index) {
  let nombreArchivo = index === undefined
    ? (file || SpreadsheetApp.getActive()).getName().toLowerCase()
    : "undefined";

  const tiposArchivo = {
    1: 'Panel de control',
    2: 'Cuadro Superficies',
    3: 'Exportación Superficies',
    4: 'Cuadro Mediciones'
  };

  for (const [key, value] of Object.entries(tiposArchivo)) {
    if (nombreArchivo.includes(value.toLowerCase()) || index === parseInt(key)) {
      return value;
    }
  }

  return "none";
}

function getCuadroTypes() {
  var cuadroTypes = {
    'selectPanelControl': 1,
    'selectCuadroSuperficies': 2,
    'selectCuadroExportacion': 3,
    'selectCuadroMediciones': 4
  };
  return cuadroTypes;
}

/**
 * Obtiene el objeto JSON de configuración de la hoja de plantilla.
 * Este objeto contiene información sobre tipos de proyecto, hojas maestras y hojas secundarias.
 *
 * @return {Object} El objeto JSON de configuración de la hoja de plantilla.
 */
function templateSheetConfigObject(booleanGetTipoArchivo, tipoArchivo) {

  var templateSheetConfigObject;

  tipoArchivo = booleanGetTipoArchivo ? getTipoArchivo() : tipoArchivo;

  // Browser.msgBox(tipoArchivo);

  switch (tipoArchivo) {
    case 'Panel de control':

      templateSheetConfigObject = {
        "fileType": [ tipoArchivo ],
        "projectTypes": [
          {
            "nombre": "Plurifamiliar",
            "code": "plurifami",
            "settings": {
              "minimumCuadro": ["instrucio", "datosproy", "contactos", "registros", "planifica", "registros", "pconhoras", "duedilige", "premiscli", "cosdesglo", "cosprecio", "xbddproye", "xcontacto", "xcuadrsup", "xempresas", "xhorastar", "xmasuputi", "xnormativ", "xproyecoo", "xrosciuda", "xroscimix", "xrosedren", "xvariable", "xvarianor", "xvariacos", "xvariapyt"],
              "fullCuadro": ["instrucio", "datosproy", "contactos", "registros", "planifica", "registros", "pconhoras", "duedilige", "panconkpi", "cosdesglo", "cosprecio", "premiscli", "predimmix", "predimsup", "predimcos", "xbddproye", "xcontacto", "xcuadrsup", "xempresas", "xhorastar", "xmasuputi", "xnormativ", "xproyecoo", "xrosciuda", "xroscimix", "xrosedren", "xvariable", "xvarianor", "xvariacos", "xvariapyt"],
              "revitCuadro": [],
              "hiddenSheets": ["contactos", "xbddproye", "xcontacto", "xcuadrsup", "xempresas", "xhorastar", "xmasuputi", "xnormativ", "xproyecoo", "xrosciuda", "xroscimix", "xrosedren", "xvariable", "xvarianor", "xvariacos", "xvariapyt"]
            }
          },
          {
            "nombre": "Unifamiliar",
            "code": "unifamili",
            "settings": {
              "minimumCuadro": ["instrucio", "datosproy", "contactos", "registros", "planifica", "registros", "pconhoras", "duedilige", "premiscli", "cosdesglo", "cosprecio", "xbddproye", "xcontacto", "xcuadrsup", "xempresas", "xhorastar", "xmasuputi", "xnormativ", "xproyecoo", "xrosciuda", "xroscimix", "xrosedren", "xvariable", "xvarianor", "xvariacos", "xvariapyt"],
              "fullCuadro": ["instrucio", "datosproy", "contactos", "registros", "planifica", "registros", "pconhoras", "duedilige", "panconkpi", "cosdesglo", "cosprecio", "premiscli", "predimmix", "predimsup", "predimcos", "xbddproye", "xcontacto", "xcuadrsup", "xempresas", "xhorastar", "xmasuputi", "xnormativ", "xproyecoo", "xrosciuda", "xroscimix", "xrosedren", "xvariable", "xvarianor", "xvariacos", "xvariapyt"],
              "revitCuadro": [],
              "hiddenSheets": ["contactos", "xbddproye", "xcontacto", "xcuadrsup", "xempresas", "xhorastar", "xmasuputi", "xnormativ", "xproyecoo", "xrosciuda", "xroscimix", "xrosedren", "xvariable", "xvarianor", "xvariacos", "xvariapyt"]
            }
          }
        ],
        "masterSheets": [
          {
            "name": "Instrucciones",
            "code": "instrucio",
            "description": "Datos mínimos e información sobre el panel de control",
            "tabColor": "#ffff00",
            "settings": {}
          },
          {
            "name": "Datos proyecto",
            "code": "datosproy",
            "description": "Datos generales del proyecto con id único",
            "tabColor": "#1155cc",
            "settings": {}
          },
          {
            "name": "Contactos",
            "code": "contactos",
            "description": "Contactos del proyecto asignados en la aplicación interna",
            "tabColor": "#00ffff",
            "settings": {}
          },
          {
            "name": "Coordinación",
            "code": "pccoordin",
            "description": "Tabla de incidencias en el desarrollo del proyecto",
            "tabColor": "#00ffff",
            "settings": {}
          },
          {
            "name": "Contenido Documental",
            "code": "planifica",
            "description": "Tabla de planificación de proyecto",
            "tabColor": "#00ffff",
            "settings": {}
          },
          {
            "name": "Registros",
            "code": "registros",
            "description": "Actas y otros registros",
            "tabColor": "#00ffff",
            "settings": {}
          },
          {
            "name": "Control Horas",
            "code": "pconhoras",
            "description": "Plantilla para la planificación de horas de proyecto",
            "tabColor": "#00ffff",
            "settings": {}
          },
          {
            "name": "Due Diligence",
            "code": "duedilige",
            "description": "Reporte de diligencias al inicio de proyecto",
            "tabColor": "#ff9900",
            "settings": {}
          },
          {
            "name": "Normativa URB",
            "code": "norurbani",
            "description": "Check automatizado de la normativa urbanística del proyecto",
            "tabColor": "#e69138",
            "settings": {}
          },
          {
            "name": "KPIs",
            "code": "panconkpi",
            "description": "Tabla de principales KPI de proyecto",
            "tabColor": "#ff00ff",
            "settings": {}
          },
          {
            "name": "Costes desglosados",
            "code": "cosdesglo",
            "description": "Tablas de costes de predimensionados y superficies BIM",
            "tabColor": "#ff00ff",
            "settings": {}
          },
          {
            "name": "Costes precios",
            "code": "cosprecio",
            "description": "Hoja maestra de la base de datos de precios €/m2",
            "tabColor": "#ff00ff",
            "settings": {}
          },
          {
            "name": "Premisas cliente",
            "code": "premiscli",
            "description": "Mix de premisas del cliente para unidades y ZZCC",
            "tabColor": "#ff00ff",
            "settings": {}
          },
          {
            "name": "Predim Mix Sup y PEC",
            "code": "predimmix",
            "description": "Tabla de configuración para el predimensionado inicial",
            "tabColor": "#ff00ff",
            "settings": {}
          },
          {
            "name": "Predim Sup VT",
            "code": "predimsup",
            "description": "Tabla de valores para el predimensionado inicial",
            "tabColor": "#ff00ff",
            "settings": {}
          },
          {
            "name": "Predim Costes",
            "code": "predimcos",
            "description": "Tablas de costes de predimensionados y superficies BIM",
            "tabColor": "#ff00ff",
            "settings": {}
          },
          {
            "name": "X BDD Proyectos",
            "code": "xbddproye",
            "description": "",
            "tabColor": "#00ff00",
            "settings": {}
          },
          {
            "name": "X Contactos",
            "code": "xcontacto",
            "description": "",
            "tabColor": "#00ff00",
            "settings": {}
          },
          {
            "name": "X Cuadro Sup",
            "code": "xcuadrsup",
            "description": "",
            "tabColor": "#00ff00",
            "settings": {}
          },
          {
            "name": "X Empresas",
            "code": "xempresas",
            "description": "",
            "tabColor": "#00ff00",
            "settings": {}
          },
          {
            "name": "X Horas tareas",
            "code": "xhorastar",
            "description": "",
            "tabColor": "#00ff00",
            "settings": {}
          },
          {
            "name": "X Matriz superficies útiles",
            "code": "xmasuputi",
            "description": "",
            "tabColor": "#00ff00",
            "settings": {}
          },
          {
            "name": "X Normativa",
            "code": "xnormativ",
            "description": "",
            "tabColor": "#00ff00",
            "settings": {}
          },
          {
            "name": "X PROYECTOS Coordinación",
            "code": "xproyecoo",
            "description": "",
            "tabColor": "#00ff00",
            "settings": {}
          },
          {
            "name": "X Roseta de ciudades",
            "code": "xrosciuda",
            "description": "",
            "tabColor": "#00ff00",
            "settings": {}
          },
          {
            "name": "X ROSETA CIUDADES y mixes",
            "code": "xroscimix",
            "description": "",
            "tabColor": "#00ff00",
            "settings": {}
          },
          {
            "name": "X Roseta edificabilidad y rentas",
            "code": "xrosedren",
            "description": "",
            "tabColor": "#00ff00",
            "settings": {}
          },
          {
            "name": "X Variables",
            "code": "xvariable",
            "description": "",
            "tabColor": "#00ff00",
            "settings": {}
          },
          {
            "name": "X Variables normativa",
            "code": "xvarianor",
            "description": "",
            "tabColor": "#00ff00",
            "settings": {}
          },
          {
            "name": "X Variables Coste",
            "code": "xvariacos",
            "description": "",
            "tabColor": "#00ff00",
            "settings": {}
          },
          {
            "name": "X Variables PT",
            "code": "xvariapyt",
            "description": "",
            "tabColor": "#00ff00",
            "settings": {}
          }
        ],
        "secondarySheets": []
      }

      break;

    case 'Cuadro Superficies':

      try {
        var valoresUnicos = obtenerValoresUnicosDeRangoNombrado('X Variables', '.', 'Desplegables/ Nombre unidades');
      } catch (error) { }

      templateSheetConfigObject = {
        "fileType": [ tipoArchivo ],
        "projectTypes": [
          {
            "nombre": "Plurifamiliar",
            "code": "plurifami",
            "settings": {
              "fullCuadro": ["csupelink", "cschivato", "conedific", "chektipom", "chektipos", "conplanta", "cuadrogar", "cuadrvtft", "sisectors", "siocupaci", "txtoperac", "xexportac", "xvariable", "xvariabsi", "txtsuperf"],
              "minimumCuadro": ["csupelink", "cschivato", "conedific", "chektipom", "chektipos", "conplanta", "txtoperac", "xexportac", "xvariable", "xvariabsi", "txtsuperf"],
              "revitCuadro": ["rvtsisect", "rvtsiocup", "rvtconare", "rvtutiesp", "rvtutihab"],
              "hiddenSheets": ["sisectors", "siocupaci", "aytojusti", "aytoconst", "aytomadrd", "cuadrvtft", "cuadrogar", "histconst", "txtoperac", "xexportac", "xvariable", "xvariabsi", "txtsuperf", "rvtconare", "rvtutiesp", "rvtutihab"]
            }
          },
          {
            "nombre": "Unifamiliar",
            "code": "unifamili",
            "settings": {
              "fullCuadro": ["csupelink", "cschivato", "conedific", "chektipom", "chektipos", "conplanta", "cuadrogar", "cuadrvtft", "sisectors", "siocupaci", "txtoperac", "xexportac", "xvariable", "xvariabsi", "txtsuperf"],
              "minimumCuadro": ["csupelink", "cschivato", "conedific", "chektipom", "chektipos", "conplanta", "txtoperac", "xexportac", "xvariable", "xvariabsi", "txtsuperf"],
              "revitCuadro": ["rvtsisect", "rvtsiocup", "rvtconare", "rvtutiesp", "rvtutihab"],
              "hiddenSheets": ["sisectors", "siocupaci", "aytojusti", "aytoconst", "aytomadrd", "cuadrvtft", "cuadrogar", "histconst", "txtoperac", "xexportac", "xvariable", "xvariabsi", "txtsuperf", "rvtconare", "rvtutiesp", "rvtutihab"]
            }
          }
        ],
        "masterSheets": [
          {
            "name": "LINK",
            "code": "csupelink",
            "description": "Hoja de datos generales y conexión con otros cuadros",
            "tabColor": "#ffff00",
            "settings": {
              "showInList": true,
            }
          },
          {
            "name": "CHIVATOS",
            "code": "cschivato",
            "description": "Hoja resumen de los principales chivatos del cuadro",
            "tabColor": "#ff0000",
            "settings": {
              "showInList": true,
            }
          },
          {
            "name": "Construida_Edificable",
            "code": "conedific",
            "description": "Hoja con las principales superficies construidas y edificables",
            "tabColor": "#3c78d8",
            "settings": {
              "showInList": true,
            }
          },
          {
            "name": "Check TIPO M",
            "code": "chektipom",
            "description": "Hoja de chequeo de las superficies generales por tipología",
            "tabColor": "#3c78d8",
            "settings": {
              "showInList": true,
              "filaFormulas": true,
              "portalSwitch": {
                "range": "BLOQUE;2",
                "action": "setBoolean"
              },
              "deleteColumnGroup": {
                "range": { "range": "A:1", }, // El rango es la primera letra del rango del grupo y la profundidad
                "action": "deleteColumnGroup"
              },
              "addColumnGroup": {
                "range": { "range": "A:D", },
                "action": "addColumnGroup"
              },
              "resizeColumns": {
                "range": { "multiple": "K", },
                "action": "resizeColumns"
              }
            }
          },
          {
            "name": "Check TIPO S",
            "code": "chektipos",
            "description": "Hoja de chequeo de las superficies de estancias por tipología",
            "tabColor": "#3c78d8",
            "settings": {
              "showInList": true,
              "filaFormulas": true,
              "portalSwitch": {
                "range": "BLOQUE;2",
                "action": "setBoolean"
              },
              "deleteColumnGroup": {
                "range": { "range": "A:1", }, // El rango es la primera letra del rango del grupo y la profundidad
                "action": "deleteColumnGroup"
              },
              "addColumnGroup": {
                "range": { "range": "A:D", },
                "action": "addColumnGroup"
              },
              "resizeColumns": {
                "range": { "multiple": "P", },
                "action": "resizeColumns"
              }
            }
          },
          {
            "name": "CONST PLANTAS",
            "code": "conplanta",
            "description": "Resúmenes de las superficies por planta, incluyendo aparcamiento y trasteros",
            "tabColor": "#ffd966",
            "settings": {
              "showInList": true,
              "filaFormulas": true,
              "tipoSuperficie": {
                "range": "PLANTAS;2",
                "action": "setValue"
              },
            }
          },
          {
            "name": "Cuadro garaje",
            "code": "cuadrogar",
            "description": "Cuadro de chequeo de las superficies y plazas de garaje",
            "tabColor": "#6aa84f",
            "settings": {
              "showInList": true,
              "filaFormulas": true,
              "switch": {
                "range": "BLOQUE;2",
                "action": "setValue"
              },
            }
          },
          {
            "name": "Cuadro VT-FT",
            "code": "cuadrvtft",
            "description": "Cuadro justificativo de ventanas y falsos techos",
            "tabColor": "#6aa84f",
            "settings": {
              "showInList": true,
              "filaFormulas": true,
            }
          },
          {
            "name": "SUP AUTO",
            "code": "cssupauto",
            "description": "Desglose de superficies automáticas",
            "tabColor": "#6aa84f",
            "settings": {
              "showInList": true,
            }
          },
          {
            "name": "SI SECTORES Y LRE",
            "code": "sisectors",
            "description": "Cuadro de sectores de incendios y LRE",
            "tabColor": "#cc4125",
            "settings": {
              "showInList": true,
              "filaFormulas": true,
            }
          },
          {
            "name": "SI OCUPACIÓN",
            "code": "siocupaci",
            "description": "Cuadro general de ocupación para Seguridad contra Incendios",
            "tabColor": "#cc4125",
            "settings": {
              "showInList": true,
              "filaFormulas": true,
            }
          },
          {
            "name": "Cuadros justificativos AYTO",
            "code": "aytojusti",
            "description": "Cuadro resumen de justificación de superficies para el ayuntamiento",
            "tabColor": "#ff00ff",
            "settings": {
              "showInList": true,
            }
          },
          {
            "name": "AYTO CONST USO",
            "code": "aytoconst",
            "description": "Cuadro de ayuntamiento de superficies construidas por uso",
            "tabColor": "#ff00ff",
            "settings": {
              "showInList": true,
              "typeSwitch": {
                "range": "A1",
                "action": "setValue"
              }
            }
          },
          {
            "name": "AYTO MADRID",
            "code": "aytomadrd",
            "description": "Cuadro de ayuntamiento de superficies computables y no computables",
            "tabColor": "#ff00ff",
            "settings": {
              "showInList": true,
              "portal": {
                "range": "B1",
                "action": "setValue"
              },
              "srbr": {
                "range": "D1",
                "action": "setValue"
              },
              "portalResume": {
                "range": { "range": "T:Y", },
                "action": "deleteRange"
              }
            }
          },
          {
            "name": "CÓMPUTO UNIDAD",
            "code": "computviv",
            "description": "Resumen de unidades de vivienda por tipología y planta",
            "tabColor": "#ff00ff",
            "settings": {
              "showInList": true,
            }
          },
          {
            "name": "IT 03 CB Cuarto Basuras",
            "code": "aytobasur",
            "description": "Cuadro de ayuntamiento de justificación de los cuartos de basuras",
            "tabColor": "#ff00ff",
            "settings": {
              "showInList": true,
            }
          },
          {
            "name": "Histórico CONST",
            "code": "histconst",
            "description": "Hoja que documenta los cambios de superficies a lo largo del proyecto",
            "tabColor": "#9900ff",
            "settings": {
              "showInList": true,
            }
          },
          {
            "name": "TXT OPERACIONES",
            "code": "txtoperac",
            "description": "Hoja que documenta los cambios de superficies a lo largo del proyecto",
            "tabColor": "#00ff00",
            "settings": {
              "showInList": true,
            }
          },
          {
            "name": "X Exportación",
            "code": "xexportac",
            "description": "Hoja que documenta los cambios de superficies a lo largo del proyecto",
            "tabColor": "#00ff00",
            "settings": {
              "showInList": true,
            }
          },
          {
            "name": "X Variables",
            "code": "xvariable",
            "description": "Hoja que documenta los cambios de superficies a lo largo del proyecto",
            "tabColor": "#00ff00",
            "settings": {
              "showInList": true,
            }
          },
          {
            "name": "X Variables SI",
            "code": "xvariabsi",
            "description": "Hoja que documenta los cambios de superficies a lo largo del proyecto",
            "tabColor": "#00ff00",
            "settings": {
              "showInList": true,
            }
          },
          {
            "name": "TXT SUPERFICIES",
            "code": "txtsuperf",
            "description": "Hoja que documenta los cambios de superficies a lo largo del proyecto",
            "tabColor": "#00ff00",
            "settings": {
              "showInList": true,
            }
          },
          {
            "name": "RVT_SI SECTORES Y LRE",
            "code": "rvtsisect", 
            "description": "Cuadro de sectores de incendios y LRE (Revit)",
            "tabColor": "#cc4125",
            "settings": {
              "showInList": false,
              "filaFormulas": true,
            }
          },
          {
            "name": "RVT_SI OCUPACIÓN",
            "code": "rvtsiocup",
            "description": "Cuadro general de ocupación para Seguridad contra Incendios (Revit)",
            "tabColor": "#cc4125",
            "settings": {
              "showInList": false,
              "filaFormulas": true,
            }
          },
          {
            "name": "RVT_TXT OPERACIONES",
            "code": "rvttxtope",
            "description": "Hoja que documenta los cambios de superficies a lo largo del proyecto (Revit)",
            "tabColor": "#00ff00",
            "settings": {
              "showInList": false
            }
          },
          {
            "name": "RVT_TXT CONST AREAS",
            "code": "rvtconare",
            "description": "",
            "tabColor": "#00ff00",
            "settings": {
              "showInList": false
            }
          },
          {
            "name": "RVT_TXT UTIL ESPACIOS",
            "code": "rvtutiesp",
            "description": "",
            "tabColor": "#00ff00",
            "settings": {
              "showInList": false
            }
          },
          {
            "name": "RVT_TXT UTIL HAB",
            "code": "rvtutihab",
            "description": "",
            "tabColor": "#00ff00",
            "settings": {
              "showInList": false
            }
          },
        ],
        "secondarySheets": [
          {
            "name": "Check Bloque-Letra M",
            "description": "Cuadro de superficies M desglosado por bloques",
            "masterSheet": "Check TIPO M",
            "relativePosition": 1,
            "settings": {
              "filaFormulas": true,
              "portalSwitch": {
                "value": true
              },
              "deleteColumnGroup": {
                "value": true
              },
              "addColumnGroup": {
                "value": true
              },
              "resizeColumns": {
                "value": [50]
              }
            }
          },
          ...(valoresUnicos !== undefined
            ? valoresUnicos.map(function(valor) {
                return {
                  "name": "Check " + valor + " TIPO M",
                  "masterSheet": "Check TIPO M",
                  "relativePosition": 2,
                  "settings": {
                    "filaFormulas": true,
                  }
                };
              })
            : []),
          {
            "name": "Check Bloque-Letra S",
            "description": "Cuadro de superficies S desglosado por bloques",
            "masterSheet": "Check TIPO S",
            "relativePosition": 1,
            "settings": {
              "filaFormulas": true,
              "portalSwitch": {
                "value": true
              },
              "deleteColumnGroup": {
                "value": true
              },
              "addColumnGroup": {
                "value": true
              },
              "resizeColumns": {
                "value": [50]
              }
            }
          },
          ...(valoresUnicos !== undefined
            ? valoresUnicos.map(function(valor) {
                return {
                  "name": "Check " + valor + " TIPO S",
                  "masterSheet": "Check TIPO S",
                  "relativePosition": 2,
                  "settings": {
                    "filaFormulas": true,
                  }
                };
              })
            : []),
          {
            "name": "Aparcamiento_Trastero",
            "description": "Cuadro de superficies de aparcamientos y trasteros",
            "masterSheet": "RESUMEN PLANTAS",
            "relativePosition": 1,
            "settings": {
              "tipoResumen": { "value": 'Aparcamiento_Trastero' },
              "hideConfig": { "value": true }
            }
          },
          {
            "name": "PLANTAS SR",
            "description": "Superficies por planta construidas sobre rasante",
            "masterSheet": "RESUMEN PLANTAS",
            "relativePosition": 2,
            "settings": {
              "tipoResumen": { "value": 'PLANTAS SR' },
              "hideConfig": { "value": true }
            }
          },
          {
            "name": "PLANTAS BR",
            "description": "Superficies por planta bajo rasante",
            "masterSheet": "RESUMEN PLANTAS",
            "relativePosition": 3,
            "settings": {
              "tipoResumen": { "value": 'PLANTAS BR' },
              "hideConfig": { "value": true }
            }
          },
          {
            "name": "PLANTAS URB",
            "description": "Superficies por planta de urbanización",
            "masterSheet": "RESUMEN PLANTAS",
            "relativePosition": 4,
            "settings": {
              "tipoResumen": { "value": 'PLANTAS URB' },
              "hideConfig": { "value": true }
            }
          },
          {
            "name": "ZZCC SR",
            "description": "Superficies por planta de zonas comunes sobre rasante",
            "masterSheet": "RESUMEN PLANTAS",
            "relativePosition": 5,
            "settings": {
              "tipoResumen": { "value": 'ZZCC SR' },
              "hideConfig": { "value": true }
            }
          },
          {
            "name": "COMP PLANTAS",
            "description": "Superficies por planta computables sobre rasante y bajo rasante",
            "masterSheet": "CONST PLANTAS",
            "relativePosition": 1,
            "settings": {
              "tipoSuperficie": { "value": 'COMPUTABLE' },
              "hideConfig": { "value": true }
            }
          },
          {
            "name": "REP PLANTAS",
            "description": "Superficies por planta repercutibles sobre rasante y bajo rasante",
            "masterSheet": "CONST PLANTAS",
            "relativePosition": 2,
            "settings": {
              "tipoSuperficie": { "value": 'REPERCUTIBLE' },
              "hideConfig": { "value": true }
            }
          },
          {
            "name": "RESUMEN ÚTIL VIV TIPO interior",
            "description": "Resumen de superficies interiores de vivienda por tipo",
            "masterSheet": "RESUMEN TIPOS",
            "relativePosition": 1,
            "settings": {
              "tipoResumen": { "value": "RESUMEN ÚTIL VIV TIPO interior" }
            }
          },
          {
            "name": "RESUMEN ÚTIL VIV SUBTIPO",
            "description": "Superficies de superficies interiores y exteriores de vivienda por tipo y subtipo",
            "masterSheet": "RESUMEN TIPOS",
            "relativePosition": 2,
            "settings": {
              "tipoResumen": { "value": "RESUMEN ÚTIL VIV SUBTIPO" }
            }
          },
          {
            "name": "RESUMEN BLOQUE-LETRA M",
            "description": "Superficies de superficies interiores y exteriores de vivienda por tipo y subtipo",
            "masterSheet": "RESUMEN TIPOS",
            "relativePosition": 2,
            "settings": {
              "tipoResumen": { "value": "RESUMEN BLOQUE-LETRA M" }
            }
          },
          {
            "name": "Cuadro trasteros",
            "description": "Cuadro de superficies de trasteros",
            "masterSheet": "Cuadro garaje",
            "relativePosition": 1,
            "settings": {
              "switch": {
                "value": "Trasteros",
              }
            }
          },
          {
            "name": "AYTO COMP USO",
            "description": "Cuadro de ayuntamiento de superficies computables por uso",
            "masterSheet": "AYTO CONST USO",
            "relativePosition": 1,
            "settings": {
              "typeSwitch": {
                "value": "COMPUTABLE POR USOS",
              }
            }
          },
          {
            "name": "AYTO MADRID SR",
            "description": "Cuadro de ayuntamiento de superficies computables y no computables sobre rasante",
            "masterSheet": "AYTO MADRID",
            "relativePosition": 1,
            "settings": {
              "portal": { "value": "TODO", },
              "srbr": { "value": "SR", },
              "portalResume": { "value": true, }
            }
          },
          {
            "name": "AYTO MADRID BR",
            "description": "Cuadro de ayuntamiento de superficies computables y no computables bajo rasante",
            "masterSheet": "AYTO MADRID",
            "relativePosition": 2,
            "settings": {
              "portal": { "value": "TODO", },
              "srbr": { "value": "BR", },
              "portalResume": { "value": true, }
            }
          },
        ]
      }

      break;

    case 'Exportación Superficies':

      templateSheetConfigObject = {
        "fileType": [ tipoArchivo ],
        "projectTypes": [
          {
            "nombre": "Plurifamiliar",
            "code": "plurifami",
            "settings": {
              "minimumCuadro": ["cexpolink", "resconfig", "resvivtip", "resvivpla"],
              "fullCuadro": ["cexpolink", "resconfig", "resvivtip", "resvivpla"],
              "revitCuadro": [],
              "hiddenSheets": ["resconfig", "resvivtip", "resvivpla"]
            }
          },
          {
            "nombre": "Unifamiliar",
            "code": "unifamili",
            "settings": {
              "minimumCuadro": ["cexpolink", "resconfig", "resvivtip", "resvivpla"],
              "fullCuadro": ["cexpolink", "resconfig", "resvivtip", "resvivpla"],
              "revitCuadro": [],
              "hiddenSheets": ["resconfig", "resvivtip", "resvivpla"]
            }
          }
        ],
        "masterSheets": [
          {
            "name": "LINK",
            "code": "cexpolink",
            "description": "Hoja de datos generales y conexión con otros cuadros",
            "tabColor": "#ffff00",
            "settings": {
            }
          },
          {
            "name": "CONFIG",
            "code": "resconfig",
            "description": "Hoja de configuración para las plantillas",
            "tabColor": "#00ff00",
            "settings": {
            }
          },
          {
            "name": "PLANTILLA TIPOS",
            "code": "resvivtip",
            "description": "Hoja de resúmenes de superficies por tipos",
            "tabColor": "#ff00ff",
            "settings": {
              "tipoResumen": {
                "range": "B2",
                "action": "setValue"
              }
            }
          },
          {
            "name": "PLANTILLA PLANTAS",
            "code": "resvivpla",
            "description": "Hoja de resúmenes de superficies por planta",
            "tabColor": "#ffd966",
            "settings": {
              "tipoResumen": {
                "range": "B2",
                "action": "setValue"
              }
            }
          },
        ],
        "secondarySheets": [
          {
            "name": "RESUMEN VIV TIPO",
            "description": "Resumen de superficies interiores de vivienda por tipo",
            "masterSheet": "PLANTILLA TIPOS",
            "relativePosition": 1,
            "settings": {
              "tipoResumen": { "value": "RESUMEN VIV TIPO" }
            }
          },
          {
            "name": "RESUMEN BLOQUE-LETRA M",
            "description": "Resumen de superficies interiores de vivienda por bloque-letra",
            "masterSheet": "PLANTILLA TIPOS",
            "relativePosition": 1,
            "settings": {
              "tipoResumen": { "value": "RESUMEN BLOQUE-LETRA M" }
            }
          },
          {
            "name": "RESUMEN ÚTIL VIV SUBTIPO",
            "description": "Resumen de superficies útiles interiores y exteriores de vivienda por tipo y subtipo",
            "masterSheet": "PLANTILLA TIPOS",
            "relativePosition": 2,
            "settings": {
              "tipoResumen": { "value": "ÚTIL VIV SUBTIPO" }
            }
          },
          {
            "name": "RESUMEN BLOQUE-LETRA ESTANCIAS",
            "description": "Resumen de superficies interiores y exteriores de vivienda por tipo y subtipo",
            "masterSheet": "PLANTILLA TIPOS",
            "relativePosition": 2,
            "settings": {
              "tipoResumen": { "value": "RESUMEN BLOQUE-LETRA ESTANCIAS" }
            }
          },
          {
            "name": "Aparcamiento_Trastero",
            "description": "Cuadro de superficies de aparcamientos y trasteros",
            "masterSheet": "PLANTILLA PLANTAS",
            "relativePosition": 1,
            "settings": {
              "tipoResumen": { "value": 'Aparcamiento_Trastero' },
            }
          },
          {
            "name": "PLANTAS SR",
            "description": "Superficies por planta construidas sobre rasante",
            "masterSheet": "PLANTILLA PLANTAS",
            "relativePosition": 2,
            "settings": {
              "tipoResumen": { "value": 'PLANTAS SR' },
            }
          },
          {
            "name": "PLANTAS BR",
            "description": "Superficies por planta bajo rasante",
            "masterSheet": "PLANTILLA PLANTAS",
            "relativePosition": 3,
            "settings": {
              "tipoResumen": { "value": 'PLANTAS BR' },
            }
          },
          {
            "name": "PLANTAS URB",
            "description": "Superficies por planta de urbanización",
            "masterSheet": "PLANTILLA PLANTAS",
            "relativePosition": 4,
            "settings": {
              "tipoResumen": { "value": 'PLANTAS URB' },
            }
          },
          {
            "name": "ZZCC SR",
            "description": "Superficies por planta de zonas comunes sobre rasante",
            "masterSheet": "PLANTILLA PLANTAS",
            "relativePosition": 5,
            "settings": {
              "tipoResumen": { "value": 'ZZCC SR' },
            }
          },
        ]
      }

      break;

    case 'Cuadro Mediciones':

      templateSheetConfigObject = {
        "fileType": [ tipoArchivo ],
        "projectTypes": [
          {
            "nombre": "Plurifamiliar",
            "code": "plurifami",
            "settings": {
              "minimumCuadro": ["cexpolink", "cuamedvar", "bctrespre", "materibim", "calprebim", "calpreglo", "calpreuni", "materibim", "materibim"],
              "fullCuadro": ["cexpolink", "resconfig", "resvivtip", "resvivpla"],
              "revitCuadro": [],
              "hiddenSheets": ["resconfig", "resvivtip", "resvivpla"]
            }
          },
          {
            "nombre": "Unifamiliar",
            "code": "unifamili",
            "settings": {
              "minimumCuadro": ["cexpolink", "resconfig", "resvivtip", "resvivpla"],
              "fullCuadro": ["cexpolink", "resconfig", "resvivtip", "resvivpla"],
              "revitCuadro": [],
              "hiddenSheets": ["resconfig", "resvivtip", "resvivpla"]
            }
          }
        ],
        "masterSheets": [
          {
            "name": "LINK",
            "code": "cexpolink",
            "description": "Hoja de datos generales y conexión con otros cuadros",
            "tabColor": "#ffff00",
            "settings": {
            }
          },
          {
            "name": "DATA",
            "code": "cuamedvar",
            "description": "Hoja de configuración, datos y variables del cuadro de mediciones",
            "tabColor": "#00ff00",
            "settings": {
            }
          },
          {
            "name": "BC3",
            "code": "bctrespre",
            "description": "Plantilla BC3 del presupuesto para importar en Presto",
            "tabColor": "#ff00ff",
            "settings": {
            }
          },
          {
            "name": "Materiales BIM",
            "code": "materibim",
            "description": "Listado de materiales que vienen del modelo BIM",
            "tabColor": "#ff9900",
            "settings": {
            }
          },
          {
            "name": "Calidades y presupuesto CALCULADOS",
            "code": "calprebim",
            "description": "Hoja de control de los datos extraídos en BIM",
            "tabColor": "#1155cc",
            "settings": {
            }
          },
          {
            "name": "Calidades y presupuesto GLOBAL",
            "code": "calpreglo",
            "description": "Hoja de control general para mediciones BIM",
            "tabColor": "#1155cc",
            "settings": {
            }
          },
          {
            "name": "Calidades y presupuesto UNICOS",
            "code": "calpreuni",
            "description": "Hoja de control general para mediciones BIM",
            "tabColor": "#1155cc",
            "settings": {
            }
          },
          {
            "name": "X BDD Variables",
            "code": "xbddvaria",
            "description": "Hoja de conexión con la base de datos",
            "tabColor": "#00ff00",
            "settings": { }
          },
          {
            "name": "X BDD Componentes",
            "code": "xbddcompo",
            "description": "Hoja de conexión con la base de datos",
            "tabColor": "#00ff00",
            "settings": { }
          },
          {
            "name": "X BDD Compuestos",
            "code": "xbddcompu",
            "description": "Hoja de conexión con la base de datos",
            "tabColor": "#00ff00",
            "settings": { }
          },
          {
            "name": "X CYPE_EST",
            "code": "xcypeestr",
            "description": "Hoja de conexión con la base de datos",
            "tabColor": "#00ff00",
            "settings": { }
          },
        ],
        "secondarySheets": []
      }

      break;
    
    case 'none':

      templateSheetConfigObject = {
        "fileType": ["none"],
        "projectTypes": [
          {
            "nombre": "Plurifamiliar",
            "code": "plurifami",
            "settings": {
              "minimumCuadro": [],
              "fullCuadro": [],
              "revitCuadro": [],
              "hiddenSheets": []
            }
          },
          {
            "nombre": "Unifamiliar",
            "code": "unifamili",
            "settings": {
              "excludedMasterSheets": [],
              "hiddenSheets": [],
              "excludedOtherSheets": []
            }
          }
        ],
        "masterSheets": [],
        "secondarySheets": []
      }

      break;
  }

  return templateSheetConfigObject;
}
