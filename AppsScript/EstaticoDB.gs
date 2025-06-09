/*---------------------------------------------------------- *** DECLARACIÓN DE VARIABLES GLOBALES *** ------------------------------------------- */

const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

const columnaFamilia = 4; 
const columnaGenero = 5 ;
const columnaMEXU = 12 ;
const columnaLocalidad = 27; 
const columnaMunicipio = 28;  
const columnaEstado = 29; 
const columnaPaises = 30; 
const columnaDias = 32;       
const columnaMeses = 33;   


const colorCorrecto = '#B4D3B2'; // Verde
const colorIncorrecto = '#FF0000'; // Rojo

const ultimaFila = hoja.getLastRow() - 1;

/* ------------------------------------------------------------ %%% FUNCIÓN PRINCIPAL %%% ------------------------------------------------------ */

/**
 * Función principal que llama a las demás para crear hacer las verificaciones.
 * Utiliza un bloque try-catch para mejorar la robustez y añade alertas que indican lo sucedido.
 */
function main() {
  try {
    SpreadsheetApp.getUi().alert("Se inician las validaciones, por favor espere al siguiente mensaje.");
    
    validarFamilias();
    validarPaises();
    validarLocalidadMunicipio();
    validarFechas();
    validarDatosMEXU()

    SpreadsheetApp.getUi().alert("Validación completa de Familias, Países, Localidades, Municipios, Fechas y Números MEXU. Revisa las celdas resaltadas.");
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Error en la validación: ${error.message}`);
  }
}

/*------------------------------------------------------------- !!! VERIFICACION DE FAMILIAS  !!! ----------------------------------*/

/**
 * Función para verificar las familias válidas respecto a una lista que ya está verificada.
 */
function validarFamilias (){
  const datos = hoja.getRange(2, columnaFamilia, ultimaFila).getValues();

//IGUALMENTE FALTAN VERIFICAR BIEN LAS FAMILIAS, TENGO ESPECIES Y GENEROS DENTRO DE ESTA LISTA.
  const familiasValidas = [
    "Acanthaceae", "Achariaceae", "Achatocarpaceae", "Acoraceae", "Acorales", "Actinidiaceae", "Adoxaceae", "Aextoxicaceae", "Aizoaceae", "Akaniaceae", "Alismataceae", "Alismatales", "Alseuosmiaceae", "Alstroemeriaceae", "Altingiaceae", "Alzateaceae", "Amaranthaceae", "Amaryllidaceae", "Amborellaceae", "Amborellales", "Amphorogynaceae", "Anacampserotaceae", "Anacardiaceae", "Anarthriaceae", "Ancistrocladaceae", "Angiosperms*********", "Anisophylleaceae", "Annonaceae", "Aphanopetalaceae", "Aphloiaceae", "Apiaceae", "Apiales", "Apocynaceae", "Apodanthaceae", "Aponogetonaceae", "Aquifoliaceae", "Aquifoliales", "Araceae", "Araliaceae", "Arecaceae", "Arecales", "Argophyllaceae", "Aristolochiaceae", "Asparagaceae", "Asparagales", "Asphodelaceae", "Asteliaceae", "Asteraceae", "Asterales", "asterids", "Asteropeiaceae", "Atherospermataceae", "Austrobaileyaceae", "Austrobaileyales", "Balanopaceae", "Balanophoraceae", "Balsaminaceae", "Barbeuiaceae", "Barbeyaceae", "Basellaceae", "Bataceae", "Begoniaceae", "Berberidaceae", "Berberidopsidaceae", "Berberidopsidales", "Bersamaceae", "Betulaceae", "Biebersteiniaceae", "Bignoniaceae", "Bixaceae", "Blandfordiaceae", "Bonnetiaceae", "Boraginaceae", "Boraginales", "Borthwickiaceae", "Boryaceae", "Brassicaceae", "Brassicales", "Bromeliaceae", "Brunelliaceae", "Bruniaceae", "Bruniales", "Burmanniaceae", "Burseraceae", "Butomaceae", "Buxaceae", "Buxales", "Byblidaceae", "Cabombaceae", "Cactaceae", "Calceolariaceae", "Calophyllaceae", "Calycanthaceae", "Calyceraceae", "Campanulaceae", "campanulids", "Campynemataceae", "Canellaceae", "Canellales", "Cannabaceae", "Cannaceae", "Capparaceae", "Caprifoliaceae", "Cardiopteridaceae", "Caricaceae", "Carlemanniaceae", "Caryocaraceae", "Caryophyllaceae", "Caryophyllales", "Casuarinaceae", "Celastraceae", "Celastrales", "Centrolepidaceae", "Centroplacaceae", "Cephalotaceae", "Ceratophyllaceae", "Ceratophyllales", "Cercidiphyllaceae", "Cervantesiaceae", "Chloranthaceae*****", "Chloranthales********", "Chrysobalanaceae", "Circaeasteraceae", "Cistaceae", "Cleomaceae", "Clethraceae", "Clusiaceae", "Codonaceae", "Colchicaceae", "Columelliaceae", "Comandraceae", "Combretaceae", "Commelinaceae", "Commelinales", "Compositae", "Connaraceae", "Convolvulaceae", "core eudicots***************", "Coriariaceae", "Cornaceae", "Cornales", "Corsiaceae", "Corynocarpaceae", "Costaceae", "Crassulaceae", "Crossosomataceae", "Crossosomatales", "Cruciferae", "Crypteroniaceae", "Ctenolophonaceae", "Cucurbitaceae", "Cucurbitales", "Cunoniaceae", "Curtisiaceae", "Cyclanthaceae", "Cymodoceaceae", "Cynomoriaceae", "Cyperaceae", "Cyrillaceae", "Cytinaceae", "Daphniphyllaceae", "Dasypogonaceae", "Datiscaceae", "Degeneriaceae", "Diapensiaceae", "Dichapetalaceae", "Didiereaceae", "Dilleniaceae", "Dilleniales", "Dioncophyllaceae", "Dioscoreaceae", "Dioscoreales", "Dipentodontaceae", "Dipsacales", "Dipterocarpaceae", "Dirachmaceae", "Doryanthaceae", "Droseraceae", "Drosophyllaceae", "Ebenaceae", "Ecdeiocoleaceae", "Elaeagnaceae", "Elaeocarpaceae", "Elatinaceae", "Emblingiaceae", "Ericaceae", "Ericales", "Eriocaulaceae", "Erythroxylaceae", "Escalloniaceae", "Escalloniales", "Eucommiaceae", "Eudicots", "Euphorbiaceae", "Euphroniaceae", "Eupomatiaceae", "Eupteleaceae", "Fabaceae", "Fabales", "Fabids***********", "Fagaceae", "Fagales", "Flagellariaceae", "Fouquieriaceae", "Francoaceae", "Frankeniaceae", "Garryaceae", "Garryales", "Geissolomataceae", "Gelsemiaceae", "Gentianaceae", "Gentianales", "Geraniaceae", "Geraniales", "Gerrardinaceae", "Gesneriaceae", "Gisekiaceae", "Gomortegaceae", "Goodeniaceae", "Goupiaceae", "Gramineae", "Greyiaceae", "Griseliniaceae", "Grossulariaceae", "Grubbiaceae", "Guamatelaceae", "Gunneraceae", "Gunnerales", "Guttiferae", "Gyrostemonaceae", "Haemodoraceae", "Halophytaceae", "Haloragaceae", "Hamamelidaceae", "Hanguanaceae", "Haptanthaceae", "Heliconiaceae", "Helwingiaceae", "Hernandiaceae", "Himantandraceae", "Huaceae", "Huerteales", "Humiriaceae", "Hydatellaceae", "Hydnoraceae", "Hydrangeaceae", "Hydrocharitaceae", "Hydroleaceae", "Hydrostachyaceae", "Hypericaceae", "Hypoxidaceae", "Icacinaceae", "Icacinales", "Iridaceae", "Irvingiaceae", "Iteaceae", "Ixioliriaceae", "Ixonanthaceae", "Joinvilleaceae", "Juglandaceae", "Juncaceae", "Juncaginaceae", "Kewaceae", "Kirkiaceae", "Koeberliniaceae", "Krameriaceae", "Labiatae", "Lacistemataceae", "Lactoridaceae", "Lamiaceae", "Lamiales", "lamiids", "Lanariaceae", "Lardizabalaceae", "Lauraceae", "Laurales", "Lecythidaceae", "Ledocarpaceae", "Leguminosae", "Lentibulariaceae", "Lepidobotryaceae", "Liliaceae", "Liliales", "Limeaceae", "Limnanthaceae", "Linaceae", "Lindenbergiaceae", "Linderniaceae", "Loasaceae", "Loganiaceae", "Lophiocarpaceae", "Lophopyxidaceae", "Loranthaceae", "Lowiaceae", "Lythraceae", "Macarthuriaceae", "Magnoliaceae", "Magnoliales", "magnoliids", "Malpighiaceae", "Malpighiales", "Malvaceae", "Malvales", "malvids", "Marantaceae", "Marcgraviaceae", "Martyniaceae", "Maundiaceae", "Mayacaceae", "Mazaceae", "Melanthiaceae", "Melastomataceae", "Meliaceae", "Melianthaceae", "Menispermaceae", "Menyanthaceae", "Metteniusaceae", "Metteniusales", "Microteaceae", "Misodendraceae", "Mitrastemonaceae", "Molluginaceae", "Monimiaceae", "monocots", "Montiaceae", "Montiniaceae", "Moraceae", "Moringaceae", "Muntingiaceae", "Musaceae", "Myodocarpaceae", "Myricaceae", "Myristicaceae", "Myrothamnaceae", "Myrtaceae", "Myrtales", "Nanodeaceae", "Nartheciaceae", "Nelumbonaceae", "Nepenthaceae", "Neuradaceae", "Nitrariaceae", "Nothofagaceae", "Nyctaginaceae", "Nymphaeaceae", "Nymphaeales", "Nyssaceae", "Ochnaceae", "Olacaceae", "Oleaceae", "Onagraceae", "Oncothecaceae", "Opiliaceae", "Orchidaceae", "Orobanchaceae", "Oxalidaceae", "Oxalidales", "Paeoniaceae", "Palmae", "Pandaceae", "Pandanaceae", "Pandanales", "Papaveraceae", "Paracryphiaceae", "Paracryphiales", "Passifloraceae", "Paulowniaceae", "Pedaliaceae", "Penaeaceae", "Pennantiaceae", "Pentadiplandraceae", "Pentaphragmataceae", "Pentaphylacaceae", "Penthoraceae", "Peraceae", "Peridiscaceae", "Petenaeaceae", "Petermanniaceae", "Petrosaviaceae", "Petrosaviales", "Phellinaceae", "Philesiaceae", "Philydraceae", "Phrymaceae", "Phyllanthaceae", "Phyllonomaceae", "Physenaceae", "Phytolaccaceae", "Picramniaceae", "Picramniales", "Picrodendraceae", "Piperaceae", "Piperales", "Pittosporaceae", "Plantae", "Plantaginaceae", "Platanaceae", "Plocospermataceae", "Plumbaginaceae", "Poaceae", "Poales", "Podostemaceae", "Polemoniaceae", "Polygalaceae", "Polygonaceae", "Pontederiaceae", "Portulacaceae", "Posidoniaceae", "Potamogetonaceae", "Primulaceae", "Proteaceae", "Proteales", "Pteleocarpaceae", "Putranjivaceae", "Quillajaceae", "Rafflesiaceae", "Ranunculaceae", "Ranunculales", "Rapateaceae", "Resedaceae", "Restionaceae", "Rhabdodendraceae", "Rhamnaceae", "Rhizophoraceae", "Rhynchothecaceae", "Ripogonaceae", "Rivinaceae", "Roridulaceae", "Rosaceae", "Rosales", "rosids", "Rousseaceae", "Rubiaceae", "Ruppiaceae", "Rutaceae", "Sabiaceae", "Salicaceae", "Salvadoraceae", "Santalaceae", "Santalales", "Sapindaceae", "Sapindales", "Sapotaceae", "Sarcobataceae", "Sarcolaenaceae", "Sarraceniaceae", "Saururaceae", "Saxifragaceae", "Saxifragales", "Scheuchzeriaceae", "Schisandraceae", "Schlegeliaceae", "Schoepfiaceae", "scientificName", "Scrophulariaceae", "Setchellanthaceae", "Simaroubaceae", "Simmondsiaceae", "Siparunaceae", "Sladeniaceae", "Smilacaceae", "Solanaceae", "Solanales", "Sphaerosepalaceae", "Sphenocleaceae", "Stachyuraceae", "Staphyleaceae", "Stegnospermataceae", "Stemonaceae", "Stemonuraceae", "Stilbaceae", "Stixidaceae", "Strasburgeriaceae", "Strelitziaceae", "Stylidiaceae", "Styracaceae", "superasterids", "superrosids", "Surianaceae", "Symplocaceae", "Talinaceae", "Tamaricaceae", "Tapisciaceae", "Tecophilaeaceae", "Tetracarpaeaceae", "Tetrachondraceae", "Tetramelaceae", "Tetrameristaceae", "Theaceae", "Thomandersiaceae", "Thurniaceae", "Thymelaeaceae", "Ticodendraceae", "Tofieldiaceae", "Torricelliaceae", "Tovariaceae", "Trigoniaceae", "Trimeniaceae", "Triuridaceae", "Trochodendraceae", "Trochodendrales", "Tropaeolaceae", "Typhaceae", "Ulmaceae", "Umbelliferae", "Urticaceae", "Vahliaceae", "Vahliales", "Velloziaceae", "Verbenaceae", "Violaceae", "Vitaceae", "Vitales", "Vivianiaceae", "Vochysiaceae", "Winteraceae", "Xanthocerataceae", "Xanthorrhoeaceae", "Xeronemataceae", "Xyridaceae", "Zingiberaceae", "Zingiberales", "Zosteraceae", "Zygophyllaceae", "Zygophyllales"
  ];

//Ciclo para 
datos.forEach((fila, i) => {
    let familia = String(fila[0]); // Asegurar que es String y eliminar espacios
    const familiaNormalizada = familia.charAt(0).toUpperCase() + familia.slice(1).toLowerCase();
    
    let casillaFamilia = hoja.getRange(i + 2, columnaFamilia) ;

    if (!familia){
      limpiarCelda(casillaFamilia) ; 
      return ;
    } // Si la celda está vacía, saltar la fila

    if (familiasValidas.includes(familiaNormalizada)) {
      // Si la familia es válida, pintar en verde y eliminar comentario si existe
      casillaFamilia.setBackground(colorCorrecto).setComment(null);
    } else {
      // Si la familia no es válida, pintar en rojo y sugerir opciones
      casillaFamilia.setBackground(colorIncorrecto);

      // Generar sugerencias automáticas
      const sugerencias = familiasValidas.filter(f => f.toLowerCase().startsWith(familia.slice(0, 3).toLowerCase()));
      const sugerenciaTexto = sugerencias.length > 0 ? sugerencias.join(", ") : "No se encontró sugerencia";

      // Agregar comentario con la sugerencia
      casillaFamilia.setComment(`Sugerencias: ${sugerenciaTexto}`);
    }

  });

}

/*----------------------------------------------------- ### VERIFICACIÓN DE LOS PAÍSES BIEN ESCRITOS ### -------------------------------------*/

/**
 * Función para verificar los países válidos respecto a una lista de países.
 */
function validarPaises() {
  const rango = hoja.getRange(2, columnaPaises, hoja.getLastRow() - 1); 
  const datos = rango.getValues();

  const paisesValidos = [
   "África", "América Central", "Argentina", "Australia", "Austria",
    "Bahamas", "Belice", "Bolivia", "Brasil", "Canadá", "Ceylon", "Checoslovaquia", "Colombia", "Costa Rica", "Cuba",
    "Ecuador", "España", "Estados Unidos Americanos", "Europa", "Filipinas", "Francia", "Guatemala", "Guiana",
    "Honduras", "India", "Indonesia", "Inglaterra", "Israel", "Italia", "Jamaica", "Japón", "Malasia", "México", "NA",
    "Nicaragua", "Nueva Zelanda", "País", "Países Bajos", "Paraguay", "Perú", "Puerto Rico",
    "República Cooperativa de Guyana", "República de Filipinas", "República de Surinam",
    "República Democrática del Congo", "República Democrática Socialista de Sri Lanka", "Santa Lucía", "Sudáfrica",
    "Suiza", "Trinidad", "Venezuela", "Zaire"
  ];

  // Validar los datos
  datos.forEach((fila, i) => {
    let pais = String(fila[0]); // Asegurar que es String y eliminar espacios

    // Normalizar formato
    const paisNormalizado = pais.charAt(0).toUpperCase() + pais.slice(1).toLowerCase();

    let casillaPais = hoja.getRange(i + 2, columnaPaises) ; 

    if(!pais){
      limpiarCelda(casillaPais) ;
      return ; 
    }

    if (paisesValidos.includes(paisNormalizado)) {
      // Si el país es válido, pintar en verde y eliminar comentario si existe
      casillaPais.setBackground("green").setComment(null);
    } else {
      // Si el país no es válido, pintar en rojo y sugerir opciones

      // Generar sugerencias automáticas
      const sugerencias = paisesValidos.filter(p => p.toLowerCase().startsWith(pais.slice(0, 3).toLowerCase()));
      const sugerenciaTexto = sugerencias.length > 0 ? sugerencias.join(", ") : "No se encontró sugerencia";

      // Agregar comentario con la sugerencia
      casillaPais.setBackground("red").setComment(`Sugerencias: ${sugerenciaTexto}`);
    }
    
  });

}

/*------------------------------------------ ^^^ VERIFICACION DE MÉXICO Y SUS LOCALIDADES Y MUNICIPIOS ^^^ ----------------------------------*/

/**
 * Función para verificar la localidad y municipio para que no sean ningún estado del país.
 */
function validarLocalidadMunicipio() {

  const estadosMexico = [
    "Aguascalientes", "Baja California", "Baja California Sur", "Campeche", "Chiapas", "Chihuahua",
    "Ciudad de México", "Coahuila", "Colima", "Durango", "Estado de México", "Guanajuato",
    "Guerrero", "Hidalgo", "Jalisco", "Michoacán", "Morelos", "Nayarit", "Nuevo León",
    "Oaxaca", "Puebla", "Querétaro", "Quintana Roo", "San Luis Potosí", "Sinaloa",
    "Sonora", "Tabasco", "Tamaulipas", "Tlaxcala", "Veracruz", "Yucatán", "Zacatecas"
  ];

  const ultimaFila = hoja.getLastRow() - 1;

  const rangoPais = hoja.getRange(2, columnaPaises, ultimaFila);
  const rangoLocalidad = hoja.getRange(2, columnaLocalidad, ultimaFila);
  const rangoMunicipio = hoja.getRange(2, columnaMunicipio, ultimaFila);
  const rangoEstado = hoja.getRange(2, columnaEstado, ultimaFila);

  const valoresPais = rangoPais.getValues();
  const valoresLocalidad = rangoLocalidad.getValues();
  const valoresMunicipio = rangoMunicipio.getValues();
  const valoresEstado = rangoEstado.getValues();


  for (let i = 0; i < ultimaFila; i++) {
    const pais = String(valoresPais[i][0]).trim();
    const localidad = String(valoresLocalidad[i][0]).trim();
    const municipio = String(valoresMunicipio[i][0]).trim();

    //const estado = String(valoresEstado[i][0]).trim();
    // se añadiria !estadosMexico.includes(estado) si es que si hay un estado deje tener el mismo nombre

    const celdaLocalidad = hoja.getRange(i + 2, columnaLocalidad);
    const celdaMunicipio = hoja.getRange(i + 2, columnaMunicipio);

    if (pais === "México") {
      const localidadEsEstado = estadosMexico.includes(localidad);
      const municipioEsEstado = estadosMexico.includes(municipio);

      celdaLocalidad.setBackground(localidadEsEstado ? colorIncorrecto : null) ;
      celdaMunicipio.setBackground(municipioEsEstado ? colorIncorrecto : null) ;

    } else {
      // Restaurar colores si el país no es México
      celdaLocalidad.setBackground(null);
      celdaMunicipio.setBackground(null);
    }
  }
}

/*----------------------------------------------------- @@@ VERIFICACIÓN DE FECHAS @@@ -------------------------------------*/

/**
 * Función para verificar las fechas que pueden ser posibles respecto al mes. No se toman en cuentan los años bisiestos.
 */
function validarFechas() {

  const rangoDias = hoja.getRange(2, columnaDias, ultimaFila);
  const rangoMeses = hoja.getRange(2, columnaMeses, ultimaFila);

  const valoresDias = rangoDias.getValues();
  const valoresMeses = rangoMeses.getValues();

  const colorIncorrecto = "#FFCCCC"; // Color de fondo para errores

  const diasEnMes = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

  for (let i = 0; i < ultimaFila; i++) {
    let dias = valoresDias[i][0];
    let meses = valoresMeses[i][0];

    let celdaDias = hoja.getRange(i + 2, columnaDias);
    let celdaMeses = hoja.getRange(i + 2, columnaMeses);

    let diasValidos = true;
    let mesesValidos = true;

    // Validar día básico
    if (dias < 1 || dias > 31) {
      celdaDias.setBackground(colorIncorrecto).setComment("ES UN NÚMERO DE DÍA INVÁLIDO");
      diasValidos = false;
    } else {
      limpiarCelda(celdaDias) ;
    }

    // Validar mes
    if (meses < 1 || meses > 12) {
      celdaMeses.setBackground(colorIncorrecto).setComment("ES UN NÚMERO DE MES INVÁLIDO");
      mesesValidos = false;
    } else {
      limpiarCelda(celdaMeses) ;
    }

    // Validar combinación días/mes si ambos son válidos
    if (diasValidos && mesesValidos) {
      if (dias > diasEnMes[meses - 1]) {
        celdaDias.setBackground(colorIncorrecto).setComment("Ese número de días no puede estar en ese mes.");
      } else {
        limpiarCelda(celdaMeses) ;
        limpiarCelda(celdaDias) ;
      }
      
    }

    // Limpiar si celda de días está vacía
    if (dias === "" || dias === null || dias === "NA") {
      limpiarCelda(celdaDias) ;
      
    }

    // Limpiar si celda de meses está vacía
    if (meses === "" || meses === null || meses === "NA") {
      limpiarCelda(celdaMeses) ;
    }
  }
}

/*----------------------------------------------------------- ___ VERIFICACION EN BASE A MEXU ___ -----------------------------------------------*/

/**
 * Función para verificar que ni la familia ni el género estén vacíos si ya hay algún número MEXUw.
 */
function validarDatosMEXU() {

  const rangoMEXU = hoja.getRange(2, columnaMEXU, ultimaFila);
  const rangoGenero = hoja.getRange(2, columnaGenero, ultimaFila);
  const rangoFamilia = hoja.getRange(2, columnaFamilia, ultimaFila);

  const valoresMEXU = rangoMEXU.getValues();
  const valoresGenero = rangoGenero.getValues();
  const valoresFamilia = rangoFamilia.getValues();

  for (let i = 0; i < ultimaFila; i++) {
    const valorMEXU = String(valoresMEXU[i][0]).trim();
    const valorGenero = String(valoresGenero[i][0]).trim();
    const valorFamilia = String(valoresFamilia[i][0]).trim();

    const celdaMEXU = hoja.getRange(i + 2, columnaMEXU);
    const celdaGenero = hoja.getRange(i + 2, columnaGenero);
    const celdaFamilia = hoja.getRange(i + 2, columnaFamilia);


    // Validar contenido de MEXU
    if (valorMEXU.includes("MEXUw")) {
      validarCeldaObligatoria(valorFamilia, celdaFamilia, "Error: Esta celda no puede estar vacía si MEXUw tiene valor.");
      validarCeldaObligatoria(valorGenero, celdaGenero, "Error: Esta celda no puede estar vacía si MEXUw tiene valor.");

      if (valorGenero !== "") {
        limpiarCelda(celdaGenero);
      }

    } else {
      celdaMEXU.setBackground(colorIncorrecto).setComment("Esta celda no puede ser vacía");
    }
  }
}

// Función para limpiar una celda
function limpiarCelda(celda) {
  celda.setBackground(null).setComment(null);
}

// Función para validar si una celda debe tener un valor
function validarCeldaObligatoria(valor, celda, mensajeError) {
  if (valor === "") {
    celda.setBackground(colorIncorrecto).setComment(mensajeError);
  }
}
