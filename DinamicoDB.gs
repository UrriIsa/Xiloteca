
/**
 * Función que se ejecuta cuando se edita una celda en la hoja de cálculo.
 * Verifica la validez de la familia ingresada y cambia el color de la celda.
 * @param {Event} e - Evento de edición de Google Sheets.
 */
function onEdit(e) {

 //------------------------------------------------------------------- ||| GENERALES |||  -------------------------------------------------

  if (!e) {
    Logger.log('El evento e no se ha pasado correctamente.') ;
    return ;
  }

  
  var hoja = e.source.getActiveSheet() ;
  var rango = e.range ;
  var fila = rango.getRow() ;
  var columnaEditada = rango.getColumn() ;

  /**
   * Definición de las columnas clave en la hoja de cálculo. Si borras columnas debes verificar el número.
   */
  var columnaPais = 30 ; 
  var columnaMunicipio = 28 ;
  var columnaLocalidad = 27 ;
  var columnaFamilia = 4 ;
  var columnaMEXU = 12 ;
  var columnaGenero = 5 ;

  /**
   * Código de color en formato hexadecimal.
   * /
  var colorCorrecto = '#B4D3B2' ; // Verde
  var colorIncorrecto = '#FF0000' ; // Rojo


  //------------------------------------------------------- !!! VERIFICACION DE FAMILIAS  !!! ----------------------------------

  /**
   * Lista de familias de plantas válidas para la verificación.
   */
  var familiasValidas = [
    "Acanthaceae", "Achariaceae", "Achatocarpaceae", "Acoraceae", "Actinidiaceae", "Aextoxicaceae", "Aizoaceae", "Akaniaceae", "Alismataceae", "Alseuosmiaceae", "Alstroemeriaceae", "Altingiaceae", "Alzateaceae", "Amaranthaceae", "Amaryllidaceae", "Amborellaceae", "Amphorogynaceae", "Anacampserotaceae", "Anacardiaceae", "Anarthriaceae", "Ancistrocladaceae", "Anisophylleaceae", "Annonaceae", "Aphanopetalaceae", "Aphloiaceae", "Apiaceae","Apocynaceae", "Apodanthaceae", "Aponogetonaceae", "Aquifoliaceae", "Araceae", "Araliaceae", "Araucariaceae", "Arecaceae", "Argophyllaceae", "Aristolochiaceae", "Asparagaceae", "Asphodelaceae", "Asteliaceae", "Asteraceae", "Asteropeiaceae", "Atherospermataceae", "Austrobaileyaceae","Balanopaceae", "Balanophoraceae", "Balsaminaceae", "Barbeuiaceae", "Barbeyaceae", "Basellaceae", "Bataceae", "Begoniaceae", "Berberidaceae", "Berberidopsidaceae", "Bersamaceae", "Betulaceae", "Biebersteiniaceae", "Bignoniaceae", "Bixaceae", "Blandfordiaceae", "Bonnetiaceae", "Boraginaceae", "Borthwickiaceae", "Boryaceae", "Brassicaceae", "Bromeliaceae", "Brunelliaceae", "Bruniaceae", "Burmanniaceae", "Burseraceae", "Butomaceae", "Buxaceae", "Byblidaceae", "Cabombaceae", "Cactaceae", "Calceolariaceae", "Calophyllaceae", "Calycanthaceae", "Calyceraceae", "Campanulaceae","Campynemataceae", "Canellaceae", "Cannabaceae", "Cannaceae", "Capparaceae", "Caprifoliaceae", "Cardiopteridaceae", "Caricaceae", "Carlemanniaceae", "Caryocaraceae", "Caryophyllaceae", "Casuarinaceae", "Celastraceae", "Centrolepidaceae", "Centroplacaceae", "Cephalotaceae", "Ceratophyllaceae", "Cercidiphyllaceae", "Cervantesiaceae", "Chloranthaceae", "Chrysobalanaceae", "Circaeasteraceae", "Cistaceae", "Cleomaceae", "Clethraceae", "Clusiaceae", "Codonaceae", "Colchicaceae", "Columelliaceae", "Comandraceae", "Combretaceae", "Commelinaceae","Connaraceae", "Convolvulaceae", "Coriariaceae", "Cornaceae", "Corsiaceae", "Corynocarpaceae", "Costaceae", "Crassulaceae", "Crossosomataceae", "Crypteroniaceae", "Ctenolophonaceae", "Cucurbitaceae", "Cunoniaceae", "Cupressaceae" ,"Curtisiaceae", "Cyclanthaceae", "Cymodoceaceae", "Cynomoriaceae", "Cyperaceae", "Cyrillaceae", "Cytinaceae", "Daphniphyllaceae", "Dasypogonaceae", "Datiscaceae", "Degeneriaceae", "Diapensiaceae", "Dichapetalaceae", "Didiereaceae", "Dilleniaceae", "Dioncophyllaceae", "Dioscoreaceae", "Dipentodontaceae", "Dipterocarpaceae", "Dirachmaceae", "Doryanthaceae", "Droseraceae", "Drosophyllaceae", "Ebenaceae", "Ecdeiocoleaceae", "Elaeagnaceae", "Elaeocarpaceae", "Elatinaceae", "Emblingiaceae", "Ericaceae", "Eriocaulaceae", "Erythroxylaceae", "Escalloniaceae", "Eucommiaceae", "Euphorbiaceae", "Euphroniaceae", "Eupomatiaceae", "Eupteleaceae", "Fabaceae", "Fagaceae", "Flagellariaceae", "Fouquieriaceae", "Francoaceae", "Frankeniaceae", "Garryaceae", "Geissolomataceae", "Gelsemiaceae", "Gentianaceae","Geraniaceae","Gerrardinaceae", "Gesneriaceae", "Gisekiaceae", "Gomortegaceae", "Goodeniaceae", "Goupiaceae", "Gramineae", "Greyiaceae", "Griseliniaceae", "Grossulariaceae", "Grubbiaceae", "Guamatelaceae", "Gunneraceae", "Gyrostemonaceae", "Haemodoraceae", "Halophytaceae", "Haloragaceae", "Hamamelidaceae", "Hanguanaceae", "Haptanthaceae", "Heliconiaceae", "Helwingiaceae", "Hernandiaceae", "Himantandraceae", "Huaceae", "Humiriaceae", "Hydatellaceae", "Hydnoraceae", "Hydrangeaceae", "Hydrocharitaceae", "Hydroleaceae", "Hydrostachyaceae", "Hypericaceae", "Hypoxidaceae", "Icacinaceae","Iridaceae", "Irvingiaceae", "Iteaceae", "Ixioliriaceae", "Ixonanthaceae", "Joinvilleaceae", "Juglandaceae", "Juncaceae", "Juncaginaceae", "Kewaceae", "Kirkiaceae", "Koeberliniaceae", "Krameriaceae",  "Lacistemataceae", "Lactoridaceae", "Lamiaceae", "Lanariaceae", "Lardizabalaceae", "Lauraceae", "Lecythidaceae", "Ledocarpaceae", "Leguminosae", "Lentibulariaceae", "Lepidobotryaceae", "Liliaceae", "Limeaceae", "Limnanthaceae", "Linaceae", "Lindenbergiaceae", "Linderniaceae", "Loasaceae", "Loganiaceae", "Lophiocarpaceae", "Lophopyxidaceae", "Loranthaceae", "Lowiaceae", "Lythraceae", "Macarthuriaceae", "Magnoliaceae", "Malpighiaceae", "Malvaceae", "Marantaceae", "Marcgraviaceae", "Martyniaceae", "Maundiaceae", "Mayacaceae", "Mazaceae", "Melanthiaceae", "Melastomataceae", "Meliaceae", "Melianthaceae", "Menispermaceae", "Menyanthaceae", "Metteniusaceae", "Microteaceae", "Misodendraceae", "Mitrastemonaceae", "Molluginaceae", "Monimiaceae", "Montiaceae", "Montiniaceae", "Moraceae", "Moringaceae", "Muntingiaceae", "Musaceae", "Myodocarpaceae", "Myricaceae", "Myristicaceae", "Myrothamnaceae", "Myrtaceae", "Nanodeaceae", "Nartheciaceae", "Nelumbonaceae", "Nepenthaceae", "Neuradaceae", "Nitrariaceae", "Nothofagaceae", "Nyctaginaceae", "Nymphaeaceae", "Nyssaceae", "Ochnaceae", "Olacaceae", "Oleaceae", "Onagraceae", "Oncothecaceae", "Opiliaceae", "Orchidaceae", "Orobanchaceae", "Oxalidaceae", "Paeoniaceae", "Palmae", "Pandaceae", "Pandanaceae",  "Papaveraceae", "Paracryphiaceae", "Passifloraceae", "Paulowniaceae", "Pedaliaceae", "Penaeaceae", "Pennantiaceae", "Pentadiplandraceae", "Pentaphragmataceae", "Pentaphylacaceae", "Penthoraceae", "Peraceae", "Peridiscaceae", "Petenaeaceae", "Petermanniaceae", "Petrosaviaceae", "Phellinaceae", "Philesiaceae", "Philydraceae", "Phrymaceae", "Phyllanthaceae", "Phyllonomaceae", "Physenaceae", "Phytolaccaceae", "Picramniaceae", "Picrodendraceae", "Piperaceae", "Pittosporaceae", "Plantaginaceae", "Platanaceae", "Plocospermataceae", "Plumbaginaceae", "Poaceae", "Podostemaceae", "Polemoniaceae", "Polygalaceae", "Polygonaceae", "Pontederiaceae", "Portulacaceae", "Posidoniaceae", "Potamogetonaceae", "Primulaceae", "Proteaceae", "Pteleocarpaceae", "Putranjivaceae", "Quillajaceae", "Rafflesiaceae", "Ranunculaceae",  "Rapateaceae", "Resedaceae", "Restionaceae", "Rhabdodendraceae", "Rhamnaceae", "Rhizophoraceae", "Rhynchothecaceae", "Ripogonaceae", "Rivinaceae", "Roridulaceae", "Rosaceae",  "Rousseaceae", "Rubiaceae", "Ruppiaceae", "Rutaceae", "Sabiaceae", "Salicaceae", "Salvadoraceae", "Santalaceae", "Sapindaceae", "Sapotaceae", "Sarcobataceae", "Sarcolaenaceae", "Sarraceniaceae", "Saururaceae", "Saxifragaceae", "Scheuchzeriaceae", "Schisandraceae", "Schlegeliaceae", "Schoepfiaceae", "scientificName", "Scrophulariaceae", "Setchellanthaceae", "Simaroubaceae", "Simmondsiaceae", "Siparunaceae", "Sladeniaceae", "Smilacaceae", "Solanaceae", "Sphaerosepalaceae", "Sphenocleaceae", "Stachyuraceae", "Staphyleaceae", "Stegnospermataceae", "Stemonaceae", "Stemonuraceae", "Stilbaceae", "Stixidaceae", "Strasburgeriaceae", "Strelitziaceae", "Stylidiaceae", "Styracaceae", "Surianaceae", "Symplocaceae", "Talinaceae", "Tamaricaceae", "Tapisciaceae", "Tecophilaeaceae", "Tetracarpaeaceae", "Tetrachondraceae", "Tetramelaceae", "Tetrameristaceae", "Theaceae", "Thomandersiaceae", "Thurniaceae", "Thymelaeaceae", "Ticodendraceae", "Tofieldiaceae", "Torricelliaceae", "Tovariaceae", "Trigoniaceae", "Trimeniaceae", "Triuridaceae", "Trochodendraceae",  "Tropaeolaceae", "Typhaceae", "Ulmaceae", "Urticaceae", "Vahliaceae", "Velloziaceae", "Verbenaceae", "Viburnaceae", "Violaceae", "Vitaceae", "Vivianiaceae", "Vochysiaceae", "Winteraceae", "Xanthocerataceae", "Xanthorrhoeaceae", "Xeronemataceae", "Xyridaceae", "Zingiberaceae", "Zosteraceae", "Zygophyllaceae"
  ] ;

  /**
   * Verifica si la celda editada pertenece a la columna de familia y valida el contenido.
   */
  var celdaFamilia = hoja.getRange(fila, columnaFamilia) ;

  if (columnaEditada === columnaFamilia) {
    var valorFamilia = rango.getValue().trim() ;
    var famVerdad = familiasValidas.includes(valorFamilia) ;

    if(!famVerdad){
      celdaFamilia.setBackground(colorIncorrecto) ; 
      celdaFamilia.setComment("Error : esta familia no está en la lista válida.") ; 
    }else{
      celdaFamilia.setBackground(colorCorrecto) ; 
      celdaFamilia.setComment(null) ;
    }
    if(valorFamilia === ""){
      celdaFamilia.setBackground(null) ;
      rango.setComment(null) ;
    }
  }


  //------------------------------------------ VERIFICACION EN BASE A MEXU ----------------------------------}

  var celdaMEXU = hoja.getRange(fila, columnaMEXU) ;
  var celdaGenero = hoja.getRange(fila, columnaGenero) ;

  var valorMEXU = celdaMEXU.getValue().trim() ;
  var valorGenero = celdaGenero.getValue().trim() ;

  if (columnaEditada === columnaFamilia || columnaEditada === columnaMEXU || columnaEditada === columnaGenero) {
      if(valorMEXU === ""){
        celdaMEXU.setBackground(null).setComment(null) ; 
        celdaGenero.setBackground(null).setComment(null) ;
        celdaFamilia.setBackground(null).setComment(null) ;
      }

      if (valorMEXU.includes("MEXUw")) {
          // Validar que Familia y Género no estén vacíos
          if (valorFamilia === "" || valorGenero === "") {
              celdaFamilia.setBackground(colorIncorrecto) ;
              celdaGenero.setBackground(colorIncorrecto) ;
              celdaFamilia.setComment("Error: Esta celda no puede estar vacía si MEXUw tiene valor.") ;
              celdaGenero.setComment("Error: Esta celda no puede estar vacía si MEXUw tiene valor.") ;
          } 
          if(valorGenero != ""){
              celdaGenero.setBackground(null) ;
              celdaGenero.setComment(null) ;
          }
          // Validar que Familia esté en la lista válida
          if (!familiasValidas.includes(valorFamilia)) {
              celdaFamilia.setBackground(colorIncorrecto) ;
              celdaFamilia.setComment("Error: Esta familia no está en la lista válida.") ;
          }
      }else if(valorMEXU != ""){
        celdaMEXU.setBackground(colorIncorrecto) ;
        celda.setComment("Checa lo que escribiste.") ; 
      }

  }


  //----------------------------------------------------- ### VERIFICACIÓN DE LOS PAÍSES BIEN ESCRITOS ### -------------------------------------

  /**
   * Lista de países válidos para la verificación.
   */
  var paisesValidos = [
    "América Central", "Argentina", "Australia", "Austria",
    "Bahamas", "Belice", "Bolivia", "Brasil", "Canadá", "Ceylon", "Checoslovaquia", "Colombia", "Costa Rica", "Costa de Marfil", "Cuba",
    "Ecuador", "España", "Estados Unidos Americanos", "Filipinas", "Francia", "Gabón", "Guatemala", "Guiana", "Guinea",
    "Honduras", "India", "Indonesia", "Inglaterra", "Israel", "Italia", "Jamaica", "Japón", "Malasia", "México", "NA",
    "Nicaragua", "Nigeria", "Nueva Zelanda", "Países Bajos", "Paraguay", "Perú", "Puerto Rico",
    "República Cooperativa de Guyana", "República de Filipinas", "República de Surinam",
    "República Democrática del Congo", "República Democrática Socialista de Sri Lanka", "Santa Lucía", "Sudáfrica",
    "Suiza", "Trinidad", "Venezuela", "Zaire"
  ] ;
  
  /**
   * Valida la entrada de datos en la columna de país y colorea la celda según su validez.
   * Además, gestiona comentarios de error y limpia el formato cuando la celda está vacía.
   */
  if (rango.getColumn() == columnaPais) {
    var valorPais = rango.getValue() ;


    // Verificar si el país está en la lista de países válidos
    if (paisesValidos.includes(valorPais)) {
      // Poner la celda en verde y quitar cualquier comentario
      rango.setBackground(colorCorrecto) ;
      rango.setComment('') ;
    } else {
      // Poner la celda en rojo y agregar un comentario
      rango.setBackground(colorIncorrecto) ;
      rango.setComment('País no válido. Verifique.') ;
    }

     // Poner la celda en blanco si no hay nada escrito
    if (valorPais === '') {
        rango.setBackground(null) ; // Esto dejará la celda sin color de fondo
        rango.setComment('') ; // Eliminar cualquier comentario
    }

  }

  //------------------------------------------ VERIFICACION DE MÉXICO Y SUS LOCALIDADES Y MUNICIPIOS ----------------------------------

  /**
   * Valida que las localidades y municipios sean correctos si el país es "México".
   * 
   */
  var estadosMexico = [
    "Aguascalientes", "Baja California", "Baja California Sur", "Campeche", "Coahuila", "Colima", "Chiapas", "Chihuahua", 
    "Ciudad de México", "Durango", "Guanajuato", "Guerrero", "Hidalgo", "Jalisco", "México", "Michoacán", "Morelos", 
    "Nayarit", "Nuevo León", "Oaxaca", "Puebla", "Querétaro", "Quintana Roo", "San Luis Potosí", "Sinaloa", "Sonora", 
    "Tabasco", "Tamaulipas", "Tlaxcala", "Veracruz", "Yucatán", "Zacatecas"
  ] ;

  // Obtener referencias a las celdas de la fila actual
  var celdaPais = hoja.getRange(fila, columnaPais) ;
  var celdaLocalidad = hoja.getRange(fila, columnaLocalidad) ;
  var celdaMunicipio = hoja.getRange(fila, columnaMunicipio) ;

  var valorLocalidad = celdaLocalidad.getValue().trim() ;
  var valorMunicipio = celdaMunicipio.getValue().trim() ;
  var valorPais = celdaPais.getValue().trim() ;

  // Solo aplicar la validación si se edita País, Localidad o Municipio
  if (columnaEditada === columnaPais || columnaEditada === columnaLocalidad || columnaEditada === columnaMunicipio) {
    // Si el país es "México", verificar Localidad y Municipio
    if (valorPais === "México") {
      var errorLocalidad = estadosMexico.includes(valorLocalidad) ;
      var errorMunicipio = estadosMexico.includes(valorMunicipio) ;

      if (errorLocalidad) {
        celdaLocalidad.setBackground(colorIncorrecto) ; // Rojo
      } else {
        celdaLocalidad.setBackground(null) ; // Restaurar color
      }

      if (errorMunicipio) {
        celdaMunicipio.setBackground(colorIncorrecto) ; // Rojo
      } else {
        celdaMunicipio.setBackground(null) ; // Restaurar color
      }
    } else {
      // Si el país NO es México, restaurar colores
      celdaLocalidad.setBackground(null) ;
      celdaMunicipio.setBackground(null) ;
    }
  }







//VERIFICACION DE LOS AUTORES.
//VERIFICACION EN NEGATIVO DE GENEROS.

