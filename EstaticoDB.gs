
function main() {
    // Llamada a la función de validación de países
  validarFamilias(); // Llamada a la función de validación de familias
  validarPaises();
  SpreadsheetApp.getUi().alert("Validación completa de Familia y Países. Revisa las celdas resaltadas.");
}

function validarFamilias (){
  
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var columnaFamilia = 4; // Corregido: Usar columnaFamilia en lugar de columnaPais
  const rango = hoja.getRange(2, columnaFamilia, hoja.getLastRow() - 1); // Rango correcto
  const datos = rango.getValues();

  // Definir colores
  var colorCorrecto = '#B4D3B2'; // Verde
  var colorIncorrecto = '#FF0000'; // Rojo

//IGUALMENTE FALTAN VERIFICAR BIEN LAS FAMILIAS, TENGO ESPECIES Y GENEROS DENTRO DE ESTA LISTA.
  const familiasValidas = [
    "Acanthaceae", "Achariaceae", "Achatocarpaceae", "Acoraceae", "Acorales", "Actinidiaceae", "Adoxaceae", "Aextoxicaceae", "Aizoaceae", "Akaniaceae", "Alismataceae", "Alismatales", "Alseuosmiaceae", "Alstroemeriaceae", "Altingiaceae", "Alzateaceae", "Amaranthaceae", "Amaryllidaceae", "Amborellaceae", "Amborellales", "Amphorogynaceae", "Anacampserotaceae", "Anacardiaceae", "Anarthriaceae", "Ancistrocladaceae", "Angiosperms*********", "Anisophylleaceae", "Annonaceae", "Aphanopetalaceae", "Aphloiaceae", "Apiaceae", "Apiales", "Apocynaceae", "Apodanthaceae", "Aponogetonaceae", "Aquifoliaceae", "Aquifoliales", "Araceae", "Araliaceae", "Arecaceae", "Arecales", "Argophyllaceae", "Aristolochiaceae", "Asparagaceae", "Asparagales", "Asphodelaceae", "Asteliaceae", "Asteraceae", "Asterales", "asterids", "Asteropeiaceae", "Atherospermataceae", "Austrobaileyaceae", "Austrobaileyales", "Balanopaceae", "Balanophoraceae", "Balsaminaceae", "Barbeuiaceae", "Barbeyaceae", "Basellaceae", "Bataceae", "Begoniaceae", "Berberidaceae", "Berberidopsidaceae", "Berberidopsidales", "Bersamaceae", "Betulaceae", "Biebersteiniaceae", "Bignoniaceae", "Bixaceae", "Blandfordiaceae", "Bonnetiaceae", "Boraginaceae", "Boraginales", "Borthwickiaceae", "Boryaceae", "Brassicaceae", "Brassicales", "Bromeliaceae", "Brunelliaceae", "Bruniaceae", "Bruniales", "Burmanniaceae", "Burseraceae", "Butomaceae", "Buxaceae", "Buxales", "Byblidaceae", "Cabombaceae", "Cactaceae", "Calceolariaceae", "Calophyllaceae", "Calycanthaceae", "Calyceraceae", "Campanulaceae", "campanulids", "Campynemataceae", "Canellaceae", "Canellales", "Cannabaceae", "Cannaceae", "Capparaceae", "Caprifoliaceae", "Cardiopteridaceae", "Caricaceae", "Carlemanniaceae", "Caryocaraceae", "Caryophyllaceae", "Caryophyllales", "Casuarinaceae", "Celastraceae", "Celastrales", "Centrolepidaceae", "Centroplacaceae", "Cephalotaceae", "Ceratophyllaceae", "Ceratophyllales", "Cercidiphyllaceae", "Cervantesiaceae", "Chloranthaceae*****", "Chloranthales********", "Chrysobalanaceae", "Circaeasteraceae", "Cistaceae", "Cleomaceae", "Clethraceae", "Clusiaceae", "Codonaceae", "Colchicaceae", "Columelliaceae", "Comandraceae", "Combretaceae", "Commelinaceae", "Commelinales", "Compositae", "Connaraceae", "Convolvulaceae", "core eudicots***************", "Coriariaceae", "Cornaceae", "Cornales", "Corsiaceae", "Corynocarpaceae", "Costaceae", "Crassulaceae", "Crossosomataceae", "Crossosomatales", "Cruciferae", "Crypteroniaceae", "Ctenolophonaceae", "Cucurbitaceae", "Cucurbitales", "Cunoniaceae", "Curtisiaceae", "Cyclanthaceae", "Cymodoceaceae", "Cynomoriaceae", "Cyperaceae", "Cyrillaceae", "Cytinaceae", "Daphniphyllaceae", "Dasypogonaceae", "Datiscaceae", "Degeneriaceae", "Diapensiaceae", "Dichapetalaceae", "Didiereaceae", "Dilleniaceae", "Dilleniales", "Dioncophyllaceae", "Dioscoreaceae", "Dioscoreales", "Dipentodontaceae", "Dipsacales", "Dipterocarpaceae", "Dirachmaceae", "Doryanthaceae", "Droseraceae", "Drosophyllaceae", "Ebenaceae", "Ecdeiocoleaceae", "Elaeagnaceae", "Elaeocarpaceae", "Elatinaceae", "Emblingiaceae", "Ericaceae", "Ericales", "Eriocaulaceae", "Erythroxylaceae", "Escalloniaceae", "Escalloniales", "Eucommiaceae", "Eudicots", "Euphorbiaceae", "Euphroniaceae", "Eupomatiaceae", "Eupteleaceae", "Fabaceae", "Fabales", "Fabids***********", "Fagaceae", "Fagales", "Flagellariaceae", "Fouquieriaceae", "Francoaceae", "Frankeniaceae", "Garryaceae", "Garryales", "Geissolomataceae", "Gelsemiaceae", "Gentianaceae", "Gentianales", "Geraniaceae", "Geraniales", "Gerrardinaceae", "Gesneriaceae", "Gisekiaceae", "Gomortegaceae", "Goodeniaceae", "Goupiaceae", "Gramineae", "Greyiaceae", "Griseliniaceae", "Grossulariaceae", "Grubbiaceae", "Guamatelaceae", "Gunneraceae", "Gunnerales", "Guttiferae", "Gyrostemonaceae", "Haemodoraceae", "Halophytaceae", "Haloragaceae", "Hamamelidaceae", "Hanguanaceae", "Haptanthaceae", "Heliconiaceae", "Helwingiaceae", "Hernandiaceae", "Himantandraceae", "Huaceae", "Huerteales", "Humiriaceae", "Hydatellaceae", "Hydnoraceae", "Hydrangeaceae", "Hydrocharitaceae", "Hydroleaceae", "Hydrostachyaceae", "Hypericaceae", "Hypoxidaceae", "Icacinaceae", "Icacinales", "Iridaceae", "Irvingiaceae", "Iteaceae", "Ixioliriaceae", "Ixonanthaceae", "Joinvilleaceae", "Juglandaceae", "Juncaceae", "Juncaginaceae", "Kewaceae", "Kirkiaceae", "Koeberliniaceae", "Krameriaceae", "Labiatae", "Lacistemataceae", "Lactoridaceae", "Lamiaceae", "Lamiales", "lamiids", "Lanariaceae", "Lardizabalaceae", "Lauraceae", "Laurales", "Lecythidaceae", "Ledocarpaceae", "Leguminosae", "Lentibulariaceae", "Lepidobotryaceae", "Liliaceae", "Liliales", "Limeaceae", "Limnanthaceae", "Linaceae", "Lindenbergiaceae", "Linderniaceae", "Loasaceae", "Loganiaceae", "Lophiocarpaceae", "Lophopyxidaceae", "Loranthaceae", "Lowiaceae", "Lythraceae", "Macarthuriaceae", "Magnoliaceae", "Magnoliales", "magnoliids", "Malpighiaceae", "Malpighiales", "Malvaceae", "Malvales", "malvids", "Marantaceae", "Marcgraviaceae", "Martyniaceae", "Maundiaceae", "Mayacaceae", "Mazaceae", "Melanthiaceae", "Melastomataceae", "Meliaceae", "Melianthaceae", "Menispermaceae", "Menyanthaceae", "Metteniusaceae", "Metteniusales", "Microteaceae", "Misodendraceae", "Mitrastemonaceae", "Molluginaceae", "Monimiaceae", "monocots", "Montiaceae", "Montiniaceae", "Moraceae", "Moringaceae", "Muntingiaceae", "Musaceae", "Myodocarpaceae", "Myricaceae", "Myristicaceae", "Myrothamnaceae", "Myrtaceae", "Myrtales", "Nanodeaceae", "Nartheciaceae", "Nelumbonaceae", "Nepenthaceae", "Neuradaceae", "Nitrariaceae", "Nothofagaceae", "Nyctaginaceae", "Nymphaeaceae", "Nymphaeales", "Nyssaceae", "Ochnaceae", "Olacaceae", "Oleaceae", "Onagraceae", "Oncothecaceae", "Opiliaceae", "Orchidaceae", "Orobanchaceae", "Oxalidaceae", "Oxalidales", "Paeoniaceae", "Palmae", "Pandaceae", "Pandanaceae", "Pandanales", "Papaveraceae", "Paracryphiaceae", "Paracryphiales", "Passifloraceae", "Paulowniaceae", "Pedaliaceae", "Penaeaceae", "Pennantiaceae", "Pentadiplandraceae", "Pentaphragmataceae", "Pentaphylacaceae", "Penthoraceae", "Peraceae", "Peridiscaceae", "Petenaeaceae", "Petermanniaceae", "Petrosaviaceae", "Petrosaviales", "Phellinaceae", "Philesiaceae", "Philydraceae", "Phrymaceae", "Phyllanthaceae", "Phyllonomaceae", "Physenaceae", "Phytolaccaceae", "Picramniaceae", "Picramniales", "Picrodendraceae", "Piperaceae", "Piperales", "Pittosporaceae", "Plantae", "Plantaginaceae", "Platanaceae", "Plocospermataceae", "Plumbaginaceae", "Poaceae", "Poales", "Podostemaceae", "Polemoniaceae", "Polygalaceae", "Polygonaceae", "Pontederiaceae", "Portulacaceae", "Posidoniaceae", "Potamogetonaceae", "Primulaceae", "Proteaceae", "Proteales", "Pteleocarpaceae", "Putranjivaceae", "Quillajaceae", "Rafflesiaceae", "Ranunculaceae", "Ranunculales", "Rapateaceae", "Resedaceae", "Restionaceae", "Rhabdodendraceae", "Rhamnaceae", "Rhizophoraceae", "Rhynchothecaceae", "Ripogonaceae", "Rivinaceae", "Roridulaceae", "Rosaceae", "Rosales", "rosids", "Rousseaceae", "Rubiaceae", "Ruppiaceae", "Rutaceae", "Sabiaceae", "Salicaceae", "Salvadoraceae", "Santalaceae", "Santalales", "Sapindaceae", "Sapindales", "Sapotaceae", "Sarcobataceae", "Sarcolaenaceae", "Sarraceniaceae", "Saururaceae", "Saxifragaceae", "Saxifragales", "Scheuchzeriaceae", "Schisandraceae", "Schlegeliaceae", "Schoepfiaceae", "scientificName", "Scrophulariaceae", "Setchellanthaceae", "Simaroubaceae", "Simmondsiaceae", "Siparunaceae", "Sladeniaceae", "Smilacaceae", "Solanaceae", "Solanales", "Sphaerosepalaceae", "Sphenocleaceae", "Stachyuraceae", "Staphyleaceae", "Stegnospermataceae", "Stemonaceae", "Stemonuraceae", "Stilbaceae", "Stixidaceae", "Strasburgeriaceae", "Strelitziaceae", "Stylidiaceae", "Styracaceae", "superasterids", "superrosids", "Surianaceae", "Symplocaceae", "Talinaceae", "Tamaricaceae", "Tapisciaceae", "Tecophilaeaceae", "Tetracarpaeaceae", "Tetrachondraceae", "Tetramelaceae", "Tetrameristaceae", "Theaceae", "Thomandersiaceae", "Thurniaceae", "Thymelaeaceae", "Ticodendraceae", "Tofieldiaceae", "Torricelliaceae", "Tovariaceae", "Trigoniaceae", "Trimeniaceae", "Triuridaceae", "Trochodendraceae", "Trochodendrales", "Tropaeolaceae", "Typhaceae", "Ulmaceae", "Umbelliferae", "Urticaceae", "Vahliaceae", "Vahliales", "Velloziaceae", "Verbenaceae", "Violaceae", "Vitaceae", "Vitales", "Vivianiaceae", "Vochysiaceae", "Winteraceae", "Xanthocerataceae", "Xanthorrhoeaceae", "Xeronemataceae", "Xyridaceae", "Zingiberaceae", "Zingiberales", "Zosteraceae", "Zygophyllaceae", "Zygophyllales"
  ];

//Ciclo para 
datos.forEach((fila, i) => {
    let familia = String(fila[0]).trim(); // Asegurar que es String y eliminar espacios

    // Normalizar formato
    const familiaNormalizada = familia.charAt(0).toUpperCase() + familia.slice(1).toLowerCase();

    let casillaFamilia = hoja.getRange(i + 2, columnaFamilia) ;

    if (!familia){
      casillaFamilia.setBackground(null).setComment(null);
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


function validarPaises() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const columnaPaises = 30; 
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
    let pais = String(fila[0]).trim(); // Asegurar que es String y eliminar espacios

    // Normalizar formato
    const paisNormalizado = pais.charAt(0).toUpperCase() + pais.slice(1).toLowerCase();

    let casillaPais = hoja.getRange(i + 2, columnaPaises) ; 

    if(!pais){
      casillaPais.setBackground(null).setComment(null);
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

