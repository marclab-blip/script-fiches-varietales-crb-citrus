// ======================================================
// SCRIPT DE GÉNÉRATION DES FICHES VARIÉTALES SIMPLES
// ======================================================
// Objectif général :
// Ce script génère automatiquement des fiches variétales Excel
// à partir :
// - d'un fichier Excel contenant les données des variétés ;
// - d'un modèle de fiche Excel ;
// - d'un dossier contenant les photographies ;
// - de marqueurs PNG utilisés pour les frises visuelles.
//
// Fonctionnement général :
// 1. Le script lit le tableau de données.
// 2. Il ouvre une copie du modèle de fiche.
// 3. Il remplace les balises {{nom_colonne}} par les données Excel.
// 4. Il insère les photos correspondant à la variété.
// 5. Il place les marqueurs sur les frises.
// 6. Il enregistre une fiche Excel par variété.
//
// Attention :
// La conversion HEIC -> PNG n'est pas faite par ce script.
// Elle doit être lancée manuellement avec le script prévu pour cela.
// ======================================================

// ======================================================
// 1. CHARGEMENT DES MODULES
// ======================================================
// Ces lignes chargent les bibliothèques nécessaires :
// - exceljs : lecture et modification des fichiers Excel ;
// - fs / fs.promises : accès aux fichiers et dossiers ;
// - path : construction de chemins compatibles Windows ;
// - image-size : lecture des dimensions des images avant insertion.
// À modifier uniquement si les dépendances du projet changent.
// ======================================================

const ExcelJS = require('exceljs');
const fs = require('fs');
const fsp = require('fs/promises');
const path = require('path');

let sizeOf = require('image-size');
sizeOf = sizeOf.imageSize || sizeOf;

const BASE_DIR = path.resolve(__dirname, '..');

// ======================================================
// 2. CONFIGURATION GÉNÉRALE DU SCRIPT
// ======================================================
// Cette zone est la partie à regarder en premier si quelqu'un reprend le script.
// Elle regroupe les chemins des fichiers, les dossiers utilisés, les images
// de marqueurs, les frises et les emplacements des photographies.
//
// À modifier si :
// - le fichier de données change de nom ;
// - le nom de la feuille Excel change ;
// - le modèle de fiche est déplacé ou remplacé ;
// - les dossiers d'entrée/sortie changent ;
// - les marqueurs ou les frises sont modifiés.
// ======================================================

const CONFIG = {
  // ------------------------------------------------------
  // 2.1 Chemins des fichiers et dossiers principaux
  // ------------------------------------------------------
  // sourceExcelPath : fichier Excel contenant les données à importer.
  // sourceSheetName : nom de la feuille du fichier source à lire.
  // templateExcelPath : modèle Excel de fiche à remplir automatiquement.
  // outputDir : dossier où seront enregistrées les fiches générées.
  // photosDir : dossier contenant les photos déjà converties en JPG/PNG/WebP.
  // marker...ImagePath : images PNG utilisées comme marqueurs sur les frises.
  // ------------------------------------------------------
  sourceExcelPath: path.join(BASE_DIR, 'donnees_varietes.xlsx'),
  sourceSheetName: 'Feuil_Bilan',
  templateExcelPath: path.join(BASE_DIR, 'modeles', 'modele_fiche_simple.xlsx'),
  outputDir: path.join(BASE_DIR, 'sortie_fiches', 'simple'),
  photosDir: path.join(BASE_DIR, 'photos'),
markerRougeImagePath: path.join(BASE_DIR, 'modeles', 'marqueur_rouge.png'),
markerJauneImagePath: path.join(BASE_DIR, 'modeles', 'marqueur_jaune.png'),
markerVertImagePath: path.join(BASE_DIR, 'modeles', 'marqueur_vert.png'),

// ------------------------------------------------------
// 2.2 Paramétrage des frises visuelles
// ------------------------------------------------------
// Chaque bloc correspond à une frise présente dans le modèle Excel.
//
// Signification des paramètres :
// - label : nom interne de la frise, utile pour les messages de contrôle ;
// - sourceKey : nom exact de la colonne Excel utilisée pour placer le marqueur ;
// - range : plage de cellules correspondant à la frise dans le modèle ;
// - min / max : bornes de l'échelle visuelle ;
// - markerWidthPx / markerHeightPx : dimensions du marqueur ;
// - offsetXPx / offsetYPx : petits décalages manuels pour ajuster la position.
//
// À modifier si :
// - une frise change de place dans le modèle ;
// - les valeurs minimales ou maximales de l'échelle changent ;
// - un marqueur apparaît légèrement décalé dans la fiche finale.
// ------------------------------------------------------

scaleSlots: [
  {
    label: 'taux_sucre',
    sourceKey: 'sucre_indice',
    range: 'B22:F22',
    min: 6,
    max: 14,
    markerWidthPx: 6,
    markerHeightPx: 34,
    offsetXPx: 18,
    offsetYPx: 0,
  },
  {
    label: 'taux_jus',
    sourceKey: 'rendement_jus_pct',
    range: 'B24:F24',
    min: 20,
    max: 100,
    markerWidthPx: 6,
    markerHeightPx: 34,
    offsetXPx: 10,
    offsetYPx: 0,
  },
  {
    label: 'acidite',
    sourceKey: 'acidite',
    range: 'B26:F26',
    min: 0,
    max: 8,
    markerWidthPx: 6,
    markerHeightPx: 34,
    offsetXPx: 2,
    offsetYPx: 0,
  },
  {
    label: 'rapport_hd',
    sourceKey: 'Rapport H/D',
    range: 'B28:F28',
    min: 0,
    max: 2,
    markerWidthPx: 6,
    markerHeightPx: 34,
    offsetXPx: 0,
    offsetYPx: 0,
  },
],

// ------------------------------------------------------
// 2.3 Paramétrage de la recherche et de l'insertion des images
// ------------------------------------------------------
// keyColumn : colonne utilisée comme identifiant principal de la variété.
// allowedExtensions : formats d'images acceptés par le script.
//
// Les images sont recherchées dans le dossier photosDir.
// Le nom du fichier image doit contenir le code ou le nom de la variété.
// Les mots présents dans fileHints aident le script à savoir si l'image
// correspond plutôt à l'arbre, au fruit ou à une coupe.
//
// À modifier si :
// - la colonne d'identification change ;
// - d'autres formats d'image doivent être acceptés ;
// - les photos changent d'emplacement dans le modèle ;
// - les mots utilisés dans les noms de fichiers photo changent.
// ------------------------------------------------------

  keyColumn: 'code_variete',
  allowedExtensions: ['.jpg', '.jpeg', '.png', '.webp'],

  // Haut = arbre / Bas = fruit ou coupe
 imageSlots: [
  {
    label: 'arbre',
    fileHints: ['arbre', 'tree', 'port'],
    range: 'G4:H11',
    alignX: 'right',
    alignY: 'top',
  },
  {
    label: 'fruit',
    fileHints: ['coupe', 'cut', 'fruit', 'fruits'],
    range: 'G15:H23',
    alignX: 'right',
    alignY: 'top',
  },
],
};

// ======================================================
// 3. VARIÉTÉS DE RÉFÉRENCE ET MARQUEURS ASSOCIÉS
// ======================================================
// Cette partie indique quelle variété doit servir de repère selon le groupe.
// Exemple :
// - les citrons sont comparés au citron Frost Eureka ;
// - les limes sont comparées à la lime mexicaine.
//
// Le script utilise ensuite ces codes pour retrouver les données de référence
// dans le fichier Excel source et placer un marqueur jaune ou vert sur les frises.
//
// À modifier si :
// - la variété de référence change ;
// - le code_variete de la variété de référence change dans le fichier Excel ;
// - de nouveaux groupes doivent être ajoutés.
// ======================================================

const REFERENCE_VARIETIES = {
  citron: 'citron_frost_eureka',
  lime: 'lime_mexicaine',
};

// ------------------------------------------------------
// 3.1 Identification du groupe de la variété
// ------------------------------------------------------
// Ces fonctions servent à déterminer si la fiche concerne un citron ou une lime,
// puis à choisir la variété de référence et la couleur du marqueur associé.
// Le lien avec les références se fait par les colonnes espece ou groupe.
// ------------------------------------------------------

function getSpeciesKey(varietyData) {
  const raw = slugify(varietyData.espece || varietyData.groupe || '');

  if (raw.includes('lime')) return 'lime';
  if (raw.includes('citron')) return 'citron';

  return null;
}

function getReferenceCode(varietyData) {
  const speciesKey = getSpeciesKey(varietyData);
  if (!speciesKey) return null;
  return REFERENCE_VARIETIES[speciesKey] || null;
}

function getReferenceMarkerPath(varietyData) {
  const speciesKey = getSpeciesKey(varietyData);

  if (speciesKey === 'lime') return CONFIG.markerVertImagePath;
  if (speciesKey === 'citron') return CONFIG.markerJauneImagePath;

  return null;
}

function findVarietyByCode(allVarieties, code) {
  const target = slugify(code || '');

  return allVarieties.find((row) => {
    const rowCode = slugify(row[CONFIG.keyColumn] || row.code_variete || '');
    return rowCode === target;
  }) || null;
}


// ======================================================
// 4. OUTILS GÉNÉRAUX DE NETTOYAGE ET DE TEXTE
// ======================================================
// Ces fonctions ne sont pas spécifiques aux fiches.
// Elles servent à :
// - normaliser les noms et codes pour faciliter les comparaisons ;
// - créer les dossiers manquants ;
// - lire proprement le contenu des cellules Excel ;
// - remplacer les balises du modèle par les données de la variété.
//
// À modifier seulement si le format des données ou des balises change.
// ======================================================

function slugify(text) {
  return String(text || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

async function ensureDir(dirPath) {
  await fsp.mkdir(dirPath, { recursive: true });
}

function getCellStringValue(cell) {
  if (!cell || cell.value == null) return '';

  if (typeof cell.value === 'object') {
    if (cell.value.richText) {
      return cell.value.richText.map((part) => part.text || '').join('');
    }
    if (cell.value.text) {
      return String(cell.value.text);
    }
    if (cell.value.result != null) {
      return String(cell.value.result);
    }
  }

  return String(cell.value);
}

function replacePlaceholdersInText(text, data) {
  if (!text || typeof text !== 'string') return text;

  return text.replace(/\{\{(.*?)\}\}/g, (_, key) => {
    const cleanKey = String(key).trim();
    const value = data[cleanKey];
    return value != null ? String(value) : '';
  });
}

// ======================================================
// 5. OUTILS DE POSITIONNEMENT DANS EXCEL
// ======================================================
// Ces fonctions convertissent les plages de cellules Excel en dimensions
// utilisables pour placer précisément les images et les marqueurs.
//
// Elles permettent au script de comprendre une plage comme B22:F22,
// de calculer sa largeur/hauteur en pixels, puis de placer un élément
// au bon endroit dans cette plage.
//
// À modifier uniquement si le système de placement dans Excel doit changer.
// ======================================================

function columnLetterToNumber(col) {
  let num = 0;
  for (let i = 0; i < col.length; i++) {
    num = num * 26 + (col.charCodeAt(i) - 64);
  }
  return num;
}

function parseRange(range) {
  const match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
  if (!match) {
    throw new Error(`Plage invalide : ${range}`);
  }

  const [, startColL, startRow, endColL, endRow] = match;

  return {
    startCol: columnLetterToNumber(startColL.toUpperCase()),
    startRow: Number(startRow),
    endCol: columnLetterToNumber(endColL.toUpperCase()),
    endRow: Number(endRow),
  };
}

function colWidthToPixels(width) {
  const w = width || 8.43;
  return Math.floor(w * 7 + 5);
}

function rowHeightToPixels(height) {
  const h = height || 15;
  return Math.floor((h * 96) / 72);
}

function getRangePixelSize(worksheet, range) {
  const { startCol, startRow, endCol, endRow } = parseRange(range);

  let widthPx = 0;
  for (let c = startCol; c <= endCol; c++) {
    widthPx += colWidthToPixels(worksheet.getColumn(c).width);
  }

  let heightPx = 0;
  for (let r = startRow; r <= endRow; r++) {
    heightPx += rowHeightToPixels(worksheet.getRow(r).height);
  }

  return { widthPx, heightPx, startCol, startRow };
}

function pixelsToColumnOffset(worksheet, startCol, pixels) {
  let remaining = pixels;
  let offset = 0;
  let col = startCol;

  while (remaining > 0) {
    const colPx = colWidthToPixels(worksheet.getColumn(col).width);

    if (remaining >= colPx) {
      offset += 1;
      remaining -= colPx;
      col += 1;
    } else {
      offset += remaining / colPx;
      remaining = 0;
    }
  }

  return offset;
}

function pixelsToRowOffset(worksheet, startRow, pixels) {
  let remaining = pixels;
  let offset = 0;
  let row = startRow;

  while (remaining > 0) {
    const rowPx = rowHeightToPixels(worksheet.getRow(row).height);

    if (remaining >= rowPx) {
      offset += 1;
      remaining -= rowPx;
      row += 1;
    } else {
      offset += remaining / rowPx;
      remaining = 0;
    }
  }

  return offset;
}

// ======================================================
// 6. INSERTION DES PHOTOGRAPHIES DANS LE MODÈLE
// ======================================================
// Cette fonction ajoute une image dans une plage de cellules donnée.
// L'image est redimensionnée automatiquement pour rentrer dans la zone
// sans être déformée. Elle est centrée dans l'emplacement prévu.
//
// À modifier si :
// - les images doivent remplir toute la zone quitte à être coupées ;
// - le mode d'ancrage dans Excel doit changer ;
// - le comportement de centrage doit être adapté.
// ======================================================

function addImageContain(workbook, worksheet, imagePath, range) {
  if (!imagePath || !fs.existsSync(imagePath)) return;

  const ext = path.extname(imagePath).replace('.', '').toLowerCase();
  const excelExt = ext === 'jpg' ? 'jpeg' : ext;

  const fileBuffer = fs.readFileSync(imagePath);
  const imgSize = sizeOf(fileBuffer);

  const imgWidth = imgSize.width;
  const imgHeight = imgSize.height;

  if (!imgWidth || !imgHeight) {
    throw new Error(`Impossible de lire la taille de l'image : ${imagePath}`);
  }

  const { widthPx, heightPx, startCol, startRow } = getRangePixelSize(
    worksheet,
    range
  );

  const scale = Math.min(widthPx / imgWidth, heightPx / imgHeight);
  const finalWidth = Math.max(1, Math.round(imgWidth * scale));
  const finalHeight = Math.max(1, Math.round(imgHeight * scale));

  const offsetXPx = Math.max(0, Math.floor((widthPx - finalWidth) / 2));
  const offsetYPx = Math.max(0, Math.floor((heightPx - finalHeight) / 2));

  const colOffset = pixelsToColumnOffset(worksheet, startCol, offsetXPx);
  const rowOffset = pixelsToRowOffset(worksheet, startRow, offsetYPx);

  const imageId = workbook.addImage({
    filename: imagePath,
    extension: excelExt,
  });

  worksheet.addImage(imageId, {
    tl: {
      col: startCol - 1 + colOffset,
      row: startRow - 1 + rowOffset,
    },
    ext: {
      width: finalWidth,
      height: finalHeight,
    },
    editAs: 'oneCell',
  });
}

// ======================================================
// 7. LECTURE DES VALEURS NUMÉRIQUES ET MARQUEURS DE FRISE
// ======================================================
// Ces fonctions servent à lire les valeurs numériques du fichier Excel,
// même lorsqu'elles sont écrites avec une virgule ou un symbole %.
// Elles calculent ensuite la position du marqueur sur la frise selon :
// position = (valeur - min) / (max - min).
//
// À modifier si :
// - de nouveaux formats de nombres doivent être acceptés ;
// - les marqueurs doivent être placés autrement ;
// - les frises ne sont plus linéaires.
// ======================================================

function parseNumericValue(value) {
  if (value == null || value === '') return null;

  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }

  if (typeof value === 'object') {
    if (value.result != null) return parseNumericValue(value.result);
    if (value.text != null) return parseNumericValue(value.text);
    if (value.richText) {
      const joined = value.richText.map((part) => part.text || '').join('');
      return parseNumericValue(joined);
    }
  }

  const text = String(value)
    .trim()
    .replace(/\s+/g, '')
    .replace('%', '')
    .replace(',', '.');

  const num = Number(text);
  return Number.isFinite(num) ? num : null;
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function addMarkerToScale(
  workbook,
  worksheet,
  markerImagePath,
  range,
  rawValue,
  min,
  max,
  options = {}
) {
  if (!markerImagePath || !fs.existsSync(markerImagePath)) return;

  const value = parseNumericValue(rawValue);
  if (value == null) return;
  if (!Number.isFinite(min) || !Number.isFinite(max) || max <= min) return;

  const { widthPx, heightPx, startCol, startRow } = getRangePixelSize(
    worksheet,
    range
  );

  const markerWidthPx = options.markerWidthPx || 6;
const markerHeightPx = options.markerHeightPx || heightPx;
const manualOffsetXPx = options.offsetXPx || 0;
const manualOffsetYPx = options.offsetYPx || 0;

const ratio = clamp((value - min) / (max - min), 0, 1);

const usableWidthPx = Math.max(1, widthPx - markerWidthPx);
const offsetXPx = Math.max(
  0,
  Math.min(usableWidthPx, Math.round(ratio * usableWidthPx) + manualOffsetXPx)
);

const offsetYPx = Math.max(
  0,
  Math.floor((heightPx - markerHeightPx) / 2) + manualOffsetYPx
);

const colOffset = pixelsToColumnOffset(worksheet, startCol, offsetXPx);
const rowOffset = pixelsToRowOffset(worksheet, startRow, offsetYPx);

  const imageId = workbook.addImage({
    filename: markerImagePath,
    extension: 'png',
  });

  worksheet.addImage(imageId, {
    tl: {
      col: startCol - 1 + colOffset,
      row: startRow - 1 + rowOffset,
    },
    ext: {
      width: markerWidthPx,
      height: markerHeightPx,
    },
    editAs: 'oneCell',
  });
}
// ======================================================
// 8. RECHERCHE DES IMAGES ASSOCIÉES À UNE VARIÉTÉ
// ======================================================
// Cette partie cherche les images correspondant à chaque variété.
// Le script compare le nom des fichiers images avec le code ou le nom
// de la variété après simplification par slugify().
//
// Les mots-clés dans fileHints permettent ensuite d'attribuer une image
// au bon emplacement : arbre, fruit, coupe, etc.
//
// À modifier si :
// - les noms des images suivent une nouvelle convention ;
// - de nouveaux emplacements image sont ajoutés ;
// - de nouveaux mots-clés doivent être reconnus.
// ======================================================

function isAllowedImageFile(fileName) {
  const ext = path.extname(fileName).toLowerCase();
  return CONFIG.allowedExtensions.includes(ext);
}

function getVarietyKeys(varietyData) {
  const keys = [
    varietyData.code_variete,
    varietyData.nom_variete,
    varietyData[CONFIG.keyColumn],
  ]
    .map((v) => slugify(v))
    .filter(Boolean);

  return [...new Set(keys)];
}

function findImagesForVariety(photosDir, varietyData) {
  if (!fs.existsSync(photosDir)) {
    return {};
  }

  const allFiles = fs.readdirSync(photosDir).filter(isAllowedImageFile);
  const keys = getVarietyKeys(varietyData);

  if (keys.length === 0) {
    return {};
  }

  const matchingFiles = allFiles.filter((file) => {
    const name = slugify(path.parse(file).name);
    return keys.some((key) => name.includes(key));
  });

  const unused = [...matchingFiles];
  const assigned = {};

  // ------------------------------------------------------
// 11.4 Insertion des photos dans les emplacements du modèle
// ------------------------------------------------------
// Pour chaque emplacement image défini dans CONFIG.imageSlots, le script
// ajoute l'image trouvée dans la plage Excel correspondante.
// ------------------------------------------------------

for (const slot of CONFIG.imageSlots) {
    let foundIndex = unused.findIndex((file) => {
      const name = slugify(path.parse(file).name);
      return slot.fileHints.some((hint) => name.includes(slugify(hint)));
    });

    if (foundIndex === -1 && unused.length > 0) {
      foundIndex = 0;
    }

    if (foundIndex !== -1) {
      const selected = unused.splice(foundIndex, 1)[0];
      assigned[slot.label] = path.join(photosDir, selected);
    } else {
      assigned[slot.label] = null;
    }
  }

  return assigned;
}

// ======================================================
// 9. LECTURE DU TABLEAU DE DONNÉES EXCEL
// ======================================================
// Cette partie transforme la feuille Excel source en liste d'objets JavaScript.
// La première ligne du tableau est utilisée comme ligne d'en-têtes.
// Chaque ligne suivante correspond à une variété.
//
// Important :
// Les noms des colonnes doivent correspondre aux balises du modèle Excel.
// Exemple : une cellule contenant {{nom_variete}} sera remplacée par la
// valeur de la colonne nom_variete.
//
// À modifier si :
// - les en-têtes ne sont plus sur la première ligne ;
// - la colonne d'identification change ;
// - il faut ignorer certaines lignes du tableau.
// ======================================================

function extractRowsFromWorksheet(worksheet) {
  const headerRow = worksheet.getRow(1);
  const headers = headerRow.values
    .slice(1)
    .map((header) => String(header || '').trim());

  const rows = [];

  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
    const row = worksheet.getRow(rowNumber);
    const obj = {};
    let hasData = false;

    headers.forEach((header, index) => {
      const cell = row.getCell(index + 1);
      const rawValue = cell.value;

      let value = rawValue;

      if (rawValue && typeof rawValue === 'object') {
        if (rawValue.richText) {
          value = rawValue.richText.map((part) => part.text || '').join('');
        } else if (rawValue.text) {
          value = rawValue.text;
        } else if (rawValue.result != null) {
          value = rawValue.result;
        }
      }

      if (value != null && value !== '') {
        hasData = true;
      }

      obj[header] = value != null ? value : '';
    });

    if (!hasData) continue;

    const keyValue = String(
      obj[CONFIG.keyColumn] || obj.nom_variete || ''
    ).trim();

    if (keyValue !== '') {
      rows.push(obj);
    }
  }

  return rows;
}

async function loadSourceData() {
  if (!fs.existsSync(CONFIG.sourceExcelPath)) {
    throw new Error(
      `Fichier de données introuvable : ${CONFIG.sourceExcelPath}`
    );
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(CONFIG.sourceExcelPath);

  const worksheet = workbook.getWorksheet(CONFIG.sourceSheetName);
  if (!worksheet) {
    throw new Error(`Feuille introuvable : ${CONFIG.sourceSheetName}`);
  }

  return extractRowsFromWorksheet(worksheet);
}

// ======================================================
// 10. REMPLISSAGE DU MODÈLE EXCEL
// ======================================================
// Cette fonction parcourt toutes les cellules du modèle et remplace
// les balises écrites sous la forme {{nom_colonne}} par les valeurs
// correspondantes dans le fichier de données.
//
// À modifier si le format des balises change.
// Exemple actuel : {{acidite}}, {{sucre_indice}}, {{ref_acidite}}.
// ======================================================

function replaceAllPlaceholders(worksheet, varietyData) {
  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber++) {
    const row = worksheet.getRow(rowNumber);

    for (let colNumber = 1; colNumber <= worksheet.columnCount; colNumber++) {
      const cell = row.getCell(colNumber);
      const cellText = getCellStringValue(cell);

      if (cellText.includes('{{') && cellText.includes('}}')) {
        cell.value = replacePlaceholdersInText(cellText, varietyData);
      }
    }
  }
}

function buildOutputFileName(varietyData) {
  const baseName = slugify(
    varietyData.code_variete || varietyData.nom_variete || 'fiche'
  );
  return `${baseName}_fiche.xlsx`;
}

// ======================================================
// 11. GÉNÉRATION D'UNE FICHE POUR UNE VARIÉTÉ
// ======================================================
// Cette fonction est le cœur du script.
// Pour une variété donnée, elle :
// - ouvre le modèle de fiche ;
// - récupère la variété de référence si elle existe ;
// - prépare les données à insérer dans le modèle ;
// - remplace les balises ;
// - insère les images ;
// - place les marqueurs sur les frises ;
// - enregistre la fiche générée.
//
// À modifier avec prudence, car cette partie relie toutes les fonctions
// précédentes entre elles.
// ======================================================

async function generateOneSheet(varietyData, allVarieties) {
  if (!fs.existsSync(CONFIG.templateExcelPath)) {
    throw new Error(`Modèle introuvable : ${CONFIG.templateExcelPath}`);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(CONFIG.templateExcelPath);

  const worksheet = workbook.getWorksheet(1);
  if (!worksheet) {
    throw new Error('Impossible de lire la première feuille du modèle.');
  }

  const currentCode = slugify(varietyData[CONFIG.keyColumn] || varietyData.code_variete || '');
const referenceCode = getReferenceCode(varietyData);
const referenceData = referenceCode ? findVarietyByCode(allVarieties, referenceCode) : null;
const referenceMarkerPath = getReferenceMarkerPath(varietyData);

// ------------------------------------------------------
// 11.1 Préparation des données envoyées au modèle
// ------------------------------------------------------
// templateData contient les données de la variété en cours, puis ajoute
// les valeurs de référence avec le préfixe ref_.
// Ces noms doivent correspondre aux balises présentes dans le modèle Excel.
// Exemple : {{ref_acidite}} utilisera templateData.ref_acidite.
// ------------------------------------------------------

const templateData = {
  ...varietyData,
  ref_nom_variete: referenceData?.nom_variete || '',
  ref_code_variete: referenceData?.code_variete || '',
  ref_vigueur: referenceData?.vigueur || '',
  ref_port: referenceData?.port || '',
  ref_epines: referenceData?.epines || '',
  ref_particularite: referenceData?.particularite || '',
  ref_porte_greffe: referenceData?.porte_greffe || '',
  ref_sensibilite_froid: referenceData?.sensibilite_froid || '',
  ref_floraison: referenceData?.floraison || '',
  ref_nouaison: referenceData?.nouaison || '',
  ref_ecart_floraison_recolte: referenceData?.ecart_floraison_recolte || '',
  ref_periode_recolte: referenceData?.periode_recolte || '',
  ref_rapport_hd: referenceData?.['Rapport H/D'] || '',
  ref_calibre: referenceData?.calibre || '',
  ref_peau: referenceData?.peau || '',
  ref_aromes: referenceData?.aromes || '',
  ref_pepins: referenceData?.pepins || '',
  ref_coloration: referenceData?.coloration || '',
  ref_epluchage_ecorce: referenceData?.epluchage_ecorce || '',
  ref_poids_moyen_g: referenceData?.poids_moyen_g || '',
  ref_sucre_indice: referenceData?.sucre_indice || '',
  ref_rendement_jus_pct: referenceData?.rendement_jus_pct || '',
  ref_acidite: referenceData?.acidite || '',
  ref_ph: referenceData?.pH || '',
  ref_forme_fruit: referenceData?.forme_fruit || '',
};

replaceAllPlaceholders(worksheet, templateData);

// ------------------------------------------------------
// 11.2 Messages de contrôle pour les valeurs de référence
// ------------------------------------------------------
// Ces console.log servent à vérifier que la variété de référence est bien
// retrouvée et que les valeurs ref_ sont correctement récupérées.
// Ils peuvent être supprimés ou commentés si la console devient trop chargée.
// ------------------------------------------------------

console.log('referenceCode =', referenceCode);
console.log('ref_vigueur =', templateData.ref_vigueur);
console.log('ref_port =', templateData.ref_port);
console.log('ref_epines =', templateData.ref_epines);
console.log('ref_calibre =', templateData.ref_calibre);
console.log('ref_peau =', templateData.ref_peau);
console.log('ref_coloration =', templateData.ref_coloration);
console.log('ref_pepins =', templateData.ref_pepins);
console.log('ref_sucre_indice =', templateData.ref_sucre_indice);
console.log('ref_rendement_jus_pct =', templateData.ref_rendement_jus_pct);
console.log('ref_acidite =', templateData.ref_acidite);
console.log('ref_rapport_hd =', templateData.ref_rapport_hd);

replaceAllPlaceholders(worksheet, templateData);

// ------------------------------------------------------
// 11.3 Recherche des images correspondant à la variété
// ------------------------------------------------------
// assignedImages contient les chemins des photos sélectionnées pour chaque
// emplacement défini dans CONFIG.imageSlots.
// ------------------------------------------------------

const assignedImages = findImagesForVariety(CONFIG.photosDir, varietyData);

console.log('--- DEBUG IMAGES ---');
console.log('Variété :', varietyData.code_variete || varietyData.nom_variete);
console.log('assignedImages =', assignedImages);

console.log('--- DEBUG FICHE ---');
console.log('Colonnes disponibles :', Object.keys(varietyData));
console.log('Valeur sucre_indice :', varietyData['sucre_indice']);
console.log('PNG existe ? :', fs.existsSync(CONFIG.markerImagePath));
console.log('Chemin PNG :', CONFIG.markerImagePath);

for (const slot of CONFIG.imageSlots) {
  const imagePath = assignedImages[slot.label];
  console.log('Ajout image slot =', slot.label, 'range =', slot.range, 'path =', imagePath);

  if (imagePath) {
    addImageContain(workbook, worksheet, imagePath, slot.range);
  }
}

// ------------------------------------------------------
// 11.5 Ajout des marqueurs sur les frises
// ------------------------------------------------------
// Pour chaque frise définie dans CONFIG.scaleSlots, le script ajoute :
// - le marqueur de référence jaune ou vert si une référence existe ;
// - le marqueur rouge correspondant à la variété de la fiche.
//
// Les valeurs utilisées proviennent des colonnes indiquées par sourceKey.
// ------------------------------------------------------

for (const scale of CONFIG.scaleSlots) {
  console.log(
    'SCALE DEBUG =>',
    scale.label,
    'range =', scale.range,
    'value =', varietyData[scale.sourceKey],
    'min =', scale.min,
    'max =', scale.max,
    'offsetXPx =', scale.offsetXPx,
    'offsetYPx =', scale.offsetYPx,
    'pct =',
    (((Number(varietyData[scale.sourceKey]) - scale.min) / (scale.max - scale.min)) * 100).toFixed(1) + '%'
  );

  // marqueur de référence : jaune pour citron, vert pour lime
  if (
    referenceData &&
    referenceMarkerPath &&
    slugify(referenceCode) !== currentCode
  ) {
    addMarkerToScale(
      workbook,
      worksheet,
      referenceMarkerPath,
      scale.range,
      referenceData[scale.sourceKey],
      scale.min,
      scale.max,
      {
        markerWidthPx: scale.markerWidthPx,
        markerHeightPx: scale.markerHeightPx,
        offsetXPx: scale.offsetXPx,
        offsetYPx: scale.offsetYPx,
      }
    );
  }

  // marqueur rouge : variété de la fiche
  addMarkerToScale(
    workbook,
    worksheet,
    CONFIG.markerRougeImagePath,
    scale.range,
    varietyData[scale.sourceKey],
    scale.min,
    scale.max,
    {
      markerWidthPx: scale.markerWidthPx,
      markerHeightPx: scale.markerHeightPx,
      offsetXPx: scale.offsetXPx,
      offsetYPx: scale.offsetYPx,
    }
  );
}

  // ------------------------------------------------------
  // 11.6 Enregistrement de la fiche générée
  // ------------------------------------------------------
  // Le fichier final est enregistré dans CONFIG.outputDir avec un nom
  // construit à partir du code ou du nom de la variété.
  // ------------------------------------------------------

  const outputPath = path.join(
    CONFIG.outputDir,
    buildOutputFileName(varietyData)
  );

  await workbook.xlsx.writeFile(outputPath);
  return outputPath;
}

// ======================================================
// 12. EXÉCUTION GÉNÉRALE DU SCRIPT
// ======================================================
// countImages() compte les images disponibles dans le dossier photo.
// main() lance toutes les étapes : création du dossier de sortie, lecture
// des données, génération de chaque fiche, puis affichage des messages finaux.
//
// À modifier si :
// - il faut générer seulement certaines variétés ;
// - il faut ajouter un filtre ;
// - il faut changer le comportement général du script.
// ======================================================

async function countImages() {
  if (!fs.existsSync(CONFIG.photosDir)) return 0;
  const files = await fsp.readdir(CONFIG.photosDir);
  return files.filter(isAllowedImageFile).length;
}

async function main() {
  try {
    await ensureDir(CONFIG.outputDir);

    const varietes = await loadSourceData();
    const numberOfImages = await countImages();

    console.log(`Nombre de variétés : ${varietes.length}`);
    console.log(`Nombre d'images : ${numberOfImages}`);

    for (const variete of varietes) {
      const outputPath = await generateOneSheet(variete, varietes);
      console.log(`Fiche générée : ${outputPath}`);
    }

    console.log('Terminé.');
  } catch (error) {
    console.error(`Erreur : ${error.message}`);
  }
}

main();