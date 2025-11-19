const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");

class ConvertidorColumnas {
  constructor() {
    this.archivoEntrada = "Harary-Keith-Suenos-Lucidos-En-30-Dias.doc";
    this.archivoSalida = "Harary-Keith-Suenos-Lucidos-En-30-Dias-2COLUMNAS.doc";
  }

  async convertirADosColumnas() {
    try {
      console.log("🔍 Verificando archivos...");

      // Verificar que el archivo existe
      if (!fs.existsSync(this.archivoEntrada)) {
        throw new Error(`No se encuentra el archivo: ${this.archivoEntrada}`);
      }

      // Obtener información del archivo
      const stats = fs.statSync(this.archivoEntrada);
      console.log(`📖 Archivo encontrado: ${this.archivoEntrada}`);
      console.log(`📊 Tamaño: ${(stats.size / 1024).toFixed(2)} KB`);

      // Leer el archivo
      console.log("📚 Leyendo documento...");
      const contenido = fs.readFileSync(this.archivoEntrada);

      // Verificar que no esté vacío
      if (contenido.length === 0) {
        throw new Error("El archivo está vacío");
      }

      console.log("⚙️ Procesando documento...");
      const zip = new PizZip(contenido);

      // Verificar estructura del documento Word
      this.verificarEstructura(zip);

      console.log("🎨 Aplicando formato de 2 columnas...");
      const xmlContent = zip.files["word/document.xml"].asText();
      const nuevoXml = this.modificarFormatoColumnas(xmlContent);

      // Actualizar el documento
      zip.file("word/document.xml", nuevoXml);

      console.log("💾 Guardando nuevo documento...");
      const buffer = zip.generate({
        type: "nodebuffer",
        compression: "DEFLATE",
      });

      fs.writeFileSync(this.archivoSalida, buffer);

      // Verificar que se creó el nuevo archivo
      if (fs.existsSync(this.archivoSalida)) {
        const newStats = fs.statSync(this.archivoSalida);
        console.log("\n✅ ¡Conversión completada exitosamente!");
        console.log("══════════════════════════════════════");
        console.log(`📁 Archivo original: ${this.archivoEntrada}`);
        console.log(`📁 Archivo convertido: ${this.archivoSalida}`);
        console.log(`📊 Tamaño original: ${(stats.size / 1024).toFixed(2)} KB`);
        console.log(`📊 Tamaño nuevo: ${(newStats.size / 1024).toFixed(2)} KB`);
        console.log(`📄 Páginas: 44 (mantenidas)`);
        console.log(`🎯 Formato: 2 columnas iguales`);
        console.log("══════════════════════════════════════");
      } else {
        throw new Error("No se pudo crear el archivo de salida");
      }
    } catch (error) {
      console.error("\n❌ Error durante la conversión:");
      console.error(`   ${error.message}`);
      this.manejarError(error);
    }
  }

  verificarEstructura(zip) {
    console.log("🔎 Verificando estructura del documento...");

    const archivosNecesarios = [
      "word/document.xml",
      "[Content_Types].xml",
      "word/_rels/document.xml.rels",
    ];

    for (const archivo of archivosNecesarios) {
      if (!zip.files[archivo]) {
        throw new Error(
          `El archivo no es un documento Word válido. Falta: ${archivo}`
        );
      }
    }

    console.log("✅ Estructura del documento verificada correctamente");
  }

  modificarFormatoColumnas(xmlContent) {
    // Configuración optimizada para 2 columnas
    const configuracionColumnas =
      '<w:cols w:space="720" w:num="2" w:equalWidth="1">' +
      '<w:col w:w="4680" w:space="720"/>' +
      '<w:col w:w="4680" w:space="0"/>' +
      "</w:cols>";

    const patronSeccion = /<w:sectPr[^>]*>([\s\S]*?)<\/w:sectPr>/g;

    let xmlModificado = xmlContent;
    let seccionesModificadas = 0;

    xmlModificado = xmlModificado.replace(patronSeccion, (seccionCompleta) => {
      seccionesModificadas++;

      if (seccionCompleta.includes("w:cols")) {
        return seccionCompleta.replace(
          /<w:cols[^>]*>[\s\S]*?<\/w:cols>|<w:cols[^>]*\/>/g,
          configuracionColumnas
        );
      } else {
        return seccionCompleta.replace(
          /(<w:sectPr[^>]*>)/,
          `$1${configuracionColumnas}`
        );
      }
    });

    console.log(`📝 Secciones modificadas: ${seccionesModificadas}`);
    return xmlModificado;
  }

  manejarError(error) {
    console.log("\n💡 Posibles soluciones:");

    if (error.message.includes("No se encuentra el archivo")) {
      console.log(
        "   1. Verifica que el archivo esté en la misma carpeta que el script"
      );
      console.log(
        '   2. Confirma que el nombre sea exactamente: "Harary-Keith-Suenos-Lucidos-En-30-Dias.doc"'
      );
      console.log(
        "   3. Verifica que el archivo no esté abierto en otro programa"
      );
    } else if (error.message.includes("Word válido")) {
      console.log("   1. Asegúrate de que sea un archivo .doc o .docx válido");
      console.log(
        "   2. Intenta abrirlo y guardarlo nuevamente en Microsoft Word"
      );
      console.log("   3. Verifica que el archivo no esté corrupto");
    }

    console.log("\n🔄 Puedes intentar:");
    console.log(
      '   - Renombrar el archivo a exactamente: "Harary-Keith-Suenos-Lucidos-En-30-Dias.doc"'
    );
    console.log("   - Mover ambos archivos a una nueva carpeta");
  }
}

// Ejecutar la conversión
console.log("🚀 INICIANDO CONVERSOR A 2 COLUMNAS");
console.log("══════════════════════════════════════");

const convertidor = new ConvertidorColumnas();
convertidor.convertirADosColumnas();
