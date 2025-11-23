const mammoth = require("mammoth");
const fs = require("fs");

class ConvertidorColumnas {
  constructor() {
    this.archivoEntrada = "Harary-Keith-Suenos-Lucidos-En-30-Dias.docx";
    this.archivoSalida =
      "Harary-Keith-Suenos-Lucidos-En-30-Dias-2COLUMNAS.html";
  }

  async convertirADosColumnas() {
    try {
      console.log("🚀 CONVERSOR A 2 COLUMNAS CON MAMMOTH");
      console.log("══════════════════════════════════════");
      console.log("🔍 Verificando archivos...");

      // Verificar que el archivo existe
      if (!fs.existsSync(this.archivoEntrada)) {
        throw new Error(`No se encuentra el archivo: ${this.archivoEntrada}`);
      }

      // Obtener información del archivo
      const stats = fs.statSync(this.archivoEntrada);
      console.log(`📖 Archivo encontrado: ${this.archivoEntrada}`);
      console.log(`📊 Tamaño: ${(stats.size / 1024).toFixed(2)} KB`);

      console.log("📚 Procesando documento DOC con Mammoth...");

      // Convertir el documento DOC a HTML
      const result = await mammoth.convertToHtml({
        path: this.archivoEntrada,
      });

      console.log("✅ Documento convertido a HTML correctamente");

      // Mostrar advertencias si las hay
      if (result.messages && result.messages.length > 0) {
        console.log("⚠️  Advertencias durante la conversión:");
        result.messages.forEach((message) => {
          console.log(`   - ${message.message}`);
        });
      }

      console.log("🎨 Aplicando formato de 2 columnas...");
      const htmlConColumnas = this.crearHTMLConColumnas(result.value);

      console.log("💾 Guardando documento HTML...");
      fs.writeFileSync(this.archivoSalida, htmlConColumnas);

      // Verificar que se creó el nuevo archivo
      if (fs.existsSync(this.archivoSalida)) {
        const newStats = fs.statSync(this.archivoSalida);
        console.log("\n✅ ¡Conversión completada exitosamente!");
        console.log("══════════════════════════════════════");
        console.log(`📁 Archivo original: ${this.archivoEntrada}`);
        console.log(`📁 Archivo con columnas: ${this.archivoSalida}`);
        console.log(`📊 Tamaño original: ${(stats.size / 1024).toFixed(2)} KB`);
        console.log(`📊 Tamaño nuevo: ${(newStats.size / 1024).toFixed(2)} KB`);
        console.log(`🎯 Formato: HTML con 2 columnas`);
        console.log("══════════════════════════════════════");
        console.log("\n💡 INSTRUCCIONES PARA TERMINAR:");
        console.log(
          "   1. Abre el archivo HTML en tu navegador para verificar"
        );
        console.log("   2. Si está correcto, ábrelo en Microsoft Word");
        console.log("   3. En Word, guarda como DOCX o PDF");
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

  crearHTMLConColumnas(htmlContent) {
    return `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Libro - 2 Columnas</title>

<style>
    body {
        font-family: "Times New Roman", serif;
        background: #eee;
        padding: 30px 0;
        display: flex;
        justify-content: center;
    }

    .pagina {
        width: 210mm;       /* Tamaño A4 */
        min-height: 297mm;
        padding: 25mm;
        background: white;
        box-shadow: 0 0 8px rgba(0,0,0,0.2);
        column-count: 2;
        column-gap: 20mm;
        column-rule: 1px solid #bbb;
        font-size: 12pt;
        text-align: justify;
        line-height: 1.4;
    }

    p {
        margin-bottom: 12px;
        text-indent: 1.5em;     /* Sangría de primera línea */
    }

    h1, h2, h3 {
        column-span: all;
        text-align: center;
        margin: 20px 0 10px;
        text-indent: 0;
    }

    h1 { font-size: 20pt; font-weight: bold; }
    h2 { font-size: 16pt; font-style: italic; }
    h3 { font-size: 14pt; }

    img, table {
        column-span: all;
        display: block;
        margin: 15px auto;
        max-width: 100%;
    }

    @media print {
        body {
            background: white;
            padding: 0;
        }
        .pagina {
            box-shadow: none;
            margin: 0;
        }
    }
</style>
</head>

<body>
    <div class="pagina">
        ${htmlContent}
    </div>
</body>
</html>`;
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
    } else if (
      error.message.includes("format") ||
      error.message.includes("corrupt")
    ) {
      console.log(
        "   1. El archivo podría estar corrupto o en formato no soportado"
      );
      console.log("   2. Intenta abrirlo y guardarlo nuevamente en Word");
      console.log(
        "   3. Verifica que sea un documento Word válido (.doc o .docx)"
      );
    } else {
      console.log("   1. Revisa que mammoth esté instalado: npm list mammoth");
      console.log(
        "   2. Verifica que el archivo no esté protegido con contraseña"
      );
      console.log(
        "   3. Intenta con un archivo más pequeño primero para probar"
      );
    }

    console.log("\n🔄 Comandos útiles:");
    console.log("   - Verificar instalación: npm list mammoth");
    console.log("   - Reinstalar mammoth: npm install mammoth@latest");
  }
}

// Ejecutar la conversión
console.log("Iniciando proceso de conversión...");
const convertidor = new ConvertidorColumnas();
convertidor.convertirADosColumnas();
