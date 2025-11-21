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
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Harary Keith - Sueños Lúcidos en 30 Días - 2 Columnas</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Times New Roman', Times, serif;
            font-size: 12pt;
            line-height: 1.6;
            color: #000;
            background: #fff;
            padding: 20px;
        }
        
        .contenedor-columnas {
            column-count: 2;
            column-gap: 25px;
            column-rule: 1px solid #ccc;
            text-align: justify;
        }
        
        p {
            margin-bottom: 12px;
            text-align: justify;
            orphans: 2; /* Mínimo de líneas al final de columna */
            widows: 2;  /* Mínimo de líneas al inicio de columna */
        }
        
        h1, h2, h3, h4, h5, h6 {
            column-span: all; /* Los títulos ocupan ambas columnas */
            margin: 20px 0 10px 0;
            text-align: center;
        }
        
        h1 {
            font-size: 18pt;
            margin-bottom: 30px;
        }
        
        h2 {
            font-size: 16pt;
        }
        
        h3 {
            font-size: 14pt;
        }
        
        /* Mejoras para impresión */
        @media print {
            body {
                padding: 10mm;
                margin: 0;
            }
            
            .contenedor-columnas {
                column-gap: 15mm;
            }
            
            @page {
                margin: 20mm;
            }
        }
        
        /* Estilos para tablas e imágenes */
        table, img {
            max-width: 100%;
            height: auto;
        }
        
        table {
            border-collapse: collapse;
            margin: 10px 0;
        }
        
        td, th {
            border: 1px solid #000;
            padding: 5px;
            font-size: 11pt;
        }
        
        /* Manejo de bloques de código o texto especial */
        pre, code {
            font-family: Consolas, monospace;
            font-size: 11pt;
            background: #f5f5f5;
            padding: 5px;
            border-radius: 3px;
        }
    </style>
</head>
<body>
    <div class="contenedor-columnas">
        ${htmlContent}
    </div>
    
    <script>
        // Script opcional para mejorar la visualización
        document.addEventListener('DOMContentLoaded', function() {
            // Asegurar que los párrafos tengan sangría
            const paragraphs = document.querySelectorAll('p');
            paragraphs.forEach(p => {
                if (!p.style.textIndent && p.textContent.trim().length > 0) {
                    p.style.textIndent = '1.5em';
                }
            });
            
            console.log('Documento formateado en 2 columnas listo');
        });
    </script>
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
