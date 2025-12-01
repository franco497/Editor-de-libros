const mammoth = require("mammoth");
const fs = require("fs-extra");
const puppeteer = require("puppeteer");

class ConvertidorColumnasPDF {
  constructor() {
    this.archivoEntrada = "libro-prueba.docx";
    this.archivoSalida = "libro-prueba-dos-columnas.pdf";
    this.archivoTempHTML = "temp-libro.html";
  }

  async convertir() {
    try {
      console.log("📄 Leyendo DOCX...");
      const resultado = await mammoth.convertToHtml({
        path: this.archivoEntrada,
      });

      console.log("🎨 Creando HTML con 2 columnas...");
      const htmlFinal = this.generarHTML(resultado.value);

      await fs.writeFile(this.archivoTempHTML, htmlFinal, "utf8");

      console.log("🖨 Generando PDF...");
      await this.crearPDF();

      console.log(`\n✅ PDF generado: ${this.archivoSalida}`);
    } catch (e) {
      console.error("❌ ERROR:", e);
    }
  }

  generarHTML(contenidoHTML) {
    return `
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Documento</title>

<style>
  body {
    font-family: "Times New Roman", serif;
    margin: 0;
    padding: 0;
    background: white;
  }

  .pagina {
    width: 210mm;
    height: 297mm;
    padding: 25mm;
    box-sizing: border-box;
  }

  .dos-columnas {
    column-count: 2;
    column-gap: 20mm;
    text-align: justify;
    line-height: 1.35;
    font-size: 12pt;
  }

  @page {
    size: A4;
    margin: 20mm;
  }
</style>
</head>

<body>
  <div class="pagina">
    <div class="dos-columnas">
      ${contenidoHTML}
    </div>
  </div>
</body>
</html>`;
  }

  async crearPDF() {
    const browser = await puppeteer.launch({
      headless: "new",
    });

    const page = await browser.newPage();
    await page.goto("file://" + process.cwd() + "/" + this.archivoTempHTML, {
      waitUntil: "networkidle0",
    });

    await page.pdf({
      path: this.archivoSalida,
      format: "A4",
      printBackground: true,
    });

    await browser.close();
  }
}

new ConvertidorColumnasPDF().convertir();
