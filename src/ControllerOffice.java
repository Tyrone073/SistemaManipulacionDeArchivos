import java.io.*;
import java.util.zip.*;
import java.nio.charset.StandardCharsets;

public class ControllerOffice {

    public static void crearDocx(String nombreArchivo, Contenido contenido) throws IOException {
       // variables
        byte[] contentTypes = generarContentTypes("Docx").getBytes(StandardCharsets.UTF_8);
        byte[] rels = generarRels().getBytes(StandardCharsets.UTF_8);
        byte[] document = generarDocumento(contenido).getBytes(StandardCharsets.UTF_8);
        byte[] header = generarEncabezado(contenido).getBytes(StandardCharsets.UTF_8);
        byte[] documentRels = generarDocumentRels().getBytes(StandardCharsets.UTF_8);

        // archivo ZIP con estructura necesaria para documento docx
        try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(nombreArchivo + ".docx"))) {
            // Archivos obligatorios
            agregarAlZip(zos, "[Content_Types].xml", contentTypes);
            agregarAlZip(zos, "_rels/.rels", rels);
            agregarAlZip(zos, "word/document.xml", document);
            agregarAlZip(zos, "word/header1.xml", header);
            agregarAlZip(zos, "word/_rels/document.xml.rels", documentRels);
        }
    }
    public static void crearXlsx(String nombreArchivo, Contenido contenido) throws IOException {
        byte[] contentTypes = generarContentTypes("xlsx").getBytes(StandardCharsets.UTF_8);
        byte[] rels = generarRels().getBytes(StandardCharsets.UTF_8);
        byte[] workbook = generarLibroDeTrabajoXml(contenido).getBytes(StandardCharsets.UTF_8);

        try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(nombreArchivo + ".xlsx"))) {
            agregarAlZip(zos, "[Content_Types].xml", contentTypes);
            agregarAlZip(zos, "_rels/.rels", rels);
            agregarAlZip(zos, "xl/workbook.xml", workbook);
            // falta añadir más archivos necesarios (sheets, styles, etc.)
        }
    }
    public static void crearPptx(String nombreArchivo, Contenido contenido) throws IOException {

        try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(nombreArchivo + ".pptx"))) {
            // Archivos necesarios
            agregarAlZip(zos, "[Content_Types].xml", generarContentTypes("pptx").getBytes(StandardCharsets.UTF_8));
            agregarAlZip(zos, "_rels/.rels", generarRelsPptx().getBytes(StandardCharsets.UTF_8));
            agregarAlZip(zos, "ppt/presentation.xml", generarHolaDePresentationXml().getBytes(StandardCharsets.UTF_8));
            agregarAlZip(zos, "ppt/slides/slide1.xml", generarDiapositiva1Xml(contenido).getBytes(StandardCharsets.UTF_8));
            agregarAlZip(zos, "ppt/slideMasters/slideMaster1.xml", generarPlantilladeDiapositivasXml().getBytes(StandardCharsets.UTF_8));
            agregarAlZip(zos, "ppt/_rels/presentation.xml.rels", generarRelsPptx().getBytes(StandardCharsets.UTF_8));
        }
    }

    private static void agregarAlZip(ZipOutputStream zos, String nombreEntry, byte[] contenido) throws IOException {
        ZipEntry entry = new ZipEntry(nombreEntry);
        zos.putNextEntry(entry);
        zos.write(contenido);
        zos.closeEntry();
    }

    public static String generarContentTypes(String tipoDocumento) {
        String contentType = "";

        switch (tipoDocumento.toLowerCase()) {
            case "docx":
                contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
                break;
            case "xlsx":
                contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
                break;
            case "pptx":
                contentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
                break;
            default:
                throw new IllegalArgumentException("Tipo de documento no soportado: " + tipoDocumento);
        }

        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n" +
                "  <Default Extension=\"rels\" contentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n" +
                "  <Default Extension=\"xml\" contentType=\"application/xml\"/>\n" +
                "  <Override PartName=\"/" + tipoDocumento.split("x")[0] + "/document.xml\" contentType=\"" + contentType + "\"/>\n" +
                "</Types>";
    }

    private static String generarRels() {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
                "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>\n" +
                "</Relationships>";
    }

    private static String generarDocumento(Contenido contenido) {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"\n" +
                "            xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n" +
                "  <w:body>\n" +
                "    <w:p>\n" +
                "      <w:r>\n" +
                "        <w:t>" + escapeXml(contenido.getMensaje()) + "</w:t>\n" +
                "      </w:r>\n" +
                "    </w:p>\n" +
                "    <w:sectPr>\n" +
                "      <w:headerReference w:type=\"default\" r:id=\"rId1\"/>\n" +
                "      <w:pgSz w:w=\"12240\" w:h=\"15840\"/>\n" +
                "      <w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/>\n" +
                "    </w:sectPr>\n" +
                "  </w:body>\n" +
                "</w:document>";
    }

    private static String generarEncabezado(Contenido contenido) {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\n" +
                "  <w:p>\n" +
                "    <w:r>\n" +
                "      <w:t>" + "Creado el " + escapeXml(contenido.getFechaActual()) + "</w:t>\n" +
                "    </w:r>\n" +
                "  </w:p>\n" +
                "</w:hdr>";
    }

    private static String generarDocumentRels() {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
                "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>\n" +
                "</Relationships>";
    }

    private static String escapeXml(String input) {
        return input.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&apos;");
    }

    private static String generarLibroDeTrabajoXml(Contenido contenido) {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n" +
                "  <sheets>\n" +
                "    <sheet name=\"Hoja1\" sheetId=\"1\" r:id=\"rId1\"/>\n" +
                "  </sheets>\n" +
                "</workbook>";
    }

    private static String generarHolaDePresentationXml() {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"\n" +
                "                xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n" +
                "  <p:sldMasterIdLst>\n" +
                "    <p:sldMasterId id=\"2147483648\" r:id=\"rId1\"/>\n" +  // Requerido: referencia al slide master
                "  </p:sldMasterIdLst>\n" +
                "  <p:sldIdLst>\n" +
                "    <p:sldId id=\"256\" r:id=\"rId2\"/>\n" +  // ID de la primera diapositiva
                "  </p:sldIdLst>\n" +
                "  <p:sldSz cx=\"9144000\" cy=\"6858000\" type=\"screen4x3\"/>\n" +  // Tamaño de diapositiva (16:9 sería cx=9144000, cy=5143500)
                "  <p:notesSz cx=\"6858000\" cy=\"9144000\"/>\n" +  // Tamaño de notas
                "</p:presentation>";
    }

    private static String generarPlantilladeDiapositivasXml() {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<p:sldMaster xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">\n" +
                "  <p:cSld>\n" +
                "    <p:bg>\n" +
                "      <p:bgRef idx=\"1001\"/>\n" +  // Fondo por defecto
                "    </p:bg>\n" +
                "  </p:cSld>\n" +
                "  <p:sldLayoutIdLst>\n" +
                "    <p:sldLayoutId id=\"2147483649\" r:id=\"rId1\"/>\n" +  // Referencia al layout
                "  </p:sldLayoutIdLst>\n" +
                "</p:sldMaster>";
    }

    private static String generarDiapositiva1Xml(Contenido contenido) {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"\n" +
                "       xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"\n" +
                "       xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n" +
                "  <p:cSld>\n" +
                "    <p:spTree>\n" +
                "      <p:sp>\n" +
                "        <p:txBody>\n" +
                "          <a:p>\n" +
                "            <a:r>\n" +
                "              <a:t>" + escapeXml(contenido.getMensaje()) + "</a:t>\n" +
                "            </a:r>\n" +
                "          </a:p>\n" +
                "        </p:txBody>\n" +
                "      </p:sp>\n" +
                "    </p:spTree>\n" +
                "  </p:cSld>\n" +
                "  <p:clrMapOvr>\n" +
                "    <a:masterClrMapping/>\n" +
                "  </p:clrMapOvr>\n" +
                "</p:sld>";
    }

    private static String generarRelsPptx() {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
                "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"slideMasters/slideMaster1.xml\"/>\n" +
                "  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide1.xml\"/>\n" +
                "</Relationships>";
    }
}