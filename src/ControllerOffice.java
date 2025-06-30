import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

public class ControllerOffice {

    static Charset caracteres = StandardCharsets.UTF_8;

    private static void agregarAlZip(ZipOutputStream zos, String nombreEntry, byte[] contenido) throws IOException {
        ZipEntry entry = new ZipEntry(nombreEntry);
        zos.putNextEntry(entry);
        zos.write(contenido);
        zos.closeEntry();
    }

    public static void crearDocx(String nombreArchivo, Contenido contenido) throws IOException {
        byte[] contentTypes = generarContentTypes("docx").getBytes(caracteres);
        byte[] rels = generarRels("docx").getBytes(caracteres);
        byte[] document = generarDocumentoXml(contenido).getBytes(caracteres);
        byte[] header = generarEncabezadoXml(contenido).getBytes(caracteres);
        byte[] documentRels = generarRelsHeaderConDocumentDocx().getBytes(caracteres);

        try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(nombreArchivo + ".docx"))) {
            agregarAlZip(zos, "[Content_Types].xml", contentTypes);
            agregarAlZip(zos, "_rels/.rels", rels);
            agregarAlZip(zos, "word/header1.xml", header);
            agregarAlZip(zos, "word/document.xml", document);
            agregarAlZip(zos, "word/_rels/document.xml.rels", documentRels);
        }
    }
    public static String leerArchivoOffice(String nombreArchivo) {
        try {
            FileInputStream fis = new FileInputStream(nombreArchivo);
            ZipInputStream zis = new ZipInputStream(fis);
            ZipEntry ze;
            List<String> contenido = new ArrayList<>();

            while ((ze = zis.getNextEntry()) != null) {
                if (ze.getName().equals("word/header1.xml") || ze.getName().equals("word/document.xml")) {
                    String ad = new String(zis.readAllBytes(), caracteres).replaceAll("<[^>]+>", "").trim();
//                    return ad;
                    if (!ad.isEmpty()) {
                        contenido.add(ad);
                    }
                }
                if (ze.getName().equals("xl/workbook.xml")) {
                    String ad = new String(zis.readAllBytes(), caracteres).replaceAll("<[^>]+>", "").trim();
                    if (!ad.isEmpty()) {
                        contenido.add(ad);
                    }
                }
                if (ze.getName().equals("ppt/presentation.xml")) {
                    String ad = new String(zis.readAllBytes(), caracteres).replaceAll("<[^>]+>", "").trim();
                    if (!ad.isEmpty()) {
                        contenido.add(ad);
                    }
                }
            }

            return String.join("\n", contenido);

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
    public static String leerArchivo(String nombreArchivo) throws IOException {

        FileReader fileR = new FileReader(nombreArchivo);
        BufferedReader bR = new BufferedReader(fileR);
        String linea;

        List<String> contenido = new ArrayList<>();

        while ((linea = bR.readLine()) != null) {
            String ad = new String(linea.getBytes(), caracteres).replaceAll("<[^>]+>", "").trim();
            if (!ad.isEmpty()) {
                contenido.add(ad);
            }
        }
        fileR.close();
        bR.close();
        return String.join("\n", contenido);
    }


    /* === XLSX === */
    public static void crearXlsx(String nombreArchivo, Contenido contenido) throws IOException {
        try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(nombreArchivo + ".xlsx"))) {
            agregarAlZip(zos, "[Content_Types].xml", generarContentTypes("xlsx").getBytes(caracteres));
            agregarAlZip(zos, "_rels/.rels", generarRels("xlsx").getBytes(caracteres));
            agregarAlZip(zos, "xl/workbook.xml", generarWorkbookXml(contenido).getBytes(caracteres));
        }
    }

    private static String generarWorkbookXml(Contenido contenido) {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n" +
                "  <sheets>\n" +
                "    <sheet name=\"Hoja1\" sheetId=\"1\" r:id=\"rId1\"/>\n" +
                "  </sheets>\n" +
                "</workbook>";
    }

    /* === PPTX === */
    public static void crearPptx(String nombreArchivo, Contenido contenido) throws IOException {
        try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(nombreArchivo + ".pptx"))) {
            agregarAlZip(zos, "[Content_Types].xml", generarContentTypes("pptx").getBytes(caracteres));
            agregarAlZip(zos, "_rels/.rels", generarRels("pptx").getBytes(caracteres));
            agregarAlZip(zos, "ppt/presentation.xml", generarPresentationXml(contenido).getBytes(caracteres));
        }
    }

//    public static void crearXml(String nombreArchivo, Contenido contenido) throws IOException {
//        String formato = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
//                "<documento tipo=\"mensaje\">\n" +
//                "  <contenido>\n" +
//                "    <texto>" + escapeXml(contenido.getMensaje()) + "</texto>\n" +
//                "    <fecha>" + escapeXml(contenido.getFechaActual()) + "</fecha>\n" +
//                "  </contenido>\n" +
//                "</documento>";
//
//        try (FileOutputStream fos = new FileOutputStream(nombreArchivo + ".xml")) {
//            fos.write(formato.getBytes(caracteres));
//        }
//    }
public static void crearSvg(String nombreArchivo, Contenido contenido) throws IOException {
    String svg = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
            "<svg width=\"600\" height=\"850\" xmlns=\"http://www.w3.org/2000/svg\">\n" +
            "  <rect width=\"600\" height=\"850\" fill=\"#ffffff\"/>\n" +
            "  <text x=\"70\" y=\"50\" font-family=\"Arial\" font-size=\"12\" fill=\"gray\">" + "Creado el " + escapeXml(contenido.getFechaActual()) + "</text>\n" +
            "  <text x=\"70\" y=\"80\" font-family=\"Arial\" font-size=\"12\" fill=\"black\">" + escapeXml(contenido.getMensaje()) + "</text>\n" +

            "</svg>";

    try (FileOutputStream fos = new FileOutputStream(nombreArchivo + ".svg")) {
        fos.write(svg.getBytes(caracteres));
    }
}


    private static String generarPresentationXml(Contenido contenido) {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\">\n" +
                "  <p:sldIdLst>\n" +
                "    <p:sldId id=\"256\" r:id=\"rId1\"/>\n" +
                "  </p:sldIdLst>\n" +
                "</p:presentation>";
    }

    /* === Utilitario para escapar caracteres en XML === */
    private static String escapeXml(String input) {
        return input.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;")
                .replace("'", "&apos;");
    }

    private static String generarContentTypes(String tipo) {
        String overrides = "";
        switch (tipo) {
            case "docx":
                overrides =
                        "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>\n" +
                                "<Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>";
                break;
            case "xlsx":
                overrides =
                        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>";
                break;
            case "pptx":
                overrides =
                        "<Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>";
                break;
            default:
                throw new IllegalArgumentException("Tipo no soportado: " + tipo);
        }

        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n" +
                "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n" +
                "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n" +
                "  " + overrides + "\n" +
                "</Types>";
    }

    private static String generarRels(String tipo) {
        String target;
        switch (tipo) {
            case "docx":
                target = "word/document.xml";
                break;
            case "xlsx":
                target = "xl/workbook.xml";
                break;
            case "pptx":
                target = "ppt/presentation.xml";
                break;
            default:
                throw new IllegalArgumentException("Tipo no soportado: " + tipo);
        }

        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
                "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"" + target + "\"/>\n" +
                "</Relationships>";
    }
    private static String generarDocumentoXml(Contenido contenido) {
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
                "    </w:sectPr>\n" +
                "  </w:body>\n" +
                "</w:document>";
    }

    private static String generarEncabezadoXml(Contenido contenido) {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\n" +
                "  <w:p>\n" +
                "    <w:r>\n" +
                "      <w:t>" + "Creado el " + escapeXml(contenido.getFechaActual()) + "</w:t>\n" +
                "    </w:r>\n" +
                "  </w:p>\n" +
                "</w:hdr>";
    }

    private static String generarRelsHeaderConDocumentDocx() {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
                "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>\n" +
                "</Relationships>";
    }

}