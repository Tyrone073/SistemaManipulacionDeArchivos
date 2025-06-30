import java.io.*;
import java.util.Scanner;

public class Main {
    static Scanner scanner = new Scanner(System.in);

    public static void main(String[] args) throws IOException {
        int opcion;
        do {
            System.out.println("\nSistema de procesamiento de archivos");
            System.out.println("1. Crear archivo");
            System.out.println("2. Leer archivo");
            System.out.println("3. Modificar archivo");
            System.out.println("4. Eliminar archivo");
            System.out.println("0. Salir");
            System.out.print("Ingrese una opción: ");
            opcion = Integer.parseInt(System.console().readLine());

            switch (opcion) {
                case 1:
                    crearArchivo();
                    break;
                case 2:
                    leerArchivo();
                    break;
                case 3:
                    modificarArchivo();
                    break;
                case 4:
                    eliminarArchivo();
                    break;
                case 0:
                    System.out.println("Programa finalizado.");
                    break;
                default:
                    System.out.println("Opción inválida.");
            }
        } while (opcion != 0);
    }

    public static void crearArchivo() throws IOException {
        System.out.println("Seleccione el tipo de archivo:");
        System.out.println("1. txt" +
                "\n2. docx" +
                "\n3. pptx" +
                "\n4. xlsx" +
                "\n5. xml");

        int tipo = System.in.read();
        scanner.nextLine();
        String extension = switch ((char)tipo) {
            case '1' -> "txt";
            case '2' -> "docx";
            case '3' -> "pptx";
            case '4' -> "xlsx";
            case '5' -> "xml";
            default -> null;
        };

        if (extension == null) {
            System.out.println("Extensión no válida.");
            return;
        }

        System.out.print("Ingrese el nombre del archivo: ");
        String nombre = scanner.nextLine();
        File archivo = new File(nombre + "." + extension);
        try{
            if (archivo.exists()) {
                System.out.println("El archivo ya existe.");
                return;
            }else {

                Contenido contenido = new Contenido("Hola este es un archivo tipo ." + extension + " creado en java");

                switch (extension) {
                    case "txt":
                        PrintWriter out = new PrintWriter(archivo);
                        out.println("Archivo creado en la fecha: " + contenido.getFechaActual());
                        out.write(contenido.getMensaje());
                        out.flush();
                        out.close();
                        System.out.println("Archivo txt creado exitosamente.");
                        break;
                    case "docx":
                        ControllerOffice.crearDocx(nombre, contenido);
                        System.out.println("Archivo docx creado exitosamente.");
                        break;
                    case "pptx":
                        ControllerOffice.crearPptx(nombre, contenido);
                        System.out.println("Archivo pptx creado exitosamente.");
                        break;
                    case "xlsx":
                        ControllerOffice.crearXlsx(nombre, contenido);
                        System.out.println("Archivo xlsx creado exitosamente.");
                        break;
                }

                System.out.println("Archivo creado exitosamente.");
                System.out.println("su ubucacion es:"+ archivo.getAbsolutePath());
                return;
            }

        }catch(IOException e){
            System.out.println("Error al crear el archivo: " + e.getMessage());
            return;
        }

    }

    public static void leerArchivo() throws IOException {
        System.out.println("--- Listado de archivos ---");
        listarArchivosFiltrados();
        System.out.print("Ingrese el nombre del archivo con su extension: ");
        String nombre = scanner.nextLine();
        File archivo = new File(nombre);

        if (!archivo.exists()) {
            System.out.println("El archivo no existe.");
        }else {
            System.out.println("\n--- Contenido del archivo ---");
            switch (archivo.getName().substring(archivo.getName().lastIndexOf(".") + 1)){
                case "txt":
                    System.out.println("Archivo txt.");
                    FileReader fileR = new FileReader(nombre);
                    BufferedReader bR = new BufferedReader(fileR);
                    System.out.println(bR.readLine());
                    bR.close();
                    break;
                case "docx":
                    System.out.println("Archivo docx.");
                    String contenidoDocx = ControllerOffice.leerArchivo(nombre);
                    System.out.println(contenidoDocx.replaceAll("<[^>]+>", "")
                    );
                    break;
                case "pptx":
                    System.out.println("Archivo pptx.");
                    break;
            }
        }
    }

    public static void modificarArchivo() {
        System.out.print("Ingrese el nombre del archivo a modificar (sin extensión): ");
        String nombre = scanner.nextLine();
        System.out.print("Ingrese la extensión del archivo (ej: txt): ");
        String extension = scanner.nextLine();
        File archivo = new File(nombre + "." + extension);

        if (!archivo.exists()) {
            System.out.println("El archivo no existe.");
            return;
        }

        System.out.print("Escribe el nuevo mensaje a agregar: ");
        String mensaje = scanner.nextLine();
        Contenido contenido = new Contenido(mensaje);

        try (BufferedWriter bw = new BufferedWriter(new FileWriter(archivo, true))) {
            bw.newLine();
            bw.write(contenido.mensaje);
            System.out.println("Mensaje agregado exitosamente.");
        } catch (IOException e) {
            System.out.println("Error al escribir en el archivo: " + e.getMessage());
        }
    }

    public static void eliminarArchivo() {
        System.out.print("Ingrese el nombre del archivo a eliminar (sin extensión): ");
        String nombre = scanner.nextLine();
        System.out.print("Ingrese la extensión del archivo (ej: txt): ");
        String extension = scanner.nextLine();
        File archivo = new File(nombre + "." + extension);

        if (archivo.exists()) {
            if (archivo.delete()) {
                System.out.println("Archivo eliminado exitosamente.");
            } else {
                System.out.println("No se pudo eliminar el archivo.");
            }
        } else {
            System.out.println("El archivo no existe.");
        }
    }


    public static void listarArchivosFiltrados() {
        File directorio = new File(".");
        String[] extensionesPermitidas = {"txt", "docx", "pptx", "xlsx", "xml"};

        String[] archivos = directorio.list((dir, nombreArchivo) -> {
            for (String ext : extensionesPermitidas) {
                if (nombreArchivo.toLowerCase().endsWith("." + ext)) {
                    return true;
                }
            }
            return false;
        });

        if (archivos != null && archivos.length > 0) {
            System.out.println("Archivos encontrados en el directorio:");
            for (String archivo : archivos) {
                System.out.println(archivo);
            }
        } else {
            System.out.println("No hay archivos en el directorio.");
        }
    }

}
