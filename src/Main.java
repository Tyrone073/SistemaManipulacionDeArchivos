import java.io.*;
import java.util.Scanner;

public class Main {
    static Scanner scanner = new Scanner(System.in);

    public static void main(String[] args) throws IOException {
        int opcion;
        do {
            System.out.println("\n====Sistema de procesamiento de archivos====");
            System.out.println("1. Crear archivo");
            System.out.println("2. Leer archivo");
            System.out.println("3. Contar palabras de un archivo");
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
                    ContarPalabrasArchivoYPasarloATxt();
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
                "\n5. svg");

        int tipo = System.in.read();
        scanner.nextLine();
        String extension = switch ((char)tipo) {
            case '1' -> "txt";
            case '2' -> "docx";
            case '3' -> "pptx";
            case '4' -> "xlsx";
            case '5' -> "svg";
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
            }else {

                Contenido contenido = new Contenido("Hola este es un archivo tipo ." + extension + " creado en java");

                switch (extension) {
                    case "txt":
                        PrintWriter out = new PrintWriter(archivo);
                        out.println("Creado el " + contenido.getFechaActual());
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
                    case "svg":
                        ControllerOffice.crearSvg(nombre, contenido);
                        System.out.println("Archivo svg creado exitosamente.");
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
            if (nombre.endsWith(".docx") || nombre.endsWith(".pptx") || nombre.endsWith(".xlsx")){
                String contenidoOffi = ControllerOffice.leerArchivoOffice(nombre);
                System.out.println(contenidoOffi);
                return;
            }else if (nombre.endsWith(".svg") || nombre.endsWith(".txt")){
                String contenidoNormal = ControllerOffice.leerArchivo(nombre);
                System.out.println(contenidoNormal);
            }else {
                System.out.println("El archivo no es compatible.");
            }
        }
    }

    public static void ContarPalabrasArchivoYPasarloATxt() throws IOException {
        System.out.println("--- Listado de archivos ---");
        listarArchivosFiltrados();
        System.out.print("Ingrese el nombre del archivo con su extension: ");
        String nombre = scanner.nextLine();
        File archivo = new File(nombre);
        String nPalabras = "";

        if (archivo.exists()) {
            if (nombre.endsWith(".docx") || nombre.endsWith(".pptx") || nombre.endsWith(".xlsx")){
                nPalabras = "Este archivo tiene "+ ControllerOffice.leerArchivoOffice(nombre).split("\\W+").length + " palabras.";
                System.out.println(nPalabras);

            }else if (nombre.endsWith(".svg") || nombre.endsWith(".txt")){
                nPalabras = "Este archivo tiene "+ ControllerOffice.leerArchivo(nombre).split("\\W+").length + " palabras.";
                System.out.println(nPalabras);

            }else {
                System.out.println("El archivo no es compatible.");
            }
            
            File conteo = new File("PalabrasDe-"+ nombre+"-A.txt");
            System.out.println("Se agrega el conteo de palabras al archivo: "+ conteo);

            try {
                PrintWriter out = new PrintWriter(conteo);
                out.println("Archivo txt de conteo de palabras del archivo: " + nombre);
                out.println(nPalabras);
                out.flush();
                out.close();
                System.out.println("Conteo de palabras agregado exitosamente.");
            } catch (FileNotFoundException e) {
                System.out.println("Error al escribir en el archivo: " + e.getMessage());
            }
        }else {
            System.out.println("El archivo no existe.");
        }
    }

    public static void eliminarArchivo() {
        System.out.println("\n--- Listado de archivos ---");
        listarArchivosFiltrados();
        System.out.print("Ingrese el nombre del archivo con su extension: ");
        String nombre = scanner.nextLine();
        File archivo = new File(nombre);

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
        String[] extensionesPermitidas = {"txt", "docx", "pptx", "xlsx", "svg"};

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
