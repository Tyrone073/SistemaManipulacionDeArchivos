import java.util.Date;

public class Contenido {
    Date fechaActual;
    String mensaje;

    public Contenido(String mensaje) {
        this.fechaActual = new Date();
        this.mensaje = mensaje;
    }

    public Date getFechaActual() {
        return fechaActual;
    }
    public String getMensaje() {
        return mensaje;
    }
    public void setMensaje(String mensaje) {
        this.mensaje = mensaje;
    }
    public void setFechaActual(Date fechaActual) {
        this.fechaActual = fechaActual;
    }
}
