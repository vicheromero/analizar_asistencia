# **Reporte de Asistencia (Google Apps Script)**

Este es un script de Google Apps Script diseñado para automatizar el análisis, procesamiento y generación de reportes de marcaciones de asistencia de empleados.

El script lee los datos de marcaciones (fichajes) desde una hoja de cálculo de Google, los procesa para calcular atrasos, horas de almuerzo y días laborados, y genera dos hojas de reporte con un análisis detallado y un resumen. Además, puede generar borradores de correo electrónico en Gmail para cada empleado con su reporte individual.

## **Características Principales**

* **Menú Personalizado:** Crea un menú "Reportes de Asistencia" en la UI de Google Sheets para un acceso fácil.  
* **Generación de Reporte Detallado:** Crea una hoja llamada Marcaciones Reorganizadas que muestra el estado de cada empleado por cada día del mes.  
* **Análisis Completo:**  
  * Calcula **atrasos** en el ingreso (con tolerancia de 10 min).  
  * Calcula el **tiempo de almuerzo** (marcando los que superan 1 hora).  
  * Identifica **"Falta marcación"** cuando los registros están incompletos.  
  * Calcula el total de **días laborados**.  
* **Manejo de Casos Especiales:** Identifica y etiqueta automáticamente:  
  * Día de descanso  
  * Feriado  
  * Permiso  
  * Comisión  
  * Compensación  
  * Horas extras (cuando se detectan marcaciones en días no laborables).  
* **Formato Condicional:** Aplica un formato de color automático en el reporte para una fácil identificación visual de problemas (rojo para faltas, amarillo/rojo para atrasos, magenta para compensación, cian para comisión, etc.).  
* **Generación de Resumen:** Crea una hoja Resumen de Asistencia que totaliza los días laborados por empleado.  
* **Notificación por Correo:** Incluye una función para generar borradores de correo en Gmail para cada empleado (listado en la hoja PERSONAL), adjuntando su resumen y detalle de marcaciones en formato de tabla HTML.

## **Requisitos de Configuración**

Para que el script funcione correctamente, tu hoja de cálculo de Google **DEBE** contener las siguientes hojas con estos nombres y estructuras:

1. **Marcaciones**  
   * La hoja con los datos de fichaje en crudo.  
   * Columnas requeridas: Nombre y Apellido, Fecha, Tipo de Registro (ej. "Ingreso", "Salida", "Inicio descanso", "Fin descanso"), Hora.  
2. **Turnos**  
   * Hoja para definir los días laborables de cada empleado.  
   * Columnas requeridas:  
     * Col A: Nombre y Apellido  
     * Cols B-H: lunes, martes, miercoles, jueves, viernes, sabado, domingo.  
     * Usar un 1 para día laborable y 0 para día de descanso.  
3. **Ausentismo**  
   * Hoja para registrar permisos, vacaciones, comisiones, etc.  
   * Columnas requeridas:  
     * Nombre Empleado  
     * Inicio de validez (Fecha de inicio)  
     * Fin de validez (Fecha de fin)  
     * Días de absentismo (Número de días)  
     * **Columna F (Tipo de Absentismo)**: ¡Importante\! El script lee esta columna.  
       * Si el texto es "Comisión", se usará "Comisión".  
       * Si el texto es "Compensación", se usará "Compensación".  
       * Cualquier otro texto (ej. "Vacaciones", "Permiso Médico", o celda vacía) será tratado como "Permiso".  
4. **Feriado**  
   * Una lista simple de días feriados.  
   * El script solo lee la **Columna A**. Cada celda en la Columna A debe contener una fecha que se considerará feriado.  
5. **PERSONAL**  
   * Hoja requerida para la función "Generar correos".  
   * Columnas requeridas:  
     * **Columna B**: Nombre y Apellido (debe coincidir con el nombre en Marcaciones).  
     * **Columna F**: Correo (la dirección de email del empleado).

## **Instalación**

1. Abre tu hoja de cálculo de Google.  
2. Ve a Extensiones \> Apps Script.  
3. Borra cualquier código existente en el editor (Code.gs).  
4. Copia todo el contenido del archivo ReporteAsistencia.gs y pégalo en el editor de Apps Script.  
5. Haz clic en el ícono de **Guardar** 💾.  
6. La primera vez que ejecutes una función (o al recargar la hoja), Google te pedirá permisos. Debes autorizar el script para que pueda modificar la hoja de cálculo (SpreadsheetApp) y generar borradores de correo (GmailApp).

## **Modo de Uso**

1. Asegúrate de que todas las hojas de requisitos (ver arriba) estén creadas y con datos.  
2. Recarga tu hoja de cálculo de Google.  
3. Aparecerá un nuevo menú llamado **"Reportes de Asistencia"**.  
4. **Paso 1:** Haz clic en Reportes de Asistencia \> Generar Reporte.  
   * El script se ejecutará (puede tardar unos segundos) y creará/actualizará las hojas Marcaciones Reorganizadas y Resumen de Asistencia.  
5. **Paso 2:** (Opcional) Haz clic en Reportes de Asistencia \> Generar correos.  
   * El script generará los borradores de correo en tu cuenta de Gmail. Revisa tu carpeta de "Borradores" en Gmail para enviarlos.

## **Vistas Previas**

**Hoja de entrada Marcaciones (Ejemplo):**

![image](https://github.com/user-attachments/assets/584dc460-9793-4211-8bdf-3bb1180ab614)

**Hoja de salida Marcaciones Reorganizadas (Ejemplo con formato):**

<img width="1291" height="423" alt="image" src="https://github.com/user-attachments/assets/e985052e-98d5-48fb-b29f-ec434966eb04" />


**Hoja de salida Resumen de Asistencia (Ejemplo):**

![image](https://github.com/user-attachments/assets/4b3e0ea3-a095-405d-bf47-e070580019c1)
