# MODELO INTEGRAL DE GESTIÓN DE LA DEMANDA Y CAPACIDAD

Modelo cuantitativo en Python para el procesamiento, balance y optimización mensual de la demanda asistencial frente a la capacidad instalada de Talento Humano e Infraestructura física.

---

## 📌 Arquitectura del Pipeline (8 Bloques)

1. **Bloque 1 - Contratos Base:** Estandarización a 27 columnas de Cápita, PGP y Techos con cruce CUPS y expansión cartesiana por sede y población.
2. **Bloque 2 - Programas Especiales:** Integración de notas técnicas programáticas, techos presupuestales y homologación de profesional.
3. **Bloque 3 - Planes Premium:** Prorrateo equitativo de actividades por UEN/sedes y consolidación del maestro general (`Base_General_Consolidada.zip`).
4. **Bloque 4 - Analítica y Resúmenes:** Tablas dinámicas y cubos agregados por CUPS, Ciudad, Centro Médico y Contrato.
5. **Bloque 5 - Oferta Talento Humano:** Estimación de horas y minutos netos mensuales descontando factores reductores y festivos oficiales de Colombia (`holidays.Colombia`).
6. **Bloque 6 - Oferta Infraestructura:** Capacidad física mensual por consultorio según disponibilidad semanal y calendario no festivo de Colombia.
7. **Bloque 7 - Balance Demanda vs Oferta:** Matriz comparativa mensual de minutos requeridos frente a oferta neta disponible.
8. **Bloque 8 - Motor de Optimización y Asignación:** Asignación secuencial mes a mes por prioridad clínica, cobertura, margen y costo; diagnóstico de cuellos de botella (TH vs Infraestructura) y reporte de citas y minutos asignados y no asignados.

---

## 🛠️ Instalación y Dependencias

Clona el repositorio e instala los paquetes requeridos:

```bash
git clone [https://github.com/lauratatianadizcastro-ops/MODELO-GESTION-DEMANDA-1709.git](https://github.com/lauratatianadizcastro-ops/MODELO-GESTION-DEMANDA-1709.git)
cd MODELO-GESTION-DEMANDA-1709
pip install -r requirements.txt
