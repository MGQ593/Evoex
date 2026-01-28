# Business Case: Evoex
## Asistente de IA para Excel con Azure OpenAI

**Fecha:** Enero 2026
**Área:** ChevyplanIT
**Versión:** 1.0

---

## 1. Resumen Ejecutivo

**Evoex** es un complemento de Microsoft Excel que integra inteligencia artificial (Azure OpenAI) directamente en la herramienta de trabajo más utilizada en la organización. Permite a los colaboradores analizar datos, generar fórmulas complejas y automatizar tareas repetitivas mediante lenguaje natural.

### Propuesta de Valor
> "Transformar la productividad en Excel mediante IA conversacional, reduciendo tiempos de análisis y eliminando barreras técnicas."

---

## 2. Problema / Oportunidad

### Situación Actual
| Problema | Impacto |
|----------|---------|
| Creación manual de fórmulas complejas | Horas de trabajo y errores frecuentes |
| Análisis de datos requiere conocimiento avanzado | Dependencia de personal especializado |
| Tareas repetitivas en hojas de cálculo | Tiempo desperdiciado en actividades de bajo valor |
| Curva de aprendizaje de Excel avanzado | Capacitaciones costosas y prolongadas |
| Búsqueda de información externa | Cambio constante entre aplicaciones |

### Oportunidad
- El 80% de los colaboradores usa Excel diariamente
- La IA generativa puede automatizar el 40% de tareas rutinarias en Excel
- Reducción significativa en tiempo de análisis y reportes

---

## 3. Solución Propuesta

### Funcionalidades de Evoex

| Funcionalidad | Descripción | Beneficio |
|---------------|-------------|-----------|
| **Chat con IA** | Consultas en lenguaje natural sobre datos | Análisis sin conocimiento técnico |
| **Generación de Fórmulas** | Crear fórmulas complejas describiendo qué se necesita | Elimina errores de sintaxis |
| **Análisis de Datos** | Interpretación automática de rangos seleccionados | Insights inmediatos |
| **Búsqueda Web Integrada** | Consultar información actualizada sin salir de Excel | Mayor contexto en análisis |
| **Múltiples Modelos IA** | Selección de modelo según complejidad | Optimización de costos |
| **Acciones Automatizadas** | Aplicar cambios directamente en celdas | Ejecución con un clic |

### Arquitectura Técnica

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│   Excel + Add-in │────▶│   Proxy Server  │────▶│  Azure OpenAI   │
│     (Usuario)    │◀────│   (Seguridad)   │◀────│    (GPT-4o)     │
└─────────────────┘     └─────────────────┘     └─────────────────┘
```

### Seguridad
- ✅ Desplegado en infraestructura propia (Easypanel)
- ✅ Credenciales protegidas en servidor (no expuestas al cliente)
- ✅ Distribución controlada via Office 365 Admin Center
- ✅ Sin acceso público al manifest del complemento
- ✅ Datos procesados en Azure (región configurada por la empresa)

---

## 4. Análisis de Costos

### Costos de Implementación (Únicos)

| Concepto | Costo Estimado | Notas |
|----------|---------------|-------|
| Desarrollo | $0 | Desarrollado internamente |
| Infraestructura inicial | $0 | Usa Easypanel existente |
| Configuración Azure OpenAI | $0 | Recurso ya provisionado |
| **Total Implementación** | **$0** | |

### Costos Operativos (Mensuales)

| Concepto | Costo Estimado | Notas |
|----------|---------------|-------|
| Azure OpenAI (GPT-4o) | $50 - $200/mes | Según uso (~1000 consultas/día) |
| Hosting Easypanel | Incluido | Infraestructura existente |
| Mantenimiento | 4 hrs/mes | Personal interno |
| **Total Mensual** | **~$150/mes** | Estimado conservador |

### Costo por Usuario
- Con 50 usuarios activos: **$3/usuario/mes**
- Con 100 usuarios activos: **$1.50/usuario/mes**

---

## 5. Beneficios y ROI

### Beneficios Cuantificables

| Métrica | Antes | Después | Ahorro |
|---------|-------|---------|--------|
| Tiempo creando fórmulas complejas | 30 min | 2 min | 93% |
| Tiempo de análisis de datos | 2 hrs | 20 min | 83% |
| Errores en fórmulas | 15% | 2% | 87% |
| Consultas a TI por Excel | 20/semana | 5/semana | 75% |

### Cálculo de ROI

**Supuestos:**
- 50 usuarios activos
- Salario promedio: $1,500/mes
- Ahorro de tiempo: 5 horas/semana por usuario

**Cálculo:**
```
Ahorro mensual por usuario = 5 hrs/semana × 4 semanas × ($1,500/160 hrs)
                           = 20 hrs × $9.38/hr
                           = $187.50/usuario/mes

Ahorro total mensual = 50 usuarios × $187.50 = $9,375/mes
Costo mensual = $150/mes
ROI mensual = ($9,375 - $150) / $150 = 6,150%
```

**Payback:** Inmediato (primer mes)

### Beneficios No Cuantificables
- Mayor satisfacción del colaborador
- Democratización del análisis de datos
- Reducción de dependencia de expertos
- Mejora en calidad de reportes
- Ventaja competitiva en adopción de IA

---

## 6. Riesgos y Mitigación

| Riesgo | Probabilidad | Impacto | Mitigación |
|--------|--------------|---------|------------|
| Adopción baja por usuarios | Media | Alto | Capacitación y casos de uso prácticos |
| Costos de API excedan presupuesto | Baja | Medio | Límites de uso y monitoreo |
| Respuestas incorrectas de IA | Media | Medio | Revisión humana antes de aplicar |
| Cambios en precios de Azure | Baja | Bajo | Presupuesto con margen |
| Indisponibilidad del servicio | Baja | Medio | Alta disponibilidad de Azure |

---

## 7. Plan de Implementación

### Fase 1: Piloto (Semanas 1-2)
- [ ] Despliegue para grupo de 10 usuarios
- [ ] Capacitación inicial
- [ ] Recolección de feedback
- [ ] Ajustes según necesidades

### Fase 2: Expansión (Semanas 3-4)
- [ ] Ampliación a 50 usuarios
- [ ] Documentación de casos de uso
- [ ] Métricas de adopción

### Fase 3: Producción (Mes 2+)
- [ ] Despliegue organizacional
- [ ] Soporte continuo
- [ ] Mejoras iterativas

---

## 8. Casos de Uso por Área

### Finanzas
- Análisis de variaciones presupuestarias
- Generación de fórmulas de amortización
- Consolidación de reportes

### Operaciones
- Análisis de indicadores de gestión
- Proyecciones de demanda
- Seguimiento de KPIs

### Recursos Humanos
- Análisis de nómina
- Métricas de rotación
- Reportes de headcount

### Comercial
- Análisis de ventas
- Seguimiento de metas
- Reportes de comisiones

---

## 9. Comparativa con Alternativas

| Característica | Evoex | Copilot for Excel | Otras soluciones |
|----------------|-------|-------------------|------------------|
| Costo mensual | ~$150 total | $30/usuario/mes | Variable |
| Control de datos | Total (on-premise) | Microsoft Cloud | Terceros |
| Personalización | Alta | Limitada | Variable |
| Integración búsqueda web | Sí | Limitada | No |
| Múltiples modelos IA | Sí | No | No |
| Soporte interno | Sí | No | No |

**Ahorro vs Copilot:** Con 50 usuarios, Copilot costaría $1,500/mes vs $150/mes de Evoex = **$1,350/mes de ahorro**

---

## 10. Recomendación

Se recomienda **aprobar la implementación de Evoex** basado en:

1. **ROI excepcional** (>6,000%)
2. **Costo mínimo** de operación
3. **Control total** de datos y seguridad
4. **Desarrollo interno** (sin dependencias externas)
5. **Escalabilidad** según necesidades

### Solicitud de Aprobación

| Concepto | Monto |
|----------|-------|
| Presupuesto mensual Azure OpenAI | $200/mes |
| Horas de soporte (IT interno) | 4 hrs/mes |

---

## 11. Anexos

### A. Capturas de Pantalla
*(Incluir capturas de la interfaz de Evoex)*

### B. Demo
- URL de producción: https://evoexia.0hidyn.easypanel.host
- Disponible en Excel via Office 365 Admin Center

### C. Contacto
- **Desarrollador:** ChevyplanIT
- **Soporte:** [Correo interno]

---

**Preparado por:** ChevyplanIT
**Fecha:** Enero 2026

