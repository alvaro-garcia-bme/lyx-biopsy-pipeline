# 🔬 Lyx Urología — AI Clinical Data Pipeline

> **Proof of Concept (PoC)** desarrollada para la Unidad de Inteligencia Clínica & Automatización de **Lyx Urología** (Director Médico: Dr. Juan I. Martínez-Salamanca).

---

## 📋 Descripción del Proyecto
Plataforma de extracción estructurada, normalización y estratificación automatizada de informes de patología y biopsias prostáticas mediante Modelos de Lenguaje Médico (LLMs).

### 🚀 Funcionalidades Principales:
* **Extracción Inteligente Multi-Formato:** Procesa informes estructurados en tablas, checklists y dictados médicos en prosa narrativa continua.
* **Normalización Oncológica EAU:** Clasificación automática del riesgo según las guías de la *European Association of Urology* (EAU) y grupos pronósticos ISUP (1 a 5).
* **Consolidación Relacional & Exportación Excel:** Generación instantánea de base de datos clínica unificada descargable en formato Excel nativo (`.xlsx`) con estilos corporativos.
* **Privacidad & Cumplimiento RGPD:** Diseñado con arquitectura de anonimización local en memoria y compatibilidad con inferencia *on-premise* sin salida de datos clínicos a la nube.

---

## 🛠️ Tecnologías Utilizadas
* **Lenguaje:** Python 3.11+
* **Framework Web:** Streamlit
* **Motor de Inferencia:** Groq API / LLM Llama 3 & GPT-OSS
* **Procesamiento de Documentos:** PyPDF, ReportLab, OpenPyXL
* **Visualización de Datos:** Plotly Express

---

## 👨‍💻 Autor
**Álvaro García Casas**  
*Graduando en Ingeniería Biomédica (Universidad Carlos III de Madrid / Georgia Tech)*  
*Email:* alvaro.garcia.bme@gmail.com | *LinkedIn:* [alvaro-garcia-casas](https://www.linkedin.com/in/alvaro-garcia-casas)
