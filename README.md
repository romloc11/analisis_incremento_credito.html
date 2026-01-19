# 📈 Análisis Automatizado de Incremento de Crédito

Sistema desarrollado en **Python** para evaluar el comportamiento histórico
de clientes y definir decisiones automáticas sobre el límite de crédito:
incremento, mantenimiento o decremento.

Este proyecto simula un motor de decisión crediticia basado en reglas,
scoring y métricas de riesgo.

---

## 🧠 Objetivo del Proyecto

- Automatizar decisiones de crédito
- Reducir riesgo financiero
- Estandarizar criterios de evaluación
- Facilitar análisis masivo de clientes

---

## ⚙️ Lógica de Decisión

El sistema evalúa variables como:

- Score interno
- Historial de pagos
- Uso de línea de crédito
- Comportamiento reciente
- Límite actual

Y asigna una decisión:
- ✅ Incremento
- ⚠️ Revisión manual
- ➖ Mantener
- ❌ Decremento

---

## 🛠️ Tecnologías

- Python
- Pandas
- NumPy
- Excel (input/output)
- Scoring rules

---

## 📊 Ejemplo Visual

![Preview](images/foto_incremento.png)

---

## 🚀 Uso

```bash
pip install -r requirements.txt
python analisis_incremento_credito.py
