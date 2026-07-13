# DCD 1.2.2 estable - corrección de despliegue Streamlit Cloud

Este paquete no modifica la lógica funcional de DCD.

Cambio realizado:
- Se fija requirements.txt con versiones estables para evitar que Streamlit Cloud instale automáticamente versiones demasiado recientes de Python/librerías.

Instrucciones de despliegue recomendadas:
1. Subir este paquete a la rama actualmente usada para el test.
2. En Streamlit Community Cloud, crear una app nueva o redeplegar seleccionando Python 3.12 en Advanced settings.
3. No usar Python 3.14 para esta fase de test real.
4. Copiar los mismos Secrets de la app anterior.
5. Reboot app.

Motivo:
El log de Streamlit muestra que el entorno se creó con Python 3.14.6 y paquetes sin fijar, y el proceso terminó con Segmentation fault antes de llegar a una excepción Python controlable.
