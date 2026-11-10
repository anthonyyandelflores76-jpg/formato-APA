# Convertidor a formato APA 7

Sube un documento de Word (.docx) y descarga una copia con formato APA 7:

- Times New Roman 12 (reemplaza letra manuscrita o decorativa), interlineado 1,5, texto justificado
- Márgenes de 2,5 cm por todos los lados
- Portada APA 7 opcional (título, autor, institución, carrera, curso, asesor y fecha)
- Número de página arriba a la derecha
- Títulos nivel 1 a 5 con estilo APA (nivel 1 centrado y en negrita, mayúscula tipo título)
- Sangría de primera línea de 1,27 cm
- Referencias con sangría francesa y orden alfabético
- Tablas sin líneas verticales (también convierte tablas escritas con barras `|`)
- Citas en bloque (40 palabras o más) con sangría
- Elimina párrafos vacíos

## Dos modos

- **Completo (por defecto):** para monografías, ensayos y trabajos. Aplica todo lo anterior, borra párrafos vacíos y da estilo APA a los títulos.
- **Conservar estructura** (casilla en la web, o `--conservar` en terminal): para formularios e informes con formato propio (por ejemplo, el Informe de Práctica de SENATI). Mantiene espacios, carátula, títulos y bordes de tablas, y aplica fuente, interlineado, justificado, márgenes, número de página, sangría y tablas alineadas a la izquierda.

## Instalación

1. Instala [Python 3.9 o superior](https://www.python.org/downloads/) (marca "Add Python to PATH").
2. Instala las dependencias:

```
pip install -r requirements.txt
```

## Uso (versión web, con contraseña)

```
python app.py
```

Abre http://127.0.0.1:5000, escribe la contraseña, sube tu .docx, llena la portada si la necesitas y descarga el resultado.

La contraseña por defecto es `consola`. Para cambiarla sin tocar el código:

```
set APA_PASSWORD=mi_nueva_clave        (Windows, cmd)
$env:APA_PASSWORD="mi_nueva_clave"     (Windows, PowerShell)
```

## Uso (terminal, sin portada ni contraseña)

```
python convertir_apa.py mi_trabajo.docx
python convertir_apa.py mi_trabajo.docx --conservar
```

## Limitaciones

- Solo funciona con .docx (si es .doc, guárdalo como .docx desde Word).
- Detecta títulos por el estilo de Word (Título 1, 2, 3...) y por nombres comunes (Introducción, Conclusiones, Referencias...).
- No convierte imágenes de texto manuscrito ni redacta las referencias: solo las ordena y les da formato.
