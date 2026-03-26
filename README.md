# HistoMatch
HistoMatch: programa que automatiza la evaluación de compatibilidad HLA en los cinco loci clásicos clasificando cada uno en cinco categorías y generando un puntaje global (1–10). 

Contiene 4 archivos: 
1. main.py
2. hla_finder.py (analiza un pdf/word y exporta los HLA a un .csv)
3. histomatch_analisis.py (analiza el .csv y eporta los resultados de compatibilidad a un .docx llamado "Analisis_HLA")
4. histomatch_reporte.py (analiza el .docx y resume la informacion para dejar lo importante en otro .docx llamado "Reporte_HLA"). 


HistoMatch tiene 6 posibilidades de diagnostico de cada loci:

1. Idéntico
Se define cuando hay una coincidencia total y absoluta entre los alelos comparados. Condición: La cantidad de diferencias es cero (len(dif) == 0) y existe al menos un alelo coincidente (len(inter) > 0). Significado: Los perfiles son exactamente iguales en ese locus.

2. Incluido
Es un estado especial de alta resolución donde un conjunto de alelos es un subconjunto casi total del otro. Condición: Se activa cuando la coincidencia es baja (menos del 40%), pero uno de los conjuntos de alelos está contenido en el otro en un 90% o más (cobertura >= 0.9). Significado: Aunque no son idénticos, la información de uno está prácticamente contenida en el otro.

3. Compatible 
Es el estado por defecto cuando no hay conflictos graves pero no se llega a la identidad total. Condición: Se alcanza si no se cumplen las reglas de "Idéntico", "Incierto" o "No compatible". También se asigna si faltan detalles de alta resolución pero los grupos base coinciden. Significado: Existe una compatibilidad aceptable bajo los parámetros del script.

4. Incierto
Se utiliza cuando los datos son ambiguos o la coincidencia es muy baja pero no nula. Condición: Se define si la coincidencia es menor al 40% (porc_iguales < 0.4) pero hay al menos un alelo igual (len(inter) >= 1), y no llega a calificar como "Incluido". Significado: Los datos sugieren una posible compatibilidad, pero los detalles no son suficientes para confirmarla.

5. No compatible
Se define cuando existe una discrepancia clara en la estructura genética. Condición: Si los Grupos Base (la parte antes de los dos puntos :) son diferentes. Si la coincidencia es menor al 40% y no hay ni un solo alelo en común. Significado: Existe un rechazo o falta de concordancia genética en ese locus.

6. Sin Datos / No detectado
Estados de error o falta de información. Condición: Se activa si alguno de los campos de grupo está vacío o si la función de mapeo no encuentra las columnas DQA/DQB. Significado: No se pudo realizar el análisis por falta de parámetros de entrada.
