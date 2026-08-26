# Pruebas de ContratoBot

Scripts sueltos (sin pytest ni otro framework -- solo `assert` + `print`),
pensados para correrse a mano después de tocar la lógica de negocio, antes
de reiniciar el servidor y darle el cambio por bueno.

Cada archivo se corre solo:

```
py tests\test_formatear_monto_letras.py
py tests\test_paso2_lote.py
py tests\test_segundo_titular.py
py tests\test_uuuuuuu_ciudad.py
py tests\test_imagen_en_blanco.py
py tests\test_historial_contratos.py
```

O todos juntos (PowerShell):

```
Get-ChildItem tests\test_*.py | ForEach-Object { py $_.FullName }
```

Todos terminan con `TODO OK` si todo salió bien, o un `AssertionError` con
el mensaje de qué falló.

## Qué NO hay acá

Pruebas que abren Excel/Word de verdad (COM, vía `word_com_bridge.py`) --
requieren Windows con Excel y Word instalados, tardan bastante (a veces
más de un minuto), y abren ventanas reales. No están automatizadas acá a
propósito. Si necesitás probar algo de la automatización COM en concreto
(pegado de imágenes, recálculo de fórmulas, conversión a PDF), pedíselo a
Claude en el chat -- arma un script puntual, lo corre, y te cuenta el
resultado, sin dejarlo como prueba permanente.
