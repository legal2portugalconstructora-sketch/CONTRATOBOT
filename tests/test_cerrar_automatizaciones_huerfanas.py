# -*- coding: utf-8 -*-
"""Prueba de word_com_bridge._cerrar_automatizaciones_huerfanas: cierra
procesos EXCEL.EXE/WINWORD.EXE lanzados vía COM (con la bandera
'/automation' en su línea de comando) que quedaron huérfanos de un intento
anterior, SIN tocar nunca una ventana que el usuario haya abierto a mano.

Esto importa porque Excel/Word se registran como servidores COM
'multi-use': si ya hay un proceso corriendo (aunque sea un huérfano
trabado, ej. mostrando un diálogo bloqueado que nadie puede cerrar porque
la automatización es headless), Dispatch("Excel.Application") se ENGANCHA
a ESE MISMO proceso en vez de abrir uno nuevo -- heredando lo que sea que
lo tenga trabado. Se vio en la práctica: un huérfano así hacía fallar
hasta 'excel.Visible = False' con un AttributeError ("Property
'Excel.Application.Visible' can not be set"), sin que ningún reintento
ayudara (seguía siendo el mismo objeto roto). Por eso hay que sacar el
huérfano de en medio ANTES de abrir uno nuevo.

Requiere Windows + Excel instalado (abre un Excel de verdad, vía COM y
vía línea de comandos, para probar los dos casos). NO requiere Word para
esta prueba puntual.

Corré con: py tests\\test_cerrar_automatizaciones_huerfanas.py
"""
import json
import os
import subprocess
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from word_com_bridge import _cerrar_automatizaciones_huerfanas, _en_windows

RUTA_EXCEL = r"C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE"


def _excel_corriendo():
    resultado = subprocess.run(
        ["powershell", "-NoProfile", "-Command",
         "Get-CimInstance Win32_Process -Filter \"Name='EXCEL.EXE'\" "
         "| Select-Object ProcessId,CommandLine | ConvertTo-Json -Compress"],
        capture_output=True, text=True,
    )
    salida = resultado.stdout.strip()
    if not salida:
        return []
    procesos = json.loads(salida)
    return [procesos] if isinstance(procesos, dict) else procesos


if __name__ == "__main__":
    if not _en_windows() or not os.path.exists(RUTA_EXCEL):
        print("AVISO: no es Windows o no se encontró Excel en la ruta esperada -- se salta la prueba.")
        sys.exit(0)

    # --- Caso 1: un huérfano vía COM (con '/automation') SÍ se cierra. ---
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    excel_huerfano = win32.Dispatch("Excel.Application")
    excel_huerfano.Visible = False
    time.sleep(1)
    try:
        assert any("/automation" in (p.get("CommandLine") or "").lower() for p in _excel_corriendo()), (
            "no se pudo crear el huérfano de prueba -- ¿cambió la bandera que usa Windows?"
        )
        _cerrar_automatizaciones_huerfanas()
        time.sleep(1)
        assert not _excel_corriendo(), "FALLO: el huérfano vía COM seguía vivo después de la limpieza"
        print("Huérfano vía COM ('/automation'): se cierra -- OK")
    finally:
        pythoncom.CoUninitialize()
        for p in _excel_corriendo():
            try:
                subprocess.run(["taskkill", "/PID", str(p["ProcessId"]), "/F"], capture_output=True)
            except Exception:
                pass

    # --- Caso 2: una ventana abierta "a mano" (sin '/automation') NUNCA se toca. ---
    proceso_manual = subprocess.Popen([RUTA_EXCEL])
    try:
        time.sleep(4)
        procesos_antes = _excel_corriendo()
        assert procesos_antes, "no se pudo abrir Excel 'a mano' para la prueba"
        assert not any("/automation" in (p.get("CommandLine") or "").lower() for p in procesos_antes), (
            "el Excel abierto 'a mano' no debería tener la bandera '/automation'"
        )
        _cerrar_automatizaciones_huerfanas()
        time.sleep(1)
        assert _excel_corriendo(), "FALLO: cerró una ventana de Excel abierta 'a mano' -- esto NUNCA debería pasar"
        print("Ventana abierta 'a mano' (sin '/automation'): NO se toca -- OK")
    finally:
        try:
            subprocess.run(["taskkill", "/PID", str(proceso_manual.pid), "/F"], capture_output=True)
        except Exception:
            pass

    print("TODO OK")
