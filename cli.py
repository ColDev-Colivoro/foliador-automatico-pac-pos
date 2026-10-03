"""
Interfaz de Línea de Comandos (CLI) para Foliador Automático PAC / POS.
Permite consultar folios, generar tandas y ajustar números base.
"""

import sys
import os
import argparse
import subprocess

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from core.folio_engine import FolioEngine


def mostrar_tabla_folios(engine):
    print("\n" + "=" * 78)
    print(f"{'CÓDIGO':<12} | {'CATEGORÍA':<10} | {'ÚLTIMO FOLIO':<14} | {'PLANILLA'}")
    print("-" * 78)
    for p in engine.plantillas:
        cod = p["codigo"]
        cat = p["categoria"]
        dig = p.get("digitos", 5)
        ultimo = engine.get_ultimo_folio(cod)
        ultimo_str = f"{ultimo:0{dig}d}"
        nom = p["nombre"]
        print(f"{cod:<12} | {cat:<10} | {ultimo_str:<14} | {nom[:38]}")
    print("=" * 78 + "\n")


def abrir_carpeta(ruta):
    if sys.platform == "win32":
        os.startfile(ruta)
    elif sys.platform == "darwin":
        subprocess.run(["open", ruta])
    else:
        subprocess.run(["xdg-open", ruta])


def modo_interactivo(engine):
    while True:
        print("\n" + "=" * 50)
        print("   FOLIADOR AUTOMÁTICO PAC / POS - COLDEV")
        print("=" * 50)
        print("1. Ver estado actual de folios")
        print("2. Generar planillas foliadas")
        print("3. Ajustar / Sincronizar último folio con planillero")
        print("4. Abrir carpeta de planillas generadas")
        print("5. Salir")
        
        opc = input("\nSeleccione una opción (1-5): ").strip()
        
        if opc == "1":
            mostrar_tabla_folios(engine)

        elif opc == "2":
            mostrar_tabla_folios(engine)
            codigo = input("Ingrese el CÓDIGO de la planilla (ej: PAC_PCC1): ").strip().upper()
            info = engine.get_plantilla_info(codigo)
            if not info:
                print(f"❌ Error: El código '{codigo}' no existe.")
                continue

            ultimo = engine.get_ultimo_folio(codigo)
            dig = info.get("digitos", 5)
            print(f"\nPlanilla: {info['nombre']}")
            print(f"Último folio emitido: {ultimo:0{dig}d}")
            print(f"El próximo folio será: {ultimo + 1:0{dig}d}")

            cant_str = input("¿Cuántas planillas desea generar? [Default: 1]: ").strip()
            cant = int(cant_str) if cant_str.isdigit() and int(cant_str) > 0 else 1

            sublote = input("Nombre de lote/carpeta opcional (Enter para automático por fecha): ").strip()
            sublote = sublote if sublote else None

            try:
                res = engine.generar_planillas(codigo, cant, carpeta_sublote=sublote)
                print("\n✅ ¡Planillas generadas con éxito!")
                print(f"   Rango emitido: {res['folio_inicial']} al {res['folio_final']} ({res['cantidad']} hojas)")
                print(f"   Carpeta: {res['carpeta_salida']}")
                
                abrir = input("\n¿Desea abrir la carpeta ahora para imprimirlas? (s/n): ").strip().lower()
                if abrir in ["s", "si", "y", "yes"]:
                    abrir_carpeta(res["carpeta_salida"])
            except Exception as e:
                print(f"❌ Error al generar: {e}")

        elif opc == "3":
            mostrar_tabla_folios(engine)
            codigo = input("Ingrese el CÓDIGO a ajustar (ej: PAC_PCC1): ").strip().upper()
            info = engine.get_plantilla_info(codigo)
            if not info:
                print(f"❌ Error: El código '{codigo}' no existe.")
                continue

            ultimo = engine.get_ultimo_folio(codigo)
            print(f"Folio actual en sistema: {ultimo}")
            nuevo_str = input("Ingrese el NUEVO último folio (ej: el que ya se usó en físico): ").strip()
            if not nuevo_str.isdigit():
                print("❌ Error: Debe ser un número entero.")
                continue

            engine.set_ultimo_folio(codigo, int(nuevo_str), "Ajuste manual desde CLI")
            print(f"✅ Folio actualizado. El próximo generado será: {int(nuevo_str) + 1}")

        elif opc == "4":
            abrir_carpeta(engine.salida_dir)

        elif opc == "5":
            print("\nHasta luego.")
            break
        else:
            print("Opción inválida.")


def main():
    parser = argparse.ArgumentParser(description="Foliador Automático PAC / POS")
    subparsers = parser.add_subparsers(dest="comando")

    # Subcomando list
    subparsers.add_parser("list", help="Muestra el estado de todas las planillas")

    # Subcomando gen
    gen_parser = subparsers.add_parser("gen", help="Genera planillas foliadas")
    gen_parser.add_argument("codigo", type=str, help="Código de la planilla (ej: PAC_PCC1)")
    gen_parser.add_argument("cantidad", type=int, nargs="?", default=1, help="Cantidad de hojas a emitir")
    gen_parser.add_argument("--lote", type=str, default=None, help="Nombre de carpeta de lote opcional")

    # Subcomando set
    set_parser = subparsers.add_parser("set", help="Ajusta el último folio registrado")
    set_parser.add_argument("codigo", type=str, help="Código de la planilla")
    set_parser.add_argument("numero", type=int, help="Número de folio base")

    # Subcomando open
    subparsers.add_parser("open", help="Abre la carpeta de salida")

    args = parser.parse_args()
    engine = FolioEngine()

    if not args.comando:
        modo_interactivo(engine)
        return

    if args.comando == "list":
        mostrar_tabla_folios(engine)

    elif args.comando == "gen":
        codigo = args.codigo.upper()
        res = engine.generar_planillas(codigo, args.cantidad, carpeta_sublote=args.lote)
        print(f"✅ Generadas {res['cantidad']} planillas [{res['folio_inicial']} a {res['folio_final']}] en: {res['carpeta_salida']}")

    elif args.comando == "set":
        codigo = args.codigo.upper()
        engine.set_ultimo_folio(codigo, args.numero, "Ajuste por parámetro CLI")
        print(f"✅ Folio de {codigo} fijado en {args.numero}. Próximo será {args.numero + 1}.")

    elif args.comando == "open":
        abrir_carpeta(engine.salida_dir)


if __name__ == "__main__":
    main()
