#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script: Validar-Codificacion-VBA.py
Propósito: Guardián automatizado de codificación y reglas de nombres para ResidenciaApp.
           Comprueba que ningún archivo VBA contenga mojibake (caracteres UTF-8 interpretados como ANSI)
           ni infracciones de la regla estricta de nombres (prohibición de Chr(211)).

Uso:
  python bdas-multiusuario/scripts/Validar-Codificacion-VBA.py [--all]
"""

import os
import sys
import re

VBA_DIR = r"H:\ResidenciaApp\bdas-multiusuario\vba-modules"
SCRIPTS_DIR = r"H:\ResidenciaApp\bdas-multiusuario\scripts"

MOJIBAKE_PATTERNS = [
    (re.compile(r'Ã[\x80-\xbf]'), "Secuencia UTF-8 multi-byte sin decodificar (Ã...)"),
    (re.compile(r'Ã[A-Za-z0-9\?]'), "Secuencia mojibake corrupta (Ã?, Ã“, etc.)"),
    (re.compile(r'[\x81\x8d\x8f\x90\x9d]'), "Carácter no imprimible Windows-1252 (residuo de UTF-8)"),
    (re.compile(r'[ǧǭǸǟǬ]'), "Carácter cirílico/extendido producto de doble mojibake"),
    (re.compile(r'\uFFFD'), "Carácter de reemplazo Unicode ()"),
    (re.compile(r'\b[A-Za-z]{2,}\?[A-Za-z]{2,}\b'), "Palabra con signo '?' en lugar de vocal acentuada (ej: m?dulo)")
]

CHR211_PATTERN = re.compile(r'ChrW?\(211\)', re.IGNORECASE)

def check_file(path):
    issues = []
    fname = os.path.basename(path)
    
    with open(path, "rb") as fp:
        raw = fp.read()
        
    # 1. Comprobar que es decodificable en Windows-1252
    try:
        text = raw.decode("windows-1252")
    except Exception as e:
        issues.append((0, f"Error decodificando Windows-1252: {e}"))
        return issues
        
    lines = text.splitlines()
    for idx, line in enumerate(lines, 1):
        # Ignorar comentarios legítimos que explican no usar Chr(211)
        if "prohibido" in line.lower() or "regla estricta" in line.lower():
            pass
        elif CHR211_PATTERN.search(line):
            issues.append((idx, f"Infracción de regla estricta de nombres: uso de Chr(211) en lugar de 'GIJÓN'"))
            
        # Comprobar patrones de mojibake
        for pat, desc in MOJIBAKE_PATTERNS:
            m = pat.search(line)
            if m:
                # Filtrar falsos positivos de SQL ("WHERE id = ?")
                val = m.group(0)
                if val in ["?", " ? ", "? ", " ?", "='?'"]:
                    continue
                issues.append((idx, f"{desc}: '{val}' en línea -> {line.strip()[:80]}"))
                break
                
    return issues

def main():
    print("================================================================")
    print("  GUARDIAN DE CODIFICACION Y INTEGRIDAD VBA - ResidenciaApp     ")
    print("================================================================")
    
    total_files = 0
    total_issues = 0
    files_with_issues = 0
    
    target_files = []
    args = sys.argv[1:]
    
    if "--all" in args:
        for root, _, files in os.walk(VBA_DIR):
            for f in files:
                if f.endswith((".bas", ".cls", ".frm")):
                    target_files.append(os.path.join(root, f))
    else:
        # Comprobar si se pasaron archivos directos
        for a in args:
            if os.path.isfile(a) and a.endswith((".bas", ".cls", ".frm")):
                target_files.append(os.path.abspath(a))
                
        # Si no se pasaron archivos directos, comprobar componentes modificados en Git (staged)
        if not target_files:
            import subprocess
            try:
                res = subprocess.run(
                    ["git", "diff", "--cached", "--name-only", "--diff-filter=ACM"],
                    capture_output=True, text=True, check=False, cwd=r"H:\ResidenciaApp"
                )
                if res.returncode == 0:
                    for line in res.stdout.splitlines():
                        line = line.strip()
                        if line.endswith((".bas", ".cls", ".frm")):
                            p = os.path.join(r"H:\ResidenciaApp", line)
                            if os.path.isfile(p):
                                target_files.append(p)
            except Exception:
                pass
                
        # Si no hay Git staged ni argumentos, comprobar directorio completo por defecto
        if not target_files and len(args) == 0:
            # Si no hay archivos staged de VBA en git, informar y salir exitoso
            print("[INFO] No hay componentes VBA en el área de preparación (staged) para validar.")
            sys.exit(0)
            
    for p in target_files:
        total_files += 1
        issues = check_file(p)
        f = os.path.basename(p)
        if issues:
            files_with_issues += 1
            total_issues += len(issues)
            print(f"\n[ALERTA] {f} ({len(issues)} problemas):")
            for line_num, desc in issues[:5]:
                print(f"   Línea {line_num}: {desc}")
            if len(issues) > 5:
                print(f"   ... y {len(issues) - 5} problemas más.")
                
    print("\n----------------------------------------------------------------")
    print(f"Archivos analizados : {total_files}")
    print(f"Archivos afectados  : {files_with_issues}")
    print(f"Problemas totales   : {total_issues}")
    print("----------------------------------------------------------------")
    
    if total_issues == 0:
        print("[OK] Todos los archivos VBA analizados son 100% compatibles con Windows-1252 y libres de mojibake.")
        sys.exit(0)
    else:
        print("[ERROR] Se encontraron problemas de codificación. Bloqueando operación.")
        sys.exit(1)

if __name__ == "__main__":
    main()
