# -*- coding: utf-8 -*-
"""Processa cada .rvt de MODELOS ATIVOS chamando o pyRevit CLI.
Copia para pasta local ASCII quando o nome tem caracteres problematicos
(o pyRevit CLI nao abre modelos com path non-ASCII em drives de rede).

Sintaxe correta do pyRevit CLI (confirmada via `pyrevit run --help`):
    pyrevit run <script.py> <modelo.rvt> --revit=<ano> --purge --allowdialogs
"""
import os, sys, time, shutil, ctypes, subprocess

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from bim_paths import PASTA_MODELOS, PASTA_SAIDA, LOG_BATCH

PYREVIT_CLI = r"C:\Program Files\pyRevit CLI\bin\pyrevit.exe"
SCRIPT = u"X:/06_COMPARTILHADOS DE PROJETOS/ARQUITETURA/DIEGO/PYREVIT SCRIPTS/BIM.extension/BIM.tab/Analise.panel/Extrair.pushbutton/script.py"

REVIT_VERSAO = "2020"
PAUSA = 5
DEBUG_UM_SO = False         

TEMP_ASCII = r"C:\Temp\bim_ascii"


def only_ascii(s):
    return all(ord(c) < 128 for c in s)


def get_short_path(p):
    b = ctypes.create_unicode_buffer(1024)
    if ctypes.windll.kernel32.GetShortPathNameW(p, b, 1024):
        v = b.value
        if v and only_ascii(v) and v != p:
            return v
    return None


def sanitize(name):
    mapa = {u"Ô":u"O", u"Ã":u"A", u"Á":u"A", u"À":u"A", u"Â":u"A", u"Ä":u"A",
            u"É":u"E", u"Ê":u"E", u"È":u"E",
            u"Í":u"I", u"Î":u"I",
            u"Ó":u"O", u"Õ":u"O", u"Ô":u"O",
            u"Ú":u"U", u"Û":u"U",
            u"Ç":u"C", u"Ñ":u"N"}
    for k, v in mapa.items():
        name = name.replace(k, v)
    return u"".join(c if ord(c) < 128 else u"_" for c in name)


def copiar_ascii(origem, nome_original):
    if not os.path.isdir(TEMP_ASCII):
        os.makedirs(TEMP_ASCII)
    destino = os.path.join(TEMP_ASCII, sanitize(nome_original))
    shutil.copy2(origem, destino)
    return destino


def main():
    os.makedirs(PASTA_SAIDA, exist_ok=True)

    for p, msg in [(PYREVIT_CLI, "pyRevit.exe"), (SCRIPT, "script.py")]:
        if not os.path.isfile(p):
            print(u"[ERRO] {} nao encontrado: {}".format(msg, p))
            sys.exit(1)
    if not os.path.isdir(PASTA_MODELOS):
        print(u"[ERRO] Pasta de modelos nao existe: {}".format(PASTA_MODELOS))
        sys.exit(1)

    rvts = sorted(f for f in os.listdir(PASTA_MODELOS) if f.lower().endswith(".rvt"))
    if DEBUG_UM_SO:
        rvts = rvts[:1]
    print(u"Modelos a processar: {}".format(len(rvts)))

    for i, rvt in enumerate(rvts, 1):
        origem = os.path.join(PASTA_MODELOS, rvt)
        print(u"[{}/{}] {}".format(i, len(rvts), rvt))

        # ---- cascata de estrategia de path ----
        alvo, nota, copiado = None, u"", False

        if only_ascii(origem):
            alvo, nota = origem, u"path ASCII original"
        else:
            sp = get_short_path(origem)
            if sp:
                alvo, nota = sp, u"short path 8.3"
            else:
                try:
                    alvo = copiar_ascii(origem, rvt)
                    copiado = True
                    nota = u"copiado p/ {}".format(alvo)
                except Exception as e:
                    print(u"  [ERRO] copia falhou: {}".format(e))
                    with open(LOG_BATCH, "a", encoding="utf-8") as logf:
                        logf.write(u"\n=== {} ===\n".format(rvt))
                        logf.write(u"[SKIP] copia falhou: {}\n".format(e))
                    continue

        # ---- SINTAXE CORRETA (modelo posicional + --allowdialogs) ----
        cmd = [
            PYREVIT_CLI,
            "run",
            SCRIPT,
            alvo,                                    # modelo como 3o argumento
            u"--revit={}".format(REVIT_VERSAO),
            # u"--purge",   
            u"--allowdialogs",
        ]

        with open(LOG_BATCH, "a", encoding="utf-8") as logf:
            logf.write(u"\n=== {} ===\n".format(rvt))
            logf.write(u"MODO: {}\n".format(nota))
            logf.write(u"ALVO: {}\n".format(alvo))
            logf.write(u"CMD : {}\n".format(u" ".join(cmd)))
            try:
                r = subprocess.run(cmd, stdout=logf, stderr=logf, check=False)
                logf.write(u"exit code: {}\n".format(r.returncode))
            except Exception as e:
                logf.write(u"[EXCEPTION] {}\n".format(e))

        # remove copia temporaria
        if copiado and os.path.isfile(alvo):
            try:
                os.remove(alvo)
                print(u"  copia temporaria removida")
            except Exception as e:
                print(u"  [aviso] nao consegui apagar copia: {}".format(e))

        time.sleep(PAUSA)

    print(u"\nConcluido. Saida em: {}".format(PASTA_SAIDA))


if __name__ == "__main__":
    main()
