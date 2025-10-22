# start.py
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WBS Filling From IFC - buildingSMART Portugal
Entry point da aplicação
"""

import sys
import os

if sys.version_info < (3, 9):
    print("ERRO: Python 3.9 ou superior é necessário")
    print(f"Versão atual: {sys.version}")
    sys.exit(1)

try:
    from app import __version__
    if '--version' in sys.argv:
        print(f"WBS Filling From IFC v{__version__}")
        sys.exit(0)
except ImportError:
    pass

from app.gui.main import main

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nAplicação interrompida pelo usuário")
        sys.exit(0)
    except Exception as e:
        print(f"\nERRO FATAL: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)