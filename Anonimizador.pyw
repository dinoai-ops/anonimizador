"""Anonimizador Juridico LGPD — lancador do app local.

Duplo clique neste arquivo abre o app sem janela de terminal (o Windows o
executa com o pythonw, ja instalado junto com o Python). Diferente de um
.bat/.exe, um arquivo .pyw nao e tratado como executavel suspeito pelo
SmartScreen/Controle de Aplicativo Inteligente quando baixado da internet.
"""
import os
import sys

_BASE = os.path.dirname(os.path.abspath(__file__))
os.chdir(_BASE)
sys.path.insert(0, _BASE)

import app  # noqa: E402

app.main()
