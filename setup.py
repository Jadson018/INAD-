import sys
from cx_Freeze import setup, Executable

# Inclua qualquer arquivo adicional necessário aqui
files = ["polo.png"]

# Dependências adicionais podem ser especificadas se necessário
build_exe_options = {
    "packages": ["tkinter", "tkcalendar", "pandas", "matplotlib", "PIL"],
    "include_files": files,
}

# Definição do executável
base = None
if sys.platform == "win32":
    base = "Win32GUI"

# Configuração do setup
setup(name='INAD+',
      version='1.0',
      description='Analisador de planilhas especificas',
      options={'build_exe': build_exe_options},
      executables=[Executable("INAD+.py", base=base)]
      )

