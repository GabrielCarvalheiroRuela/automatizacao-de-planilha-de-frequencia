# pip install cx_Freeze
# python setup.py build
from cx_Freeze import setup, Executable
#, base="Win32GUI"
executables = [Executable("planilhafrequencia.py", base="Win32GUI")]

setup(
    name="NomeDoExecutavel",
    version="1.0",
    description="Descrição do seu programa",
    executables=executables
)
