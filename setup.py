
'''
from cx_Freeze import setup, Executable
base = None
# //Remplacer "monprogramme.py" par le nom du script qui lance votre programme
executables = [Executable("main.py", base=base)]
# //Renseignez ici la liste complète des packages utilisés par votre application
# packages = ['idna']
packages = ['numpy', 'docx', 'bs4', 'seaborn', 'matplotlib', 'pandas', 'PIL']
options = {
    'build_exe': {
        'packages': packages,
    },
}
# //Adaptez les valeurs des variables "name", "version", "description" à votre programme.
setup(
    name='ReportGenerator',
    options=options,
    version='1.0',
    description='Program that generates report',
    executables=executables
)
'''