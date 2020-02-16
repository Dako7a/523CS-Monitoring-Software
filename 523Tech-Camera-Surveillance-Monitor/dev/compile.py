from cx_Freeze import *
#pip install cx_Freeze
#pip install wmi
#pip install pywin32

exe = [Executable("523CamSurv.py", base = "Win32GUI", icon="favicon.ico",)] # NO CONSOLE
# exe = [Executable("523CamSurv.py")]  #WITH CONSOLE (debugging)

setup(
    name = "523CamSurv",
    options = {
        'build_exe':{
            'include_msvcr': True, 
            'packages': ['wmi', 'pywin'],   
            'add_to_path': True     
        }
    },
    version= '1.1',
    executables = exe
)
