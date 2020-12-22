from cx_Freeze import setup, Executable

base = None    

executables = [Executable("main.py", base=base)]

packages = ["idna", "smtplib", "email.mime.multipart", "email.mime.text", "base64", "email", "email.mime.base", "sys", "tkinter", "PIL", "os"]
options = {
    'build_exe': {    
        'packages':packages,
    },    
}

setup(
    name = "Email-Sender",
    options = options,
    version = "1.0",
    description = 'Fast email sender',
    executables = executables
)