@echo off


:start
cls

set python_ver=37

python ./get-pip.py

cd \
cd \python%python_ver%\Scripts\
pip install Pillow==7.0.0

pause
exit