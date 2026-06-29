@echo off
title ERB Manager
echo =======================================
echo  ERB Manager - Rapprochement Bancaire
echo =======================================
echo.
echo Installation des dependances...
pip install streamlit pandas openpyxl xlrd numpy --quiet
echo.
echo Lancement de l'application...
echo Ouvrez votre navigateur sur : http://localhost:8501
echo Pour arreter : fermez cette fenetre ou CTRL+C
echo.
streamlit run app.py --server.fileWatcherType none
pause
