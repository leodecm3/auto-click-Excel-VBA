ECHO OFF
CLS
ECHO.
ECHO.
ECHO.
ECHO.
ECHO.
ECHO.
ECHO.
ECHO.
ECHO  VOU APAGAR O QUE TA AQUI NESSA PASTA E COLAR O CONTEUDO DO GITHUB POR CIMA
ECHO.
ECHO.
ECHO.
ECHO.
ECHO.
ECHO.
ECHO.
ECHO.
ECHO.
set /p opcao=.  TEM CERTEZA???   (y/n): 


if "%opcao%"=="y" goto tacaFogoNessaMerda
if "%opcao%"=="Y" goto tacaFogoNessaMerda
if "%opcao%"=="s" goto tacaFogoNessaMerda
if "%opcao%"=="S" goto tacaFogoNessaMerda

CLS


ECHO.
ECHO.
ECHO.
ECHO.
ECHO.
ECHO    saindo SEM FAZER NADA  :))))
ECHO.
ECHO.
ECHO.
ECHO.
ECHO.
ECHO.


timeout 2 > NUL

exit
pause

:tacaFogoNessaMerda
ECHO ON
CLS
git pull origin master
pause