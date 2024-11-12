set gs="C:\Program Files\gs\gs9.52\bin\gswin64c.exe"
set tesseract="\Program Files\Tesseract-OCR\tesseract.exe"

if '%1'=='' goto :badParams
if '%2'=='' goto :badParams
mkdir %temp%\ocr\
set nm=%~n1
SETLOCAL ENABLEDELAYEDEXPANSION 

rem split pdf into multiple jpeg
%gs% -dSAFER -dBATCH -dNOPAUSE -sDEVICE=jpeg -r300 -dTextAlphaBits=4 -o "%temp%\ocr\ocr_%nm%_%%04d.jpg" -f "%1"

rem ocr each jpeg
for %%i in (%temp%\ocr\ocr_%nm%_*.jpg) do %tesseract% -l eng "%%i" %%~pni pdf
del %temp%\ocr\ocr_%nm%_*.jpg

rem combine pdfs
set ff=#
for %%i in (%temp%\ocr\ocr_%nm%_*.pdf) do set ff=!ff! %%i
set ff=%ff:#=%
%gs% -dNOPAUSE -dQUIET -dBATCH -dNOPAUSE -q -sDEVICE=pdfwrite -o "%2" %ff%
del %temp%\ocr\ocr_%nm%_*.pdf

goto :eof

:badParams
echo usage %0 pdf-In pdf-Out