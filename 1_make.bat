call SET_VB_PATH
COPY Form3.frm Form3.frm.tmp
fart Form3.frm $BUILD_TIME_STAMP$ "Build Date=%date:~10,4%-%date:~4,2%-%date:~7,2% Time=%time:~0,2%-%time:~3,2%-%time:~6,2%-%time:~9,2%"
"%VB_PATH%\vb6" /make Project1.vbp
COPY Form3.frm.tmp Form3.frm
sleep 1