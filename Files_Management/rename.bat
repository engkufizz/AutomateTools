@ECHO OFF
PUSHD .
FOR /R %%d IN (.) DO (
    cd "%%d"
    IF EXIST *.cfg (
       REN *.cfg *.txt
    )
)
POPD