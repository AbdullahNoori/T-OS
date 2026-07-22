@echo off
setlocal
pushd "%~dp0"

echo Build script for Windows
echo.

where nasm >nul 2>&1 || (
  echo ERROR: NASM was not found in PATH.
  goto error
)

where powershell >nul 2>&1 || (
  echo ERROR: Windows PowerShell was not found in PATH.
  goto error
)

where imdisk >nul 2>&1 || (
  echo ERROR: ImDisk was not found in PATH.
  goto error
)

if exist B:\ (
  echo ERROR: Drive B: is already in use. Free it before building TOS.
  goto error
)

echo Assembling bootloader...
nasm -O0 -f bin -o source\bootload\bootload.bin source\bootload\bootload.asm || goto error

echo Assembling MikeOS kernel...
nasm -O0 -f bin -Isource\ -o source\kernel.bin source\kernel.asm || goto error

echo Assembling programs...
for %%i in (programs\*.asm) do (
  nasm -O0 -f bin -Iprograms\ -o "programs\%%~ni.bin" "%%i" || goto error
)

echo Adding bootsector to disk image...
powershell -NoProfile -Command "$boot = [IO.File]::ReadAllBytes('source\bootload\bootload.bin'); if ($boot.Length -ne 512) { throw 'Bootloader must be exactly 512 bytes.' }; $image = [IO.File]::Open('disk_images\mikeos.flp', [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite); try { $image.Write($boot, 0, $boot.Length) } finally { $image.Dispose() }" || goto error

echo Mounting disk image...
imdisk -a -f "%CD%\disk_images\mikeos.flp" -s 1440K -m B: || goto error
set "image_mounted=1"

echo Copying kernel and applications to disk image...
copy /Y source\kernel.bin B:\ >nul || goto error
copy /Y programs\*.bin B:\ >nul || goto error
copy /Y programs\sample.pcx B:\ >nul || goto error
copy /Y programs\vedithlp.txt B:\ >nul || goto error
copy /Y programs\gen.4th B:\ >nul || goto error
copy /Y programs\hello.512 B:\ >nul || goto error
copy /Y programs\*.bas B:\ >nul || goto error

echo Dismounting disk image...
imdisk -D -m B: || goto error
set "image_mounted="

echo Done!
popd
exit /b 0

:error
if defined image_mounted imdisk -D -m B: >nul 2>&1
echo.
echo Build failed. Review the first error above.
popd
exit /b 1
