# TOS setup and run guide for Windows

This guide covers Windows 10 and 11. The quickest path is to boot the prebuilt image in QEMU. Install the build tools only if you want to compile changes from source.

## Quick start: run the prebuilt image

1. Download and install the 64-bit Windows build of [QEMU](https://www.qemu.org/download/).
2. Open Command Prompt in the TOS repository folder.
3. Confirm that QEMU is available:

   ```bat
   qemu-system-i386 --version
   ```

   If Windows cannot find the command, use its full path (normally `C:\Program Files\qemu\qemu-system-i386.exe`) or add the QEMU installation folder to `PATH`.

4. Boot TOS from the supplied floppy image:

   ```bat
   qemu-system-i386 -m 16M -drive format=raw,file=disk_images\mikeos.flp,index=0,if=floppy -boot a -no-reboot
   ```

TOS opens at the command line. Type `HELP` to see built-in commands, `DIR` to list files, or a program name such as `PONG` to launch it. Use `Ctrl+Alt+G` to release the mouse or keyboard from QEMU, then close the QEMU window to stop the machine.

## Build from source

### 1. Install the required tools

- **NASM:** download the Windows ZIP or installer from the [official NASM releases](https://www.nasm.us/pub/nasm/releasebuilds/), then add the folder containing `nasm.exe` to `PATH`. The build uses NASM's flat-binary (`-f bin`) output format.
- **ImDisk Toolkit:** install the 64-bit package from the [ImDisk Toolkit files page](https://sourceforge.net/projects/imdisk-toolkit/files/). It supplies `imdisk.exe`, which temporarily mounts the floppy image as drive `B:`.
- **Windows PowerShell:** included with Windows 10 and 11. The build script uses it to place the 512-byte boot sector into the floppy image.
- **QEMU:** optional for building, but required for the run and verification steps above.

Open a new Command Prompt and verify the required build tools:

```bat
nasm -v
where imdisk
powershell -NoProfile -Command "$PSVersionTable.PSVersion"
```

### 2. Build TOS

The build temporarily uses drive `B:`. Make sure that drive letter is free, then open **Command Prompt as Administrator**, change to the repository folder, and run:

```bat
buildwin.bat
```

The script:

1. Assembles the bootloader, kernel, and assembly programs with NASM.
2. Writes the bootloader into `disk_images\mikeos.flp`.
3. Mounts that image as `B:` with ImDisk.
4. Copies the rebuilt kernel, programs, and data files into the image.
5. Dismounts `B:` before exiting.

The successful build ends with `Done!`. It updates `source\kernel.bin`, the compiled files under `programs\`, and `disk_images\mikeos.flp`.

### 3. Run the rebuilt image

```bat
qemu-system-i386 -m 16M -drive format=raw,file=disk_images\mikeos.flp,index=0,if=floppy -boot a -no-reboot
```

You do not need to rebuild `disk_images\mikeos.iso` to run the floppy image in QEMU. The current Windows build script updates only `mikeos.flp`.

## Troubleshooting

### A command is not recognized

Close and reopen Command Prompt after installing a tool. Run `where nasm`, `where imdisk`, or `where qemu-system-i386` to confirm that Windows can find it. If not, add the tool's installation directory to the system or user `PATH`.

### Drive B: is already in use

The script stops before mounting the image if `B:` is occupied. Disconnect or reassign the existing `B:` drive, then rerun the build.

### ImDisk reports access denied

Run Command Prompt as Administrator. Installing and mounting the ImDisk driver requires elevated permissions.

### A failed build leaves B: mounted

The script normally dismounts the image during error handling. If it remains mounted, run this from an Administrator Command Prompt:

```bat
imdisk -D -m B:
```

### QEMU opens but TOS does not boot

Run the command from the repository root and confirm that `disk_images\mikeos.flp` exists. Keep `format=raw` and `if=floppy` in the QEMU command so the image is treated as a floppy disk.

### Start over with the repository image

The build modifies the tracked floppy image. Before rebuilding, make a backup if the current image contains files you want to keep.
