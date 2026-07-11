#!/bin/bash -u

if ! [ "$(id -u)" -eq 0 ]; then
 	echo "[halt] Not running with superuser privileges."
 	exit 1
fi

echo "[okay] running as superuser."

vercomp() {
    [ "$1" = "$2" ] && return 0

    local IFS=.
    local i ver1=($1) ver2=($2)
    unset IFS
    if [ "${#ver1[@]}" -ne "${#ver2[@]}" ]; then
        echo "[halt] versions being compared don't have the same form!" >&2
        echo "[halt] '$1' vs '$2'" >&2
        exit 1
    fi

    for ((i = 0; i < ${#ver1[@]}; ++i)); do
        (( ${ver1[i]} > ${ver2[i]} )) && return 1
        (( ${ver1[i]} < ${ver2[i]} )) && return 2
    done

    return 0
}

nasm_version_check () {
    vercomp $1 $2
    case $? in
        0) op='=';;
        1) op='>';;
        2) op='<';;
    esac
    if [ $op = '=' ] || [ $op = '>' ]; then
        echo "[okay] nasm version at least '$2'"
        return 0
    else
        echo "[halt] nasm version is too low" >&2
        return 1
    fi
}

# Main
declare -r NASM_VER_REGEX='[0-9]+\.[0-9]+(\.[0-9]+)?'
declare -r MINIMUM_NASM_VERSION=2.10.09
declare -r NASM_PATH=$(which nasm)

declare current_nasm_version=

if [ -z "$NASM_PATH" ]; then
	echo "[halt] nasm was not found on the system! Make sure it is named nasm and is in the path."
	exit 1
fi

current_nasm_version=$(nasm -v)

[ $? -ne 0 ] && echo '[halt] error calling nasm' >&2 && exit 1

if [[ "$current_nasm_version" =~ $NASM_VER_REGEX ]]; then
    current_nasm_version=${BASH_REMATCH[0]}

    # NASM 3.x returns versions like 3.02, but the script expects X.Y.Z
    if [[ "$current_nasm_version" =~ ^[0-9]+\.[0-9]+$ ]]; then
        current_nasm_version="${current_nasm_version}.00"
    fi

    echo "[okay] found current nasm version of '$current_nasm_version'"
else
    echo "[halt] could not determine nasm version" >&2
    exit 1
fi

nasm_version_check "$current_nasm_version" "$MINIMUM_NASM_VERSION"

[ $? -ne 0 ] && echo "[halt] nasm not found or version is incompatible" >&2 && exit 1

"$NASM_PATH" -O0 -f bin -o source/bootload/bootload.bin source/bootload/bootload.asm || exit 1
echo "[okay] assembled bootloader"

cd source
"$NASM_PATH" -O0 -f bin -o kernel.bin kernel.asm || exit 1
echo "[okay] assembled kernel"
cd ..

cd programs
for i in *.asm; do
	"$NASM_PATH" -O0 -f bin $i -o "$(basename $i .asm).bin" || exit 1
	echo "[okay] assembled program: $i"
done
echo "[okay] assembled all programs"
cd ..

cp disk_images/mikeos.flp disk_images/mikeos.dmg
echo "[okay] copied floppy image"

dd conv=notrunc if=source/bootload/bootload.bin of=disk_images/mikeos.dmg || exit 1
echo "[okay] added bootloader to image"

tmp_file=$(mktemp -d /tmp/$(basename $0).XXXXXX)
[ $? -ne 0 ] && echo "[halt] error creating a temp file" >&2 && exit 1

dev=$(echo -n $(hdid -nobrowse -nomount disk_images/mikeos.dmg))
[ $? -ne 0 ] && echo "[halt] could not create disk from image" >&2 && exit 1

mount -t msdos "$dev" "$tmp_file"
[ $? -ne 0 ] && echo "[halt] could not mount "$dev"" >&2 && exit 1

cp source/kernel.bin "$tmp_file/"
cp programs/*.bin programs/*.bas programs/sample.pcx programs/vedithlp.txt programs/gen.4th programs/hello.512 "$tmp_file"
echo "[okay] added programs to image"

diskutil umount "$tmp_file"
hdiutil detach "$dev"
rm -rf "$tmp_file"
echo "[okay] unmounted floppy image"

rm -f disk_images/mikeos.iso
if ! command -v mkisofs >/dev/null 2>&1; then
	echo "[halt] mkisofs was not found; install it with: brew install cdrtools" >&2
	exit 1
fi
mkisofs -quiet -V 'MIKEOS' -input-charset iso8859-1 -o disk_images/mikeos.iso -b mikeos.dmg disk_images/ || exit 1
echo "[okay] converted floppy to ISO image"
echo "[done] build completed"
