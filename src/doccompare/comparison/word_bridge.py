"""Isolated, serialized Word jobs using the installed macOS scripting API."""
from contextlib import contextmanager
from pathlib import Path
import subprocess
import time


def apple_string(value):
    # Never interpolate raw paths or author names into executable AppleScript.
    value = str(value)
    if any(c in value for c in '\r\n\x00'):
        raise ValueError('Sökväg och författarnamn får inte innehålla radbrytningar.')
    return '"' + value.replace('\\', '\\\\').replace('"', '\\"') + '"'


def word_workspace():
    root = Path.home() / 'Library/Group Containers/UBF8T346G9.Office/DocCompare'
    root.mkdir(parents=True, exist_ok=True)
    return root


@contextmanager
def word_lock(root, timeout=180):
    import fcntl
    with open(root / 'word.lock', 'a') as lock:
        deadline = time.monotonic() + timeout
        while True:
            try:
                fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
                break
            except BlockingIOError:
                if time.monotonic() >= deadline:
                    raise RuntimeError('Word används av en annan jämförelse. Försök igen när den är klar.')
                time.sleep(0.2)
        try:
            yield
        finally:
            fcntl.flock(lock, fcntl.LOCK_UN)


def comparison_script(folder, author):
    old, new, tracked, pdf, accepted, rejected = [apple_string(folder / (folder.name + '-' + name)) for name in (
        'original.docx', 'modified.docx', 'tracked.docx', 'document.pdf', 'accepted.docx', 'rejected.docx')]
    display_settings = {
        'inserted text color': 'blue', 'deleted text color': 'red',
        'inserted text mark': 'inserted text mark underline',
        'deleted text mark': 'deleted text mark strike through',
        'move from text color': 'green', 'move to text color': 'green',
        'move from text mark': 'move from text mark double strike through',
        'move to text mark': 'move to text mark double underline',
        'revised lines color': 'black', 'revised lines mark': 'revised lines mark outside border',
        'revised properties color': 'violet',
        'revised properties mark': 'revised properties mark color only',
    }
    saved = ', '.join(key + ' of settings' for key in display_settings)
    setup = '\n'.join(f'        set {key} of settings to {value}' for key, value in display_settings.items())
    restore = '\n'.join(f'    set {key} of settings to item {i} of savedSettings'
                        for i, key in enumerate(display_settings, 1))
    cleanup = '\n'.join(f'    my closeOwnedDocument({apple_string(folder / (folder.name + "-" + name))}, {apple_string(folder.name + "-" + name)})'
                        for name in ('original.docx', 'modified.docx', 'tracked.docx', 'accepted.docx', 'rejected.docx'))
    # All paths live in a unique directory inside Office's sandbox. Even when
    # opening an arbitrary user path works, saving there can return -1708.
    return f'''
on waitForDocument(expectedPath, expectedName)
    tell application "Microsoft Word"
        repeat 300 times
            try
                if (posix full name of document expectedName) is expectedPath then return expectedName
            end try
            delay 0.1
        end repeat
    end tell
    error "Word öppnade inte arbetskopian i tid."
end waitForDocument

on closeOwnedDocument(expectedPath, expectedName)
    tell application "Microsoft Word"
        try
            if (posix full name of document expectedName) is expectedPath then close document expectedName saving no
        end try
    end tell
end closeOwnedDocument

tell application "Microsoft Word"
    set sourceName to ""
    set resultName to ""
    set savedSettings to {{{saved}}}
    try
{setup}
        open POSIX file {old} add to recent files false
        set sourceName to my waitForDocument({old}, {apple_string(folder.name + "-original.docx")})
        compare document sourceName path {new} author name {apple_string(author)} target compare target current detect format changes true ignore all comparison warnings false add to recent files false
        set resultName to sourceName
        set sourceName to ""
        set print revisions of document resultName to true
        set show revisions of document resultName to true
        set revisions mode of view of active window of document resultName to in line revisions
        set revisions view of view of active window of document resultName to revisions view final
        set show revisions and comments of view of active window of document resultName to true
        save as document resultName file name {tracked} file format format document add to recent files false
        set resultName to my waitForDocument({tracked}, {apple_string(folder.name + "-tracked.docx")})
        repaginate document resultName
        save as document resultName file name {pdf} file format format PDF add to recent files false
        accept all revisions document resultName
        save as document resultName file name {accepted} file format format document add to recent files false
        set resultName to my waitForDocument({accepted}, {apple_string(folder.name + "-accepted.docx")})
        close document resultName saving no
        set resultName to ""
        open POSIX file {tracked} add to recent files false
        set resultName to my waitForDocument({tracked}, {apple_string(folder.name + "-tracked.docx")})
        reject all revisions document resultName
        save as document resultName file name {rejected} file format format document add to recent files false
        set resultName to my waitForDocument({rejected}, {apple_string(folder.name + "-rejected.docx")})
        close document resultName saving no
        set resultName to ""
        if sourceName is not "" then close document sourceName saving no
        set sourceName to ""
        set resultVersion to application version
    on error messageText number errorNumber
        try
            if resultName is not "" then close document resultName saving no
        end try
        try
            if sourceName is not "" then close document sourceName saving no
        end try
{cleanup}
{restore}
        error messageText number errorNumber
    end try
{cleanup}
{restore}
    return resultVersion
end tell
'''


def run_comparison(folder, author='DocCompare'):
    script = comparison_script(Path(folder), author)
    script = script.replace('\ntell application', '\nwith timeout of 150 seconds\ntell application', 1) + '\nend timeout'
    try:
        # Word receives its own shorter timeout so its error handler restores
        # settings before Python terminates a stalled AppleScript process.
        result = subprocess.run(['osascript', '-e', script],
                                capture_output=True, text=True, timeout=170)
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError('Word svarade inte i tid. Kontrollera öppna Word-dialoger och försök igen. '
                           'Ingen förenklad PDF har skapats.') from exc
    if result.returncode:
        if '-1743' in result.stderr:
            raise RuntimeError('Tillåt DocCompare att styra Microsoft Word under '
                               'Systeminställningar > Integritet och säkerhet > Automation.')
        raise RuntimeError('Word-jämförelsen misslyckades: ' + result.stderr.strip())
    for name in ('tracked.docx', 'document.pdf', 'accepted.docx', 'rejected.docx'):
        if not (Path(folder) / (Path(folder).name + '-' + name)).is_file():
            raise RuntimeError(f'Word skapade inte den förväntade arbetsfilen {name}.')
    return result.stdout.strip()
