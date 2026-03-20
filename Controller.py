import pytchat
import vboxapi
import time
import threading
import re

# --- НАСТРОЙКИ МАКСИМА (БОССА) ---
VIDEO_ID = "yW-z9tUW844"
VM_NAME = "Chatusesxp2"
SNAPSHOT_NAME = "chatuseswinxp"

# Таблица скан-кодов (нажать и отпустить)
SCANCODES = {
    'w': [0x11, 0x91], 'i': [0x17, 0x97], 'n': [0x31, 0xb1], 'r': [0x13, 0x93],
    'e': [0x12, 0x92], 'x': [0x2d, 0xad], 'p': [0x19, 0x99], 'o': [0x18, 0x98],
    'a': [0x1e, 0x9e], 's': [0x1f, 0x9f], 'd': [0x20, 0xa0], 'c': [0x2e, 0xae],
    't': [0x14, 0x94], 'm': [0x32, 0xb2], 'u': [0x16, 0x96], 'l': [0x26, 0xa6],
    'v': [0x2f, 0xaf], 'k': [0x25, 0xa5], ' ': [0x39, 0xb9], '.': [0x34, 0xb4],
    '\n': [0x1c, 0x9c], 'backspace': [0x0e, 0x8e], 'win': [0xdb, 0x5b]
}

mgr = vboxapi.VirtualBoxManager(None, None)
vbox = mgr.getVirtualBox()

def get_session():
    return mgr.getSessionObject(vbox)

def handle_vbox(full_message):
    try:
        mach = vbox.findMachine(VM_NAME)
        commands = re.findall(r'!\w+[^!]*', full_message)
        
        for part in commands:
            args = part.strip().split(' ')
            cmd = args[0].lower()
            val = " ".join(args[1:]) if len(args) > 1 else ""

            # --- 1. ПИТАНИЕ (ФИКС SAFEARRAY ДЛЯ VBOX 7.x) ---
            if cmd in ["!startvm", "!launchvm", "!start_mc"]:
                if mach.state < 5:
                    # КРИТИЧЕСКИЙ ФИКС: передаем пустые списки [] для env и параметров
                    # win32com требует их как последовательности для SAFEARRAY
                    mach.launchVMProcess(get_session(), "gui", []) 
                    print(f"БОСС, Chatusesxp2 ПОЕХАЛА!")
            
            elif cmd == "!restartvm":
                session = get_session()
                if mach.state >= 5:
                    mach.lockMachine(session, 1)
                    session.console.powerDown()
                    session.unlockMachine()
                    time.sleep(2)
                mach.launchVMProcess(get_session(), "gui", [])

            elif cmd == "!revert":
                session = get_session()
                if mach.state >= 5:
                    mach.lockMachine(session, 1)
                    session.console.powerDown()
                    session.unlockMachine()
                    while mach.state >= 5: time.sleep(0.5)
                snapshot = mach.findSnapshot(SNAPSHOT_NAME)
                mach.lockMachine(session, 1)
                session.console.restoreSnapshot(snapshot)
                session.unlockMachine()

            # --- 2. МЫШЬ ---
            elif cmd in ["!move", "!mouse", "!abs", "!cursor"]:
                session = get_session()
                mach.lockMachine(session, 1)
                coords = val.split()
                if len(coords) >= 2:
                    x, y = int(coords[0]), int(coords[1])
                    if "abs" in cmd or "cursor" in cmd:
                        session.console.mouse.putMouseEventAbsolute(x, y, 0, 0, 0)
                    else:
                        session.console.mouse.putMouseEvent(x, y, 0, 0, 0)
                session.unlockMachine()

            elif cmd in ["!click", "!lclick", "!rclick"]:
                session = get_session()
                mach.lockMachine(session, 1)
                btn = 2 if "r" in cmd else 1
                session.console.mouse.putMouseEvent(0, 0, 0, 0, btn)
                session.console.mouse.putMouseEvent(0, 0, 0, 0, 0)
                session.unlockMachine()

            # --- 3. КЛАВИАТУРА (ФИКС SAFEARRAY) ---
            elif cmd in ["!type", "!send", "!typeenter"]:
                session = get_session()
                mach.lockMachine(session, 1)
                text = val + ("\n" if "send" in cmd else "")
                for char in text.lower():
                    if char in SCANCODES:
                        # Передаем список [код], это и есть SAFEARRAY
                        session.console.keyboard.PutScancodes(SCANCODES[char])
                session.unlockMachine()

            elif cmd in ["!wait", "!pause"]:
                time.sleep(min(float(val), 5.0))

    except Exception as e:
        print(f"ОШИБКА ЛАВАКА: {e}")

# ЗАПУСК
chat = pytchat.create(video_id=VIDEO_ID)
print(f"БОСС МАКСИМ, КОД ПЕРЕЗАРЯЖЕН! Слушаю чат {VIDEO_ID}")

while chat.is_alive():
    for c in chat.get().sync_items():
        if "!" in c.message:
            threading.Thread(target=handle_vbox, args=(c.message,)).start()
