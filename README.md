# Python Tools

A small collection of Python command-line tools.
- **Drip Writer**: types out text you paste into it, into any window, with randomized pauses between keystrokes.
- **AI Solver**: reads questions off your screen with an AI model and fills in the answers for you.

All commands below are typed into the command line:
- **Windows:** press the Start button, type `Terminal`, and open it. (On older Windows, open `PowerShell` instead.)
- **Mac:** press `Cmd + Space`, type `Terminal`, and open it.

Copy and paste each command **one at a time**, pressing Enter after each one.

> Windows and Mac use different commands. Only follow the steps for your system.

---

## Step 1: Install Python

Download and run the installer from **[python.org/downloads](https://www.python.org/downloads/)** (it works for both Windows and Mac). Use the default options.

**Mac only:** when the installer finishes, a Finder window opens. Double-click **`Install Certificates.command`** in it. (You can also find it later in **Applications → Python 3.x**.) This prevents connection errors later.

Then **close Terminal and open a new one**, and check that it worked.

**Windows:**
```
py --version
```

**Mac:**
```
python3 --version
```

If it shows **Python 3.10 or newer**, you're ready. Open the tool you want below and continue with Step 2.

> **Mac:** if it shows an older version like `Python 3.9.6`, that's the old Python built into macOS. Quit Terminal completely (`Cmd + Q`), open it again, and check again.

---

<details>
<summary><strong>Step 2: Drip Writer</strong> (click to open)</summary>

### 2.1 Download it
1. On this repo's GitHub page, click the **Code** button, then **Download ZIP**, or use https://u2l.ai/6ztWrn.
2. Unzip it. Inside the unzipped folder, find the **`dripwriter`** folder and move it into your **Downloads** folder.

### 2.2 Install its dependencies (one time only)

**Windows:**
```
py -m pip install keyboard python-docx selenium pywin32
```

**Mac:**
```
python3 -m pip install keyboard pyobjc-framework-Quartz python-docx selenium
```

### 2.3 Go to the dripwriter folder in Downloads

**Windows:**
```
cd Downloads\dripwriter
```

**Mac:**
```
cd ~/Downloads/dripwriter
```

### 2.4 Run it

**Windows:**
```
py dripwriter.py
```

**Mac:**
```
python3 dripwriter.py
```

### From now on
You only need steps **2.3** and **2.4** each time you want to use it: go to the folder, then run it.

> **Mac:** macOS has to allow Terminal to type for you. Go to **System Settings → Privacy & Security → Accessibility** and turn it on for **Terminal**. Then quit Terminal completely (`Cmd + Q`) and open it again. If Backspace or typing does nothing, this is the fix.

</details>

---

<details>
<summary><strong>Step 2: AI Solver</strong> (click to open)</summary>

### 2.1 Download it
1. On this repo's GitHub page, click the **Code** button, then **Download ZIP**, or use https://u2l.ai/4Xcypz.
2. Unzip it. Inside the unzipped folder, find the **`ai-solver`** folder and move it into your **Downloads** folder.

### 2.2 Install its dependencies (one time only)

**Windows:**
```
py -m pip install keyboard pyautogui pyperclip python-dotenv google-genai pillow pydantic pymupdf
```

**Mac:**
```
python3 -m pip install pynput pyautogui pyperclip python-dotenv google-genai pillow pydantic pymupdf
```

(The two lists differ on purpose: on a Mac, `pynput` replaces `keyboard`. `pymupdf` reads PDF textbooks for the optional Textbooks and References feature.)

> Already installed an older version? Run the line again; it only adds what's missing.

### 2.3 Get a free Gemini API key (one time only)
1. Go to **[aistudio.google.com/app/apikey](https://aistudio.google.com/app/apikey)** and sign in with a Google account.
2. Accept the terms if it asks.
3. Click **Create API key**, then click the copy icon next to your new key.

Keep this key private, like a password.

### 2.4 Go to the ai-solver folder in Downloads

**Windows:**
```
cd Downloads\ai-solver
```

**Mac:**
```
cd ~/Downloads/ai-solver
```

### 2.5 Create your .env file and add your key (one time only)

The `ai-solver` folder comes with a file called `.env.example`. Rename it to `.env`:

**Windows:**
```
ren .env.example .env
```

**Mac:**
```
mv .env.example .env
```

Then open `.env` in a text editor:

**Windows:**
```
notepad .env
```

**Mac:**
```
open -e .env
```

It contains this line:
```
GEMINI_API_KEY=ENTER_KEY_HERE
```

Replace `ENTER_KEY_HERE` with the key you copied (no spaces, no quotes). Save the file (`Ctrl + S` on Windows, `Cmd + S` on Mac) and close the editor.

> The file name starts with a dot, so it may be hidden in File Explorer or Finder. That's normal; the commands above work on it anyway.

### 2.6 Mac only: allow Terminal to see the screen and use the keyboard (one time only)

Open **System Settings → Privacy & Security** and turn **Terminal** on in all three of these lists (if you run the solver from another app, like iTerm or VS Code, turn on that app instead):

| Setting | Why the solver needs it |
| --- | --- |
| **Accessibility** | to click and type the answers |
| **Input Monitoring** | to notice when you press the hotkey |
| **Screen Recording** (called **Screen & System Audio Recording** on newer macOS) | to take the screenshot the AI reads |

If Terminal isn't in a list, click **+**, go to **Applications → Utilities**, and choose **Terminal**. Afterwards **quit Terminal completely (`Cmd + Q`)** and open it again. The permissions only take effect after a restart.

> The solver checks these when it starts. If one is missing, macOS may pop up its own permission prompt, and the solver prints exactly which one to turn on.

### 2.7 Run it

**Windows:**
```
py solver.py
```

**Mac:**
```
python3 solver.py
```

### 2.8 Use it
1. Leave the Terminal window open and switch to your browser with the questions. If Textbooks and References is on, wait for the line `[REFERENCES] … files loaded - ready, you can start` first.
2. Press the hotkey:
   - **Windows:** `Ctrl + Alt + S`
   - **Mac:** `Control + Option + S`
3. Wait for the chime: a rising chime means the answers were filled in; a low tone means the AI didn't give a usable answer (press again). Your mouse goes back to where it was.
4. For questions whose next part appears after you answer, just press the hotkey again.

**Emergency stop:** quickly move your mouse into the **top-left corner** of the screen.
**Quit the program:** click the Terminal window and press `Ctrl + C` (on Mac too: `Control + C`).

### From now on
You only need steps **2.4** and **2.7** each time you want to use it: go to the folder, then run it.

### Textbooks and References (optional)
For **non-math multiple-choice** questions (definitions, concepts, "which term does this scenario describe?"), the solver can look terms up in your own course files. Math questions are answered exactly the same either way.

1. Run the solver once. It creates a folder called **`Textbooks and References`** inside the `ai-solver` folder.
2. Put your files in that folder: **PDFs** (textbooks, notes, slides saved as PDF), **PowerPoint** (`.pptx`, including speaker notes), **Word** (`.docx`), or **`.txt`** files. Long file names are fine.
3. Set `USE_TEXTBOOKS_AND_REFERENCES = True` (see Settings below) and run the solver again.

On start it shows one line counting up (`1/3`, `2/3`, `3/3`) and then `… files loaded - ready, you can start`. The first start reads each file once; after that it starts almost instantly. If you add or change a file, restart the solver.

> Old `.doc` / `.ppt` files and scanned PDFs (pictures of pages) can't be read. Open old files and use **Save As → .docx / .pptx**.

### Settings (optional)
The settings are simple `True`/`False` switches near the top of `solver.py`. Open it with `notepad solver.py` (Windows) or `open -e solver.py` (Mac), change a value, save, and run the program again.

| Setting | What it does |
| --- | --- |
| `SAVE_SCREENSHOTS` | `True` = saves a screenshot after every press, so you can check what was filled in. |
| `AUTO_ZIP_SCREENSHOTS` | `True` = puts those screenshots into one `tempscreenshots.zip` file instead of a `tempscreenshots` folder. |
| `PLAY_SOUNDS` | turns the chime / error tone on or off. |
| `RETURN_MOUSE_WHEN_DONE` | `True` = puts your mouse back where it was after the answers are filled in. |
| `USE_TEXTBOOKS_AND_REFERENCES` | `True` = looks up terms in your `Textbooks and References` folder for non-math multiple-choice questions (see above). |

### Troubleshooting

| Problem | Fix |
| --- | --- |
| Windows: `py` is not recognized | Run the python.org installer again with the default options, then open a **new** Terminal. |
| Mac: `error: externally-managed-environment` when installing | You're using Homebrew's Python. Install Python from python.org (Step 1), open a new Terminal, and try again. |
| `No such file or directory` / `can't open file` | You're in the wrong folder, or the folder isn't in Downloads. Check the folder name and redo the `cd` step. |
| `WARNING: GEMINI_API_KEY ... ENTER_KEY_HERE` | You haven't renamed `.env.example` to `.env` and put your key in it yet (step 2.5), or didn't save the file. |
| Mac: pressing the hotkey does nothing, or `macOS isn't letting this program watch the keyboard` | Redo step 2.6. All three permissions must be on for Terminal, and Terminal must be fully restarted (`Cmd + Q`). |
| Mac: `WARNING: macOS hasn't given the app you run this from … these permissions yet` | Turn on the permissions it lists (step 2.6), then fully restart Terminal (`Cmd + Q`). |
| Mac: the screenshot shows only your wallpaper | Screen Recording permission is missing (step 2.6). |
| `[REFERENCES] PDF files need PyMuPDF` | Run the install line in step 2.2 again, then restart the solver. |
| `[REFERENCES] … old .doc format` (or `.ppt`) | Open the file and use **Save As → .docx** (or **.pptx**), then restart the solver. |
| `[REFERENCES] … almost no text - probably scanned images` | That PDF is pictures of pages, which can't be searched. Use a text PDF if you have one. |
| `429` / `RESOURCE_EXHAUSTED` / "limit used up" | That model's free daily limit is used up. It resets at midnight Pacific time. |

</details>
