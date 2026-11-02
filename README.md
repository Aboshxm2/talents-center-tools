# Talents Center tools

AI tools for مركز الموهبة والإبداع والتميز (Talents Center), Deanship of Student Affairs,
Taibah University.

## Official letters skill (`talents-center-letters`)

Write what you want to say in everyday Arabic (dialect, notes, half sentences) and Claude
turns it into a formal official letter on the center's letterhead (الكليشة). The letter
gets the correct titles and honorifics, the standard greeting and closing, today's Hijri
date and the signature block. It comes as a Word file, ready to print and sign.

Example request:

> ابغى خطاب لعميد شؤون الطلاب نطلب منه يخاطب موهبة عشان يشاركون في برنامج الجاهزية التنافسية

### Install (once, about a minute)

1. Download [talents-center-letters.zip](https://github.com/Aboshxm2/talents-center-tools/raw/main/talents-center-letters.zip).
2. In Claude (web or desktop app), go to **Customize → Skills → + → Create skill → Upload a
   skill** and choose the zip.
3. Make sure **Code execution and file creation** is on, under **Settings → Capabilities**.
   The skill needs it to build the Word file.

Notes:

- Skills are per account: upload it from the account that will write the letters. On Team
  and Enterprise plans, an admin can install it for the whole organization instead.
- Menu names can change. Anthropic's current guide:
  [Use skills in Claude](https://support.claude.com/en/articles/12512180-use-skills-in-claude).
- The letterhead fonts **Tosh A** and **Alexandria** must be installed on the computer that
  opens the Word file, or Word substitutes other fonts.

### Use

Just ask in a normal chat; there's no need to mention the skill. Claude replies with:

- the letter text,
- the Word file,
- a short «للمراجعة» list of anything to check, such as missing details (an event date,
  for example). These are highlighted in yellow in the letter.

You can also ask for changes («اختصره»، «كبّر الخط»), the same letter for several
recipients, or a formal email instead of a letter.

### If skills can't be installed on the account

Attach `talents-center-letters.zip` to the chat and start with:

> استخدم المهارة الموجودة في الملف المرفق (فكّ الضغط واتبع تعليمات SKILL.md)، ثم اكتب خطابًا …

Attach it again in each new chat. Linking to this GitHub page instead is not reliable.
Claude can read the instructions from a link, but the Word file needs the template and
script inside Claude's sandbox. An installed or attached skill puts them there.

### Settings

Edit `talents-center-letters/assets/config.json`, then rebuild the zip (below) and upload
it again.

| Key | Controls |
|---|---|
| `signature_title`, `signature_name` | The signature block |
| `header_unit_line` | The unit name printed under «عمادة شؤون الطلاب» |
| `body_font_size` | Body text size in points (default 12) |
| `default_closing` | The closing phrase |
| `greeting_font` | Optional calligraphic font for the السلام عليكم line |
| `contact` | Default name and mobile for the contact line |

### Updating the skill

Claude keeps the copy you uploaded, so changes in this repo reach no one until the new zip
is uploaded. After changing anything in `talents-center-letters/`, rebuild the zip from
the repo root:

```bash
rm -f talents-center-letters.zip
zip -r talents-center-letters.zip talents-center-letters -x "talents-center-letters/evals/*" "*/__pycache__/*"
```

Without the `zip` command, compress the `talents-center-letters` folder with your file
manager. The zip must contain the folder itself, not just its contents.

## What's in this repo

- `talents-center-letters/` holds the skill:
  - `SKILL.md`: the instructions
  - `references/`: writing guide and example letters
  - `scripts/build_letter.py`: fills the template (Python standard library only)
  - `assets/`: the official template and settings
  - `evals/`: test prompts
- `talents-center-letters.zip`: the packaged skill to upload.
- `samples/`: letters produced by the skill.
- `كليشة الخطابات (1).docx`: the original letterhead template.

Signed letters are kept out of this public repo: `*.pdf` is in `.gitignore`.
