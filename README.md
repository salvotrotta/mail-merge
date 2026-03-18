# 📄 Mail Merge

Applicazione Windows per generare documenti personalizzati in serie a partire da un template Word e un foglio Excel.

![Version](https://img.shields.io/badge/versione-1.0.0-blue)
![Platform](https://img.shields.io/badge/piattaforma-Windows%2010%2F11-blue)
![Python](https://img.shields.io/badge/Python-3.8%2B-yellow)
![License](https://img.shields.io/badge/licenza-MIT-green)

---

## ✨ Funzionalità

- 📝 **Template Word** con segnaposto `{{campo}}` — mantiene tutta la formattazione originale (header, footer, font, allineamenti)
- 📊 **Origine dati Excel** (.xlsx) con selezione/deselezione dei record
- 📤 **Esportazione** in PDF, PDF/A o Word (.docx)
- 🏷️ **Nome file personalizzabile** con parte fissa, campi dinamici dall'Excel e separatori
- 💶 **Formato valuta** configurabile per colonne numeriche (es. `€ 1.234,56`)
- ⏹️ **Interruzione** della generazione in qualsiasi momento
- 📈 **Barra di avanzamento** con contatore in tempo reale

---

## 🖥️ Requisiti

- Windows 10 o 11
- [Python 3.8+](https://www.python.org/downloads/) — durante l'installazione spuntare **"Add Python to PATH"**
- [LibreOffice](https://www.libreoffice.org/download/) — necessario solo per esportare in PDF o PDF/A

---

## 🚀 Installazione

### Metodo 1 — Installer (consigliato)

1. Vai alla sezione [Releases](../../releases)
2. Scarica `MailMerge_Setup_v1.0.0.exe`
3. Esegui l'installer e segui le istruzioni

### Metodo 2 — Da sorgente

```bash
git clone https://github.com/salvotrotta/mail-merge.git
cd mail-merge
pip install -r requirements.txt
python mail_merge_gui.py
```

---

## 📖 Come si usa

1. **Carica il file Excel** (.xlsx) con i dati — la prima riga deve contenere le intestazioni
2. Clicca **"Vedi record"** per selezionare/deselezionare le righe da elaborare
3. **Carica il template Word** (.docx) con i segnaposto `{{NomeColonna}}`
4. **Seleziona la cartella** di output
5. **Configura il nome file** con parte fissa e campi dinamici
6. **Scegli il formato** di esportazione (PDF, PDF/A, Word)
7. Clicca **"Avvia generazione"**

### Esempio di template Word

```
Spett.le {{Ragione_Sociale}}
Via {{Indirizzo}}, {{CAP}} {{Città}}

Oggetto: Approvazione budget

Con la presente comunichiamo che il budget approvato
ammonta a {{budget_totale_approvato}}.
```

---

## 🗂️ Struttura del repository

```
mail-merge/
├── mail_merge_gui.py        # Applicazione principale
├── avvia.bat                # Launcher Windows
├── crea_icona.py            # Generatore icona
├── build.bat                # Script compilazione EXE
├── setup.iss                # Script installer Inno Setup
├── requirements.txt         # Dipendenze Python
├── icon.ico                 # Icona applicazione
├── .gitignore               # File ignorati da Git
└── README.md                # Questo file
```

---

## 🔧 Build dell'installer

Per compilare l'installer in autonomia:

1. Installa [Inno Setup 6](https://jrsoftware.org/isdl.php)
2. Esegui `build.bat` — genera `dist\MailMerge\`  
   *(oppure salta questo passo se usi il metodo senza PyInstaller)*
3. Apri `setup.iss` con Inno Setup → premi **F9**
4. L'installer viene creato in `installer_output\`

---

## 📋 Licenza

Distribuito sotto licenza [MIT](LICENSE).

---

## 🤝 Contributi

Pull request e segnalazioni di bug sono benvenuti!  
Apri una [Issue](../../issues) per qualsiasi problema o suggerimento.
