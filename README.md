# WEB APP 04_07 - FinancePlus Email Azienda Streamlit PRO

Pacchetto Streamlit pronto per GitHub e Streamlit Cloud.

## Cosa contiene

- `WEB_APP_04_07.py` - file unico principale del progetto.
- `streamlit_app.py` - entrypoint consigliato per Streamlit Cloud.
- `requirements.txt` - dipendenze Python.
- `runtime.txt` - versione Python consigliata.
- `.streamlit/config.toml` - tema grafico e configurazione upload.
- `.streamlit/secrets.toml.example` - esempio credenziali IMAP senza password reali.
- `financeplus_data/` - struttura locale dati, clienti, export e temporanei.

## Funzioni principali

- Dashboard operativa.
- Comando **Cerca Azienda**.
- Ricerca email via IMAP per oggetto, mittente, corpo e testo.
- **Vedi Tutto**: anteprima email, sintesi, elenco allegati.
- **Scarica Tutto**: archivia email e allegati nella cartella cliente.
- Scarico selettivo di una o più email.
- Salvataggio email originale `.eml`.
- Esclusione automatica immagini.
- Rinomina documenti in formato: `AZIENDA_tipologia_GG-MM-AAAA.ext`.
- Archivio cliente automatico.
- Database SQLite locale.
- Import documenti singoli o multipli.
- Import cartella locale con sottocartelle quando l'app gira su PC/server.
- Download ZIP cartella cliente.

## Avvio locale

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

In alternativa:

```bash
streamlit run WEB_APP_04_07.py
```

## Configurazione email in locale

Puoi inserire i dati dalla sidebar dell'app oppure creare il file:

```text
.streamlit/secrets.toml
```

partendo da:

```text
.streamlit/secrets.toml.example
```

Esempio:

```toml
FP_IMAP_HOST = "imap.gmail.com"
FP_IMAP_PORT = "993"
FP_IMAP_USER = "nome@email.it"
FP_IMAP_PASSWORD = "app_password"
FP_IMAP_FOLDER = "INBOX"
```

Per Gmail usa una **App Password**, non la password normale dell'account.

## Pubblicazione su GitHub

1. Crea una repository GitHub.
2. Carica tutti i file e le cartelle del pacchetto.
3. Non caricare `.streamlit/secrets.toml` con password reali.
4. Mantieni `streamlit_app.py` nella root del repository.

## Pubblicazione su Streamlit Cloud

1. Vai su Streamlit Cloud.
2. Crea una nuova app collegando la repository GitHub.
3. Imposta **Main file path**:

```text
streamlit_app.py
```

4. In **Settings > Secrets** inserisci le credenziali IMAP.
5. Avvia il deploy.

## Nota importante su archiviazione dati

Su Streamlit Cloud il filesystem può essere temporaneo. Per uso professionale stabile conviene integrare un archivio persistente, per esempio Google Drive, pCloud, S3 o database esterno. La versione attuale usa SQLite locale e cartelle locali, ottime per PC/server o demo Cloud.

## Comandi rapidi

### Test sintassi

```bash
python -m py_compile WEB_APP_04_07.py streamlit_app.py
```

### Avvio

```bash
streamlit run streamlit_app.py
```
