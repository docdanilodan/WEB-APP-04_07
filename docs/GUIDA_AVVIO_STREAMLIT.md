# Guida rapida - WEB APP 04_07 Streamlit

## 1. Apri la web app

Avvia il comando:

```bash
streamlit run streamlit_app.py
```

## 2. Configura la mail

Nella sidebar inserisci:

| Campo | Valore esempio |
|---|---|
| IMAP host | imap.gmail.com |
| Porta | 993 |
| Email utente | nome@email.it |
| Password / App password | app password generata |
| Cartella IMAP | INBOX |

## 3. Cerca azienda

Vai su **Cerca Azienda**, scrivi la ragione sociale o una parola chiave presente nelle email.

## 4. Vedi tutto

Il sistema mostra email, allegati, sintesi e dati essenziali.

## 5. Scarica tutto

Con il comando **Scarica tutto** l'app crea la cartella cliente e archivia:

- email originale `.eml`;
- testo/sintesi email;
- allegati documentali;
- record nel database SQLite.

## 6. Importa documenti locali

Vai su **Importa Documenti** per caricare PDF, Word, Excel, TXT, XML, EML e altri file.

## 7. Archivio clienti

Vai su **Archivio Clienti** per vedere documenti salvati, email archiviate e generare ZIP cliente.
