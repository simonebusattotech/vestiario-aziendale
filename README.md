# Vestiario aziendale – scansione barcode

Web app per smartphone (nessuna installazione) per tracciare il vestiario aziendale inviato in lavanderia.
I dati vengono salvati su un **Foglio Google** condiviso, scaricabile in Excel in qualsiasi momento.

**Funzioni**
- **Registra**: associa un capo (codice a barre) al CID del dipendente.
- **Invio**: segna i capi consegnati alla lavanderia (i capi nuovi vengono registrati automaticamente).
- **Rientro**: verifica i capi tornati. Segnala i capi di altri CID, i capi sconosciuti e quelli mancanti.

Per scansionare si può usare la fotocamera del telefono, un lettore bluetooth o l'inserimento manuale.

---

## Configurazione (una volta sola, a cura dell'amministratore)

### 1. Foglio Google + script
1. Su Google Drive crea un nuovo **Foglio Google**, per esempio "Vestiario".
2. Apri **Estensioni → Apps Script**, cancella il codice presente e incolla il contenuto di
   [`apps-script/Code.gs`](apps-script/Code.gs).
3. Nella riga `const API_KEY = 'CAMBIA-QUESTA-CHIAVE';` sostituisci la chiave con una parola segreta a tua scelta.
4. Salva, scegli la funzione **`setup`** nel menu in alto e premi **Esegui**. Autorizza l'accesso quando richiesto
   (se compare "App non verificata": *Avanzate → Vai a … (non sicuro)*, perché lo script è tuo).
   Nel foglio compaiono le schede **Capi**, **Movimenti** e **Utenti**.
5. **Distribuisci → Nuovo deployment** → icona ingranaggio → **App web**
   - Esegui come: **Me**
   - Chi ha accesso: **Chiunque**
   - Premi **Distribuisci** e copia l'**URL dell'app web** (termina con `/exec`).

> Se in seguito modifichi `Code.gs`, usa **Distribuisci → Gestisci deployment → Modifica → Nuova versione**.
> In questo modo l'URL resta lo stesso.

### 2. Configurare i telefoni
Apri l'app, tocca **⚙**, inserisci l'URL `/exec` e la chiave, poi premi **Salva**.
Con **"Copia link di configurazione"** ottieni un link da mandare ai colleghi (per esempio su WhatsApp):
aprendolo, il loro telefono si configura da solo.

Su Android/iPhone puoi usare **"Aggiungi a schermata Home"** per avere l'icona come una normale app.

### 3. (Facoltativo) CID ammessi
Nella scheda **Utenti** del foglio inserisci i CID abilitati (colonna A) e il nome (colonna B).
Se la scheda è vuota, l'app accetta qualsiasi CID.

## Scaricare l'Excel
Dal Foglio Google: **File → Scarica → Microsoft Excel (.xlsx)**.

## Struttura
| File | Descrizione |
|---|---|
| `index.html`, `app.js`, `style.css` | Web app (pubblicata con GitHub Pages) |
| `apps-script/Code.gs` | Backend da incollare in Google Apps Script |
| `versione-pc/` | Versione desktop offline (Python, file Excel locale) |
