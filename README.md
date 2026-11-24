# Libri di Natale · Antonio Catalano

Piccolo e-commerce natalizio per i libri di Antonio Catalano, con scroll orizzontale, scheda di ogni libro, carrello, conferma d'ordine con codice e QR, e appuntamenti a Bergamo.

## File
- `index.html` — il sito (tutti i dati sono in cima allo script: `NEGOZIO`, `LIBRI`, `APPUNTAMENTI`, `LIBRERIE`)
- `qr.js` — generatore di QR code (Kazuhiko Arase, licenza MIT)
- `img/` — copertine dei libri
- `ordini-google-sheet.gs` — script per il foglio Google degli ordini (istruzioni in fondo al file)
- `extra/storie-kami.jpg` — copertina della versione kamishibai di *Storie piccolissime*, non usata

## Da completare prima di andare online
In `index.html`, oggetto `NEGOZIO`:
1. `paypalClientId` — client-id del conto PayPal Business (attiva PayPal e "Paga con carta")
2. `ordiniURL` — URL `/exec` dell'app web del foglio Google
3. `ordiniChiave` — la stessa parola segreta impostata come `CHIAVE` nello script
4. `spedizione` — costo di spedizione (ora 5 €, di esempio)

Le mail degli ordini (ogni 3 ore) e il riepilogo per la biglietteria (ogni 10 ore) arrivano a universisensibili@sensibili.it, solo se c'è qualcosa di nuovo.
