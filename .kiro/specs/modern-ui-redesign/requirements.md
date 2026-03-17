# Documento dei Requisiti

## Introduzione

Questo documento descrive i requisiti per la modernizzazione completa dell'interfaccia utente dell'applicazione desktop Windows Forms "Auser Excel Transformer" (AuserGestioneTrasporti). L'obiettivo è trasformare l'attuale UI funzionale ma visivamente datata in un'interfaccia moderna, elegante e piacevole, mantenendo tutte le funzionalità esistenti. Il redesign adotta una palette cromatica definita (bianco primario, #393939 carbone scuro, #009246 verde italiano, #FAB900 ambra/oro), include il logo aziendale, e garantisce che tutti gli elementi si adattino correttamente a qualsiasi dimensione di finestra con layout centrato.

## Glossario

- **MainForm**: Il form principale dell'applicazione che ospita i controlli di trasformazione CSV/Excel e il VolunteerPanel.
- **VolunteerPanel**: Il pannello utente incorporato nel MainForm che gestisce contatti volontari, credenziali Gmail, selezione Excel e invio email.
- **ThemeManager**: Il componente responsabile della gestione centralizzata dei colori, font e stili dell'applicazione.
- **ModernButton**: Un controllo Button personalizzato con stile moderno, angoli arrotondati e feedback visivo hover/press.
- **ModernTextBox**: Un controllo TextBox personalizzato con bordo stilizzato e placeholder text.
- **ModernGroupBox**: Un controllo GroupBox personalizzato con intestazione stilizzata e bordo moderno.
- **HeaderPanel**: Il pannello superiore del MainForm che contiene il logo aziendale e il titolo dell'applicazione.
- **ContentPanel**: Il pannello centrale scrollabile che contiene tutti i controlli funzionali.
- **Palette**: L'insieme dei colori ufficiali: Bianco (#FFFFFF), Carbone (#393939), Verde (#009246), Ambra (#FAB900).
- **Responsive Layout**: Un layout che ridimensiona e riposiziona i controlli proporzionalmente al variare delle dimensioni della finestra.

## Requisiti

### Requisito 1: Sistema di Tema Centralizzato

**User Story:** Come sviluppatore, voglio un sistema di tema centralizzato, in modo da poter applicare e modificare lo stile visivo dell'intera applicazione da un unico punto.

#### Criteri di Accettazione

1. THE ThemeManager SHALL definire i colori primari dell'applicazione: sfondo principale Bianco (#FFFFFF), testo primario Carbone (#393939), accento primario Verde (#009246), accento secondario Ambra (#FAB900).
2. THE ThemeManager SHALL definire la tipografia dell'applicazione usando il font "Segoe UI" con dimensioni 24pt per titoli, 12pt per sottotitoli, 10pt per testo normale e 9pt per testo secondario.
3. THE ThemeManager SHALL esporre metodi statici per applicare lo stile a ogni tipo di controllo (Button, TextBox, Label, GroupBox, ListView, ComboBox, ProgressBar).
4. WHEN un controllo viene stilizzato tramite ThemeManager, THE ThemeManager SHALL applicare colori, font e dimensioni coerenti con la Palette definita.
5. THE ThemeManager SHALL definire tre varianti di stile per i pulsanti: Primary (sfondo Verde #009246, testo Bianco), Secondary (sfondo Carbone #393939, testo Bianco) e Accent (sfondo Ambra #FAB900, testo Carbone).

---

### Requisito 2: Header con Logo e Titolo

**User Story:** Come utente, voglio vedere il logo aziendale e il titolo dell'applicazione in cima alla finestra, in modo da riconoscere immediatamente l'applicazione e percepirne la professionalità.

#### Criteri di Accettazione

1. THE HeaderPanel SHALL essere posizionato nella parte superiore del MainForm con altezza fissa di 80 pixel e sfondo Carbone (#393939).
2. THE HeaderPanel SHALL caricare e visualizzare l'immagine da `logo/Auser_logo.png` ridimensionata proporzionalmente a un'altezza massima di 60 pixel, allineata a sinistra con margine di 20 pixel.
3. IF il file `logo/Auser_logo.png` non è accessibile, THEN THE HeaderPanel SHALL visualizzare solo il testo del titolo senza interrompere l'avvio dell'applicazione.
4. THE HeaderPanel SHALL visualizzare il testo "Auser Gestione Trasporti" in font "Segoe UI" 18pt Bold, colore Bianco (#FFFFFF), centrato verticalmente e posizionato a destra del logo.
5. WHILE la finestra viene ridimensionata, THE HeaderPanel SHALL mantenere la propria altezza fissa e adattare la larghezza all'intera larghezza del form con Anchor Left+Right+Top.

---

### Requisito 3: Layout Responsive e Centrato del MainForm

**User Story:** Come utente, voglio che tutti i controlli si adattino correttamente a qualsiasi dimensione di finestra, in modo da poter usare l'applicazione su schermi di diverse dimensioni senza perdere funzionalità o leggibilità.

#### Criteri di Accettazione

1. THE MainForm SHALL impostare una dimensione minima di 700x600 pixel per garantire la visibilità di tutti i controlli essenziali.
2. THE ContentPanel SHALL essere un pannello scrollabile posizionato sotto l'HeaderPanel, con Anchor Top+Bottom+Left+Right, che occupa tutto lo spazio rimanente del form.
3. THE ContentPanel SHALL contenere un pannello interno centrato con larghezza massima di 900 pixel, centrato orizzontalmente nel ContentPanel.
4. WHILE la larghezza del ContentPanel supera 940 pixel, THE ContentPanel SHALL mantenere il pannello interno centrato con margini laterali uguali.
5. WHILE la larghezza del ContentPanel è inferiore a 940 pixel, THE ContentPanel SHALL espandere il pannello interno fino alla larghezza disponibile meno 40 pixel di margine totale.
6. THE MainForm SHALL impostare lo sfondo Bianco (#FFFFFF) come colore di sfondo principale.
7. WHEN il MainForm viene ridimensionato, THE MainForm SHALL ridistribuire i controlli interni mantenendo le proporzioni e la centratura definite.

---

### Requisito 4: Stile Moderno dei Pulsanti

**User Story:** Come utente, voglio pulsanti visivamente moderni con feedback visivo chiaro, in modo da capire immediatamente quali azioni sono disponibili e quale sto per eseguire.

#### Criteri di Accettazione

1. THE ModernButton SHALL avere angoli arrotondati con raggio di 6 pixel, ottenuti tramite override del metodo OnPaint.
2. THE ModernButton SHALL visualizzare un effetto hover cambiando la luminosità del colore di sfondo del 15% quando il cursore entra nel controllo.
3. THE ModernButton SHALL visualizzare un effetto press scurendo il colore di sfondo del 20% quando il pulsante viene premuto.
4. THE ModernButton SHALL avere altezza minima di 40 pixel e padding orizzontale di 20 pixel.
5. WHEN un ModernButton è disabilitato, THE ModernButton SHALL visualizzare sfondo grigio (#CCCCCC) e testo grigio scuro (#888888) per indicare chiaramente lo stato non interattivo.
6. THE ModernButton di tipo Primary SHALL usare sfondo Verde (#009246) e testo Bianco (#FFFFFF).
7. THE ModernButton di tipo Secondary SHALL usare sfondo Carbone (#393939) e testo Bianco (#FFFFFF).
8. THE ModernButton di tipo Accent SHALL usare sfondo Ambra (#FAB900) e testo Carbone (#393939).

---

### Requisito 5: Stile Moderno dei Campi di Testo

**User Story:** Come utente, voglio campi di testo con uno stile moderno e pulito, in modo da avere un'esperienza visiva coerente e professionale durante l'inserimento dei dati.

#### Criteri di Accettazione

1. THE ModernTextBox SHALL avere un bordo inferiore di 2 pixel in colore Verde (#009246) al posto del bordo standard di sistema, ottenuto tramite override del metodo OnPaint.
2. WHEN un ModernTextBox riceve il focus, THE ModernTextBox SHALL evidenziare il bordo inferiore con colore Ambra (#FAB900) e spessore 2 pixel.
3. WHEN un ModernTextBox perde il focus, THE ModernTextBox SHALL ripristinare il bordo inferiore in colore Verde (#009246).
4. THE ModernTextBox SHALL avere sfondo Bianco (#FFFFFF) e testo Carbone (#393939).
5. WHERE un ModernTextBox ha un placeholder configurato, THE ModernTextBox SHALL visualizzare il testo placeholder in colore grigio (#AAAAAA) quando il campo è vuoto e non ha il focus.

---

### Requisito 6: Stile Moderno dei GroupBox

**User Story:** Come utente, voglio sezioni visivamente distinte e moderne nel pannello volontari, in modo da navigare facilmente tra le diverse aree funzionali.

#### Criteri di Accettazione

1. THE ModernGroupBox SHALL visualizzare un'intestazione con sfondo Carbone (#393939), testo Bianco (#FFFFFF) in font "Segoe UI" 10pt Bold, con padding di 8 pixel.
2. THE ModernGroupBox SHALL avere un bordo di 1 pixel in colore grigio chiaro (#E0E0E0) con angoli arrotondati di 4 pixel.
3. THE ModernGroupBox SHALL avere sfondo interno Bianco (#FFFFFF).
4. WHILE il ModernGroupBox viene ridimensionato, THE ModernGroupBox SHALL adattare la larghezza dell'intestazione all'intera larghezza del controllo.

---

### Requisito 7: Sezione Trasformazione CSV/Excel nel MainForm

**User Story:** Come utente, voglio una sezione di trasformazione file visivamente chiara e moderna, in modo da selezionare i file e avviare l'elaborazione con facilità.

#### Criteri di Accettazione

1. THE MainForm SHALL organizzare i controlli di trasformazione in una card con sfondo Bianco, bordo grigio chiaro (#E0E0E0) di 1 pixel e ombra leggera, con padding interno di 24 pixel.
2. THE MainForm SHALL visualizzare un'etichetta di sezione "Trasformazione File" in font "Segoe UI" 14pt Bold, colore Carbone (#393939), sopra i controlli di selezione file.
3. THE MainForm SHALL disporre il pulsante "Seleziona CSV" e il relativo percorso file su una riga orizzontale, con il pulsante a sinistra (larghezza 180px) e il percorso a destra che si espande.
4. THE MainForm SHALL disporre il pulsante "Seleziona Excel" e il relativo percorso file su una riga orizzontale separata, con la stessa struttura del CSV.
5. THE MainForm SHALL disporre i pulsanti "Elabora" e "Scarica" affiancati orizzontalmente, entrambi con stile Primary (Verde).
6. THE MainForm SHALL visualizzare l'etichetta di stato con font "Segoe UI" 9pt, colore Verde (#009246) per successo e Rosso (#D32F2F) per errore.
7. WHEN il percorso file selezionato supera la larghezza disponibile, THE MainForm SHALL troncare il testo con ellissi (...) mantenendo visibile la parte finale del percorso.

---

### Requisito 8: Sezione Volontari Modernizzata

**User Story:** Come utente, voglio che il pannello di gestione volontari abbia lo stesso stile moderno del resto dell'applicazione, in modo da avere un'esperienza visiva coerente.

#### Criteri di Accettazione

1. THE VolunteerPanel SHALL applicare il ThemeManager a tutti i propri controlli durante l'inizializzazione.
2. THE VolunteerPanel SHALL sostituire i GroupBox standard con ModernGroupBox per le quattro sezioni: Contatti Volontari, Credenziali Gmail, Selezione Excel, Invio Email.
3. THE VolunteerPanel SHALL applicare lo stile ModernButton Secondary ai pulsanti "Aggiungi Volontari", "Aggiungi Contatto", "Elimina Tutti" e "Cancella Credenziali".
4. THE VolunteerPanel SHALL applicare lo stile ModernButton Primary al pulsante "Invia Email".
5. THE VolunteerPanel SHALL applicare lo stile ModernTextBox ai campi "Email Gmail" e "Password App".
6. THE VolunteerPanel SHALL stilizzare la ListView dei volontari con intestazioni di colonna in sfondo Carbone (#393939) e testo Bianco, righe alternate in Bianco e grigio chiarissimo (#F5F5F5).
7. THE VolunteerPanel SHALL stilizzare la ProgressBar con colore di riempimento Verde (#009246) su sfondo grigio chiaro (#E0E0E0).
8. WHILE il VolunteerPanel viene ridimensionato, THE VolunteerPanel SHALL adattare la larghezza di tutti i controlli interni mantenendo i margini definiti tramite Anchor Left+Right.

---

### Requisito 9: Adattabilità Responsive a Tutte le Dimensioni

**User Story:** Come utente, voglio che l'applicazione sia utilizzabile su schermi di qualsiasi dimensione, da laptop piccoli a monitor grandi, in modo da non avere problemi di visualizzazione indipendentemente dall'hardware.

#### Criteri di Accettazione

1. THE MainForm SHALL supportare ridimensionamento da una dimensione minima di 700x600 pixel fino a dimensioni massime illimitate senza perdita di funzionalità.
2. WHEN la finestra viene ridimensionata a larghezze superiori a 1200 pixel, THE MainForm SHALL mantenere il contenuto centrato con larghezza massima di 900 pixel e margini laterali uguali.
3. WHEN la finestra viene ridimensionata a larghezze inferiori a 740 pixel, THE MainForm SHALL attivare lo scroll orizzontale per consentire l'accesso a tutti i controlli.
4. THE VolunteerPanel SHALL usare Anchor AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top per tutti i controlli che devono espandersi orizzontalmente.
5. THE MainForm SHALL usare AutoScroll = true per consentire lo scroll verticale quando il contenuto supera l'altezza visibile.
6. WHEN la finestra viene ridimensionata, THE MainForm SHALL aggiornare la posizione e larghezza del pannello interno centrato tramite l'evento Resize.

---

### Requisito 10: Icona e Identità Visiva dell'Applicazione

**User Story:** Come utente, voglio che l'applicazione abbia un'identità visiva coerente inclusa l'icona nella barra delle applicazioni, in modo da riconoscerla facilmente tra le applicazioni aperte.

#### Criteri di Accettazione

1. THE MainForm SHALL caricare l'icona dell'applicazione dal file `logo/favicon_auser.png` convertita in formato ICO, oppure dall'icona embedded esistente `Resources/app_icon.ico`.
2. THE MainForm SHALL impostare il titolo della finestra come "Auser Gestione Trasporti" coerentemente con il branding.
3. THE MainForm SHALL visualizzare il logo `logo/Auser_logo.png` nell'HeaderPanel con rendering di alta qualità (InterpolationMode.HighQualityBicubic).
4. IF il logo non è disponibile al percorso specificato, THEN THE MainForm SHALL continuare l'avvio normalmente visualizzando solo il testo del titolo nell'HeaderPanel.

---

### Requisito 11: Dialogo "Aggiungi Contatto" Modernizzato

**User Story:** Come utente, voglio che il dialogo per aggiungere un contatto volontario abbia lo stesso stile moderno dell'applicazione principale, in modo da avere un'esperienza visiva coerente.

#### Criteri di Accettazione

1. THE VolunteerPanel SHALL creare il dialogo "Aggiungi Contatto" con sfondo Bianco (#FFFFFF), bordo fisso e posizione centrata rispetto al form padre.
2. THE VolunteerPanel SHALL applicare ModernTextBox ai campi "Cognome" ed "Email" nel dialogo.
3. THE VolunteerPanel SHALL applicare ModernButton Primary al pulsante "OK" e ModernButton Secondary al pulsante "Annulla" nel dialogo.
4. THE VolunteerPanel SHALL impostare la dimensione del dialogo a 420x220 pixel con font "Segoe UI" 9pt per tutte le etichette.
