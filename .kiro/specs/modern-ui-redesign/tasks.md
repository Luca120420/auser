# Piano di Implementazione: Modern UI Redesign

## Tasks

- [x] 1. Creare ThemeManager
  - [x] 1.1 Creare `UI/ThemeManager.cs` con costanti di colore e font della palette
  - [x] 1.2 Implementare metodi statici `ApplyPrimary`, `ApplySecondary`, `ApplyAccent` per `ModernButton`
  - [x] 1.3 Implementare metodi statici `ApplyStyle` per `Label`, `ListView`, `ComboBox`, `ProgressBar`

- [x] 2. Creare controlli custom
  - [x] 2.1 Creare `UI/Controls/ModernButton.cs` con override `OnPaint` per angoli arrotondati (raggio 6px), effetti hover/press e stati disabilitato
  - [x] 2.2 Creare `UI/Controls/ModernTextBox.cs` con bordo inferiore colorato, cambio colore al focus/blur e supporto placeholder
  - [x] 2.3 Creare `UI/Controls/ModernGroupBox.cs` con override `OnPaint` per intestazione Carbone, bordo grigio chiaro e angoli arrotondati

- [x] 3. Creare HeaderPanel
  - [x] 3.1 Creare `UI/HeaderPanel.cs` con altezza fissa 80px, sfondo Carbone, caricamento logo con fallback graceful e label titolo

- [x] 4. Aggiornare MainForm
  - [x] 4.1 Aggiungere `HeaderPanel` (Anchor Top+Left+Right) e `ContentPanel` scrollabile (Anchor Top+Bottom+Left+Right)
  - [x] 4.2 Creare `InnerPanel` centrato (max 900px) dentro `ContentPanel` e gestire l'evento `Resize` per ricalcolare posizione e larghezza
  - [x] 4.3 Creare `TransformCard` con bordo grigio chiaro, etichetta sezione "Trasformazione File" e layout orizzontale per pulsanti e percorsi file
  - [x] 4.4 Sostituire i `Button` esistenti con `ModernButton` applicando stile Primary a "Elabora" e "Scarica", Secondary a "Seleziona CSV" e "Seleziona Excel"
  - [x] 4.5 Impostare `MinimumSize = new Size(700, 600)`, `BackColor = Color.White` e titolo "Auser Gestione Trasporti"
  - [x] 4.6 Abilitare `AutoEllipsis = true` sui label dei percorsi file

- [x] 5. Aggiornare VolunteerPanel
  - [x] 5.1 Sostituire i quattro `GroupBox` con `ModernGroupBox` (Contatti Volontari, Credenziali Gmail, Selezione Excel, Invio Email)
  - [x] 5.2 Sostituire i `Button` con `ModernButton`: Secondary per "Aggiungi Volontari", "Aggiungi Contatto", "Elimina Tutti", "Cancella Credenziali"; Primary per "Invia Email"
  - [x] 5.3 Sostituire i `TextBox` credenziali Gmail con `ModernTextBox` con placeholder configurato
  - [x] 5.4 Applicare `ThemeManager.ApplyStyle` alla `ListView` (intestazioni Carbone, righe alternate) e alla `ProgressBar`
  - [x] 5.5 Impostare Anchor `Left | Right | Top` su tutti i controlli espandibili orizzontalmente
  - [x] 5.6 Aggiornare il dialogo "Aggiungi Contatto" con `ModernTextBox`, `ModernButton` Primary/Secondary e dimensioni 420x220px

- [x] 6. Registrare i nuovi file nel csproj
  - [x] 6.1 Aggiungere `UI/ThemeManager.cs`, `UI/Controls/ModernButton.cs`, `UI/Controls/ModernTextBox.cs`, `UI/Controls/ModernGroupBox.cs`, `UI/HeaderPanel.cs` agli `<ItemGroup>` di compilazione in `AuserExcelTransformer.csproj`

- [x] 7. Scrivere i test
  - [x] 7.1 Creare `Tests/ModernUIDesignTests.cs` con unit test per: costanti ThemeManager, proprietà HeaderPanel, MinimumSize MainForm, tipi controlli VolunteerPanel, proprietà dialogo Aggiungi Contatto, titolo form, avvio senza logo
  - [x] 7.2 Creare `Tests/ModernUIDesignPropertyTests.cs` con property test FsCheck per le 7 proprietà definite nel design
  - [x] 7.3 Aggiungere i due file di test agli `<ItemGroup>` di compilazione in `AuserExcelTransformer.csproj`

- [x] 8. Verificare la build
  - [x] 8.1 Eseguire `dotnet publish AuserExcelTransformer.csproj -c Release -r win-x64 --self-contained true -o build_output` e correggere eventuali errori di compilazione
