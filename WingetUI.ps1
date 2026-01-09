Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Improve Unicode handling for external process output (winget) and our own logs.
try {
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [Console]::OutputEncoding = $utf8
    $OutputEncoding = $utf8
} catch {
    # Best-effort only
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# ============================================================================
# Localization / Language Support
# ============================================================================
$script:Strings = @{
    'de' = @{
        WindowTitle = 'WingetUpdateUI'
        MenuSettings = '_Einstellungen'
        MenuProxy = '_Proxy-Einstellungen...'
        MenuLanguage = '_Sprache'
        BtnRefresh = '⟳ Aktualisieren'
        BtnUpdate = '⬆ Update'
        BtnCancel = '✕ Abbrechen'
        BtnSelectAll = '☑ Alle'
        BtnSelectNone = '☐ Keine'
        ColName = 'Name'
        ColVersion = 'Version'
        ColAvailable = 'Verfügbar'
        StatusReady = 'Bereit'
        StatusLoading = 'Lade Updates...'
        StatusUpdating = 'Update läuft...'
        StatusError = 'Fehler'
        StatusLoadingDetails = 'Details laden: {0}/{1}'
        LogLoading = 'Lade winget Updates...'
        LogFoundUpdates = 'Gefundene Updates: {0}'
        LogDetailsLoaded = 'Details geladen: {0}'
        LogAllDetailsLoaded = 'Alle Details geladen.'
        LogNoSelection = 'Keine Programme ausgewählt.'
        LogStartingUpdates = 'Starte Updates für {0} Programme...'
        LogDone = 'Fertig.'
        LogCancelled = 'Update abgebrochen.'
        LogErrorLoading = 'Fehler beim Laden: '
        LogErrorUpdating = 'Fehler beim Updaten: '
        LogOperationRunning = 'Es läuft bereits eine Operation.'
        DetailPublisher = 'Herausgeber'
        DetailAuthor = 'Autor'
        DetailDescription = 'Beschreibung'
        DetailLicense = 'Lizenz'
        DetailHomepage = 'Homepage'
        DetailInstaller = 'Installer'
        DetailDownload = 'Download'
        DetailReleaseDate = 'Veröffentlicht'
        DetailLoading = 'Details werden geladen...'
        DetailInstalled = 'Installiert: '
        DetailAvailable = 'Verfügbar: '
        LinkLicense = 'Lizenz'
        LinkSupport = 'Support'
        LinkVersionInfo = 'Versionsinfo'
        ProxyTitle = 'Proxy-Einstellungen'
        ProxyEnable = 'Proxy aktivieren'
        ProxyServer = 'Server:'
        ProxyPort = 'Port:'
        ProxyType = 'Typ:'
        ProxyUser = 'Benutzer (optional):'
        ProxyPassword = 'Passwort (optional):'
        ProxySave = 'Speichern'
        ProxyCancel = 'Abbrechen'
        ProxyValidationServer = 'Bitte geben Sie einen gültigen Server ein.'
        ProxyValidationPort = 'Bitte geben Sie einen gültigen Port (1-65535) ein.'
        MenuHelp = '_Hilfe'
        MenuAbout = 'Über WingetUpdateUI...'
        AboutTitle = 'Über WingetUpdateUI'
        AboutVersion = 'Version {0}'
        AboutDescription = 'A minimalistic GUI wrapper for winget update.'
        AboutCopyright = '© 2026 WingetUpdateUI'
        AboutAuthor = 'Autor: ThisLimn0'
        AboutGitHub = 'GitHub'
        AboutOK = 'OK'
        MenuTimeFormat = 'Zeitformat'
        TimeFormatDateTime = 'Datum und Zeit'
        TimeFormatTimeOnly = 'Nur Zeit'
        TimeFormatNone = 'Kein Zeitstempel'
        TooltipLoading = 'Lade Details für {0}...'
        TooltipPublisher = 'Herausgeber: {0}'
        TooltipLicense = 'Lizenz: {0}'
        ExpanderLog = 'Log'
        MenuInstallMode = '_Installationsmodus'
        InstallModeNone = 'Standard (keine Option)'
        InstallModeSilent = 'Leise (--silent)'
        InstallModeInteractive = 'Interaktiv (--interactive)'
        MenuForce = 'Erzwingen (--force)'
        ForceTooltip = 'Erzwingt die Installation auch bei Fehlern'
        LogAllDetailsFromCache = 'Alle Details aus dem Cache geladen.'
        MenuView = '_Ansicht'
        MenuDarkMode = 'Dunkelmodus'
        ColStatus = 'Status'
        TooltipVersion = 'Version: {0} → {1}'
        UpdateOK = 'OK'
        UpdateError = 'Fehler ({0})'
        LogWingetNotFound = 'winget wurde nicht gefunden. Bitte installieren Sie Windows Package Manager.'
        LogStarting = 'Starte...'
        LogRetrying = 'Erneuter Versuch mit erhöhten Rechten...'
        DetailLoadError = '(Details konnten nicht geladen werden)'
    }
    'en' = @{
        WindowTitle = 'WingetUpdateUI'
        MenuSettings = '_Settings'
        MenuProxy = '_Proxy Settings...'
        MenuLanguage = '_Language'
        BtnRefresh = '⟳ Refresh'
        BtnUpdate = '⬆ Update'
        BtnCancel = '✕ Cancel'
        BtnSelectAll = '☑ All'
        BtnSelectNone = '☐ None'
        ColName = 'Name'
        ColVersion = 'Version'
        ColAvailable = 'Available'
        StatusReady = 'Ready'
        StatusLoading = 'Loading updates...'
        StatusUpdating = 'Updating...'
        StatusError = 'Error'
        StatusLoadingDetails = 'Loading details: {0}/{1}'
        LogLoading = 'Loading winget updates...'
        LogFoundUpdates = 'Found updates: {0}'
        LogDetailsLoaded = 'Details loaded: {0}'
        LogAllDetailsLoaded = 'All details loaded.'
        LogNoSelection = 'No programs selected.'
        LogStartingUpdates = 'Starting updates for {0} programs...'
        LogDone = 'Done.'
        LogCancelled = 'Update cancelled.'
        LogErrorLoading = 'Error loading: '
        LogErrorUpdating = 'Error updating: '
        LogOperationRunning = 'An operation is already running.'
        DetailPublisher = 'Publisher'
        DetailAuthor = 'Author'
        DetailDescription = 'Description'
        DetailLicense = 'License'
        DetailHomepage = 'Homepage'
        DetailInstaller = 'Installer'
        DetailDownload = 'Download'
        DetailReleaseDate = 'Release Date'
        DetailLoading = 'Loading details...'
        DetailInstalled = 'Installed: '
        DetailAvailable = 'Available: '
        LinkLicense = 'License'
        LinkSupport = 'Support'
        LinkVersionInfo = 'Version Info'
        ProxyTitle = 'Proxy Settings'
        ProxyEnable = 'Enable Proxy'
        ProxyServer = 'Server:'
        ProxyPort = 'Port:'
        ProxyType = 'Type:'
        ProxyUser = 'User (optional):'
        ProxyPassword = 'Password (optional):'
        ProxySave = 'Save'
        ProxyCancel = 'Cancel'
        ProxyValidationServer = 'Please enter a valid server.'
        ProxyValidationPort = 'Please enter a valid port (1-65535).'
        MenuHelp = '_Help'
        MenuAbout = 'About WingetUpdateUI...'
        AboutTitle = 'About WingetUpdateUI'
        AboutVersion = 'Version {0}'
        AboutDescription = 'A minimalistic GUI wrapper for winget update.'
        AboutCopyright = '© 2026 WingetUpdateUI'
        AboutAuthor = 'Author: ThisLimn0'
        AboutGitHub = 'GitHub'
        AboutOK = 'OK'
        MenuTimeFormat = 'Time Format'
        TimeFormatDateTime = 'Date and Time'
        TimeFormatTimeOnly = 'Time Only'
        TimeFormatNone = 'No Timestamp'
        TooltipLoading = 'Loading details for {0}...'
        TooltipPublisher = 'Publisher: {0}'
        TooltipLicense = 'License: {0}'
        ExpanderLog = 'Log'
        MenuInstallMode = '_Install Mode'
        InstallModeNone = 'Default (no option)'
        InstallModeSilent = 'Silent (--silent)'
        InstallModeInteractive = 'Interactive (--interactive)'
        MenuForce = 'Force (--force)'
        ForceTooltip = 'Forces the installation even on errors'
        LogAllDetailsFromCache = 'All details loaded from cache.'
        MenuView = '_View'
        MenuDarkMode = 'Dark Mode'
        ColStatus = 'Status'
        TooltipVersion = 'Version: {0} → {1}'
        UpdateOK = 'OK'
        UpdateError = 'Error ({0})'
        LogWingetNotFound = 'winget was not found. Please install Windows Package Manager.'
        LogStarting = 'Starting...'
        LogRetrying = 'Retrying with elevated privileges...'
        DetailLoadError = '(Details could not be loaded)'
    }
    'fr' = @{
        WindowTitle = 'WingetUpdateUI'
        MenuSettings = '_Paramètres'
        MenuProxy = '_Paramètres Proxy...'
        MenuLanguage = '_Langue'
        BtnRefresh = '⟳ Actualiser'
        BtnUpdate = '⬆ Mettre à jour'
        BtnCancel = '✕ Annuler'
        BtnSelectAll = '☑ Tous'
        BtnSelectNone = '☐ Aucun'
        ColName = 'Nom'
        ColVersion = 'Version'
        ColAvailable = 'Disponible'
        StatusReady = 'Prêt'
        StatusLoading = 'Chargement...'
        StatusUpdating = 'Mise à jour...'
        StatusError = 'Erreur'
        StatusLoadingDetails = 'Chargement: {0}/{1}'
        LogLoading = 'Chargement des mises à jour winget...'
        LogFoundUpdates = 'Mises à jour trouvées: {0}'
        LogDetailsLoaded = 'Détails chargés: {0}'
        LogAllDetailsLoaded = 'Tous les détails chargés.'
        LogNoSelection = 'Aucun programme sélectionné.'
        LogStartingUpdates = 'Démarrage des mises à jour pour {0} programmes...'
        LogDone = 'Terminé.'
        LogCancelled = 'Mise à jour annulée.'
        LogErrorLoading = 'Erreur de chargement: '
        LogErrorUpdating = 'Erreur de mise à jour: '
        LogOperationRunning = 'Une opération est déjà en cours.'
        DetailPublisher = 'Éditeur'
        DetailAuthor = 'Auteur'
        DetailDescription = 'Description'
        DetailLicense = 'Licence'
        DetailHomepage = 'Site web'
        DetailInstaller = 'Installeur'
        DetailDownload = 'Téléchargement'
        DetailReleaseDate = 'Date de sortie'
        DetailLoading = 'Chargement des détails...'
        DetailInstalled = 'Installé: '
        DetailAvailable = 'Disponible: '
        LinkLicense = 'Licence'
        LinkSupport = 'Support'
        LinkVersionInfo = 'Notes de version'
        ProxyTitle = 'Paramètres Proxy'
        ProxyEnable = 'Activer le proxy'
        ProxyServer = 'Serveur:'
        ProxyPort = 'Port:'
        ProxyType = 'Type:'
        ProxyUser = 'Utilisateur (optionnel):'
        ProxyPassword = 'Mot de passe (optionnel):'
        ProxySave = 'Enregistrer'
        ProxyCancel = 'Annuler'
        ProxyValidationServer = 'Veuillez entrer un serveur valide.'
        ProxyValidationPort = 'Veuillez entrer un port valide (1-65535).'
        MenuHelp = 'Aid_e'
        MenuAbout = 'À propos de WingetUpdateUI...'
        AboutTitle = 'À propos de WingetUpdateUI'
        AboutVersion = 'Version {0}'
        AboutDescription = 'A minimalistic GUI wrapper for winget update.'
        AboutCopyright = '© 2026 WingetUpdateUI'
        AboutAuthor = 'Auteur: ThisLimn0'
        AboutGitHub = 'GitHub'
        AboutOK = 'OK'
        MenuTimeFormat = 'Format de l''heure'
        TimeFormatDateTime = 'Date et heure'
        TimeFormatTimeOnly = 'Heure seulement'
        TimeFormatNone = 'Pas d''horodatage'
        TooltipLoading = 'Chargement des détails pour {0}...'
        TooltipPublisher = 'Éditeur: {0}'
        TooltipLicense = 'Licence: {0}'
        ExpanderLog = 'Journal'
        MenuInstallMode = '_Mode d''installation'
        InstallModeNone = 'Par défaut (aucune option)'
        InstallModeSilent = 'Silencieux (--silent)'
        InstallModeInteractive = 'Interactif (--interactive)'
        MenuForce = 'Forcer (--force)'
        ForceTooltip = 'Force l''installation même en cas d''erreurs'
        LogAllDetailsFromCache = 'Tous les détails chargés du cache.'
        MenuView = '_Affichage'
        MenuDarkMode = 'Mode sombre'
        ColStatus = 'Statut'
        TooltipVersion = 'Version: {0} → {1}'
        UpdateOK = 'OK'
        UpdateError = 'Erreur ({0})'
        LogWingetNotFound = 'winget introuvable. Veuillez installer Windows Package Manager.'
        LogStarting = 'Démarrage...'
        LogRetrying = 'Nouvelle tentative avec droits élevés...'
        DetailLoadError = '(Les détails n''ont pas pu être chargés)'
    }
    'es' = @{
        WindowTitle = 'WingetUpdateUI'
        MenuSettings = '_Configuración'
        MenuProxy = '_Configuración de Proxy...'
        MenuLanguage = '_Idioma'
        BtnRefresh = '⟳ Actualizar'
        BtnUpdate = '⬆ Actualizar'
        BtnCancel = '✕ Cancelar'
        BtnSelectAll = '☑ Todos'
        BtnSelectNone = '☐ Ninguno'
        ColName = 'Nombre'
        ColVersion = 'Versión'
        ColAvailable = 'Disponible'
        StatusReady = 'Listo'
        StatusLoading = 'Cargando...'
        StatusUpdating = 'Actualizando...'
        StatusError = 'Error'
        StatusLoadingDetails = 'Cargando: {0}/{1}'
        LogLoading = 'Cargando actualizaciones de winget...'
        LogFoundUpdates = 'Actualizaciones encontradas: {0}'
        LogDetailsLoaded = 'Detalles cargados: {0}'
        LogAllDetailsLoaded = 'Todos los detalles cargados.'
        LogNoSelection = 'Ningún programa seleccionado.'
        LogStartingUpdates = 'Iniciando actualizaciones para {0} programas...'
        LogDone = 'Hecho.'
        LogCancelled = 'Actualización cancelada.'
        LogErrorLoading = 'Error al cargar: '
        LogErrorUpdating = 'Error al actualizar: '
        LogOperationRunning = 'Ya hay una operación en curso.'
        DetailPublisher = 'Editor'
        DetailAuthor = 'Autor'
        DetailDescription = 'Descripción'
        DetailLicense = 'Licencia'
        DetailHomepage = 'Sitio web'
        DetailInstaller = 'Instalador'
        DetailDownload = 'Descarga'
        DetailReleaseDate = 'Fecha de lanzamiento'
        DetailLoading = 'Cargando detalles...'
        DetailInstalled = 'Instalado: '
        DetailAvailable = 'Disponible: '
        LinkLicense = 'Licencia'
        LinkSupport = 'Soporte'
        LinkVersionInfo = 'Info de versión'
        ProxyTitle = 'Configuración de Proxy'
        ProxyEnable = 'Habilitar proxy'
        ProxyServer = 'Servidor:'
        ProxyPort = 'Puerto:'
        ProxyType = 'Tipo:'
        ProxyUser = 'Usuario (opcional):'
        ProxyPassword = 'Contraseña (opcional):'
        ProxySave = 'Guardar'
        ProxyCancel = 'Cancelar'
        ProxyValidationServer = 'Por favor ingrese un servidor válido.'
        ProxyValidationPort = 'Por favor ingrese un puerto válido (1-65535).'
        MenuHelp = 'Ay_uda'
        MenuAbout = 'Acerca de WingetUpdateUI...'
        AboutTitle = 'Acerca de WingetUpdateUI'
        AboutVersion = 'Versión {0}'
        AboutDescription = 'A minimalistic GUI wrapper for winget update.'
        AboutCopyright = '© 2026 WingetUpdateUI'
        AboutAuthor = 'Autor: ThisLimn0'
        AboutGitHub = 'GitHub'
        AboutOK = 'Aceptar'
        MenuTimeFormat = 'Formato de hora'
        TimeFormatDateTime = 'Fecha y hora'
        TimeFormatTimeOnly = 'Solo hora'
        TimeFormatNone = 'Sin marca de tiempo'
        TooltipLoading = 'Cargando detalles para {0}...'
        TooltipPublisher = 'Editor: {0}'
        TooltipLicense = 'Licencia: {0}'
        ExpanderLog = 'Registro'
        MenuInstallMode = '_Modo de instalación'
        InstallModeNone = 'Predeterminado (sin opción)'
        InstallModeSilent = 'Silencioso (--silent)'
        InstallModeInteractive = 'Interactivo (--interactive)'
        MenuForce = 'Forzar (--force)'
        ForceTooltip = 'Fuerza la instalación incluso con errores'
        LogAllDetailsFromCache = 'Todos los detalles cargados del caché.'
        MenuView = '_Vista'
        MenuDarkMode = 'Modo oscuro'
        ColStatus = 'Estado'
        TooltipVersion = 'Versión: {0} → {1}'
        UpdateOK = 'OK'
        UpdateError = 'Error ({0})'
        LogWingetNotFound = 'winget no encontrado. Por favor instale Windows Package Manager.'
        LogStarting = 'Iniciando...'
        LogRetrying = 'Reintentando con privilegios elevados...'
        DetailLoadError = '(No se pudieron cargar los detalles)'
    }
    'it' = @{
        WindowTitle = 'WingetUpdateUI'
        MenuSettings = '_Impostazioni'
        MenuProxy = '_Impostazioni Proxy...'
        MenuLanguage = '_Lingua'
        BtnRefresh = '⟳ Aggiorna'
        BtnUpdate = '⬆ Aggiorna'
        BtnCancel = '✕ Annulla'
        BtnSelectAll = '☑ Tutti'
        BtnSelectNone = '☐ Nessuno'
        ColName = 'Nome'
        ColVersion = 'Versione'
        ColAvailable = 'Disponibile'
        StatusReady = 'Pronto'
        StatusLoading = 'Caricamento...'
        StatusUpdating = 'Aggiornamento...'
        StatusError = 'Errore'
        StatusLoadingDetails = 'Caricamento: {0}/{1}'
        LogLoading = 'Caricamento aggiornamenti winget...'
        LogFoundUpdates = 'Aggiornamenti trovati: {0}'
        LogDetailsLoaded = 'Dettagli caricati: {0}'
        LogAllDetailsLoaded = 'Tutti i dettagli caricati.'
        LogNoSelection = 'Nessun programma selezionato.'
        LogStartingUpdates = 'Avvio aggiornamenti per {0} programmi...'
        LogDone = 'Fatto.'
        LogCancelled = 'Aggiornamento annullato.'
        LogErrorLoading = 'Errore di caricamento: '
        LogErrorUpdating = 'Errore di aggiornamento: '
        LogOperationRunning = 'Un''operazione è già in corso.'
        DetailPublisher = 'Editore'
        DetailAuthor = 'Autore'
        DetailDescription = 'Descrizione'
        DetailLicense = 'Licenza'
        DetailHomepage = 'Sito web'
        DetailInstaller = 'Installer'
        DetailDownload = 'Scarica'
        DetailReleaseDate = 'Data di rilascio'
        DetailLoading = 'Caricamento dettagli...'
        DetailInstalled = 'Installato: '
        DetailAvailable = 'Disponibile: '
        LinkLicense = 'Licenza'
        LinkSupport = 'Supporto'
        LinkVersionInfo = 'Info versione'
        ProxyTitle = 'Impostazioni Proxy'
        ProxyEnable = 'Abilita proxy'
        ProxyServer = 'Server:'
        ProxyPort = 'Porta:'
        ProxyType = 'Tipo:'
        ProxyUser = 'Utente (opzionale):'
        ProxyPassword = 'Password (opzionale):'
        ProxySave = 'Salva'
        ProxyCancel = 'Annulla'
        ProxyValidationServer = 'Inserire un server valido.'
        ProxyValidationPort = 'Inserire una porta valida (1-65535).'
        MenuHelp = '_Aiuto'
        MenuAbout = 'Informazioni su WingetUpdateUI...'
        AboutTitle = 'Informazioni su WingetUpdateUI'
        AboutVersion = 'Versione {0}'
        AboutDescription = 'A minimalistic GUI wrapper for winget update.'
        AboutCopyright = '© 2026 WingetUpdateUI'
        AboutAuthor = 'Autore: ThisLimn0'
        AboutGitHub = 'GitHub'
        AboutOK = 'OK'
        MenuTimeFormat = 'Formato ora'
        TimeFormatDateTime = 'Data e ora'
        TimeFormatTimeOnly = 'Solo ora'
        TimeFormatNone = 'Nessun timestamp'
        TooltipLoading = 'Caricamento dettagli per {0}...'
        TooltipPublisher = 'Editore: {0}'
        TooltipLicense = 'Licenza: {0}'
        ExpanderLog = 'Log'
        MenuInstallMode = '_Modalità di installazione'
        InstallModeNone = 'Predefinito (nessuna opzione)'
        InstallModeSilent = 'Silenzioso (--silent)'
        InstallModeInteractive = 'Interattivo (--interactive)'
        MenuForce = 'Forza (--force)'
        ForceTooltip = 'Forza l''installazione anche in caso di errori'
        LogAllDetailsFromCache = 'Tutti i dettagli caricati dalla cache.'
        MenuView = '_Visualizza'
        MenuDarkMode = 'Modalità scura'
        ColStatus = 'Stato'
        TooltipVersion = 'Versione: {0} → {1}'
        UpdateOK = 'OK'
        UpdateError = 'Errore ({0})'
        LogWingetNotFound = 'winget non trovato. Installa Windows Package Manager.'
        LogStarting = 'Avvio...'
        LogRetrying = 'Nuovo tentativo con privilegi elevati...'
        DetailLoadError = '(Impossibile caricare i dettagli)'
    }
    'nl' = @{
        WindowTitle = 'WingetUpdateUI'
        MenuSettings = '_Instellingen'
        MenuProxy = '_Proxy-instellingen...'
        MenuLanguage = '_Taal'
        BtnRefresh = '⟳ Vernieuwen'
        BtnUpdate = '⬆ Bijwerken'
        BtnCancel = '✕ Annuleren'
        BtnSelectAll = '☑ Alles'
        BtnSelectNone = '☐ Geen'
        ColName = 'Naam'
        ColVersion = 'Versie'
        ColAvailable = 'Beschikbaar'
        StatusReady = 'Gereed'
        StatusLoading = 'Updates laden...'
        StatusUpdating = 'Bijwerken...'
        StatusError = 'Fout'
        StatusLoadingDetails = 'Details laden: {0}/{1}'
        LogLoading = 'Winget updates laden...'
        LogFoundUpdates = 'Gevonden updates: {0}'
        LogDetailsLoaded = 'Details geladen: {0}'
        LogAllDetailsLoaded = 'Alle details geladen.'
        LogNoSelection = 'Geen programma''s geselecteerd.'
        LogStartingUpdates = 'Updates starten voor {0} programma''s...'
        LogDone = 'Klaar.'
        LogCancelled = 'Update geannuleerd.'
        LogErrorLoading = 'Fout bij laden: '
        LogErrorUpdating = 'Fout bij bijwerken: '
        LogOperationRunning = 'Er loopt al een bewerking.'
        DetailPublisher = 'Uitgever'
        DetailAuthor = 'Auteur'
        DetailDescription = 'Beschrijving'
        DetailLicense = 'Licentie'
        DetailHomepage = 'Homepage'
        DetailInstaller = 'Installatieprogramma'
        DetailDownload = 'Download'
        DetailReleaseDate = 'Uitgebracht'
        DetailLoading = 'Details worden geladen...'
        DetailInstalled = 'Geïnstalleerd: '
        DetailAvailable = 'Beschikbaar: '
        LinkLicense = 'Licentie'
        LinkSupport = 'Ondersteuning'
        LinkVersionInfo = 'Versie-info'
        ProxyTitle = 'Proxy-instellingen'
        ProxyEnable = 'Proxy inschakelen'
        ProxyServer = 'Server:'
        ProxyPort = 'Poort:'
        ProxyType = 'Type:'
        ProxyUser = 'Gebruiker (optioneel):'
        ProxyPassword = 'Wachtwoord (optioneel):'
        ProxySave = 'Opslaan'
        ProxyCancel = 'Annuleren'
        ProxyValidationServer = 'Voer een geldige server in.'
        ProxyValidationPort = 'Voer een geldige poort in (1-65535).'
        MenuHelp = '_Help'
        MenuAbout = 'Over WingetUpdateUI...'
        AboutTitle = 'Over WingetUpdateUI'
        AboutVersion = 'Versie {0}'
        AboutDescription = 'A minimalistic GUI wrapper for winget update.'
        AboutCopyright = '© 2026 WingetUpdateUI'
        AboutAuthor = 'Auteur: ThisLimn0'
        AboutGitHub = 'GitHub'
        AboutOK = 'OK'
        MenuTimeFormat = 'Tijdformaat'
        TimeFormatDateTime = 'Datum en tijd'
        TimeFormatTimeOnly = 'Alleen tijd'
        TimeFormatNone = 'Geen tijdstempel'
        TooltipLoading = 'Details laden voor {0}...'
        TooltipPublisher = 'Uitgever: {0}'
        TooltipLicense = 'Licentie: {0}'
        ExpanderLog = 'Logboek'
        MenuInstallMode = '_Installatiemodus'
        InstallModeNone = 'Standaard (geen optie)'
        InstallModeSilent = 'Stil (--silent)'
        InstallModeInteractive = 'Interactief (--interactive)'
        MenuForce = 'Forceren (--force)'
        ForceTooltip = 'Forceert de installatie zelfs bij fouten'
        LogAllDetailsFromCache = 'Alle details uit cache geladen.'
        MenuView = '_Weergave'
        MenuDarkMode = 'Donkere modus'
        ColStatus = 'Status'
        TooltipVersion = 'Versie: {0} → {1}'
        UpdateOK = 'OK'
        UpdateError = 'Fout ({0})'
        LogWingetNotFound = 'winget niet gevonden. Installeer Windows Package Manager.'
        LogStarting = 'Starten...'
        LogRetrying = 'Opnieuw proberen met verhoogde rechten...'
        DetailLoadError = '(Details konden niet worden geladen)'
    }
}

# Current language
$script:CurrentLanguage = 'en'

# Settings file path
$script:SettingsPath = Join-Path $env:APPDATA 'WingetUpdateUI\settings.json'

# Proxy settings
$script:ProxySettings = @{
    Enabled = $false
    Server = ''
    Port = ''
    Type = 'HTTP'
    User = ''
    Password = ''
}

# Installation mode: 'none', 'silent', 'interactive'
$script:InstallMode = 'none'

# Time format for log: 'datetime', 'timeonly', 'none'
$script:TimeFormat = 'datetime'

# Force installation flag
$script:ForceInstall = $false

# Dark mode setting
$script:DarkMode = $false

function Get-String {
    param([string] $Key)
    $lang = $script:CurrentLanguage
    if ($script:Strings[$lang] -and $script:Strings[$lang][$Key]) {
        return $script:Strings[$lang][$Key]
    }
    # Fallback to English
    if ($script:Strings['en'][$Key]) {
        return $script:Strings['en'][$Key]
    }
    return $Key
}

function Save-Settings {
    # Create proxy settings copy with encrypted password
    $proxyToSave = @{
        Enabled = $script:ProxySettings.Enabled
        Server = $script:ProxySettings.Server
        Port = $script:ProxySettings.Port
        Type = $script:ProxySettings.Type
        User = $script:ProxySettings.User
        EncryptedPassword = ''
    }
    
    # Encrypt password using DPAPI (Windows Data Protection API)
    if (-not [string]::IsNullOrWhiteSpace($script:ProxySettings.Password)) {
        try {
            $secureString = ConvertTo-SecureString -String $script:ProxySettings.Password -AsPlainText -Force
            $proxyToSave.EncryptedPassword = ConvertFrom-SecureString -SecureString $secureString
        } catch {
            # If encryption fails, don't save password
            $proxyToSave.EncryptedPassword = ''
        }
    }
    
    $settings = @{
        Language = $script:CurrentLanguage
        Proxy = $proxyToSave
        InstallMode = $script:InstallMode
        ForceInstall = $script:ForceInstall
        DarkMode = $script:DarkMode
        TimeFormat = $script:TimeFormat
    }
    
    $dir = Split-Path $script:SettingsPath -Parent
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    
    $settings | ConvertTo-Json | Set-Content -Path $script:SettingsPath -Encoding UTF8
}

function Import-Settings {
    if (Test-Path $script:SettingsPath) {
        try {
            $settings = Get-Content -Path $script:SettingsPath -Raw | ConvertFrom-Json
            if ($settings.Language -and $script:Strings.ContainsKey($settings.Language)) {
                $script:CurrentLanguage = $settings.Language
            }
            if ($settings.Proxy) {
                $script:ProxySettings.Enabled = [bool]$settings.Proxy.Enabled
                $script:ProxySettings.Server = [string]$settings.Proxy.Server
                $script:ProxySettings.Port = [string]$settings.Proxy.Port
                $script:ProxySettings.User = [string]$settings.Proxy.User
                
                # Decrypt password using DPAPI
                $script:ProxySettings.Password = ''
                if (-not [string]::IsNullOrWhiteSpace($settings.Proxy.EncryptedPassword)) {
                    try {
                        $secureString = ConvertTo-SecureString -String $settings.Proxy.EncryptedPassword
                        $script:ProxySettings.Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto(
                            [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString)
                        )
                    } catch {
                        # Decryption failed (different user/machine), ignore
                        $script:ProxySettings.Password = ''
                    }
                } elseif (-not [string]::IsNullOrWhiteSpace($settings.Proxy.Password)) {
                    # Legacy: plain text password (migrate on next save)
                    $script:ProxySettings.Password = [string]$settings.Proxy.Password
                }
            }
            if ($settings.InstallMode -and $settings.InstallMode -in @('none','silent','interactive')) {
                $script:InstallMode = $settings.InstallMode
            }
            if ($null -ne $settings.ForceInstall) {
                $script:ForceInstall = [bool]$settings.ForceInstall
            }
            if ($null -ne $settings.DarkMode) {
                $script:DarkMode = [bool]$settings.DarkMode
            }
            if ($settings.TimeFormat -and $settings.TimeFormat -in @('datetime','timeonly','none')) {
                $script:TimeFormat = $settings.TimeFormat
            }
        } catch {
            # Ignore errors, use defaults
        }
    }
}

function Set-ProxySettings {
    if ($script:ProxySettings.Enabled -and -not [string]::IsNullOrWhiteSpace($script:ProxySettings.Server)) {
        $proxyUri = "http://$($script:ProxySettings.Server)"
        if (-not [string]::IsNullOrWhiteSpace($script:ProxySettings.Port)) {
            $proxyUri += ":$($script:ProxySettings.Port)"
        }
        
        $proxy = New-Object System.Net.WebProxy($proxyUri, $true)
        
        if (-not [string]::IsNullOrWhiteSpace($script:ProxySettings.User)) {
            $proxy.Credentials = New-Object System.Net.NetworkCredential(
                $script:ProxySettings.User,
                $script:ProxySettings.Password
            )
        }
        
        [System.Net.WebRequest]::DefaultWebProxy = $proxy
        $env:HTTP_PROXY = $proxyUri
        $env:HTTPS_PROXY = $proxyUri
    } else {
        [System.Net.WebRequest]::DefaultWebProxy = $null
        Remove-Item Env:HTTP_PROXY -ErrorAction SilentlyContinue
        Remove-Item Env:HTTPS_PROXY -ErrorAction SilentlyContinue
    }
}

# Load settings at startup
Import-Settings
Set-ProxySettings

# Check if winget is available
function Test-WingetAvailable {
    try {
        $null = Get-Command winget -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

if (-not (Test-WingetAvailable)) {
    [System.Windows.MessageBox]::Show(
        (Get-String 'LogWingetNotFound'),
        'WingetUpdateUI',
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error
    ) | Out-Null
    exit 1
}

Add-Type -TypeDefinition @"
using System;
using System.ComponentModel;

public class WingetPackage : INotifyPropertyChanged
{
    public string Name { get; set; }
    public string Id { get; set; }
    public string Version { get; set; }
    public string AvailableVersion { get; set; }
    public string Source { get; set; }
    
    // Enriched details from winget show
    public string Publisher { get; set; }
    public string Author { get; set; }
    public string Description { get; set; }
    public string License { get; set; }
    public string LicenseUrl { get; set; }
    public string SupportUrl { get; set; }
    public string ReleaseNotesUrl { get; set; }
    public string Homepage { get; set; }
    public string InstallerType { get; set; }
    public string InstallerUrl { get; set; }
    public string ReleaseDate { get; set; }
    public bool DetailsLoaded { get; set; }

    private bool _selected;
    public bool Selected
    {
        get { return _selected; }
        set
        {
            if (_selected != value)
            {
                _selected = value;
                OnPropertyChanged("Selected");
            }
        }
    }

    private string _tooltipText;
    public string TooltipText
    {
        get { return _tooltipText ?? "..."; }
        set
        {
            if (_tooltipText != value)
            {
                _tooltipText = value;
                OnPropertyChanged("TooltipText");
            }
        }
    }

    // Update status: empty, "✓", or error message
    private string _updateStatus = "";
    public string UpdateStatus
    {
        get { return _updateStatus; }
        set
        {
            if (_updateStatus != value)
            {
                _updateStatus = value;
                OnPropertyChanged("UpdateStatus");
            }
        }
    }

    public event PropertyChangedEventHandler PropertyChanged;
    
    private void OnPropertyChanged(string propertyName)
    {
        if (PropertyChanged != null)
        {
            PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
"@

function Add-LogLine {
    param(
        [Parameter(Mandatory)] [System.Windows.Controls.TextBox] $LogBox,
        [Parameter(Mandatory)] [string] $Message
    )

    $line = switch ($script:TimeFormat) {
        'datetime' { "[" + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + "] $Message" }
        'timeonly' { "[" + (Get-Date).ToString('HH:mm:ss') + "] $Message" }
        'none'     { $Message }
        default    { "[" + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + "] $Message" }
    }

    $LogBox.Dispatcher.Invoke([Action] {
        if ([string]::IsNullOrWhiteSpace($LogBox.Text)) {
            $LogBox.Text = $line
        } else {
            $LogBox.AppendText([Environment]::NewLine + $line)
        }
        $LogBox.UpdateLayout()
        $LogBox.ScrollToEnd()
    })
}

function Set-Busy {
    param(
        [Parameter(Mandatory)] [System.Windows.Controls.ProgressBar] $BusyBar,
        [Parameter(Mandatory)] [System.Windows.Controls.TextBlock] $StatusText,
        [Parameter(Mandatory)] [bool] $IsBusy,
        [string] $Status = $null
    )
    
    if ($null -eq $Status) { $Status = Get-String 'StatusReady' }

    $BusyBar.Dispatcher.Invoke([Action] {
        $BusyBar.IsIndeterminate = $IsBusy
        $StatusText.Text = $Status
    })
}

function Get-WingetUpgradesJson {
    # winget must be in PATH
    $winget = Get-Command winget -ErrorAction Stop
    $raw = & $winget.Source upgrade --output json --disable-interactivity 2>&1
    $rawText = ($raw | ForEach-Object { [string]$_ }) -join "`n"

    if ($LASTEXITCODE -ne 0) {
        throw "winget upgrade failed (exit $LASTEXITCODE): $rawText"
    }

    try {
        return $rawText | ConvertFrom-Json -Depth 20
    } catch {
        throw "Failed to parse winget JSON output. Output was: $rawText"
    }
}

function Convert-WingetJsonToPackages {
    param(
        [Parameter(Mandatory)] $WingetJson
    )

    $packages = New-Object System.Collections.Generic.List[WingetPackage]

    # winget schema typically: { Sources: [ { Name, Packages: [ { PackageIdentifier, PackageName, InstalledVersion, AvailableVersion, Source } ] } ] }
    $sources = @($WingetJson.Sources)

    foreach ($src in $sources) {
        $srcName = [string]$src.Name
        foreach ($p in @($src.Packages)) {
            $item = [WingetPackage]::new()
            $item.Name = [string]$p.PackageName
            $item.Id = [string]$p.PackageIdentifier
            $item.Version = [string]$p.InstalledVersion
            $item.AvailableVersion = [string]$p.AvailableVersion
            $item.Source = if (-not [string]::IsNullOrWhiteSpace([string]$p.Source)) { [string]$p.Source } else { $srcName }
            $item.Selected = $false
            $packages.Add($item)
        }
    }

    return $packages
}

# Load XAML
$root = Split-Path -Parent $PSCommandPath
$xamlPath = Join-Path $root 'ui\MainWindow.xaml'

if (-not (Test-Path $xamlPath)) {
    throw "XAML not found: $xamlPath"
}

[xml]$xaml = Get-Content -LiteralPath $xamlPath -Raw
$reader = New-Object System.Xml.XmlNodeReader($xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Find controls
$RefreshButton = $window.FindName('RefreshButton')
$UpdateSelectedButton = $window.FindName('UpdateSelectedButton')
$CancelButton = $window.FindName('CancelButton')
$SelectAllButton = $window.FindName('SelectAllButton')
$SelectNoneButton = $window.FindName('SelectNoneButton')
$PackagesGrid = $window.FindName('PackagesGrid')
$BusyBar = $window.FindName('BusyBar')
$StatusText = $window.FindName('StatusText')
$LogBox = $window.FindName('LogBox')
$ButtonDivider = $window.FindName('ButtonDivider')

# Menu controls
$MenuSettings = $window.FindName('MenuSettings')
$MenuProxy = $window.FindName('MenuProxy')
$MenuLanguage = $window.FindName('MenuLanguage')
$MenuLangDE = $window.FindName('MenuLangDE')
$MenuLangEN = $window.FindName('MenuLangEN')
$MenuLangFR = $window.FindName('MenuLangFR')
$MenuLangES = $window.FindName('MenuLangES')
$MenuLangIT = $window.FindName('MenuLangIT')
$MenuLangNL = $window.FindName('MenuLangNL')
$MenuInstallMode = $window.FindName('MenuInstallMode')
$MenuModeNone = $window.FindName('MenuModeNone')
$MenuModeSilent = $window.FindName('MenuModeSilent')
$MenuModeInteractive = $window.FindName('MenuModeInteractive')
$MenuForce = $window.FindName('MenuForce')

# View menu controls
$MenuView = $window.FindName('MenuView')
$MenuDarkMode = $window.FindName('MenuDarkMode')
$MenuTimeFormat = $window.FindName('MenuTimeFormat')
$MenuTimeDateTime = $window.FindName('MenuTimeDateTime')
$MenuTimeOnly = $window.FindName('MenuTimeOnly')
$MenuTimeNone = $window.FindName('MenuTimeNone')

# Main layout controls for theming
$RootPanel = $window.FindName('RootPanel')
$MainGrid = $window.FindName('MainGrid')
$TopMenu = $window.FindName('TopMenu')

# Column name references
$ColName = $window.FindName('ColName')
$ColVersion = $window.FindName('ColVersion')
$ColAvailable = $window.FindName('ColAvailable')
$ColStatus = $window.FindName('ColStatus')

# Log Panel and Toggle Button
$LogToggleButton = $window.FindName('LogToggleButton')
$LogPanel = $window.FindName('LogPanel')

# Detail panel controls
$DetailPanel = $window.FindName('DetailPanel')
$DetailSplitter = $window.FindName('DetailSplitter')
$DetailTitle = $window.FindName('DetailTitle')
$DetailId = $window.FindName('DetailId')
$DetailVersion = $window.FindName('DetailVersion')
$DetailAvailable = $window.FindName('DetailAvailable')
$DetailLoading = $window.FindName('DetailLoading')
$DetailContent = $window.FindName('DetailContent')
$DetailPublisher = $window.FindName('DetailPublisher')
$DetailAuthor = $window.FindName('DetailAuthor')
$DetailDescription = $window.FindName('DetailDescription')
$DetailLicense = $window.FindName('DetailLicense')
$DetailHomepage = $window.FindName('DetailHomepage')
$DetailInstallerType = $window.FindName('DetailInstallerType')
$DetailInstallerUrl = $window.FindName('DetailInstallerUrl')
$DetailReleaseDate = $window.FindName('DetailReleaseDate')

# Quick Links panel controls
$QuickLinksPanel = $window.FindName('QuickLinksPanel')
$LinkLicense = $window.FindName('LinkLicense')
$LinkSeparator1 = $window.FindName('LinkSeparator1')
$LinkSupport = $window.FindName('LinkSupport')
$LinkSeparator2 = $window.FindName('LinkSeparator2')
$LinkReleaseNotes = $window.FindName('LinkReleaseNotes')

# Detail panel labels for localization
$LabelInstalled = $window.FindName('LabelInstalled')
$LabelAvailable = $window.FindName('LabelAvailable')
$LabelPublisher = $window.FindName('LabelPublisher')
$LabelAuthor = $window.FindName('LabelAuthor')
$LabelDescription = $window.FindName('LabelDescription')
$LabelLicense = $window.FindName('LabelLicense')
$LabelHomepage = $window.FindName('LabelHomepage')
$LabelInstaller = $window.FindName('LabelInstaller')
$LabelInstallerUrl = $window.FindName('LabelInstallerUrl')
$LabelReleaseDate = $window.FindName('LabelReleaseDate')

# Detail panel labels (find by walking the tree)
$script:DetailLabels = @{}

# Get the detail column definition for resizing
$mainGrid = $DetailPanel.Parent
$DetailColumnDef = $mainGrid.ColumnDefinitions[2]

# Data source
$Packages = [System.Collections.ObjectModel.ObservableCollection[WingetPackage]]::new()
$PackagesGrid.ItemsSource = $Packages

# Cache for enriched package details (Id -> details hashtable)
$script:PackageDetailsCache = @{}

# Background enrichment state - support for 4 parallel workers
$script:EnrichmentWorkers = @($null, $null, $null, $null)
$script:EnrichmentQueue = [System.Collections.Generic.Queue[string]]::new()
$script:MaxEnrichmentWorkers = 4

# Async operation state (run in a dedicated runspace; UI thread polls completion)
$script:PendingOp = $null
$script:CancelRequested = $false
$script:UpdateProgress = $null
$script:UpdateLastProcessedIndex = -1

# Cancel the current update operation
function Stop-UpdateOperation {
    if ($null -eq $script:PendingOp) { return }
    if ($script:PendingOp.Kind -ne 'update') { return }
    
    $script:CancelRequested = $true
    
    # Signal cancel to the update progress
    if ($null -ne $script:UpdateProgress) {
        $script:UpdateProgress.Cancelled = $true
    }
    
    try {
        # Stop the PowerShell runspace
        $script:PendingOp.PS.Stop()
    } catch {
        # Ignore errors during stop
    }
    
    try {
        $script:PendingOp.PS.Dispose()
    } catch {}
    try {
        $script:PendingOp.Runspace.Close()
        $script:PendingOp.Runspace.Dispose()
    } catch {}
    
    $script:PendingOp = $null
    $script:UpdateProgress = $null
    $script:UpdateLastProcessedIndex = -1
    $BusyBar.Value = 0
    
    Add-LogLine -LogBox $LogBox -Message (Get-String 'LogCancelled')
    Set-Busy -BusyBar $BusyBar -StatusText $StatusText -IsBusy:$false -Status (Get-String 'StatusReady')
    Set-ButtonsEnabled -enabled:$true
    $CancelButton.Visibility = 'Collapsed'
    
    $script:CancelRequested = $false
}

# Helper to log errors to console
function Write-ConsoleError {
    param([string] $Message)
    try {
        [Console]::Error.WriteLine("[ERROR] $Message")
    } catch {
        # Ignore console errors
    }
}

function Start-AsyncOperation {
    param(
        [Parameter(Mandatory)] [ValidateSet('refresh','update')] [string] $Kind,
        [Parameter(Mandatory)] [string] $ScriptText,
        [object[]] $Arguments = @()
    )

    if ($null -ne $script:PendingOp) {
        Add-LogLine -LogBox $LogBox -Message (Get-String 'LogOperationRunning')
        return $false
    }

    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $rs.Open()

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    $null = $ps.AddScript($ScriptText)
    foreach ($a in $Arguments) { $null = $ps.AddArgument($a) }

    $async = $ps.BeginInvoke()
    $script:PendingOp = [pscustomobject]@{
        Kind    = $Kind
        Runspace= $rs
        PS      = $ps
        Async   = $async
    }
    return $true
}

function Complete-PendingOperation {
    if ($null -eq $script:PendingOp) { return }
    if (-not $script:PendingOp.Async.IsCompleted) { return }

    $kind = $script:PendingOp.Kind
    $ps = $script:PendingOp.PS
    $rs = $script:PendingOp.Runspace

    try {
        $result = $ps.EndInvoke($script:PendingOp.Async)

        if ($ps.Streams.Error.Count -gt 0) {
            $msg = ($ps.Streams.Error | ForEach-Object { $_.ToString() }) -join "`n"
            Write-ConsoleError $msg
            throw $msg
        }

        switch ($kind) {
            'refresh' {
                $Packages.Clear()
                $script:EnrichmentQueue.Clear()
                
                # Count how many packages need enrichment
                $needsEnrichment = 0
                
                foreach ($row in @($result)) {
                    $p = [WingetPackage]::new()
                    $p.Name = [string]$row.Name
                    $p.Id = [string]$row.Id
                    $p.Version = [string]$row.Version
                    $p.AvailableVersion = [string]$row.AvailableVersion
                    $p.Source = [string]$row.Source
                    $p.Selected = $true
                    
                    # Check if we already have cached details for this package
                    $pkgId = [string]$row.Id
                    if ($script:PackageDetailsCache.ContainsKey($pkgId)) {
                        $cached = $script:PackageDetailsCache[$pkgId]
                        $p.Publisher = [string]$cached.Publisher
                        $p.Author = [string]$cached.Author
                        $p.Description = [string]$cached.Description
                        $p.License = [string]$cached.License
                        $p.LicenseUrl = [string]$cached.LicenseUrl
                        $p.SupportUrl = [string]$cached.SupportUrl
                        $p.ReleaseNotesUrl = [string]$cached.ReleaseNotesUrl
                        $p.Homepage = [string]$cached.Homepage
                        $p.InstallerType = [string]$cached.InstallerType
                        $p.InstallerUrl = [string]$cached.InstallerUrl
                        $p.ReleaseDate = [string]$cached.ReleaseDate
                        $p.DetailsLoaded = $true
                        
                        # Build tooltip from cache
                        $tooltipLines = @()
                        $tooltipLines += "$($p.Name) ($($p.Id))"
                        $tooltipLines += ((Get-String 'TooltipVersion') -f $p.Version, $p.AvailableVersion)
                        if (-not [string]::IsNullOrWhiteSpace($p.Publisher)) { $tooltipLines += ((Get-String 'TooltipPublisher') -f $p.Publisher) }
                        if (-not [string]::IsNullOrWhiteSpace($p.License)) { $tooltipLines += ((Get-String 'TooltipLicense') -f $p.License) }
                        if (-not [string]::IsNullOrWhiteSpace($p.Description)) {
                            $desc = $p.Description
                            if ($desc.Length -gt 200) { $desc = $desc.Substring(0, 200) + "..." }
                            $tooltipLines += "`n$desc"
                        }
                        $p.TooltipText = ($tooltipLines -join "`n")
                    } else {
                        $p.TooltipText = "Lade Details für $($p.Name)..."
                        # Queue for background enrichment
                        $script:EnrichmentQueue.Enqueue($pkgId)
                        $needsEnrichment++
                    }
                    
                    $Packages.Add($p)
                }

                Add-LogLine -LogBox $LogBox -Message ((Get-String 'LogFoundUpdates') -f $Packages.Count)
                Set-ButtonsEnabled -enabled:$true
                
                # Initialize enrichment progress tracking
                $script:EnrichmentTotal = $needsEnrichment
                $script:EnrichmentDone = ($Packages.Count - $needsEnrichment)
                
                if ($needsEnrichment -gt 0) {
                    # Show enrichment progress
                    $BusyBar.IsIndeterminate = $false
                    $BusyBar.Maximum = $Packages.Count
                    $BusyBar.Value = $script:EnrichmentDone
                    $StatusText.Text = (Get-String 'StatusLoadingDetails') -f $script:EnrichmentDone, $Packages.Count
                    
                    # Start background enrichment
                    Start-NextEnrichment
                } else {
                    # All details already cached
                    Add-LogLine -LogBox $LogBox -Message (Get-String 'LogAllDetailsFromCache')
                    Set-Busy -BusyBar $BusyBar -StatusText $StatusText -IsBusy:$false -Status (Get-String 'StatusReady')
                }
            }
            'update' {
                # Hide cancel button
                $CancelButton.Visibility = 'Collapsed'
                
                # Process results from the progress hashtable
                if ($null -ne $script:UpdateProgress) {
                    foreach ($r in @($script:UpdateProgress.Results)) {
                        # Update the package status in the UI
                        $pkg = $Packages | Where-Object { $_.Id -eq $r.Id } | Select-Object -First 1
                        if ($null -ne $pkg) {
                            if ($r.ExitCode -eq 0) {
                                $pkg.UpdateStatus = "✓"
                                Add-LogLine -LogBox $LogBox -Message ("{0} {1} ({2}) ExitCode={3}" -f $r.Name, $pkg.AvailableVersion, $r.Id, $r.ExitCode)
                                Add-LogLine -LogBox $LogBox -Message ("successfully updated to version {0}." -f $pkg.AvailableVersion)
                            } else {
                                # Store full error message for tooltip display
                                $errMsg = if (-not [string]::IsNullOrWhiteSpace($r.Error)) { $r.Error } else { "Exit: $($r.ExitCode)" }
                                $pkg.UpdateStatus = $errMsg
                                Add-LogLine -LogBox $LogBox -Message ("{0} {1} ({2}) ExitCode={3}" -f $r.Name, $pkg.AvailableVersion, $r.Id, $r.ExitCode)
                            }
                        } else {
                            Add-LogLine -LogBox $LogBox -Message ("{0} ({1}) ExitCode={2}" -f $r.Name, $r.Id, $r.ExitCode)
                        }
                    }
                    $script:UpdateProgress = $null
                }

                Add-LogLine -LogBox $LogBox -Message (Get-String 'LogDone')
                $BusyBar.Value = 0
                Set-Busy -BusyBar $BusyBar -StatusText $StatusText -IsBusy:$false -Status (Get-String 'StatusReady')
                Set-ButtonsEnabled -enabled:$true

                # Clean up resources BEFORE refresh
                try { $ps.Dispose() } catch {}
                try { $rs.Close() } catch {}
                try { $rs.Dispose() } catch {}
                $script:PendingOp = $null
                
                # refresh list after updates (delayed to avoid race condition)
                $window.Dispatcher.BeginInvoke([Action]{
                    Update-PackageList
                }, [System.Windows.Threading.DispatcherPriority]::Background)
                
                # Return early to skip finally block cleanup (already done)
                return
            }
        }
    } catch {
        $err = $_.Exception.Message
        if ([string]::IsNullOrWhiteSpace($err)) { $err = $_.ToString() }
        Write-ConsoleError "[$kind] $err"

        if ($kind -eq 'refresh') {
            Add-LogLine -LogBox $LogBox -Message ((Get-String 'LogErrorLoading') + $err)
            Set-Busy -BusyBar $BusyBar -StatusText $StatusText -IsBusy:$false -Status (Get-String 'StatusError')
        } else {
            Add-LogLine -LogBox $LogBox -Message ((Get-String 'LogErrorUpdating') + $err)
            Set-Busy -BusyBar $BusyBar -StatusText $StatusText -IsBusy:$false -Status (Get-String 'StatusError')
        }

        Set-ButtonsEnabled -enabled:$true
        $script:PendingOp = $null
    } finally {
        # Only cleanup if not already done (for update success path)
        if ($null -ne $script:PendingOp) {
            try { $ps.Dispose() } catch {}
            try { $rs.Close() } catch {}
            try { $rs.Dispose() } catch {}
            $script:PendingOp = $null
        }
    }
}

# Enrichment completion handler - checks all workers
function Complete-EnrichmentOperation {
    for ($workerIdx = 0; $workerIdx -lt $script:MaxEnrichmentWorkers; $workerIdx++) {
        $worker = $script:EnrichmentWorkers[$workerIdx]
        if ($null -eq $worker) { continue }
        if (-not $worker.Async.IsCompleted) { continue }

        $ps = $worker.PS
        $rs = $worker.Runspace
        # PackageId is stored in worker but used from result

        try {
            $result = $ps.EndInvoke($worker.Async)

            if ($ps.Streams.Error.Count -gt 0) {
                $errMsg = ($ps.Streams.Error | ForEach-Object { $_.ToString() }) -join "`n"
                Write-ConsoleError "[enrich] $errMsg"
            }

            if ($result -and $result.Count -gt 0) {
                $details = $result[0]
                $loadedPkgId = [string]$details.Id
                
                # Cache the details
                $script:PackageDetailsCache[$loadedPkgId] = $details
                
                # Update the package in the collection
                foreach ($pkg in $Packages) {
                    if ($pkg.Id -eq $loadedPkgId) {
                        $pkg.Publisher = [string]$details.Publisher
                        $pkg.Author = [string]$details.Author
                        $pkg.Description = [string]$details.Description
                        $pkg.License = [string]$details.License
                        $pkg.LicenseUrl = [string]$details.LicenseUrl
                        $pkg.SupportUrl = [string]$details.SupportUrl
                        $pkg.ReleaseNotesUrl = [string]$details.ReleaseNotesUrl
                        $pkg.Homepage = [string]$details.Homepage
                        $pkg.InstallerType = [string]$details.InstallerType
                        $pkg.InstallerUrl = [string]$details.InstallerUrl
                        $pkg.ReleaseDate = [string]$details.ReleaseDate
                        $pkg.DetailsLoaded = $true
                        
                        # Build tooltip text
                        $tooltipLines = @()
                        $tooltipLines += "$($pkg.Name) ($($pkg.Id))"
                        $tooltipLines += (Get-String 'TooltipVersion') -f $pkg.Version, $pkg.AvailableVersion
                        if (-not [string]::IsNullOrWhiteSpace($pkg.Publisher)) { $tooltipLines += (Get-String 'TooltipPublisher') -f $pkg.Publisher }
                        if (-not [string]::IsNullOrWhiteSpace($pkg.License)) { $tooltipLines += (Get-String 'TooltipLicense') -f $pkg.License }
                        if (-not [string]::IsNullOrWhiteSpace($pkg.Description)) {
                            $desc = $pkg.Description
                            if ($desc.Length -gt 200) { $desc = $desc.Substring(0, 200) + "..." }
                            $tooltipLines += "`n$desc"
                        }
                        $pkg.TooltipText = ($tooltipLines -join "`n")
                        
                        # Log the loaded details
                        Add-LogLine -LogBox $LogBox -Message ((Get-String 'LogDetailsLoaded') -f $pkg.Name)
                        
                        # If this package is currently selected, update detail panel
                        $selectedItem = $PackagesGrid.SelectedItem
                        if ($null -ne $selectedItem -and $selectedItem.Id -eq $loadedPkgId) {
                            Update-DetailPanel -Package $pkg
                        }
                        break
                    }
                }
            }
        } catch {
            Write-ConsoleError "[enrich] $($_.Exception.Message)"
        } finally {
            try { $ps.Dispose() } catch {}
            try { $rs.Close() } catch {}
            try { $rs.Dispose() } catch {}
            $script:EnrichmentWorkers[$workerIdx] = $null
            
            # Update enrichment progress (only if no update is running)
            $script:EnrichmentDone++
            if ($null -eq $script:PendingOp -or $script:PendingOp.Kind -ne 'update') {
                $BusyBar.Value = $script:EnrichmentDone
                
                # Check if all workers are done and queue is empty
                $activeWorkers = @($script:EnrichmentWorkers | Where-Object { $null -ne $_ }).Count
                if ($script:EnrichmentQueue.Count -eq 0 -and $activeWorkers -eq 0) {
                    # Enrichment complete
                    $BusyBar.IsIndeterminate = $false
                    $BusyBar.Value = 0
                    $StatusText.Text = Get-String 'StatusReady'
                    Add-LogLine -LogBox $LogBox -Message (Get-String 'LogAllDetailsLoaded')
                } else {
                    $total = $Packages.Count
                    $StatusText.Text = (Get-String 'StatusLoadingDetails') -f $script:EnrichmentDone, $total
                }
            }
            
            # Start next item in this worker slot
            Start-NextEnrichment
        }
    }
}

# Update progress display for running updates
function Update-UpdateProgress {
    if ($null -eq $script:UpdateProgress) { return }
    if ($null -eq $script:PendingOp -or $script:PendingOp.Kind -ne 'update') { return }
    
    $progress = $script:UpdateProgress
    
    # Update progress bar with current percentage
    $BusyBar.Value = $progress.CurrentPercent
    
    # Update status text with current package info
    $pkgIndex = $progress.CurrentPackageIndex + 1
    $total = $progress.TotalPackages
    $pkgName = $progress.CurrentPackageName
    $percent = $progress.CurrentPercent
    
    if (-not [string]::IsNullOrWhiteSpace($pkgName)) {
        $StatusText.Text = "Update $pkgIndex/$total`: $pkgName ($percent%)"
    }
    
    # Update status for completed packages
    $resultsCount = $progress.Results.Count
    $lastProcessedIndex = -1
    if ($null -ne $script:UpdateLastProcessedIndex) {
        $lastProcessedIndex = $script:UpdateLastProcessedIndex
    }
    
    for ($i = $lastProcessedIndex + 1; $i -lt $resultsCount; $i++) {
        $r = $progress.Results[$i]
        $pkg = $Packages | Where-Object { $_.Id -eq $r.Id } | Select-Object -First 1
        $availVer = if ($null -ne $pkg) { $pkg.AvailableVersion } else { '' }
        if ($null -ne $pkg) {
            if ($r.ExitCode -eq 0) {
                Add-LogLine -LogBox $LogBox -Message ("{0} {1} ({2}) ExitCode={3}" -f $r.Name, $availVer, $r.Id, $r.ExitCode)
                Add-LogLine -LogBox $LogBox -Message ("successfully updated to version {0}." -f $availVer)
                # Remove successfully updated package from list
                $Window.Dispatcher.Invoke([Action] {
                    $Packages.Remove($pkg) | Out-Null
                })
            } else {
                Add-LogLine -LogBox $LogBox -Message ("{0} {1} ({2}) ExitCode={3}" -f $r.Name, $availVer, $r.Id, $r.ExitCode)
                $errMsg = if (-not [string]::IsNullOrWhiteSpace($r.Error)) { $r.Error } else { "Exit: $($r.ExitCode)" }
                # Clean up error message - extract meaningful part but keep full text
                if ($errMsg -match '0x[0-9a-fA-F]+\s*:\s*(.+)') {
                    $errMsg = $Matches[1]
                }
                # Store full error for tooltip display
                $pkg.UpdateStatus = $errMsg
            }
        } else {
            Add-LogLine -LogBox $LogBox -Message ("{0} ({1}) ExitCode={2}" -f $r.Name, $r.Id, $r.ExitCode)
        }
    }
    $script:UpdateLastProcessedIndex = $resultsCount - 1
}

# Variable to track processed results
$script:UpdateLastProcessedIndex = -1

# Start enriching packages - fills up to MaxEnrichmentWorkers slots
function Start-NextEnrichment {
    # Find free worker slots and start new operations
    $workerIdx = 0
    while ($workerIdx -lt $script:MaxEnrichmentWorkers) {
        if ($null -ne $script:EnrichmentWorkers[$workerIdx]) { 
            $workerIdx++
            continue 
        }
        if ($script:EnrichmentQueue.Count -eq 0) { return }
        
        $nextId = $script:EnrichmentQueue.Dequeue()
        
        # Skip if already cached
        if ($script:PackageDetailsCache.ContainsKey($nextId)) {
            # Don't increment - try this slot again with next item
            continue
        }
        
        Start-EnrichmentWorker -WorkerIndex $workerIdx -PackageId $nextId
        $workerIdx++
    }
}

function Start-EnrichmentWorker {
    param(
        [int] $WorkerIndex,
        [string] $PackageId
    )
    
    $enrichScript = @'
param($PackageId)

$ErrorActionPreference = 'Stop'

function Invoke-WingetUtf8 {
    param([Parameter(Mandatory)] [string[]] $WingetArgs)

    $wingetPath = (Get-Command winget -ErrorAction Stop).Source

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $wingetPath
    $psi.Arguments = ($WingetArgs -join ' ')
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true

    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $psi.StandardOutputEncoding = $utf8
    $psi.StandardErrorEncoding = $utf8

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi

    if (-not $p.Start()) {
        throw 'Failed to start winget.'
    }

    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    $p.WaitForExit()

    return [pscustomobject]@{
        ExitCode = $p.ExitCode
        Stdout   = $stdout
        Stderr   = $stderr
    }
}

function Parse-WingetShowOutput {
    param([string] $Text)
    
    $details = @{
        Id = $PackageId
        Publisher = ''
        Author = ''
        Description = ''
        License = ''
        LicenseUrl = ''
        SupportUrl = ''
        ReleaseNotesUrl = ''
        Homepage = ''
        InstallerType = ''
        InstallerUrl = ''
        ReleaseDate = ''
    }
    
    $lines = $Text -split "`r?`n"
    $inDescription = $false
    $descLines = @()
    
    foreach ($line in $lines) {
        # Publisher: DE, EN, FR, ES, IT, NL
        if ($line -match '^(Herausgeber|Publisher|Éditeur|Editor|Editore|Uitgever):\s*(.+)$') {
            $details.Publisher = $Matches[2].Trim()
            $inDescription = $false
        }
        # Author: DE, EN, FR, ES, IT, NL
        elseif ($line -match '^(Autor|Author|Auteur|Autore):\s*(.+)$') {
            $details.Author = $Matches[2].Trim()
            $inDescription = $false
        }
        # Description header: DE, EN, FR, ES, IT, NL
        elseif ($line -match '^(Beschreibung|Description|Descripción|Descrizione|Beschrijving):\s*$') {
            $inDescription = $true
        }
        # License: DE, EN, FR, ES, IT, NL
        elseif ($line -match '^(Lizenz|License|Licence|Licencia|Licenza|Licentie):\s*(.+)$') {
            $details.License = $Matches[2].Trim()
            $inDescription = $false
        }
        # License URL: DE, EN, FR, ES, IT, NL
        elseif ($line -match '^(Lizenz-URL|License\s*Url|URL.*licence|URL.*licencia|URL.*licenza|Licentie-URL):\s*(.+)$') {
            $details.LicenseUrl = $Matches[2].Trim()
            $inDescription = $false
        }
        # Support URL: DE, EN, FR, ES, IT, NL
        elseif ($line -match '^(Herausgeber-Support-URL|Publisher\s*Support\s*Url|URL.*support|URL.*soporte|URL.*supporto|Ondersteunings-URL):\s*(.+)$') {
            $details.SupportUrl = $Matches[2].Trim()
            $inDescription = $false
        }
        # Release Notes URL: DE, EN, FR, ES, IT, NL
        elseif ($line -match '^(URL der Versionshinweise|Release\s*Notes\s*Url|URL.*notes|URL.*notas|URL.*rilascio|URL.*release):\s*(.+)$') {
            $details.ReleaseNotesUrl = $Matches[2].Trim()
            $inDescription = $false
        }
        # Homepage: DE, EN, FR, ES, IT, NL
        elseif ($line -match '^(Startseite|Homepage|Page d''accueil|Página de inicio|Pagina iniziale|Startpagina):\s*(.+)$') {
            $details.Homepage = $Matches[2].Trim()
            $inDescription = $false
        }
        # Installer Type: DE, EN, FR, ES, IT, NL
        elseif ($line -match '^(Installertyp|Installer\s*Type|Type d''installateur|Tipo de instalador|Tipo di programma|Installertype):\s*(.+)$') {
            $details.InstallerType = $Matches[2].Trim()
            $inDescription = $false
        }
        # Installer URL: DE, EN, FR, ES, IT, NL
        elseif ($line -match '^\s*(Installer-URL|Installer\s*Url|URL.*installateur|URL.*instalador|URL.*programma|Installatie-URL):\s*(.+)$') {
            $details.InstallerUrl = $Matches[2].Trim()
            $inDescription = $false
        }
        # Release Date: DE, EN, FR, ES, IT, NL
        elseif ($line -match '^(Freigabedatum|Release\s*Date|Date de sortie|Fecha de lanzamiento|Data di rilascio|Uitgavedatum):\s*(.+)$') {
            $details.ReleaseDate = $Matches[2].Trim()
            $inDescription = $false
        }
        elseif ($inDescription) {
            if ($line -match '^\S+:' -and $line -notmatch '^\s') {
                $inDescription = $false
                $details.Description = ($descLines -join ' ').Trim()
            } else {
                $descLines += $line.Trim()
            }
        }
    }
    
    if ($inDescription -and $descLines.Count -gt 0) {
        $details.Description = ($descLines -join ' ').Trim()
    }
    
    return [pscustomobject]$details
}

$resp = Invoke-WingetUtf8 -WingetArgs @('show', $PackageId, '--disable-interactivity')
if ($resp.ExitCode -eq 0) {
    Parse-WingetShowOutput -Text $resp.Stdout
} else {
    [pscustomobject]@{
        Id = $PackageId
        Publisher = ''
        Author = ''
        Description = "(Details konnten nicht geladen werden)"
        License = ''
        LicenseUrl = ''
        SupportUrl = ''
        ReleaseNotesUrl = ''
        Homepage = ''
        InstallerType = ''
        InstallerUrl = ''
        ReleaseDate = ''
    }
}
'@
    
    # Create runspace and execute
    $rs = [RunspaceFactory]::CreateRunspace()
    $rs.Open()

    $ps = [PowerShell]::Create()
    $ps.Runspace = $rs
    $null = $ps.AddScript($enrichScript)
    $null = $ps.AddArgument($PackageId)

    $async = $ps.BeginInvoke()

    $script:EnrichmentWorkers[$WorkerIndex] = @{
        PS = $ps
        Runspace = $rs
        Async = $async
        PackageId = $PackageId
    }
}

# Update the detail panel with package info
function Update-DetailPanel {
    param([WingetPackage] $Package)
    
    $DetailTitle.Text = $Package.Name
    $DetailId.Text = $Package.Id
    $DetailVersion.Text = $Package.Version
    $DetailAvailable.Text = $Package.AvailableVersion
    
    if ($Package.DetailsLoaded) {
        $DetailLoading.Visibility = 'Collapsed'
        $DetailContent.Visibility = 'Visible'
        $DetailPublisher.Text = if ([string]::IsNullOrWhiteSpace($Package.Publisher)) { "-" } else { $Package.Publisher }
        $DetailAuthor.Text = if ([string]::IsNullOrWhiteSpace($Package.Author)) { "-" } else { $Package.Author }
        $DetailDescription.Text = if ([string]::IsNullOrWhiteSpace($Package.Description)) { "-" } else { $Package.Description }
        $DetailLicense.Text = if ([string]::IsNullOrWhiteSpace($Package.License)) { "-" } else { $Package.License }
        $DetailHomepage.Text = if ([string]::IsNullOrWhiteSpace($Package.Homepage)) { "-" } else { $Package.Homepage }
        $DetailInstallerType.Text = if ([string]::IsNullOrWhiteSpace($Package.InstallerType)) { "-" } else { $Package.InstallerType }
        $DetailInstallerUrl.Text = if ([string]::IsNullOrWhiteSpace($Package.InstallerUrl)) { "-" } else { $Package.InstallerUrl }
        $DetailReleaseDate.Text = if ([string]::IsNullOrWhiteSpace($Package.ReleaseDate)) { "-" } else { $Package.ReleaseDate }
        
        # Build Quick Links dynamically
        $hasLicense = -not [string]::IsNullOrWhiteSpace($Package.LicenseUrl)
        $hasSupport = -not [string]::IsNullOrWhiteSpace($Package.SupportUrl)
        $hasReleaseNotes = -not [string]::IsNullOrWhiteSpace($Package.ReleaseNotesUrl)
        
        # Store URLs for click handlers
        $script:CurrentLicenseUrl = $Package.LicenseUrl
        $script:CurrentSupportUrl = $Package.SupportUrl
        $script:CurrentReleaseNotesUrl = $Package.ReleaseNotesUrl
        
        # Show/hide links and separators based on availability
        $LinkLicense.Visibility = if ($hasLicense) { 'Visible' } else { 'Collapsed' }
        $LinkLicense.Text = Get-String 'LinkLicense'
        
        $LinkSupport.Visibility = if ($hasSupport) { 'Visible' } else { 'Collapsed' }
        $LinkSupport.Text = Get-String 'LinkSupport'
        
        $LinkReleaseNotes.Visibility = if ($hasReleaseNotes) { 'Visible' } else { 'Collapsed' }
        $LinkReleaseNotes.Text = Get-String 'LinkVersionInfo'
        
        # Show separator1 only if license AND (support OR releaseNotes) are visible
        $LinkSeparator1.Visibility = if ($hasLicense -and ($hasSupport -or $hasReleaseNotes)) { 'Visible' } else { 'Collapsed' }
        
        # Show separator2 only if support AND releaseNotes are visible
        $LinkSeparator2.Visibility = if ($hasSupport -and $hasReleaseNotes) { 'Visible' } else { 'Collapsed' }
        
        # Show panel only if at least one link is available
        $QuickLinksPanel.Visibility = if ($hasLicense -or $hasSupport -or $hasReleaseNotes) { 'Visible' } else { 'Collapsed' }
    } else {
        $DetailLoading.Visibility = 'Visible'
        $DetailContent.Visibility = 'Collapsed'
        $QuickLinksPanel.Visibility = 'Collapsed'
    }
}

# ============================================================================
# Dark Mode
# ============================================================================
function Set-DarkMode {
    param([bool] $IsDark)
    
    $script:DarkMode = $IsDark
    $MenuDarkMode.IsChecked = $IsDark
    
    if ($IsDark) {
        # Dark Mode colors
        $bgDark = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#1E1E1E')
        $bgMedium = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#252526')
        $bgLight = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#2D2D30')
        $bgButton = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#3C3C3C')
        $fgLight = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#E0E0E0')
        $fgMuted = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#A0A0A0')
        $borderColor = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#3F3F46')
        $accentColor = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#0078D4')
        
        # Create dark mode ScrollBar style via Application Resources
        $scrollBarStyle = @"
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <Style x:Key="DarkScrollBarThumb" TargetType="Thumb">
        <Setter Property="OverridesDefaultStyle" Value="True"/>
        <Setter Property="IsTabStop" Value="False"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="Thumb">
                    <Border Background="#555555" CornerRadius="4"/>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <Style TargetType="ScrollBar">
        <Setter Property="Background" Value="#2D2D30"/>
        <Style.Triggers>
            <Trigger Property="Orientation" Value="Vertical">
                <Setter Property="Width" Value="12"/>
            </Trigger>
            <Trigger Property="Orientation" Value="Horizontal">
                <Setter Property="Height" Value="12"/>
            </Trigger>
        </Style.Triggers>
    </Style>
    <Style TargetType="CheckBox">
        <Setter Property="Foreground" Value="#E0E0E0"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="CheckBox">
                    <BulletDecorator Background="Transparent">
                        <BulletDecorator.Bullet>
                            <Border x:Name="Border" Width="16" Height="16" 
                                    Background="#333333" BorderBrush="#666666" BorderThickness="1" CornerRadius="2">
                                <Path x:Name="CheckMark" Width="10" Height="10" Data="M 0 5 L 3 8 L 9 2" 
                                      Stroke="#E0E0E0" StrokeThickness="2" Visibility="Collapsed"
                                      HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                        </BulletDecorator.Bullet>
                        <ContentPresenter Margin="4,0,0,0" VerticalAlignment="Center" HorizontalAlignment="Left" RecognizesAccessKey="True"/>
                    </BulletDecorator>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsChecked" Value="True">
                            <Setter TargetName="CheckMark" Property="Visibility" Value="Visible"/>
                            <Setter TargetName="Border" Property="Background" Value="#0078D4"/>
                            <Setter TargetName="Border" Property="BorderBrush" Value="#0078D4"/>
                        </Trigger>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter TargetName="Border" Property="BorderBrush" Value="#888888"/>
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <Style TargetType="Button">
        <Setter Property="Foreground" Value="#E0E0E0"/>
        <Setter Property="Background" Value="#3C3C3C"/>
        <Setter Property="BorderBrush" Value="#555555"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="Button">
                    <Border x:Name="border" Background="{TemplateBinding Background}" 
                            BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="2">
                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsEnabled" Value="False">
                            <Setter TargetName="border" Property="Background" Value="#2A2A2A"/>
                            <Setter TargetName="border" Property="BorderBrush" Value="#3A3A3A"/>
                            <Setter Property="Foreground" Value="#555555"/>
                        </Trigger>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter TargetName="border" Property="Background" Value="#4A4A4A"/>
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <Style TargetType="Menu">
        <Setter Property="Background" Value="#252526"/>
        <Setter Property="Foreground" Value="#E0E0E0"/>
        <Setter Property="BorderBrush" Value="#252526"/>
        <Setter Property="BorderThickness" Value="0"/>
    </Style>
    <Style TargetType="MenuItem">
        <Setter Property="Background" Value="#252526"/>
        <Setter Property="Foreground" Value="#E0E0E0"/>
        <Setter Property="BorderBrush" Value="#252526"/>
        <Setter Property="BorderThickness" Value="0"/>
        <Setter Property="Padding" Value="6,3"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="MenuItem">
                    <Border x:Name="Border" Background="{TemplateBinding Background}" 
                            BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="0">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto" SharedSizeGroup="Icon"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto" SharedSizeGroup="Shortcut"/>
                                <ColumnDefinition Width="13"/>
                            </Grid.ColumnDefinitions>
                            <ContentPresenter x:Name="Icon" Grid.Column="0" Margin="6,0,6,0" VerticalAlignment="Center" ContentSource="Icon"/>
                            <Border x:Name="Check" Grid.Column="0" Width="16" Height="16" Margin="6,0,6,0" Visibility="Collapsed" 
                                    Background="#0078D4" CornerRadius="2">
                                <Path x:Name="CheckMark" Width="10" Height="10" Data="M 0 5 L 3 8 L 9 2" 
                                      Stroke="White" StrokeThickness="2" Visibility="Collapsed"/>
                            </Border>
                            <ContentPresenter x:Name="HeaderHost" Grid.Column="1" Margin="6,3" ContentSource="Header" RecognizesAccessKey="True"/>
                            <TextBlock x:Name="InputGestureText" Grid.Column="2" Margin="5,2,0,2" Text="{TemplateBinding InputGestureText}" DockPanel.Dock="Right"/>
                            <Path x:Name="SubMenuArrow" Grid.Column="3" Fill="#E0E0E0" VerticalAlignment="Center" 
                                  Data="M 0 0 L 0 7 L 4 3.5 Z" Visibility="Collapsed"/>
                            <Popup x:Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsSubmenuOpen}" 
                                   AllowsTransparency="True" Focusable="False" PopupAnimation="Fade">
                                <Border x:Name="SubmenuBorder" Background="#2D2D30" BorderBrush="#3F3F46" BorderThickness="1" SnapsToDevicePixels="True">
                                    <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Cycle"/>
                                </Border>
                            </Popup>
                        </Grid>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="Role" Value="TopLevelHeader">
                            <Setter TargetName="SubMenuArrow" Property="Visibility" Value="Collapsed"/>
                            <Setter TargetName="Popup" Property="Placement" Value="Bottom"/>
                        </Trigger>
                        <Trigger Property="Role" Value="TopLevelItem">
                            <Setter TargetName="SubMenuArrow" Property="Visibility" Value="Collapsed"/>
                        </Trigger>
                        <Trigger Property="Role" Value="SubmenuHeader">
                            <Setter TargetName="SubMenuArrow" Property="Visibility" Value="Visible"/>
                            <Setter TargetName="Popup" Property="Placement" Value="Right"/>
                        </Trigger>
                        <Trigger Property="IsHighlighted" Value="True">
                            <Setter TargetName="Border" Property="Background" Value="#3E3E42"/>
                        </Trigger>
                        <Trigger Property="Icon" Value="{x:Null}">
                            <Setter TargetName="Icon" Property="Visibility" Value="Collapsed"/>
                        </Trigger>
                        <Trigger Property="IsCheckable" Value="True">
                            <Setter TargetName="Check" Property="Visibility" Value="Visible"/>
                            <Setter TargetName="Icon" Property="Visibility" Value="Collapsed"/>
                        </Trigger>
                        <Trigger Property="IsChecked" Value="True">
                            <Setter TargetName="CheckMark" Property="Visibility" Value="Visible"/>
                        </Trigger>
                        <Trigger Property="IsChecked" Value="False">
                            <Setter TargetName="Check" Property="Background" Value="#333333"/>
                            <Setter TargetName="Check" Property="BorderBrush" Value="#666666"/>
                            <Setter TargetName="Check" Property="BorderThickness" Value="1"/>
                        </Trigger>
                        <Trigger Property="IsEnabled" Value="False">
                            <Setter Property="Foreground" Value="#656565"/>
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <Style TargetType="Separator">
        <Setter Property="Background" Value="#4A4A4A"/>
        <Setter Property="Margin" Value="0,2"/>
        <Setter Property="Height" Value="1"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="Separator">
                    <Border Background="{TemplateBinding Background}" Height="1" Margin="4,2"/>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <Style TargetType="ToggleButton">
        <Setter Property="Foreground" Value="#E0E0E0"/>
        <Setter Property="Background" Value="#3C3C3C"/>
        <Setter Property="BorderBrush" Value="#555555"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="ToggleButton">
                    <Border x:Name="border" Background="{TemplateBinding Background}" 
                            BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="2">
                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsChecked" Value="True">
                            <Setter TargetName="border" Property="Background" Value="#0078D4"/>
                            <Setter TargetName="border" Property="BorderBrush" Value="#0078D4"/>
                            <Setter Property="Foreground" Value="White"/>
                        </Trigger>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter TargetName="border" Property="Background" Value="#4A4A4A"/>
                        </Trigger>
                        <MultiTrigger>
                            <MultiTrigger.Conditions>
                                <Condition Property="IsChecked" Value="True"/>
                                <Condition Property="IsMouseOver" Value="True"/>
                            </MultiTrigger.Conditions>
                            <Setter TargetName="border" Property="Background" Value="#1084D8"/>
                        </MultiTrigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
</ResourceDictionary>
"@
        
        try {
            $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($scrollBarStyle))
            $darkResources = [System.Windows.Markup.XamlReader]::Load($reader)
            $Window.Resources = $darkResources
        } catch {
            # Fallback if XAML parsing fails
            Write-Host "Dark mode resources warning: $_"
        }
        
        # Window and main containers
        $Window.Background = $bgDark
        $RootPanel.Background = $bgDark
        $MainGrid.Background = $bgDark
        
        # Menu styling - comprehensive dark theme
        $TopMenu.Background = $bgMedium
        $TopMenu.Foreground = $fgLight
        
        # Style all menu items recursively for dark mode
        $menuPopupBg = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#2D2D30')
        $separatorBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#4A4A4A')
        $TopMenu.BorderBrush = $bgMedium
        $TopMenu.BorderThickness = [System.Windows.Thickness]::new(0)
        foreach ($menuItem in $TopMenu.Items) {
            if ($menuItem -is [System.Windows.Controls.MenuItem]) {
                $menuItem.Background = $bgMedium
                $menuItem.Foreground = $fgLight
                $menuItem.BorderBrush = $bgMedium
                $menuItem.BorderThickness = [System.Windows.Thickness]::new(0)
                foreach ($subItem in $menuItem.Items) {
                    if ($subItem -is [System.Windows.Controls.MenuItem]) {
                        $subItem.Background = $menuPopupBg
                        $subItem.Foreground = $fgLight
                        $subItem.BorderBrush = $menuPopupBg
                        $subItem.BorderThickness = [System.Windows.Thickness]::new(0)
                    } elseif ($subItem -is [System.Windows.Controls.Separator]) {
                        $subItem.Background = $separatorBrush
                        $subItem.Foreground = $separatorBrush
                    }
                }
            }
        }
        
        # Button divider
        if ($null -ne $ButtonDivider) {
            $ButtonDivider.Fill = $borderColor
        }
        
        # Progress bar border
        $BusyBar.BorderBrush = $borderColor
        
        # DataGrid styling
        $PackagesGrid.Background = $bgMedium
        $PackagesGrid.Foreground = $fgLight
        $PackagesGrid.RowBackground = $bgMedium
        $PackagesGrid.AlternatingRowBackground = $bgLight
        $PackagesGrid.BorderBrush = $borderColor
        
        # DataGrid column header style
        $headerStyle = [System.Windows.Style]::new([System.Windows.Controls.Primitives.DataGridColumnHeader])
        $headerStyle.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.Control]::BackgroundProperty, $bgLight))
        $headerStyle.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.Control]::ForegroundProperty, $fgLight))
        $headerStyle.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.Control]::BorderBrushProperty, $borderColor))
        $headerStyle.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.Control]::PaddingProperty, [System.Windows.Thickness]::new(8,4,8,4)))
        $PackagesGrid.ColumnHeaderStyle = $headerStyle
        
        # Detail panel (Border with ScrollViewer > StackPanel)
        $DetailPanel.Background = $bgMedium
        $DetailPanel.BorderBrush = $borderColor
        
        # Detail panel text elements
        $DetailTitle.Foreground = $fgLight
        $DetailId.Foreground = $fgMuted
        $DetailVersion.Foreground = $fgLight
        $DetailLoading.Foreground = $fgMuted
        $DetailPublisher.Foreground = $fgLight
        $DetailAuthor.Foreground = $fgLight
        $DetailDescription.Foreground = $fgLight
        $DetailLicense.Foreground = $fgLight
        $DetailHomepage.Foreground = $accentColor
        $DetailInstallerType.Foreground = $fgLight
        $DetailInstallerUrl.Foreground = $accentColor
        $DetailReleaseDate.Foreground = $fgLight
        $DetailAvailable.Foreground = $accentColor
        
        # Quick Links styling
        $LinkLicense.Foreground = $accentColor
        $LinkSupport.Foreground = $accentColor
        $LinkReleaseNotes.Foreground = $accentColor
        $LinkSeparator1.Foreground = $fgMuted
        $LinkSeparator2.Foreground = $fgMuted
        
        # Detail panel labels
        $LabelInstalled.Foreground = $fgLight
        $LabelAvailable.Foreground = $fgLight
        $LabelPublisher.Foreground = $fgLight
        $LabelAuthor.Foreground = $fgLight
        $LabelDescription.Foreground = $fgLight
        $LabelLicense.Foreground = $fgLight
        $LabelHomepage.Foreground = $fgLight
        $LabelInstaller.Foreground = $fgLight
        $LabelInstallerUrl.Foreground = $fgLight
        $LabelReleaseDate.Foreground = $fgLight
        
        # Status text
        $StatusText.Foreground = $fgLight
        
        # Progress bar
        $BusyBar.Background = $bgLight
        $BusyBar.Foreground = $accentColor
        
        # Splitter
        $DetailSplitter.Background = $borderColor
        
        # Log panel and toggle button
        if ($null -ne $LogToggleButton) {
            $LogToggleButton.Background = $bgButton
            $LogToggleButton.Foreground = $fgLight
            $LogToggleButton.BorderBrush = $borderColor
        }
        if ($null -ne $LogPanel) {
            $LogPanel.Background = $bgMedium
            $LogPanel.BorderBrush = $borderColor
            $LogBox.Background = $bgLight
            $LogBox.Foreground = $fgLight
            $LogBox.BorderBrush = $borderColor
        }
        
        # Buttons with better dark mode styling
        $RefreshButton.Background = $accentColor
        $RefreshButton.Foreground = [System.Windows.Media.Brushes]::White
        $RefreshButton.BorderBrush = $accentColor
        
        $UpdateSelectedButton.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#388E3C')
        $UpdateSelectedButton.Foreground = [System.Windows.Media.Brushes]::White
        $UpdateSelectedButton.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#388E3C')
        
        $CancelButton.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#D32F2F')
        $CancelButton.Foreground = [System.Windows.Media.Brushes]::White
        $CancelButton.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#D32F2F')
        
        $SelectAllButton.Background = $bgButton
        $SelectAllButton.Foreground = $fgLight
        $SelectAllButton.BorderBrush = $borderColor
        
        $SelectNoneButton.Background = $bgButton
        $SelectNoneButton.Foreground = $fgLight
        $SelectNoneButton.BorderBrush = $borderColor
        
    } else {
        # Light Mode colors
        $bgWhite = [System.Windows.Media.Brushes]::White
        $bgLight = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#F5F5F5')
        $bgGray = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#E8E8E8')
        $fgDark = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#333333')
        $borderColor = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#CCCCCC')
        $accentColor = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#0078D4')
        
        # Clear dark mode resources and use light mode styles
        $lightStyle = @"
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <Style TargetType="ScrollBar">
        <Setter Property="Background" Value="#F0F0F0"/>
    </Style>
    <Style TargetType="CheckBox">
        <Setter Property="Foreground" Value="#333333"/>
    </Style>
    <Style TargetType="Button">
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="Button">
                    <Border x:Name="border" Background="{TemplateBinding Background}" 
                            BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="2">
                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsEnabled" Value="False">
                            <Setter TargetName="border" Property="Background" Value="#E8E8E8"/>
                            <Setter TargetName="border" Property="BorderBrush" Value="#CCCCCC"/>
                            <Setter Property="Foreground" Value="#999999"/>
                        </Trigger>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter TargetName="border" Property="Background" Value="#D0D0D0"/>
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <Style TargetType="ToggleButton">
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="ToggleButton">
                    <Border x:Name="border" Background="{TemplateBinding Background}" 
                            BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="2">
                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsChecked" Value="True">
                            <Setter TargetName="border" Property="Background" Value="#0078D4"/>
                            <Setter TargetName="border" Property="BorderBrush" Value="#0078D4"/>
                            <Setter Property="Foreground" Value="White"/>
                        </Trigger>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter TargetName="border" Property="Background" Value="#D0D0D0"/>
                        </Trigger>
                        <MultiTrigger>
                            <MultiTrigger.Conditions>
                                <Condition Property="IsChecked" Value="True"/>
                                <Condition Property="IsMouseOver" Value="True"/>
                            </MultiTrigger.Conditions>
                            <Setter TargetName="border" Property="Background" Value="#1084D8"/>
                        </MultiTrigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <Style TargetType="MenuItem">
        <Setter Property="Foreground" Value="#333333"/>
    </Style>
</ResourceDictionary>
"@
        
        try {
            $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($lightStyle))
            $lightResources = [System.Windows.Markup.XamlReader]::Load($reader)
            $Window.Resources = $lightResources
        } catch {
            $Window.Resources = [System.Windows.ResourceDictionary]::new()
        }
        
        # Window and main containers
        $Window.Background = $bgLight
        $RootPanel.Background = $bgLight
        $MainGrid.Background = $bgLight
        
        # Menu styling - reset for light mode
        $TopMenu.Background = $bgGray
        $TopMenu.Foreground = $fgDark
        
        # Reset all menu items for light mode
        $lightBorder = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#CCCCCC')
        $lightSeparator = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#D0D0D0')
        $transparentBrush = [System.Windows.Media.Brushes]::Transparent
        $TopMenu.BorderBrush = $transparentBrush
        foreach ($menuItem in $TopMenu.Items) {
            if ($menuItem -is [System.Windows.Controls.MenuItem]) {
                $menuItem.Background = $bgGray
                $menuItem.Foreground = $fgDark
                $menuItem.BorderBrush = $transparentBrush
                foreach ($subItem in $menuItem.Items) {
                    if ($subItem -is [System.Windows.Controls.MenuItem]) {
                        $subItem.Background = $bgWhite
                        $subItem.Foreground = $fgDark
                        $subItem.BorderBrush = $lightBorder
                    } elseif ($subItem -is [System.Windows.Controls.Separator]) {
                        $subItem.Background = $lightSeparator
                        $subItem.Foreground = $lightSeparator
                    }
                }
            }
        }
        
        # Button divider
        if ($null -ne $ButtonDivider) {
            $ButtonDivider.Fill = $lightBorder
        }
        
        # Progress bar border
        $BusyBar.BorderBrush = $lightBorder
        
        # DataGrid styling
        $PackagesGrid.Background = $bgWhite
        $PackagesGrid.Foreground = $fgDark
        $PackagesGrid.RowBackground = $bgWhite
        $PackagesGrid.AlternatingRowBackground = $bgLight
        $PackagesGrid.BorderBrush = $borderColor
        
        # DataGrid column header style - reset to default
        $headerStyle = [System.Windows.Style]::new([System.Windows.Controls.Primitives.DataGridColumnHeader])
        $headerStyle.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.Control]::BackgroundProperty, $bgGray))
        $headerStyle.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.Control]::ForegroundProperty, $fgDark))
        $headerStyle.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.Control]::BorderBrushProperty, $borderColor))
        $headerStyle.Setters.Add([System.Windows.Setter]::new([System.Windows.Controls.Control]::PaddingProperty, [System.Windows.Thickness]::new(8,4,8,4)))
        $PackagesGrid.ColumnHeaderStyle = $headerStyle
        
        # Detail panel (Border with ScrollViewer > StackPanel)
        $DetailPanel.Background = $bgWhite
        $DetailPanel.BorderBrush = $borderColor
        
        # Detail panel text elements
        $fgGray = [System.Windows.Media.BrushConverter]::new().ConvertFromString('Gray')
        $DetailTitle.Foreground = $fgDark
        $DetailId.Foreground = $fgGray
        $DetailVersion.Foreground = $fgDark
        $DetailLoading.Foreground = $fgGray
        $DetailPublisher.Foreground = $fgDark
        $DetailAuthor.Foreground = $fgDark
        $DetailDescription.Foreground = $fgDark
        $DetailLicense.Foreground = $fgDark
        $DetailHomepage.Foreground = $accentColor
        $DetailInstallerType.Foreground = $fgDark
        $DetailInstallerUrl.Foreground = $accentColor
        $DetailReleaseDate.Foreground = $fgDark
        $DetailAvailable.Foreground = $accentColor
        
        # Quick Links styling
        $LinkLicense.Foreground = $accentColor
        $LinkSupport.Foreground = $accentColor
        $LinkReleaseNotes.Foreground = $accentColor
        $LinkSeparator1.Foreground = $fgGray
        $LinkSeparator2.Foreground = $fgGray
        
        # Detail panel labels
        $LabelInstalled.Foreground = $fgDark
        $LabelAvailable.Foreground = $fgDark
        $LabelPublisher.Foreground = $fgDark
        $LabelAuthor.Foreground = $fgDark
        $LabelDescription.Foreground = $fgDark
        $LabelLicense.Foreground = $fgDark
        $LabelHomepage.Foreground = $fgDark
        $LabelInstaller.Foreground = $fgDark
        $LabelInstallerUrl.Foreground = $fgDark
        $LabelReleaseDate.Foreground = $fgDark
        
        # Status text
        $StatusText.Foreground = $fgDark
        
        # Progress bar
        $BusyBar.Background = $bgGray
        $BusyBar.Foreground = $accentColor
        
        # Splitter
        $DetailSplitter.Background = $borderColor
        
        # Log panel and toggle button
        if ($null -ne $LogToggleButton) {
            $LogToggleButton.Background = $bgGray
            $LogToggleButton.Foreground = $fgDark
            $LogToggleButton.BorderBrush = $borderColor
        }
        if ($null -ne $LogPanel) {
            $LogPanel.Background = $bgLight
            $LogPanel.BorderBrush = $borderColor
            $LogBox.Background = $bgWhite
            $LogBox.Foreground = $fgDark
            $LogBox.BorderBrush = $borderColor
        }
        
        # Buttons
        $RefreshButton.Background = $accentColor
        $RefreshButton.Foreground = [System.Windows.Media.Brushes]::White
        $RefreshButton.BorderBrush = $accentColor
        
        $UpdateSelectedButton.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#4CAF50')
        $UpdateSelectedButton.Foreground = [System.Windows.Media.Brushes]::White
        $UpdateSelectedButton.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#4CAF50')
        
        $CancelButton.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#F44336')
        $CancelButton.Foreground = [System.Windows.Media.Brushes]::White
        $CancelButton.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#F44336')
        
        $SelectAllButton.Background = $bgGray
        $SelectAllButton.Foreground = $fgDark
        $SelectAllButton.BorderBrush = $borderColor
        
        $SelectNoneButton.Background = $bgGray
        $SelectNoneButton.Foreground = $fgDark
        $SelectNoneButton.BorderBrush = $borderColor
    }
    
    Save-Settings
}

# ============================================================================
# Language Application
# ============================================================================
function Set-Language {
    param([string] $Lang)
    
    $script:CurrentLanguage = $Lang
    
    # Update menu checks
    $MenuLangDE.IsChecked = ($Lang -eq 'de')
    $MenuLangEN.IsChecked = ($Lang -eq 'en')
    $MenuLangFR.IsChecked = ($Lang -eq 'fr')
    $MenuLangES.IsChecked = ($Lang -eq 'es')
    $MenuLangIT.IsChecked = ($Lang -eq 'it')
    $MenuLangNL.IsChecked = ($Lang -eq 'nl')
    
    # Update UI elements
    # Menu headers
    $MenuSettings.Header = Get-String 'MenuSettings'
    $MenuProxy.Header = Get-String 'MenuProxy'
    $MenuLanguage.Header = Get-String 'MenuLanguage'
    $MenuInstallMode.Header = Get-String 'MenuInstallMode'
    $MenuModeNone.Header = Get-String 'InstallModeNone'
    $MenuModeSilent.Header = Get-String 'InstallModeSilent'
    $MenuModeInteractive.Header = Get-String 'InstallModeInteractive'
    $MenuForce.Header = Get-String 'MenuForce'
    $MenuForce.ToolTip = Get-String 'ForceTooltip'
    $MenuView.Header = Get-String 'MenuView'
    $MenuDarkMode.Header = Get-String 'MenuDarkMode'
    $MenuTimeFormat.Header = Get-String 'MenuTimeFormat'
    $MenuTimeDateTime.Header = Get-String 'TimeFormatDateTime'
    $MenuTimeOnly.Header = Get-String 'TimeFormatTimeOnly'
    $MenuTimeNone.Header = Get-String 'TimeFormatNone'
    $MenuHelp.Header = Get-String 'MenuHelp'
    $MenuAbout.Header = Get-String 'MenuAbout'
    
    # Buttons
    $RefreshButton.Content = Get-String 'BtnRefresh'
    $UpdateSelectedButton.Content = Get-String 'BtnUpdate'
    $CancelButton.Content = Get-String 'BtnCancel'
    $SelectAllButton.Content = Get-String 'BtnSelectAll'
    $SelectNoneButton.Content = Get-String 'BtnSelectNone'
    
    $StatusText.Text = Get-String 'StatusReady'
    
    if ($null -ne $LogToggleButton) {
        $LogToggleButton.Content = "📋 " + (Get-String 'ExpanderLog')
    }
    
    # Update DataGrid column headers
    $ColName.Header = Get-String 'ColName'
    $ColVersion.Header = Get-String 'ColVersion'
    $ColAvailable.Header = Get-String 'ColAvailable'
    $ColStatus.Header = Get-String 'ColStatus'
    
    # Update detail panel labels
    $LabelInstalled.Text = Get-String 'DetailInstalled'
    $LabelAvailable.Text = Get-String 'DetailAvailable'
    $LabelPublisher.Text = Get-String 'DetailPublisher'
    $LabelAuthor.Text = Get-String 'DetailAuthor'
    $LabelDescription.Text = Get-String 'DetailDescription'
    $LabelLicense.Text = Get-String 'DetailLicense'
    $LabelHomepage.Text = Get-String 'DetailHomepage'
    $LabelInstaller.Text = Get-String 'DetailInstaller'
    $LabelInstallerUrl.Text = Get-String 'DetailDownload'
    $LabelReleaseDate.Text = Get-String 'DetailReleaseDate'
    $DetailLoading.Text = Get-String 'DetailLoading'
    
    Save-Settings
}

# ============================================================================
# Proxy Settings Dialog
# ============================================================================
function Show-ProxyDialog {
    $proxyXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$(Get-String 'ProxyTitle')" Height="320" Width="450" 
        WindowStartupLocation="CenterOwner" ResizeMode="NoResize">
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        
        <CheckBox x:Name="ProxyEnabled" Grid.Row="0" Grid.ColumnSpan="2" 
                  Content="$(Get-String 'ProxyEnable')" Margin="0,0,0,15"/>
        
        <TextBlock Grid.Row="1" Grid.Column="0" Text="$(Get-String 'ProxyServer')" 
                   VerticalAlignment="Center" Margin="0,0,10,8"/>
        <TextBox x:Name="ProxyServer" Grid.Row="1" Grid.Column="1" Margin="0,0,0,8"/>
        
        <TextBlock Grid.Row="2" Grid.Column="0" Text="$(Get-String 'ProxyPort')" 
                   VerticalAlignment="Center" Margin="0,0,10,8"/>
        <StackPanel Grid.Row="2" Grid.Column="1" Orientation="Horizontal" Margin="0,0,0,8">
            <TextBox x:Name="ProxyPort" Width="80"/>
            <TextBlock Text="$(Get-String 'ProxyType')" VerticalAlignment="Center" Margin="20,0,10,0"/>
            <ComboBox x:Name="ProxyType" Width="100">
                <ComboBoxItem Content="HTTP" IsSelected="True"/>
                <ComboBoxItem Content="HTTPS"/>
                <ComboBoxItem Content="SOCKS4"/>
                <ComboBoxItem Content="SOCKS5"/>
            </ComboBox>
        </StackPanel>
        
        <TextBlock Grid.Row="3" Grid.Column="0" Text="$(Get-String 'ProxyUser')" 
                   VerticalAlignment="Center" Margin="0,0,10,8"/>
        <TextBox x:Name="ProxyUser" Grid.Row="3" Grid.Column="1" Margin="0,0,0,8"/>
        
        <TextBlock Grid.Row="4" Grid.Column="0" Text="$(Get-String 'ProxyPassword')" 
                   VerticalAlignment="Center" Margin="0,0,10,8"/>
        <PasswordBox x:Name="ProxyPassword" Grid.Row="4" Grid.Column="1" Margin="0,0,0,8"/>
        
        <TextBlock x:Name="ValidationMessage" Grid.Row="5" Grid.ColumnSpan="2" 
                   Foreground="Red" FontSize="11" TextWrapping="Wrap" Margin="0,0,0,8" Visibility="Collapsed"/>
        
        <StackPanel Grid.Row="7" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="BtnSave" Content="$(Get-String 'ProxySave')" Width="80" Margin="0,0,10,0" IsDefault="True"/>
            <Button x:Name="BtnCancel" Content="$(Get-String 'ProxyCancel')" Width="80" IsCancel="True"/>
        </StackPanel>
    </Grid>
</Window>
"@
    
    $reader = New-Object System.Xml.XmlNodeReader([xml]$proxyXaml)
    $proxyWindow = [Windows.Markup.XamlReader]::Load($reader)
    $proxyWindow.Owner = $window
    
    $chkEnabled = $proxyWindow.FindName('ProxyEnabled')
    $txtServer = $proxyWindow.FindName('ProxyServer')
    $txtPort = $proxyWindow.FindName('ProxyPort')
    $cmbType = $proxyWindow.FindName('ProxyType')
    $txtUser = $proxyWindow.FindName('ProxyUser')
    $txtPassword = $proxyWindow.FindName('ProxyPassword')
    $txtValidation = $proxyWindow.FindName('ValidationMessage')
    $btnSave = $proxyWindow.FindName('BtnSave')
    $btnCancel = $proxyWindow.FindName('BtnCancel')
    
    # Apply dark mode styling directly to controls
    if ($script:DarkMode) {
        $bgDark = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#1E1E1E')
        $bgInput = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#333333')
        $fgLight = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#E0E0E0')
        $borderColor = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#555555')
        
        # Window background
        $proxyWindow.Background = $bgDark
        
        # Find the main Grid and style it
        $mainGrid = $proxyWindow.Content
        if ($mainGrid -is [System.Windows.Controls.Grid]) {
            $mainGrid.Background = $bgDark
            
            # Style all children including nested StackPanels
            foreach ($child in $mainGrid.Children) {
                if ($child -is [System.Windows.Controls.TextBlock]) {
                    $child.Foreground = $fgLight
                }
                elseif ($child -is [System.Windows.Controls.StackPanel]) {
                    $child.Background = $bgDark
                    # Style TextBlocks inside StackPanel (like the Type label)
                    foreach ($spChild in $child.Children) {
                        if ($spChild -is [System.Windows.Controls.TextBlock]) {
                            $spChild.Foreground = $fgLight
                        }
                    }
                }
            }
        }
        
        # CheckBox
        $chkEnabled.Foreground = $fgLight
        
        # TextBoxes
        $txtServer.Background = $bgInput
        $txtServer.Foreground = $fgLight
        $txtServer.BorderBrush = $borderColor
        $txtServer.CaretBrush = $fgLight
        
        $txtPort.Background = $bgInput
        $txtPort.Foreground = $fgLight
        $txtPort.BorderBrush = $borderColor
        $txtPort.CaretBrush = $fgLight
        
        $txtUser.Background = $bgInput
        $txtUser.Foreground = $fgLight
        $txtUser.BorderBrush = $borderColor
        $txtUser.CaretBrush = $fgLight
        
        # ComboBox - need to style dropdown items too
        $cmbType.Background = $bgInput
        $cmbType.Foreground = $fgLight
        $cmbType.BorderBrush = $borderColor
        
        # Style ComboBox items for dark mode dropdown
        foreach ($item in $cmbType.Items) {
            if ($item -is [System.Windows.Controls.ComboBoxItem]) {
                $item.Background = $bgInput
                $item.Foreground = $fgLight
            }
        }
        
        # Add style for dropdown popup via resources
        $itemContainerStyle = New-Object System.Windows.Style([System.Windows.Controls.ComboBoxItem])
        $itemContainerStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BackgroundProperty, $bgInput)))
        $itemContainerStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::ForegroundProperty, $fgLight)))
        $itemContainerStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BorderBrushProperty, $borderColor)))
        
        # Hover trigger
        $hoverTrigger = New-Object System.Windows.Trigger
        $hoverTrigger.Property = [System.Windows.Controls.ComboBoxItem]::IsMouseOverProperty
        $hoverTrigger.Value = $true
        $hoverBg = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#3E3E42')
        $hoverTrigger.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Control]::BackgroundProperty, $hoverBg)))
        $itemContainerStyle.Triggers.Add($hoverTrigger)
        
        $cmbType.ItemContainerStyle = $itemContainerStyle
        
        # PasswordBox
        $txtPassword.Background = $bgInput
        $txtPassword.Foreground = $fgLight
        $txtPassword.BorderBrush = $borderColor
        $txtPassword.CaretBrush = $fgLight
        
        # Buttons - use darker backgrounds for better visibility
        $btnSaveBg = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#0066B8')
        $btnSave.Background = $btnSaveBg
        $btnSave.Foreground = [System.Windows.Media.Brushes]::White
        $btnSave.BorderBrush = $btnSaveBg
        
        $btnCancelBg = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#2D2D30')
        $btnCancel.Background = $btnCancelBg
        $btnCancel.Foreground = $fgLight
        $btnCancel.BorderBrush = $borderColor
    }
    
    # Load current settings
    $chkEnabled.IsChecked = $script:ProxySettings.Enabled
    $txtServer.Text = $script:ProxySettings.Server
    $txtPort.Text = $script:ProxySettings.Port
    $txtUser.Text = $script:ProxySettings.User
    $txtPassword.Password = $script:ProxySettings.Password
    
    # Set proxy type from saved settings
    $savedType = $script:ProxySettings.Type
    if (-not $savedType) { $savedType = 'HTTP' }
    for ($i = 0; $i -lt $cmbType.Items.Count; $i++) {
        if ($cmbType.Items[$i].Content -eq $savedType) {
            $cmbType.SelectedIndex = $i
            break
        }
    }
    
    $btnSave.Add_Click({
        # Validation
        $txtValidation.Visibility = 'Collapsed'
        
        if ($chkEnabled.IsChecked) {
            # Validate server
            $server = $txtServer.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($server)) {
                $txtValidation.Text = (Get-String 'ProxyValidationServer')
                $txtValidation.Visibility = 'Visible'
                $txtServer.Focus()
                return
            }
            
            # Validate port (1-65535)
            $portText = $txtPort.Text.Trim()
            $port = 0
            if (-not [int]::TryParse($portText, [ref]$port) -or $port -lt 1 -or $port -gt 65535) {
                $txtValidation.Text = (Get-String 'ProxyValidationPort')
                $txtValidation.Visibility = 'Visible'
                $txtPort.Focus()
                return
            }
        }
        
        $script:ProxySettings.Enabled = $chkEnabled.IsChecked
        $script:ProxySettings.Server = $txtServer.Text.Trim()
        $script:ProxySettings.Port = $txtPort.Text.Trim()
        $script:ProxySettings.Type = $cmbType.SelectedItem.Content
        $script:ProxySettings.User = $txtUser.Text
        $script:ProxySettings.Password = $txtPassword.Password
        
        Save-Settings
        Set-ProxySettings
        
        $proxyWindow.DialogResult = $true
        $proxyWindow.Close()
    })
    
    $btnCancel.Add_Click({
        $proxyWindow.DialogResult = $false
        $proxyWindow.Close()
    })
    
    $proxyWindow.ShowDialog() | Out-Null
}

# ============================================================================
# About Dialog
# ============================================================================
function Show-AboutDialog {
    $script:AppVersion = '1.0.0'
    
    $aboutXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$(Get-String 'AboutTitle')" Height="280" Width="400" 
        WindowStartupLocation="CenterOwner" ResizeMode="NoResize">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <TextBlock Grid.Row="0" Text="WingetUpdateUI" FontSize="24" FontWeight="Bold" HorizontalAlignment="Center" Margin="0,0,0,5"/>
        <TextBlock x:Name="VersionText" Grid.Row="1" FontSize="14" HorizontalAlignment="Center" Margin="0,0,0,10"/>
        <TextBlock x:Name="DescriptionText" Grid.Row="2" TextWrapping="Wrap" TextAlignment="Center" Margin="0,0,0,10"/>
        <TextBlock x:Name="AuthorText" Grid.Row="3" FontSize="12" HorizontalAlignment="Center" Margin="0,0,0,5"/>
        <TextBlock x:Name="GitHubLink" Grid.Row="4" FontSize="12" HorizontalAlignment="Center" Margin="0,0,0,10" Cursor="Hand">
            <Hyperlink x:Name="GitHubHyperlink" NavigateUri="https://github.com/ThisLimn0/WingetUpdateUI">
                <Run Text="$(Get-String 'AboutGitHub'): github.com/ThisLimn0/WingetUpdateUI"/>
            </Hyperlink>
        </TextBlock>
        <TextBlock x:Name="CopyrightText" Grid.Row="5" FontSize="11" HorizontalAlignment="Center" VerticalAlignment="Bottom" Margin="0,0,0,15"/>
        
        <Button x:Name="BtnOK" Grid.Row="6" Content="$(Get-String 'AboutOK')" Width="80" HorizontalAlignment="Center" IsDefault="True" IsCancel="True"/>
    </Grid>
</Window>
"@
    
    $reader = New-Object System.Xml.XmlNodeReader([xml]$aboutXaml)
    $aboutWindow = [Windows.Markup.XamlReader]::Load($reader)
    $aboutWindow.Owner = $window
    
    $txtVersion = $aboutWindow.FindName('VersionText')
    $txtDescription = $aboutWindow.FindName('DescriptionText')
    $txtAuthor = $aboutWindow.FindName('AuthorText')
    $txtCopyright = $aboutWindow.FindName('CopyrightText')
    $gitHubHyperlink = $aboutWindow.FindName('GitHubHyperlink')
    $btnOK = $aboutWindow.FindName('BtnOK')
    
    $txtVersion.Text = (Get-String 'AboutVersion') -f $script:AppVersion
    $txtDescription.Text = Get-String 'AboutDescription'
    $txtAuthor.Text = Get-String 'AboutAuthor'
    $txtCopyright.Text = Get-String 'AboutCopyright'
    
    # Handle hyperlink click
    $gitHubHyperlink.Add_RequestNavigate({
        param($sender, $e)
        [System.Diagnostics.Process]::Start($e.Uri.AbsoluteUri) | Out-Null
        $e.Handled = $true
    })
    
    # Apply dark mode styling
    if ($script:DarkMode) {
        $bgDark = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#1E1E1E')
        $fgLight = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#E0E0E0')
        
        $aboutWindow.Background = $bgDark
        
        $mainGrid = $aboutWindow.Content
        if ($mainGrid -is [System.Windows.Controls.Grid]) {
            $mainGrid.Background = $bgDark
            foreach ($child in $mainGrid.Children) {
                if ($child -is [System.Windows.Controls.TextBlock]) {
                    $child.Foreground = $fgLight
                }
            }
        }
        
        $btnOKBg = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#0078D4')
        $btnOK.Background = $btnOKBg
        $btnOK.Foreground = [System.Windows.Media.Brushes]::White
        $btnOK.BorderBrush = $btnOKBg
    }
    
    $btnOK.Add_Click({
        $aboutWindow.Close()
    })
    
    $aboutWindow.ShowDialog() | Out-Null
}

# ============================================================================
# Menu Event Handlers
# ============================================================================
$MenuProxy.Add_Click({ Show-ProxyDialog })

$MenuHelp = $window.FindName('MenuHelp')
$MenuAbout = $window.FindName('MenuAbout')
$MenuAbout.Add_Click({ Show-AboutDialog })

$MenuLangDE.Add_Click({ Set-Language 'de' })
$MenuLangEN.Add_Click({ Set-Language 'en' })
$MenuLangFR.Add_Click({ Set-Language 'fr' })
$MenuLangES.Add_Click({ Set-Language 'es' })
$MenuLangIT.Add_Click({ Set-Language 'it' })
$MenuLangNL.Add_Click({ Set-Language 'nl' })

# Install Mode handlers
function Set-InstallMode {
    param([string] $Mode)
    
    $script:InstallMode = $Mode
    
    # Update menu checks (radio button behavior)
    $MenuModeNone.IsChecked = ($Mode -eq 'none')
    $MenuModeSilent.IsChecked = ($Mode -eq 'silent')
    $MenuModeInteractive.IsChecked = ($Mode -eq 'interactive')
    
    Save-Settings
}

$MenuModeNone.Add_Click({ Set-InstallMode 'none' })
$MenuModeSilent.Add_Click({ Set-InstallMode 'silent' })
$MenuModeInteractive.Add_Click({ Set-InstallMode 'interactive' })

# Force toggle
$MenuForce.Add_Click({
    $script:ForceInstall = $MenuForce.IsChecked
    Save-Settings
})

# Dark Mode toggle
$MenuDarkMode.Add_Click({
    Set-DarkMode (-not $script:DarkMode)
})

# Time Format toggle functions
function Set-TimeFormat {
    param([string]$Format)
    $script:TimeFormat = $Format
    $MenuTimeDateTime.IsChecked = ($Format -eq 'datetime')
    $MenuTimeOnly.IsChecked = ($Format -eq 'timeonly')
    $MenuTimeNone.IsChecked = ($Format -eq 'none')
    Save-Settings
}

$MenuTimeDateTime.Add_Click({ Set-TimeFormat 'datetime' })
$MenuTimeOnly.Add_Click({ Set-TimeFormat 'timeonly' })
$MenuTimeNone.Add_Click({ Set-TimeFormat 'none' })

# Apply saved settings on startup
Set-Language $script:CurrentLanguage
Set-InstallMode $script:InstallMode
$MenuForce.IsChecked = $script:ForceInstall
Set-DarkMode $script:DarkMode

# Apply time format setting
$MenuTimeDateTime.IsChecked = ($script:TimeFormat -eq 'datetime')
$MenuTimeOnly.IsChecked = ($script:TimeFormat -eq 'timeonly')
$MenuTimeNone.IsChecked = ($script:TimeFormat -eq 'none')

$timer = [System.Windows.Threading.DispatcherTimer]::new()
$timer.Interval = [TimeSpan]::FromMilliseconds(200)
$timer.Add_Tick({ 
    Complete-PendingOperation 
    Complete-EnrichmentOperation
    Update-UpdateProgress
})
$timer.Start()

# Handle selection change to show/hide detail panel
$PackagesGrid.Add_SelectionChanged({
    $selectedPkg = $PackagesGrid.SelectedItem
    if ($selectedPkg -ne $null) {
        # Show detail panel
        $DetailColumnDef.Width = [System.Windows.GridLength]::new(300)
        $DetailSplitter.Visibility = 'Visible'
        $DetailPanel.Visibility = 'Visible'
        Update-DetailPanel -Package $selectedPkg
    } else {
        # Hide detail panel
        $DetailColumnDef.Width = [System.Windows.GridLength]::new(0)
        $DetailSplitter.Visibility = 'Collapsed'
        $DetailPanel.Visibility = 'Collapsed'
    }
})

# Click handlers for clickable links
$DetailHomepage.Add_MouseLeftButtonUp({
    $url = $DetailHomepage.Text
    if (-not [string]::IsNullOrWhiteSpace($url) -and $url -ne '-' -and $url -match '^https?://') {
        [System.Diagnostics.Process]::Start($url) | Out-Null
    }
})

$DetailInstallerUrl.Add_MouseLeftButtonUp({
    $url = $DetailInstallerUrl.Text
    if (-not [string]::IsNullOrWhiteSpace($url) -and $url -ne '-' -and $url -match '^https?://') {
        [System.Diagnostics.Process]::Start($url) | Out-Null
    }
})

# Quick Links click handlers
$LinkLicense.Add_MouseLeftButtonUp({
    if (-not [string]::IsNullOrWhiteSpace($script:CurrentLicenseUrl)) {
        [System.Diagnostics.Process]::Start($script:CurrentLicenseUrl) | Out-Null
    }
})

$LinkSupport.Add_MouseLeftButtonUp({
    if (-not [string]::IsNullOrWhiteSpace($script:CurrentSupportUrl)) {
        [System.Diagnostics.Process]::Start($script:CurrentSupportUrl) | Out-Null
    }
})

$LinkReleaseNotes.Add_MouseLeftButtonUp({
    if (-not [string]::IsNullOrWhiteSpace($script:CurrentReleaseNotesUrl)) {
        [System.Diagnostics.Process]::Start($script:CurrentReleaseNotesUrl) | Out-Null
    }
})

function Set-ButtonsEnabled {
    param([bool] $enabled)
    $window.Dispatcher.Invoke([Action] {
        $RefreshButton.IsEnabled = $enabled
        $UpdateSelectedButton.IsEnabled = $enabled
        $SelectAllButton.IsEnabled = $enabled
        $SelectNoneButton.IsEnabled = $enabled
    })
}

function Update-PackageList {
    Set-ButtonsEnabled -enabled:$false
    Set-Busy -BusyBar $BusyBar -StatusText $StatusText -IsBusy:$true -Status (Get-String 'StatusLoading')
    Add-LogLine -LogBox $LogBox -Message (Get-String 'LogLoading')

    $scriptText = @'
param()

$ErrorActionPreference = 'Stop'

function Invoke-WingetUtf8 {
    param([Parameter(Mandatory)] [string[]] $WingetArgs)

    $wingetPath = (Get-Command winget -ErrorAction Stop).Source

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $wingetPath
    $psi.Arguments = ($WingetArgs -join ' ')
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true

    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $psi.StandardOutputEncoding = $utf8
    $psi.StandardErrorEncoding = $utf8

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi

    if (-not $p.Start()) {
        throw 'Failed to start winget.'
    }

    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    $p.WaitForExit()

    return [pscustomobject]@{
        ExitCode = $p.ExitCode
        Stdout   = $stdout
        Stderr   = $stderr
        Output   = ($stdout + $stderr)
    }
}

function Get-WingetUpgradesJson {
    $resp = Invoke-WingetUtf8 -WingetArgs @('upgrade','--output','json','--disable-interactivity')
    if ($resp.ExitCode -ne 0) { throw "winget upgrade failed (exit $($resp.ExitCode)): $($resp.Output)" }
    return ($resp.Stdout | ConvertFrom-Json -Depth 20)
}

function Convert-WingetJsonToRows {
    param([Parameter(Mandatory)] $WingetJson)
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($src in @($WingetJson.Sources)) {
        $srcName = [string]$src.Name
        foreach ($p in @($src.Packages)) {
            $rows.Add([pscustomobject]@{
                Name = [string]$p.PackageName
                Id = [string]$p.PackageIdentifier
                Version = [string]$p.InstalledVersion
                AvailableVersion = [string]$p.AvailableVersion
                Source = if (-not [string]::IsNullOrWhiteSpace([string]$p.Source)) { [string]$p.Source } else { $srcName }
            })
        }
    }
    return $rows
}

function Convert-WingetUpgradeTableToRows {
    param([Parameter(Mandatory)] [string] $Text)

    $rows = New-Object System.Collections.Generic.List[object]
    $lines = @($Text -split "`r?`n")

    # Find separator line under the header (----  ----  ---- ... or continuous dashes)
    $sepIndex = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^-{10,}') { $sepIndex = $i; break }
    }

    if ($sepIndex -lt 0) {
        return $rows
    }

    for ($i = $sepIndex + 1; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        # Skip obvious non-table lines
        if (
            $line -match '^Weitere Hilfe' -or
            $line -match '^Verwendung:' -or
            $line -match '^Die folgenden' -or
            $line -match 'Aktualisierungen verf' -or
            $line -match '^Mindestens\s+\d+\s+Paket' -or
            $line -match '^\d+\s+Aktualisierung'
        ) { break }

        $parts = @($line -split '\s{2,}' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($parts.Count -lt 4) { continue }

        # Expect either:
        #  - Name | Id | Version | Verf fcgbar
        #  - Name | Id | Version | Verf fcgbar | Quelle
        if ($parts.Count -eq 4) {
            $rows.Add([pscustomobject]@{
                Name = [string]$parts[0]
                Id = [string]$parts[1]
                Version = [string]$parts[2]
                AvailableVersion = [string]$parts[3]
                Source = 'winget'
            })
            continue
        }

        if ($parts.Count -ge 5) {
            $nameParts = $parts[0..($parts.Count - 5)]
            $name = ($nameParts -join ' ').Trim()
            $id = [string]$parts[$parts.Count - 4]
            $version = [string]$parts[$parts.Count - 3]
            $available = [string]$parts[$parts.Count - 2]
            $source = [string]$parts[$parts.Count - 1]

            $rows.Add([pscustomobject]@{
                Name = $name
                Id = $id
                Version = $version
                AvailableVersion = $available
                Source = $source
            })
        }
    }

    return $rows
}

function Get-WingetUpgradesRows {
    $help = (Invoke-WingetUtf8 -WingetArgs @('upgrade','--help')).Output
    $supportsJson = ($help -match '(?m)^\s*--output\b')

    if ($supportsJson) {
        $json = Get-WingetUpgradesJson
        return (Convert-WingetJsonToRows -WingetJson $json)
    }

    # WinGet v1.12 has no `--output json`; to keep output stable and fast, use winget source.
    $resp = Invoke-WingetUtf8 -WingetArgs @('upgrade','--source','winget','--accept-source-agreements','--disable-interactivity')
    if ($resp.ExitCode -ne 0) {
        throw "winget upgrade failed (exit $($resp.ExitCode)): $($resp.Output)"
    }

    return (Convert-WingetUpgradeTableToRows -Text $resp.Stdout)
}

Get-WingetUpgradesRows
'@

    Start-AsyncOperation -Kind 'refresh' -ScriptText $scriptText
}

function Select-All([bool] $value) {
    foreach ($p in $Packages) { $p.Selected = $value }
}

$RefreshButton.Add_Click({ Update-PackageList })
$SelectAllButton.Add_Click({ Select-All $true })
$SelectNoneButton.Add_Click({ Select-All $false })
$CancelButton.Add_Click({ Stop-UpdateOperation })

# Log Toggle Button - show/hide log panel
$LogToggleButton.Add_Click({
    if ($LogToggleButton.IsChecked) {
        $LogPanel.Visibility = 'Visible'
        $LogBox.ScrollToEnd()
    } else {
        $LogPanel.Visibility = 'Collapsed'
    }
})

$UpdateSelectedButton.Add_Click({
    $selected = @($Packages | Where-Object { $_.Selected })
    if ($selected.Count -eq 0) {
        Add-LogLine -LogBox $LogBox -Message (Get-String 'LogNoSelection')
        return
    }

    Set-ButtonsEnabled -enabled:$false
    Set-Busy -BusyBar $BusyBar -StatusText $StatusText -IsBusy:$false -Status (Get-String 'StatusUpdating')
    Add-LogLine -LogBox $LogBox -Message ((Get-String 'LogStartingUpdates') -f $selected.Count)
    
    # Show cancel button during update
    $CancelButton.Visibility = 'Visible'
    
    # Reset all status fields and progress tracking
    foreach ($pkg in $Packages) {
        $pkg.UpdateStatus = ""
    }
    $script:UpdateLastProcessedIndex = -1
    
    # Setup progress bar for package progress
    $BusyBar.IsIndeterminate = $false
    $BusyBar.Maximum = 100
    $BusyBar.Value = 0

    $ids = @($selected | ForEach-Object { [string]$_.Id })
    $namesById = @{}
    foreach ($pkg in $selected) { $namesById[[string]$pkg.Id] = [string]$pkg.Name }
    
    $installMode = $script:InstallMode
    $forceInstall = $script:ForceInstall
    
    # Get localized strings for the runspace
    $localizedStarting = Get-String 'LogStarting'
    $localizedRetrying = Get-String 'LogRetrying'
    
    # Create synchronized hashtable for progress communication
    $script:UpdateProgress = [hashtable]::Synchronized(@{
        CurrentPackageIndex = 0
        TotalPackages = $ids.Count
        LocalizedStarting = $localizedStarting
        LocalizedRetrying = $localizedRetrying
        CurrentPackageId = ''
        CurrentPackageName = ''
        CurrentPercent = 0
        CurrentStatus = ''
        LastError = ''
        Results = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
        Completed = $false
        Cancelled = $false
    })

    $scriptText = @'
param($Ids, $NamesById, $InstallMode, $ForceInstall, $Progress)

$ErrorActionPreference = 'Stop'

function Invoke-WingetUpgradeByIdWithProgress {
    param(
        [Parameter(Mandatory)] [string] $Id,
        [string] $Mode = 'none',
        [bool] $Force = $false,
        [hashtable] $Progress
    )
    
    $wingetPath = (Get-Command winget -ErrorAction Stop).Source
    
    $upgradeArgs = @('upgrade','--id',$Id,'--accept-source-agreements','--accept-package-agreements','--include-unknown')
    
    switch ($Mode) {
        'silent' { $upgradeArgs += '--silent' }
        'interactive' { $upgradeArgs += '--interactive' }
        default { $upgradeArgs += '--disable-interactivity' }
    }
    
    if ($Force) {
        $upgradeArgs += '--force'
    }
    
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $wingetPath
    $psi.Arguments = ($upgradeArgs -join ' ')
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $psi.StandardOutputEncoding = $utf8
    $psi.StandardErrorEncoding = $utf8
    
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $psi
    
    $outputBuilder = New-Object System.Text.StringBuilder
    $errorBuilder = New-Object System.Text.StringBuilder
    $lastError = ""
    
    if (-not $p.Start()) {
        throw 'Failed to start winget.'
    }
    
    # Read output line by line for progress
    while (-not $p.StandardOutput.EndOfStream) {
        if ($Progress.Cancelled) {
            try { $p.Kill() } catch {}
            break
        }
        
        $line = $p.StandardOutput.ReadLine()
        if ($line -ne $null) {
            $null = $outputBuilder.AppendLine($line)
            
            # Parse progress percentage (e.g., "  ██████████▒▒▒▒▒▒▒▒▒▒  34%")
            if ($line -match '(\d+)%\s*$') {
                $Progress.CurrentPercent = [int]$Matches[1]
            }
            
            # Parse status messages
            if ($line -match 'Installer-Hash|Hash verified|Paketinstallation|Installation started|wird heruntergeladen|Downloading|Installing') {
                $Progress.CurrentStatus = $line.Trim()
            }
            
            # Capture error messages
            if ($line -match 'Fehler|Error|failed|fehlgeschlagen|0x[0-9a-fA-F]+') {
                $lastError = $line.Trim()
            }
        }
    }
    
    # Read any remaining stderr
    $stderr = $p.StandardError.ReadToEnd()
    if (-not [string]::IsNullOrWhiteSpace($stderr)) {
        $null = $errorBuilder.Append($stderr)
        if ($stderr -match 'Fehler|Error|failed|0x[0-9a-fA-F]+') {
            $lastError = $stderr.Trim()
        }
    }
    
    $p.WaitForExit()
    
    $Progress.LastError = $lastError
    
    return [pscustomobject]@{
        ExitCode = $p.ExitCode
        Output = ($outputBuilder.ToString() + $errorBuilder.ToString()).Trim()
        Error = $lastError
    }
}

# Fallback function for packages that need UAC elevation - batch mode (single UAC prompt)
function Invoke-WingetUpgradeElevatedBatch {
    param(
        [Parameter(Mandatory)] [array] $PackageIds,
        [hashtable] $NamesById,
        [bool] $Force = $false,
        [hashtable] $Progress
    )
    
    # Validate all package IDs to prevent command injection
    # Valid winget IDs: alphanumeric, dots, hyphens, underscores
    foreach ($pkgId in $PackageIds) {
        if ($pkgId -notmatch '^[a-zA-Z0-9._-]+$') {
            $name = if ($NamesById.ContainsKey($pkgId)) { [string]$NamesById[$pkgId] } else { $pkgId }
            $null = $Progress.Results.Add([pscustomobject]@{
                Id = $pkgId
                Name = $name
                ExitCode = -1
                Output = ""
                Error = "Invalid package ID format - skipped for security"
            })
        }
    }
    
    # Filter to only valid IDs
    $validIds = @($PackageIds | Where-Object { $_ -match '^[a-zA-Z0-9._-]+$' })
    if ($validIds.Count -eq 0) { return }
    
    # Create temp BAT file for safe execution
    $tempBat = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "winget_batch_$([guid]::NewGuid().ToString('N')).bat")
    
    try {
        # Build BAT file content
        $batLines = @('@echo off')
        foreach ($pkgId in $validIds) {
            $cmd = "winget upgrade --id `"$pkgId`" --accept-source-agreements --accept-package-agreements --include-unknown --disable-interactivity"
            if ($Force) { $cmd += " --force" }
            $batLines += $cmd
        }
        $batLines += 'pause'
        
        # Write BAT file
        $batLines | Set-Content -Path $tempBat -Encoding ASCII
        
        # Run elevated
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $tempBat
        $psi.UseShellExecute = $true
        $psi.Verb = 'runas'
        $psi.WindowStyle = 'Normal'
        
        $p = New-Object System.Diagnostics.Process
        $p.StartInfo = $psi
        
        $started = $false
        try {
            $started = $p.Start()
        } catch {
            $started = $false
        }
        
        if (-not $started) {
            # UAC was cancelled
            foreach ($pkgId in $validIds) {
                $name = if ($NamesById.ContainsKey($pkgId)) { [string]$NamesById[$pkgId] } else { $pkgId }
                $null = $Progress.Results.Add([pscustomobject]@{
                    Id = $pkgId
                    Name = $name
                    ExitCode = -2147023673
                    Output = ""
                    Error = "UAC cancelled"
                })
            }
            return
        }
        
        $p.WaitForExit()
        
        # Add results for all packages
        foreach ($pkgId in $validIds) {
            $name = if ($NamesById.ContainsKey($pkgId)) { [string]$NamesById[$pkgId] } else { $pkgId }
            $null = $Progress.Results.Add([pscustomobject]@{
                Id = $pkgId
                Name = $name
                ExitCode = 0
                Output = "Elevated batch install completed"
                Error = ""
            })
            $Progress.CurrentPackageIndex++
        }
    } finally {
        # Cleanup temp BAT file
        if (Test-Path $tempBat) {
            Remove-Item -Path $tempBat -Force -ErrorAction SilentlyContinue
        }
    }
}

# Track packages that need elevation for batch processing
$packagesNeedingElevation = [System.Collections.ArrayList]::new()

$index = 0
foreach ($id in @($Ids)) {
    if ($Progress.Cancelled) { break }
    
    $name = if ($NamesById.ContainsKey($id)) { [string]$NamesById[$id] } else { $id }
    
    $Progress.CurrentPackageIndex = $index
    $Progress.CurrentPackageId = $id
    $Progress.CurrentPackageName = $name
    $Progress.CurrentPercent = 0
    $Progress.CurrentStatus = $Progress.LocalizedStarting
    $Progress.LastError = ""
    
    $res = Invoke-WingetUpgradeByIdWithProgress -Id $id -Mode $InstallMode -Force $ForceInstall -Progress $Progress
    
    # If failed with certain error codes, queue for batch elevated install
    # -2147023673 = ERROR_CANCELLED (often UAC related)
    # -1978335189 = APPINSTALLER_CLI_ERROR_INSTALL_SYSTEM_NOT_SUPPORTED
    # 0x80070005 = E_ACCESSDENIED
    if ($res.ExitCode -ne 0) {
        $needsElevation = ($res.ExitCode -eq -2147023673) -or 
                         ($res.ExitCode -eq -1978335189) -or 
                         ($res.ExitCode -eq -2147024891) -or  # 0x80070005 as signed int
                         ($res.Output -match 'Administrator|admin|elevation|UAC|Zugriff verweigert|Access denied')
        
        if ($needsElevation) {
            # Queue for batch elevated install instead of individual UAC prompt
            $null = $packagesNeedingElevation.Add([pscustomobject]@{
                Id = $id
                Name = $name
            })
            $index++
            continue  # Skip adding to results now, will be added after batch elevation
        }
    }
    
    $null = $Progress.Results.Add([pscustomobject]@{
        Id = $id
        Name = $name
        ExitCode = $res.ExitCode
        Output = $res.Output
        Error = $res.Error
    })
    
    $index++
}

# Process all packages that need elevation in a single batch (one UAC prompt)
if ($packagesNeedingElevation.Count -gt 0 -and -not $Progress.Cancelled) {
    $Progress.CurrentStatus = $Progress.LocalizedRetrying
    $Progress.CurrentPercent = 0
    
    $elevatedIds = @($packagesNeedingElevation | ForEach-Object { $_.Id })
    Invoke-WingetUpgradeElevatedBatch -PackageIds $elevatedIds -NamesById $NamesById -Force $ForceInstall -Progress $Progress
}

$Progress.Completed = $true
'@

    Start-AsyncOperation -Kind 'update' -ScriptText $scriptText -Arguments @($ids, $namesById, $installMode, $forceInstall, $script:UpdateProgress)
})

# Initial load
$window.Add_ContentRendered({ Update-PackageList })

# Show
$window.ShowDialog() | Out-Null
