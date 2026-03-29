# SPI Open Access Retro

Freie Retro-Reimplementation eines historischen integrierten Office-Pakets im Stil von SPI Open Access.

Wichtig:
- Dieses Repository ist ein eigenstaendiges Community-Projekt.
- Es steht in keiner Verbindung zum urspruenglichen Hersteller.
- Historische Produktnamen dienen nur zur Einordnung und Kompatibilitaetsbeschreibung.

Die aktuelle Version liefert einen lauffaehigen Textmodus-Host mit separaten Modulen fuer:

- Desktop-Datenbank
- Tabellenkalkulation
- Textverarbeitung
- Kommunikation
- Mail
- Reporting
- PRO-inspirierte Programmierung

## Starten

```powershell
dotnet run --project .\src\SpiOpenAccess.App
```

Einzelnes Modul direkt anzeigen:

```powershell
dotnet run --project .\src\SpiOpenAccess.App -- db
```

Tests:

```powershell
dotnet test .\SpiOpenAccess.sln
```

## Architektur

- `src/SpiOpenAccess.Core`: Kernabstraktionen, Workspace und Suite-Modell
- `src/SpiOpenAccess.Infrastructure`: Suite-Zusammenbau und Default-Workspace
- `src/SpiOpenAccess.Modules.*`: Fachmodule fuer das Office-Paket
- `tests/SpiOpenAccess.Tests`: Basistests fuer Registrierung und Kernlogik

## Projektziel

Das Ziel ist kein moderner Office-Klon, sondern eine oeffentliche Retro-Reimplementation mit:

- DOS-aehnlicher Bedienoberflaeche
- Datenbank-, Formular- und Reporting-Kern
- Tabellenkalkulation, Textverarbeitung, Mail und Kommunikation
- PRO-inspirierter Skriptfaehigkeit
- schrittweiser historischer Atmosphaere statt rein moderner UI

## Veroeffentlichung

Geeignet fuer ein oeffentliches GitHub-Repository als Retro-/Preservation-Projekt. Vor einer breiteren Veroeffentlichung sollten originale Marken, Assets, Hilfetexte und Binardaten des historischen Produkts nicht unbesehen uebernommen werden.

## Naechste Ausbaustufen

- echte Dateiformate fuer Tabellen, Dokumente und Reports
- persistente PRO-Laufzeit mit Parser/AST/VM
- Formular- und Maskendesigner
- Druckspooler und Netzwerk-/Locking-Modell
- Mailbox-Store, Kommunikationsadapter und Batch-Ausfuehrung
