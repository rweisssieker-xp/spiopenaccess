# Contributing

## Ziel

Dieses Projekt baut eine freie Retro-Reimplementation eines historischen Office-Pakets. Beitraege sollten auf Kompatibilitaet, Testbarkeit und historisches Bediengefuehl einzahlen.

## Regeln

- Keine originalen proprietaeren Quelltexte, Assets oder Hilfedateien uebernehmen.
- Neue Funktionen mit Tests absichern, wenn Kernlogik betroffen ist.
- Oberflaechen aendern nur, wenn sie das Retro-Ziel staerken oder technische Schulden reduzieren.
- Oeffentliche APIs und Dateiformate moeglichst stabil halten.

## Setup

```powershell
dotnet test .\SpiOpenAccess.sln
```

```powershell
dotnet run --project .\src\SpiOpenAccess.App
```
