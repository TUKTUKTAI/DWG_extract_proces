# DXF Pipeline (NL)

Deze pipeline is een lokale/prototype vervanging voor de Meridian/Mosaic DataExtract-stap:

1. `DWG -> DXF` (via ODA File Converter)
2. `DXF -> extract .xlsx -> Java naverwerking`
3. (Optioneel) samenvoegen naar 1 eindbestand

Doel:
- zo min mogelijk handwerk
- duidelijke vaste mappen
- fouttolerant per stap

## Benodigd (eenmalig)

- ODA File Converter (geinstalleerd)
- Python packages:

```powershell
python -m pip install -U -r .\broncode\tools\requirements-dxf-pipeline.txt
```

## ODA op een andere pc (pad / installatie)

De scripts gebruiken standaard `pipeline_defaults.ps1` en proberen daar automatisch een ODA-installatie te vinden (auto-detect van veelvoorkomende paden).

### Als ODA direct werkt

Dan hoef je niets aan te passen.

### ODA-pad snel vinden (Windows)

Als je het exacte pad van `ODAFileConverter.exe` nodig hebt, dit werkt meestal het snelst:

Alternatief via PowerShell (zoekt in de standaard ODA-map):

```powershell
Get-ChildItem "C:\Program Files\ODA" -Recurse -Filter ODAFileConverter.exe -ErrorAction SilentlyContinue |
  Select-Object -ExpandProperty FullName
```

### Als ODA niet gevonden wordt

Je hebt dan 2 opties:

1. **Eenmalig aanpassen in code (aanbevolen voor vaste setup)**
   - open `.\broncode\tools\pipeline_defaults.ps1`
   - zoek de lijst `odaCandidates` (bovenin het bestand)
   - voeg jouw volledige pad toe, bijvoorbeeld:

```powershell
"D:\Tools\ODA\ODAFileConverter 26.12.0\ODAFileConverter.exe",
```

   - bij voorkeur bovenaan de lijst zetten, zodat die eerst geprobeerd wordt

2. **Pad meegeven via parameter (handig voor testen)**

```powershell
powershell -ExecutionPolicy Bypass -File ".\broncode\tools\convert_dwg_to_dxf.ps1" -OdaExe "D:\Tools\ODA\ODAFileConverter.exe"
```

### Eerste keer op een nieuwe pc

Open ODA File Converter 1 keer handmatig als dat nodig is (bijv. voor first-run/licentie-instellingen). Daarna werken de scripts meestal direct.

## Verwachte projectmappen

Run de commando's vanuit de projectroot (de map met):
- `assets`
- `broncode`
- `dwgs`
- `DXF`
- `Bron`
- `Doel`
- `Eindresultaat`
- `jdk-13.0.2`

De scripts gebruiken standaardinstellingen uit:
- `broncode/tools/pipeline_defaults.ps1`

Daar staan o.a.:
- input DWG map (`.\dwgs`)
- output DXF map (`.\DXF`)
- extract output (`.\Bron`)
- Java naverwerking output (`.\Doel`)
- merge output (`.\Eindresultaat`)
- ODA pad + versie

## Snelstart (aanbevolen, 3 losse stappen)

Ga eerst naar de projectroot:

```powershell
cd "C:\Users\tarun\Desktop\NBD dataextractie"
```

### 1) DWG -> DXF (batch, alles in 1 keer)

Dit is de simpele/klassieke ODA batch-run (alles tegelijk):

```powershell
powershell -ExecutionPolicy Bypass -File .\broncode\tools\convert_dwg_to_dxf.ps1
```

### 1a) Alternatief: DWG -> DXF per bestand (meer foutisolatie)

Deze variant draait ODA per DWG en logt fouten naar `.\Doel\oda_per_file_errors.csv`:

```powershell
powershell -ExecutionPolicy Bypass -File .\broncode\tools\convert_dwg_to_dxf_per_file.ps1
```

Gebruik deze variant als stabiliteit belangrijker is dan snelheid.

### 2) DXF -> extract -> Java naverwerking

Dit script doet:
- `DXF -> DataExtract-achtige .xlsx` (naar `.\Bron`)
- Java compile + run (`P22_0002_Main.java`)
- naverwerking-output naar `.\Doel`

```powershell
powershell -ExecutionPolicy Bypass -File .\broncode\tools\run_dxf_pipeline.ps1
```

Met automatische merge op het einde (optioneel):

```powershell
powershell -ExecutionPolicy Bypass -File .\broncode\tools\run_dxf_pipeline.ps1 -MergeAfter
```

### 3)  merge van alle losse Java-outputbestanden

Voegt alle `Extraheren resultaat ...xlsx` bestanden uit `.\Doel` samen naar 1 bestand in `.\Eindresultaat`.

```powershell
python .\broncode\tools\merge_naverwerking_results.py
```

## Wrapper (optioneel, 1 script voor handmatig gebruik)

Als je de pipeline handmatig in 1 command wilt draaien (intern nog steeds via DXF):

```powershell
powershell -ExecutionPolicy Bypass -File ".\broncode\tools\run_dwg_to_excel.ps1"
```

Met automatische merge op het einde:

```powershell
powershell -ExecutionPolicy Bypass -File ".\broncode\tools\run_dwg_to_excel.ps1" -MergeAfter
```

Let op:
- `-MergeAfter` werkt op zowel `run_dwg_to_excel.ps1` als `run_dxf_pipeline.ps1`
- Voor FME blijft 3 losse stappen meestal het duidelijkst

## Resultaat (output)

Na stap 1 + 2:
- `.\DXF\` = geconverteerde DXF-bestanden
- `.\Bron\` = gegenereerde extract `.xlsx` bestanden
- `.\Doel\` = `Extraheren resultaat ...xlsx` bestanden (Java output)

Na stap 3:
- `.\Eindresultaat\Extraheren resultaat SAMENGEVOEGD.xlsx`

## Foutafhandeling (logbestanden)

- ODA per-file fouten:
  - `.\Doel\oda_per_file_errors.csv`
- ODA batch `.err` bestanden (door ODA aangemaakte foutbestanden in `DXF`):
  - `.\Doel\oda_batch_errors.csv`
- DXF parse / validatie fouten:
  - `.\Doel\dxf_extract_errors.csv`
- Java naverwerking fouten:
  - `.\Doel\errors.csv`
- Merge fouten:
  - `.\Eindresultaat\merge_errors.csv`

Bij fouten in een bestand hoort de batch door te gaan met de rest.

Opmerking:
- ODA batch-conversie kan naast `.dxf` ook `.err` bestanden in `.\DXF` aanmaken (bij problematische input).
- Dat verklaart waarom het totaal aantal bestanden in `.\DXF` soms hoger is dan het aantal `.dwg` bestanden.
- Deze `.err` bestanden worden nu expliciet gemeld en gelogd.

## Opmerking over validatie

`run_dxf_pipeline.ps1` gebruikt standaard `Strict NBD` validatie (via defaults), zodat niet-NBD/rare DXF's sneller worden afgekeurd en gelogd.

## FME Configuratie (aanbevolen: 3 SystemCallers)

Gebruik in FME 3 losse `SystemCaller` stappen. Dat geeft betere controle per stap en maakt foutanalyse eenvoudiger.

### SystemCaller 1 - DWG -> DXF

Batch ODA-run:

```powershell
powershell -ExecutionPolicy Bypass -File ".\broncode\tools\convert_dwg_to_dxf.ps1"
```

Of per-file ODA-run (meer foutisolatie):

```powershell
powershell -ExecutionPolicy Bypass -File ".\broncode\tools\convert_dwg_to_dxf_per_file.ps1"
```

### SystemCaller 2 - DXF pipeline (extract + Java naverwerking)

```powershell
powershell -ExecutionPolicy Bypass -File ".\broncode\tools\run_dxf_pipeline.ps1"
```

### SystemCaller 3 - Merge naverwerking-output

```powershell
python ".\broncode\tools\merge_naverwerking_results.py"
```

### Aanbevolen FME instellingen

- Working directory (heel belangrijk):
  - `C:\Users\tarun\Desktop\NBD dataextractie`
- Gebruik in `SystemCaller` bij voorkeur relatieve paden (zoals hierboven), zodat de workflow overdraagbaar blijft naar collega's/andere werkstations met dezelfde projectstructuur.
- Bewaak na elke stap de juiste logbestanden (zie hierboven)
- Gebruik per-file ODA script als batchstabiliteit belangrijker is
- FME Form (desktop) is meestal makkelijker dan FME Flow voor ODA (extern programma)
