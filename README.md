# Invoice Compare Tool 📊

Værktøj til at sammenligne PDF-filer fra Azure Blob Storage med fakturaer fra Excel (Autocore-udtræk).

## Funktionalitet

- ✅ Læser fakturaer fra Excel-fil
- ✅ Downloader PDF'er fra Azure Blob Storage baseret på dealer-koder
- ✅ Parses PDF'er for at ekstrahere:
  - Fakturadato
  - Moms-beløb
  - Totalt beløb ("i alt")
  - Kundenummer
- ✅ Sammenligner data fra Excel vs PDF
- ✅ Detekterer kreditnotaer automatisk
- ✅ Genererer farvekodet Excel-rapport med matches og afvigelser

## Installation

### 1. Klon repositoryet
```bash
git clone https://github.com/mojoh-semler/Invoice-comparetool.git
cd Invoice-comparetool
```

### 2. Opret virtuel miljø
```bash
python -m venv venv
source venv/bin/activate  # Mac/Linux
# eller
venv\Scripts\activate  # Windows
```

### 3. Installer dependencies
```bash
pip install -r requirements.txt
```

### 4. Konfigurer Azure-forbindelse
Kopier `.env.example` til `.env` og udfyld dine Azure-credentials:
```bash
cp .env.example .env
```

Rediger `.env`:
```
AZURE_STORAGE_ACCOUNT_NAME=documentserver
AZURE_SAS_TOKEN=your_sas_token_here
DOWNLOAD_FOLDER=downloads
```

## Brug

### Grundlæggende kørsel
```bash
python rapport_FINAL_v24.py
```

### Inputfiler
- **Excel-fil (v24)**: `invoice_extract_prod.xlsx` i den konfigurerede `BATCH_FOLDER` med AutoCore-kolonner:
  - `INVOICENO` → `faktura nr` - Unik fakturaidentifikator
  - `KUNNUM` → `kunde nr` - Kundenummer
  - `FAKTURADATO` → `faktura dato` - Dato på fakturaen
  - `MOMS_BELØB_ENDLIG` → `moms` - Moms-beløb
  - `FAKBLØ` → `i alt` - Totalt beløb

v24 normaliserer og omdøber kolonnerne ved indlæsning, før dublet-rensning og sammenligning. Resten af behandlingen er uændret fra v23.

### Output
- **Excel-rapport**: `sammenligningsrapport_finance.xlsx`
  - 🟢 Grøn = Match uden forskelle
  - 🔴 Rød = Forskel mellem Excel og PDF
  - 🟡 Gul = Kreditnota
- **Log-fil**: `rapport_log_YYYYMMDD_HHMMSS.txt`
- **Downloaded PDFer**: I `downloaded_pdfs/` mappen

## Rapport-output

Rapporten indeholder følgende kolonner:
| Kolonne | Beskrivelse |
|---------|-------------|
| Faktura nr | Fakturaens unikke ID |
| PDF fil | Filnavn på downloaded PDF |
| Type | Faktura eller Kreditnota |
| Kunde nr Excel | Kundenummer fra Excel |
| Kunde nr PDF | Kundenummer ekstraheret fra PDF |
| Faktura dato Excel | Dato fra Excel |
| Faktura dato PDF | Dato ekstraheret fra PDF |
| moms Excel | Moms fra Excel |
| moms PDF | Moms ekstraheret fra PDF |
| i alt Excel | Totalt beløb fra Excel |
| i alt PDF | Totalt beløb ekstraheret fra PDF |
| Kommentar | Beskrivelse af evt. forskelle |

## Features

### Dato-normalisering
Håndterer flere dato-formater:
- `DD.MM.YYYY`, `DD-MM-YYYY`, `DD/MM/YYYY`
- `YYYY.MM.DD`, `YYYY-MM-DD`, `YYYY/MM/DD`
- `DDMMYYYY`, `YYYYMMDD`

### Dansk talformat
Konverterer danske talformater automatisk:
- `1.234,56` → `1234.56`
- `1234,56` → `1234.56`

### Kreditnota-detektion
Detekterer automatisk kreditnotaer og negerer beløb hvis nødvendigt.

### Dublet-håndtering
Fjerner dubletter baseret på fakturanummer, bevarer rækker med mest data.

## Fejlhåndtering

Scriptet logger alle fejl og advarsler:
- PDF'er uden dato
- PDF'er uden beløb
- PDF'er uden moms
- Fakturaer uden matchende PDF
- Azure-forbindelsesfejl

Se `rapport_log_*.txt` for detaljer.

## Arkitektur

```
rapport_FINAL_v21.py
├── Azure Blob Storage (Download PDFer)
├── Local PDF Directory (Parsing)
├── Excel Reader (Input)
└── Excel Report Generator (Output)
```

### Nøgle-komponenter

1. **PDF Matching** - Matcher PDF-filnavn med fakturanumre baseret på dealer-koder
2. **Date Extraction** - Finder fakturadatoer i PDF-tekst
3. **Amount Extraction** - Parses beløb (moms, total) fra PDF
4. **Data Comparison** - Validerer matches mellem Excel og PDF
5. **Report Generation** - Laver farvekodet Excel-rapport

## Opsummering efter kørsel

Efter scriptets afslutning vises:
```
Total fakturaer i Excel:           474
  ✅ Match uden forskelle:         414
  🟡 Match MED forskelle:          60
  ❌ Uden matchende PDF:           0
  📋 Heraf kreditnotaer:           X
```

## Fejlfinding

**Problem**: "Tjek .env!" fejl
- **Løsning**: Sørg for at `.env` eksisterer med `AZURE_STORAGE_ACCOUNT_NAME` og `AZURE_SAS_TOKEN`

**Problem**: Ingen PDF'er downloaded
- **Løsning**: Kontroller at SAS token er gyldig og har passende permissions (Read, List, Create)

**Problem**: Moms ekstraheres ikke fra nogle PDF'er
- **Løsning**: PDF'en kan have et uventet format. Se `rapport_log_*.txt` for detaljer

## Udvikling

### Kommende features
- [ ] Database-integrering for persistent lagring
- [ ] Batch-rapport med gruppering pr. dealer
- [ ] Email-notifikation ved fejl
- [ ] Web-interface for upload og visning
- [ ] Incremental updates (kun nye fakturaer)

## Licens
Privat projekt

## Kontakt
mojoh@semler.dk
