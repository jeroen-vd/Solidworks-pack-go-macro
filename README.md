# SOLIDWORKS Pack & Rename Macro

Deze repository bevat een VBA-macro die het standaard Pack & Go-proces uitbreidt met automatische naamgeving voor assemblies, parts en tekeningen. De macro kan direct in de SOLIDWORKS VBA-editor worden geplakt of als .frm/.bas-bestanden worden geïmporteerd.

## Functionaliteit
- Prefix instelbaar via pop-up (bijv. `XXX`).
- Automatische nummering:
  - Bovenste assembly: `PREFIX-A00`.
  - Subassemblies: `PREFIX-A01`, `A02`, ...
  - Parts: `PREFIX-P01`, `P02`, ... (gelijke parts delen hetzelfde nummer).
- Optioneel tekeningen meenemen; tekeningen krijgen dezelfde basisnaam als het bijbehorende model.
- Alles wordt naar één doelmap gekopieerd, vergelijkbaar met Pack & Go.

## Gebruiksaanwijzing
1. Open de bovenste assembly in SOLIDWORKS.
2. Importeer `macros/PackAndRename.bas` en `macros/PackAndRenameForm.frm` in de VBA-editor (Tools → Macro → New/ Edit → File → Import File).
3. Start de macro (`Main`).
4. Vul in de pop-up:
   - **Prefix**: gedeelde prefix voor alle bestanden.
   - **Exportmap**: doelmap voor de Pack & Go-uitvoer.
   - **Tekeningen meenemen**: vink aan indien je tekeningen wilt exporteren.
5. Klik **Uitvoeren** om het Pack & Go-proces met nieuwe namen te starten.

## Aanpassen
- Pas eventueel de startnummers aan in `PackAndRename.bas` (variabelen `partCounter` en `asmCounter`).
- De helperfunctie `BuildNumber` bepaalt het nummerformaat (`00`).

## Vereisten
- SOLIDWORKS desktop met VBA-macro-ondersteuning.
- Toegang tot de doelmap (schrijfrechten).
