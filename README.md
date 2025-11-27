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

## Snelle testexport uitvoeren
Wil je snel controleren of de macro in jouw SOLIDWORKS-omgeving correct exporteert? Gebruik dan deze stappen:

1. Maak een tijdelijke map aan, bijvoorbeeld `C:\Temp\PackRenameTest`.
2. Open een kleine testassembly (een paar parts/subassemblies is genoeg) en sla deze op zodat het bestandspad bekend is.
3. Start de macro en vul in de pop-up:
   - **Prefix**: kies bijvoorbeeld `TEST`.
   - **Exportmap**: kies de zojuist aangemaakte map.
   - **Tekeningen meenemen**: vink alleen aan als er tekeningen bij je test horen.
4. Klik **Uitvoeren**. In de exportmap verschijnt één pakket met automatisch hernoemde bestanden (bijv. `TEST-A00`, `TEST-A01`, `TEST-P01`, ...). De nieuwe namen kun je direct openen in SOLIDWORKS om te controleren of alles klopt.

## Macro omzetten naar .SWP of .dll
SOLIDWORKS accepteert kant-en-klare macro’s als `.swp` (gecompileerde macro) of als COM-invoegtoepassing (`.dll`). Deze repository bevat de bronbestanden (`.bas` en `.frm`). Zo maak je er zelf een bruikbaar bestand van in SOLIDWORKS:

### Opslaan als .SWP
1. Open SOLIDWORKS en kies **Tools → Macro → New...**.
2. Sla het lege macro-bestand op als `PackAndRename.swp` (willekeurige locatie).
3. De VBA-editor opent. Importeer de bestanden uit deze repository via **File → Import File**:
   - `macros/PackAndRename.bas`
   - `macros/PackAndRenameForm.frm`
4. Controleer of de module `PackAndRename` en het formulier `PackAndRenameForm` zichtbaar zijn.
5. Kies **File → Save** om de macro in dezelfde `.swp` op te slaan. Dit `.swp`-bestand kun je daarna direct laden en uitvoeren via **Tools → Macro → Run...**.

### Bouwen als .dll (optioneel)
Alleen nodig als je een COM-add-in wilt met extra UI-logica. Stappen in hoofdlijnen:
1. Open Visual Studio en maak een nieuw **Class Library**-project (vb.net of C#) met **.NET Framework** dat past bij jouw SOLIDWORKS-versie.
2. Voeg referenties toe aan de SOLIDWORKS interop-assemblies (`SolidWorks.Interop.sldworks`, `SolidWorks.Interop.swconst`, `SolidWorks.Interop.swpublished`).
3. Kopieer de logica uit `macros/PackAndRename.bas` naar een klas en schrijf een `ISwAddin`-implementatie die de functionaliteit aanroept.
4. Compileer de `.dll` en registreer deze via **Tools → Add-ins → Add...** of met `regasm.exe`.

## Aanpassen
- Pas eventueel de startnummers aan in `PackAndRename.bas` (variabelen `partCounter` en `asmCounter`).
- De helperfunctie `BuildNumber` bepaalt het nummerformaat (`00`).

## Vereisten
- SOLIDWORKS desktop met VBA-macro-ondersteuning.
- Toegang tot de doelmap (schrijfrechten).
