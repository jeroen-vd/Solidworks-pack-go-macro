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
2. Ga naar **Tools → Macro → New...** en sla een lege macro op als `PackAndRename.swp` (of andere naam).
3. De VBA-editor opent. Zorg dat **Project Explorer** zichtbaar is (View → Project Explorer). Klik in de tree op je nieuwe `.swp`.
4. Importeer via **File → Import File** achtereenvolgens:
   - `macros/PackAndRename.bas`
   - `macros/PackAndRenameForm.frm`
5. Verwijder lege standaarditems (bijv. `Module1`) zodat alleen `PackAndRename` (module) en `PackAndRenameForm` (UserForm) overblijven. Als je ze niet ziet: klik in Project Explorer op het project en kies **View → Code** of dubbelklik de module.
6. Ga naar **File → Save**. Je kunt de macro nu direct runnen zonder opnieuw te openen.
7. Start de macro (`Main`).
8. Vul in de pop-up:
   - **Prefix**: gedeelde prefix voor alle bestanden.
   - **Exportmap**: doelmap voor de Pack & Go-uitvoer.
   - **Tekeningen meenemen**: vink aan indien je tekeningen wilt exporteren.
9. Klik **Uitvoeren** om het Pack & Go-proces met nieuwe namen te starten.

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
4. Controleer in **Project Explorer** of `PackAndRename` (module) en `PackAndRenameForm` (UserForm) zichtbaar zijn en verwijder eventueel lege standaarditems zoals `Module1`.
5. Kies **File → Save** om de macro in dezelfde `.swp` op te slaan. Dit `.swp`-bestand kun je daarna direct laden en uitvoeren via **Tools → Macro → Run...**.

**Belangrijke tip bij het importeren**
- Importeer de `.bas` en `.frm` alleen in het nieuwe macrobestand dat je net hebt aangemaakt. Verwijder eventueel automatisch toegevoegde items zoals `Module1` als die leeg is, zodat alleen `PackAndRename` (module) en `PackAndRenameForm` (UserForm) overblijven.
- Als je de bestanden per ongeluk meerdere keren importeert, kunnen er dubbele modules of formulieren ontstaan en zie je in SOLIDWORKS meerdere entries. Laat dan alleen één set staan (module + UserForm) en verwijder de duplicaten; sla daarna opnieuw op als `.swp`.

### Bouwen als .dll (optioneel)
Alleen nodig als je een COM-add-in wilt met extra UI-logica. Stappen in hoofdlijnen:
1. Open Visual Studio en maak een nieuw **Class Library**-project (vb.net of C#) met **.NET Framework** dat past bij jouw SOLIDWORKS-versie.
2. Voeg referenties toe aan de SOLIDWORKS interop-assemblies (`SolidWorks.Interop.sldworks`, `SolidWorks.Interop.swconst`, `SolidWorks.Interop.swpublished`).
3. Kopieer de logica uit `macros/PackAndRename.bas` naar een klas en schrijf een `ISwAddin`-implementatie die de functionaliteit aanroept.
4. Compileer de `.dll` en registreer deze via **Tools → Add-ins → Add...** of met `regasm.exe`.

## Foutopsporing bij laden in SOLIDWORKS
- **Geen code zichtbaar na importeren**: controleer of Project Explorer het juiste project toont en dubbelklik op `PackAndRename` of `PackAndRenameForm`. Staat de code er niet? Verwijder de lege standaardmodule (`Module1`), importeer de `.bas` en `.frm` opnieuw in hetzelfde project en sla op.
- **Krijg je meerdere macro-namen te zien bij het starten?** Open de VBA-editor, controleer of er dubbele modules/userforms staan, verwijder de duplicaten en sla opnieuw op.
- **Referentiefout bij het runnen?** Controleer in de VBA-editor onder **Tools → References** of alle standaard SOLIDWORKS-referenties actief zijn (`SolidWorks <versie> Type Library`, `SldWorks <versie> Constant Type Library`). Zet ontbrekende referenties aan, sla op en test opnieuw.
- **Macro doet niets**: zorg dat het actieve document een opgeslagen assembly is (geen part), dat de exportmap bestaat en schrijfrechten heeft en dat in de pop-up een prefix én exportmap zijn ingevuld.

## Aanpassen
- Pas eventueel de startnummers aan in `PackAndRename.bas` (variabelen `partCounter` en `asmCounter`).
- De helperfunctie `BuildNumber` bepaalt het nummerformaat (`00`).

## Vereisten
- SOLIDWORKS desktop met VBA-macro-ondersteuning.
- Toegang tot de doelmap (schrijfrechten).
