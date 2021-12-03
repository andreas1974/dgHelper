# dgHelper
Beskrivning
===========
DgHelper är ett AutoHotKey-skript som underlättar vid inmatning i det svenska släktforskningsprogrammet Disgen 8.2. Med ett tangentbords-snabbkommando (ctrl+k) delas en Källhänvisning, kopierad från ArkivDigital, upp i sina beståndsdelar såsom sidnummer, AID, eventuell egen citattext med mera, och kopieras till respektive fält i Disgens fönster för Källa eller Hänvisning. Skriptet innehåller även en rutin för uppdelning och inmatning av kopierade koordinater för ort.

Skriptet är utvecklat för Disgen 8.2 men synes fungera även i 2016, då de flesta av inmatningsrutorna för källor och orter är snarlika och heter samma sak i de två versionerna av programmet. Mina få tester (i utgånget och begränsat demo-läge) visar på att skriptet även fungerar i Disgen 2016. <br />
Källorna kan kopieras antingen från Arkiv Digitals klientprogram (gamla "ArkivDigital Online") eller från deras nyare webbversion http://app.arkivdigital.se/.

Stumfilm som visar hur skriptet fungerar: https://www.facebook.com/andreas.jansson.5817/videos/1420301858055503

Förberedelser
=============
För att kunna använda skriptet behöver du installera AutoHotKey på din dator. Välj "Download AutoHotkey Installer" och installera (ver. 1.1.x) från  https://autohotkey.com/. Läs gärna på om AutoHotKey på externa webbplatser, för att bilda er en egen uppfattning om programmet som "kör" automatiseringsskriptet!

Ladda ner DgHelper.ahk från fillistan här på GitHub [eller den gröna knappen], starta sedan skriptet genom att dubbelklicka på det.
Lägg skriptet eller en genväg till det på ditt skrivbord för att ha snabb tillgång till det, eller schemalägg det så att det startar tillsammans med datorn om du alltid vill ha det igång. Att skriptet är igång märks inte på annat vis än att en grön fyrkantig ikon med ett vitt H visar sig i Windows Taskbar (längst nere till höger); det är vad som krävs för att snabbkommandot Ctrl+K nu ska utföra sitt jobb i Disgens fönster för källor, hänvisningar och ort.

(Det finns även möjlighet att kompilera dgHelper till en fristående exe-fil, men jag avhåller mig från att distribuera en sådan, då jag vill att ni som använder skriptet ska kunna öppna källkoden och undersöka den. En okänd exe-fil kan göra vad som helst!)

Användande av dgHelper.ahk
==========================
Ny källa
-------------
1. Kopiera källhänvisningen från ArkivDigital (välj "kopiera källhänvisning" i AD:s meny)
2. Skapa en ny (tom) källa i Disgen 8.2.
3. Tryck Ctrl+k (d.v.s. håll nere knappen "Ctrl" och tryck på tangenten "k". Släpp sedan upp båda.)
Resultat: Kort titel läggas in (länsbokstaven exkluderas, se specialinställningar nedan) och källans årtal läggs in i fältet Publicering.

Ny hänvisning
-------------
1. Kopiera källhänvisningen från ArkivDigital. (Detta moment hoppar du över om du redan har kopierat källhänvisning och skapat källa enligt ovan – den kopierade hänvisningen ligger kvar!)
2. Skapa en ny (tom) hänvisning i Disgen 8.2.
3. Tryck Ctrl+k
Resultat: Prefix p (pagina) väljs om sidnummer finns. Sidnummer läggs in i fältet hänvisningstext. Om sidnummer saknas skrivs istället "AD: xx" som hänvisningstext, där xx är Arkiv Digitals bildnummer (detta kan enkelt anpassas i skriptet). Om datum finns på rad 2 i urklippet (din kopierade källhänvisning/text) läggs detta in som datum och källans kvalitet sätts då även till "primär". Koppling till Arkiv Digital läggs in. Citat från källan läggs in om det finns med i din kopierade källhänvisning/text (rad 2 om ej datum samt därpå följande rader). Som Anteckning läggs hänvisningen sådan som den kopierades från Arkiv Digital, d.v.s. inklusive eventuell länsbokstav.

Utökat användningsområde vid transkribering (hänvisning inkl citat)
-------------------------------------------------------------------
Vid transkribering av text från Arkiv Digital, till källor, kan ett enkelt program såsom Anteckningar / Notepad rekommenderas som "mellanprogram", i synnerhet om du inte har en mycket bred bildskärm (eller två stycken). Öppna "Anteckningar" och förminska dess fönster det så att det tar upp en lämplig del av skärmytan, t.ex. nedre tredjedelen.
Kopiera Arkiv Digitals källa och lägg på rad 1 i Anteckningar.
Om det är en födelse eller vigselnotis som skrivs av finns det ofta ett exakt datum. Skriv in datumet på rad 2. Formatet ska vara ÅÅÅÅMMDD (inga bindestreck eller andra tecken, endast 8 siffror, nollor om dag eller månad saknas).
Skriv på följande rader av texten från Arkiv Digital. Detta blir citat-texten som hamnar som "Citat" i hänvisningen när du senare trycker Ctrl+k enligt beskrivningen "Ny hänvisning" ovan. Om exakt datum saknas lägger du citatet med början på rad 2.

Istället för att bara kopiera källhänvisningen direkt från Arkiv Digital kan texten alltså kopieras från ett valfritt program. Markera hela texten i Anteckningar och tryck Ctrl+C (för Kopiera / Copy). Tryck sedan Ctrl+k för att ta in hänvisning inkl citat och ev. datum i öppen hänvisningsruta i Disgen, enligt ovan.

Exempel på vad som kan kopieras och tas omhand av dgHelper
----------------------------------------------------------
> Ås (P) C:5 (1801-1841) Bild 1 / sid 4 (AID: v706.b1.s4, NAD: SE/GLA/130)<br />
> 18381221<br />
> [Död] 2. [Begr] 9. Enkl. Jonas Jonasson på KlippeSvedjorna under Lidagärde. Ålderdom. 78 [år] 8 m. n.d.<br />
> [Vid sin död boende med sonen Per Jonsson och dennes familj, på torpet Klippesvedjorna]<br />

*Rad 1 används som källa, rad 2 läggs in som datum (om raden består av exakt 8 siffror), rad 3 och därpå följande rader hamnar som citattext i Disgens källhänvisning.*

Koordinater (ort)
=================
Skriptet kan även användas för att underlätta inmatning av koordinater för en ny ort, t.ex. en gård, i DisGen. Kopiera koordinater från t.ex. "hitta.se", där X- och Y-koordinater de ligger med kommatecken mellan sig och tryck Ctrl+k för att ta in dem i ortens resp. fält.
Anledningen till denna funktion är att undertecknad många gånger råkat ut för en bugg i ortsträdet för Disgen 8.2, där växling av program med Ctrl+TAB, medan ortsredigeringen är öppen, får orter att byta information med varandra m.fl. liknande effekter. Se beskrivning av en av dessa buggar i meddelande #15 i denna diskussionstråd: https://forum.dis.se/vb/showthread.php/1231-Ikoner-och-orter-byts-ut-i-ortstr%E4det-(g%E5rd-blir-socken-socknen-borta)?p=6730&viewfull=15#post6730
Att kopiera koordinater och trycka ctrl+k enligt beskrivning nedan gör att man slipper växla fram och tillbaka mellan DisGen och webbläsaren med kartan, för att kopiera först X-koordinaten och sedan Y-koordinaten. Detaljerade instruktioner:
1. Gå in på kartan på hitta.se och leta upp den plats du är intresserad av. 
2. Klicka på den lilla "hamburgermenyn" (tre horisontella streck) och välj "Koordinater".
3. Justera eventuellt siktet, så att det ligger över orten vars koordinater du vill använda.
4. Markera koordinaterna som står under rubriken RT90 till höger om kartan på hitta.se. Det är detta koordinatsystem som Disgen använder sig av. Tryck Ctrl+C för att kopiera koordinaterna.
5. Öppna "Redigera ort" genom att välja "Ny nivå..." eller "Ändra..." på en befintlig ort i ortsträdet.
6. Tryck Ctrl+k
Resultat: Koordinaterna från den kopierade texten delas upp och placeras i X resp. Y-fältet för orten. Eventuella decimaler tas bort. Om du kopierade t.ex. "6404444.522,1339731.124" från Hitta kommer 6404444 att placeras i X-fältet och 1339731 i Y-fältet.

Specialinställningar
====================
På grund av att Arkiv Digital i sin webbversion har lagt till länsbokstaven som del av källhänvisningen, ställer det till det i sorteringsordningen i källträdet – de med länsbokstav hamnar före de gamla. Exempelvis hamnar ...<br />
Od (P) AI:5<br />
... FÖRE de tidigare inmatade, äldre, böckerna ...<br />
Od AI:1<br />
Od AI:3<br />
Od AI:4<br />
Skriptet tar i sitt originalutförande bort länsbokstaven och dess omslutande parenteser, för att skapa "Kort titel" för källan. Om du inte har många gamla källor och dessutom endast arbetar med den webb-baserade versionen av ArkivDigital (den som körs i webbläsare t.ex. Chrome), så att du ALLTID får med länsbokstaven i de källor du kopierar (den kommer inte med från den äldre versionen av Arkiv Digital Online – den som installeras och körs lokalt på din dator) kan det möjligen ligga en poäng i att behålla länsbokstaven i källan(?). Du kan enkelt stänga av avlägsnandet av Länsbokstav genom att öppna dhHelper.ahk i Anteckningar och ändra raden ...<br />
RemoveLänsbokstav := true<br />
... till ... <br />
RemoveLänsbokstav := false<br />

Ctrl+k aktiverar inmatningsskriptet för källor och källhänvisningar även från ArkivDigital och Notepad samt Notepad++. Det räcker alltså att kopiera källhänvisningen och trycka ctrl+k i ArkivDigital för att skriptet ska växla över till Disgen och  kopiera in uppgifterna till den ruta som är aktiv där ("Egenskaper för källhänvisning" eller "Egenskaper för källa"). Observera dock att rutan för källa el. källhänvisning måste vara öppen i Disgen för att något ska hända när du trycker ctrl+k i de övriga programmen.

Kända buggar
============
Datumet (från rad 2 i kopierad källtext, om sådan finns) kopieras ibland in med en förskjutning om ett tecken, till datumrutan i Disgens hänvisning. Kontrollera alltså alltid att datum hamnar rätt – datumet blir emellertid så felaktigt att Disgen inte accepterar det, varför ingen nämnvärd risk föreligger att det ska sparas något felaktigt.


Upphov / kontakt
================
Andreas Jansson<br />
Ön Tärby den 23/7 2017<br />

Det är nog lättast att få kontakt med mig genom att ni letar upp mig på Facebook (det finns inlägg om DGHelper i gruppen "Vi som gillar Disgen". Jag misstänker att jag även får en avisering om ni skriver något under Pull Requests här på GitHub. Jag har inte lyckats logga in på Dis Forum på ett tag – jag hade fått en fråga där om skriptet där beträffande Disgen 2021 tror jag, men inloggningen krånglade som sagt, och jag har inte fixat ny än (dec 2021).
