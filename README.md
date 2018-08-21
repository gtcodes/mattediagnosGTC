# Kort Beskrivning

Detta är ett project för att generera matte-recept för förstaårs deltagare på GTG.

# Körning av program

För att köra programmet, kalla på ./diagnos.py med en mapp som har namnet angivet i RESULT\_FOLDER med alla csv filer med resultat från alla klasser. Standard namnet är resultat. Då kommer det att genereras en folder (med det namn som variabeln RECIPE\_FOLDER har) med en folder för varje klass som innehåller alla recept för den klassen.

# Antaganden som görs på input fil.

Den tar som input en csv fil som innehåller följande fält:

namn;negativa tal;Potenser;bråk;Procent;Algebra;Geometri;Ekvationer

Det är viktigt att fälten är i den ordningen eftersom recepten är specifika och just nu matchar vi inte mot kolumnen och väljer rätt recept. Istället antas filen följa ovanstående format.

Första cellen på första raden antas innehålla klassnamn. Dem andra columnerna på första raden ignoreras. Datan antas börja på andra raden.

Om namnen har formatet lastname, firstname kan man ändra STANDARD\_NAME\_FORMAT till 1 så görs det om till att bli firstname lastname.

för att behålla å,ä,ö i csv bör man spara som CSV (Kommaavgränsad) i excell.
