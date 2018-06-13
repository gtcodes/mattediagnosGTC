# Kort Beskrivning

Detta är ett project för att generera matte-recept för förstaårs deltagare på GTG.

# Körning av program

För att köra programmet så kalla på ./diagnos.py med en mapp som har namnet angivet i RESULT\_FOLDER med alla csv filer på resultat på alla klasser. Standard namnet är resultat. Då kommer det att genereas en map (med det namn som variabeln RECIPE\_FOLDER har) med en map för varje klass som innehåller alla recept för den klassen.

# Antaganden som görs på input fil.

Den tar som input en csv fil som innehåller följande fält:

Efternamn;Förnamn;negativa tal;Potenser;bråk;Procent;Algebra;Geometri;Ekvationer

Det är viktigt att fälten är i den ordningen eftersom recepten är specifika och just nu matchar vi inte mot kolumnen och väljer rätt recept. Istället antas filen följa ovanstående format.

Första cellen på första raden antas innehålla klassnamn
Andra raden antas innehålla fieldheaders
datan antas börja på tredje raden
Fyll i START\_DATA för att ange vilken kolumn datan börjar.


