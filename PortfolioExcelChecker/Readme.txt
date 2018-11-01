== Uso di PortfolioExcelChecker
Basta lanciarlo senza nessun parametro


== Info
Le quotazioni dei fondi ETF li prendo usando ruby. Le url che usa lo script ruby sono fisse dentro allo script.

Il file preposto 
si trova sotto D:\PC_Jim_2016\Projects\ruby\GitHub\ruby_scratch\finanz_net\get_quote.rb
Questo mi genera un file di testo separato da ; con le quote dei fondi.
Per esempio:
LU0292106167;XetraETF;18,39;24.02./17:36;-0,05/-0,27%

Il link dove si trova fisso in get_quote.rb.

== PortfolioExcelChecker

PortfolioExcelChecker lancia il processo in ruby che prende i valori attuali dei fondi e li scrive in quote.csv.
PortfolioExcelChecker riconosce che il processo di ruby è terminato e analizza quote.csv. Metti i valori
nel file excel Portfolio.xlsm che si trova su D:\Documents\easybank e lo apre.
Dentro al foglio excel esiste la possibilità di navigare direttamente al sito dei chart.
I path dove si trova rubx.exe, lo script get_quote.rb e il file excel sono fissi nel codice e non sono configurabili.
L'ordine dei tab in excel così come la posizione della prima riga con i dati sono fisse nel codice. (row 4)
PortfolioExcelChecker è un programma in c#. 
