# ImapFatturaElettronica

Scaricare le fatture Elettroniche dalla posta certificata con la libreria MailKit (https://www.nuget.org/packages/MailKit/).

Le prove le ho fatte con una pec del mio dominio presso Aruba.

prima di provare il codice installare la libreria.

Valorizzare le variabili:

-server IMAP

-porta IMAP

-mail

-password

ci sono due variabili:

ID_UltimaFattura e DataUltimoControlloFatture che servono per identificare l'ultima fattura scaricata per non scaricarle ogni volta tutte.
Questo è un suggerimento dall'amico Lucio. 

Per usarli, va salvato il valore nei settings dell'applicazione ogni volta che si scarica la mail (nuovo ID e nuova data/ora).

Quindi il valore di default non deve esistere tranne nel caso della prima volta che viene scaricata la posta.
