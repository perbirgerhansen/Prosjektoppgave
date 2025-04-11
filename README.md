
# Prosjekt oppgave i PY1010  2025 Per Hansen 
Jeg har utviklet et helpdesk-program med et grafisk brukergrensesnitt (GUI) ved hjelp av Tkinter-modulen i Python. Valget av Tkinter var bevisst for å skape en mer interaktiv brukeropplevelse sammenlignet med et konsollbasert program.

Jeg forsto det slik at oppgaven var fri,  jeg har likevel lagt inn noen av funksjonene dere har beskrevet i oppgaveteksten. 

## Programmet muliggjør registrering av henvendelser med følgende informasjon:
Kunde/Navn: Navn på kunden og kontaktperson.<br/>
Kontaktinformasjon: Adresse og telefonnummer.<br/>
Feilbeskrivelse: Kategori (type feil) og en detaljert beskrivelse.<br/>
Håndtering: Prioritet, lisenspris og tilfredshetsnivå.<br/>
Tidsregistrering: Automatisk registrering av starttidspunkt ved opprettelse og en knapp for å stoppe tiden. Fakturering beregnes ved å runde opp tidsbruken til nærmeste hele minutt (minimum 1 minutt) og multiplisere med en timesats på kr 6000,- (100 kr/minutt).

## Funksjonalitet for datainnsamling og lagring:
Alle felter er obligatoriske og endringer må lagres via en lagre-funksjon.
En Treeview-widget viser en oversikt over registrerte henvendelser med en integrert scrollbar. Ved å klikke på en rad i Treeview, populeres de tilhørende tekstboksene og andre input-felt med dataen fra den valgte henvendelsen.
Data lagres i en Excel-fil ved navn prosjektdata.xls, som forventes å ligge i mappen <ins>C:\prosjektoppgave sammen med programfilen prosjektoppgave.py.<ins>

## Utviklingsmiljø og plattformbegrensninger:
Programmet er utviklet og testet på Windows 10/11 platform i Spyder (versjon 5.5.1 og 6.0.5).
Det er viktig å merke seg at Tkinter ikke støttes i Jupyter Notebook eller Google Colab.

## Oppgaveløsning:
Jeg anser at programmet mitt svarer på oppgaven ved å implementere følgende konsepter og strukturer:<br/>
Datastrukturer: Bruk av lister og dictionaries for datamanipulering, som demonstrert i bla. funksjonen oppdater_tabell(file_path).<br/>
Visualisering: Generering av diagram som viser fordelingen av prioriteringsnivåer, tilfredshet, henvendelser, tidsbruk etc. som demonstrert i bla. funksjonen  priokake_graf(file_path).<br/>
Kontrollflyt: Implementering av betingede setninger (if/else), som vist i bla. funksjonen sisterekke().<br/>
Iterasjon: Bruk av løkker (for eller while) for å håndtere data, for eksempel ved oppdatering av tabellen i funksjonen oppdater_tabell(file_path).<br/>


