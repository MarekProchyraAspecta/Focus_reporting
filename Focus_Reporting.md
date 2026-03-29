# 📊 Focus Reporting – Knowledge Transfer Documentation

## 1. O aplikacii

Focus Reporting je Python batch aplikacia urcena na:
- zber usage dat z OCI (Oracle Cloud Infrastructure) pomocou standardnej Oracle FOCUS API
- transformaciu a agregaciu dat
- generovanie reportov (CSV)
- (historicky) distribuciu reportov emailom

### Architektura (logical)
- Entry point: `ShowUsage.py`
- Mail modul: `SendMail.py` (momentálne deaktivovaný)
- State management: `scripts/last_saved_date.txt`
- Logovanie: `ShowUsage.log`

### Typ spracovania
- batch processing (cron-driven)
- incremental processing (based on last_saved_date)

#### Prikaz na extrakciu dat
 - kompletny prikaz so vsetkymi parametrami: ( stiahne USAGE-ove data s hodinovou granularitou za volitelne obdobie a odosle vysledne CSV emailom )
    - "python3 /home/opc/python_projects/focus_reporting/venv/scripts/ShowUsage.py -qrt USAGE -freq ADHOC -g HOURLY -ds 2026-02-15 -de 2026-03-01 -csv"

#### Vstupne parametre
 -qrt ( typ extrahovanych dat ) = USAGE ( default ) | COST 
 -g ( granularita ) = HOURLY ( default ) | DAILY | MONTHLY
 
 -ds ( date start ) = datum zaciatku extrakcie ( formatovany na 00:00 )

 -de ( date end ) = datum konca extrakcie ( formatovany na 00:00 ) - realne sa stiahnu data az po "de" -1 00:00

 -freq REGULAR_WEEKLY: kazdy pondelok v akt. mesiaci sa stiahnu data od zac. akt. mesiaca po nedelu 00:00 pred tymto pondelkom ( tzn. od 1. v mesiaci az po akt. pondelok -1 ) 

   - FROM: 1. den v akt. mesiaci  
   - TO: [( akt. den ( pondelok ) -1 ) 00:00 ] => TO = nedela 00:00

 -freq REGULAR_MONTHLY: data za poslednych 30 dni
 
  - FROM: ( den spustenia app - 30 ) 00:00   
  - TO: ( den spustenia app - 1 ) 00:00
  Poznamka: 
 
 -freq ADHOC: data za volitelne obdobie ( max 30 dnova historia )
  - FROM: -ds 00:00  
  - TO: ( -de -1 ) 00:00

-csv: posli extrahovane a finalizovane data vo formate CSV mailom


#### Aktualne Cron Jobs:
    1) logika: kazdy 1. den v mesiaci o 06:00 stiahni data na hodinovej baze z SP OCI za predosly mesiac ( 30 dnova historia )

    # every first of month at 06:00 AM extract from [socialnapoistovna] tenant usage data for previous month -- DISABLED: and send CSV file
    0 6 1 * * /usr/bin/python3 /home/opc/python_projects/focus_reporting/venv/scripts/ShowUsage.py -g HOURLY -freq "REGULAR_MONTHLY" -csv >> /home/opc/python_projects/focus_reporting/venv/scripts/ShowUsage.log 2>&1; echo "Exit status: $?" >> /home/opc/python_projects/focus_reporting/venv/scripts/ShowUsage.log

    2) logika: kazdy pondelok v akt. mesiaci stiahni data na hodinovej baze z SP OCI za akt. mesiac az po den-1 00:00 volania app       

    # every monday extract from [socialnapoistovna] tenant usage data for current month
    0 0 * * 1 /usr/bin/python3 /home/opc/python_projects/focus_reporting/venv/scripts/ShowUsage.py -g HOURLY -freq "REGULAR_WEEKLY" -csv >> /home/opc/python_projects/focus_reporting/venv/scripts/ShowUsage.log 2>&1; echo "Exit status: $?" >> /home/opc/python_projects/focus_reporting/venv/scripts/ShowUsage.log


## 2. Connectivity & Security

### 2.1 Runtime prostredie

- OS: Linux (remote server cez SSH)
- Runtime: Python 3.x
- Virtual env: `venv/`
- Spúšťanie: cron



### 2.2 OCI Connectivity (FOCUS API)

Aplikácia používa OCI SDK + Focus API na získanie usage dát.

#### Autentifikácia:
- typ: API key-based authentication
- config:
```
~/.oci/config
```

#### Obsah configu:
- tenancy OCID
- user OCID
- fingerprint
- private key path

---

### 2.3 Tenancy model

#### 1️⃣ execution tenant = Aspecta cloud ( tenant: aspectask )
- tenant, pod ktorým beží aplikácia
- obsahuje:
  - OCI credentials 
  - compute instance (server)
  - config súbory

Poznamka: 
  - v sucasnej verzii sa tento tenant nepouziva, pre spravnu funkciu app nie je nutny
  - tenant bol pripraveny na mozne ukladanie CSV do Aspecta cloudu

#### 2️⃣ source tenant (data source) = SP cloud ( tenant: SocialnaPoistovna )
- tenant, z ktorého sa čítajú usage dáta
- prístup:
  - cez IAM policies
  - cross-tenant access

---

### 2.4 Security riziká

| Oblasť | Riziko | Odporúčanie |
|------|-------|------------|
| ~/.oci/config | obsahuje private key path | necommitovať |
| private keys | citlivé | secure storage |
| email creds | hardcoded | presun do .env |

---

## 3. Focus API špecifiká

### Granularita
  #### Popis logiky: 
    - kazdy mesiac ( 2. v akt. mesiaci ) potrebujeme reportovat do SP data za SP spotrebu za cely predosly mesiac s hodinovou granularitou 
    - tuto granularitu nam ostatne dostatne typy agregacii ( DAILY, MONTHLY ) neposkytuju, preto je default nastavena HOURLY granularita

### Limity standardneho Oracle FOCUS API

#### max 8 hodin v jednom volani:
  - FOCUS API vracia max. 8 hodin v jednom volani
  - dosledok: aplikacia vola API v niekolkych cykloch => sklada cely den z viacerych batchov

---

#### max 30-dňové okno:
  - FOCUS API vracia data max za poslednych 30 dni (rolling window)
  - dosledok: aplikaciu je potrebne spustat min. 2x mesacne ( aby sme si odlozili data za prechody medzi mesiacmi - pre pripad, ze predosly reportovany mesiac ma 31 dni )

---

#### oneskorenie dát v OCI:
  - data o infrastrukture v OCI sa objavuju oneskorene
  - typicky:
    - nekompletne v 1. den v mesiaci
    - niekedy nekompletne ani v 2. den rano

---

Poznamka:
  - z dovodov obmedzeni na strane Oracle FOCUS API ( vyssie ) je nutne vysledky manuálne validovat a archivovat


### Odporúčané min nastavenie extrakcie dat

| Termin            | Parameter -freq | Rozsah extrahovanych dat | 
| ----------------- | --------------- | ------------------------ |
|  2. den v mesiaci | REGULAR_MONTHLY | data za poslednych 30 dni ( spolu s datami zbieranymi a zalovanymi WEEKLY = kompletna historia za predosly mesiac )                        |
| 15. den v mesiaci | REGULAR_WEEKLY  | data za akt. mesiac [ od 1. v mesiaci po 14. v mesiaci 00:00 ] ( realne sa ale nestiahnu vsetky data za 14. den => pozri oneskorenie dat ) |

---

### Business constraint

- Alex S. potrebuje mat report vygenerovany max. do 2. dna v mesiaci (SP reporting)
- potencialny konflikt:
  - kvoli obmedzeniam na strane Oracle FOCUS API ( pozri vyssie ) dáta nemusia byť do tohto datumu kompletne
  - je nutna manualna intervencia ( stiahnutie dat za volitelne obdobie )
    - => volat manualne python script s parametrom -freq = ADHOC a parametrami -ds (start) a -de (end)


---

## 4. Process Flow

### High-level flow

| Krok | Popis kroku | nazov funkcie |
| ---- |------------ | ------------- |
| 0 | Hlavna funkcia ( start app ) | fn_Main() |
| 1 | Validacia vstupnych parametrov | fn_ProcessInputParams() |
| 2 | Ziskanie FOCUS API usage klienta | fn_GetUsageClient() |
| 3 | Extrakcia dat | fn_GetData_Main () |
| 4 | Formatovanie dat, generovanie CSV | fn_PrintSaveData() |
| 5 | Save last extracted date | priamo v fn_Main() |
| 6 | Send CSV via email ( akt. disabled) | fn_SendMail_SendGridApp() |
| 7 | Save CSV to GitHub | fn_CommitCSVToGitHub() |

- Poznamka: generovane CSV je ukladane na server: python_projects/focus_reporting/save a zaroven aj na GiHub

---

## 5. Email funkcionalita

### Stav
- momentalne pozastavena
- dovody:
    1. nekonzistentne data - pozri vyssie ( obmedzenia Oracle FOCUS API )
    2. potreba manualnej revizie
    3. expirovany SendGrid subscription

---

## 6. Deployment

```
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## 7. Known Issues

- nekonzistentne data z OCI
- zavislost na manualnom spracovani

---

## 8. Odporucane zlepsenia

- retry mechanizmus pre API
- data completeness check
- automaticka archivacia CSV

---

## 9. Zaver

Aplikacia je funkcny reporting nastroj, ale:
- silne zavisi od kvality OCI dat
- vyzaduje manualny dohlad
- potrebuje stabilizaciu pre production-grade pouzitie

---

📅 Generated: 2026-03-29 19:36:23
