# **Dashboard: Nieruchomości - Baza Danych**

### Narzędzie służy do administrowania danymi wynajmowanych nieruchomości: 
   ##### - zmiany danych najemców, nieruchomości, właścicieli, aktualizacji umów
   ##### - wystawiania i generowania bieżących faktur
   ##### - sporządzania rejestru wystawionych faktur
   ##### - generowania raportów dotyczących przychodu oraz kwoty należnego podatku
   ##### - automatyzacji wysyłania dokumentów mailowo w do najemców, którzy wolą otrzymywać faktury drogą elektroniczną.
___
#### Zawartość:  
   ##### - plik MainMenu.xlsm - plik główny służacy do obsługi całej aplikacji. Arkusz programu MS Excel (o rozszerzeniu .xlms) współpracujący z innymi aplikacjami pakietu MS Office za pomocą kodu VBA.
    
##### - folder add-inns: zawiera plik WinDatePicker.xlam - wtyczka DatePicker pochodząca ze strony autora: https://www.rondebruin.nl/win/addins/datepicker.htm

##### - folder BazaDanych: znajduje się w nim plik MainDB.accdb (MS Access), który zawiera przykładowe dane 20 nieruchomości, 20 najemców i jednego właściciela.
##### - folder templates: zawiera plik InvTempl01.docx - szablon faktury, służący do generowania dokumentów
##### - folder Faktury: służy do przechowywania wygenerowanych faktur
##### - folder Raporty: służy do przechowywania wygenerowanych raportów
___
#### Zanim zaczniesz korzystać - zainstaluj dodatki:
1) Dodaj WinDatePicker z folderu Add-Ins do swojego Excela. Pełna instrukcja instalacji - 
[Ron de Bruin Excel Automation](https://www.rondebruin.nl/win/addins/datepicker.htm)
2) Następnie w edytorze VBA (ALT + F11) w zakładce Tools -> References upewnij się, że zaznaczyłeś następujące opcje:
    - Visual Basic For Applications,
    - Microsoft Excel 16.0 Object Library,
    - OLE Automation,
    - Microsoft Office 16.0 Object Library,
    - Microsoft ActiveX Data Objects 6.1 Library,
    - Nicrosoft Outlook 16.0 Object Library,
    - Microsoft Forms 2.0 Object Librarym
    - Microsoft Word 16.0 Object Library
___
#### Zasada działania:
1) Wczytywanie danych i ew. korekta:
    - W pliku MainMenu.xlsm w arkuszu Menu główne za pomocą przycisku "Wczytaj dane" ładujemy aktualne dane najemców, nieruchomości, umów oraz właścicieli - niezbędne do wystawienia faktur. Tutaj możemy też zmienić ręcznie dowolne dane, tj. kwotę konkretnej faktury lub sposób płatności.
     - Jeśli w pliku znajdują się już dane, a chcemy mieć pewność, że są one aktualne - najpierw naciskamy "Wyczyść dane" a następnie "Wczytaj dane".
2) Wystawianie faktur:
    - Zaznaczamy osoby, którym chcemy wystawić fakturę i wybieramy datę wystawienia faktury
    - Klikamy przycisk "Wystaw faktury"
3) Rejestr faktur:
    - W arkuszu Rejestr faktur znajdują się wszystkie dane dotyczące wystawionych faktur
    - Faktury w postaci plików .pdf znajdują się w folderze Faktury, w podfolderach odpowiednich dla daty ich wystawienia.
4) Wysyłanie faktur pocztą elektroniczną
    - W arkuszu Rejestr faktur, jeśli w kolumnie Status znajduje się wpis  "Nie wysłano" - można ustawić opcję Zaznacz na TAK
    - Następnie wcisnąć przycisk "Wyślij faktury" - wszystkie zaznaczone faktury zostaną wysłane na adres zawarty w bazie danych.

    **UWAGA: Opcja ta wymaga skonfigurowania programu Outlook oraz uzupełnienia adresów email w arkuszu Opcje!**

5) Generowanie raportów:
    - W arkuszu Raporty znajdują się zestawione tabelarycznie dane przedstawiające ilość faktur wystawionych w danym miesiącu, kwotę przychodu z tego tytułu, sumę narastającą przychodu od początku danego roku, kwotę podatku do zapłacenia za dany miesiąc oraz aktualną stawkę podatkową
    - Po każdorazowym wystawieniu nowych faktur należy zaktualizować dane raportów za pomocą przycisku "Odśwież dane"
    - Następnie w polu Zaznacz należy ustawić TAK w miesiącu, w którym chcemy wystawić raport
    - Po zaznaczeniu odpowiednich pozycji klikamy przycisk "Generuj Raport"
    - Raporty w postaci plików .pdf są zapisywane w folderze Raporty.
___
