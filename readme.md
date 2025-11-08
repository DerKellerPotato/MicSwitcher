# MicSwitcher 🎙️

MicSwitcher ist eine kleine Windows-System-Tray-Anwendung, die dein Standard-Mikrofon automatisch umschaltet, basierend darauf, ob dein Headset-Mikrofon stumm geschaltet ist oder nicht.

## 🎯 Der Anwendungsfall

Du sitzt mit Freunden im Discord und benutzt dein hochwertiges Standmikrofon. Dafür ist dein Headset-Mikrofon hochgeklappt (stumm geschaltet).
Jetzt stehst du auf, um dir etwas zu holen, möchtest aber weiterreden. Du klappst dein Headset-Mikrofon herunter (entmutest es).

**MicSwitcher erkennt das Entmuten und schaltet das Standard-Mikrofon von Windows automatisch auf dein Headset um.**

Wenn du dich wieder hinsetzt und das Headset-Mikrofon hochklappst (stummschaltest), schaltet die App sofort wieder auf dein Standmikrofon zurück.

![Screenshot des Einstellungsfensters](https://i.imgur.com/DEIN-SCREENSHOT-HIER.png)
*(Tipp: Mache einen Screenshot vom Einstellungsfenster, lade ihn z.B. bei [Imgur](https://imgur.com/upload) hoch und füge den Link hier ein)*

## 🚀 Hauptfunktionen

* **System-Tray:** Läuft unauffällig im System-Tray (neben der Uhr).
* **Mute-Erkennung:** Hört auf Mute/Unmute-Ereignisse von ausgewählten Headsets (basierend auf der Windows-Audio-API via `NAudio`).
* **Auto-Switch:** Automatisches Umschalten zwischen einem primären "Stand-Mikrofon" (wenn das Headset stumm ist) und dem "Headset-Mikrofon" (wenn es aktiv ist).
* **Benachrichtigungen:** Zeigt eine Windows-Benachrichtigung an, sobald das Mikrofon umgeschaltet wurde.
* **GUI:** Einfaches Einstellungsfenster zur Auswahl der Geräte.
* **Autostart:** Kann im Einstellungsfenster so konfiguriert werden, dass es mit Windows startet.
* **Live-Update:** Einstellungen werden beim Speichern sofort "live" übernommen, ohne dass die App neu gestartet werden muss.

## ⚠️ Wichtiger Hinweis für Discord, Teams, Gamebar etc.

Viele Kommunikations-Apps (wie Discord, Teams, OBS oder die Xbox Gamebar) "merken" sich beim Start, welches Mikrofon sie verwenden. Sie folgen nicht immer automatisch dem Windows-Standardgerät, *nachdem* sie gestartet wurden.

**Die Lösung:** Stelle sicher, dass du in den Audio-Einstellungen deiner App (z.B. Discord) als **Eingabegerät** nicht dein Headset oder dein Standmikrofon fest auswählst, sondern die Option namens "**Standard**" (oder "Default"). Nur dann folgt die App den Änderungen, die MicSwitcher vornimmt.



## ⚙️ Verwendung (Für Anwender)

1.  Lade die `MicSwitcher.exe` aus dem [Releases-Tab](https://github.com/DEIN-USERNAME/MicSwitcher/releases) herunter. *(Du musst zuerst eine Release erstellen, damit dieser Link funktioniert)*
2.  Starte die `MicSwitcher.exe`. Beim ersten Start öffnet sich automatisch das Einstellungsfenster.
3.  Wähle im Dropdown-Menü dein **Stand-Mikrofon** aus (das Gerät, das aktiv sein soll, wenn das Headset stumm ist).
4.  Wähle in der Checkbox-Liste das **Headset-Mikrofon** (oder mehrere), das die Umschaltung auslösen soll.
5.  Setze den Haken bei "Mit Windows starten", wenn gewünscht.
6.  Klicke auf "Speichern & Schließen".
7.  Die App läuft nun im Hintergrund. Teste es, indem du dein Headset mutest/entmutest!

## 💻 Kompilieren (Für Entwickler)

Dieses Projekt wurde mit Visual Studio 2022 und .NET 8 erstellt.

1.  Klone dieses Repository: `git clone https://github.com/DEIN-USERNAME/MicSwitcher.git`
2.  Öffne die `MicSwitcher.sln`-Datei in Visual Studio 2022.
3.  Stelle sicher, dass die ".NET-Desktopentwicklung"-Arbeitslast in Visual Studio installiert ist.
4.  Klicke mit der rechten Maustaste auf das Projekt und wähle "Manage NuGet Packages...".
5.  Stelle sicher, dass das `NAudio`-Paket installiert/wiederhergestellt ist.
6.  Stelle die Konfiguration oben von `Debug` auf `Release`.
7.  Klicke im Menü auf **Erstellen** > **Projektmappe erstellen** (oder F6).

### Als einzelne EXE-Datei veröffentlichen

Um eine **einzelne, portable `.exe`-Datei** zu erstellen (empfohlen):
1.  Rechtsklick auf das `MicSwitcher`-Projekt im Solution Explorer.
2.  Wähle **Veröffentlichen...** (Publish...).
3.  Wähle "Ordner" als Ziel.
4.  Klicke bei den Profileinstellungen auf "Bearbeiten":
    * **Bereitstellungsmodus:** `Eigenständig` (Self-contained)
    * **Datei-Veröffentlichungsoptionen:** Haken bei `In einer Einzeldatei veröffentlichen` (Produce single file)
5.  Klicke "Speichern" und dann "Veröffentlichen". Die fertige `.exe` liegt im Ordner `bin/Release/net8.0-windows/publish`.

## 🛠️ Verwendete Technologien

* **C# (.NET 8)**
* **Windows Forms** (für das Einstellungs-GUI und das Tray-Icon)
* **NAudio** (Zum Abhören von Windows-Audio-Ereignissen wie Mute/Volume)
* **Windows Core Audio (COM-Schnittstellen)** (Zum programmatischen Umschalten des Standard-Audiogeräts)