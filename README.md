# MidiTimecode
VB.NET Module to send and receive MIDI Timecode as per definition in Wikipedia https://en.wikipedia.org/wiki/MIDI_timecode

The file "MidiFunctions.vb" contains everything needed to send and receive timecode. There are no external dependencies reqired.



VB.NET Modul zum senden und empfangen con MIDI Timecode

Die Datei "MidiFunctions.vb" enthält alles was zum senden und empfangen von timecode erforderlich ist. Es sind keine weiteren Abhängigkeiten vorhanden.

Der Sender und Empfänger hält sich exact an die MIDI Spezifikationen und ermöglicht extrem hohe genauigkeit von ca +- 2 ms.
Die Mitgelieferte Grafische Oberfläche ist eine vereinfachte Darstellung um die MIDI Funktionen direkt testen zu können.
Zum senden von MIDI signalen über Netzwerk empfehlen wir RTP Midi von Tobias Erichsen: https://www.tobias-erichsen.de/software/rtpmidi.html. In kombination mit unserem TimeCode Sender ist so die Synchronisation von MIDI Anwendungen auf PCs auch ohne MIDI-Hardware möglich.
