# Post-Maintenance Checkliste für Exchange-Umgebungen

**Zweck:** Nach jedem Wartungsfenster, CU-Update, Security Update, Hotfix oder Drittanbieter-Update aktiv prüfen, ob die Umgebung wirklich vollständig wiederhergestellt ist, nicht nur, ob der Dienststatus wieder "Running" zeigt.

## Wartungsprotokoll

| Feld | Eintrag |
|---|---|
| Update-Typ | ☐ CU &nbsp;☐ SU &nbsp;☐ HU &nbsp;☐ Third-Party &nbsp;☐ Sonstiges |
| Version / KB-Nummer | ______________ |
| Betroffene Server / DAG | ______________ |
| Datum der Durchführung | ______________ |
| Durchgeführt von | ______________ |

---

## 1. Transport und Mailflow

- [ ] Testnachricht durch jede interne Transportstrecke geschickt und Zustellung bestätigt
- [ ] Testnachricht von On-Premises nach Exchange Online geschickt und Zustellung inkl. Laufzeit geprüft
- [ ] Testnachricht von Exchange Online nach On-Premises geschickt und Zustellung inkl. Laufzeit geprüft (Rückweg separat testen, nicht nur den Hinweg)
- [ ] Testnachricht ins Internet (ausgehend) geschickt und Zustellung bestätigt
- [ ] Transport Queues auf allen Servern geprüft, insbesondere auf Nachrichten im Retry-Status oder ungewöhnlich hohe Queue-Längen
- [ ] Sichergestellt, dass keine Nachrichten im Poison-Queue-Status hängen

## 2. Protokolle und Authentifizierung

- [ ] Testlogin für jedes in der Umgebung genutzte Protokoll durchgeführt (MAPI over HTTP, EWS, Autodiscover, OWA, ActiveSync, ggf. POP3/IMAP4)
- [ ] Bei jedem Test die vollständige Authentifizierung geprüft, nicht nur die Erreichbarkeit des Endpunkts
- [ ] Bei OWA zusätzlich geprüft: Weiterleitung nach Login, Ladezeit der Postfachansicht, Zugriff auf Kalender

## 3. DAG und Datenbanken

- [ ] Kopierstatus aller Datenbankkopien auf allen DAG-Mitgliedern geprüft (gesund, gemountet wie erwartet, keine ungewöhnliche Kopierverzögerung)
- [ ] Geprüft, ob nach einem Failover die aktive Kopie auf dem erwarteten oder zumindest einem geeigneten Server liegt
- [ ] Ressourcenauslastung (CPU, RAM, Disk) auf dem Server, der eine Datenbank übernommen hat, mit dem Normalzustand vor dem Update verglichen
- [ ] DAG-Quorum-Status geprüft

## 4. Hybrid und Zertifikate

- [ ] Status des Hybridconnectors geprüft (aktiv, korrekt authentifiziert)
- [ ] Gültigkeit und korrekte Bindung aller relevanten TLS-Zertifikate geprüft, insbesondere wenn während des Wartungsfensters eine Erneuerung stattgefunden hat
- [ ] Frei/Besetzt-Informationen (Free/Busy) zwischen On-Premises und Exchange Online getestet
- [ ] Organisationsbeziehung und Autodiscover-Antworten für Hybrid-Postfächer geprüft

## 5. Monitoring selbst

- [ ] Alert-Unterdrückung/Wartungsmodus im Monitoring-System aktiv deaktiviert, nicht auf automatisches Timeout verlassen
- [ ] Geprüft, ob alle Monitoring-Agenten und Sensoren nach dem Update wieder korrekt melden
- [ ] Kurzfristig (in den ersten Stunden nach dem Update) Trendwerte statt nur Momentaufnahmen beobachtet

## 6. Abschluss

- [ ] Ergebnis der Checkliste dokumentiert (Datum, durchgeführte Tests, Auffälligkeiten)
- [ ] Wartungsfenster erst nach vollständig positivem Ergebnis offiziell als abgeschlossen markiert

---

*Diese Checkliste ist eine Vorlage und sollte an die individuelle Umgebung angepasst werden, insbesondere hinsichtlich genutzter Protokolle, Anzahl der DAG-Mitglieder und Hybrid-Konfiguration.*
