# Post-Maintenance Checklist for Exchange Environments

**Purpose:** After every maintenance window, CU update, security update, hotfix, or third-party update, actively verify that the environment is truly fully restored, not just that the service status reads "Running" again.

## Maintenance Record

| Field | Entry |
|---|---|
| Update type | ☐ CU &nbsp;☐ SU &nbsp;☐ HU &nbsp;☐ Third-party &nbsp;☐ Other |
| Version / KB number | ______________ |
| Affected servers / DAG | ______________ |
| Date performed | ______________ |
| Performed by | ______________ |

---

## 1. Transport and mail flow

- [ ] Sent a test message through every internal transport path and confirmed delivery
- [ ] Sent a test message from on-premises to Exchange Online and checked delivery, including transit time
- [ ] Sent a test message from Exchange Online to on-premises and checked delivery, including transit time (test the return path separately, not just the outbound path)
- [ ] Sent a test message to the internet (outbound) and confirmed delivery
- [ ] Checked transport queues on all servers, especially for messages stuck in retry or unusually long queue lengths
- [ ] Confirmed no messages are stuck in the poison queue

## 2. Protocols and authentication

- [ ] Performed a test login for every protocol used in the environment (MAPI over HTTP, EWS, Autodiscover, OWA, ActiveSync, POP3/IMAP4 if applicable)
- [ ] Verified full authentication for each test, not just endpoint reachability
- [ ] For OWA, additionally checked: post-login redirect, mailbox view load time, calendar access

## 3. DAG and databases

- [ ] Checked copy status of all database copies across all DAG members (healthy, mounted as expected, no unusual copy queue length)
- [ ] Verified that after a failover, the active copy sits on the expected or at least a suitable server
- [ ] Compared resource usage (CPU, RAM, disk) on the server that took over a database against the baseline before the update
- [ ] Checked DAG quorum status

## 4. Hybrid and certificates

- [ ] Checked hybrid connector status (active, authenticating correctly)
- [ ] Verified validity and correct binding of all relevant TLS certificates, especially if a renewal occurred during the maintenance window
- [ ] Tested Free/Busy information between on-premises and Exchange Online
- [ ] Checked organization relationship and Autodiscover responses for hybrid mailboxes

## 5. Monitoring itself

- [ ] Actively disabled alert suppression / maintenance mode in the monitoring system, rather than relying on automatic timeout
- [ ] Verified that all monitoring agents and sensors are reporting correctly again after the update
- [ ] Watched trend values, not just point-in-time snapshots, in the first few hours after the update

## 6. Sign-off

- [ ] Documented the checklist outcome (date, tests performed, anything unusual)
- [ ] Only marked the maintenance window as officially closed once the outcome was fully positive

---

*This checklist is a template and should be adapted to your individual environment, particularly with regard to the protocols in use, the number of DAG members, and your hybrid configuration.*
