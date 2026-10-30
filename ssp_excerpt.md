# System Security Plan Excerpt: Finance and Payroll ERP (FPE)

*Sample document for portfolio use. The system, organization and people are fictional.*

## 1. System identification

| Item | Value |
|---|---|
| System name | Finance and Payroll ERP (FPE) |
| System type | Major application |
| Operational status | Operational |
| System owner | Finance Controller |
| Information system security officer | Manager, IT Security |
| Authorizing official | Chief Financial Officer |
| Hosting | Organization-managed data center with a separately managed payroll service |

## 2. System description and purpose

FPE supports four business processes: accounts payable, general ledger, procurement and payroll. Finance staff use it to record vendor invoices, approve and issue payments, post journal entries and close each month. The payroll module calculates pay and produces direct deposit files. About 45 finance, procurement and IT users hold accounts, and a small number of administrators manage the application, database and payroll configuration.

## 3. Authorization boundary

Inside the boundary: the ERP application server, the ERP database server, the payroll application and its database, and the user access and change management records for these components. Outside the boundary: the corporate network, the identity provider, the HR system, the ticketing system and the VPN. Inherited controls from those services are noted in the SSP statements where relevant.

## 4. Security categorization (FIPS 199)

| Objective | Impact | Rationale |
|---|---|---|
| Confidentiality | Moderate | Payroll records contain personal and bank account data, and financial data is non-public. Disclosure would cause serious but not catastrophic harm. |
| Integrity | Moderate | Altered payments, vendor records or payroll could cause financial loss and misstated reports. Errors are detectable through reconciliation but could be costly. |
| Availability | Moderate | A multi-day outage would delay vendor payments, payroll and the monthly close. Manual workarounds exist for short outages. |

Overall system impact level: **Moderate**, the highest of the three objectives. The Moderate baseline in NIST SP 800-53 Rev. 5 applies.

## 5. Organization-defined parameters used in this sample

| Control | Parameter |
|---|---|
| AC-2, PS-4 | Disable accounts of terminated personnel within 24 hours; review access every six months |
| CM-3 | Independent approval before implementation of every production change |
| CP-4 | Test the contingency plan, including restoration, every 12 months |
| RA-5 | Scan monthly with authenticated scans |
| SI-2 | Remediate critical flaws within 30 days |
| AU-6 | Review audit records weekly |
| IR-8 | Review the incident response plan every 12 months |

## 6. Control implementation statements

The statements for the 12 sampled controls are in the `SSP` sheet of [`control_assessment.xlsx`](control_assessment.xlsx). Each states who is responsible, whether the control is implemented or partially implemented, and how it works. The assessment then tests each statement against evidence.

## 7. Notes on use

This excerpt covers sections an assessor reads first (identification, boundary, categorization). A full SSP also includes data flows, interconnection agreements, a full inventory and statements for every control in the baseline.
