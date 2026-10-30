# NIST SP 800-53 Control Assessment: SSP Excerpt, Assessment and POA&M

## Project Description

This project is a sample security control assessment for a fictional Moderate-impact finance and payroll system, following the structure used in the NIST Risk Management Framework. It has three linked parts: a System Security Plan excerpt with the system's categorization and what the system owner says each control does; an assessment of 12 NIST SP 800-53 Rev. 5 controls that tests those statements against evidence; and a Plan of Action and Milestones (POA&M) that the workbook generates automatically from every control rated Other Than Satisfied. The findings reuse results from my ITGC testing and risk register projects, so the same weaknesses appear as audit exceptions, as risks and as framework control failures. This is the kind of work done for federal and regulated environments, and I built it to show I can write and test controls in that format. It is a sample, not an authorized assessment.

## Files

| File | Purpose |
|---|---|
| [`ssp_excerpt.md`](ssp_excerpt.md) | System identification, boundary, FIPS 199 categorization with rationale, and the parameters used |
| [`control_assessment.xlsx`](control_assessment.xlsx) | Workbook: Scope, SSP statements, Assessment, POAM and Summary sheets (211 formulas, no errors) |
| [`control_assessment_export.csv`](control_assessment_export.csv) | Assessment results readable on GitHub |
| [`scripts/build_assessment.py`](scripts/build_assessment.py) | Python (openpyxl) script that generates the workbook |

## How the workbook works

- **SSP sheet:** for each control, the responsible role, implementation status and the statement of how the control is met.
- **Assessment sheet:** the method (Examine, Interview, Test), objective, evidence, result, finding, risk rating, remediation plan, owner and due date. A formula numbers each Other Than Satisfied control and builds its POA&M reference.
- **POAM sheet:** every cell is a formula. It pulls each failed control's weakness, plan, owner and date from the Assessment sheet, and calculates days to completion and an OVERDUE flag against the as-of date. Change a result and the POA&M updates itself.
- **Summary sheet:** results, open weaknesses by risk rating, POA&M status and results by control family, with a chart.

## Results

Of 12 controls, 3 were Satisfied and 9 were Other Than Satisfied (25% Satisfied). Five weaknesses are rated High and four Moderate. Two POA&M items are overdue as of 2026-10-30.

| Control | Result | Risk | What was found |
|---|---|---|---|
| AC-2 Account Management | Other Than Satisfied | High | 2 of 4 terminated users still active; 2 of 14 users overdue for review |
| PS-4 Personnel Termination | Other Than Satisfied | High | Access not disabled within 24 hours of termination |
| AC-5 Separation of Duties | Other Than Satisfied | Moderate | One user holds AP Clerk and Purchasing Agent roles |
| IA-2 Identification and Authentication | Other Than Satisfied | High | No multi-factor authentication on remote access |
| CM-3 Configuration Change Control | Other Than Satisfied | High | 4 of 8 changes lacked valid independent approval |
| SI-2 Flaw Remediation | Other Than Satisfied | High | Critical ERP vulnerability unpatched past 30 days |
| SC-28 Protection of Information at Rest | Other Than Satisfied | Moderate | Payroll data in unencrypted spreadsheets |
| AU-6 Audit Record Review | Other Than Satisfied | Moderate | Log review is ad hoc, with no weekly record or alerting |
| IR-8 Incident Response Plan | Other Than Satisfied | Moderate | No tested ransomware runbook |
| CP-4 Contingency Plan Testing | Satisfied | | Restoration test completed and documented |
| RA-5 Vulnerability Monitoring and Scanning | Satisfied | | Monthly scans run and found the critical ERP flaw |
| AT-2 Literacy Training and Awareness | Satisfied | | Training material covers phishing and malware |

Two of the findings share a root cause. AC-2 and PS-4 both fail because terminated users kept access, so one fix (HR to IT deprovisioning) addresses both POA&M items. RA-5 passes while SI-2 fails: the scans work, and the problem is acting on what they find. Observations on the Satisfied controls (restoration was the only contingency test, and training completion metrics were unavailable) are recorded in the findings, not turned into weaknesses.

## Connections to other projects

- [`itgc-controls-testing`](https://github.com/mmanzanares7/itgc-controls-testing) is the source for AC-2, PS-4, AC-5 and CM-3.
- [`risk-register-excel`](https://github.com/mmanzanares7/risk-register-excel) is the source for the other findings and the POA&M dates.
- [`nist-csf-gap-assessment`](https://github.com/mmanzanares7/nist-csf-gap-assessment) rolls the same weaknesses up to the framework level.

## Limitations

The system, organization and findings are fictional, and the organization-defined parameters were chosen for the sample. Real assessments follow NIST SP 800-53A objectives for each control part and each enhancement, assess far more than 12 controls, and are performed by an independent assessor for an authorizing official. The IA-2 finding addresses the multi-factor requirement only, not every part of the control. I have not performed an authorization.

## Summary

I wrote the system context, categorized it, stated how 12 controls are implemented, tested those statements, and generated a POA&M from the failures. The result shows the SSP, SAR and POA&M relationship in practice, and how a single weakness is traced from an audit test to a risk to a control failure.
