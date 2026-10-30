"""Build control_assessment.xlsx: SSP implementation statements, a control assessment (SAR style)
and a POA&M for 12 NIST SP 800-53 Rev. 5 controls on a fictional system. The POA&M, results
counts, overdue flags and summary are formulas driven by the Assessment sheet.

Usage: python3 build_assessment.py ../control_assessment.xlsx
"""
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

OUT = sys.argv[1] if len(sys.argv) > 1 else "control_assessment.xlsx"
NAVY, BLUE = "1F3864", "0000FF"
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD = PatternFill("solid", start_color=NAVY, end_color=NAVY)
RED = PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")
AMB = PatternFill("solid", start_color="FFEB9C", end_color="FFEB9C")
GRN = PatternFill("solid", start_color="C6EFCE", end_color="C6EFCE")


def f(bold=False, color="000000", size=10, italic=False):
    return Font(name="Arial", bold=bold, color=color, size=size, italic=italic)


def head(ws, row, labels):
    for i, lab in enumerate(labels, 1):
        c = ws.cell(row=row, column=i, value=lab)
        c.font, c.fill, c.border = f(True, "FFFFFF"), HEAD, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 36


def title(ws, text, sub):
    ws["A1"], ws["A2"] = text, sub
    ws["A1"].font, ws["A2"].font = f(True, NAVY, 14), f(italic=True, color="595959")


def put(ws, r, c, v, inp=False, center=False, bold=False):
    x = ws.cell(row=r, column=c, value=v)
    x.font = f(bold=bold, color=BLUE if inp else "000000")
    x.border = BOX
    x.alignment = Alignment(wrap_text=True, vertical="top", horizontal="center" if center else None)
    return x


# id, name, family, role, ssp status, ssp statement, method, objective, evidence, result, finding, risk, source, plan, owner, due, plan status
C = [
    ("AC-2", "Account Management", "AC", "IT Security", "Implemented",
     "Accounts are created on manager approval. Accounts for terminated personnel are disabled within 24 hours of HR notification. Department managers review access every six months.",
     "Examine, Test", "Terminated users have no active account; every active user was reviewed in the last six months.",
     "User access list (16 users), HR termination dates, access review records, test date 2026-06-15",
     "Other Than Satisfied",
     "2 of 4 terminated users (U1003, U1010) were still Active. 2 of 14 active users (U1010, U1016) were past the six-month review cycle.",
     "High", "itgc-controls-testing, Controls 1 and 2",
     "Disable both accounts, review their activity, automate deprovisioning from HR events, complete overdue reviews.", "IT Security", date(2026, 10, 15), "In progress"),
    ("PS-4", "Personnel Termination", "PS", "HR and IT Security", "Implemented",
     "HR notifies IT Security of terminations the same day. System access is disabled within 24 hours and credentials are retrieved.",
     "Examine, Interview", "System access is disabled within 24 hours of termination.",
     "HR termination report, user access list, interview with IT Security lead",
     "Other Than Satisfied",
     "Terminations on 2026-05-15 and 2026-06-02 had not resulted in disabled accounts by 2026-06-15, well past the 24 hour requirement.",
     "High", "itgc-controls-testing, Control 1",
     "Add a same-day HR to IT notification step with a ticket, and reconcile terminations to active accounts monthly.", "HR", date(2026, 10, 15), "In progress"),
    ("AC-5", "Separation of Duties", "AC", "Finance Controller", "Implemented",
     "Roles are defined so purchasing and payment functions are separated. Conflicting role combinations are not assigned.",
     "Examine, Test", "No user holds a combination of roles listed as conflicting.",
     "Role assignment report, segregation of duties conflict matrix",
     "Other Than Satisfied",
     "User U1016 holds AP Clerk and Purchasing Agent, so one person can create a purchase order and process its payment. No compensating review found.",
     "Moderate", "itgc-controls-testing, Control 4",
     "Remove the Purchasing Agent role; until then, supervisor review of all payments by this user.", "Finance Controller", date(2026, 11, 15), "Open"),
    ("IA-2", "Identification and Authentication (Organizational Users)", "IA", "IT Security", "Partially implemented",
     "Users authenticate with unique IDs and passwords. Multi-factor authentication is required for privileged accounts. MFA for remote access is planned.",
     "Examine, Test", "Multi-factor authentication is required for remote access and privileged accounts.",
     "VPN authentication configuration, privileged account list",
     "Other Than Satisfied",
     "Remote access over VPN does not require multi-factor authentication.",
     "High", "risk-register-excel, R-10",
     "Enforce MFA for all VPN users.", "IT Security", date(2026, 12, 1), "Open"),
    ("CM-3", "Configuration Change Control", "CM", "IT Infrastructure", "Implemented",
     "All production changes are recorded in the ticketing system and approved by someone other than the requester before implementation.",
     "Examine, Test", "Each production change has independent approval dated before implementation.",
     "Change management log (8 changes), approval dates",
     "Other Than Satisfied",
     "4 of 8 changes failed: CH-2203 and CH-2207 had no approver, CH-2206 was approved after implementation, CH-2208 was self-approved.",
     "High", "itgc-controls-testing, Control 3",
     "Block deployment until independent approval is logged; review the four changes after the fact.", "IT Infrastructure", date(2026, 12, 15), "Open"),
    ("CP-4", "Contingency Plan Testing", "CP", "IT Infrastructure", "Implemented",
     "The contingency plan is tested every year, including restoration of the financial database from backup.",
     "Examine", "A restoration test was performed and documented in the last 12 months.",
     "Restoration test report and sign-off",
     "Satisfied",
     "A full restoration test of the financial database was completed and documented. Scope was restoration only; exercises of the wider plan were not evidenced (observation, not a finding).",
     "", "risk-register-excel, R-13 (closed)", "", "", None, ""),
    ("RA-5", "Vulnerability Monitoring and Scanning", "RA", "IT Security", "Implemented",
     "Authenticated vulnerability scans run monthly and results are reported to the system owner.",
     "Examine, Test", "Scans run monthly and identify vulnerabilities in the ERP environment.",
     "Scan reports for the last three months",
     "Satisfied",
     "Scans run monthly and identified a critical ERP vulnerability, which shows the monitoring works. Remediation is assessed under SI-2.",
     "", "risk-register-excel, R-09", "", "", None, ""),
    ("SI-2", "Flaw Remediation", "SI", "IT Infrastructure", "Implemented",
     "Critical flaws are remediated within 30 days of vendor patch release.",
     "Examine, Test", "Critical flaws are patched within 30 days.",
     "Scan reports, patch records, change tickets",
     "Other Than Satisfied",
     "A critical vulnerability on the production ERP application server remains unpatched past the 30 day requirement.",
     "High", "risk-register-excel, R-09",
     "Apply the vendor patch in the next maintenance window; add a compensating WAF rule until then.", "IT Infrastructure", date(2026, 11, 20), "In progress"),
    ("SC-28", "Protection of Information at Rest", "SC", "Finance Controller", "Implemented",
     "Sensitive financial and payroll data is encrypted at rest in the ERP database and in approved storage.",
     "Examine, Test", "Payroll and financial data is encrypted wherever it is stored.",
     "Shared drive file review, database encryption settings",
     "Other Than Satisfied",
     "Payroll data was found in unencrypted spreadsheets on a shared drive.",
     "Moderate", "risk-register-excel, R-14",
     "Move the data to an access-controlled system and encrypt existing files.", "Finance Controller", date(2026, 12, 10), "Open"),
    ("AU-6", "Audit Record Review, Analysis, and Reporting", "AU", "IT Security", "Implemented",
     "Audit logs are reviewed weekly and anomalies are investigated.",
     "Examine, Interview", "Logs are reviewed on a defined schedule and anomalies are followed up.",
     "Review records, interview with IT Security lead, detection script output",
     "Other Than Satisfied",
     "Log review is done with an ad hoc script on exported logs. There is no weekly review record or alerting.",
     "Moderate", "log-anomaly-detection; nist-csf-gap-assessment DE.CM",
     "Centralize logs in a SIEM, schedule detections and keep weekly review records.", "IT Security", date(2027, 2, 15), "Open"),
    ("IR-8", "Incident Response Plan", "IR", "IT Security", "Partially implemented",
     "The incident response plan covers phishing and malware and is reviewed every year. A ransomware runbook is planned.",
     "Examine", "The plan covers the incident types the system faces and was reviewed in the last 12 months.",
     "Incident response plan, phishing playbook, review history",
     "Other Than Satisfied",
     "No tested runbook exists for ransomware, the most likely high-impact incident for this system.",
     "Moderate", "risk-register-excel, R-11; phishing-incident-response",
     "Write the ransomware runbook and test it in a tabletop exercise.", "IT Security", date(2027, 1, 15), "Open"),
    ("AT-2", "Literacy Training and Awareness", "AT", "IT Security", "Implemented",
     "All users complete annual security awareness training. A plain-language guide covers phishing and malware.",
     "Examine", "Users receive awareness training and reference material on phishing and malware.",
     "Reference guide, training policy",
     "Satisfied",
     "Training material exists and covers phishing and malware with reporting steps. Completion and phishing simulation metrics were not available (observation).",
     "", "security-awareness-phishing-malware-guide", "", "", None, ""),
]
N = len(C)
FIRST, LAST = 5, 5 + N - 1        # 5..16

wb = Workbook()

# ---------------------------------------------------------------- Scope
sc = wb.active
sc.title = "Scope"
title(sc, "NIST SP 800-53 Rev. 5 Control Assessment: Sample", "Fictional system and findings for portfolio use. Not an authorized assessment.")
rows = [
    ("System", "Finance and Payroll ERP (FPE), a fictional system supporting accounts payable, general ledger, procurement and payroll."),
    ("Categorization", "FIPS 199 Moderate for confidentiality, integrity and availability. See ssp_excerpt.md for the rationale."),
    ("Baseline", "NIST SP 800-53 Rev. 5 Moderate baseline. This sample covers 12 controls."),
    ("Method", "NIST SP 800-53A style: Examine, Interview and Test. Each control is rated Satisfied or Other Than Satisfied."),
    ("Parameters", "Values such as 24 hours, six months and 30 days are organization-defined parameters chosen for this fictional system."),
    ("Sheets", "SSP (what the system owner says), Assessment (what was found), POAM (formula-driven from Assessment), Summary."),
    ("Colour key", "Blue text is an input. Black text is a formula."),
    ("Sources", "Findings for AC-2, PS-4, AC-5 and CM-3 come from my itgc-controls-testing repo. The rest correspond to risks in my risk-register-excel repo."),
    ("Limits", "All data is synthetic. This sample shows the structure and logic of an SSP, SAR and POA&M. It does not represent an actual authorization."),
]
for i, (k, v) in enumerate(rows):
    sc.cell(row=4 + i, column=1, value=k).font = f(True)
    c = sc.cell(row=4 + i, column=2, value=v)
    c.font, c.alignment = f(), Alignment(wrap_text=True, vertical="top")
    sc.cell(row=4 + i, column=1).alignment = Alignment(vertical="top")
sc["A15"] = "Settings"
sc["A15"].font = f(True, NAVY, 11)
sc["A16"] = "Report as-of date"
sc["A16"].font = f(True)
sc["B16"] = date(2026, 10, 30)
sc["B16"].font = f(color=BLUE)
sc["B16"].number_format = "yyyy-mm-dd"
sc["B16"].alignment = Alignment(horizontal="left")
sc["A17"] = "Fixed so results are reproducible. Replace with =TODAY() for a live view."
sc["A17"].font = f(italic=True, color="595959")
sc.column_dimensions["A"].width = 20
sc.column_dimensions["B"].width = 110
sc.sheet_view.showGridLines = False
ASOF = "Scope!$B$16"

# ---------------------------------------------------------------- SSP
ss = wb.create_sheet("SSP")
title(ss, "System Security Plan: Control Implementation Statements", "What the system owner states. The Assessment sheet tests these statements.")
head(ss, 4, ["Control", "Name", "Family", "Responsible role", "Implementation status", "Implementation statement"])
for i, c in enumerate(C):
    r = FIRST + i
    put(ss, r, 1, c[0], center=True, bold=True)
    put(ss, r, 2, c[1], inp=True)
    put(ss, r, 3, c[2], inp=True, center=True)
    put(ss, r, 4, c[3], inp=True, center=True)
    put(ss, r, 5, c[4], inp=True, center=True)
    put(ss, r, 6, c[5], inp=True)
    ss.row_dimensions[r].height = 56
for col, w in zip("ABCDEF", (10, 36, 8, 18, 18, 95)):
    ss.column_dimensions[col].width = w
ss.freeze_panes = "C5"
dvi = DataValidation(type="list", formula1='"Implemented,Partially implemented,Planned,Not applicable"')
ss.add_data_validation(dvi)
dvi.add(f"E{FIRST}:E{LAST}")

# ---------------------------------------------------------------- Assessment
asx = wb.create_sheet("Assessment")
title(asx, "Control Assessment (SAR)", "Blue cells are inputs. POA&M reference and sequence are formulas.")
head(asx, 4, ["Control", "Name", "Method", "Assessment objective", "Evidence reviewed", "Result", "Finding", "Risk rating",
              "Source in portfolio", "POA&M ref", "Remediation plan", "Owner", "Scheduled completion", "Plan status", "OTS seq (helper)"])
for i, c in enumerate(C):
    r = FIRST + i
    (cid, nm, fam, role, st, stmt, meth, obj, ev, res, fin, risk, src, plan, own, due, pst) = c
    put(asx, r, 1, f"=SSP!A{r}", center=True, bold=True)
    put(asx, r, 2, f"=SSP!B{r}")
    put(asx, r, 3, meth, inp=True, center=True)
    put(asx, r, 4, obj, inp=True)
    put(asx, r, 5, ev, inp=True)
    put(asx, r, 6, res, inp=True, center=True, bold=True)
    put(asx, r, 7, fin, inp=True)
    put(asx, r, 8, risk or None, inp=True, center=True)
    put(asx, r, 9, src, inp=True)
    put(asx, r, 10, f'=IF(O{r}="","","POAM-"&TEXT(O{r},"00"))', center=True)
    put(asx, r, 11, plan or None, inp=True)
    put(asx, r, 12, own or None, inp=True, center=True)
    x = put(asx, r, 13, due, inp=True, center=True)
    x.number_format = "yyyy-mm-dd"
    put(asx, r, 14, pst or None, inp=True, center=True)
    put(asx, r, 15, f'=IF(F{r}="Other Than Satisfied",COUNTIF($F${FIRST}:F{r},"Other Than Satisfied"),"")', center=True)
    asx.row_dimensions[r].height = 82
for col, w in zip("ABCDEFGHIJKLMNO", (9, 30, 13, 36, 32, 14, 48, 11, 24, 11, 44, 16, 13, 12, 9)):
    asx.column_dimensions[col].width = w
asx.freeze_panes = "C5"
dvr = DataValidation(type="list", formula1='"Satisfied,Other Than Satisfied,Not assessed"')
dvk = DataValidation(type="list", formula1='"Low,Moderate,High"', allow_blank=True)
dvp = DataValidation(type="list", formula1='"Open,In progress,Complete"', allow_blank=True)
for dv in (dvr, dvk, dvp):
    asx.add_data_validation(dv)
dvr.add(f"F{FIRST}:F{LAST}")
dvk.add(f"H{FIRST}:H{LAST}")
dvp.add(f"N{FIRST}:N{LAST}")
for word, fill in (("Other Than Satisfied", RED), ("Satisfied", GRN)):
    asx.conditional_formatting.add(f"F{FIRST}:F{LAST}", CellIsRule(operator="equal", formula=[f'"{word}"'], fill=fill))
for word, fill in (("High", RED), ("Moderate", AMB), ("Low", GRN)):
    asx.conditional_formatting.add(f"H{FIRST}:H{LAST}", CellIsRule(operator="equal", formula=[f'"{word}"'], fill=fill))
AR = lambda col: f"Assessment!${col}${FIRST}:${col}${LAST}"

# ---------------------------------------------------------------- POAM
pm = wb.create_sheet("POAM")
title(pm, "Plan of Action and Milestones", "Every cell is a formula. A row appears for each control rated Other Than Satisfied.")
head(pm, 4, ["POA&M ref", "Control", "Weakness", "Risk rating", "Remediation plan", "Owner", "Scheduled completion", "Status",
             "Days to completion", "Overdue?"])
for n in range(1, N + 1):
    r = 4 + n
    m = f"MATCH({n},{AR('O')},0)"
    put(pm, r, 1, f'=IFERROR(INDEX({AR("J")},{m}),"")', center=True, bold=True)
    put(pm, r, 2, f'=IFERROR(INDEX({AR("A")},{m})&" "&INDEX({AR("B")},{m}),"")')
    put(pm, r, 3, f'=IFERROR(INDEX({AR("G")},{m}),"")')
    put(pm, r, 4, f'=IFERROR(INDEX({AR("H")},{m}),"")', center=True)
    put(pm, r, 5, f'=IFERROR(INDEX({AR("K")},{m}),"")')
    put(pm, r, 6, f'=IFERROR(INDEX({AR("L")},{m}),"")', center=True)
    put(pm, r, 7, f'=IFERROR(INDEX({AR("M")},{m}),"")', center=True).number_format = "yyyy-mm-dd"
    put(pm, r, 8, f'=IFERROR(INDEX({AR("N")},{m}),"")', center=True)
    put(pm, r, 9, f'=IF(OR(A{r}="",H{r}="Complete"),"",G{r}-{ASOF})', center=True)
    put(pm, r, 10, f'=IF(OR(A{r}="",H{r}="Complete"),"",IF(G{r}<{ASOF},"OVERDUE",""))', center=True, bold=True)
    pm.row_dimensions[r].height = 70
for col, w in zip("ABCDEFGHIJ", (11, 32, 52, 11, 46, 17, 13, 11, 11, 11)):
    pm.column_dimensions[col].width = w
pm.freeze_panes = "C5"
for word, fill in (("High", RED), ("Moderate", AMB), ("Low", GRN)):
    pm.conditional_formatting.add(f"D5:D{4 + N}", CellIsRule(operator="equal", formula=[f'"{word}"'], fill=fill))
pm.conditional_formatting.add(f"J5:J{4 + N}", CellIsRule(operator="equal", formula=['"OVERDUE"'], fill=RED))

# ---------------------------------------------------------------- Summary
sm = wb.create_sheet("Summary")
title(sm, "Assessment Summary", "All figures are formulas.")
head(sm, 4, ["Result", "Controls"])
for i, k in enumerate(["Satisfied", "Other Than Satisfied", "Not assessed"]):
    put(sm, 5 + i, 1, k, bold=True)
    put(sm, 5 + i, 2, f'=COUNTIF({AR("F")},A{5 + i})', center=True)
put(sm, 8, 1, "Total controls", bold=True)
put(sm, 8, 2, "=SUM(B5:B7)", center=True, bold=True)
put(sm, 9, 1, "Percent Satisfied", bold=True)
put(sm, 9, 2, "=IF(B8=0,0,B5/B8)", center=True, bold=True).number_format = "0%"
head(sm, 11, ["Open weaknesses by risk", "Count"])
for i, k in enumerate(["High", "Moderate", "Low"]):
    put(sm, 12 + i, 1, k, bold=True)
    put(sm, 12 + i, 2, f'=COUNTIFS({AR("F")},"Other Than Satisfied",{AR("H")},A{12 + i},{AR("N")},"<>Complete")', center=True)
head(sm, 16, ["POA&M status", "Count"])
put(sm, 17, 1, "Open or in progress", bold=True)
put(sm, 17, 2, f'=COUNTIF(POAM!H5:H{4 + N},"Open")+COUNTIF(POAM!H5:H{4 + N},"In progress")', center=True)
put(sm, 18, 1, "Overdue", bold=True)
put(sm, 18, 2, f'=COUNTIF(POAM!J5:J{4 + N},"OVERDUE")', center=True)
head(sm, 20, ["Control family", "Controls", "Satisfied", "Other Than Satisfied"])
fams = ["AC", "AT", "AU", "CM", "CP", "IA", "IR", "PS", "RA", "SC", "SI"]
for i, fm in enumerate(fams):
    r = 21 + i
    put(sm, r, 1, fm, bold=True)
    put(sm, r, 2, f'=COUNTIF(SSP!$C${FIRST}:$C${LAST},A{r})', center=True)
    put(sm, r, 3, f'=COUNTIFS(SSP!$C${FIRST}:$C${LAST},A{r},{AR("F")},"Satisfied")', center=True)
    put(sm, r, 4, f'=COUNTIFS(SSP!$C${FIRST}:$C${LAST},A{r},{AR("F")},"Other Than Satisfied")', center=True)
for col, w in zip("ABCD", (30, 11, 11, 20)):
    sm.column_dimensions[col].width = w
sm.sheet_view.showGridLines = False
ch = BarChart()
ch.type = "bar"
ch.grouping = "stacked"
ch.overlap = 100
ch.title = "Control results by family"
ch.add_data(Reference(sm, min_col=3, max_col=4, min_row=20, max_row=20 + len(fams)), titles_from_data=True)
ch.set_categories(Reference(sm, min_col=1, min_row=21, max_row=20 + len(fams)))
ch.series[0].graphicalProperties.solidFill = "70AD47"
ch.series[1].graphicalProperties.solidFill = "C00000"
ch.height, ch.width = 8.5, 13
ch.x_axis.delete = False
ch.y_axis.delete = False
sm.add_chart(ch, "F4")

wb._sheets = [wb[n] for n in ["Scope", "SSP", "Assessment", "POAM", "Summary"]]
wb.active = 2
wb.save(OUT)
print("saved", OUT)
