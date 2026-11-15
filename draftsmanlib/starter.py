"""The starter library: working templates and clauses, written the first time
Draftsman meets an empty library and topped up when a new version brings more.
They are starting points - plain Word files to reword, restyle, rename or
delete. A starter you delete stays deleted: the library keeps a list of the
ones it has already been given.

Each template below is a list of lines in a small shorthand, so that the
wording is all that has to be read here:

    "text"              a paragraph            "# text" / "## text"   headings
    "! text"            a bold paragraph       "* text"               a bullet
    ""                  an empty paragraph
    ("letter", subject, who)    letterhead, address block for {{<who> Name}}, date, greeting, subject
    ("sign",)                   Yours sincerely + the handler's details
    ("kv", [(label, text), ...])            a two-column table
    ("grid", [headings], rows)              a table: rows is a number of blank rows or a list of rows
    ("ctl",)                                the document-control table
    ("foot", text)                          the page footer
"""
from __future__ import annotations

import os

from .library import CLAUSES, STANDING, TEMPLATES

MANIFEST = ".draftsman-starters"
README_NAME = "READ ME.txt"

README = """DRAFTSMAN LIBRARY
=================

Templates   Word documents with {{tags}} in them. Any folders you like.
Clauses     Reusable wording. A file straight in Clauses is a single clause;
            a folder in Clauses is a tick-list of the clauses inside it.
            Number the files (1 ..., 2 ...) to fix the order they go in.
standing values.txt
            Name = value, one per line. Any field with that name is filled
            in for you on every template (your organisation, your own name).

TAGS
  {{Name}}                 a field - becomes a box on the form
  {{Name|default}}         a field with a starting value
  {{Due|today+1m}}         a date worked out for you: today, today+14,
                           today+2w, today+1m, today+1y (and - instead of +)
  {{Outcome=Upheld|Partly upheld|Not upheld}}
                           a field with a drop-down of choices
  {{Address+}}             a field with room for several lines
  {{Middle Name?}}         a field that is fine to leave empty (no warning)
  {{TODAY}}  {{TODAY_SHORT}}  {{YEAR}}
  {{DOC_REF}} {{DOC_TITLE}} {{VERSION}} {{OWNER}} {{ISSUE_DATE}} {{REVIEW_DATE}}

  Each of these alone in its own paragraph, in the body of the document:
  {{clause:Complaint rights}}      drops in that clause
  {{clauses:DSAR exemptions}}      a tick-list of the clauses in that folder
  {{#if Extended}} ... {{/if}}     keeps what is between them only when the
                                   field (or tick-box) is filled in / "yes"
  {{#unless Extended}} ... {{/if}} the opposite
  {{#else}}                        between the two: the other case

The templates here are starters. Check the wording against your own
procedures before the first real use, then make them yours: they are
ordinary Word files and nothing in the app depends on their names. Delete
the ones you do not want - they will not come back.

(This file is rewritten by Draftsman; keep your own notes somewhere else.)
"""

STANDING_TEXT = """# Filled in automatically wherever a template has a field with the same name.
# Name = value. Indent the following lines to carry a value over several lines.
Organisation =
Organisation Address =
Handler Name =
Handler Title =
Contact Email =
Contact Phone =
"""

CLAUSE_TEXT = {
    "Complaint rights.txt": (
        "If you are unhappy with how we have handled your request, please tell us first so that we can put it right: "
        "contact {{Handler Name}} at {{Contact Email}}.\n"
        "You also have the right to complain to the Information Commissioner's Office (ICO). You can do this at "
        "ico.org.uk/make-a-complaint or by calling 0303 123 1113. You may also be able to enforce your rights through the courts."
    ),
    "Rights summary.txt": (
        "Under data protection law you have the right to ask for a copy of the information we hold about you, to ask us "
        "to correct it, and in some circumstances to ask us to erase it, to restrict what we do with it or to object to "
        "our using it. Where we rely on your consent you can withdraw it at any time. To use any of these rights, "
        "contact {{Handler Name}} at {{Contact Email}}."
    ),
    "Confidentiality notice.txt": (
        "This document contains confidential information and is intended only for the person or organisation it is "
        "addressed to. If you have received it in error, please tell us and delete it."
    ),
    "DSAR exemptions/1 Information about other people.txt": (
        "Some of the information relates to other people as well as to you. We have removed details that would identify "
        "them where they have not agreed to it being shared and it would not be reasonable to share it without their agreement."
    ),
    "DSAR exemptions/2 Legal professional privilege.txt": (
        "Some of the information consists of confidential communications with our legal advisers. It is covered by legal "
        "professional privilege and is exempt from the right of access."
    ),
    "DSAR exemptions/3 Confidential references.txt": (
        "Some of the information is a reference given in confidence for employment, training or education purposes. "
        "Confidential references are exempt from the right of access."
    ),
    "DSAR exemptions/4 Management planning.txt": (
        "Some of the information was recorded for management forecasting or planning, and giving it to you now would be "
        "likely to prejudice the conduct of our business."
    ),
    "DSAR exemptions/5 Negotiations.txt": (
        "Some of the information records our intentions in negotiations with you, and giving it to you now would be "
        "likely to prejudice those negotiations."
    ),
    "DSAR exemptions/6 Not your personal data.txt": (
        "Some of the documents we found also contain information that is not about you. The right of access covers your "
        "own personal data, so that other information has been removed."
    ),
    "Breach advice/1 Passwords.txt": (
        "Change the password for any account that used the same or a similar password, and turn on two-step "
        "verification wherever it is offered."
    ),
    "Breach advice/2 Suspicious contact.txt": (
        "Be cautious about unexpected emails, calls or texts that mention us or ask for personal or financial details. "
        "We will never ask you for a password or for bank details in this way. If in doubt, contact us using the "
        "details in this letter."
    ),
    "Breach advice/3 Bank and credit.txt": (
        "Keep an eye on your bank and card statements and tell your bank straight away about anything you do not "
        "recognise. You can check your credit report free of charge with the main credit reference agencies."
    ),
    "Breach advice/4 Report fraud.txt": (
        "If you think you have been a victim of fraud, report it to Action Fraud at actionfraud.police.uk or on "
        "0300 123 2040."
    ),
    "Erasure exceptions/1 Legal obligation.txt": (
        "We are required by law to keep this information for a set period, so we are not able to erase it before then."
    ),
    "Erasure exceptions/2 Legal claims.txt": (
        "We need to keep this information to establish, exercise or defend legal claims."
    ),
    "Erasure exceptions/3 Still needed.txt": (
        "The information is still needed for the purpose it was collected for and we have a lawful basis for keeping "
        "it, so none of the grounds for erasure applies."
    ),
}

SUPPLEMENTARY = [
    "! About how we use your information",
    ("labelled", [
        ("Why we hold it", "{{Purposes+|To deliver our services to you and to meet our legal and contractual obligations.}}"),
        ("The kinds of information", "{{Categories+|Contact details, records of the support you received and correspondence with you.}}"),
        ("Who it has been shared with", "{{Recipients+|Our funders and delivery partners, where the service requires it.}}"),
        ("How long we keep it", "{{Retention Period|In line with our retention schedule}}"),
        ("Where it came from", "{{Source|From you, and from the organisation that referred you to us}}"),
    ]),
    "You have the right to ask us to correct information that is wrong, and in some circumstances to erase it, to "
    "restrict what we do with it, or to object to our using it.",
]

SAR = "Subject access requests/"
RIGHTS = "Individual rights/"
DP = "Data protection/"
SEC = "Information security/"
QA = "Quality and audit/"
GOV = "Governance/"
POL = "Policies/"
COMP = "Complaints/"
LET = "Letters and notes/"
PART = "Participants/"

T = {}

# ------------------------------------------------------- subject access
T[SAR + "1 Acknowledgement"] = [
    ("letter", "Your request for your personal information - {{Reference}}", "Requester"),
    "Thank you for your request, which we received on {{Date Received}}, for a copy of the personal information "
    "{{Organisation}} holds about you.",
    "We are dealing with it as a subject access request under the UK General Data Protection Regulation. We will send "
    "you our response without undue delay and by {{Due Date}} at the latest.",
    "{{#if Clarification Needed}}",
    "To help us find the information you want, please tell us: {{Clarification Needed+}}",
    "The time we have to respond is paused from today until we hear from you, and the date above will move back by "
    "the same number of days.",
    "{{/if}}",
    "{{#if Identity Needed}}",
    "Before we can release personal information we need to be sure who we are sending it to. Please send us: {{Identity Needed+}}",
    "The time we have to respond starts when we receive it.",
    "{{/if}}",
    "If you have any questions in the meantime, please contact me using the details below and quote the reference above.",
    ("sign",), ("foot", "{{Reference}} - acknowledgement"),
]
T[SAR + "2 Extension notice"] = [
    ("letter", "Your request for your personal information - more time needed - {{Reference}}", "Requester"),
    "I am writing about your subject access request, which we received on {{Date Received}}.",
    "We normally have one month to respond. The law allows this to be extended by up to a further two months where a "
    "request is complex or where a person has made a number of requests. We need to extend the time in your case "
    "because: {{Extension Reason+}}",
    "We will now respond by {{Extended Due Date}} at the latest, and sooner if we can.",
    "{{clause:Complaint rights}}",
    ("sign",), ("foot", "{{Reference}} - extension notice"),
]
T[SAR + "3 Response"] = [
    ("letter", "Response to your subject access request - {{Reference}}", "Requester"),
    "Thank you for your request, received on {{Date Received}}, for a copy of the personal information "
    "{{Organisation}} holds about you. We have now completed it.",
    "{{Search Summary+|We searched our email system, our case records and our shared files for information relating to you.}}",
    "A copy of your personal information is enclosed as {{Enclosure Description|a single indexed document}}. It is "
    "being sent to you {{Delivery Method|by secure email}}.",
    "{{#if Redactions Applied}}",
    "In places, information has been removed. Each place is marked REDACTED. This is because:",
    "{{clauses:DSAR exemptions}}",
    "{{/if}}",
    "{{#if Information Withheld}}",
    "A small number of documents have been withheld in full, for the reasons given above.",
    "{{/if}}",
    "{{clause:Supplementary information}}",
    "{{clause:Complaint rights}}",
    ("sign",), ("foot", "{{Reference}} - response"),
]
T[SAR + "4 Proof of identity request"] = [
    ("letter", "Your request for your personal information - confirming who you are - {{Reference}}", "Requester"),
    "Thank you for your request, which we received on {{Date Received}}, for a copy of the personal information "
    "{{Organisation}} holds about you.",
    "Before we release personal information we have to be sure we are sending it to the right person. We are not able "
    "to confirm that from the details we have, so please send us: {{Identity Needed+|a copy of one document showing "
    "your name and current address, such as a recent utility bill or bank statement}}",
    "You can send it {{Return Method|by replying to this email}}. We will use it only to confirm your identity and "
    "will delete it once we have done so.",
    "The one month we have to respond will start on the day we receive it. If we have not heard from you by "
    "{{Reply By|today+1m}} we will treat the request as closed; you are welcome to make a new request at any time.",
    ("sign",), ("foot", "{{Reference}} - proof of identity request"),
]
T[SAR + "5 Third-party consent request"] = [
    ("letter", "Request for your agreement to share information - {{Reference}}", "Third Party"),
    "{{Organisation}} has received a request from a person for a copy of the information we hold about them. They "
    "have a legal right to this under data protection law.",
    "Some of that information also mentions you: {{What Mentions Them+}}",
    "We are writing to ask whether you agree to this information being shared with the person who made the request. "
    "You do not have to agree, and you do not have to give a reason.",
    "Please let us know by {{Reply By|today+7}}. If you do not agree, or we do not hear from you, we will decide "
    "whether it is reasonable to share the information without your agreement, taking into account any duty of "
    "confidence owed to you and anything you tell us. Where we can, we will remove details that identify you.",
    ("sign",), ("foot", "{{Reference}} - third-party consent request"),
]
T[SAR + "6 No information held"] = [
    ("letter", "Response to your subject access request - {{Reference}}", "Requester"),
    "Thank you for your request, received on {{Date Received}}, for a copy of the personal information "
    "{{Organisation}} holds about you.",
    "{{Search Summary+|We searched our email system, our case records and our shared files using the details you gave us.}}",
    "We can confirm that we do not hold any personal information about you{{Scope Note?}}.",
    "If you think we should hold information about you, for example because you have used one of our services under "
    "a different name or address, please let us know and we will search again.",
    "{{clause:Complaint rights}}",
    ("sign",), ("foot", "{{Reference}} - no information held"),
]
T[SAR + "7 Refusal - unfounded or excessive"] = [
    ("letter", "Your subject access request - {{Reference}}", "Requester"),
    "Thank you for your request, received on {{Date Received}}.",
    "We have considered it carefully. Data protection law allows an organisation to refuse a request that is "
    "manifestly unfounded or excessive. We have decided that your request is "
    "{{Ground=manifestly unfounded|manifestly excessive|manifestly unfounded and excessive}} and we will not be "
    "responding to it further. Our reasons are:",
    "{{Reasons+}}",
    "{{#if What We Can Offer}}",
    "We would still like to help where we can: {{What We Can Offer+}}",
    "{{/if}}",
    "{{clause:Complaint rights}}",
    ("sign",), ("foot", "{{Reference}} - refusal"),
]

# ---------------------------------------------------- individual rights
T[RIGHTS + "Erasure request response"] = [
    ("letter", "Your request to erase your personal information - {{Reference}}", "Requester"),
    "Thank you for your request, received on {{Date Received}}, asking {{Organisation}} to erase the personal "
    "information we hold about you.",
    "{{#if Erased}}",
    "We have erased the following: {{What Was Erased+}}",
    "Where we had shared this information with other organisations, we have told them about your request.",
    "{{/if}}",
    "{{#if Kept}}",
    "We have not erased the following: {{What Was Kept+}}",
    "This is because:",
    "{{clauses:Erasure exceptions}}",
    "It will be kept securely, used only for that purpose, and destroyed {{Destruction Point|at the end of our retention period}}.",
    "{{/if}}",
    "{{clause:Complaint rights}}",
    ("sign",), ("foot", "{{Reference}} - erasure response"),
]
T[RIGHTS + "Rectification request response"] = [
    ("letter", "Your request to correct your personal information - {{Reference}}", "Requester"),
    "Thank you for your request, received on {{Date Received}}, asking {{Organisation}} to correct the information we "
    "hold about you.",
    "{{#if Corrected}}",
    "We have made the following changes: {{What Was Corrected+}}",
    "Where we had shared the earlier information with other organisations, we have told them about the correction.",
    "{{/if}}",
    "{{#if Not Corrected}}",
    "We have not changed the following: {{What Was Not Corrected+}}",
    "This is because we are satisfied the record is accurate: {{Why Accurate+}}",
    "We have added a note to the record to show that you disagree with it.",
    "{{/if}}",
    "{{clause:Complaint rights}}",
    ("sign",), ("foot", "{{Reference}} - rectification response"),
]
T[RIGHTS + "Objection or restriction response"] = [
    ("letter", "Your {{Request Type=objection to our use of|request to restrict our use of}} your personal information - {{Reference}}", "Requester"),
    "Thank you for your request, received on {{Date Received}}.",
    "You asked us to: {{What Was Asked+}}",
    "Our decision: {{Decision=We have agreed to your request|We have agreed to your request in part|We are not able to agree to your request}}.",
    "{{Decision Detail+}}",
    "{{#if Reasons}}",
    "Our reasons are: {{Reasons+}}",
    "{{/if}}",
    "{{clause:Complaint rights}}",
    ("sign",), ("foot", "{{Reference}} - response"),
]

# ------------------------------------------------------ data protection
T[DP + "Personal data breach record"] = [
    "# Personal data breach record",
    ("kv", [("Reference", "{{Breach Reference}}"), ("Recorded by", "{{Handler Name}}"), ("Date recorded", "{{TODAY}}"),
            ("When it happened", "{{Date Of Breach}}"), ("When we became aware", "{{Date Aware}}"),
            ("What happened", "{{What Happened+}}"), ("Information involved", "{{Data Involved+}}"),
            ("People affected (number and who)", "{{People Affected+}}"), ("Likely consequences", "{{Consequences+}}"),
            ("Containment and recovery", "{{Actions Taken+}}"),
            ("Risk to people's rights and freedoms", "{{Risk Level=Unlikely|Likely|High}} - {{Risk Reasoning+}}"),
            ("Reported to the ICO?", "{{Reported To ICO=No|Yes}}"), ("ICO reference and date", "{{ICO Reference}}"),
            ("Affected people told?", "{{People Told=No|Yes}}"),
            ("Controller, funder or prime contractor told?", "{{Partner Told=Not applicable|No|Yes}}"),
            ("Reason for those decisions", "{{Decision Reasons+}}"),
            ("Lessons and follow-up actions", "{{Follow Up+}}"), ("Signed off by", "{{Signed Off By}}")]),
    "",
    "A breach that is likely to result in a risk to people's rights and freedoms must be reported to the ICO without "
    "undue delay and, where feasible, within 72 hours of becoming aware of it.",
    ("foot", "{{Breach Reference}} - personal data breach record"),
]
T[DP + "Breach notification to individuals"] = [
    ("letter", "Important: an incident involving your personal information", "Recipient"),
    "I am writing to let you know about an incident that has affected some of the personal information "
    "{{Organisation}} holds about you. I am sorry that this has happened.",
    "! What happened",
    "{{What Happened+}}",
    "! What information was involved",
    "{{Data Involved+}}",
    "! What this could mean for you",
    "{{Consequences+}}",
    "! What we have done",
    "{{Actions Taken+}}",
    "{{#if Reported To ICO}}",
    "We have reported the incident to the Information Commissioner's Office.",
    "{{/if}}",
    "! What you can do",
    "{{clauses:Breach advice}}",
    "! Who to contact",
    "If you have any questions, or would like to talk this through, please contact {{Handler Name}}, "
    "{{Handler Title}}, at {{Contact Email}} or on {{Contact Phone}}.",
    ("sign",), ("foot", "{{Breach Reference}}"),
]
T[DP + "Breach notification to controller or prime contractor"] = [
    ("letter", "Notification of a personal data breach - {{Contract Name}} - {{Breach Reference}}", "Recipient"),
    "In line with our agreement, I am notifying you of a personal data breach affecting information we process "
    "{{Capacity=on your behalf|under the contract}}. We became aware of it on {{Date Aware}}.",
    ("kv", [("What happened", "{{What Happened+}}"), ("When", "{{Date Of Breach}}"),
            ("Information involved", "{{Data Involved+}}"),
            ("People affected", "{{People Affected+}}"),
            ("Likely consequences", "{{Consequences+}}"),
            ("What we have done", "{{Actions Taken+}}"),
            ("What we plan to do", "{{Next Steps+}}"),
            ("Our assessment of risk", "{{Risk Level=Unlikely|Likely|High}} - {{Risk Reasoning+}}")]),
    "",
    "This is {{Report Stage=our initial report; we will update you as we learn more|our final report}}. "
    "We have not contacted the Information Commissioner's Office or the people affected about this"
    "{{Contact Caveat|, and will not do so without your instruction}}.",
    "Our contact for this incident is {{Handler Name}}, {{Handler Title}}, {{Contact Email}}, {{Contact Phone}}.",
    ("sign",), ("foot", "{{Breach Reference}} - notification"),
]
T[DP + "DPIA"] = [
    "# Data protection impact assessment - {{Project Name}}",
    "Reference {{DOC_REF}}   |   Version {{VERSION}}   |   Owner {{OWNER}}   |   {{ISSUE_DATE}}",
    "## 1. What is proposed", "{{Proposal+}}",
    "## 2. Why a DPIA is needed", "{{Screening Outcome+}}",
    "## 3. The processing - nature, scope, context and purposes", "{{Processing Description+}}",
    "## 4. Lawful basis and, for special category data, the condition relied on", "{{Lawful Basis+}}",
    "## 5. Who has been consulted", "{{Consultation+}}",
    "## 6. Necessity and proportionality", "{{Necessity+}}",
    "## 7. Risks and measures",
    ("grid", ["Risk to individuals", "Likelihood and severity", "Measure to reduce it", "Residual risk"], 4),
    "## 8. Sign-off",
    "DPO advice: {{DPO Advice+}}",
    "Accepted by (senior responsible owner): {{Accepted By}}     Date: {{Acceptance Date}}",
    "Review date: {{REVIEW_DATE}}",
    ("foot", "{{DOC_REF}} v{{VERSION}} - DPIA - {{Project Name}}"),
]
T[DP + "DPIA screening"] = [
    "# DPIA screening - {{Project Name}}",
    ("kv", [("Completed by", "{{Handler Name}}"), ("Date", "{{TODAY}}"), ("Project lead", "{{Project Lead}}"),
            ("What is proposed", "{{Proposal+}}")]),
    "## Does the proposal involve any of these?",
    ("grid", ["Question", "Yes / No", "Notes"], [
        ["Special category or criminal offence data, or data about vulnerable people, on a large scale", "", ""],
        ["Systematic monitoring or tracking of people", "", ""],
        ["Profiling, scoring or automated decisions with a significant effect on people", "", ""],
        ["New technology, or a new use of existing technology", "", ""],
        ["Matching or combining data from different sources", "", ""],
        ["Processing that could stop someone getting a service or exercising a right", "", ""],
        ["Collecting personal data from a source other than the person, without telling them", "", ""],
        ["A new supplier or system that will hold our personal data", "", ""],
        ["Sending personal data outside the UK", "", ""],
    ]),
    "## Outcome",
    "{{Outcome=A full DPIA is required|A full DPIA is not required}}",
    "Reasons: {{Outcome Reasons+}}",
    "DPO: {{DPO Name}}     Date: {{TODAY}}",
    ("foot", "DPIA screening - {{Project Name}}"),
]
T[DP + "Legitimate interests assessment"] = [
    "# Legitimate interests assessment - {{Processing Name}}",
    ("kv", [("Completed by", "{{Handler Name}}"), ("Date", "{{TODAY}}"), ("Review date", "{{Review Date|today+1y}}")]),
    "## 1. Purpose test - is there a legitimate interest?",
    "What we want to do and why: {{Purpose+}}",
    "Who benefits, and how: {{Benefits+}}",
    "What would happen if we could not do it: {{Impact If Not+}}",
    "## 2. Necessity test - is the processing necessary?",
    "Does it actually help achieve the purpose, and is it proportionate: {{Necessity+}}",
    "Could the same be achieved with less data, or in a less intrusive way: {{Alternatives+}}",
    "## 3. Balancing test - do the individual's interests override ours?",
    "Our relationship with the people concerned, and what they would reasonably expect: {{Expectations+}}",
    "The kind of data, including any special category data or data about children or vulnerable people: {{Data Sensitivity+}}",
    "The possible impact on them, and how likely and severe it is: {{Impact+}}",
    "Safeguards, including whether they can easily opt out: {{Safeguards+}}",
    "## 4. Outcome",
    "{{Outcome=We can rely on legitimate interests|We cannot rely on legitimate interests}}",
    "{{Outcome Reasons+}}",
    "Approved by: {{Approved By}}",
    ("foot", "LIA - {{Processing Name}}"),
]
T[DP + "Data sharing agreement (short form)"] = [
    "# Data sharing agreement",
    "Between {{Organisation}} and {{Other Party}} (together, the parties).",
    ("kv", [("Reference", "{{Reference}}"), ("Start date", "{{Start Date|today}}"), ("Review date", "{{Review Date|today+1y}}"),
            ("Relationship", "{{Relationship=Controller to controller|Joint controllers|Controller to processor}}")]),
    "## 1. Why information is being shared", "{{Purpose+}}",
    "## 2. What will be shared", "{{Data Shared+}}",
    "Whose information it is: {{Data Subjects+}}",
    "## 3. Lawful basis",
    "{{Lawful Basis+}}",
    "For special category data, the condition relied on: {{Special Category Condition|Not applicable}}",
    "## 4. How it will be shared and how often", "{{Transfer Method+}}",
    "## 5. Keeping it secure",
    "Each party will keep the shared information secure, limit access to those who need it, and not pass it to "
    "anyone else without the other party's written agreement, unless required to by law.",
    "{{Security Measures+}}",
    "## 6. Keeping it accurate and for no longer than needed",
    "The shared information will be kept for {{Retention Period}} and then securely destroyed. Each party will tell "
    "the other about any correction that needs to be made.",
    "## 7. People's rights",
    "People will be told about the sharing through {{Privacy Information|each party's privacy notice}}. A request "
    "from a person about their information will be dealt with by the party that receives it, with the other's help "
    "where needed. Each party will tell the other within {{Notice Days|2}} working days of receiving one that "
    "concerns the shared information.",
    "## 8. If something goes wrong",
    "Each party will tell the other without undue delay, and within {{Breach Notice Hours|24}} hours, of becoming "
    "aware of a personal data breach involving the shared information, and will help with any report that has to be made.",
    "## 9. Review and ending",
    "This agreement will be reviewed on the date above and whenever the sharing changes. Either party may end it by "
    "giving {{Notice Period|30 days'}} written notice, after which the shared information will be returned or destroyed.",
    "## Signed",
    ("grid", ["", "{{Organisation}}", "{{Other Party}}"], [["Name", "", ""], ["Role", "", ""], ["Signature", "", ""], ["Date", "", ""]]),
    ("foot", "{{Reference}} - data sharing agreement - {{Organisation}} and {{Other Party}}"),
]
T[DP + "Processor due diligence questionnaire"] = [
    "# Supplier data protection questionnaire",
    "{{Organisation}} is considering using {{Supplier Name}} for {{Service Description}}. Because this would "
    "involve personal information, we need to understand how it would be protected. Please complete and return "
    "this by {{Reply By|today+2w}} to {{Contact Email}}.",
    ("grid", ["Question", "Answer"], [
        ["Company name, registered address and ICO registration number", ""],
        ["Who is responsible for data protection? Name and contact details", ""],
        ["What personal data would you process for us, and where would it be stored (country and provider)?", ""],
        ["Would any of it be transferred or accessible outside the UK? If so, where, and under what safeguard?", ""],
        ["Which sub-processors would be involved? Please list them and what each does", ""],
        ["What security certifications do you hold (ISO 27001, Cyber Essentials Plus, SOC 2)? Please attach certificates", ""],
        ["How is data encrypted, in transit and at rest?", ""],
        ["How is access controlled, and is multi-factor authentication enforced?", ""],
        ["How are your staff vetted and trained in data protection?", ""],
        ["When was your last penetration test, and by whom?", ""],
        ["How quickly would you tell us about a personal data breach?", ""],
        ["How would you help us respond to a subject access or erasure request?", ""],
        ["What happens to our data when the contract ends?", ""],
        ["Have you had a reportable breach or regulatory action in the last three years?", ""],
        ["Is our data used to train or improve any AI model?", ""],
    ]),
    "Completed by: ____________________     Role: ____________________     Date: ____________",
    ("foot", "Supplier data protection questionnaire - {{Supplier Name}}"),
]
T[DP + "Record of processing entry"] = [
    "# Record of processing activities - {{Processing Name}}",
    ("kv", [("Reference", "{{Reference}}"), ("Service area", "{{Service Area}}"), ("Owner", "{{Process Owner}}"),
            ("Our role", "{{Role=Controller|Joint controller|Processor}}"),
            ("Controller (if we are the processor)", "{{Controller Name|Not applicable}}"),
            ("Purpose of the processing", "{{Purpose+}}"),
            ("Lawful basis", "{{Lawful Basis=Consent|Contract|Legal obligation|Vital interests|Public task|Legitimate interests}}"),
            ("Special category condition", "{{Special Category Condition|Not applicable}}"),
            ("Whose data", "{{Data Subjects+}}"), ("Categories of data", "{{Data Categories+}}"),
            ("Where it comes from", "{{Source+}}"), ("Who it is shared with", "{{Recipients+}}"),
            ("Transfers outside the UK and the safeguard", "{{Transfers|None}}"),
            ("How long it is kept", "{{Retention Period}}"), ("Where it is held", "{{Systems+}}"),
            ("Security measures", "{{Security Measures+}}"),
            ("DPIA", "{{DPIA=Not required|Completed|Required - not yet done}}"),
            ("Last reviewed", "{{TODAY}}")]),
    ("foot", "ROPA - {{Reference}} - {{Processing Name}}"),
]
T[DP + "Privacy notice"] = [
    "# Privacy notice - {{Service Name}}",
    "Last updated {{TODAY}}",
    "## Who we are",
    "{{Organisation}} is the {{Role=controller|joint controller with|processor for}} {{Other Controller?}} for the "
    "information described here. You can contact our Data Protection Officer, {{DPO Name}}, at {{Contact Email}}.",
    "## What we collect", "{{Data Collected+}}",
    "## Why we collect it and our lawful basis", "{{Purposes And Basis+}}",
    "## Who we share it with", "{{Recipients+}}",
    "{{#if Transfers}}",
    "## Transfers outside the UK", "{{Transfers+}}",
    "{{/if}}",
    "## How long we keep it", "{{Retention+}}",
    "## Your rights",
    "{{clause:Rights summary}}",
    "## If you have a concern",
    "Please contact us first so that we can put it right. You also have the right to complain to the Information "
    "Commissioner's Office at ico.org.uk/make-a-complaint or on 0303 123 1113.",
    ("foot", "{{DOC_REF}}  |  Privacy notice - {{Service Name}}  |  Version {{VERSION}}"),
]
T[DP + "Data protection complaint response"] = [
    ("letter", "Your data protection complaint - {{Reference}}", "Complainant"),
    "Thank you for your complaint, which we received on {{Date Received}}, about how {{Organisation}} has handled "
    "your personal information.",
    "! What you told us",
    "{{Complaint Summary+}}",
    "! What we did to look into it",
    "{{Investigation+}}",
    "! What we found",
    "{{Outcome=We have upheld your complaint|We have partly upheld your complaint|We have not upheld your complaint}}.",
    "{{Findings+}}",
    "{{#if Remedy}}",
    "! What we are doing about it",
    "{{Remedy+}}",
    "{{/if}}",
    "If you are not satisfied with this response, you have the right to complain to the Information Commissioner's "
    "Office at ico.org.uk/make-a-complaint or on 0303 123 1113.",
    ("sign",), ("foot", "{{Reference}} - complaint response"),
]
T[DP + "Record of destruction"] = [
    "# Record of destruction",
    ("kv", [("Reference", "{{Reference}}"), ("Date destroyed", "{{Date Destroyed|today}}"),
            ("Service area", "{{Service Area}}"),
            ("What was destroyed", "{{Records Description+}}"),
            ("Date range of the records", "{{Record Dates}}"),
            ("Format", "{{Format=Paper|Electronic|Paper and electronic|Device or media}}"),
            ("Retention rule applied", "{{Retention Rule}}"),
            ("Method", "{{Method=Cross-cut shredding on site|Confidential waste contractor|Secure deletion|Device wipe|Physical destruction of media}}"),
            ("Contractor and certificate number", "{{Certificate|Not applicable}}"),
            ("Any legal hold, open request or complaint checked", "{{Holds Checked=Yes - none found|Yes - see note}}"),
            ("Carried out by", "{{Destroyed By}}"), ("Authorised by", "{{Authorised By}}")]),
    ("foot", "{{Reference}} - record of destruction"),
]

# -------------------------------------------------- information security
T[SEC + "Security incident report"] = [
    "# Information security incident report",
    ("kv", [("Incident reference", "{{Incident Reference}}"), ("Reported by", "{{Reported By}}"),
            ("Date and time reported", "{{Date Reported|today}}"), ("Date and time it happened", "{{Date Of Incident}}"),
            ("Type", "{{Incident Type=Phishing or social engineering|Malware|Lost or stolen device|Unauthorised access|Misdirected information|System outage|Physical security|Other}}"),
            ("Severity", "{{Severity=Low|Medium|High|Critical}}"),
            ("What happened", "{{What Happened+}}"), ("Systems, sites or information affected", "{{Affected+}}"),
            ("Personal data involved?", "{{Personal Data Involved=No|Yes - breach record opened|Not yet known}}"),
            ("Immediate actions", "{{Actions Taken+}}"), ("Root cause", "{{Root Cause+}}"),
            ("Corrective actions, owner and date", "{{Corrective Actions+}}"),
            ("Who has been told", "{{Notified+}}"), ("Lessons learned", "{{Lessons+}}"),
            ("Closed by and date", "{{Closed By}}")]),
    ("foot", "{{Incident Reference}} - security incident report"),
]
T[SEC + "Access request (starter, mover, leaver)"] = [
    "# Access request - {{Request Type=New starter|Change of role|Leaver}}",
    ("kv", [("Name", "{{Staff Name}}"), ("Job title", "{{Job Title}}"), ("Service area", "{{Service Area}}"),
            ("Line manager", "{{Line Manager}}"), ("Effective date", "{{Effective Date}}"),
            ("Requested by", "{{Handler Name}}"), ("Date requested", "{{TODAY}}")]),
    "## Access",
    ("grid", ["System or group", "Add / Remove / No change", "Level", "Done by and date"], [
        ["Microsoft 365 account and licence", "", "", ""], ["Email and shared mailboxes", "", "", ""],
        ["SharePoint and Teams", "", "", ""], ["Case management system", "", "", ""],
        ["Finance or HR system", "", "", ""], ["Multi-factor authentication", "", "", ""],
        ["Building access", "", "", ""], ["", "", "", ""]]),
    "## Equipment",
    ("grid", ["Item", "Asset number", "Issued / Returned", "Date"], [["Laptop", "", "", ""], ["Mobile phone", "", "", ""], ["", "", "", ""]]),
    "## Checks",
    "* Data protection and information security induction {{Induction=booked|completed|not applicable}}",
    "* Acceptable use policy signed {{Policy Signed=yes|no|not applicable}}",
    "* For a leaver: account disabled, mailbox handled, shared passwords changed, equipment returned {{Leaver Checks=not applicable|completed|outstanding}}",
    "Notes: {{Notes+}}",
    "Approved by: {{Approved By}}",
    ("foot", "Access request - {{Staff Name}} - {{TODAY_SHORT}}"),
]
T[SEC + "Asset disposal record"] = [
    "# IT asset disposal record",
    ("kv", [("Reference", "{{Reference}}"), ("Date", "{{Disposal Date|today}}"),
            ("Disposal route", "{{Route=Certified disposal contractor|Returned to lessor|Donated after wipe|Reissued after wipe|Destroyed on site}}"),
            ("Contractor", "{{Contractor|Not applicable}}"), ("Certificate number", "{{Certificate|Not applicable}}")]),
    ("grid", ["Asset number", "Type and model", "Serial number", "Held data?", "Wipe or destruction method", "Verified by"], 6),
    "The asset register has been updated: {{Register Updated=Yes|No}}",
    "Carried out by: {{Handler Name}}     Authorised by: {{Authorised By}}",
    ("foot", "{{Reference}} - asset disposal record"),
]
T[SEC + "Business continuity test report"] = [
    "# Business continuity test report",
    ("kv", [("Test reference", "{{Reference}}"), ("Date", "{{Test Date}}"),
            ("Type", "{{Test Type=Desktop walk-through|Call cascade|Backup restore|System failover|Site loss simulation}}"),
            ("Scenario", "{{Scenario+}}"), ("Objectives", "{{Objectives+}}"), ("Who took part", "{{Participants+}}")]),
    "## What happened", "{{Narrative+}}",
    "## Results against objectives",
    ("grid", ["Objective", "Target", "Achieved", "Met?"], 4),
    "## What worked", "{{Worked Well+}}",
    "## What did not", "{{Issues+}}",
    "## Actions",
    ("grid", ["Action", "Owner", "Due", "Status"], 4),
    "Overall result: {{Result=Passed|Passed with actions|Failed}}     Next test due: {{Next Test|today+1y}}",
    "Report by: {{Handler Name}}, {{TODAY}}",
    ("foot", "{{Reference}} - business continuity test report"),
]
T[SEC + "Supplier security review"] = [
    "# Supplier security review - {{Supplier Name}}",
    ("kv", [("Service provided", "{{Service Description+}}"), ("Information they hold or can reach", "{{Data Access+}}"),
            ("Criticality", "{{Criticality=Low|Medium|High}}"), ("Contract owner", "{{Contract Owner}}"),
            ("Reviewed by", "{{Handler Name}}"), ("Date", "{{TODAY}}"), ("Next review", "{{Next Review|today+1y}}")]),
    ("grid", ["Check", "Evidence seen", "Result"], [
        ["Current ISO 27001 / Cyber Essentials Plus certificate", "", ""], ["Data processing agreement in place and current", "", ""],
        ["Hosting location and sub-processors unchanged", "", ""], ["Incidents or breaches since last review", "", ""],
        ["Penetration test or audit report in the last 12 months", "", ""], ["Service levels met", "", ""],
        ["Access by their staff to our systems still appropriate", "", ""], ["Exit and data return arrangements", "", ""]]),
    "Findings: {{Findings+}}",
    "Outcome: {{Outcome=Approved|Approved with actions|Not approved}}",
    "Actions: {{Actions+}}",
    ("foot", "Supplier security review - {{Supplier Name}} - {{TODAY_SHORT}}"),
]

# ------------------------------------------------------ quality and audit
T[QA + "Internal audit plan"] = [
    "# Internal audit plan - {{Audit Title}}",
    ("kv", [("Audit reference", "{{Reference}}"), ("Standard", "{{Standard=ISO 9001|ISO 27001|ISO 9001 and ISO 27001|Contract requirements|Internal procedure}}"),
            ("Area or process", "{{Audit Scope+}}"), ("Objective", "{{Objective+}}"),
            ("Criteria (clauses, controls, procedures)", "{{Criteria+}}"),
            ("Auditor", "{{Auditor}}"), ("Auditee", "{{Auditee}}"), ("Date and time", "{{Audit Date}}"),
            ("Method", "{{Method=Interview and document review|Remote review|Observation|Sampling of records}}")]),
    "## Schedule",
    ("grid", ["Time", "Activity", "Who"], [["", "Opening meeting", ""], ["", "", ""], ["", "", ""], ["", "Closing meeting", ""]]),
    "## Documents and records to have ready", "{{Documents Needed+}}",
    ("foot", "{{Reference}} - audit plan"),
]
T[QA + "Internal audit report"] = [
    "# Internal audit report - {{Audit Title}}",
    ("kv", [("Audit reference", "{{Reference}}"), ("Standard", "{{Standard=ISO 9001|ISO 27001|ISO 9001 and ISO 27001|Contract requirements|Internal procedure}}"),
            ("Area or process", "{{Audit Scope+}}"), ("Criteria", "{{Criteria+}}"), ("Auditor", "{{Auditor}}"),
            ("Auditee", "{{Auditee}}"), ("Date of audit", "{{Audit Date}}"), ("Date of report", "{{TODAY}}")]),
    "## Summary", "{{Summary+}}",
    "## What was looked at", "{{Evidence Reviewed+}}",
    "## Good practice", "{{Good Practice+}}",
    "## Findings",
    ("grid", ["No.", "Finding", "Clause or control", "Grade", "Evidence"], 5),
    "Grades: major nonconformity, minor nonconformity, observation, opportunity for improvement.",
    "## Conclusion",
    "{{Conclusion=The area audited conforms to the criteria|The area audited conforms, with the minor findings above|The area audited does not fully conform}}.",
    "{{Conclusion Detail+}}",
    "Corrective actions are to be agreed with the auditee by {{Actions Due|today+2w}} and recorded on the corrective action log.",
    ("foot", "{{Reference}} - internal audit report"),
]
T[QA + "Nonconformity and corrective action"] = [
    "# Nonconformity and corrective action report",
    ("kv", [("Reference", "{{Reference}}"), ("Raised by", "{{Raised By}}"), ("Date raised", "{{Date Raised|today}}"),
            ("Source", "{{Source=Internal audit|External audit|Complaint|Incident|Monitoring|Staff suggestion|Supplier}}"),
            ("Grade", "{{Grade=Major|Minor|Observation}}"), ("Clause, control or procedure", "{{Requirement}}"),
            ("What was found", "{{Description+}}"), ("Evidence", "{{Evidence+}}"),
            ("Immediate correction", "{{Correction+}}"), ("Root cause", "{{Root Cause+}}"),
            ("Could it exist elsewhere?", "{{Extent+}}"),
            ("Corrective action", "{{Corrective Action+}}"), ("Owner", "{{Action Owner}}"), ("Due", "{{Due Date|today+1m}}"),
            ("Date completed", "{{Date Completed}}"), ("Check that it worked (how, when, by whom)", "{{Effectiveness Review+}}"),
            ("Closed by and date", "{{Closed By}}")]),
    ("foot", "{{Reference}} - corrective action report"),
]
T[QA + "Management review minutes"] = [
    "# Management review - {{Review Period}}",
    ("kv", [("Date", "{{Meeting Date}}"), ("Chair", "{{Chair}}"), ("Present", "{{Attendees+}}"), ("Apologies", "{{Apologies}}"),
            ("Systems covered", "{{Scope=Quality and information security management systems|Quality management system|Information security management system}}")]),
    "## 1. Actions from the last review", "{{Previous Actions+}}",
    "## 2. Changes that affect the management system",
    "Internal and external issues, and the needs of interested parties: {{Context Changes+}}",
    "## 3. Performance",
    "Customer and stakeholder feedback, including complaints: {{Feedback+}}",
    "Objectives - progress: {{Objectives Progress+}}",
    "Process performance and service delivery against contract: {{Process Performance+}}",
    "Nonconformities and corrective actions: {{Nonconformities+}}",
    "Monitoring and measurement results: {{Monitoring Results+}}",
    "Audit results, internal and external: {{Audit Results+}}",
    "Supplier performance: {{Supplier Performance+}}",
    "Information security incidents and personal data breaches: {{Incidents+}}",
    "## 4. Resources", "{{Resources+}}",
    "## 5. Risks and opportunities",
    "Risk assessment results, the risk treatment plan, and how well earlier actions worked: {{Risks+}}",
    "## 6. Opportunities for improvement", "{{Improvements+}}",
    "## 7. Decisions and actions",
    ("grid", ["Decision or action", "Owner", "Due"], 5),
    "Conclusion on whether the management system remains suitable, adequate and effective: {{Conclusion+}}",
    "Next review: {{Next Review|today+1y}}",
    ("foot", "Management review - {{Review Period}} - minutes"),
]
T[QA + "Risk assessment"] = [
    "# Risk assessment - {{Assessment Title}}",
    ("kv", [("Reference", "{{Reference}}"), ("Scope", "{{Risk Scope+}}"), ("Assessed by", "{{Handler Name}}"),
            ("Date", "{{TODAY}}"), ("Review date", "{{Review Date|today+1y}}")]),
    "Scoring: likelihood 1 (rare) to 5 (almost certain) multiplied by impact 1 (negligible) to 5 (severe). "
    "{{Risk Appetite|Scores of 12 or more need a treatment plan agreed by the senior management team.}}",
    ("grid", ["Risk", "Cause and consequence", "Existing controls", "L", "I", "Score", "Treatment", "Owner", "Target date"], 6),
    "Summary and recommendation: {{Summary+}}",
    "Accepted by: {{Accepted By}}",
    ("foot", "{{Reference}} - risk assessment - {{Assessment Title}}"),
]
T[QA + "Change request"] = [
    "# Change request",
    ("kv", [("Reference", "{{Reference}}"), ("Requested by", "{{Handler Name}}"), ("Date", "{{TODAY}}"),
            ("What is changing", "{{Change Description+}}"), ("Why", "{{Reason+}}"),
            ("Type", "{{Change Type=System or software|Process or procedure|Supplier|Premises|Organisational}}"),
            ("Who and what is affected", "{{Impact+}}"),
            ("Information security and data protection impact", "{{Security Impact+}}"),
            ("DPIA screening needed?", "{{DPIA Needed=No|Yes - completed|Yes - outstanding}}"),
            ("Risks and how they are managed", "{{Risks+}}"), ("Testing", "{{Testing+}}"),
            ("How to back out", "{{Rollback+}}"), ("Planned date", "{{Planned Date}}"),
            ("Documents and training to update", "{{Documents To Update+}}"),
            ("Decision", "{{Decision=Approved|Approved with conditions|Rejected|Deferred}}"),
            ("Approved by and date", "{{Approved By}}"), ("Post-change review", "{{Post Review+}}")]),
    ("foot", "{{Reference}} - change request"),
]
T[QA + "Supplier evaluation"] = [
    "# Supplier evaluation - {{Supplier Name}}",
    ("kv", [("Goods or service", "{{Service Description+}}"), ("Evaluation type", "{{Evaluation Type=New supplier|Annual re-evaluation|After a problem}}"),
            ("Evaluated by", "{{Handler Name}}"), ("Date", "{{TODAY}}")]),
    ("grid", ["Criterion", "Score 1-5", "Comments"], [
        ["Quality of goods or service", "", ""], ["Delivery and reliability", "", ""], ["Price and value", "", ""],
        ["Responsiveness and communication", "", ""], ["Certifications and compliance", "", ""],
        ["Information security and data protection", "", ""], ["Financial stability and insurance", "", ""],
        ["Social value and sustainability", "", ""]]),
    "Overall: {{Outcome=Approved|Approved with conditions|Not approved}}",
    "Comments and conditions: {{Comments+}}",
    "Next evaluation: {{Next Review|today+1y}}",
    ("foot", "Supplier evaluation - {{Supplier Name}} - {{TODAY_SHORT}}"),
]

# ------------------------------------------------------------- governance
T[GOV + "Board or SMT paper"] = [
    "# {{Paper Title}}",
    ("kv", [("Meeting", "{{Meeting Name}}"), ("Date", "{{Meeting Date}}"), ("Author", "{{Handler Name}}, {{Handler Title}}"),
            ("Paper for", "{{Paper For=Decision|Discussion|Information|Assurance}}"),
            ("Classification", "{{Classification=Internal|Confidential|Public}}")]),
    "## 1. Purpose", "{{Purpose+}}",
    "## 2. Recommendation", "{{Recommendation+}}",
    "## 3. Background", "{{Background+}}",
    "## 4. Options considered", "{{Options+}}",
    "## 5. Implications",
    ("labelled", [("Financial", "{{Financial Implications+}}"), ("Legal and contractual", "{{Legal Implications+}}"),
                  ("Data protection and information security", "{{Data Implications+}}"),
                  ("People and equality", "{{People Implications+}}"), ("Risk", "{{Risk Implications+}}")]),
    "## 6. Next steps", "{{Next Steps+}}",
    ("foot", "{{Meeting Name}} - {{Meeting Date}} - {{Paper Title}}"),
]
T[GOV + "Meeting agenda"] = [
    "# {{Meeting Name}} - agenda",
    ("kv", [("Date and time", "{{Meeting Date}}"), ("Where", "{{Location|Microsoft Teams}}"), ("Chair", "{{Chair}}"),
            ("Invited", "{{Attendees+}}")]),
    ("grid", ["No.", "Item", "Lead", "Paper", "Time"], [
        ["1", "Welcome, apologies and declarations of interest", "", "", ""],
        ["2", "Minutes of the last meeting and matters arising", "", "", ""],
        ["3", "Action log", "", "", ""], ["4", "", "", "", ""], ["5", "", "", "", ""], ["6", "", "", "", ""],
        ["7", "Any other business", "", "", ""], ["8", "Date of next meeting", "", "", ""]]),
    "Papers are {{Papers Location|attached}}. Please send apologies to {{Contact Email}}.",
    ("foot", "{{Meeting Name}} - {{Meeting Date}} - agenda"),
]
T[GOV + "Meeting minutes"] = [
    "# {{Meeting Name}} - minutes",
    ("kv", [("Date and time", "{{Meeting Date}}"), ("Where", "{{Location|Microsoft Teams}}"), ("Chair", "{{Chair}}"),
            ("Present", "{{Attendees+}}"), ("Apologies", "{{Apologies}}"), ("Minutes taken by", "{{Handler Name}}")]),
    "## 1. Welcome, apologies and declarations of interest", "{{Declarations+|No interests were declared.}}",
    "## 2. Minutes of the last meeting", "{{Previous Minutes+|The minutes of the last meeting were agreed as a true record.}}",
    "## 3. Matters arising and action log", "{{Matters Arising+}}",
    "## 4. Items discussed", "{{Discussion+}}",
    "## 5. Decisions", "{{Decisions+}}",
    "## 6. Any other business", "{{Other Business+|None.}}",
    "## Actions",
    ("grid", ["No.", "Action", "Owner", "Due", "Status"], 6),
    "Next meeting: {{Next Meeting}}",
    ("foot", "{{Meeting Name}} - {{Meeting Date}} - minutes"),
]
T[GOV + "Briefing note"] = [
    "# Briefing note - {{Subject}}",
    ("kv", [("For", "{{Briefing For}}"), ("From", "{{Handler Name}}, {{Handler Title}}"), ("Date", "{{TODAY}}")]),
    "## Summary", "{{Summary+}}",
    "## Background", "{{Background+}}",
    "## Key points", "{{Key Points+}}",
    "## What this means for us", "{{Implications+}}",
    "## Recommended action", "{{Recommendation+}}",
    "## Sources", "{{Sources+}}",
    ("foot", "Briefing note - {{Subject}} - {{TODAY_SHORT}}"),
]
T[GOV + "Terms of reference"] = [
    "# {{Group Name}} - terms of reference",
    ("ctl",),
    "## 1. Purpose", "{{Purpose+}}",
    "## 2. Responsibilities", "{{Responsibilities+}}",
    "## 3. Authority", "{{Authority+}}",
    "## 4. Membership",
    "Chair: {{Chair}}", "Members: {{Members+}}", "In attendance: {{In Attendance+}}",
    "## 5. Quorum", "{{Quorum}}",
    "## 6. Meetings", "The group meets {{Frequency=monthly|every two months|quarterly|twice a year|annually}}. {{Meeting Arrangements+}}",
    "## 7. Reporting", "{{Reporting+}}",
    "## 8. Review", "These terms of reference are reviewed {{Review Cycle|annually}}.",
    ("foot", "{{DOC_REF}}  |  {{Group Name}} terms of reference  |  Version {{VERSION}}  |  Review: {{REVIEW_DATE}}"),
]

# ---------------------------------------------------------------- policies
T[POL + "Policy skeleton"] = [
    ("title", "{{DOC_TITLE}}"), ("ctl",),
    "# 1. Purpose", "{{Purpose+}}",
    "# 2. Scope", "{{Scope+|This policy applies to all staff, volunteers and contractors.}}",
    "# 3. Policy statement", "{{Policy Statement+}}",
    "# 4. Responsibilities", "{{Responsibilities+}}",
    "# 5. Procedure", "{{Procedure+}}",
    "# 6. Related documents", "{{Related Documents+}}",
    "# 7. Monitoring and review",
    "This policy is reviewed at least every {{Review Cycle|12 months}}, and sooner if the law, our contracts or our practice changes.",
    ("foot", "{{DOC_REF}}  |  {{DOC_TITLE}}  |  Version {{VERSION}}  |  Owner: {{OWNER}}  |  Review: {{REVIEW_DATE}}"),
]
T[POL + "Procedure skeleton"] = [
    ("title", "{{DOC_TITLE}}"), ("ctl",),
    "# 1. Purpose", "{{Purpose+}}",
    "# 2. Scope", "{{Scope+}}",
    "# 3. Who does what",
    ("grid", ["Role", "Responsibility"], 4),
    "# 4. Procedure",
    ("grid", ["Step", "What happens", "Who", "Record kept"], 6),
    "# 5. What to do when something goes wrong", "{{Exceptions+}}",
    "# 6. Records",
    ("grid", ["Record", "Where it is kept", "How long"], 3),
    "# 7. Related documents", "{{Related Documents+}}",
    "# 8. Change history",
    ("grid", ["Version", "Date", "Change", "Approved by"], [["{{VERSION}}", "{{ISSUE_DATE}}", "{{Change Summary|First issue}}", "{{Approved By}}"]]),
    ("foot", "{{DOC_REF}}  |  {{DOC_TITLE}}  |  Version {{VERSION}}  |  Owner: {{OWNER}}  |  Review: {{REVIEW_DATE}}"),
]

# -------------------------------------------------------------- complaints
T[COMP + "Complaint acknowledgement"] = [
    ("letter", "Your complaint - {{Reference}}", "Complainant"),
    "Thank you for getting in touch on {{Date Received}}. I am sorry that you have had cause to complain.",
    "I understand your complaint to be: {{Complaint Summary+}}",
    "If I have misunderstood, please let me know.",
    "{{Investigator Name|I}} will look into this and we will write to you with the outcome by {{Response Due|today+2w}}. "
    "If we need longer, we will tell you why and when you can expect to hear.",
    "If you would like to add anything, or need this in a different format, please contact me using the details below.",
    ("sign",), ("foot", "{{Reference}} - complaint acknowledgement"),
]
T[COMP + "Complaint outcome"] = [
    ("letter", "Outcome of your complaint - {{Reference}}", "Complainant"),
    "I am writing with the outcome of the complaint you made on {{Date Received}}.",
    "! Your complaint",
    "{{Complaint Summary+}}",
    "! How we looked into it",
    "{{Investigation+}}",
    "! What we found",
    "{{Outcome=Your complaint is upheld|Your complaint is partly upheld|Your complaint is not upheld}}.",
    "{{Findings+}}",
    "{{#if Remedy}}",
    "! What we are doing to put things right",
    "{{Remedy+}}",
    "{{/if}}",
    "! If you are not satisfied",
    "{{Escalation+|You can ask for this decision to be reviewed by writing to us within 20 working days, saying why you are not satisfied.}}",
    ("sign",), ("foot", "{{Reference}} - complaint outcome"),
]

# --------------------------------------------------------- letters and notes
T[LET + "General letter"] = [
    ("letter", "{{Subject}}", "Recipient"),
    "{{Body+}}",
    ("sign",),
]
T[LET + "Memo"] = [
    "# Memo",
    ("kv", [("To", "{{Memo To}}"), ("From", "{{Handler Name}}, {{Handler Title}}"), ("Date", "{{TODAY}}"), ("Subject", "{{Subject}}")]),
    "", "{{Body+}}",
    "{{#if Action Required}}",
    "! Action required",
    "{{Action Required+}}",
    "{{/if}}",
]
T[LET + "File note"] = [
    "# File note",
    ("kv", [("Reference", "{{Reference}}"), ("Date and time", "{{Note Date|today}}"), ("Written by", "{{Handler Name}}"),
            ("Type", "{{Note Type=Telephone call|Meeting|Email exchange|Decision|Observation}}"),
            ("With", "{{Note With}}"), ("About", "{{Subject}}")]),
    "## What was said or done", "{{Note Body+}}",
    "## What was agreed", "{{Agreed+}}",
    "## Follow-up", "{{Follow Up+}}",
    ("foot", "{{Reference}} - file note - {{TODAY_SHORT}}"),
]
T[LET + "Handover note"] = [
    "# Handover - {{Handover Subject}}",
    ("kv", [("From", "{{Handler Name}}"), ("To", "{{Handover To}}"), ("Date", "{{TODAY}}"), ("Covers", "{{Handover Period}}")]),
    "## What this covers", "{{Overview+}}",
    "## In progress",
    ("grid", ["Task", "Where it is up to", "Next step", "Deadline"], 5),
    "## Regular tasks",
    ("grid", ["Task", "How often", "How to do it or where the procedure is"], 4),
    "## Where things are", "{{Locations+}}",
    "## Who to ask", "{{Contacts+}}",
    "## Things to watch", "{{Watch Points+}}",
    ("foot", "Handover - {{Handover Subject}} - {{TODAY_SHORT}}"),
]
T[LET + "Policy acknowledgement form"] = [
    "# Policy acknowledgement",
    "I confirm that I have read and understood the following, and that I will follow {{Them=it|them}}:",
    "{{Policies Listed+}}",
    "I understand that I can ask {{Handler Name}} if anything is unclear, and that not following "
    "{{Them=it|them}} may be dealt with under the disciplinary procedure.",
    ("kv", [("Name", "{{Staff Name}}"), ("Job title", "{{Job Title}}"), ("Signature", ""), ("Date", "")]),
    ("foot", "Policy acknowledgement - {{Staff Name}}"),
]

# -------------------------------------------------------------- participants
T[PART + "Consent form"] = [
    "# {{Service Name}} - your consent",
    "{{Organisation}} would like to {{What We Want To Do+}}. We need your agreement to do this. Please read each "
    "statement and tick the ones you agree to. You do not have to agree to any of them, and saying no will not "
    "affect the support you receive.",
    ("grid", ["I agree to...", "Yes", "No"], [["{{Consent Statement 1}}", "", ""], ["{{Consent Statement 2?}}", "", ""], ["{{Consent Statement 3?}}", "", ""]]),
    "! What you should know",
    "* We will use your information only for what is described above.",
    "* We will keep it for {{Retention Period}} and then securely destroy it.",
    "* We will share it only with {{Recipients|the organisations named above}}.",
    "* You can change your mind at any time by contacting {{Handler Name}} at {{Contact Email}} or on {{Contact Phone}}. "
    "It is as easy to withdraw your consent as it is to give it.",
    "{{clause:Rights summary}}",
    ("kv", [("Your name", "{{Participant Name}}"), ("Signature", ""), ("Date", "")]),
    ("foot", "{{Service Name}} - consent form - version {{VERSION}}"),
]
T[PART + "Consent withdrawal confirmation"] = [
    ("letter", "Confirmation that you have withdrawn your consent", "Recipient"),
    "Thank you for letting us know on {{Date Received}} that you no longer agree to: {{What Was Consented To+}}",
    "We have updated our records and stopped from {{Effective Date|today}}. Withdrawing your consent does not affect "
    "anything we did before that date.",
    "{{#if What Continues}}",
    "We will continue to: {{What Continues+}}",
    "This is because we do not rely on your consent for it: {{Why It Continues+}}",
    "{{/if}}",
    "You can give your consent again at any time by contacting me.",
    ("sign",),
]


# ------------------------------------------------------------- the builder
def _new():
    import docx
    from docx.shared import Cm, Pt
    d = docx.Document()
    style = d.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(11)
    for s in d.sections:
        s.left_margin = s.right_margin = Cm(2.5)
        s.top_margin = s.bottom_margin = Cm(2.2)
    return d


def _table(d, rows, header=False, label_column=False):
    cols = max(len(r) for r in rows)
    table = d.add_table(rows=len(rows), cols=cols)
    table.style = "Table Grid"
    for i, row in enumerate(rows):
        for j in range(cols):
            cell = table.cell(i, j)
            cell.text = row[j] if j < len(row) else ""
            if cell.text and ((header and i == 0) or (label_column and j == 0)):
                for run in cell.paragraphs[0].runs:
                    run.bold = True
    return table


def make(spec):
    """A Document from the shorthand described at the top of this file."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt
    d = _new()
    for item in spec:
        if isinstance(item, str):
            if item.startswith("## "):
                d.add_heading(item[3:], level=2)
            elif item.startswith("# "):
                d.add_heading(item[2:], level=1)
            elif item.startswith("! "):
                d.add_paragraph().add_run(item[2:]).bold = True
            elif item.startswith("* "):
                d.add_paragraph(item[2:], style="List Bullet")
            else:
                d.add_paragraph(item)
            continue
        kind = item[0]
        if kind == "letter":
            who = item[2]
            p = d.add_paragraph("{{Organisation}}")
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.runs[0].bold = True
            d.add_paragraph("{{Organisation Address+}}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
            d.add_paragraph("{{%s Name}}" % who)
            d.add_paragraph("{{%s Address+}}" % who)
            d.add_paragraph("{{TODAY}}")
            if "{{Reference}}" in item[1]:
                d.add_paragraph("Our reference: {{Reference}}")
            d.add_paragraph("Dear {{Salutation|Sir or Madam}}")
            d.add_paragraph().add_run(item[1]).bold = True
        elif kind == "sign":
            d.add_paragraph("Yours sincerely")
            d.add_paragraph()
            for line in ("{{Handler Name}}", "{{Handler Title}}", "{{Contact Email}}"):
                d.add_paragraph(line)
        elif kind == "title":
            d.add_heading(item[1], level=0)
        elif kind == "kv":
            _table(d, [list(r) for r in item[1]], label_column=True)
            d.add_paragraph()
        elif kind == "grid":
            rows = item[2]
            body = [[""] * len(item[1]) for _ in range(rows)] if isinstance(rows, int) else [list(r) for r in rows]
            _table(d, [list(item[1])] + body, header=True)
            d.add_paragraph()
        elif kind == "labelled":
            for label, text in item[1]:
                p = d.add_paragraph()
                p.add_run(label + ": ").bold = True
                p.add_run(text)
        elif kind == "ctl":
            _table(d, [["Document reference", "{{DOC_REF}}"], ["Version", "{{VERSION}}"], ["Owner", "{{OWNER}}"],
                       ["Issue date", "{{ISSUE_DATE}}"], ["Next review", "{{REVIEW_DATE}}"]], label_column=True)
            d.add_paragraph()
        elif kind == "foot":
            run = d.sections[0].footer.paragraphs[0].add_run(item[1])
            run.font.size = Pt(8)
        else:
            raise ValueError(f"unknown block {kind!r}")
    return d


def starter_files() -> dict:
    """relative path (with / separators) -> a function that writes it to a path."""
    out = {}
    for name, spec in T.items():
        out[f"{TEMPLATES}/{name}.docx"] = (lambda path, spec=spec: make(spec).save(path))
    out[f"{CLAUSES}/Supplementary information.docx"] = lambda path: make(SUPPLEMENTARY).save(path)
    for rel, text in CLAUSE_TEXT.items():
        out[f"{CLAUSES}/{rel}"] = (lambda path, text=text: _write(path, text + "\n"))
    out[STANDING] = lambda path: _write(path, STANDING_TEXT)
    return out


def _write(path, text):
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(text)


# v1.0 kept these two elsewhere; treat the old copies as the same starter so
# an existing library is not given a second one.
MOVED = {f"{TEMPLATES}/{LET}General letter.docx": f"{TEMPLATES}/Letters/General letter.docx"}


def is_ours(root: str) -> bool:
    """True for a folder Draftsman set up (it has our manifest or our READ ME)."""
    if os.path.isfile(os.path.join(root, MANIFEST)):
        return True
    try:
        with open(os.path.join(root, README_NAME), encoding="utf-8") as fh:
            return fh.read(40).startswith("DRAFTSMAN LIBRARY")
    except OSError:
        return False


def build(root: str) -> list:
    """Write the starter files a library has not been given before. Never
    overwrites, and never brings back a starter that was deleted. Returns the
    paths written."""
    os.makedirs(root, exist_ok=True)
    manifest_path = os.path.join(root, MANIFEST)
    try:
        with open(manifest_path, encoding="utf-8") as fh:
            given = {ln.strip() for ln in fh if ln.strip()}
    except OSError:
        given = set()
    made = []
    for rel, writer in starter_files().items():
        if rel in given:
            continue
        path = os.path.join(root, *rel.split("/"))
        old = MOVED.get(rel)
        if not os.path.exists(path) and not (old and os.path.exists(os.path.join(root, *old.split("/")))):
            os.makedirs(os.path.dirname(path), exist_ok=True)
            writer(path)
            made.append(path)
        given.add(rel)
    _write(os.path.join(root, README_NAME), README)
    _write(manifest_path, "\n".join(sorted(given)) + "\n")
    return made
