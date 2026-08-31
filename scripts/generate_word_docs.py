"""
Generate Word (.docx) documents for Xero and CBA applications.
Usage: python scripts/generate_word_docs.py
"""

import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BLUE   = RGBColor(0, 114, 178)
GREY   = RGBColor(90, 90, 90)
BLACK  = RGBColor(0, 0, 0)
CONTACT = "vrandesh@gmail.com  |  0468 450 792  |  linkedin.com/in/vrandesh"

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def set_margins(doc, top=0.75, bottom=0.75, left=0.75, right=0.75):
    for section in doc.sections:
        section.top_margin    = Inches(top)
        section.bottom_margin = Inches(bottom)
        section.left_margin   = Inches(left)
        section.right_margin  = Inches(right)


def add_rule(doc, color="0072B2"):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(2)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"),   "single")
    bottom.set(qn("w:sz"),    "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


def add_name_header(doc, name):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(3)
    r = p.add_run(name)
    r.bold = True
    r.font.size = Pt(22)
    r.font.color.rgb = BLUE

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after  = Pt(6)
    r2 = p2.add_run(CONTACT)
    r2.font.size = Pt(9)
    r2.font.color.rgb = GREY


def add_section(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after  = Pt(1)
    r = p.add_run(title.upper())
    r.bold = True
    r.font.size = Pt(10.5)
    r.font.color.rgb = BLUE
    add_rule(doc)


def body(doc, text, before=1, after=3, size=10):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after  = Pt(after)
    r = p.add_run(text)
    r.font.size = Pt(size)
    return p


def competency(doc, label, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(2)
    p.paragraph_format.left_indent  = Inches(0.15)
    rl = p.add_run(label + ": ")
    rl.bold = True
    rl.font.size = Pt(10)
    rt = p.add_run(text)
    rt.font.size = Pt(10)


def role(doc, dates, title, company, location, desc, bullets):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(7)
    p.paragraph_format.space_after  = Pt(1)
    rt = p.add_run(title)
    rt.bold = True
    rt.font.size = Pt(10.5)
    p.add_run("  —  ").font.size = Pt(10.5)
    rc = p.add_run(company)
    rc.bold = True
    rc.font.size = Pt(10.5)
    rc.font.color.rgb = BLUE
    meta = p.add_run(f"     {location}  |  {dates}")
    meta.font.size = Pt(9)
    meta.font.color.rgb = GREY
    meta.italic = True

    if desc:
        pd = doc.add_paragraph()
        pd.paragraph_format.space_before = Pt(1)
        pd.paragraph_format.space_after  = Pt(2)
        pd.paragraph_format.left_indent  = Inches(0.0)
        pd.add_run(desc).font.size = Pt(10)

    for b in bullets:
        pb = doc.add_paragraph(style="List Bullet")
        pb.paragraph_format.space_before     = Pt(1)
        pb.paragraph_format.space_after      = Pt(1)
        pb.paragraph_format.left_indent      = Inches(0.25)
        pb.paragraph_format.first_line_indent = Inches(-0.15)
        pb.add_run(b).font.size = Pt(10)


def education_entry(doc, dates, degree, institution, location, desc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after  = Pt(1)
    rd = p.add_run(degree)
    rd.bold = True
    rd.font.size = Pt(10.5)
    p.add_run("  —  ").font.size = Pt(10.5)
    ri = p.add_run(institution)
    ri.bold = True
    ri.font.size = Pt(10.5)
    ri.font.color.rgb = BLUE
    meta = p.add_run(f"     {location}  |  {dates}")
    meta.font.size = Pt(9)
    meta.font.color.rgb = GREY
    meta.italic = True
    if desc:
        body(doc, desc, before=1, after=3)


def simple_bullet(doc, text):
    pb = doc.add_paragraph(style="List Bullet")
    pb.paragraph_format.space_before     = Pt(1)
    pb.paragraph_format.space_after      = Pt(1)
    pb.paragraph_format.left_indent      = Inches(0.25)
    pb.paragraph_format.first_line_indent = Inches(-0.15)
    pb.add_run(text).font.size = Pt(10)


# ---------------------------------------------------------------------------
# Xero CV
# ---------------------------------------------------------------------------

def build_xero_cv():
    doc = Document()
    set_margins(doc)

    add_name_header(doc, "Vrandesh Bandikatti")

    # Profile
    add_section(doc, "Profile")
    body(doc,
         "Senior Full Stack Engineer with 17+ years building enterprise platforms, specialising in "
         "production-grade agentic AI systems. Delivered multi-agent pipelines, RAG architectures and "
         "MCP-integrated workflows using Azure OpenAI, LangGraph and LangChain — embedded into full-stack "
         "enterprise products, not deployed as standalone prototypes. Designs for auditability: separates AI "
         "extraction from deterministic rule execution, builds evaluation and monitoring frameworks, and "
         "implements human-in-the-loop controls that production teams trust. Cloud-agnostic across Azure, "
         "AWS and GCP; strong delivery record in financial services and SaaS.")

    # Core Competencies
    add_section(doc, "Core Competencies")
    competency(doc, "AI Agents & Workflow Automation",
               "Azure OpenAI, Azure AI Foundry, AWS Bedrock, LangGraph, LangChain, MCP integrations, "
               "Copilot Studio, Power Automate; RAG pipelines, structured outputs, function calling, agent "
               "orchestration, evaluation frameworks and human-in-the-loop controls integrated into "
               "full-stack production platforms.")
    competency(doc, "Enterprise Platform Integration",
               "REST APIs, GraphQL, webhooks and SDKs; Xero, MYOB, Workday Adaptive Planning and ADF "
               "integrations; real-time financial data movement and enterprise connectivity across Azure, "
               "AWS and GCP with Salesforce, Slack and NetSuite-class systems.")
    competency(doc, "Full Stack Engineering",
               "Python, FastAPI, Django, TypeScript, React, Vue.js, Node.js; microservices, distributed "
               "systems and event-driven architecture; PostgreSQL, MongoDB, SQL Server, Firestore; "
               "Docker, CI/CD and DevSecOps.")
    competency(doc, "Evaluation & Observability",
               "Grafana dashboards, structured logging and AI workflow monitoring; bounded self-repair and "
               "graceful degradation patterns; deterministic rule separation, row-level security and "
               "full audit trails.")
    competency(doc, "Engineering Leadership",
               "Discovery through production delivery; cross-functional teams of 4–19; solution "
               "architecture, coaching and executive stakeholder engagement while remaining hands-on "
               "in design, coding, troubleshooting and release.")

    # Professional Experience
    add_section(doc, "Professional Experience")

    role(doc, "Mar 2022 – Present",
         "Senior Full Stack Developer & Cloud/AI Solution Architect",
         "Findex Australia", "Melbourne, VIC",
         "Architect and hands-on engineer across Findex Data Science platforms, owning delivery from "
         "discovery through full-stack implementation, cloud deployment, UAT and production support "
         "using Python/Django, FastAPI, GraphQL, Vue/React, Azure and Microsoft Fabric.",
         [
             "Led cross-functional teams of 4–19 across Education Portal, PerformX and KRM/STI "
             "initiatives; partnered with executives to manage scope, priorities, delivery risks and "
             "stakeholder expectations.",
             "Built PerformX Agent: conversational assistant that answers plain-language questions against "
             "a governed Power BI model via React and FastAPI, with triple permission enforcement and a "
             "bounded self-repair loop that rewrites failed queries or exits gracefully rather than "
             "returning incorrect answers. Deployed in browser and Microsoft Teams; developed "
             "end-to-end using Claude Code.",
             "Built Payroll AI: four-agent pipeline that reads Modern Award PDFs, extracts rules with "
             "clause references, generates client-specific calculation code, runs it against live payroll "
             "data and reports results into Power BI with full audit trails. Registered as an eligible "
             "R&D activity under the DISR R&D Tax Incentive.",
             "Architected KRM Remuneration Portal: full-stack STI, referral and audit platform using "
             "Python/Django, GraphQL and Vue/React with Entra ID SSO/RBAC, row-level security and "
             "governed Microsoft Fabric foundations. Onboarded 443 users and 129 people leaders.",
             "Designed EduPort and WealthAI TMD Processing; established CI/CD, Grafana observability and "
             "AI workflow evaluation frameworks across all platforms; stepped in as interim DevOps lead "
             "during a critical production incident.",
         ])

    role(doc, "Oct 2017 – Mar 2022",
         "Senior Software Engineer, Product Innovation",
         "Quest Payment Systems", "Melbourne, VIC",
         "Regulated payments engineering. Led backend and cloud delivery from concept to production "
         "for Donation Point Go, a contactless donation channel. Built 12+ microservices on AWS and GCP.",
         [
             "Designed and delivered end-to-end serverless architecture on AWS (Lambda, Step Functions, S3) "
             "and GCP (Firebase, Firestore) for Donation Point Go; built for high availability and "
             "unpredictable charitable-event load.",
             "Built Xero and MYOB integrations for real-time financial data movement and automated "
             "reconciliation — hands-on experience with the same accounting integration layer that "
             "Xero's AI platform extends.",
             "Led a team of three across Airpay mobile POS; designed high-availability architectures on "
             "Node.js, Django and Ember.js including Elasticsearch and serverless APIs.",
         ])

    role(doc, "Jun 2013 – Jul 2017",
         "Technology Lead / Full Stack Developer",
         "Apptarix Mobility", "Bangalore, India",
         "Filed provisional patent 201641021557 — a recommendation engine combining commercial and "
         "learnt rules with a method to measure rule effectiveness. Built an early NLP sentiment "
         "engine in R and Python. Managed a team of five.",
         [
             "Led architecture and full-stack development across OTT, analytics and EdTech product lines; "
             "designed scalable data ingestion, scraping and processing pipelines.",
             "Built early ML/NLP tooling (cosine similarity, market basket analysis, content extraction) "
             "predating modern LLM frameworks.",
         ])

    role(doc, "Jan 2020 – Aug 2021",
         "Technology Consultant", "BuyPal", "Melbourne, VIC",
         "Concurrent engagement alongside Quest. Led knowledge transfer and legacy platform migration "
         "to modern web standards.", [])

    # Earlier experience as a compact list
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(7)
    p.paragraph_format.space_after  = Pt(1)
    r = p.add_run("Earlier Experience")
    r.bold = True
    r.font.size = Pt(10.5)
    for line in [
        "Esri Australia (Aug–Sep 2017) — Full Stack Developer, AngularJS and Node.js, AWS, ArcGIS.",
        "Ionos Networks (Nov 2014–Mar 2015) — Resident Technology Lead via Apptarix; Node.js, PHP, Python, AWS S3.",
        "Studio 360eye (Sep 2012–May 2013) — Founder; distributed team of six, client management and delivery.",
        "UNI Group (Feb 2011–Sep 2012) — Senior Software Engineer and project lead.",
        "Hewlett-Packard / Mphasis (Jan 2009–Mar 2010) — Infrastructure Engineer, 70+ Linux/HP-UX servers, Perl automation.",
        "Robert Bosch (Mar–Jun 2007) — Project Trainee, Electronics Circuits Boards Division.",
    ]:
        simple_bullet(doc, line)

    # Education
    add_section(doc, "Education")
    education_entry(doc, "Jun 2023 – Aug 2023",
                    "Certificate in Business Analytics",
                    "Harvard Business School Online", "Online",
                    "Business analytics, strategy and commercial decision-making.")
    education_entry(doc, "2004 – 2008",
                    "BE, Electronics & Communication",
                    "Visvesvaraya Technological University", "Karnataka, India",
                    "First Class Honours. Electronics, embedded systems and signal processing.")

    # Languages
    add_section(doc, "Languages")
    body(doc, "English (Expert), Hindi (Fluent), Kannada (Native), Konkani (Native).")

    # Patents
    add_section(doc, "Patents & IP")
    body(doc, "Patent application 201641021557 — \"A system and method to measure the effectiveness and "
         "performance of commercial rules and learnt rules that recommend content to user(s).\" "
         "Filed 2016, Apptarix Mobility.")

    # References
    add_section(doc, "References")
    body(doc, "Available upon request.")

    return doc


# ---------------------------------------------------------------------------
# CBA CV
# ---------------------------------------------------------------------------

def build_cba_cv():
    doc = Document()
    set_margins(doc)

    add_name_header(doc, "Vrandesh Bandikatti")

    add_section(doc, "Profile")
    body(doc,
         "Hands-on Senior Software Engineer and Cloud Solution Architect with 17+ years designing, "
         "building and operating production platforms across financial services, payments and enterprise AI. "
         "Combines architecture ownership with full-stack delivery: designing platform foundations, "
         "establishing engineering standards and reusable golden paths, then building the services, "
         "CI/CD pipelines and cloud infrastructure to run them in production. Multi-cloud across Azure "
         "(primary delivery) and AWS (12+ microservices in regulated production at Quest Payment Systems); "
         "recent work spans Python, FastAPI, containerised workloads, DevSecOps, Grafana observability and "
         "AI-powered engineering tooling. Stepped in as interim DevOps lead during a critical production "
         "incident and remained hands-on through resolution.")

    add_section(doc, "Core Competencies")
    competency(doc, "Platform Engineering & DevSecOps",
               "CI/CD pipeline design and governance, Infrastructure as Code, Docker and Kubernetes, "
               "Azure DevOps, Grafana observability, security scanning and automated testing; self-service "
               "developer tooling, golden paths and engineering standards across multi-product "
               "platform estates.")
    competency(doc, "Cloud Architecture — Azure, AWS & GCP",
               "Azure (Fabric, OneLake, Entra ID SSO/RBAC, Durable Functions, Key Vault, DevOps, Blob "
               "Storage); AWS (Lambda, Step Functions, Bedrock, S3, EC2); GCP (Firebase, Firestore, Cloud "
               "Functions). Serverless, event-driven and distributed patterns; identity and access "
               "management; cost optimisation and reliability engineering.")
    competency(doc, "Full Stack Engineering",
               "Python, Django, FastAPI, Node.js, TypeScript, React, Vue.js, GraphQL and REST APIs; "
               "microservices and distributed systems; PostgreSQL, MongoDB, SQL Server, Firestore; "
               "SOLID design, secure coding and compliance-grade implementations.")
    competency(doc, "AI-Powered Engineering",
               "Azure OpenAI, Azure AI Foundry, LangGraph, LangChain, MCP integrations; RAG and agent "
               "orchestration in full-stack platforms; Claude Code and AI-assisted development practices "
               "that accelerate delivery and reduce integration risk across complex engineering workflows.")
    competency(doc, "Engineering Leadership",
               "Technical leadership across teams of 4–19; solution architecture, engineering standards, "
               "cross-functional delivery and executive stakeholder engagement while remaining hands-on "
               "in design, coding, troubleshooting and production support.")

    add_section(doc, "Professional Experience")

    role(doc, "Mar 2022 – Present",
         "Senior Full Stack Developer & Cloud/AI Solution Architect",
         "Findex Australia", "Melbourne, VIC",
         "Architect and hands-on engineer across Findex Data Science platforms, owning delivery from "
         "discovery through full-stack implementation, cloud deployment, UAT and production support "
         "using Python/Django, FastAPI, GraphQL, Vue/React, Azure and Microsoft Fabric.",
         [
             "Stepped in as interim DevOps lead during a critical production incident; established CI/CD "
             "pipelines, DevSecOps practices, Grafana observability and reusable deployment patterns "
             "across all Data Science platforms, improving release confidence and reducing manual "
             "operational overhead.",
             "Led cross-functional teams of 4–19 across Education Portal, PerformX and KRM/STI "
             "initiatives; partnered with executives to manage scope, delivery risks and stakeholder "
             "expectations from discovery through UAT and release.",
             "Built Payroll AI: four-agent pipeline using Azure OpenAI and LangGraph that reads Modern "
             "Award PDFs, extracts rules, generates client-specific calculation code and reports into "
             "Power BI with full audit trails. Registered as an eligible R&D activity under the "
             "DISR R&D Tax Incentive.",
             "Built PerformX Agent: conversational AI assistant on React/FastAPI with triple permission "
             "enforcement and bounded self-repair on query failure; deployed in browser and Microsoft "
             "Teams using AI-assisted development tooling including Claude Code.",
             "Architected KRM Remuneration Portal: Python/Django, GraphQL, Vue/React with Entra ID "
             "SSO/RBAC, row-level security and governed Microsoft Fabric foundations. Onboarded "
             "443 users and 129 people leaders.",
         ])

    role(doc, "Oct 2017 – Mar 2022",
         "Senior Software Engineer, Product Innovation",
         "Quest Payment Systems", "Melbourne, VIC",
         "Regulated payments engineering. Designed, built and operated 12+ microservices in regulated "
         "financial services production on AWS and Google Cloud.",
         [
             "Designed end-to-end serverless infrastructure on AWS (Lambda, Step Functions, S3) and GCP "
             "(Firebase, Firestore) for Donation Point Go; owned infrastructure, security and production "
             "operations for the full microservices estate across high-availability, "
             "unpredictable-load payment events.",
             "Built and maintained 12+ microservices and integrations including Xero and MYOB on "
             "Node.js, Django and Ember.js using Lambda, Firestore and Elasticsearch; led a team of "
             "three across Airpay mobile POS delivery.",
             "Delivered full-lifecycle platform engineering: IaC, automated testing, release pipelines "
             "and incident response for a PCI-adjacent contactless payments stack.",
         ])

    role(doc, "Jun 2013 – Jul 2017",
         "Technology Lead / Full Stack Developer",
         "Apptarix Mobility", "Bangalore, India",
         "Filed provisional patent 201641021557 — a recommendation engine combining commercial and "
         "learnt rules with a method to measure rule effectiveness. Built an early NLP sentiment "
         "engine in R and Python. Managed a team of five.",
         [
             "Led architecture and full-stack development across OTT, analytics and EdTech product lines; "
             "managed a team of five from greenfield to production.",
             "Designed scalable data ingestion, content extraction and automation pipelines; delivered "
             "early ML/NLP tooling and analytics dashboards.",
         ])

    role(doc, "Jan 2020 – Aug 2021",
         "Technology Consultant", "BuyPal", "Melbourne, VIC",
         "Concurrent engagement alongside Quest. Led knowledge transfer and legacy platform migration "
         "to modern web standards.", [])

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(7)
    p.paragraph_format.space_after  = Pt(1)
    r = p.add_run("Earlier Experience")
    r.bold = True
    r.font.size = Pt(10.5)
    for line in [
        "Esri Australia (Aug–Sep 2017) — Full Stack Developer, AngularJS and Node.js, AWS, ArcGIS.",
        "Ionos Networks (Nov 2014–Mar 2015) — Resident Technology Lead via Apptarix; Node.js, PHP, Python, AWS S3.",
        "Studio 360eye (Sep 2012–May 2013) — Founder; distributed team of six, client management and delivery.",
        "UNI Group (Feb 2011–Sep 2012) — Senior Software Engineer and project lead.",
        "Hewlett-Packard / Mphasis (Jan 2009–Mar 2010) — Infrastructure Engineer, 70+ Linux/HP-UX servers, Perl automation.",
        "Robert Bosch (Mar–Jun 2007) — Project Trainee, Electronics Circuits Boards Division.",
    ]:
        simple_bullet(doc, line)

    add_section(doc, "Education")
    education_entry(doc, "Jun 2023 – Aug 2023",
                    "Certificate in Business Analytics",
                    "Harvard Business School Online", "Online",
                    "Business analytics, strategy and commercial decision-making.")
    education_entry(doc, "2004 – 2008",
                    "BE, Electronics & Communication",
                    "Visvesvaraya Technological University", "Karnataka, India",
                    "First Class Honours. Electronics, embedded systems and signal processing.")

    add_section(doc, "Languages")
    body(doc, "English (Expert), Hindi (Fluent), Kannada (Native), Konkani (Native).")

    add_section(doc, "Patents & IP")
    body(doc, "Patent application 201641021557 — \"A system and method to measure the effectiveness and "
         "performance of commercial rules and learnt rules that recommend content to user(s).\" "
         "Filed 2016, Apptarix Mobility.")

    add_section(doc, "References")
    body(doc, "Available upon request.")

    return doc


# ---------------------------------------------------------------------------
# Cover letter builder
# ---------------------------------------------------------------------------

def build_cover(paragraphs_before_bullets, bullets, paragraphs_after, closing_name):
    """
    paragraphs_before_bullets: list of (text, is_intro_label) tuples
    bullets: list of (bold_label, rest_of_text) tuples
    paragraphs_after: list of plain text strings
    """
    doc = Document()
    set_margins(doc, top=0.85, bottom=0.85, left=0.85, right=0.85)

    add_name_header(doc, "Vrandesh Bandikatti")

    # Date
    from datetime import date
    p_date = doc.add_paragraph()
    p_date.paragraph_format.space_before = Pt(6)
    p_date.paragraph_format.space_after  = Pt(10)
    d = date.today()
    date_str = f"{d.day} {d.strftime('%B %Y')}"
    p_date.add_run(date_str).font.size = Pt(10)

    for text, is_label in paragraphs_before_bullets:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(8)
        r = p.add_run(text)
        r.font.size = Pt(10)

    # Bullets
    for bold_label, rest in bullets:
        pb = doc.add_paragraph(style="List Bullet")
        pb.paragraph_format.space_before     = Pt(2)
        pb.paragraph_format.space_after      = Pt(2)
        pb.paragraph_format.left_indent      = Inches(0.25)
        pb.paragraph_format.first_line_indent = Inches(-0.15)
        if bold_label:
            rb = pb.add_run(bold_label + ": ")
            rb.bold = True
            rb.font.size = Pt(10)
        rb2 = pb.add_run(rest)
        rb2.font.size = Pt(10)

    for text in paragraphs_after:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after  = Pt(0)
        p.add_run(text).font.size = Pt(10)

    # Closing
    p_close = doc.add_paragraph()
    p_close.paragraph_format.space_before = Pt(16)
    p_close.paragraph_format.space_after  = Pt(0)
    p_close.add_run("Kind regards,").font.size = Pt(10)

    p_sig = doc.add_paragraph()
    p_sig.paragraph_format.space_before = Pt(14)
    p_sig.paragraph_format.space_after  = Pt(0)
    rs = p_sig.add_run(closing_name)
    rs.bold = True
    rs.font.size = Pt(10)

    return doc


# ---------------------------------------------------------------------------
# Xero cover letter content
# ---------------------------------------------------------------------------

XERO_BEFORE = [
    ("Dear Xero hiring team,", False),
    ("I am very interested in the Senior Engineer, AI Workflows role at Xero. For the past four years I "
     "have been building exactly the systems this role describes at Findex Australia: production AI agents "
     "using Azure OpenAI and LangGraph with MCP integrations, triple permission enforcement, bounded "
     "self-repair loops and Grafana-instrumented evaluation frameworks, all shipped inside a full-stack "
     "enterprise platform, not deployed as standalone prototypes.", False),
    ("Three deliveries speak directly to this role:", False),
]

XERO_BULLETS = [
    ("PerformX Agent (Findex)",
     "conversational AI assistant on React/FastAPI that answers plain-language data questions against a "
     "governed Power BI model, with three independent permission checks and a self-repair loop that "
     "rewrites failed queries or exits gracefully rather than hallucinating. Developed end-to-end "
     "using Claude Code."),
    ("Payroll AI (Findex)",
     "four-agent pipeline from PDF ingestion through rule extraction, calculation code generation and "
     "Power BI reporting, registered as an eligible R&D activity under the DISR Tax Incentive."),
    ("Xero and MYOB integrations (Quest Payment Systems)",
     "real-time financial data reconciliation across payment and accounting systems. I know the Xero "
     "API from the other side of the integration."),
]

XERO_AFTER = [
    "Xero’s purpose, making life better for people in small business and their advisers, maps directly "
    "to what I have been building. The integration targets in this role (Workday, Slack, NetSuite, "
    "Salesforce) are the same class of webhook-and-SDK enterprise connectivity I have built against in "
    "regulated financial services environments. With direct Xero and MYOB API experience, I know the "
    "platform your AI layer is extending. Your coaching mandate fits naturally: at Findex I embedded "
    "alongside Finance and HR teams to build AI fluency as we co-designed and delivered, rather than "
    "documenting and handing off.",

    "I would welcome the opportunity to discuss how my experience across agentic AI systems, enterprise "
    "integration and full-stack platform delivery could contribute to Xero’s AI engineering team.",
]

# ---------------------------------------------------------------------------
# CBA cover letter content
# ---------------------------------------------------------------------------

CBA_BEFORE = [
    ("Dear Hiring Manager,", False),
    ("I am applying for the Principal Software Engineer role in Business Banking Technology. The mandate "
     "— hands-on technical leadership designing platform foundations, engineering standards and "
     "developer tooling for a major bank — aligns closely with what I have been delivering across "
     "the past decade in regulated financial services.", False),
    ("Three areas of experience I would bring directly:", False),
]

CBA_BULLETS = [
    ("Platform engineering and DevSecOps (Findex)",
     "stepped in as interim DevOps lead during a critical production incident; established CI/CD "
     "pipelines, Grafana observability, DevSecOps practices and reusable deployment patterns across a "
     "multi-product enterprise platform estate, improving release confidence and reducing manual "
     "operational overhead."),
    ("AWS production at scale (Quest Payment Systems)",
     "designed, built and operated 12+ microservices on AWS (Lambda, Step Functions, S3) and GCP for a "
     "regulated, PCI-adjacent contactless payment platform; owned infrastructure, security and incident "
     "response for high-availability financial services workloads."),
    ("AI-powered engineering practices",
     "integrated Claude Code and Azure OpenAI into development workflows at Findex, accelerating agent "
     "evaluation and integration testing across complex multi-product delivery; accustomed to championing "
     "AI tooling adoption within engineering teams."),
]

CBA_AFTER = [
    "Commonwealth Bank’s scale and regulatory complexity is what makes this role compelling. I have "
    "spent my career in environments where engineering decisions carry compliance weight and where platform "
    "reliability directly affects financial outcomes for real people. Business Banking Technology sits at "
    "exactly that intersection, and the Principal mandate — designing golden paths, raising "
    "engineering standards and mentoring senior engineers while staying hands-on — is how I have led "
    "throughout my career.",

    "I would welcome the opportunity to discuss how my background maps to your current priorities.",

    "I look forward to hearing from you.",
]

# ---------------------------------------------------------------------------
# Preacta CV
# ---------------------------------------------------------------------------

def build_preacta_cv():
    doc = Document()
    set_margins(doc)

    add_name_header(doc, "Vrandesh Bandikatti")

    add_section(doc, "Profile")
    body(doc,
         "Senior Full Stack Engineer with 17+ years building enterprise platforms, specialising in "
         "production-grade agentic AI systems. Takes architecture ownership from day one: delivered "
         "multi-agent pipelines, RAG architectures, function calling, structured outputs and "
         "human-in-the-loop controls using Azure OpenAI, AWS Bedrock, LangGraph and LangChain — "
         "integrated into full-stack enterprise products rather than deployed as standalone prototypes. "
         "Track record of greenfield AI builds in regulated financial services environments, where the "
         "compliance outcome is explainable and repeatable, not a model's opinion. Cloud-agnostic "
         "across Azure, AWS and GCP.")

    add_section(doc, "Core Competencies")
    competency(doc, "Agentic AI Systems & RAG",
               "Multi-agent orchestration, RAG pipelines, function calling, structured outputs and "
               "evaluation frameworks using Azure OpenAI, Azure AI Foundry, AWS Bedrock, LangGraph, "
               "LangChain and MCP integrations. Human-in-the-loop controls, bounded self-repair and "
               "deterministic rule separation — shipped in production financial services platforms.")
    competency(doc, "Full Stack Engineering",
               "Python, FastAPI, Django, TypeScript, React, Vue.js, Node.js; GraphQL and REST APIs; "
               "microservices and event-driven systems; PostgreSQL, MongoDB, SQL Server, Firestore; "
               "Docker, CI/CD, DevSecOps.")
    competency(doc, "Cloud Architecture & Model Deployment",
               "Azure (Durable Functions, Functions, Blob Storage, Entra ID, DevOps, ADF), AWS "
               "(Lambda, Step Functions, Bedrock, S3), GCP (Firebase, Firestore, Cloud Functions). "
               "Serverless and event-driven patterns for AI workload deployment; API-based model "
               "serving via Azure OpenAI and AWS Bedrock at production scale.")
    competency(doc, "Observability & Incident Response",
               "Grafana dashboards, structured logging and AI workflow monitoring; bounded graceful "
               "degradation patterns; full production incident ownership — including stepping in as "
               "interim DevOps lead during a critical outage and stabilising systems under pressure.")
    competency(doc, "FinTech Platform Engineering",
               "Compliance-grade AI design in regulated environments; payroll compliance, payment "
               "microservices and financial data movement; audit trails, row-level security and "
               "explainability as design patterns across all production AI systems.")

    add_section(doc, "Professional Experience")

    role(doc, "Mar 2022 – Present",
         "Senior Full Stack Developer & Cloud/AI Solution Architect",
         "Findex Australia", "Melbourne, VIC",
         "Architect and hands-on engineer across Findex Data Science platforms, owning delivery from "
         "discovery through full-stack implementation, cloud deployment and production support "
         "using Python/Django, FastAPI, GraphQL, Vue/React, Azure and Microsoft Fabric.",
         [
             "Led cross-functional teams of 4–19 across Education Portal, PerformX and KRM/STI "
             "initiatives; partnered with executives to manage scope, priorities, delivery risks and "
             "stakeholder expectations from discovery through release.",
             "Built Payroll AI: four-agent pipeline reading Modern Award PDFs, extracting rules with "
             "clause references, generating client-specific payroll calculation code and reporting into "
             "Power BI with full audit trails. Registered as an eligible R&D activity under the DISR "
             "R&D Tax Incentive. Developed end-to-end using Claude Code.",
             "Built PerformX Agent: conversational AI assistant with triple permission enforcement and "
             "bounded self-repair loop on query failure, deployed in browser and Microsoft Teams. "
             "Self-repair rewrites failed queries or exits gracefully rather than returning incorrect answers.",
             "Built Copilot Studio BI Agent: Teams bot grounded on SharePoint and Excel knowledge "
             "sources with a custom DAX connector into the governed Power BI semantic model; Entra ID "
             "and Power Platform DLP governance; consumed by two production platforms; removed a "
             "manual Finance adjustment step running every cycle.",
             "Designed WealthAI TMD Processing (serverless Durable Functions pipeline for TMD "
             "document classification feeding a deterministic rule engine) and EduPort; established "
             "CI/CD, Grafana observability and human-in-the-loop patterns; stepped in as interim "
             "DevOps lead during a critical production incident.",
         ])

    role(doc, "Oct 2017 – Mar 2022",
         "Senior Software Engineer, Product Innovation",
         "Quest Payment Systems", "Melbourne, VIC",
         "Regulated payments engineering. Led backend and cloud delivery from concept to production "
         "for Donation Point Go, a contactless donation channel. Built and maintained 12+ microservices "
         "on AWS and Google Cloud.",
         [
             "Designed and delivered end-to-end serverless architecture on AWS (Lambda, Step Functions, "
             "S3) and GCP (Firebase, Firestore, Cloud Functions) for high availability and "
             "unpredictable charitable-event load.",
             "Built and maintained third-party integrations including Xero and MYOB for real-time "
             "financial data movement and reconciliation; designed high-availability architectures on "
             "Node.js, Django and Ember.js with Elasticsearch.",
             "Led a team of three; remained hands-on in backend engineering for Airpay mobile payments "
             "and POS applications.",
         ])

    role(doc, "Jun 2013 – Jul 2017",
         "Technology Lead / Full Stack Developer",
         "Apptarix Mobility", "Bangalore, India",
         "Filed provisional patent 201641021557 — a recommendation engine combining commercial and "
         "learnt rules. Built an early NLP sentiment engine using cosine similarity in R and Python. "
         "Managed a team of five across OTT, analytics and EdTech product lines.",
         [
             "Led architecture and full-stack development of analytics dashboards, OTT platforms and "
             "EdTech products; designed scalable data ingestion and processing pipelines.",
             "Built early ML/NLP tooling (cosine similarity, market basket analysis, content extraction) "
             "predating modern LLM frameworks.",
         ])

    role(doc, "Jan 2020 – Aug 2021",
         "Technology Consultant", "BuyPal", "Melbourne, VIC",
         "Concurrent engagement alongside Quest. Led knowledge transfer and legacy platform migration "
         "to modern web standards.", [])

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(7)
    p.paragraph_format.space_after  = Pt(1)
    r = p.add_run("Earlier Experience")
    r.bold = True
    r.font.size = Pt(10.5)
    for line in [
        "Esri Australia (Aug–Sep 2017) — Full Stack Developer, AngularJS and Node.js, AWS, ArcGIS.",
        "Ionos Networks (Nov 2014–Mar 2015) — Resident Technology Lead via Apptarix; Node.js, PHP, Python, AWS S3.",
        "Studio 360eye (Sep 2012–May 2013) — Founder; distributed team of six, client management and delivery.",
        "UNI Group (Feb 2011–Sep 2012) — Senior Software Engineer and project lead.",
        "Hewlett-Packard / Mphasis (Jan 2009–Mar 2010) — Infrastructure Engineer, 70+ Linux/HP-UX servers, Perl automation.",
        "Robert Bosch (Mar–Jun 2007) — Project Trainee, Electronics Circuits Boards Division.",
    ]:
        simple_bullet(doc, line)

    add_section(doc, "Education")
    education_entry(doc, "Jun 2023 – Aug 2023",
                    "Certificate in Business Analytics",
                    "Harvard Business School Online", "Online",
                    "Business analytics, strategy and commercial decision-making.")
    education_entry(doc, "2004 – 2008",
                    "BE, Electronics & Communication",
                    "Visvesvaraya Technological University", "Karnataka, India",
                    "First Class Honours. Electronics, embedded systems and signal processing.")

    add_section(doc, "Languages")
    body(doc, "English (Expert), Hindi (Fluent), Kannada (Native), Konkani (Native).")

    add_section(doc, "Patents & IP")
    body(doc, "Patent application 201641021557 — \"A system and method to measure the effectiveness and "
         "performance of commercial rules and learnt rules that recommend content to user(s).\" "
         "Filed 2016, Apptarix Mobility.")

    add_section(doc, "References")
    body(doc, "Available upon request.")

    return doc


# ---------------------------------------------------------------------------
# iGoDirect CV
# ---------------------------------------------------------------------------

def build_igodirect_cv():
    doc = Document()
    set_margins(doc)

    add_name_header(doc, "Vrandesh Bandikatti")

    add_section(doc, "Profile")
    body(doc,
         "Senior Full Stack Engineer and Cloud Solution Architect with 17+ years building enterprise "
         "platforms, specialising in production-grade agentic AI systems. Sets and owns AI architecture "
         "standards across engineering teams: designed and delivered agent orchestration, RAG, function "
         "calling and human-in-the-loop controls inside regulated financial services platforms, remaining "
         "hands-on in design, coding and release. Brings direct experience in payment-regulated financial "
         "services and proprietary platform engineering. Cloud-agnostic across Azure, AWS and GCP; "
         "consistent track record of greenfield AI capability built from the ground up.")

    add_section(doc, "Core Competencies")
    competency(doc, "Agentic AI Architecture & Standards",
               "Designed and owned AI architecture standards: agent orchestration, RAG, function "
               "calling, structured outputs, evaluation frameworks and human-in-the-loop controls "
               "using Azure OpenAI, AWS Bedrock, LangGraph, LangChain, Copilot Studio and MCP "
               "integrations — shipped in production financial services platforms, not as standalone prototypes.")
    competency(doc, "Regulated FinTech Platform Engineering",
               "AFSL-adjacent regulated systems and payment microservices; Visa-adjacent payment "
               "integrations (mobile POS, contactless donation channels); real-time financial data "
               "movement and reconciliation; compliance-grade design with audit trails, row-level "
               "security and explainability.")
    competency(doc, "Full Stack Engineering",
               "Python, FastAPI, Django, TypeScript, React, Vue.js, Node.js; GraphQL and REST APIs; "
               "microservices, distributed systems and event-driven architecture; PostgreSQL, MongoDB, "
               "SQL Server, Firestore; Docker, CI/CD, DevSecOps.")
    competency(doc, "Cloud Architecture",
               "Azure (Fabric, OneLake, Entra ID SSO/RBAC, Power BI Embedded, Durable Functions, ADF, "
               "DevOps), AWS (Lambda, Step Functions, Bedrock, S3), GCP (Firebase, Firestore, Cloud "
               "Functions). Serverless, event-driven and auto-scaling patterns for regulated "
               "production environments.")
    competency(doc, "Engineering Leadership & Standards-Setting",
               "Architecture discovery and standards-setting for high-ownership engineering teams; "
               "solution design, cross-functional team leadership (4–19), stakeholder engagement and "
               "hands-on delivery from concept through production.")

    add_section(doc, "Professional Experience")

    role(doc, "Mar 2022 – Present",
         "Senior Full Stack Developer & Cloud/AI Solution Architect",
         "Findex Australia", "Melbourne, VIC",
         "Set architecture and engineering standards for all AI-enabled platforms at Findex; owned "
         "delivery from discovery through full-stack implementation, cloud deployment and production "
         "support using Python/Django, FastAPI, GraphQL, Vue/React, Azure and Microsoft Fabric.",
         [
             "Led cross-functional teams of 4–19 across Education Portal, PerformX and KRM/STI "
             "initiatives; partnered with executives on scope, priorities, delivery risks and "
             "stakeholder expectations.",
             "Built Payroll AI: four-agent pipeline converting Modern Award PDFs into payroll "
             "calculation code with full clause traceability and audit trails. Registered as an "
             "eligible R&D activity under the DISR R&D Tax Incentive. Developed end-to-end using Claude Code.",
             "Built PerformX Agent: conversational AI assistant answering plain-language data questions "
             "against a governed Power BI model with triple permission enforcement and bounded self-repair. "
             "Deployed in browser and Microsoft Teams.",
             "Built Copilot Studio BI Agent: Teams bot with a custom DAX connector into the governed "
             "Power BI semantic model; Entra ID and DLP governance; consumed by two production "
             "platforms; removed a manual Finance adjustment step running every cycle.",
             "Architected KRM Remuneration Portal (443 users, 129 people leaders) and WealthAI TMD "
             "Processing; established CI/CD, Grafana observability and human-in-the-loop patterns; "
             "stepped in as interim DevOps lead during a critical production incident.",
         ])

    role(doc, "Oct 2017 – Mar 2022",
         "Senior Software Engineer, Product Innovation",
         "Quest Payment Systems", "Melbourne, VIC",
         "Regulated payments engineering. Led backend and cloud delivery from concept to production "
         "for Donation Point Go, a contactless donation channel for charities, and Airpay, a "
         "mobile POS application.",
         [
             "Built 12+ microservices on AWS (Lambda, Step Functions, S3) and GCP (Firebase, "
             "Firestore, Cloud Functions) for high-availability, event-driven payment processing "
             "under unpredictable load.",
             "Built Visa-adjacent payment integrations for Airpay mobile POS: real-time transaction "
             "processing, financial data movement and automated reconciliation with Xero and MYOB.",
             "Led a team of three while remaining hands-on in backend engineering; designed "
             "high-availability architectures on Node.js, Django and Ember.js with Elasticsearch.",
         ])

    role(doc, "Jun 2013 – Jul 2017",
         "Technology Lead / Full Stack Developer",
         "Apptarix Mobility", "Bangalore, India",
         "Filed provisional patent 201641021557 — a recommendation engine combining commercial and "
         "learnt rules. Built an early NLP sentiment engine in R and Python. Managed a team of five "
         "across OTT, analytics and EdTech product lines.",
         [
             "Led architecture and full-stack development of analytics dashboards, OTT platforms and "
             "EdTech products; designed scalable data ingestion and processing pipelines.",
             "Built early ML/NLP tooling (cosine similarity, market basket analysis, content extraction) "
             "predating modern LLM frameworks.",
         ])

    role(doc, "Jan 2020 – Aug 2021",
         "Technology Consultant", "BuyPal", "Melbourne, VIC",
         "Concurrent engagement alongside Quest. Led knowledge transfer and legacy platform migration "
         "to modern web standards.", [])

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(7)
    p.paragraph_format.space_after  = Pt(1)
    r = p.add_run("Earlier Experience")
    r.bold = True
    r.font.size = Pt(10.5)
    for line in [
        "Esri Australia (Aug–Sep 2017) — Full Stack Developer, AngularJS and Node.js, AWS, ArcGIS.",
        "Ionos Networks (Nov 2014–Mar 2015) — Resident Technology Lead via Apptarix; Node.js, PHP, Python, AWS S3.",
        "Studio 360eye (Sep 2012–May 2013) — Founder; distributed team of six, client management and delivery.",
        "UNI Group (Feb 2011–Sep 2012) — Senior Software Engineer and project lead.",
        "Hewlett-Packard / Mphasis (Jan 2009–Mar 2010) — Infrastructure Engineer, 70+ Linux/HP-UX servers, Perl automation.",
        "Robert Bosch (Mar–Jun 2007) — Project Trainee, Electronics Circuits Boards Division.",
    ]:
        simple_bullet(doc, line)

    add_section(doc, "Education")
    education_entry(doc, "Jun 2023 – Aug 2023",
                    "Certificate in Business Analytics",
                    "Harvard Business School Online", "Online",
                    "Business analytics, strategy and commercial decision-making.")
    education_entry(doc, "2004 – 2008",
                    "BE, Electronics & Communication",
                    "Visvesvaraya Technological University", "Karnataka, India",
                    "First Class Honours. Electronics, embedded systems and signal processing.")

    add_section(doc, "Languages")
    body(doc, "English (Expert), Hindi (Fluent), Kannada (Native), Konkani (Native).")

    add_section(doc, "Patents & IP")
    body(doc, "Patent application 201641021557 — \"A system and method to measure the effectiveness and "
         "performance of commercial rules and learnt rules that recommend content to user(s).\" "
         "Filed 2016, Apptarix Mobility.")

    add_section(doc, "References")
    body(doc, "Available upon request.")

    return doc


# ---------------------------------------------------------------------------
# Preacta cover letter content
# ---------------------------------------------------------------------------

PREACTA_BEFORE = [
    ("Dear Hiring Manager,", False),
    ("I am very interested in the AI Systems Developer opportunity being represented by Preacta "
     "Recruitment. With 17+ years of software engineering experience, I have spent the past three "
     "years building exactly this: production-grade agentic AI systems with full architecture "
     "ownership, from RAG and multi-agent pipelines through to model deployment and incident "
     "response — inside regulated financial services environments where the compliance outcome "
     "must be explainable and repeatable.", False),
    ("My most relevant delivery:", False),
]

PREACTA_BULLETS = [
    ("Payroll AI (Findex)",
     "Four-agent pipeline from Modern Award PDF ingestion through rule extraction, payroll "
     "calculation code generation and Power BI reporting with full audit trails. Registered as an "
     "eligible R&D activity under the DISR Tax Incentive. Developed end-to-end using Claude Code."),
    ("PerformX Agent (Findex)",
     "Conversational AI assistant on React/FastAPI with triple permission enforcement and bounded "
     "self-repair — designed so the system degrades gracefully rather than returning incorrect answers."),
    ("Copilot Studio BI Agent (Findex)",
     "Teams-based AI agent with a custom DAX connector, Entra ID governance and Power Platform "
     "DLP; consumed by two production platforms and removed a manual Finance process running every cycle."),
]

PREACTA_AFTER = [
    "The greenfield framing aligns precisely with how I have worked best. Payroll AI, PerformX and "
    "WealthAI TMD Processing were all built from scratch with architecture ownership from the first "
    "whiteboard session — observability, audit trails and human-in-the-loop controls built in, "
    "not bolted on afterwards. I bring the same discipline to this build.",

    "I would welcome the opportunity to discuss how my experience across agentic AI architecture, "
    "full-stack delivery and regulated FinTech platforms could contribute to this role.",
]


# ---------------------------------------------------------------------------
# iGoDirect cover letter content
# ---------------------------------------------------------------------------

IGODIRECT_BEFORE = [
    ("Dear iGoDirect team,", False),
    ("I am very interested in the Principal AI Software Engineer opportunity at iGoDirect. With "
     "17+ years of software engineering experience, I have spent the past three years building "
     "production-grade agentic AI systems at Findex Australia — setting architecture standards "
     "and owning end-to-end delivery, not just implementing what others specify. The mandate to "
     "lead how iGoDirect builds software with AI is precisely the kind of greenfield ownership "
     "I do best.", False),
    ("My most relevant delivery:", False),
]

IGODIRECT_BULLETS = [
    ("Payroll AI (Findex)",
     "Four-agent pipeline from Modern Award PDF ingestion through rule extraction, payroll "
     "calculation code generation and Power BI reporting — every compliance decision traceable "
     "and repeatable. Developed end-to-end using Claude Code."),
    ("PerformX Agent (Findex)",
     "Conversational AI assistant with triple permission enforcement and bounded self-repair, "
     "ensuring the system degrades gracefully rather than returning incorrect answers."),
    ("Airpay and Donation Point Go (Quest Payment Systems)",
     "Led backend delivery of Visa-adjacent payment platforms with 12+ microservices on AWS and "
     "GCP — the same class of regulated, real-time financial transaction infrastructure that "
     "underpins card-issuing programmes."),
]

IGODIRECT_AFTER = [
    "iGoDirect's position as an AFSL holder and Visa Principal Member operating proprietary "
    "platform technology is a combination I understand deeply. I have built inside regulated "
    "financial services environments where audit trails, access controls and explainable system "
    "behaviour are requirements, not afterthoughts — and I have led payment engineering teams "
    "that work at the same intersection of compliance and product velocity.",

    "I would welcome the opportunity to discuss how my experience across agentic AI architecture, "
    "regulated FinTech delivery and full-stack platform engineering could contribute to "
    "iGoDirect's AI direction.",
]


# ---------------------------------------------------------------------------
# Railway CV
# ---------------------------------------------------------------------------

def build_railway_cv():
    doc = Document()
    set_margins(doc)

    add_name_header(doc, "Vrandesh Bandikatti")

    add_section(doc, "Profile")
    body(doc,
         "Senior Full Stack Engineer with 17+ years building production-grade platforms, specialising in "
         "end-to-end product delivery from UI to backend workflows. Designs and ships TypeScript and "
         "GraphQL APIs with strong data-modelling guarantees, builds React frontends that handle complex "
         "data-fetching and mutation patterns, and orchestrates complex async backend pipelines — all owned "
         "from technical specification through production monitoring. Accustomed to high-ownership, "
         "high-autonomy environments where the distance between idea and shipped product is short. Applies "
         "AI-assisted tooling including Claude Code to accelerate delivery and reduce integration risk "
         "across complex engineering workflows.")

    add_section(doc, "Core Competencies")
    competency(doc, "Full Stack Engineering",
               "TypeScript, React, Vue.js, Node.js, Python, FastAPI, Django; GraphQL and REST API design "
               "and implementation; microservices, distributed systems and event-driven architecture; "
               "PostgreSQL, MongoDB, SQL Server, Firestore; Docker, CI/CD and DevSecOps.")
    competency(doc, "API Design & Async Pipeline Orchestration",
               "GraphQL APIs with strong data-modelling guarantees; complex async backend job orchestration "
               "using AWS Step Functions, Azure Durable Functions and Lambda (Temporal-equivalent patterns); "
               "build/deploy and event-processing pipelines at scale; internal and external API design "
               "with typed contracts.")
    competency(doc, "Platform Engineering & Observability",
               "CI/CD pipeline design and governance, DevSecOps practices, Grafana observability dashboards, "
               "structured logging and AI workflow monitoring; production incident ownership from detection "
               "through resolution; bounded graceful degradation patterns.")
    competency(doc, "AI-Assisted Engineering",
               "Claude Code and Azure OpenAI integrated into development workflows; agent orchestration "
               "(LangGraph, LangChain, MCP integrations) and RAG pipelines in production full-stack "
               "platforms; AI tooling adoption championed across engineering teams.")
    competency(doc, "Engineering Leadership & End-to-End Ownership",
               "Full project lifecycle from Engineering Requirement Document to production monitoring; "
               "cross-functional team leadership (4–19); high-agency delivery with executive stakeholder "
               "engagement; written technical design and async communication practices.")

    add_section(doc, "Professional Experience")

    role(doc, "Mar 2022 – Present",
         "Senior Full Stack Developer & Cloud/AI Solution Architect",
         "Findex Australia", "Melbourne, VIC",
         "Architect and hands-on engineer across Findex Data Science platforms, owning delivery from "
         "technical specification through full-stack implementation, cloud deployment, UAT and production "
         "monitoring using TypeScript, Python/Django, FastAPI, GraphQL, Vue/React, Azure and Microsoft Fabric.",
         [
             "Built PerformX Agent end-to-end: React/FastAPI product with GraphQL data layer answering "
             "plain-language questions against a governed Power BI model, with triple permission enforcement "
             "and a bounded self-repair loop that rewrites failed queries or exits gracefully rather than "
             "returning incorrect answers. Owned from Engineering Requirement Document through implementation, "
             "deployment and production monitoring; developed using Claude Code.",
             "Architected KRM Remuneration Portal: full-stack STI, referral and audit platform using "
             "Python/Django, GraphQL and Vue/React with Entra ID SSO/RBAC, row-level security and governed "
             "Microsoft Fabric foundations. Onboarded 443 users and 129 people leaders.",
             "Built Payroll AI: four-agent async pipeline from Modern Award PDF ingestion through rule "
             "extraction, calculation code generation and Power BI reporting with full audit trails — "
             "owned from planning and technical specification through implementation and production "
             "monitoring. Registered as an eligible R&D activity under the DISR R&D Tax Incentive.",
             "Led cross-functional teams of 4–19 across Education Portal, PerformX and KRM/STI initiatives; "
             "partnered with executives to manage scope, delivery risks and stakeholder expectations.",
             "Established CI/CD pipelines, Grafana observability dashboards and DevSecOps practices across "
             "all Data Science platforms; stepped in as interim DevOps lead during a critical production "
             "incident, stabilising systems and improving observability tooling under pressure.",
         ])

    role(doc, "Oct 2017 – Mar 2022",
         "Senior Software Engineer, Product Innovation",
         "Quest Payment Systems", "Melbourne, VIC",
         "Regulated payments engineering. Designed, built and operated 12+ microservices on AWS and Google "
         "Cloud for a contactless donation platform and mobile POS. Owned full product lifecycle from ERD "
         "and architecture through implementation, release and incident response.",
         [
             "Designed high-availability event-processing architecture for Donation Point Go on AWS Lambda "
             "and Step Functions, handling unpredictable transaction volumes for a regulated contactless "
             "payment channel deployed across charities and venues nationally; full serverless estate "
             "extended across GCP (Firebase, Firestore, Cloud Functions).",
             "Built TypeScript/Node.js and Django APIs with Xero, MYOB and Elasticsearch integrations; "
             "designed GraphQL and REST contracts consumed by internal and external clients; led backend "
             "delivery for Airpay mobile POS.",
             "Led a team of three while remaining hands-on in backend engineering, architecture and "
             "production operations across the full 12+ microservice estate.",
         ])

    role(doc, "Jun 2013 – Jul 2017",
         "Technology Lead / Full Stack Developer",
         "Apptarix Mobility", "Bangalore, India",
         "Filed provisional patent 201641021557 — a recommendation engine combining commercial and learnt "
         "rules. Built an early NLP engine in R and Python. Managed a team of five across OTT, analytics "
         "and EdTech product lines.",
         [
             "Led architecture and full-stack development of analytics dashboards, OTT platforms and EdTech "
             "products; designed scalable data ingestion, scraping and async processing pipelines.",
             "Built early ML/NLP tooling (cosine similarity, market basket analysis, content extraction) "
             "predating modern LLM frameworks.",
         ])

    role(doc, "Jan 2020 – Aug 2021",
         "Technology Consultant", "BuyPal", "Melbourne, VIC",
         "Concurrent engagement alongside Quest. Led knowledge transfer and legacy platform migration "
         "to modern web standards.", [])

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(7)
    p.paragraph_format.space_after  = Pt(1)
    r = p.add_run("Earlier Experience")
    r.bold = True
    r.font.size = Pt(10.5)
    for line in [
        "Esri Australia (Aug–Sep 2017) — Full Stack Developer, AngularJS and Node.js, AWS, ArcGIS.",
        "Ionos Networks (Nov 2014–Mar 2015) — Resident Technology Lead via Apptarix; Node.js, PHP, Python, AWS S3.",
        "Studio 360eye (Sep 2012–May 2013) — Founder; distributed team of six, client management and delivery.",
        "UNI Group (Feb 2011–Sep 2012) — Senior Software Engineer and project lead.",
        "Hewlett-Packard / Mphasis (Jan 2009–Mar 2010) — Infrastructure Engineer, 70+ Linux/HP-UX servers, Perl automation.",
        "Robert Bosch (Mar–Jun 2007) — Project Trainee, Electronics Circuits Boards Division.",
    ]:
        simple_bullet(doc, line)

    add_section(doc, "Education")
    education_entry(doc, "Jun 2023 – Aug 2023",
                    "Certificate in Business Analytics",
                    "Harvard Business School Online", "Online",
                    "Business analytics, strategy and commercial decision-making.")
    education_entry(doc, "2004 – 2008",
                    "BE, Electronics & Communication",
                    "Visvesvaraya Technological University", "Karnataka, India",
                    "First Class Honours. Electronics, embedded systems and signal processing.")

    add_section(doc, "Languages")
    body(doc, "English (Expert), Hindi (Fluent), Kannada (Native), Konkani (Native).")

    add_section(doc, "Patents & IP")
    body(doc, "Patent application 201641021557 — \"A system and method to measure the effectiveness and "
         "performance of commercial rules and learnt rules that recommend content to user(s).\" "
         "Filed 2016, Apptarix Mobility.")

    add_section(doc, "References")
    body(doc, "Available upon request.")

    return doc


# ---------------------------------------------------------------------------
# Railway cover letter content
# ---------------------------------------------------------------------------

RAILWAY_BEFORE = [
    ("Dear Railway hiring team,", False),
    ("I am very interested in the Senior Full-Stack Engineer, Product opportunity at Railway. With 17+ "
     "years of full-stack engineering experience across financial services, payments and AI platforms, "
     "the role's model of end-to-end ownership from UI through async backend pipelines reflects exactly "
     "how I have been working. For the past four-plus years at Findex Australia I have been building "
     "exactly this: TypeScript and GraphQL APIs with strong data-modelling guarantees, React frontends "
     "handling complex data-fetching and mutation patterns, and async backend pipelines owned from "
     "technical specification through production monitoring, not handed off once shipped.", False),
    ("Three deliveries speak directly to this role:", False),
]

RAILWAY_BULLETS = [
    ("PerformX Agent (Findex)",
     "React/FastAPI product with GraphQL data layer answering plain-language questions against a governed "
     "Power BI model, with triple permission enforcement and a self-repair loop that rewrites failed "
     "queries or exits gracefully. Owned from Engineering Requirement Document through implementation "
     "and production monitoring using Claude Code."),
    ("Payroll AI (Findex)",
     "Four-agent async pipeline from PDF ingestion through rule extraction, calculation code generation "
     "and reporting, registered as an R&D activity under the DISR Tax Incentive. Full lifecycle "
     "ownership from spec to production."),
    ("Quest Payment Systems",
     "Designed and operated 12+ microservices on AWS Lambda and Step Functions for a regulated "
     "contactless payment platform; built Xero and MYOB integrations and led backend delivery "
     "of Airpay mobile POS."),
]

RAILWAY_AFTER = [
    "Railway's purpose of rebuilding the cloud computing stack to make software creation faster and "
    "more accessible maps directly to what I find most compelling: tooling that removes friction and "
    "gives builders more leverage. A team of 42 shipping weekly to two million users is a context I "
    "am ready for. Temporal is not in my current stack, but I have delivered equivalent orchestration "
    "patterns with AWS Step Functions and Azure Durable Functions, and the mental model transfers "
    "cleanly. Rust is an honest gap; I would start by reading the CLI and Nixpacks repositories, "
    "running builds locally and working through open issues before attempting contributions, the same "
    "pattern I used when adopting LangGraph and Azure Durable Functions for production workloads.",

    "I would welcome the opportunity to discuss how my experience across TypeScript/GraphQL product "
    "delivery, async pipeline orchestration and full-stack platform engineering could contribute to "
    "Railway's product team.",
]


# ---------------------------------------------------------------------------
# Generate all ten files
# ---------------------------------------------------------------------------

base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

xero_cv_path         = os.path.join(base, "cv",           "Vrandesh_Bandikatti_CV_Xero_AI_Workflows.docx")
xero_cl_path         = os.path.join(base, "cover_letters", "Vrandesh_Bandikatti_Cover_Letter_Xero.docx")
cba_cv_path          = os.path.join(base, "cv",           "Vrandesh_Bandikatti_CV_CBA_Principal_SE.docx")
cba_cl_path          = os.path.join(base, "cover_letters", "Vrandesh_Bandikatti_Cover_Letter_CBA.docx")
preacta_cv_path      = os.path.join(base, "cv",           "Vrandesh_Bandikatti_CV_Preacta_AI_Systems_Developer.docx")
preacta_cl_path      = os.path.join(base, "cover_letters", "Vrandesh_Bandikatti_Cover_Letter_Preacta.docx")
igodirect_cv_path    = os.path.join(base, "cv",           "Vrandesh_Bandikatti_CV_iGoDirect_Principal_AI_Engineer.docx")
igodirect_cl_path    = os.path.join(base, "cover_letters", "Vrandesh_Bandikatti_Cover_Letter_iGoDirect.docx")
railway_cv_path      = os.path.join(base, "cv",           "Vrandesh_Bandikatti_CV_Railway_Senior_FullStack_Engineer.docx")
railway_cl_path      = os.path.join(base, "cover_letters", "Vrandesh_Bandikatti_Cover_Letter_Railway.docx")

build_xero_cv().save(xero_cv_path)
print(f"Saved: {xero_cv_path}")

build_cba_cv().save(cba_cv_path)
print(f"Saved: {cba_cv_path}")

build_cover(XERO_BEFORE, XERO_BULLETS, XERO_AFTER, "Vrandesh Bandikatti").save(xero_cl_path)
print(f"Saved: {xero_cl_path}")

build_cover(CBA_BEFORE, CBA_BULLETS, CBA_AFTER, "Vrandesh Bandikatti").save(cba_cl_path)
print(f"Saved: {cba_cl_path}")

build_preacta_cv().save(preacta_cv_path)
print(f"Saved: {preacta_cv_path}")

build_cover(PREACTA_BEFORE, PREACTA_BULLETS, PREACTA_AFTER, "Vrandesh Bandikatti").save(preacta_cl_path)
print(f"Saved: {preacta_cl_path}")

build_igodirect_cv().save(igodirect_cv_path)
print(f"Saved: {igodirect_cv_path}")

build_cover(IGODIRECT_BEFORE, IGODIRECT_BULLETS, IGODIRECT_AFTER, "Vrandesh Bandikatti").save(igodirect_cl_path)
print(f"Saved: {igodirect_cl_path}")

build_railway_cv().save(railway_cv_path)
print(f"Saved: {railway_cv_path}")

build_cover(RAILWAY_BEFORE, RAILWAY_BULLETS, RAILWAY_AFTER, "Vrandesh Bandikatti").save(railway_cl_path)
print(f"Saved: {railway_cl_path}")

print("\nAll ten documents generated successfully.")
