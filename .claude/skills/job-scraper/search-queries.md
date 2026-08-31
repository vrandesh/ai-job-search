# Search Queries for Job Scraper

<!-- Populated by /setup for Vrandesh Bandikatti — Melbourne, Australia -->

## Installed portal CLIs (primary for `/scrape`)

`/scrape` discovers every portal skill under `.agents/skills/*/SKILL.md` and runs its CLI first. Shipped country-agnostic CLIs include `linkedin-search` and `freehire-search`; any skill you add with `/add-portal` is included the same way. You do **not** need a matching `site:` line below for those CLIs to run.

The `site:` query templates in this file are the **WebSearch fallback** — for portals without a CLI, company career pages, or when a CLI fails.

**Language scope:** All queries below are in English (CV language and primary market language). The Language Gate in `04-job-evaluation.md` handles any postings with non-English language requirements.

## Search Sites

Primary (Australian job boards):
- **seek.com.au** — Australia's largest general job board
- **linkedin.com/jobs** — LinkedIn job listings (filter: Melbourne / Australia); also covered by `linkedin-search` CLI
- **au.indeed.com** — Indeed Australia
- **gradconnection.com.au** — skip (graduate-only, excluded by deal-breaker)

Secondary (company career pages via Google):
- Direct Google searches with `site:` filters for target companies

## Query Categories

### Priority 1: AI/Agentic Engineering + Cloud Architecture
*Strongest and most desired direction — production AI systems on cloud platforms*

```
site:seek.com.au "AI Engineer" Melbourne
site:seek.com.au "AI Solutions Architect" Melbourne
site:seek.com.au "Agentic AI" Melbourne
site:seek.com.au "Cloud AI Engineer" Melbourne
site:seek.com.au "AI Platform Engineer" Melbourne
site:au.indeed.com "AI Engineer" "Azure OR AWS OR GCP" Melbourne
site:linkedin.com/jobs "AI Engineer" Melbourne Australia
site:linkedin.com/jobs "Cloud Solution Architect" "AI" Melbourne
site:seek.com.au "Machine Learning Engineer" Melbourne
site:seek.com.au "LLM Engineer" Melbourne
site:seek.com.au "GenAI Engineer" Melbourne
```

### Priority 2: Senior Full Stack Engineering
*Core strength — enterprise full-stack platform delivery*

```
site:seek.com.au "Senior Full Stack Developer" Melbourne
site:seek.com.au "Senior Software Engineer" Melbourne
site:seek.com.au "Staff Engineer" Melbourne
site:seek.com.au "Principal Engineer" Melbourne
site:seek.com.au "Senior Full Stack Engineer" "Python OR Django OR FastAPI" Melbourne
site:au.indeed.com "Senior Software Engineer" "Python" "Azure OR AWS" Melbourne
site:linkedin.com/jobs "Senior Full Stack Developer" Melbourne Australia
site:linkedin.com/jobs "Staff Engineer" Melbourne Australia
site:seek.com.au "Senior Python Developer" Melbourne
site:seek.com.au "Senior Backend Engineer" "Python" Melbourne
```

### Priority 3: Cloud Architecture + Solutions Architecture
*Architecture ownership with hands-on delivery*

```
site:seek.com.au "Cloud Solution Architect" Melbourne
site:seek.com.au "Solutions Architect" Melbourne
site:seek.com.au "Cloud Architect" "Azure OR AWS" Melbourne
site:seek.com.au "Enterprise Architect" Melbourne
site:seek.com.au "Technical Architect" "Python OR Full Stack" Melbourne
site:au.indeed.com "Solutions Architect" "Azure" Melbourne
site:linkedin.com/jobs "Solutions Architect" Melbourne Australia
site:linkedin.com/jobs "Cloud Architect" "Azure" Melbourne
```

### Priority 4: Technical Leadership
*Leadership roles that keep engineering accountability*

```
site:seek.com.au "Technical Lead" Melbourne
site:seek.com.au "Engineering Lead" Melbourne
site:seek.com.au "Head of Engineering" Melbourne
site:seek.com.au "Lead Engineer" "AI OR Cloud OR Full Stack" Melbourne
site:seek.com.au "Engineering Manager" "hands-on OR technical" Melbourne
site:linkedin.com/jobs "Technical Lead" "Python OR Azure" Melbourne Australia
site:linkedin.com/jobs "Head of Engineering" Melbourne Australia
```

### Priority 5: Domain-specific (Financial Services + AI)
*Sector-specific searches leveraging FinTech/financial services background*

```
site:seek.com.au "AI Engineer" "fintech OR financial OR payments" Melbourne
site:seek.com.au "Senior Developer" "Azure" "financial services" Melbourne
site:seek.com.au "Platform Engineer" "Azure OR Python" Melbourne
site:seek.com.au "Data Platform Engineer" "Microsoft Fabric OR Databricks" Melbourne
site:linkedin.com/jobs "AI Engineer" "financial services" Melbourne
site:seek.com.au "Full Stack Developer" "Azure OpenAI OR LangChain OR LangGraph" Melbourne
```

### Target Company Searches
*Monitor specific employers for openings*

```
site:atlassian.com/company/careers "engineer" OR "architect"
site:canva.com/careers "engineer" OR "architect"
site:seek.com.au/companies "Deloitte Digital" engineer Melbourne
site:seek.com.au/companies "Thoughtworks" Melbourne
site:seek.com.au/companies "Xero" engineer Melbourne
site:seek.com.au/companies "REA Group" engineer Melbourne
site:seek.com.au/companies "Afterpay" OR "Block" engineer Melbourne
site:microsoft.com/en-au/jobs "engineer" OR "architect" Melbourne
```

## Location Filter

When evaluating results, verify the job location fits the following tiers:

- **Ideal:** Melbourne CBD, inner suburbs (3-15km radius), hybrid with ≥3 days WFH
- **Acceptable:** Greater Melbourne metro, full remote (anywhere in Australia)
- **Borderline:** Occasional interstate presence required (Sydney/Brisbane/Perth/Adelaide travel is fine; flag if >2 days/month average)
- **FAIL (deal-breaker):** Relocation required outside Melbourne

## Salary Filter

Flag any role where the advertised or estimated salary is below **AUD $150,000 base + superannuation**. Do not auto-exclude — flag for the user's review in case the range is negotiable or the posting is understated.

## Language Filter

Working languages and levels are in CLAUDE.md's Languages table. Apply `04-job-evaluation.md`'s Language Gate: a posting requiring a language not declared (e.g. Mandarin required) is a hard FAIL; a posting requiring a higher level than declared in a language that is listed is a FLAG (not excluded). Postings simply *written* in a language not worked in, that don't require it on the job, are fine.

## Date Filter

Only include jobs posted within the last 14 days, or with an application deadline that has not yet passed. If a posting date cannot be determined, include it but flag as "date unknown".

## Adapting Queries

If the user specifies a focus area, select queries from the matching category and also generate 2-3 custom queries for that focus. For example:
- `/scrape AI` → Priority 1 queries + custom agentic/LLM-specific queries
- `/scrape fintech` → Priority 5 queries + target company searches for ANZ/NAB/Xero/Afterpay
- `/scrape architecture` → Priority 3 queries + any architect openings at target companies
