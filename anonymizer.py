import re
import time
from typing import List, Dict, Tuple, Set, Optional
from lxml import etree
import spacy
from presidio_analyzer import AnalyzerEngine
from presidio_anonymizer import AnonymizerEngine
from presidio_anonymizer.entities import OperatorConfig


CLINICAL_TRIAL_PATTERNS = {
    "NCT_ID": r"\bNCT\d{8}\b",
    "EUDRACT_ID": r"\b\d{4}-\d{6}-\d{2}\b",
    "EU_CT_ID": r"\bEU[-\s]?CT\s*\d{4}[-\s]?\d{4,6}[-\s]?\d{2}[-\s]?\d{2}\b",
    "ISRCTN_ID": r"\bISRCTN\d{8}\b",
    "JRCT_ID": r"\bjRCT[a-zA-Z]?\d{7,10}\b",
    "CTRI_ID": r"\bCTRI[/\-]\d{4}[/\-]\d{2,3}[/\-]\d{5,6}\b",
    "ANZCTR_ID": r"\bACTRN\d{14}\b",
    "CHICTR_ID": r"\bChiCTR[-]?(?:[A-Z]{2,4}[-]?)?\d{7,10}\b",
    "DRKS_ID": r"\bDRKS\d{8}\b",
    "IRCT_ID": r"\bIRCT\d{14,}\b",
    "UMIN_ID": r"\bUMIN\d{9}\b",
    "KCT_ID": r"\bKCT\d{7}\b",
    "PROTOCOL_ID": r"\b(?:Protocol(?:\s+ID)?|Study)[\s:]*[A-Z]{2,5}[-\s]?\d{3,5}\b",
    "SUBJECT_ID": r"\b(?:Subject|Patient|SUBJ)[\s#:-]*\d{3,}(?:[-]\d{3})?\b",
    "SITE_NUMBER": r"\b(?:Site|Center)[\s#:-]*\d{2,4}\b",
    "LOT_BATCH_ID": r"\b(?:Lot|Batch|Kit)[\s:]*[A-Z0-9]{2,}[-]?[A-Z0-9]{2,}\b|(?:Sample(?:\s+ID))[\s:]*[A-Z0-9]{2,}[-]?[A-Z0-9]{2,}\b",
    "STRUCTURED_ID": r"\b[A-Z]{2,6}[-_/]\d{2,8}\b",
    "STRUCTURED_ID_NUM_FIRST": r"\b\d{2,6}[-_/][A-Z]{2,6}\b",
    "PRODUCT_CODE": r"\b[A-Z]{2,5}[-][A-Z]{2,5}[-][A-Z0-9]{2,5}(?:[-][A-Z0-9]{1,3})?\b",
    "NIF_CIF_ES": r"\b(?:NIF|CIF|DNI|NIE)[\s:.-]*[A-Z]?\d{7,8}[-]?[A-Z]?\b",
    "NIF_STANDALONE": r"\b[A-Z]\d{7,8}[-]?[A-Z]?\b",
    "DNI_STANDALONE": r"\b\d{8}[-]?[A-Za-z]\b",
    "LONG_NUMBER_ID": r"\b\d{3,}(?:[-./]\d{2,})+\b|\b\d{6,}[A-Za-z]*\b",
    "LICENSE_PLATE": r"\b\d{4}[\s-]?[A-Z]{2,3}\b|\b[A-Z]{1,2}[\s-]?\d{4}[\s-]?[A-Z]{2,3}\b",
    "IBAN_CODE": r"\b[A-Z]{2}\d{2}[\s]?\d{4}[\s]?\d{4}[\s]?\d{4}[\s]?\d{4}[\s]?\d{0,4}\b",
    "SPACED_DIGIT_SEQ": r"\b\d{2,}(?:\s\d{2,}){2,}\b",
}

CASE_SENSITIVE_PATTERNS = {
    "ALPHANUMERIC_CODE": r"\b[A-Z]{1,}\d{6,}[A-Za-z0-9]*\b|\b[A-Z]{3,}\d{5,}\b",
}

SAFE_ACRONYMS = {
    "DNA", "RNA", "mRNA", "cDNA", "siRNA", "miRNA", "tRNA", "rRNA",
    "PCR", "qPCR", "ELISA", "HPLC", "GLP", "GMP", "GCP", "ICH",
    "AUC", "BMI", "ECG", "EKG", "MRI", "PET", "CBC", "WBC", "RBC",
    "HIV", "AIDS", "HBV", "HCV", "HPV", "CMV", "EBV", "RSV", "HSV",
    "SAE", "ADE", "ADR", "SUSAR", "DLT", "MTD", "ORR", "DOR",
    "PFS", "DFS", "RFS", "TTR", "TTP", "HRR",
    "RECIST", "WHO", "FDA", "EMA", "TGA", "ICF", "SAP",
    "ITT", "PPS", "FAS", "SOC", "MedDRA", "CTCAE",
    "BID", "TID", "QID", "QHS", "PRN",
    "IND", "NDA", "BLA", "MAA", "CTA", "CSR",
    "SOC", "IEC", "IRB", "DMC", "CRF", "eCRF",
    "PII", "PHI", "HIPAA", "GDPR",
    "PDF", "XML", "HTML", "CSS", "SQL", "API", "URL", "HTTP", "HTTPS",
    "UTC", "GMT", "ISO", "ICD", "CPT", "ATC",
    "QOL", "VAS", "NYHA", "ECOG", "FACT",
    "NOT", "AND", "FOR", "THE", "HAS", "HAD", "ARE", "WAS", "HER",
    "HIS", "OUR", "ALL", "BUT", "NOR", "YET", "ANY", "CAN", "MAY",
    "USE", "PER", "VIA", "SET", "GET", "PUT", "RUN", "END",
    "USA", "EUR", "GBP", "USD",
    "NOTE", "ALSO", "MUST", "WILL", "DOES", "EACH", "BOTH", "ONLY",
    "WHEN", "THEN", "THAN", "SUCH", "THAT", "THIS", "WHAT", "WITH",
    "FROM", "INTO", "OVER", "SOME", "BEEN", "HAVE", "WERE", "HERE",
    "USED", "ONCE", "TYPE", "FORM", "PART", "SIDE", "SAME", "LAST",
    "NEXT", "FULL", "MADE", "CASE", "MORE", "MOST", "LESS", "VERY",
    "WELL", "JUST", "LIKE", "ALSO", "EVEN", "BACK", "LONG", "HIGH",
    "STILL", "AFTER", "ABOUT", "ABOVE", "BELOW", "UNDER", "OTHER",
    "EVERY", "FIRST", "FOUND", "GIVEN", "BASED", "USING", "WHILE",
    "THESE", "THOSE", "WHICH", "WHERE", "THERE", "THEIR", "SHALL",
    "WOULD", "COULD", "SHOULD", "BEING", "DURING", "BEFORE",
    "BETWEEN", "HOWEVER", "WITHOUT", "THROUGH", "AGAINST", "BECAUSE",
    "ANOTHER", "WHETHER", "WITHIN", "EITHER", "NEITHER",
    "TOTAL", "TABLE", "VALUE", "LEVEL", "GROUP", "STUDY", "TRIAL",
    "PHASE", "VISIT", "DAILY", "PRIOR", "AFTER", "EARLY", "FINAL",
    "MAJOR", "MINOR", "UPPER", "LOWER", "LOCAL", "POINT", "RANGE",
    "SCORE", "RATIO", "ORGAN", "BLOOD", "LIVER", "RENAL", "VIRAL",
}

BIOMEDICAL_PATTERNS = {
    "DRUG_SUFFIX": r"\b[A-Z][a-z]+(?:mab|nib|zole|pril|statin|vir|cillin|cycline|sartan|olol|dipine|mycin|floxacin|fenac|afil|dronate|lukast|prazole|setron|triptan|vastatin|zodone)\b",
    "CHEMICAL_COMPOUND": r"\b\d*[A-Z]?[-]?(?:methyl|ethyl|propyl|butyl|acetyl|hydroxy|chloro|fluoro|amino|phenyl|benzyl)[a-z]+\b",
    "DNA_RNA": r"\b(?:DNA|RNA|mRNA|cDNA|siRNA|miRNA|tRNA|rRNA)\b",
    "PROBEMIX_PRODUCT": r"\bProbemix\s+[A-Z]?\d{2,4}\s*[A-Z0-9]+\b",
    "TRADEMARK_NAME": r"\b[A-Za-zÀ-ÿ][A-Za-zÀ-ÿ0-9\-]*\s*[®™](?!\w)",
    "COPYRIGHT_NAME": r"[©]\s*(?:\d{4}\s+)?[A-Za-zÀ-ÿ][A-Za-zÀ-ÿ0-9\-]+(?!\w)",
}

INVALID_POS_FOR_REDACTION = {"VERB", "ADV", "ADJ", "DET", "PRON", "ADP", "CCONJ", "SCONJ", "AUX", "PART", "INTJ"}

# Stopwords funcionales (solo palabras gramaticales comunes) - v10.0
STOPWORDS_FUNCTIONAL_EN = {
    "the", "a", "an", "this", "that", "these", "those",
    "of", "in", "on", "at", "to", "from", "by", "for", "with", "as", "into", "over", "under", "between", "through",
    "and", "or", "but", "nor", "so", "yet", "if", "because", "although", "while",
    "i", "you", "he", "she", "it", "we", "they", "me", "him", "her", "us", "them",
    "my", "your", "his", "its", "our", "their", "mine", "yours", "hers", "ours", "theirs",
    "is", "are", "was", "were", "be", "been", "being",
    "have", "has", "had", "do", "does", "did",
    "will", "would", "shall", "should", "can", "could", "may", "might", "must",
    "all", "any", "some", "each", "every", "no", "not", "none", "both", "either", "neither",
    "very", "too", "also", "just", "only", "even", "still", "already"
}

STOPWORDS_FUNCTIONAL_ES = {
    "el", "la", "los", "las", "un", "una", "unos", "unas",
    "este", "esta", "estos", "estas", "ese", "esa", "esos", "esas",
    "mi", "mis", "tu", "tus", "su", "sus", "nuestro", "nuestra", "nuestros", "nuestras",
    "de", "del", "en", "a", "por", "para", "con", "sin", "sobre", "entre", "desde", "hasta", "hacia",
    "y", "o", "pero", "ni", "aunque", "porque", "si", "mientras",
    "yo", "tú", "él", "ella", "nosotros", "nosotras", "vosotros", "vosotras", "ellos", "ellas",
    "me", "te", "se", "nos", "os", "lo", "le", "les",
    "es", "son", "era", "eran", "fue", "ser", "estar",
    "he", "has", "ha", "hemos", "han", "haber",
    "todo", "toda", "todos", "todas", "algún", "alguna", "algunos", "algunas",
    "ningún", "ninguna", "ningunos", "ningunas", "no"
}

# Sustantivos comunes que suelen aparecer capitalizados por estilo (falsos positivos frecuentes)
COMMON_SINGLETON_BLOCK_EN = {
    "system", "study", "protocol", "trial", "switch", "data", "process",
    "document", "file", "report", "table", "section", "chapter", "version",
    "note", "warning", "caution", "important", "example", "figure", "appendix",
    "patient", "treatment", "medication", "disease", "adverse", "event", "consent",
    "endpoint", "efficacy", "safety", "drug", "dose", "subject", "placebo",
    "sample", "analysis", "result", "outcome", "therapy", "diagnosis", "symptom",
    "procedure", "baseline", "screening", "randomization", "cohort", "arm",
    "visit", "investigator", "sponsor", "monitor", "deviation", "amendment",
    "inclusion", "exclusion", "enrollment", "withdrawal", "discontinuation",
    "toxicity", "tolerability", "pharmacokinetics", "bioavailability",
    "compliance", "regulation", "submission", "approval", "authorization",
    "agreement", "contract", "liability", "indemnification", "confidentiality",
    "disclosure", "obligation", "jurisdiction", "arbitration", "governance",
    "oversight", "audit", "inspection", "certificate", "license", "patent",
    "trademark", "regulatory", "guideline", "directive", "statute", "provision",
    "clause", "warranty", "termination", "notification", "declaration",
    "assessment", "evaluation", "monitoring", "surveillance", "pharmacovigilance",
    "response", "infection", "inflammation", "mortality", "morbidity", "survival",
    "incidence", "prevalence", "population", "intervention", "comparator",
}

COMMON_SINGLETON_BLOCK_ES = {
    "sistema", "estudio", "protocolo", "ensayo", "cambio", "interruptor", "datos", "proceso",
    "documento", "archivo", "informe", "tabla", "sección", "capítulo", "versión",
    "nota", "advertencia", "precaución", "importante", "ejemplo", "figura", "apéndice",
    "paciente", "tratamiento", "medicación", "enfermedad", "adverso", "evento",
    "consentimiento", "eficacia", "seguridad", "fármaco", "dosis", "sujeto", "placebo",
    "muestra", "análisis", "resultado", "terapia", "diagnóstico", "síntoma",
    "procedimiento", "aleatorización", "cohorte", "brazo", "visita",
    "investigador", "promotor", "monitor", "desviación", "enmienda",
    "inclusión", "exclusión", "reclutamiento", "retirada", "discontinuación",
    "toxicidad", "tolerabilidad", "farmacocinética", "biodisponibilidad",
    "cumplimiento", "regulación", "presentación", "aprobación", "autorización",
    "acuerdo", "contrato", "responsabilidad", "indemnización", "confidencialidad",
    "divulgación", "obligación", "jurisdicción", "arbitraje", "gobernanza",
    "supervisión", "auditoría", "inspección", "certificado", "licencia", "patente",
    "regulatorio", "directriz", "directiva", "estatuto", "disposición",
    "cláusula", "garantía", "rescisión", "notificación", "declaración",
    "evaluación", "monitorización", "vigilancia", "farmacovigilancia",
    "respuesta", "infección", "inflamación", "mortalidad", "morbilidad",
    "supervivencia", "incidencia", "prevalencia", "población", "intervención",
    "comparador",
}

# NER "fuerte" que permitimos (si hay NER, no bloqueamos por listas)
ALLOWED_NER_LABELS = {"PERSON", "ORG", "GPE", "LOC", "FAC", "PRODUCT"}

CLINICAL_ABBREVIATIONS_WITH_VALUE = [
    r"\b(?:CRN|MRN|PRN|IFU|HDD|SSN|ID|REF|NO)[-#:\s]+[A-Z0-9]{2,}[-]?\d{2,}\b",
    r"\b(?:Protocol|Study|Subject|Patient|Site|Center|Kit|Vial)[-\s]*(?:ID|No|Number|#)?[-:\s]*[A-Z0-9]{2,}[-/]?\d{2,}\b",
    r"\bSample[-\s]+(?:ID|No|Number|#)[-:\s]*[A-Z0-9]{2,}[-/]?\d{2,}\b",
]

PHARMA_COMPANY_PATTERNS = [
    r"\b(?:Pfizer|Novartis|Roche|Sanofi|Merck|MSD|AstraZeneca|GSK|GlaxoSmithKline|Johnson\s*&?\s*Johnson|J&J|AbbVie|Bristol[- ]?Myers[- ]?Squibb|BMS|Eli\s*Lilly|Lilly|Amgen|Gilead|Bayer|Novo\s*Nordisk|Takeda|Boehringer\s*Ingelheim|Boehringer|Biogen|Regeneron|Moderna|BioNTech)\b",
    r"\b(?:Teva|Viatris|CSL\s*Behring|CSL|Astellas|Daiichi\s*Sankyo|Otsuka|Eisai|Chugai|Kyowa\s*Kirin|Shionogi|Alexion|Vertex|Incyte|BioMarin|Alnylam|Seagen|Horizon|UCB|Ipsen|Lundbeck|Grifols|Ferring|Servier|Menarini|Chiesi|Almirall|Gr[üu]nenthal|Fresenius|Bausch\s*Health|Jazz\s*Pharma|Hikma|Lupin|Cipla|Zydus|Sun\s*Pharma|Dr\.?\s*Reddy|Celltrion|Samsung\s*Bioepis)\b",
]

CRO_PATTERNS = [
    r"\b(?:IQVIA|Covance|PPD|ICON|Syneos\s*Health|Syneos|Parexel|PRA\s*Health|Quintiles|Labcorp|LabCorp|Charles\s*River)\b",
    r"\b(?:Medpace|Fortrea|WuXi\s*AppTec|WuXi|Caidya|Clinipace|dMed|Novotech|Ergomed|Allucent|Veristat|Synteract|CTI\s*Clinical|Inotiv|Velocity\s*Clinical|CMIC|EPS\s*International|Worldwide\s*Clinical\s*Trials|Thermo\s*Fisher|Q2\s*Solutions|Frontage|Celerion|SGS|Almac|Eurofins|Quest\s*Diagnostics)\b",
]

BLOCKBUSTER_DRUGS = [
    r"\b(?:Keytruda|pembrolizumab|Ozempic|Wegovy|semaglutide|Dupixent|dupilumab|Biktarvy|Eliquis|apixaban|Skyrizi|risankizumab|Mounjaro|Zepbound|tirzepatide|Darzalex|daratumumab|Stelara|ustekinumab|Trikafta|Kaftrio|Alyftrek)\b",
    r"\b(?:Opdivo|nivolumab|Humira|adalimumab|Eylea|aflibercept|Rinvoq|upadacitinib|Xtandi|enzalutamide|Imbruvica|ibrutinib|Tagrisso|osimertinib|Revlimid|lenalidomide|Tremfya|guselkumab|Cosentyx|secukinumab|Entresto|sacubitril)\b",
    r"\b(?:Xarelto|rivaroxaban|Trulicity|dulaglutide|Jardiance|empagliflozin|Farxiga|dapagliflozin|Ocrevus|ocrelizumab|Vyndaqel|Vyndamax|tafamidis|Paxlovid|nirmatrelvir|Comirnaty|Spikevax|Prevnar|Shingrix|Gardasil)\b",
    r"\b(?:Kisunla|donanemab|Leqembi|lecanemab|Rezdiffra|resmetirom|Datroway|Pluvicto|Enhertu|Padcev|Tecvayli|Talvey|Columvi|glofitamab|Epkinly|epcoritamab|Elahere|mirvetuximab)\b",
]

CLINICAL_TECH_PLATFORMS = [
    r"\b(?:Veeva|Veeva\s*CTMS|Veeva\s*Vault|Veeva\s*eTMF|Veeva\s*RTSM|Veeva\s*EDC|SiteVault|Veeva\s*CRM|Veeva\s*Clinical\s*Platform)\b",
    r"\b(?:Medidata|Medidata\s*Rave|Rave\s*EDC|Medidata\s*CTMS|Medidata\s*eTMF|Clinical\s*Data\s*Studio|Medidata\s*eConsent|Medidata\s*eCOA)\b",
    r"\b(?:Oracle\s*Clinical|Oracle\s*Clinical\s*One|Oracle\s*InForm|Oracle\s*CTMS|Oracle\s*Safety|Oracle\s*Health\s*Sciences|Oracle\s*Argus)\b",
    r"\b(?:IQVIA\s*RTSM|IQVIA\s*EDC|IQVIA\s*CTMS|IQVIA\s*eTMF|IQVIA\s*OCE|Orchestrated\s*Clinical\s*Trials|IQVIA\s*OneKey|Citeline|Trialtrove|Pharmaprojects)\b",
    r"\b(?:Clario|Signant\s*Health|Signant|Castor\s*EDC|Castor|OpenClinica|REDCap|Clinical\s*Conductor|Advarra|WCG|WIRB|Copernicus)\b",
    r"\b(?:Bioclinica|ERT|PHT|Exponent|Trial\s*Interactive|TransPerfect|Lionbridge|RWS|CSOFT|Welocalize)\b",
]

CENTRAL_LABS = [
    r"\b(?:Q2\s*Solutions|Q\s*Squared|Labcorp\s*Drug\s*Development|Labcorp\s*Central\s*Labs|Covance\s*Central\s*Labs)\b",
    r"\b(?:ICON\s*Central\s*Labs|PPD\s*Laboratories|Eurofins\s*Central\s*Lab|ACM\s*Global\s*Laboratories|BioAgilytix|Frontage\s*Labs|WuXi\s*Clinical)\b",
    r"\b(?:BARC\s*Global|MLM\s*Medical\s*Labs|Sonic\s*Clinical\s*Trials|Clinical\s*Reference\s*Laboratory|CRL)\b",
]

LAB_PRODUCT_PATTERNS = [
    r"\b[A-Z][a-z]+(?:amp|pure|zard)\s+(?:DNA|RNA|Genomic)?\s*(?:mini|midi|maxi|LS)?\s*(?:kit|purification)?\b",
    r"\b[A-Z][a-z]*(?:mix|lyser)(?:\.[A-Za-z]+)?\b",
    r"\b\d{4,5}[-]L\d{4,6}\b",
    r"\bP\d{3}[-]?\d{0,3}[A-Z]?\b",
    r"\b(?:RUO|IVD|CE[-]?IVD)\b",
    r"\b[a-z]+(?:MLPA|LPA|NER|PCR|DNA|RNA)\b",
]

ACRONYM_PATTERNS = [
    r"\b[A-Z]{2,5}[-]?[A-Z]{0,3}\s+(?:Select|Plus|Pro|Kit|System|Assay|Panel|Array)\b",
]

ACRONYM_ALLCAPS_RE = re.compile(r'\b[A-Z]{3,}\b')

SOFTWARE_PRODUCT_PATTERNS = [
    r"\b(?:Microsoft|Windows)\s+(?:Windows\s+)?(?:Server|Azure|Office|Teams|SharePoint|OneDrive|Dynamics|Exchange|SQL\s*Server|Visual\s*Studio|Power\s*BI|Outlook|Word|Excel|PowerPoint|Access|Publisher|OneNote|Project|Visio|Intune|Defender|Sentinel|Entra|Copilot|365|2012|2016|2019|2022|2025)(?:\s+(?:R2|Standard|Enterprise|Datacenter|Professional|Home|Pro|Ultimate|Education|Business|Premium|\d+(?:\.\d+)*))?\b",
    r"\b(?:Oracle|SAP|Salesforce|ServiceNow|Workday|Kronos|ADP|Concur|Ariba|SuccessFactors|Tableau|Snowflake|Databricks|Splunk|Elastic|MongoDB|Redis|PostgreSQL|MySQL|MariaDB|Teradata|Informatica|Talend|MicroStrategy|Qlik|Domo|Looker|Alteryx)\b",
    r"\b(?:VMware|Citrix|Red\s*Hat|SUSE|Canonical|Ubuntu|Debian|CentOS|Rocky\s*Linux|AlmaLinux|Fedora|openSUSE)(?:\s+[A-Za-z]+(?:\s+[A-Za-z0-9]+)*)?\b",
    r"\b(?:IBM|Dell|HP|HPE|Hewlett[-\s]?Packard|Cisco|Juniper|Arista|Palo\s*Alto|Fortinet|F5|Netscaler|Checkpoint|Zscaler|Crowdstrike|SentinelOne|Carbon\s*Black|Symantec|McAfee|Trend\s*Micro|Sophos|ESET|Kaspersky|Bitdefender|Avast|Norton|Malwarebytes)\b",
    r"\b(?:AWS|Amazon\s*Web\s*Services|Amazon|Google\s*Cloud|GCP|Google|Azure|IBM\s*Cloud|Alibaba\s*Cloud|Alibaba|DigitalOcean|Linode|Vultr|Heroku|Vercel|Netlify|Cloudflare|Akamai|Fastly)\b",
    r"\b(?:SpaceX|Tesla|Neuralink|Boring\s*Company|OpenAI|Anthropic|DeepMind|Meta|Facebook|Instagram|WhatsApp|TikTok|ByteDance|Twitter|X\s*Corp|Snapchat|Pinterest|LinkedIn|Reddit|Discord|Telegram|Signal)\b",
    r"\b(?:Netflix|Spotify|Disney|Hulu|HBO|Warner\s*Bros|Universal|Paramount|Sony\s*Pictures|MGM|Lionsgate|DreamWorks|Pixar|Lucasfilm|Marvel|DC\s*Comics|Activision|Blizzard|EA|Ubisoft|Epic\s*Games|Valve|Riot\s*Games|Rockstar|Bethesda|CD\s*Projekt)\b",
    r"\b(?:Uber|Lyft|Airbnb|Booking|Expedia|TripAdvisor|Yelp|DoorDash|Instacart|Grubhub|Postmates|Deliveroo|Just\s*Eat|Rappi|Glovo|iFood)\b",
    r"\b(?:PayPal|Stripe|Square|Block|Adyen|Klarna|Affirm|Afterpay|Venmo|Cash\s*App|Revolut|N26|Monzo|Chime|Robinhood|Coinbase|Binance|Kraken|FTX|Gemini)\b",
    r"\b(?:Kubernetes|Docker|Podman|OpenShift|Rancher|Tanzu|EKS|AKS|GKE|Helm|Terraform|Ansible|Puppet|Chef|SaltStack|Vagrant|Packer)\b",
    r"\b(?:Jenkins|GitLab|GitHub|Bitbucket|Azure\s*DevOps|CircleCI|Travis\s*CI|TeamCity|Bamboo|Octopus\s*Deploy|ArgoCD|Flux|Spinnaker)\b",
    r"\b(?:Jira|Confluence|Trello|Asana|Monday|Notion|Slack|Zoom|Webex|GoToMeeting|BlueJeans|RingCentral|Twilio|Vonage|Genesys|Five9|NICE|Avaya|Mitel)\b",
    r"\b(?:Adobe|Acrobat|Photoshop|Illustrator|InDesign|Premiere|After\s*Effects|XD|Figma|Sketch|InVision|Canva|Miro|Lucidchart|Draw\.io|Visio)\b",
    r"\b(?:AutoCAD|SolidWorks|CATIA|NX|Creo|Inventor|Revit|ArchiCAD|SketchUp|Rhino|Blender|Maya|3ds\s*Max|Cinema\s*4D|Houdini|ZBrush)\b",
    r"\b(?:NVIDIA|AMD|Intel|Qualcomm|Broadcom|Texas\s*Instruments|NXP|Infineon|STMicroelectronics|Microchip|Xilinx|Altera|Lattice)\b",
    r"\b(?:Apple|iPhone|iPad|MacBook|iMac|Mac\s*Pro|Mac\s*Mini|Apple\s*Watch|AirPods|HomePod|Apple\s*TV|macOS|iOS|iPadOS|watchOS|tvOS|Safari|Xcode|Swift|Objective-C)\b",
    r"\b(?:Dell|Lenovo|ASUS|Acer|HP|Hewlett[-\s]?Packard|Sony|Toshiba|Fujitsu|Panasonic|LG|Samsung|Huawei|Xiaomi|OPPO|Vivo|OnePlus|Motorola|Nokia|Ericsson|ZTE|HTC|BlackBerry|Palm)\b",
    r"\b(?:ThinkPad|ThinkCentre|ThinkStation|IdeaPad|IdeaCentre|Legion|Yoga|ROG|ZenBook|VivoBook|TUF|Strix|ProBook|EliteBook|ZBook|Latitude|Precision|OptiPlex|PowerEdge|Inspiron|XPS|Alienware|Vostro)\b",
    r"\b(?:PlayStation|Xbox|Nintendo\s+Switch|Steam\s*Deck|Oculus|Quest|HoloLens|Kindle|Fire\s*TV|Roku|Chromecast|Apple\s*TV|Sonos|Bose|JBL|Harman|Bang\s*&\s*Olufsen|Sennheiser|Audio[-\s]?Technica)\b",
    r"\b(?:Canon|Nikon|Sony|Fujifilm|Olympus|Panasonic|Leica|Hasselblad|GoPro|DJI|Garmin|Fitbit|Polar|Suunto|Wahoo|Zwift)\b",
    r"\b(?:SAP\s*S/4HANA|SAP\s*ECC|SAP\s*BW|SAP\s*HANA|SAP\s*Fiori|SAP\s*Ariba|SAP\s*Concur|SAP\s*SuccessFactors|SAP\s*Hybris|SAP\s*C/4HANA)\b",
    r"\b(?:Epic|Cerner|MEDITECH|Allscripts|athenahealth|eClinicalWorks|NextGen|Greenway|Aprima|DrChrono|Practice\s*Fusion|Kareo|CureMD)\b",
    r"\bVelocity(?:\s+\d+(?:\.\d+)*)?\s+GRID(?:\s+Server)?\b",
    r"\b(?:Gmail|Hotmail|Outlook\.com|Yahoo\s*Mail|Yahoo|AOL|ProtonMail|Proton\s*Mail|Zoho\s*Mail|Zoho|iCloud|Mail\.ru|Yandex\s*Mail|Yandex|GMX|Tutanota|FastMail|Mailfence|Runbox|Posteo|StartMail|Hushmail|Comcast|Comcast\.net|Xfinity|Verizon|AT&T|T-Mobile|Vodafone|Movistar|Orange|Telefonica|Telef[oó]nica)\b",
    r"\b(?:Aliyun|Alibaba\s*Mail|QQ\s*Mail|NetEase|163\.com|126\.com|Sina|Sohu|Baidu|Tencent|WeChat|Weibo|Douyin|Kuaishou|Bilibili|JD\.com|Pinduoduo|Meituan|Didi|ByteDance|Line|KakaoTalk|Kakao|Naver|Daum)\b",
    r"\b(?:Outlook|Teams|OneDrive|OneNote|SharePoint|Skype|Bing|Cortana|Edge|Internet\s*Explorer|Firefox|Chrome|Chromium|Opera|Brave|Vivaldi|Tor\s*Browser|DuckDuckGo|Ecosia)\b",
    r"\b(?:Dropbox|Box|WeTransfer|MediaFire|Mega|pCloud|Tresorit|Sync\.com|SpiderOak|Backblaze|Carbonite|CrashPlan|IDrive|Wasabi)\b",
    r"\b(?:Carestream|Kodak|Agfa|Fujifilm|Siemens\s*Healthineers|Siemens|Philips|GE\s*Healthcare|Medtronic|Stryker|Zimmer\s*Biomet|Becton\s*Dickinson|Baxter|Edwards\s*Lifesciences|Abbott|Boston\s*Scientific|Danaher|Hologic|Intuitive\s*Surgical|ResMed|Varian|Elekta)\b",
]

# Common instruction verbs to exclude from product name detection
INSTRUCTION_VERBS = {
    'install', 'click', 'open', 'select', 'choose', 'press', 'enter', 'type',
    'go', 'run', 'start', 'stop', 'close', 'save', 'load', 'download', 'upload',
    'create', 'delete', 'remove', 'add', 'edit', 'update', 'view', 'see',
    'check', 'find', 'search', 'browse', 'navigate', 'access', 'use', 'try',
    'test', 'verify', 'confirm', 'accept', 'decline', 'cancel', 'submit',
    'send', 'receive', 'copy', 'paste', 'cut', 'move', 'drag', 'drop',
    'configure', 'setup', 'enable', 'disable', 'activate', 'deactivate',
    'instalar', 'abrir', 'seleccionar', 'elegir', 'pulsar', 'escribir',
    'ejecutar', 'iniciar', 'parar', 'cerrar', 'guardar', 'cargar', 'descargar',
    'crear', 'eliminar', 'borrar', 'añadir', 'editar', 'actualizar', 'ver',
    'comprobar', 'buscar', 'navegar', 'acceder', 'usar', 'probar',
    'verificar', 'confirmar', 'aceptar', 'rechazar', 'cancelar', 'enviar',
}

URL_PATTERNS = [
    r"(?:https?://)?(?:www\.)?[a-zA-Z0-9][-a-zA-Z0-9]*(?:\.[a-zA-Z]{2,})+(?:/[^\s]*)?\b",
]

CLINICAL_URL_PATTERNS = [
    r"(?:https?://)?(?:www\.)?clinicaltrials\.gov/(?:ct2/show/|study/)?NCT\d{8}\b",
    r"(?:https?://)?(?:www\.)?eudract\.ema\.europa\.eu[^\s]*\b",
    r"(?:https?://)?(?:www\.)?euclinicaltrials\.eu[^\s]*\b",
    r"(?:https?://)?(?:www\.)?who\.int/(?:clinical-?trials|ictrp)[^\s]*\b",
    r"(?:https?://)?(?:www\.)?isrctn\.com/ISRCTN\d{8}\b",
    r"(?:https?://)?jrct\.niph\.go\.jp[^\s]*\b",
    r"(?:https?://)?ctri\.nic\.in[^\s]*\b",
    r"(?:https?://)?(?:www\.)?anzctr\.org\.au[^\s]*\b",
    r"(?:https?://)?drks\.de[^\s]*\b",
    r"(?:https?://)?(?:www\.)?chictr\.org\.cn[^\s]*\b",
    r"\bclinicaltrials\.gov\s*/?\s*(?:ct2/show/|study/)?NCT\d{8}\b",
    r"\b(?:see|visit|refer\s+to|available\s+at)?\s*(?:https?://)?clinicaltrials\.gov[^\s]*\b",
]

ADDRESS_PATTERNS = [
    r"\b\d{1,5}\s+[A-Z][a-zÀ-ÿ]+(?:\s+[A-Z][a-zÀ-ÿ]+)*\s+(?:Street|St\.?|Avenue|Ave\.?|Road|Rd\.?|Boulevard|Blvd\.?|Drive|Dr\.?|Lane|Ln\.?|Way|Court|Ct\.?|Place|Pl\.?|Parkway|Pkwy\.?|Highway|Hwy\.?)\b",
    r"\b\d{1,5}\s+[A-Z][a-zÀ-ÿ]+(?:straat|weg|laan|plein|gracht|kade|singel|dijk|straße|stra[ßs]e|platz|allee)\b",
    r"\b\d{4,5}\s*[A-Z]{2}\b,?\s*[A-ZÀ-ÿ][a-zÀ-ÿ]+",
    r"\b[A-Z][a-zÀ-ÿ]+\s+\d+[-/]?\d*\s*,\s*\d{4,5}\s*[A-Z]?[A-Z]?\s+[A-ZÀ-ÿ][a-zÀ-ÿ]+\b",
    r"\b[A-Z]{1,2}\d{1,2}\s*\d[A-Z]{2}\b",
    r"\b[A-Z]\d[A-Z]\s*\d[A-Z]\d\b",
    r"\b(?:Building|Bldg|Block|Tower|Floor|Room|Office)\s*[#:]?\s*[A-Z0-9]+\s*,\s*\d{1,5}\s+[A-Za-z]+\b",
]

INVESTIGATOR_PATTERNS = [
    r"\b(?:Dr|Prof|Professor|Doctor)\.?\s+[A-ZÀ-ÿ][a-zÀ-ÿ]+(?:\s+[A-ZÀ-ÿ][a-zÀ-ÿ]+){0,2}\b",
    r"\b[A-ZÀ-ÿ][a-zÀ-ÿ]+(?:\s+[A-ZÀ-ÿ][a-zÀ-ÿ]+){0,2},?\s*(?:MD|M\.D\.|PhD|Ph\.D\.|PharmD|Pharm\.D\.|DO|D\.O\.|RN|R\.N\.|NP|PA|MBBS|FRCP|FACP|FACS)\b",
    r"\b(?:Principal\s+Investigator|Sub-?Investigator|Study\s+Coordinator|Site\s+Director|Medical\s+Monitor|Study\s+Physician)[\s:]+[A-ZÀ-ÿ][a-zÀ-ÿ]+(?:\s+[A-ZÀ-ÿ][a-zÀ-ÿ]+){0,2}\b",
    r"\b(?:PI|Co-?PI|SI)[\s:]+[A-ZÀ-ÿ][a-zÀ-ÿ]+(?:\s+[A-ZÀ-ÿ][a-zÀ-ÿ]+){0,2}\b",
]

CLINICAL_DATE_PATTERNS = [
    r"\b[Qq][1-4]\s*[-/]?\s*(?:20[1-3]\d|FY\s*20[1-3]\d)\b",
    r"\b(?:FY|CY)\s*20[1-3]\d\b",
    r"\b(?:H1|H2|1H|2H)\s*[-/]?\s*20[1-3]\d\b",
    r"\b(?:Visit|Day|Week|Month|Year)\s*[-#:]?\s*\d{1,3}\b",
    r"\b(?:V|D|W|M)\d{1,3}\b(?=\s|$|[,;.])",
    r"\b(?:Screening|End\s*of\s*(?:Study|Treatment|Trial)|Follow[\s-]?up|Randomization|Enrollment)\s*(?:Visit)?\b",
    r"\bBaseline\s+(?:Visit|Date|Period|Assessment|Evaluation|Examination|Value|Measurement)\b",
    r"\b(?:At|From|Since|After|Before|Pre|Post)[\s-]?Baseline\b",
    r"\b(?:Study|Protocol|Amendment)\s*(?:Start|End|Initiation|Completion|Termination)\s*(?:Date)?[\s:]+\d{1,2}[-/]\d{1,2}[-/]20[1-3]\d\b",
]

LOT_BATCH_PATTERNS = [
    r"\b(?:LOT|Lot|BATCH|Batch)[-#:\s]*[A-Z0-9]{2,}[-/]?[A-Z0-9]{2,}[-/]?[A-Z0-9]*\b",
    r"\b(?:Kit|Vial|Bottle|Ampule|Ampoule)[-#:\s]*(?:ID|No|Number)?[-#:\s]*[A-Z0-9]{4,}\b",
    r"\bSample\s+(?:ID|No|Number)[-#:\s]*[A-Z0-9]{4,}\b",
    r"\b(?:Exp(?:iry)?|MFG|Manufacturing)[-:\s]*(?:Date)?[-:\s]*\d{2}[-/]\d{2}[-/]?\d{2,4}\b",
    r"\b[A-Z]{2,3}[-]?\d{4,8}[-]?[A-Z]?\d{0,4}\b(?=\s+(?:lot|batch|expir))",
]

SITE_CODE_PATTERNS = [
    r"\b\d{3}[-]?[A-Z]{2}[-]?[A-ZÀ-ÿ][a-zÀ-ÿ]+\b",
    r"\b(?:Site|Centro|Centre|Center|Sitio)[-#:\s]*\d{2,4}[-]?[A-Z]{0,3}\b",
    r"\b[A-Z]{2,3}[-]?\d{3,4}[-]?(?:Site|Centro|Centre)?\b",
    r"\bS\d{3,4}[-]?[A-Z]{0,2}\b",
    r"\b(?:Investigator|Inv)[-#:\s]*(?:Site|ID)?[-#:\s]*\d{3,5}\b",
    r"\b\d{4}[-][A-Z]{2,3}[-]\d{3}\b",
    r"\b\d{4}[-]\d{3,4}[-]\d{2,4}\b",
]

IRB_ETHICS_PATTERNS = [
    r"\b(?:IRB|CEIC|CEISH|CEIM|CEI|EC|REB|HREC|IEC|DSMB|DMC)[-#:\s]*\d{4,}\b",
    r"\b(?:IRB|Ethics|Ethical)[-\s]*(?:Committee|Board|Approval|Review)?[-#:\s]*(?:No|Number|ID)?[-#:\s]*[A-Z0-9]{2,}[-/]?\d{2,}\b",
    r"\b(?:Protocol|Study)[-\s]*(?:Approval|Authorization)[-#:\s]*[A-Z0-9]{2,}[-/]?\d{4,}\b",
    r"\b(?:Comité\s+(?:de\s+)?[ÉE]tica|Ethics\s+Committee|Institutional\s+Review\s+Board)\s+(?:Approval|Reference)?[-#:\s]*[A-Z0-9/-]+\b",
    r"\b(?:FDA|EMA|PMDA|TGA|Health\s+Canada|ANVISA|NMPA|MHRA)[-\s]*(?:Approval|IND|NDA|BLA|MAA)?[-#:\s]*\d{4,}\b",
]

AGE_RANGE_PATTERNS = [
    r"\b(?:aged?|ages?)\s*(?:\d{1,3}\s*[-–to]+\s*\d{1,3}|\d{1,3}\s*(?:years?|y\.?o\.?|yo))\b",
    r"\b[≥≤><]\s*\d{1,3}\s*(?:years?(?:\s+(?:of\s+)?age)?|y\.?o\.?)\b",
    r"\b(?:between|from)\s+\d{1,3}\s+(?:and|to)\s+\d{1,3}\s+(?:years?(?:\s+(?:of\s+)?age)?)\b",
    r"\b\d{1,3}\s*[-–]\s*\d{1,3}\s*(?:years?(?:\s+(?:of\s+)?age)?|y\.?o\.?)\b",
    r"\b(?:pediatric|paediatric|adult|elderly|geriatric)\s+(?:patients?|subjects?|population|cohort)\s*\(?\d{1,3}\s*[-–to]+\s*\d{1,3}(?:\s*years?)?\)?\b",
    r"\b(?:inclusion|exclusion|eligibility)[-:\s]*.*?(?:age|años|âge)\s*[≥≤><]?\s*\d{1,3}\b",
]

HOSPITAL_PATTERNS = [
    r"\b(?:Hospital|Hosp\.?)\s+(?:Universitario|University|General|Regional|Central|Municipal|Nacional|National|Metropolitan|Memorial|Community|Children'?s|Infantil|Pedi[aá]trico|Cl[ií]nico|Teaching)\s+(?:de\s+|del\s+)?[A-ZÀ-ÿ][a-zÀ-ÿ]+(?:\s+[A-ZÀ-ÿ][a-zÀ-ÿ]+){0,3}\b",
    r"\b[A-ZÀ-ÿ][a-zÀ-ÿ]+(?:\s+[A-ZÀ-ÿ][a-zÀ-ÿ]+){0,2}\s+(?:Hospital|Medical\s+Center|Clinic|Cl[ií]nica|Krankenhaus|Klinik|Hôpital|Ospedale|Ziekenhuis|Sjukhus|Sykehus)\b",
    r"\b(?:Medical\s+Center|Centro\s+M[eé]dico|Centre\s+M[eé]dical|Medizinisches\s+Zentrum)\s+[A-ZÀ-ÿ][a-zÀ-ÿ]+(?:\s+[A-ZÀ-ÿ][a-zÀ-ÿ]+){0,2}\b",
    r"\b(?:Cl[ií]nica|Clinic|Klinik)\s+[A-ZÀ-ÿ][a-zÀ-ÿ]+(?:\s+[A-ZÀ-ÿ][a-zÀ-ÿ]+){0,2}\b",
    r"\b(?:University|Universidad|Universit[äa]t|Universit[eé])\s+(?:of\s+)?[A-ZÀ-ÿ][a-zÀ-ÿ]+(?:\s+[A-ZÀ-ÿ][a-zÀ-ÿ]+){0,2}\s+(?:Hospital|Medical\s+Center|School\s+of\s+Medicine|Faculty\s+of\s+Medicine)\b",
    r"\b(?:NHS|VA|Veterans\s+Affairs|Kaiser|Mayo|Cleveland|Johns\s+Hopkins|Mass\s+General|MGH|Cedars[- ]Sinai|Mount\s+Sinai|NYU\s+Langone|Stanford\s+Health|UCSF|UCLA\s+Health)\b",
    r"\b(?:Charit[eé]|AP-HP|Karolinska|Huddinge|Sahlgrenska|Rigshospitalet|Erasmus\s+MC|UMC\s+Utrecht|LMU\s+Klinikum|Heidelberg\s+University\s+Hospital)\b",
    r"\b(?:Institut|Institute|Centro|Centre|Center)\s+(?:of\s+|de\s+|für\s+)?(?:Oncology|Oncolog[ií]a|Cancer|Cardiology|Cardiolog[ií]a|Neurology|Neurolog[ií]a|Research|Investigaci[oó]n)\s*[A-ZÀ-ÿ]?[a-zÀ-ÿ]*\b",
]

COUNTRY_NATIONALITY_PATTERNS = []

STUDY_NAME_PATTERNS = [
    r"\b(?:KEYNOTE|CHECKMATE|JAVELIN|IMPOWER|HIMALAYA|TOPAZ|DESTINY|EMBRACE|MONALEESA|PALOMA|MONARCH|OLYMPIA|PEARL|PRIMA|PROFOUND|TRITON|TALAPRO|MAGNITUDE|PROPEL|ENZAMET|TITAN|ARASENS|SPARTAN|PROSPER|GALAHAD|VISION|PSMA|TAILOR|ACORN|ARROW|BEACON|CASCADE|ELEVATE|FALCON|GEMINI|HORIZON|INFINITY|JUPITER|LOTUS|NAUTILUS|ORION|PIONEER|QUEST|RADIANT|STELLAR|TRINITY|ULTRA|VERTEX|WISDOM|XENON|ZENITH)[-]?\d{0,4}\b",
    r"\b(?:SOLO|NOVA|ARIEL|POLO|DELTA)[-]\d{1,4}\b",
    r"\b(?:Study|Trial|Protocol)\s+[A-Z]{2,5}[-]?\d{3,5}(?:[-][A-Z]{1,3})?\b",
    r"\b[A-Z]{2,4}[-]?\d{3,5}[-]?(?:Study|Trial)\b",
    r"\b(?:Phase\s+)?[IViv]{1,3}[ab]?\s+(?:Study|Trial)\s+[A-Z0-9]{2,8}\b",
]

RANDOMIZATION_PATTERNS = [
    r"\b(?:Randomization|Randomisation|RANDO|RAND)[-#:\s]*(?:No|Number|ID|Code)?[-#:\s]*[A-Z0-9]{2,}[-]?\d{4,}\b",
    r"\b(?:IVRS|IWRS|IRT|IxRS)[-#:\s]*(?:No|Number|ID|Code)?[-#:\s]*[A-Z0-9]{4,}\b",
    r"\b(?:Subject|Patient|Screening)[-\s]*(?:Randomization|Randomisation)[-#:\s]*[A-Z0-9]{2,}[-]?\d{3,}\b",
    r"\b(?:Treatment|Arm|Group)[-\s]*(?:Assignment|Allocation)[-#:\s]*[A-Z0-9]{2,}\b",
    r"\b(?:Kit|Drug|Medication)[-\s]*(?:Number|No|ID|Code)[-#:\s]*[A-Z0-9]{4,}\b",
    r"\b(?:Blinded|Unblinded)[-\s]*(?:Code|ID|Number)[-#:\s]*[A-Z0-9]{4,}\b",
]

MEDICAL_DEVICE_PATTERNS = [
    r"\b(?:Medtronic|Boston\s*Scientific|Abbott|Edwards\s*Lifesciences|Stryker|Zimmer\s*Biomet|Johnson\s*&\s*Johnson|Siemens\s*Healthineers|Philips|GE\s*Healthcare|BD|Becton\s*Dickinson|Baxter|Fresenius|B\.?\s*Braun|Smith\s*&\s*Nephew|Intuitive\s*Surgical|Dexcom|Insulet|Tandem|Terumo|Olympus|Cook\s*Medical|Hologic)\b",
    r"\b(?:DaVinci|da\s*Vinci|Optima|Signa|Ingenia|Artis|Somatom|Navigator|Libre|Guardian|Enlite|G6|G7|OmniPod|Minimed|Accu-Chek|OneTouch|Contour|FreeStyle)\b",
    r"\b(?:Stent|Pacemaker|Defibrillator|ICD|CRT|Catheter|Valve|Prosthesis|Implant|Pump|Monitor|Sensor|Electrode|Lead|Graft|Mesh|Scaffold|Coil|Filter|Balloon|Shunt|Port|Reservoir)\s+[A-Z][a-zA-Z0-9]*(?:\s+[A-Z][a-zA-Z0-9]*)?\b",
    r"\b[A-Z][a-zA-Z]*(?:Stent|Valve|Cath|Pump|Lead|Graft|Mesh|Coil|Flow|Guard|Seal|Fix|Flex|Plus|Pro|Ultra|Max|Elite|Prime|Neo|Next|One|360|3D)\b",
    r"\b(?:CE[-\s]?marked|510\(k\)|PMA|De\s*Novo|Class\s+[IViv]{1,3})\s+(?:device|product|system)\b",
]

INTERNATIONAL_PHONE_PATTERNS = [
    r"\+\d{1,3}[-.\s]?\(?\d{1,4}\)?[-.\s]?\d{2,4}[-.\s]?\d{2,4}[-.\s]?\d{0,4}\b",
    r"\b00\d{1,3}[-.\s]?\d{2,4}[-.\s]?\d{2,4}[-.\s]?\d{2,4}\b",
    r"\b\(\d{2,4}\)\s*\d{3,4}[-.\s]?\d{3,4}\b",
    r"\b\d{3}[-.\s]\d{3}[-.\s]\d{4}\b",
    r"\b\d{4}[-.\s]\d{3}[-.\s]\d{3}\b",
    r"\b\d{2}[-.\s]\d{4}[-.\s]\d{4}\b",
    r"\b(?:Tel|Phone|Fax|Mobile|Cell)[-:\s]*[+]?\d[\d\s\-().]{7,18}\b",
    r"\b\d{2}\s+\d{3}\s+\d{3}\b",
    r"\b\d{3}\s+\d{3}\s+\d{3}\b",
    r"\b\d{3}\s+\d{2}\s+\d{2}\s+\d{2}\b",
    r"\b9\d{2}[-.\s]?\d{3}[-.\s]?\d{3}\b",
    r"\b[67]\d{2}[-.\s]?\d{3}[-.\s]?\d{3}\b",
]

CORPORATE_EMAIL_PATTERNS = [
    r"\b[A-Za-z0-9._%+-]+@(?:pfizer|novartis|roche|sanofi|merck|msd|astrazeneca|gsk|glaxosmithkline|abbvie|bms|bristol-?myers|lilly|amgen|gilead|bayer|novonordisk|novo-nordisk|takeda|boehringer|biogen|regeneron|moderna|biontech|teva|astellas|eisai|vertex|alexion|iqvia|ppd|icon|syneos|parexel|covance|labcorp|medpace|fortrea|wuxi|pra-health|quintiles)\.com\b",
    r"\b[A-Za-z0-9._%+-]+@(?:veeva|medidata|oracle|citeline|clario|signant|castor|advarra|wcg|transperfect|lionbridge|rws|csoft|welocalize)\.com\b",
    r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.(?:pharma|clinical|medical|health|research|trials?|study|bio|med)\.[A-Za-z]{2,}\b",
]

POSTAL_CODE_PATTERNS = [
    r"\b\d{5}(?:[-\s]?\d{4})?\b(?=\s*(?:USA?|United\s+States|[A-Z]{2}\s|,))",
    r"\b[A-Z]{1,2}\d{1,2}[A-Z]?\s*\d[A-Z]{2}\b",
    r"\b[A-Z]\d[A-Z]\s*\d[A-Z]\d\b",
    r"\b\d{3}[-\s]?\d{4}\b(?=\s*(?:Japan|JP|Tokyo|Osaka))",
    r"\b\d{4}\s*[A-Z]{2}\b",
    r"\b\d{5}\s+[A-ZÀ-ÿ][a-zÀ-ÿ]+\b(?=\s*(?:Germany|France|Spain|Italy|DE|FR|ES|IT))",
    r"\b(?:CP|C\.P\.|Código\s+Postal|Postcode|ZIP|PLZ)[-:\s]*\d{4,6}[-\s]?[A-Z]{0,2}\b",
]

DOB_PATTERNS = [
    r"\b(?:DOB|D\.O\.B\.|Date\s+of\s+Birth|Birth\s*Date|Fecha\s+de\s+Nacimiento|Geburtsdatum|Date\s+de\s+Naissance)[-:\s]*\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4}\b",
    r"\b(?:DOB|D\.O\.B\.|Date\s+of\s+Birth|Birth\s*Date)[-:\s]*(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[-.\s]*\d{1,2}[-,.\s]*\d{2,4}\b",
    r"\b(?:Born|Nacido|Né|Geboren)[-:\s]*(?:on\s+)?(?:\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[-.\s]*\d{1,2}[-,.\s]*\d{2,4})\b",
    r"\b(?:Age|Edad|Âge|Alter)[-:\s]*\d{1,3}\s*(?:years?|años|ans|Jahre)?\s*(?:old)?\b",
]

XLIFF_NS = {
    "xliff": "urn:oasis:names:tc:xliff:document:1.2",
    "mq": "MQXliff"
}

INLINE_TAG_NAMES = {"ph", "bpt", "ept", "it", "bx", "ex", "x", "g", "mrk", "sub"}

CRITICAL_PII_PATTERNS = [
    r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b",
    r"\+\d{1,3}[-.\s]?\(?\d{1,4}\)?[-.\s]?\d{2,4}[-.\s]?\d{2,4}",
    r"(?:https?://)[^\s]+",
    r"\bNCT\d{8}\b",
    r"\b\d{4}-\d{6}-\d{2}\b",
]

# ============================================================================
# PRECOMPILED REGEX PATTERNS (for performance)
# ============================================================================

def _compile_pattern_list(patterns, flags=re.IGNORECASE):
    """Compile a list of pattern strings into regex objects."""
    return [re.compile(p, flags) for p in patterns]

def _compile_pattern_dict(patterns, flags=re.IGNORECASE):
    """Compile a dict of pattern strings into regex objects."""
    return {name: re.compile(p, flags) for name, p in patterns.items()}

# Compile all pattern lists
CLINICAL_TRIAL_PATTERNS_RE = _compile_pattern_dict(CLINICAL_TRIAL_PATTERNS)
CASE_SENSITIVE_PATTERNS_RE = _compile_pattern_dict(CASE_SENSITIVE_PATTERNS, flags=0)  # No IGNORECASE - uppercase only
BIOMEDICAL_PATTERNS_RE = _compile_pattern_dict(BIOMEDICAL_PATTERNS)
CLINICAL_ABBREVIATIONS_RE = _compile_pattern_list(CLINICAL_ABBREVIATIONS_WITH_VALUE)
PHARMA_COMPANY_RE = _compile_pattern_list(PHARMA_COMPANY_PATTERNS)
CRO_RE = _compile_pattern_list(CRO_PATTERNS)
BLOCKBUSTER_DRUGS_RE = _compile_pattern_list(BLOCKBUSTER_DRUGS)
CLINICAL_TECH_RE = _compile_pattern_list(CLINICAL_TECH_PLATFORMS)
CENTRAL_LABS_RE = _compile_pattern_list(CENTRAL_LABS)
LAB_PRODUCT_RE = _compile_pattern_list(LAB_PRODUCT_PATTERNS)
ACRONYM_RE = _compile_pattern_list(ACRONYM_PATTERNS, flags=0)  # No IGNORECASE - pattern requires uppercase
SOFTWARE_PRODUCT_RE = _compile_pattern_list(SOFTWARE_PRODUCT_PATTERNS, flags=0)
URL_RE = _compile_pattern_list(URL_PATTERNS)
CLINICAL_URL_RE = _compile_pattern_list(CLINICAL_URL_PATTERNS)
ADDRESS_RE = _compile_pattern_list(ADDRESS_PATTERNS)
INVESTIGATOR_RE = _compile_pattern_list(INVESTIGATOR_PATTERNS, flags=0)  # No IGNORECASE - pattern defines case explicitly
CLINICAL_DATE_RE = _compile_pattern_list(CLINICAL_DATE_PATTERNS)
LOT_BATCH_RE = _compile_pattern_list(LOT_BATCH_PATTERNS)
SITE_CODE_RE = _compile_pattern_list(SITE_CODE_PATTERNS)
IRB_ETHICS_RE = _compile_pattern_list(IRB_ETHICS_PATTERNS, flags=0)
AGE_RANGE_RE = _compile_pattern_list(AGE_RANGE_PATTERNS)
HOSPITAL_RE = _compile_pattern_list(HOSPITAL_PATTERNS, flags=0)
COUNTRY_NATIONALITY_RE = _compile_pattern_list(COUNTRY_NATIONALITY_PATTERNS)
STUDY_NAME_RE = _compile_pattern_list(STUDY_NAME_PATTERNS)
RANDOMIZATION_RE = _compile_pattern_list(RANDOMIZATION_PATTERNS)
MEDICAL_DEVICE_RE = _compile_pattern_list(MEDICAL_DEVICE_PATTERNS, flags=0)
INTERNATIONAL_PHONE_RE = _compile_pattern_list(INTERNATIONAL_PHONE_PATTERNS)
CORPORATE_EMAIL_RE = _compile_pattern_list(CORPORATE_EMAIL_PATTERNS)
POSTAL_CODE_RE = _compile_pattern_list(POSTAL_CODE_PATTERNS)
DOB_RE = _compile_pattern_list(DOB_PATTERNS)
CRITICAL_PII_RE = _compile_pattern_list(CRITICAL_PII_PATTERNS)

# Helper functions for working with compiled patterns
def _findall_compiled_list(compiled_list, text):
    """Find all matches from a list of compiled patterns."""
    results = []
    for compiled in compiled_list:
        results.extend(compiled.findall(text))
    return results

def _sub_compiled_list(compiled_list, replacement, text):
    """Apply substitutions from a list of compiled patterns, returns (text, count)."""
    count = 0
    for compiled in compiled_list:
        matches = compiled.findall(text)
        if matches:
            count += len(matches)
            text = compiled.sub(replacement, text)
    return text, count

def _sub_compiled_dict(compiled_dict, replacement, text):
    """Apply substitutions from a dict of compiled patterns, returns (text, count)."""
    count = 0
    for name, compiled in compiled_dict.items():
        matches = compiled.findall(text)
        if matches:
            count += len(matches)
            text = compiled.sub(replacement, text)
    return text, count

ORG_CONTEXT_WORDS = ["sponsor", "vendor", "company", "platform", "trial", "protocol", "drug", "device", "pharma", "bio", "lab", "manufacturer", "supplier", "contractor"]


class MQXLIFFAnonymizer:
    def __init__(self, replacement_token: str = "REDACTED"):
        self.replacement_token = replacement_token
        self.nlp_en = None
        self.nlp_es = None
        self.nlp_biomedical = None
        self.presidio_analyzer = None
        self.presidio_anonymizer = None
        self.stats = {
            "regex_ct": 0,
            "presidio_pii": 0,
            "biomedical": 0,
            "dictionary": 0,
            "critical_pii_remaining": {},
            "timings": {
                "total_ms": 0.0,
                "segments": 0,
                "last_segment_ms": 0.0
            }
        }
        self.terms_cache = set()
        self.singleton_candidates = {}
        self.lowercase_words = set()
        self._cache_regex = None
        self._cache_dirty = True
        self.strict_validation = False
        self.enable_benchmark = False
    
    def _load_nlp_models(self):
        if self.nlp_en is None:
            try:
                self.nlp_en = spacy.load("en_core_web_lg")
            except OSError:
                raise RuntimeError(
                    "Missing model: en_core_web_lg. "
                    "Install with: python -m spacy download en_core_web_lg"
                )
        if self.nlp_es is None:
            try:
                self.nlp_es = spacy.load("es_core_news_md")
            except OSError:
                self.nlp_es = None
                import logging
                logging.warning(
                    "Missing model: es_core_news_md — Spanish NLP features disabled. "
                    "Install with: python -m spacy download es_core_news_md"
                )
        if self.nlp_biomedical is None:
            try:
                self.nlp_biomedical = spacy.load("en_ner_bc5cdr_md")
            except OSError:
                self.nlp_biomedical = None
                import logging
                logging.warning(
                    "Missing model: en_ner_bc5cdr_md — Biomedical entity detection (CHEMICAL/DISEASE) disabled. "
                    "Install with: pip install scispacy && pip install https://s3-us-west-2.amazonaws.com/ai2-s2-scispacy/releases/v0.5.4/en_ner_bc5cdr_md-0.5.4.tar.gz"
                )
        if self.presidio_analyzer is None:
            try:
                self.presidio_analyzer = AnalyzerEngine()
                self.presidio_anonymizer = AnonymizerEngine()
            except Exception:
                self.presidio_analyzer = None
                self.presidio_anonymizer = None
                import logging
                logging.warning(
                    "Missing dependency: presidio — PII detection disabled. "
                    "Install with: pip install presidio-analyzer presidio-anonymizer"
                )
    
    def _get_nlp(self, lang: str = "en"):
        self._load_nlp_models()
        if lang == "es" and self.nlp_es is not None:
            return self.nlp_es
        return self.nlp_en
    
    def reset_stats(self):
        self.stats = {
            "regex_ct": 0,
            "presidio_pii": 0,
            "biomedical": 0,
            "dictionary": 0,
            "critical_pii_remaining": {},
            "timings": {
                "total_ms": 0.0,
                "segments": 0,
                "last_segment_ms": 0.0
            }
        }
        self.terms_cache = set()
        self.singleton_candidates = {}
        self.lowercase_words = set()
        self._cache_regex = None
        self._cache_dirty = True
    
    def _scan_document_for_lowercase(self, tree):
        """Scan document to find all words that appear in lowercase.
        Used to validate cache - if a term appears lowercase anywhere, it's not a proper noun.
        Excludes content from emails and URLs to avoid false lowercase detection."""
        trans_units = tree.xpath("//xliff:trans-unit", namespaces=XLIFF_NS)
        
        email_url_pattern = re.compile(
            r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}|'
            r'https?://[^\s]+|'
            r'www\.[^\s]+|'
            r'[a-zA-Z0-9.-]+\.(com|org|net|edu|gov|io|co|es|de|fr|uk|eu)[^\s]*',
            re.IGNORECASE
        )
        
        for tu in trans_units:
            for elem in tu.xpath(".//xliff:source | .//xliff:target", namespaces=XLIFF_NS):
                text = "".join(elem.itertext())
                if text:
                    clean_text = email_url_pattern.sub(' ', text)
                    words = re.findall(r'\b[a-záéíóúñüàèìòùäëïöü]+\b', clean_text, re.IGNORECASE)
                    for word in words:
                        if word.islower() and len(word) > 2:
                            self.lowercase_words.add(word.lower())
    
    def _should_block_cache_candidate(self, span_text: str, doc_span, lang: str) -> bool:
        """
        Devuelve True si NO deberíamos cachear un candidato (para evitar falsos positivos).
        doc_span: idealmente un spaCy Span/Doc con tokens.
        lang: "en" o "es"
        v10.0: Separación stopwords funcionales vs sustantivos comunes
        """
        if not span_text:
            return True

        # Tokenización robusta usando spaCy si está disponible
        tokens = []
        if doc_span is not None:
            for t in doc_span:
                if getattr(t, "is_space", False) or getattr(t, "is_punct", False):
                    continue
                tokens.append(t)
        else:
            # fallback simple si no hay doc_span
            tokens = span_text.strip().split()

        if not tokens:
            return True

        # Helpers
        stopwords = STOPWORDS_FUNCTIONAL_EN if lang == "en" else STOPWORDS_FUNCTIONAL_ES
        common_block = COMMON_SINGLETON_BLOCK_EN if lang == "en" else COMMON_SINGLETON_BLOCK_ES

        # Señales de entidad: NER del original es fiable (ORG, PERSON, GPE...);
        # PROPN sin NER se verifica en minúscula para evitar falsos por capitalización
        has_ner = False
        has_propn = False

        if doc_span is not None:
            for t in doc_span:
                if getattr(t, "ent_type_", "") in ALLOWED_NER_LABELS:
                    has_ner = True

        if not has_ner:
            self._load_nlp_models()
            nlp = self._get_nlp(lang)
            if nlp and span_text:
                lc_doc = nlp(span_text.lower())
                for t in lc_doc:
                    if t.is_space:
                        continue
                    if t.pos_ == "PROPN":
                        has_propn = True
                    if t.ent_type_ in ALLOWED_NER_LABELS:
                        has_ner = True

        # Si hay NER fuerte, NO bloquear
        if has_ner:
            return False

        # Normalizados
        if doc_span is not None:
            lower_tokens = [t.lower_ for t in tokens if hasattr(t, 'lower_')]
        else:
            lower_tokens = [str(x).lower() for x in tokens]

        # Regla 1: 1 token que es stopword funcional → bloquear
        if len(lower_tokens) == 1 and lower_tokens[0] in stopwords:
            return True

        # Regla 2: 1 token que es sustantivo común bloqueado → bloquear SIEMPRE
        # (ya verificamos NER arriba, si no hay NER fuerte, bloqueamos aunque sea PROPN)
        if len(lower_tokens) == 1 and lower_tokens[0] in common_block:
            return True

        # Regla 3: 2 tokens con stopword + common_block → bloquear (ej: "The System", "El Sistema")
        if len(lower_tokens) == 2:
            has_stopword = any(t in stopwords for t in lower_tokens)
            has_common_block = any(t in common_block for t in lower_tokens)
            all_common_block = all(t in common_block for t in lower_tokens)
            if has_stopword and has_common_block:
                return True
            if all_common_block:
                return True
            # Si tiene stopword y no tiene PROPN → bloquear
            if has_stopword and not has_propn:
                return True

        return False
    
    def _add_to_cache(self, term: str, lang: str = "en", doc_span=None):
        """Add a detected term to the cache for consistent anonymization.
        
        v10.0: Integra _should_block_cache_candidate() para mejor filtrado.
        doc_span: opcional (Span/Doc de spaCy) del candidato.
        """
        term_strip = term.strip()
        if len(term_strip) < 4 or self.replacement_token in term_strip:
            return

        words = term_strip.split()
        num_words = len(words)
        key = term_strip.lower()

        # v10.0: Usar el nuevo sistema de bloqueo
        if self._should_block_cache_candidate(term_strip, doc_span, lang):
            return

        def is_titlecase_word(w: str) -> bool:
            return len(w) >= 2 and w[0].isupper() and w[1:].islower()

        title_case_count = sum(1 for w in words if is_titlecase_word(w))
        has_digits_hyphens = bool(re.search(r"[\d_-]", term_strip))
        is_acronym = term_strip.isupper() and len(term_strip) >= 3

        only_first_capitalized = (
            num_words > 1
            and words[0][:1].isupper()
            and all(w[:1].islower() for w in words[1:] if w)
        )

        # Evidencia PROPN/NER: NER del original es fiable (ORG, PERSON, etc.);
        # PROPN sin NER se verifica en minúscula para evitar inflación por capitalización
        propn_ratio = 0.0
        propn_count = 0
        if not (has_digits_hyphens or is_acronym or title_case_count >= 2):
            self._load_nlp_models()
            nlp = self._get_nlp(lang)
            if nlp:
                orig_doc = nlp(term_strip)
                orig_tokens = [t for t in orig_doc if not t.is_space]
                if not orig_tokens:
                    return
                has_orig_ner = any(t.ent_type_ in ALLOWED_NER_LABELS for t in orig_tokens)
                if has_orig_ner:
                    propn_count = sum(1 for t in orig_tokens if (t.pos_ == "PROPN" or t.ent_type_ in ALLOWED_NER_LABELS) and not t.is_stop)
                else:
                    lc_doc = nlp(term_strip.lower())
                    lc_tokens = [t for t in lc_doc if not t.is_space]
                    if not lc_tokens:
                        return
                    propn_count = sum(1 for t in lc_tokens if t.pos_ == "PROPN" and not t.is_stop)
                    orig_tokens = lc_tokens
                propn_ratio = propn_count / len(orig_tokens)

        if num_words >= 2:
            if title_case_count >= 2:
                self._load_nlp_models()
                nlp = self._get_nlp(lang)
                if nlp:
                    orig_tc_doc = nlp(term_strip)
                    orig_tc_tokens = [t for t in orig_tc_doc if not t.is_space]
                    has_tc_ner = any(t.ent_type_ in ALLOWED_NER_LABELS for t in orig_tc_tokens)
                    if not has_tc_ner:
                        lc_tc_doc = nlp(term_strip.lower())
                        lc_tc_tokens = [t for t in lc_tc_doc if not t.is_space]
                        has_real_propn = any(
                            t.pos_ == "PROPN" or t.ent_type_ in ALLOWED_NER_LABELS
                            for t in lc_tc_tokens
                        )
                        if not has_real_propn:
                            return
                
                if key not in self.terms_cache:
                    self.terms_cache.add(key)
                    self._cache_dirty = True
                return
            
            has_strong_evidence = (
                has_digits_hyphens
                or is_acronym
                or propn_ratio >= 0.50
                or propn_count >= 1
            )

            # Inicio de oración (solo primera palabra capitalizada): más estricto
            if only_first_capitalized and not (has_digits_hyphens or is_acronym or propn_ratio >= 0.67):
                return

            if not has_strong_evidence:
                return

        else:
            # single-word: muy estricto
            if len(term_strip) < 3:
                return
            if not (is_acronym or has_digits_hyphens):
                # exige PROPN ratio = 1.0
                if propn_ratio < 1.0:
                    return

        # Añadir al cache
        if key not in self.terms_cache:
            self.terms_cache.add(key)
            self._cache_dirty = True
    
    def _build_cache_regex(self):
        """Build a single compiled regex from all cached terms."""
        if not self.terms_cache or len(self.terms_cache) > 10000:
            self._cache_regex = None
            return
        
        sorted_terms = sorted(self.terms_cache, key=len, reverse=True)
        escaped_terms = [r'(?<![A-Za-z0-9À-ÿ])' + re.escape(t) + r'(?![A-Za-z0-9À-ÿ])' for t in sorted_terms]
        pattern = '|'.join(escaped_terms)
        self._cache_regex = re.compile(pattern, re.IGNORECASE)
        self._cache_dirty = False

    def _apply_cache(self, text: str) -> str:
        """Apply cached terms using optimized single-pass regex."""
        if not self.terms_cache:
            return text
        
        if self._cache_dirty or self._cache_regex is None:
            self._build_cache_regex()
        
        if self._cache_regex is None:
            # Fallback for very large caches (>10k terms)
            result = text
            for term in sorted(self.terms_cache, key=len, reverse=True):
                pattern = r'(?<![A-Za-z0-9À-ÿ])' + re.escape(term) + r'(?![A-Za-z0-9À-ÿ])'
                def _fallback_replace(m):
                    matched = m.group(0)
                    if matched.islower():
                        return matched
                    fb_words = [w for w in re.findall(r'[a-záéíóúñüàèìòùäëïöü]+', matched.lower(), re.IGNORECASE) if len(w) > 2]
                    if fb_words and self.lowercase_words and all(w in self.lowercase_words for w in fb_words):
                        return matched
                    self.stats["regex_ct"] += 1
                    return self.replacement_token
                result = re.sub(pattern, _fallback_replace, result, flags=re.IGNORECASE)
            return result
        
        def replace_match(m):
            matched = m.group(0)
            if matched.islower():
                return matched
            match_words = [w for w in re.findall(r'[a-záéíóúñüàèìòùäëïöü]+', matched.lower(), re.IGNORECASE) if len(w) > 2]
            if match_words and self.lowercase_words and all(w in self.lowercase_words for w in match_words):
                return matched
            self.stats["regex_ct"] += 1
            return self.replacement_token
        
        return self._cache_regex.sub(replace_match, text)

    def _apply_cache_to_element(self, element):
        """Apply accumulated cache to all text nodes in an XML element (second pass)."""
        if not self.terms_cache:
            return
        if self._cache_dirty or self._cache_regex is None:
            self._build_cache_regex()
        if self._cache_regex is None:
            return
        
        def apply_to_text(text):
            if not text or not text.strip():
                return text
            def replace_match(m):
                matched = m.group(0)
                if matched.islower():
                    return matched
                match_words = [w for w in re.findall(r'[a-záéíóúñüàèìòùäëïöü]+', matched.lower(), re.IGNORECASE) if len(w) > 2]
                if match_words and self.lowercase_words and all(w in self.lowercase_words for w in match_words):
                    return matched
                self.stats["regex_ct"] += 1
                return self.replacement_token
            return self._cache_regex.sub(replace_match, text)
        
        if element.text:
            element.text = apply_to_text(element.text)
        for child in element:
            if child.tail:
                child.tail = apply_to_text(child.tail)
            local_name = etree.QName(child.tag).localname if isinstance(child.tag, str) else None
            if local_name not in INLINE_TAG_NAMES:
                self._apply_cache_to_element(child)

    def _mask_pii_example(self, text: str, pii_type: str) -> str:
        """Mask PII for safe logging (e.g., j***@d***.com, ***1234 for IDs)."""
        if pii_type == "email" and "@" in text:
            parts = text.split("@")
            user = parts[0][0] + "***" if parts[0] else "***"
            domain_parts = parts[1].split(".") if len(parts) > 1 else ["***", "com"]
            domain = domain_parts[0][0] + "***" if domain_parts[0] else "***"
            tld = domain_parts[-1] if len(domain_parts) > 1 else "com"
            return f"{user}@{domain}.{tld}"
        elif pii_type == "phone":
            digits = re.sub(r'\D', '', text)
            return "***" + digits[-3:] if len(digits) >= 3 else "***"
        elif pii_type == "url":
            match = re.match(r'(https?://)?([^/]+)', text)
            if match:
                return (match.group(1) or "") + match.group(2)[:10] + "***"
            return text[:10] + "***"
        elif pii_type in ["clinical_id", "eudract_id"]:
            # Show only last 4 characters for clinical IDs (e.g., ***1234)
            return "***" + text[-4:] if len(text) >= 4 else "***"
        else:
            return text[:3] + "***" + text[-2:] if len(text) > 5 else "***"

    def _validate_no_critical_pii(self, text: str) -> dict:
        """Post-scan to detect any PII that escaped the pipeline."""
        categories = {
            "email": CRITICAL_PII_RE[0],
            "phone": CRITICAL_PII_RE[1],
            "url": CRITICAL_PII_RE[2],
            "clinical_id": CRITICAL_PII_RE[3],
            "eudract_id": CRITICAL_PII_RE[4],
        }
        
        counts = {}
        examples = {}
        
        for cat_name, pattern in categories.items():
            matches = pattern.findall(text)
            # Filter out matches that are just the replacement token
            real_matches = [m for m in matches if self.replacement_token not in m]
            counts[cat_name] = len(real_matches)
            if real_matches:
                examples[cat_name] = [self._mask_pii_example(m, cat_name) for m in real_matches[:3]]
        
        return {"counts": counts, "examples": examples}
    
    def _detect_proper_nouns_spacy(self, text: str, lang: str = "en") -> str:
        """Detect and anonymize proper nouns based on capitalization rules.
        
        v10.0: Mid-sentence requiere PROPN o NER válido para reducir falsos positivos.
        
        Rules:
        - At sentence start: only anonymize if 2+ consecutive capitalized words (compound name)
        - Mid-sentence: SOLO tokens PROPN o con ent_type_ permitido (v10.0)
        - Consecutive capitalized words are grouped together
        """
        self._load_nlp_models()
        nlp = self._get_nlp(lang)
        if not nlp:
            return text
        
        doc = nlp(text)
        result = text
        
        skip_pos = {"DET", "PRON", "ADP", "CCONJ", "SCONJ", "AUX", "PART", "VERB"}
        skip_words = {'i', 'me', 'my', 'you', 'he', 'she', 'it', 'we', 'they', 'this', 'that', 'these', 'those'}
        
        propn_spans = []
        i = 0
        
        while i < len(doc):
            token = doc[i]
            
            if not token.text or not token.text[0].isupper() or token.is_space or token.is_punct:
                i += 1
                continue
            
            if token.text in SAFE_ACRONYMS:
                i += 1
                continue
            
            if token.text.lower() in skip_words or token.pos_ in skip_pos:
                i += 1
                continue
            
            is_sentence_start = (i == 0) or (i > 0 and doc[i-1].text in '.!?:')
            
            start_idx = token.idx
            end_idx = token.idx + len(token.text)
            consecutive_caps = 1
            span_start_i = i
            
            j = i + 1
            while j < len(doc):
                next_tok = doc[j]
                if next_tok.is_space:
                    j += 1
                    continue
                if next_tok.text and next_tok.text[0].isupper() and next_tok.pos_ not in skip_pos:
                    end_idx = next_tok.idx + len(next_tok.text)
                    consecutive_caps += 1
                    j += 1
                else:
                    break
            
            span_end_i = j
            
            should_anonymize = False
            is_mid_sentence_propn = False
            if is_sentence_start:
                if consecutive_caps >= 2:
                    span_tokens_orig = [doc[k] for k in range(span_start_i, span_end_i) if not doc[k].is_space]
                    has_orig_ner_start = any(
                        t.ent_type_ in ALLOWED_NER_LABELS for t in span_tokens_orig
                    )
                    if has_orig_ner_start:
                        should_anonymize = True
                    else:
                        span_text_raw = text[start_idx:end_idx].strip()
                        lc_doc = nlp(span_text_raw.lower())
                        lc_tokens = [t for t in lc_doc if not t.is_space]
                        has_real_propn = any(
                            t.pos_ == "PROPN" or t.ent_type_ in ALLOWED_NER_LABELS
                            for t in lc_tokens
                        )
                        should_anonymize = has_real_propn
            else:
                span_tokens_orig = [doc[k] for k in range(span_start_i, span_end_i) if not doc[k].is_space]
                has_orig_ner = any(
                    t.ent_type_ in ALLOWED_NER_LABELS for t in span_tokens_orig
                )
                if has_orig_ner:
                    should_anonymize = token.pos_ not in skip_pos
                else:
                    span_text_raw = text[start_idx:end_idx].strip()
                    lc_span_doc = nlp(span_text_raw.lower())
                    lc_span_tokens = [t for t in lc_span_doc if not t.is_space]
                    has_real_propn = any(
                        t.pos_ == "PROPN" or t.ent_type_ in ALLOWED_NER_LABELS
                        for t in lc_span_tokens
                    )
                    should_anonymize = has_real_propn and (token.pos_ not in skip_pos)
                is_mid_sentence_propn = True
            
            if should_anonymize:
                span_text = text[start_idx:end_idx].strip()
                if span_text:
                    # Guardar también los índices del doc para pasar doc_span
                    propn_spans.append((start_idx, end_idx, span_text, is_mid_sentence_propn, span_start_i, span_end_i))
            
            i = j if j > i + 1 else i + 1
        
        propn_spans.sort(key=lambda x: x[0], reverse=True)
        
        for start, end, span_text, is_mid_sentence, doc_start, doc_end in propn_spans:
            term_lower = span_text.lower()
            span_words = [w for w in re.findall(r'[a-záéíóúñüàèìòùäëïöü]+', term_lower, re.IGNORECASE) if len(w) > 2]
            if span_words and all(w in self.lowercase_words for w in span_words):
                continue
            
            # v10.0: Pasar doc_span a _add_to_cache para mejor validación
            doc_span = doc[doc_start:doc_end] if doc_start < doc_end else None
            
            # ALWAYS use _add_to_cache for validation - never add directly to cache
            cache_size_before = len(self.terms_cache)
            self._add_to_cache(span_text, lang, doc_span=doc_span)
            
            # Only anonymize if term was actually added to cache (passed validation)
            if len(self.terms_cache) > cache_size_before or term_lower in self.terms_cache:
                result = result[:start] + self.replacement_token + result[end:]
                self.stats["regex_ct"] += 1
        
        return result
    
    def _should_skip_entity(self, text: str, lang: str = "en", entity_type: str = None) -> bool:
        if re.match(r'^__TAG\d{4}__$', text):
            return True
        
        if '\x00' in text or re.search(r'PH\d+', text):
            clean = re.sub(r'\x00?PH\d+\x00?', '', text).strip()
            if not clean or clean.lower() in STOPWORDS_FUNCTIONAL_EN or clean.lower() in STOPWORDS_FUNCTIONAL_ES:
                return True
            text = clean
        
        always_process_types = {"EMAIL_ADDRESS", "URL", "PHONE_NUMBER", "IBAN_CODE", "CREDIT_CARD", "IP_ADDRESS"}
        if entity_type and entity_type in always_process_types:
            return False
        
        if '@' in text or text.startswith('www.') or text.startswith('http'):
            return False
        
        nlp = self._get_nlp(lang)
        doc = nlp(text)
        
        tokens = [t for t in doc if not t.is_space]
        if len(tokens) == 1 and tokens[0].is_stop:
            return True
        
        has_uppercase = any(c.isupper() for c in text)
        has_digit = any(c.isdigit() for c in text)
        if not has_uppercase and not has_digit:
            return True
        
        if len(text) < 4 and not has_digit:
            return True
        
        return False
    
    def _is_valid_pos_for_redaction(self, text: str, lang: str = "en") -> bool:
        nlp = self._get_nlp(lang)
        doc = nlp(text)
        
        if len(doc) == 0:
            return False
        
        non_space_tokens = [t for t in doc if not t.is_space]
        total_count = len(non_space_tokens)
        
        if total_count == 0:
            return False
        
        if total_count == 1:
            token = non_space_tokens[0]
            if token.pos_ in INVALID_POS_FOR_REDACTION:
                return False
        
        if re.search(r'[A-Z]{2,}[-_/]\d+|\d+[-_/][A-Z]{2,}', text):
            return True
        
        if re.search(r'^[A-Z]{2,}[a-z]+\d+$', text):
            return True
        
        if self._is_name_like(text, lang):
            return True
        
        noun_propn_count = sum(1 for token in doc if token.pos_ in ["NOUN", "PROPN", "NUM", "X"])
        return noun_propn_count / total_count >= 0.5
    
    def _is_structured_id(self, text: str) -> bool:
        patterns = [
            r'^[A-Z]{2,6}[-_/]\d{2,}$',
            r'^\d{2,}[-_/][A-Z]{2,6}$',
            r'^[A-Z0-9]{2,}[-_/][A-Z0-9]{2,}$',
            r'^\d{4}-\d{6}-\d{2}$',
            r'^NCT\d{8}$',
            r'^[A-Z]+\d+[A-Z]*\d*$',
        ]
        for pattern in patterns:
            if re.match(pattern, text, re.IGNORECASE):
                return True
        return False
    
    
    def _has_local_context(self, full_text: str, entity_start: int, entity_end: int, context_words: List[str], window: int = 40) -> bool:
        start_window = max(0, entity_start - window)
        end_window = min(len(full_text), entity_end + window)
        
        local_text = full_text[start_window:end_window].lower()
        
        return any(cw in local_text for cw in context_words)
    
    def _is_name_like(self, text: str, lang: str = "en") -> bool:
        nlp = self._get_nlp(lang)
        doc = nlp(text)
        tokens = [t for t in doc if not t.is_space]
        
        if len(tokens) == 0:
            return False
        
        # v10.0: Verificar si contiene stopwords + common_block (ej: "The System")
        stopwords = STOPWORDS_FUNCTIONAL_EN if lang == "en" else STOPWORDS_FUNCTIONAL_ES
        common_block = COMMON_SINGLETON_BLOCK_EN if lang == "en" else COMMON_SINGLETON_BLOCK_ES
        
        lower_tokens = [t.lower_ for t in tokens]
        has_stopword = any(t in stopwords for t in lower_tokens)
        has_common_block = any(t in common_block for t in lower_tokens)
        
        # Si tiene stopword + common_block → NO es nombre real
        if has_stopword and has_common_block:
            return False
        
        # Si es 1 token y está en common_block → NO es nombre real
        if len(lower_tokens) == 1 and lower_tokens[0] in common_block:
            return False
        
        has_version_pattern = bool(re.search(r'v\d|V\d|\d+\.\d+|\d{4}\s*R\d|[A-Z]{2,}[-_]\d+|\d+[-_][A-Z]{2,}', text))
        if has_version_pattern:
            return True
        
        has_numbers_hyphens = bool(re.search(r'[-_]\d|\d[-_]|[A-Z]+\d+', text))
        if has_numbers_hyphens:
            return True
        
        has_internal_caps = bool(re.search(r'[a-z][A-Z]', text))
        if has_internal_caps:
            return True
        
        has_mixed_case_with_digits = bool(re.search(r'^[A-Z]{2,}[a-z]+\d+$|^[A-Z][a-z]+[A-Z][a-z]*\d*$', text))
        if has_mixed_case_with_digits:
            return True
        
        words = text.split()
        propn_count = sum(1 for t in tokens if t.pos_ == "PROPN")
        propn_ratio = propn_count / len(tokens) if len(tokens) > 0 else 0
        
        if len(words) >= 2:
            title_case_count = sum(1 for w in words if w and w[0].isupper())
            if title_case_count >= 1 and (propn_ratio >= 0.5 or has_numbers_hyphens or has_internal_caps):
                return True
            if title_case_count >= 2:
                return True
        
        if propn_ratio >= 0.5:
            return True
        
        noun_propn_count = sum(1 for t in tokens if t.pos_ in ["NOUN", "PROPN"])
        if len(tokens) >= 2 and noun_propn_count == len(tokens):
            has_any_upper = any(c.isupper() for c in text)
            if has_any_upper:
                return True
        
        return False
    
    def anonymize_with_regex_ct(self, text: str, lang: str = "en") -> str:
        """Anonymize clinical trial IDs using precompiled patterns."""
        result = text
        
        # Use compiled clinical trial patterns
        result, ct_count = _sub_compiled_dict(CLINICAL_TRIAL_PATTERNS_RE, self.replacement_token, result)
        self.stats["regex_ct"] += ct_count
        
        # Use case-sensitive patterns (alphanumeric codes) - must be uppercase only
        result, cs_count = _sub_compiled_dict(CASE_SENSITIVE_PATTERNS_RE, self.replacement_token, result)
        self.stats["regex_ct"] += cs_count
        
        acronym_re = ACRONYM_ALLCAPS_RE
        for m in reversed(list(acronym_re.finditer(result))):
            matched = m.group(0)
            if matched in SAFE_ACRONYMS or matched == self.replacement_token.upper():
                continue
            result = result[:m.start()] + self.replacement_token + result[m.end():]
            self.stats["regex_ct"] += 1
        
        # Use compiled clinical abbreviations
        result, abbrev_count = _sub_compiled_list(CLINICAL_ABBREVIATIONS_RE, self.replacement_token, result)
        self.stats["regex_ct"] += abbrev_count
        
        # Software products - collect matches using compiled patterns
        all_matches = []
        for compiled in SOFTWARE_PRODUCT_RE:
            matches = compiled.findall(result)
            for match in matches:
                if match not in all_matches:
                    all_matches.append(match)
        
        all_matches.sort(key=len, reverse=True)
        
        for match in all_matches:
            if match not in result:
                continue
            words = match.split()
            if words and words[0].lower() in INSTRUCTION_VERBS:
                clean_match = ' '.join(words[1:])
                if clean_match:
                    self._add_to_cache(clean_match, lang=lang)
                    esc_re = re.compile(re.escape(match), re.IGNORECASE)
                    result = esc_re.sub(words[0] + ' ' + self.replacement_token, result)
                    self.stats["regex_ct"] += 1
            else:
                self._add_to_cache(match, lang=lang)
                self.stats["regex_ct"] += 1
                esc_re = re.compile(re.escape(match), re.IGNORECASE)
                result = esc_re.sub(self.replacement_token, result)
        
        return result
    
    def anonymize_with_presidio(self, text: str, lang: str = "en") -> str:
        self._load_nlp_models()
        if self.presidio_analyzer is None:
            return text
        
        entities = [
            "PERSON",
            "EMAIL_ADDRESS", 
            "PHONE_NUMBER",
            "CREDIT_CARD",
            "US_SSN",
            "US_PASSPORT",
            "US_DRIVER_LICENSE",
            "US_ITIN",
            "US_BANK_NUMBER",
            "UK_NHS",
            "IP_ADDRESS",
            "IBAN_CODE",
            "MEDICAL_LICENSE",
            "URL",
            "CRYPTO",
            "SG_NRIC_FIN",
            "AU_ABN",
            "AU_ACN",
            "AU_TFN",
            "AU_MEDICARE",
        ]
        score_threshold = 0.5
        
        try:
            analysis_results = self.presidio_analyzer.analyze(
                text=text,
                language=lang,
                entities=entities,
                score_threshold=score_threshold
            )
            
            filtered_results = []
            for r in analysis_results:
                entity_text = text[r.start:r.end]
                
                if '\x00' in entity_text:
                    continue
                
                if self._should_skip_entity(entity_text, lang, entity_type=r.entity_type):
                    continue
                
                if r.entity_type == "PERSON":
                    entity_words = [w for w in re.findall(r'[a-záéíóúñüàèìòùäëïöü]+', entity_text.lower(), re.IGNORECASE) if len(w) > 2]
                    if entity_words and all(w in self.lowercase_words for w in entity_words):
                        continue
                
                if r.entity_type in ["PERSON", "LOCATION"] and not self._is_valid_pos_for_redaction(entity_text, lang):
                    if r.score < 0.85:
                        continue
                
                # Propagación de tokens individuales para PERSON
                # Si detectamos "Zayab Thomas", guardar "Thomas" en cache para encontrarlo en target
                if r.entity_type == "PERSON":
                    words = entity_text.split()
                    for word in words:
                        word_clean = word.strip()
                        # Solo guardar tokens que parezcan apellidos/nombres propios
                        # (longitud >= 4, empieza con mayúscula, no es stopword común)
                        if (len(word_clean) >= 4 and 
                            word_clean[0].isupper() and 
                            word_clean.lower() not in self.lowercase_words and
                            word_clean.lower() not in {'with', 'from', 'this', 'that', 'have', 'been', 'will', 'would', 'could', 'should'}):
                            self._add_to_cache(word_clean, lang)
                
                filtered_results.append(r)
            
            if filtered_results:
                self.stats["presidio_pii"] += len(filtered_results)
                
                operators = {
                    entity_type: OperatorConfig("replace", {"new_value": self.replacement_token})
                    for entity_type in set(r.entity_type for r in filtered_results)
                }
                
                anonymized = self.presidio_anonymizer.anonymize(
                    text=text,
                    analyzer_results=filtered_results,
                    operators=operators
                )
                return anonymized.text
        except Exception:
            pass
        
        return text
    
    def anonymize_with_biomedical(self, text: str, lang: str = "en") -> str:
        """Anonymize biomedical entities using precompiled patterns."""
        result = text
        self._load_nlp_models()
        
        entities_to_replace = []
        
        # SciSpacy biomedical NER - ONLY for English (model generates garbage on Spanish)
        if self.nlp_biomedical is not None and lang == 'en':
            doc_bio = self.nlp_biomedical(text)
            for ent in doc_bio.ents:
                if ent.label_ in ["CHEMICAL", "DISEASE"]:
                    if len(ent.text) > 4:
                        if not self._should_skip_entity(ent.text, lang):
                            if self._is_valid_pos_for_redaction(ent.text, lang):
                                entities_to_replace.append(ent.text)
        
        # Collect entities using compiled patterns
        entities_to_replace.extend(_findall_compiled_list(PHARMA_COMPANY_RE, result))
        entities_to_replace.extend(_findall_compiled_list(CRO_RE, result))
        entities_to_replace.extend(_findall_compiled_list(BLOCKBUSTER_DRUGS_RE, result))
        entities_to_replace.extend(_findall_compiled_list(CLINICAL_TECH_RE, result))
        entities_to_replace.extend(_findall_compiled_list(CENTRAL_LABS_RE, result))
        
        for compiled in LAB_PRODUCT_RE:
            matches = compiled.findall(result)
            for match in matches:
                if self._is_structured_id(match) or self._is_valid_pos_for_redaction(match, lang):
                    entities_to_replace.append(match)
        entities_to_replace.extend(_findall_compiled_list(ACRONYM_RE, result))
        
        # Apply substitutions using compiled patterns
        result, cnt = _sub_compiled_list(CLINICAL_URL_RE, self.replacement_token, result)
        self.stats["biomedical"] += cnt
        
        result, cnt = _sub_compiled_list(URL_RE, self.replacement_token, result)
        self.stats["biomedical"] += cnt
        
        result, cnt = _sub_compiled_list(ADDRESS_RE, self.replacement_token, result)
        self.stats["biomedical"] += cnt
        
        result, cnt = _sub_compiled_list(INVESTIGATOR_RE, self.replacement_token, result)
        self.stats["biomedical"] += cnt
        
        result, cnt = _sub_compiled_list(CLINICAL_DATE_RE, self.replacement_token, result)
        self.stats["biomedical"] += cnt
        
        result, cnt = _sub_compiled_list(LOT_BATCH_RE, self.replacement_token, result)
        self.stats["biomedical"] += cnt
        
        result, cnt = _sub_compiled_list(SITE_CODE_RE, self.replacement_token, result)
        self.stats["biomedical"] += cnt
        
        result, cnt = _sub_compiled_list(IRB_ETHICS_RE, self.replacement_token, result)
        self.stats["biomedical"] += cnt
        
        result, cnt = _sub_compiled_list(AGE_RANGE_RE, self.replacement_token, result)
        self.stats["biomedical"] += cnt
        
        result, cnt = _sub_compiled_list(HOSPITAL_RE, self.replacement_token, result)
        self.stats["biomedical"] += cnt
        
        result, cnt = _sub_compiled_list(COUNTRY_NATIONALITY_RE, self.replacement_token, result)
        self.stats["biomedical"] += cnt
        
        result, cnt = _sub_compiled_list(STUDY_NAME_RE, self.replacement_token, result)
        self.stats["biomedical"] += cnt
        
        result, cnt = _sub_compiled_list(RANDOMIZATION_RE, self.replacement_token, result)
        self.stats["biomedical"] += cnt
        
        for compiled in MEDICAL_DEVICE_RE:
            matches = compiled.findall(result)
            for match in matches:
                if self._is_valid_pos_for_redaction(match, lang):
                    escaped = re.escape(match)
                    pattern_re = re.compile(r'(?<![A-Za-z0-9À-ÿ])' + escaped + r'(?![A-Za-z0-9À-ÿ])', re.IGNORECASE)
                    match_count = len(pattern_re.findall(result))
                    if match_count > 0:
                        self.stats["biomedical"] += match_count
                        result = pattern_re.sub(self.replacement_token, result)
        
        result, cnt = _sub_compiled_list(INTERNATIONAL_PHONE_RE, self.replacement_token, result)
        self.stats["biomedical"] += cnt
        
        result, cnt = _sub_compiled_list(CORPORATE_EMAIL_RE, self.replacement_token, result)
        self.stats["biomedical"] += cnt
        
        result, cnt = _sub_compiled_list(POSTAL_CODE_RE, self.replacement_token, result)
        self.stats["biomedical"] += cnt
        
        result, cnt = _sub_compiled_list(DOB_RE, self.replacement_token, result)
        self.stats["biomedical"] += cnt
        
        # Biomedical patterns using compiled dict
        for pattern_name, compiled in BIOMEDICAL_PATTERNS_RE.items():
            if pattern_name == "DNA_RNA":
                continue
            matches = compiled.findall(result)
            for match in matches:
                if self._should_skip_entity(match, lang):
                    continue
                if pattern_name in ["DRUG_SUFFIX", "CHEMICAL_COMPOUND"]:
                    if self._is_valid_pos_for_redaction(match, lang):
                        entities_to_replace.append(match)
                elif pattern_name in ["TRADEMARK_NAME", "COPYRIGHT_NAME", "PROBEMIX_PRODUCT"]:
                    entities_to_replace.append(match)
        
        # spaCy NER for ORG/PRODUCT entities
        nlp = self._get_nlp(lang)
        doc = nlp(result)
        for ent in doc.ents:
            if ent.label_ in ["ORG", "PRODUCT", "WORK_OF_ART"]:
                if self._should_skip_entity(ent.text, lang):
                    continue
                if len(ent.text) < 4:
                    continue
                
                if not self._is_valid_pos_for_redaction(ent.text, lang):
                    continue
                
                is_name_like = self._is_name_like(ent.text, lang)
                
                has_context = self._has_local_context(
                    result, 
                    ent.start_char, 
                    ent.end_char, 
                    ORG_CONTEXT_WORDS, 
                    window=40
                )
                
                if is_name_like or has_context:
                    entities_to_replace.append(ent.text)
        
        entities_to_replace = list(set(entities_to_replace))
        entities_to_replace.sort(key=len, reverse=True)
        
        for entity in entities_to_replace:
            if len(entity) > 3:
                escaped = re.escape(entity)
                pattern_re = re.compile(r'(?<![A-Za-z0-9À-ÿ])' + escaped + r'(?![A-Za-z0-9À-ÿ])', re.IGNORECASE)
                match_count = len(pattern_re.findall(result))
                if match_count > 0:
                    self.stats["biomedical"] += match_count
                    result = pattern_re.sub(self.replacement_token, result)
                    self._add_to_cache(entity, lang)
        
        return result
    
    def anonymize_with_dictionary(self, text: str, terms: Set[str]) -> str:
        if not terms:
            return text
        
        result = text
        sorted_terms = sorted(terms, key=len, reverse=True)
        
        for term in sorted_terms:
            if not term.strip():
                continue
            escaped_term = re.escape(term.strip())
            pattern = r'(?<![a-zA-Z0-9À-ÿ])' + escaped_term + r'(?![a-zA-Z0-9À-ÿ])'
            matches = re.findall(pattern, result, re.IGNORECASE)
            if matches:
                self.stats["dictionary"] += len(matches)
                result = re.sub(pattern, self.replacement_token, result, flags=re.IGNORECASE)
        
        return result
    
    def _make_wl_placeholder(self, counter: int) -> str:
        digits = f"{counter:06d}"
        encoded = ''.join(chr(0xE010 + int(d)) for d in digits)
        return f"\uE000{encoded}\uE001"

    def _protect_whitelist_terms(self, text: str, whitelist_terms: Set[str]) -> Tuple[str, List[Tuple[str, str]]]:
        if not whitelist_terms:
            return text, []
        
        replacements = []
        sorted_terms = sorted(whitelist_terms, key=len, reverse=True)
        counter = 0
        
        for term in sorted_terms:
            pattern = re.compile(re.escape(term), re.IGNORECASE)
            matches = list(pattern.finditer(text))
            if matches:
                for match in reversed(matches):
                    placeholder = self._make_wl_placeholder(counter)
                    counter += 1
                    original_match = match.group(0)
                    replacements.append((placeholder, original_match))
                    text = text[:match.start()] + placeholder + text[match.end():]
        
        return text, replacements
    
    def _restore_whitelist_placeholders(self, text: str, replacements: List[Tuple[str, str]]) -> str:
        for placeholder, original in replacements:
            text = text.replace(placeholder, original)
        return text
    
    def process_text_node(self, text: str, 
                          lang: str = "en",
                          use_regex: bool = True,
                          use_presidio: bool = True,
                          use_biomedical: bool = True,
                          use_dictionary: bool = True,
                          dictionary_terms: Set[str] = None,
                          whitelist_terms: Set[str] = None) -> str:
        start_time = time.perf_counter()
        
        if not text or not text.strip():
            return text
        
        wl_placeholders = []
        if whitelist_terms:
            text, wl_placeholders = self._protect_whitelist_terms(text, whitelist_terms)
        
        result = text
        
        if use_regex:
            result = self.anonymize_with_regex_ct(result, lang=lang)
        
        if use_presidio:
            result = self.anonymize_with_presidio(result, lang=lang)
        
        if use_biomedical:
            result = self.anonymize_with_biomedical(result, lang=lang)
        
        if use_dictionary and dictionary_terms:
            result = self.anonymize_with_dictionary(result, dictionary_terms)
        
        if use_regex:
            result = self._detect_proper_nouns_spacy(result, lang=lang)
            result = self._apply_cache(result)
        
        if wl_placeholders:
            result = self._restore_whitelist_placeholders(result, wl_placeholders)
        
        # Validate no critical PII remaining
        pii_check = self._validate_no_critical_pii(result)
        if any(pii_check["counts"].values()):
            self.stats["critical_pii_remaining"] = pii_check
            # Strict mode: raise error if critical PII remains
            if self.strict_validation:
                raise ValueError(f"Critical PII remained after anonymization: {pii_check['counts']}")

        # Record timing
        elapsed_ms = round((time.perf_counter() - start_time) * 1000, 2)
        self.stats["timings"]["last_segment_ms"] = elapsed_ms
        
        # Accumulative timings (when benchmark enabled)
        if self.enable_benchmark:
            self.stats["timings"]["total_ms"] += elapsed_ms
            self.stats["timings"]["segments"] += 1
        
        return result
    
    def _is_inline_tag(self, element) -> bool:
        local_name = etree.QName(element.tag).localname if isinstance(element.tag, str) else None
        return local_name in INLINE_TAG_NAMES if local_name else False
    
    def _process_segment_element(self, element, lang: str = "en", **kwargs):
        """
        Procesa un elemento de segmento usando pipeline linearize→process→rebuild.
        Usa placeholders únicos para preservar la estructura de inline tags.
        """
        children = list(element)
        has_inline = any(self._is_inline_tag(c) for c in children)
        
        if not has_inline:
            if element.text:
                element.text = self.process_text_node(element.text, lang=lang, **kwargs)
            for child in children:
                self._process_segment_element(child, lang=lang, **kwargs)
            return
        
        PH_START = "\uE002PH"
        PH_END = "\uE003"
        
        linearized = ""
        segments = []
        
        if element.text:
            segments.append(("element_text", element.text))
            linearized += element.text
        
        for idx, child in enumerate(children):
            placeholder = f"{PH_START}{idx}{PH_END}"
            segments.append(("placeholder", placeholder, child))
            linearized += placeholder
            
            if child.tail:
                segments.append(("tail", child.tail, child))
                linearized += child.tail
        
        processed = self.process_text_node(linearized, lang=lang, **kwargs)
        
        if processed == linearized:
            return
        
        import re
        ph_pattern = re.compile(f"{re.escape(PH_START)}(\\d+){re.escape(PH_END)}")
        
        parts = []
        last_end = 0
        for m in ph_pattern.finditer(processed):
            if m.start() > last_end:
                parts.append(("text", processed[last_end:m.start()]))
            parts.append(("ph", int(m.group(1))))
            last_end = m.end()
        if last_end < len(processed):
            parts.append(("text", processed[last_end:]))
        
        new_element_text = ""
        child_tails = {child: "" for child in children}
        
        current_target = "element"
        current_child_idx = -1
        
        for part in parts:
            if part[0] == "text":
                text = part[1]
                if current_target == "element":
                    new_element_text += text
                else:
                    child = children[current_child_idx]
                    child_tails[child] += text
            elif part[0] == "ph":
                child_idx = part[1]
                if current_target == "element":
                    current_target = "tail"
                current_child_idx = child_idx
        
        element.text = new_element_text if new_element_text else None
        
        for child in children:
            child.tail = child_tails.get(child) or None
        
        for child in children:
            if not self._is_inline_tag(child):
                self._process_segment_element(child, lang=lang, **kwargs)
    
    def _remove_segment_history(self, tree) -> int:
        """Remove segment version history from MQXLIFF to prevent data leakage.
        
        Removes:
        - <mq:historical-unit> elements (contain old segment versions)
        - Empty <mq:minorversions> elements after cleanup
        - Updates mq:hashistory to "false" for consistency
        
        Returns count of removed elements.
        """
        removed_count = 0
        
        historical_units = tree.xpath("//mq:historical-unit", namespaces=XLIFF_NS)
        for hu in historical_units:
            parent = hu.getparent()
            if parent is not None:
                parent.remove(hu)
                removed_count += 1
        
        # Keep empty mq:minorversions elements - MemoQ requires them for RowHistory
        
        if removed_count > 0:
            mq_ns = XLIFF_NS.get("mq", "MQXliff")
            docinfo_elements = tree.xpath("//mq:docinformation", namespaces=XLIFF_NS)
            for docinfo in docinfo_elements:
                hashistory_attr = f"{{{mq_ns}}}hashistory"
                if hashistory_attr in docinfo.attrib:
                    docinfo.attrib[hashistory_attr] = "false"
        
        return removed_count
    
    def anonymize_mqxliff(self, xml_content: bytes, 
                          process_source: bool = True,
                          process_target: bool = True,
                          use_regex: bool = True,
                          use_presidio: bool = True,
                          use_biomedical: bool = True,
                          use_dictionary: bool = True,
                          dictionary_terms: Set[str] = None,
                          whitelist_terms: Set[str] = None) -> Tuple[bytes, Dict[str, int], List[Dict]]:
        self.reset_stats()
        previews = []
        
        try:
            xml_content = self._normalize_xml_input(xml_content)
            parser = etree.XMLParser(remove_blank_text=False, strip_cdata=False)
            tree = etree.fromstring(xml_content, parser=parser)
        except etree.XMLSyntaxError as e:
            raise ValueError(f"Error parsing MQXLIFF: {str(e)}")
        
        trans_units = tree.xpath("//xliff:trans-unit", namespaces=XLIFF_NS)
        
        if use_regex:
            self._scan_document_for_lowercase(tree)
        
        process_kwargs = {
            "use_regex": use_regex,
            "use_presidio": use_presidio,
            "use_biomedical": use_biomedical,
            "use_dictionary": use_dictionary,
            "dictionary_terms": dictionary_terms,
            "whitelist_terms": whitelist_terms
        }
        
        original_sources = {}
        original_targets = {}
        
        for i, tu in enumerate(trans_units):
            source_before = ""
            target_before = ""
            source_after = ""
            target_after = ""
            
            sources = tu.xpath("xliff:source", namespaces=XLIFF_NS)
            targets = tu.xpath("xliff:target", namespaces=XLIFF_NS)
            
            if sources and process_source:
                source = sources[0]
                source_before = self._get_element_text_content(source)
                original_sources[i] = source_before
                self._process_segment_element(source, lang="en", **process_kwargs)
                source_after = self._get_element_text_content(source)
            
            if targets and process_target:
                target = targets[0]
                target_before = self._get_element_text_content(target)
                original_targets[i] = target_before
                self._process_segment_element(target, lang="es", **process_kwargs)
                target_after = self._get_element_text_content(target)
            
            if source_before != source_after or target_before != target_after:
                previews.append({
                    "segment": i + 1,
                    "source_before": source_before,
                    "source_after": source_after,
                    "target_before": target_before,
                    "target_after": target_after
                })
        
        if whitelist_terms:
            wl_lower = {t.lower() for t in whitelist_terms}
            self.terms_cache -= wl_lower
        
        if use_regex and self.terms_cache:
            self._build_cache_regex()
            for i, tu in enumerate(trans_units):
                if process_source:
                    sources = tu.xpath("xliff:source", namespaces=XLIFF_NS)
                    if sources:
                        source = sources[0]
                        text_before_pass2 = self._get_element_text_content(source)
                        self._apply_cache_to_element(source)
                        text_after_pass2 = self._get_element_text_content(source)
                        if text_before_pass2 != text_after_pass2 and i in original_sources:
                            existing = next((p for p in previews if p["segment"] == i + 1), None)
                            if existing:
                                existing["source_after"] = text_after_pass2
                            else:
                                previews.append({
                                    "segment": i + 1,
                                    "source_before": original_sources[i],
                                    "source_after": text_after_pass2,
                                    "target_before": original_targets.get(i, ""),
                                    "target_after": self._get_element_text_content(
                                        tu.xpath("xliff:target", namespaces=XLIFF_NS)[0]
                                    ) if tu.xpath("xliff:target", namespaces=XLIFF_NS) else ""
                                })
                
                if process_target:
                    targets = tu.xpath("xliff:target", namespaces=XLIFF_NS)
                    if targets:
                        target = targets[0]
                        text_before_pass2 = self._get_element_text_content(target)
                        self._apply_cache_to_element(target)
                        text_after_pass2 = self._get_element_text_content(target)
                        if text_before_pass2 != text_after_pass2:
                            existing = next((p for p in previews if p["segment"] == i + 1), None)
                            if existing:
                                existing["target_after"] = text_after_pass2
                            else:
                                previews.append({
                                    "segment": i + 1,
                                    "source_before": original_sources.get(i, ""),
                                    "source_after": self._get_element_text_content(
                                        tu.xpath("xliff:source", namespaces=XLIFF_NS)[0]
                                    ) if tu.xpath("xliff:source", namespaces=XLIFF_NS) else "",
                                    "target_before": original_targets.get(i, ""),
                                    "target_after": text_after_pass2
                                })
            
            previews.sort(key=lambda p: p["segment"])
        
        history_removed = self._remove_segment_history(tree)
        if history_removed > 0:
            self.stats["history_removed"] = history_removed
        
        self._clean_xml_tree(tree)
        result_xml = etree.tostring(tree, xml_declaration=True, encoding="UTF-8")
        
        result_xml = self._normalize_xml_format(result_xml)
        
        return result_xml, self.stats.copy(), previews
    
    def _normalize_xml_format(self, xml_bytes: bytes) -> bytes:
        """Normalize XML format for memoQ compatibility.
        
        Fixes lxml serialization quirks:
        - Uses double quotes in XML declaration
        - Restores space before /> in self-closing tags
        - Preserves version attribute position in xliff element
        """
        import re
        
        result = xml_bytes.replace(
            b"<?xml version='1.0' encoding='UTF-8'?>",
            b'<?xml version="1.0" encoding="UTF-8"?>'
        )
        
        result = re.sub(b'([^ "\'])/>', rb'\1 />', result)
        
        result = re.sub(
            b'<xliff xmlns="([^"]+)" xmlns:mq="([^"]+)" xmlns:xsi="([^"]+)" version="([^"]+)" xsi:schemaLocation="([^"]+)">',
            rb'<xliff version="\4" xmlns="\1" xmlns:mq="\2" xmlns:xsi="\3" xsi:schemaLocation="\5">',
            result
        )
        
        return result
    
    def _get_element_text_content(self, element) -> str:
        return "".join(element.itertext())

    def _normalize_xml_input(self, xml_content: bytes) -> bytes:
        if xml_content[:3] == b'\xef\xbb\xbf':
            xml_content = xml_content[3:]
        elif xml_content[:2] in (b'\xff\xfe', b'\xfe\xff'):
            try:
                encoding = 'utf-16-le' if xml_content[:2] == b'\xff\xfe' else 'utf-16-be'
                text = xml_content[2:].decode(encoding)
                xml_content = text.encode('utf-8')
                xml_content = re.sub(rb'encoding=["\'][^"\']*["\']', rb'encoding="UTF-8"', xml_content)
            except (UnicodeDecodeError, UnicodeEncodeError):
                pass
        xml_content = re.sub(rb'[\x00-\x08\x0b\x0c\x0e-\x1f]', b'', xml_content)
        return xml_content

    def _clean_xml_tree(self, tree):
        _ctrl_re = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')
        for element in tree.iter():
            if element.text and _ctrl_re.search(element.text):
                element.text = _ctrl_re.sub('', element.text)
            if element.tail and _ctrl_re.search(element.tail):
                element.tail = _ctrl_re.sub('', element.tail)

    def _scan_document_for_lowercase_tmx(self, tree):
        trans_units = tree.xpath("//tu")
        
        email_url_pattern = re.compile(
            r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}|'
            r'https?://[^\s]+|'
            r'www\.[^\s]+|'
            r'[a-zA-Z0-9.-]+\.(com|org|net|edu|gov|io|co|es|de|fr|uk|eu)[^\s]*',
            re.IGNORECASE
        )
        
        for tu in trans_units:
            for tuv in tu.xpath("tuv"):
                for seg in tuv.xpath("seg"):
                    text = "".join(seg.itertext())
                    if text:
                        clean_text = email_url_pattern.sub(' ', text)
                        words = re.findall(r'\b[a-záéíóúñüàèìòùäëïöü]+\b', clean_text, re.IGNORECASE)
                        for word in words:
                            if word.islower() and len(word) > 2:
                                self.lowercase_words.add(word.lower())

    def _get_tmx_tuv_by_lang(self, tu, lang_code):
        for tuv in tu.xpath("tuv"):
            tuv_lang = tuv.get("{http://www.w3.org/XML/1998/namespace}lang", "")
            if not tuv_lang:
                tuv_lang = tuv.get("lang", "")
            if tuv_lang.lower().startswith(lang_code.lower()):
                segs = tuv.xpath("seg")
                if segs:
                    return segs[0]
        return None

    def anonymize_tmx(self, xml_content: bytes,
                      process_source: bool = True,
                      process_target: bool = True,
                      source_lang: str = "en",
                      target_lang: str = "es",
                      use_regex: bool = True,
                      use_presidio: bool = True,
                      use_biomedical: bool = True,
                      use_dictionary: bool = True,
                      dictionary_terms: Set[str] = None,
                      whitelist_terms: Set[str] = None) -> Tuple[bytes, Dict[str, int], List[Dict]]:
        self.reset_stats()
        previews = []
        
        try:
            xml_content = self._normalize_xml_input(xml_content)
            parser = etree.XMLParser(remove_blank_text=False, strip_cdata=False)
            tree = etree.fromstring(xml_content, parser=parser)
        except etree.XMLSyntaxError as e:
            raise ValueError(f"Error parsing TMX: {str(e)}")
        
        trans_units = tree.xpath("//tu")
        
        if use_regex:
            self._scan_document_for_lowercase_tmx(tree)
        
        process_kwargs = {
            "use_regex": use_regex,
            "use_presidio": use_presidio,
            "use_biomedical": use_biomedical,
            "use_dictionary": use_dictionary,
            "dictionary_terms": dictionary_terms,
            "whitelist_terms": whitelist_terms
        }
        
        original_sources = {}
        original_targets = {}
        
        for i, tu in enumerate(trans_units):
            source_before = ""
            target_before = ""
            source_after = ""
            target_after = ""
            
            source_seg = self._get_tmx_tuv_by_lang(tu, source_lang)
            target_seg = self._get_tmx_tuv_by_lang(tu, target_lang)
            
            if source_seg is not None and process_source:
                source_before = self._get_element_text_content(source_seg)
                original_sources[i] = source_before
                self._process_segment_element(source_seg, lang=source_lang[:2], **process_kwargs)
                source_after = self._get_element_text_content(source_seg)
            
            if target_seg is not None and process_target:
                target_before = self._get_element_text_content(target_seg)
                original_targets[i] = target_before
                self._process_segment_element(target_seg, lang=target_lang[:2], **process_kwargs)
                target_after = self._get_element_text_content(target_seg)
            
            if source_before != source_after or target_before != target_after:
                previews.append({
                    "segment": i + 1,
                    "source_before": source_before,
                    "source_after": source_after,
                    "target_before": target_before,
                    "target_after": target_after
                })
        
        if whitelist_terms:
            wl_lower = {t.lower() for t in whitelist_terms}
            self.terms_cache -= wl_lower
        
        if use_regex and self.terms_cache:
            self._build_cache_regex()
            for i, tu in enumerate(trans_units):
                if process_source:
                    source_seg = self._get_tmx_tuv_by_lang(tu, source_lang)
                    if source_seg is not None:
                        text_before_pass2 = self._get_element_text_content(source_seg)
                        self._apply_cache_to_element(source_seg)
                        text_after_pass2 = self._get_element_text_content(source_seg)
                        if text_before_pass2 != text_after_pass2 and i in original_sources:
                            existing = next((p for p in previews if p["segment"] == i + 1), None)
                            if existing:
                                existing["source_after"] = text_after_pass2
                            else:
                                target_seg = self._get_tmx_tuv_by_lang(tu, target_lang)
                                previews.append({
                                    "segment": i + 1,
                                    "source_before": original_sources[i],
                                    "source_after": text_after_pass2,
                                    "target_before": original_targets.get(i, ""),
                                    "target_after": self._get_element_text_content(target_seg) if target_seg is not None else ""
                                })
                
                if process_target:
                    target_seg = self._get_tmx_tuv_by_lang(tu, target_lang)
                    if target_seg is not None:
                        text_before_pass2 = self._get_element_text_content(target_seg)
                        self._apply_cache_to_element(target_seg)
                        text_after_pass2 = self._get_element_text_content(target_seg)
                        if text_before_pass2 != text_after_pass2:
                            existing = next((p for p in previews if p["segment"] == i + 1), None)
                            if existing:
                                existing["target_after"] = text_after_pass2
                            else:
                                source_seg = self._get_tmx_tuv_by_lang(tu, source_lang)
                                previews.append({
                                    "segment": i + 1,
                                    "source_before": original_sources.get(i, ""),
                                    "source_after": self._get_element_text_content(source_seg) if source_seg is not None else "",
                                    "target_before": original_targets.get(i, ""),
                                    "target_after": text_after_pass2
                                })
            
            previews.sort(key=lambda p: p["segment"])
        
        self._clean_xml_tree(tree)
        result_xml = etree.tostring(tree, xml_declaration=True, encoding="UTF-8")
        
        result_xml = result_xml.replace(
            b"<?xml version='1.0' encoding='UTF-8'?>",
            b'<?xml version="1.0" encoding="UTF-8"?>'
        )
        
        return result_xml, self.stats.copy(), previews


def _strip_control_chars(text: str) -> str:
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)


def load_dictionary_terms(file_content: str) -> Set[str]:
    file_content = _strip_control_chars(file_content)
    terms = set()
    for line in file_content.strip().split('\n'):
        if ',' in line:
            for part in line.split(','):
                term = part.strip()
                if term:
                    terms.add(term)
        else:
            term = line.strip()
            if term:
                terms.add(term)
    return terms
