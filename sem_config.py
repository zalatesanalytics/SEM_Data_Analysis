"""Configuration for the EMERGE+ Newcomer SEM analysis app."""

LATENT_VARIABLES = {
    "CRB": {
        "name": "Credential Recognition Barriers",
        "indicators": [
            "cred_req_difficulty", "cred_info_access", "cred_financial_burden",
            "cred_emotional_burden", "cred_delay_employment",
        ],
    },
    "WD": {
        "name": "Workplace Discrimination",
        "indicators": [
            "disc_hiring", "disc_workplace", "disc_network_exclusion",
            "disc_foreign_education", "disc_accent",
        ],
    },
    "FP": {
        "name": "Financial Pressure",
        "indicators": [
            "rent_burden", "debt_burden", "remittance_burden",
            "housing_insecurity", "food_insecurity", "ability_save",
        ],
    },
    "SC": {
        "name": "Social Capital",
        "indicators": [
            "professional_contacts", "mentors_count", "networking_participation",
            "trust_support_systems",
        ],
    },
    "MHB": {
        "name": "Mental Health Burden",
        "indicators": [
            "burnout", "employment_anxiety", "hopeless_career", "exhaustion",
            "social_isolation", "overwhelmed",
        ],
    },
    "DR": {
        "name": "Digital Readiness",
        "indicators": [
            "online_job_confidence", "linkedin_confidence", "ai_career_confidence",
            "online_learning_participation",
        ],
    },
    "FR": {
        "name": "Family Responsibilities",
        "indicators": ["family_employment_impact", "family_stress", "caregiving_hours"],
    },
    "ME": {
        "name": "Mentorship Effectiveness",
        "indicators": ["mentor_sector_match", "mentor_background_alignment", "mentor_referral_support"],
    },
}

OBSERVED_OUTCOMES = [
    "employment_confidence", "job_satisfaction", "financial_stability",
    "mental_wellbeing", "future_confidence", "belonging_canada",
    "employment_integration", "economic_stability", "career_satisfaction",
]

STRUCTURAL_PATHS = [
    ("CRB", "MHB"), ("CRB", "employment_confidence"),
    ("WD", "MHB"), ("WD", "job_satisfaction"),
    ("FP", "MHB"), ("FP", "financial_stability"),
    ("SC", "employment_confidence"), ("SC", "job_satisfaction"),
    ("DR", "employment_confidence"), ("MHB", "employment_confidence"),
    ("MHB", "job_satisfaction"), ("ME", "SC"),
    ("ME", "employment_confidence"), ("FR", "MHB"),
    ("FR", "employment_confidence"),
]

MEDIATION_PATHS = [
    {"name": "Discrimination → Mental Health → Employment", "iv": "WD", "mediator": "MHB", "dv": "employment_confidence"},
    {"name": "Credential Barriers → Mental Health → Job Satisfaction", "iv": "CRB", "mediator": "MHB", "dv": "job_satisfaction"},
    {"name": "Mentorship → Social Capital → Employment", "iv": "ME", "mediator": "SC", "dv": "employment_confidence"},
    {"name": "Financial Pressure → Mental Health → Employment", "iv": "FP", "mediator": "MHB", "dv": "employment_confidence"},
]

FIT_THRESHOLDS = {"CFI": 0.90, "TLI": 0.90, "RMSEA": 0.08, "SRMR": 0.08}
