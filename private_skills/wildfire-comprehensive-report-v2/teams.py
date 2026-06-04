"""
teams.py — TEAMS alias table for the Fire Risk Modeling Exercise.

Each entry maps one of the 20 submitting teams to its various names across
the source files (NDP workspace folder, evaluator xlsx, tracking xlsx). The
team_id column matches the Deep_Synthesis__<team_id>_(scenario) folder naming
used by the rubric-review skill.

This table is project-specific and lives inside the v2 comprehensive-report
skill. Next year's cohort will need a fresh table.
"""

TEAMS = [
    # --- 14 teams with NDP workspaces ---
    {
        "team_id": "Minerva_University",
        "tracking_org": "Minerva University",
        "tracking_model": "Cell2Fire W",
        "evaluator_orgs": ["Minerva University"],
        "evaluator_model_filter": None,
        "combined_report": False,
    },
    {
        "team_id": "UCI_CHRS",
        "tracking_org": "UC Irvine",
        "tracking_model": "Wildfire Intelligence System",
        "evaluator_orgs": ["UC Irvine"],
        "evaluator_model_filter": None,
        "combined_report": False,
    },
    {
        "team_id": "IMAGECAT",
        "tracking_org": "ImageCat",
        "tracking_model": "FireImpact",
        "evaluator_orgs": ["IMAGECAT", "ImageCat"],
        "evaluator_model_filter": None,
        "combined_report": False,
    },
    {
        "team_id": "ELMFIRE",
        "tracking_org": "UC Berkley",
        "tracking_model": "ElmFire",
        "evaluator_orgs": ["UC Berkley / ElmFire", "UC Berkley", "UC Berkeley"],
        "evaluator_model_filter": None,
        "combined_report": False,
    },
    {
        "team_id": "Pyrologix_and_Vibrant_Planet",
        "tracking_org": "Vibrant Planet",
        "tracking_model": "Vibrant Planet",
        "evaluator_orgs": ["Vibrant Planet", "Pyrologix", "Vibrant Planet / Pyrologix"],
        "evaluator_model_filter": None,
        "combined_report": False,
    },
    {
        "team_id": "San_Jose_State",
        "tracking_org": "San Jose State University",
        "tracking_model": "WRF-SFIRE",
        "evaluator_orgs": ["San Jose State University", "SJSU", "WIRC/SJSU"],
        "evaluator_model_filter": None,
        "combined_report": False,
    },
    {
        "team_id": "OroraTech",
        "tracking_org": "OroraTech",
        "tracking_model": "Wildfire Property Risk",
        "evaluator_orgs": ["OroraTech"],
        "evaluator_model_filter": None,
        "combined_report": False,
    },
    {
        "team_id": "SIG",
        "tracking_org": "Spatial Informatics Group",
        "tracking_model": "Ensemble: Pyretechnics, Pyregence, QARA/HVRA post-processing, Ember Wash",
        "evaluator_orgs": ["Spatial Informatics Group", "SIG", "Pyregence"],
        "evaluator_model_filter": None,
        "combined_report": True,
    },
    {
        "team_id": "Firesafe",
        "tracking_org": "FireSafe Analytics",
        "tracking_model": "FireSafe Analytics",
        "evaluator_orgs": ["FireSafe Analytics", "Firesafe"],
        "evaluator_model_filter": None,
        "combined_report": False,
    },
    {
        "team_id": "SkyTL",
        "tracking_org": "SkyTL",
        "tracking_model": "WindTL AGNI-NAR",
        "evaluator_orgs": ["SkyTL"],
        "evaluator_model_filter": None,
        "combined_report": False,
    },
    {
        "team_id": "UNR",
        "tracking_org": "University of Nevada",
        "tracking_model": "Ensemble Risk Assessment: fire spread + ML + EAL",
        "evaluator_orgs": ["University of Nevada", "Univ of Nevada", "UNR"],
        "evaluator_model_filter": None,
        "combined_report": False,
    },
    {
        "team_id": "FiSci",
        "tracking_org": "FiSci",
        "tracking_model": "FiSci",
        "evaluator_orgs": ["FiSci"],
        "evaluator_model_filter": None,
        "combined_report": False,
    },
    {
        "team_id": "French_National_Center_for_Scientific_Research_team",
        "tracking_org": "French National Centre for Scientific Research",
        "tracking_model": "FireCaster",
        "evaluator_orgs": [
            "French National Center for Science",
            "French National Centre for Scientific Research",
            "CNRS",
        ],
        "evaluator_model_filter": None,
        "combined_report": True,
    },
    {
        "team_id": "XyloPlan",
        "tracking_org": "XyloPlan",
        "tracking_model": "Urban Fire Model",
        "evaluator_orgs": ["XyloPlan"],
        "evaluator_model_filter": None,
        "combined_report": False,
    },

    # --- 6 teams without NDP workspaces (synthetic workspaces in missing_rubrics/) ---
    {
        "team_id": "PinePeak",
        "tracking_org": "PinePeak",
        "tracking_model": "FlameSight",
        "evaluator_orgs": ["PinePeak", "Pinepeak"],
        "evaluator_model_filter": None,
        "combined_report": True,
    },
    {
        "team_id": "Vanderbilt_WUI_ANN",
        "tracking_org": "Vanderbilt University",
        "tracking_model": "WUI - Artifical Neural Network Model",
        "evaluator_orgs": ["Vanderbilt University"],
        "evaluator_model_filter": "Artificial Neural",
        "combined_report": True,
    },
    {
        "team_id": "Vanderbilt_WUI_AGNI_NAR",
        "tracking_org": "Vanderbilt University",
        "tracking_model": "WUI - AGNI-NAR Model",
        "evaluator_orgs": ["Vanderbilt University"],
        "evaluator_model_filter": "AGNI-NAR",
        "combined_report": True,
    },
    {
        "team_id": "FlameMapper",
        "tracking_org": "FlameMapper",
        "tracking_model": "EmberCast",
        "evaluator_orgs": ["FlameMapper", "EmberCast"],
        "evaluator_model_filter": None,
        "combined_report": True,
    },
    {
        "team_id": "Zesty_AI",
        "tracking_org": "Zesty AI",
        "tracking_model": "Z-Fire L2 Model",
        "evaluator_orgs": ["Zesty AI", "ZestyAI"],
        "evaluator_model_filter": None,
        "combined_report": True,
    },
    {
        "team_id": "TenaxAI",
        "tracking_org": "TenaxAI",
        "tracking_model": "Tenax ai Structural Vulnerability Score",
        "evaluator_orgs": ["Tenax AI", "TenaxAI", "Tenex.ai"],
        "evaluator_model_filter": None,
        "combined_report": False,
    },
]


def team_for(team_id: str) -> dict | None:
    for t in TEAMS:
        if t["team_id"] == team_id:
            return t
    return None


def parse_workspace_folder(folder_name: str) -> tuple[str | None, str | None]:
    """Parse 'Deep_Synthesis__<team_id>_(Forest|Prairie)' → (team_id, scenario)."""
    import re
    m = re.match(r"^Deep_Synthesis__(.+)_\((Forest|Prairie)\)$", folder_name)
    if not m:
        return (None, None)
    team_id, scenario = m.group(1), m.group(2)
    if team_for(team_id) is None:
        return (None, scenario)
    return (team_id, scenario)
