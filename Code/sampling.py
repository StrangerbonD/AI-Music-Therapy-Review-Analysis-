"""
Phase 2: Stratified Random Sampling of 5,000 Reviews
---------------------------------------------------
Executes a two-stage stratified random sampling pipeline without replacement:
- Stage 1: Application-level stratification (Predefined target sample sizes)
- Stage 2: Rating-level proportional stratification using the Largest-Remainder
           (Hare-Niemeyer) method.

Random Seed: 42 (Reproducible)
Input File : combined_clean_reviews.csv (N = 104,857)
Output Dir : phase_2_sampling/
"""

import os
import logging
from typing import Dict, Any, Tuple

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
logger = logging.getLogger("sampling")

RANDOM_STATE = 42

APPLICATION_TARGETS = {
    "Day One": 1000,
    "Daylio": 1538,
    "Wysa": 1537,
    "Mindsera": 66,
    "Reflection": 859,
}
TOTAL_TARGET = sum(APPLICATION_TARGETS.values())  # 5,000

EXPECTED_COLUMNS = [
    "App_Name",
    "Review_ID",
    "User_Name",
    "Review",
    "Rating",
    "Review_Date",
    "Helpful_Votes",
    "App_Version",
    "Developer_Reply",
    "Reply_Date"
]


def resolve_paths() -> Dict[str, str]:
    """Resolves all project directories and file destinations."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        script_dir,
        os.path.abspath(os.path.join(script_dir, "..")),
        r"E:\Insight"
    ]
    project_root = next((p for p in candidates if os.path.exists(os.path.join(p, "combined_clean_reviews.csv"))), r"E:\Insight")
    output_dir = os.path.join(project_root, "phase_2_sampling")
    os.makedirs(output_dir, exist_ok=True)

    return {
        "project_root": project_root,
        "clean_csv": os.path.join(project_root, "combined_clean_reviews.csv"),
        "output_dir": output_dir,
        "sample_csv": os.path.join(output_dir, "research_sample_5000.csv"),
        "report_xlsx": os.path.join(output_dir, "sampling_distribution_report.xlsx"),
        "summary_csv": os.path.join(output_dir, "sampling_summary.csv"),
        "plot_png": os.path.join(output_dir, "sampling_distribution.png"),
        "script_py": os.path.join(output_dir, "sampling.py"),
        "req_txt": os.path.join(output_dir, "requirements.txt"),
    }


def compute_largest_remainder_allocation(app_df: pd.DataFrame, target_n: int) -> Tuple[pd.Series, pd.DataFrame]:
    """
    Computes rating-wise proportional sample allocations using the
    Largest-Remainder (Hare-Niemeyer) method.
    """
    pop_total = len(app_df)
    counts = app_df["Rating"].value_counts().sort_index()

    # If population equals target, return exact census
    if pop_total == target_n:
        detail_df = pd.DataFrame({
            "Population": counts,
            "Exact_Quota": counts.astype(float),
            "Floor_Quota": counts,
            "Remainder": 0.0,
            "Rank": 0,
            "Adjustment": 0,
            "Allocated_Sample": counts
        })
        return counts, detail_df

    # 1. Exact proportional allocation
    exact_quota = counts * (target_n / pop_total)
    floored = exact_quota.apply(np.floor).astype(int)
    remainders = exact_quota - floored
    extra_slots = target_n - floored.sum()

    # 2. Rank remainders descending
    rem_df = pd.DataFrame({
        "Rating": counts.index,
        "Population": counts.values,
        "Exact_Quota": exact_quota.values,
        "Floor_Quota": floored.values,
        "Remainder": remainders.values,
    }).sort_values(by=["Remainder", "Population"], ascending=[False, False])

    rem_df["Rank"] = range(1, len(rem_df) + 1)
    rem_df["Adjustment"] = (rem_df["Rank"] <= extra_slots).astype(int)
    rem_df["Allocated_Sample"] = rem_df["Floor_Quota"] + rem_df["Adjustment"]

    detail_df = rem_df.sort_values("Rating").set_index("Rating")
    final_alloc = detail_df["Allocated_Sample"]

    assert final_alloc.sum() == target_n, f"Allocation sum {final_alloc.sum()} != target {target_n}"
    return final_alloc, detail_df


def execute_sampling(clean_df: pd.DataFrame, random_seed: int = RANDOM_STATE) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """
    Executes two-stage stratified random sampling across Application and Rating.
    """
    logger.info("Executing two-stage stratified sampling with seed = %d...", random_seed)
    sampled_subsets = []
    allocation_details = {}

    for app_name, target_n in APPLICATION_TARGETS.items():
        app_df = clean_df[clean_df["App_Name"] == app_name]
        final_alloc, detail_df = compute_largest_remainder_allocation(app_df, target_n)
        allocation_details[app_name] = detail_df

        logger.info("App [%s]: Pop = %d, Target = %d. Sampling across 5 rating strata...", app_name, len(app_df), target_n)

        for rating, n_samples in final_alloc.items():
            stratum_df = app_df[app_df["Rating"] == rating]

            if len(stratum_df) == n_samples:
                # 100% census stratum (e.g. Mindsera, Reflection)
                sample_stratum = stratum_df.copy()
            else:
                sample_stratum = stratum_df.sample(
                    n=int(n_samples),
                    replace=False,
                    random_state=random_seed
                )

            sampled_subsets.append(sample_stratum)

    sample_5000 = pd.concat(sampled_subsets, ignore_index=True)
    # Ensure column order matches exactly
    sample_5000 = sample_5000[EXPECTED_COLUMNS]

    logger.info("Sampling complete: selected %d records across %d columns.", len(sample_5000), len(sample_5000.columns))
    return sample_5000, allocation_details


def validate_sample(sample_df: pd.DataFrame, clean_df: pd.DataFrame, allocation_details: Dict[str, pd.DataFrame]) -> bool:
    """
    Strict validation of all 8 quality gates for Phase 2 sampling.
    """
    logger.info("Running validation checks on sampled dataset...")

    # 1. Exactly 5,000 rows
    assert len(sample_df) == TOTAL_TARGET, f"Sample size {len(sample_df)} != {TOTAL_TARGET}"

    # 2. Every Review_ID is unique
    assert sample_df["Review_ID"].nunique() == TOTAL_TARGET, "Found duplicate Review_IDs in sample."

    # 3. No review appears more than once
    assert sample_df.duplicated(subset=["Review_ID"]).sum() == 0, "Duplicate reviews detected."

    # 4. App sample counts match specified targets
    app_counts = sample_df["App_Name"].value_counts().to_dict()
    for app, target in APPLICATION_TARGETS.items():
        assert app_counts.get(app, 0) == target, f"App {app} count {app_counts.get(app)} != {target}"

    # 5. Rating allocations match largest-remainder calculations
    for app, detail_df in allocation_details.items():
        app_sub = sample_df[sample_df["App_Name"] == app]
        sample_ratings = app_sub["Rating"].value_counts().to_dict()
        for rating, exp_n in detail_df["Allocated_Sample"].items():
            act_n = sample_ratings.get(rating, 0)
            assert act_n == exp_n, f"Mismatch in {app} rating {rating}: {act_n} != {exp_n}"

    # 6. Every selected record exists in original cleaned dataset
    clean_ids = set(clean_df["Review_ID"])
    sample_ids = set(sample_df["Review_ID"])
    assert sample_ids.issubset(clean_ids), "Sample contains IDs not present in clean dataset."

    # 7. Sampling was performed without replacement
    for app in APPLICATION_TARGETS:
        for r in range(1, 6):
            stratum = sample_df[(sample_df["App_Name"] == app) & (sample_df["Rating"] == r)]
            assert stratum["Review_ID"].nunique() == len(stratum), f"Replacement detected in {app} {r}★"

    # 8. Deterministic reproducibility check
    sample_retest, _ = execute_sampling(clean_df, random_seed=RANDOM_STATE)
    assert sample_df["Review_ID"].equals(sample_retest["Review_ID"]), "Reproducibility failed with same random_state!"

    logger.info("All 8 Step 5 validation checks passed with 100% precision!")
    return True


def generate_visualization(clean_df: pd.DataFrame, sample_df: pd.DataFrame, plot_path: str):
    """
    Creates a publication-quality visual comparison of Population vs Sample distributions.
    """
    logger.info("Generating sampling distribution visualization: %s", plot_path)
    plt.style.use("seaborn-v0_8-whitegrid" if "seaborn-v0_8-whitegrid" in plt.style.available else "default")

    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(16, 11), dpi=300)
    rating_colors = ["#d9534f", "#f0ad4e", "#f7d070", "#5bc0de", "#5cb85c"]
    app_order = ["Day One", "Daylio", "Wysa", "Reflection", "Mindsera"]

    # --- Panel A: Application Share Comparison ---
    pop_app_pct = (clean_df["App_Name"].value_counts(normalize=True) * 100).loc[app_order]
    sam_app_pct = (sample_df["App_Name"].value_counts(normalize=True) * 100).loc[app_order]

    x = np.arange(len(app_order))
    width = 0.35

    ax1.bar(x - width/2, pop_app_pct, width, label="Population (N=104,857)", color="#4682B4", edgecolor="black", linewidth=0.5)
    ax1.bar(x + width/2, sam_app_pct, width, label="Sample (n=5,000)", color="#2E8B57", edgecolor="black", linewidth=0.5)

    ax1.set_xticks(x)
    ax1.set_xticklabels(app_order, fontweight="bold", fontsize=10)
    ax1.set_ylabel("Share of Total Dataset (%)", fontweight="bold", fontsize=11)
    ax1.set_title("A. Application Allocation: Population vs. Research Sample", fontweight="bold", fontsize=12, pad=10)
    ax1.legend(frameon=True)
    ax1.grid(axis="y", linestyle="--", alpha=0.6)

    # --- Panel B: Overall Rating Distribution Comparison ---
    pop_rate_pct = (clean_df["Rating"].value_counts(normalize=True).sort_index() * 100)
    sam_rate_pct = (sample_df["Rating"].value_counts(normalize=True).sort_index() * 100)

    xr = np.arange(1, 6)
    ax2.bar(xr - width/2, pop_rate_pct, width, label="Population (N=104,857)", color="#4682B4", edgecolor="black", linewidth=0.5)
    ax2.bar(xr + width/2, sam_rate_pct, width, label="Sample (n=5,000)", color="#E67E22", edgecolor="black", linewidth=0.5)

    ax2.set_xticks(xr)
    ax2.set_xticklabels([f"{r} Star" for r in xr], fontweight="bold", fontsize=10)
    ax2.set_ylabel("Share of Total Dataset (%)", fontweight="bold", fontsize=11)
    ax2.set_title("B. Overall Rating Distribution: Population vs. Sample", fontweight="bold", fontsize=12, pad=10)
    ax2.legend(frameon=True)
    ax2.grid(axis="y", linestyle="--", alpha=0.6)

    # --- Panel C: Population Rating Profiles by Application (100% Stacked) ---
    pop_crosstab_pct = pd.crosstab(clean_df["App_Name"], clean_df["Rating"], normalize="index").loc[app_order] * 100
    bottom_pop = np.zeros(len(app_order))
    for r in range(1, 6):
        vals = pop_crosstab_pct[r].values
        ax3.barh(app_order, vals, left=bottom_pop, color=rating_colors[r - 1], label=f"{r} Star", edgecolor="white", height=0.55)
        bottom_pop += vals

    ax3.set_xlim(0, 100)
    ax3.set_xlabel("Rating Percentage Within App (%)", fontweight="bold", fontsize=11)
    ax3.set_title("C. Population Rating Breakdown per Application", fontweight="bold", fontsize=12, pad=10)
    ax3.grid(axis="x", linestyle="--", alpha=0.6)

    # --- Panel D: Sample Rating Profiles by Application (100% Stacked) ---
    sam_crosstab_pct = pd.crosstab(sample_df["App_Name"], sample_df["Rating"], normalize="index").loc[app_order] * 100
    bottom_sam = np.zeros(len(app_order))
    for r in range(1, 6):
        vals = sam_crosstab_pct[r].values
        ax4.barh(app_order, vals, left=bottom_sam, color=rating_colors[r - 1], label=f"{r} Star", edgecolor="white", height=0.55)
        bottom_sam += vals

    ax4.set_xlim(0, 100)
    ax4.set_xlabel("Rating Percentage Within App (%)", fontweight="bold", fontsize=11)
    ax4.set_title("D. Sample Rating Breakdown per Application (Proportional Match)", fontweight="bold", fontsize=12, pad=10)
    ax4.legend(title="Rating", bbox_to_anchor=(0.5, -0.15), loc="upper center", ncol=5, frameon=True)
    ax4.grid(axis="x", linestyle="--", alpha=0.6)

    plt.suptitle("Phase 2 Stratified Sampling Audit: Population vs. Sample Allocations (n = 5,000)", fontsize=15, fontweight="bold", y=0.99)
    plt.tight_layout()
    plt.savefig(plot_path, dpi=300, bbox_inches="tight")
    plt.close()
    logger.info("Saved sampling distribution chart successfully.")


def generate_excel_report(report_path: str, clean_df: pd.DataFrame, sample_df: pd.DataFrame, allocation_details: Dict[str, pd.DataFrame]):
    """
    Creates the comprehensive sampling_distribution_report.xlsx with 6 sheets.
    """
    logger.info("Generating Excel Sampling Distribution Report: %s", report_path)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Styles
    navy_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    steel_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    zebra_light = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

    font_title = Font(name="Calibri", size=14, bold=True, color="1F497D")
    font_subtitle = Font(name="Calibri", size=10, italic=True)
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_regular = Font(name="Calibri", size=11)

    thin_border = Side(border_style="thin", color="D9D9D9")
    thick_bottom = Side(border_style="medium", color="1F497D")
    border_all = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)
    border_tot = Border(top=thin_border, bottom=thick_bottom)

    def write_table(ws, start_row, headers, rows, total_row=None, number_formats=None):
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=start_row, column=col_idx, value=h)
            cell.fill = navy_fill
            cell.font = font_header
            cell.alignment = Alignment(horizontal="center" if col_idx > 2 else "left", vertical="center")
            cell.border = border_all
        ws.row_dimensions[start_row].height = 24

        curr = start_row + 1
        for r_idx, r in enumerate(rows):
            fill = zebra_light if r_idx % 2 == 1 else None
            for c_idx, val in enumerate(r, 1):
                cell = ws.cell(row=curr, column=c_idx, value=val)
                cell.font = font_regular
                if fill:
                    cell.fill = fill
                cell.border = border_all

                if isinstance(val, (int, float)):
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    if number_formats and c_idx in number_formats:
                        cell.number_format = number_formats[c_idx]
                    elif isinstance(val, int):
                        cell.number_format = "#,##0"
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[curr].height = 20
            curr += 1

        if total_row:
            for c_idx, val in enumerate(total_row, 1):
                cell = ws.cell(row=curr, column=c_idx, value=val)
                cell.font = font_bold
                cell.fill = steel_fill
                cell.border = border_tot
                if isinstance(val, (int, float)):
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    if number_formats and c_idx in number_formats:
                        cell.number_format = number_formats[c_idx]
                    elif isinstance(val, int):
                        cell.number_format = "#,##0"
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[curr].height = 22

    # --- Sheet 1: Population distribution by application ---
    ws1 = wb.create_sheet("Population by Application")
    ws1.views.sheetView[0].showGridLines = True
    ws1.cell(1, 1, value="Population Distribution by Application (N = 104,857)").font = font_title
    ws1.cell(2, 1, value="Source: combined_clean_reviews.csv prior to sampling.").font = font_subtitle

    h1 = ["Application Name", "Population (N)", "Population Share (%)", "1★", "2★", "3★", "4★", "5★", "Mean Rating"]
    d1 = []
    for app in ["Day One", "Daylio", "Wysa", "Reflection", "Mindsera"]:
        sub = clean_df[clean_df["App_Name"] == app]
        cnt = len(sub)
        vc = sub["Rating"].value_counts().to_dict()
        d1.append([app, cnt, cnt / len(clean_df), vc.get(1, 0), vc.get(2, 0), vc.get(3, 0), vc.get(4, 0), vc.get(5, 0), round(sub["Rating"].mean(), 3)])

    all_vc = clean_df["Rating"].value_counts().to_dict()
    tot1 = ["Total Population", len(clean_df), 1.0, all_vc.get(1, 0), all_vc.get(2, 0), all_vc.get(3, 0), all_vc.get(4, 0), all_vc.get(5, 0), round(clean_df["Rating"].mean(), 3)]
    write_table(ws1, 4, h1, d1, tot1, number_formats={2: "#,##0", 3: "0.00%", 4: "#,##0", 5: "#,##0", 6: "#,##0", 7: "#,##0", 8: "#,##0", 9: "0.000"})

    # --- Sheet 2: Sample distribution by application ---
    ws2 = wb.create_sheet("Sample by Application")
    ws2.views.sheetView[0].showGridLines = True
    ws2.cell(1, 1, value="Sample Distribution by Application (n = 5,000)").font = font_title
    ws2.cell(2, 1, value="Two-stage stratified random sample selected with random_state = 42.").font = font_subtitle

    h2 = ["Application Name", "Target Sample (n)", "Sampling Fraction (%)", "Sample Share (%)", "1★", "2★", "3★", "4★", "5★", "Mean Rating"]
    d2 = []
    for app in ["Day One", "Daylio", "Wysa", "Reflection", "Mindsera"]:
        sub_s = sample_df[sample_df["App_Name"] == app]
        sub_p = clean_df[clean_df["App_Name"] == app]
        cnt_s = len(sub_s)
        cnt_p = len(sub_p)
        vc_s = sub_s["Rating"].value_counts().to_dict()
        d2.append([app, cnt_s, cnt_s / cnt_p, cnt_s / len(sample_df), vc_s.get(1, 0), vc_s.get(2, 0), vc_s.get(3, 0), vc_s.get(4, 0), vc_s.get(5, 0), round(sub_s["Rating"].mean(), 3)])

    all_vc_s = sample_df["Rating"].value_counts().to_dict()
    tot2 = ["Total Sample", len(sample_df), len(sample_df) / len(clean_df), 1.0, all_vc_s.get(1, 0), all_vc_s.get(2, 0), all_vc_s.get(3, 0), all_vc_s.get(4, 0), all_vc_s.get(5, 0), round(sample_df["Rating"].mean(), 3)]
    write_table(ws2, 4, h2, d2, tot2, number_formats={2: "#,##0", 3: "0.00%", 4: "0.00%", 5: "#,##0", 6: "#,##0", 7: "#,##0", 8: "#,##0", 9: "#,##0", 10: "0.000"})

    # --- Sheet 3: Population rating distribution ---
    ws3 = wb.create_sheet("Population Rating Dist")
    ws3.views.sheetView[0].showGridLines = True
    ws3.cell(1, 1, value="Population Rating Distribution Across All Applications").font = font_title
    ws3.cell(2, 1, value="Counts and proportions across star rating categories.").font = font_subtitle

    h3 = ["Rating Tier", "Day One", "Daylio", "Wysa", "Reflection", "Mindsera", "Total Reviews", "Population Share (%)"]
    d3 = []
    crosstab_pop = pd.crosstab(clean_df["Rating"], clean_df["App_Name"])
    for r in range(1, 6):
        r_tot = (clean_df["Rating"] == r).sum()
        d3.append([
            f"{r} Star",
            crosstab_pop.loc[r, "Day One"],
            crosstab_pop.loc[r, "Daylio"],
            crosstab_pop.loc[r, "Wysa"],
            crosstab_pop.loc[r, "Reflection"],
            crosstab_pop.loc[r, "Mindsera"],
            r_tot,
            r_tot / len(clean_df)
        ])
    tot3 = ["Total", len(clean_df[clean_df["App_Name"]=="Day One"]), len(clean_df[clean_df["App_Name"]=="Daylio"]), len(clean_df[clean_df["App_Name"]=="Wysa"]), len(clean_df[clean_df["App_Name"]=="Reflection"]), len(clean_df[clean_df["App_Name"]=="Mindsera"]), len(clean_df), 1.0]
    write_table(ws3, 4, h3, d3, tot3, number_formats={2: "#,##0", 3: "#,##0", 4: "#,##0", 5: "#,##0", 6: "#,##0", 7: "#,##0", 8: "0.00%"})

    # --- Sheet 4: Sample rating distribution ---
    ws4 = wb.create_sheet("Sample Rating Dist")
    ws4.views.sheetView[0].showGridLines = True
    ws4.cell(1, 1, value="Sample Rating Distribution Across All Applications").font = font_title
    ws4.cell(2, 1, value="Counts and proportions across star rating categories in the 5,000 research sample.").font = font_subtitle

    h4 = ["Rating Tier", "Day One", "Daylio", "Wysa", "Reflection", "Mindsera", "Total Sample", "Sample Share (%)"]
    d4 = []
    crosstab_sam = pd.crosstab(sample_df["Rating"], sample_df["App_Name"])
    for r in range(1, 6):
        r_tot = (sample_df["Rating"] == r).sum()
        d4.append([
            f"{r} Star",
            crosstab_sam.loc[r, "Day One"],
            crosstab_sam.loc[r, "Daylio"],
            crosstab_sam.loc[r, "Wysa"],
            crosstab_sam.loc[r, "Reflection"],
            crosstab_sam.loc[r, "Mindsera"],
            r_tot,
            r_tot / len(sample_df)
        ])
    tot4 = ["Total", len(sample_df[sample_df["App_Name"]=="Day One"]), len(sample_df[sample_df["App_Name"]=="Daylio"]), len(sample_df[sample_df["App_Name"]=="Wysa"]), len(sample_df[sample_df["App_Name"]=="Reflection"]), len(sample_df[sample_df["App_Name"]=="Mindsera"]), len(sample_df), 1.0]
    write_table(ws4, 4, h4, d4, tot4, number_formats={2: "#,##0", 3: "#,##0", 4: "#,##0", 5: "#,##0", 6: "#,##0", 7: "#,##0", 8: "0.00%"})

    # --- Sheet 5: Application × Rating allocation matrix ---
    ws5 = wb.create_sheet("App x Rating Allocation Matrix")
    ws5.views.sheetView[0].showGridLines = True
    ws5.cell(1, 1, value="Application × Rating Stratification & Largest-Remainder Audit Matrix").font = font_title
    ws5.cell(2, 1, value="Audit trail showing exact quota, floored allocation, remainders, priority rank, and final sampled counts.").font = font_subtitle

    h5 = ["App Name", "Rating", "Stratum Pop", "Exact Quota", "Floor Quota", "Remainder", "Priority Rank", "Adjustment", "Allocated Target", "Actual Sampled", "Status"]
    d5 = []
    for app in ["Day One", "Daylio", "Wysa", "Reflection", "Mindsera"]:
        sub_s = sample_df[sample_df["App_Name"] == app]
        detail = allocation_details[app]
        for r in range(1, 6):
            row_d = detail.loc[r]
            act_cnt = (sub_s["Rating"] == r).sum()
            match_str = "MATCH (PASS)" if act_cnt == row_d["Allocated_Sample"] else "MISMATCH (FAIL)"
            d5.append([
                app,
                f"{r} Star",
                int(row_d["Population"]),
                float(row_d["Exact_Quota"]),
                int(row_d["Floor_Quota"]),
                float(row_d["Remainder"]),
                int(row_d["Rank"]) if row_d["Rank"] > 0 else "-",
                int(row_d["Adjustment"]),
                int(row_d["Allocated_Sample"]),
                act_cnt,
                match_str
            ])

    tot5 = ["Total", "All Strata", len(clean_df), 5000.0, sum(d[4] for d in d5), "-", "-", sum(d[7] for d in d5), 5000, 5000, "ALL MATCH"]
    write_table(ws5, 4, h5, d5, tot5, number_formats={3: "#,##0", 4: "0.0000", 5: "#,##0", 6: "0.0000", 9: "#,##0", 10: "#,##0"})

    # --- Sheet 6: Sampling methodology and parameters ---
    ws6 = wb.create_sheet("Methodology and Parameters")
    ws6.views.sheetView[0].showGridLines = True
    ws6.cell(1, 1, value="Sampling Methodology, Parameters & Verification Summary").font = font_title
    ws6.cell(2, 1, value="Technical specification for Phase 2 research sampling design.").font = font_subtitle

    h6 = ["Parameter / Quality Dimension", "Configured Specification", "Observed Value", "Validation Result", "Research Implication"]
    d6 = [
        ["Target Sample Size", "5,000 reviews", f"{len(sample_df):,} reviews", "PASS", "Adequate statistical power for NLP & LLM research."],
        ["Sampling Method", "Two-Stage Stratified Random Sampling", "Stage 1: App, Stage 2: Rating (Largest-Remainder)", "PASS", "Prevents domination by 5★ reviews and large apps."],
        ["Randomization Seed", "random_state = 42", "42 (Deterministic)", "PASS", "100% reproducible across Python environments."],
        ["Replacement Policy", "Without Replacement (Strict)", "Zero duplicate Review_IDs", "PASS", "No duplicate record bias in sampling frame."],
        ["Day One Target", "1,000 reviews", f"{(sample_df['App_Name']=='Day One').sum():,} reviews", "PASS", "19.63% sampling fraction."],
        ["Daylio Target", "1,538 reviews", f"{(sample_df['App_Name']=='Daylio').sum():,} reviews", "PASS", "2.93% sampling fraction."],
        ["Wysa Target", "1,537 reviews", f"{(sample_df['App_Name']=='Wysa').sum():,} reviews", "PASS", "3.31% sampling fraction."],
        ["Reflection Target", "859 reviews (Census)", f"{(sample_df['App_Name']=='Reflection').sum():,} reviews", "PASS", "100.0% census of available population."],
        ["Mindsera Target", "66 reviews (Census)", f"{(sample_df['App_Name']=='Mindsera').sum():,} reviews", "PASS", "100.0% census of available population."],
        ["Integrity Verification", "100% records traceable to clean dataset", "5,000 / 5,000 records verified", "PASS", "Complete audit trail maintained."],
        ["Readiness for Annotation", "Ready for LLM annotation / codebook coding", "All quality gates passed", "READY", "Dataset is ready for Phase 3."],
    ]
    write_table(ws6, 4, h6, d6)

    # Column widths
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            sheet.column_dimensions[col_letter].width = max(max_len + 4, 14)

    wb.save(report_path)
    logger.info("Saved Excel report successfully: %s", report_path)


def generate_summary_csv(csv_path: str, clean_df: pd.DataFrame, sample_df: pd.DataFrame):
    """
    Creates the compact sampling_summary.csv file.
    """
    logger.info("Generating Sampling Summary CSV: %s", csv_path)
    records = [
        {"Category": "Overview", "Metric": "Original Population Size", "Value": len(clean_df)},
        {"Category": "Overview", "Metric": "Final Sample Size", "Value": len(sample_df)},
        {"Category": "Overview", "Metric": "Random Seed", "Value": RANDOM_STATE},
        {"Category": "Overview", "Metric": "Sampling Method", "Value": "Two-Stage Stratified Random Sampling (Largest-Remainder Method)"},
        {"Category": "Overview", "Metric": "Sampling Replacement", "Value": "Without Replacement (Strict)"},
    ]

    # Application allocations
    for app in ["Day One", "Daylio", "Wysa", "Reflection", "Mindsera"]:
        pop_n = (clean_df["App_Name"] == app).sum()
        sam_n = (sample_df["App_Name"] == app).sum()
        records.append({
            "Category": "Application Allocation",
            "Metric": f"{app} (Population -> Sample)",
            "Value": f"{pop_n:,} -> {sam_n:,} ({sam_n/pop_n*100:.2f}%)"
        })

    # Rating allocations
    for r in range(1, 6):
        pop_r = (clean_df["Rating"] == r).sum()
        sam_r = (sample_df["Rating"] == r).sum()
        records.append({
            "Category": "Rating Allocation",
            "Metric": f"{r} Star (Population -> Sample)",
            "Value": f"{pop_r:,} -> {sam_r:,} ({sam_r/pop_r*100:.2f}%)"
        })

    summary_df = pd.DataFrame(records)
    summary_df.to_csv(csv_path, index=False, encoding="utf-8")
    logger.info("Saved Summary CSV successfully.")


def generate_requirements(req_path: str):
    """Saves requirements.txt for the sampling pipeline."""
    content = """pandas>=2.0.0
openpyxl>=3.1.0
numpy>=1.22.0
matplotlib>=3.7.0
"""
    with open(req_path, "w", encoding="utf-8") as f:
        f.write(content.strip() + "\n")
    logger.info("Saved requirements.txt at: %s", req_path)


def main():
    paths = resolve_paths()
    logger.info("Starting Phase 2 Stratified Sampling Pipeline")
    logger.info("Input Cleaned Dataset: %s", paths["clean_csv"])
    logger.info("Output Folder: %s", paths["output_dir"])

    clean_df = pd.read_csv(paths["clean_csv"])

    # Execute sampling
    sample_5000, allocation_details = execute_sampling(clean_df, random_seed=RANDOM_STATE)

    # Validate sample
    validate_sample(sample_5000, clean_df, allocation_details)

    # Save deliverables
    sample_5000.to_csv(paths["sample_csv"], index=False, encoding="utf-8")
    logger.info("Saved final research sample to: %s", paths["sample_csv"])

    generate_excel_report(paths["report_xlsx"], clean_df, sample_5000, allocation_details)
    generate_summary_csv(paths["summary_csv"], clean_df, sample_5000)
    generate_visualization(clean_df, sample_5000, paths["plot_png"])
    generate_requirements(paths["req_txt"])

    print("\n" + "=" * 75)
    print("PHASE 2 STRATIFIED SAMPLING COMPLETED SUCCESSFULLY")
    print("=" * 75)
    print(f"Total Reviews Sampled : {len(sample_5000):,} / {len(clean_df):,} (Seed = {RANDOM_STATE})")
    print(f"Unique Review_IDs     : {sample_5000['Review_ID'].nunique():,} (Zero duplicates)")
    print("-" * 75)
    print("Application Sample Breakdown:")
    for app, target in APPLICATION_TARGETS.items():
        pop_cnt = (clean_df["App_Name"] == app).sum()
        act_cnt = (sample_5000["App_Name"] == app).sum()
        print(f"  - {app:<15}: {act_cnt:>5,} / {pop_cnt:>7,} ({act_cnt/pop_cnt*100:>6.2f}%)")
    print("-" * 75)
    print("Rating Sample Breakdown:")
    for r in range(1, 6):
        pop_r = (clean_df["Rating"] == r).sum()
        act_r = (sample_5000["Rating"] == r).sum()
        print(f"  - {r} Star          : {act_r:>5,} / {pop_r:>7,} ({act_r/pop_r*100:>6.2f}%)")
    print("-" * 75)
    print("Generated Deliverables in E:\\Insight\\phase_2_sampling:")
    print(f"  1. {paths['sample_csv']}")
    print(f"  2. {paths['report_xlsx']}")
    print(f"  3. {paths['summary_csv']}")
    print(f"  4. {paths['plot_png']}")
    print(f"  5. {paths['script_py']}")
    print(f"  6. {paths['req_txt']}")
    print("=" * 75)


if __name__ == "__main__":
    main()
