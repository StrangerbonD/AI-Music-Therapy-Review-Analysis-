import os
import glob
import csv
from collections import Counter


def analyze_dataset_reviews(dataset_dir: str):
    """
    Scans all CSV files in dataset_dir, calculates total reviews,
    and counts 1, 2, 3, 4, and 5-star ratings from the 'score' column.
    """
    csv_pattern = os.path.join(dataset_dir, "*.csv")
    csv_files = sorted(glob.glob(csv_pattern))

    if not csv_files:
        print(f"No CSV files found in: {dataset_dir}")
        return

    results = []

    for file_path in csv_files:
        filename = os.path.basename(file_path)
        star_counts = Counter()
        total_reviews = 0
        unknown_scores = Counter()

        # 'utf-8-sig' cleanly handles UTF-8 Byte Order Marks (\ufeff)
        with open(file_path, mode="r", encoding="utf-8-sig", errors="replace") as f:
            reader = csv.DictReader(f)
            for row in reader:
                total_reviews += 1
                score_str = row.get("score", "").strip()

                if score_str in ("1", "2", "3", "4", "5"):
                    star_counts[int(score_str)] += 1
                else:
                    unknown_scores[score_str] += 1

        results.append({
            "filename": filename,
            "total": total_reviews,
            "1_star": star_counts[1],
            "2_star": star_counts[2],
            "3_star": star_counts[3],
            "4_star": star_counts[4],
            "5_star": star_counts[5],
            "unknown": dict(unknown_scores)
        })

    # Print Table Header
    print(f"{'File Name':<50} {'Total':>10} {'1 Star':>10} {'2 Star':>10} {'3 Star':>10} {'4 Star':>10} {'5 Star':>10}")
    print("-" * 115)

    # Print Rows
    for r in results:
        print(f"{r['filename']:<50} {r['total']:>10,} {r['1_star']:>10,} {r['2_star']:>10,} {r['3_star']:>10,} {r['4_star']:>10,} {r['5_star']:>10,}")

    print("-" * 115)

    # Print Totals
    total_all = sum(r["total"] for r in results)
    total_1 = sum(r["1_star"] for r in results)
    total_2 = sum(r["2_star"] for r in results)
    total_3 = sum(r["3_star"] for r in results)
    total_4 = sum(r["4_star"] for r in results)
    total_5 = sum(r["5_star"] for r in results)

    print(f"{'Total':<50} {total_all:>10,} {total_1:>10,} {total_2:>10,} {total_3:>10,} {total_4:>10,} {total_5:>10,}")

    if total_all > 0:
        print("\nPercentage Distribution:")
        print(f"  5 Stars: {total_5:,} ({total_5 / total_all * 100:.2f}%)")
        print(f"  4 Stars: {total_4:,} ({total_4 / total_all * 100:.2f}%)")
        print(f"  3 Stars: {total_3:,} ({total_3 / total_all * 100:.2f}%)")
        print(f"  2 Stars: {total_2:,} ({total_2 / total_all * 100:.2f}%)")
        print(f"  1 Star : {total_1:,} ({total_1 / total_all * 100:.2f}%)")


if __name__ == "__main__":
    # Resolve Dataset path relative to the script location or fallback to E:\Insight\Dataset
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(script_dir, "..", "Dataset"),
        os.path.join(script_dir, "Dataset"),
        r"E:\Insight\Dataset",
    ]
    dataset_folder = next((os.path.abspath(p) for p in candidates if os.path.exists(p)), r"E:\Insight\Dataset")

    analyze_dataset_reviews(dataset_folder)
