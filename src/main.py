"""CLI entry point for python-excel-report."""

import argparse
from pathlib import Path

from rich.console import Console
from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn

from src.data_loader import load_folder
from src.analyzer import summarize_by_category, monthly_totals, top_products, kpi_summary
from src.report_builder import build_report

console = Console()

# Default data directory relative to this file's package root
_DEFAULT_DATA = Path(__file__).parent.parent / "data"
_DEFAULT_OUTPUT = Path("report.xlsx")


def parse_args():
    parser = argparse.ArgumentParser(
        description="Generate a professional Excel sales report from CSV/xlsx files."
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=_DEFAULT_DATA,
        help="Folder containing CSV/xlsx input files (default: ./data)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=_DEFAULT_OUTPUT,
        help="Output Excel file path (default: report.xlsx)",
    )
    parser.add_argument(
        "--top",
        type=int,
        default=5,
        help="Number of top products to include (default: 5)",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(),
        TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
        console=console,
        transient=True,
    ) as progress:
        task = progress.add_task("Loading data...", total=4)

        # Step 1: Load
        df = load_folder(args.input)
        progress.update(task, advance=1, description="Analyzing data...")

        # Step 2: Analyze
        cat_sum = summarize_by_category(df)
        monthly = monthly_totals(df)
        top_prods = top_products(df, n=args.top)
        kpis = kpi_summary(df)
        progress.update(task, advance=1, description="Building report...")

        # Step 3: Build
        build_report(df, cat_sum, monthly, top_prods, kpis, args.output)
        progress.update(task, advance=1, description="Done!")
        progress.update(task, advance=1)

    console.print(f"\n[bold green]Report saved:[/bold green] {args.output.resolve()}")
    console.print(f"[cyan]Total Revenue:[/cyan]  ${kpis['total_revenue']:,.0f}")
    console.print(f"[cyan]Top Product:[/cyan]    {kpis['top_product']}")
    console.print(f"[cyan]Months covered:[/cyan] {kpis['num_months']}")


if __name__ == "__main__":
    main()
