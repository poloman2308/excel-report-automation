import argparse
from report_generator.sales_report import SalesReport

def main():
    parser = argparse.ArgumentParser(description="Generate Sales Excel Report")

    parser.add_argument(
        '--input', required=True, help='Path to input CSV file (e.g., data/sales_march.csv)'
    )
    parser.add_argument(
        '--output_dir', default='reports', help='Directory to save the Excel report'
    )
    parser.add_argument(
        '--logo', default=None, help='Path to logo image (optional)'
    )

    args = parser.parse_args()

    report = SalesReport(
        input_file=args.input,
        output_dir=args.output_dir,
        logo_path=args.logo
    )

    report.generate()

if __name__ == "__main__":
    main()
