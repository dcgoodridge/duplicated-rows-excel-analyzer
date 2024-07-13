import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import chardet
import logging
import argparse

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def load_excel(file_path, encoding=None):
    try:
        if encoding:
            # return pd.read_csv(file_path, encoding=encoding)
            return pd.read_excel(file_path)
        else:
            return pd.read_excel(file_path)
    except UnicodeDecodeError as e:
        # Attempt to detect encoding
        with open(file_path, 'rb') as f:
            result = chardet.detect(f.read())
            detected_encoding = result['encoding']
            logging.warning(f"Failed to load with provided encoding. Detected encoding: {detected_encoding}")
            return pd.read_excel(file_path, encoding=detected_encoding)
    except Exception as e:
        logging.error(f"An error occurred: {e}")
        return None

def group_similarities(df, column_index, threshold):
    unique_values = df.iloc[:, column_index].unique()
    groups = {}
    group_index = 0

    logging.info(f"Grouping {len(unique_values)} unique values from column index {column_index}")

    for idx, value in enumerate(unique_values):
        if not any(value in group for group in groups.values()):
            matches = process.extractBests(value, unique_values, scorer=fuzz.token_sort_ratio, score_cutoff=threshold)
            groups[group_index] = [match[0] for match in matches]
            group_index += 1

        if idx % 100 == 0:
            logging.info(f"Processed {idx} out of {len(unique_values)} unique values")

    return groups

def add_group_column(df, column_index, groups):
    group_map = {}
    for group_index, group_values in groups.items():
        for value in group_values:
            group_map[value] = group_index

    df.insert(0, 'Group', df.iloc[:, column_index].map(group_map))
    df.sort_values(by='Group', inplace=True)
    return df

def save_excel(df, output_path):
    df.to_excel(output_path, index=False)

def main(file_path, column_index, threshold, output_path, encoding=None):
    logging.info("Starting the similarity grouping process")
    df = load_excel(file_path, encoding)
    if df is not None:
        groups = group_similarities(df, column_index, threshold)
        df = add_group_column(df, column_index, groups)
        save_excel(df, output_path)
        logging.info(f"Processed file saved to {output_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Excel Similarity Grouping")
    parser.add_argument("--input-file", default="input.xlsx", help="Path to the input Excel file")
    parser.add_argument("--output-file", default="output.xlsx", help="Path to the output Excel file")
    parser.add_argument("--column-index", type=int, default=0, help="Index of the column to analyze (0-based)")
    parser.add_argument("--threshold", type=int, default=80, help="Similarity threshold (0-100)")
    parser.add_argument("--encoding", default='utf-8', help="Encoding of the input file")
    args = parser.parse_args()

    main(args.input_file, args.column_index, args.threshold, args.output_file, args.encoding)
