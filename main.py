import argparse
import logging
from collections import Counter

import pandas as pd
from fuzzywuzzy import fuzz

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def load_excel(file_path):
    try:
        logging.info(f"Input file: '{file_path}'..")
        return pd.read_excel(file_path)
    except Exception as e:
        logging.error(f"An error occurred: {e}")
        return None


def compute_similarity_groups_index(df, column_index, threshold=80):
    column_values = df.iloc[:, column_index].astype(str)
    num_rows = len(column_values)
    similarity_groups_index = [-1] * num_rows  # -1 indicates unassigned
    current_group = 0
    for i in range(num_rows):
        if i % 100 == 0:
            logging.info(f"Processing {i} out of {num_rows} rows..")
        if similarity_groups_index[i] != -1:  # This row is already part of a Similarity-Group. Skip
            continue
        current_group += 1
        similarity_groups_index[i] = current_group  # Assign a new Similarity-Group index
        for j in range(i + 1, num_rows):
            if similarity_groups_index[j] == -1:
                similarity = fuzz.ratio(column_values[i], column_values[j])
                if similarity >= threshold:
                    similarity_groups_index[j] = current_group
    logging.info(f"Finished processing {num_rows} rows")
    total_count = num_rows
    groups_count = current_group
    duplicated_rows = total_count - groups_count
    duplicated_rows_percent = 100 * (duplicated_rows / total_count)
    logging.info(f"Found {duplicated_rows} duplicated rows ({duplicated_rows_percent:.1f}%)")
    return similarity_groups_index


def save_excel(df, output_path):
    df.to_excel(output_path, index=False)


def add_group_indices_to_df(df, group_col_name, column_index, group_indices):
    df.insert(column_index, group_col_name, group_indices)
    return df


def reorder_rows_by_group_size(df, group_col_name):
    group_counts = df[group_col_name].value_counts().to_dict()
    df['GroupSize'] = df[group_col_name].map(group_counts)
    df = df.sort_values(by=['GroupSize', group_col_name], ascending=[False, True]).drop(columns=['GroupSize'])
    return df


def reorder_rows_by_group_index(df, group_col_name):
    df = df.sort_values(by=[group_col_name], ascending=True)
    return df


def reindex_groups_by_size(group_indices):
    group_counts = Counter(group_indices)
    sorted_groups = sorted(group_counts.keys(), key=lambda k: group_counts[k], reverse=True)
    new_index_map = {old_idx: new_idx for new_idx, old_idx in enumerate(sorted_groups, 1)}
    new_group_indices = [new_index_map[old_idx] for old_idx in group_indices]
    return new_group_indices


def main(file_path, column_index, threshold, output_path, encoding=None):
    logging.info("Starting the similarity grouping process")
    df = load_excel(file_path)
    if df is not None:
        similarity_groups_index = compute_similarity_groups_index(df, column_index, threshold)
        similarity_groups_index = reindex_groups_by_size(similarity_groups_index)
        df_edit = add_group_indices_to_df(df, 'GroupIndex', column_index, similarity_groups_index)
        df_edit = reorder_rows_by_group_index(df_edit, 'GroupIndex')
        save_excel(df_edit, output_path)
        logging.info(f"Output file: '{output_path}'")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Excel Similarity Grouping")
    parser.add_argument("--input-file", default="input.xlsx", help="Path to the input Excel file")
    parser.add_argument("--output-file", default="output.xlsx", help="Path to the output Excel file")
    parser.add_argument("--column-index", type=int, default=0, help="Index of the column to analyze (0-based)")
    parser.add_argument("--threshold", type=int, default=80, help="Similarity threshold (0-100)")
    parser.add_argument("--encoding", default='utf-8', help="Encoding of the input file")
    args = parser.parse_args()

    main(args.input_file, args.column_index, args.threshold, args.output_file, args.encoding)
