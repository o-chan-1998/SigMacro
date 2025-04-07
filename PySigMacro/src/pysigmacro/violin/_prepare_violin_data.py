#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-04-01 22:20:15 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/violin/_prepare_violin_data.py
# ----------------------------------------
import os
__FILE__ = (
    "/home/ywatanabe/win/documents/SigMacro/PySigMacro/src/pysigmacro/violin/_prepare_violin_data.py"
)
__DIR__ = os.path.dirname(__FILE__)
# ----------------------------------------
import numpy as np
import pandas as pd
from scipy import stats

def prepare_violin_data(data_dict, space=0.05, n=256, width_factor=0.5):
    """
    Calculate kernel density data for violin plots in a format compatible with SigmaPlot.
    Ensures x-axis positions match box plots exactly.

    Parameters:
    - data_dict: dictionary structured like the output from _gen_data_box/violin
                 with keys like 'X 0', 'Y 0', 'BGRA 0', etc.
                 Or pandas DataFrame with similar structure
    - space: spacing between violins (as proportion of position)
    - n: number of intervals for kernel density function (default 256)
    - width_factor: controls width of the violin (0.5 means half-width on each side)

    Returns:
    - DataFrame ready to export for SigmaPlot
    """
    # Convert DataFrame to dict if needed
    if isinstance(data_dict, pd.DataFrame):
        data_dict = data_dict.to_dict('list')

    # Debug the input
    print(f"Input data keys: {list(data_dict.keys())}")

    # Extract data columns - flexible pattern matching to find all data arrays
    y_columns = {}
    x_values = {}
    colors = {}

    # More flexible pattern matching for data extraction
    for key in data_dict:
        # Look for column names that have Y or y and a number
        if isinstance(key, str) and ('Y' in key or 'y' in key) and any(c.isdigit() for c in key):
            # Extract numeric index - take the first number found
            idx_str = ''.join(c for c in key if c.isdigit())
            if idx_str:
                idx = int(idx_str)
                # Convert to numpy array if it's a list
                data = data_dict[key]
                if isinstance(data, list):
                    data = np.array(data)
                # Only include if it's numeric data
                if isinstance(data, np.ndarray) and len(data) > 0:
                    y_columns[idx] = data
                    print(f"Found Y data for index {idx}: shape={data.shape}")

                    # Look for corresponding X name pattern
                    x_patterns = [f"X {idx}", f"x {idx}", f"X{idx}", f"x{idx}",
                                f"Category {idx}", f"category {idx}"]
                    for x_pattern in x_patterns:
                        if x_pattern in data_dict:
                            x_values[idx] = data_dict[x_pattern]
                            print(f"Found X label for index {idx}: {x_values[idx]}")
                            break

                    # Look for corresponding color pattern
                    color_patterns = [f"BGRA {idx}", f"bgra {idx}", f"BGRA{idx}", f"bgra{idx}",
                                    f"Color {idx}", f"color {idx}"]
                    for color_pattern in color_patterns:
                        if color_pattern in data_dict:
                            colors[idx] = data_dict[color_pattern]
                            print(f"Found color for index {idx}")
                            break

    # Also try dict with 'Col1', 'Col2' pattern - SigmaPlot format
    if not y_columns:
        for key in data_dict:
            if isinstance(key, str) and key.startswith("Col"):
                try:
                    col_num = int(key[3:])
                    # Every 9th column from Col10 is X value (index 0, 9, 18...)
                    # Every 9th column from Col14 is Y value (index 4, 13, 22...)
                    col_index = (col_num - 10) % 9
                    group_idx = (col_num - 10) // 9

                    if col_index == 4:  # Y value - category name
                        # Extract actual data from the rest of the file
                        for j in range(1, n+1):
                            # Each group has three columns of data after the initial columns
                            data_col = f"Col{len(data_dict) + group_idx*3 + 1}"
                            if data_col in data_dict and data_dict[data_col][j] is not None:
                                if group_idx not in y_columns:
                                    y_columns[group_idx] = []
                                y_columns[group_idx].append(data_dict[data_col][j])

                        if group_idx in y_columns:
                            y_columns[group_idx] = np.array(y_columns[group_idx])
                            print(f"Found SigmaPlot Y data for group {group_idx}")

                        # Store category name
                        if data_dict[key] and data_dict[key][0] is not None:
                            x_values[group_idx] = data_dict[key][0]
                            print(f"Found SigmaPlot category for group {group_idx}: {x_values[group_idx]}")

                    elif col_index == 8:  # BGRA value
                        if len(data_dict[key]) >= 4:
                            colors[group_idx] = data_dict[key][:4]
                            print(f"Found SigmaPlot color for group {group_idx}")
                except (ValueError, TypeError, IndexError):
                    continue

    # If no valid data found, return empty DataFrame with error message
    if not y_columns:
        print("WARNING: No suitable Y data columns found in input.")
        return pd.DataFrame()

    print(f"Processing {len(y_columns)} data groups for violin plot")

    # Prepare output data structure
    output_data = {}

    # Format for graph properties section
    output_data["_PLOT_TYPE_EXPLANATION"] = ["Plot Type"]
    output_data["PLOT_TYPE"] = ["Violin Plot"]
    output_data["_GRAPH_PARAMS_EXPLANATION"] = ["Graph Parameters"]
    output_data["GRAPH_PARAMS"] = ["Values"]

    # Calculate category positions (match exactly with box plot)
    category_positions = {}
    for j, idx in enumerate(sorted(y_columns.keys()), 0):
        # Box plots use integer positions (1, 2, 3...)
        category_positions[idx] = j + 1

    # Initialize data columns with empty lists - using same structure as box plot
    for j, idx in enumerate(sorted(y_columns.keys()), 1):
        col_index = 10 + (j-1) * 9

        # Get category name or use index
        category_name = x_values.get(idx, f"Category {idx}")
        if not isinstance(category_name, str):
            category_name = f"Category {idx}"

        # Use exact same positions as box plot (integers 1, 2, 3...)
        position = category_positions[idx]

        # Create columns for X, Y and RGBA values
        output_data[f"Col{col_index}"] = [position]  # X column (exact category position)
        output_data[f"Col{col_index+1}"] = [None] # X error
        output_data[f"Col{col_index+2}"] = [None] # X upper
        output_data[f"Col{col_index+3}"] = [None] # X lower
        output_data[f"Col{col_index+4}"] = [category_name] # Y column (category name)
        output_data[f"Col{col_index+5}"] = [None] # Y error
        output_data[f"Col{col_index+6}"] = [None] # Y upper
        output_data[f"Col{col_index+7}"] = [None] # Y lower

        # RGBA column - use provided color or default
        if idx in colors:
            color = colors[idx]
            output_data[f"Col{col_index+8}"] = [int(color[0])]    # B value
            output_data[f"Col{col_index+8}"].append(int(color[1]))  # G value
            output_data[f"Col{col_index+8}"].append(int(color[2]))  # R value
            output_data[f"Col{col_index+8}"].append(int(color[3]) if len(color) > 3 else 255)  # Alpha
        else:
            output_data[f"Col{col_index+8}"] = [0]    # B value
            output_data[f"Col{col_index+8}"].append(0)  # G value
            output_data[f"Col{col_index+8}"].append(0)  # R value
            output_data[f"Col{col_index+8}"].append(255)  # Alpha value

    # Pre-allocate result columns for all groups
    max_rows = n + 2  # Add some buffer
    for j, idx in enumerate(sorted(y_columns.keys()), 1):
        res_col = j * 3 + len(y_columns) * 9 + 10

        output_data[f"Col{res_col}"] = [None] * max_rows
        output_data[f"Col{res_col+1}"] = [None] * max_rows
        output_data[f"Col{res_col+2}"] = [None] * max_rows

    # Make sure all columns have the same length by padding with None
    max_length = max(len(column) for column in output_data.values())
    for key in output_data:
        if len(output_data[key]) < max_length:
            output_data[key].extend([None] * (max_length - len(output_data[key])))

    # Calculate kernel density data for each group
    for j, idx in enumerate(sorted(y_columns.keys()), 1):
        # Get data array
        data = y_columns[idx]
        position = category_positions[idx]

        # Sort the data
        a = np.sort(data)

        # Calculate parameters
        sigma = np.std(a)
        m = len(a)

        # Compute Inter Quartile Range
        q1 = np.percentile(a, 25)
        q3 = np.percentile(a, 75)
        iqr = q3 - q1

        # Compute bandwidth (Silverman's rule of thumb)
        scaled_iqr = iqr / 1.34
        min_val = min(sigma, scaled_iqr) if scaled_iqr > 0 else sigma
        bandwidth = 0.9 * min_val / (m ** 0.2)

        # Define range for kernel density
        z_start = np.min(a) - 3.5 * bandwidth
        z_end = np.max(a) + 3.5 * bandwidth

        # Compute kernel density values
        z_values = np.linspace(z_start, z_end, n+1)
        kde = stats.gaussian_kde(a, bw_method=bandwidth/sigma)
        density = kde(z_values) * bandwidth

        # Scale density to match desired width
        # Box plots typically have width=1, so scale to match
        max_density = np.max(density) if len(density) > 0 else 0.1
        scaled_density = density * (width_factor / max_density)

        # Calculate result column indices
        res_col = j * 3 + len(y_columns) * 9 + 10

        # Fill data into output columns
        # For exact position matching: center violin at category position (position)
        for i, (z, dens) in enumerate(zip(z_values, scaled_density)):
            output_data[f"Col{res_col}"][i+1] = z
            output_data[f"Col{res_col+1}"][i+1] = position + dens
            output_data[f"Col{res_col+2}"][i+1] = position - dens

    # Add the All Groups X-tics column (exact match with box plot positions)
    last_col = max(int(k.replace("Col", "")) for k in output_data if k.startswith("Col")) + 1
    output_data[f"Col{last_col}"] = ["categories"] + [category_positions[idx] for idx in sorted(y_columns.keys())]
    output_data[f"Col{last_col}"].extend([None] * (max_length - len(output_data[f"Col{last_col}"])))

    # Create DataFrame
    df = pd.DataFrame(output_data)

    return df

if __name__ == "__main__":
    # Example usage
    np.random.seed(42)
    data = {}
    for i in range(3):
        # Create base data with uniform distribution
        low = 3 * (i + 1)
        high = 8 * (i + 1)
        base_data = np.random.uniform(low, high, 30)

        # Add outliers for better visualization
        outliers_low = np.random.uniform(low - 2, low - 1, 1)
        outliers_high = np.random.uniform(high + 1, high + 2, 1)
        data[f'Category {i+1}'] = np.concatenate([base_data, outliers_low, outliers_high])

    # Original sample points
    data_df = pd.DataFrame(data)

    # Create the data formatted for SigmaPlot
    violin_df = prepare_violin_data(data)

    # Concatenation
    import pysigmacro as psm
    data_and_violin_df = psm.data.create_padded_df(data_df, violin_df)

    __import__("ipdb").set_trace()
    # Export to CSV for SigmaPlot
    # violin_df.to_csv('violin_plot_data.csv', index=False)
    data_and_violin_df.to_csv('violin_plot_data.csv', index=False)
    print("Data exported to 'violin_plot_data.csv'")

# EOF