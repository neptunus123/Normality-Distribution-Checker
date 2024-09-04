import pandas as pd
import numpy as np
import scipy.stats as stats
import tkinter as tk
from tkinter import Toplevel, Button
from tkinter.filedialog import askopenfilename, asksaveasfilename


def load_data():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = askopenfilename(
        title="Select an Excel file",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
    )
    if not file_path:
        root.destroy()
        return None, None
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names

    def select_sheet(name):
        nonlocal sheet_name
        sheet_name = name
        top.destroy()
        root.destroy()

    top = Toplevel(root)
    sheet_name = None
    for sheet in sheet_names:
        button = Button(
            top, text=sheet, command=lambda sheet=sheet: select_sheet(sheet)
        )
        button.pack()
    top.mainloop()
    return file_path, sheet_name


def save_data():
    root = tk.Tk()
    root.withdraw()
    file_path = asksaveasfilename(
        title="Save the Excel file",
        filetypes=[("Excel files", "*.xlsx")],
        defaultextension=".xlsx",
    )
    root.destroy()
    return file_path


def analyze_data(input_file, sheet_name, output_file):
    df = pd.read_excel(input_file, sheet_name=sheet_name)

    if "ID" in df.columns:
        df = df.drop(columns=["ID"])

    results = []
    for col in df.columns:
        data = df[col].dropna()
        if pd.api.types.is_numeric_dtype(data) and len(set(data)) > 10:
            if len(data) > 3:
                try:
                    stat, p_value = stats.shapiro(data)
                    if p_value > 0.05:
                        results.append(
                            (
                                col,
                                "Normal",
                                "Mean and Std Dev",
                                f"Mean={data.mean():.2f}, Std Dev={data.std():.2f}",
                            )
                        )
                    else:
                        results.append(
                            (
                                col,
                                "Not Normal",
                                "Median and Min-Max",
                                f"Median={data.median()}, Min-Max=({data.min()}, {data.max()})",
                            )
                        )
                except Exception as e:
                    results.append((col, "Error in Test", str(e), ""))
            else:
                results.append((col, "Insufficient Data", "Not Applicable", ""))
        elif pd.api.types.is_numeric_dtype(data):
            value_counts = data.value_counts()
            percentages = value_counts / len(data) * 100
            formatted_values = ", ".join(
                [
                    f"{int(idx)} values: n={count} (%{percent:.2f})"
                    for idx, count, percent in zip(
                        value_counts.index, value_counts, percentages
                    )
                ]
            )
            results.append(
                (
                    col,
                    "Categorical or Non-numeric",
                    "Counts and Percentages",
                    formatted_values,
                )
            )
        else:
            results.append((col, "String type data content", "", ""))

    results_df = pd.DataFrame(
        results, columns=["Column", "Type", "Assumed Measure", "Measure Values"]
    )
    results_df.to_excel(output_file, index=False)


def main():
    input_file, sheet_name = load_data()
    if input_file and sheet_name:
        output_file = save_data()
        if output_file:
            analyze_data(input_file, sheet_name, output_file)
            print("Analysis completed and file saved.")


if __name__ == "__main__":
    main()
