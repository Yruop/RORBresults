import csv
import os
import subprocess
import tkinter as tk
from tkinter import filedialog


def process_line(line):
    parts = line.split()

    for i, part in enumerate(parts):
        if part in ["min", "hour"]:
            if i > 0 and parts[i - 1].replace('.', '').isdigit():
                parts[i - 1:i + 1] = [' '.join(parts[i - 1:i + 1])]

    for x, y in enumerate(parts):
        if y == "in":
            if x > 0 and parts[x + 2].isdigit():
                parts[x - 1:x + 2] = [' '.join(parts[x - 1:x + 2])]

    return parts


def main():
    root = tk.Tk()
    root.withdraw()

    print("Select input batch .out file")
    input_file_path = filedialog.askopenfilename(title="Select RORB batch.out file",
                                                 filetypes=[("Text files", "*.out"), ("All files", "*.*")])

    if not input_file_path:
        print("No input file selected. Exiting.")
        return

    input_filename = os.path.basename(input_file_path)

    output_file_path = os.path.join(os.path.dirname(input_file_path), f"{os.path.splitext(input_filename)[0]}.csv")

    if not output_file_path:
        print("No output file selected. Exiting.")
        return

    print("Reading file contents")
    with open(input_file_path, 'r') as infile:
        lines = infile.readlines()

    processed_data = [[' ']+process_line(line) for line in lines]

    with open(output_file_path, 'w', newline='') as csvfile:
        csv_writer = csv.writer(csvfile)
        csv_writer.writerows(processed_data)

    print(f"File successfully processed. Output saved to {output_file_path}")
    subprocess.run(["start", "excel", output_file_path], shell=True)


if __name__ == "__main__":
    main()
